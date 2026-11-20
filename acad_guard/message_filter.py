"""OLE message filter for a COM client thread: retry calls that AutoCAD rejects while it is busy.

AutoCAD answers RPC_E_CALL_REJECTED ("Call was rejected by callee" / "Вызов был отклонен") while it is busy -
hatching, regenerating, switching documents. The refused call did not run, so it is safe to send it again,
writes included. COM does that by itself when the calling thread has a message filter whose
RetryRejectedCall asks for a retry; pywin32 cannot register one, so this does it with ctypes.

    message_filter.register(timeout_s=20)   # once, on the thread that makes the COM calls (an STA)
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes

SERVERCALL_ISHANDLED = 0
SERVERCALL_RETRYLATER = 2
PENDINGMSG_WAITDEFPROCESS = 2
S_OK = 0
E_NOINTERFACE = 0x80004002 - 2**32          # as a signed HRESULT
_IIDS = ("{00000000-0000-0000-C000-000000000046}",   # IUnknown
         "{00000016-0000-0000-C000-000000000046}")   # IMessageFilter


class _GUID(ctypes.Structure):
    _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_ubyte * 8)]

    def __str__(self) -> str:
        d4 = bytes(self.Data4)
        return f"{{{self.Data1:08X}-{self.Data2:04X}-{self.Data3:04X}-{d4[:2].hex().upper()}-{d4[2:].hex().upper()}}}"


_QI = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p))
_REF = ctypes.WINFUNCTYPE(wintypes.ULONG, ctypes.c_void_p)
_INCOMING = ctypes.WINFUNCTYPE(wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD,
                               ctypes.c_void_p)
_RETRY = ctypes.WINFUNCTYPE(wintypes.DWORD, ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD)
_PENDING = ctypes.WINFUNCTYPE(wintypes.DWORD, ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD)


class _VTable(ctypes.Structure):
    _fields_ = [("QueryInterface", _QI), ("AddRef", _REF), ("Release", _REF), ("HandleInComingCall", _INCOMING),
                ("RetryRejectedCall", _RETRY), ("MessagePending", _PENDING)]


class _Filter(ctypes.Structure):
    _fields_ = [("vtbl", ctypes.POINTER(_VTable))]


_alive: list[object] = []      # the filters and their callbacks must outlive the threads they serve
stats = {"retried": 0, "gave_up": 0}


def register(timeout_s: float = 20.0) -> bool:
    """Register the filter on the calling thread. A rejected call is retried every 0.1 s for up to timeout_s,
    then it fails as before. Returns False if COM refused (e.g. not an STA thread)."""
    limit_ms = int(timeout_s * 1000)

    def query_interface(this, riid, ppv):
        if str(riid.contents) in _IIDS:
            ppv[0] = this
            return S_OK
        ppv[0] = None
        return E_NOINTERFACE

    def retry_rejected_call(this, htask, tick_count, reject_type):
        if reject_type == SERVERCALL_RETRYLATER and tick_count < limit_ms:
            stats["retried"] += 1
            return 100                 # ms until COM sends the call again
        stats["gave_up"] += 1
        return 0xFFFFFFFF              # cancel: the caller gets RPC_E_CALL_REJECTED

    vtbl = _VTable(_QI(query_interface), _REF(lambda this: 1), _REF(lambda this: 1),
                   _INCOMING(lambda this, call_type, htask, tick, info: SERVERCALL_ISHANDLED),
                   _RETRY(retry_rejected_call),
                   _PENDING(lambda this, htask, tick, pending: PENDINGMSG_WAITDEFPROCESS))
    obj = _Filter(ctypes.pointer(vtbl))
    _alive.extend((vtbl, obj))
    old = ctypes.c_void_p()
    return ctypes.windll.ole32.CoRegisterMessageFilter(ctypes.byref(obj), ctypes.byref(old)) == S_OK
