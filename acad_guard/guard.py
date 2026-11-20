"""AutoCAD session guard: blocking windows, health, crash recovery and (re)start - as plain AutoCAD or
through SOFiCAD's launcher (start profile in guard_config.json).

Used by the `autocad` MCP wrapper (autocad_guarded.py) around every tool call, and by the
`soficad` server. All functions are synchronous and COM-thread-safe (they initialise COM
themselves); call them through asyncio.to_thread from async code.

Configuration and the dialog rules live in guard_config.json next to this file; every handled
window, crash and restart is appended to events.jsonl.
"""
from __future__ import annotations

import contextlib
import ctypes
import json
import os
import queue
import re
import subprocess
import threading
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as w
import win32gui
import win32process

GUARD_DIR = Path(__file__).resolve().parent
WINDOWS_PS1 = GUARD_DIR / "windows.ps1"
CONFIG_FILE = GUARD_DIR / "guard_config.json"
EVENTS_FILE = GUARD_DIR / "events.jsonl"
STATE_FILE = GUARD_DIR / "session_state.json"
OPEN_WITH_FILE = GUARD_DIR / "open_with_expected.txt"   # when the last .docx export started (expect_open_with)
SHOTS_DIR = GUARD_DIR / "shots"
PROGID = "AutoCAD.Application"
_user32 = ctypes.windll.user32
_lock = threading.RLock()  # one guard action at a time (two MCP servers may race on the same dialog)


class GuardError(RuntimeError):
    pass


# --------------------------------------------------------------------------- config / log

def config() -> dict[str, Any]:
    return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))


def _save_config(cfg: dict[str, Any]) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def log_event(kind: str, **data: Any) -> dict[str, Any]:
    ev = {"time": time.strftime("%Y-%m-%d %H:%M:%S"), "kind": kind, **data}
    with contextlib.suppress(OSError), EVENTS_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(ev, ensure_ascii=False) + "\n")
    return ev


def events(limit: int = 30) -> list[dict[str, Any]]:
    if not EVENTS_FILE.exists():
        return []
    lines = EVENTS_FILE.read_text(encoding="utf-8").splitlines()[-limit:]
    return [json.loads(ln) for ln in lines if ln.strip()]


# --------------------------------------------------------------------------- windows (PowerShell/UIA)

_PS_CMD = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(WINDOWS_PS1)]
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class _PsWorker:
    """windows.ps1 -Mode serve kept running: a request then costs only its UI Automation work (~0.1 s)
    instead of ~0.6 s for starting PowerShell and compiling the helper types (a SOFiCAD question answered
    during a build waited ~1.5 s for list + press). Restarted after a failure; _ps falls back to a
    one-shot run for that request."""

    def __init__(self) -> None:
        self._p: subprocess.Popen | None = None
        self._q: queue.Queue = queue.Queue()
        self._lock = threading.Lock()

    def _start(self) -> None:
        self._p = subprocess.Popen(_PS_CMD + ["-Mode", "serve"], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                   stderr=subprocess.DEVNULL, creationflags=_NO_WINDOW)
        self._q = q = queue.Queue()
        out = self._p.stdout

        def read() -> None:
            for ln in iter(out.readline, b""):
                q.put(ln)
            q.put(None)
        threading.Thread(target=read, daemon=True, name="guard-ps-reader").start()

    def stop(self) -> None:
        if self._p is not None:
            with contextlib.suppress(Exception):
                self._p.kill()
            self._p = None

    def request(self, req: dict[str, Any], timeout: float) -> dict[str, Any]:
        with self._lock:
            try:
                if self._p is None or self._p.poll() is not None:
                    self._start()
                    timeout += 20                       # first request: PowerShell compiles the types
                self._p.stdin.write((json.dumps(req, ensure_ascii=False) + "\n").encode("utf-8"))
                self._p.stdin.flush()
                t0 = time.time()
                while True:
                    ln = self._q.get(timeout=max(0.1, timeout - (time.time() - t0)))
                    if ln is None:
                        raise GuardError("windows.ps1 worker ended")
                    s = ln.decode("utf-8", errors="replace").strip()
                    if s.startswith("{"):
                        res = json.loads(s)
                        if "error" in res and len(res) == 1:
                            raise GuardError(f"windows.ps1 {req.get('mode')}: {res['error'][:400]}")
                        return res
            except (OSError, ValueError, queue.Empty, GuardError):
                self.stop()
                raise


_worker = _PsWorker()


def _ps(mode: str, hwnd: int = 0, button: str = "", out: str = "", timeout: float = 90) -> dict[str, Any]:
    with contextlib.suppress(OSError, ValueError, queue.Empty, GuardError):
        return _worker.request({"mode": mode, "hwnd": hwnd, "button": button, "out": out}, timeout)
    args = ["-Mode", mode] + (["-Hwnd", str(hwnd)] if hwnd else []) + (["-Button", button] if button else []) \
        + (["-Out", out] if out else [])
    r = subprocess.run(_PS_CMD + args, capture_output=True, timeout=timeout, creationflags=_NO_WINDOW)
    text = r.stdout.decode("utf-8", errors="replace").strip()
    if r.returncode != 0 or "{" not in text:
        raise GuardError(f"windows.ps1 {mode} failed: {r.stderr.decode('utf-8', errors='replace')[-400:]}")
    return json.loads(text[text.index("{"):])


def list_windows() -> dict[str, Any]:
    """AutoCAD-related windows with their UI elements (+ blocked/hung flags)."""
    return _ps("list")


def press(hwnd: int, button: str) -> dict[str, Any]:
    return _ps("press", hwnd, button)


def close_window(hwnd: int) -> dict[str, Any]:
    return _ps("close", hwnd)


def screenshot(hwnd: int) -> str:
    SHOTS_DIR.mkdir(exist_ok=True)
    path = SHOTS_DIR / f"{time.strftime('%Y%m%d_%H%M%S')}_{hwnd}.png"
    _ps("shot", hwnd, out=str(path))
    return str(path)


def _is_palette(win: dict[str, Any]) -> bool:
    """AutoCAD's own untitled WPF palettes (floating command line etc.) - never touched."""
    return (win.get("class") or "").startswith("HwndWrapper[") and not win.get("title")


def match_rule(win: dict[str, Any], rules: list[dict[str, Any]] | None = None) -> dict[str, Any] | None:
    for rule in rules if rules is not None else config()["dialog_rules"]:
        if "process" in rule and (win.get("process") or "").lower() != rule["process"].lower():
            continue
        if "title" in rule and not re.search(rule["title"], win.get("title") or ""):
            continue
        if "class" in rule and not re.search(rule["class"], win.get("class") or ""):
            continue
        if "text" in rule and not re.search(rule["text"], " | ".join(e["name"] for e in win.get("elements") or [])):
            continue
        return rule
    return None


_unknown_seen: dict[int, tuple[float, str | None]] = {}   # hwnd -> (when reported, screenshot)
UNKNOWN_REPEAT_S = 300    # an unknown window still open is logged and photographed again only after this


def _buttons(win: dict[str, Any]) -> list[str]:
    return [e["name"] for e in win.get("elements") or [] if e.get("type") in ("Button", "Pane", "Hyperlink")]


def handle_dialogs(only_if_blocked: bool = False) -> dict[str, Any]:
    """Answer every window covered by a rule; dismiss pure info boxes; report (with screenshot) the rest."""
    with _lock:
        cfg = config()
        state = list_windows()
        if only_if_blocked and not state.get("autocad_blocked") and not any(
                match_rule(wn, cfg["dialog_rules"]) for wn in state.get("windows") or []):
            return {"handled": [], "unknown": [], "crash_seen": False, "state": state}
        handled, unknown, crash_seen = [], [], False
        info_buttons = set(cfg.get("info_buttons", []))
        palettes = set(cfg.get("palette_titles", []))
        blocked = bool(state.get("autocad_blocked"))
        for win in state.get("windows") or []:
            if _is_palette(win) or (win.get("title") or "") in palettes:
                continue
            if not win.get("title") and not win.get("elements") and not match_rule(win, cfg["dialog_rules"]):
                continue  # invisible helper window (e.g. cer_dialog's 'Tao Thread Event Target')
            rule = match_rule(win, cfg["dialog_rules"])
            if rule is None:
                if not blocked and win.get("process") == "acad":
                    continue  # a non-modal AutoCAD window (palette, tool window): nothing is blocked
                btns = [b for b in _buttons(win) if b in info_buttons]
                all_btn = [e["name"] for e in win.get("elements") or [] if e.get("type") == "Button"]
                if (cfg.get("auto_dismiss_info_dialogs") and blocked
                        and len(btns) == 1 and len(all_btn) <= 1 and win.get("owned_by_autocad")):
                    rule = {"name": f"info box '{win.get('title')}'", "buttons": btns, "close": True, "kind": "info"}
                else:
                    # wait loops look every 2 s: one log line and screenshot per window, not one per look
                    seen_at, shot = _unknown_seen.get(win["hwnd"], (0.0, None))
                    repeat = time.time() - seen_at < UNKNOWN_REPEAT_S
                    if not repeat:
                        shot = None
                        with contextlib.suppress(Exception):
                            shot = screenshot(win["hwnd"])
                        _unknown_seen[win["hwnd"]] = (time.time(), shot)
                    entry = {k: win.get(k) for k in ("hwnd", "process", "class", "title")}
                    entry |= {"texts": [e["name"] for e in win.get("elements") or [] if e.get("type") == "Text"][:6],
                              "buttons": _buttons(win)[:15], "screenshot": shot}
                    unknown.append(entry)
                    if not repeat:
                        log_event("unknown_dialog", **entry)
                    continue
            if rule.get("kind") == "crash":
                crash_seen = True
            info_shot = None
            if rule.get("kind") == "info":
                with contextlib.suppress(Exception):
                    info_shot = screenshot(win["hwnd"])
            steps = []
            for b in rule.get("buttons", []):
                res = press(win["hwnd"], b)
                steps.append(res)
                if not res.get("still_open"):
                    break
            if (not steps or steps[-1].get("still_open")) and rule.get("close"):
                steps.append(close_window(win["hwnd"]))
            closed = bool(steps) and not steps[-1].get("still_open")
            ev = log_event("dialog", rule=rule["name"], title=win.get("title"), process=win.get("process"),
                           texts=[e["name"] for e in win.get("elements") or [] if e.get("type") == "Text"][:3],
                           actions=[s.get("button") or s.get("method") for s in steps], closed=closed,
                           screenshot=info_shot)
            handled.append(ev)
        return {"handled": handled, "unknown": unknown, "crash_seen": crash_seen}


class DialogWatcher:
    """While a call runs, answer known windows that pop up mid-call (licence notice when a module
    loads, proxy/reference dialogs when a drawing opens, ...). Unknown windows are left alone.

        with guard.DialogWatcher() as wd:
            ...  # blocking COM call
        wd.handled  # what was answered

    A modal question blocks the call until it is answered, so the check runs every 0.2 s: quick_state
    (WinAPI) costs 0.2 ms. At 2.5 s, SOFiCAD's 5 "similar to POS n?" questions cost the Var track drawing
    17 s of its build.
    """

    def __init__(self, interval: float = 0.2) -> None:
        self.interval = interval
        self.handled: list[dict[str, Any]] = []
        self.unknown: list[dict[str, Any]] = []
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True, name="acad-dialog-watcher")

    def _run(self) -> None:
        pythoncom.CoInitialize()
        last_open_with = 0.0
        try:
            while not self._stop.wait(self.interval):
                with contextlib.suppress(Exception):
                    if quick_state().get("blocked"):
                        r = handle_dialogs()
                        self.handled += r["handled"]
                        self.unknown = r["unknown"] or self.unknown
                    if time.time() - last_open_with > 1.0:     # a few ms: once a second is enough
                        last_open_with = time.time()
                        close_open_with()
        finally:
            pythoncom.CoUninitialize()

    def __enter__(self) -> "DialogWatcher":
        self._thread.start()
        return self

    def __exit__(self, *exc: Any) -> None:
        self._stop.set()
        self._thread.join(timeout=30)


# --------------------------------------------------------------------------- Windows "open with"

def expect_open_with() -> None:
    """Call when an action is about to make Windows ask which program opens a file (SOFiCAD's .docx schedule
    with no Word installed). OpenWith windows whose process starts within 2 min of it are closed whenever the
    guard looks (ensure_ready, DialogWatcher), also from another process: left open, one holds the
    foreground and no keystroke reaches AutoCAD (seen: one stayed for an hour after an export)."""
    with contextlib.suppress(OSError):
        OPEN_WITH_FILE.write_text(f"{time.time():.3f}", encoding="utf-8")


def _process_start(hp) -> float | None:
    times = [ctypes.c_ulonglong() for _ in range(4)]                  # FILETIME: 100 ns since 1601
    if not ctypes.windll.kernel32.GetProcessTimes(hp, *[ctypes.byref(t) for t in times]):
        return None
    return times[0].value / 1e7 - 11644473600.0


def close_open_with(window_s: float = 120.0) -> int:
    """Close the OpenWith.exe windows started within window_s after the last expect_open_with (never an
    "open with" the user brought up at another time). Returns how many were closed (WinAPI, ms)."""
    try:
        since = float(OPEN_WITH_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return 0
    # Windows 11's dialog (class "Open With") is not listed by EnumWindows, although visible and in front
    wins: list[int] = []
    win32gui.EnumWindows(lambda h, _: wins.append(h) or True, None)
    h = 0
    while h := win32gui.FindWindowEx(None, h, "Open With", None):
        wins.append(h)
    closed = 0
    for h in dict.fromkeys(wins):
        if not win32gui.IsWindowVisible(h):
            continue
        pid = win32process.GetWindowThreadProcessId(h)[1]
        hp = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)   # PROCESS_QUERY_LIMITED_INFORMATION
        if hp:
            buf, n = ctypes.create_unicode_buffer(512), ctypes.c_ulong(512)
            if ctypes.windll.kernel32.QueryFullProcessImageNameW(hp, 0, buf, ctypes.byref(n)) \
                    and os.path.basename(buf.value).lower() == "openwith.exe":
                started = _process_start(hp)
                if started is not None and since - 2.0 <= started <= since + window_s:
                    win32gui.PostMessage(h, 0x0010, 0, 0)               # WM_CLOSE
                    closed += 1
            ctypes.windll.kernel32.CloseHandle(hp)
    if closed:
        log_event("dialog", rule="Windows 'open with' (after .docx export)", actions=["WM_CLOSE"], closed=True,
                  count=closed)
    return closed


def learn_rule(hwnd: int, button: str | None, close: bool, name: str | None = None) -> dict[str, Any]:
    """Store 'for this window press <button>' (matched by exact title, or process if untitled)."""
    state = list_windows()
    win = next((x for x in state.get("windows") or [] if x["hwnd"] == hwnd), None)
    if win is None:
        raise GuardError(f"window {hwnd} not found")
    rule: dict[str, Any] = {"name": name or f"learned: {win.get('title') or win.get('process')}"}
    if win.get("title"):
        rule["title"] = "^" + re.escape(win["title"]) + "$"
    else:
        rule["process"] = win.get("process")
        rule["class"] = "^" + re.escape(win.get("class") or "") + "$"
    rule["buttons"] = [button] if button else []
    rule["close"] = bool(close or not button)
    cfg = config()
    cfg["dialog_rules"].insert(0, rule)
    _save_config(cfg)
    log_event("rule_learned", rule=rule)
    return rule


# --------------------------------------------------------------------------- fast state (WinAPI, ms)

def _main_windows() -> list[int]:
    found: list[int] = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h).startswith("AfxMDIFrame") \
                and "AutoCAD" in win32gui.GetWindowText(h):
            found.append(h)
        return True

    win32gui.EnumWindows(cb, None)
    return found


def acad_process_running() -> bool:
    r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq acad.exe", "/NH"], capture_output=True,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return b"acad.exe" in r.stdout.lower()


def quick_state() -> dict[str, Any]:
    mains = _main_windows()
    if not mains:
        return {"running": acad_process_running(), "main_window": False}
    h = mains[0]
    return {"running": True, "main_window": True, "title": win32gui.GetWindowText(h),
            "pid": win32process.GetWindowThreadProcessId(h)[1],
            "blocked": not win32gui.IsWindowEnabled(h), "hung": bool(_user32.IsHungAppWindow(h))}


# --------------------------------------------------------------------------- COM probes

def _com_initialised() -> bool:
    kind, qualifier = ctypes.c_int(), ctypes.c_int()
    return ctypes.windll.ole32.CoGetApartmentType(ctypes.byref(kind), ctypes.byref(qualifier)) == 0


@contextlib.contextmanager
def _com():
    """COM for this thread - left alone when the thread already has it: pythoncom.CoInitialize on an
    initialised thread adds no reference, so the CoUninitialize after it ended the thread's COM (seen: the
    snapshot every 30 s in ensure_ready broke a script's main thread - CO_E_NOTINITIALIZED on its next
    GetActiveObject)."""
    if _com_initialised():
        yield
        return
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def com_probe(with_docs: bool = True) -> dict[str, Any]:
    """Ask AutoCAD over COM (fails while a modal window or command-line input blocks it)."""
    with _com():
        try:
            app = w.GetActiveObject(PROGID)
            n = app.Documents.Count  # 0 = only the Start tab is open: a normal, idle state
            doc = app.ActiveDocument if n else None
            res: dict[str, Any] = {"com_ok": True, "version": app.Version,
                                   "active": (doc.FullName or doc.Name) if doc else None}
            if with_docs:
                res["documents"] = [{"name": d.Name, "path": d.FullName, "saved": bool(d.Saved),
                                     "read_only": bool(d.ReadOnly)}
                                    for d in (app.Documents.Item(i) for i in range(n))]
                res["command"] = (doc.GetVariable("CMDNAMES") or None) if doc else None
                arx = [m.lower() for m in app.ListArx()]
                res["soficad_ui_loaded"] = any(m.startswith("sof") and m.endswith("_ui.arx") for m in arx)
            return res
        except Exception as ex:  # noqa: BLE001
            return {"com_ok": False, "error": str(ex)[:200]}


def snapshot() -> dict[str, Any] | None:
    """Remember which drawings are open (to reopen them after a crash)."""
    p = com_probe()
    if not p.get("com_ok"):
        return None
    docs = [d for d in p["documents"] if d["path"] and os.path.isabs(d["path"])]
    if not docs:
        return last_snapshot()  # keep the last real list (e.g. AutoCAD just restarted on the Start tab)
    snap = {"time": time.strftime("%Y-%m-%d %H:%M:%S"), "active": p["active"], "documents": docs}
    STATE_FILE.write_text(json.dumps(snap, ensure_ascii=False, indent=1), encoding="utf-8")
    return snap


def last_snapshot() -> dict[str, Any] | None:
    with contextlib.suppress(OSError, ValueError):
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return None


def health() -> dict[str, Any]:
    qs = quick_state()
    res: dict[str, Any] = {"quick": qs}
    with contextlib.suppress(GuardError):
        res["start_profile"] = start_profile()[0]
    if qs.get("main_window") and not qs.get("blocked") and not qs.get("hung"):
        res["com"] = com_probe()
    if qs.get("blocked") or not qs.get("main_window"):
        with contextlib.suppress(Exception):
            st = list_windows()
            res["windows"] = [{k: x.get(k) for k in ("hwnd", "process", "title")} |
                              {"rule": (match_rule(x) or {}).get("name"), "buttons": _buttons(x)[:10]}
                              for x in st.get("windows") or [] if not _is_palette(x)]
    res["last_snapshot"] = last_snapshot()
    return res


# --------------------------------------------------------------------------- start / restart / recover

def find_acad_exe() -> str | None:
    """The newest installed AutoCAD: <Program Files>\\Autodesk\\AutoCAD 20xx\\acad.exe."""
    base = Path(os.environ.get("ProgramW6432") or os.environ.get("ProgramFiles") or r"C:\Program Files")
    found = sorted((base / "Autodesk").glob("AutoCAD 20*/acad.exe"))
    return str(found[-1]) if found else None


def start_profile(st: dict[str, Any] | None = None) -> tuple[str, dict[str, Any]]:
    """('soficad' | 'autocad', its settings): ACAD_GUARD_PROFILE, else start.profile; 'auto' = SOFiCAD if its
    launcher is installed. A config with the start settings at the top level (before profiles) is 'soficad'."""
    st = st if st is not None else config()["start"]
    if "soficad" not in st and "autocad" not in st:
        return "soficad", st
    name = (os.environ.get("ACAD_GUARD_PROFILE") or st.get("profile") or "auto").strip().lower()
    if name == "auto":
        launcher = st.get("soficad", {}).get("launcher")
        name = "soficad" if launcher and os.path.exists(launcher) else "autocad"
    if name not in ("soficad", "autocad"):
        raise GuardError(f"unknown start profile {name!r}: use auto, soficad or autocad")
    return name, st.get(name, {})


def start_command() -> list[str]:
    name, p = start_profile()
    if name == "soficad" and p.get("launcher") and os.path.exists(p["launcher"]):
        return [p["launcher"], *p.get("launcher_args", [])]
    exe = p.get("acad_exe") or "auto"
    if name == "autocad" and (exe == "auto" or not os.path.exists(exe)):
        exe = find_acad_exe() or exe
    if not os.path.exists(exe):
        raise GuardError(f"AutoCAD not found ({exe!r}): set start.{name}.acad_exe in {CONFIG_FILE}")
    return [exe, *p.get("acad_args", [])]


def start_autocad() -> dict[str, Any]:
    """Start AutoCAD by the start profile: SOFiCAD's launcher (else acad.exe with SOFiCAD's arguments), or
    plain acad.exe. Never via COM Dispatch: that starts AutoCAD without SOFiCAD's UI and API."""
    cmd = start_command()
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(cmd, creationflags=flags, close_fds=True)
    return log_event("start", command=cmd)


def wait_ready(timeout: float | None = None, require_soficad_ui: bool | None = None) -> dict[str, Any]:
    """Wait until AutoCAD answers COM and is not blocked (and, for the soficad profile, has SOFiCAD's UI
    loaded); answer known dialogs meanwhile."""
    st = config()["start"]
    timeout = timeout or st.get("ready_timeout_s", 300)
    name, p = start_profile(st)
    need_ui = p.get("require_soficad_ui", name == "soficad") if require_soficad_ui is None else require_soficad_ui
    t0, handled, unknown, last = time.time(), [], [], {}
    while time.time() - t0 < timeout:
        qs = quick_state()
        if qs.get("main_window") and not qs.get("blocked") and not qs.get("hung"):
            p = com_probe()
            if p.get("com_ok") and (not need_ui or p.get("soficad_ui_loaded")):
                return {"ready": True, "seconds": round(time.time() - t0), "handled": handled,
                        "unknown": unknown, "com": p}
            last = p
        if qs.get("running") or qs.get("main_window"):
            with contextlib.suppress(Exception):
                r = handle_dialogs()
                handled += r["handled"]
                unknown = r["unknown"] or unknown
        time.sleep(2)
    return {"ready": False, "seconds": round(time.time() - t0), "handled": handled, "unknown": unknown,
            "last_probe": last, "quick": quick_state()}


def _wait_exit(timeout: float) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout:
        if not acad_process_running():
            return True
        time.sleep(1)
    return False


def _kill_acad() -> None:
    subprocess.run(["taskkill", "/IM", "acad.exe", "/F"], capture_output=True,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))


def reopen(docs: list[dict[str, Any]], active: str | None = None) -> list[dict[str, Any]]:
    out = []
    with _com():
        app = w.GetActiveObject(PROGID)
        open_now = {app.Documents.Item(i).FullName.lower() for i in range(app.Documents.Count)}
        for d in docs:
            path = d["path"]
            if path.lower() in open_now:
                out.append({"path": path, "status": "already open"})
            elif not os.path.exists(path):
                out.append({"path": path, "status": "file missing"})
            else:
                try:
                    app.Documents.Open(path, bool(d.get("read_only")))
                    out.append({"path": path, "status": "reopened"})
                except Exception as ex:  # noqa: BLE001
                    out.append({"path": path, "status": f"failed: {str(ex)[:120]}"})
        if active:
            for i in range(app.Documents.Count):
                if app.Documents.Item(i).FullName.lower() == active.lower():
                    with contextlib.suppress(Exception):
                        app.Documents.Item(i).Activate()
    return out


def restart(reopen_docs: bool = True, force: bool = False) -> dict[str, Any]:
    """Close AutoCAD cleanly and start it again by the start profile; reopen the drawings.

    Refuses if a drawing has unsaved changes, unless force=True (then they are discarded).
    A hung AutoCAD is killed only with force=True."""
    with _lock:
        snap = snapshot() or last_snapshot()
        p = com_probe()
        if acad_process_running():
            if p.get("com_ok"):
                unsaved = [d["name"] for d in p["documents"] if not d["saved"]]
                if unsaved and not force:
                    raise GuardError(f"unsaved drawings {unsaved}: save them or restart with force=True")
                with _com():
                    app = w.GetActiveObject(PROGID)
                    for i in reversed(range(app.Documents.Count)):
                        with contextlib.suppress(Exception):
                            app.Documents.Item(i).Close(False)
                    with contextlib.suppress(Exception):
                        app.Quit()
                if not _wait_exit(60):
                    if not force:
                        raise GuardError("AutoCAD did not exit within 60 s; use force=True to kill it")
                    _kill_acad()
            else:
                handle_dialogs()
                if not force:
                    raise GuardError("AutoCAD is not answering COM (blocked/hung); use force=True to kill it")
                _kill_acad()
            _wait_exit(30)
        log_event("restart", reason="requested", force=force)
        start_autocad()
        ready = wait_ready()
        res = {"ready": ready, "snapshot": snap}
        if ready.get("ready") and reopen_docs and snap:
            res["reopened"] = reopen(snap["documents"], snap.get("active"))
        return res


def recover(reason: str) -> dict[str, Any]:
    """AutoCAD is gone (crash) or showing a crash report: close reports, start it, reopen drawings."""
    with _lock:
        snap = last_snapshot()
        handled: list[dict[str, Any]] = []
        unknown: list[dict[str, Any]] = []
        # a crash shows a chain of windows (Error Aborting -> Error Report -> 'cannot continue, save?'
        # -> ...): keep answering until the process is gone
        t0 = time.time()
        while acad_process_running() and time.time() - t0 < 120:
            r = handle_dialogs()
            handled += r["handled"]
            unknown = r["unknown"] or unknown
            if unknown and not r["handled"]:
                break  # stuck on a window we do not know - report it instead of killing AutoCAD
            time.sleep(2)
        if acad_process_running():
            qs = quick_state()
            if qs.get("main_window") and not qs.get("hung") and not qs.get("blocked"):
                return {"recovered": False, "note": "AutoCAD is still running and usable", "handled": handled}
            if unknown:
                return {"recovered": False, "note": "AutoCAD waits on an unknown window after the crash",
                        "handled": handled, "unknown": unknown}
            _kill_acad()
            _wait_exit(20)
        log_event("crash_recovery", reason=reason, snapshot=snap)
        start_autocad()
        ready = wait_ready()
        res = {"recovered": bool(ready.get("ready")), "reason": reason,
               "handled_before": [{"rule": h.get("rule"), "title": h.get("title"), "actions": h.get("actions")}
                                  for h in handled],
               "ready": ready, "snapshot": snap}
        if ready.get("ready") and snap:
            res["reopened"] = reopen(snap["documents"], snap.get("active"))
            res["note"] = ("drawings reopened from their last saved state; unsaved changes made before the "
                           "crash are only in AutoCAD's autosave (.sv$) / Drawing Recovery Manager")
        return res


_last_snapshot_at = 0.0


def ensure_ready(start_if_missing: bool = True) -> dict[str, Any]:
    """Pre-call check (fast path: a few ms). Fixes what it can; raises GuardError otherwise.

    Returns a dict describing anything it did (empty 'actions' when all was fine)."""
    global _last_snapshot_at
    actions: list[Any] = []
    with contextlib.suppress(Exception):
        if n := close_open_with():                  # a late "open with" from a schedule export
            actions.append({"closed_open_with": n})
    qs = quick_state()
    if not qs.get("running"):
        if not start_if_missing:
            raise GuardError("AutoCAD is not running")
        snap = last_snapshot()
        actions.append(start_autocad())
        ready = wait_ready()
        if not ready.get("ready"):
            raise GuardError(f"AutoCAD did not become ready: {json.dumps(ready, ensure_ascii=False)[:600]}")
        actions.append({"started": True, "seconds": ready["seconds"], "dialogs": ready["handled"]})
        if snap:
            actions.append({"reopened": reopen(snap["documents"], snap.get("active"))})
        return {"actions": actions, "restarted": True}
    if not qs.get("main_window"):
        ready = wait_ready(timeout=120)  # starting up, or only a crash report is left on screen
        if not ready.get("ready"):
            if any(e.get("rule", "").startswith("AutoCAD crash") for e in ready["handled"]) or not acad_process_running():
                rec = recover("crash report / no main window")
                return {"actions": [rec], "restarted": True}
            raise GuardError(f"AutoCAD has no usable main window: {json.dumps(ready, ensure_ascii=False)[:600]}")
        return {"actions": [{"waited_for_startup": ready["seconds"], "dialogs": ready["handled"]}], "restarted": False}
    if qs.get("blocked"):
        r = handle_dialogs()
        actions += r["handled"]
        if r["crash_seen"]:
            rec = recover("crash dialog")
            return {"actions": actions + [rec], "restarted": True}
        time.sleep(0.5)
        if quick_state().get("blocked"):
            raise GuardError("AutoCAD is blocked by a window the guard does not know: "
                             + json.dumps(r["unknown"], ensure_ascii=False)[:900]
                             + " - decide what to press (acad_dialog_press, remember=True to keep the rule).")
    if qs.get("hung"):
        raise GuardError("AutoCAD is not responding (busy or hung). Wait, or acad_restart(force=True).")
    if time.time() - _last_snapshot_at > 30:
        with contextlib.suppress(Exception):
            if snapshot():
                _last_snapshot_at = time.time()
    return {"actions": actions, "restarted": False}


TOOLS_DIR = Path(__file__).resolve().parent.parent / "tools"


def command_line(tail: int = 6) -> list[str]:
    """Last lines of AutoCAD's command line, the pending input included (UI Automation, ~1 s)."""
    with contextlib.suppress(Exception):
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                            str(TOOLS_DIR / "cmdline.ps1"), "-Tail", str(tail)], capture_output=True, timeout=30,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return [ln for ln in r.stdout.decode("utf-8-sig", errors="replace").splitlines() if ln.strip()]
    return []


def clear_command_line() -> dict[str, Any]:
    """Cancel whatever AutoCAD's command line holds - typed text that was never entered (a stray key
    blocks every COM call: '<unknown>.Count'), or a command waiting for a point - with Esc, Esc.
    These are real keystrokes: AutoCAD comes to the front for a moment (window messages do not reach
    the command line)."""
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                        str(TOOLS_DIR / "acad_keys.ps1"), "-Keys", "{ESC}{ESC}"], capture_output=True, timeout=30,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    time.sleep(0.5)
    res = {"keys": r.stdout.decode("utf-8-sig", errors="replace").strip(), "com_ok": com_probe(with_docs=False).get("com_ok"),
           "command_line": command_line(3)}
    log_event("command_line_cleared", **res)
    return res


def diagnose_failure() -> dict[str, Any]:
    """After a failed tool call: was it a window, a crash, or a waiting command line?"""
    qs = quick_state()
    if not qs.get("running"):
        return {"cause": "crash", "recovery": recover("AutoCAD process disappeared during a call")}
    r = handle_dialogs()
    if r["crash_seen"]:
        return {"cause": "crash", "handled": r["handled"], "recovery": recover("crash dialog after a call")}
    if r["handled"] or r["unknown"]:
        return {"cause": "dialog", "handled": r["handled"], "unknown": r["unknown"]}
    p = com_probe(with_docs=False)
    if not p.get("com_ok"):
        return {"cause": "busy", "command_line": command_line(4),
                "detail": "AutoCAD rejects COM calls without a visible dialog - the command line holds typed "
                "text or waits for input (e.g. '(_>' or a point); clear it with acad_clear_input (Esc, Esc)"}
    return {"cause": "none-detected"}
