"""Entry point of the `autocad` MCP server: unmodified autocad-mcp-pro + the AutoCAD session guard.

autocad-mcp-pro (site-packages/server.py) is imported and run as is, with additions:
  - GuardMiddleware around every tool call:
      before - fast check (ms); answers known dialogs; starts AutoCAD if it is not running (plain,
               or through SOFiCAD's launcher: start profile in guard_config.json); recovers from a
               crash; reconnects the COM backend when AutoCAD's process changed (no more
               "/mcp reconnect" after an AutoCAD restart);
      after a failure that looks AutoCAD-related - diagnoses it (dialog / crash / waiting command
               line), recovers if it was a crash, and adds the diagnosis to the error;
               a call that failed only on a stale COM connection is retried once.
  - acad_* tools: acad_health, acad_dialogs, acad_dialog_press, acad_restart, acad_events.
  - an OLE message filter on autocad-mcp-pro's COM thread (calls a busy AutoCAD rejects are retried) and
    English fallback linetypes for localized AutoCAD.
  - autocad-mcp-pro's own fallback that starts AutoCAD through COM Dispatch is disabled: the guard
    starts it (a COM-started AutoCAD would come up without SOFiCAD's UI and API).
Rules for windows and the start command: guard_config.json; history: events.jsonl.
"""
from __future__ import annotations

import asyncio
import contextlib
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import chain_layout  # noqa: E402
import guard  # noqa: E402
import message_filter  # noqa: E402
import server as acad_server  # noqa: E402  autocad-mcp-pro (site-packages/server.py)
import backends.com_backend as com_backend  # noqa: E402
import config as acad_config  # noqa: E402
import engineering.arch.draw as arch_draw  # noqa: E402
import win32com.client  # noqa: E402
from fastmcp import Context  # noqa: E402
from fastmcp.exceptions import ToolError  # noqa: E402
from fastmcp.server.middleware import Middleware, MiddlewareContext  # noqa: E402
from mcp.types import TextContent  # noqa: E402


from importlib.metadata import version as _pkg_version  # noqa: E402

# autocad-mcp-pro internals this wrapper relies on: server.mcp, server.main, server._registered_tools,
# backends.com_backend._acad_app / _COM_STATE. Pinned in C:\autocad-mcp\constraints.txt (pip.ini).
TESTED_VERSIONS = {"1.6.0"}
ACAD_MCP_VERSION = _pkg_version("autocad-mcp-pro")
if ACAD_MCP_VERSION not in TESTED_VERSIONS:
    guard.log_event("version_warning", package="autocad-mcp-pro", version=ACAD_MCP_VERSION,
                    tested=sorted(TESTED_VERSIONS))
    print(f"[acad-guard] WARNING: autocad-mcp-pro {ACAD_MCP_VERSION} is untested with this wrapper "
          f"(tested: {sorted(TESTED_VERSIONS)})", file=sys.stderr)


def _acad_app_no_dispatch():
    """Replacement for com_backend._acad_app: attach only, never COM-launch AutoCAD."""
    if "app" not in com_backend._COM_STATE:
        try:
            com_backend._COM_STATE["app"] = win32com.client.GetActiveObject(acad_config.settings.cad_progid)
        except Exception as exc:
            raise RuntimeError("AutoCAD is not running (the guard starts it before tool calls - retry, "
                               "or call acad_restart)") from exc
    return com_backend._COM_STATE["app"]


com_backend._acad_app = _acad_app_no_dispatch

# A localized AutoCAD names the standard linetypes in its own language (Russian acadiso.lin: CENTER = "осевая"),
# so autocad-mcp-pro's Linetypes.Load("CENTER", "acadiso.lin") fails there with "undefined linetype". The
# standard names are then loaded from English copies in lin/ (same patterns).
LIN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lin")
_load_linetype_orig = com_backend._load_linetype


def _load_linetype_any_language(doc, name: str, lin_file: str) -> None:
    try:
        _load_linetype_orig(doc, name, lin_file)
    except RuntimeError as exc:
        base = os.path.basename(lin_file).lower()
        if base not in ("acad.lin", "acadiso.lin"):
            raise
        try:
            _load_linetype_orig(doc, name, os.path.join(LIN_DIR, base.replace(".lin", "_en.lin")))
        except RuntimeError:
            raise exc from None


com_backend._load_linetype = _load_linetype_any_language

# autocad-mcp-pro makes every COM call on one executor thread and retries a rejected call only in a few reads,
# so a busy AutoCAD ("Вызов был отклонен" / RPC_E_CALL_REJECTED, e.g. while it hatches the walls of
# arch_plan_from_spec) failed a whole tool call. With a message filter on that thread COM retries the call by
# itself (up to 20 s); the executor's initializer is looked up when the backend connects, i.e. after this.
_com_init_orig = com_backend._com_init


def _com_init_with_retry() -> None:
    _com_init_orig()
    if not message_filter.register(timeout_s=20):
        guard.log_event("message_filter_failed")


com_backend._com_init = _com_init_with_retry

# arch chains (arch_dimension_chains, arch_plan_from_spec) set the DIMSCALE system variable to the plot scale
# while they draw, but dimensions added through COM do not take it: they kept ScaleFactor 1, so at 1:100 their
# 2.5 text and arrows were 2.5 mm in the model - invisible on the sheet (AutoCAD 2027). Give each chain
# dimension the scale itself. Both callers look draw_chains up in the module at call time.
_draw_chains_orig = arch_draw.draw_chains


ARCH_TICK = 4             # AcDimArrowheadType acArrowArchTick
MOVE_TEXT_NO_LEADER = 2   # AcDimTextMovement acMoveTextNoLeader


def _dim_text(dim) -> str:
    """The text a dimension shows: its override, else the measurement at the style's precision, trailing zeros
    suppressed as DIMZIN says (ISO-25: precision 2, "900" not "900.00") and with its decimal separator."""
    if dim.TextOverride:
        return dim.TextOverride
    text = f"{float(dim.Measurement):.{int(dim.PrimaryUnitsPrecision)}f}"
    if dim.SuppressTrailingZeros and "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", dim.DecimalSeparator or ".")


def _untangle_chain_texts(doc, res: dict, factor: float) -> int:
    """Lay out the texts of each chain row with chain_layout.layout_row (no two overlap): measure each text in
    the drawing, keep fitting ones inside their segment (TextInside - AutoCAD's best fit would otherwise move
    some of them out) and set the moved ones' positions. Returns how many texts were moved off their centre."""
    rows: dict[str, list] = {}
    for h, d in zip(res["handles"], res["dimensions"]):
        rows.setdefault(d["row"], []).append((doc.HandleToObject(h), d))
    first = next(iter(rows.values()))[0][0]
    height, gap = float(first.TextHeight) * factor, float(first.TextGap) * factor
    pad = max(gap, 0.4 * height)          # free space kept on each side of a text
    widths: dict[str, float] = {}

    def width(text: str) -> float:
        if text not in widths:
            probe = doc.ModelSpace.AddText(text, com_backend._apoint(0, 0), height)
            probe.StyleName = first.TextStyle
            lo, hi = probe.GetBoundingBox()
            widths[text] = hi[0] - lo[0]
            probe.Delete()
        return widths[text]

    def axis_of(items) -> int:
        f, t = items[0][1]["from"], items[0][1]["to"]
        return 0 if abs(t[0] - f[0]) >= abs(t[1] - f[1]) else 1

    bounds = {row: {p for _, d in items for p in (d["from"][axis_of(items)], d["to"][axis_of(items)])}
              for row, items in rows.items()}
    moved = 0
    for row, items in rows.items():
        axis = axis_of(items)
        segments = []
        for dim, d in items:
            a, b = sorted((d["from"][axis], d["to"][axis]))
            segments.append((a, b, width(_dim_text(dim)) + 2 * pad))
            dim.TextInside = True
        layout = chain_layout.layout_row(segments, bounds[row], bounds[row] | chain_layout.farther_lines(bounds, row))
        for (dim, _), (c, off_centre) in zip(items, layout):
            if off_centre:
                p = list(dim.TextPosition)
                p[axis] = c
                dim.TextMovement = MOVE_TEXT_NO_LEADER
                dim.TextPosition = com_backend._apoint(p[0], p[1])
                moved += 1
    return moved


async def _draw_chains_scaled(backend, *args, **kwargs):
    res = await _draw_chains_orig(backend, *args, **kwargs)
    factor = float(res.get("dimscale") or 1.0)
    if getattr(backend, "name", "") == "com" and res.get("handles"):
        def apply() -> int:
            doc = com_backend._acad_doc()
            for h in res["handles"]:
                dim = doc.HandleToObject(h)
                if abs(float(dim.ScaleFactor) - factor) > 1e-9:
                    dim.ScaleFactor = factor
                # architectural ticks: arrows do not fit a 120 mm segment and flip onto the neighbours
                dim.Arrowhead1Type = dim.Arrowhead2Type = ARCH_TICK
            return _untangle_chain_texts(doc, res, factor)
        res["texts_moved_beside"] = await backend._run(apply)
    return res


arch_draw.draw_chains = _draw_chains_scaled

# Tables (entity_create_table and the arch schedules) kept the table style's top alignment: the text sat on
# the top rule of each 7-mm row. Centre every cell, title and header rows included.
_create_table_orig = com_backend.ComBackend.entity_create_table
ROW_TYPES, MIDDLE_CENTER = (1, 2, 4), 5   # AcRowType data / title / header (one per call: a combined mask
#                                           is refused, "Неверный ввод"), AcCellAlignment acMiddleCenter


def centre_table(table) -> None:
    for row_type in ROW_TYPES:
        table.SetAlignment(row_type, MIDDLE_CENTER)


async def _create_table_centred(self, *args, **kwargs):
    info = await _create_table_orig(self, *args, **kwargs)
    await self._run(lambda: centre_table(com_backend._acad_doc().HandleToObject(info.handle)))
    return info


com_backend.ComBackend.entity_create_table = _create_table_centred

GUARD_TOOLS = {"acad_health", "acad_dialogs", "acad_dialog_press", "acad_restart", "acad_events"}
# errors that mean "AutoCAD did not answer" rather than "the tool refused / bad input"
ACAD_TROUBLE = re.compile(
    r"did not respond|RPC|rejected|отклон|недоступ|not running|-21474(23174|23170|18111|21246|21005|17842)|"
    r"not accepting COM|modal dialog", re.I)
STALE_PROXY = re.compile(r"-2147023174|RPC server is unavailable|Сервер RPC недоступен", re.I)


async def _reset_backend(fctx) -> bool:
    backend = fctx.lifespan_context.get("backend") if fctx is not None else None
    if backend is None or getattr(backend, "name", "") != "com":
        return False
    with contextlib.suppress(Exception):
        await backend.disconnect()
    com_backend._COM_STATE.pop("app", None)
    await backend.connect()
    return True


def _note(result, text: str):
    with contextlib.suppress(Exception):
        result.content.append(TextContent(type="text", text=text))
    return result


class GuardMiddleware(Middleware):
    def __init__(self) -> None:
        self._pid: int | None = None

    async def on_call_tool(self, context: MiddlewareContext, call_next):
        if context.message.name in GUARD_TOOLS:
            return await call_next(context)
        fctx = context.fastmcp_context
        try:
            pre = await asyncio.to_thread(guard.ensure_ready)
        except guard.GuardError as exc:
            raise ToolError(f"[acad-guard] {exc}") from exc
        pid = (await asyncio.to_thread(guard.quick_state)).get("pid")
        if pre.get("restarted") or (pid and self._pid and pid != self._pid):
            await _reset_backend(fctx)
        self._pid = pid or self._pid
        watcher = guard.DialogWatcher()  # answers known windows that pop up while the tool runs
        try:
            with watcher:
                result = await call_next(context)
        except Exception as exc:
            msg = str(exc)
            qs = await asyncio.to_thread(guard.quick_state)
            trouble = (ACAD_TROUBLE.search(msg) or not qs.get("running") or qs.get("blocked") or qs.get("hung"))
            if not trouble:
                raise
            if STALE_PROXY.search(msg) and qs.get("main_window") and not qs.get("blocked"):
                # the backend still held a COM proxy of a previous AutoCAD process: reconnect, retry once
                await _reset_backend(fctx)
                result = await call_next(context)
                return _note(result, "[acad-guard] reconnected to the current AutoCAD and retried the call")
            diag = await asyncio.to_thread(guard.diagnose_failure)
            if watcher.handled:
                diag["answered_during_call"] = [h.get("rule") for h in watcher.handled]
            if diag.get("cause") == "crash":
                await _reset_backend(fctx)
            raise ToolError(f"{msg}\n[acad-guard] {json.dumps(diag, ensure_ascii=False, default=str)[:2000]}") from exc
        if pre.get("actions"):
            _note(result, "[acad-guard] before this call: "
                  + json.dumps(pre["actions"], ensure_ascii=False, default=str)[:1500])
        if watcher.handled:
            _note(result, "[acad-guard] answered during this call: "
                  + json.dumps([{"rule": h.get("rule"), "actions": h.get("actions")} for h in watcher.handled],
                               ensure_ascii=False)[:1000])
        return result


acad_server.mcp.add_middleware(GuardMiddleware())  # added last = innermost, next to the tool


@acad_server.mcp.tool(annotations={"readOnlyHint": True})
async def acad_health() -> dict:
    """AutoCAD session state: running / blocked by a window / hung, COM answering, open drawings
    (saved or not), active command, start profile (autocad / soficad), SOFiCAD UI loaded, blocking
    windows with the matching guard rule, and the last snapshot of open drawings (used to reopen them
    after a crash)."""
    res = await asyncio.to_thread(guard.health)
    res["wrapper"] = {"autocad_mcp_pro": ACAD_MCP_VERSION, "tested": ACAD_MCP_VERSION in TESTED_VERSIONS}
    return res


@acad_server.mcp.tool(annotations={"readOnlyHint": False})
async def acad_dialogs(handle: bool = False) -> dict:
    """Windows of AutoCAD / owned by AutoCAD / crash reporters / Windows 'open with'.

    handle=False: list them with their buttons and the guard rule that matches (if any).
    handle=True: answer every window covered by a rule (guard_config.json), dismiss single-button
    info boxes, and return the unknown ones with a screenshot path (read it to see the window).
    Unknown windows are never pressed automatically - decide, then use acad_dialog_press."""
    if handle:
        return await asyncio.to_thread(guard.handle_dialogs)
    st = await asyncio.to_thread(guard.list_windows)
    for win in st.get("windows") or []:
        rule = guard.match_rule(win)
        win["rule"] = rule["name"] if rule else None
    return st


@acad_server.mcp.tool(annotations={"readOnlyHint": False})
async def acad_dialog_press(hwnd: int, button: str | None = None, close: bool = False,
                            remember: bool = False, rule_name: str | None = None) -> dict:
    """Press a button (its visible name, e.g. 'Yes', 'Ignore', 'Закрыть') in an AutoCAD-related window,
    or close it (close=True = the title-bar X). hwnd and button names come from acad_dialogs.
    remember=True first stores a rule (exact window title, or process+class if untitled) so the guard
    answers this window the same way from now on."""
    st = await asyncio.to_thread(guard.list_windows)
    if hwnd not in {x["hwnd"] for x in st.get("windows") or []}:
        raise ToolError(f"hwnd {hwnd} is not an AutoCAD-related window (see acad_dialogs)")
    if not button and not close:
        raise ToolError("pass button=... or close=True")
    learned = await asyncio.to_thread(guard.learn_rule, hwnd, button, close, rule_name) if remember else None
    if button:
        res = await asyncio.to_thread(guard.press, hwnd, button)
        if res.get("still_open") and close:
            res = await asyncio.to_thread(guard.close_window, hwnd)
    else:
        res = await asyncio.to_thread(guard.close_window, hwnd)
    guard.log_event("dialog_manual", hwnd=hwnd, button=button, close=close, result=res, learned=learned)
    return {"result": res, "learned_rule": learned}


@acad_server.mcp.tool(annotations={"readOnlyHint": False, "destructiveHint": True})
async def acad_restart(ctx: Context, reopen: bool = True, force: bool = False) -> dict:
    """Restart AutoCAD by the start profile (plain, or through SOFiCAD's launcher and wait until SOFiCAD
    is loaded), reopen the drawings that were open, and reconnect this server. Refuses if a drawing has unsaved changes unless
    force=True (they are discarded); a hung AutoCAD is killed only with force=True.
    Takes 1-3 minutes."""
    try:
        res = await asyncio.to_thread(guard.restart, reopen, force)
    except guard.GuardError as exc:
        raise ToolError(f"[acad-guard] {exc}") from exc
    await _reset_backend(ctx)
    return res


@acad_server.mcp.tool(annotations={"readOnlyHint": False})
async def acad_clear_input() -> dict:
    """Cancel what AutoCAD's command line holds when calls keep failing with '<unknown>.Count'-style
    errors (or acad_health shows com_ok false) and no dialog is open: text typed into the command line
    but never entered (a stray key), or a command waiting for a point. Sends Esc twice as real
    keystrokes - AutoCAD comes to the front for a moment. Returns whether COM answers again."""
    return await asyncio.to_thread(guard.clear_command_line)


@acad_server.mcp.tool(annotations={"readOnlyHint": True})
async def acad_events(limit: int = 30) -> list:
    """Recent guard events: dialogs answered (rule, title, button), unknown windows (with screenshot
    paths), starts, restarts, crash recoveries, learned rules."""
    return await asyncio.to_thread(guard.events, limit)


if __name__ == "__main__":
    acad_server.main()
