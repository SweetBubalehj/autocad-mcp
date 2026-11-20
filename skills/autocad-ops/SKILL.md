---
name: autocad-ops
description: Operating a live AutoCAD (full AutoCAD on Windows, not LT) through the MCP server autocad (autocad-mcp-pro + a session guard) - starting it, dialogs, crashes and restarts, transient COM errors, a stuck command line, interactive commands, LISP through the command line, big drawings, plotting to custom paper sizes and plot styles. Use before any work with the autocad server and whenever an AutoCAD call hangs, returns a COM/RPC error, or AutoCAD crashed. With SOFiCAD, skill soficad-ops (repo soficad-mcp) adds SOFiCAD's API, units, dialogs and schedule.
---

# AutoCAD: staying unstuck

This skill covers plain AutoCAD. With SOFiCAD, skill **soficad-ops** (repo soficad-mcp) adds SOFiCAD's API, units, dialogs and schedule, and **soficad-drawing** builds reinforcement drawings.

## Setup
- **Server `autocad`:** autocad-mcp-pro 1.6.0, started through the wrapper `C:\autocad-mcp\acad_guard\autocad_guarded.py`.
  - Before **every** call a guard (`guard.py`):
    - answers known dialogs;
    - starts AutoCAD if needed;
    - recovers after a crash and reopens the drawings;
    - reconnects COM when the AutoCAD process changed.
  - Extra tools: `acad_health`, `acad_dialogs`, `acad_dialog_press`, `acad_clear_input`, `acad_restart`, `acad_events`.
- Other servers can share the guard: soficad-mcp's `soficad` server imports it from `C:\autocad-mcp\acad_guard`.
- **Guard files:**
  - dialog rules and the start profile: `C:\autocad-mcp\acad_guard\guard_config.json` (re-read on every call);
  - log `events.jsonl`; screenshots of unknown windows in `shots\`.
- **After editing server code:** `/mcp` → Reconnect that server (ask the user).

## Starting
1. `acad_health`: running, blocked, hung, `com_ok`, open drawings (saved or not), `start_profile`, `soficad_ui_loaded`.
2. If AutoCAD is not running, do nothing special: the guard starts it on the first call (or use `acad_restart`).
- **Start profile** (`start.profile` in `guard_config.json`; the environment variable `ACAD_GUARD_PROFILE` overrides it):
  - `auto` (default): through SOFiCAD's launcher if SOFiCAD is installed, else plain AutoCAD;
  - `autocad`: `acad.exe /nologo`. `acad_exe: "auto"` takes the newest `Program Files\Autodesk\AutoCAD 20xx\acad.exe`; give a path to pick another version;
  - `soficad`: `soficad_launcher.exe --application soficad`; the guard then waits until SOFiCAD's UI has loaded (see soficad-ops).
- **Never start AutoCAD through COM** (`Dispatch("AutoCAD.Application")`, `system_launch` with no AutoCAD running). The wrapper disables autocad-mcp-pro's own COM start: the guard starts AutoCAD by the profile, and a COM-started AutoCAD has no SOFiCAD UI or API.

## Dialogs
- **Known windows are answered by the guard** (rules in `guard_config.json`):
  - Autodesk trial / licensing banner → close;
  - Drawing Recovery → Close;
  - unsigned `.lsp` → Do Not Load;
  - missing references → Ignore;
  - proxy → Show proxy graphics;
  - Windows "open with" → close;
  - crash report → close;
  - SOFiCAD's windows (see soficad-ops).
- **How fast the guard answers:**
  - `DialogWatcher` checks every 0.2 s while a call runs (`quick_state`, WinAPI, 0.2 ms). At 2.5 s, five modal questions during one build cost 17 s.
  - `windows.ps1` runs as one persistent `-Mode serve` process per Python process (`guard._PsWorker`, one JSON line per request). A window list takes 0.12 s instead of 0.6 s, because PowerShell's start and the `Add-Type` compile happen once.
  - Press and close poll for the window to go away instead of sleeping 700 ms.
  - If the worker fails, `_ps` runs the old one-shot PowerShell for that request. The worker ends when its parent process closes the pipe.
  - Result: a question costs about 0.5 s.
- **COM initialisation in Python:**
  - `pythoncom.CoInitialize()` on a thread that already has COM (every thread that imported pythoncom is MAINSTA) adds no reference. The `CoUninitialize()` after it then ends that thread's COM.
  - Seen: the guard's 30-s snapshot (`com_probe`) broke a script's main thread, and its next `GetActiveObject` failed with CO_E_NOTINITIALIZED ("AutoCAD is not running").
  - `guard._com()` therefore leaves a thread that has COM alone (`CoGetApartmentType`). Init/uninit pairs belong only in threads you start yourself.
- **Never send input while a dialog is open.** AutoCAD accepts `SendCommand` in a modal state but loses the input (seen: 16 objects vanished while a question stayed open for 10 min). Check `guard.quick_state()["blocked"]` before every command and answer or report the window first.
- A window with a single OK/Close button is dismissed as an info box.
- **Unknown windows are left alone.** The call fails with "[acad-guard] … blocked by a window", giving the title, texts, buttons and a screenshot path.
  1. Look at the screenshot (Read) or call `acad_dialogs`.
  2. Make safe choices yourself (Ignore / Cancel / Do not load / Close). Ask the user when the choice saves, overwrites, deletes or sends data.
  3. For a recurring window, use `acad_dialog_press(hwnd, button, remember=True)` to store a rule.

## Localized AutoCAD (checked in Russian AutoCAD 2027, 2026-11-20)
- **Language:** `app.LocaleId` (1049 = Russian, 1033 = English); the Start tab is "Начало".
- **Commands:** send English names with `_` (`_LINE`, `_.UNDO`, `_-LAYOUT`); without it a localized AutoCAD answers "Неизвестная команда". autocad-mcp-pro does this already; do the same in `system_run_command` and LISP `(command "_.…")`.
- **Linetypes:** the localized `acadiso.lin` / `acad.lin` name the standard ones in the local language (Russian: CENTER = "осевая", HIDDEN = "невидимая", DASHED = "штриховая", PHANTOM = "фантом"; only `ACAD_ISO*` keep their names). The wrapper retries a failed load from `acad_guard\lin\acadiso_en.lin` / `acad_en.lin` (the same patterns under the English names), so `linetype="CENTER"` works in any language.
- **Dialogs:** titles and buttons are localized. The guard rules list both languages ("Безопасность — исполняемый файл без подписи" → "Не загружать", "Восстановление чертежа", "Ссылки — файлы не найдены", "Информация о прокси-объектах"). A rule's buttons are tried in order and a missing name is skipped, so add the other language's button after the English one. The Russian texts are in `<AutoCAD>\ru-RU\AcTaskDialogs.resources.dll` and `acadres.dll` (UTF-16 strings).
  - The titles use an em dash (`—`), not a hyphen.
  - Not known in Russian yet: "AutoCAD Alert … cannot continue" and the crash-report titles (a crash cannot be brought up on purpose). Look at the screenshot when a crash comes, and add them.
- **AutoCAD remembers "Не загружать" / "Do Not Load" for a file for the rest of the session:** a second `(load)` of it is cancelled silently, with no window.
- **Opening through COM** (`Documents.Open`) showed no proxy notice and no references dialog. A missing xref opens the non-modal "Внешние ссылки" palette, which blocks nothing.
- **Command-line text** is read as UTF-8 (`cmdline.ps1` and `acad_keys.ps1` set `[Console]::OutputEncoding`; PowerShell otherwise writes Cyrillic in the OEM code page 866).
- **Cyrillic text, layer names and file paths** work through COM (`Стены`, `Проверка чертежа.dwg`, PDF export).

## Stuck command line (no dialog, but COM refuses)
- **Symptom:** every call fails with `<unknown>.Count` / `<unknown>.FullName` for more than 15 s, and `acad_health` shows `com_ok: false` while `blocked: false`.
- **Cause:** the command line holds text typed but never entered (a stray key: "Ъ" = `]` on a Russian layout; once a "„" after layout switches), or a command waiting for a point, or LISP waiting for a closing bracket `(_>`.
- **Fix:** `acad_clear_input`. It sends Esc, Esc as real keystrokes, so AutoCAD comes to the front for a moment; mention it to the user. Then retry.
- **"AutoCAD is not in the foreground - nothing typed"** (from `acad_clear_input` / `acad_keys.ps1`): some other window holds the focus.
  - Seen three times: a Windows "open with" (`OpenWith.exe`) left by an export that wrote a .docx with no Word installed; one came over an hour after it.
  - **The guard closes these:** `guard.expect_open_with()` before such an export, `close_open_with()` in `ensure_ready` and in `DialogWatcher` once a second. It closes only OpenWith processes that started within 2 min of the last `expect_open_with`, never an "open with" the user brought up at another time. The time is kept in `acad_guard/open_with_expected.txt`, so other processes close them too.
  - Windows 11's "open with" dialog (class `Open With`) is **not listed by `EnumWindows`**, although it is visible and in front; `FindWindowEx(None, h, "Open With", None)` finds it.
  - By hand: `Get-Process OpenWith | Stop-Process` (it is only the app chooser).
  - `acad_keys.ps1` accepts any window of acad.exe in front. While a command waits for input, the focus sits in AutoCAD's floating command line (an `HwndWrapper` window), not in the main frame.
- **See what is pending:** `C:\autocad-mcp\tools\cmdline.ps1` (or `guard.command_line()`).
- **Window messages** (`PostMessage WM_CHAR`) never reach the command line, but they can arrive later once the window gets focus. Do not use them.

## Crashes
- **Recovery is automatic:** the guard closes the crash report, starts AutoCAD by the start profile, waits until it answers (and, for SOFiCAD, until its UI is loaded) and reopens the last snapshot's drawings (about 1 min). `acad_events` shows what happened.
- **Unsaved changes** survive only in autosave (`%TEMP%\*.sv$`) and the Drawing Recovery Manager. Tell the user.
- **The crash dialog chain is handled:** Error Aborting → OK; Error Report → close; "AutoCAD cannot continue… save?" → Yes (`*_recover.dwg`; the original is untouched).
- **A Python process (not AutoCAD) crashing with exit code 5** (access violation): `CoUninitialize` between two COM sessions in the same thread. Initialise COM once per thread and never uninitialise it.
  - `faulthandler` also prints harmless `0x80010001` (call rejected) exceptions that COM handles itself.
- **Transient COM errors** (`<unknown>.FullName`, `AutoCAD.Application.ActiveDocument`, `Property 'Item.X' can not be set.`, "Вызов был отклонен" / RPC_E_CALL_REJECTED) mean AutoCAD briefly rejected a call while busy after a big operation, a save or a layout switch; win32com reports some of them as an AttributeError.
  - The `autocad` server's COM thread has an OLE message filter (`acad_guard\message_filter.py`, since 2026-11-20): COM itself retries a rejected call every 0.1 s for up to 20 s. Before it, `arch_plan_from_spec` failed halfway with "Вызов был отклонен" and left a partial plan (its rollback did not run).
  - Measured: 20 Documents.Add/read/Close rounds gave 5 errors without the filter, 0 with it (14 retries).
  - Own scripts: `message_filter.register()` once on the thread after `CoInitialize`. Errors that persist mean a stuck command line (see above).

## LISP through the command line
- **At most 240 chars per INPUT LINE.** A longer line is truncated and AutoCAD hangs at `(_>`.
  - One expression may span many input lines in one `SendCommand` (`"(progn\n…\n…\n)\n"`): 2209 chars over 11 lines worked.
  - A `(defun …)` is one line, so split long helpers into several defuns.
- **LISP symbols ignore case:** `mcp-gh` and `mcp-gH` are the same function, and the later defun silently replaces the earlier one (seen: "too few arguments").
- **`(load "temp.lsp")` triggers the SECURELOAD dialog.** Send commands instead, or send data: `(open f "r" "utf8")` + `(read-line)` + `(read)` reads lists without evaluating them. That handled a 41,000-char line, and entmake took a 2000-vertex polyline from it.
- **Files AutoLISP writes:** open them with `"utf8"` and read them as UTF-8; the ANSI code page here is 1251, which has no umlauts.
- **COM per entity is slow:** every property is a cross-process call (20–50 ms per item). Many objects: write entmake data to a file and let one in-process LISP loop create them (about 1 ms per object).
- **entmake versus COM Add\*:**
  - The same: omitted colour, linetype and lineweight follow CECOLOR / CELTYPE / CELWEIGHT.
  - Different: a TEXT does not take its style's width factor (give group 41).
  - A text's annotative flag comes from group 7 (the style), whichever style is current.
  - For a hatch, DXF wants the pattern's line definitions, so make hatches in-process with `vla-AddHatch` on an entmade boundary.
- **`(command "_.ERASE" ...)` inside a `(progn …)` silently erased nothing.** Use `entdel`.

## Big drawings
- **20+ MB, PDF underlays:** `analysis_layer_stats` / `entity_list` time out (one COM call per entity). Read them with one LISP pass instead (soficad-mcp has `tools\dwg_survey.py`, `dwg_window_dump.py` and `dwg_entget.py`, see soficad-ops).
- Open a COPY of such a drawing and close it without saving.
- **`ssget "C"` only finds what the view shows:** zoom to a window before selecting in it.
- **Geometry hidden in blocks:** plans pasted as big anonymous blocks (A$C…) do not show up in a window selection. The insert's DXF data plus the block definition transformed by the insert's position and rotation gives world coordinates.
- **Save fails with "Error saving the document" / "One or more objects in this drawing cannot be saved to the specified format":** broken `AcDbPdfReference` objects (PDF underlays). `_.AUDIT _Y` fixes it, then the save works (seen twice after big builds).

## Experiments and undo
- **Save the drawing before experimenting;** afterwards close it without saving and reopen it (`Close(False)` + `Documents.Open`). Undo after other commands may undo something else.
- **Rebuilding by script is quicker than repairing.**

## Interactive commands
- **A command waiting for a point from LISP blocks COM** (`com_ok: false`, `SendCommand` does not return). Answer it with real keystrokes:
  - `C:\autocad-mcp\tools\acad_keys.ps1 -Keys "..."`: brings AutoCAD to the front and checks the focus;
  - or cancel with `acad_clear_input`.
- **To explore a command:** `C:\autocad-mcp\tools\cmd_probe.py "COMMAND\n"`, with `cmdline.ps1` / `guard.list_windows()` to see what it asks. Save the drawing first.

## Screenshots
- `view_zoom_and_screenshot` with a layout active zooms the layout, so switch to model first.
- The real window, even when covered: `C:\autocad-mcp\tools\shot_acad.ps1 -Out x.png`.

## Layouts, long sheets, plot styles
- **Paper larger than DWG To PDF's largest size** (4A0, 2378 × 1682): `tools\pc3_custom_size.py "DWG To PDF.pc3" Langformat 3950x914`.
  - It writes a NEW `DWG To PDF - Langformat.pc3` plus a `.pmp` with that user-defined size in the profile's Plotters folder. AutoCAD lists it at once ("UserDefinedMetric (3950.00 x 914.00MM)").
  - A new plotter file is a profile change: ask the user first.
  - File format of .pc3 / .pmp / .ctb: a 48-byte header, then adler32 of the compressed data, the uncompressed length, the compressed length, and a zlib text.
- **monochrome.ctb prints white text masks black.** Blocks with a solid hatch in colour 255 behind their text (e.g. elevation markers) lose the text.
  - `tools\ctb_variant.py monochrome.ctb monochrome_grau.ctb 8 9 250 251 252 253 254 255` makes a copy in which those colours plot in their own colour: 255 white, the greys grey.
- **Layouts and viewports over COM:**
  - `vla-CopyObjects` returns the copies NOT in the input order: copy objects you need by handle one at a time.
  - `Coordinates` of a polyline cannot be set through late-bound COM ("can not be set"); use `vla-put-Coordinates` in LISP.
  - `MSpace = True` fails with "There is no active modelspace viewport" when the new viewport is off screen: zoom the paper first.
  - `SetVariable("LAYOUTCREATEVIEWPORT", 0)` needs a short (`VARIANT(VT_I2, 0)`).

## Architectural plans (`arch_*`; a 100 m² house checked 2026-11-20 in Russian AutoCAD)
- **`arch_plan_from_spec`** draws walls (mm, axes, junctions mitred), doors, windows, room labels with MEASURED areas, the exterior chains and the door / window / room schedules in one call (about 20 s for 12 walls, 18 openings, 9 rooms). Separate exterior walls per side keep opening `offset`s simple; give them in CCW order so `swing: "in"` opens inwards.
- **Language:** `lang` is `en` or `tr` only. Room names are free text (Cyrillic works); the area label is "5.73 m²" and the schedules have English headers. `text_find_replace` reaches TEXT/MTEXT (labels: " m²" → " м²", "." → ","), not TABLE cells: translate those over COM (`table.SetText(row, col, …)`, `DeleteColumns` for the internal WALL column, `SetColumnWidth`, `RecomputeTableBlock`).
- **Font:** the Standard style uses `txt.shx`, which has no "²" (shown as "?"). ISOCPEUR (Windows Fonts) has Cyrillic, "²" and "№": `doc.TextStyles.Item("Standard").SetFont("ISOCPEUR", False, False, 204, 34)`. `system_run_lisp` refuses every `vla-*` expression (denylist), so do it over COM.
- **Chain dimensions** (wrapper, since 2026-11-20; `arch_dimension_chains` and `arch_plan_from_spec` both get it):
  - The chains set DIMSCALE while drawing, but COM-added dimensions keep ScaleFactor 1 (2.5 mm text in the model: no numbers on a 1:100 sheet). The wrapper sets each one's ScaleFactor to the scale.
  - Architectural ticks instead of arrows (arrows do not fit a 120 mm segment and flip onto the neighbours).
  - The texts are laid out per row by `acad_guard\chain_layout.py` (pure, tested offline): a text that fits stays inside its segment (`TextInside`; AutoCAD's best fit moved some out on its own), a 120 / 250 / 400 wall's text goes beside its segment and a neighbour slides within its own segment to make room. Texts never overlap and never cross their own row's extension lines; the farther rows' extension lines (walls, overall run through the openings row) are avoided where there is room. The result reports `texts_moved_beside`.
  - Width = the text AutoCAD shows (ISO-25 has precision 2 with trailing zeros suppressed: "900", not "900.00"), measured with a probe TEXT in the dimension's style.
  - Check a row objectively: take each dimension's `TextPosition` and the measured width of its text, cluster the rows by the text line, and look for overlapping intervals.
- **Tables** (`entity_create_table`, the arch schedules): the wrapper centres every cell (`SetAlignment(row_type, acMiddleCenter)` once per row type 1 / 2 / 4 - a combined mask is refused with "Неверный ввод"); before, the text sat on the top rule of each row.
- **Sizes at 1:100:** label names are 350 mm high and start at `at`, left-aligned: in rooms narrower than about 2.2 m scale the three label texts by 0.8 about the name's insertion and move them (`entity_batch_modify`). A schedule row is 0.7 m: 9 rooms ≈ 8.4 m of table, so place tables from their real `bbox`.
- **Sheet:** a new layout gets AutoCAD's default viewport when it is first shown or plotted (`viewport_list` showed none before): list again after the first plot and delete it, or the plan appears twice. DWG To PDF's "ISO A2" has margins (17 / 5 mm), so a 594 × 420 frame at the paper origin is clipped: use `ISO_full_bleed_A2_(594.00_x_420.00_MM)` (`layout.CanonicalMediaName` over COM; `CenterPlot` is refused for layout plots, `PlotOrigin` 0,0). Put the viewport on a layer with `Plottable = False`, or its frame prints around the drawing.
- **Check the PDF, not only the model screenshot:** render it (`pypdfium2` is in the venv) and look; the doubled viewport and the missing dimension texts only showed there.

## Files
- `C:\autocad-mcp\acad_guard\`: `guard.py`, `windows.ps1`, `guard_config.json`, `autocad_guarded.py`, `message_filter.py`, `chain_layout.py`, `test_start_profile.py`, `lin\`, `events.jsonl`, `shots\`.
- `C:\autocad-mcp\tools\`: `shot_acad.ps1`, `cmdline.ps1`, `acad_keys.ps1`, `cmd_probe.py`, `pc3_custom_size.py`, `ctb_variant.py`, `sync_skills.py`.
- The session scratchpad is temporary: move anything useful into `tools` or a skill.
