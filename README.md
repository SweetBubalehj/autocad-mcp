# AutoCAD MCP with a session guard

An MCP server that lets Claude Code (or any MCP client) drive a **running AutoCAD** on Windows, plus a Claude
Code skill for operating it. It is [autocad-mcp-pro](https://github.com/U-C4N/Autocad-MCP) (MIT, about 250
tools for drawing, editing, layers, dimensions, layouts, plotting, used unmodified as a dependency) wrapped in
a **session guard** that keeps a live AutoCAD usable without a person at the keyboard:

- before every tool call it answers known dialogs, starts AutoCAD if it is not running, recovers from a crash
  (closes the crash reports, restarts AutoCAD, reopens the drawings) and reconnects COM to a new AutoCAD
  process, so no MCP reconnect is needed after a restart;
- while a call runs it answers known dialogs that pop up (every 0.2 s), and leaves unknown ones alone: the
  call fails with the window's title, texts, buttons and a screenshot, and `acad_dialog_press(...,
  remember=True)` turns the answer into a rule;
- after a failed call it tells a dialog from a crash from a command line that waits for input, and
  `acad_clear_input` cancels the latter;
- it works with **localized AutoCAD**: the dialog rules carry English and Russian titles and buttons, and the
  standard linetypes (CENTER, HIDDEN, ...) load under their English names even where the local `acadiso.lin`
  translates them (checked in Russian AutoCAD 2027);
- an OLE message filter on autocad-mcp-pro's COM thread makes COM retry the calls a busy AutoCAD rejects
  ("Call was rejected by callee"), which otherwise failed long operations such as a whole plan halfway;
- architectural chain dimensions get the plot scale they are drawn for (without it their text was too small
  to see on the sheet), architectural ticks, and a text layout in which no two numbers overlap (a short wall
  segment's number goes beside it, a neighbour slides aside to make room);
- table cells are centred.

Extra tools: `acad_health`, `acad_dialogs`, `acad_dialog_press`, `acad_clear_input`, `acad_restart`,
`acad_events`.

| Folder | What |
|---|---|
| `acad_guard/` | `autocad_guarded.py` (the `autocad` server: autocad-mcp-pro + the guard as middleware, plus small patches), `guard.py` (windows, health, start, crash recovery), `windows.ps1` (UI Automation), `message_filter.py` (COM retries for a busy AutoCAD), `chain_layout.py` (where chain dimension texts go), `guard_config.json` (start profile and dialog rules), `lin/` (English linetype files for localized AutoCAD), `test_start_profile.py` (offline test). |
| `skills/autocad-ops/` | Claude Code skill: starting AutoCAD, dialogs, crashes, a stuck command line, LISP through the command line, big drawings, plot styles, localized AutoCAD. |
| `tools/` | `cmdline.ps1` (command-line text), `acad_keys.ps1` (real keystrokes), `shot_acad.ps1` (window screenshot), `cmd_probe.py` (explore an interactive command), `pc3_custom_size.py` (paper sizes beyond DWG To PDF's), `ctb_variant.py` (plot-style variants), `sync_skills.py`. |

**SOFiCAD:** the separate repo [soficad-mcp](https://github.com/SweetBubalehj/soficad-mcp) builds SOFiCAD reinforcement drawings on top of this one (its
own `soficad` server uses this guard).

## Requirements

- Windows with full AutoCAD (COM and AutoLISP; AutoCAD LT has neither). Tested with AutoCAD 2027, English
  and Russian.
- Python 3.15 (64-bit; tested) and Claude Code.

## Install

The skill and `pip.ini` expect the repo in `C:\autocad-mcp`; the server finds its own files relative to
itself. The repo root doubles as the venv:

```powershell
git clone https://github.com/SweetBubalehj/autocad-mcp.git C:\autocad-mcp
cd C:\autocad-mcp
py -3.15 -m venv .
Scripts\pip install -r requirements.lock.txt      # pip.ini pins autocad-mcp-pro (constraints.txt)
Scripts\python.exe acad_guard\test_start_profile.py
```

Register the server in your Claude Code project: copy `mcp.example.json` to `<project>\.mcp.json`, or

```powershell
claude mcp add autocad -e AUTOCAD_MCP_BACKEND=com -- C:\autocad-mcp\Scripts\python.exe C:\autocad-mcp\acad_guard\autocad_guarded.py
```

Then install the skill into that project:

```powershell
Scripts\python.exe tools\sync_skills.py install <project folder>
```

In Claude Code, `acad_health` shows AutoCAD's state and the start profile.

## How AutoCAD is started

`start.profile` in `acad_guard/guard_config.json` (the environment variable `ACAD_GUARD_PROFILE` overrides it):

- `auto` (default): through SOFiCAD's launcher if SOFiCAD is installed, else plain AutoCAD;
- `autocad`: `acad.exe /nologo`. `acad_exe: "auto"` picks the newest `Program Files\Autodesk\AutoCAD 20xx\acad.exe`;
  give a path to pick another version, and `acad_args` for a profile (`/p`) or template (`/t`);
- `soficad`: SOFiCAD's launcher; the guard waits until SOFiCAD's UI has loaded.

AutoCAD is never started through COM (autocad-mcp-pro's own COM start is disabled): the guard starts it by the
profile, so it comes up the way a person starts it.

## Dialog rules

`dialog_rules` in `guard_config.json` match a window by title, process, class or text (regular expressions)
and press the listed buttons in order; a button that is not there is skipped, so one rule can list the English
and the Russian name. Rules learned with `acad_dialog_press(remember=True)` are added at the top. The file is
re-read on every call. Windows of other languages are not known yet: when the guard reports an unknown window,
add its title and button (the screenshot path is in the error and in `acad_events`).

## Working on it

- Server code changes need `/mcp` → Reconnect in Claude Code.
- The skill is loaded from the project's `.claude\skills`. After editing it there, copy it back before
  committing: `Scripts\python.exe tools\sync_skills.py export <project folder>`.
- GitHub Actions runs `acad_guard/test_start_profile.py` on every push (Windows, Python 3.15, pywin32).
- Runtime state (`acad_guard/events.jsonl`, `session_state.json`, `shots/`) is ignored.

## License

MIT, see [LICENSE](LICENSE). The dependency autocad-mcp-pro is MIT as well. AutoCAD is a product of Autodesk;
it is not part of this repo and needs its own licence.
