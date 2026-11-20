# Explore an interactive AutoCAD command: feed raw input to the command line (each argument = one
# SendCommand string, "\n" = Enter), then show the command-line tail and any windows that popped up.
# Careful: a command that waits for a point from LISP blocks COM until it gets input - answer it with
# acad_keys.ps1 (real keystrokes, AutoCAD must be in front) or Esc. Save the drawing before experiments.
import json, os, subprocess, sys, threading, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "acad_guard"))
import pythoncom
import win32com.client as w
import guard

HERE = os.path.dirname(os.path.abspath(__file__))
inputs = [a.replace("\\n", "\n") for a in sys.argv[1:]]
state = {"sent": 0}


def send():
    pythoncom.CoInitialize()
    try:
        d = w.GetActiveObject("AutoCAD.Application").ActiveDocument
        for s in inputs:
            d.SendCommand(s)
            state["sent"] += 1
            time.sleep(0.8)
    except Exception as ex:  # noqa: BLE001
        state["error"] = str(ex)[:300]
    finally:
        pythoncom.CoUninitialize()


t = threading.Thread(target=send, daemon=True)
t.start()
t.join(20)
time.sleep(1.0)
print("sent", state, "| thread still blocked:", t.is_alive())
for win in guard.list_windows().get("windows") or []:
    print("WINDOW:", json.dumps({k: win.get(k) for k in ("hwnd", "title", "class")}, ensure_ascii=False),
          [(e["type"], e["name"]) for e in win.get("elements", [])][:40])
out = subprocess.run(["powershell", "-NoProfile", "-File", os.path.join(HERE, "cmdline.ps1"), "-Tail", "12"],
                     capture_output=True, text=True, encoding="mbcs", errors="replace").stdout
print("CMDLINE:\n" + out)
