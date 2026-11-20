# Offline test of the guard's configuration (no AutoCAD needed; Windows with pywin32, as the guard itself):
# start profile and command, dialog rules in English and Russian, the English fallback linetype files.
# Run: C:\autocad-mcp\Scripts\python.exe C:\autocad-mcp\acad_guard\test_start_profile.py
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import guard  # noqa: E402

failed = 0


def check(name: str, cond: bool, detail: object = "") -> None:
    global failed
    print(f"  {'ok  ' if cond else 'FAIL'} {name}" + (f"  ({detail})" if detail and not cond else ""))
    failed += not cond


os.environ.pop("ACAD_GUARD_PROFILE", None)
tmp = tempfile.mkdtemp()
launcher = os.path.join(tmp, "soficad_launcher.exe")
acad = os.path.join(tmp, "acad.exe")
for p in (launcher, acad):
    open(p, "w").close()
ST = {"profile": "auto", "soficad": {"launcher": launcher, "launcher_args": ["--application", "soficad"],
                                     "acad_exe": acad, "acad_args": ["/p", "SOFiCAD"], "require_soficad_ui": True},
      "autocad": {"acad_exe": acad, "acad_args": ["/nologo"]}}

print("start profile:")
check("auto with SOFiCAD's launcher -> soficad", guard.start_profile(ST)[0] == "soficad")
no_sof = dict(ST, soficad=dict(ST["soficad"], launcher=os.path.join(tmp, "missing.exe")))
check("auto without it -> autocad", guard.start_profile(no_sof)[0] == "autocad")
for env, want in (("autocad", "autocad"), ("SOFICAD", "soficad"), (" auto ", "soficad")):
    os.environ["ACAD_GUARD_PROFILE"] = env
    check(f"ACAD_GUARD_PROFILE={env!r} -> {want}", guard.start_profile(ST)[0] == want)
os.environ["ACAD_GUARD_PROFILE"] = "bogus"
try:
    guard.start_profile(ST)
    check("unknown profile is refused", False)
except guard.GuardError:
    check("unknown profile is refused", True)
os.environ.pop("ACAD_GUARD_PROFILE", None)
legacy = {"launcher": launcher, "acad_exe": acad, "require_soficad_ui": True}
check("config before profiles (flat) -> soficad", guard.start_profile(legacy) == ("soficad", legacy))

print("start command:")
real_config = guard.config
guard.config = lambda: {"start": ST}
check("soficad: the launcher", guard.start_command() == [launcher, "--application", "soficad"])
guard.config = lambda: {"start": dict(ST, profile="autocad")}
check("autocad: acad.exe /nologo", guard.start_command() == [acad, "/nologo"])
auto_exe = dict(ST, profile="autocad", autocad={"acad_exe": "auto", "acad_args": []})
guard.config = lambda: {"start": auto_exe}
found = guard.find_acad_exe()
try:
    cmd = guard.start_command()
    check("acad_exe 'auto' -> the newest installed AutoCAD", found is not None and cmd == [found])
except guard.GuardError as ex:
    check("acad_exe 'auto' without AutoCAD -> a clear error", found is None and "acad_exe" in str(ex), ex)
guard.config = real_config

print("dialog rules:")
cfg = guard.config()
for rule in cfg["dialog_rules"]:
    for key in ("title", "text", "class"):
        if key in rule:
            try:
                re.compile(rule[key])
            except re.error as ex:
                check(f"regex of {rule['name']!r}.{key}", False, ex)
check("every rule regex compiles", True)
WINDOWS = [  # (window as windows.ps1 lists it, rule name expected)
    ({"title": "Security - Unsigned Executable File", "process": "acad"}, "SECURELOAD: unsigned file"),
    ({"title": "Безопасность — исполняемый файл без подписи", "process": "acad"}, "SECURELOAD: unsigned file"),
    ({"title": "Drawing Recovery", "process": "acad"}, "Drawing Recovery notice (previous session crashed)"),
    ({"title": "Восстановление чертежа", "process": "acad"}, "Drawing Recovery notice (previous session crashed)"),
    ({"title": "References - Unresolved Reference Files", "process": "acad"}, "missing references (xref/images)"),
    ({"title": "Ссылки — файлы не найдены", "process": "acad"}, "missing references (xref/images)"),
    ({"title": "Proxy Information", "process": "acad"}, "proxy information"),
    ({"title": "Информация о прокси-объектах", "process": "acad"}, "proxy information"),
    ({"title": "Some other window", "process": "acad"}, None),
]
for win, want in WINDOWS:
    got = (guard.match_rule(win, cfg["dialog_rules"]) or {}).get("name")
    check(f"{win['title']!r} -> {want}", got == want, got)
for rule in cfg["dialog_rules"]:
    if rule["name"] in ("SECURELOAD: unsigned file", "missing references (xref/images)", "proxy information"):
        check(f"{rule['name']!r} has a Cyrillic button", any(re.search("[а-яА-Я]", b) for b in rule["buttons"]))

print("fallback linetypes:")
lin_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lin")
for fname, scale in (("acad_en.lin", 1.0), ("acadiso_en.lin", 25.4)):
    defs = re.findall(r"^\*([A-Z0-9]+),[^\n]*\r?\n(A,[^\r\n]+)", open(os.path.join(lin_dir, fname), encoding="ascii").read(), re.M)
    names = {n for n, _ in defs}
    check(f"{fname}: 24 standard linetypes", len(defs) == 24 and {"CENTER", "HIDDEN", "PHANTOM", "DASHED", "DASHDOT"} <= names)
    center = [float(v) for v in dict(defs)["CENTER"][2:].split(",")]
    check(f"{fname}: CENTER = 1.25, -.25, .25, -.25 inch", all(abs(a - b * scale) < 1e-9 for a, b in zip(center, [1.25, -.25, .25, -.25])))

print("RESULT:", "all checks passed" if not failed else f"{failed} check(s) FAILED")
sys.exit(1 if failed else 0)
