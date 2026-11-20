# Copy the Claude Code skills between this repo (skills/) and a Claude Code project (<project>\.claude\skills).
#   Scripts\python.exe tools\sync_skills.py export <project>    project -> repo  (before a commit)
#   Scripts\python.exe tools\sync_skills.py install <project>   repo -> project  (after a clone)
# Claude Code loads the skills from the project; the copy in the repo is what gets versioned.
import os
import shutil
import sys

SKILLS = ("autocad-ops",)  # soficad-mcp syncs its own skills (soficad-ops, soficad-drawing)
REPO_SKILLS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "skills")


def main() -> None:
    if len(sys.argv) != 3 or sys.argv[1] not in ("export", "install"):
        sys.exit(__doc__ or "usage: sync_skills.py export|install <project folder>")
    project_skills = os.path.join(os.path.abspath(sys.argv[2]), ".claude", "skills")
    src, dst = (project_skills, REPO_SKILLS) if sys.argv[1] == "export" else (REPO_SKILLS, project_skills)
    for name in SKILLS:
        s, d = os.path.join(src, name), os.path.join(dst, name)
        if not os.path.isfile(os.path.join(s, "SKILL.md")):
            sys.exit(f"no skill {name!r} in {src}")
        if os.path.isdir(d):
            shutil.rmtree(d)  # a mirror: files deleted at the source go here too
        shutil.copytree(s, d, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
        print(f"{s} -> {d}")


if __name__ == "__main__":
    main()
