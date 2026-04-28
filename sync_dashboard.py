"""
StartHub Africa Dashboard – Auto-Sync Agent
============================================
Watches TWO files for changes:
  • The Impact Excel file  → data update
  • generate_dashboard.py → design/code update

When either changes it:
  1. Regenerates index.html
  2. Commits & pushes to GitHub Pages

Run this script at Windows startup (see start_sync.bat).
It runs silently in the background until you close the window.

CONFIGURE THE TWO SETTINGS BELOW before first run.
"""

import sys
import time
import subprocess
import pathlib
import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION  ← edit these if needed, then leave forever
# ─────────────────────────────────────────────────────────────────────────────

# Folder containing generate_dashboard.py and the Excel file (auto-detected)
WORK_DIR = pathlib.Path(__file__).parent.resolve()

# Name of your Excel file
EXCEL_FILE = "StartHub Africa Impact Dashboard (2).xlsx"

# How often to check for changes, in seconds
CHECK_INTERVAL = 60

# ─────────────────────────────────────────────────────────────────────────────

EXCEL_PATH = WORK_DIR / EXCEL_FILE
GENERATOR  = WORK_DIR / "generate_dashboard.py"
DASHBOARD  = WORK_DIR / "index.html"
LOG_FILE   = WORK_DIR / "sync_log.txt"

def log(msg: str):
    ts   = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def run(cmd: list, cwd=None):
    result = subprocess.run(cmd, cwd=cwd or WORK_DIR, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())
    return result.stdout.strip()

def git_push(reason: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    run(["git", "add", "index.html"])
    run(["git", "commit", "-m", f"Auto-update dashboard {ts} ({reason})"])
    run(["git", "push"])

def regenerate(reason: str):
    log(f"{reason} – regenerating dashboard…")
    try:
        result = subprocess.run(
            [sys.executable, str(GENERATOR)],
            cwd=WORK_DIR, capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or result.stdout)
        log("Dashboard generated successfully.")
    except Exception as e:
        log(f"ERROR generating dashboard: {e}")
        return False
    return True

def push(reason: str):
    log("Pushing to GitHub Pages…")
    try:
        git_push(reason)
        log("✅ Dashboard published – team link is live.")
    except Exception as e:
        log(f"ERROR pushing to GitHub: {e}")

def check_git_ready():
    try:
        run(["git", "--version"])
    except Exception:
        log("ERROR: git is not installed or not in PATH.")
        log("Download from: https://git-scm.com/download/win")
        return False
    try:
        run(["git", "status"])
    except Exception:
        log("ERROR: This folder is not a git repository.")
        log("Run setup_github.bat first.")
        return False
    return True

def get_mtime(path: pathlib.Path) -> float:
    try:
        return path.stat().st_mtime
    except FileNotFoundError:
        return 0.0

def main():
    log("=" * 55)
    log("StartHub Africa Dashboard Auto-Sync Agent started")
    log(f"Watching Excel:     {EXCEL_PATH.name}")
    log(f"Watching Generator: {GENERATOR.name}")
    log(f"Check interval:     {CHECK_INTERVAL}s")
    log("=" * 55)

    if not EXCEL_PATH.exists():
        log(f"ERROR: Excel file not found: {EXCEL_PATH}")
        input("Press Enter to exit.")
        return

    if not check_git_ready():
        input("Press Enter to exit.")
        return

    last_excel_mtime = get_mtime(EXCEL_PATH)
    last_gen_mtime   = get_mtime(GENERATOR)

    while True:
        time.sleep(CHECK_INTERVAL)

        excel_mtime = get_mtime(EXCEL_PATH)
        gen_mtime   = get_mtime(GENERATOR)

        excel_changed = excel_mtime != last_excel_mtime and excel_mtime != 0
        gen_changed   = gen_mtime   != last_gen_mtime   and gen_mtime   != 0

        if excel_changed or gen_changed:
            # Update tracked times before regenerating
            last_excel_mtime = excel_mtime
            last_gen_mtime   = gen_mtime

            # Grace period — let OneDrive/editor finish writing
            time.sleep(5)

            if excel_changed and gen_changed:
                reason = "data + design update"
            elif excel_changed:
                reason = "data update"
            else:
                reason = "design update"

            ok = regenerate(reason)
            if ok:
                push(reason)

if __name__ == "__main__":
    main()
