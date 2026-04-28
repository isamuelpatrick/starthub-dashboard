"""
StartHub Africa Dashboard – Auto-Sync Agent
============================================
Watches the Impact Excel file for changes.
When a change is detected it:
  1. Regenerates dashboard.html
  2. Commits & pushes to GitHub Pages

Run this script at Windows startup (see start_sync.bat).
It runs silently in the background until you close the window.

CONFIGURE THE THREE SETTINGS BELOW before first run.
"""

import os
import sys
import time
import subprocess
import pathlib
import datetime
import shutil

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION  ← edit these three lines once, then never touch again
# ─────────────────────────────────────────────────────────────────────────────

# Full path to the folder that contains generate_dashboard.py and the Excel file
WORK_DIR = pathlib.Path(__file__).parent.resolve()   # auto-detects this folder

# Name of your Excel file (just the filename, not the full path)
EXCEL_FILE = "StartHub Africa Impact Dashboard (2).xlsx"

# How often to check for changes, in seconds (60 = once per minute)
CHECK_INTERVAL = 60

# ─────────────────────────────────────────────────────────────────────────────

EXCEL_PATH  = WORK_DIR / EXCEL_FILE
GENERATOR   = WORK_DIR / "generate_dashboard.py"
DASHBOARD   = WORK_DIR / "index.html"
LOG_FILE    = WORK_DIR / "sync_log.txt"

def log(msg: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def run(cmd: list, cwd=None):
    result = subprocess.run(
        cmd, cwd=cwd or WORK_DIR,
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())
    return result.stdout.strip()

def git_push():
    """Commit the new dashboard.html and push to GitHub."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    run(["git", "add", "index.html"])
    run(["git", "commit", "-m", f"Auto-update dashboard {ts}"])
    run(["git", "push"])

def regenerate():
    log("Excel change detected – regenerating dashboard…")
    try:
        result = subprocess.run(
            [sys.executable, str(GENERATOR)],
            cwd=WORK_DIR,
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or result.stdout)
        log("Dashboard generated successfully.")
    except Exception as e:
        log(f"ERROR generating dashboard: {e}")
        return False
    return True

def push():
    log("Pushing to GitHub Pages…")
    try:
        git_push()
        log("✅ Dashboard published – team link is live.")
    except Exception as e:
        log(f"ERROR pushing to GitHub: {e}")

def check_git_ready():
    """Verify git is installed and repo is initialised."""
    try:
        run(["git", "--version"])
    except Exception:
        log("ERROR: git is not installed or not in PATH.")
        log("Please install Git for Windows: https://git-scm.com/download/win")
        return False
    try:
        run(["git", "status"])
    except Exception:
        log("ERROR: This folder is not a git repository.")
        log("Run setup_github.bat first to initialise the repository.")
        return False
    return True

def main():
    log("=" * 55)
    log("StartHub Africa Dashboard Auto-Sync Agent started")
    log(f"Watching: {EXCEL_PATH}")
    log(f"Check interval: {CHECK_INTERVAL}s")
    log("=" * 55)

    if not EXCEL_PATH.exists():
        log(f"ERROR: Excel file not found at {EXCEL_PATH}")
        log("Check the WORK_DIR and EXCEL_FILE settings at the top of this script.")
        input("Press Enter to exit.")
        return

    if not check_git_ready():
        input("Press Enter to exit.")
        return

    last_mtime = EXCEL_PATH.stat().st_mtime

    while True:
        time.sleep(CHECK_INTERVAL)
        try:
            current_mtime = EXCEL_PATH.stat().st_mtime
        except FileNotFoundError:
            log("WARNING: Excel file temporarily not found (may be syncing). Retrying…")
            continue

        if current_mtime != last_mtime:
            last_mtime = current_mtime
            # Small grace period: OneDrive may still be writing the file
            time.sleep(5)
            ok = regenerate()
            if ok:
                push()
        # else: no change, stay silent

if __name__ == "__main__":
    main()
