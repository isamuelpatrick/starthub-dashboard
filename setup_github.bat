@echo off
:: ============================================================
:: StartHub Africa Dashboard – GitHub Setup (run once only)
:: ============================================================
:: Before running:
::   1. Create a FREE GitHub account at https://github.com
::   2. Create a new repository called "starthub-dashboard"
::      (tick "Add a README file" when creating it)
::   3. Enable GitHub Pages:
::        Repo → Settings → Pages → Source: "Deploy from branch"
::        Branch: main  /  Folder: / (root)  → Save
::   4. Edit the two lines marked EDIT below
::   5. Double-click this file to run
:: ============================================================

:: ── EDIT THESE TWO LINES ─────────────────────────────────────
set GITHUB_USER=isamuelpatrick
set REPO_URL=https://github.com/isamuelpatrick/starthub-dashboard.git
:: ─────────────────────────────────────────────────────────────

echo.
echo  StartHub Africa Dashboard – GitHub Setup
echo  ==========================================

:: Check git is installed
git --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Git is not installed.
    echo  Download it from: https://git-scm.com/download/win
    echo  Then re-run this script.
    pause
    exit /b 1
)

:: Move into the Impact Data folder (same folder as this .bat file)
cd /d "%~dp0"

:: Initialise git repo if not already done
if not exist ".git" (
    echo.
    echo  Initialising git repository…
    git init
    git branch -M main
)

:: Set the remote (or update it)
git remote remove origin 2>nul
git remote add origin %REPO_URL%

:: Set your identity (GitHub will use this for commits)
git config user.email "%GITHUB_USER%@users.noreply.github.com"
git config user.name "StartHub Dashboard Bot"

:: Create a .gitignore so only the right files are tracked
echo *.xlsx > .gitignore
echo sync_log.txt >> .gitignore
echo __pycache__/ >> .gitignore
echo *.pyc >> .gitignore

:: Add dashboard + scripts and push
echo.
echo  Adding files and pushing to GitHub…
git add dashboard.html generate_dashboard.py sync_dashboard.py setup_github.bat
git commit -m "Initial dashboard publish" 2>nul || echo  (nothing new to commit – already up to date)
git push -u origin main

echo.
echo  ============================================================
echo   Setup complete!
echo.
echo   Your live dashboard URL will be:
echo   https://%GITHUB_USER%.github.io/starthub-dashboard/dashboard.html
echo.
echo   GitHub Pages can take 1-2 minutes to activate the first time.
echo   Share that link in Teams – it never changes.
echo  ============================================================
echo.
pause
