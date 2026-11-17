@echo off
setlocal EnableDelayedExpansion

REM ---------- decide which folder to work on ----------
REM 1) a folder dragged onto this .bat  2) otherwise the current directory
REM (double-clicking sets the current directory to this file's folder)
if not "%~1"=="" (
  if exist "%~1\" cd /d "%~1"
)

echo ==========================================
echo   GIT UPLOAD
echo   Folder: %CD%
echo ==========================================
echo.

git --version >nul 2>&1
if errorlevel 1 (
  echo ERROR: Git is not installed or not on PATH.
  pause
  exit /b 1
)

REM ---------- are we inside a repo? ----------
set "inrepo="
for /f "delims=" %%r in ('git rev-parse --is-inside-work-tree 2^>nul') do set "inrepo=%%r"

if "!inrepo!"=="true" (
  for /f "delims=" %%t in ('git rev-parse --show-toplevel 2^>nul') do set "root=%%t"
  cd /d "!root!"
  echo Repo root: !CD!
  echo.
) else (
  echo No git repository found here.
  echo A new one will be created and linked to GitHub.
  echo.
  set "url="
  set /p "url=Paste GitHub repo URL (Enter to cancel): "
  if "!url!"=="" (
    echo Cancelled.
    pause
    exit /b 1
  )
  git init
  if errorlevel 1 (
    echo ERROR: git init failed.
    pause
    exit /b 1
  )
  git branch -M main
  git remote add origin "!url!"
  echo.
  echo Repository initialised.
  echo.
)

REM ---------- make sure an 'origin' remote exists ----------
git remote get-url origin >nul 2>&1
if errorlevel 1 (
  echo No 'origin' remote is configured.
  set "url="
  set /p "url=Paste GitHub repo URL (Enter to cancel): "
  if "!url!"=="" (
    echo Cancelled.
    pause
    exit /b 1
  )
  git remote add origin "!url!"
  echo.
)

REM ---------- detect current branch ----------
set "branch="
for /f "delims=" %%b in ('git branch --show-current 2^>nul') do set "branch=%%b"
if "!branch!"=="" set "branch=main"

echo Branch: !branch!
echo.
echo ---------- Changes detected ----------
git status --short
echo --------------------------------------
echo.

git add -A

git diff --cached --quiet
if errorlevel 1 (
  set "msg="
  set /p "msg=Commit message (Enter for default): "
  if "!msg!"=="" set "msg=Update %DATE% %TIME%"

  set "note="
  set /p "note=Extra notes (optional, Enter to skip): "

  echo.
  if "!note!"=="" (
    git commit -m "!msg!"
  ) else (
    git commit -m "!msg!" -m "!note!"
  )

  if errorlevel 1 (
    echo.
    echo ERROR: Commit failed.
    pause
    exit /b 1
  )
) else (
  echo No new changes to commit - checking for unpushed commits...
)

echo.
echo Pushing to origin/!branch! ...
echo.
git push -u origin "!branch!"

if errorlevel 1 (
  echo.
  echo ==========================================
  echo   PUSH FAILED
  echo ==========================================
  echo The remote probably has commits you don't have yet.
  echo Run 0_pull.bat first, then run this script again.
) else (
  echo.
  echo ==========================================
  echo   DONE - files uploaded to GitHub
  echo ==========================================
)

echo.
pause
endlocal
