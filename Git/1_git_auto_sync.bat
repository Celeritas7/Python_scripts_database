@echo off
cd /d "%~dp0"

if not exist ".git" (
  echo ERROR: Not a git repository
  pause
  exit /b
)

set "msg="
set /p "msg=Commit message (press Enter for default): "
if "%msg%"=="" set "msg=Auto sync: update files"

git add -A

git diff --cached --quiet
if errorlevel 1 (
  git commit -m "%msg%"
  git push origin main
) else (
  echo No changes to commit.
)

pause
