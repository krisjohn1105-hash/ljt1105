@echo off
chcp 65001 > nul
set SCRIPT_DIR=%~dp0
echo ========================================================
echo [Auto Git Sync] Windows 작업 스케줄러 등록 해제
echo ========================================================
python "%SCRIPT_DIR%auto_git_sync.py" --unregister
echo.
pause
