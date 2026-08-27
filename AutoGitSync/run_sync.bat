@echo off
chcp 65001 > nul
set SCRIPT_DIR=%~dp0
echo ========================================================
echo [Auto Git Sync] 수동 동기화 실행 중...
echo ========================================================
python "%SCRIPT_DIR%auto_git_sync.py" --once
echo.
echo 작업이 완료되었습니다.
pause
