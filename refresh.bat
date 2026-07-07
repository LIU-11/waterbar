@echo off
chcp 65001 >nul 2>&1
echo.
echo  === 看板数据手动刷新 ===
echo.
"C:\Users\47\.workbuddy\binaries\python\envs\default\Scripts\python.exe" "%~dp0auto_refresh.py"
echo.
pause
