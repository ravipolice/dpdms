@echo off
setlocal
cd /d "%~dp0"
start "Employee Transfer Dashboard Server" /min "C:\Users\ravip\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "%~dp0server.py"
timeout /t 2 /nobreak >nul
start http://127.0.0.1:8765
endlocal
