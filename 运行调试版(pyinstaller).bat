@echo off
setlocal
cd /d "%~dp0"
set WOA_DEBUG=1
start "" "WOA_Debug.exe"
endlocal