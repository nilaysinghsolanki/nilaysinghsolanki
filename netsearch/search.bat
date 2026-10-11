@echo off
rem Double-click: interactive search in this window.
cd /d "%~dp0"
py netsearch.py %*
pause
