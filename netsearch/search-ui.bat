@echo off
rem Double-click: opens the search page in your browser (local only).
cd /d "%~dp0"
py netsearch.py serve
