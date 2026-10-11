@echo off
rem Re-index everything in the data folder (only changed files are re-read).
cd /d "%~dp0"
if not exist data mkdir data
py netsearch.py index data
py netsearch.py prune
pause
