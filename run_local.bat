@echo off
title DCTtool Cloud Web Server
cd /d "%~dp0"
echo ========================================================
echo   Starting DCTtool Cloud Web Server...
echo   Open in Google Chrome: http://localhost:5000
echo ========================================================
python server.py
pause
