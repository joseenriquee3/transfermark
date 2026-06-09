@echo off
REM Double-click this to view the website properly.
REM Browsers block reading players.xlsx when you open index.html directly from
REM your hard drive ("Failed to fetch"). This starts a tiny local web server
REM so the page loads through http://localhost instead, which works.
cd /d "%~dp0"
echo Starting local server...
echo When you are done, close this black window to stop the server.
start "" "http://localhost:8000/index.html"
python -m http.server 8000
pause
