@echo off
REM Opens the Missing-IG helper in your browser through a local server.
REM (Browsers block reading players.xlsx if you open the page directly from disk.)
cd /d "%~dp0"
echo Starting local server. Close this black window when you're done.
start "" "http://localhost:8000/missing_ig.html"
python -m http.server 8000
pause
