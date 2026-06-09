@echo off
REM Double-click this to scrape real players from Transfermarkt into players.xlsx.
REM This is SLOW on purpose (to avoid getting blocked). Let it finish.
cd /d "%~dp0"
python scraper.py
echo.
echo If players.xlsx was created, opening the website via a local server...
start "" "http://localhost:8000/index.html"
python -m http.server 8000
pause
