@echo off
REM Default launcher: Web control deck.
call conda activate visa

echo Starting N9918A Web Control Deck...
echo Open http://127.0.0.1:5000 in your browser.
python web_app.py

pause
