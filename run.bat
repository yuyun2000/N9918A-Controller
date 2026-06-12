@echo off
REM Default launcher: Web control deck.
set N9918A_WEB_URL=http://127.0.0.1:5000
call conda activate visa

echo Starting N9918A Web Control Deck...
echo Browser will open: %N9918A_WEB_URL%
python web_app.py

pause
