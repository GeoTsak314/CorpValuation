@echo off
cd /d %~dp0
python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name CorpValuation app.py
pause
