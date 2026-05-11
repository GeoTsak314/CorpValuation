@echo off
cd /d %~dp0
python -m pip install -r requirements.txt
pyinstaller --noconfirm --windowed --name CorpValueAnalyzer app59.py
pause
