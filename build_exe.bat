@echo off
cd /d %~dp0
python -m pip install -r requirements.txt
pyinstaller --noconfirm --windowed --name CORPValueAnalyzer app591.py
pause
