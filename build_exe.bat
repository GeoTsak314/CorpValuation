@echo off
cd /d %~dp0
python -m pip install -r requirements.txt
pyinstaller --noconfirm --onefile --windowed --name EnterpriseValueAnalyzer app.py
pause
