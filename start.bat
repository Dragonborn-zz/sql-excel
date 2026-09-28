@echo off
cd /d "%~dp0"
if not exist .venv (
  python -m venv .venv
)
call .venv\Scripts\activate.bat
python -c "import webview, pandas, openpyxl, xlrd" 2>nul
if errorlevel 1 (
  python -m pip install -r requirements.txt
  if errorlevel 1 python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
)
python main.py
