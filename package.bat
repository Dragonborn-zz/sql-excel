@echo off
cd /d "%~dp0"

if not exist .venv (
  python -m venv .venv
  if errorlevel 1 goto :fail
)
call .venv\Scripts\activate.bat

python -c "import webview, pandas, openpyxl, xlrd" 2>nul
if errorlevel 1 (
  python -m pip install -r requirements.txt
  if errorlevel 1 python -m pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
  if errorlevel 1 goto :fail
)

python -c "import PyInstaller" 2>nul
if errorlevel 1 (
  python -m pip install pyinstaller
  if errorlevel 1 python -m pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com
  if errorlevel 1 goto :fail
)

python -c "import PyInstaller.__main__ as m; m.run(['--noconfirm','--clean','--windowed','--name','SQLExcel','--add-data','static;static','main.py'])"
if errorlevel 1 goto :fail

echo.
python -c "print('\u6253\u5305\u5b8c\u6210\uff1adist\\SQLExcel\\SQLExcel.exe')"
python -c "print('\u8bf7\u8fde\u540c\u6574\u4e2a dist\\SQLExcel \u6587\u4ef6\u5939\u4e00\u8d77\u62f7\u8d1d\uff0c\u4e0d\u8981\u53ea\u590d\u5236 exe\u3002')"
python -c "print('\u76ee\u6807\u7535\u8111\u9700\u8981\u5df2\u5b89\u88c5 WebView2\u3002')"
pause
exit /b 0

:fail
echo.
python -c "print('\u6253\u5305\u5931\u8d25\u3002')"
pause
exit /b 1
