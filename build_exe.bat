@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Cài công cụ build...
python -m pip install -q pyinstaller pillow

echo Tạo icon assets\app.ico ...
python build_icon.py
if errorlevel 1 exit /b 1

echo Đóng gói AppPrintPO.exe (PyInstaller)...
python -m PyInstaller --noconfirm AppPrintPO.spec

if errorlevel 1 (
  echo Lỗi PyInstaller.
  exit /b 1
)

echo.
echo Xong: dist\AppPrintPO.exe
pause
