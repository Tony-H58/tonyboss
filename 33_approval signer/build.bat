@echo off
chcp 65001 >nul
cd /d "%~dp0"
pip install -r requirements.txt || goto :err
pyinstaller --noconfirm --onefile --windowed --name "簽核工具" app.py || goto :err
copy /y config.ini dist\config.ini >nul
echo.
echo 完成：dist\簽核工具.exe （config.ini 需放在同一資料夾）
pause
exit /b 0
:err
echo 打包失敗
pause
exit /b 1
