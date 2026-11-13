@echo off
chcp 65001 >nul
cd /d "%~dp0"
pip install -r requirements.txt || goto :err
pyinstaller --noconfirm --onefile --windowed --name "簽核工具" app.py || goto :err
copy /y config.ini dist\config.ini >nul
set "OUT=E:\88. Claude\88_temp\品管系統驗報簽核"
if not exist "%OUT%" mkdir "%OUT%"
copy /y "dist\簽核工具.exe" "%OUT%\" >nul || goto :err
copy /y "config.ini" "%OUT%\" >nul || goto :err
echo.
echo 完成：成品已複製到 %OUT%
echo   簽核工具.exe 與 config.ini 需放在同一資料夾
pause
exit /b 0
:err
echo 打包失敗
pause
exit /b 1
