@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo [錯誤] 找不到虛擬環境，請先雙擊 setup_install.bat 進行環境設置。
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

echo === 開始執行 run_d365_rpa.py ===
python run_d365_rpa.py

echo.
echo === 程式執行完畢 ===
pause
