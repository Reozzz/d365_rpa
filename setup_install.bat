@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   D365 RPA 環境設置 / 套件檢查工具
echo ============================================
echo.

REM --- 檢查 Python 是否已安裝 ---
where python >nul 2>nul
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.10 以上版本。
    echo 下載頁面：https://www.python.org/downloads/
    echo 安裝時請務必勾選 "Add Python to PATH"
    pause
    exit /b 1
)

echo [OK] 偵測到 Python：
python --version
echo.

REM --- 建立虛擬環境（若尚未建立） ---
if not exist "venv\Scripts\activate.bat" (
    echo [設置] 找不到 venv，正在建立虛擬環境...
    python -m venv venv
    if errorlevel 1 (
        echo [錯誤] 建立虛擬環境失敗。
        pause
        exit /b 1
    )
) else (
    echo [OK] 已偵測到現有的 venv 虛擬環境。
)
echo.

REM --- 啟用虛擬環境 ---
call venv\Scripts\activate.bat

REM --- 升級 pip ---
echo [設置] 正在升級 pip...
python -m pip install --upgrade pip

REM --- 安裝 requirements.txt 內的套件 ---
if not exist "requirements.txt" (
    echo [錯誤] 找不到 requirements.txt，請確認檔案是否與此工具放在同一資料夾。
    pause
    exit /b 1
)

echo.
echo [設置] 正在安裝所需套件（requirements.txt）...
pip install -r requirements.txt
if errorlevel 1 (
    echo [錯誤] 套件安裝失敗，請檢查上方錯誤訊息。
    pause
    exit /b 1
)

REM --- 安裝 Playwright 瀏覽器核心（pip 只裝 Python 套件，不含瀏覽器本體） ---
echo.
echo [設置] 正在下載 Playwright 瀏覽器（Chromium）...
python -m playwright install chromium
if errorlevel 1 (
    echo [錯誤] Playwright 瀏覽器下載失敗，請檢查網路連線後重新執行本工具。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   環境設置完成！請雙擊 run_d365_rpa.bat 執行程式
echo ============================================
pause
