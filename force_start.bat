@echo off
setlocal enabledelayedexpansion

echo ==========================================
echo   BizStrategy AI Partner (Force Start)
echo ==========================================

echo [1/4] Killing existing processes...
taskkill /F /IM streamlit.exe /T 2>nul
taskkill /F /IM python.exe /T 2>nul
timeout /t 2 /nobreak >nul

echo [2/4] Activating virtual environment...
if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment .venv not found.
    pause
    exit /b
)
call .venv\Scripts\activate.bat

echo [3/4] Running system diagnosis...
python diagnose.py
if %ERRORLEVEL% neq 0 (
    echo [WARNING] Diagnosis found issues. Check output.
    pause
)

echo [4/4] Launching application in new window...
start "BizStrategy AI Partner" cmd /k "streamlit run app.py"

echo Done. This window will close automatically.
timeout /t 3 >nul
exit
