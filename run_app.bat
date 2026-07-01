@echo off
cd /d "%~dp0"
echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Python was not found on this computer.
    echo Please install Python and tick "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo Installing/updating requirements...
python -m pip install -r requirements.txt

echo.
echo Starting Laurel Wreath Template Builder...
echo If the browser does not open, copy the Local URL shown below.
python -m streamlit run app.py
pause
