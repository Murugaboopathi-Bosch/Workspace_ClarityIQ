@echo off
echo ================================================
echo   CODE TO CLARITY — Local Runner (Windows)
echo   Bosch Digital Innovation League 2026
echo ================================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Download from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Install dependencies if needed
echo [1/3] Installing dependencies...
pip install -r requirements.txt -q

REM Check for API key
if "%GEMINI_API_KEY%"=="" (
    echo.
    echo [WARNING] GEMINI_API_KEY not set as environment variable.
    echo           Make sure you have pasted your key in code_to_clarity.py CONFIG section.
    echo.
)

REM Run the agent
echo [2/3] Running Code to Clarity Agent...
echo.
python code_to_clarity.py %1 %2

echo.
echo [3/3] Done! Check the 'output' folder for your BRD and FRD documents.
pause
