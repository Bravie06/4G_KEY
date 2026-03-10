@echo off
echo Starting 4G KPI Report Generator...

:: Check if Python is installed
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed or not added to your PATH.
    echo Please install Python from https://www.python.org/downloads/
    pause
    goto :eof
)

:: Create a virtual environment if it doesn't exist
IF NOT EXIST "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

:: Activate the virtual environment
call venv\Scripts\activate

:: Install required packages
echo Installing dependencies...
pip install -r requirements.txt

:: Run the application
echo Starting the application...
python app.py

:: If the app closes unexpectedly, keep the window open to read errors
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with an error. Please read the error message above.
    pause
)

:: Deactivate after closing
deactivate