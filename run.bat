@echo off
setlocal EnableDelayedExpansion
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
    python -m venv --system-site-packages venv
)

:: Activate the virtual environment
call venv\Scripts\activate

:: Try installing dependencies normally using python -m pip
:: (Using just 'pip' can sometimes cause the batch script to exit prematurely on Windows)
echo Installing dependencies...
python -m pip install -r requirements.txt

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ----------------------------------------------------------------------
    echo WARNING: Dependency installation failed.
    echo It looks like your network might be blocking internet access for Python.
    echo ----------------------------------------------------------------------
    echo If you are behind a corporate proxy, please enter it below.
    echo Leave blank to skip and try running the app anyway.
    echo Format example: http://your.proxy.address:8080
    echo ----------------------------------------------------------------------
    set /p PROXY="Enter proxy URL or press Enter to skip: "

    IF NOT "!PROXY!"=="" (
        echo Retrying installation with proxy...
        python -m pip install --proxy="!PROXY!" -r requirements.txt
    )
)

:: Check if we have the modules installed
python -c "import pandas; import openpyxl" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ----------------------------------------------------------------------
    echo CRITICAL ERROR: Could not install 'pandas' and 'openpyxl'.
    echo Your company network is blocking the Python package manager.
    echo Please contact your IT department to allow python to download packages.
    echo Or install them manually using: python -m pip install pandas openpyxl
    echo ----------------------------------------------------------------------
    echo.
    pause
    deactivate
    goto :eof
)

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