@echo off

REM Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python is not installed. Please install Python using the exe file in this folder.
    echo Make sure to check the box that says "Add Python to PATH".
    pause
    exit /b 1
)

REM Check Python version
for /f "tokens=2 delims= " %%v in ('python --version') do set PYTHON_VERSION=%%v
for /f "tokens=1,2 delims=." %%a in ("%PYTHON_VERSION%") do (
    set MAJOR=%%a
    set MINOR=%%b
)
if %MAJOR% lss 3 (
    echo Python version must be at least 3.7. Current version: %PYTHON_VERSION%.
    pause
    exit /b 1
)
if %MAJOR%==3 if %MINOR% lss 7 (
    echo Python version must be at least 3.7. Current version: %PYTHON_VERSION%.
    pause
    exit /b 1
)

echo Python version %PYTHON_VERSION% is sufficient.

REM Install requirements from requirements.txt
echo Installing requirements...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo Failed to install requirements.
    pause
    exit /b 1
)

echo Requirements installed successfully.

REM Create a new batch file to run main.py
echo python main.py >> run.bat
echo You can now run the program by typing "run" into the command line or by double-clicking on run.bat.
pause