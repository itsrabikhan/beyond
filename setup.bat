@echo off
echo Installing required packages...
pip install -r requirements.txt
if %ERRORLEVEL%==0 (
    echo All packages installed successfully.
echo Creating startup script...
echo python main.py > run.bat
echo Startup script created.
echo Setup complete. You can now run the program by running typing "run" into your terminal or clicking the run file in File Explorer!
) else (
    echo An error occurred while installing packages. Please try again or ask for help.
    exit /b
)