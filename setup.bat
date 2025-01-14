echo "Installing required packages..."
pip install -r requirements.txt
echo "Installation complete. If any errors occurred"
if %ERRORLEVEL% != 0 (
    echo "Command failed. Please ask for help."
    exit /b %ERRORLEVEL%
)
echo "All packages installed successfully."
echo "Creating startup script..."
echo "python main.py" > run.bat
echo "Startup script created."
echo "Setup complete. You can now run the program by running typing ""run"" into your terminal or clicking the run file in File Explorer!"