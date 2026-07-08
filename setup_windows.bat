@echo off
echo Creating virtual environment...
python -m venv venv
echo Activating virtual environment...
call venv\Scripts\activate
echo Installing dependencies...
pip install -r requirements.txt
echo.
echo Setup complete!
echo Please remember to select the Python interpreter in VS Code:
echo 1. Press Ctrl+Shift+P
echo 2. Type "Python: Select Interpreter"
echo 3. Select the one in .\venv\Scripts\python.exe
echo.
pause
