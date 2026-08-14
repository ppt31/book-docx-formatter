@echo off
echo ========================================================
echo Setting up Python Virtual Environment (venv)
echo ========================================================

if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)

echo Installing dependencies...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

echo.
echo ========================================================
echo Environment setup completed successfully!
echo You can run 'run_cli.bat' or 'run_web.bat'.
echo ========================================================
pause
