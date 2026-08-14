@echo off
echo Starting Web UI...
call venv\Scripts\activate.bat
streamlit run app.py
pause
