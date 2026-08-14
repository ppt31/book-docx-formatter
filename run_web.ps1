# PowerShell Launcher for Web UI
Write-Host "Starting Web UI..." -ForegroundColor Green
if (Test-Path ".\venv\Scripts\Activate.ps1") {
    & ".\venv\Scripts\Activate.ps1"
}
streamlit run app.py
