# PowerShell Launcher for CLI
if (Test-Path ".\venv\Scripts\Activate.ps1") {
    & ".\venv\Scripts\Activate.ps1"
}
python main.py $args
