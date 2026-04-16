# setup.ps1

if (!(Test-Path ".venv")) {
    python -m venv .venv
}

& ".venv\Scripts\Activate.ps1"

pip install --upgrade pip

if (Test-Path "requirements.txt") {
    pip install -r requirements.txt
} else {
    Write-Host "⚠️ requirements.txt 없음"
}

Write-Host "🚀 Ready to go"