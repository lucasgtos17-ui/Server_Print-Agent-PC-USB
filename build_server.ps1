$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Test-Path "$root\.venv")) {
  python -m venv .venv
}

. "$root\.venv\Scripts\Activate.ps1"

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller --noconfirm --clean --onefile --name PrintServerDashboard `
  --windowed `
  --icon "icon.ico" `
  --collect-all reportlab `
  --collect-all openpyxl `
  server_main.py

pyinstaller --noconfirm --clean --onefile --name PrintServerDashboardService `
  --icon "icon.ico" `
  --collect-all reportlab `
  --collect-all openpyxl `
  service.py

if (-not (Test-Path "$root\dist\config.json")) {
  Copy-Item "$root\config.example.json" "$root\dist\config.json" -Force
}

if (-not (Test-Path "$root\dist\data")) {
  New-Item -ItemType Directory -Path "$root\dist\data" | Out-Null
}

Copy-Item "$root\stop_server_service.ps1" "$root\dist\stop_server_service.ps1" -Force
Copy-Item "$root\stop_server_service.cmd" "$root\dist\stop_server_service.cmd" -Force

Write-Host "Build concluido:"
Write-Host " - $root\dist\PrintServerDashboard.exe"
Write-Host " - $root\dist\PrintServerDashboardService.exe"
Write-Host "Edite: $root\dist\config.json"
