# Build Windows executable for Print Client Agent
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Test-Path "$root\.venv")) {
  python -m venv .venv
}
. "$root\.venv\Scripts\Activate.ps1"

pip install -r requirements.txt
pip install pyinstaller

pyinstaller --noconfirm --onefile --windowed --name PrintClientAgent `
  --icon "icon.ico" `
  --add-data "config.example.json;." `
  config_ui.py

pyinstaller --noconfirm --onefile --name PrintClientAgentService `
  --icon "icon.ico" `
  --add-data "config.example.json;." `
  --hidden-import "win32timezone" `
  service.py

Write-Host "Build complete. EXE in dist\PrintClientAgent.exe and dist\PrintClientAgentService.exe"
