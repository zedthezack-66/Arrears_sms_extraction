# Build script: creates a venv, installs dependencies, and builds a single-file exe with PyInstaller
# Usage: Right-click > Run with PowerShell, or execute from PowerShell prompt.

$here = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $here

if (-Not (Test-Path -Path .venv)) {
    python -m venv .venv
}

$py = Join-Path $here ".venv\Scripts\python.exe"

& $py -m pip install --upgrade pip
& $py -m pip install pyinstaller pandas openpyxl

# Clean previous builds
if (Test-Path -Path build) { Remove-Item -Recurse -Force build }
if (Test-Path -Path dist) { Remove-Item -Recurse -Force dist }
if (Test-Path -Path "Xtenda_SMS_Extractor.exe") { Remove-Item -Force "Xtenda_SMS_Extractor.exe" }

# Build single-file windowed exe (no console). If you want console, remove --windowed
& $py -m PyInstaller --noconfirm --clean --onefile --windowed --name Xtenda_SMS_Extractor gui_runner.py

if (Test-Path -Path dist\Xtenda_SMS_Extractor.exe) {
    Copy-Item -Path dist\Xtenda_SMS_Extractor.exe -Destination . -Force
    Write-Host "Build succeeded: Xtenda_SMS_Extractor.exe placed in $here"
} else {
    Write-Error "Build failed. See dist and build folders for logs." 
}
