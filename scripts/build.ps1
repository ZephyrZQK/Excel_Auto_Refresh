param(
    [string]$AppName = "ExcelAutoRefresh"
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ScriptDir ".venv"
$PythonExe = Join-Path $VenvDir "Scripts\python.exe"
$EntryPoint = Join-Path $ScriptDir "excel_auto_refresh_app.py"
$Requirements = Join-Path $ScriptDir "requirements.txt"
$DistDir = Join-Path $ScriptDir "dist"
$WorkDir = Join-Path $ScriptDir "build"
$SpecDir = $ScriptDir

if (-not (Test-Path $PythonExe)) {
    python -m venv $VenvDir
}

& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install -r $Requirements
& $PythonExe -m PyInstaller `
    --onefile `
    --windowed `
    --name $AppName `
    --clean `
    --distpath $DistDir `
    --workpath $WorkDir `
    --specpath $SpecDir `
    $EntryPoint

Write-Host "Built: $(Join-Path $DistDir "$AppName.exe")"
