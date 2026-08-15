# Fetch the data, then open the screener in the default browser.
# The Microsoft Store python.exe stub in WindowsApps is skipped on purpose.
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

$candidates = @(
    "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe"
)
$python = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $python) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notlike '*WindowsApps*') { $python = $cmd.Source }
}
if (-not $python) {
    Write-Host "找不到 Python，請先安裝：winget install Python.Python.3.13" -ForegroundColor Red
    exit 1
}

& $python "twstock_screener.py" @args
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$latest = Get-ChildItem "output\*.html" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latest) { Start-Process $latest.FullName }
