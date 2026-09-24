$ErrorActionPreference = 'Stop'
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $python)) {
    Write-Error 'Ambiente virtual ausente. Execute primeiro: .venv\Scripts\python.exe -m pip install -r requirements.txt'
}

Set-Location $PSScriptRoot
& $python app.py