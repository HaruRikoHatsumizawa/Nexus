$ErrorActionPreference = 'Stop'
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $python)) {
    Write-Error 'Ambiente virtual ausente. Execute primeiro: .venv\Scripts\python.exe -m pip install -r requirements.txt'
}

Set-Location $PSScriptRoot
if (Test-Path (Join-Path $PSScriptRoot '.env')) {
    'NEXUS_ENV', 'PORT', 'SESSION_COOKIE_SECURE', 'SECRET_KEY', 'ADMIN_USER', 'ADMIN_PASS_HASH' |
        ForEach-Object { Remove-Item "Env:$_" -ErrorAction SilentlyContinue }
}
& $python app.py