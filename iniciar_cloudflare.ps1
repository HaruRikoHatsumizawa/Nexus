$ErrorActionPreference = 'Stop'
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
$logs = Join-Path $PSScriptRoot 'logs'

if (-not (Test-Path $python)) {
    throw 'Ambiente virtual ausente. Execute .venv\Scripts\python.exe -m pip install -r requirements.txt'
}
if (-not (Test-Path (Join-Path $PSScriptRoot '.env'))) {
    throw 'Arquivo .env ausente. Configure-o copiando e preenchendo .env.example.'
}

$cloudflared = Get-Command cloudflared -ErrorAction SilentlyContinue
if (-not $cloudflared) {
    throw 'cloudflared nao encontrado. Instale-o antes de iniciar o tunnel.'
}
if ([string]::IsNullOrWhiteSpace($env:CLOUDFLARE_TUNNEL_TOKEN)) {
    throw 'Defina CLOUDFLARE_TUNNEL_TOKEN nas variaveis de ambiente do Windows.'
}

$port = 5000
if (-not [string]::IsNullOrWhiteSpace($env:PORT)) {
    if (-not [int]::TryParse($env:PORT, [ref]$port) -or $port -lt 1 -or $port -gt 65535) {
        throw 'PORT deve ser um numero entre 1 e 65535.'
    }
}

New-Item -ItemType Directory -Path $logs -Force | Out-Null
$env:NEXUS_ENV = 'production'
$env:SESSION_COOKIE_SECURE = 'true'
$env:PORT = [string]$port

$server = Start-Process `
    -FilePath $python `
    -ArgumentList @('-m', 'waitress', '--listen', "127.0.0.1:$port", 'app:app') `
    -WorkingDirectory $PSScriptRoot `
    -RedirectStandardOutput (Join-Path $logs 'waitress.log') `
    -RedirectStandardError (Join-Path $logs 'waitress-error.log') `
    -PassThru

try {
    $ready = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $server.Refresh()
        if ($server.HasExited) {
            throw "O servidor encerrou durante a inicializacao. Consulte $logs\waitress-error.log."
        }
        try {
            Invoke-WebRequest -Uri "http://127.0.0.1:$port/login" -TimeoutSec 2 | Out-Null
            $ready = $true
            break
        }
        catch {
            Start-Sleep -Seconds 1
        }
    }
    if (-not $ready) {
        throw "O servidor nao respondeu em 30 segundos. Consulte $logs\waitress-error.log."
    }

    Write-Host "Servidor local ativo em http://127.0.0.1:$port. Pressione Ctrl+C para encerrar."
    & $cloudflared.Source tunnel run --token $env:CLOUDFLARE_TUNNEL_TOKEN
    if ($LASTEXITCODE -ne 0) {
        throw "cloudflared encerrou com codigo $LASTEXITCODE."
    }
}
finally {
    $server.Refresh()
    if (-not $server.HasExited) {
        Stop-Process -Id $server.Id
        $server.WaitForExit(10000) | Out-Null
    }
}
