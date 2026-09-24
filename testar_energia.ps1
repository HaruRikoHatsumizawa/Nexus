$ErrorActionPreference = 'Stop'
$url = 'http://127.0.0.1:5000/api/energia/leitura'

$leitura = @{
    dispositivo = 'ESP32-simulacao'
    energia_kwh = 0.125
    custo_brl = 0.12
    potencia_w = 275.0
    tensao_v = 220.0
    corrente_a = 1.25
    timestamp = (Get-Date).ToString('s')
} | ConvertTo-Json

$resposta = Invoke-RestMethod -Uri $url -Method Post -ContentType 'application/json' -Body $leitura
Write-Output $resposta.mensagem
Write-Output ('Dispositivo: ' + $resposta.leitura.dispositivo)
Write-Output ('Energia: ' + $resposta.leitura.energia_kwh + ' kWh')

$resumo = Invoke-RestMethod -Uri 'http://127.0.0.1:5000/api/energia'
Write-Output ('Total acumulado: ' + $resumo.resumo.energia_kwh + ' kWh')
Write-Output ('Leituras recebidas: ' + $resumo.resumo.leituras)