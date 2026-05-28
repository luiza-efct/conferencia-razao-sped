$ErrorActionPreference = 'Stop'
$proj = $PSScriptRoot
$py = Join-Path $proj '.venv\Scripts\python.exe'

if (-not (Test-Path $py)) {
    Write-Host "ERRO: venv nao encontrado em $py" -ForegroundColor Red
    Write-Host "Rode: python -m venv .venv ; .\.venv\Scripts\pip install -r requirements.txt"
    exit 1
}

$env:PORT = '5000'
$env:FLASK_ENV = 'development'

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Flask EFCT — Conferencia Razao x SPED" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  URL local: http://localhost:5000/" -ForegroundColor Cyan
Write-Host "  Para parar: feche esta janela ou Ctrl+C" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

Set-Location $proj
& $py app.py
