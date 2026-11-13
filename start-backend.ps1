# Run this script in PowerShell to install and start the RoDTEP Scrip Management Backend on port 5060
$ErrorActionPreference = "Stop"

$BackendDir = Join-Path $PSScriptRoot "backend"
Set-Location $BackendDir

Write-Host "1. Installing backend dependencies..." -ForegroundColor Cyan
npm install

if (Get-Command pm2 -ErrorAction SilentlyContinue) {
    Write-Host "2. Starting backend with PM2 on Port 5060..." -ForegroundColor Cyan
    pm2 delete sap-rodtep 2>$null
    pm2 start server.js --name "sap-rodtep"
    pm2 save
    Write-Host "RoDTEP Backend is running in background via PM2 on http://127.0.0.1:5060" -ForegroundColor Green
} else {
    Write-Host "2. Starting backend with Node.js on Port 5060..." -ForegroundColor Cyan
    node server.js
}
