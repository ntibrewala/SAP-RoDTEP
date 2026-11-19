$ErrorActionPreference = "Continue"

Write-Host "1. Stopping existing RoDTEP backend on port 5065..." -ForegroundColor Cyan
Stop-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -ErrorAction SilentlyContinue
schtasks /End /TN "SAP-RoDTEP-Backend" 2>$null | Out-Null
Get-NetTCPConnection -LocalPort 5065 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object {
    if ($_ -and $_ -gt 0) {
        Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue
        taskkill /F /PID $_ /T 2>$null | Out-Null
    }
}
Start-Sleep -Seconds 2

Write-Host "2. Setting up D:\BHV\rod\backend and D:\BHV\rod\frontend..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path "D:\BHV\rod\backend" | Out-Null
New-Item -ItemType Directory -Force -Path "D:\BHV\rod\frontend" | Out-Null

if (Test-Path "D:\BHV\rodtep\backend\node_modules") {
    if (-not (Test-Path "D:\BHV\rod\backend\node_modules")) {
        Write-Host "Copying node_modules from D:\BHV\rodtep\backend..." -ForegroundColor Yellow
        Copy-Item -Path "D:\BHV\rodtep\backend\node_modules" -Destination "D:\BHV\rod\backend\node_modules" -Recurse -Force
    }
}

if (Test-Path "D:\BHV\rodtep\rodtep_local_store.json") {
    if (-not (Test-Path "D:\BHV\rod\rodtep_local_store.json")) {
        Copy-Item -Path "D:\BHV\rodtep\rodtep_local_store.json" -Destination "D:\BHV\rod\rodtep_local_store.json" -Force
    }
}

Copy-Item -Path "N:\SAP-RoDTEP\SAP-RoDTEP\backend\*" -Destination "D:\BHV\rod\backend\" -Recurse -Force
Copy-Item -Path "N:\SAP-RoDTEP\SAP-RoDTEP\frontendbuilt\*" -Destination "D:\BHV\rod\frontend\" -Recurse -Force
Copy-Item -Path "N:\SAP-RoDTEP\SAP-RoDTEP\frontendbuilt\*" -Destination "C:\inetpub\wwwroot\rod\" -Recurse -Force

try {
    Import-Module WebAdministration -ErrorAction SilentlyContinue
    if (Test-Path "IIS:\Sites\Default Web Site\rod") {
        Set-ItemProperty "IIS:\Sites\Default Web Site\rod" -Name physicalPath -Value "D:\BHV\rod\frontend" -ErrorAction SilentlyContinue
        Write-Host "Pointed IIS /rod physical path directly to D:\BHV\rod\frontend" -ForegroundColor Green
    }
} catch {}

if (-not (Test-Path "D:\BHV\rod\backend\node_modules")) {
    Write-Host "Installing backend npm packages in D:\BHV\rod\backend..." -ForegroundColor Yellow
    Push-Location "D:\BHV\rod\backend"
    npm install
    Pop-Location
}

Write-Host "3. Starting SAP-RoDTEP-Backend from D:\BHV\rod\backend..." -ForegroundColor Cyan
$nodePath = (Get-Command node).Source
try {
    $action = New-ScheduledTaskAction -Execute $nodePath -Argument "server.js" -WorkingDirectory "D:\BHV\rod\backend"
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 365) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
    Register-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -Action $action -Trigger $trigger -Settings $settings -User "SYSTEM" -RunLevel Highest -Force -ErrorAction SilentlyContinue | Out-Null
    Start-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -ErrorAction SilentlyContinue
} catch {}

Start-Sleep -Seconds 2

# Verify port 5065 is listening; if Scheduled Task did not start (e.g. non-Admin PowerShell), launch node.exe directly in background
$listening = Get-NetTCPConnection -LocalPort 5065 -State Listen -ErrorAction SilentlyContinue
if (-not $listening) {
    Write-Host "Starting node.js backend process directly on port 5065 from D:\BHV\rod\backend..." -ForegroundColor Yellow
    Start-Process -FilePath $nodePath -ArgumentList "server.js" -WorkingDirectory "D:\BHV\rod\backend" -WindowStyle Hidden
    Start-Sleep -Seconds 2
}

$listeningAfter = Get-NetTCPConnection -LocalPort 5065 -State Listen -ErrorAction SilentlyContinue
if ($listeningAfter) {
    Write-Host "SUCCESS! Backend is actively listening on Port 5065 (PID: $($listeningAfter[0].OwningProcess)) and Frontend is deployed to D:\BHV\rod\frontend!" -ForegroundColor Green
} else {
    Write-Host "WARNING: Port 5065 is still not listening. Running node server.js to check for startup errors:" -ForegroundColor Red
    Push-Location "D:\BHV\rod\backend"
    node server.js
    Pop-Location
}
