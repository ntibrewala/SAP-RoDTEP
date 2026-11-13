$ErrorActionPreference = "Continue"

Write-Host "1. Stopping existing RoDTEP backend..." -ForegroundColor Cyan
Stop-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -ErrorAction SilentlyContinue
Get-NetTCPConnection -LocalPort 5065 -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 1

Write-Host "2. Creating D:\BHV\rod\backend and D:\BHV\rod\frontend..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path "D:\BHV\rod\backend" | Out-Null
New-Item -ItemType Directory -Force -Path "D:\BHV\rod\frontend" | Out-Null

if (Test-Path "D:\BHV\rodtep\backend\node_modules") {
    if (-not (Test-Path "D:\BHV\rod\backend\node_modules")) {
        Write-Host "Copying node_modules from D:\BHV\rodtep\backend..." -ForegroundColor Yellow
        Copy-Item -Path "D:\BHV\rodtep\backend\node_modules" -Destination "D:\BHV\rod\backend\node_modules" -Recurse -Force
    }
}

if (Test-Path "D:\BHV\rodtep\rodtep_local_store.json") {
    Copy-Item -Path "D:\BHV\rodtep\rodtep_local_store.json" -Destination "D:\BHV\rod\rodtep_local_store.json" -Force
}

Copy-Item -Path "N:\SAP-RoDTEP\SAP-RoDTEP\backend\*" -Destination "D:\BHV\rod\backend\" -Recurse -Force
Copy-Item -Path "N:\SAP-RoDTEP\SAP-RoDTEP\frontendbuilt\*" -Destination "D:\BHV\rod\frontend\" -Recurse -Force
Copy-Item -Path "N:\SAP-RoDTEP\SAP-RoDTEP\frontendbuilt\*" -Destination "C:\inetpub\wwwroot\rod\" -Recurse -Force

if (-not (Test-Path "D:\BHV\rod\backend\node_modules")) {
    Write-Host "Installing backend npm packages in D:\BHV\rod\backend..." -ForegroundColor Yellow
    Push-Location "D:\BHV\rod\backend"
    npm install
    Pop-Location
}

Write-Host "3. Starting SAP-RoDTEP-Backend from D:\BHV\rod\backend..." -ForegroundColor Cyan
$nodePath = (Get-Command node).Source
$action = New-ScheduledTaskAction -Execute $nodePath -Argument "server.js" -WorkingDirectory "D:\BHV\rod\backend"
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 365) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -Action $action -Trigger $trigger -Settings $settings -User "SYSTEM" -RunLevel Highest -Force | Out-Null
Start-ScheduledTask -TaskName "SAP-RoDTEP-Backend"

Start-Sleep -Seconds 2
Write-Host "DONE! D:\BHV\rod is updated and backend is running on port 5065." -ForegroundColor Green
