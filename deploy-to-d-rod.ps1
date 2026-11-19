$ErrorActionPreference = "Continue"

Write-Host "1. Locating or installing NSSM (Windows Service Manager)..." -ForegroundColor Cyan
New-Item -ItemType Directory -Force -Path "D:\BHV\rod\backend" | Out-Null
New-Item -ItemType Directory -Force -Path "D:\BHV\rod\frontend" | Out-Null

$nssmCandidates = @(
    (Get-Command nssm -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -ErrorAction SilentlyContinue),
    "D:\BHV\rod\nssm.exe",
    "D:\BHV\nssm.exe",
    "C:\nssm\nssm.exe",
    "C:\nssm\win64\nssm.exe",
    "C:\Windows\System32\nssm.exe",
    "C:\ProgramData\chocolatey\bin\nssm.exe"
) | Where-Object { $_ -and (Test-Path $_) }

$nssmPath = $nssmCandidates | Select-Object -First 1

if (-not $nssmPath) {
    Write-Host "Downloading nssm.exe to D:\BHV\rod\nssm.exe..." -ForegroundColor Yellow
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $zipPath = "D:\BHV\rod\nssm.zip"
        $extractDir = "D:\BHV\rod\nssm_tmp"
        Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $zipPath -UseBasicParsing
        Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
        $foundExe = Get-ChildItem -Path $extractDir -Filter "nssm.exe" -Recurse | Where-Object { $_.FullName -like "*win64*" } | Select-Object -First 1
        if ($foundExe) {
            Copy-Item -Path $foundExe.FullName -Destination "D:\BHV\rod\nssm.exe" -Force
            $nssmPath = "D:\BHV\rod\nssm.exe"
        }
        Remove-Item -Path $zipPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Note: Could not auto-download nssm.cc ($($_.Exception.Message))." -ForegroundColor Yellow
    }
}

Write-Host "2. Stopping existing RoDTEP backend on port 5065..." -ForegroundColor Cyan
Stop-Service -Name "SAP-RoDTEP-Backend" -Force -ErrorAction SilentlyContinue
Stop-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName "SAP-RoDTEP-Backend" -Confirm:$false -ErrorAction SilentlyContinue
Get-NetTCPConnection -LocalPort 5065 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object {
    if ($_ -and $_ -gt 0) {
        Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue
        taskkill /F /PID $_ /T 2>$null | Out-Null
    }
}
Start-Sleep -Seconds 2

Write-Host "3. Deploying Backend and Frontend to D:\BHV\rod..." -ForegroundColor Cyan
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

try {
    if (-not (Get-NetFirewallRule -DisplayName "SAP-RoDTEP-Backend-5065" -ErrorAction SilentlyContinue)) {
        New-NetFirewallRule -DisplayName "SAP-RoDTEP-Backend-5065" -Direction Inbound -LocalPort 5065 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null
    }
} catch {}

if (-not (Test-Path "D:\BHV\rod\backend\node_modules")) {
    Write-Host "Installing backend npm packages in D:\BHV\rod\backend..." -ForegroundColor Yellow
    Push-Location "D:\BHV\rod\backend"
    npm install
    Pop-Location
}

$nodePath = (Get-Command node).Source

if ($nssmPath -and (Test-Path $nssmPath)) {
    Write-Host "4. Configuring & Starting SAP-RoDTEP-Backend Windows Service via NSSM ($nssmPath)..." -ForegroundColor Cyan
    & $nssmPath install "SAP-RoDTEP-Backend" $nodePath "server.js" 2>$null | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" Application $nodePath | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" AppParameters "server.js" | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" AppDirectory "D:\BHV\rod\backend" | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" DisplayName "SAP RoDTEP Scrip Backend (Port 5065)" | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" Description "Node.js SAP HANA & Service Layer Backend for RoDTEP/RoSCTL Scrip Management (D:\BHV\rod\backend)" | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" Start SERVICE_AUTO_START | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" AppStdout "D:\BHV\rod\backend\service.log" | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" AppStderr "D:\BHV\rod\backend\service-error.log" | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" AppRotateFiles 1 | Out-Null
    & $nssmPath set "SAP-RoDTEP-Backend" AppExit Default Restart | Out-Null
    & $nssmPath start "SAP-RoDTEP-Backend" 2>$null | Out-Null
    Start-Service -Name "SAP-RoDTEP-Backend" -ErrorAction SilentlyContinue
}

Start-Sleep -Seconds 2

$listening = Get-NetTCPConnection -LocalPort 5065 -State Listen -ErrorAction SilentlyContinue
if (-not $listening) {
    Write-Host "Starting node.js backend process directly on port 5065 from D:\BHV\rod\backend..." -ForegroundColor Yellow
    Start-Process -FilePath $nodePath -ArgumentList "server.js" -WorkingDirectory "D:\BHV\rod\backend" -WindowStyle Hidden
    Start-Sleep -Seconds 2
}

$listeningAfter = Get-NetTCPConnection -LocalPort 5065 -State Listen -ErrorAction SilentlyContinue
if ($listeningAfter) {
    $svc = Get-Service -Name "SAP-RoDTEP-Backend" -ErrorAction SilentlyContinue
    $svcStatus = if ($svc) { "Windows Service Status: $($svc.Status)" } else { "Direct Process" }
    Write-Host "SUCCESS! SAP-RoDTEP-Backend is listening on Port 5065 (PID: $($listeningAfter[0].OwningProcess) | $svcStatus)" -ForegroundColor Green
} else {
    Write-Host "WARNING: Port 5065 is still not listening. Check D:\BHV\rod\backend\service-error.log" -ForegroundColor Red
}
