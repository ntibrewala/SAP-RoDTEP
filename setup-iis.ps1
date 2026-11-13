# Run this script as Administrator in PowerShell on 10.10.0.15
$ErrorActionPreference = "Stop"

$AppName = "rod"
$AppPoolName = "SAP-RoDTEP-Pool"
$BackendPort = 5065
$SourceRoot = $PSScriptRoot
$DTargetRoot = "D:\BHV\rodtep"
$DTargetBackend = "D:\BHV\rodtep\backend"
$DTargetFrontend = "D:\BHV\rodtep\frontendbuilt"
$IisTargetFolder = "C:\inetpub\wwwroot\rod"

Write-Host "1. Shifting built files (frontendbuilt & backend) to $DTargetRoot and $IisTargetFolder..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $DTargetBackend -Force | Out-Null
New-Item -ItemType Directory -Path $DTargetFrontend -Force | Out-Null
New-Item -ItemType Directory -Path $IisTargetFolder -Force | Out-Null

Copy-Item -Path "$SourceRoot\backend\*" -Destination $DTargetBackend -Recurse -Force
Copy-Item -Path "$SourceRoot\frontendbuilt\*" -Destination $DTargetFrontend -Recurse -Force
Copy-Item -Path "$SourceRoot\frontendbuilt\*" -Destination $IisTargetFolder -Recurse -Force

Write-Host "2. Setting IIS folder permissions on $IisTargetFolder..." -ForegroundColor Cyan
icacls $IisTargetFolder /grant "IIS_IUSRS:(OI)(CI)RX" /grant "IUSR:(OI)(CI)RX" /T /Q | Out-Null

Import-Module WebAdministration

Write-Host "3. Configuring IIS Application Pool ($AppPoolName)..." -ForegroundColor Cyan
if (!(Test-Path "IIS:\AppPools\$AppPoolName")) {
    New-WebAppPool -Name $AppPoolName | Out-Null
}
Set-ItemProperty "IIS:\AppPools\$AppPoolName" -Name "managedRuntimeVersion" -Value ""

# Enable ARR Reverse Proxy in IIS if Application Request Routing is installed
try {
    Set-WebConfigurationProperty -pspath 'MACHINE/WEBROOT/APPHOST' -filter "system.webServer/proxy" -name "enabled" -value "True" -ErrorAction SilentlyContinue
} catch {}

Write-Host "4. Creating IIS Sub-Application '/$AppName' under Default Web Site (http://10.10.0.15/rod)..." -ForegroundColor Cyan
$DefaultSite = (Get-Website | Where-Object { $_.ID -eq 1 -or $_.Name -eq "Default Web Site" } | Select-Object -First 1).Name
if (!$DefaultSite) { $DefaultSite = "Default Web Site" }

if (Get-WebApplication -Site $DefaultSite -Name $AppName -ErrorAction SilentlyContinue) {
    Remove-WebApplication -Site $DefaultSite -Name $AppName
}
New-WebApplication -Site $DefaultSite -Name $AppName -PhysicalPath $IisTargetFolder -ApplicationPool $AppPoolName -Force | Out-Null

Write-Host "5. Opening Windows Firewall for LAN access (Port 80 & Backend Port $BackendPort)..." -ForegroundColor Cyan
if (!(Get-NetFirewallRule -DisplayName "SAP-RoDTEP-IIS-80" -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName "SAP-RoDTEP-IIS-80" -Direction Inbound -LocalPort 80 -Protocol TCP -Action Allow | Out-Null
}
if (!(Get-NetFirewallRule -DisplayName "SAP-RoDTEP-API-$BackendPort" -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName "SAP-RoDTEP-API-$BackendPort" -Direction Inbound -LocalPort $BackendPort -Protocol TCP -Action Allow | Out-Null
}

Write-Host "6. Starting Backend Server from $DTargetBackend on Port $BackendPort..." -ForegroundColor Cyan
Set-Location $DTargetBackend
if (!(Test-Path "$DTargetBackend\node_modules")) {
    npm install
}

if (Get-Command pm2 -ErrorAction SilentlyContinue) {
    pm2 delete sap-rodtep 2>$null
    pm2 start server.js --name "sap-rodtep"
    pm2 save
    Write-Host ""
    Write-Host "SUCCESS! RoDTEP is live on the LAN at: http://10.10.0.15/rod" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "SUCCESS! IIS '/rod' is configured. Starting Node backend on port $BackendPort..." -ForegroundColor Green
    Write-Host "Access on LAN at: http://10.10.0.15/rod" -ForegroundColor Green
    node server.js
}
