# Read OFCSCAN.INI in C:\Program Files (x86)\Trend Micro\OfficeScan Client\ directory
# find line beginning ServerAddr=
$OsceServerAddrLine = Get-Content -Path "C:\Program Files (x86)\Trend Micro\OfficeScan Client\OFCSCAN.INI" | Where-Object { $_ -match 'ServerAddr=' }
# Split on = symbol and take second item (OSCE FQDN)
#VARIABLE - OSCE FQDN
$OsceServerAddr = $OsceServerAddrLine.Split('=')[1]
Write-Host "`nThe OSCE FQDN is $OsceServerAddr `n" -ForegroundColor Yellow

# find line beginning Master_DomainPort=
$OsceServerMasterDomainPortLine = Get-Content -Path "C:\Program Files (x86)\Trend Micro\OfficeScan Client\OFCSCAN.INI" | Where-Object { $_ -match 'Master_DomainPort=' }
# Split on = symbol and take second item (OSCE Master Domain Port)
#VARIABLE - OSCE Server Port
$OsceServerMasterDomainPort = $OsceServerMasterDomainPortLine.Split('=')[1]
Write-Host "`nThe OSCE Server Master Domain Port is $OsceServerMasterDomainPort `n" -ForegroundColor Yellow

# find line beginning Client_LocalServer_Port=
$OsceClientLocalServerPortLine = Get-Content -Path "C:\Program Files (x86)\Trend Micro\OfficeScan Client\OFCSCAN.INI" | Where-Object { $_ -match 'Client_LocalServer_Port=' }
# Split on = symbol and take second item (OSCE Client Local Server Port)
#VARIABLE - OSCE Client Listening Port
$OsceClientLocalServerPort = $OsceClientLocalServerPortLine.Split('=')[1] #VARIABLE
Write-Host "`nThe OSCE Client Local Server Port is $OsceClientLocalServerPort `n" -ForegroundColor Yellow

#VARIABLE - Certificate Path
$OsceCertPath = "\\" + $OsceServerAddr + "\ofcscan\pccnt\common\ofcntcer.dat"
Write-Host "`nThe Certificate Path is $OsceCertPath `n" -ForegroundColor Yellow

#VARIABLE - Ipxfer exe Path
$IpxferPath = "\\" + $OsceServerAddr + "\ofcscan\Admin\Utility\IpXfer\IpXfer_x64.exe"
Write-Host "`nIpXfer Path is $IpxferPath `n" -ForegroundColor Yellow

#VARIABLE - OfficeScan Password
$OscePassword = Read-Host "`nPlease enter OSCE Agent Password"

#VARIABLE - IpXfer Command
$IpxferCommand = $IpxferPath + " -s " + $OsceServerAddr + " -p " + $OsceServerMasterDomainPort + " -c " + $OsceClientLocalServerPort + " -e " + $OsceCertPath + " -pwd " + $OscePassword

Write-Warning "`nReady to run the IpXfer Refresh to the current OSCE Server `n"
cmd /c "pause"

Write-Host "`nIpXfer is running... Please wait... You can check Task Manager to see the status of IpXfer_x64.exe`n"

Invoke-Expression $IpxferCommand

# Wait for ntrtscan to start
while ( (Get-Service -Name ntrtscan).Status -ne "Running" ) {
"Waiting for ntrtscan to start ..."
Start-Sleep -Seconds 10
}
"ntrtscan is Up!"

# final check for ntrtscan
Write-Host "`nChecking Trend Micro OfficeScan RealTime Scan Service... `n"
Get-Service | Where {$_.Name -eq "ntrtscan"}

Write-Host "`nCompleted!" -ForegroundColor Yellow