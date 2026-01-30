#This is a modification of FRCS-HW_SW_Inventory_v1.05 by an unknown author.
#Author of Modification Daniel Barker daniel.barker.27@spaceforce.mil
#Released under GPL 2.0 Liscence

<#
Change Log
V.5 Added Local User and Log File Scan
V.6 Changed report Directory so it does not dump into the script root directory.
V.65 Set to fix file permissions from SCAP and Copy over the SFC log
V.66 Added Prompt to run scap and sfc.
V.7 Cleaned it up a bit.
V.8 Added CSV output for users
V.85 Added License Info for Windows and Office
V.90 Completely re-wrote script to allow it to run on Windows 7 with Powershell 2.0. Output for everything except SCAP is in one txt file.
V.92 Changing output to multiple CSV Adding Differing SFC and Removing Components and Programs
V.95 Hardware is output into Basic Information and Installed Software is added to its own csv
Known Working Systems:
Windows 7 Powershell 2.0
Windows 10 Powershell 5.0
#>



<#Notes:
To change where this script looks for SCAP change $ScapLocation It is currently around line 79ish by default it looks in $ScriptDir recursivly.
#>


Write-Host -ForegroundColor Green "Initalizing Variables"
$osInfo = Get-WmiObject -Class Win32_OperatingSystem | Select-Object Caption, BuildNumber, Manufacturer
$csInfo = Get-WmiObject -Class Win32_ComputerSystem | Select-Object Manufacturer, Model, Name
$currentUser = Get-WmiObject Win32_ComputerSystem | Select-Object -ExpandProperty UserName
$networkInfo = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.ipaddress -ne $Null}).IPAddress | Select-Object -First 1
$networkMac = (Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.ipenabled -eq $true }).MacAddress | Select-Object -First 1
$systemInfo = Get-WmiObject -Class Win32_ComputerSystemProduct | Select-Object IdentifyingNumber, Name, Vendor, Version
$biosInfo = Get-WmiObject -Class Win32_BIOS | Select-Object SMBIOSBIOSVersion,Manufacturer,Version
$Liscense = Get-WmiObject -Class softwarelicensingproduct | Where-Object {$_.PartialProductKey} | Select-Object Name, Description, LicenseStatus
$Model = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty Model
$CPU = Get-WmiObject -Class win32_processor | Select-Object -ExpandProperty Name
$RAM = Get-WmiObject -Class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | ForEach-Object { [math]::Round(($_.sum / 1GB),2) }
$Storage = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$env:systemdrive'" | ForEach-Object { [math]::Round($_.Size / 1GB,2) }
$ScriptDir= Split-Path -Parent -Path $MyInvocation.MyCommand.Path
$now = Get-Date -Format 'yyyyMMdd-HHmm'
$MOASPrompt = "MOAS#>"

Write-Host -ForegroundColor Green "Creating the Save Folder"
#Creat the Save Directory for this scan
$ShortDate = Get-Date -Format "yyyyMMdd"
$ScanSaveDir = "$ScriptDir\$ShortDate-$env:COMPUTERNAME"
$null = New-Item -ItemType Directory -Force -Path $ScanSaveDir
$scriptLocation = $PSScriptRoot
# This is the output file
$CSVBasicInfo = "$ScanSaveDir\$env:COMPUTERNAME-BasicInfo-$now.csv"
$CSVUpdatesandHotfixes = "$ScanSaveDir\$env:COMPUTERNAME-UpdateandHotfixes-$now.csv"
$CSVLocalUserList = "$ScanSaveDir\$env:COMPUTERNAME-LocalUsers-$now.csv"
$CSVPPS = "$ScanSaveDir\$env:COMPUTERNAME-PPS-$now.csv"
$CSVLogs = "$ScanSaveDir\$env:COMPUTERNAME-Logs-$now.csv"
$TXTSfc = "$ScanSaveDir\$env:COMPUTERNAME-SFC-$now.txt"
$CSVInstalledSoftware = "$ScanSaveDir\$env:COMPUTERNAME-InstalledSoftware-$now.csv"
#Add location of SCAP here. This assumes scap has been extracted to the Script Root\scc_#.#
$ScapLocation = Get-ChildItem $ScriptDir -Filter cscc.exe -Recurse -ErrorAction SilentlyContinue| % { $_.FullName }

Write-Host -ForegroundColor Red "Would you Like to Run SCAP?"
Write-Host -ForegroundColor Red "1. Yes"
Write-Host -ForegroundColor Red "2. No"
$RunSCAP = Read-Host -Prompt $MOASPrompt
if ($RunSCAP -ne "1" -and $RunSCAP -ne "2"){
    Read-Host -Prompt "Please Enter 1 or 2"
}

Write-Host -ForegroundColor Red "Would you Like to Run SFC?"
Write-Host -ForegroundColor Red "1. SCANNOW"
Write-Host -ForegroundColor Red "2. VERIFYONLY"
Write-Host -ForegroundColor Red "3. No"
$RunSFC = Read-Host -Prompt $MOASPrompt
if ($RunSFC -ne "1" -and $RunSFC -ne "2"-and $RunSFC -ne "3"){
    Read-Host -Prompt "Please Enter 1,2 or 3"
}

Write-Host -ForegroundColor Green "Writing Basic Information"
#Add the basic information
$BasicInfo = @()

$BasicInfo += New-Object PSObject -Property @{Title="Computer Manufacturer"; Data=$csInfo.Manufacturer}

$BasicInfo += New-Object PSObject -Property @{Title="Comuter Model"; Data=$csInfo.Model}

$BasicInfo += New-Object PSObject -Property @{Title="Computer Name"; Data=$csInfo.Name}

$BasicInfo += New-Object PSObject -Property @{Title="Identifying Number"; Data=$systemInfo.IdentifyingNumber}

$BasicInfo += New-Object PSObject -Property @{Title="OS Caption"; Data=$osInfo.Caption}

$BasicInfo += New-Object PSObject -Property @{Title="BuildNumber"; Data=$osInfo.BuildNumber}

$BasicInfo += New-Object PSObject -Property @{Title="OS Manufaturer"; Data=$osInfo.Manufacturer}

$BasicInfo += New-Object PSObject -Property @{Title="Current User"; Data=$currentUser}

$BasicInfo += New-Object PSObject -Property @{Title="IP Address"; Data=$networkInfo}

$BasicInfo += New-Object PSObject -Property @{Title="MAC Address"; Data=$networkMac}

$BasicInfo += New-Object PSObject -Property @{Title="Bios SMBIOSBIOSVersion"; Data=$biosInfo.SMBIOSBIOSVersion}

$BasicInfo += New-Object PSObject -Property @{Title="Bios Manufacturer"; Data=$biosInfo.Manufacturer}

$BasicInfo += New-Object PSObject -Property @{Title="Bios Version"; Data=$biosInfo.Version}

$BasicInfo += New-Object PSObject -Property @{Title="CPU"; Data=$CPU}

$BasicInfo += New-Object PSObject -Property @{Title="RAM"; Data=$RAM}

$BasicInfo += New-Object PSObject -Property @{Title="Storage"; Data=$Storage}




$Liscense | ForEach-Object {
    # Create a custom object for each string
    $BasicInfo += New-Object PSObject -Property @{
        "Title" = "Software Liscense Name"
        "Data"= $_.Name
        }
    $BasicInfo += New-Object PSObject -Property @{
    "Title" = "Software Liscense Description"
    "Data"= $_.Description
        }
    $BasicInfo += New-Object PSObject -Property @{
    "Title" = "Software Liscense Status"
    "Data"= $_.LicenseStatus
        }

}


$BasicInfo | Select-Object Title, Data | Export-Csv -Path $CSVBasicInfo -NoTypeInformation


Write-Host -ForegroundColor Green "Getting Local Users"


#Add User List
$adsi = [ADSI]"WinNT://$env:COMPUTERNAME"
$LocalUsers = $adsi.Children | where {$_.SchemaClassName -eq 'user'} | Foreach-Object {
    $groups = $_.Groups() | Foreach-Object {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
    $_ | Select-Object @{n='UserName';e={$_.Name}}, @{n='Description';e={$_.Description}}, @{n='Groups';e={$groups -join ';'}}
}
$LocalUsers | Export-Csv -Path $CSVLocalUserList -NoTypeInformation

Write-Host -ForegroundColor Green "Getting HotFixes"

# Updates and Hot Fixes
$InstalledUpdates = Get-HotFix | Select-Object Description, HotFixID, Installedby, InstalledOn | Export-Csv -Path $CSVUpdatesandHotfixes -NoTypeInformation

#Software List

$InstalledSoftware = Get-WmiObject -Query "SELECT * FROM Win32_Product" | Select-Object Name, Vendor, Version | Export-Csv -Path $CSVInstalledSoftware -NoTypeInformation


#PPS

$Processes = @{}
$Processes = Get-Process
Write-Host -ForegroundColor Green "Getting TCP and UDP Ports"
    $netstat = netstat -a -n -o | findstr /R /C:"^  TCP"
        $Ports = $netstat | ForEach-Object {
            $Parts = $_.Trim() -split "\s+"
            New-Object PSObject -Property @{
                Proto = $Parts[0]
                LocalAddress = $Parts[1].Substring(0,$Parts[1].LastIndexOf(':')) -replace '\[','' -replace '\]',''
                RemoteAddress = $Parts[2].Substring(0,$Parts[2].LastIndexOf(':')) -replace '\[','' -replace '\]',''
                State = $Parts[3]
                PID = $Parts[4]
                LocalPort = $Parts[1].Split(':')[-1]
                RemotePort = $Parts[2].Split(':')[-1]
                ProcessName = $Processes[[int]$Parts[4]].ProcessName
                }
             } | Select-Object LocalAddress,
        RemoteAddress,
        @{Name="Proto";Expression={"TCP"}},
        LocalPort,RemotePort,State,PID,
       ProcessName |Where-Object { $_.LocalAddress -notmatch "127.0.0.1|0.0.0.0|::1|::" -and $_.RemoteAddress -notmatch "127.0.0.1|0.0.0.0|::1|::" -and $_.State -eq "ESTABLISHED"} | Sort-Object -Property ProcessName

$Ports | add-member –membertype NoteProperty –name FRCS_Protocols –value n/a -ErrorAction SilentlyContinue
foreach ($P in $Ports) {
        if( $P.RemotePort -eq "80") {$P.FRCS_Protocols = "HTTP"}
        if( $P.RemotePort -eq "443") {$P.FRCS_Protocols = "HTTPS"}
        if( $P.RemotePort -eq "1089") {$P.FRCS_Protocols = "Foundation Fieldbus HSE"}
        if( $P.RemotePort -eq "1090") {$P.FRCS_Protocols = "Foundation Fieldbus HSE"}
        if( $P.RemotePort -eq "1091") {$P.FRCS_Protocols = "Foundation Fieldbus HSE"}
        if( $P.RemotePort -eq "1541") {$P.FRCS_Protocols = "Foxboro/Invensys Foxboro DCS Informix"}
        if( $P.RemotePort -eq "20000") {$P.FRCS_Protocols = "DNP3"}
        if( $P.RemotePort -eq "34962") {$P.FRCS_Protocols = "PROFINET"}
        if( $P.RemotePort -eq "34963") {$P.FRCS_Protocols = "PROFINET"}
        if( $P.RemotePort -eq "34964") {$P.FRCS_Protocols = "PROFINET"}
        if( $P.RemotePort -eq "44818") {$P.FRCS_Protocols = "EtherNet/IP"}
        if( $P.RemotePort -eq "45678") {$P.FRCS_Protocols = "Foxboro/Invensys Foxboro DCS AIMAPI"}
        if( $P.RemotePort -eq "55555") {$P.FRCS_Protocols = "Foxboro/Invensys Foxboro DCS FoxAPI"}
        if( $P.RemotePort -eq "3480") {$P.FRCS_Protocols = "OPC UA Discovery Server"}
        if( $P.RemotePort -eq "3480") {$P.FRCS_Protocols = "OPC UA Discovery Server"}
    if( $P.RemotePort -eq "102") {$P.FRCS_Protocols = "ICCP"}
    if( $P.RemotePort -eq "502") {$P.FRCS_Protocols = "Modbus TCP"}
if( $P.RemotePort -eq "3480") {$P.FRCS_Protocols = "OPC UA Discovery Server"}
if( $P.RemotePort -eq "4000") {$P.FRCS_Protocols = "Emerson/Fisher ROC Plus"}
if( $P.RemotePort -eq "4840") {$P.FRCS_Protocols = "OPC UA Discovery Server"}
if( $P.RemotePort -eq "5052") {$P.FRCS_Protocols = "Telvent OASyS DNA"}
if( $P.RemotePort -eq "5065") {$P.FRCS_Protocols = "Telvent OASyS DNA"}
if( $P.RemotePort -eq "5450") {$P.FRCS_Protocols = "OSIsoft PI Server"}
if( $P.RemotePort -eq "10307") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10311") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10364") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10365") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10407") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10409") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10410") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10412") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10414") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10415") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10428") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10431") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10432") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10447") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10449") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "10450") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "11001") {$P.FRCS_Protocols = "Johnson Controls Metasys N1"}
if( $P.RemotePort -eq "12135") {$P.FRCS_Protocols = "Telvent OASyS DNA"}
if( $P.RemotePort -eq "12137") {$P.FRCS_Protocols = "Telvent OASyS DNA"}
if( $P.RemotePort -eq "12316") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "12645") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "12647") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "12648") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "13722") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "13724") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "13782") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "13783") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "18000") {$P.FRCS_Protocols = "Iconic Genesis32 GenBroker (TCP)"}
if( $P.RemotePort -eq "38000") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38001") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38011") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38012") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38014") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38015") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38200") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38210") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38301") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38400") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38589") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "38593") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "38600") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "38700") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "38971") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "39129") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "39278") {$P.FRCS_Protocols = "ABB Ranger 2003"}
if( $P.RemotePort -eq "50001") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50002") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50003") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50004") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50005") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50006") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50007") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50008") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50009") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50010") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50011") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50012") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50013") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50014") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50015") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50016") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50018") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50019") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50025") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50026") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50027") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50028") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50110") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "50111") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
if( $P.RemotePort -eq "62900") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62911") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62924") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62930") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62938") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62956") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62957") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62963") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62981") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62982") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62985") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "62992") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63012") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63041") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63075") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63079") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63082") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63088") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "63094") {$P.FRCS_Protocols = "SNC GENe"}
if( $P.RemotePort -eq "65443") {$P.FRCS_Protocols = "SNC GENe"}

if( $P.RemotePort -ge "56001" -AND $P.RemotePort -le "56099" ) {$P.FRCS_Protocols = "Telvent OASyS DNA"}
if( $P.RemotePort -ge "63027" -AND $P.RemotePort -le "63036" ) {$P.FRCS_Protocols = "SNC GENe"}
    }


######################################################################
# Query Listening UDP Ports (No Connections in UDP)
 $netstat = netstat -a -n -o | findstr /R /C:"^  UDP"
        $UDPPorts = $netstat | ForEach-Object {
            $Parts = $_.Trim() -split "\s+"
            New-Object PSObject -Property @{
                Proto = $Parts[0]
                LocalAddress = $Parts[1].Substring(0,$Parts[1].LastIndexOf(':')) -replace '\[','' -replace '\]',''
                RemoteAddress = $Parts[2].Substring(0,$Parts[2].LastIndexOf(':')) -replace '\[','' -replace '\]',''
                State = $Parts[3]
                PID = $Parts[4]
                LocalPort = $Parts[1].Split(':')[-1]
                RemotePort = $Parts[2].Split(':')[-1]
                ProcessName = $Processes[[int]$Parts[4]].ProcessName
                }
             } | Select-Object LocalAddress,State,PID,RemotePort,
        @{Name="Proto";Expression={"UDP"}},
        LocalPort, ProcessName| Where-Object {$_.LocalAddress -notmatch "127.0.0.1|0.0.0.0|::1|::"} | Sort-Object -Property ProcessName

$UDPPorts | add-member –membertype NoteProperty –name FRCS_Protocols –value n/a -ErrorAction SilentlyContinue
foreach ($P in $UDPPorts) {
        if( $P.Proto -eq "UDP") {$P.RemotePort = $P.LocalPort}
        if( $P.RemotePort -eq "1089") {$P.FRCS_Protocols = "Foundation Fieldbus HSE"}
        if( $P.RemotePort -eq "1090") {$P.FRCS_Protocols = "Foundation Fieldbus HSE"}
        if( $P.RemotePort -eq "1091") {$P.FRCS_Protocols = "Foundation Fieldbus HSE"}
        if( $P.RemotePort -eq "1541") {$P.FRCS_Protocols = "Foxboro/Invensys Foxboro DCS Informix"}
        if( $P.RemotePort -eq "20000") {$P.FRCS_Protocols = "DNP3"}
        if( $P.RemotePort -eq "34962") {$P.FRCS_Protocols = "PROFINET"}
        if( $P.RemotePort -eq "34963") {$P.FRCS_Protocols = "PROFINET"}
        if( $P.RemotePort -eq "34964") {$P.FRCS_Protocols = "PROFINET"}
        if( $P.RemotePort -eq "44818") {$P.FRCS_Protocols = "EtherNet/IP"}
        if( $P.RemotePort -eq "45678") {$P.FRCS_Protocols = "Foxboro/Invensys Foxboro DCS AIMAPI"}
        if( $P.RemotePort -eq "55555") {$P.FRCS_Protocols = "Foxboro/Invensys Foxboro DCS FoxAPI"}
        if( $P.RemotePort -eq "2222") {$P.FRCS_Protocols = "EtherNet/IP"}
        if( $P.RemotePort -eq "5050") {$P.FRCS_Protocols = "Telvent OASyS DNA"}
        if( $P.RemotePort -eq "5051") {$P.FRCS_Protocols = "Telvent OASyS DNA"}
        if( $P.RemotePort -eq "34980") {$P.FRCS_Protocols = "EtherCAT"}
        if( $P.RemotePort -eq "47808") {$P.FRCS_Protocols = "BACnet/IP"}
        if( $P.RemotePort -eq "50020") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
        if( $P.RemotePort -eq "50021") {$P.FRCS_Protocols = "Siemens Spectrum Power TG"}
        if( $P.RemotePort -eq "55000") {$P.FRCS_Protocols = "FL-net Reception"}
        if( $P.RemotePort -eq "55001") {$P.FRCS_Protocols = "FL-net Reception"}
        if( $P.RemotePort -eq "55002") {$P.FRCS_Protocols = "FL-net Reception"}
        if( $P.RemotePort -eq "55003") {$P.FRCS_Protocols = "FL-net Transmission"}

            }

$Ports += $UDPPorts

$ShowPorts = $Ports | Select-Object LocalAddress,RemoteAddress,Proto,LocalPort,RemotePort,PID,ProcessName,FRCS_Protocols | Export-Csv -Path $CSVPPS -NoTypeInformation

Write-Host -ForegroundColor Green "Pulling Log Files: This takes quite a bit"
#Add Log Files
#Add Days should be -90 but shorten for testing
if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $EventID = '4624','4704','4740','1102','4946','6412','4672' #What Events We Care About
    $Logs = 'Application','Security','System','Windows PowerShell' #Logs to Search
    $Date = (Get-Date).AddDays(0)
    $AllLogResults = Get-WinEvent -WarningAction SilentlyContinue -FilterHashtable @{LogName=$Logs; StartTime=$Date; Level=1,2,3,4,0; ID=$EventID}
    $AllLogResults | Export-Csv -Path $CSVLogs -NoTypeInformation
    }

If ($RunSCAP -eq "1"){
    Write-Host -ForegroundColor Green "Running SCAP: Get some coffee"
    $ScapSaveLocation = "$ScanSaveDir\SCAP"
    $null = New-Item -ItemType Directory -Force -Path $ScapSaveLocation
    Start-Process -NoNewWindow  -Wait -Path $ScapLocation -ArgumentList "-u $ScapSaveLocation"
    Write-Host -ForegroundColor Green "SCAP is Finally Done!"
}

If ($RunSFC -eq "1"){
    Write-Host -ForegroundColor Green "Running SFC: Time for Coffee and maybe a nap. Note: SFC tends to take a long time at 22%"
    Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList '/scannow' -Wait -NoNewWindow
    Get-Content "C:\Windows\Logs\CBS\CBS.log" | Out-String | Out-File -FilePath $TXTSfc
}

If ($RunSFC -eq "2"){
    Write-Host -ForegroundColor Green "Running SFC: Time for Coffee and maybe a nap. Note: SFC tends to take a long time at 22%"
    Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList '/verifyonly' -Wait -NoNewWindow
    Get-Content "C:\Windows\Logs\CBS\CBS.log" | Out-String | Out-File -FilePath $TXTSfc
}

Write-Host -ForegroundColor Green "Fixing Permissions"
$AllItems=Get-ChildItem -Path $ScanSaveDir -Recurse -ErrorAction SilentlyContinue
foreach ($Item in $AllItems){
    $Acl = Get-Acl -ErrorAction SilentlyContinue -Path $Item.FullName
    $Acl.SetAccessRuleProtection($false,$true)
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","FullControl","Containerinherit,Objectinherit","none","Allow")
    $Acl.AddAccessRule($AccessRule)
    Set-Acl $Item.FullName $Acl -ErrorAction SilentlyContinue
    Write-Host $Item
}

Write-Host -ForegroundColor Green "Script Complete"
