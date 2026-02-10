#This is a modification of FRCS-HW_SW_Inventory_v1.05 by an unknown author.
#Author of Modification Dan B
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
V.96 Added GUI for SCAP/SFC/Log options with file picker for SCAP executable
V.97 Fixed PS 2.0 compatibility and permissions handling for files vs directories
V.98 Added additional ICS/SCADA protocol port detection
V.99 Added admin check, enhanced disk/network info, progress indicator, silent mode, summary report
V1.00 Enhanced non-admin mode: detailed skip/collect list, graceful degradation for SFC, GUI shows SFC as disabled when not admin
V1.01 Added -Help command-line flag with comprehensive usage documentation
V1.02 Fixed en-dash encoding bug: replaced Unicode en-dashes (U+2013) with ASCII hyphens on Add-Member calls
Known Working Systems:
Windows 7 Powershell 2.0
Windows 10 Powershell 5.0
Windows 11 Powershell 5.1
Windows Server 2012-2022
#>

#region Command-Line Parameters
# Silent mode parameters: -Silent -RunSCAP -ScapPath "path" -RunSFC [1|2] -LogDays 90
param(
    [switch]$Help,
    [switch]$Silent,
    [switch]$RunSCAPParam,
    [string]$ScapPathParam = "",
    [string]$RunSFCParam = "3",
    [int]$LogDaysParam = 90
)
#endregion

#region Help Display
if ($Help) {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "  MOAS - System Inventory and Audit Tool v1.02" -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Collects system inventory data including hardware, software,"
    Write-Host "  network configuration, local users, event logs, and optionally"
    Write-Host "  runs SCAP compliance and SFC scans."
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\MOAS.ps1                    # Interactive GUI mode"
    Write-Host "  .\MOAS.ps1 -Help              # Display this help message"
    Write-Host "  .\MOAS.ps1 -Silent [options]  # Silent/batch mode"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -Help              Display this help message and exit"
    Write-Host ""
    Write-Host "  -Silent            Run in silent mode (no GUI, no prompts)"
    Write-Host ""
    Write-Host "  -RunSCAPParam      Enable SCAP scan (use with -Silent)"
    Write-Host ""
    Write-Host "  -ScapPathParam     Path to SCAP executable (cscc.exe)"
    Write-Host "                     Default: searches script directory"
    Write-Host ""
    Write-Host "  -RunSFCParam       SFC scan mode (use with -Silent)"
    Write-Host "                     1 = SFC /SCANNOW"
    Write-Host "                     2 = SFC /VERIFYONLY"
    Write-Host "                     3 = Do not run SFC (default)"
    Write-Host ""
    Write-Host "  -LogDaysParam      Number of days of event logs to collect"
    Write-Host "                     Range: 1-365, Default: 90"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  # Interactive mode with GUI"
    Write-Host "  .\MOAS.ps1" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  # Silent mode with defaults (90 days logs, no SCAP/SFC)"
    Write-Host "  .\MOAS.ps1 -Silent" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  # Silent mode with SCAP scan"
    Write-Host "  .\MOAS.ps1 -Silent -RunSCAPParam -ScapPathParam 'C:\SCAP\cscc.exe'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  # Silent mode with SFC verify and 30 days of logs"
    Write-Host "  .\MOAS.ps1 -Silent -RunSFCParam 2 -LogDaysParam 30" -ForegroundColor Gray
    Write-Host ""
    Write-Host "OUTPUT:" -ForegroundColor Yellow
    Write-Host "  Creates a dated folder in the script directory containing:"
    Write-Host "    - BasicInfo-*.csv        System/hardware information"
    Write-Host "    - LocalUsers-*.csv       Local user accounts"
    Write-Host "    - UpdateandHotfixes-*.csv Installed updates"
    Write-Host "    - InstalledSoftware-*.csv Installed applications"
    Write-Host "    - PPS-*.csv              Network ports and processes"
    Write-Host "    - Logs-*.csv             Event log entries"
    Write-Host "    - SFC-*.txt              SFC scan results (if run)"
    Write-Host "    - SCAP\                  SCAP results folder (if run)"
    Write-Host ""
    Write-Host "REQUIREMENTS:" -ForegroundColor Yellow
    Write-Host "  - PowerShell 2.0 or later"
    Write-Host "  - Administrator privileges recommended for full functionality"
    Write-Host "  - Without admin: Security logs and SFC scans are skipped"
    Write-Host ""
    Write-Host "SUPPORTED SYSTEMS:" -ForegroundColor Yellow
    Write-Host "  Windows 7, 8, 10, 11"
    Write-Host "  Windows Server 2012, 2016, 2019, 2022"
    Write-Host ""
    exit
}
#endregion

#region Administrator Check
$isAdmin = (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host "  WARNING: Script is NOT running as Administrator" -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  WILL BE SKIPPED (Requires Administrator):" -ForegroundColor Red
    Write-Host "    [X] Security Event Log collection" -ForegroundColor Red
    Write-Host "    [X] SFC (System File Checker) scans" -ForegroundColor Red
    Write-Host ""
    Write-Host "  WILL STILL COLLECT (No Admin Required):" -ForegroundColor Green
    Write-Host "    [+] Basic System Information (Computer, BIOS, CPU, RAM)" -ForegroundColor Green
    Write-Host "    [+] Disk Information (All drives with free space)" -ForegroundColor Green
    Write-Host "    [+] Network Adapter Information (IP, MAC, Gateway, DNS)" -ForegroundColor Green
    Write-Host "    [+] Local User Accounts" -ForegroundColor Green
    Write-Host "    [+] Installed Updates and Hotfixes" -ForegroundColor Green
    Write-Host "    [+] Installed Software (Win32_Product)" -ForegroundColor Green
    Write-Host "    [+] Active Network Connections (TCP/UDP ports)" -ForegroundColor Green
    Write-Host "    [+] Application, System, and PowerShell Event Logs" -ForegroundColor Green
    Write-Host "    [+] Software License Information" -ForegroundColor Green
    Write-Host "    [+] SCAP Scan (if selected and tool permits)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  To run with full capabilities:" -ForegroundColor Cyan
    Write-Host "    Right-click PowerShell -> 'Run as Administrator'" -ForegroundColor Cyan
    Write-Host "    Then run this script again" -ForegroundColor Cyan
    Write-Host ""

    if (-not $Silent) {
        $continue = Read-Host "Continue with limited scan? (Y/N)"
        if ($continue -ne "Y" -and $continue -ne "y") {
            Write-Host "Exiting..." -ForegroundColor Red
            exit
        }
        Write-Host ""
        Write-Host "  Continuing with limited scan..." -ForegroundColor Yellow
    } else {
        Write-Host "  Silent mode: Continuing with limited scan..." -ForegroundColor Yellow
    }
    Write-Host ""
}
#endregion

# Load Windows Forms Assembly (compatible with PowerShell 2.0+)
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#region GUI Configuration Form
function Show-MOASConfigForm {
    param(
        [string]$DefaultScapPath = "",
        [bool]$IsAdministrator = $false
    )

    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "MOAS System Inventory Configuration"
    $form.Size = New-Object System.Drawing.Size(500, 380)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    # SCAP GroupBox
    $scapGroup = New-Object System.Windows.Forms.GroupBox
    $scapGroup.Text = "SCAP Configuration"
    $scapGroup.Location = New-Object System.Drawing.Point(10, 10)
    $scapGroup.Size = New-Object System.Drawing.Size(460, 100)
    $form.Controls.Add($scapGroup)

    # SCAP Checkbox
    $chkScap = New-Object System.Windows.Forms.CheckBox
    $chkScap.Text = "Run SCAP Scan"
    $chkScap.Location = New-Object System.Drawing.Point(15, 25)
    $chkScap.Size = New-Object System.Drawing.Size(150, 20)
    $scapGroup.Controls.Add($chkScap)

    # SCAP Path Label
    $lblScapPath = New-Object System.Windows.Forms.Label
    $lblScapPath.Text = "SCAP Executable (cscc.exe):"
    $lblScapPath.Location = New-Object System.Drawing.Point(15, 50)
    $lblScapPath.Size = New-Object System.Drawing.Size(200, 20)
    $scapGroup.Controls.Add($lblScapPath)

    # SCAP Path TextBox
    $txtScapPath = New-Object System.Windows.Forms.TextBox
    $txtScapPath.Location = New-Object System.Drawing.Point(15, 70)
    $txtScapPath.Size = New-Object System.Drawing.Size(340, 20)
    $txtScapPath.Text = $DefaultScapPath
    $txtScapPath.Enabled = $false
    $scapGroup.Controls.Add($txtScapPath)

    # SCAP Browse Button
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse..."
    $btnBrowse.Location = New-Object System.Drawing.Point(360, 68)
    $btnBrowse.Size = New-Object System.Drawing.Size(85, 25)
    $btnBrowse.Enabled = $false
    $scapGroup.Controls.Add($btnBrowse)

    # Enable/disable SCAP path controls based on checkbox
    $chkScap.Add_CheckedChanged({
        $txtScapPath.Enabled = $chkScap.Checked
        $btnBrowse.Enabled = $chkScap.Checked
    })

    # Browse button click handler
    $btnBrowse.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "Select SCAP Executable (cscc.exe)"
        $openFileDialog.Filter = "SCAP Executable (cscc.exe)|cscc.exe|All Executables (*.exe)|*.exe"
        $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        if ($openFileDialog.ShowDialog() -eq "OK") {
            $txtScapPath.Text = $openFileDialog.FileName
        }
    })

    # SFC GroupBox
    $sfcGroup = New-Object System.Windows.Forms.GroupBox
    $sfcGroup.Text = "System File Checker (SFC)"
    $sfcGroup.Location = New-Object System.Drawing.Point(10, 115)
    $sfcGroup.Size = New-Object System.Drawing.Size(460, 80)
    $form.Controls.Add($sfcGroup)

    # SFC Radio Buttons
    $rbSfcNone = New-Object System.Windows.Forms.RadioButton
    $rbSfcNone.Text = "Do not run SFC"
    $rbSfcNone.Location = New-Object System.Drawing.Point(15, 25)
    $rbSfcNone.Size = New-Object System.Drawing.Size(130, 20)
    $rbSfcNone.Checked = $true
    $sfcGroup.Controls.Add($rbSfcNone)

    $rbSfcScannow = New-Object System.Windows.Forms.RadioButton
    $rbSfcScannow.Text = "SFC /SCANNOW"
    $rbSfcScannow.Location = New-Object System.Drawing.Point(150, 25)
    $rbSfcScannow.Size = New-Object System.Drawing.Size(130, 20)
    $sfcGroup.Controls.Add($rbSfcScannow)

    $rbSfcVerify = New-Object System.Windows.Forms.RadioButton
    $rbSfcVerify.Text = "SFC /VERIFYONLY"
    $rbSfcVerify.Location = New-Object System.Drawing.Point(290, 25)
    $rbSfcVerify.Size = New-Object System.Drawing.Size(140, 20)
    $sfcGroup.Controls.Add($rbSfcVerify)

    # SFC Description Label
    $lblSfcDesc = New-Object System.Windows.Forms.Label
    $lblSfcDesc.Location = New-Object System.Drawing.Point(15, 50)
    $lblSfcDesc.Size = New-Object System.Drawing.Size(430, 20)

    # Disable SFC options if not running as Administrator
    if (-not $IsAdministrator) {
        $rbSfcNone.Enabled = $false
        $rbSfcScannow.Enabled = $false
        $rbSfcVerify.Enabled = $false
        $lblSfcDesc.Text = "SFC requires Administrator privileges (not available)"
        $lblSfcDesc.ForeColor = [System.Drawing.Color]::Red
    } else {
        $lblSfcDesc.Text = "Note: SFC scans can take a long time (especially at 22%)"
        $lblSfcDesc.ForeColor = [System.Drawing.Color]::Gray
    }
    $sfcGroup.Controls.Add($lblSfcDesc)

    # Log Collection GroupBox
    $logGroup = New-Object System.Windows.Forms.GroupBox
    $logGroup.Text = "Event Log Collection"
    $logGroup.Location = New-Object System.Drawing.Point(10, 200)
    $logGroup.Size = New-Object System.Drawing.Size(460, 80)
    $form.Controls.Add($logGroup)

    # Log Days Label
    $lblLogDays = New-Object System.Windows.Forms.Label
    $lblLogDays.Text = "Collect logs from the past (days):"
    $lblLogDays.Location = New-Object System.Drawing.Point(15, 30)
    $lblLogDays.Size = New-Object System.Drawing.Size(200, 20)
    $logGroup.Controls.Add($lblLogDays)

    # Log Days NumericUpDown (using TextBox for PS 2.0 compatibility)
    $txtLogDays = New-Object System.Windows.Forms.TextBox
    $txtLogDays.Location = New-Object System.Drawing.Point(220, 28)
    $txtLogDays.Size = New-Object System.Drawing.Size(60, 20)
    $txtLogDays.Text = "90"
    $txtLogDays.TextAlign = "Right"
    $logGroup.Controls.Add($txtLogDays)

    # Validate numeric input
    $txtLogDays.Add_KeyPress({
        param($sender, $e)
        if (-not [char]::IsDigit($e.KeyChar) -and $e.KeyChar -ne [char]8) {
            $e.Handled = $true
        }
    })

    # Log Days Description
    $lblLogDesc = New-Object System.Windows.Forms.Label
    $lblLogDesc.Text = "Events: Logon (4624), Audit Policy (4704), Lockout (4740), Log Clear (1102), etc."
    $lblLogDesc.Location = New-Object System.Drawing.Point(15, 55)
    $lblLogDesc.Size = New-Object System.Drawing.Size(430, 20)
    $lblLogDesc.ForeColor = [System.Drawing.Color]::Gray
    $logGroup.Controls.Add($lblLogDesc)

    # OK Button
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Start Scan"
    $btnOK.Location = New-Object System.Drawing.Point(290, 295)
    $btnOK.Size = New-Object System.Drawing.Size(85, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    # Cancel Button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(385, 295)
    $btnCancel.Size = New-Object System.Drawing.Size(85, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    # Show the form
    $result = $form.ShowDialog()

    # Determine SFC selection
    $sfcChoice = "3"  # Default: No
    if ($rbSfcScannow.Checked) { $sfcChoice = "1" }
    elseif ($rbSfcVerify.Checked) { $sfcChoice = "2" }

    # Validate log days
    $logDays = 90
    if ($txtLogDays.Text -match '^\d+$') {
        $logDays = [int]$txtLogDays.Text
        if ($logDays -lt 1) { $logDays = 1 }
        if ($logDays -gt 365) { $logDays = 365 }
    }

    # Return configuration as hashtable (PS 2.0 compatible)
    $runScapValue = "2"
    if ($chkScap.Checked) { $runScapValue = "1" }

    $configResult = @{
        DialogResult = $result
        RunSCAP = $runScapValue
        ScapPath = $txtScapPath.Text
        RunSFC = $sfcChoice
        LogDays = $logDays
    }
    return $configResult
}
#endregion

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

Write-Host -ForegroundColor Green "Creating the Save Folder"
#Creat the Save Directory for this scan
$ShortDate = Get-Date -Format "yyyyMMdd"
$ScanSaveDir = "$ScriptDir\$ShortDate-$env:COMPUTERNAME"
$null = New-Item -ItemType Directory -Force -Path $ScanSaveDir
# This is the output file
$CSVBasicInfo = "$ScanSaveDir\$env:COMPUTERNAME-BasicInfo-$now.csv"
$CSVUpdatesandHotfixes = "$ScanSaveDir\$env:COMPUTERNAME-UpdateandHotfixes-$now.csv"
$CSVLocalUserList = "$ScanSaveDir\$env:COMPUTERNAME-LocalUsers-$now.csv"
$CSVPPS = "$ScanSaveDir\$env:COMPUTERNAME-PPS-$now.csv"
$CSVLogs = "$ScanSaveDir\$env:COMPUTERNAME-Logs-$now.csv"
$TXTSfc = "$ScanSaveDir\$env:COMPUTERNAME-SFC-$now.txt"
$CSVInstalledSoftware = "$ScanSaveDir\$env:COMPUTERNAME-InstalledSoftware-$now.csv"

# Try to find SCAP in script directory as default
$DefaultScapPath = Get-ChildItem $ScriptDir -Filter cscc.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 | ForEach-Object { $_.FullName }
if (-not $DefaultScapPath) { $DefaultScapPath = "" }

# Initialize tracking variables for summary report
$script:CollectedItems = @()
$script:Warnings = @()
$script:StartTime = Get-Date

# Silent mode or GUI mode
if ($Silent) {
    Write-Host -ForegroundColor Cyan "Running in Silent Mode..."
    # Use command-line parameters
    $RunSCAP = "2"
    if ($RunSCAPParam) { $RunSCAP = "1" }
    $ScapLocation = $ScapPathParam
    if (-not $ScapLocation -and $RunSCAP -eq "1") { $ScapLocation = $DefaultScapPath }
    $RunSFC = $RunSFCParam
    $LogDays = $LogDaysParam
    if ($LogDays -lt 1) { $LogDays = 1 }
    if ($LogDays -gt 365) { $LogDays = 365 }
} else {
    # Show GUI Configuration Dialog
    Write-Host -ForegroundColor Green "Opening Configuration Dialog..."
    $Config = Show-MOASConfigForm -DefaultScapPath $DefaultScapPath -IsAdministrator $isAdmin

    # Check if user cancelled
    if ($Config.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host -ForegroundColor Yellow "Operation cancelled by user."
        exit
    }

    # Set variables from GUI
    $RunSCAP = $Config.RunSCAP
    $ScapLocation = $Config.ScapPath
    $RunSFC = $Config.RunSFC
    $LogDays = $Config.LogDays
}

# Display configuration (PS 2.0 compatible)
Write-Host -ForegroundColor Green "Configuration:"
$scapDisplay = "No"
if ($RunSCAP -eq "1") { $scapDisplay = "Yes" }
Write-Host "  Run SCAP: $scapDisplay"
if ($RunSCAP -eq "1") { Write-Host "  SCAP Path: $ScapLocation" }
$sfcDisplay = "No"
if ($RunSFC -eq "1") { $sfcDisplay = "SCANNOW" }
if ($RunSFC -eq "2") { $sfcDisplay = "VERIFYONLY" }
Write-Host "  Run SFC: $sfcDisplay"
Write-Host "  Log Days: $LogDays"

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

# Enhanced Disk Information - All drives with free space
Write-Host -ForegroundColor Green "Getting Disk Information"
$AllDisks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
foreach ($Disk in $AllDisks) {
    $diskSize = [math]::Round($Disk.Size / 1GB, 2)
    $diskFree = [math]::Round($Disk.FreeSpace / 1GB, 2)
    $diskUsedPercent = if ($Disk.Size -gt 0) { [math]::Round((($Disk.Size - $Disk.FreeSpace) / $Disk.Size) * 100, 1) } else { 0 }
    $BasicInfo += New-Object PSObject -Property @{Title="Disk $($Disk.DeviceID) Size (GB)"; Data=$diskSize}
    $BasicInfo += New-Object PSObject -Property @{Title="Disk $($Disk.DeviceID) Free (GB)"; Data=$diskFree}
    $BasicInfo += New-Object PSObject -Property @{Title="Disk $($Disk.DeviceID) Used (%)"; Data=$diskUsedPercent}
}
$script:CollectedItems += "Disk Information"

# Enhanced Network Information - All adapters
Write-Host -ForegroundColor Green "Getting Network Adapter Information"
$AllNetAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction SilentlyContinue
$adapterIndex = 1
foreach ($Adapter in $AllNetAdapters) {
    $BasicInfo += New-Object PSObject -Property @{Title="Network Adapter $adapterIndex Description"; Data=$Adapter.Description}
    if ($Adapter.IPAddress) {
        $ipList = $Adapter.IPAddress -join "; "
        $BasicInfo += New-Object PSObject -Property @{Title="Network Adapter $adapterIndex IP Address"; Data=$ipList}
    }
    if ($Adapter.MACAddress) {
        $BasicInfo += New-Object PSObject -Property @{Title="Network Adapter $adapterIndex MAC Address"; Data=$Adapter.MACAddress}
    }
    if ($Adapter.DefaultIPGateway) {
        $gwList = $Adapter.DefaultIPGateway -join "; "
        $BasicInfo += New-Object PSObject -Property @{Title="Network Adapter $adapterIndex Gateway"; Data=$gwList}
    }
    if ($Adapter.DNSServerSearchOrder) {
        $dnsList = $Adapter.DNSServerSearchOrder -join "; "
        $BasicInfo += New-Object PSObject -Property @{Title="Network Adapter $adapterIndex DNS Servers"; Data=$dnsList}
    }
    if ($Adapter.DHCPEnabled) {
        $dhcpStatus = "Enabled"
        if ($Adapter.DHCPServer) { $dhcpStatus = "Enabled (Server: $($Adapter.DHCPServer))" }
    } else {
        $dhcpStatus = "Disabled (Static IP)"
    }
    $BasicInfo += New-Object PSObject -Property @{Title="Network Adapter $adapterIndex DHCP"; Data=$dhcpStatus}
    $adapterIndex++
}
$script:CollectedItems += "Network Adapter Information"

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
$script:CollectedItems += "Basic System Information"

Write-Host -ForegroundColor Green "Getting Local Users"


#Add User List
$adsi = [ADSI]"WinNT://$env:COMPUTERNAME"
$LocalUsers = $adsi.Children | where {$_.SchemaClassName -eq 'user'} | Foreach-Object {
    $groups = $_.Groups() | Foreach-Object {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}
    $_ | Select-Object @{n='UserName';e={$_.Name}}, @{n='Description';e={$_.Description}}, @{n='Groups';e={$groups -join ';'}}
}
$LocalUsers | Export-Csv -Path $CSVLocalUserList -NoTypeInformation
$script:CollectedItems += "Local Users ($(@($LocalUsers).Count) users)"

Write-Host -ForegroundColor Green "Getting HotFixes"

# Updates and Hot Fixes
$InstalledUpdates = Get-HotFix -ErrorAction SilentlyContinue | Select-Object Description, HotFixID, Installedby, InstalledOn
$InstalledUpdates | Export-Csv -Path $CSVUpdatesandHotfixes -NoTypeInformation
$script:CollectedItems += "Installed Updates ($(@($InstalledUpdates).Count) hotfixes)"

#Software List
Write-Host -ForegroundColor Green "Getting Installed Software (this may take several minutes)..."
Write-Host -ForegroundColor Gray "  Querying Win32_Product - please wait..."

# Progress indicator for slow WMI query
$softwareJob = Start-Job -ScriptBlock {
    Get-WmiObject -Query "SELECT * FROM Win32_Product" | Select-Object Name, Vendor, Version
}

# Show progress while waiting (PS 2.0 compatible - dots instead of spinner)
Write-Host -NoNewline "  Processing"
$dotCount = 0
while ($softwareJob.State -eq 'Running') {
    Write-Host -NoNewline "."
    $dotCount++
    if ($dotCount -ge 60) {
        # Start a new line after 60 dots to prevent very long lines
        Write-Host ""
        Write-Host -NoNewline "  Processing"
        $dotCount = 0
    }
    Start-Sleep -Milliseconds 500
}
Write-Host " Done!"

$InstalledSoftware = Receive-Job -Job $softwareJob
Remove-Job -Job $softwareJob
$InstalledSoftware | Export-Csv -Path $CSVInstalledSoftware -NoTypeInformation
$script:CollectedItems += "Installed Software ($(@($InstalledSoftware).Count) items)"


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

$Ports | add-member -membertype NoteProperty -name FRCS_Protocols -value n/a -ErrorAction SilentlyContinue
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

# Additional ICS/SCADA Protocols
if( $P.RemotePort -eq "789") {$P.FRCS_Protocols = "Red Lion Crimson v3"}
if( $P.RemotePort -eq "1911") {$P.FRCS_Protocols = "Niagara Fox (Tridium)"}
if( $P.RemotePort -eq "1962") {$P.FRCS_Protocols = "PCWorx"}
if( $P.RemotePort -eq "2222") {$P.FRCS_Protocols = "EtherNet/IP"}
if( $P.RemotePort -eq "2404") {$P.FRCS_Protocols = "IEC 60870-5-104"}
if( $P.RemotePort -eq "2455") {$P.FRCS_Protocols = "WAGO I/O (PCWorx)"}
if( $P.RemotePort -eq "4712") {$P.FRCS_Protocols = "Siemens WinCC OA"}
if( $P.RemotePort -eq "4713") {$P.FRCS_Protocols = "Siemens WinCC OA"}
if( $P.RemotePort -eq "4911") {$P.FRCS_Protocols = "Niagara Fox SSL (Tridium)"}
if( $P.RemotePort -eq "5006") {$P.FRCS_Protocols = "Mitsubishi MELSEC-Q"}
if( $P.RemotePort -eq "5007") {$P.FRCS_Protocols = "Mitsubishi MELSEC-Q"}
if( $P.RemotePort -eq "5094") {$P.FRCS_Protocols = "HART-IP"}
if( $P.RemotePort -eq "5095") {$P.FRCS_Protocols = "HART-IP"}
if( $P.RemotePort -eq "9600") {$P.FRCS_Protocols = "OMRON FINS"}
if( $P.RemotePort -eq "18245") {$P.FRCS_Protocols = "GE SRTP"}
if( $P.RemotePort -eq "18246") {$P.FRCS_Protocols = "GE SRTP"}
if( $P.RemotePort -eq "19999") {$P.FRCS_Protocols = "DNP3"}
if( $P.RemotePort -eq "20256") {$P.FRCS_Protocols = "Unitronics PCOM"}
if( $P.RemotePort -eq "20547") {$P.FRCS_Protocols = "ProConOS (PCWorx)"}
if( $P.RemotePort -eq "41100") {$P.FRCS_Protocols = "Yokogawa CENTUM"}
if( $P.RemotePort -eq "44818") {$P.FRCS_Protocols = "EtherNet/IP CIP"}
if( $P.RemotePort -eq "48898") {$P.FRCS_Protocols = "Niagara Fox Secure"}
if( $P.RemotePort -eq "57176") {$P.FRCS_Protocols = "CODESYS Runtime"}

# Siemens S7 uses COTP/ISO-TSAP on port 102 (same as ICCP)
if( $P.RemotePort -eq "102") {$P.FRCS_Protocols = "ICCP/Siemens S7 COTP"}

# Honeywell Experion
if( $P.RemotePort -eq "51000") {$P.FRCS_Protocols = "Honeywell Experion PKS"}
if( $P.RemotePort -eq "51001") {$P.FRCS_Protocols = "Honeywell Experion PKS"}
if( $P.RemotePort -eq "51002") {$P.FRCS_Protocols = "Honeywell Experion PKS"}

# Schneider Electric
if( $P.RemotePort -eq "1541") {$P.FRCS_Protocols = "Foxboro/Schneider DCS"}
if( $P.RemotePort -eq "6000") {$P.FRCS_Protocols = "Schneider ClearSCADA"}
if( $P.RemotePort -eq "6543") {$P.FRCS_Protocols = "Schneider Modicon"}

# Rockwell/Allen-Bradley
if( $P.RemotePort -eq "2221") {$P.FRCS_Protocols = "Rockwell Allen-Bradley DF1"}
if( $P.RemotePort -eq "2223") {$P.FRCS_Protocols = "Rockwell Allen-Bradley EtherNet/IP"}
if( $P.RemotePort -eq "17185") {$P.FRCS_Protocols = "Rockwell RSLinx"}

# OPC Classic (DCOM)
if( $P.RemotePort -eq "135") {$P.FRCS_Protocols = "OPC Classic (DCOM RPC)"}

# IEC 61850
if( $P.RemotePort -eq "102") {$P.FRCS_Protocols = "IEC 61850 MMS/ICCP/S7"}

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

$UDPPorts | add-member -membertype NoteProperty -name FRCS_Protocols -value n/a -ErrorAction SilentlyContinue
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

        # Additional ICS/SCADA UDP Protocols
        if( $P.RemotePort -eq "69") {$P.FRCS_Protocols = "TFTP (ICS Firmware)"}
        if( $P.RemotePort -eq "161") {$P.FRCS_Protocols = "SNMP"}
        if( $P.RemotePort -eq "162") {$P.FRCS_Protocols = "SNMP Trap"}
        if( $P.RemotePort -eq "1911") {$P.FRCS_Protocols = "Niagara Fox (Tridium)"}
        if( $P.RemotePort -eq "4000") {$P.FRCS_Protocols = "Emerson/Fisher ROC Plus"}
        if( $P.RemotePort -eq "4911") {$P.FRCS_Protocols = "Niagara Fox SSL (Tridium)"}
        if( $P.RemotePort -eq "5094") {$P.FRCS_Protocols = "HART-IP"}
        if( $P.RemotePort -eq "5095") {$P.FRCS_Protocols = "HART-IP"}
        if( $P.RemotePort -eq "9600") {$P.FRCS_Protocols = "OMRON FINS"}
        if( $P.RemotePort -eq "18245") {$P.FRCS_Protocols = "GE SRTP"}
        if( $P.RemotePort -eq "18246") {$P.FRCS_Protocols = "GE SRTP"}
        if( $P.RemotePort -eq "19999") {$P.FRCS_Protocols = "DNP3"}
        if( $P.RemotePort -eq "41794") {$P.FRCS_Protocols = "Crestron (Building Automation)"}
        if( $P.RemotePort -eq "47809") {$P.FRCS_Protocols = "BACnet/IP Secure"}
        if( $P.RemotePort -eq "48898") {$P.FRCS_Protocols = "Niagara Fox Secure"}
        if( $P.RemotePort -eq "57176") {$P.FRCS_Protocols = "CODESYS Runtime"}

            }

$Ports += $UDPPorts

$ShowPorts = $Ports | Select-Object LocalAddress,RemoteAddress,Proto,LocalPort,RemotePort,PID,ProcessName,FRCS_Protocols | Export-Csv -Path $CSVPPS -NoTypeInformation
$script:CollectedItems += "Network Ports/Processes ($(@($Ports).Count) connections)"

Write-Host -ForegroundColor Green "Pulling Log Files: This takes quite a bit (collecting $LogDays days of logs)"
#Add Log Files
if ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $EventID = '4624','4625','4672','4704','4740','4946','6412','1102','10000','11010' #What Events We Care About
    $Logs = 'Application','Security','System','Windows PowerShell' #Logs to Search
    $Date = (Get-Date).AddDays(-$LogDays)
    $AllLogResults = Get-WinEvent -WarningAction SilentlyContinue -FilterHashtable @{LogName=$Logs; StartTime=$Date; Level=1,2,3,4,0; ID=$EventID}
    $AllLogResults | Export-Csv -Path $CSVLogs -NoTypeInformation
    $script:CollectedItems += "Event Logs ($(@($AllLogResults).Count) events, $LogDays days)"
    } else {
    Write-Host -ForegroundColor Yellow "Warning: Not running as Administrator - skipping Security log collection"
    $script:Warnings += "Security log collection skipped (requires Administrator)"
    }

If ($RunSCAP -eq "1"){
    if ($ScapLocation -and (Test-Path $ScapLocation)) {
        Write-Host -ForegroundColor Green "Running SCAP: Get some coffee"
        $ScapSaveLocation = "$ScanSaveDir\SCAP"
        $null = New-Item -ItemType Directory -Force -Path $ScapSaveLocation
        Start-Process -NoNewWindow -Wait -FilePath $ScapLocation -ArgumentList "-u $ScapSaveLocation"
        Write-Host -ForegroundColor Green "SCAP is Finally Done!"
        $script:CollectedItems += "SCAP Compliance Scan"
    } else {
        Write-Host -ForegroundColor Red "SCAP executable not found at: $ScapLocation"
        $script:Warnings += "SCAP scan skipped (executable not found)"
        Write-Host -ForegroundColor Red "Skipping SCAP scan."
    }
}

If ($RunSFC -eq "1"){
    if ($isAdmin) {
        Write-Host -ForegroundColor Green "Running SFC: Time for Coffee and maybe a nap. Note: SFC tends to take a long time at 22%"
        Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList '/scannow' -Wait -NoNewWindow
        Get-Content "C:\Windows\Logs\CBS\CBS.log" -ErrorAction SilentlyContinue | Out-String | Out-File -FilePath $TXTSfc
        $script:CollectedItems += "SFC Scan (SCANNOW)"
    } else {
        Write-Host -ForegroundColor Yellow "Skipping SFC SCANNOW - requires Administrator privileges"
        $script:Warnings += "SFC SCANNOW skipped (requires Administrator)"
    }
}

If ($RunSFC -eq "2"){
    if ($isAdmin) {
        Write-Host -ForegroundColor Green "Running SFC: Time for Coffee and maybe a nap. Note: SFC tends to take a long time at 22%"
        Start-Process -FilePath "${env:Windir}\System32\SFC.EXE" -ArgumentList '/verifyonly' -Wait -NoNewWindow
        Get-Content "C:\Windows\Logs\CBS\CBS.log" -ErrorAction SilentlyContinue | Out-String | Out-File -FilePath $TXTSfc
        $script:CollectedItems += "SFC Scan (VERIFYONLY)"
    } else {
        Write-Host -ForegroundColor Yellow "Skipping SFC VERIFYONLY - requires Administrator privileges"
        $script:Warnings += "SFC VERIFYONLY skipped (requires Administrator)"
    }
}

Write-Host -ForegroundColor Green "Fixing Permissions"
# Fix permissions on the scan directory and all contents
# Handle the root scan directory first
try {
    $Acl = Get-Acl -Path $ScanSaveDir -ErrorAction Stop
    $Acl.SetAccessRuleProtection($false, $true)
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $Acl.AddAccessRule($AccessRule)
    Set-Acl -Path $ScanSaveDir -AclObject $Acl -ErrorAction Stop
    Write-Host "  $ScanSaveDir"
} catch {
    Write-Host -ForegroundColor Yellow "  Warning: Could not set permissions on $ScanSaveDir"
}

# Process all items in the directory
$AllItems = Get-ChildItem -Path $ScanSaveDir -Recurse -ErrorAction SilentlyContinue
foreach ($Item in $AllItems) {
    try {
        $Acl = Get-Acl -Path $Item.FullName -ErrorAction Stop
        $Acl.SetAccessRuleProtection($false, $true)

        if ($Item.PSIsContainer) {
            # Directory: use inheritance flags
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        } else {
            # File: no inheritance flags needed
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "None", "None", "Allow")
        }

        $Acl.AddAccessRule($AccessRule)
        Set-Acl -Path $Item.FullName -AclObject $Acl -ErrorAction Stop
        Write-Host "  $($Item.FullName)"
    } catch {
        Write-Host -ForegroundColor Yellow "  Warning: Could not set permissions on $($Item.FullName)"
    }
}

#region Summary Report
$script:EndTime = Get-Date
$duration = $script:EndTime - $script:StartTime

Write-Host ""
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "             MOAS SCAN SUMMARY REPORT" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Computer Name:    $env:COMPUTERNAME" -ForegroundColor White
Write-Host "  Scan Date:        $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host "  Duration:         $([math]::Round($duration.TotalMinutes, 2)) minutes" -ForegroundColor White
Write-Host "  Output Directory: $ScanSaveDir" -ForegroundColor White
Write-Host ""

Write-Host "  Data Collected:" -ForegroundColor Green
foreach ($item in $script:CollectedItems) {
    Write-Host "    [+] $item" -ForegroundColor Green
}
Write-Host ""

# List output files
Write-Host "  Output Files:" -ForegroundColor Cyan
$outputFiles = Get-ChildItem -Path $ScanSaveDir -File -ErrorAction SilentlyContinue
foreach ($file in $outputFiles) {
    $sizeKB = [math]::Round($file.Length / 1KB, 1)
    Write-Host "    - $($file.Name) ($sizeKB KB)" -ForegroundColor White
}
Write-Host ""

# Show warnings if any
if ($script:Warnings.Count -gt 0) {
    Write-Host "  Warnings:" -ForegroundColor Yellow
    foreach ($warning in $script:Warnings) {
        Write-Host "    [!] $warning" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Admin status reminder
if (-not $isAdmin) {
    Write-Host "  Note: Script ran without Administrator privileges." -ForegroundColor Yellow
    Write-Host "        Some data may be incomplete (Security logs, SFC)." -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "                    SCAN COMPLETE" -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""
#endregion
