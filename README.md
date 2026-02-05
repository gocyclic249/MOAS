# MOAS - System Inventory and Audit Tool

A PowerShell-based system inventory and audit tool for Windows environments. Originally based on FRCS-HW_SW_Inventory, extensively modified and enhanced for modern use.

**Version:** 1.02
**Author:** Dan B
**License:** GPL 2.0

## Features

- **System Information Collection**
  - Computer manufacturer, model, and serial number
  - BIOS information
  - CPU and RAM details
  - All disk drives with free space and usage percentage
  - Software license information

- **Network Information**
  - All network adapters with IP, MAC, gateway, and DNS
  - DHCP status for each adapter
  - Active TCP/UDP connections with process information
  - ICS/SCADA protocol detection (Modbus, DNP3, EtherNet/IP, BACnet, OPC UA, and 60+ more)

- **User and Security**
  - Local user accounts with group memberships
  - Event log collection (configurable date range)
  - Security event monitoring (logon, audit policy, lockouts, log clearing)

- **Software Inventory**
  - Installed updates and hotfixes
  - Installed software (Win32_Product)

- **Optional Scans**
  - SCAP compliance scanning (requires SCAP tool)
  - SFC (System File Checker) scans

## Supported Systems

- Windows 7 (PowerShell 2.0)
- Windows 8/8.1
- Windows 10 (PowerShell 5.0)
- Windows 11 (PowerShell 5.1)
- Windows Server 2012, 2012 R2, 2016, 2019, 2022

## Installation

1. Download `MOAS.ps1` to your desired location
2. (Optional) Place your SCAP executable (cscc.exe) in the same directory or a subdirectory

## Usage

### Interactive Mode (GUI)

Simply run the script to open the configuration dialog:

```powershell
.\MOAS.ps1
```

The GUI allows you to:
- Enable/disable SCAP scanning with file browser
- Select SFC scan mode (SCANNOW, VERIFYONLY, or none)
- Configure how many days of event logs to collect (default: 90)

### Silent/Batch Mode

Run without GUI prompts for automation:

```powershell
# Basic silent mode with defaults
.\MOAS.ps1 -Silent

# With SCAP scan
.\MOAS.ps1 -Silent -RunSCAPParam -ScapPathParam "C:\SCAP\cscc.exe"

# With SFC verify and 30 days of logs
.\MOAS.ps1 -Silent -RunSFCParam 2 -LogDaysParam 30

# Full example with all options
.\MOAS.ps1 -Silent -RunSCAPParam -ScapPathParam "C:\SCAP\cscc.exe" -RunSFCParam 1 -LogDaysParam 60
```

### Help

Display command-line help:

```powershell
.\MOAS.ps1 -Help
```

## Command-Line Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-Help` | Display help message and exit | - |
| `-Silent` | Run in silent mode (no GUI, no prompts) | Off |
| `-RunSCAPParam` | Enable SCAP scan | Off |
| `-ScapPathParam` | Path to SCAP executable (cscc.exe) | Auto-detect |
| `-RunSFCParam` | SFC mode: 1=SCANNOW, 2=VERIFYONLY, 3=None | 3 |
| `-LogDaysParam` | Days of event logs to collect (1-365) | 90 |

## Output

The script creates a dated folder in the script directory (e.g., `20260130-COMPUTERNAME`) containing:

| File | Description |
|------|-------------|
| `BasicInfo-*.csv` | System hardware and configuration |
| `LocalUsers-*.csv` | Local user accounts and groups |
| `UpdateandHotfixes-*.csv` | Installed Windows updates |
| `InstalledSoftware-*.csv` | Installed applications |
| `PPS-*.csv` | Network ports, processes, and protocols |
| `Logs-*.csv` | Event log entries |
| `SFC-*.txt` | SFC scan results (if run) |
| `SCAP\` | SCAP compliance results (if run) |

## Administrator Privileges

The script works with or without Administrator privileges:

### With Administrator
- Full functionality
- Security event log collection
- SFC scans available

### Without Administrator
The script will still collect:
- Basic system information
- Disk and network information
- Local user accounts
- Installed updates and software
- Network connections
- Application, System, and PowerShell event logs
- Software license information

**Skipped without admin:**
- Security event log collection
- SFC scans

## ICS/SCADA Protocol Detection

The script identifies industrial control system protocols on network connections, including:

- **Fieldbus:** Foundation Fieldbus HSE, PROFINET, EtherCAT
- **Industrial Ethernet:** EtherNet/IP, Modbus TCP, OPC UA
- **Building Automation:** BACnet/IP, Niagara Fox, Johnson Controls Metasys
- **SCADA:** DNP3, IEC 60870-5-104, IEC 61850 MMS
- **Vendor-Specific:** Siemens S7/WinCC, Honeywell Experion, ABB Ranger, Rockwell/Allen-Bradley, Schneider Electric, GE SRTP, OMRON FINS, Mitsubishi MELSEC, Yokogawa CENTUM, and many more

## Version History

| Version | Changes |
|---------|---------|
| 1.02 | Fixed en-dash encoding bug that caused parsing errors on some systems |
| 1.01 | Added -Help command-line flag |
| 1.00 | Enhanced non-admin mode with detailed skip/collect list |
| 0.99 | Admin check, enhanced disk/network info, progress indicator, silent mode, summary |
| 0.98 | Added 60+ ICS/SCADA protocol port detection |
| 0.97 | Fixed PS 2.0 compatibility and permissions handling |
| 0.96 | Added GUI for SCAP/SFC/Log configuration |
| 0.95 | Hardware output to BasicInfo, separate InstalledSoftware CSV |

## Troubleshooting

### Script won't run
Ensure your PowerShell execution policy allows scripts:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### SFC options grayed out in GUI
SFC requires Administrator privileges. Run PowerShell as Administrator.

### Missing Security logs
Security event logs require Administrator privileges. The script will collect Application, System, and PowerShell logs without admin.

### SCAP scan not running
Verify the path to cscc.exe is correct and the file exists.

### Parsing errors or "Expressions are only allowed as the first element of a pipeline"

This is typically caused by file encoding issues. PowerShell 5.1 defaults to the system's ANSI encoding (usually Windows-1252) when a script file lacks a UTF-8 BOM (Byte Order Mark). If the file contains any non-ASCII characters (such as Unicode en-dashes, smart quotes, or accented characters from copy-paste), they can be misinterpreted and corrupt the parser state.

**To check and fix encoding:**

1. Open the script in Notepad (or Notepad++)
2. Go to **File > Save As**
3. In the **Encoding** dropdown, select **UTF-8 with BOM** (or **UTF-8-BOM** in Notepad++)
4. Save and overwrite the file

**To verify encoding in PowerShell:**
```powershell
# Check the first bytes of the file for a UTF-8 BOM (should be EF BB BF)
Format-Hex .\MOAS.ps1 | Select-Object -First 1
```

**Prevention tips:**
- When editing in any text editor, always save as **UTF-8 with BOM** for PowerShell scripts
- Avoid copy-pasting code from web pages or Word documents, which can introduce smart quotes (`"` `"`) and en-dashes (`–`) that look like regular characters but are different Unicode code points
- If transferring the script between systems, use Git or binary-safe file transfer to preserve encoding

## License

This project is released under the GPL 2.0 License.
