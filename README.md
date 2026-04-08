# SLIM Tool (Service Line Interface Manager)

**Created by:** Brad Arrowood

**Originally created:** 2019

**Last updated:** 2026.04.08

**Script name:** `slim.ps1` (formerly `IT_service_desk_masterkey_v2.ps1`)

## Overview

SLIM Tool is a PowerShell-based remote administration toolkit originally built for Tier 2 desktop support. It provides a text-based menu interface for remotely managing Windows PCs, POS registers, and back office computers across retail store environments.

The script evolved from a collection of `.bat` and `.cmd` files into a full-featured PowerShell tool with nested menus, automatic device-type detection (POS vs. standard PC), and functions for common support tasks -- all designed so that non-technical team members could use it by simply entering a computer name or store number.

## Requirements

- **PowerShell 5.1+** (uses CIM cmdlets, `Get-CimInstance`, `Invoke-CimMethod`, etc.)
- **Active Directory module** (for account management features)
- **Network access** to target computers via WinRM / administrative shares (`\\computer\c$`)
- **nircmd.exe** (required for volume control features only; the script copies it to the target machine, runs the command, then removes it automatically)

### Execution Policy

If you receive an error about script execution being disabled, run the following in an **Administrator** PowerShell session:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
```

When prompted, choose **Yes to All**. You can then run the script from a regular PowerShell session.

## Main Menu

| Option | Description |
|--------|-------------|
| 1 | Connect to a remote computer (auto-detects POS vs. standard PC) |
| 2 | Employee account info (lookup, unlock, password reset) |
| 3 | I.T. service desk employee contact list |
| 4 | List all devices at a store (by 4-digit store number) |
| 5 | Production Control monitoring dashboard |

## Remote PC Menu (Standard PCs)

After connecting to a non-POS computer, the following options are available:

| Option | Description |
|--------|-------------|
| 1 | System overview (general) -- user, OS, system info, uptime, disk, memory |
| 2 | System overview (detailed) -- 14 individual reports (see below) |
| 3 | Clear cache and temporary files |
| 4 | Reboot PC |
| 5 | Shutdown PC |
| 6 | Change volume (requires nircmd.exe) |
| 7 | Change clock time zone (10 US/Canada zones supported) |
| 8 | Restart printer spooler and clear print queue |
| 9 | Speed up a PC (clear prefetch and temp files) |
| 10 | Send custom pop-up message to PC |
| 11 | Kill a running process (interactive process list with PID selection) |
| 12 | Kill BlockKeyboard/BlockMouse processes |
| 14 | Backup Desktop and/or Documents to network drive |

### Detailed System Reports (Option 2)

| Report | Details |
|--------|---------|
| PC Serial Number | BIOS serial number |
| Printer Info | Device ID, driver, port |
| Current User | Currently logged-in user |
| OS Info | OS name, serial number, architecture |
| System Info | Name, domain, manufacturer, model, system type |
| Add/Remove Programs | Full installed software list |
| Process List | Running processes with PIDs |
| Service List | All services with state, start mode, and exit codes |
| USB Devices | Connected USB device details |
| Uptime | Current time, last boot, and calculated uptime |
| Disk Space | All drives with size and free space |
| Memory Info | Bank label, capacity, data width per DIMM |
| Processor Info | CPU name, manufacturer, cores, architecture |
| Monitor Serial Numbers | Pulled from EDID registry data (not available on all-in-ones) |

## POS Register Menu

When connecting to a device with a hostname starting with `POS-` or `LAB-`, a POS-specific menu is presented:

| Option | Description |
|--------|-------------|
| 1 | System overview (general) |
| 2 | Stop POS services (Epicor, Secure Data, Store, SQL) |
| 3 | Restart POS services (full stop/start cycle with prompts) |
| 4 | Resync barcode scanner (restarts scanner-related POS services) |
| 5 | Reboot PC (stops POS services first, then reboots) |
| 6 | Shutdown PC (stops POS services first, then shuts down) |
| 7 | Change volume (requires nircmd.exe) |
| 8 | Change clock time zone |
| 9 | Restart printer spooler and clear print queue |
| 10 | Stop only POS Client and Shell |
| 11 | Send custom pop-up message |

## Employee Account Management

| Feature | Description |
|---------|-------------|
| Account lookup | Displays full name, comment, account status, and password info |
| Account unlock | Unlocks a locked Active Directory account |
| Password reset | Resets an AD account password (minimum 7 characters) |

## Production Control

A monitoring dashboard that polls network file share directories on a configurable timer (1-24 hours). It tracks:

- **TokenData** folder -- most recent file and timestamp
- **PollData** folders (3 monitored) -- most recent files and timestamps
- **BulkInventoryLoad** -- file count and individual file sizes

If timestamps between monitored folders diverge by more than 54 minutes, a Windows message box alert is triggered and Explorer windows are opened to the affected directories.

## How It Works

1. Launch the script -- the Main Menu is displayed
2. Choose **Connect to a computer** and enter a hostname or IP address
3. The script pings the target to verify it's online
4. Based on the hostname prefix (`POS-`, `LAB-`, or other), the appropriate menu is loaded
5. Select an action -- the script executes it remotely and returns to the menu

All remote operations use standard Windows management tools: CIM (WMI successor), PowerShell remoting (`Invoke-Command`), administrative shares, and `psexec` for certain nircmd operations.

## Version History

- **v1** -- `IT_service_desk_masterkey_v1.bat` -- Original batch file collection
- **v2** -- `slim.ps1` -- Full rewrite in PowerShell with menu system, functions, POS support, and expanded feature set
- **2026.04** -- Modernized: deprecated WMI cmdlets replaced with CIM equivalents, aliases expanded to full cmdlet names, encoding issues from migration corrected

## Known Linter Warnings (Safe to Ignore)

If you open this script in VS Code or another editor with PSScriptAnalyzer, you will see approximately 15 warnings stating that various variables are "assigned but never used." These are **false positives** and the script functions correctly as-is.

This happens because the script uses PowerShell's **scope inheritance** -- a variable assigned in one function is readable by any function called from within it. For example, `$MenuSelection` is assigned in `GetMenu` and then read by `GetInfo` (which is called from inside `GetMenu`). The linter analyzes each function in isolation and doesn't follow the call chain, so it reports the variable as unused within the function where it's assigned.

This is an intentional design pattern used throughout the script for variables like `$compname`, `$pcip`, `$MenuSelection`, `$TZSelection`, `$BACKUPSelection`, and others. No changes are needed.

## References

- [PowerShell Remote Management Commands](https://techtalk.gfi.com/11-most-useful-powershell-commands-for-remote-management/)
- [PowerShell Running Executables](https://social.technet.microsoft.com/wiki/contents/articles/7703.powershell-running-executables.aspx)
- [PowerShell Comparison Operators](https://ss64.com/ps/syntax-compare.html)
- [PowerShell Remoting Cheatsheet](https://blog.netspi.com/powershell-remoting-cheatsheet/)
- [NirCmd - NirSoft](https://nircmd.nirsoft.net/)
