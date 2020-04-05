# IT_service_desk_masterkey_v2

Created by:     reurbo

Created on:     2019.06.01

Last updated:   2019.10.29

Script name:		IT_service_desk_masterkey_v2.ps1 (aka slim.ps1)

Description:		An updated script from my previous IT_service_desk_masterkey_v1.bat to include new features and potentially speed up previously slower tasks from the last version. Some features require nircmd.exe (listed in the References) to do certain tasks. I have the file set to be copied from a specific folder to the target computer for those tasks, run what is needed, then remove the file to clean up after itself.

NOTICE:			I use PowerShell ISE for this and you can hide the script pane using Ctrl+R. Be aware, if you are getting an error you aren't allowed to run scripts in PowerShell you'll need to do the following:

    1. Open an Administrator instance of PowerShell ISE
    
    2. Enter the following: Set-ExecutionPolicy -ExecutionPolicy Unrestricted
    
    3. When prompted, choose "Yes to All"

You can then close out of the Administrator PowerShell ISE, open a regular PowerShell ISE, open the script and run it.


#List of features in this script:
1. Remotely clear the cache and temporary files on another PC
2. Remotely reboot or shutdown another PC
3. Remotely change the volume or clock time zone of another PC
4. Remotely restart the printer spooler and clear the print queue of another PC
5. Remotely kill a process of another PC
6. Remotely make a backup of all files in specific folders on a PC to their network drive in a unquie folder
7. Send a custom pop-up message to a remote PC (does not 
8. Get general or detailed system information of a remote PC including (where possible):
    1. Current logged in user
    2. OS info (operating system, serial number, architecture)
    3. System info (PC name, domain, manufacturer, model, system type)
    4. Current time, last boot up time, and system uptime
    5. Disk space (lists all local drives, their size, and free space)
    6. Memory info (bank label, capacity, data width, and device locator)
    7. PC serial number
    8. PC printer info
    9. Add/Remove Program list
    10. Process list
    11. Service list
    12. USB devices
    13. CPU info (name, make/model, manufacturer, processor ID, number of cores, architecture)
    14. Monitor serial numbers (note: all-in-one computers will not list any monitors)

References:

  https://techtalk.gfi.com/11-most-useful-powershell-commands-for-remote-management/
  
  https://social.technet.microsoft.com/wiki/contents/articles/7703.powershell-running-executables.aspx
	
  https://ss64.com/ps/syntax-compare.html
	
  https://blog.netspi.com/powershell-remoting-cheatsheet/
	
  https://nircmd.nirsoft.net/

This script was created using Notepad++.
