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

References:

  https://techtalk.gfi.com/11-most-useful-powershell-commands-for-remote-management/
  
  https://social.technet.microsoft.com/wiki/contents/articles/7703.powershell-running-executables.aspx
	
  https://ss64.com/ps/syntax-compare.html
	
  https://blog.netspi.com/powershell-remoting-cheatsheet/
	
  https://nircmd.nirsoft.net/

This script was created using Notepad++.
