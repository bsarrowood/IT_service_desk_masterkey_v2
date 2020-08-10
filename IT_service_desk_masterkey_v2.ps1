# Created by:		Brad Arrowood
# Created on:		2019.06.01
# Last updated:		2019.10.29
# Script name:		IT_service_desk_masterkey_v2.ps1 (aka slim.ps1)
# Description:		An updated script from my previous IT_service_desk_masterkey_v1.bat to include
#			new features and potentially speed up previously slower tasks from the last version.
#			Some features require nircmd.exe (listed in the References) to do certain tasks.
#			I have the file set to be copied from a specific folder to the target computer for
#			those tasks, run what is needed, then remove the file to clean up after itself.

# NOTICE:		I use PowerShell ISE for this and you can hide the script pane using Ctrl+R.
#			Be aware, if you are getting an error you aren't allowed to run scripts in PowerShell
#			you'll need to do the following:
#				1. Open an Administrator instance of PowerShell ISE
#				2. Enter the following: Set-ExecutionPolicy -ExecutionPolicy Unrestricted
#				3. When prompted, choose "Yes to All"
#			You can then close out of the Administrator PowerShell ISE, open a regular 
#			PowerShell ISE, open the script and run it.

# References:
#	https://techtalk.gfi.com/11-most-useful-powershell-commands-for-remote-management/
#	https://social.technet.microsoft.com/wiki/contents/articles/7703.powershell-running-executables.aspx
#	https://ss64.com/ps/syntax-compare.html  <-- IF statement comparision operations list
#	https://blog.netspi.com/powershell-remoting-cheatsheet/
#	https://nircmd.nirsoft.net/

function iniStart {
	#each time the script runs, this is called to clear all variables beforhand... less $compname
    $addHours = ''
    $BACKUPSelection = ''	
    $checkstring = ''
    $connection = ''
    $device = ''
	$DID = ''
	$DID_KEY_NAME = ''
    $dirDateTime = ''
    $dirNameDateTime = ''
	$EDID = ''
	$EDID_String = ''
    $ErrorActionPreference = ''
    $ErrorActionPreference_Backup = ''
    $filesBULK_sizes = ''
    $fileTOKEN = ''
    $folderBULKcount = ''
    $folderPOLL1 = ''
    $folderPOLL1_str = ''
    $folderPOLL2 = ''
    $folderPOLL2_str = ''
    $folderPOLL3 = ''
    $folderPOLL3_str = ''
    $folderPOLL4 = ''
    $folderPOLL4_str = ''
    $folderTOKEN = ''
	$HID = ''
	$HID_KEY_NAME = ''
	$int = ''
    $keytype = ''
    $lastbootuptime = ''
    $localdatetime = ''
	$LOGGEDUSER1 = ''
    $LOGGEDUSER2 = ''
	$matches = ''
	$matchfound = ''
    $MenuSelection = ''
    $MenuSelectionDeatiled = ''
    $MODDEDUSER = ''
    $moncol1 = ''
    $moncol2 = ''
    $moncol3 = ''
    $moncol4 = ''
	$monrow = ''
	$montable = ''
    $msg = ''
    $msgBoxInput = ''
    $nameLength = ''
    $null = ''
    $numLength = ''
    $pathBULK = ''
    $pathFROMStore1 = ''
    $pathFROMStore2 = ''
    $pathname1 = ''
    $pathname2 = ''
    $pathname3 = ''
    $pathname4 = ''
    $pathname5 = ''
    $pathname6 = ''
    $pathname7 = ''
    $pathname8 = ''
    $pathPOLL1 = ''
    $pathPOLL2 = ''
    $pathPOLL3 = ''
    $pathPOLL4 = ''
    $pathTOKEN = ''
    $pathTOServer = ''
    $pathUSERS1 = ''
    $pathUSERS2 = ''
    $PIDSelection = ''
    $pcip = ''
    $ping = ''
    $recentfileTOKEN = ''
    $recentfilePOLL1 = ''
    $recentfilePOLL2 = ''
    $recentfilePOLL3 = ''
    $recentfilePOLL4 = ''
    $recentfileTOKENcheck = ''
    $recentfilePOLL1check = ''
    $recentfilePOLL2check = ''
    $recentfilePOLL3check = ''
    $recentfilePOLL4check = ''
	$reg = ''
	$regKey = ''
    $pathFROMEmpDesktop1 = ''
	$pathFROMEmpDesktop2 = ''
    $pathFROMEmpDocuments1 = ''
	$pathFROMEmpDocuments2 = ''
    $pathFROMEmpDownloads1 = ''
	$pathFROMEmpDownloads2 = ''
    $pathFROMEmpPictures1 = ''
	$pathFROMEmpPictures2 = ''
    $size = ''
    $storeNum = ''
    $timeEnd = ''
    $TimeNow = ''
    $timespan_TokenPoll1 = ''
    $timespan_TokenPoll3 = ''
    $timespan_Poll1Poll3 = ''
    $timeStart = ''
    $TZSelection = ''
    $uptime = ''
    $wmi = ''            
}	

function funcPause { 
	"" 
	Read-Host -Prompt "Press Enter to continue" 
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 
}

function GetCompName{
    #every time a new pc name is input, all variables are cleared
    iniStart #everytime a new pc is chosen all other variables are cleared

	Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
	""
    $pcname = Read-Host "Please enter a computer name or IP" 
    $compname = $pcname.ToUpper()
    if ($compname -eq "") {
        MainMenu
    }
    else {
        CheckHost
    }
}

function CheckHost{ 
    #this checks if the compname is online before loading the menu options
    $compnameUPPER = $compname.ToUpper()

	$ping = gwmi Win32_PingStatus -filter "Address='$compname'" 
	if($ping.StatusCode -eq 0){$pcip=$ping.ProtocolAddress; GetMenu} 
	else{Read-Host -Prompt "Host $compname is offline...Press any key to continue"; GetCompName} 
}

function GetMenu { 
    iniStart #everytime the remote pc menu loads all variables other than compname are cleared
    Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
    "  $compname ($pcip)" 
    "" 
    "1)  System overview (General)" 
    "2)  System overview (Detailed)" 
    "3)  Clear cache and temporary files" 
    "4)  Reboot PC" 
    "5)  Shutdown PC" 
    "6)  Change volume"
    "7)  Change clock time zone"
    "8)  Restart printer spooler and clear print queue"
    "9)  Speed up a PC (i.e. HP 6005)"
    "10) Send custom message to PC"
    "11) Kill a running process"
    "12) Kill any 'Block' processes (CWPC and SFPC Only)"
    "13) Make a backup of everything on the Desktop and/or My Documents to their network drive"
    "" 
    "C)  Switch to a different computer" 
    "X)  Exit to the Main Menu" 
    "" 
    $MenuSelection = Read-Host "Enter Selection" 
    GetInfo 
} 

function GetInfo {
	Clear-Host 
    switch ($MenuSelection){ 
		1 {
            #general overview of a pc
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			
			#Current User
			gwmi -computer $compname Win32_ComputerSystem | Format-Table @{Expression={$_.Username};Label="Current User"} 
			"------------------------------"
			#OS Info
			gwmi -computer $compname Win32_OperatingSystem | Format-List @{Expression={$_.Caption};Label="OS Name"},SerialNumber,OSArchitecture 
			"------------------------------"
			#System Info
			"System Info:"
			gwmi -computer $compname Win32_ComputerSystem | Format-List Name,Domain,Manufacturer,Model,SystemType 
			"------------------------------"
			#Uptime
            ""
			$wmi = gwmi -computer $compname Win32_OperatingSystem 
            $localdatetime = $wmi.ConvertToDateTime($wmi.LocalDateTime) 
            $lastbootuptime = $wmi.ConvertToDateTime($wmi.LastBootUpTime) 
            "Current Time:      $localdatetime" 
            "Last Boot Up Time: $lastbootuptime" 
            $uptime = $localdatetime - $lastbootuptime 
            "Uptime: $uptime" 
            ""
			"------------------------------"
			#Disk Space
            ""
			"Disk Space:"
			$wmi = gwmi -computer $compname Win32_logicaldisk 
            foreach($device in $wmi){ 
                    Write-Host "Drive: " $device.name    
                    Write-Host -NoNewLine "Size: "; "{0:N2}" -f ($device.Size/1Gb) + " Gb" 
                    Write-Host -NoNewLine "FreeSpace: "; "{0:N2}" -f ($device.FreeSpace/1Gb) + " Gb" 
                    "" 
             } 
			"------------------------------"
			#Memory Info
			"Memory Info:"
			$wmi = gwmi -computer $compname Win32_PhysicalMemory 
            foreach($device in $wmi){ 
                Write-Host "Bank Label:     " $device.BankLabel 
                Write-Host "Capacity:       " ($device.Capacity/1MB) "Mb" 
                Write-Host "Data Width:     " $device.DataWidth 
                Write-Host "Device Locator: " $device.DeviceLocator     
                ""         
            } 
			#"------------------------------"
			#Monitor Serial Number(s)
            #removed from general sys info report but leaving in detailed options
            ""
            Read-Host -Prompt "Press Enter to continue" 
            CheckHost 
		}
        2 {
            #detailed overview of a pc

            function CheckHostDetailed{ 
                $ping = gwmi Win32_PingStatus -filter "Address='$compname'" 
                if($ping.StatusCode -eq 0){$pcip=$ping.ProtocolAddress; DetailedMenu} 
                else{Read-Host -Prompt "Host $compname is offline...Press any key to continue"; GetCompName} 
            }

            function DetailedMenu {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
			    ""
                "Detailed System Overview Options:"
                ""
                "1)  PC Serial Number" 
                "2)  PC Printer Info" 
                "3)  Current User" 
                "4)  OS Info" 
                "5)  System Info" 
                "6)  Add/Remove Program List" 
                "7)  Process List" 
                "8)  Service List" 
                "9)  USB Devices" 
                "10) Uptime" 
                "11) Disk Space" 
                "12) Memory Info" 
                "13) Processor Info" 
                "14) Monitor Serial Numbers" 
                "" 
                "X)  Cancel" 
                "" 
                $MenuSelectionDeatiled = Read-Host "Enter Selection" 
                GetInfoDetailed
            }

            function GetInfoDetailed { 
                Clear-Host 
                switch ($MenuSelectionDeatiled){ 
                    1 { #PC Serial Number
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_BIOS | Select-Object SerialNumber | Format-List 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
           
                    2 { #PC Printer Information 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_Printer | Select-Object DeviceID,DriverName, PortName | Format-List 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed           
                        } 
           
                    3 { #Current User 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_ComputerSystem | Format-Table @{Expression={$_.Username};Label="Current User"} 
                        "" 
                        #May take a very long time if on a domain with many users 
                        #"All Users" 
                        #"------------" 
                        #gwmi -computer $compname Win32_UserAccount | foreach{$_.Caption} 
                        Read-Host -Prompt "Press Enter to continue"
                        CheckHostDetailed           
                        } 
           
                    4 { #OS Info 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_OperatingSystem | Format-List @{Expression={$_.Caption};Label="OS Name"},SerialNumber,OSArchitecture 
                        Read-Host -Prompt "Press Enter to continue"
                        CheckHostDetailed        
                        } 
           
                    5 { #System Info 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_ComputerSystem | Format-List Name,Domain,Manufacturer,Model,SystemType 
                        Read-Host -Prompt "Press Enter to continue"
                        CheckHostDetailed          
                        }         
           
                    6 { #Add/Remove Program List 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_Product | Sort-Object Name | Format-Table Name,Vendor,Version,InstallDate 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
           
                    7 { #Process Listx 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_Process | Select-Object Caption,Handle | Sort-Object Caption | Format-Table 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed          
                        } 
           
                    8 { #Service List 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_Service | Select-Object Name,State,Status,StartMode,ProcessID, ExitCode | Sort-Object Name | Format-Table 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed         
                        } 
         
                    9 { #USB Devices 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_USBControllerDevice | %{[wmi]($_.Dependent)} | Select-Object Caption, Manufacturer, DeviceID | Format-List 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed           
                        } 
           
                    10 { #Uptime 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        $wmi = gwmi -computer $compname Win32_OperatingSystem 
                        $localdatetime = $wmi.ConvertToDateTime($wmi.LocalDateTime) 
                        $lastbootuptime = $wmi.ConvertToDateTime($wmi.LastBootUpTime) 
             
                        "Current Time:      $localdatetime" 
                        "Last Boot Up Time: $lastbootuptime" 
             
                        $uptime = $localdatetime - $lastbootuptime 
                        "" 
                        "Uptime: $uptime" 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
                    11 { #Disk Info 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        $wmi = gwmi -computer $compname Win32_logicaldisk 
                        foreach($device in $wmi){ 
                                Write-Host "Drive: " $device.name    
                                Write-Host -NoNewLine "Size: "; "{0:N2}" -f ($device.Size/1Gb) + " Gb" 
                                Write-Host -NoNewLine "FreeSpace: "; "{0:N2}" -f ($device.FreeSpace/1Gb) + " Gb" 
                                "" 
                            } 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
                    12 { #Memory Info 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        $wmi = gwmi -computer $compname Win32_PhysicalMemory 
                        foreach($device in $wmi){ 
                            Write-Host "Bank Label:     " $device.BankLabel 
                            Write-Host "Capacity:       " ($device.Capacity/1MB) "Mb" 
                            Write-Host "Data Width:     " $device.DataWidth 
                            Write-Host "Device Locator: " $device.DeviceLocator     
                            ""         
                        } 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
                    13 { #Processor Info 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
                        gwmi -computer $compname Win32_Processor | Format-List Caption,Name,Manufacturer,ProcessorId,NumberOfCores,AddressWidth   
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
                    14 { #Monitor Info 
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            "" 
             
                        #Turn off Error Messages 
                        $ErrorActionPreference_Backup = $ErrorActionPreference 
                        $ErrorActionPreference = "SilentlyContinue" 
 
 
                        $keytype=[Microsoft.Win32.RegistryHive]::LocalMachine 
                        if($reg=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($keytype,$compname)){ 
                            #Create Table To Hold Info 
                            $montable = New-Object system.Data.DataTable "Monitor Info" 
                            #Create Columns for Table 
                            $moncol1 = New-Object system.Data.DataColumn Name,([string]) 
                            $moncol2 = New-Object system.Data.DataColumn Serial,([string]) 
                            $moncol3 = New-Object system.Data.DataColumn Ascii,([string]) 
                            #Add Columns to Table 
                            $montable.columns.add($moncol1) 
                            $montable.columns.add($moncol2) 
                            $montable.columns.add($moncol3) 
 
 
 
                            $regKey= $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\DISPLAY" ) 
                            $HID = $regkey.GetSubKeyNames() 
                            foreach($HID_KEY_NAME in $HID){ 
                                $regKey= $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\DISPLAY\\$HID_KEY_NAME" ) 
                                $DID = $regkey.GetSubKeyNames() 
                                foreach($DID_KEY_NAME in $DID){ 
                                    $regKey= $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\DISPLAY\\$HID_KEY_NAME\\$DID_KEY_NAME\\Device Parameters" ) 
                                    $EDID = $regKey.GetValue("EDID") 
                                    foreach($int in $EDID){ 
                                        $EDID_String = $EDID_String+([char]$int) 
                                    } 
                                    #Create new row in table 
                                    $monrow=$montable.NewRow() 
                         
                                    #MonitorName 
                                    $checkstring = [char]0x00 + [char]0x00 + [char]0x00 + [char]0xFC + [char]0x00            
                                    $matchfound = $EDID_String -match "$checkstring([\w ]+)" 
                                    if($matchfound){$monrow.Name = [string]$matches[1]} else {$monrow.Name = '-'} 
 
                         
                                    #Serial Number 
                                    $checkstring = [char]0x00 + [char]0x00 + [char]0x00 + [char]0xFF + [char]0x00            
                                    $matchfound =  $EDID_String -match "$checkstring(\S+)" 
                                    if($matchfound){$monrow.Serial = [string]$matches[1]} else {$monrow.Serial = '-'} 
                                                 
                                    #AsciiString 
                                    $checkstring = [char]0x00 + [char]0x00 + [char]0x00 + [char]0xFE + [char]0x00            
                                    $matchfound = $EDID_String -match "$checkstring([\w ]+)" 
                                    if($matchfound){$monrow.Ascii = [string]$matches[1]} else {$monrow.Ascii = '-'}          
 
                                 
                                    $EDID_String = '' 
                         
                                    $montable.Rows.Add($monrow) 
                                } 
                            } 
                            $montable | select-object  -unique Serial,Name,Ascii | Where-Object {$_.Serial -ne "-"} | Format-Table  
                        } else {  
                            Write-Host "Access Denied - Check Permissions" 
                        } 
                        $ErrorActionPreference = $ErrorActionPreference_Backup #Reset Error Messages 
                        Read-Host -Prompt "Press Enter to continue" 
                        CheckHostDetailed 
                        } 
                    x {CheckHost} 
                    default{CheckHostDetailed} 
                } 
            } 
            CheckHostDetailed
        }
		3 {
			#clear cache and temporary files

            function copyNIRCMD {
                #copy the nircmd.exe file over from network drive to pc being worked on
                #this file is needed for the script to work
                $filesFROM = "I:\Isd\System Support\scripts\tools\nircmd.exe"
                $filesTO = "\\$compname\c$\users\public"
                Copy-Item -Path $filesFROM -Destination $filesTO

            }

            function removeNIRCMD {
                #cleanup fuction to remove the EXE once task compelted
                Remove-Item -Path "\\$compname\c$\users\public\nircmd.exe"
            }
			
            function clearCache {
                if ($compname.Contains('SRPC')) {
                    Remove-Item "\\$compname\c$\Users\STA$storeNum\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\ST$storeNum\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\STM$storeNum\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                }
                else {
                    Remove-Item "\\$compname\c$\Users\$MODDEDUSER\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                }
            }

            function MrCleanCWPCSFPC {
                #this takes the compname string and extracts the 4-digit store number into a new substring called storeNum
			    $storeNum = $compname.Substring(5,4)

                #setting path names to remote device
                $pathname1 = "\\$compname\c$\Users\STA$storeNum\AppData\local\temp"
                $pathname5 = "\\$compname\c$\Users\ST$storeNum\AppData\local\temp"
                $pathname7 = "\\$compname\c$\Users\STM$storeNum\AppData\local\temp"

                if ($compname.Contains('SRPC')) {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
			        ""
                    Write-Host "It is expected to see multiple listings of red error code appear while this script runs."
                    Write-Host "It is attempting multiple temporary file directories with some not existing."
                    Start-Sleep 8
                    ""
                    
                    #function to clear cache
                    Write-Host "Clearing cache (Stage 1-of-4)...." 
                    clearCache
                    ""

                    Write-Host "Connecting to temporary file directories (Stage 2-of-4)...." 
			        ""

                    #adding network drives
                    New-PSDrive –Name “U” –PSProvider FileSystem –Root $pathname1 –Persist
                    New-PSDrive –Name “V” –PSProvider FileSystem –Root $pathname5 –Persist
                    New-PSDrive –Name “W” –PSProvider FileSystem –Root $pathname7 –Persist
                    Start-Sleep 2
                    ""

			        Write-Host "Drive(s) mounted"
			        ""

                    #clearing temp files
                    Write-Host "Clearing temporary files (Stage 3-of-4)...."
                    Remove-Item "U:\*.*" -Force -Recurse
                    Remove-Item "V:\*.*" -Force -Recurse
                    Remove-Item "W:\*.*" -Force -Recurse
                    ""
				
                    #removing network drives
			        Write-Host "Unmounting drive(s) (Stage 4-of-4)....."
                    Get-PSDrive U | Remove-PSDrive
                    Get-PSDrive V | Remove-PSDrive
                    Get-PSDrive W | Remove-PSDrive
                    Start-Sleep 2
			        Write-Host "Unmounted."
                }
                else {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
			        ""
                    Write-Host "It is expected to see multiple listings of red error code appear while this script runs."
                    Write-Host "It is attempting multiple temporary file directories with some not existing."
                    Start-Sleep 8
                    ""

                    #function to clear cache
                    Write-Host "Clearing cache (Stage 1-of-4)...." 
                    clearCache
                    ""

                    Write-Host "Connecting to temporary file directories (Stage 2-of-4)...." 
			        ""

                    #adding network drives
                    New-PSDrive –Name “U” –PSProvider FileSystem –Root $pathname1 –Persist
                    New-PSDrive –Name “V” –PSProvider FileSystem –Root $pathname2 –Persist
                    Start-Sleep 3
                    ""

			        Write-Host "Drive(s) mounted"
			        ""

                    #clearing temp files
                    Write-Host "Clearing temporary files (Stage 3-of-4)...."
                    Remove-Item "U:\*.*" -Force -Recurse
                    Remove-Item "V:\*.*" -Force -Recurse
                    ""
				
                    #removing network drives
			        Write-Host "Unmounting drive(s) (Stage 4-of-4)....."
			        Get-PSDrive U | Remove-PSDrive
			        Get-PSDrive V | Remove-PSDrive
			        Start-Sleep 2
			        Write-Host "Unmounted."
                }
            }

            function MrClean {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
				""
                #$emp = gwmi -computer $compname Win32_ComputerSystem | Format-Table @{Expression={$_.Username};Label="Current User"}
                $LOGGEDUSER = Read-Host "Enter the account name or user name (i.e. ST1234, jsmith)"
                #this will take any letters in the string and convert them to uppercase 
                $MODDEDUSER = $LOGGEDUSER.ToUpper()
                ""

                #setting path names to remote device
                $pathname1 = "\\$compname\c$\Users\$MODDEDUSER\AppData\local\temp"

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
				""
                Write-Host "It is expected to see multiple listings of red error code appear while this script runs."
                Write-Host "It is attempting multiple temporary file directories with some not existing."
                Start-Sleep 8
                ""

                #function to clear cache
                Write-Host "Clearing cache (Stage 1-of-4)...." 
                clearCache
                ""

                Write-Host "Connecting to temporary file directories (Stage 2-of-4)...." 
				""
                
                #added this 2019.10.29 as it runs the del cmd locally from the pc, which is SIGNIFICANTLY faster in deleting files
                #still allowing to mount and delete sub-folders afterwards to cleanup any extra files it may miss or need to force delete
                psexec \\$compname C:\users\public\nircmd.exe filldelete "c:\Users\$MODDEDUSER\AppData\Local\temp\*.*"

                #adding network drives
                New-PSDrive –Name “U” –PSProvider FileSystem –Root $pathname1 –Persist
                New-PSDrive –Name “V” –PSProvider FileSystem –Root $pathname2 –Persist
                Start-Sleep 2
                ""

				Write-Host "Drive(s) mounted"
				""

                #clearing temp files
                Write-Host "Clearing temporary files (Stage 3-of-4)...."
                Remove-Item "U:\*.*" -Force -Recurse
                Remove-Item "V:\*.*" -Force -Recurse
                ""
				
                #removing network drives
				Write-Host "Unmounting drive(s) (Stage 4-of-4)....."
				Get-PSDrive U | Remove-PSDrive
				Get-PSDrive V | Remove-PSDrive
				Start-Sleep 2
				Write-Host "Unmounted."
            }

            if (($compname.Contains('CW')) -and ($compname.Contains('STORE')))  {
                clearCache
                MrCleanCWPCSFPC

                funcPause
                CheckHost
            }
            if ($compname.Contains('ISK')) {
                clearCache
                MrCleanCWPCSFPC

                funcPause
                CheckHost
            }
            if (($compname.Contains('SRPC')) -and ($compname.Contains('STORE')))  {
                clearCache
                #this will need to check for STA, ST, and STM accts because of the diversity of how they are imaged
                MrCleanCWPCSFPC
                
                funcPause
                CheckHost
            }
            if ($compname.Contains('BOPC')) {
                #this is part of the new naming convention for the bopcs. they have been "Store####" but since 2019.04 they will be "Store####BOPC"
                #a number may be added at the end if "Store####BOPC" if/when the bopc needs to be replaced
                copyNIRCMD
                MrClean
                removeNIRCMD

                funcPause
                CheckHost
            }

            #if the device isn't a CWPC, SFPC, or SRPC it will run the MrClean function to run the script like on any other pc (i.e. home office, regional manager, etc...)
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "Clearing cache and temp files on $compane...."
            ""
            MrClean

            funcPause
            CheckHost
		}		
		4 {
			#reboot pc
			$msg = "We are rebooting this device. Please leave it alone until it has reloaded completely. Thank you."
			Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $compname
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			shutdown /r /m \\$compname /t 5
			Write-Host "Remote reboot signal sent to $compname"

			funcPause
			CheckHost
		}
		5 {
			#shutdown pc
            $msg = "We are shutting down this device. Please leave it alone during this process. Thank you."
			Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $compname
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			shutdown /s /m \\$compname /t 5
			Write-Host "Remote reboot signal sent to $compname"
			
            funcPause
			CheckHost
		}
        6 {
            #change volume

            function copyNIRCMD {
                #copy the nircmd.exe file over from network drive to pc being worked on
                #this file is needed for the script to work
                $filesFROM = "I:\Isd\System Support\scripts\tools\nircmd.exe"
                $filesTO = "\\$compname\c$\users\public"
                Copy-Item -Path $filesFROM -Destination $filesTO
            }

            function removeNIRCMD {
                #cleanup fuction to remove the EXE once task compelted
                Remove-Item -Path "\\$compname\c$\users\public\nircmd.exe"
            }

            function setVOLUME {
                #every time this function is run, the needed EXE is copied to the device to ensure it is there
                copyNIRCMD

                # 0 (silence) and 65535 (full volume) / 655.35 per 1%
                $volConvert = 655.35 * $volSelection
                #uses the math string modifier while also rounding up decimal points to the next whole number
                [math]::Round($volConvert)

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Attempting to unmute device (Stage 1-of-3)...."
                # 0 to unmute; 1 to mute ; 2 to toggle mute/unmute
                psexec \\$compname C:\users\public\nircmd.exe mutesysvolume 0 

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                ""
                Write-Host "Attempting to set device master volume to $volSelection% (Stage 2-of-3)...." 
                #to change volume percent
                psexec \\$compname C:\users\public\nircmd.exe setsysvolume $volConvert 

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Device volume set to $volSelection%. Sending test tone to $compname  (Stage 3-of-3)...."
                psexec \\$compname C:\users\public\nircmd.exe beep 400 3500
                removeNIRCMD
                
                $msg = "The volume of this device has been remotely adjusted."
			    Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $compname
                Start-Sleep -s 3

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Audio test tone sent to $compname."
                funcPause
                askPERCENT
            }

            function askPERCENT {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Enter the master volume percent below (i.e. 20 for 20%)."
                Write-Host "Leave blank and press Enter to go back to the main menu."
                [int]$volSelection = Read-Host "Enter percent (0-100)"

                if ($volSelection -eq "") {
                    CheckHost
                }
                #trying to do dual -ge with -le made null or letter responses valid for this IF statement
                #needed to set to -gt with -lt to make the IF statment work properly
                if (($volSelection -gt "-1") -and ($volSelection -lt "101")) {
                    setVOLUME
                    askPERCENT
                }
                ""
                Write-Host "Please input a percent value between 0-100"
                funcPause
                askPERCENT
                }

            askPERCENT
        }
        7 {
            #change clock time zone

            #reference of timezone IDs
            #Hawaiian Standard Time
            #Alaskan Standard Time
            #Pacific Standard Time
            #US Mountain Standard Time = Arizona
            #Mountain Standard Time
            #Central Standard Time
            #Canada Central Standard Time = Saskatchewan
            #Eastern Standard Time
            #US Eastern Standard Time = Indiana (East)
            #Atlantic Standard Time = Atlantic Time (Canada)
            
            function listofTZ {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                "What time zone would you like to change the PC to?" 
                "1)  Hawaiian Standard Time" 
                "2)  Alaskan Standard Time" 
                "3)  Pacific Standard Time" 
                "4)  Arizona" 
                "5)  Mountain Standard Time" 
                "6)  Central Standard Time"
                "7)  Canada Central Standard Time (Saskatchewan)" 
                "8)  Eastern Standard Time" 
                "9)  Eastern Standard Time (East Indiana)" 
                "10) Atlantic Standard Time (Canada)"
                "" 
                "X)  Cancel" 
                "" 
                $TZSelection = Read-Host "Enter Selection"
                changingTZ
            }
            
            function changingTZ {
                $connection = New-PSSession -ComputerName $compname
                switch ($TZSelection){ 
		            1 {
                        #Hawaiian Standard Time
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
			            ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Hawaiian Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Hawaiian Standard Time on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    2 {
                        #Alaskan Standard Time
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Alaskan Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Alaskan Standard Time on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    3 {
                        #Pacific Standard Time
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Pacific Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Pacific Standard Time on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    4 {
                        #US Mountain Standard Time = Arizona
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "US Mountain Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Arizona on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    5 {
                        #Mountain Standard Time
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Mountain Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Mountain Standard Time on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    6 {
                        #Central Standard Time
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Central Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Central Standard Time on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    7 {
                        #Canada Central Standard Time = Saskatchewan
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Canada Central Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Saskatchewan on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    8 {
                        #Eastern Standard Time
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Eastern Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Eastern Standard Time on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    9 {
                        #US Eastern Standard Time = Indiana (East)
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "US Eastern Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Indiana (East) on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    10 {
                        #Atlantic Standard Time = Atlantic Time (Canada)
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
				        ""
                        Write-Host "Changing time zone...."
                        Invoke-Command -Session $connection -Scriptblock {tzutil /s "Atlantic Standard Time"}
                        Start-Sleep -s 3
                        Remove-PSSession @connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Atlantic Time (Canada) on $compname."
                        Write-Host "If the PC is a CWPC or SFPC you may need to reboot it as well for the time zone modification to take affect."

                        funcPause
                        CheckHost
                    }
                    x {Clear-Host;CheckHost} 
                    default{listofTZ} 
                }
            }
            listofTZ
		}
        8 {
			#restart printer spooler and clear queue
            $msg = "We are working on this device. Please leave it alone until it has reloaded completely. Thank you."
			Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $compname
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""

            Write-Host  "Stopping the printer spooler on $compname...."
            Get-Service -Name SPOOLER -ComputerName $compname | Stop-Service -force
		    Start-Sleep -s 6
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host  "Clearing the printer print queue on $compname...."
            Remove-Item \\$compname\c$\Windows\System32\spool\printers\* -force
            Start-Sleep -s 6
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host  "Starting the printer spooler on $compname...."
            Get-Service -Name SPOOLER -ComputerName $compname | Set-Service -Status Running
            Start-Sleep -s 6
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
	        Write-Host "Script complete."
			
            funcPause
			CheckHost
		}
        9 {
            #an attempt to speed up an old PC still in use (i.e. HP 6005)
            
            function clearCacheOLDPC {
                Remove-Item "\\$compname\c$\Windows\Prefetch\*.*" -Force -Recurse
                Remove-Item "\\$compname\c$\Windows\Temp\*.*" -Force -Recurse
            }

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "Attempting to clear unnecessary files...."
            clearCacheOLDPC
            ""
            Write-Host "Script complete."

            funcPause
            CheckHost
        }
        10 {
		    #send a custom message
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
            ""
            $msg = Read-Host "Enter your message or input zero (0) to cancel"

            if ($msg -eq 0){
                Clear-Host
                CheckHost
            } else {
                Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $compname
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                ""
                Write-Host "Your message has been sent."
            }
            
            funcPause
            CheckHost
		}
        11 {
            #kill a running process
            
            function killPID {
                #grabs the username of the current person logged into the computer attempting to use the script
                $LOGGEDUSER1 = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                $LOGGEDUSER2 = $LOGGEDUSER1.Substring(8)
                taskkill /F /S $compname /U $LOGGEDUSER2 /PID $PIDSelection
            }
            
            function reloadPIDs {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
			    "" 
                Write-Host "The process has been remotely killed. Reloading process list...."
                ""
                gwmi -computer $compname Win32_Process | Select-Object Caption,Handle | Sort-Object Caption | Format-Table 
                ""
                Write-Host "Enter the PID # of the process to kill (Leave blank to exit to the Main Menu)"
                $PIDSelection = Read-Host "PID #"

                if ($PIDSelection -eq "") {
                    CheckHost
                }
                else {
                    killPID
                    reloadPIDs
                }
            }

            function loadPIDs {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
			    "" 
                Write-Host "Loading process list...."
                ""
                gwmi -computer $compname Win32_Process | Select-Object Caption,Handle | Sort-Object Caption | Format-Table 
                ""
                Write-Host "Enter the PID # of the process to kill (Leave blank to exit to the Main Menu)"
                $PIDSelection = Read-Host "PID #"

                if ($PIDSelection -eq "") {
                    CheckHost
                }
                else {
                    killPID
                    reloadPIDs
                }
            }

            loadPIDs

            #not really needed as option choice leads into function that exits to main menu or loops into reload function
            #so as to reload the services again to kill another or confirm task no longer running
            funcPause
            CheckHost
        }
		12 {
			#kill 'Block' processes
            #grabs the username of the current person logged into the computer attempting to use the script
            $LOGGEDUSER1 = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $LOGGEDUSER2 = $LOGGEDUSER1.Substring(8)
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			Write-Host "Attempting to remotely kill the block processes on $compname...."
			""
			taskkill /F /S $compname /U $LOGGEDUSER2 /IM BlockKeyboard.exe
	        taskkill /F /S $compname /U $LOGGEDUSER2 /IM BlockMouse.exe
            Write-Host ""
	        Write-Host "Script complete."

			funcPause
			CheckHost
		}
        13 {
            #copy all files and folders on the desktop, Documents, Downloads, and Pictures of the ST, STM, or personal emp
            #then save them to their respective network drive under a newly made, unique folder

            function backupPC {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 

                #this replaces telling spacifically which server and where by searching AD for their HomeDirectory and saving it to a var
                $homedir = (Get-ADUser $MODDEDUSER -properties *).HomeDirectory
                
                #making variables for My Documents to backup (Desktop, Documents, Downloads, and Pictures)
                $pathFROMEmpDesktop1 = "\\$compname\c$\Users\$MODDEDUSER\Desktop"
                $pathFROMEmpDocuments1 = "\\$compname\c$\Users\$MODDEDUSER\Documents"
                $pathFROMEmpDownloads1 = "\\$compname\c$\Users\$MODDEDUSER\Downloads"
                $pathFROMEmpPictures1 = "\\$compname\c$\Users\$MODDEDUSER\Pictures"

                #gets todays date in the order of year month day as numeric and saves it to a string
                $dirDateTime = (Get-Date).ToString("yyyy.MM.dd_HHmm")
                #new folder name plus the date generated from previous string to create unique, dated folder
                $dirNameDateTime = "$compname documents backup" + " " + $dirDateTime
                
                if ( Test-Path $homedir -PathType Container ) {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    "Confirmed server and network directory both exist. Creating a backup folder on $MODDEDUSER network folder...."
                    #making the dir the backup will be stored to
                    New-Item -Path "$homedir\$dirNameDateTime" -type directory -Force

                    #after confirming server 1 and the directory are available, checking which path is used for My Documents.. with or without ".PIERFW1"
                    if ( Test-Path $pathFROMEmpDesktop1 -PathType Container ) {
                        ""
                        Write-Host "Copying documents to the newly made network backup folder...."
                        Copy-Item -Path $pathFROMEmpDesktop1 -Recurse -Destination "$homedir\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDocuments1 -Recurse -Destination "$homedir\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDownloads1 -Recurse -Destination "$homedir\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpPictures1 -Recurse -Destination "$homedir\$dirNameDateTime" -Container

                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        "" 
                        Write-Host "The document backup for $MODDEDUSER has completed."

                        funcPause
                        backupMenu
                    }
                }

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Either the server does not exit or the network folder for the account doesn't. Backup not created."

                funcPause
                backupMenu
            }

            function backupInput {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Enter the network username (i.e. ST1234, jsmith, or 444555)"
                Write-Host "Leave blank and press Enter to go back to the main menu."
                $LOGGEDUSER = Read-Host "Username"

                if ($LOGGEDUSER -eq "") {
                    CheckHost
                }
                else {
                    #this will take any letters in the string and convert them to uppercase 
                    $MODDEDUSER = $LOGGEDUSER.ToUpper()
                    backupPC

                    backupInput
                }
            }

            backupMenu
	    }

        c {Clear-Host; GetCompName} 
        x {Clear-Host; MainMenu} 
        default{CheckHost} 
    }
}

function acctCheck {
    $empID = ""
    $empIDUPPER = ""
    $empChoice = ""
    Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
    ""
    Write-Host "Input the employee ID or user name to search."
    Write-Host "Leave blank and press Enter to go back to the main menu."
    ""
    $empID = Read-Host "User name"

    if ($empID -eq "") {
        MainMenu
    } else {
        $empIDUPPER = $empID.ToUpper()

        function empMenuORIGINAL {
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "    Employee: $empIDUPPER" 
            ""
            net user $empIDUPPER /domain | findstr /b "Full. Comment. Account. Password"
            ""
            Read-Host -Prompt "Press Enter to continue"
            acctCheck
        }

        function empMenu {
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "    Employee: $empIDUPPER" 
            ""
            net user $empIDUPPER /domain | findstr /b "Full. Comment. Account. Password"
            ""
            "1)  Account Unlock"
            "2)  [COMING SOON] Password Reset"
            ""
            "C)  Switch to a different employee"
            "X)  Exit to the Main Menu" 
            ""
            $empChoice = Read-Host "Enter Selection"

            switch ($empChoice){
                1 {
                    $LOGGEDUSER1 = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    $LOGGEDUSER2 = $LOGGEDUSER1.Substring(8)
                    #acct unlock
                    #Unlock-ADAccount -Credential $LOGGEDUSER2 -Identity $empIDUPPER
                    Unlock-ADAccount -Identity $empIDUPPER
                    ""
                    Read-Host -Prompt "The account has been unlocked"
                    empMenu
                }
                2 {
                    empMenu

                    $newpwd = ""
                    $newpwdSECURE = ""
                    #pwd reset
                    $newpwd = Read-Host "Enter the new password"
                    $newpwdCount = $newpwd.Length
                    if ($newpwdCount -gt '6') {
                        $LOGGEDUSER1 = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                        $LOGGEDUSER2 = $LOGGEDUSER1.Substring(8)
                        $newpwdSECURE = ConvertTo-SecureString -AsPlainText 'New password' -Force
                        Set-ADAccountPassword $empIDUPPER -NewPassword $newpwd –Reset
                        ""
                        Read-Host -Prompt "The new password has been set"
                        #clearing varibles before leaving screen
                        $newpwd = ""
                        $newpwdCount = ""
                        empMenu
                    } else {
                        ""
                        Read-Host -Prompt "The password entered is under 7 characters. Please try again"
                        #clearing varibles before leaving screen
                        $newpwd = ""
                        $newpwdCount = ""
                        empMenu
                    }
                    #clearing varibles before leaving screen
                    $newpwd = ""
                    $newpwdCount = ""
                    empMenu
                }

                c {Clear-Host; acctCheck}
                x {Clear-Host; MainMenu}
            }
            empMenu
        }
        empMenu
    }
}
    
    $compname = $env:computername
    $ping = gwmi Win32_PingStatus -filter "Address='$compname'"
    $pcip = $ping.ProtocolAddress
    
    function getDuration {
        #removing this function and just having the loop run until 7am
        $compname = $env:computername
        $ping = gwmi Win32_PingStatus -filter "Address='$compname'"
        $pcip = $ping.ProtocolAddress
        CLS
        "  /----------------------\" 
        "  |      SLIM TOOL       |" 
        "  \----------------------/" 
        "  $compname ($pcip)" 
        ""
        [int] $addHours = Read-Host "How many hours do you want to monitor (1-24)"

        if ($addHours -eq ""){
            MainMenu
        }

        if (($addHours -gt 0) -and ($addHours -lt 25)) {
            #collects the current time without a seconds value, zeros out the seconds
            $timeStart = Get-Date -Date (Get-Date -Format “yyyy-MM-dd HH:mm”)
            #adds input number of hours to the current time to set as the end time
            $timeEnd = $timeStart.AddHours($addHours)
            #$timeEnd = $timeStart.AddMinutes($addHours)

            folderCheck
        }
        else {
            ""
            "The number you entered is out of bounds. Please try again."
                
            funcPause
            getDuration
        }
    }

    function folderCheck {
        
        #this fuction has the do loop so every time the loop is run, it gets the 3 folder names and their latest time stamps to display

        Do { 
            #at the start of each loop, gets the most recent time to compare against set loop end time
            $TimeNow = Get-Date -Date (Get-Date -Format “yyyy-MM-dd HH:mm”)
            $TimeNow_Time = Get-Date -Format HHmm
            #[str]$TimeEnd_Hour = "0700"
            $recentfileTOKEN = ""
            $recentfilePOLL1 = ""
            $recentfilePOLL2 = ""
            $recentfilePOLL3 = ""

            #if ($TimeNow_Time -gt "0700") {
            #    $TimeEnd = $TimeNow.AddDays(1)
            #}

            #if ($TimeNow_Time.StartsWith("05")) {
            #
            #}

            if ($TimeNow -le $TimeEnd) {
                
                ##### getting TokenData and PollData folder data #####
                #get-childitem path | where filter only folders | sort most recent folder to top | select the first folder
                $folderTOKEN = gci \\vmfwposaomp3\EPICOR\aw_pier1\TokenData | ? { $_.PSIsContainer } | sort CreationTime -desc | select -f 1
                #$folderTOKEN = gci \\vmfwposaomp3\EPICOR\aw_pier1\TokenData | ? { $_.PSIsContainer } | sort LastWriteTime -desc | select -f 1
                $pathTOKEN = "\\vmfwposaomp3\EPICOR\aw_pier1\TokenData" + "\$folderTOKEN"
                #gets the file name of the most recent file in the folder; then takes it and outputs the date as a string to be usable in IF statement
                $fileTOKEN = gci $pathTOKEN | sort CreationTime -desc | select -f 1
                $recentfileTOKEN = $fileTOKEN.LastWriteTime

                #get-childitem path | where filter only folders | sort most recent folder to top | select the first folder
                $folderPOLL1 = gci \\vmfwposaomp3\EPICOR\aw_pier1\PollData | ? { $_.PSIsContainer } | sort CreationTime -desc | select -f 1
                $pathPOLL1 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL1"
                $filePOLL1 = gci $pathPOLL1 | sort CreationTime -desc | select -f 1
                $recentfilePOLL1 = $filePOLL1.LastWriteTime

                #does same as above but skips the first folder to grab the second folder name
                $folderPOLL2 = gci \\vmfwposaomp3\EPICOR\aw_pier1\PollData | ? { $_.PSIsContainer } | sort CreationTime -desc | select-object -skip 1 | select-object -f 1
                $pathPOLL2 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL2"
                $filePOLL2 = gci $pathPOLL2 | sort CreationTime -desc | select -f 1
                $recentfilePOLL2 = $filePOLL2.LastWriteTime

                #does same as above but skips the first 2 folders to grab the third folder name
                $folderPOLL3 = gci \\vmfwposaomp3\EPICOR\aw_pier1\PollData | ? { $_.PSIsContainer } | sort CreationTime -desc | select-object -skip 2 | select-object -f 1
                $pathPOLL3 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL3"
                $filePOLL3 = gci $pathPOLL3 | sort CreationTime -desc | select -f 1
                $recentfilePOLL3 = $filePOLL3.LastWriteTime

                #does same as above but skips the first 2 folders to grab the third folder name
                $folderPOLL4 = gci \\vmfwposaomp3\EPICOR\aw_pier1\PollData | ? { $_.PSIsContainer } | sort CreationTime -desc | select-object -skip 3 | select-object -f 1
                $pathPOLL4 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL4"
                $filePOLL4 = gci $pathPOLL4 | sort CreationTime -desc | select -f 1
                $recentfilePOLL4 = $filePOLL4.LastWriteTime

                #comparing if most recent date is newer than previous run captured date


                #creating a checksum for each dir so a var will hold the last captured timestamp in case the following loop returns null
                #this SHOULD also take the var created at the end of the loop to compare that if the last time the loop ran compared to the next,
                #if the file date/time stamp from the previous loop is newer, it doesn't revert back to an earlier date/time stamp but keeps the newest one
                if ($recentfileTOKEN -ne $null) {
                    #if (($recentfileTOKEN_previous -gt $recentfileTOKEN) -and ($recentfileTOKEN -ne "No recent files created")) {
                    #    $recentfileTOKEN = $recentfileTOKEN_previous
                    #}
                    $recentfileTOKENcheck = $recentfileTOKEN
                }
                if ($recentfilePOLL1 -ne $null) {
                    #if (($recentfilePOLL1_previous -gt $recentfilePOLL1) -and ($recentfilePOLL1 -ne "No recent files created")) {
                    #    $recentfilePOLL1 = $recentfilePOLL1_previous
                    #}
                    $recentfilePOLL1check = $recentfilePOLL1
                }
                if ($recentfilePOLL2 -ne $null) {
                    #if (($recentfilePOLL2_previous -gt $recentfilePOLL2) -and ($recentfilePOLL2 -ne "No recent files created")) {
                    #    $recentfilePOLL2 = $recentfilePOLL2_previous
                    #}
                    $recentfilePOLL2check = $recentfilePOLL2
                }
                if ($recentfilePOLL3 -ne $null) {
                    #if (($recentfilePOLL3_previous -gt $recentfilePOLL3) -and ($recentfilePOLL3 -ne "No recent files created")) {
                    #    $recentfilePOLL3 = $recentfilePOLL3_previous
                    #}
                    $recentfilePOLL3check = $recentfilePOLL3
                }
                if ($recentfilePOLL4 -ne $null) {
                    #if (($recentfilePOLL3_previous -gt $recentfilePOLL3) -and ($recentfilePOLL3 -ne "No recent files created")) {
                    #    $recentfilePOLL3 = $recentfilePOLL3_previous
                    #}
                    $recentfilePOLL4check = $recentfilePOLL4
                }

                #now comparing the most recent run against the last known previous captured timestamp
                #if the latest lopp is null but the checksum isn't, the checksum overrights the latest run var
                #if the latest lopp is null and the checksum is too, the latest run var is filled with the statement of no no files to compare to
                if ($recentfileTOKEN -eq $null) {
                    if ($recentfileTOKENcheck -ne $null) {
                        $recentfileTOKEN = $recentfileTOKENcheck
                    }
                    else {
                        $recentfileTOKEN = "No recent files created"
                    }
                }
                if ($recentfilePOLL1 -eq $null) {
                    if ($recentfilePOLL1check -ne $null) {
                        $recentfilePOLL1 = $recentfilePOLL1check
                    }
                    else {
                        $recentfilePOLL1 = "No recent files created"
                    }
                }
                if ($recentfilePOLL2 -eq $null) {
                    if ($recentfilePOLL2check -ne $null) {
                        $recentfilePOLL2 = $recentfilePOLL2check
                    }
                    else {
                        $recentfilePOLL2 = "No recent files created"
                    }
                }
                if ($recentfilePOLL3 -eq $null) {
                    if ($recentfilePOLL3check -ne $null) {
                        $recentfilePOLL3 = $recentfilePOLL3check
                    }
                    else {
                        $recentfilePOLL3 = "No recent files created"
                    }
                }
                if ($recentfilePOLL4 -eq $null) {
                    if ($recentfilePOLL4check -ne $null) {
                        $recentfilePOLL4 = $recentfilePOLL3check
                    }
                    else {
                        $recentfilePOLL4 = "No recent files created"
                    }
                }

                function timeCompare_Active {
                    
                    #this will make the line of text
                    #$text = " this is a test of the system "
                    #write-host $text -ForegroundColor Black -BackgroundColor Red

                    ### make all variables in here global by doing $global:VariableName so they can be used in the outputLists function when writing to the screen


                    if (($recentfileTOKEN -ne $null) -and ($recentfileTOKEN -ne "No recent files created")) {
                        if (($recentfilePOLL1 -ne $null) -and ($recentfilePOLL1 -ne "No recent files created")) {
                            if (($recentfilePOLL3 -ne $null) -and ($recentfilePOLL3 -ne "No recent files created")) {

                                #compares the timespan between a start/end date
                                #if the result is a negative, it flips the variables to recalc and get a positive output

                                ### Token / Poll1
                                $timespan_TokenPoll1 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL1
                                if ($timespan_TokenPoll1 -ge 0) {
                                    if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54)) {
                                        msgboxTimeAlert
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                        $Global:warningDelay = $TimeNow.AddMinutes(30)
                                    }
                                }
                                if ($timespan_TokenPoll1 -lt 0) {
                                    $timespan_TokenPoll1 = New-Timespan -Start $recentfilePoll1 -End $recentfileToken
                                    if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54)) {
                                        msgboxTimeAlert
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                        $Global:warningDelay = $TimeNow.AddMinutes(30)
                                    }
                                }

                                ### Token / Poll3
                                $timespan_TokenPoll3 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL3
                                if ($timespan_TokenPoll3 -ge 0) {
                                    if (($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                        msgboxTimeAlert
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                        $Global:warningDelay = $TimeNow.AddMinutes(30)
                                    }
                                }
                                if ($timespan_TokenPoll3 -lt 0) {
                                    $timespan_TokenPoll3 = New-Timespan -Start $recentfilePoll3 -End $recentfileToken
                                    if (($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                        msgboxTimeAlert
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                        $Global:warningDelay = $TimeNow.AddMinutes(30)
                                    }
                                }

                                ### Poll1 / Poll3
                                $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll1 -End $recentfilePOLL3
                                if ($timespan_Poll1Poll3 -ge 0) {
                                    if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                        msgboxTimeAlert
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                        $Global:warningDelay = $TimeNow.AddMinutes(30)
                                    }
                                }
                                if ($timespan_Poll1Poll3 -lt 0) {
                                    $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll3 -End $recentfilePoll1
                                    if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                        msgboxTimeAlert
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                        ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                        $Global:warningDelay = $TimeNow.AddMinutes(30)
                                    }
                                }
                            }
                        }
                    }

                }

                ### run this function to be able to pull global variables from it into this function
                #timeCompare_Active

                #after all the above, whatever the most recent captured timestamp is or the empty notification gets displayed accortingly
                #the source folder for each of variable directories are listed followed by either the last modified time or empty notification
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                ""
                Write-Host "Monitoring loop running until: " $timeEnd
                ""
                "Source:     Folder:             Time:"
                Write-Host "TokenData  " $folderTOKEN "   " $recentfileTOKEN

                ##### instead of adjusting the spacing based on the folder name character count as it seemed to be unreliable,
                ##### we're instead converting the folder names into a string to compare the 4 variations AWL/IP, AWL/TR, AWL/TR1, and AW3
                ##### with if/and statements to only output the line with the proper spacing

                ### folder 1
                $folderPOLL1_str = $folderPOLL1.ToString()
                if ($folderPOLL1_str.StartsWith("AW3")) {
                    if ($folderPOLL1_str.EndsWith("TR2")) {
                        Write-Host "PollData   " $folderPOLL1 "" $recentfilePOLL1
                    }
                    if ($folderPOLL1_str.EndsWith("IP")) {
                        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                    }
                }
                if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("IP") )) {
                    Write-Host "PollData   " $folderPOLL1 "   " $recentfilePOLL1
                }
                if ((($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR"))) -or (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR1"))) -or (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR2")))) {
                    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                    $folderPOLL1 = $folderPOLL2
                    $folderPOLL1_str = $folderPOLL1.ToString()
                    $recentfilePOLL1 = $recentfilePOLL2
                    
                    if (($folderPOLL1_str.StartsWith("AW3")) -and ($folderPOLL1_str.EndsWith("TR2"))) {
                        Write-Host "PollData   " $folderPOLL1 "" $recentfilePOLL1
                    }
                    if (($folderPOLL1_str.StartsWith("AW3")) -and ($folderPOLL1_str.EndsWith("IP"))) {
                        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                    }
                    if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("IP") )) {
                        Write-Host "PollData   " $folderPOLL1 "   " $recentfilePOLL1
                    }
                    if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR1") )) {
                        $folderPOLL1 = $folderPOLL4
                        $recentfilePOLL1 = $recentfilePOLL4
                        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                    }
                    if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR2") )) {
                        #$folderPOLL1 = $folderPOLL4
                        #$recentfilePOLL1 = $recentfilePOLL4
                        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                    }
                }
                #if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR1") )) {
                #    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                #    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                #    $folderPOLL1 = $folderPOLL2
                #    $folderPOLL1_str = $folderPOLL1.ToString()
                #    $recentfilePOLL1 = $recentfilePOLL2
                #    
                #    if ($folderPOLL1_str.StartsWith("AW3")) {
                #        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                #    }
                #    if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("IP") )) {
                #        Write-Host "PollData   " $folderPOLL1 "   " $recentfilePOLL1
                #    }
                #    if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR1") )) {
                #        $folderPOLL1 = $folderPOLL4
                #        $recentfilePOLL1 = $recentfilePOLL4
                #        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                #    }
                #}

                ### folder 3
                $folderPOLL3_str = $folderPOLL3.ToString()
                if ($folderPOLL3_str.StartsWith("AW3")) {
                    if ($folderPOLL3_str.EndsWith("TR2")) {
                        Write-Host "PollData   " $folderPOLL3 "" $recentfilePOLL3
                    }
                    if ($folderPOLL3_str.EndsWith("IP")) {
                        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                    }
                }
                if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                    Write-Host "PollData   " $folderPOLL3 "   " $recentfilePOLL3
                }
                if ((($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR"))) -or (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR1"))) -or (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR2")))) {
                    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                    $folderPOLL3 = $folderPOLL2
                    $folderPOLL3_str = $folderPOLL3.ToString()
                    $recentfilePOLL3 = $recentfilePOLL2
                    
                    if (($folderPOLL3_str.StartsWith("AW3")) -and ($folderPOLL3_str.EndsWith("TR2"))) {
                        Write-Host "PollData   " $folderPOLL3 "" $recentfilePOLL3
                    }
                    if (($folderPOLL3_str.StartsWith("AW3")) -and ($folderPOLL3_str.EndsWith("IP"))) {
                        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                        Write-Host "PollData   " $folderPOLL3 "   " $recentfilePOLL3
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR1") )) {
                        $folderPOLL3 = $folderPOLL4
                        $recentfilePOLL3 = $recentfilePOLL4
                        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR2") )) {
                        #$folderPOLL3 = $folderPOLL4
                        #$recentfilePOLL3 = $recentfilePOLL4
                        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                    }
                }
                #if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR1") )) {
                #    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                #    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                #    $folderPOLL3 = $folderPOLL2
                #    $folderPOLL3_str = $folderPOLL3.ToString()
                #    $recentfilePOLL3 = $recentfilePOLL2
                #    
                #    if ($folderPOLL3_str.StartsWith("AW3")) {
                #        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                #    }
                #    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                #        Write-Host "PollData   " $folderPOLL3 "   " $recentfilePOLL3
                #    }
                #}


                ##### getting BulkInventoryLoad folder data #####
                $pathBULK = ""
                $pathBULK = "\\pier1.com\filedrop\Demandware\Production\Outgoing\inventory\BulkInventoryLoad"
                $folderBULKcount = Get-ChildItem -Path $pathBULK | Measure-Object | %{$_.Count}

                #made to convert file sizes into something more readable
                Function Format-FileSize() {
                    Param ([int]$size)
                    #If     ($size -gt 1TB) {[string]::Format("{0:0.00} TB", $size / 1TB)}
                    #ElseIf ($size -gt 1GB) {[string]::Format("{0:0.00} GB", $size / 1GB)}
                    #ElseIf ($size -gt 1MB) {[string]::Format("{0:0.00} MB", $size / 1MB)}
                    If ($size -gt 1KB) {[string]::Format("{0:0.00} KB", $size / 1KB)}
                    ElseIf ($size -gt 0)   {[string]::Format("{0:0.00} B", $size)}
                    Else                   {"0 B"}
                }

                $filesBULK_sizes = ""
                $filesBULK_sizes = Get-ChildItem -Path $pathBULK | Select-Object Name, @{Name="Size";Expression={Format-FileSize($_.Length)}}


                ""
                ""
                "Source:             Count:"
                Write-Host "BulkInventoryLoad" " " $folderBULKcount
                ""
                $filesBULK_sizes

                function timeCompare_Notify {

                    function msgboxTimeAlert {
                        $msgboxAlert =  [System.Windows.MessageBox]::Show('Monitored folders have exceeded the time variance threshold.','Monitoring Alert','Ok','Error')
                        #Start-Sleep 5
                    }

                    function compareNotify {
                        #only if all 3 folders have date/time values will the script compare
                        if (($recentfileTOKEN -ne $null) -and ($recentfileTOKEN -ne "No recent files created")) {
                            if (($recentfilePOLL1 -ne $null) -and ($recentfilePOLL1 -ne "No recent files created")) {
                                if (($recentfilePOLL3 -ne $null) -and ($recentfilePOLL3 -ne "No recent files created")) {

                                    #compares the timespan between a start/end date
                                    #if the result is a negative, it flips the variables to recalc and get a positive output

                                    ### Token compared to Poll1 & Poll3
                                    $timespan_TokenPoll1 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL1
                                    $timespan_TokenPoll3 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL3
                                    if (($timespan_TokenPoll1 -ge 0) -or ($timespan_TokenPoll3 -ge 0)) {
    
                                        if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54) -or ($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            #opens explorer to the listed directories
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }
                                    if (($timespan_TokenPoll1 -lt 0) -or ($timespan_TokenPoll3 -lt 0)) {
                                        $timespan_TokenPoll1 = New-Timespan -Start $recentfilePoll1 -End $recentfileToken
                                        $timespan_TokenPoll3 = New-Timespan -Start $recentfilePoll3 -End $recentfileToken
    
                                        if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54) -or ($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }

                                    ### Poll1 compared to Poll3
                                    $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll1 -End $recentfilePOLL3
                                    if ($timespan_Poll1Poll3 -ge 0) {
    
                                        if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }
                                    if ($timespan_Poll1Poll3 -lt 0) {
                                        $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll3 -End $recentfilePoll1
    
                                        if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            ii \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if ($warningDelay -lt $TimeNow) {
                        compareNotify
                    }

                    #outputs only the "minute" difference, ignoring days, hours, seconds, and more
                    #$timespan1.Minutes
                }

                timeCompare_Notify
                Start-Sleep -Seconds 5
            }

            $recentfileTOKEN_previous = $recentfileTOKEN
            $recentfilePOLL1_previous = $recentfilePOLL1
            $recentfilePOLL2_previous = $recentfilePOLL2
            $recentfilePOLL3_previous = $recentfilePOLL3

        } Until ($TimeNow -eq $TimeEnd)
    }

    function msgbox {
        $msgBoxInput =  [System.Windows.MessageBox]::Show('The monitoring script has completed its run time. Would you like to continue monitoring?','Monitoring completed','YesNo','Information')

        switch  ($msgBoxInput) {
            'Yes' {
                $timeStart = Get-Date -Date (Get-Date -Format “yyyy-MM-dd HH:mm”)
                $timeEnd = $timeStart.AddHours(1)
                folderCheck
            }

            'No' {
                MainMenu
            }
        }
    }

    getDuration
    #foldercheck
    msgbox
    MainMenu
}

function lifePath {
    switch ($mainChoice){ 
		1 {
            GetCompName
        }
        2 {
            acctCheck
        }
        3 {
            function pulldeviceList {
                cd c:\
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                ""
                "Searching for all devices with '$storeNum' in the device name...."
                Get-ADComputer -properties Name -Filter * | Where-Object {$_.Name -match $storeNum} | Format-List -Property Name
                Read-Host -Prompt "Press Enter to continue"
                getDeviceName
            }

            function getDeviceName {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                ""
                Write-Host "Input part of a device name to search for all partial matches in Active Directory."
                Write-Host "Leave blank and press Enter to go back to the main menu."
                ""
                $storeNum = Read-Host "Enter what to search for"

                if ($storeNum -eq "") {
                    MainMenu
                }
                else {
                    pulldeviceList
                }
            }

            getDeviceName
        }

        x {
            exit
        }
        default{MainMenu}
    }
    getstoreNum
}

function MainMenu {
    iniStart
    Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
    ""
    "1)  Connect to a computer" 
    "2)  Employee account info"
    "4)  Search Active Directory for devices"
    ""
    "X)  Exit the program" 
    ""
    $mainChoice = Read-Host "Enter Selection"
    lifePath
}
    Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
    ""
    "1)  Amy Evans...........x8788 DTAEVANS1 DTAEVANS3"
    "2)  Brad Arrowood.......x7836 DTBARROWOOD DTBXARROWOOD1 DTBXARROWOOD2"
    "3)  Gabriela Barrett....x6289 DTGBARRETT DTGBARRETT2 DTGXQUINONES2"
    "4)  Jimmie Young........x6110 LTJYOUNG"
    "5)  Justin Jones........x6090 DTJJJONES DTSERVICEDESK02"
    "6)  Larry Fewell........x8791 LTLFEWELL DTLFEWELL1 DTLFEWELL2"
    "7)  Michael Roye........x8194 DTMROYE DTMROYE2 DTMROYEHOME"
    "8)  Paul Murry..........x8268 DTPMURRY DTPMURRY1 DTPMURRY2"
    "9)  Preston Smith.......x8475 DTPSMITH DTPSMITH23"
    "10) Raymond Garcia......x8772 DTRGARCIA DTRGARCIA2"
    "11) Rodney Gordon.......x8789 DTRGORDON1 DTRGORDON3 DTRGORDON4"
    "12) Teresa Withers......x8872 DTTWITHERS1 DTTWITHERS2"
    "13) Tony Fitch..........x8594 DTTFITCH DTTFITCH1"

    funcPause
    MainMenu
}

#---------Start Main-------------- 
Clear-Host
MainMenu
