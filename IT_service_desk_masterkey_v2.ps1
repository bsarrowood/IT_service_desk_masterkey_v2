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

#references:
#https://techtalk.gfi.com/11-most-useful-powershell-commands-for-remote-management/
#https://social.technet.microsoft.com/wiki/contents/articles/7703.powershell-running-executables.aspx
#https://ss64.com/ps/syntax-compare.html  <-- IF statement comparision operations list
#https://blog.netspi.com/powershell-remoting-cheatsheet/
#https://nircmd.nirsoft.net/


function funcPause { 
	"" 
	Read-Host -Prompt "Press Enter to continue" 
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 
}

function GetCompName{

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
    #this checks if the compname is a pos or something else to direct to different menu options
    $compnameUPPER = $compname.ToUpper()

    #added this IF statement to be able to use POSMenu on registers in the lab
    if ($compnameUPPER.StartsWith("LAB-")) {
        $ping = Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$compname'" 
        if($ping.StatusCode -eq 0){$pcip=$ping.ProtocolAddress; POSMenu} 
        else{Read-Host -Prompt "Host $compname is offline...Press any key to continue"; GetCompName} 
    }

    if ($compnameUPPER.StartsWith("POS-")) {
        $ping = Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$compname'" 
        if($ping.StatusCode -eq 0){$pcip=$ping.ProtocolAddress; POSMenu} 
        else{Read-Host -Prompt "Host $compname is offline...Press any key to continue"; GetCompName} 
    }
    else {
        $ping = Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$compname'" 
        if($ping.StatusCode -eq 0){$pcip=$ping.ProtocolAddress; GetMenu} 
        else{Read-Host -Prompt "Host $compname is offline...Press any key to continue"; GetCompName} 
    }
}

function GetMenu {
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
    "6)  Change volume (Requires update to PowerShell. See Brad before using.)"
    "7)  Change clock time zone"
    "8)  Restart printer spooler and clear print queue"
    "9)  Speed up a PC (i.e. HP 6005)"
    "10) Send custom message to PC"
    "11) Kill a running process"
    "12) Kill any 'Block' processes (CWPC and SFPC Only)"
    "13) [COMING SOON] Correct device to auto login (CWPC and SFPC Only)"
    "14) Make a backup of everything on the Desktop and/or My Documents to their network drive"
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
			Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-Table @{Expression={$_.Username};Label="Current User"} 
			"------------------------------"
			#OS Info
			Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $compname | Format-List @{Expression={$_.Caption};Label="OS Name"},SerialNumber,OSArchitecture 
			"------------------------------"
			#System Info
			"System Info:"
			Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-List Name,Domain,Manufacturer,Model,SystemType 
			"------------------------------"
			#Uptime
            ""
			$wmi = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $compname 
            $localdatetime = $wmi.LocalDateTime 
            $lastbootuptime = $wmi.LastBootUpTime 
            "Current Time:      $localdatetime" 
            "Last Boot Up Time: $lastbootuptime" 
            $uptime = $localdatetime - $lastbootuptime 
            "Uptime: $uptime" 
            ""
			"------------------------------"
			#Disk Space
            ""
			"Disk Space:"
			$wmi = Get-CimInstance -ClassName Win32_logicaldisk -ComputerName $compname 
            foreach($device in $wmi){ 
                    Write-Host "Drive: " $device.name    
                    Write-Host -NoNewLine "Size: "; "{0:N2}" -f ($device.Size/1Gb) + " Gb" 
                    Write-Host -NoNewLine "FreeSpace: "; "{0:N2}" -f ($device.FreeSpace/1Gb) + " Gb" 
                    "" 
             } 
			"------------------------------"
			#Memory Info
			"Memory Info:"
			$wmi = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $compname 
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
                $ping = Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$compname'" 
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
                        Get-CimInstance -ClassName Win32_BIOS -ComputerName $compname | Select-Object SerialNumber | Format-List 
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
                        Get-CimInstance -ClassName Win32_Printer -ComputerName $compname | Select-Object DeviceID,DriverName, PortName | Format-List 
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
                        Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-Table @{Expression={$_.Username};Label="Current User"} 
                        "" 
                        #May take a very long time if on a domain with many users 
                        #"All Users" 
                        #"------------" 
                        #Get-CimInstance -ClassName Win32_UserAccount -ComputerName $compname | foreach{$_.Caption} 
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
                        Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $compname | Format-List @{Expression={$_.Caption};Label="OS Name"},SerialNumber,OSArchitecture 
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
                        Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-List Name,Domain,Manufacturer,Model,SystemType 
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
                        Get-CimInstance -ClassName Win32_Product -ComputerName $compname | Sort-Object Name | Format-Table Name,Vendor,Version 
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
                        Get-CimInstance -ClassName Win32_Process -ComputerName $compname | Select-Object Caption,Handle | Sort-Object Caption | Format-Table 
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
                        Get-CimInstance -ClassName Win32_Service -ComputerName $compname | Select-Object Name,State,Status,StartMode,ProcessID, ExitCode | Sort-Object Name | Format-Table 
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
                        Get-CimInstance -ClassName Win32_USBControllerDevice -ComputerName $compname | ForEach-Object { Get-CimInstance -InputObject $_.Dependent -ComputerName $compname } | Select-Object Caption, Manufacturer, DeviceID | Format-List 
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
                        $wmi = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $compname 
                        $localdatetime = $wmi.LocalDateTime 
                        $lastbootuptime = $wmi.LastBootUpTime 
             
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
                        $wmi = Get-CimInstance -ClassName Win32_logicaldisk -ComputerName $compname 
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
                        $wmi = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $compname 
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
                        Get-CimInstance -ClassName Win32_Processor -ComputerName $compname | Format-List Caption,Name,Manufacturer,ProcessorId,NumberOfCores,AddressWidth   
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
                            $montable | Select-Object-object  -unique Serial,Name,Ascii | Where-Object {$_.Serial -ne "-"} | Format-Table  
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
			
            function clearCache {
                if ($compname.Contains('SRPC')) {
                    Remove-Item "\\$compname\c$\Users\STA$storeNum\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\STA$storeNum.PIERFW1\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\ST$storeNum\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\ST$storeNum.PIERFW1\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\STM$storeNum\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\STM$storeNum.PIERFW1\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                }
                else {
                    Remove-Item "\\$compname\c$\Users\$MODDEDUSER\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                    Remove-Item "\\$compname\c$\Users\$MODDEDUSER.PIERFW1\AppData\Local\Microsoft\Windows\Caches\*.*" -Force -Recurse
                }
            }

            function MrCleanCWPCSFPC {
                #this takes the compname string and extracts the 4-digit store number into a new substring called storeNum
			    $storeNum = $compname.Substring(5,4)

                #setting path names to remote device
                $pathname1 = "\\$compname\c$\Users\STA$storeNum\AppData\local\temp"
                $pathname2 = "\\$compname\c$\Users\STA$storeNum.PIERFW1\AppData\local\temp"
                #$pathname3 = "\\$compname\c$\Users\STA$storeNum\AppData\Local\Packages\feb3bb8e-0ba8-404c-948a-0c1d34a04074-Pier1-Imports_f5fycwrf0hd7g\AC"
                #$pathname4 = "\\$compname\c$\Users\STA$storeNum.PIERFW1\AppData\Local\Packages\feb3bb8e-0ba8-404c-948a-0c1d34a04074-Pier1-Imports_f5fycwrf0hd7g\AC"
                $pathname5 = "\\$compname\c$\Users\ST$storeNum\AppData\local\temp"
                $pathname6 = "\\$compname\c$\Users\ST$storeNum.PIERFW1\AppData\local\temp"
                $pathname7 = "\\$compname\c$\Users\STM$storeNum\AppData\local\temp"
                $pathname8 = "\\$compname\c$\Users\STM$storeNum.PIERFW1\AppData\local\temp"


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
                    New-PSDrive -Name "J" -PSProvider FileSystem -Root $pathname1 -Persist
                    New-PSDrive -Name "K" -PSProvider FileSystem -Root $pathname2 -Persist
                    New-PSDrive -Name "W" -PSProvider FileSystem -Root $pathname5 -Persist
                    New-PSDrive -Name "X" -PSProvider FileSystem -Root $pathname6 -Persist
                    New-PSDrive -Name "Y" -PSProvider FileSystem -Root $pathname7 -Persist
                    New-PSDrive -Name "Z" -PSProvider FileSystem -Root $pathname8 -Persist
                    Start-Sleep 2
                    ""

			        Write-Host "Drive(s) mounted"
			        ""
				
			        #getting file counts
                    #skipping getting file counts if half will always error because of mapped network drive will fail since path won't exist
			        #Write-Host "File count on drive 'J:': " ( Get-ChildItem J:\ | Measure-Object ).Count
			        #Write-Host "File count on drive 'K:': " ( Get-ChildItem K:\ | Measure-Object ).Count
                    #Write-Host "File count on drive 'W:': " ( Get-ChildItem K:\ | Measure-Object ).Count
                    #Write-Host "File count on drive 'X:': " ( Get-ChildItem K:\ | Measure-Object ).Count
                    #Write-Host "File count on drive 'Y:': " ( Get-ChildItem K:\ | Measure-Object ).Count
                    #Write-Host "File count on drive 'Z:': " ( Get-ChildItem K:\ | Measure-Object ).Count
                    #""

                    #clearing temp files
                    Write-Host "Clearing temporary files (Stage 3-of-4)...."
                    Remove-Item "J:\*.*" -Force -Recurse
                    Remove-Item "K:\*.*" -Force -Recurse
                    Remove-Item "W:\*.*" -Force -Recurse
                    Remove-Item "X:\*.*" -Force -Recurse
                    Remove-Item "Y:\*.*" -Force -Recurse
                    Remove-Item "Z:\*.*" -Force -Recurse
                    ""
				
                    #removing network drives
			        Write-Host "Unmounting drive(s) (Stage 4-of-4)....."
			        Get-PSDrive J | Remove-PSDrive
			        Get-PSDrive K | Remove-PSDrive
                    Get-PSDrive W | Remove-PSDrive
                    Get-PSDrive X | Remove-PSDrive
                    Get-PSDrive Y | Remove-PSDrive
                    Get-PSDrive Z | Remove-PSDrive
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
                    New-PSDrive -Name "J" -PSProvider FileSystem -Root $pathname1 -Persist
                    New-PSDrive -Name "K" -PSProvider FileSystem -Root $pathname2 -Persist
                    Start-Sleep 3
                    ""

			        Write-Host "Drive(s) mounted"
			        ""
				
			        #getting file counts
                    #skipping getting file counts if half will always error because of mapped network drive will fail since path won't exist
			        #Write-Host "File Count: " ( Get-ChildItem J:\ | Measure-Object ).Count
			        #Write-Host "File Count: " ( Get-ChildItem K:\ | Measure-Object ).Count
                    #""

                    #clearing temp files
                    Write-Host "Clearing temporary files (Stage 3-of-4)...."
                    Remove-Item "J:\*.*" -Force -Recurse
                    Remove-Item "K:\*.*" -Force -Recurse
                    ""
				
                    #removing network drives
			        Write-Host "Unmounting drive(s) (Stage 4-of-4)....."
			        Get-PSDrive J | Remove-PSDrive
			        Get-PSDrive K | Remove-PSDrive
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
                #$emp = Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-Table @{Expression={$_.Username};Label="Current User"}
                $LOGGEDUSER = Read-Host "Enter the account name or user name (i.e. ST1234, jsmith)"
                #this will take any letters in the string and convert them to uppercase 
                $MODDEDUSER = $LOGGEDUSER.ToUpper()
                ""

                #setting path names to remote device
                $pathname1 = "\\$compname\c$\Users\$MODDEDUSER\AppData\local\temp"
                $pathname2 = "\\$compname\c$\Users\$MODDEDUSER.PIERFW1\AppData\local\temp"

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
                New-PSDrive -Name "J" -PSProvider FileSystem -Root $pathname1 -Persist
                New-PSDrive -Name "K" -PSProvider FileSystem -Root $pathname2 -Persist
                Start-Sleep 2
                ""

				Write-Host "Drive(s) mounted"
				""
				
				#getting file counts
                #skipping getting file counts if half will always error because of mapped network drive will fail since path won't exist
				#Write-Host "File count on drive 'J:': " ( Get-ChildItem J:\ | Measure-Object ).Count
				#Write-Host "File count on drive 'K:': " ( Get-ChildItem K:\ | Measure-Object ).Count
                #""

                #clearing temp files
                Write-Host "Clearing temporary files (Stage 3-of-4)...."
                Remove-Item "J:\*.*" -Force -Recurse
                Remove-Item "K:\*.*" -Force -Recurse
                ""
				
                #removing network drives
				Write-Host "Unmounting drive(s) (Stage 4-of-4)....."
				Get-PSDrive J | Remove-PSDrive
				Get-PSDrive K | Remove-PSDrive
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
                MrClean

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
            Write-Host "Clearing cache and temp files on $compname...."
            ""
            MrClean

            funcPause
            CheckHost
		}		
		4 {
			#reboot pc
			$msg = "We are rebooting this device. Please leave it alone until it has reloaded completely. Thank you."
			Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
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
			Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
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
                #$fileNIRCMDloc = "C:\users\public\nircmd.exe"
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
			    Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
                        Remove-PSSession $connection
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
			Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
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
            #an attempt to speed up a PC (i.e. HP 6005)
            
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
                Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
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
                Get-CimInstance -ClassName Win32_Process -ComputerName $compname | Select-Object Caption,Handle | Sort-Object Caption | Format-Table 
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
                Get-CimInstance -ClassName Win32_Process -ComputerName $compname | Select-Object Caption,Handle | Sort-Object Caption | Format-Table 
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
            GetMenu
            #correct device to auto login (cwpc and sfpc only)

            #copy 3 files over from network drive to pc being worked on
            $filesFROM1 = "\\vmfwsccmdpp1\Deployments\CWPC-SSD-Updates\Production Staging\Fixes\Boot to Screen Prompting Password\AutoLogonSTAuser\Autologon.exe"
            $filesFROM2 = "\\vmfwsccmdpp1\Deployments\CWPC-SSD-Updates\Production Staging\Fixes\Boot to Screen Prompting Password\AutoLogonSTAuser\Autologon.reg"
            $filesFROM3 = "\\vmfwsccmdpp1\Deployments\CWPC-SSD-Updates\Production Staging\Fixes\Boot to Screen Prompting Password\AutoLogonSTAuser\Pier1.ESS.SetAutoLogon.exe"
            $filesTO = "\\$compname\c$\"

            Copy-Item -Path $filesFROM1 -Destination $filesTO
            Copy-Item -Path $filesFROM2 -Destination $filesTO
            Copy-Item -Path $filesFROM3 -Destination $filesTO

            #grabs the username of the current person logged into the computer attempting to use the script
            #$LOGGEDUSER1 = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            #$LOGGEDUSER2 = $LOGGEDUSER1.Substring(8)
            #Invoke-Command -ComputerName $compname -Credential $LOGGEDUSER2 -ScriptBlock { Start-Process "C:\Windows\notepad.exe" } 

            Invoke-Command -ComputerName $compname -ScriptBlock { Start-Process "C:\Windows\notepad.exe" } 

            #returns code showing success but tried on a lab pos and nothing loaded up
            #Invoke-WMIMethod -Class Win32_Process -Name Create -Computername $compname -ArgumentList Notepad.exe


        }
        14 {
            #copy all files and folders on the desktop of the ST, STM, or personal emp
            #then save them to their respective network drives under a newly made, unique folder

            function backupST {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 

			    $storeNum = $compname.Substring(5,4)

                $ServerStore1 = "\\ntfwstrvfsp1\strusers\ST$storeNum"
                #below servers are not the correct ones. AD confirmed server location of ST and STM H:\ drives
                #$ServerStore2 = "\\ntfwcorpvfsp2\strusers\ST$storeNum"
                #$ServerStore3 = "\\ntfwcorpvfsp3\strusers\ST$storeNum"
                #$ServerStore4 = "\\ntfwcorpvfsp4\strusers\ST$storeNum"
                #$ServerStore5 = "\\ntfwcorpvfsp5\strusers\ST$storeNum"
                #$ServerStore6 = "\\ntfwcorpvfsp6\strusers\ST$storeNum"
                #$ServerStore7 = "\\ntfwcorpvfsp7\strusers\ST$storeNum"
                #$ServerStore8 = "\\ntfwcorpvfsp8\strusers\ST$storeNum"
                
                #gets todays date in the order of year month day as numeric and saves it to a string
                $dirDateTime = (Get-Date).ToString("yyyy.MM.dd_HHmm")
                #new folder name plus the date generated from previous string to create unique, dated folder
                $dirNameDateTime = "bopc desktop backup" + " " + $dirDateTime
                #path to store Desktop
                $pathFROMStore1 = "\\$compname\c$\Users\ST$storeNum\Desktop"
                $pathFROMStore2 = "\\$compname\c$\Users\ST$storeNum.PIERFW1\Desktop"

                if ( Test-Path $ServerStore1 -PathType Container ) {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    "Confirmed server and ST network directory both exist. Creating backup folder on the network folder for the ST account...."
                    #making the dir the backup will be stored to
                    New-Item -Path "$ServerStore1\$dirNameDateTime" -type directory -Force

                    #as there are 2 possible directories to copy from, this checks both and only copies from the one found to exist
                    #i've never seen both directories exist on a single PC so there should be a slim chance the second IF will override the first
                    if ( Test-Path $pathFROMStore1 -PathType Container ) {
                        ""
                        Write-Host "Copying everything from the primary ST account Desktop to their newly made network backup folder...."
                        Copy-Item -Path $pathFROMStore1 -Recurse -Destination "$ServerStore1\$dirNameDateTime" -Container
                    }
                    if ( Test-Path $pathFROMStore2 -PathType Container ) {
                        ""
                        Write-Host "Copying everything from the alternative ST account Desktop location to their newly made network backup folder...."
                        Copy-Item -Path $pathFROMStore2 -Recurse -Destination "$ServerStore1\$dirNameDateTime" -Container
                    }

                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    Write-Host "ST Desktop backup copy completed."

                    funcPause
                    backupMenu
                }
                else {
                    
                    Write-Host "Either the server (ntfwstrvfsp1) does not exit or the network folder for the account doesn't. Backup not created."

                    funcPause
                    backupMenu

                }

            }

            function backupSTM {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 

			    $storeNum = $compname.Substring(5,4)

                $ServerStore1 = "\\ntfwstrvfsp1\strusers\STM$storeNum"
                #below servers are not the correct ones. AD confirmed server location of ST and STM H:\ drives
                #$ServerStore2 = "\\ntfwcorpvfsp2\strusers\STM$storeNum"
                #$ServerStore3 = "\\ntfwcorpvfsp3\strusers\STM$storeNum"
                #$ServerStore4 = "\\ntfwcorpvfsp4\strusers\STM$storeNum"
                #$ServerStore5 = "\\ntfwcorpvfsp5\strusers\STM$storeNum"
                #$ServerStore6 = "\\ntfwcorpvfsp6\strusers\STM$storeNum"
                #$ServerStore7 = "\\ntfwcorpvfsp7\strusers\STM$storeNum"
                #$ServerStore8 = "\\ntfwcorpvfsp8\strusers\STM$storeNum"
                
                #gets todays date in the order of year month day as numeric and saves it to a string
                $dirDateTime = (Get-Date).ToString("yyyy.MM.dd_HHmm")
                #new folder name plus the date generated from previous string to create unique, dated folder
                $dirNameDateTime = "bopc desktop backup" + " " + $dirDateTime
                #path to store Desktop
                $pathFROMStore1 = "\\$compname\c$\Users\STM$storeNum\Desktop"
                $pathFROMStore2 = "\\$compname\c$\Users\STM$storeNum.PIERFW1\Desktop"

                if ( Test-Path $ServerStore1 -PathType Container ) {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    "Confirmed server and STM network directory both exist. Creating backup folder on the network folder for the STM account...."
                    #making the dir the backup will be stored to
                    New-Item -Path "$ServerStore1\$dirNameDateTime" -type directory -Force

                    #as there are 2 possible directories to copy from, this checks both and only copies from the one found to exist
                    #i've never seen both directories exist on a single PC so there should be a slim chance the second IF will override the first
                    if ( Test-Path $pathFROMStore1 -PathType Container ) {
                        ""
                        Write-Host "Copying everything from the primary STM account Desktop to their newly made network backup folder...."
                        Copy-Item -Path $pathFROMStore1 -Recurse -Destination "$ServerStore1\$dirNameDateTime" -Container
                    }
                    if ( Test-Path $pathFROMStore2 -PathType Container ) {
                        ""
                        Write-Host "Copying everything from the alternative STM account Desktop location to their newly made network backup folder...."
                        Copy-Item -Path $pathFROMStore2 -Recurse -Destination "$ServerStore1\$dirNameDateTime" -Container
                    }

                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    Write-Host "STM Desktop backup copy completed."

                    funcPause
                    backupMenu
                }
                else {
                    
                    Write-Host "Either the server (ntfwstrvfsp1) does not exit or the network folder for the account doesn't. Backup not created."

                    funcPause
                    backupMenu

                }

            }

            function backupEMP {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 

                #making variables for each server and both directories on each server
                #commented out servers are not the correct ones. AD confirmed server location of emp H:\ drives
                #$ServerEMP1a = "\\ntfwcorpvfsp1\users\$MODDEDUSER"
                #$ServerEMP1b = "\\ntfwcorpvfsp1\users2\$MODDEDUSER"
                #$ServerEMP2a = "\\ntfwcorpvfsp2\users\$MODDEDUSER"
                #$ServerEMP2b = "\\ntfwcorpvfsp2\users2\$MODDEDUSER"
                $ServerEMP3a = "\\ntfwcorpvfsp3\users\$MODDEDUSER"
                $ServerEMP3b = "\\ntfwcorpvfsp3\users2\$MODDEDUSER"
                #$ServerEMP4a = "\\ntfwcorpvfsp4\users\$MODDEDUSER"
                #$ServerEMP4b = "\\ntfwcorpvfsp4\users2\$MODDEDUSER"
                #$ServerEMP5a = "\\ntfwcorpvfsp5\users\$MODDEDUSER"
                #$ServerEMP5b = "\\ntfwcorpvfsp5\users2\$MODDEDUSER"
                #$ServerEMP6a = "\\ntfwcorpvfsp6\users\$MODDEDUSER"
                #$ServerEMP6b = "\\ntfwcorpvfsp6\users2\$MODDEDUSER"
                #$ServerEMP7a = "\\ntfwcorpvfsp7\users\$MODDEDUSER"
                #$ServerEMP7b = "\\ntfwcorpvfsp7\users2\$MODDEDUSER"
                #$ServerEMP8a = "\\ntfwcorpvfsp8\users\$MODDEDUSER"
                #$ServerEMP8b = "\\ntfwcorpvfsp8\users2\$MODDEDUSER"
                
                #making variables for My Documents to backup (Desktop, Documents, Downloads, and Pictures)
                $pathFROMEmpDesktop1 = "\\$compname\c$\Users\$MODDEDUSER\Desktop"
                $pathFROMEmpDesktop2 = "\\$compname\c$\Users\$MODDEDUSER.PIERFW1\Desktop"
                $pathFROMEmpDocuments1 = "\\$compname\c$\Users\$MODDEDUSER\Documents"
                $pathFROMEmpDocuments2 = "\\$compname\c$\Users\$MODDEDUSER.PIERFW1\Documents"
                $pathFROMEmpDownloads1 = "\\$compname\c$\Users\$MODDEDUSER\Downloads"
                $pathFROMEmpDownloads2 = "\\$compname\c$\Users\$MODDEDUSER.PIERFW1\Downloads"
                $pathFROMEmpPictures1 = "\\$compname\c$\Users\$MODDEDUSER\Pictures"
                $pathFROMEmpPictures2 = "\\$compname\c$\Users\$MODDEDUSER.PIERFW1\Pictures"

                #gets todays date in the order of year month day as numeric and saves it to a string
                $dirDateTime = (Get-Date).ToString("yyyy.MM.dd_HHmm")
                #new folder name plus the date generated from previous string to create unique, dated folder
                $dirNameDateTime = "$compname documents backup" + " " + $dirDateTime
                
                if ( Test-Path $ServerEMP3a -PathType Container ) {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    "Confirmed server and network directory both exist. Creating a backup folder on $MODDEDUSER network folder...."
                    #making the dir the backup will be stored to
                    New-Item -Path "$ServerEMP3a\$dirNameDateTime" -type directory -Force

                    #after confirming server 1 and the directory are available, checking which path is used for My Documents.. with or without ".PIERFW1"
                    if ( Test-Path $pathFROMEmpDesktop1 -PathType Container ) {
                        ""
                        Write-Host "Copying documents to the newly made network backup folder...."
                        Copy-Item -Path $pathFROMEmpDesktop1 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDocuments1 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDownloads1 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpPictures1 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container

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
                    if ( Test-Path $pathFROMEmpDesktop2 -PathType Container ) {
                        ""
                        Write-Host "Copying documents to the newly made network backup folder...."
                        Copy-Item -Path $pathFROMEmpDesktop2 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDocuments2 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDownloads2 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpPictures2 -Recurse -Destination "$ServerEMP3a\$dirNameDateTime" -Container

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
                if ( Test-Path $ServerEMP3b -PathType Container ) {
                    Clear-Host
                    "  /----------------------\" 
                    "  |      SLIM TOOL       |" 
                    "  \----------------------/" 
                    "  $compname ($pcip)" 
                    "" 
                    "Confirmed server and network directory both exist. Creating a backup folder on $MODDEDUSER network folder...."
                    #making the dir the backup will be stored to
                    New-Item -Path "$ServerEMP3b\$dirNameDateTime" -type directory -Force

                    #after confirming server 1 and the directory are available, checking which path is used for My Documents.. with or without ".PIERFW1"
                    if ( Test-Path $pathFROMEmpDesktop1 -PathType Container ) {
                        ""
                        Write-Host "Copying documents to the newly made network backup folder...."
                        Copy-Item -Path $pathFROMEmpDesktop1 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDocuments1 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDownloads1 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpPictures1 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container

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
                    if ( Test-Path $pathFROMEmpDesktop2 -PathType Container ) {
                        ""
                        Write-Host "Copying documents to the newly made network backup folder...."
                        Copy-Item -Path $pathFROMEmpDesktop2 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDocuments2 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpDownloads2 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container
                        Copy-Item -Path $pathFROMEmpPictures2 -Recurse -Destination "$ServerEMP3b\$dirNameDateTime" -Container

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
                Write-Host "Either the server (ntfwcorpvfsp3) does not exit or the network folder for the account doesn't. Backup not created."

                funcPause
                backupMenu
            }

            function backupMenu {

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                "1)  Backup ST account only" 
                "2)  Backup STM account only" 
                "3)  Backup an EMP account (i.e. empID 489000 or JSMITH)" 
                "" 
                "X)  Cancel" 
                "" 
                $BACKUPSelection = Read-Host "Enter Selection"
                backupChoice

            }

            function backupChoice {
                Clear-Host 
                switch ($BACKUPSelection){ 
		            1 {
                        backupST
                    }
                    2 {
                        backupSTM
                    }
                    3 {
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        ""
                        $LOGGEDUSER = Read-Host "Enter their Active Directory network username (i.e. jsmith or 012345)"
                        #this will take any letters in the string and convert them to uppercase 
                        $MODDEDUSER = $LOGGEDUSER.ToUpper()
                        backupEMP

                        backupMenu
                    }

                    x {Clear-Host;CheckHost} 
                    default {backupMenu}
                }
            }

            backupMenu
	    }
        order66 {
            # order 66, some fun and useful nircmd commands. i have tested many in win10 and most do not work
            # https://nircmd.nirsoft.net/

            function copyNIRCMD {
                #copy the nircmd.exe file over from network drive to pc being worked on
                #this file is needed for the script to work
                $filesFROM = "I:\Isd\System Support\scripts\tools\nircmd.exe"
                $filesTO = "\\$compname\c$\users\public"
                Copy-Item -Path $filesFROM -Destination $filesTO
                #$fileNIRCMDloc = "C:\users\public\nircmd.exe"
            }

            function removeNIRCMD {
                #cleanup fuction to remove the EXE once task compelted
                Remove-Item -Path "\\$compname\c$\users\public\nircmd.exe"
            }

            function listofNIRCMD {
                #everytime the list loads, the nircmd file is copied over to ensure it is always there
                copyNIRCMD

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                "1)  Open all optical disk trays of remote PC" 
                "2)  Close all optical disk trays of remote PC" 
                "3)  Empty the recycle bin in all drives"
                "4)  Mute the system volume" 
                "5)  Unmute the system volume"
                "" 
                "X)  Cancel" 
                "" 
                $NIRCMDSelection = Read-Host "Enter Selection"
                unleashNIRCMD
            }

            function unleashNIRCMD {
                switch ($NIRCMDSelection){ 
		            1 {
                        #open cd tray(s)
                        copyNIRCMD
                        psexec \\$compname C:\users\public\nircmd.exe cdrom open d:
                        psexec \\$compname C:\users\public\nircmd.exe cdrom open e:
                        psexec \\$compname C:\users\public\nircmd.exe cdrom open f:
                        psexec \\$compname C:\users\public\nircmd.exe cdrom open j:
                        removeNIRCMD

                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        "" 
                        Write-Host "The CD tray(s) have been opened."

                        funcPause
                        listofNIRCMD
                    }

                    2 {
                        #close cd tray(s)
                        copyNIRCMD
                        psexec \\$compname C:\users\public\nircmd.exe cdrom close d:
                        psexec \\$compname C:\users\public\nircmd.exe cdrom close e:
                        psexec \\$compname C:\users\public\nircmd.exe cdrom close f:
                        psexec \\$compname C:\users\public\nircmd.exe cdrom close j:
                        removeNIRCMD
                        
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        "" 
                        Write-Host "The CD tray(s) have been closed."

                        funcPause
                        listofNIRCMD
                    }

                    3 {
                        #empty recycle bin
                        copyNIRCMD
                        psexec \\$compname C:\users\public\nircmd.exe emptybin
                        removeNIRCMD
                        
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        "" 
                        Write-Host "The recycle bin has been emptied."

                        funcPause
                        listofNIRCMD
                    }

                    4 {
                        #mute remote pc volume
                        copyNIRCMD
                        psexec \\$compname C:\users\public\nircmd.exe mutesysvolume 1
                        removeNIRCMD
                        
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        "" 
                        Write-Host "The volume on the remote computer has been muted."

                        funcPause
                        listofNIRCMD
                    }

                    5 {
                        #unmute remote pc volume
                        copyNIRCMD
                        psexec \\$compname C:\users\public\nircmd.exe mutesysvolume 0
                        removeNIRCMD
                        
                        Clear-Host
                        "  /----------------------\" 
                        "  |      SLIM TOOL       |" 
                        "  \----------------------/" 
                        "  $compname ($pcip)" 
                        "" 
                        Write-Host "The volume on the remote computer has been unmuted."

                        funcPause
                        listofNIRCMD
                    }
                    
	                x {Clear-Host; CheckHost} 
	                default{listofNIRCMD} 
                }
            }

            listofNIRCMD
        }
        
        c {Clear-Host; GetCompName} 
        x {Clear-Host; MainMenu} 
        default{CheckHost} 
    }
}

function POSMenu {
    Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
    "  $compname ($pcip)" 
    "" 
    "1)  System overview (General)" 
    "2)  Stop POS services" 
    "3)  Restart POS services" 
    "4)  Resync barcode scanner"
    "5)  Reboot PC (POS version)" 
    "6)  Shutdown PC (POS version)" 
    "7)  Change volume (Requires update to PowerShell. See Brad before using.)"
    "8)  Change clock time zone"
    "9)  Restart printer spooler and clear print queue"
    "10) Stop only the POS Client and Shell"
    "11) Send custom message to PC"
    "" 
    "C)  Switch to a different computer" 
    "X)  Exit to the Main Menu" 
    "" 
    $MenuSelection = Read-Host "Enter Selection" 
    GetPOSInfo 
} 

function GetPOSInfo {
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
			Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-Table @{Expression={$_.Username};Label="Current User"} 
			"------------------------------"
			#OS Info
			Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $compname | Format-List @{Expression={$_.Caption};Label="OS Name"},SerialNumber,OSArchitecture 
			"------------------------------"
			#System Info
			"System Info:"
			Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $compname | Format-List Name,Domain,Manufacturer,Model,SystemType 
			"------------------------------"
			#Uptime
            ""
			$wmi = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $compname 
            $localdatetime = $wmi.LocalDateTime 
            $lastbootuptime = $wmi.LastBootUpTime 
            "Current Time:      $localdatetime" 
            "Last Boot Up Time: $lastbootuptime" 
            $uptime = $localdatetime - $lastbootuptime 
            "Uptime: $uptime" 
            ""
			"------------------------------"
			#Disk Space
            ""
			"Disk Space:"
			$wmi = Get-CimInstance -ClassName Win32_logicaldisk -ComputerName $compname 
            foreach($device in $wmi){ 
                    Write-Host "Drive: " $device.name    
                    Write-Host -NoNewLine "Size: "; "{0:N2}" -f ($device.Size/1Gb) + " Gb" 
                    Write-Host -NoNewLine "FreeSpace: "; "{0:N2}" -f ($device.FreeSpace/1Gb) + " Gb" 
                    "" 
             } 
			"------------------------------"
			#Memory Info
			"Memory Info:"
			$wmi = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $compname 
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
            #stop pos services

            #Epicor Authorization Service
            #Epicor Cash Office Service
            #Epicor Customer Order Service
            #Epicor Customer Service
            #Epicor DataConnect Service
            #Epicor Electronic Journal Service
            #Epicor Employee Service
            #Epicor File Distribution Service
            #Epicor Item Management Service
            #Epicor POS Service
            #Epicor POS Scanner Device Service
            #Epicor Promotion Service
            #Epicor Replication Service
            #Epicor Retail Transaction Service
            #Epicor Stock Service
            #Epicor Store Maintenance Service
            #Epicor Store Management Service
            #Epicor Store Message Queue Manager
            #Epicor System Control Service
            #Epicor Transaction Processing Service
            #Epicor Unique ID Service
            #Secure Data Encryption Certificate Installer
            #Secure Data Service
            #Secure Data Web Service
            #Store Customer Payment
            #Store SAF Service
            #NsbCustomerRelationshipManagementService
            #MSSQL$SQLEXPRESS

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "Killing the POS Client (Stage 1-of-2)...."
            Invoke-Command -computername $compname -ScriptBlock {Stop-Process -Name "NSB.POS.Client" -Force -Verbose -ErrorAction SilentlyContinue}

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			Write-Host "Stopping all necessary services (Stage 2-of-2)...."
            ""
            #this will get all processes starting with 'Epicor' on a remote pos and stop each processes
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Epicor")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Secure Data")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Store")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("NsbCustomer")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("MSSQL")} | Stop-Service -Force

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "All necessary services have been stopped."

            funcPause
            CheckHost
        }
        3 {
            #restart pos services

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "Killing the POS Client (Stage 1-of-4)...."
            Invoke-Command -computername $compname -ScriptBlock {Stop-Process -Name "NSB.POS.Client" -Force -Verbose -ErrorAction SilentlyContinue}

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			Write-Host "Stopping all necessary services (Stage 2-of-4)...."
            ""
            #this will get all processes starting with 'Epicor' on a remote pos and stop each processes
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Epicor")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Secure Data")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Store")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("NsbCustomer")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("MSSQL")} | Stop-Service -Force

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "All necessary services have been stopped. Waiting before launching...."
            Start-Sleep -s 8

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "Starting all necessary services (Stage 3-of-4)...."

            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("MSSQL")} | Start-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("NsbCustomer")} | Start-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Store")} | Start-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Secure Data")} | Start-Service -Force

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "Requesting to start POS Client (Stage 4-of-4)...."
            $msg = "The POS services have been reset. Please press the 'Start POS' button to finish loading the register."
            Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
            Write-Host "POS services reset completed."
            Write-Host "Have the caller press the 'Start POS' button on the register to complete loading the POSClient."

            funcPause
            CheckHost
        }
        4 {
            #resync the pos barcode scanner

            #code copied from Scanner-ResyncP1.ps1 and modified to my variable names
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
		    ""
            Write-Host "Connecting...."
	        Start-Sleep 5
            ""
            Invoke-Command -computername $compname -ScriptBlock {Stop-Process -Name "NSB.POS.Client"  -Force -Verbose -ErrorAction SilentlyContinue
	            Stop-Service -name "Epicor POS Service","Epicor POS Scanner Device Service"
	            Start-Sleep 10
	            Start-Service -name "Epicor POS Service"
	            Start-Sleep 15
	            Start-Service -name "Epicor POS Scanner Device Service"
	            Start-Sleep 15
            }

            #old way to start the shell and posclient that doesn't work this way
            #Run-Shell -machine $compname | out-null
            #Run-POSClient -machine $compname | out-null
            
            $msg = "The scanner service has been reset. Please press the 'Start POS' button to finish loading the register."
            #sends msg to the pos to have the caller press the Start POS button
            Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
		    ""
            Write-Host "POS services reset completed."
            Write-Host "Have the caller press the 'Start POS' button on the register to complete loading the POSClient."

            funcPause
            CheckHost
        }
        5 {
            #reboot pos

            $msg = "We are rebooting this device. Please leave it alone until it has reloaded completely. Thank you."
			Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			Write-Host "Stopping all necessary services"

            #this will get all processes starting with 'Epicor' on a remote pos and stop each processes
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Epicor")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Secure Data")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Store")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("NsbCustomer")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("MSSQL")} | Stop-Service -Force

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
		    ""
            Write-Host "Sending reboot signal"
            shutdown /r /m \\$compname /t 5
            ""
			Write-Host "Remote reboot signal sent to $compname"

			funcPause
			CheckHost
        }
        6 {
            #shutdown pos

            $msg = "We are shutting down this device. Please leave it alone until it has shut down completely. Thank you."
			Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
			""
			Write-Host "Stopping all necessary services"

            ""
            #this will get all processes starting with 'Epicor' on a remote pos and stop each processes
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Epicor")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Secure Data")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("Store")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("NsbCustomer")} | Stop-Service -Force
            Get-Service -ComputerName $compname | Where-Object {$_.displayName.StartsWith("MSSQL")} | Stop-Service -Force

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
		    ""
            Write-Host "Sending shutdown signal"
            shutdown /s /m \\$compname /t 5
            ""
			Write-Host "Remote shutdown signal sent to $compname"

			funcPause
			CheckHost
        }
        7 {
            #change volume

            function copyNIRCMD {
                #copy the nircmd.exe file over from network drive to pc being worked on
                #this file is needed for the script to work
                $filesFROM = "I:\Isd\System Support\scripts\tools\nircmd.exe"
                $filesTO = "\\$compname\c$\users\public"
                Copy-Item -Path $filesFROM -Destination $filesTO
                #$fileNIRCMDloc = "C:\users\public\nircmd.exe"
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
                [math]::Round($volConvert)

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Attempting to unmute device (Stage 1-of-2)...."
                # 0 to unmute; 1 to mute ; 2 to toggle mute/unmute
                psexec \\DTMROYE C:\nircmd.exe mutesysvolume 0 

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                Write-Host "Attempting to set device master volume to $volSelection% (Stage 2-of-2)...." 
                #to change volume percent
                psexec \\DTMROYE C:\nircmd.exe setsysvolume $volConvert 
                removeNIRCMD

                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                "  $compname ($pcip)" 
                "" 
                Write-Host "Device volume set to $volSelection%."
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
                $volSelection = Read-Host "Enter percent (0-100)"

                if ($volSelection -eq "") {
                    CheckHost
                }
                if ($volSelection -ge 0) {
                    if ($volSelection -le 100) {
                        setVOLUME
                        askPERCENT
                    }
                    else {
                        ""
                        Write-Host "Please input a percent value between 0-100"
                        funcPause
                        askPERCENT
                    }
                }
                else {
                    ""
                    Write-Host "Please input a percent value between 0-100"
                    funcPause
                    askPERCENT
                }
            }
            askPERCENT
        }
        8 {
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
            
            function listofTZPOS {
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
                changingTZPOS
            }
            
            function changingTZPOS {
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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Hawaiian Standard Time on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Alaskan Standard Time on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Pacific Standard Time on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Arizona on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Mountain Standard Time on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Central Standard Time on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Saskatchewan on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Eastern Standard Time on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Indiana (East) on $compname."

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
                        Remove-PSSession $connection
                        Start-Sleep -s 2
                        Write-Host "Time zone set to Atlantic Time (Canada) on $compname."

                        funcPause
                        CheckHost
                    }
                    x {Clear-Host;CheckHost} 
                }
            }
            listofTZPOS
		}
        9 {
			#restart printer spooler and clear queue

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
        10 {
            #kill the pos client and shell only
            
            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
            ""
            Write-Host "Stopping the POS Client and Shell...."
            ""
            Invoke-Command -computername $compname -ScriptBlock {Stop-Process -Name "NSB.POS.Client"  -Force -Verbose -ErrorAction SilentlyContinue
                Start-Sleep -s 5
                Stop-Process -Name "NSB.Shell"  -Force -Verbose -ErrorAction SilentlyContinue
            }

            Clear-Host
            "  /----------------------\" 
            "  |      SLIM TOOL       |" 
            "  \----------------------/" 
            "  $compname ($pcip)" 
            ""
            Write-Host "POS Client and the Shell had been stopped."

            funcPause
            CheckHost
        }
        11 {
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
                Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine="msg * $msg"} -ComputerName $compname
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
        
        c {Clear-Host;$compname="";GetCompName} 
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
            "3)  [COMING SOON] Clear groupPriority Attribute field"
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
                        Set-ADAccountPassword $empIDUPPER -NewPassword $newpwd -Reset
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
                3 {
                    empMenu

                    #clear attribute groupPriority
                    Set-ADUser $empIDUPPER -groupPriority $null
                    ""
                    Read-Host -Prompt "The groupPriority attribute has been cleared"
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

function ProdCtrl {

    function getDuration {
        $compname = $env:computername
        $ping = Get-CimInstance -ClassName Win32_PingStatus -Filter "Address='$compname'"
        $pcip = $ping.ProtocolAddress
        Clear-Host
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
            $timeStart = Get-Date -Date (Get-Date -Format "yyyy-MM-dd HH:mm")
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
        
        #this fuction has the do loop so that every time the loop is run, it gets the 3 folder names and their latest time stamps to display

        Do { 
            #at the start of each loop, gets the most recent time to compare against set loop end time
            $TimeNow = Get-Date -Date (Get-Date -Format "yyyy-MM-dd HH:mm")
            $recentfileTOKEN = ""
            $recentfilePOLL1 = ""
            $recentfilePOLL2 = ""
            $recentfilePOLL3 = ""
            
            if ($TimeNow -le $TimeEnd) {
                
                ##### getting TokenData and PollData folder data #####
                #get-childitem path | where filter only folders | Sort-Object most recent folder to top | Select-Object the first folder
                $folderTOKEN = Get-ChildItem \\vmfwposaomp3\EPICOR\aw_pier1\TokenData | Where-Object { $_.PSIsContainer } | Sort-Object CreationTime -desc | Select-Object -f 1
                $pathTOKEN = "\\vmfwposaomp3\EPICOR\aw_pier1\TokenData" + "\$folderTOKEN"
                #gets the file name of the most recent file in the folder; then takes it and outputs the date as a string to be usable in IF statement
                $fileTOKEN = Get-ChildItem $pathTOKEN | Sort-Object CreationTime -desc | Select-Object -f 1
                $recentfileTOKEN = $fileTOKEN.LastWriteTime

                #get-childitem path | where filter only folders | Sort-Object most recent folder to top | Select-Object the first folder
                $folderPOLL1 = Get-ChildItem \\vmfwposaomp3\EPICOR\aw_pier1\PollData | Where-Object { $_.PSIsContainer } | Sort-Object CreationTime -desc | Select-Object -f 1
                $pathPOLL1 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL1"
                $filePOLL1 = Get-ChildItem $pathPOLL1 | Sort-Object CreationTime -desc | Select-Object -f 1
                $recentfilePOLL1 = $filePOLL1.LastWriteTime

                #does same as above but skips the first folder to grab the second folder name
                $folderPOLL2 = Get-ChildItem \\vmfwposaomp3\EPICOR\aw_pier1\PollData | Where-Object { $_.PSIsContainer } | Sort-Object CreationTime -desc | Select-Object-object -skip 1 | Select-Object-object -f 1
                $pathPOLL2 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL2"
                $filePOLL2 = Get-ChildItem $pathPOLL2 | Sort-Object CreationTime -desc | Select-Object -f 1
                $recentfilePOLL2 = $filePOLL2.LastWriteTime

                #does same as above but skips the first 2 folders to grab the third folder name
                $folderPOLL3 = Get-ChildItem \\vmfwposaomp3\EPICOR\aw_pier1\PollData | Where-Object { $_.PSIsContainer } | Sort-Object CreationTime -desc | Select-Object-object -skip 2 | Select-Object-object -f 1
                $pathPOLL3 = "\\vmfwposaomp3\EPICOR\aw_pier1\PollData" + "\$folderPOLL3"
                $filePOLL3 = Get-ChildItem $pathPOLL3 | Sort-Object CreationTime -desc | Select-Object -f 1
                $recentfilePOLL3 = $filePOLL3.LastWriteTime

                #creating a checksum for each dir so a var will hold the last captured timestamp in case the following loop returns null
                if ($null -ne $recentfileTOKEN) {
                    $recentfileTOKENcheck = $recentfileTOKEN
                }
                if ($null -ne $recentfilePOLL1) {
                    $recentfilePOLL1check = $recentfilePOLL1
                }
                if ($null -ne $recentfilePOLL2) {
                    $recentfilePOLL2check = $recentfilePOLL2
                }
                if ($null -ne $recentfilePOLL3) {
                    $recentfilePOLL3check = $recentfilePOLL3
                }

                #now comparing the most recent run against the last known previous captured timestamp
                #if the latest lopp is null but the checksum isn't, the checksum overrights the latest run var
                #if the latest lopp is null and the checksum is too, the latest run var is filled with the statement of no no files to compare to
                if ($null -eq $recentfileTOKEN) {
                    if ($null -ne $recentfileTOKENcheck) {
                        $recentfileTOKEN = $recentfileTOKENcheck
                    }
                    else {
                        $recentfileTOKEN = "No recent files created"
                    }
                }
                if ($null -eq $recentfilePOLL1) {
                    if ($null -ne $recentfilePOLL1check) {
                        $recentfilePOLL1 = $recentfilePOLL1check
                    }
                    else {
                        $recentfilePOLL1 = "No recent files created"
                    }
                }
                if ($null -eq $recentfilePOLL2) {
                    if ($null -ne $recentfilePOLL2check) {
                        $recentfilePOLL2 = $recentfilePOLL2check
                    }
                    else {
                        $recentfilePOLL2 = "No recent files created"
                    }
                }
                if ($null -eq $recentfilePOLL3) {
                    if ($null -ne $recentfilePOLL3check) {
                        $recentfilePOLL3 = $recentfilePOLL3check
                    }
                    else {
                        $recentfilePOLL3 = "No recent files created"
                    }
                }

                ##### getting BulkInventoryLoad folder data #####
                $pathBULK = "\\pier1.com\filedrop\Demandware\Production\Outgoing\inventory\BulkInventoryLoad"
                $folderBULKcount = Get-ChildItem -Path $pathBULK | Measure-Object | ForEach-Object {$_.Count}

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

                $filesBULK_sizes = Get-ChildItem -Path $pathBULK | Select-Object Name, @{Name="Size";Expression={Format-FileSize($_.Length)}}

                function timeCompare_Active {
                    
                    #this will make the line of text
                    #$text = " this is a test of the system "
                    #write-host $text -ForegroundColor Black -BackgroundColor Red

                    ### make all variables in here global by doing $global:VariableName so they can be used in the outputLists function when writing to the screen


                    if (($null -ne $recentfileTOKEN) -and ($recentfileTOKEN -ne "No recent files created")) {
                        if (($null -ne $recentfilePOLL1) -and ($recentfilePOLL1 -ne "No recent files created")) {
                            if (($null -ne $recentfilePOLL3) -and ($recentfilePOLL3 -ne "No recent files created")) {

                                #compares the timespan between a start/end date
                                #if the result is a negative, it flips the variables to recalc and get a positive output

                                ### Token / Poll1
                                $timespan_TokenPoll1 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL1
                                if ($timespan_TokenPoll1 -ge 0) {
    
                                    # DO SOMETHING
                                    if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54)) {
                                        #####
                                    }
                                }
                                if ($timespan_TokenPoll1 -lt 0) {
                                    $timespan_TokenPoll1 = New-Timespan -Start $recentfilePoll1 -End $recentfileToken
    
                                    # DO SOMETHING
                                    if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54)) {
                                        #####
                                    }
                                }

                                ### Token / Poll3
                                $timespan_TokenPoll3 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL3
                                if ($timespan_TokenPoll3 -ge 0) {
    
                                    # DO SOMETHING
                                    if (($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                        #####
                                    }
                                }
                                if ($timespan_TokenPoll3 -lt 0) {
                                    $timespan_TokenPoll3 = New-Timespan -Start $recentfilePoll3 -End $recentfileToken
    
                                    # DO SOMETHING
                                    if (($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                        #####
                                    }
                                }

                                ### Poll1 / Poll3
                                $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll1 -End $recentfilePOLL3
                                if ($timespan_Poll1Poll3 -ge 0) {
    
                                    # DO SOMETHING
                                    if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                        #####
                                    }
                                }
                                if ($timespan_Poll1Poll3 -lt 0) {
                                    $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll3 -End $recentfilePoll1
    
                                    # DO SOMETHING
                                    if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                        #####
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
                    Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                }
                if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("IP") )) {
                    Write-Host "PollData   " $folderPOLL1 "   " $recentfilePOLL1
                }
                if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR") )) {
                    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                    $folderPOLL1 = $folderPOLL2
                    $folderPOLL1_str = $folderPOLL1.ToString()
                    $recentfilePOLL1 = $recentfilePOLL2
                    
                    if ($folderPOLL1_str.StartsWith("AW3")) {
                        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                        Write-Host "PollData   " $folderPOLL1 "   " $recentfilePOLL1
                    }
                }
                if (($folderPOLL1_str.StartsWith("AWL")) -and ($folderPOLL1_str.EndsWith("TR1") )) {
                    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                    $folderPOLL1 = $folderPOLL2
                    $folderPOLL1_str = $folderPOLL1.ToString()
                    $recentfilePOLL1 = $recentfilePOLL2
                    
                    if ($folderPOLL1_str.StartsWith("AW3")) {
                        Write-Host "PollData   " $folderPOLL1 " " $recentfilePOLL1
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                        Write-Host "PollData   " $folderPOLL1 "   " $recentfilePOLL1
                    }
                }

                ### folder 3
                $folderPOLL3_str = $folderPOLL3.ToString()
                if ($folderPOLL3_str.StartsWith("AW3")) {
                    Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                }
                if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                    Write-Host "PollData   " $folderPOLL3 "   " $recentfilePOLL3
                }
                if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR") )) {
                    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                    $folderPOLL3 = $folderPOLL2
                    $folderPOLL3_str = $folderPOLL3.ToString()
                    $recentfilePOLL3 = $recentfilePOLL2
                    
                    if ($folderPOLL3_str.StartsWith("AW3")) {
                        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                        Write-Host "PollData   " $folderPOLL3 "   " $recentfilePOLL3
                    }
                }
                if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("TR1") )) {
                    ### because the creation time gap from the AWL/IP's and AWL/TR and AWL/TR1's jumps dramatically when the new AW3 is made,
                    ### when this happens the AWL/TR and AWL/TR1 from the selection will be swapped by the 2nd in the list
                    $folderPOLL3 = $folderPOLL2
                    $folderPOLL3_str = $folderPOLL3.ToString()
                    $recentfilePOLL3 = $recentfilePOLL2
                    
                    if ($folderPOLL3_str.StartsWith("AW3")) {
                        Write-Host "PollData   " $folderPOLL3 " " $recentfilePOLL3
                    }
                    if (($folderPOLL3_str.StartsWith("AWL")) -and ($folderPOLL3_str.EndsWith("IP") )) {
                        Write-Host "PollData   " $folderPOLL3 "   " $recentfilePOLL3
                    }
                }

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
                        if (($null -ne $recentfileTOKEN) -and ($recentfileTOKEN -ne "No recent files created")) {
                            if (($null -ne $recentfilePOLL1) -and ($recentfilePOLL1 -ne "No recent files created")) {
                                if (($null -ne $recentfilePOLL3) -and ($recentfilePOLL3 -ne "No recent files created")) {

                                    #compares the timespan between a start/end date
                                    #if the result is a negative, it flips the variables to recalc and get a positive output

                                    ### Token / Poll1
                                    $timespan_TokenPoll1 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL1
                                    if ($timespan_TokenPoll1 -ge 0) {
    
                                        # DO SOMETHING
                                        if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            #opens explorer to the listed directories
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }
                                    if ($timespan_TokenPoll1 -lt 0) {
                                        $timespan_TokenPoll1 = New-Timespan -Start $recentfilePoll1 -End $recentfileToken
    
                                        # DO SOMETHING
                                        if (($timespan_TokenPoll1.Days -gt 0) -or ($timespan_TokenPoll1.Hours -gt 0) -or ($timespan_TokenPoll1.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }

                                    ### Token / Poll3
                                    $timespan_TokenPoll3 = New-Timespan -Start $recentfileTOKEN -End $recentfilePOLL3
                                    if ($timespan_TokenPoll3 -ge 0) {
    
                                        # DO SOMETHING
                                        if (($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }
                                    if ($timespan_TokenPoll3 -lt 0) {
                                        $timespan_TokenPoll3 = New-Timespan -Start $recentfilePoll3 -End $recentfileToken
    
                                        # DO SOMETHING
                                        if (($timespan_TokenPoll3.Days -gt 0) -or ($timespan_TokenPoll3.Hours -gt 0) -or ($timespan_TokenPoll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }

                                    ### Poll1 / Poll3
                                    $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll1 -End $recentfilePOLL3
                                    if ($timespan_Poll1Poll3 -ge 0) {
    
                                        # DO SOMETHING
                                        if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\PollData
                                            $Global:warningDelay = $TimeNow.AddMinutes(30)
                                        }
                                    }
                                    if ($timespan_Poll1Poll3 -lt 0) {
                                        $timespan_Poll1Poll3 = New-Timespan -Start $recentfilePoll3 -End $recentfilePoll1
    
                                        # DO SOMETHING
                                        if (($timespan_Poll1Poll3.Days -gt 0) -or ($timespan_Poll1Poll3.Hours -gt 0) -or ($timespan_Poll1Poll3.Minutes -gt 54)) {
                                            msgboxTimeAlert
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\TokenData
                                            Invoke-Item \\vmfwposaomp3\EPICOR\aw_pier1\PollData
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
        } Until ($TimeNow -ge $TimeEnd)
    }

    function msgbox {
        $msgBoxInput =  [System.Windows.MessageBox]::Show('The monitoring script has completed its run time. Would you like to continue monitoring?','Monitoring completed','YesNo','Information')

        switch  ($msgBoxInput) {
            'Yes' {
                getDuration
            }

            'No' {
                MainMenu
            }
        }
    }

    getDuration
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
            empList
        }
        4 {

            function pulldeviceList {
                Set-Location c:\
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                ""
                "Searching for all devices with '$storeNum' in the device name...."
                Get-ADComputer -properties Name -Filter * | Where-Object {$_.Name -match $storeNum} | Format-List -Property Name
                Read-Host -Prompt "Press Enter to continue"
                getStoreNum
            }

            function getstoreNum {
                Clear-Host
                "  /----------------------\" 
                "  |      SLIM TOOL       |" 
                "  \----------------------/" 
                ""
                Write-Host "Input the 4-digit store number to list all devices at that store."
                Write-Host "Leave blank and press Enter to go back to the main menu."
                ""
                $storeNum = Read-Host "4-digit store number"
                $numLength = $storeNum.Length

                if ($storeNum -eq "") {
                    MainMenu
                }
                if ($numLength -eq "4") {
                    pulldeviceList
                }
                ""
                Write-Host "You did not enter a 4-digit number. Please try again."
                Read-Host -Prompt "Press Enter to continue"
                getstoreNum
            }

            getstoreNum
        }
        5 {
            ProdCtrl
        }

        x {
            Invoke-Expression (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
            exit
        }
        default{MainMenu}
    }
}

function MainMenu {
    #PSSpeak
    Clear-Host
    "  /----------------------\" 
    "  |      SLIM TOOL       |" 
    "  \----------------------/" 
    ""
    "1)  Connect to a computer" 
    "2)  Employee account info"
    "3)  I.T. service desk employee list"
    "4)  List all devices at a store"
    "5)  Production Control"
    ""
    "X)  Exit the program" 
    ""
    $mainChoice = Read-Host "Enter Selection"
    lifePath
}

function empList {
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

function PSSpeak {
    #Add-Type -AssemblyName System.speech
    #$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
    #$speak.Speak('Hello, Dave. I am HAL 9000')

    Add-Type -AssemblyName System.speech
    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $tts.Rate   = -5  # -10 to 10; -10 is slowest, 10 is fastest
    $speak.Speak('Hello, Dave. I am HAL 9000')
}

function RickRoll {
    Invoke-Expression (New-Object Net.WebClient).DownloadString("http://bit.ly/e0Mw9w")
    Clear-Host
    Write-Host "You've been Rick Rolled!"
}

#---------Start Main-------------- 
Clear-Host
MainMenu
