' ===================================================   
'Script to install a program and change the environment on a local computer
' Start with DEBUG or RUN
' Creates a logfile in C:\Autodesk\InstallLog\...

'ToDo- means that it is not implemented yet
	 
' Author: Marc Sleegers
' Created: 2016-12-11
' ===================================================  
		
	Dim fso, LOGFILE, LOGFILE2, Wshshell, APO, CM2012Logfile
	DIM GDIRGLOBAL, GDIR, COMMAND, ADSK_DDIRGLOBAL, DDIR
	Dim ReplicationFolderCount, Temp_Folder, G_Base, Copy_Action, ScriptVersion, adsk_User_Who_Started_the_Script
	Dim strSafeDate, strSafeTime, strDateTime
	Dim languagePacks 
	' Create list supported Language Packs
	Set languagePacks = CreateObject( "System.Collections.ArrayList" )
	' Add languages
	languagePacks.Add "ENU" 
	
	ScriptVersion = "3.0 2017-05-05"
		
	Set wshShell = CreateObject( "WScript.Shell" )
	 
	
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	set net = Wscript.CreateObject("Wscript.Network")		
	Set oWMI = GetObject("winmgmts:" & "{impersonationLevel=Impersonate}!\\.\root\cimv2")
	Set objArgs = WScript.Arguments
	IF objArgs.Count = 0 THEN
		RC = Msgbox ("Syntax in script seems right but" & vbcrlf & "You have to start the script with: " & vbcrlf & Wscript.Scriptname  & " RUN" & vbcrlf & "or" & vbcrlf & Wscript.Scriptname  & " DEBUG" & vbcrlf & "DEBUG just creates the logfile and test your syntax. No changes is done", 64, "Version=" & ScriptVersion)		
		Wscript.quit 99
	ELSE
		Run_Mode = uCase(objArgs(0))
		IF Run_mode <> "RUN" and Run_Mode <> "DEBUG" then
			RC = Msgbox ("Syntax in script seems right but" & vbcrlf & "You have to start the script with: " & vbcrlf & Wscript.Scriptname  & " RUN" & vbcrlf & "or" & vbcrlf & Wscript.Scriptname  & " DEBUG"  & vbcrlf & "DEBUG just creates the logfile and test your syntax. No changes is done", 64, "Version=" & ScriptVersion)
			Wscript.quit 99
		END IF
	END IF
	
	IF objArgs.Count = 2 THEN
		Language_pack = uCase(objArgs(1))
	ELSE
		Language_pack = "ENU"
	END IF
	
	APO = Chr(34)   ' = "
	
	strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
	strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
	'Set strDateTime equal to a string representation of the current date and time, for use as part of a valid Windows filename
	strDateTime = strSafeDate & "-" & strSafeTime
	
' ----------------  Start Edit after this line -----------------------------------------------------	
' ----------------  Start Edit after this line -----------------------------------------------------	
' ----------------  Start Edit after this line -----------------------------------------------------	
' ----------------  Start Edit after this line -----------------------------------------------------	


'Name log file (Mandatory) the Path is C:\Autodesk\InstallLog
	LOGFILEPATH = "C:\Autodesk\InstallLog"
	LOGFILENAME =  "Set License server autodesknlm"
	
'Init the functions	(Mandatory)
	Init_functions LOGFILENAME  
' Check required language pack
	If Not languagePacks.Contains(Language_pack) Then
		StatusMessage(" Required Language Pack (" & Language_pack & ") is not included in this installation. The English version will be installed")
		Language_pack = "ENU"
	End If	
	
' clean old registry keys
  adsk_CleanUpLicenseRegKeys "autodesknlm.unilever.com", "autodesknlm.unilever.com", "Clean old License server Regkey values"	
	
' Adding FLEXLM License Manager keys
    adsk_SetLicenseRegKeys "autodesknlm.unilever.com", "autodesknlm.unilever.com", "Adding License server Regkey values"

msgbox "Addint FLEXLM License"

' Restarting computer after installation
   ' adsk_restart_computer 30, "Computer is restarting in 30 seconds, save your work now" , "Restarting due to installation new software"
		
' -------------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------------
 ' !!! =========== END OF CHANGE ======================
 ' !!! =========== END OF CHANGE ======================
 ' !!! =========== END OF CHANGE ======================
 ' !!! =========== END OF CHANGE ======================
 ' !!! =========== END OF CHANGE ======================
 
 
'                  Do NOT change below this line	
' -----------------------------------------------------------------						
' -----------------------------------------------------------------										 	
	
	' Run_Mode = RUN or DEBUG

	
	'Copy_Action can be "FORCE", "COPY", ("UPDATE", "DELETE", "MIRROR" Next version maybe)
	' FORCE = Copy even if local folders not exist = create the folders and copy files
	' COPY = Just copy if the local folder exist
	
'                                         ---- Functions You can use below this line --------------
'                                         All commands ends with a string that will end upp in the LOG file

'---	'Command: adsk_add_Registry_Logentry_installed
		'Comment: Place a log entry in HKLM\Software\Wow6432node\WSP\SWC_Run that the software is installed
		'Syntax: adsk_add_Registry_Logentry_installed "ProgramID/Name", "Programversion", "Yes/no","TheLogMessage"
		'Example: adsk_add_Registry_Logentry_installed "Dolly", "Version 2.0", "Yes","Adding WSP log entry for Mandatory program Dolly to Registry" 

'todo	'Command: adsk_add_Registry_Logentry_Remove  NOT WORKING YET
		'Comment: Remove an entry when You uninstall a program
		'Syntax: adsk_add_Registry_Logentry_Remove "ProgramID","TheLogMessage"
		'Example: adsk_add_Registry_Logentry_Remove "Dolly", "Removing WSP log entry for program Dolly from Registry"

'--- 	'Command: adsk_Folder_Copy
		'Comment: Copy folders including files and subfolders
		'Syntax: adsk_Folder_Copy The folderToCopy, PathToTheToFolder, "FORCE"/"COPY", TheLogMessage
		'Example: adsk_Folder_Copy "Symbols&Fonts", "c:\temp\", "FORCE", "Copy files for Fonts&Symbols"
		
'--- 	'Command: adsk_Folder_Create
'	 	'Comment: Create folders including subfolders
		'Syntax: adsk_Folder_Create TheFolderToCreate, TheLogMessage
		'Example: adsk_Folder_Create "c:\Program Files (x86)\Common Files\DataEast\Licenses", "Creating folder for Xtools License files"
	
'--- 	'Command: adsk_Folder_Delete
'     	'Comment: Delete folder and subfolders and files within them
		'Syntax: adsk_Folder_Delete TheFolderToDelete, TheLogMessage
		'Example: adsk_Folder_Delete "C:\ArcGIS10.1", "Deleting old Arcgis10.1 root folder if found"

'--- 	'Command: adsk_Folder_Check_Exist
		'Comment: Check if a folder exists or not. It it a function that returns True or False
		'Syntax: adsk_Folder_Check_Exist TheFolderName,TheMessage
		'Example: IF adsk_Folder_Check_Exist("c:\_WSPdata","Check if c:\_WSPdata exists") THEN adsk_File_Delete("C:\Progradata\WSP", "kalle.txt" ,"Deleting Kalle.txt")
		
'--- 	'Command: adsk_File_Copy
		'Comment: Copy files with wildcards like *.doc, CopyAction UPDATE och COPY
		'Syntax: adsk_File_Copy FolderToCopyFrom, FilesToCopy , FolderToCopyTo, CopyAction, TheLogMessage
		'Example: adsk_File_Copy "Standardmallar", "*.*", "c:\Arcgis10.3\Desktop10.3\MapTemplates\WSP", "COPY", "Copy Standardmallar"
		
'--- 	'Command: adsk_File_Check_Exist
		'Comment: Check if a file exists or not. It it a function that returns True or False
		'Syntax: adsk_File_Check_Exist TheFileName,TheMessage
		'Example: IF adsk_File_Check_Exist("c:\_WSPdata\apa.txt","Check if apa.txt exists") THEN adsk_File_Delete("C:\Progradata\WSP","kalle.txt","Removing Kalle.txt")
	
'---	'Command: adsk_File_Delete
		'Comment: Delete files
		'Syntax: adsk_File_Delete ThePathToFiles, TheFiles, TheLogMessage
		'Example: adsk_File_Delete "c:\Program Files (x86)\Common Files\DataEast\Licenses", "*.licinfo","Removing old Xtools license files"
	 
'ToDo- 	'Command: adsk_File_Get_File_Info
		'Comment: Get info about a files (TheWantedFileInfo = "CREATIONDATE" | "DATELASTACCESSED" | "DATELASTMODIFIED" | "READONLY" | "HIDDEN" | "SYSTEM"
		'Comment: The Command gives a return value that can be used in an IF clause. 
		'Syntax: adsk_File_Get_File_Info TheWantedFileInfo, TheLogMessage
		'Example: TheFileversion = adsk_File_Get_File_Info ("C:\folder\file.txt", "DATELASTMODIFIED" , "Checking Lastmodified date on File.txt")
		'Example: If (adsk_File_Get_File_Info "C:\folder\file.txt", "DATELASTMODIFIED" , "Checking Lastmodified date on File.txt" = "2016-05-19" Then adsk_File_Delete "C:\folder\file.txt", "Delting File"
	
'--- 	'adsk_Get_File_Version
		'Comment:   Get the version of a file. It is a function and returns the version of a file
		'Syntax: adsk_Get_File_Version :  (TheFile2Check)
		'Example: adsk_Get_File_Version ("c:\Program Files (x86)\ProgramX\Program.exe")
	
'--- 	'adsk_File_Check_Version
		'Comment:   Check the file version of a file. It is a function and returns True or False
		'			EqualOrHigher possible values 0 or 1. If 1 function will return True if current version >= required version
		'Syntax: adsk_File_Check_Version :  (TheFile2Check, TheFileVersion, TheLogMessage, EqualOrHigher)
		'Example: IF (adsk_File_Check_Version ("c:\Program Files (x86)\ProgramX\Program.exe" , "2.1", "Checking if Program.exe is the right version") Then 
						'adsk_file_copy ......
				  'END IF
				  
 '--- 	'Commnad: adsk_File_Rename
 	 	'Comment: Rename a file
 		'Syntax: adsk_File_Rename TheOldFileName,TheNewFileName, TheLogMessage		
 	 	'Example: adsk_File_Rename "TheOldFileName.txt", "TheNewFileName.txt", "Renaming the file TheOldFileName.txt to TheNewFileName.txt"			  
	
'--- 	'Command: adsk_Registry_Add_Regfile
		'Comment: Adding a .REG file to the registry
		'Syntax: adsk_Registry_Add_Regfile TheRegfile, TheOption, TheLogMessage			'TheOption can be /reg:64 eller /reg:32  (32 goes to Wow6432node)
		'Example: adsk_Registry_Add_Regfile "\\se.wspgroup.com\deployment\Media\ForDomainUsers\CQI\R19G\HKCU_TestRegFile.reg", "/reg:64", "Registry Change:"
	
'---	'Command: adsk_Registry_AddKey 
		'Comment: Add a Key to the registry. TheHive is HKLM/HKCU/HKCR
		'Syntax: adsk_Registry_AddKey TheHive, TheBaseKey, TheKeyName, TheLogMessage
		'Example adsk_Registry_AddKey "HKLM", "SOFTWARE\FLEXlm License Manager", "ADSKFLEX_LICENSE_FILE", "Adding ADSKFLEX_LICENSE_FILE key"
	
	
'--- 	'Command: adsk_Registry_AddValue
		'Comment Add a vaule to the registry. TheHive is HKLM/HKCU/HKCR, RegType can be REG_SZ/REG_EXPAND_SZ/REG_DWORD/REG_BINARY
		'Syntax: adsk_Registry_AddValue TheHive, TheKey, TheValueName, TheValue, RegType, TheLogMessage	
		'Example: adsk_Registry_AddValue "HKLM", "SOFTWARE\FLEXlm License Manager", "ADSKFLEX_LICENSE_FILE", "2080@LICSERVER1;2080@LICSERVER2;2080@LICSERVER1","REG_SZ", "Adding the REGvalue"	
	
'ToDo- 	'adsk_Registry_Deletekey TheHive, TheKey, TheLogMessage
	
'--- 	'Command: adsk_Registry_DeleteValue
		'Comment: Delete a Reistry Value. TheHive is HKLM/HKCU/HKCR
		'Syntax: adsk_Registry_DeleteValue TheHive,TheKey,TheValue, TheLogMessage
		'Example: adsk_Registry_DeleteValue "HKLM","SOFTWARE\Wow6432Node\WSP","ArcGIS 10","Deleting old value from registry"
	
'---	'Command: adsk_Registry_Get_Key
		'Comment: Get a registry key from the registry. TheHive is HKLM/HKCU/HKCR
		'Syntax: adsk_Registry_Get_Key TheHive,TheKey, TheLogMessage 
		'Example: adsk_Registry_Get_Key "HKLM", "Software\WSP\Installdate", "Get the WSP Install date from registry" 
	
'ToDo- 	'adsk_Registry_Check_Key_Exist TheHive,TheKey, TheLogMessage
	
'ToDo- 	'adsk_Registry_Check_Value_Exist TheHive,TheKey,TheValue, TheLogMessage
	
'--- 	'Command: adsk_Registry_Get_Value
		'Comment: Dunction that Get a value from the registry
		'Syntax: adsk_Registry_Get_Value TheHive,TheKey,TheValue, TheLogMessage
		'Example: IF adsk_Registry_Get_Value ("HKLM", "Software\WSP", "TheDate", "Getting the WSP date from Registry") = "1958-05-19" then
					'adsk_file_copy ......
				  'END IF
	
'---	'Command: adsk_MSI_Install
		'Comment: Run an MSI install command with options
		'Syntax:  adsk_MSI_Install TheMSIfile.msi, TheParameters, TheLogMessage
		'Example: adsk_MSI_Install "Dolly.msi", "transforms=Dolly.mst /qb", "Installing program Dolly"
	
'---	'Command: adsk_MSP_Install
		'Comment: Run an MSP install command with options
		'Syntax:  adsk_MSP_Install TheMSIfile.msp, TheParameters, TheLogMessage
		'Example: adsk_MSP_Install "Dolly.msp", "transforms=Dolly.mst /qb", "Installing patch for program Dolly"	
	
'---	'Command: adsk_WMI_Uninstall
		'Comment: Uninstall a peogram with WMI
		'Syntax: adsk_WMI_Uninstall CommandString, TheLogMessage
		'Example: adsk_WMI_Uninstall "ET Geo%%", "Removing all programs starting with ET Geo"
		
'---	'Command: adsk_MSI_Uninstall
		'Comment: Uninstall a peogram with MSI
		'Syntax: adsk_MSI_Uninstall CommandString, TheLogMessage
		'Example: adsk_MSI_Uninstall "Dolly.msi", "/qb", "Removing program Dolly"		
			
'--- 	'Command: adsk_EXE_Run
		'Comment: Rune an EXE file with options
		'Syntax: adsk_EXE_Run TheExeFile, TheOptions, TheLogMessage
		'Example: adsk_EXE_Run "Dolly.exe", "/Run", "Run Dolly.exe"
	
'--- 	'Command: adsk_VBS_Run
		'Comment: Run a VB-script with parameters
		'Syntax: adsk_VBS_Run VBscriptName, Parameters, TheLogMessage
		'Example: adsk_VBS_Run "Change permissions in registry.vbs","", "Changing permissions in registry for ArcGis Administrator"
	
'--- 	'adsk_Run_Command 
		'Command: adsk_Run_Command
		'Comment: Run a Command
		'Syntax: adsk_Run_Command TheCommand, TheLogMessage
		'Example: adsk_Run_Command "netsh advfirewall firewall delete rule name=" & APO & "ArcCatalog" & chr(34), "Remove ArcCatalog exception from Firewall"
		' will run the command netsh advfirewall firewall delete rule name="ArcCatalog"
		'Example: adsk_Run_Command "ipfig /all > c:\temp\iplist.txt", "Create a text file with the IP confi"	

'---	'Command: adsk_restart_computer
		'Comment: Restart the computer, WaitTime is in seconds
		'Syntax: adsk_restart_computer WaitTime, UserMessage, TheMessage
		'Example adsk_restart_computer 120, "Computer is restarting in 2 minutes, save your work now" , "Restarting because of installation"

'--- 	'Command: adsk_CopyFile_to_all_userdirs
		'Comment: Copy Files to all userdirs
		'Syntax: adsk_CopyFile_to_all_userdirs TheRelativeFolderName, TheFileName, TheUserFilter ,TheLogMessage
		'Example: adsk_CopyFile_to_all_userdirs "Appdata\Local\ESRI", "TheFile.txt", "SE" ,"Copying a file to all user folders"
		
'---	'Command: adsk_EraseFile_from_all_userdirs
		'Comment: Erase a File from all userdirs
		'Syntax: adsk_EraseFile_from_all_userdirs TheRelativeFolderName, TheFileName, TheUserFilter ,TheLogMessage
		'Example: adsk_EraseFile_from_all_userdirs "AppData\Roaming", "untitled.upf", "SE" ,"Remove \AppData\Roaming\untitled.upf from all users"
		
'--- 	'adsk_Create_Shortcut
		'Comment: Create a shortcut to a file or application
		'Syntax: adsk_Create_Shortcut ShortcutName, Target, ShortCutDestination, Icon, StartDir, Desc, Args, TheLogMessage
		'Syntax: ShortCutDesination can be: "AllUsersDesktop"/"Desktop"/"AllUsersStartMenu"/"Start Menu"/"Programs"/"AllUsersPrograms"/"AllUsersStartup"/"Startup" or Path to a folder in the Start Menu
		'Example: adsk_Create_Shortcut "Shortcut to Dolly", "C:\Program Files (x86\Dolly\Dolly.exe", "AllUsersDesktop", "C:\Program Files (x86\Dolly\Dolly.ico", "C:\Program Files (x86\Dolly", "The Clone Program Dolly", "/startme", "Creating shortcut to Dolly"
			
		
'---	'adsk_Show_Message
		'Comment: Show a message to the user
		'Syntax: adsk_Show_Message "The Message",TheMessageType,"The Header"
		'TheMessageType = 16 (Critical)
		'TheMessageType = 32 (Warning query)
		'TheMessageType = 48 (Warning Message)
		'TheMessageType = 64 (Informationmessage)
		'Example: adsk_Show_Message "Hello World", 64, "The Message Heading
		
'-- Some help for VBS --
'You can concatenate values like this
' Value1 = "Hello"
' Value2 = " World"
' TotalValue = Value1 & Value2
' Now the variable TotalValue = "Hello World"
	
' APO = " 
' If You need  to specify  apa="Parameter"	then You can write apa = APO & "Parameter" & APO
	
	
 

 
 
 SUB Init_functions (TheMessage)
	adsk_User_Who_Started_the_Script = net.UserName
	PathToCache = left(Wscript.ScriptFullName,len(Wscript.ScriptFullName)-len(Wscript.ScriptName))	
'Get the Temp folder	
	Temp_Folder = wshShell.ExpandEnvironmentStrings( "%TEMP%" )	
' The Copy command	
	COMMAND = "c:\Windows\System32\Robocopy.exe"
'Logfiles
	LOGFILENAME =  LOGFILEPATH & "\" & LOGFILENAME & " " & strDateTime & ".log"
	LOGFILE2 = Temp_Folder & "\" & Wscript.Scriptname & "_RoboCopy.log"
	
	If fso.FileExists(LOGFILENAME) then fso.DeleteFile(LOGFILENAME)		'Delete logfile if exist
	If fso.FileExists(LOGFILE2) then fso.DeleteFile(LOGFILE2)		'Delete logfile if exist	
	
	adsk_Folder_Create_Without_Logfile LOGFILEPATH
	 
	Set Logfile= fso.CreateTextFile(LOGFILENAME, True, True)	
	
	Logfile.WriteLine("Start logging installation: ------  Date " & date & " " & Time & " --------------------")	
	StatusMessage(ScriptVersion & " Autodesk")	
	StatusMessage("Script Run Mode = " & Run_Mode)
	StatusMessage("Started by = " & adsk_User_Who_Started_the_Script)
	StatusMessage(TheMessage)		
	
 END SUB  'END of Init_functions
 
 
 SUB adsk_Show_Message (TheMessage,TheMessageType,TheHeader)
	StatusMessage("adsk_Show_Message: ><((((º>  --------------------")
	StatusMessage("adsk_Show_Message: Showing information " & TheMessage)
	StatusMessage("adsk_Show_Message: Showing Header " & TheHeader)
	Svar = msgbox (TheMessage, TheMessageType,TheHeader)
	StatusMessage("adsk_Show_Message: --------------------  ><((((º>" )	
 END SUB
 
 SUB adsk_Registry_Logentry_installed_Add (TheProgramID, TheProgramNameAndVersion, IsItMandatory, TheLogMessage)
	RC = 0	
	StatusMessage("adsk_Registry_Logentry_installed_Add: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_Logentry_installed_Add: Adding installed value to the registry")
	StatusMessage("adsk_Registry_Logentry_installed_Add: Adding " & TheProgramID & " " & TheProgramNameAndVersion)
	StatusMessage("adsk_Registry_Logentry_installed_Add: Adding " & TheLogMessage)
	StatusMessage("adsk_Registry_Logentry_installed_Add: Adding " & "When=" & Now & " - " & "Mandatory=" & IsItMandatory & " - " & TheProgramNameAndVersion)
on error resume next	
	IF Run_Mode <> "DEBUG" THEN adsk_Registry_AddKey "HKLM", "Software\Wow6432node", "WSP", TheLogMessage				
	IF Run_Mode <> "DEBUG" THEN adsk_Registry_AddKey "HKLM", "Software\Wow6432node\WSP", "SWC", TheLogMessage				
	on error goto 0		
	strTemp = "When=" & Now & " - " & "Mandatory=" & IsItMandatory & " - " & TheProgramNameAndVersion		
	IF Run_Mode <> "DEBUG" THEN adsk_Registry_AddValue "HKLM", "Software\Wow6432node\WSP\SWC", TheProgramID, strTemp, "REG_SZ", TheLogMessage				
	StatusMessage("adsk_Registry_Logentry_installed_Add: --------------------  ><((((º>" )	
 END SUB
 
 
  SUB adsk_Registry_Logentry_installed_Remove (TheProgramID, TheLogMessage)  ' Ffungerar inte
	RC = 0	
	StatusMessage("adsk_Registry_Logentry_installed_Remove: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_Logentry_installed_Remove: Removing installed value " & TheProgramID & " from the registry")
on error resume next	
	RC = WshShell.RegRead("HKLM\Software\Wow6432node\WSP\SWC\" & TheProgramID & "\")	
	RC = err.nummer
	on error goto 0	
	IF RC = 0 THEN
		StatusMessage("adsk_Registry_Logentry_installed_Remove: " & TheLogMessage)
		StatusMessage("adsk_Registry_Logentry_installed_Remove: The Value exist")
		IF RunMode <> "DEBUG" THEN WshShell.RegDelete ("HKLM\Software\Wow6432node\WSP\SWC\" & TheProgramID )		
		StatusMessage("adsk_Registry_Logentry_installed_Remove: Removed " & "HKLM\Software\Wow6432node\WSP\SWC\" & TheProgramID )
	ELSE
		StatusMessage("adsk_Registry_Logentry_installed_Remove: " & TheProgramID & " Is not found in the registry")		
	END IF
	StatusMessage("adsk_Registry_Logentry_installed_Remove: --------------------  ><((((º>" )	
 END SUB

 
 
 SUB adsk_File_Rename (TheOldFile, TheNewFile, TheMessage)
	StatusMessage("adsk_File_Rename: ><((((º>  --------------------")	
	StatusMessage("adsk_File_Rename: " & TheMessage)	
	StatusMessage("adsk_File_Rename: rename from: " & TheOldFile )	
	StatusMessage("adsk_File_Rename: rename to: " & TheNewFile )	
	IF fso.FileExists(TheNewFile) THEN
		StatusMessage("adsk_File_Rename: Attention: " & TheNewFile & " already exists. Remove it first with command adsk_File_Delete !!" )			
	END IF
	IF fso.FileExists(TheOldFile) THEN
		IF Run_Mode <> "DEBUG" THEN RC = Fso.MoveFile (TheOldFile, TheNewFile)
	ELSE		
		StatusMessage("adsk_File_Rename: " & TheOldFile & " is Missing !!" )	
	END IF
	StatusMessage("adsk_File_Rename: --------------------  ><((((º>" )	
 END SUB 'End adsk_File_Rename
 
 
 SUB adsk_WMI_Uninstall (TheProgramToUninstall, TheMessage)
	RC = 0
	If Trim(TheCommand) = "" then TheCommand = "/qb"	
	StatusMessage("adsk_WMI_Uninstall: ><((((º>  --------------------")	
	StatusMessage("adsk_WMI_Uninstall: " & TheMessage)	
	StatusMessage("adsk_WMI_Uninstall: Starting uninstall")		
	StatusMessage("adsk_WMI_Uninstall: Product to uninstall: " & TheProgramToUninstall)	
	StatusMessage("adsk_WMI_Uninstall: Command to use: CMD.EXE /C wmic product where " & APO & "name like '" & TheProgramToUninstall & "'" & APO & " call uninstall")		
	'"wmic product where " & APO & "name like '" & TheProgramToUninstall & "'" & APO & " call uninstall"
	IF Run_Mode <> "DEBUG" THEN RC = wshshell.run("CMD.EXE /C wmic product where " & APO & "name like '" & TheProgramToUninstall & "'" & APO & " call uninstall",0,True)
	StatusMessage("adsk_WMI_Uninstall: Uninstall Returncode: " & RC )	
	StatusMessage("adsk_WMI_Uninstall: Uninstall of " & TheProgramToUninstall & " ended" )	
	StatusMessage("adsk_WMI_Uninstall: --------------------  ><((((º>" )	
 END SUB  'end of adsk_WMI_Uninstall
 
 SUB adsk_Run_Command(TheCommand,TheMessage)
	RC = 0
		StatusMessage("adsk_Run_Command: ><((((º>  --------------------")			
		StatusMessage("adsk_Run_Command: " & TheMessage)
		StatusMessage("adsk_Run_Command: Running the command: " & TheCommand)
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run(TheCommand,0,True)
		StatusMessage("adsk_Run_Command: Returncode: " & RC)	
		StatusMessage("adsk_Run_Command: --------------------  ><((((º>")			
 END SUB  'End of adsk_Run_Command
 
 
  SUB adsk_MSI_Install(TheMsiFile,TheParameters,TheMessage)
	RC = 0
		MsiCommand = "C:\windows\System32\msiexec.exe /i "
		IF Instr(TheMsiFile," ") <> 0 then TheMsiFile = APO & TheMsiFile & APO
		StatusMessage("adsk_MSI_Install: ><((((º>  --------------------")			
		StatusMessage("adsk_MSI_Install: " & TheMessage)
		StatusMessage("adsk_MSI_Install: Running the command: " & MsiCommand & TheMsiFile & " " & TheParameters)
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run(MsiCommand & TheMsiFile & " " & TheParameters,0,True)
		StatusMessage("adsk_MSI_Install: Returncode: " & RC)	
		StatusMessage("adsk_MSI_Install: --------------------  ><((((º>")			
 END SUB  'End of adsk_MSI_Install

 SUB adsk_MSP_Install(TheMspFile,TheParameters,TheMessage)
	RC = 0
		MsiCommand = "C:\windows\System32\msiexec.exe /p "
		IF Instr(TheMspFile," ") <> 0 then TheMspFile = APO & TheMspFile & APO
		StatusMessage("adsk_MSP_Install: ><((((º>  --------------------")			
		StatusMessage("adsk_MSP_Install: " & TheMessage)
		StatusMessage("adsk_MSP_Install: Running the command: " & MsiCommand & TheMspFile & " " & TheParameters)
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run(MsiCommand & TheMspFile & " " & TheParameters,0,True)
		StatusMessage("adsk_MSP_Install: Returncode: " & RC)	
		StatusMessage("adsk_MSP_Install: --------------------  ><((((º>")			
 END SUB  'End of adsk_MSP_Install
 
 
 

  SUB adsk_MSI_UnInstall(TheMsiFile,TheParameters,TheMessage)
	RC = 0
		MsiCommand = "C:\windows\System32\msiexec.exe /x "
		IF Instr(TheMsiFile," ") <> 0 then TheMsiFile = APO & TheMsiFile & APO
		StatusMessage("adsk_MSI_UnInstall: ><((((º>  --------------------")			
		StatusMessage("adsk_MSI_UnInstall: " & TheMessage)
		StatusMessage("adsk_MSI_UnInstall: Running the command: " & MsiCommand & TheMsiFile & " " & TheParameters)
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run(MsiCommand & TheMsiFile & " " & TheParameters,0,True)
		StatusMessage("adsk_MSI_UnInstall: Returncode: " & RC)	
		StatusMessage("adsk_MSI_UnInstall: --------------------  ><((((º>")			
 END SUB  'End of adsk_MSI_UnInstall


 
  SUB adsk_EXE_Run(TheExeFile,TheParameters,TheMessage)
	RC = 0
		IF Instr(TheExeFile," ") <> 0 then TheExeFile = APO & TheExeFile & APO
		StatusMessage("adsk_EXE_Run: ><((((º>  --------------------")			
		StatusMessage("adsk_EXE_Run: " & TheMessage)
		StatusMessage("adsk_EXE_Run: Running the command: " & TheExeFile & " " & TheParameters)
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run(TheExeFile & " " & TheParameters,0,True)
		StatusMessage("adsk_EXE_Run: Returncode: " & RC)	
		StatusMessage("adsk_EXE_Run: --------------------  ><((((º>")			
 END SUB  'End of adsk_EXE_Run
 
 
 
 FUNCTION adsk_Is_Product_Code_Installed (TheProductCode, TheVersion, TheMessage)  
	StatusMessage("adsk_Is_ProductCode_Installed: ><((((º>  --------------------")	
	StatusMessage("adsk_Is_ProductCode_Installed: " & TheMessage)	
	StatusMessage("adsk_Is_ProductCode_Installed: Checking if the productcode " & TheProductCode & " is installed")	
	Dim msi
	Set msi = CreateObject("WindowsInstaller.Installer")
	On Error Resume Next
	Dim version
	version = msi.ProductInfo("{" & TheProductCode & "}", TheVersion)
	Dim installed
	installed = ( Err.Number = 0 )
	If Installed = 1 then 
		StatusMessage("adsk_Is_ProductCode_Installed: " & TheProductcode & " is installed")
		StatusMessage("adsk_Is_ProductCode_Installed: --------------------  ><((((º>")
		adsk_Is_Product_Code_Installed = True
		Exit Function
	else 
		StatusMessage("adsk_Is_ProductCode_Installed: " & TheProductcode & " is not installed")
		StatusMessage("adsk_Is_ProductCode_Installed: --------------------  ><((((º>")
		adsk_Is_Product_Code_Installed = False
		Exit Function
	end if
'	On Error GoTo 0 
 END FUNCTION  ' END of adsk_Is_Product_Code_Installed
 
 SUB adsk_Registry_Add_Regfile (TheRegFile, TheOption, TheMessage )
	RC = 0
		StatusMessage("adsk_Registry_Add_Regfile: ><((((º>  --------------------")
		StatusMessage("adsk_Registry_Add_Regfile: " & " = The Reg file " & TheRegFile & " is to be added with option: " & TheOption)		
	IF fso.FileExists(TheRegFile) then		
	'IF Run_Mode <> "DEBUG" THEN RC = wshshell.run("c:\Windows\regedit.exe /s " & TheRegFile & " " & TheOption,0,True)
	IF Run_Mode <> "DEBUG" THEN RC = wshshell.run("reg import " & APO & TheRegFile & APO & " " & TheOption,0,True)	
		StatusMessage("adsk_Registry_Add_Regfile: " & " = The Reg file " & TheRegFile & " Added to registry")		
		StatusMessage("adsk_Registry_Add_Regfile: RC = " & RC)		
		StatusMessage("adsk_Registry_Add_Regfile: --------------------  ><((((º>")		
	ELSE
		StatusMessage("adsk_Registry_Add_Regfile: " & " = The Reg file " & TheRegFile & " NOT found")		
		StatusMessage("adsk_Registry_Add_Regfile: RC = " & RC)		
		StatusMessage("adsk_Registry_Add_Regfile: --------------------  ><((((º>")		
	END IF
 END SUB   'END of adsk_Registry_Add_Regfile
 
 Function adsk_Registry_Get_Value (TheHive, TheKey, TheValue, TheMessage)
	RC = 0
	On error resume next
	StatusMessage("adsk_Registry_Get_Value: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_Get_Value: " & TheMessage)	
	StatusMessage("adsk_Registry_Get_Value: Getting the value " & TheHive & "\" & TheKey & "\" & TheValue)	
	RC = WshShell.RegRead (TheHive & "\" & TheKey & "\" & TheValue)		
	StatusMessage("adsk_Registry_Get_Value: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_Get_Value: --------------------  ><((((º>" )
	adsk_Registry_Get_Value = RC
	On error Goto 0
 END Function 'END of adsk_Registry_Get_Value

  Function adsk_Registry_Get_Key (TheHive, TheKey, TheMessage)
	RC = 0
	On error resume next
	StatusMessage("adsk_Registry_Get_Key: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_Get_Key: " & TheMessage)	
	StatusMessage("adsk_Registry_Get_Key: Getting the key " & TheHive & "\" & TheKey & "\")	
	RC = WshShell.RegRead (TheHive & "\" & TheKey & "\" )		
	StatusMessage("adsk_Registry_Get_Key: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_Get_Key: --------------------  ><((((º>" )
	adsk_Registry_Get_Key = RC
	On error Goto 0
 END Function 'END of adsk_Registry_Get_Key
 
 
 SUB adsk_Registry_DeleteValue (TheHive, TheKey, TheValue, TheMessage)
	RC = 0
	On error resume next
	StatusMessage("adsk_Registry_DeleteValue: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_DeleteValue: " & TheMessage)	
	StatusMessage("adsk_Registry_DeleteValue: Deleting " & TheHive & "\" & TheKey & "\" & TheValue)	
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.RegDelete (TheHive & "\" & TheKey & "\" & TheValue)	
	StatusMessage("adsk_Registry_DeleteValue: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_DeleteValue: --------------------  ><((((º>" )	
	On error Goto 0
 END SUB 'END of adsk_Registry_DeleteValue
 
 SUB adsk_Registry_Deletekey (TheHive, TheKey, TheMessage)
	RC = 0
	On error resume next
	StatusMessage("adsk_Registry_Deletekey: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_Deletekey: " & TheMessage)	
	StatusMessage("adsk_Registry_Deletekey: Deleting " & TheHive & "\" & TheKey & "\")
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.RegDelete (TheHive & "\" & TheKey & "\")		
	StatusMessage("adsk_Registry_Deletekey: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_Deletekey: --------------------  ><((((º>")	   
	On error Goto 0
 END SUB  'END of adsk_Registry_Deletekey
 
 
 SUB adsk_Registry_AddValue (TheHive, TheKey, TheValueName, TheValue, RegType, TheMessage)
	RC = 0
	If RegType = "" THEN RegType = "REG_SZ"
 	StatusMessage("adsk_Registry_AddValue: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_AddValue: " & TheMessage)	
	StatusMessage("adsk_Registry_AddValue: Adding " & TheHive & "\" & TheKey & "\" & TheValueName & "=" & TheValue & "(" & RegType & ")")	
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.RegWrite (TheHive & "\" & TheKey & "\" & TheValueName, TheValue, RegType)		
	StatusMessage("adsk_Registry_AddValue: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_AddValue: --------------------  ><((((º>")	   
 END SUB  'end of adsk_Add_Reg_Value
 

 
 SUB adsk_Registry_AddKey (TheHive, TheKey, TheValue, TheMessage)
	RC = 0
	If RegType = "" THEN RegType = "REG_SZ"
 	StatusMessage("adsk_Registry_AddKey: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_AddKey: " & TheMessage)	
	StatusMessage("adsk_Registry_AddKey: Adding " & TheHive & "\" & TheKey & "\" & TheValue)	
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.RegWrite (TheHive & "\" & TheKey & "\" , TheValue )		
	StatusMessage("adsk_Registry_AddKey: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_AddKey: --------------------  ><((((º>")	   
 END SUB  'end of adsk_Add_Reg_Value


 'ToDo- 	'adsk_Registry_Check_Key_Exist TheHive,TheKey, TheLogMessage
 SUB adsk_Registry_Check_Key_Exist (TheHive, TheKey, TheMessage)
	RC = 0
 	StatusMessage("adsk_Registry_Check_Key_Exist: ><((((º>  --------------------")	
	StatusMessage("adsk_Registry_Check_Key_Exist: " & TheMessage)	
	StatusMessage("adsk_Registry_Check_Key_Exist: Checking if " & TheHive & "\" & TheKey & " exists")	
	'IF Run_Mode <> "DEBUG" THEN RC = WshShell.RegWrite (TheHive & "\" & TheKey & "\" , TheValue )		
	StatusMessage("adsk_Registry_Check_Key_Exist: Returncode: " & RC )	   
	StatusMessage("adsk_Registry_Check_Key_Exist: --------------------  ><((((º>")	   
 END SUB  'end of adsk_Add_Reg_Value


 
 ' Subroutine for writing to logfiles
SUB Statusmessage(Statustext)
		LogFile.Writeline(Time & " ---> " & Statustext)
END SUB

PUBLIC FUNCTION adsk_Get_File_Version (TheFile2Check)
	RC = 0
	StatusMessage("adsk_Get_File_Version: " & Time & " ><((((º>  --------------------")	
	StatusMessage("adsk_Get_File_Version: Get version of " & TheFile2Check )
	IF fso.FileExists(TheFile2Check) then
		adsk_Get_File_Version = FSO.GetFileVersion(TheFile2Check)
		StatusMessage("adsk_Get_File_Version: " & TheFile2Check & ": " & adsk_Get_File_Version)
	ELSE
		adsk_Get_File_Version = "-1"
		StatusMessage("adsk_Get_File_Version: The File " & TheFile2Check & " Not found! ")			
	END IF
	StatusMessage("adsk_Get_File_Version: " & Time & " --------------------  ><((((º>")	   
END FUNCTION  'END of adsk_Get_File_Version


 'Check the file version
PUBLIC FUNCTION adsk_File_Check_Version (TheFile2Check, TheFileVersion, TheMessage, EqualOrHigher)
	RC = 0
	StatusMessage("adsk_File_Check_Version: " & Time & " ><((((º>  --------------------")	
	StatusMessage("adsk_File_Check_Version: " & TheMessage )
	IF fso.FileExists(TheFile2Check) then
		'currentVersion = FSO.GetFileVersion(TheFile2Check)
		currentVersion = adsk_Get_File_Version(TheFile2Check)
		StatusMessage("adsk_File_Check_Version: current version = " & currentVersion & "; required version = " & TheFileVersion )
		IF EqualOrHigher = 1 Then	
			IF (currentVersion >= TheFileVersion) then
				StatusMessage("adsk_File_Check_Version: Version OK ")
				adsk_File_Check_Version = True
			ELSE
				StatusMessage("adsk_File_Check_Version: Version NOK")		
				adsk_File_Check_Version = false
			END IF
		ELSE
			IF (currentVersion = TheFileVersion) then
				StatusMessage("adsk_File_Check_Version: Version OK ")
				adsk_File_Check_Version = True
			ELSE
				StatusMessage("adsk_File_Check_Version: Version NOK")		
				adsk_File_Check_Version = false
			END IF
		END IF
	ELSE
		adsk_File_Check_Version = False
		StatusMessage("adsk_File_Check_Version: The File " & TheFile2Check & " Not found! ")			
	END IF
	StatusMessage("adsk_File_Check_Version: " & Time & " --------------------  ><((((º>")	   
END FUNCTION  'END of adsk_File_Check_Version


 'compare Versions
PUBLIC Function adsk_Compare_Versions(Version1, Version2)
  Ver1 = adsk_Get_Version_StringAsArray(Version1)
  Ver2 = adsk_Get_Version_StringAsArray(Version2)
  'StatusMessage("adsk_Get_Version_StringAsArray: Version part " & 0 & " = " & Ver1(0))
  
  If Ver1(0) < Ver2(0) Then
    adsk_Compare_Versions = -1
  ElseIf Ver1(0) = Ver2(0) Then
    If Ver1(1) < Ver2(1) Then
      adsk_Compare_Versions = -1
    ElseIf Ver1(1) = Ver2(1) Then
      adsk_Compare_Versions = 0
    Else
      adsk_Compare_Versions = 1
    End If
  Else
    adsk_Compare_Versions = 1
  End If
End Function  'END of adsk_Compare_Versions


PUBLIC Function adsk_Get_Version_StringAsArray(Version)
  VersionAll = Array(0, 0, 0, 0)
  VersionParts = Split(Version, ".")
  For N = 0 To UBound(VersionParts)
    VersionAll(N) = CLng(VersionParts(N))
	'StatusMessage("adsk_Get_Version_StringAsArray: Version part " & N & " = " & VersionAll(N))
  Next
   Hi = Lsh(VersionAll(0), 16) + VersionAll(1)
   Lo = Lsh(VersionAll(2), 16) + VersionAll(3)
   
   'StatusMessage("adsk_Get_Version_StringAsArray: Hi = " & Hi)
   'StatusMessage("adsk_Get_Version_StringAsArray: Lo = " & Lo)
   adsk_Get_Version_StringAsArray = Array(Hi, Lo)
End Function


Function Lsh(N, Bits)
  ' Bitwise left shift
  Lsh = N * (2 ^ Bits)
End Function
 ' Delete file 		
SUB adsk_File_Delete (ThePath ,TheFile, TheMessage)
	RC = 0
	StatusMessage("adsk_File_Delete: ><((((º>  --------------------")	
	'Wildcards ?
	IF (Instr(TheFile,"*") <> 0) then
		StatusMessage("adsk_File_Delete: Delete with wildcards: " & ThePath & "\" & TheFile)
		WildCard = inStr(TheFile,"\")						
		StatusMessage("adsk_File_Delete: Check if folder exists  -> " & ThePath & " = " & fso.FolderExists(ThePath))		'If not exist = False	
		IF fso.FolderExists(ThePath) then
			StatusMessage("adsk_File_Delete: Deleting files " & ThePath & "\" & TheFile)
			IF Run_Mode <> "DEBUG" THEN RC = fso.DeleteFile(ThePath & "\" & TheFile, True)
			StatusMessage("adsk_File_Delete: Returncode: " & RC)
			StatusMessage("adsk_File_Delete: --------------------  ><((((º>")
		ELSE
			StatusMessage("adsk_File_Delete: ----- Folder not found--------------")		
			StatusMessage("adsk_File_Delete: ----- Can not delete files ---------")		
			StatusMessage("adsk_File_Delete: --------------------  ><((((º>")
		END IF
	ELSE	
		StatusMessage("adsk_File_Delete: Check if file exists  -> " & ThePath & " = " & fso.FileExists(ThePath & "\" & TheFile))	'If not exist = False	
		IF (fso.FileExists(ThePath & "\" & TheFile))  THEN
			IF fso.FileExists(ThePath & "\" & TheFile) then			
				StatusMessage("adsk_File_Delete: Delete the file => " & ThePath & "\" & TheFile)
				'StatusMessage("adsk_File_Delete: -------------- Next File --------------")			
				StatusMessage("adsk_File_Delete: --------------------  ><((((º>")
				IF Run_Mode <> "DEBUG" THEN RC = fso.DeleteFile( ThePath & "\" & TheFile, True) 	
				StatusMessage("adsk_File_Delete: Returncode: " & RC)
			ELSE
				StatusMessage("adsk_File_Delete: -------------- File not found--------------")		
				StatusMessage("adsk_File_Delete: " & ThePath & "\" & TheFile)		
			END IF
		ELSE
			'StatusMessage("adsk_File_Delete: ----- Next File ------------------")		
			StatusMessage("adsk_File_Delete: --------------------  ><((((º>")
		END IF
	END IF
END SUB  'END of adsk_File_Delete


 ' Delete folder and files within
SUB adsk_Folder_Delete (TheFolder2Delete, TheMessage)
	RC = 0
	StatusMessage("adsk_Folder_Delete: ><((((º>  --------------------")	
	StatusMessage("adsk_Folder_Delete: Check if folder exists  -> " & TheFolder2Delete & " = " & fso.FolderExists(TheFolder2Delete))		'If not exist = False	
	IF (fso.FolderExists(TheFolder2Delete))  THEN
		IF fso.FolderExists(TheFolder2Delete) then			
			StatusMessage("adsk_Folder_Delete: Delete the folder => " & TheFolder2Delete)
			StatusMessage("adsk_Folder_Delete:  -----------------------------  ><((((º>  --")
			IF Run_Mode <> "DEBUG" THEN RC = fso.DeleteFolder (TheFolder2Delete, True )
			StatusMessage("adsk_Folder_Delete:  Returncode: " & RC )
		ELSE
			StatusMessage("adsk_Folder_Delete: -------------- Folder not found--------------")	
		END IF
	ELSE
		'StatusMessage("adsk_Folder_Delete: -------------- Next Folder ----------------")				
		StatusMessage("adsk_Folder_Delete: --------------------  ><((((º>")
	END IF

END SUB 'END of adsk_Folder_Delete


SUB adsk_Folder_Create (TheFolder, TheMessage)
	RC = 0
	StatusMessage("adsk_Folder_Create: ><((((º>  --------------------")	
	StatusMessage("adsk_Folder_Create: " & TheMessage )	
	Dim arrDirs, i, idxFirst, strDir, strDirBuild
    ' Convert relative to absolute path
    strDir = FSO.GetAbsolutePathName( TheFolder )

    ' Split a multi level path in its "components"
    arrDirs = Split( strDir, "\" )

    ' Check if the absolute path is UNC or not
    If Left( strDir, 2 ) = "\\" Then
        strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
        idxFirst    = 4
    Else
        strDirBuild = arrDirs(0) & "\"
        idxFirst    = 1
    End If

    ' Check each (sub)folder and create it if it doesn't exist
    For i = idxFirst to Ubound( arrDirs )
        strDirBuild = FSO.BuildPath( strDirBuild, arrDirs(i) )
		StatusMessage("adsk_Folder_Create: Check for existence of: " & strDirBuild & ". Exist = " &  fso.FolderExists(strDirBuild))		'If Exist = True
        If Not FSO.FolderExists( strDirBuild ) Then 
			StatusMessage("adsk_Folder_Create: Creating Folder")
			IF Run_Mode <> "DEBUG" THEN RC = FSO.CreateFolder (strDirBuild)
			StatusMessage("adsk_Folder_Create: Returncode: " & RC )
		else
			StatusMessage("adsk_Folder_Create: Folder already exist")		
        End if
    Next    
		StatusMessage("adsk_Folder_Create: --------------------  ><((((º>")	
END SUB  'end of adsk_Folder_Create


SUB adsk_Folder_Create_Without_Logfile (TheFolder)
	RC = 0
	Dim arrDirs, i, idxFirst, strDir, strDirBuild
    ' Convert relative to absolute path
    strDir = FSO.GetAbsolutePathName( TheFolder )
    ' Split a multi level path in its "components"
    arrDirs = Split( strDir, "\" )
    ' Check if the absolute path is UNC or not
    If Left( strDir, 2 ) = "\\" Then
        strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
        idxFirst    = 4
    Else
        strDirBuild = arrDirs(0) & "\"
        idxFirst    = 1
    End If
    ' Check each (sub)folder and create it if it doesn't exist
    For i = idxFirst to Ubound( arrDirs )
        strDirBuild = FSO.BuildPath( strDirBuild, arrDirs(i) )
	'If Exist = True
        If Not FSO.FolderExists( strDirBuild ) Then 
			RC = FSO.CreateFolder (strDirBuild)
        End if
    Next    
END SUB  'end of adsk_Folder_Create



 ' Copy folders
SUB adsk_Folder_Copy(Fromfolder, Tofolder, Copy_Action, LogText)   'Copy the folders
	RC = 0
	Fromfolder = PathToCache & Fromfolder
	StatusMessage("adsk_Folder_Copy: ><((((º>  --------------------")	
	StatusMessage("adsk_Folder_Copy: " & LogText)	
	StatusMessage("adsk_Folder_Copy: Check Fromfolder " & Fromfolder & " = " & fso.FolderExists(Fromfolder))		'If Exist = True
	StatusMessage("adsk_Folder_Copy: Check Tofolder   " & Tofolder & " = " & fso.FolderExists(Tofolder))		'If not exist = False
	IF NOT (fso.FolderExists(ToFolder)) Then adsk_Folder_Create Tofolder,"Creating missing folder: " & Tofolder
	IF (fso.FolderExists(Fromfolder) AND (fso.FolderExists(ToFolder))) Then  	'Both folders exists
		'IF Updated_Files(Fromfolder, Tofolder, "adsk_Folder_Copy") Then				'Is UPDATE.TXT updated?
		IF Ucase(Copy_Action) = "UPDATE" Then				'Is UPDATE.TXT updated?
			StatusMessage("adsk_Folder_Copy: Action = " & Copy_Action)	
			StatusMessage("adsk_Folder_Copy: UPDATE " & ToFolder & " from " & Fromfolder)	
			IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run (COMMAND & " " & APO & FromFolder & APO & " " & APO & ToFolder & APO & " /TEE /E /NP /r:0 /w:0 /FFT /COPY:DT /XD RECYCLER /log+:" & LOGFILE2,1,True)
			StatusMessage("adsk_Folder_Copy: Returncode: " & RC)
			StatusMessage("adsk_Folder_Copy: -------------- Next Folder --------------")			
			StatusMessage("adsk_Folder_Copy: --------------------  ><((((º>")
			StatusMessage(" ")
		ELSE
			StatusMessage("-------------- No Update ----------------")			'Yes = No files to update
			StatusMessage("adsk_Folder_Copy: UPDATE " & ToFolder & " from " & Fromfolder)	
		END IF	
	
		IF Ucase(Copy_Action) = "COPY" then 'Force the copy without ToFolder.			
			StatusMessage("adsk_Folder_Copy: Action = " & Copy_Action)	
			StatusMessage("adsk_Folder_Copy: COPY " & ToFolder & " from " & Fromfolder)		
			StatusMessage("adsk_Folder_Copy: UPDATE " & ToFolder & " from " & Fromfolder)	
			'RC = fso.GetFolder((Fromfolder).Copy, Tofolder)
			IF Run_Mode <> "DEBUG" THEN RC = fso.CopyFolder(Fromfolder, Tofolder)			
			StatusMessage("adsk_Folder_Copy: Returncode: " & RC)
		END IF
	
		IF Ucase(Copy_Action) = "FORCE" then 'Force the copy without ToFolder.
			StatusMessage("adsk_Folder_Copy: Action = " & Copy_Action)				
			StatusMessage("adsk_Folder_Copy: Force copy from " & Fromfolder & "to " & ToFolder)	
			StatusMessage("adsk_Folder_Copy: UPDATE " & ToFolder & " from " & Fromfolder)	
			IF Run_Mode <> "DEBUG" THEN RC = WshShell.Run (COMMAND & " " & APO & FromFolder & APO & " " & APO & ToFolder & APO & " /TEE /E /NP /r:0 /w:0 /FFT /COPY:DT /XD RECYCLER /log+:" & LOGFILE2,1,True)		
			StatusMessage("adsk_Folder_Copy: Returncode: " & RC)
		END IF		
		'StatusMessage("adsk_Folder_Copy: -------------- Next Folder --------------")		
		StatusMessage("adsk_Folder_Copy: --------------------  ><((((º>")
	END IF		
END SUB  'END of adsk_Folder_Copy

 ' Copy files
SUB adsk_File_Copy(FromFolder, TheFile, ToFolder, Copy_Action, Logtext)								'Copy some files
	RC = 0
	Fromfolder = PathToCache & Fromfolder
	If right(ToFolder,1) <> "\" then ToFolder = ToFolder & "\"
	StatusMessage("adsk_File_Copy: ><((((º>  --------------------")	
	StatusMessage("adsk_File_Copy " &LogText)
	'IF (fso.FileExists(Fromfolder & "\" & TheFile) AND (fso.FolderExists(ToFolder))) Then 
	
	  If Ucase(Copy_Action) = "UPDATE" then
		IF Updated_Files(Fromfolder, Tofolder, "adsk_File_Copy") Then   'If the file UPDATE.TXT exists then there are updated files	
			StatusMessage("adsk_File_Copy: Copy file " & Fromfolder & "\" & TheFile & " to folder " & ToFolder)
			StatusMessage("adsk_Folder_Copy: UPDATE " & ToFolder & " from " & Fromfolder)	
			IF Run_Mode <> "DEBUG" THEN RC = fso.CopyFile (FromFolder & "\" & TheFile, ToFolder)
			StatusMessage("adsk_File_Copy: Returncode " & RC)
			'RC = WshShell.Run (COMMAND & " " & APO & Fromfolder & "\" & TheFile & APO & " " & APO & ToFolder & APO & " /TEE /E /NP /r:0 /w:0 /FFT /COPY:DT /XD RECYCLER /log+:" & LOGFILE2,1,True)
			StatusMessage("adsk_File_Copy: --------------------  ><((((º>")
		ELSE
			StatusMessage("-------------- No Update ----------------")		'No files to update
		END IF
	  END IF
	  
	  If Ucase(Copy_Action) = "COPY" then		
			StatusMessage("adsk_File_Copy: " &LogText & " Start the copy")
			StatusMessage("adsk_File_Copy: Copy file " & Fromfolder & "\" & TheFile & " to folder " & ToFolder)
			IF Run_Mode <> "DEBUG" THEN RC = fso.CopyFile (FromFolder & "\" & TheFile, ToFolder)
			StatusMessage("adsk_File_Copy: Returncode " & RC)
			StatusMessage("adsk_File_Copy: --------------------  ><((((º>")
	  END IF
	  	  
	'ELSE
	'	StatusMessage("adsk_File_Copy: Fromfile or ToFolder is missing!")		
	'	StatusMessage("adsk_File_Copy: --------------------  ><((((º>")
	'END IF	

END SUB   'END of adsk_File_Copy

 ' Check i there is updated files
FUNCTION Updated_Files(FromFolder, ToFolder, Logtext)
	RC = 0
	Fromfolder = PathToCache & Fromfolder
	StatusMessage(" -- Call coming from:" & LogText)
	StatusMessage(" -- Checking UPDATE.TXT --")
	IF fso.FileExists(FromFolder & "\UPDATE.TXT") Then 			'Remote files maybe updated?
		Set FromFolderFile = fso.GetFile(FromFolder & "\UPDATE.TXT")		
		IF fso.FileExists(ToFolder & "\UPDATE.TXT") Then		'Local files has run the update before
			Set ToFolderFile = fso.GetFile(Tofolder & "\UPDATE.TXT")
			IF FromFolderFile.DateLastModified = ToFolderFile.DateLastModified then			'Remote and Local files are the same = No update
				StatusMessage("Remote UPDATE.TXT last modified=" & FromFolderFile.DateLastModified & " - Local UPDATE.TXT last modified=" & ToFolderFile.DateLastModified)
				StatusMessage(" -- No Update Needed UPDATE.TXT is up to date--")
				StatusMessage("--------------  ><((((º>   --------------")		'No files to update	
				StatusMessage(" ")				
				Updated_Files = False
				Exit Function
			ELSE
				StatusMessage(" -- Remote and local files are different = Update needed --")			'Remote and local files are different = run update
				StatusMessage("Remote UPDATE.TXT last modified=" & FromFolderFile.DateLastModified & " - Local UPDATE.TXT last modified=" & ToFolderFile.DateLastModified)
				Updated_Files = True			
				Exit Function			
			END IF
		ELSE						
			StatusMessage(" -- Remote UPDATE.TXT found, Local files never updated = Update needed --")			'Remote files update and local files never updated = Run update
			StatusMessage("Remote UPDATE.TXT last modified=" & FromFolderFile.DateLastModified)
			Updated_Files = True			
			Exit Function
		END IF		
	ELSE
			StatusMessage("Remote UPDATE.TXT missing = No Update")
			StatusMessage("--------------  ><((((º>   --------------")		'No files to update			
			StatusMessage(" ")
	END IF
END FUNCTION

Function adsk_File_Check_Exist(TheFileName,TheMessage)
	RC = 0
	StatusMessage("adsk_File_Check_Exist: ><((((º>  --------------------")
	StatusMessage("adsk_File_Check_Exist: " & TheMessage)	
	IF fso.FileExists(TheFileName) Then
		StatusMessage("adsk_File_Check_Exist: Yes the file " & TheFileName & " Exists" )
		StatusMessage("adsk_File_Check_Exist: --------------------  ><((((º>")
		adsk_File_Check_Exist = True
	ELSE
		StatusMessage("adsk_File_Check_Exist: No the file " & TheFileName & " is missing" )
		StatusMessage("adsk_File_Check_Exist: --------------------  ><((((º>")
		adsk_File_Check_Exist = False	
	END IF
END FUNCTION

Function adsk_Folder_Check_Exist(TheFolderName,TheMessage)
	RC = 0
	StatusMessage("adsk_Folder_Check_Exist: ><((((º>  --------------------")
	StatusMessage("adsk_Folder_Check_Exist: " & TheMessage)	
	IF fso.FolderExists(TheFolderName) Then
		StatusMessage("adsk_Folder_Check_Exist: Yes the folder " & TheFolderName & " Exists" )
		StatusMessage("adsk_Folder_Check_Exist: --------------------  ><((((º>")
		adsk_Folder_Check_Exist = True
	ELSE
		StatusMessage("adsk_Folder_Check_Exist: No the folder " & TheFolderName & " is missing" )
		StatusMessage("adsk_Folder_Check_Exist: --------------------  ><((((º>")
		adsk_Folder_Check_Exist = False	
	END IF
END FUNCTION


Function adsk_File_GetInfo (Thefilename, TheWantedFileInfo, TheMessage)
	RC = 0
	StatusMessage("adsk_File_GetInfo: ><((((º>  --------------------")
	StatusMessage("adsk_File_GetInfo: " & TheMessage)
	StatusMessage("adsk_File_GetInfo: Looking for: " & TheWantedFileInfo)
	Set f = fso.GetFile(Thefilename)
	TheWantedFileInfo = Ucase(TheWantedFileInfo)
	Select Case TheWantedFileInfo
		Case "CREATIONDATE" adsk_File_GetInfo = f.DateCreated
		Case "DATELASTACCESSED" adsk_File_GetInfo = f.DateLastAccessed 
		Case "DATELASTMODIFIED" adsk_File_GetInfo = f.DateLastModified 	
		Case "READONLY"  If f.attributes and 1 Then adsk_File_GetInfo = True 	
		Case "HIDDEN"  If f.attributes and 2 Then adsk_File_GetInfo = True 	
		Case "SYSTEM"  If f.attributes and 4 Then adsk_File_GetInfo = True 	
	END Select
	StatusMessage("adsk_File_GetInfo: --------------------  ><((((º>")
END FUNCTION


SUB adsk_restart_computer (WaitTime, UserMessage, TheMessage)
	RC = 0
	StatusMessage("adsk_restart_computer: ><((((º>  --------------------")	
	StatusMessage("adsk_restart_computer: "& TheMessage )	
	StatusMessage("adsk_restart_computer: Restarting this computer in " & WaitTime & "seconds")		
	StatusMessage("c:\Windows\System32\shutdown.exe -t " & WaitTime & " -r " & UserMessage)
	IF Run_Mode <> "DEBUG" THEN WshShell.run("c:\Windows\System32\shutdown.exe -t " & WaitTime & " -r ")	
	StatusMessage("adsk_restart_computer: --------------------  ><((((º>")
END SUB

SUB adsk_VBS_Run (VBscriptName, Parameter, TheMessage)
	RC = 0
	StatusMessage("adsk_VBS_Run: ><((((º>  --------------------")
	StatusMessage("adsk_VBS_Run: " & TheMessage)
	IF instr(Parameter," ") <> 0 THEN Parameter = APO & Parameter & APO
	IF instr(VBscriptName," ") <> 0 THEN VBscriptName = APO & VBscriptName & APO
	IF Run_Mode <> "DEBUG" THEN RC = WshShell.run(VBscriptName & " " & Parameter)	
	StatusMessage("adsk_VBS_Run: Returncode " & RC)
	StatusMessage("adsk_VBS_Run: --------------------  ><((((º>")
END SUB

SUB adsk_CopyFile_to_all_userdirs(TheRelativeFolderName, TheFileName, TheUserFilter ,TheMessage)
'	TheRelativeFolderName = the foldername relative to C:\Users for example: Appdata\Local\ESRI
'	TheFileName = The name of the file to copy	
' 	TheUserFilter = filter for the user folder to serach in for example: SE
	RC = 0 
	StatusMessage("adsk_CopyFile_to_all_userdirs: ><((((º>  --------------------")
	StatusMessage("adsk_CopyFile_to_all_userdirs: " & TheMessage)
	StatusMessage("adsk_CopyFile_to_all_userdirs: TheRelativeFolderName = " & TheRelativeFolderName)
	StatusMessage("adsk_CopyFile_to_all_userdirs: TheFileName = " & TheFileName)
	StatusMessage("adsk_CopyFile_to_all_userdirs: TheUserFilter = " & TheUserFilter)
	dir = "C:\Users"
	dim oFolder,oFolders,oFiles,item
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	set oFolder=oFileSys.GetFolder(dir)
	set oFolders=oFolder.SubFolders
	set oFiles=oFolder.Files

	' get all sub-folders in this folder
	For each item in oFolders
		IF Ucase(Left(item.name,2)) = Ucase(TheUserFilter) Then		' Use only if You want to limit the search
			If oFileSys.FolderExists(Dir & "\" & item.name & "\" & TheRelativeFolderName) THEN
				StatusMessage("adsk_CopyFile_to_all_userdirs: " & Dir & "\" & item.name & "\" & TheRelativeFolderName & " Exists")
				IF Run_Mode <> "DEBUG" THEN RC = oFilesys.CopyFile(TheFileName, Dir & "\" & item.name & "\" & TheRelativeFolderName & "\" , true)											
				StatusMessage("adsk_CopyFile_to_all_userdirs: Returncode " & RC)
			ELSE			
				StatusMessage("adsk_CopyFile_to_all_userdirs: " & Dir & "\" & item.name & "\" & TheRelativeFolderName & " Does not Exist")
				adsk_Folder_Create  Dir & "\" & item.name & "\" & TheRelativeFolderName ,"Creating the missing folder"
				StatusMessage("adsk_CopyFile_to_all_userdirs: " & Dir & item.name & "\" & TheRelativeFolderName & " Does now Exists")
				IF Run_Mode <> "DEBUG" THEN RC = oFilesys.CopyFile(TheFileName, Dir & "\" & item.name & "\" & TheRelativeFolderName & "\", true)
				StatusMessage("adsk_CopyFile_to_all_userdirs: Returncode " & RC)
			END IF
			' Use only if You want to limit the search
		END IF
	Next
	StatusMessage("adsk_CopyFile_to_all_userdirs: --------------------  ><((((º>")
END SUB


'adsk_EraseFile_from_all_userdirs
SUB adsk_EraseFile_from_all_userdirs(TheRelativeFolderName, TheFileName, TheUserFilter ,TheMessage)
'	TheRelativeFolderName = the foldername relative to C:\Users for example: Appdata\Local\ESRI
'	TheFileName = The name of the file to copy
	TheFilename = TheFilename
' 	TheUserFilter = filter for the user folder to serach in for example: SE
	RC = 0 
	StatusMessage("adsk_EraseFile_from_all_userdirs: ><((((º>  --------------------")
	StatusMessage("adsk_EraseFile_from_all_userdirs: " & TheMessage)
	StatusMessage("adsk_EraseFile_from_all_userdirs: TheRelativeFolderName = " & TheRelativeFolderName)
	StatusMessage("adsk_EraseFile_from_all_userdirs: TheFileName = " & TheFileName)
	StatusMessage("adsk_EraseFile_from_all_userdirs: TheUserFilter = " & TheUserFilter)
	dir = "C:\Users"
	dim oFolder,oFolders,oFiles,item
	Set oFileSys = CreateObject("Scripting.FileSystemObject")
	set oFolder=oFileSys.GetFolder(dir)
	set oFolders=oFolder.SubFolders
	set oFiles=oFolder.Files

	' get all sub-folders in this folder
	For each item in oFolders
		IF Ucase(Left(item.name,2)) = Ucase(TheUserFilter) Then		' Use only if You want to limit the search
			If oFileSys.FolderExists(Dir & "\" & item.name & "\" & TheRelativeFolderName) THEN				
				StatusMessage("adsk_EraseFile_from_all_userdirs: " & Dir & "\" & item.name & "\" & TheRelativeFolderName & " Folder Exists")				
				IF oFileSys.FileExists(Dir & "\" & item.name & "\" & TheRelativeFolderName & "\" & TheFileName) Then
					StatusMessage("adsk_EraseFile_from_all_userdirs: " & Dir & "\" & item.name & "\" & TheRelativeFolderName & "\" & TheFileName & " File Exists")
					TheFile = Dir & "\" & item.name & "\" & TheRelativeFolderName & "\" & TheFileName
					'IF Run_Mode <> "DEBUG" THEN RC = adsk_File_Delete(Dir & "\" & item.name & "\" & TheRelativeFolderName, TheFileName, "Deleting File")
					
					IF Run_Mode <> "DEBUG" THEN RC = ofilesys.DeleteFile(TheFile,true)																							
					StatusMessage("adsk_EraseFile_from_all_userdirs: Returncode " & RC)
				END IF
				
			END IF
		END IF
	Next
	StatusMessage("adsk_EraseFile_from_all_userdirs: --------------------  ><((((º>")
END SUB






SUB adsk_Create_Shortcut (SCName, Target, ScDest, Icon, StartDir, Desc, Args, TheLogMessage )
On error resume next
	StatusMessage("adsk_Create_Shortcut: ><((((º>  --------------------")
	StatusMessage("adsk_Create_Shortcut: " & TheLogMessage)
	StatusMessage("adsk_Create_Shortcut: " & SCName & "," & Target & "," & ScDest & "," & Icon & "," & StartDir & "," & Desc & "," & Args)

	IF Icon 	=	"" then Icon = Target
	IF StartDir 	= 	"" then StartDir = ".\"
	IF Desc		=	"" then Desc = "A shortcut"
	IF Args		=	"" then Args = ""

	Set WshSysEnv = WshShell.Environment("PROCESS")

	Windir=WshSysEnv("WINDIR")
	Progdir=WshSysEnv("PROGRAMFILES")

	IF (instr(ScDest,"\") = 0) then
		strDestination = WshShell.SpecialFolders(ScDest)
		set oShellLink = WshShell.CreateShortcut(strDestination & "\" & SCName & ".lnk")
	ELSE
		set oShellLink = WshShell.CreateShortcut(scDest & "\" & SCName & ".lnk")
	END IF
    oShellLink.TargetPath = Target
	'oShellLink.WindowStyle = 1 '1,3,7
    'oShellLink.Hotkey = "CTRL+SHIFT+F" 'ALT+, CTRL+, SHIFT+, EXT+. 
    oShellLink.IconLocation =  Icon
    oShellLink.Description = Desc
    oShellLink.WorkingDirectory = StartDir
	oShellLink.Arguments = Args
    oShellLink.Save
         
	StatusMessage("adsk_Create_Shortcut: --------------------  ><((((º>")
On error goto 0
END SUB 'End adsk_Create_Shortcut
SUB adsk_Remove_Shortcut (File1, TheLogMessage )

    RC = 0 
	StatusMessage("adsk_Remove_Shortcut: ><((((º>  --------------------")
	StatusMessage("adsk_Remove_Shortcut: " & TheMessage)
	StatusMessage("adsk_Remove_Shortcut: File = " & File1)

	Dim objShell:Set objShell=CreateObject("Wscript.Shell") 

' For the Removal Process, remove shortcuts/icons from the "Corporate Shortcuts" desktop folder (all users) on the local machine(s):: 
' All users
' Programs 		%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs
' Start Menu 	%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu
' Current user
' Desktop 		%USERPROFILE%\Desktop
' Favorites		%USERPROFILE%\Favorites
' Start Menu	%APPDATA%\Microsoft\Windows\Start Menu
' Startup		%APPDATA%\Microsoft\Windows\Start Menu\Programs\StartUp
' Public
' Public Desktop %PUBLIC%\Desktop
' C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Autodesk

    IF Run_Mode <> "DEBUG" THEN
		'arrFiles = Array(File1, File2, File3) 
		arrLocations = Array("AllUsersDesktop" , "AllUsersStartMenu" , "AllUsersPrograms" , "AllUsersStartup" , "Desktop" , "Favorites" , "StartMenu" , "Startup")

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objShell = CreateObject("WScript.Shell")


		For Each strLocation in arrLocations
			strLoc = objShell.SpecialFolders(strLocation)
			StatusMessage("adsk_Remove_Shortcut: Location = " & strLoc)
			If Right(strLoc, 1) <> "\" Then strLoc = strLoc & "\"
			'For Each strFile In arrFiles
				If objFSO.FileExists(strLoc & File1) = True Then 
					objFSO.DeleteFile strLoc & File1, True
					StatusMessage("adsk_Remove_Shortcut: Removed " & File1 & " from " & strLoc)
				end If
				
				If strLocation = "AllUsersPrograms" Then
				  strLoc = strLoc & "Autodesk"
				  StatusMessage("adsk_Remove_Shortcut: Location = " & strLoc)
					If Right(strLoc, 1) <> "\" Then strLoc = strLoc & "\"
					'For Each strFile In arrFiles
						If objFSO.FileExists(strLoc & File1) = True Then 
							objFSO.DeleteFile strLoc & File1, True
							StatusMessage("adsk_Remove_Shortcut: Removed " & File1 & " from " & strLoc)
						end If
				end If
				
			'Next
		Next
	END IF
	StatusMessage("adsk_Remove_Shortcut: --------------------  ><((((º>")
END SUB 'End adsk_Remove_Shortcut

'#==============================================================================
'#==============================================================================
'#  SCRIPT.........:	CleanUpStartMenuItems.vbs
'#  AUTHOR.........:	Stuart Barrett
'#  VERSION........:	1.0
'#  CREATED........:	11/11/11
'#  LICENSE........:	Freeware
'#  REQUIREMENTS...:	
'#
'#  DESCRIPTION....:	Cleans up any Start Menu folders containing broken 
'#						shortcuts
'#
'#  NOTES..........:	Untested on Windows 7!!
'#						
'#						Will ask before deleting any folders (can be amended
'#						as per noted in script)
'#						
'#						You can add as many excluded folders as required.
'# 
'#  CUSTOMIZE......:  
'#==============================================================================
'#  REVISED BY.....:  
'#  EMAIL..........:  
'#  REVISION DATE..:  
'#  REVISION NOTES.:
'#
'#==============================================================================
'#==============================================================================

Sub adsk_CleanUpStartMenuItems(intInclude, intExclude, intDelFolder, TheMessage)
    StatusMessage("adsk_CleanUpStartMenuItems: ><((((º>  --------------------")
	StatusMessage("adsk_CleanUpStartMenuItems: " & TheMessage)
	StatusMessage("adsk_CleanUpStartMenuItems: Include = " & intInclude)
	StatusMessage("adsk_CleanUpStartMenuItems: Exclude = " & intExclude)
	StatusMessage("adsk_CleanUpStartMenuItems: Delete folder = " & intDelFolder)
	
	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	strStartMenu = objShell.SpecialFolders("StartMenu")
	strAllUsersStartMenu = objShell.SpecialFolders("AllUsersStartMenu")

	adsk_DeleteStartMenuItems strStartMenu, 0, intInclude, intExclude, intDelFolder, TheMessage
	adsk_DeleteStartMenuItems strAllUsersStartMenu, 1, intInclude, intExclude, intDelFolder, TheMessage

end SUB 'End adsk_CleanUpStartMenuItems
'#--------------------------------------------------------------------------
'#	SUBROUTINE.....:	DeleteStartMenuItems(strFolder, intStartMenu)
'#	PURPOSE........:	Deletes all the folders with broken shortcuts in
'#						the specified folder
'#	ARGUMENTS......:	strFolder = full path to the folder
'#						intStartMenu = index of Start Menu type
'#	EXAMPLE........:	DeleteStartMenuItems("c:\documents and settings\all users\start menu", 0)
'#	NOTES..........:	intStartMenu Values:
'#						0 = Start Menu
'#						1 = All Users Start Menu
'#--------------------------------------------------------------------------
Sub adsk_DeleteStartMenuItems(strFolder, intStartMenu, intInclude, intExclude, intDelFolder, TheMessage)
 On Error Resume Next
    	
	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
    
	If intExclude = 1 Then
	'#--------------------------------------------------------------------------
	'#	Add Any Excluded Folders in here. You will also need to increase the
	'#	arrExcludedFolders(x) array variable by 1.
	'#--------------------------------------------------------------------------
		Dim arrExcludedFolders(3)
		
		arrExcludedFolders(0) = "\programs"
		arrExcludedFolders(1) = "\programs\accessories\*"
		arrExcludedFolders(2) = "\programs\administrative tools"
	'#--------------------------------------------------------------------------
	End If
	
	If intInclude = 1 Then
	'#--------------------------------------------------------------------------
	'#	Add Any Included Folders in here. You will also need to increase the
	'#	arrIncludedFolders(x) array variable by 1.
	'#--------------------------------------------------------------------------
		Dim arrIncludedFolders(1)
		
		arrIncludedFolders(0) = "\programs\Autodesk\*"
	'#--------------------------------------------------------------------------
	End If
	
	booDelete = 0
	booCheck = 1
	Set objFolder = objFSO.GetFolder(strFolder)
			
	If intStartMenu = 0 Then
		strSM = objShell.SpecialFolders("StartMenu")
	Else 
		strSM = objShell.SpecialFolders("AllUsersStartMenu")
	End If
	
	If intExclude = 1 Then
		booCheck = 1
		For i = 0 To UBound(arrExcludedFolders) - 1
			strExcludedFolder = strSM & arrExcludedFolders(i)

			If LCase(strFolder) = LCase(strSM) Then
				booCheck = 0
				Exit For
			End If
						
					
			If Right(strExcludedFolder, 2) = "\*" Then
				strExcludedFolder = Left(strExcludedFolder, Len(strExcludedFolder) - 2)
						
				If InStr(LCase(strFolder), LCase(strExcludedFolder)) > 0 Then 
					booCheck = 0
					Exit For
				End If
				Else
					If InStr(LCase(strExcludedFolder) & "||", LCase(strFolder) & "||") > 0 Then
						booCheck = 0
						Exit For
					End If
			End If
		Next
	End If		
	
	If intInclude = 1 Then
		booCheck = 0
		For i = 0 To UBound(arrIncludedFolders) - 1
			strIncludedFolder = strSM & arrIncludedFolders(i)

			If LCase(strFolder) = LCase(strSM) Then
				booCheck = 0
				Exit For
			End If
						
					
			If Right(strIncludedFolder, 2) = "\*" Then
				strIncludedFolder = Left(strIncludedFolder, Len(strIncludedFolder) - 2)
						
				If InStr(LCase(strFolder), LCase(strIncludedFolder)) > 0 Then 
					booCheck = 1
					Exit For
				End If
				Else
					If InStr(LCase(strIncludedFolder) & "||", LCase(strFolder) & "||") > 0 Then
						booCheck = 1
						Exit For
					End If
			End If
		Next
	End If		
	
	If booCheck = 1 Then
		For Each objItem In objFolder.Files
			strFullName = objFSO.GetAbsolutePathName(objItem)

			If Right(LCase(objItem), 3) = "lnk" Then
				Set objShortcut = objShell.CreateShortcut(strFullName)
				strTarget = LCase(objShortcut.TargetPath)

				If NOT objFSO.FileExists(strTarget) Then
					IF intDelFolder = 1 Then
						booDelete = 1
						Exit For
					ELSE
					    booDelete = 0
						objFSO.DeleteFile strFullName, True
						StatusMessage("adsk_DeleteStartMenuItems: Deleted file " & strFullName)
					End If
					'#--------------------------------------------------------------------------
				End If
			End If
		Next
				
		If booDelete = 1 Then 
		    StatusMessage("adsk_DeleteStartMenuItems: Deleted folder: " & objFolder)
			objFolder.Delete True
		end If
	End If

	For Each objItem In objFolder.SubFolders
		adsk_DeleteStartMenuItems objItem.Path, intStartMenu, intInclude, intExclude, intDelFolder, TheMessage
	Next
End Sub 'end adsk_DeleteStartMenuItems

Sub adsk_CleanUpLicenseRegKeys(licServer1, licServer2, TheMessage)
    StatusMessage("adsk_CleanUpLicenseRegKeys: ><((((º>  --------------------")
	StatusMessage("adsk_CleanUpLicenseRegKeys: License server1: " & licServer1)
	StatusMessage("adsk_CleanUpLicenseRegKeys: License server2: " & licServer2)
	StatusMessage("adsk_CleanUpLicenseRegKeys: " & TheMessage)
	

	
'Set objShell = CreateObject("WScript.Shell")

' These keys will be cleaned
'----------
strKeyClean1      = "HKEY_LOCAL_MACHINE\SOFTWARE\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE"
strKeyClean2      = "HKEY_CURRENT_USER\Software\Flexlm License Manager\ADSKFLEX_LICENSE_FILE"
strKeyClean3      = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Session Manager\Environment\ADSKFLEX_LICENSE_FILE"
strKeyClean4      = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet002\Control\Session Manager\Environment\ADSKFLEX_LICENSE_FILE"

strKeyName = "ADSKFLEX_LICENSE_FILE"
if not licServer1 = "" then 
  strKeyValue = "2080@" & licServer1
end if

if not licServer2 = "" then 
  strKeyValue = strKeyValue & ";2080@" & licServer2
end if

StatusMessage("adsk_CleanUpLicenseRegKeys: Registry Key value: " & strKeyValue)

Const HKEY_USERS = &H80000003
strComputer = "." 
ContractExpress = "\Software\FLEXlm License Manager"
'----------

'This is the new Server variable to set
'strKeyValue1 = "2080@BRBSAPP24095"

'dim checkInfoBorrow
'checkInfoBorrow = true

'Dim showMessageBox
'showMessageBox = true

Dim StrRegistry
Dim StatusRegistry
on error resume next

' ----Script execution starts here----

StrRegistry = WshShell.RegRead(strKeyClean1)
if not StrRegistry = "" then 
  If Run_Mode <> "DEBUG" then WshShell.RegDelete(strKeyClean1) end if
  StatusMessage("adsk_CleanUpLicenseRegKeys: Delete key: " & strKeyClean1 & ": " & StrRegistry)
end if

StrRegistry = WshShell.RegRead(strKeyClean2)
if not StrRegistry = "" then 
  If Run_Mode <> "DEBUG" then WshShell.RegDelete(strKeyClean2) end if
  StatusMessage("adsk_CleanUpLicenseRegKeys: Delete key: " & strKeyClean2 & ": " & StrRegistry)
end if

StrRegistry = WshShell.RegRead(strKeyClean3)
if not StrRegistry = strKeyValue then
  If Run_Mode <> "DEBUG" then WshShell.RegDelete(strKeyClean3) end if
  StatusRegistry = "D"
  StatusMessage("adsk_CleanUpLicenseRegKeys: Delete key: " & strKeyClean3 & ": " & StrRegistry)
end if

StrRegistry = WshShell.RegRead(strKeyClean4)
if not StrRegistry = strKeyValue then
  If Run_Mode <> "DEBUG" then WshShell.RegDelete(strKeyClean4) end if
  StatusMessage("adsk_CleanUpLicenseRegKeys: Delete key: " & strKeyClean4 & ": " & StrRegistry)
  StatusRegistry = "D"
end if

' --------
'Enumerate All subkeys in HKEY_USERS
objRegistry.EnumKey HKEY_USERS, "", arrSubkeys

Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

oReg.EnumKey HKEY_USERS, "", sidList
For Each sid In sidList
 	If oReg.EnumKey(HKEY_USERS, sid & ContractExpress, found) = 0 Then   		
           strKeyClean9 = "HKEY_USERS\" & sid & ContractExpress & "\" & strKeyName
           StrRegistry = WshShell.RegRead(strKeyClean9)
           if not StrRegistry = "" then 
              If Run_Mode <> "DEBUG" then WshShell.RegDelete(strKeyClean9) end if
              StatusMessage("adsk_CleanUpLicenseRegKeys: Delete key: " & strKeyClean9 & ": " & StrRegistry)
           end if
	End If
NEXT

End Sub 'adsk_CleanUpLicenseRegKeys
 
Sub adsk_SetLicenseRegKeys(licServer1, licServer2, TheMessage)
    StatusMessage("adsk_SetLicenseRegKeys: ><((((º>  --------------------")
	StatusMessage("adsk_SetLicenseRegKeys: License server1: " & licServer1)
	StatusMessage("adsk_SetLicenseRegKeys: License server2: " & licServer2)
	StatusMessage("adsk_SetLicenseRegKeys: " & TheMessage)

'This key will be modified
strHKLM = "HKLM"
strKeyModify = "System\CurrentControlSet\Control\Session Manager\Environment"
strKeyName1 = "ADSKFLEX_LICENSE_FILE"
strKeyName2 = "FLEXLM_TIMEOUT"
'This is the new Server variable to set

if not licServer1 = "" then 
  strKeyValue = "2080@" & licServer1
end if

if not licServer2 = "" then 
  strKeyValue = strKeyValue & ";2080@" & licServer2
end if
StatusMessage("adsk_SetLicenseRegKeys: Registry Key value: " & strKeyValue)
on error resume next

' --------
'Writes the main registry control key
  adsk_Registry_AddValue strHKLM, strKeyModify, strKeyName1, strKeyValue,"REG_SZ", strKeyName1 & " - " & TheMessage
  adsk_Registry_AddValue strHKLM, strKeyModify, strKeyName2, "2000000"  ,"REG_SZ", strKeyName2 & " - " & TheMessage

MSGBOX "Registry addition"
	
' MsgBox "please, restart your pc ... " 
' Enable to restart immediately
'  wshShell.run "cmd.exe /C shutdown /r /f /t 05 "
' -----------

END Sub 'adsk_SetLicenseRegKeys

	StatusMessage("  ")
	StatusMessage("----- End logging installation: -------------")
	StatusMessage("-- Closing Logfile .´¯`.¸¸.´¯`.¸ ><((((º> ---")	
	Logfile.Close
	SET FSO = Nothing	
	If Run_Mode = "DEBUG" then msgbox("OK! Finished. Look in log file " & LOGFILENAME) end if