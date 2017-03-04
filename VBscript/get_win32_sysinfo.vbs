Option Explicit
'==============================================================================
' LANG:		VBScript
' NAME:		get_sysinfo_win32.vbs
' VERSION:	1.0
' DATE: 	2016-01-22
' AUTHORS:
'	Rui Sousa <ruib.sousa@gmail.com>
'
' DESCRIPTION: 
' This script is intended to collect system information from Servers
' installed on Microsoft Operating Systems.
'
' NOTES:
' This script was developed based on version 1.0 developed by Daniel Moya
'
' USAGE:
'       cscript /nologo get_sysinfo_win32.vbs
'
'==============================================================================
' CHANGES:
'	Date		Description
' 	
'==============================================================================
' TODO:
' 	* Improve function to detect type of Hypervisor that the system is 
'	running on (function: isVM)
'	* Add arguments support for the plugin
'	* Add arguments support for the script help
'	* Add arguments support to disable Debug to output file
'==============================================================================

'==============================================================================
' Global constant and variable declarations
'==============================================================================
'------------------------------------------------------------------------------
' Constants
Const SCRIPT_VERSION = "2.0"
Const SCRIPT_NAME 	= "sysinfo"
Const OUTPUT_EXT 	= ".output"		' Output file extension

Const FORREADING 	= 1
Const FORWRITING	= 2
Const FORAPPENDING 	= 8

Const NOLOG			= 0
Const INFORMATIONAL = 1
Const ERRORS 		= 2
Const DEBUGGING 	= 3
Const OUTPUT 		= 4

'------------------------------------------------------------------------------
' Variables

' Variables for routines
Dim errGatherWMIInformation

' Variables for log routines
Dim	oLogFSO, oLogFile					' Global Log file system object
Dim oShell								' Global Shell system object
Dim sLogFile							' Global Log File name
Dim sLogFileLocation					' Global Path to the log file
Dim bEnableDebugLog						' Global control of Debug of the script to file
Dim bIncludeDateStamp					' Global boolean to control the insert of Date Time in log file
Dim bAppendDateStampInLogFileName		' Global boolean to control the insert of Date Time in log filename
Dim sLogFilePath						' Global string for full log file path and name
Dim sScriptName							' Script file name

' Objects for WMI
Dim oWMIService, colItems, oItem, oProperty

' Variables for script options
Dim bAllowErrors
Dim bInvalidArgument
Dim bDisplayHelp
Dim bSaveFile
Dim bAlternateCredentials
Dim bWMIBios
Dim bWMIApplications
Dim bWMIServices
Dim sWMIComputer						' Computer Name to connect through WMI

' Username and Password
Dim sUserName, sPassword

' Variables for Script Results 
Dim sHostname
Dim sOperatingSystem_Caption
Dim sServerModel
Dim sProcessorModel
Dim bIsVM
Dim sVMPlatform
Dim iNumberOfSockets
Dim	iNumberOfCores
Dim sSystem_IdentifyingNumber

'==============================================================================
'  Main routines
'==============================================================================
' Enable custom error handling
'On Error Resume Next

If LCase (Right (WScript.FullName, 11)) <> "cscript.exe" Then
    MsgBox "This script should be run from a command line (eg ""cscript.exe /nologo" & Wscript.ScriptName & """)", vbCritical, "Error"
    WScript.Quit (1)
End If


Set oShell = CreateObject( "WScript.Shell" )
sScriptName = oShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" ) & "_" & SCRIPT_NAME	' Get the name of the script being run
Set oShell = Nothing



GetOptions	' Get Options from useer

If (bDisplayHelp) Then
	DisplayHelp
Else

	If (sWMIComputer <> "") Then
		' Run the GatherWMIInformation() function and return the status
		' to errGatherInformation, if the function fails then the
		' rest is skipped. 
		errGatherWMIInformation = GatherWMIInformation()
		
		If errGatherWMIInformation = True Then
		
			' Send script result to output file
			DebugReportProgress "", vbCrLf & string(80,"*")
			DebugReportProgress "", " Results:"
			DebugReportProgress "", string(80,"*")
			DebugReportProgress "", "System Name: " & sHostname
			DebugReportProgress "", "System Model: " & sServerModel
			DebugReportProgress "", "Operating System: " & sOperatingSystem_Caption
			DebugReportProgress "", "CPU Model: " & sProcessorModel
			DebugReportProgress "", "Sockets: " & iNumberOfSockets
			DebugReportProgress "", "Cores per Socket: "  & iNumberOfCores
			DebugReportProgress "", "Type: "  & sVMPlatform
			DebugReportProgress "", "Serial Number: "  & sSystem_IdentifyingNumber
			DebugReportProgress "", string(80,"*")
			
			' servername|os|servermodel|cpumodel|nsockets|corespsocket|type|serialnumber
			DebugReportProgress "", sHostname & "|" & sOperatingSystem_Caption & "|" & _
				sServerModel & "|" & sProcessorModel & "|" & iNumberOfSockets & "|" & _
				iNumberOfCores & "|" & sVMPlatform & "|" & sSystem_IdentifyingNumber 
				
			
			Wscript.Echo "The script has been executed successfully."
			Wscript.Echo vbCrLf & "Please return the file " & sLogFile & " to eProseed."
		Else
		
			WScript.Echo "The script could not run successfully."
			WScript.Quit (1)
		End If
		
	Else
		WScript.Echo "ERROR: The computer to connect WMI provider is not defined. Please check script!"
		WScript.Quit (1)
	End If

End If

oLogFile.close
Set oLogFile = Nothing
Set oLogFSO = Nothing

WScript.Quit

'==============================================================================
'  End of Main routines
'==============================================================================

'==============================================================================
'  Functions and Procedures
'==============================================================================
'------------------------------------------------------------------------------
' Define default settings to run the script
Sub GetOptions()
	' Variables declaration
	Dim oArgs
	Dim nArgs
	
	' Default settings
	sWMIComputer = "."					' Set ComputerName to connect to WMI Namespace
	
	bEnableDebugLog 		= True		' You can disable logging globally by setting the bEnableDebugLog option to false.
	bInvalidArgument 		= False		' Control if the script arguments are valid
	bDisplayHelp 			= False
	bSaveFile 				= True		'
	bAllowErrors 			= False		' Define how to execution errors
	bAlternateCredentials 	= False		' Define if there's alternate credentials to connect to WMI Provider
	
	bWMIBios				= True		' With this option enabled the script will gather BIOS Information
	bWMIApplications		= True		' With this option enabled the script will gather Installed Application Information
	bWMIServices			= True		' With this option enabled the script will gather Services Information
	bIncludeDateStamp 				= False	' Setting this to true will time stamp each message that is logged to the output file with the current date and time.
	bAppendDateStampInLogFileName 	= True	' This will set the log file name to the current date and time. 
	
	
	' Check script arguments
	Set oArgs = WScript.Arguments
	If (oArgs.Count > 0) Then
		For nArgs = 0 To oArgs.Count - 1
			' Change settings based on arguments
			SetOptions objArgs(nArg)
		Next
	End If

End Sub ' GetOptions

'------------------------------------------------------------------------------
' Set script settings based on the arguments
Sub SetOptions(strOption)
	' Variables Declaration
	Dim nArguments
	
	' Handle arguments that are passed in the script
	' TO-DO: Handle argument to integrate with plugin 
	
	'
	'nArguments = Len(strOption)
	'If (nArguments < 2) Then
	'	bInvalidArgument = True
	'Else	
	'End If
	
	
End Sub

Sub DisplayHelp
End Sub


' Function to convert WMI time to "normal" time.
Function ConvertWMIDate(dUTCDate)
	ConvertWMIDate = CDate(Mid(dUTCDate, 5, 2) & "/" &  Mid(dUTCDate, 7, 2) & "/" & Left(dUTCDate, 4) & " " & _
                          Mid (dUTCDate, 9, 2) & ":" &  Mid(dUTCDate, 11, 2) & ":" & Mid(dUTCDate, 13, 2))
End Function

Function CreateLogFile (ByVal sScriptFileName)
	Dim oLogShell
	
	Dim sNow

	' If Logging is not enabled exit procedure
    If bEnableDebugLog = False Then Exit Function
 
    Set oLogFSO = CreateObject("Scripting.FileSystemObject")
   
	Set oLogShell = CreateObject("Wscript.Shell")
	sLogFileLocation = oLogShell.CurrentDirectory & "\"	
	Set oLogShell = Nothing
   
    If bAppendDateStampInLogFileName Then
        sNow = Replace(Replace(Replace(Replace(Now(),"/",""),":","")," ","T"),"-","")
        sScriptFileName =  sScriptFileName & "_" & sNow & OUTPUT_EXT
        bAppendDateStampInLogFileName = False
	Else
		 sScriptFileName = sScriptFileName & OUTPUT_EXT
    End If 
   
    sLogFile = sScriptFileName
	sLogFilePath = sLogFileLocation & sScriptFileName
		'wscript.echo sScriptFileName & vbCrLf
	'wscript.echo sLogFilePath & vbCrLf
	Set oLogFile = oLogFSO.OpenTextFile(sLogFilePath, FORWRITING, True)
	CreateLogFile = sLogFile
End Function

Sub DebugReportProgress(ByVal iLogType, ByVal sMessage)
	' Variables Declaration
	Dim sLogTypeMsg, sMessagePrefix
	dim oTempFSO
	
	' If Debug Log is not enabled exit procedure
    If bEnableDebugLog = False Then Exit Sub
	
	' Check if log file is created
	If sLogFile = "" then
		sLogFilePath = CreateLogFile (sScriptName)	' Create Log File
	End if	
	
	Select case iLogType
		Case INFORMATIONAL
			sLogTypeMsg = "Information" & VbTab
		Case ERRORS
			sLogTypeMsg = "ERROR" & VbTab
		Case OUTPUT
			sLogTypeMsg = ""
	End Select

    If bIncludeDateStamp and iLogType <> OUTPUT Then
        sMessagePrefix = sLogTypeMsg & Now & "   "
	Else
		If iLogType <> OUTPUT Then
			sMessagePrefix = sLogTypeMsg & "   "
		End if
    End If
	
	If sLogTypeMsg <> "" Then
		sMessage = sMessagePrefix & "   " & sMessage
	End If
	
	'If  oLogFSO.GetFile(sLogFilePath).Size <> 0 Then
		sMessage = vbCrLf & sMessage
	'End If
	
	' Write information to the file
	oLogFile.Write(sMessage)
	
End Sub

Function GetFileSize(strFolder, strFile ) 
	Dim lngFSize, lngDSize
	Dim oFO
	Dim oFD
	Dim OFS
	lngFSize = 0
	Set OFS = CreateObject("Scripting.FileSystemObject")

	If strFolder = "" Then strFolder = ActiveWorkbook.Path
	If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
   
	If OFS.FolderExists(strFolder) Then
		'If Not IsMissing(strFile) Then
	
		If OFS.FileExists(strFolder & strFile) Then
			Set oFO = OFS.GetFile(strFolder & strFile)
			GetFileSize = oFO.Size
		End If
	End If
   
End Function   '*** GetFileSize ***

' Procedure to send error to console
Sub ErrorHandler(ByVal sMessage)
		
	sMessage = sMessage & vbCrLf & "Error: " & Err.Number & vbCrLf _
		& "Error (Hex): " & Hex(Err.Number) & vbCrLf _
		& "Source: " &  Err.Source  & vbCrLf _
		& "Description: " &  Err.Description
		
	WScript.Echo sMessage
	DebugReportProgress ERRORS, sMessage
	
End Sub


' Function to Gather information from WMI Provider
Function GatherWMIInformation()
	Dim oSWbemLocator
	Dim dwUTCPlaceHolder
	Dim sMessage
	Dim sOS_InstallDate, arrOSystem_Name
	Dim sOS_ServicePack, sOS_LanguageCode, sOS_Version
	Dim iOSVersion
	Dim sOperatingSystem_WindowsDirectory
	Dim sSystem_Manufacturer, sSystem_Name
	Dim sBIOS_SMBIOSBIOSVersion, sBIOS_SMBIOSMajorVersion, sBIOS_SMBIOSMinorVersion
	Dim sBIOS_Manufacturer, sBIOS_Version, sBiosCharacteristics
	Dim arrBIOS_BiosCharacteristics
	ReDim arrBIOS_BiosCharacteristics(0)
	Dim i, iChassisType, iNumberOfLogicalProcessors, iTempNumberOfCores, iTempNumberOfLogicalProcessors
	Dim bProcessorHTSystem
	
	' Define how to handle errors
	If (bAllowErrors) Then
		On Error Resume Next
	End If
	
	DebugReportProgress "", "Starting exection of script: " & Wscript.ScriptName
	DebugReportProgress "", "Script Name: " & SCRIPT_NAME
	DebugReportProgress "", "Script Version: " & SCRIPT_VERSION
	DebugReportProgress "", "Log file: " & sLogFilePath
	DebugReportProgress "", "Execution start time: " & Now()
	DebugReportProgress "", "Start subroutine: GatherWMIInformation(" & sWMIComputer & ")"
	
	
	'Dim arrGroupUser
	If (bAlternateCredentials) Then
		Set oSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		Set oWMIService = oSWbemLocator.ConnectServer(sWMIComputer,"root\cimv2",sUserName,sPassword)
	Else
		Set oWMIService = GetObject("winmgmts:\\" & sWMIComputer & "\root\cimv2")
	End If
	If (Err <> 0) Then ' Failed to Bind to WMI provider
	    ReportProgress "ERROR: Unable to bind to WMI provider on " & sWMIComputer & "."
	    GatherWMIInformation = False
		
		Err.Clear
	    Exit Function
	End If

	DebugReportProgress "", vbCrLf & string(80,"*")
	DebugReportProgress "", "SYSTEM INFORMATION:"
	DebugReportProgress "", string(80,"*")
	DebugReportProgress  "", "Gathering OS information"
	
	Set colItems = oWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For Each oItem In colItems
		DebugReportProgress  "", "Current User Name: " & oItem.UserName
		DebugReportProgress  "", "Computer Name: " & oItem.Name
		DebugReportProgress  "", "Domain: " & oItem.Domain
		DebugReportProgress  "", "Workgroup: " & oItem.Workgroup
		sHostname = oItem.Name
	Next
	
	Set colItems = oWMIService.ExecQuery("Select Name, CSDVersion, InstallDate, OSLanguage, Version, WindowsDirectory from Win32_OperatingSystem",,48)
	For Each oItem in colItems
		sOS_InstallDate = oItem.InstallDate
		arrOSystem_Name = Split(oItem.Name,"|")
		sOperatingSystem_Caption = arrOSystem_Name(0)
		sOS_ServicePack = oItem.CSDVersion
		sOS_LanguageCode = Clng(oItem.OSLanguage)
		sOS_LanguageCode = Hex(sOS_LanguageCode)
		sOS_Version = oItem.Version
		sOperatingSystem_WindowsDirectory = oItem.WindowsDirectory
	Next
	
	DebugReportProgress "", "OS Name: " & sOperatingSystem_Caption
	DebugReportProgress "", "Service Pack: " & sOS_ServicePack
	DebugReportProgress "", "OS Version: " & sOS_Version
	DebugReportProgress "", "Windows is installed at " & sOperatingSystem_WindowsDirectory
	DebugReportProgress "", "Install date: " & ConvertWMIDate(sOS_InstallDate)
	DebugReportProgress "", "Operating System Language: " & ReturnOperatingSystemLanguage(sOS_LanguageCode)
	DebugReportProgress "", "Time Zone: " & GetWmiPropertyValue("root\cimv2", "Win32_TimeZone", "Description")
	
	
	' Before continue with the script check requirements
	' Check if the OS is supported
	If (checkOSSupport (sOS_Version) = False) Then
		DebugReportProgress "", "ERROR: The OS version is not supported. We won't continue with the script execution." &_
			vbCrLf & "The minimum OS supported is Windows Server 2003 (with SP1) with hotfix KB932370 installed." & vbCrLf
			
		Wscript.Echo "ERROR: The OS version is not supported. We won't continue with the script execution."
		GatherWMIInformation = False
		Exit Function
	End If
	
	' Check if when OS is WS2K3 the required hotfixes are installed
	' OS Version = 5.2
	' In Windows Server 2003 it's required HotFix KB932370 to support 
	' Number of Sockets and Number of Sockets gathering in WMI
	If ConvertOSVersion2Number(sOS_Version) = 52 then
		' Check hotfix KB932370 is installed
		If CheckHotfix ("932370") <> True Then
			DebugReportProgress "", "ERROR: The OS version is Windows Server 2003. It's required to have KB932370 installed."
			
			Wscript.Echo "ERROR: The required hotfixes KB932370 are not installed on Windows Server 2003." & _
				vbCrLf & "The minimum OS supported is Windows Server 2003 (with SP1) with hotfix KB932370 installed." & vbCrLf & _
				vbCrLf & "Please install the requirements and run the script again." & vbCrLf
			GatherWMIInformation = False
			Exit Function		
		End If
	End If
	
	DebugReportProgress "", vbCrLf & string(80,"*")
	DebugReportProgress "", "Gathering information about BIOS and Manufacturer:"
	DebugReportProgress "", string(80,"*")
	
	' Gather information from the computer system
	DebugReportProgress "", ">> Gathering computer system product information:"
	Set colItems = oWMIService.ExecQuery("Select Vendor, Name, IdentifyingNumber from Win32_ComputerSystemProduct",,48)
	For Each oItem in colItems
	    sSystem_Manufacturer = oItem.Vendor
	    sSystem_Name = oItem.Name
	    sSystem_IdentifyingNumber = Trim(oItem.IdentifyingNumber)
	Next
	sServerModel = GetWmiPropertyValue("root\cimv2", "Win32_ComputerSystem", "Model")
	
	DebugReportProgress "", "Manufacturer: " & sSystem_Manufacturer
	DebugReportProgress "", "Product name: " & sSystem_Name
	DebugReportProgress "", "Identifying Number: " & sSystem_IdentifyingNumber
	DebugReportProgress "", "Model: " & sServerModel
	DebugReportProgress "", "Is VM: " & isVM
	DebugReportProgress "", "Type: " & sVMPlatform
	
	DebugReportProgress "", vbCrLf & ">> Gathering system enclosure information:"
	Set colItems = oWMIService.ExecQuery("Select ChassisTypes from Win32_SystemEnclosure",,48)
	For Each oItem in colItems
	    For i = Lbound(oItem.ChassisTypes) to Ubound(oItem.ChassisTypes)
	        iChassisType = oItem.ChassisTypes(i)
	    Next
	Next
	DebugReportProgress "", "Chassis Type: " & getChassisType(iChassisType)
	
	
	' Gather information from BIOS
	If (bWMIBios) Then
		DebugReportProgress "", vbCrLf & ">> Gathering BIOS information:"
		Set colItems = oWMIService.ExecQuery("Select BiosCharacteristics, SMBIOSBIOSVersion, SMBIOSMajorVersion, SMBIOSMinorVersion, Version, Manufacturer from Win32_BIOS",,48)
		For Each oItem in colItems
			sBIOS_Manufacturer = oItem.Manufacturer
			sBIOS_SMBIOSBIOSVersion = oItem.SMBIOSBIOSVersion
			sBIOS_SMBIOSMajorVersion = oItem.SMBIOSMajorVersion
			sBIOS_SMBIOSMinorVersion = oItem.SMBIOSMinorVersion
			sBIOS_Version = oItem.Version
			arrBIOS_BiosCharacteristics(0) = 3
			If (IsArray(oItem.BiosCharacteristics)) Then
				For i = 0 To Ubound(oItem.BiosCharacteristics)
					ReDim Preserve arrBIOS_BiosCharacteristics(i)
					arrBIOS_BiosCharacteristics(i) = oItem.BiosCharacteristics(i)
				Next
			End If
		Next
		
		DebugReportProgress "", "BIOS Manufacturer: "  & sBIOS_Manufacturer
		DebugReportProgress "", "BIOS Version: "  & sBIOS_Version
		DebugReportProgress "", "SMBIOS Version: "  & sBIOS_SMBIOSBIOSVersion & " (Major: " & sBIOS_SMBIOSMajorVersion & ", Minor: " & sBIOS_SMBIOSMinorVersion & ")"
		DebugReportProgress "", "BIOS Characteristics:"
		
		For i = 0 To Ubound(arrBIOS_BiosCharacteristics)
			If (sBiosCharacteristics = "") Then
				sBiosCharacteristics = ReturnBiosCharacteristic(arrBIOS_BiosCharacteristics(i))
			Else
				sBiosCharacteristics = ReturnBiosCharacteristic(arrBIOS_BiosCharacteristics(i))
			End If
			DebugReportProgress "", "    " & sBiosCharacteristics
		Next
		
	End If
	
	' Gather information from Processors
	DebugReportProgress "", vbCrLf & ">> Gathering processor information:"
	
	' Architecture property specifies the processor architecture used by this platform
	iNumberOfSockets = GetWmiPropertyValue("root\cimv2", "Win32_ComputerSystem", "NumberOfProcessors")
	iNumberOfLogicalProcessors = GetWmiPropertyValue("root\cimv2", "Win32_ComputerSystem", "NumberOfLogicalProcessors")
	' Check if Processors have  hyper-threading enable
	
	bProcessorHTSystem = False
	iTempNumberOfCores = 0
	iTempNumberOfLogicalProcessors = 0
	Set colItems = oWMIService.ExecQuery("Select NumberOfCores, NumberOfLogicalProcessors from Win32_Processor",,48)
	For Each oItem in colItems
		iTempNumberOfCores = iTempNumberOfCores + oItem.NumberOfCores
		iTempNumberOfLogicalProcessors = iTempNumberOfLogicalProcessors + oItem.NumberOfLogicalProcessors
	Next
	
	' If Total Number of Logical Processor is bigger then
	' Total number of Cores then the processors have HT enabled
	If iTempNumberOfLogicalProcessors > iTempNumberOfCores Then
		bProcessorHTSystem = True
		iNumberOfCores = iTempNumberOfCores / iNumberOfSockets
	else 
		iNumberOfCores = iNumberOfLogicalProcessors / iNumberOfSockets
	End if
	DebugReportProgress "", "Number of Sockets: " & iNumberOfSockets
	DebugReportProgress "", "Number of Cores: " & iNumberOfCores
	DebugReportProgress "", "Hyper-Threading Enabled: " & bProcessorHTSystem
	
	' Get Processors details
	'
	'Laptop
	'Description                           	ExtClock  	L2CacheSize  	MaxClockSpeed  	Name 										SocketDesignation
	'Intel64 Family 6 Model 60 Stepping 3	100			256				2301			Intel(R) Core(TM) i7-4712MQ CPU @ 2.30GHz	Onboard
	'
	'Physical Server
	'Description                           ExtClock  L2CacheSize  MaxClockSpeed  NAME									SocketDesignation
	'Intel64 Family 6 Model 29 Stepping 1  1066      6144         2400           Intel(R) Xeon(R) CPU E7440  @ 2.40GHz  Proc 1
	'Intel64 Family 6 Model 29 Stepping 1  1066      6144         2400           Intel(R) Xeon(R) CPU E7440  @ 2.40GHz  Proc 2
	'Intel64 Family 6 Model 29 Stepping 1  1066      6144         2400           Intel(R) Xeon(R) CPU E7440  @ 2.40GHz  Proc 3
	'Intel64 Family 6 Model 29 Stepping 1  1066      6144         2400           Intel(R) Xeon(R) CPU E7440  @ 2.40GHz  Proc 4
	'
	'VirtualBox
	'Description                          	ExtClock  	L2CacheSize  	MaxClockSpeed  	Name										SocketDesignation
	'Intel64 Family 6 Model 60 Stepping 3								2295 			Intel(R) Core(TM) i7-4712MQ CPU @ 2.30GHz	
	'
	'VMware
	'Description                           ExtClock  L2CacheSize  MaxClockSpeed  Name                                      SocketDesignation
	'Intel64 Family 6 Model 45 Stepping 2            0            2000           Intel(R) Xeon(R) CPU E5-2650 0 @ 2.00GHz  CPU socket #0
	'Intel64 Family 6 Model 45 Stepping 2            0            2000           Intel(R) Xeon(R) CPU E5-2650 0 @ 2.00GHz  CPU socket #1
	'
	'Hyper-V
	'Description                           	ExtClock  L2CacheSize  MaxClockSpeed  Name                                      SocketDesignation
	'Intel64 Family 6 Model 15 Stepping 11	1333                   2667           Intel(R) Xeon(R) CPU X5355  @ 2.66GHz  	None
	'
	
	If ConvertOSVersion2Number(sOS_Version) = 52 then
		' Get processors information on Windows Server = 52 (=2003)
		Set colItems = oWMIService.ExecQuery("Select Description, ExtClock, L2CacheSize, Name, MaxClockSpeed, Architecture, NumberOfCores, NumberOfLogicalProcessors from Win32_Processor",,48)
		i = 0
		For Each oItem in colItems
			i = i + 1
			DebugReportProgress "", "Processor #" & i & " Information:"
			DebugReportProgress "", "   Name: " & oItem.Name
			DebugReportProgress "", "   Description: " & oItem.Description
			DebugReportProgress "", "   Speed: " & oItem.MaxClockSpeed
			DebugReportProgress "", "   L2 Cache Size: " & oItem.L2CacheSize
			DebugReportProgress "", "   External clock: " & oItem.MaxClockSpeed
			DebugReportProgress "", "   Architecture: " & getSystemArchitectureType(oItem.Architecture)
			DebugReportProgress "", "   Number of Cores: " & oItem.NumberOfCores
			DebugReportProgress "", "   Number of Logical Processors: " & oItem.NumberOfLogicalProcessors
			sProcessorModel = oItem.Name
		Next
	
	else 
		' Get processors information on Windows Version > 52 (2003)
		Set colItems = oWMIService.ExecQuery("Select Description, ExtClock, L3CacheSize, L2CacheSize, Name, MaxClockSpeed, Architecture, NumberOfCores, NumberOfLogicalProcessors from Win32_Processor",,48)
		i = 0
		For Each oItem in colItems
			i = i + 1
			DebugReportProgress "", "Processor #" & i & " Information:"
			DebugReportProgress "", "   Name: " & oItem.Name
			DebugReportProgress "", "   Description: " & oItem.Description
			DebugReportProgress "", "   Speed: " & oItem.MaxClockSpeed
			DebugReportProgress "", "   L2 Cache Size: " & oItem.L2CacheSize
			DebugReportProgress "", "   L3 Cache Size: " & oItem.L3CacheSize
			DebugReportProgress "", "   External clock: " & oItem.MaxClockSpeed
			DebugReportProgress "", "   Architecture: " & getSystemArchitectureType(oItem.Architecture)
			DebugReportProgress "", "   Number of Cores: " & oItem.NumberOfCores
			DebugReportProgress "", "   Number of Logical Processors: " & oItem.NumberOfLogicalProcessors
			sProcessorModel = oItem.Name
		Next
		
	End If
	
	' If bWMIApplications is True gather installed applications information
	If (bWMIApplications) and  ConvertOSVersion2Number(sOS_Version) > 52 Then
		DebugReportProgress "", vbCrLf & string(80,"*")
		DebugReportProgress "", "Gathering information about Installed Software:"
		DebugReportProgress "", string(80,"*")
		Set colItems = oWMIService.ExecQuery("Select Name, Vendor, Version, InstallDate from Win32_Product WHERE Name <> Null",,48)
		
		
		For Each oItem in colItems
			DebugReportProgress "", "Product Name:" & oItem.Name
			DebugReportProgress "", "   Vendor:" & oItem.Vendor
			DebugReportProgress "", "   Version:" & oItem.Version
			If (IsNull(oItem.InstallDate)) Then
				DebugReportProgress "", "   Install Date: N/A"
			Else
				DebugReportProgress "", "   Install Date: " & oItem.InstallDate
			End If
		Next
	End if
	
	' If bWMIApplications is True gather services information
	If (bWMIServices) Then
		DebugReportProgress "", vbCrLf & string(80,"*")
		DebugReportProgress "", "Gathering information about Services:"
		DebugReportProgress "", string(80,"*")
	
		Set colItems = oWMIService.ExecQuery("Select Caption, Name, Started, StartMode, StartName from Win32_Service Where ServiceType ='Share Process' Or ServiceType ='Own Process'",,48)
		For Each oItem in colItems
			DebugReportProgress "", "Caption: " & oItem.Caption
			DebugReportProgress "", "   Name: " & oItem.Name
			DebugReportProgress "", "   Started: " & oItem.Started
			DebugReportProgress "", "   Start Mode: " & oItem.StartMode
			DebugReportProgress "", "   Start Name: " & oItem.StartName
		Next
	End If
		
		
	Set oWMIService = Nothing
	Set oSWbemLocator = Nothing
	DebugReportProgress "", string(80,"*")
	DebugReportProgress "", "End subroutine: GatherWMIInformation()"
	GatherWMIInformation = True
	
End Function ' GatherWMIInformation


Function GetWmiPropertyValue(sNameSpace, sClassName, sPropertyName)
	' Variables Declaration
	Dim sLine, sPropertyValue
	
    'On Error Resume Next
    sPropertyValue = ""

    Set colItems = oWMIService.ExecQuery("Select * from " & sClassName,,48)
    For Each oItem in colItems
        For Each oProperty in oItem.Properties_
            sLine = ""
            If oProperty.Name = sPropertyName Then
                If oProperty.IsArray = True Then
                    sLine = "s" & objProperty.Name & " = Join(oItem." & oProperty.Name & ", " & Chr(34) & "," & Chr(34) & ")" & vbCrLf
                    sLine = sLine & "sPropertyValue =  s" & oProperty.Name
                Else
                    sLine =  "sPropertyValue =  oItem." & oProperty.Name
                End If
                Execute sLine
            End If
        Next
    Next
    GetWmiPropertyValue = sPropertyValue
End Function

Function ReturnOperatingSystemLanguage(strOSLanguageCode)
	Dim sOSLanguageName, sOSTempCode
	sOSTempCode = Cstr(strOSLanguageCode)
	Select Case strOSLanguageCode
		Case "1"		sOSLanguageName = "Arabic"
		Case "4"   		sOSLanguageName = "Chinese"
		Case "9"   		sOSLanguageName = "English"
		Case "401" 		sOSLanguageName = "Arabic - Saudi Arabia"
		Case "402" 		sOSLanguageName = "Bulgarian"
		Case "403" 		sOSLanguageName = "Catalan"
		Case "404"		sOSLanguageName = "Chinese - Taiwan"
		Case "405" 		sOSLanguageName = "Czech"
		Case "406" 		sOSLanguageName = "Danish"
		Case "407" 		sOSLanguageName = "German"
		Case "408" 		sOSLanguageName = "Greek"
		Case "409"		sOSLanguageName = "English" 
		Case "40A" 		sOSLanguageName = "Spanish - Traditional Sort"
		Case "40B" 		sOSLanguageName = "Finnish"
		Case "40C" 		sOSLanguageName = "French - France"
		Case "40D" 		sOSLanguageName = "Hebrew"
		Case "40E"		sOSLanguageName = "Hungarian"
		Case "40F" 		sOSLanguageName = "Icelandic"
		Case "410" 		sOSLanguageName = "Italian - Italy"
		Case "411" 		sOSLanguageName = "Japanese"
		Case "412" 		sOSLanguageName = "Korean"
		Case "413" 		sOSLanguageName = "Dutch - Netherlands"
		Case "414" 		sOSLanguageName = "Norwegian - Bokmal"
		Case "415" 		sOSLanguageName = "Polish"
		Case "416" 		sOSLanguageName = "Portuguese - Brazil"
		Case "417" 		sOSLanguageName = "Rhaeto-Romanic"
		Case "418" 		sOSLanguageName = "Romanian"
		Case "419" 		sOSLanguageName = "Russian"
		Case "41A" 		sOSLanguageName = "Croatian"
		Case "41B" 		sOSLanguageName = "Slovak"
		Case "41C" 		sOSLanguageName = "Albanian"
		Case "41D" 		sOSLanguageName = "Swedish"
		Case "41E" 		sOSLanguageName = "Thai"
		Case "41F" 		sOSLanguageName = "Turkish"
		Case "420" 		sOSLanguageName = "Urdu"
		Case "421" 		sOSLanguageName = "Indonesian"
		Case "422" 		sOSLanguageName = "Ukrainian"
		Case "423" 		sOSLanguageName = "Belarusian"
		Case "424" 		sOSLanguageName = "Slovenian"
		Case "425" 		sOSLanguageName = "Estonian"
		Case "426" 		sOSLanguageName = "Estonian"
		Case "426" 		sOSLanguageName = "Latvian"
		Case "427" 		sOSLanguageName = "Lithuanian"
		Case "429" 		sOSLanguageName = "Persion"
		Case "42A" 		sOSLanguageName = "Vietnamese"
		Case "42D" 		sOSLanguageName = "Basque"
		Case "42E"		sOSLanguageName = "Serbian"
		Case "42F" 		sOSLanguageName = "Macedonian (FYROM)"
		Case "430" 		sOSLanguageName = "Sutu"
		Case "431" 		sOSLanguageName = "Tsonga"
		Case "432" 		sOSLanguageName = "Tswana"
		Case "434" 		sOSLanguageName = "Xhosa"
		Case "435" 		sOSLanguageName = "Zulu"
		Case "436" 		sOSLanguageName = "Afrikaans"
		Case "438" 		sOSLanguageName = "Faeroese"
		Case "43A" 		sOSLanguageName = "Maltese"
		Case "43C" 		sOSLanguageName = "Gaelic"
		Case "43D" 		sOSLanguageName = "Yiddish"
		Case "43E" 		sOSLanguageName = "Malay - Malaysia"
		Case "801" 		sOSLanguageName = "Arabic - Iraq"
		Case "804" 		sOSLanguageName = "Chinese - PRC"
		Case "807" 		sOSLanguageName = "German - Switzerland"
		Case "809"		sOSLanguageName = "English - United Kingdom"
		Case "80A" 		sOSLanguageName = "Spanish - Mexico"
		Case "80C" 		sOSLanguageName = "French - Belgium"
		Case "810" 		sOSLanguageName = "Italian - Switzerland"
		Case "813" 		sOSLanguageName = "Dutch - Belgium"
		Case "814" 		sOSLanguageName = "Norwegian - Nynorsk"
		Case "816" 		sOSLanguageName = "Portuguese - Portugal"
		Case "818" 		sOSLanguageName = "Romanian - Moldova"
		Case "819" 		sOSLanguageName = "Russian - Moldova"
		Case "81A" 		sOSLanguageName = "Serbian - Latin"
		Case "81D" 		sOSLanguageName = "Swedish - Finland"
		Case "C01" 		sOSLanguageName = "Arabic - Egypt"
		Case "C04" 		sOSLanguageName = "Chinese - Hong Kong SAR"
		Case "C07" 		sOSLanguageName = "German - Austria"
		Case "C09" 		sOSLanguageName = "English - Australia"
		Case "C0A" 		sOSLanguageName = "Spanish - International Sort"
		Case "C0C" 		sOSLanguageName = "French - Canada"
		Case "C1A" 		sOSLanguageName = "Serbian - Cyrillic"
		Case "1004"		sOSLanguageName = "Chinese - Singapore"
		Case "1007"		sOSLanguageName = "German - Luxembourg"
		Case "1009"		sOSLanguageName = "English - Canada"
		Case "100A"    	sOSLanguageName = "Spanish - Guatemala"
		Case "100C"    	sOSLanguageName = "French - Switzerland"
		Case "1401"    	sOSLanguageName = "Arabic - Algeria"
		Case "1409"    	sOSLanguageName = "English - New Zealand"
		Case "140A"    	sOSLanguageName = "Spanish - Costa Rica"
		Case "140C"    	sOSLanguageName = "French - Luxembourg"
		Case "1801"    	sOSLanguageName = "Arabic - Morocco"
		Case "1809"    	sOSLanguageName = "English - Ireland"
		Case "180A"    	sOSLanguageName = "Spanish - Panama"
		Case "1C01"    	sOSLanguageName = "Arabic - Tunisia"
		Case "1C09"    	sOSLanguageName = "English - South Africa"
		Case "1C0A"    	sOSLanguageName = "Spanish - Dominican Republic"
		Case "2001"    	sOSLanguageName = "Arabic - Oman"
		Case "2009"    	sOSLanguageName = "English - Jamaica"
		Case "200A"    	sOSLanguageName = "Spanish - Venezuela"
		Case "2401"    	sOSLanguageName = "Arabic - Yemen"
		Case "240A"    	sOSLanguageName = "Spanish - Colombia"
		Case "2801"    	sOSLanguageName = "Arabic - Syria"
		Case "2809"    	sOSLanguageName = "English - Belize"
		Case "280A"    	sOSLanguageName = "Spanish - Peru"
		Case "2C01"    	sOSLanguageName = "Arabic - Jordan"
		Case "2C09"    	sOSLanguageName = "English - Trinidad"
		Case "2C0A"    	sOSLanguageName = "Spanish - Argentina"
		Case "3001"    	sOSLanguageName = "Arabic - Lebanon"
		Case "300A"    	sOSLanguageName = "Spanish - Ecuador"
		Case "3401"    	sOSLanguageName = "Arabic - Kuwait"
		Case "340A"    	sOSLanguageName = "Spanish - Chile"
		Case "3801"    	sOSLanguageName = "Arabic - U.A.E."
		Case "380A"    	sOSLanguageName = "Spanish - Uruguay"
		Case "3C01"    	sOSLanguageName = "Arabic - Bahrain"
		Case "3C0A"    	sOSLanguageName = "Spanish - Paraguay"
		Case "4001"    	sOSLanguageName = "Arabic - Qatar"
		Case "400A"    	sOSLanguageName = "Spanish - Bolivia"
		Case "440A"    	sOSLanguageName = "Spanish - El Salvador"
		Case "480A"    	sOSLanguageName = "Spanish - Honduras"
		Case "4C0A"    	sOSLanguageName = "Spanish - Nicaragua"
		Case "500A"    	sOSLanguageName = "Spanish - Puerto Rico"
		Case Else		sOSLanguageName = "Unknown"
	End Select
	
	' Return result
	ReturnOperatingSystemLanguage = sOSLanguageName	
End Function ' ReturnOperatingSystemLanguage


Function ReturnBiosCharacteristic(nBiosCharacteristic)
	' Variables declaration
	Dim sBiosCharacteristic
	
	' Get the description for the Bios Characteristic
	Select Case nBiosCharacteristic
		Case 0		sBiosCharacteristic = "Reserved" 
		Case 1		sBiosCharacteristic = "Reserved" 
		Case 2		sBiosCharacteristic = "Unknown"
		Case 3		sBiosCharacteristic = "BIOS Characteristics Not Supported"
		Case 4		sBiosCharacteristic = "ISA is supported"
		Case 5		sBiosCharacteristic = "MCA is supported"
		Case 6		sBiosCharacteristic = "EISA is supported"
		Case 7		sBiosCharacteristic = "PCI is supported"
		Case 8		sBiosCharacteristic = "PC Card (PCMCIA) is supported"
		Case 9		sBiosCharacteristic = "Plug and Play is supported"
		Case 10		sBiosCharacteristic = "APM is supported"
		Case 11		sBiosCharacteristic = "BIOS is Upgradable (Flash)"
		Case 12		sBiosCharacteristic = "BIOS shadowing is allowed"
		Case 13		sBiosCharacteristic = "VL-VESA is supported"
		Case 14		sBiosCharacteristic = "ESCD support is available"
		Case 15		sBiosCharacteristic = "Boot from CD is supported"
		Case 16		sBiosCharacteristic = "Selectable Boot is supported"
		Case 17		sBiosCharacteristic = "BIOS ROM is socketed"
		Case 18		sBiosCharacteristic = "Boot From PC Card (PCMCIA) is supported"
		Case 19		sBiosCharacteristic = "EDD (Enhanced Disk Drive) Specification is supported"
		Case 20		sBiosCharacteristic = "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported"
		Case 21		sBiosCharacteristic = "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported"
		Case 22		sBiosCharacteristic = "Int 13h - 5.25 / 360 KB Floppy Services are supported"
		Case 23		sBiosCharacteristic = "Int 13h - 5.25 /1.2MB Floppy Services are supported"
		Case 24		sBiosCharacteristic = "13h - 3.5 / 720 KB Floppy Services are supported"
		Case 25		sBiosCharacteristic = "Int 13h - 3.5 / 2.88 MB Floppy Services are supported"
		Case 26		sBiosCharacteristic = "Int 5h, Print Screen Service is supported"
		Case 27		sBiosCharacteristic = "Int 9h, 8042 Keyboard services are supported"
		Case 28		sBiosCharacteristic = "Int 14h, Serial Services are supported"
		Case 29		sBiosCharacteristic = "Int 17h, printer services are supported"
		Case 30		sBiosCharacteristic = "Int 10h, CGA/Mono Video Services are supported"
		Case 31		sBiosCharacteristic = "NEC PC-98"
		Case 32		sBiosCharacteristic = "ACPI supported"
		Case 33		sBiosCharacteristic = "USB Legacy is supported"
		Case 34		sBiosCharacteristic = "AGP is supported"
		Case 35		sBiosCharacteristic = "I2O boot is supported"
		Case 36		sBiosCharacteristic = "LS-120 boot is supported"
		Case 37		sBiosCharacteristic = "ATAPI ZIP Drive boot is supported"
		Case 38		sBiosCharacteristic = "1394 boot is supported"
		Case 39		sBiosCharacteristic = "Smart Battery supported"
		Case Else	sBiosCharacteristic = "Unknown (Undocumented)"
	End Select
	
	' Return result
	ReturnBiosCharacteristic = sBiosCharacteristic
End Function ' ReturnBiosCharacteristic

' Check if system is s Virtual Machine
Function isVM
	
	' Type of Hypervisors to detect
	'	- hyperv Microsoft Hyper-V 
	'	- kvm Linux Kernel Virtual Machine (KVM) 
	'	- openvz OpenVZ or Virtuozzo 
	'	- powervm_lx86 IBM PowerVM Lx86 Linux/x86 emulator 
	'	- qemu QEMU (unaccelerated) 
	'	- uml User-Mode Linux (UML) 
	'	- virtage Hitachi Virtualization Manager (HVM) Virtage LPAR 
	'	- virtualbox VirtualBox 
	'	- virtualpc Microsoft VirtualPC 
	'	- vmware VMware 
	'	- xen Xen 
	'	- xen-dom0 Xen dom0 (privileged domain) 
	'	- xen-domU Xen domU (paravirtualized guest domain) 
	'	- xen-hvm Xen guest fully virtualized (HVM)

	' Variables declaration
	Dim bIsVM
	
    ' Check the WMI information against known values
    bIsVM = False
    sVMPlatform = "PHYSICAL"
    
    'sModel = GetWmiPropertyValue("root\cimv2", "Win32_ComputerSystem", "Model")

	' Check of the system is a Virtual Machine
	Select Case sServerModel
		Case "Virtual Machine"
			' Microsoft virtualization technology detected, assign defaults
			sVMPlatform = "Hyper-V"
			bIsVM = True
			
		Case "VMware Virtual Platform"
			' VMware detected
			bIsVM = True
			sVMPlatform = "VMware"
			
		Case "VirtualBox"
			' VirtualBox detected
			bIsVM = True
			sVMPlatform = "VirtualBox"
		
		Case "HVM DomU"
			' Oracle Virtual Manager or Citrix XenServer
			bIsVM = True
			sVMPlatform = "Xen"
    
	End Select
    
	' Return result
	isVM = bIsVM
End Function

' Determine the description for the Chassis Type
Function getChassisType(iChassisType)
	' Variables Declaration
	Dim sChassisType
	
	' Get description for the chassis type
	Select Case iChassisType
		Case 1	sChassisType = "Other"
		Case 2	sChassisType = "Unknown"
		Case 3	sChassisType = "Desktop"
		Case 4 	sChassisType = "Low Profile Desktop"
		Case 5 	sChassisType = "Pizza Box"
		Case 6 	sChassisType = "Mini Tower"
		Case 7	sChassisType = "Tower"
		Case 8	sChassisType = "Portable"		
		Case 9	sChassisType = "Laptop"
		Case 10	sChassisType = "Notebook"
		Case 11	sChassisType = "Hand Held"
		Case 12	sChassisType = "Docking Station"
		Case 13	sChassisType = "All in One"
		Case 14	sChassisType = "Sub Notebook"
		Case 15	sChassisType = "Space-Saving"
		Case 16	sChassisType = "Lunch Box"
		Case 17	sChassisType = "Main System Chassis"
		Case 18	sChassisType = "Expansion Chassis"
		Case 19	sChassisType = "SubChassis"
		Case 20	sChassisType = "Bus Expansion Chassis"
		Case 21	sChassisType = "Peripheral Chassis"
		Case 22	sChassisType = "Storage Chassis"
		Case 23	sChassisType = "Rack Mount Chassis"
		Case 24	sChassisType = "Sealed-Case PC"
		Case Else sChassisType = "Unknown (Undocumented)"
	End Select
	
	' Return result
	getChassisType = sChassisType
End Function ' getChassisType

' Function to determine the description the processor architecture used by this platform
Function getSystemArchitectureType (nArchitectureType)

	' Variables Declaration
	Dim sArchitectureType
	
	' Get description for the chassis type
	Select Case nArchitectureType
		Case 0	sArchitectureType = "x86"
		Case 1	sArchitectureType = "MIPS"
		Case 2	sArchitectureType = "Alpha"
		Case 3 	sArchitectureType = "PowerPC"
		Case 6 	sArchitectureType = "ia64"
		Case 9 	sArchitectureType = "x64"
		Case Else sArchitectureType = "Unknown (Undocumented)"
	End Select
	
	' Return result
	getSystemArchitectureType = sArchitectureType
End Function ' End of function getSystemArchitectureType


' Function to check if a particular hotfix is installed or not. 
' This function has these 3 return options:
' TRUE, FALSE, <error description> 
Function CheckHotfix(sHotfixID)
	
	'On error resume next
	Dim oHFWMIService
	Dim sWMIforesp
	Dim colQuickFixes
	
	Set oHFWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & sWMIComputer & "\root\cimv2")
		
	if err.number <> 0 then
		CheckHotfix = "WMI could not connect to computer '" & sWMIComputer & "'"
		Exit function 'No reason to continue
	end if
	
	sWMIforesp = "Select * from Win32_QuickFixEngineering where HotFixID like 'Q" & sHotfixID &_ 
		"%' OR HotFixID like 'KB" & sHotfixID & "%'"
	Set colQuickFixes = oHFWMIService.ExecQuery (sWMIforesp)
	If err.number <> 0 Then	'if an error occurs
		CheckHotfix = "Unable to get WMI hotfix info"
	Else 'Error number 0 meaning no error occured 
		if colQuickFixes.count > 0 then
			CheckHotfix = True	'HF installed
		else 
			CheckHotfix = False	'HF not installed
		end If
	end if
	Set colQuickFixes = Nothing
	Set oHFWMIService = Nothing
	
	Err.Clear
	On Error GoTo 0
End Function ' End of Function CheckHotfix


' Check Operating System Version
' This function has these 2 return options:
' 	- True		OS is supported by the script
'	- False		OS is NOT supported by the script
Function checkOSSupport (sOSVersion)

	' OS Version:
	'	10.0	- Windows 10 or Windows Server 2016 TP
	'	6.3		- Windows 8.1 or Windows Server 2012 R2
	' 	6.2 	- Windows 8.0 or Windows Server 2012
	'	6.1 	- Windows 7 or Windows Server 2008 R2
	'	6.0		- Windows Vista or Windows Server 2008
	'	5.2		- Windows Server 2003 or Windows Server 2003 R2 or Windows XP 64-bit
	'	5.1 	- Windows XP
	' 	5.0		- Windows 2000
	' 	4.0		- Windows NT 4.0
	'	Source Info: https://msdn.microsoft.com/en-us/library/windows/desktop/ms724832(v=vs.85).aspx
	
	' Variables Declaration
	Dim iOSVersion
	
	' Convert OS Version to number
	'iOSVersion = Mid(sOSVersion,1,1) & Mid(sOSVersion,3,1)
	
	checkOSSupport = True
	
	' If OS Version previous then Windows Server 2003, then we don't support
	If ConvertOSVersion2Number (sOSVersion) < 52 then
		checkOSSupport = FALSE
	End If
	
End Function	' End of Function checkOSSupport


' Function to convert WMI time to "normal" time.
Function ConvertOSVersion2Number(sOSVersion)
	ConvertOSVersion2Number = Mid(sOSVersion,1,1) & Mid(sOSVersion,3,1)
End Function