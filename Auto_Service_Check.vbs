'==========================================================================
'
' VBScript Source File -- Created with BurnSoft BurnPad
'
' NAME: <Auto_Service_Check.vbs>
'
' AUTHOR: BurnSoft, BurnSoft
' DATE  : 3/16/2005
'
' COMMENT: <This script will check the Automatic Services and restart the
'			them if they are not up and running.>
'
'==========================================================================
Dim TimeToStop, Shell 
Dim strComputer,RunEvery
'==========================================================================
Function FindArgument(ByVal arg)
	
	dim retVal
	retVal = ""

	set ArgsObject = WScript.Arguments		
	
	for i = 0 to ArgsObject.Count - 1

		strVal = ArgsObject.Item(i)
		
		if (Left(strVal, 1) = "/") or (Left(strVal, 1) = "-") then
			position = InStr(1, strVal, "=")
			if position > 0 then				
				subString = mid(strVal, 2, position - 2)				
				if UCase(subString) = UCase(arg) then					
					retVal = mid(strVal, position+1, len(strVal))										
				end if
			end if
		end if
	next 
	
	FindArgument = retVal

End Function
'==========================================================================
Sub DisplayHelp
	WScript.Echo " "
    WScript.Echo "Copyright (c) 2004 BurnSoft."
	WScript.Echo " "	
	WScript.Echo "Auto_Service_Check.vbs"
	WScript.Echo "Check for Automatic services and restarts them"	
	WScript.Echo " "
	WScript.Echo "Usage for Remote Machine: "	
	WScript.Echo "cscript.exe Auto_Service_Check.vbs /Machine=SRV1 /interval=5"
	WScript.Echo " "
	WScript.Echo "Usage for Local Machine: "	
	WScript.Echo "cscript.exe Auto_Service_Check.vbs /interval=5"
	WScript.Echo " "
	WScript.Echo "Valid command-line arguments:"
	WScript.Echo " "
	WScript.Echo "  /Machine        - Name of the machine that you wish to monitor."
	WScript.Echo "  /interval		- number of minutes that you wish to check."
	WScript.Echo "NOTE: if you do not supply an interval it will constantly check "
	WScript.Echo "	the services which will consume 30 to 40% of your CPU."
End Sub
'==========================================================================
strComputer = FindArgument(uCase("machine"))
RunEvery = FindArgument(uCase("interval"))
WScript.echo("BurnSoft's Auto_Service_Check.vbs running.")
If RunEvery = "" Then
	RunEvery = 0
End If
WScript.echo("Will Check Auto Services Every " & RunEvery & " minutes.")
If strComputer = "" Then
	strComputer = "."
End If
If strComputer = "." Then
	WScript.echo("On Local Machine.")
Else
	WScript.echo("On Machine: " * strComputer)
End if
TimeToStop = False
Set Shell = CreateObject("Wscript.Shell")
Do Until TimeToStop
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colListOfServices = objWMIService.ExecQuery _
	 ("Select * from Win32_Service Where State = 'Stopped' and StartMode = " _
	     & "'Auto'")
	For Each objService in colListOfServices
		WScript.Echo(Now & vbtab & objService.DisplayName & " is down.")
	    objService.StartService()
	Next
	WScript.sleep 60000 * RunEvery
loop