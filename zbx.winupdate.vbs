On Error Resume Next
'----------ZABBIX PATH-----------
If WScript.Arguments.Count < 2 Then 'the path may contain spaces
	senderPath = "c:\zabbix\zabbix_sender.exe"
	confPath = "c:\zabbix\zabbix_agentd.win.conf"
Elseif WScript.Arguments.Count = 2 Then
	senderPath = WScript.Arguments(0)
	confPath = WScript.Arguments(1)
Else
	WScript.Quit(1)
End If
'----------DEFINITION------------
Dim WshShell, updateSession, updateSearcher, searchResult, commandPart, AllUpdates, CriticalUpdates, DefinitionUpdates, SecurityUpdates, ServicePacks, UpdateRollUps, RebootRequired, RebootRequiredForNewUpdates
Set WshShell = WScript.CreateObject("WScript.Shell")
Set updateSession = createObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()
commandPart = chr(34) & senderPath & chr(34) & " -c " & chr(34) & confPath & chr(34) & " -k "
'----------SEARCH ALL------------
Set searchResult = updateSearcher.Search("Type='Software' AND IsInstalled=0") 'AllUpdates
AllUpdates = -1
AllUpdates = searchResult.Updates.count
updatesDetectedByTypes = 0
'----------SEARCH CRITICAL-------
CriticalUpdates = -1
If updatesDetectedByTypes <> AllUpdates Then
	CriticalUpdates = CheckUpdatesQuantity("Software", "E6CF1350-C01B-414D-A61F-263D14D133B4") 'CriticalUpdates	
	updatesDetectedByTypes = updatesDetectedByTypes + CriticalUpdates
Else
	CriticalUpdates = 0
End If
'----------SEARCH SECURITY-------
SecurityUpdates = -1
If updatesDetectedByTypes <> AllUpdates Then
	SecurityUpdates = CheckUpdatesQuantity("Software", "0FA1201D-4330-4FA8-8AE9-B877473B6441") 'SecurityUpdates
	updatesDetectedByTypes = updatesDetectedByTypes + SecurityUpdates
Else
	SecurityUpdates = 0
End If
'----------SEARCH DEFINITION-----
DefinitionUpdates = -1
If updatesDetectedByTypes <> AllUpdates Then
	DefinitionUpdates = CheckUpdatesQuantity("Software", "E0789628-CE08-4437-BE74-2495B842F43B") 'DefinitionUpdates
	updatesDetectedByTypes = updatesDetectedByTypes + DefinitionUpdates
Else
	DefinitionUpdates = 0
End If	
'----------SEARCH SERVICEPACKS---
ServicePacks = -1
If updatesDetectedByTypes <> AllUpdates Then
	ServicePacks = CheckUpdatesQuantity("Software", "68C5B0A3-D1A6-4553-AE49-01D3A7827828") 'ServicePacks
	updatesDetectedByTypes = updatesDetectedByTypes + ServicePacks
Else
	ServicePacks = 0
End If	
'----------SEARCH ROOLUPS--------
UpdateRollUps = -1
If updatesDetectedByTypes <> AllUpdates Then
	UpdateRollUps = CheckUpdatesQuantity("Software", "28BC880E-0592-4CBF-8F95-C79B17911D5F") 'UpdateRollUps
	updatesDetectedByTypes = updatesDetectedByTypes + UpdateRollUps
Else
	UpdateRollUps = 0
End If
'----------REBOOT REQUIRED-------
RebootRequiredFlag = 0
errText = Err.Description
RebootRequired = ""
RebootRequired = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired\")
If Err.Description = errText Then RebootRequiredFlag = RebootRequiredFlag + 1

errText = Err.Description
RebootRequired = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending\")
If Err.Description = errText Then RebootRequiredFlag = RebootRequiredFlag + 1

If RebootRequiredFlag > 0 Then
	RebootRequired = "1"
Else
	RebootRequired = "0"
End If
'----------SENT DATA-------------
If AllUpdates = -1 Or CriticalUpdates = -1 Or DefinitionUpdates = -1 Or SecurityUpdates = -1 Or ServicePacks = -1 Or UpdateRollUps = - 1 Then
	strRun = commandPart & "zbx.winupdate.vbs.wsusavailability -o 0"	
Else
	strRun = commandPart & "zbx.winupdate.vbs.wsusavailability -o 1"
End If
WshShell.Run strRun,1,True

strRun = commandPart & "zbx.winupdate.vbs.all -o " & CStr(AllUpdates)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.critical -o " & CStr(CriticalUpdates)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.definition -o " & CStr(DefinitionUpdates)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.security -o " & CStr(SecurityUpdates)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.servicepacks -o " & CStr(ServicePacks)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.updaterollups -o " & CStr(UpdateRollUps)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.datetime -o " & chr(34) & CStr(Date) & " " & CStr(Time) & chr(34)
WshShell.Run strRun,1,True
strRun = commandPart & "zbx.winupdate.vbs.rebootrequired -o " & CStr(RebootRequired)
WshShell.Run strRun,1,True

Set WshShell = Nothing
Set updateSession = Nothing
Set updateSearcher = Nothing
Set searchResult = Nothing

If Len(Err.Description) > 1 Then
	WScript.Quit(1)
Else
	WScript.Quit(0)
End If

Function CheckUpdatesQuantity(updateType, updateCategoryID)
	Set searchResult = updateSearcher.Search("Type='" & updateType & "' AND IsInstalled=0 AND CategoryIDs contains '" & updateCategoryID & "'")
	CheckUpdatesQuantity = searchResult.Updates.count
	searchResult = Nothing
End Function
