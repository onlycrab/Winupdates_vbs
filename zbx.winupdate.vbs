On Error Resume Next
'----------ZABBIX PATH-----------
Dim senderPath, confPath, FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
senderPath = ""
confPath = ""
If WScript.Arguments.Count >= 2 Then
	senderPath = WScript.Arguments(0)
	confPath = WScript.Arguments(1)
	senderPath = Replace(senderFileName, chr(34), "")
	confPath = Replace(confFileName, chr(34), "")
	If Not FSO.FileExists(senderPath) Then
		senderPath = ""
	End If
	If Not FSO.FileExists(confPath) Then
		confPath = ""
	End If
End If
If Len(senderPath) < 2 Or Len(confPath) < 2 Then
	'search sender and configurations files
	'The search is performed recursively in each parent folder starting from the current.
	'F.e., if the full path of the script file is "c:\Program Files\Zabbix Agent\plugins\zbx.winupdate.vbs", then:
	'- search for files in the "c:\Program Files\Zabbix Agent\plugins\" directory
	'- if no files are found, check in the "c:\Program Files\Zabbix Agent\" directory
	'- if no files are found, check in the "c:\Program Files\" directory
	'- if no files are found, check in the "c:\" directory

	Dim senderFileName, confFileName, isSenderFound, isConfFound, parentPath
	senderFileName = "zabbix_sender.exe"
	confFileName = "zabbix_agentd.win.conf"
	isSenderFound = False
	isConfFound = False
	parentPath = FSO.GetParentFolderName(WScript.ScriptFullName)
	Do While Not isSenderFound Or Not isConfFound
		If Len(Trim(parentPath)) = 0 Then
			WScript.Quit(1)
		End If
		If Not isSenderFound Then
			If FSO.FileExists(Replace(parentPath & "\" & senderFileName, "\\", "\")) Then
				senderPath = Replace(parentPath & "\" & senderFileName, "\\", "\")
				isSenderFound = True
			End If
		End If
		If Not isConfFound Then
			If FSO.FileExists(Replace(parentPath & "\" & confFileName, "\\", "\")) Then
				confPath = Replace(parentPath & "\" & confFileName, "\\", "\")
				isConfFound = True
			End If
		End If
		parentPath = FSO.GetParentFolderName(parentPath)		
	Loop
End If
Set FSO = Nothing
'----------CHECK CONFIG FILE-----
Dim WshShell, zhost, Config, ZServerAct, useConfigFile
Set WshShell = WScript.CreateObject("WScript.Shell")

'object for ZabbixServer-ZabbixPort pairs storage
'used if the field 'Hostname' is empty in configuration file
Set ZServerAct = Null
useConfigFile = True
'read configuration file like .ini file
Set Config = new SimplyIni
Config.AddKeyValueFile(chr(34) & confPath & chr(34))
zhost = Config.GetValueByKey("hostname")
If IsNull(zhost) Or Len(zhost) = 0 Then
	useConfigFile = False
	'get hostname from OS
	zhost = ""
	zhost = WshShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	If Len(zhost) = 0 Then
		WScript.Quit(1)
	End If
	
	'parsing server and port
	'reading 'ServerActive'
	tmp = Config.GetValueByKey("serveractive")
	If IsNull(tmp) Or Len(tmp) = 0 Then
		'if 'ServerActive' is empty or missing - reading 'Server'
		tmp = Config.GetValueByKey("server")
		If IsNull(tmp) Or Len(tmp) = 0 Then
			WScript.Quit(1)
		End If
	End If
	
	'create ZServerActive object
	Set ZServerAct = New ZServerActive
	Dim tmpArr
	tmpArr = Split(tmp, ",")
	'for each value parse adding server and port to ZServerActive object
	For i = 0 To UBound(tmpArr)		
		ZServerAct.AddElementStr(tmpArr(i))
	Next
	If ZServerAct.Count = 0 Then
		WScript.Quit(1)
	End If
End If
'----------SEARCHER DEFINITION---
Dim updateSession, updateSearcher, searchResult, AllUpdates, CriticalUpdates, DefinitionUpdates, SecurityUpdates, ServicePacks, UpdateRollUps, RebootRequired, RebootRequiredForNewUpdates
Set updateSession = createObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()
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
Elseif AllUpdates <> -1 Then 
	CriticalUpdates = 0
End If
'----------SEARCH SECURITY-------
SecurityUpdates = -1
If updatesDetectedByTypes <> AllUpdates Then
	SecurityUpdates = CheckUpdatesQuantity("Software", "0FA1201D-4330-4FA8-8AE9-B877473B6441") 'SecurityUpdates
	updatesDetectedByTypes = updatesDetectedByTypes + SecurityUpdates
Elseif AllUpdates <> -1 Then 
	SecurityUpdates = 0
End If
'----------SEARCH DEFINITION-----
DefinitionUpdates = -1
If updatesDetectedByTypes <> AllUpdates Then
	DefinitionUpdates = CheckUpdatesQuantity("Software", "E0789628-CE08-4437-BE74-2495B842F43B") 'DefinitionUpdates
	updatesDetectedByTypes = updatesDetectedByTypes + DefinitionUpdates
Elseif AllUpdates <> -1 Then 
	DefinitionUpdates = 0
End If	
'----------SEARCH SERVICEPACKS---
ServicePacks = -1
If updatesDetectedByTypes <> AllUpdates Then
	ServicePacks = CheckUpdatesQuantity("Software", "68C5B0A3-D1A6-4553-AE49-01D3A7827828") 'ServicePacks
	updatesDetectedByTypes = updatesDetectedByTypes + ServicePacks
Elseif AllUpdates <> -1 Then 
	ServicePacks = 0
End If	
'----------SEARCH ROOLUPS--------
UpdateRollUps = -1
If updatesDetectedByTypes <> AllUpdates Then
	UpdateRollUps = CheckUpdatesQuantity("Software", "28BC880E-0592-4CBF-8F95-C79B17911D5F") 'UpdateRollUps
	updatesDetectedByTypes = updatesDetectedByTypes + UpdateRollUps
Elseif AllUpdates <> -1 Then 
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
'----------SEND DATA-------------
Dim sendDataArray(8), isWsusUnavailable
If AllUpdates = -1 Or CriticalUpdates = -1 Or DefinitionUpdates = -1 Or SecurityUpdates = -1 Or ServicePacks = -1 Or UpdateRollUps = - 1 Then	
	sendDataArray(0) = "zbx.winupdate.vbs.wsusavailability -o 0"
	isWsusUnavailable = True
Else
	sendDataArray(0) = "zbx.winupdate.vbs.wsusavailability -o 1"
	isWsusUnavailable = False
End If
sendDataArray(1) = "zbx.winupdate.vbs.all -o " & CStr(AllUpdates)
sendDataArray(2) = "zbx.winupdate.vbs.critical -o " & CStr(CriticalUpdates)
sendDataArray(3) = "zbx.winupdate.vbs.definition -o " & CStr(DefinitionUpdates)
sendDataArray(4) = "zbx.winupdate.vbs.security -o " & CStr(SecurityUpdates)
sendDataArray(5) = "zbx.winupdate.vbs.servicepacks -o " & CStr(ServicePacks)
sendDataArray(6) = "zbx.winupdate.vbs.updaterollups -o " & CStr(UpdateRollUps)
sendDataArray(7) = "zbx.winupdate.vbs.datetime -o " & chr(34) & CStr(Date) & " " & CStr(Time) & chr(34)
sendDataArray(8) = "zbx.winupdate.vbs.rebootrequired -o " & CStr(RebootRequired)

Dim strRun
'if filed Hostname is not empty in configuration file
If useConfigFile Then
	'send zbx.winupdate.vbs.wsusavailability and Quit if wsus is unavailable
	strRun = chr(34) & senderPath & chr(34) & " -c " & chr(34) & confPath & chr(34) & " -k " & sendDataArray(0)
	WshShell.Run strRun,1,True
	If isWsusUnavailable Then
		WScript.Quit(0)
	End If
	
	'for each value
	For i = 1 To UBound(sendDataArray)
		strRun = chr(34) & senderPath & chr(34) & " -c " & chr(34) & confPath & chr(34) & " -k " & sendDataArray(i)
		WshShell.Run strRun,1,True
	Next
'if filed Hostname is empty in configuration file
Else
	'send zbx.winupdate.vbs.wsusavailability
	'for each zabbix-server in configuration file
	For i = 0 To ZServerAct.Count - 1
		If IsNull(ZServerAct.GetZPortByIndex(i)) Then
			'zabbix_sender -z server -s host -k key -v value
			strRun = chr(34) & senderPath & chr(34) & " -z " & ZServerAct.GetZServerByIndex(i) & " -s " & zhost & " -k " & sendDataArray(0)
		Else
			'zabbix_sender -z server -p port -s host -k key -v value
			strRun = chr(34) & senderPath & chr(34) & " -z " & ZServerAct.GetZServerByIndex(i) & " -p " & ZServerAct.GetZPortByIndex(i) & " -s " & zhost & " -k " & sendDataArray(0)
		End If
		WshShell.Run strRun,1,True
	Next
	
	'Quit if wsus is unavailable
	If isWsusUnavailable Then
		WScript.Quit(0)
	End If
	
	'for each zabbix-server in configuration file
	For i = 0 To ZServerAct.Count - 1
		'for each value
		For j = 1 To UBound(sendDataArray)
			If IsNull(ZServerAct.GetZPortByIndex(i)) Then
				'zabbix_sender -z server -s host -k key -v value
				strRun = chr(34) & senderPath & chr(34) & " -z " & ZServerAct.GetZServerByIndex(i) & " -s " & zhost & " -k " & sendDataArray(j)
			Else
				'zabbix_sender -z server -p port -s host -k key -v value
				strRun = chr(34) & senderPath & chr(34) & " -z " & ZServerAct.GetZServerByIndex(i) & " -p " & ZServerAct.GetZPortByIndex(i) & " -s " & zhost & " -k " & sendDataArray(j)
			End If
			WshShell.Run strRun,1,True
		Next
	Next
End If

Set WshShell = Nothing
Set updateSession = Nothing
Set updateSearcher = Nothing
Set searchResult = Nothing

WScript.Quit(0)

Function CheckUpdatesQuantity(updateType, updateCategoryID)
	Set searchResult = updateSearcher.Search("Type='" & updateType & "' AND IsInstalled=0 AND CategoryIDs contains '" & updateCategoryID & "'")
	CheckUpdatesQuantity = searchResult.Updates.count
End Function

'Class for ZabbixServer-ZabbixPort pairs storage
'Used if the field 'Hostname' is empty in configuration file
Class ZServerActive
	Private zserverArray()
	Private zportArray()
	
	Private Sub Class_Initialize()
		Redim Preserve zserverArray(0)
		Redim Preserve zportArray(0)
		zserverArray(0) = Null
		zportArray(0) = Null
	End Sub
	
	Public Function Count()
		If IsNull(zserverArray(0)) And IsNull(zportArray(0)) Then
			Count = 0
		End If
		Count = UBound(zserverArray) + 1
	End Function
	
	Public Function AddElement(zserver, zport)
		If Len(Trim(zserver)) <> 0 Then
			If IsNull(zserverArray(0)) And IsNull(zportArray(0)) Then
				zserverArray(0) = Trim(zserver)
				If Not IsNull(zport) And Len(Trim(zport)) <> 0 Then
					zportArray(0) = Trim(zport)
				Else
					zportArray(0) = Null
				End If			
			Else
				Redim Preserve zserverArray(UBound(zserverArray) + 1)
				Redim Preserve zportArray(UBound(zportArray) + 1)
				zserverArray(UBound(zserverArray)) = Trim(zserver)			
				If Not IsNull(zport) And Len(Trim(zport)) <> 0 Then
					zportArray(UBound(zportArray)) = Trim(zport)
				Else
					zportArray(UBound(zportArray)) = Null
				End If			
			End If
		End If
	End Function
	
	'parse element from string "zabbix-server:zabbix-port"
	Public Function AddElementStr(zserverportStr)
		tmpzsp = Split(zserverportStr, ":")
		If UBound(tmpzsp) = 1 Then
			AddElement tmpzsp(0), tmpzsp(1)
		ElseIf UBound(tmpzsp) = 0 Then
			AddElement tmpzsp(0), Null
		End If
	End Function
	
	Public Function GetZServerByIndex(index)
		GetZServerByIndex = Null
		If index <= UBound(zserverArray) Then
			GetZServerByIndex = zserverArray(index)
		End If
	End Function
	
	Public Function GetZPortByIndex(index)
		GetZPortByIndex = Null
		If index <= UBound(zportArray) Then
			GetZPortByIndex = zportArray(index)
		End If
	End Function
	
End Class

'Class for configuration file data storage
Class SimplyIni
	Public name
	Private keyArray()
	Private valueArray()

	Private Sub Class_Initialize()
		Redim Preserve keyArray(0)
		Redim Preserve valueArray(0)
		keyArray(0) = Null
		valueArray(0) = Null
	End Sub

	'Full key match value search
	Public Function GetValueByKey(keyName)
		keyName = LCase(keyName)
		GetValueByKey = Null
		For i = 0 To UBound(keyArray)
			If keyArray(i) = keyName Then
				GetValueByKey = valueArray(i)
				Exit For
			End If
		Next
	End Function

	Public Function GetValueByIndex(index)
		GetValueByIndex = Null
		If index <= UBound(valueArray) Then
			GetValueByIndex = valueArray(index)
		End If
	End Function

	Public Function GetKeyByIndex(index)
		GetKeyByIndex = Null
		If index <= UBound(keyArray) Then
			GetKeyByIndex = keyArray(index)
		End If
	End Function

	Public Function Count()
		Count = UBound(keyArray) + 1
	End Function
	
	Public Function AddElement(keyName, value)
		If IsNull(keyArray(0)) And IsNull(valueArray(0)) Then
			keyArray(0) = LCase(Trim(keyName))
			valueArray(0) = Trim(value)
		Else
			Redim Preserve keyArray(UBound(keyArray) + 1)
			Redim Preserve valueArray(UBound(valueArray) + 1)
			keyArray(UBound(keyArray)) = LCase(Trim(keyName))
			valueArray(UBound(valueArray)) = Trim(value)
		End If
	End Function

	'Adds all valid elements from file
	Public Function AddKeyValueFile(filePath)
		Set SectionFSO = CreateObject("Scripting.FileSystemObject")
		Set sectionFile = SectionFSO.OpenTextFile(Replace(filePath, chr(34), ""), 1)
		Do While Not sectionFile.AtEndOfStream
			line = sectionFile.ReadLine
			
			'ignore comments
			If (Mid(line, 1, 1)) <> ";" And (Mid(line, 1, 1)) <> "#" Then
				splited = Split(line, "=")
				If UBound(splited) <> 0 Then
					AddElement Trim(splited(0)), Trim(splited(1))
				End If
			End If
		Loop
		Set SectionFSO = Nothing
		Set sectionFile = Nothing
	End Function
End Class
