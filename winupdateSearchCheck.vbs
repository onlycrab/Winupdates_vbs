On Error Resume Next
Dim senderPath, confPath, FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
senderPath = ""
confPath = ""
'search sender and configurations files
'The search is performed recursively in each parent folder starting from the current.
'F.e., if the full path of the script file is "c:\Program Files\Zabbix Agent\plugins\zbx.winupdate.vbs", then:
'- search for files in the "c:\Program Files\Zabbix Agent\plugins\" directory
'- if no files are found, check in the "c:\Program Files\Zabbix Agent\" directory
'- if no files are found, check in the "c:\Program Files\" directory
'- if no files are found, check in the "c:\" directory
Dim senderFileName, confFileName, isSenderFound, isConfFound, parentPath
senderFileName = "zabbix_sender.exe"
confFileName = "zabbix_agentd.conf"
isSenderFound = False
isConfFound = False
parentPath = FSO.GetParentFolderName(WScript.ScriptFullName)
Do While Not isSenderFound Or Not isConfFound
	If Len(Trim(parentPath)) = 0 Then
		MsgBox TestResult(senderFileName, confFileName, senderPath, confPath, isSenderFound, isConfFound)
		WScript.Quit(0)
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
MsgBox TestResult(senderFileName, confFileName, senderPath, confPath, isSenderFound, isConfFound)
WScript.Quit(0)

Function TestResult(senderFileName, confFileName, senderPath, confPath, isSenderFound, isConfFound)
	Dim testMessage
	testMessage = ""
	If isSenderFound Then
		testMessage = testMessage & "Sender file " & chr(34) & senderFileName & chr(34) & " found : " & vbCrlf & senderPath & vbCrlf
	Else
		testMessage = testMessage & "Sender file " & chr(34) & senderFileName & chr(34) & " NOT FOUND!"
	End If
	testMessage = testMessage & vbCrlf
	If isConfFound Then
		testMessage = testMessage & "Configuration file " & chr(34) & confFileName & chr(34) & " found : " & vbCrlf & confPath
	Else
		testMessage = testMessage & "Configuration file " & chr(34) & confFileName & chr(34) & " NOT FOUND!"
	End If
	TestResult = testMessage
End Function
