# Winupdates_vbs
## Description
  This script was been created to monitor Windows servers for available updates. The script uses only the Windows Update API. (https://docs.microsoft.com/en-us/windows/desktop/api/wuapi/nf-wuapi-iupdatesearcher-search). 
  The script checks the number of all available updates, and only the important types of updates separately (according to https://blogs.technet.microsoft.com/dubaisec/2016/01/28/windows-update-categories/).
  Zabbix template with the necessary metrics is attached to the script.
## Compatibility
  Tested on Windows Server 2008, 2012, 2016; zabbix 3.4.
## How to use
  The script can be called in two ways:
  1. With parameters: 1 – path to zabbix_sender.exe, 2 – path to "zabbix_agentd.win.conf".
  2. No parameters: paths to "zabbix_sender.exe" and "zabbix_agentd.win.conf" must be specified manually in the variables "senderPath", "confPath" (line 4, 5).
  In both cases, the paths to the specified files may contain spaces.
  Case 1 - item "WU - WinupdatesCheckLast 0", case 2 - "WU - WinupdatesCheckLast 1".
### Case 1
  1. Distribute the zbx.winupdate.vbs file to the required machines.
  2. Make sure that "EnableRemoteCommand=1" is enabled in "zabbix_agent.win.conf".
  3. Import the "Winupdates_vbs.xml" template.
  4. Verify that item "WU - WinupdatesCheckLast 0" is disabled.
  5. Specify the paths to "zabbix_sender.exe" and "zabbix_agentd.win.conf" by the parameters at item "WU - WinupdatesCheckLast 1".
####  Example: system.run[c:\zabbix\plugins\zbx.winupdate.vbs "c:\zabbix\zabbix_sender.exe" "c:\zabbix\zabbix_agentd.win.conf",nowait]
  6. Verify that item "WU - WinupdatesCheckLast 1" is enabled.
### Case 2
  1. In the file zbx.winupdate.vbs specify the paths to "zabbix_sender.exe" and "zabbix_agentd.win.conf" in the variables "senderPath", "confPath" (lines 4, 5).
  2. Distribute the zbx.winupdate.vbs file to the required machines.
  3. Make sure that "EnableRemoteCommand=1" is enabled in "zabbix_agent.win.conf".
  4. Import the Winupdates_vbs.xml template.
  5. Verify that item "WU - WinupdatesCheckLast 1" is disabled.
  6. Verify that item "WU - WinupdatesCheckLast 0" is enabled.
### Optional
  All items are divided into 2 groups: "Winupdates vbs" and "Winupdates vbs panel". You can create a panel with a convenient display of the current status of available updates on the servers. 
  To do this, you need to create a widget "Data overview" on the dashboard, in the application field specify "Winupdates vbs panel".
  See panel_example.jpg.
