# Winupdates_vbs
## Description
  This script was been created to monitor Windows servers for available updates. The script uses only the Windows Update API. (https://docs.microsoft.com/en-us/windows/desktop/api/wuapi/nf-wuapi-iupdatesearcher-search). 
  The script checks the number of all available updates, and only the important types of updates separately (according to https://blogs.technet.microsoft.com/dubaisec/2016/01/28/windows-update-categories/).
  Zabbix template with the necessary metrics is attached to the script.
## Compatibility
  Tested on Windows Server 2008, 2012, 2016; zabbix 5.0.
## How to use
  The script can be used in two ways:
  1. Specify macroses {$ZSENDER} (full path to the zabbix_sender.exe) and {$ZPATH} (full path to the zabbix_agentd.win.conf) in the zabbix-node.
  2. Do not specify macros. Then it will automatically search for files zabbix_sender.exe and zabbix_agentd.win.conf.
  Search algorithm:
  The search is performed recursively in each parent folder starting from the script parent folder.
  F.e., if the full path of the script file is "c:\Program Files\Zabbix Agent\plugins\zbx.winupdate.vbs", then:
    search for files in the "c:\Program Files\Zabbix Agent\plugins\" directory;
    if no files are found, check in the "c:\Program Files\Zabbix Agent\" directory;
    if no files are found, check in the "c:\Program Files\" directory;
    if no files are found, check in the "c:\" directory.
### Case 1
  1. Distribute the "zbx.winupdate.vbs" file to the required machines.
  2. Make sure that "EnableRemoteCommand=1" is enabled in "zabbix_agentd.win.conf".
  3. Import the "Winupdates_vbs.xml" template.
  4. Specify the paths to "zabbix_sender.exe" and "zabbix_agentd.win.conf" by the macroses {$ZSENDER} and {$ZPATH} respectively on each zabbix-node.
### Case 2
  1. Distribute the "zbx.winupdate.vbs" and "winupdateSearchCheck.vbs" file to the required machines.
  2. Run the script "winupdateSearchCheck.vbs" manually: make sure the automatic files search is successful.
  3. Make sure that "EnableRemoteCommand=1" is enabled in "zabbix_agentd.win.conf".
  4. Import the "Winupdates_vbs.xml" template.
### Optional
  All items are divided into 2 groups: "Winupdates vbs" and "Winupdates vbs panel". You can create a panel with a convenient display of the current status of available updates on the servers. 
  To do this, you need to create a widget "Data overview" on the dashboard, in the application field specify "Winupdates vbs panel".
  See panel_example.jpg.
