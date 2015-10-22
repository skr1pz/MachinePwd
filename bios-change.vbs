Const wbemFlagReturnImmediately = 16
Const wbemFlagForwardOnly = 32
lFlags = wbemFlagReturnImmediately + wbemFlagForwardOnly
strService = "winmgmts:{impersonationlevel=impersonate}//"
strComputer = "."
strNamespace = "/root/HP/InstrumentedBIOS"
strQuery = "select * from HP_BIOSSettingInterface"
Set objWMIService = GetObject(strService & _
strComputer & strNamespace)
Set colItems = objWMIService.ExecQuery(strQuery,,lFlags)
For each objItem in colItems
objItem.SetBiosSetting oReturn, _
"LAN/WLAN Switching", _
"Enable", _
"<kbd/>"
Next

