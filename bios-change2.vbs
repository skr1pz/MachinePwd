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

Dim strReturn
Select Case oReturn
Case 0 strReturn = "Success"
Case 1 strReturn = "Not Supported"
Case 2 strReturn = "Unspecified Error"
Case 3 strReturn = "Timeout"
Case 4 strReturn = "Failed" 
Case 5 strReturn = "Invalid Parameter"
Case 6 strReturn = "Access Denied"
Case Else strReturn = "..."
End Select
WScript.Echo "SetBiosSetting() returned: (" & oReturn _
& ") " & strReturn 