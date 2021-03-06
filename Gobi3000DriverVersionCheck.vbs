'==========================================================================
' NAME: Gobi3000DriverVersionCheck.vbs
'
' AUTHOR: Brian Gonzalez , Panasonic
' DATE  : 05/30/2013
'
' COMMENT:
' Checks on the driver version of the Gobi 3000 Module.
'==========================================================================
On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PnPSignedDriver WHERE DeviceName LIKE '%Sierra%'", "WQL", _
	bemFlagReturnImmediately + wbemFlagForwardOnly)

