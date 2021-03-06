'==========================================================================
'
' NAME: Postmdt.vbs
'
' AUTHOR: Brian Gonzalez
' DATE  : 4/17/2012
'
'
' Changlog:
' 4/17 - Fixed logic on Pro/Ent function check.
'	- Fixed Sample network portion of script.
'	- Fixed sylink call to configure SEP.
'	- Added subRoutine to re run install when 1618 is encountered.
'	- Added Resume On Next setting to top of script
'	- Added Delay before JoinDomain call, to ensure network is initialized
'	- Moved PCRS install log copy to the 4th login when vendorimaging is logged in
'	- Moved Startmenu clear and copy script to jobFinished subRoutine
' 4/25 - Added hopsdetect1_utilfiletxt Script script call.
'	- Added batch to initialize ecw client.
'	- Copied HopsDetect CMD to a subfolder.
' 4/27 - Added portion to write image version to c:\util\version.doc
'
' PURPOSE: Built to run on an image produced with MDT and started in the
'  field with connection to the VNSNY Domain.  The script will prepare a machine
'  for vnsny use.
'
'==========================================================================
On Error Resume Next
'Setup Objects, Constants and Variables.
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objWMI = GetObject("winmgmts:")
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Const ForReading = 1, ForWriting = 2, ForAppending = 8
versionNumber = "VNSY_Panasonic_1.1"
scriptFolder = objFSO.GetParentFolderName(Wscript.ScriptFullName)
currstepPath = scriptFolder & "\currStep.txt"


SELECT Description, DriverVersion FROM Win32_PnPSignedDriver WHERE DeviceName LIKE '%Sierra%