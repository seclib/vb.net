#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=tbicon.ico
#AutoIt3Wrapper_Outfile=Install.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Comment=Contact imaging@us.panasonic.com for support.
#AutoIt3Wrapper_Res_Description=OneClick Panasonic Toughbook Installer.
#AutoIt3Wrapper_Res_Fileversion=1.5.4.0
#AutoIt3Wrapper_Res_LegalCopyright=Panasonic Corporation Of North America
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
$sInstallVersion = "1.5.4"
$sLogFolderPath = @WindowsDir & "\Temp"
FileInstall("7za.exe", @WindowsDir & "\Temp\7za.exe", 1)
;** AUT2EXE settings
;================================================================================================================
; Panasonic Toughbook Parent Script
;  By Brian Gonzalez
;
; Purpose: Installs all drivers from the first discovered subfolder using
;	the PInstall.bat scripts.
;================================================================================================================
; Changelog
;================================================================================================================
; 1.5.4 - Dec 23, 2016
;	- Removed all log re-routing and HideCmdWindowEvery3Sec.exe calls.
; 1.5.3 - Apr 22, 2016
;	- Commented out code pertaining to KillBDD functionality, as it was causing errors.
; 1.5.2 - Apr 6, 2016
;	- Commented out portion of install that skips existing install.zips.
;	- Added logging for sBDDKill var.
; .......
; 1.2.3 - May 14, 2013
;	Set install.exe to perform an ArraySort on the $aDriverZips and not rely on the OS for correct sorting.
; 1.2.2 - Apr 17, 2013
;	Changed all C:\Windows strings to @WindowsDir
; 1.2.1 - Feb 4 2013
;	Changed script to dymanically search for Sub-Folder name and not rely on "src\" name.
; 1.2 - Jan 31 2013
;	Changed "##" o pre-fix used for skipped application/driver installs.
; 1.1 - Dec 21 2012
;	Added PNPCheck function, corrected count of elements in driver array.
;	Added TASKKILL call at end of script to close CmdWindowEvery3Sec.exe
;	Updated .ICOs for both HideCmdWindowEvery3Sec.exe and install.exe
;	Also compiled code with PanaConsulting icon.
;	Cleaned up Progress Bar Display by adding Bundle Name and removing File Extension
; 1.0 - Dec 12 2012
;	Added step to delete (pop) Optional Drivers from DriverZips Array.
;================================================================================================================
; AutoIt Includes
;================================================================================================================
#include <Date.au3>
#include <File.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
;================================================================================================================
; Main Routine
;================================================================================================================
AutoItSetOption("MustDeclareVars", 0)
Dim $aPNPIDContents[100] ;Array used when checking through the PNPID txt file
Local $StartTimer = TimerInit()

; Create LogFile, will reside in \Windows\Temp directory
$sLogFile = FileOpen($sLogFolderPath & "\PanaInstall_" & @YEAR & @MON & @MDAY & "_" & @HOUR & @MIN & ".log", 2)
If $sLogFile = -1 Then
	MsgBox(16, "logPath error", "Unable to access/create log file (" & $sLogFile & ").")
	Exit
EndIf

; Tag LogFile with Start Date and Time
FileWriteLine($sLogFile, "=========================================")
FileWriteLine($sLogFile, "Start Date/Time Stamp: " & _Now())

; Grab Sub-Folder path containing driver zip files
$cFoldersOnly = 2
$aSubFolders = _FileListToArray(@ScriptDir, "*", $cFoldersOnly)
$sSrcFolderName = $aSubFolders[1]
$sSrcPath = @ScriptDir & "\" & $aSubFolders[1]
;MsgBox(0, "", "SourcePath: " & $sSrcFolderName)

; Grab SystemDrive
$sSystemDrive = EnvGet("systemdrive"); Sets the System Drive for the Customer Computer

; Disable 64-Bit Redirection
DllCall("kernel32.dll", "int", "Wow64DisableWow64FsRedirection", "int", 1)
FileWriteLine($sLogFile, "Disabled 64Bit Redirection via DLL call.")

; Copy Changelogs local and pull bundle version number from "CurrentVersion=" in bundle_changelog.txt file
FileDelete($sSystemDrive & "\Drivers\*.txt")
FileCopy(@ScriptDir & "\*.txt", $sSystemDrive & "\Drivers\", 8)
If FileExists($sSystemDrive & "\Drivers\Bundle_Changelog.txt") Then
	Local $file = FileOpen($sSystemDrive & "\Drivers\Bundle_Changelog.txt", 0)
	; Read in lines of text until the EOF is reached
	While 1
		Local $line = FileReadLine($file)
		If @error = -1 Then ExitLoop
		;MsgBox(0, "Line read:", $line)
		If StringInStr($line, "CurrentVersion") Then
			$sBundleVersion = StringReplace(StringReplace($line, "CurrentVersion=", ""), " ", "")
			ExitLoop
		EndIf
	WEnd
Else
	$sBundleVersion = "Not Found"
EndIf

; Add more information to the log file
FileWriteLine($sLogFile, "Beginning to Process Bundle Name: " & $sSrcFolderName)
FileWriteLine($sLogFile, "Bundle Version: " & $sBundleVersion)
FileWriteLine($sLogFile, "Script Version: " & $sInstallVersion)
FileWriteLine($sLogFile, "=========================================")

; Populate array with all zip files inside of "src" sub-folder
$sFilesOnly = 1
If FileExists($sSrcPath) Then
	$aDriverZips = _FileListToArray($sSrcPath, "*.zip", $sFilesOnly)
	_ArraySort($aDriverZips, 0, 1)
	;_ArrayDisplay($aDriverZips)
EndIf

; Start Looping program to hide Command Line Windows
;Check if .ZIPs array was populated, if not exit
If @error = 1 Then
	MsgBox(0, "", "No Driver Folders found in \ sub-folder. Exiting script.")
	Exit
EndIf

; Delete Optional Installs (##_) from DriverZips Array
_ArrayDelete($aDriverZips, 0) ;Delete Orig Array Count from Array
$sDriverZips = _ArrayToString($aDriverZips, ",")
$sTestQuery = '\Q##_\E.*?,'
$sDriverZipsNoOptionals = StringRegExpReplace($sDriverZips, $sTestQuery, "")
$aDriverZips = StringSplit($sDriverZipsNoOptionals, ",")

; Copy 7za.exe locally to process ZIP packages
$s7ZAPath = @WindowsDir & "\Temp\7za.exe"
; First Make Sure 7za.exe Exists.
If Not (FileExists($s7ZAPath)) Then
	FileWriteLine($sLogFile, $s7ZAPath & " was not found, exiting script.")
	Exit
EndIf
FileCopy($s7ZAPath, @TempDir, 9) ; 9 = Overwrite and Create Dest Dir
FileWriteLine($sLogFile, "Copied 7za.exe to TempDir(" & @TempDir & "):" & @error)

; Begin to process ZIP files
FileWriteLine($sLogFile, "Beginning to process resource folders(src):" & $aDriverZips[0])

; Calculate amount of steps
$sTotalSteps = $aDriverZips[0] * 4
$sCurrentStep = 0
$sCurrentPercentComplete = fGrabPercentComplete(0)

; Create ProgressBar GUI
$oGUI = GUICreate("Panasonic Driver Installer (Install v" & $sInstallVersion & ") (Bundle Version v" & $sBundleVersion & ")", 600, 130, 0, 0, $WS_BORDER, $WS_EX_TOPMOST) ;width, height, top, left
$oCompleteProgressLabel = GUICtrlCreateLabel("Complete Process:", 5, 10, 100, 20) ;left, top, width, height
$oCompleteProgressBar = GUICtrlCreateProgress(100, 5, 490, 20, $PBS_SMOOTH)
$oCompletePercentLabel = GUICtrlCreateLabel("", 100, 30, 100, 20)
$oCompleteProgressText = GUICtrlCreateLabel("", 200, 30, 385, 20, $SS_RIGHT)
$oStepProgressLabel = GUICtrlCreateLabel("Current Step:", 5, 65, 100, 20)
$oStepProgressBar = GUICtrlCreateProgress(100, 60, 490, 20, $PBS_SMOOTH)
$oStepPercentLabel = GUICtrlCreateLabel("", 100, 85, 100, 20)
$oStepProgressText = GUICtrlCreateLabel("", 200, 85, 385, 20, $SS_RIGHT)
GUISetState()

fProgressBars(0, "Beginning Tbook Installer...", 0, "")

; Begin cycling through Driver ZIPs
For $i = 1 To $aDriverZips[0]
	$sDriverZipPath = $sSrcPath & "\" & $aDriverZips[$i]
	$sDriverName = StringLeft($aDriverZips[$i], StringLen($aDriverZips[$i]) - 4) ;Remove file extension when updating progress bars
	If StringLen($sDriverName) > 35 Then
		$sDriverName = StringLeft($sDriverName, 35) ;If name is longer than 8 characters, shorten the name.
	EndIf

	; fProgressBars args: complete percent, text, step percent, text
    $sCurrentStep = $sCurrentStep + 1
	$sCurrentPercentComplete = fGrabPercentComplete($sCurrentStep)
	fProgressBars($sCurrentPercentComplete, "Copying " & $i & " of " & $aDriverZips[0] & " packages in " & $sSrcFolderName & " bundle.", 0, "Copying " & $sDriverName)
	FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Beggining to copy driver zip """ & $sDriverZipPath & """ to (" & @TempDir & "\" & $sSrcFolderName & "\)")
	;MsgBox(0, "", FileGetLongName(@TempDir & "\" & $sSrcFolderName & "\"))
	$ret = FileCopy($sDriverZipPath, (@TempDir & "\" & $sSrcFolderName & "\"), 9)
	fProgressBars($sCurrentPercentComplete, "Copying " & $i & " of " & $aDriverZips[0] & " packages in " & $sSrcFolderName & " bundle.", 100, "Copying " & $sDriverName)
	; Delete ZipPath if the word "temp" is found anywhere in pathname.
	If StringInStr($sDriverZipPath, "temp") Then
		$ret = FileDelete($sDriverZipPath)
		FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Deleted driver zip """ & $sDriverZipPath & """:" & $ret)
	EndIf
	Sleep(750)
Next

; Begin cycling through Driver ZIPs
For $i = 1 To $aDriverZips[0]
	$sDriverZipPath = @TempDir & "\" & $sSrcFolderName & "\" & $aDriverZips[$i]
	$sDriverName = StringLeft($aDriverZips[$i], StringLen($aDriverZips[$i]) - 4) ;Remove file extension when updating progress bars

	;If FileExists($sSystemDrive & "\Drivers\" & $sSrcFolderName & "\" & $sDriverName) Then
	;	FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- " & $sSystemDrive & "\Drivers\" & $sSrcFolderName & "\" & $sDriverName & " already exist, jumping to next .ZIP")
	;	ContinueLoop
	;EndIf

	If StringLen($sDriverName) > 35 Then
		$sDriverName = StringLeft($sDriverName, 35) ;If name is longer than 8 characters, shorten the name.
	EndIf
	FileWriteLine($sLogFile, "Processing " & $i & " of " & $aDriverZips[0] & "... (" & @TempDir & "\" & $sSrcFolderName & "\" & $aDriverZips[$i] & ")")

	$sDriverExtractFolder = $sSystemDrive & "\Drivers\" & $sSrcFolderName & "\" & StringTrimRight($aDriverZips[$i], 4) ; Specify extract folder, driver name without extension
	$sCurrentStep = $sCurrentStep + 1
	$sCurrentPercentComplete = fGrabPercentComplete($sCurrentStep)
	fProgressBars($sCurrentPercentComplete, "Processing " & $i & " of " & $aDriverZips[0] & " packages in " & $sSrcFolderName & " bundle.", 33, "Extracting " & $sDriverName)
	FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Beggining to extract driver """ & $aDriverZips[$i] & """ from (" & $sDriverZipPath & ")")
	$sRet = RunWait("cmd.exe /c 7za.exe x """ & $sDriverZipPath & """ -o""" & $sSystemDrive & "\Drivers\" & $sSrcFolderName & "\*"" -y", @TempDir, @SW_HIDE)
	FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Completed extracting driver """ & $aDriverZips[$i] & """ to (" & $sDriverExtractFolder & "): " & $sRet)
	$ret = FileDelete($sDriverZipPath)
	FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Deleted driver zip """ & $sDriverZipPath & """:" & $ret)
	$sCurrentStep = $sCurrentStep + 1
	$sCurrentPercentComplete = fGrabPercentComplete($sCurrentStep)
	fProgressBars($sCurrentPercentComplete, "Processing " & $i & " of " & $aDriverZips[0] & " packages in " & $sSrcFolderName & " bundle.", 66, "Installing " & $sDriverName)
	If FileExists($sDriverExtractFolder & "\pnpid.txt") Then
		_FileReadToArray($sDriverExtractFolder & "\pnpid.txt", $aPNPIDContents)
		FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- PNPID file found, beginning WMI check for Device: " & $aPNPIDContents[1])
		If PNPCheck($aPNPIDContents[1]) Then
			FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- PNP Check returned successful for PNPID:" & $aPNPIDContents[1] & ": ")
			FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Beginning to execute the PInstall.bat from extracted driver as PNPID Check returned successfull (" & $sDriverExtractFolder & "): ")
			$sCmd = "cmd.exe /c pinstall.bat """" """ & $sLogFolderPath & """"
			FileWriteLine($sLogFile, "Executing: """ & $sCmd & """" & ", from the following Directory: " & $sDriverExtractFolder)
			$sRet = RunWait($sCmd, $sDriverExtractFolder, @SW_HIDE)
		Else
			FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Skipping install as PNPID Check returned Failed (" & $sDriverExtractFolder & "): ")
		EndIf
	Else
		FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Beginning to execute the PInstall.bat from extracted driver (" & $sDriverExtractFolder & "): ")
		$sCmd = "cmd.exe /c pinstall.bat """" """ & $sLogFolderPath & """"
		FileWriteLine($sLogFile, "Executing: """ & $sCmd & """" & ", from the following Directory: " & $sDriverExtractFolder)
		; Execution of PInsall
		$sRet = RunWait($sCmd, $sDriverExtractFolder, @SW_HIDE)
	EndIf
	$sCurrentStep = $sCurrentStep + 1
	$sCurrentPercentComplete = fGrabPercentComplete($sCurrentStep)
	fProgressBars($sCurrentPercentComplete, "Processing " & $i & " of " & $aDriverZips[0] & " packages in " & $sSrcFolderName & " bundle.", 100, "Install of " & $sDriverName & " is complete.")
	Sleep(750) ; Delay to let user see install completed.
	FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Completed executing the PInstall.bat from extracted driver (" & $sDriverExtractFolder & "): " & $sRet)

Next

;Delete any .log files created in the root of the SystemDrive
$AmountOfSecondsRun = TimerDiff($StartTimer/1000)
$aLogFilesOnSysDrive = _FileListToArray($sSystemDrive, "*.log", $sFilesOnly)

;Check if .ZIPs array was populated, if not exit
If NOT @error = 1 Then
   For $i = 1 To $aLogFilesOnSysDrive[0]
	   $TempLogFilePath = $sSystemDrive & "\" & $aLogFilesOnSysDrive[$i]
	   $t = FileGetTime($TempLogFilePath, $FT_CREATED, 0)
	   ;_ArrayDisplay($t)
	   Local $Date = $t[0] & '/' & $t[1] & '/' & $t[2] & ' ' & $t[3] & ':' & $t[4] & ':' & $t[5]
	   If (_DateDiff('s', $Date, _NowCalc()) <= $AmountOfSecondsRun) Then
		   FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Deleting logfile (" & $TempLogFilePath & "): " & $sRet)
		   FileDelete($TempLogFilePath)
	   EndIf
   Next
EndIf

FileWriteLine($sLogFile, @HOUR & ":" & @MIN & "--- Script Execution is complete.")

; Delete 7za.exe
FileDelete($s7ZAPath)
FileClose($sLogFile)
Exit
;================================================================================================================
; Functions and Sub Routines
;================================================================================================================
Func fGrabPercentComplete($sCurrentStep)
	Return Round(($sCurrentStep / $sTotalSteps) * 100)
EndFunc   ;==>fGrabPercentComplete

Func fProgressBars($sCompleteBarPerc, $sCompleteBarText, $sStepBarPerc, $sStepBarText) ;complete percent, text, step percent, text
	GUICtrlSetData($oCompleteProgressBar, $sCompleteBarPerc)
	GUICtrlSetData($oCompleteProgressText, $sCompleteBarText)
	GUICtrlSetData($oCompletePercentLabel, $sCompleteBarPerc & "%")
	GUICtrlSetData($oStepProgressBar, $sStepBarPerc)
	GUICtrlSetData($oStepProgressText, $sStepBarText)
	GUICtrlSetData($oStepPercentLabel, $sStepBarPerc & "%")
EndFunc   ;==>fProgressBars

Func PNPCheck($sPNPID)
	$sPNPID = StringReplace($sPNPID, '\', '%') ;Clean up PNPID string
	$objWMIService = ObjGet("winmgmts:\\.\root\cimv2")
	$sQuery = "Select * FROM Win32_PNPEntity WHERE DeviceID LIKE '%" & $sPNPID & "%'"
	$colPNPEntity = $objWMIService.ExecQuery($sQuery)
	Return $colPNPEntity.Count
EndFunc   ;==>PNPCheck