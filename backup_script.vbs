runBackup

Function runBackup()
	Dim fileName, folderToBackup, fileToBackupTo, currentDate, fileExtensions, eachFileWrite, totalFileWrite

	currentDate = DAY(date()) & "." & Month(date()) & "." & YEAR(Date())
	fileName = "uTorrent " & currentDate & ".zip"		'File name; leave everything from the first "&" onwards
	folderToBackup = "C:\Users\mlevit\AppData\Roaming\uTorrent"		' Location you wish to backup
	fileExtensions = array("dat", "old")		'Specify which extensions you wish to backup on (i.e. leave as "array()" if you want the entire folder)

	totalFileWrite = 2500

	Dim fileSystem, fileObject, getFile

	set fileSystem = CreateObject("Scripting.FileSystemObject")
	
	If Not fileSystem.FolderExists(folderToBackup) Then
		gotoError(1000)
	End If
	
	fileSystem.CreateTextFile(fileName)
	
	set fileToBackupTo = fileSystem.GetFile(fileName)
	
	If Not fileSystem.FileExists(fileName) Then
		gotoError(1001)
	End if

	Dim shellApp, zipFile, folderToZip

	set shellApp = CreateObject("Shell.Application")
	set zipFile = shellApp.NameSpace(fileSystem.GetAbsolutePathName(".") & "\" & fileName)
	set folderToZip = shellApp.NameSpace(folderToBackup)
	
	'If file extensions were specified, then run chekc
	if (ubound(fileExtensions) > 0) Then
		'Loops through each file in the folder
		For Each file In folderToZip.Items
			'Loops through all the file extensions specified above
			For Each extension In fileExtensions
				'Looks for a matching file extension within each file name
				if (checkFileExtension(fileSystem.GetExtensionName(file), extension)) Then
					newZipFileSize = 0
					
					zipFile.CopyHere(file)
					
					While (newZipFileSize < fileToBackupTo.Size)
						newZipFileSize = fileToBackupTo.Size
						WScript.Sleep 10000
					Wend
					
					Exit For
				End If
			Next
		Next	
	Else
		zipFile.CopyHere(folderToZip.Items)
	End If

	WScript.Sleep totalFileWrite

	MsgBox "Backup has finished." & chr(13) & "A new file was created: " & fileName, 48, "Backup Completed"
	WScript.Quit
End Function

Function checkFileExtension(fileExtension, extension)
	If (StrComp(fileExtension,extension, 1) = 0) Then
		checkFileExtension = true
	Else
		checkFileExtension = false
	End If 
End Function

Function gotoError(errNumber)
	Select Case errNumber
		Case 1000
			MsgBox "The folder you are trying to backup does not exist.", 16, "Missing Folder"
			WScript.Quit
			
		Case 1001
			MsgBox "Could not create the .zip file for backup.", 16, "ZIP FIle"
			WScript.Quit
	End Select
End Function
t

 