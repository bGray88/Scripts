
'******************************************************************************
' Purpose:	7z selected files to maximum zip specifications
'
' Inputs:	The full folder path to files, files with specific extension
'			to be zipped
'******************************************************************************
Sub ZipFiles7z(strFolderSelected, strFileExt)
	Const ShowWindow = 1, DontShowWindow = 0
	Const WaitUntilFinished = True, DontWaitUntilFinished = False
	Const objShellFinished = 1, objShellFailed = 2
	Dim strZipExe : strZipExe = "C:\Program Files\7-Zip\7zG.exe"
	Dim strZipExt : strZipExt = ".7z"
	Dim strZipFlags : strZipFlags = "a -t7z -m0=lzma2:d27 -mx9 -mmt8"
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderSelected)
	Dim colFiles : Set colFiles = objFolder.Files
	
	Dim objShellExec, objFile
	Dim strSelectedPath, strItemExtLoc, strExtLocPath, strExtName
	For Each objFile In colFiles
		strSelectedPath = strFolderSelected & "\"
		strItemExtLoc = (InStr(objFile, "."))
		strExtLocPath = Left(objFile, strItemExtLoc)
		strExtName = Replace(objFile, strExtLocPath, "")
		If strExtName = strFileExt Then
			Set objShellExec = objShell.Exec(DoubleQuotes(strZipExe) & " " & _
			strZipFlags & " " & _
			DoubleQuotes(strSelectedPath & Replace(objFile.Name,"." & strExtName, "") & strZipExt) & " " & _
			DoubleQuotes(strSelectedPath & objFile.Name))
			Do While objShellExec.Status <> objShellFinished
				WScript.Sleep 100
			Loop
		End If
	Next
	
	Set objFSO = Nothing
	Set objShell = Nothing
	Set objFolder = Nothing
	Set colFiles = Nothing
	Set objShellExec = Nothing
End Sub
