Option Explicit

Const ForReading = 1
Const ForWriting = 2

Dim objShell, objFSO, objFile, curDir, strText, strNewWants, strNewWantsMode

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
objShell.CurrentDirectory = ".."
Set objFile = objFSO.OpenTextFile(objShell.CurrentDirectory & _
	"\config\loginusers.vdf", ForReading)
strText = objFile.ReadAll

objFile.Close

strNewWants = Replace(strText, """WantsOfflineMode""" & "		" & """0""", _
	"""WantsOfflineMode""" & "		" & """1""")
strNewWantsMode = Replace(strNewWants, """SkipOfflineModeWarning""" & "		" & """0""", _
	"""SkipOfflineModeWarning""" & "		" & """1""")
Set objFile = objFSO.OpenTextFile(objShell.CurrentDirectory & _
	"\config\loginusers.vdf", ForWriting)

objFile.WriteLine strNewWantsMode

Set objShell = Nothing
objFile.Close
