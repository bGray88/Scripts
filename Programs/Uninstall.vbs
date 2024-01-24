On Error Resume Next

const HKLM = &H80000002
strComputer = "."

if (Wscript.Arguments.Count <> 1) Then
	Wscript.Echo "Correct Usage: Uninstall.vbs 'Software Name'"
	Wscript.Quit
End If

For Each subarg In Wscript.Arguments
	softName = subarg
Next

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
strCommand = "Select * from Win32_Product WHERE Name LIKE " _
	& chr(34) & chr(37) & softName & chr(37) & chr(34)
'Wscript.Echo strCommand
Set colSoftware = objWMIService.ExecQuery(strCommand)

Set objShell = Wscript.CreateObject ("Wscript.shell")
Set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
objReg.EnumKey HKLM, strKeyPath, arrSubKeys
For Each objSoftware in colSoftware
'	Wscript.Echo objSoftware.Caption & VBCR & objSoftware.Name
	For Each subkey In arrSubKeys
		objReg.GetStringValue HKLM, strKeyPath & "\" & subkey, "DisplayName", strNameValue
		If LCase(strNameValue) = LCase(objSoftware.Caption) Then
			objReg.GetStringValue HKLM, strKeyPath & "\" & subkey, "UninstallString", strUninstallValue
			strUninstallValue = Replace(strUninstallValue, "MsiExec.exe /I", "MsiExec.exe /X")
			strUninstallValue = Replace(strUninstallValue, "MsiExec.exe /X", "MsiExec.exe /norestart /quiet /X")
'			Wscript.Echo "Match Found: " & strNameValue & VBCR & "Uninstall String:" & VBCR & strUninstallValue
			objShell.Run strUninstallValue, 1, TRUE
		End If
	Next
Next
