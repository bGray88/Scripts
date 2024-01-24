# Software Installation NSI
# Updated: 02.29.16
# Creator: Brandon Gray

# Required for macro VerifyUserAdmin line46
!include LogicLib.nsh

# Script accessible file name
!define INSTALLFILENAME "SOFTWARE_NAME"
# Icon file name for exe
!define ICONFILENAME "${INSTALLFILENAME}.ico"

# General Setting
# -------------------------------------------
# Name of software package
Name "${INSTALLFILENAME}"

# Installer file name
OutFile "${INSTALLFILENAME}-setup.exe"

# Hide installer/Uninstaller actions
ShowInstDetails "nevershow"
ShowUninstDetails "nevershow"

# File compression type
SetCompressor "lzma"

# Update status
SetOverwrite on

# Folder Selection Page
# -------------------------------------------
# Set default install directory
InstallDir "$PROGRAMFILES\${INSTALLFILENAME}"

# User directory prompt
DirText "Choose a folder in which to install ${INSTALLFILENAME}"

# Security and permission requisites
# -------------------------------------------
# Request required admin level permissions
RequestExecutionLevel admin

!macro VerifyUserAdmin
	UserInfo::GetAccountType
	Pop $0
	${If} $0 != "admin" ;Require admin rights on NT4+
        messageBox MB_ICONSTOP "Administrator rights required"
        setErrorLevel 740 ;ERROR_ELEVATION_REQUIRED
        Quit
	${EndIf}
!macroend

# Installer Section
# -------------------------------------------
Section
	
	# Installer Files Section
	# -------------------------------------------
	# Define output path
	SetOutPath "$INSTDIR"

	# Specify file(s) to go in output path
	File /a /r "Data\"
	
	# Set install base to All Users
	setShellVarContext all
	
	CopyFiles /SILENT "$EXEDIR\${INSTALLFILENAME}.7z" "$OUTDIR"
	Nsis7z::Extract "$OUTDIR\${INSTALLFILENAME}.7z"
	Delete "$OUTDIR\${INSTALLFILENAME}.7z"
	
	# Create Shortcut(s)
	# -------------------------------------------
	# Define output path for start in shortcut address
	SetOutPath "$INSTDIR\DOSBox"
	
	# Define Start Menu shortcut(s)
	CreateDirectory "$SMPROGRAMS\DOS\${INSTALLFILENAME}\"
	CreateShortCut "$SMPROGRAMS\DOS\${INSTALLFILENAME}\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
	CreateShortCut "$SMPROGRAMS\DOS\${INSTALLFILENAME}\${INSTALLFILENAME}.lnk" \
		'"$OUTDIR\DOSBox.exe"' \
		'-conf "$INSTDIR\${INSTALLFILENAME}.conf" -noconsole -c "exit"' \
		"$INSTDIR\${ICONFILENAME}"
	
	# Define Desktop shortcut(s)
	CreateShortCut "$DESKTOP\${INSTALLFILENAME}.lnk" \
		'"$OUTDIR\DOSBox.exe"' \
		'-conf "$INSTDIR\${INSTALLFILENAME}.conf" -noconsole -c "exit"' \
		"$INSTDIR\${ICONFILENAME}"
	
	# Request Uninstaller Creation
	# -------------------------------------------
	# Define uninstaller name
	WriteUninstaller "$INSTDIR\uninstall.exe"
	
SectionEnd

# Uninstaller Section
# -------------------------------------------
Section "Uninstall"
	
	# Set install base to All Users
	setShellVarContext all
	
	${If} ${FileExists} "$INSTDIR\${INSTALLFILENAME}.conf"
		# Delete installed file(s)
		RMDir /r "$INSTDIR\*.*"
		
		; Remove the installation directory
		RMDir "$INSTDIR"
	${EndIf}
	
	# Delete Desktop shortcut(s)
	Delete "$DESKTOP\${INSTALLFILENAME}.lnk"
	
	# Delete Start Menu	shortcut(s)
	Delete "$SMPROGRAMS\DOS\${INSTALLFILENAME}\*.*"
	RMDir /r "$SMPROGRAMS\DOS\${INSTALLFILENAME}"
	RMDir "$SMPROGRAMS\DOS"

SectionEnd

# Installer Functions
# -------------------------------------------
# Check user commitment
Function .onInit
	# Set install base to All Users
	setShellVarContext all
	!insertmacro VerifyUserAdmin
FunctionEnd

Function .onInstFailed
	MessageBox MB_OK "The install has failed$\r$\nPlease check the current user \
		admin rights before making another attempt"
FunctionEnd

Function .onUserAbort
	MessageBox MB_YESNO "Abort install?" IDYES NoCancelAbort
    Abort ; Abort the cancel request
	NoCancelAbort:
FunctionEnd

Function .onInstSuccess
    MessageBox MB_OK "The program has been installed successfully$\r$\n\
		Run this program from the icon created on the Desktop or Start Menu"
FunctionEnd

# Uninstaller Functions
# -------------------------------------------
Function un.onInit
    MessageBox MB_YESNO "This will uninstall ${INSTALLFILENAME} from your system- \
		Continue?" IDYES NoAbort
	Abort ; Causes uninstaller to quit
    NoAbort:
	!insertmacro VerifyUserAdmin
FunctionEnd

Function un.onUninstFailed
	MessageBox MB_OK "The uninstall has failed$\r$\nPlease check the current user \
		admin rights before making another attempt"
FunctionEnd

Function un.onUninstSuccess
	MessageBox MB_OK "The uninstall has been successful"
FunctionEnd

Function un.onUserAbort
    MessageBox MB_YESNO "Abort uninstall?" IDYES NoCancelAbort
    Abort ; Abort the cancel request
    NoCancelAbort:
FunctionEnd
