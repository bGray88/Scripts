# Software Installation NSI
# Updated: 02.29.16
# Creator: Brandon Gray

# Required for macro VerifyUserAdmin
!include LogicLib.nsh

# Script accessible file name
!define BASEFILENAME "PACKAGE_NAME"
!define BASELONGFILENAME "SOFTWARE_FULL_NAME"
!define EXPANSLONGFILENAME "EXPANSION_FULL_NAME"
# Executable file name(s)
!define EXEFILENAME "SOFTWARE.EXE"
!define EXEEXPANSFILENAME "EXPANSION.EXE"
# Icon file name(s) for exe
!define ICONFILENAME "${BASELONGFILENAME}.ico"
!define ICONEXPANSFILENAME "${EXPANSLONGFILENAME}.ico"

# Script accessible software type descriptor
!define SOFTWARETYPE "SOFTWARE_CLASSIFICATION"

# General Setting
# -------------------------------------------
# Name of software package
Name "${BASEFILENAME}"

# Installer file name
OutFile "${BASEFILENAME}-setup.exe"

# Hide installer/Uninstaller actions
ShowInstDetails "nevershow"
ShowUninstDetails "nevershow"

# File compression type
SetCompress off

# Update status
SetOverwrite on

# Folder Selection Page
# -------------------------------------------
# Set default install directory
InstallDir "$PROGRAMFILES\${BASEFILENAME}"

# User directory prompt
DirText "Choose a folder in which to install ${BASEFILENAME}"

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
	File "Data\Data.7z"
	
	# Set install base to All Users
	setShellVarContext all
	
	Nsis7z::Extract "Data.7z"
	Delete "$INSTDIR\Data.7z"
	CopyFiles /SILENT "$EXEDIR\${BASEFILENAME}.7z" "$INSTDIR"
	Nsis7z::Extract "${BASEFILENAME}.7z"
	Delete "$INSTDIR\${BASEFILENAME}.7z"
	
	# Create Shortcut(s)
	# -------------------------------------------
	# Define output path for start in shortcut address
	SetOutPath "$INSTDIR"
	
	# Define Start Menu shortcut(s)
	CreateDirectory "$SMPROGRAMS\${SOFTWARETYPE}\${BASEFILENAME}\"
	CreateShortCut "$SMPROGRAMS\${SOFTWARETYPE}\${BASEFILENAME}\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
	
	CreateShortCut "$SMPROGRAMS\${SOFTWARETYPE}\${BASEFILENAME}\${BASELONGFILENAME}.lnk" \
		'"$INSTDIR\${EXEFILENAME}"' "" "$INSTDIR\${ICONFILENAME}"
	
	CreateShortCut "$SMPROGRAMS\${SOFTWARETYPE}\${BASEFILENAME}\${EXPANSLONGFILENAME}.lnk" \
		'"$INSTDIR\${EXEEXPANSFILENAME}"' "" "$INSTDIR\${ICONEXPANSFILENAME}"
	
	# Define Desktop shortcut(s)
	CreateShortCut "$DESKTOP\${BASELONGFILENAME}.lnk" \
		'"$INSTDIR\${EXEFILENAME}"' "" "$INSTDIR\${ICONFILENAME}"
	
	CreateShortCut "$DESKTOP\${EXPANSLONGFILENAME}.lnk" \
		'"$INSTDIR\${EXEEXPANSFILENAME}"' "" "$INSTDIR\${ICONEXPANSFILENAME}"
	
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
	
	${If} ${FileExists} "$INSTDIR\${EXEFILENAME}"
		# Delete installed file(s)
		RMDir /r "$INSTDIR\*.*"
		
		# Remove the installation directory
		RMDir "$INSTDIR"
	${EndIf}
	
	# Delete Desktop shortcut(s)
	Delete "$DESKTOP\${BASELONGFILENAME}.lnk"
	Delete "$DESKTOP\${EXPANSLONGFILENAME}.lnk"
	
	# Delete Start Menu	shortcut(s)
	Delete "$SMPROGRAMS\${SOFTWARETYPE}\${BASEFILENAME}\*.*"
	RMDir /r "$SMPROGRAMS\${SOFTWARETYPE}\${BASEFILENAME}"
	RMDir "$SMPROGRAMS\${SOFTWARETYPE}"

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
    MessageBox MB_YESNO "This will uninstall ${BASEFILENAME} from your system- \
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
