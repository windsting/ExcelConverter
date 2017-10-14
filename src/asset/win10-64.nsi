# name of the installer
OutFile "ExcelConverter.win-x64.installer.exe"

InstallDir "$PROGRAMFILES64\ExcelConverter"

# SetCompressor /SOLID lzma

# default section start: every NSIS script has least one section
Section

SetOutPath $INSTDIR

File /r ExcelConverter.win-x64\*.*

WriteUninstaller $INSTDIR\uninstaller.exe

ExecWait '"$INSTDIR\addpath.bat"'

MessageBox MB_OK "Installation Succeed! Enjoy it."

# default section end
SectionEnd

#-------
Section "Uninstall"

Delete $INSTDIR\uninstaller.exe

Delete $INSTDIR\*

SectionEnd