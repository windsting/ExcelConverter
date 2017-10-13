# name of the installer
OutFile "ExcelConverter.win-x64.installer.exe"

InstallDir "$PROGRAMFILES64\ExcelConverter"

# SetCompressor /SOLID lzma

# default section start: every NSIS script has least one section
Section

SetOutPath $INSTDIR

File *


WriteUninstaller $INSTDIR\uninstaller.exe

ExecWait '"$INSTDIR\addpath.bat"'

# default section end
SectionEnd

#-------
Section "Uninstall"

Delete $INSTDIR\uninstaller.exe

Delete $INSTDIR\*

SectionEnd