
@echo off

SET mypath=%~dp0
REM echo %mypath:~0,-1%

echo.>> a.txt
del a.txt
echo Setting environment please be patient!

setx PATH "%mypath%;%path%" /M
