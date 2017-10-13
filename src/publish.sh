
NSISEXE="/d/portable/NSISPortable/App/NSIS/Bin/makensis.exe"
SCP_DEST="wangg@192.168.8.173:soft/"

rm -rf out

dotnet publish -c Release -r win-x64 -o out/ExcelConverter.win-x64
dotnet publish -c Release -r osx-x64 -o out/ExcelConverter.osx-x64
dotnet publish -c Release -r linux-x64 -o out/ExcelConverter.linux-x64

cp asset/install.sh out/ExcelConverter.linux-x64/
cp asset/install.sh out/ExcelConverter.osx-x64/
cp asset/addpath.bat out/ExcelConverter.win-x64/
cp asset/win10-64.nsi out/ExcelConverter.win-x64/


$NSISEXE out/ExcelConverter.win-x64/win10-64.nsi

# Bandizip.exe c out/ExcelConverter.win-x64.zip out/ExcelConverter.win-x64/
Bandizip.exe c out/ExcelConverter.osx-x64.zip out/ExcelConverter.osx-x64/
Bandizip.exe c out/ExcelConverter.linux-x64.zip out/ExcelConverter.linux-x64/

scp out/*.zip out/ExcelConverter.win-x64/ExcelConverter.win-x64.installer.exe $SCP_DEST

