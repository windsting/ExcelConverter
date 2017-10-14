
BANDIZIP="/c/Program Files/Bandizip/Bandizip.exe"
NSISEXE="/d/portable/NSISPortable/App/NSIS/Bin/makensis.exe"
SCP_DEST="wangg@192.168.8.173:soft/"

rm -rf out

dotnet publish -c Release -r linux-x64 -o out/ExcelConverter.linux-x64
dotnet publish -c Release -r osx-x64 -o out/ExcelConverter.osx-x64
dotnet publish -c Release -r win-x64 -o out/ExcelConverter.win-x64

cp asset/install.sh out/ExcelConverter.linux-x64/
cp asset/install.sh out/ExcelConverter.osx-x64/
cp asset/addpath.bat out/ExcelConverter.win-x64/
cp asset/win10-64.nsi out/


$NSISEXE out/win10-64.nsi

"$BANDIZIP" c -l:9 out/ExcelConverter.linux-x64.zip out/ExcelConverter.linux-x64/
"$BANDIZIP" c -l:9 out/ExcelConverter.osx-x64.zip out/ExcelConverter.osx-x64/
"$BANDIZIP" c -l:9 out/ExcelConverter.win-x64.zip out/ExcelConverter.win-x64/

scp out/*.zip $SCP_DEST

