@echo off
setlocal
dotnet publish .\StyleWatcherWin.csproj -c Release -r win-x64 -p:PublishSingleFile=true -p:SelfContained=true -p:TrimMode=partial -p:InvariantGlobalization=true -o publish\win-x64
echo Done. See .\publish\win-x64\
