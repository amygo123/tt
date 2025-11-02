@echo off
setlocal
echo === StyleWatcherWin: Self-contained Single-File build (win-x64) ===
where dotnet >nul 2>nul
if errorlevel 1 (
  echo 未检测到 dotnet 命令。请先安装 .NET 8 SDK：https://dotnet.microsoft.com/download
  pause
  exit /b 1
)
dotnet --info | findstr /i "Version" | findstr "8." >nul
if errorlevel 1 (
  echo 警告：建议使用 .NET 8 SDK 以获得更小更稳的单文件包。继续尝试……
)
dotnet restore
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:TrimMode=partial -p:InvariantGlobalization=true -o publish\win-x64
echo.
echo 构建完成：publish\win-x64\StyleWatcher.exe
pause
