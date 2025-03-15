@echo off
echo Loading NetworkLabeler in Civil 3D...

:: Set the path to your compiled DLL
set DLL_PATH=%~dp0bin\Debug\net48\NetworkLabeler.dll

:: Create a temporary LSP file to load the DLL
echo (command "NETLOAD" "%DLL_PATH%") > "%TEMP%\LoadNetworkLabeler.lsp"

:: Start Civil 3D with the LSP file
start "" "C:\Program Files\Autodesk\AutoCAD 2024\acad.exe" /ld "C:\Program Files\Autodesk\AutoCAD 2024\AecBase.dbx" /p "<<C3D_Metric>>" /product C3D /language en-US /nologo /b "%TEMP%\LoadNetworkLabeler.lsp"

echo.
echo Civil 3D is starting. Once it's loaded, you can start debugging in Visual Studio Code.
echo.
echo Press any key to exit...
pause > nul