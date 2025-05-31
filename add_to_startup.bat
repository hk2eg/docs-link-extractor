@echo off
set "SOURCE_EXE=%~dp0pdf_spreadsheet_link_extractor.exe"
set "STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "SHORTCUT_NAME=Universal Link Extractor.lnk"

echo Set oWS = WScript.CreateObject("WScript.Shell") > temp.vbs
echo sLinkFile = "%STARTUP_FOLDER%\%SHORTCUT_NAME%" >> temp.vbs
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> temp.vbs
echo oLink.TargetPath = "%SOURCE_EXE%" >> temp.vbs
echo oLink.Arguments = "--silent" >> temp.vbs
echo oLink.WindowStyle = 1 >> temp.vbs
echo oLink.Save >> temp.vbs

cscript //nologo temp.vbs
del temp.vbs

echo ✅ Shortcut added to Startup (silent tray mode).
pause
