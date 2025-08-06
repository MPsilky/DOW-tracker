;--------------------------------
; DOW 30 Tracker Installer Script
;--------------------------------
[Setup]
; App metadata
AppName=DOW 30 Tracker
AppVersion=1.0.0
DefaultDirName={pf}\DOW 30 Tracker
DefaultGroupName=DOW 30 Tracker
OutputBaseFilename=DOW30TrackerSetup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

;--------------------------------
[Files]
; Install the EXE out of your PyInstaller dist folder
Source: "dist\DOW30_Excel_Dashboard.exe"; DestDir: "{app}"; Flags: ignoreversion
; (If you have any additional DLLs or resources, list them here too)
; Source: "dist\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

;--------------------------------
[Icons]
; Start‑menu shortcut
Name: "{group}\DOW 30 Tracker"; Filename: "{app}\DOW30_Excel_Dashboard.exe"
; Optional: desktop shortcut
Name: "{commondesktop}\DOW 30 Tracker"; Filename: "{app}\DOW30_Excel_Dashboard.exe"; Tasks: desktopicon

;--------------------------------
[Tasks]
Name: desktopicon; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

;--------------------------------
[Run]
; Offer to launch the app at the end of setup
Filename: "{app}\DOW30_Excel_Dashboard.exe"; Description: "Launch DOW 30 Tracker"; Flags: nowait postinstall skipifsilent
