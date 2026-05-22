[Setup]
; Unique identifier for your app (DO NOT CHANGE THIS ONCE PUBLISHED, it's used for uninstalling)
AppId={{94A7B1C8-5D8F-4E32-A1B9-7C6D2E1F4A9B}
AppName=GST & Income Tax Automation Suite
AppVersion=1.0
AppPublisher=AutomationCafe
DefaultDirName={userappdata}\AutomationCafe
PrivilegesRequired=lowest
DisableDirPage=yes
DefaultGroupName=AutomationCafe
DisableProgramGroupPage=yes
; Where the final Installer.exe will be saved
OutputDir=dist
; The name of the installer file your users will download
OutputBaseFilename=GST_Suite_Installer
; App Icon
SetupIconFile=studycafelogo.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; The executable we are packing
Source: "dist\AutomationCafe.exe"; DestDir: "{app}"; Flags: ignoreversion
; Include the updater executable if needed
Source: "dist\StudyCafeSuite_Updater.exe"; DestDir: "{app}"; Flags: ignoreversion
; (You can add more Source: lines here if you have extra folders like 'input' that need to be shipped alongside the EXE)

[Icons]
; Start menu shortcut
Name: "{group}\AutomationCafe"; Filename: "{app}\AutomationCafe.exe"; IconFilename: "{app}\AutomationCafe.exe"; WorkingDir: "{userdocs}"
; Desktop shortcut (if user checks the box)
Name: "{autodesktop}\AutomationCafe"; Filename: "{app}\AutomationCafe.exe"; Tasks: desktopicon; IconFilename: "{app}\AutomationCafe.exe"; WorkingDir: "{userdocs}"

[Run]
; Launch the app automatically after installation finishes
Filename: "{app}\AutomationCafe.exe"; Description: "{cm:LaunchProgram,AutomationCafe}"; Flags: nowait postinstall skipifsilent
