[Setup]
AppId=AuditFlow.Desktop
AppName=AuditFlow
AppVersion=0.1.0
AppPublisher=ToddQiu
AppPublisherURL=https://github.com/ToddQiu/AuditFlow
DefaultDirName={localappdata}\Programs\AuditFlow
DefaultGroupName=AuditFlow
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=AuditFlow-Setup-0.1.0
SetupIconFile=..\assets\audit-a-20261011.ico
UninstallDisplayIcon={app}\AuditFlow.exe
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=no

[Files]
Source: "..\dist\AuditFlow\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userdesktop}\AuditFlow"; Filename: "{app}\AuditFlow.exe"
Name: "{group}\AuditFlow"; Filename: "{app}\AuditFlow.exe"

[Run]
Filename: "{app}\AuditFlow.exe"; Description: "Open AuditFlow"; Flags: nowait postinstall skipifsilent unchecked
