#define MyAppName "CCTV Health Monitoring System"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Mr. Neeraj Kumar"
#define MyAppExeName "CCTV Health Monitoring System.exe"

[Setup]
AppId={{CCTV-Health-Monitoring-System-2026}}
AppName=CCTV Health Monitoring System
AppVersion=1.0.0
AppPublisher=Mr. Neeraj Kumar
DefaultDirName={autopf}\CCTV Health Monitoring System
DefaultGroupName=CCTV Health Monitoring System
OutputDir=installer
OutputBaseFilename=CCTV_Health_Monitoring_System_Setup_v1.0.0
SetupIconFile=CCTV.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
Uninstallable=yes

[Files]
Source: "dist\CCTV Health Monitoring System\*"; \
DestDir: "{app}"; \
Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autodesktop}\CCTV Health Monitoring System"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\CCTV Health Monitoring System"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\Uninstall CCTV Health Monitoring System"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch CCTV Health Monitoring System"; Flags: nowait postinstall skipifsilent