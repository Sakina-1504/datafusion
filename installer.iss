; DataFusion Platform - Inno Setup installer script
; This builds a proper Setup.exe with Start Menu shortcut, Desktop
; shortcut (optional), and an uninstaller.

#define MyAppName "DataFusion Platform"
#define MyAppVersion "1.0"
#define MyAppPublisher "Sakina"
#define MyAppExeName "DataFusionPlatform.exe"

[Setup]
AppId={{8F2B6E1A-9C3D-4A5E-9B7F-DATAFUSION001}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
SetupIconFile=assets\app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
; Output installer file name and location
OutputDir=installer_output
OutputBaseFilename=DataFusionPlatform_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Doesn't require admin rights - installs to the user's own folder instead
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main executable built by PyInstaller
Source: "dist\DataFusionPlatform.exe"; DestDir: "{app}"; Flags: ignoreversion
; Sample data, so users have something to test with right after installing
Source: "sample_data\*"; DestDir: "{app}\sample_data"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent