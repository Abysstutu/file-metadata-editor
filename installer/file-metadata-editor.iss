#ifndef AppVersion
  #error AppVersion must be supplied with /DAppVersion=...
#endif
#ifndef SourceDir
  #error SourceDir must be supplied with /DSourceDir=...
#endif

#define AppName "文件源数据编辑"
#define AppExe "文件源数据编辑.exe"

[Setup]
AppId={{A0739180-9FD1-4CF6-A5EE-01AB1ED62287}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Abysstutu
AppPublisherURL=https://github.com/Abysstutu/file-metadata-editor
AppSupportURL=https://github.com/Abysstutu/file-metadata-editor/issues
DefaultDirName={localappdata}\Programs\FileMetadataEditor
DefaultGroupName={#AppName}
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
OutputBaseFilename=FileMetadataEditor-v{#AppVersion}-Windows-x64-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExe}
SetupIconFile=..\assets\app-icon.ico

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"

[Run]
Filename: "{app}\{#AppExe}"; Description: "启动{#AppName}"; Flags: nowait postinstall skipifsilent
