#define MyAppName "MaktabOS"
#define MyAppVersion "3.0.0"
#define MyAppPublisher "Ашуров Абдулло"

[Setup]
AppId={{39B82C98-22AD-4550-A532-148E045ED088}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\MaktabOS
UsePreviousAppDir=no
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\MaktabOS.exe
UninstallDisplayName={#MyAppName} {#MyAppVersion}
SetupIconFile=..\assets\maktabos.ico
OutputDir=..\dist
OutputBaseFilename=MaktabOS-Setup-{#MyAppVersion}
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
CloseApplications=yes
RestartApplications=no
Uninstallable=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные ярлыки:"; Flags: unchecked

[Files]
Source: "..\dist\MaktabOS\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\MaktabOS"; Filename: "{app}\MaktabOS.exe"
Name: "{autodesktop}\MaktabOS"; Filename: "{app}\MaktabOS.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\MaktabOS.exe"; Description: "Запустить MaktabOS"; Flags: nowait postinstall skipifsilent
