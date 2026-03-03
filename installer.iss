; Inno Setup Script - 热门询价图片工具 Windows Installer

#define MyAppName "热门询价图片工具"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ZhouRui"
#define MyAppExeName "热门询价图片工具.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=HotInquiryTool_Setup_v{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes

[Tasks]
Name: "desktopicon"; Description: "Create desktop shortcut"; Flags: checked

[Files]
Source: "dist\热门询价图片工具\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "tesseract\*"; DestDir: "{app}\tesseract"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\uploads"
