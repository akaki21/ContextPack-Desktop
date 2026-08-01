#ifndef MyAppVersion
  #define MyAppVersion "2.2.0"
#endif

#define MyAppName "ContextPack Desktop"
#define MyAppPublisher "Akaki"
#define MyAppExeName "Start-ContextPack-GUI.cmd"

[Setup]
AppId={{8B492796-DADE-4DB8-915A-3BD2A1D07C64}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\ContextPack Desktop
DefaultGroupName=ContextPack Desktop
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=ContextPack-Setup
SetupIconFile=..\assets\contextpack.ico
UninstallDisplayIcon={app}\assets\contextpack.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
ChangesEnvironment=no
VersionInfoVersion={#MyAppVersion}.0
VersionInfoProductName={#MyAppName}
VersionInfoDescription=AI-ready document context packager
VersionInfoCompany={#MyAppPublisher}

[Tasks]
Name: "desktopicon"; Description: "Create a Desktop shortcut / Desktop shortcut-ის შექმნა"; GroupDescription: "Shortcuts / მალსახმობები:"; Flags: unchecked

[Files]
Source: "..\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs; Excludes: ".git\*,.github\*,.venv\*,input\*,output\*,tessdata\*,dist\*,installer\*,tests\*,__pycache__\*,*.pyc,.gitattributes,.gitignore,CONTRIBUTING.md"
Source: "bootstrap.ps1"; DestDir: "{app}\installer"; Flags: ignoreversion

[Dirs]
Name: "{app}\input"
Name: "{app}\output"
Name: "{app}\tessdata"
Name: "{app}\logs"

[Icons]
Name: "{group}\ContextPack Desktop"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\assets\contextpack.ico"
Name: "{group}\ContextPack Setup Repair"; Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{app}\installer\bootstrap.ps1"" -InstallRoot ""{app}"" -Interactive"; WorkingDir: "{app}"; IconFilename: "{app}\assets\contextpack.ico"
Name: "{group}\Uninstall ContextPack Desktop"; Filename: "{uninstallexe}"
Name: "{autodesktop}\ContextPack Desktop"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\assets\contextpack.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch ContextPack Desktop / ContextPack Desktop-ის გაშვება"; WorkingDir: "{app}"; Flags: postinstall nowait skipifsilent unchecked; Check: CanLaunchApp

[UninstallDelete]
Type: filesandordirs; Name: "{app}\.venv"
Type: filesandordirs; Name: "{app}\tessdata"
Type: filesandordirs; Name: "{app}\logs"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: dirifempty; Name: "{app}"

[Code]
var
  DependencyReady: Boolean;

function CanLaunchApp(): Boolean;
begin
  Result := DependencyReady;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  PowerShellPath: String;
  Parameters: String;
begin
  if CurStep = ssPostInstall then
  begin
    WizardForm.StatusLabel.Caption := 'Installing and checking required components... This can take several minutes. / საჭირო კომპონენტები ყენდება და მოწმდება...';
    PowerShellPath := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
    Parameters := '-NoProfile -ExecutionPolicy Bypass -File "' + ExpandConstant('{app}\installer\bootstrap.ps1') + '" -InstallRoot "' + ExpandConstant('{app}') + '"';
    DependencyReady := Exec(PowerShellPath, Parameters, ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
    if not DependencyReady then
    begin
      Log('ContextPack dependency bootstrap failed. See: ' + ExpandConstant('{app}\logs\install.log'));
      if not WizardSilent then
      MsgBox(
        'ContextPack files were installed, but one or more required components could not be prepared.' + #13#10 +
        'Open "ContextPack Setup Repair" from the Start menu to retry.' + #13#10#13#10 +
        'ContextPack-ის ფაილები დაყენდა, მაგრამ საჭირო კომპონენტების მომზადება ვერ დასრულდა.' + #13#10 +
        'ხელახლა საცდელად Start მენიუდან გახსენი "ContextPack Setup Repair".' + #13#10#13#10 +
        'Log: ' + ExpandConstant('{app}\logs\install.log'),
        mbError, MB_OK);
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  DependencyReady := False;
  Result := True;
end;
