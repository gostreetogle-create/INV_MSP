; Inno Setup script for Inventor Pult.
; Build:  ISCC.exe InventorPult.iss   (run from this folder)
; Output: installer\Output\InventorPult-Setup.exe

#define MyAppName "Inventor Pult"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "gostreetogle-create"
#define MyAppExeName "InventorPult.exe"

[Setup]
; Fixed AppId — lets Inno Setup recognise an existing install of this same app on
; upgrade, instead of installing a second copy side by side.
AppId={{2EC932D4-614F-479C-AAB8-6963822EE8F4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Per-user install under Program Files if the user has rights, otherwise falls back
; to a per-user folder — no administrator rights required either way.
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=InventorPult-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Создать значок на рабочем столе"; GroupDescription: "Дополнительные значки:"

[Files]
Source: "..\dist\InventorPult.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\catalog.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Удалить {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
// Belt-and-suspenders: even though Inno Setup already recognises a prior install via
// AppId and offers to upgrade in place, explicitly run the previous version's own
// uninstaller first (silently) so old files never linger — matches "перезаписать,
// удалить старое, установить заново" instead of just merging over it.
function InitializeSetup(): Boolean;
var
  UninstallString: String;
  ResultCode: Integer;
begin
  Result := True;
  if RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1', 'UninstallString', UninstallString) or
     RegQueryStringValue(HKLM, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1', 'UninstallString', UninstallString) then
  begin
    UninstallString := RemoveQuotes(UninstallString);
    Exec(UninstallString, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;
