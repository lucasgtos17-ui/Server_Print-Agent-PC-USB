[Setup]
AppName=Print Server Dashboard
AppVersion=1.0.0
DefaultDirName={pf}\PrintServerDashboard
DefaultGroupName=Print Server Dashboard
OutputBaseFilename=PrintServerDashboardSetup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\icon.ico
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; GroupDescription: "Atalhos:"

[Files]
Source: "dist\PrintServerDashboard.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\PrintServerDashboardService.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.example.json"; DestDir: "{app}"; DestName: "config.json"; Flags: ignoreversion onlyifdoesntexist
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}\data"

[Icons]
Name: "{group}\Dashboard (Abrir navegador)"; Filename: "http://127.0.0.1:8088/"; IconFilename: "{app}\icon.ico"
Name: "{group}\Pasta de instalacao"; Filename: "{app}"; IconFilename: "{app}\icon.ico"
Name: "{commondesktop}\Print Server Dashboard"; Filename: "http://127.0.0.1:8088/"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\PrintServerDashboardService.exe"; Parameters: "install {code:SvcCreds}"; Flags: runhidden; StatusMsg: "Instalando servico..."
Filename: "{app}\PrintServerDashboardService.exe"; Parameters: "start"; Flags: runhidden; StatusMsg: "Iniciando servico..."

[UninstallRun]
Filename: "{app}\PrintServerDashboardService.exe"; Parameters: "stop"; Flags: runhidden; RunOnceId: "StopService"
Filename: "{app}\PrintServerDashboardService.exe"; Parameters: "remove"; Flags: runhidden; RunOnceId: "RemoveService"

[Code]
var
  CredsPage: TInputQueryWizardPage;

function SvcCreds(Param: string): string;
var
  UserName: string;
  Password: string;
begin
  UserName := Trim(CredsPage.Values[0]);
  Password := CredsPage.Values[1];
  if UserName <> '' then
    Result := '-username="' + UserName + '" -password="' + Password + '"'
  else
    Result := '';
end;

procedure InitializeWizard;
begin
  CredsPage := CreateInputQueryPage(
    wpSelectTasks,
    'Credenciais do Servico',
    'Opcional',
    'Informe usuario/senha (ex: DOMINIO\usuario) para rodar o servico com credenciais especificas.'#13#10 +
    'Deixe em branco para rodar como LocalSystem.',
  );
  CredsPage.Add('Usuario (DOMINIO\Usuario):', False);
  CredsPage.Add('Senha:', True);
end;
