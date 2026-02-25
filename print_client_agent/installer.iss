[Setup]
AppName=Print Client Agent
AppVersion=1.0.0
DefaultDirName={pf}\PrintClientAgent
DefaultGroupName=Print Client Agent
OutputBaseFilename=PrintClientAgentSetup
SetupIconFile=icon.ico
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na area de trabalho"; GroupDescription: "Atalhos:"

[Files]
Source: "dist\PrintClientAgent.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\PrintClientAgentService.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.example.json"; DestDir: "{app}"; DestName: "config.json"; Flags: ignoreversion onlyifdoesntexist
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Configurar Agente"; Filename: "{app}\PrintClientAgent.exe"; IconFilename: "{app}\icon.ico"
Name: "{commondesktop}\Configurar Agente"; Filename: "{app}\PrintClientAgent.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\PrintClientAgentService.exe"; Parameters: "install {code:SvcCreds}"; Flags: runhidden; StatusMsg: "Instalando servico..."
Filename: "{app}\PrintClientAgentService.exe"; Parameters: "start"; Flags: runhidden; StatusMsg: "Iniciando servico..."

[UninstallRun]
Filename: "{app}\PrintClientAgentService.exe"; Parameters: "stop"; Flags: runhidden; RunOnceId: "StopService"
Filename: "{app}\PrintClientAgentService.exe"; Parameters: "remove"; Flags: runhidden; RunOnceId: "RemoveService"

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
    'Se desejar, informe usuario e senha para o servico rodar com credenciais administrativas.'#13#10 +
    'Deixe em branco para rodar como LocalSystem.',
  );
  CredsPage.Add('Usuario (DOMINIO\\Usuario):', False);
  CredsPage.Add('Senha:', True);
end;
