# Print Client Agent (USB / Local printers)

Este agente roda na maquina que tem a impressora USB/local e envia os jobs para o servidor.

## Requisitos
- Windows
- Python 3.10+

## Instalar
```powershell
cd C:\Elemento\Scripts\print_client_agent
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Configurar
1. Copie `config.example.json` para `config.json`
2. Rode o configurador:
```powershell
python config_ui.py
```

## Rodar agente
```powershell
python agent.py
```

## Gerar executavel (.exe)
```powershell
cd C:\Elemento\Scripts\print_client_agent
.\build_agent.ps1
```

Executaveis gerados:
- `dist\PrintClientAgent.exe` (configurador UI)
- `dist\PrintClientAgentService.exe` (servico do Windows)

## Instalar como servico (manual)
```powershell
cd C:\Elemento\Scripts\print_client_agent\dist
.\PrintClientAgentService.exe install
.\PrintClientAgentService.exe start
```

Para remover:
```powershell
.\PrintClientAgentService.exe stop
.\PrintClientAgentService.exe remove
```

Observacao: se a impressora USB estiver instalada apenas para o usuario, o servico pode nao enxergar a fila. Nesse caso, instale a impressora "para todos os usuarios" ou instale o servico com credenciais de usuario.

## Criar instalador (Inno Setup)
1. Gere os executaveis com `.\build_agent.ps1`
2. Abra `installer.iss` no Inno Setup e clique em Compile.

O instalador resultante cria atalhos e instala/inicia o servico automaticamente.

O agente vai monitorar a impressora selecionada e enviar para o servidor em:
`http://SERVIDOR:8088/api/client-jobs`
