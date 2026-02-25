# Print Server Dashboard (PaperCut MF)

Este projeto cria um painel web local para contabilizar impressões a partir dos logs do PaperCut MF no Windows Print Server.

O MVP lê apenas o print log do PaperCut (impressões). Para contabilizar copias/scan, vamos integrar com a API do PaperCut MF ou com o banco externo na proxima etapa.

## Requisitos
- Windows Print Server com PaperCut MF instalado.
- Acesso ao diretorio de logs do PaperCut no servidor.
- Python 3.10+.

## Configuracao
1. Copie `config.example.json` para `config.json` e ajuste os valores.

Campos principais:
- `papercut_log_dir`: diretorio onde o PaperCut grava os print logs.
- `papercut_log_glob`: padrao de nome dos logs.
- `papercut_xmlrpc_url`: endpoint XML-RPC do PaperCut (para integracao futura).
- `papercut_auth_token`: token de autenticacao do PaperCut.
- `db_path`: caminho do SQLite local.
- `printer_poll_enabled`: habilita coleta automatica dos contadores IP.
- `printer_poll_interval_sec`: intervalo de coleta (segundos).

Exemplo de URL do XML-RPC:
```text
https://SERVIDOR:9192/rpc/api/xmlrpc
```

## Instalacao
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Ingestao dos logs
```powershell
python -m app.ingest --since-days 7
```

## Rodar o servidor
```powershell
uvicorn app.main:app --host 0.0.0.0 --port 8088
```

Acesse:
- `http://SERVIDOR:8088/`
- `http://SERVIDOR:8088/api/summary`
- `http://SERVIDOR:8088/api/jobs`
- `http://SERVIDOR:8088/report`

## Configuracao de Setor e Modelo
O dashboard permite cadastrar manualmente:
- Setor por usuario
- Modelo por impressora

Esses dados sao usados nos relatorios e rankings. O campo `source` pode ser `manual`, `ad` ou `windows`.

## Relatorios
Relatorios podem ser gerados por:
- usuario
- setor
- impressora
- modelo

Formatos suportados:
- CSV
- Excel (XLSX)
- PDF

Endpoint:
```
GET /report?group_by=user&since=2026-01-01&until=2026-01-31&format=pdf
```

## Impressoras IP (Contadores)
O dashboard consegue ler contadores de copia/scan/print em impressoras IP, usando a pagina web de manutencao.

Exemplos:
- Brother: `http://IP/etc/mnt_info.html?kind=item`
- Samsung: `http://IP/sws/index.html`

Cadastre no painel em "Impressoras IP (Contadores)".

Para gerar relatorios baseados nos contadores:
```
GET /report-counters?group_by=printer&metric=copy&since=2026-01-01&until=2026-01-31&format=csv
```

## Observacoes
- Este MVP depende do print log do PaperCut (impressao). Copias/scan serao adicionadas via API do PaperCut MF ou leitura do banco.
- Garanta que todos os clientes imprimam via o servidor para contabilizar corretamente.
