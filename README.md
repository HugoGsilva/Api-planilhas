# Api Planilhas

API em FastAPI para consultar CNPJs na DirectD e gerar planilhas XLSX no modelo de dados do cliente.

## O que o sistema faz

- Protege o frontend com Basic Auth.
- Permite consulta individual de CNPJ.
- Permite envio de planilha XLSX em lote com uma coluna `CNPJ`.
- Consulta a API DirectD `CadastroPessoaJuridica`.
- Gera XLSX para download com os campos do modelo DirectD.
- Mantem apenas arquivos temporarios de jobs em disco, com limpeza por retencao.

## Modelo da planilha gerada

As colunas geradas seguem o CSV modelo do cliente:

```text
cnpj;razaoSocial;nomeFantasia;dataFundacao;cnaeCodigo;cnaeDescricao;cnaEsSecundarios;quantidadeFuncionarios;situacaoCadastral;naturezaJuridicaCodigo;naturezaJuridicaDescricao;naturezaJuridicaTipo;porte;faixaFuncionarios;faixaFaturamento;matriz;orgaoPublico;ramo;tipoEmpresa;telefones;enderecos;emails;ultimaAtualizacaoPJ;socios
```

Listas do JSON sao achatadas:

- `telefones`: todos os `telefoneComDDD`, separados por `, `.
- `emails`: todos os `enderecoEmail`, separados por `, `.
- `enderecos`: `logradouro, numero, complemento, bairro, cidade, uf, cep`; multiplos enderecos sao separados por ` | `.
- `socios`: `nome (cargo)`, separados por `, `.
- `matriz`: `Sim` para `true`, `Nao` para `false`, vazio quando ausente.

## Variaveis de ambiente

Configure em `.env` no desenvolvimento local, ou em `Environment variables` no Portainer/VPS.

Obrigatorias:

- `DIRECTD_TOKEN`: token da DirectD.
- `APP_BASIC_USER`: usuario do Basic Auth.
- `APP_BASIC_PASSWORD`: senha do Basic Auth.

Opcionais:

- `DIRECTD_BASE_URL`: padrao `https://apiv3.directd.com.br`.
- `DIRECTD_TIMEOUT_SECONDS`: padrao `20`.
- `DIRECTD_BATCH_DELAY_SECONDS`: padrao `0`.
- `JOB_STORAGE_DIR`: padrao `storage/jobs`.
- `JOB_RETENTION_HOURS`: padrao `48`.
- `UPLOAD_MAX_MB`: padrao `100`.

## Rodar localmente no Windows

```powershell
cd "C:\Users\Hugo\Desktop\Api-planilhas-hom\Api-planilhas"

Get-Content .env | ForEach-Object {
  $line = $_.Trim()
  if ($line -and -not $line.StartsWith("#") -and $line.Contains("=")) {
    $parts = $line.Split("=", 2)
    [Environment]::SetEnvironmentVariable($parts[0].Trim(), $parts[1].Trim(), "Process")
  }
}

.\.venv\Scripts\python.exe -m uvicorn api_planilhas.web:app --app-dir src --host 127.0.0.1 --port 8000
```

Acesse:

```text
http://127.0.0.1:8000/
```

## Docker

O `Dockerfile` usa imagem oficial Python multi-arch. Em host Arch Linux ARM64, o build local usa a arquitetura do host por padrao.

Build local:

```bash
docker build -t api-planilhas:latest .
```

Run:

```bash
docker run -d \
  --name api-planilhas \
  --restart unless-stopped \
  -p 8000:8000 \
  -e DIRECTD_TOKEN="seu-token" \
  -e APP_BASIC_USER="admin" \
  -e APP_BASIC_PASSWORD="senha" \
  -v api_planilhas_jobs:/app/storage/jobs \
  api-planilhas:latest
```

Build explicito para ARM64:

```bash
docker buildx build --platform linux/arm64 -t api-planilhas:latest .
```

## Portainer

Use `docker-compose.yml` como stack, ou configure o container manualmente.

No Portainer, informe as variaveis em `Environment variables`, sem colocar token real no repositorio:

- `DIRECTD_TOKEN`
- `APP_BASIC_USER`
- `APP_BASIC_PASSWORD`
- demais variaveis opcionais se precisar alterar os padroes

Mantenha um volume persistente em:

```text
/app/storage/jobs
```

Esse volume guarda jobs temporarios, status e planilhas geradas ate a limpeza por `JOB_RETENTION_HOURS`.

## Testes

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## Arquivos principais

- `src/api_planilhas/web.py`: rotas FastAPI e frontend.
- `src/api_planilhas/directd.py`: cliente HTTP da DirectD.
- `src/api_planilhas/converter.py`: mapeamento JSON DirectD para XLSX.
- `src/api_planilhas/xlsx_reader.py`: leitura da coluna `CNPJ` em planilhas de entrada.
- `src/api_planilhas/jobs.py`: persistencia local de jobs em SQLite.
- `src/api_planilhas/batch_processor.py`: processamento sequencial dos CNPJs em lote.
