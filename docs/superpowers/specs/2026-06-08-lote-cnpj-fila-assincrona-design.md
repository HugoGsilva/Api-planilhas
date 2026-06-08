# Lote de CNPJ com fila assincrona

## Contexto

O MVP atual consulta um CNPJ por vez na DirectD e devolve uma planilha XLSX no mesmo request. Para planilhas maiores, manter a requisicao HTTP aberta ate o fim do processamento e fragil: navegador, proxy, VPS ou servidor podem encerrar a conexao antes da conclusao.

A documentacao publica da DirectD para `CadastroPessoaJuridica` expoe o endpoint individual `GET /api/CadastroPessoaJuridica`. A FAQ oficial informa que atualmente nao ha limite de requisicoes por minuto ou hora nas APIs, mas cada CPF/CNPJ consultado conta como uma consulta, mesmo quando repetido.

## Objetivo

Permitir que o usuario envie uma planilha XLSX com uma coluna `CNPJ`, crie um trabalho de processamento em fila, acompanhe o progresso e baixe a planilha final quando o processamento terminar.

O sistema deve evitar manutencao manual na VPS usando retencao automatica de jobs antigos.

## Fora de escopo

- Integrar com um endpoint oficial de lote da DirectD, pois o contrato publico identificado para API automatizada e individual.
- Criar Redis, Celery, banco externo ou worker separado.
- Autenticacao alem do Basic Auth ja existente.
- Garantir processamento literalmente infinito. O sistema nao tera limite pequeno de linhas, mas ainda dependera de saldo DirectD, espaco em disco, memoria e tamanho maximo de upload configurado.

## Arquitetura

A aplicacao FastAPI tera uma fila local simples:

- `storage/jobs/`: pasta local para arquivos de jobs.
- SQLite local para metadados de jobs, progresso e lista de erros.
- Um worker em background dentro do mesmo processo da aplicacao.
- Processamento sequencial, um CNPJ por vez, sem paralelismo.

Essa escolha evita servicos extras no Portainer e reduz manutencao operacional.

## Configuracao

Novas variaveis de ambiente:

- `JOB_STORAGE_DIR`: pasta de armazenamento dos jobs. Padrao: `storage/jobs`.
- `JOB_RETENTION_HOURS`: tempo de retencao de jobs concluidos/falhos. Padrao: `48`.
- `UPLOAD_MAX_MB`: tamanho maximo aceito para upload. Padrao inicial: `100`.
- `DIRECTD_BATCH_DELAY_SECONDS`: atraso entre consultas. Padrao: `0`, alinhado a FAQ oficial que informa ausencia de limite por minuto/hora.

Valores invalidos devem falhar de forma clara na inicializacao ou no request correspondente.

## Fluxo do usuario

1. Usuario acessa a tela com Basic Auth.
2. Usuario escolhe uma planilha XLSX.
3. Backend valida tamanho e formato.
4. Backend le a primeira aba e procura uma coluna `CNPJ`.
5. Se nao encontrar a coluna, retorna o erro grande:
   `nao tem cnpj na sua planilha adicione uma coluna CNPJ e os cnpj em baixo`
6. Se encontrar a coluna, cria um job e retorna o `job_id`.
7. Front passa a consultar o status do job periodicamente.
8. Worker processa cada CNPJ:
   - normaliza para 14 digitos;
   - ignora celulas vazias;
   - registra erro para CNPJ invalido;
   - chama DirectD para CNPJ valido;
   - em sucesso, adiciona a linha ao resultado;
   - em falha, registra o CNPJ e a mensagem do erro.
9. Quando o job conclui, o usuario baixa a planilha final.
10. A tela informa quais CNPJs deram erro.

## Contratos HTTP

Todos os endpoints continuam protegidos por Basic Auth.

### `POST /api/lotes`

Recebe multipart/form-data com campo `file`.

Respostas:

- `202 Accepted`: cria job e retorna JSON com `job_id`.
- `400 Bad Request`: planilha invalida, sem coluna CNPJ ou arquivo fora do formato.
- `413 Payload Too Large`: arquivo acima de `UPLOAD_MAX_MB`.

Erro obrigatorio quando faltar coluna:

```text
nao tem cnpj na sua planilha adicione uma coluna CNPJ e os cnpj em baixo
```

### `GET /api/lotes/{job_id}`

Retorna status do job em JSON:

- `status`: `queued`, `processing`, `completed` ou `failed`.
- `total`: quantidade de CNPJs validos/informados para processamento.
- `processed`: quantidade ja processada.
- `success`: quantidade de consultas bem-sucedidas.
- `errors`: lista de CNPJs com erro e mensagem.
- `download_ready`: booleano.

### `GET /api/lotes/{job_id}/download`

Retorna o XLSX final quando `status=completed`.

Se ainda nao estiver pronto, retorna `409 Conflict`.

## Planilha final

A planilha final deve manter o mesmo mapeamento aprovado do MVP atual. Linhas com erro nao entram na aba principal de oportunidades.

Para informar falhas ao usuario sem poluir o layout principal, a primeira versao informa os erros apenas na tela de status. A planilha final tera somente a aba principal com o mesmo layout aprovado do MVP atual.

## Retencao automatica

Ao iniciar a aplicacao e periodicamente enquanto ela roda, o sistema remove jobs antigos cujo `created_at` ou `finished_at` ultrapasse `JOB_RETENTION_HOURS`.

A limpeza deve remover:

- planilha original;
- planilha final;
- registros SQLite associados ao job.

Falhas de limpeza devem ser registradas de forma controlada e nao devem derrubar a aplicacao.

## Tratamento de erros

- Planilha sem coluna `CNPJ`: erro grande obrigatorio na tela.
- CNPJ invalido: job continua e registra erro daquele valor.
- Erro DirectD em um CNPJ: job continua e registra erro daquele CNPJ.
- Falha estrutural do job, como arquivo ilegivel apos criacao: job vira `failed`.
- Se nenhum CNPJ gerar sucesso, o job ainda pode concluir com planilha vazia e lista de erros, desde que a falha nao seja estrutural.

## Frontend

A tela atual deve ganhar uma area de lote:

- input para upload XLSX;
- botao para iniciar processamento;
- indicador de status/progresso;
- lista de CNPJs com erro;
- botao de download quando pronto.

O fluxo individual de CNPJ deve continuar existindo.

## Testes

Cobertura minima:

- extrair coluna `CNPJ` de planilha XLSX;
- rejeitar planilha sem coluna `CNPJ` com a mensagem obrigatoria;
- criar job e consultar status;
- processar lote com sucessos e falhas sem parar no primeiro erro;
- gerar planilha final com apenas sucessos;
- proteger endpoints de lote com Basic Auth;
- limpeza automatica remove jobs expirados;
- limites/configuracoes invalidas falham de forma clara.
