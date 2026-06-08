# Design: MVP Web DirectD para planilha

## Contexto

O projeto ja possui um conversor local validado que transforma JSONs de CNPJ da DirectD em uma planilha `.xlsx` com o layout aprovado. O proximo passo e criar uma aplicacao web simples para consultar um CNPJ diretamente na API DirectD e entregar a planilha ao usuario.

## Objetivo

Criar um MVP web com FastAPI e uma pagina HTML simples. O usuario acessa a tela com Basic Auth, informa um CNPJ, o backend consulta a DirectD, gera a planilha em memoria e o navegador inicia o download automaticamente. A tela tambem deve manter um botao para baixar novamente caso o download automatico falhe.

## Fora de escopo

- Consulta em lote de multiplos CNPJs.
- API DirectD por lote.
- Frontend separado com build JavaScript.
- Banco de dados.
- Cache persistente em disco.
- Armazenamento permanente de JSONs retornados.
- Armazenamento permanente de planilhas geradas.
- Exibicao de JSON bruto ou resumo detalhado dos dados ao usuario.
- Definicao de regras comerciais para colunas ainda vazias, como `IE`, `Segmento`, `Valor Estimado`, `*Vendedor`, `Titulo da oportunidade` e `Data Ganho/Perdido`.

## Stack

- Python.
- FastAPI.
- HTML, CSS e JavaScript simples servidos pelo backend.
- Conversor atual em `src/api_planilhas/converter.py`.

## Configuracao

A aplicacao deve ler configuracoes por variaveis de ambiente:

- `DIRECTD_TOKEN`: token da API DirectD.
- `APP_BASIC_USER`: usuario do Basic Auth.
- `APP_BASIC_PASSWORD`: senha do Basic Auth.
- `DIRECTD_BASE_URL`: opcional, padrao `https://apiv3.directd.com.br`.

Em producao, essas variaveis serao configuradas no Portainer. Localmente, a aplicacao pode carregar um `.env` se isso for implementado, mas nao deve depender dele em producao.

## Rotas

### `GET /`

Serve a pagina HTML principal protegida por Basic Auth.

A pagina deve conter:

- campo para CNPJ;
- botao `Gerar planilha`;
- estado de carregamento;
- mensagem de sucesso;
- botao `Baixar novamente` apos uma consulta bem-sucedida;
- mensagem de erro quando houver falha.

Nao deve mostrar JSON bruto nem resumo detalhado da empresa.

### `POST /api/planilha`

Recebe um CNPJ e retorna uma planilha `.xlsx` como download.

Regras:

- rota protegida por Basic Auth;
- aceitar apenas um CNPJ por chamada;
- normalizar entrada removendo caracteres nao numericos;
- validar exatamente 14 digitos;
- chamar a API DirectD sempre que o usuario solicitar;
- gerar a planilha em memoria;
- retornar `Content-Type` de XLSX;
- retornar nome de arquivo previsivel, por exemplo `importacao_oportunidade_<cnpj>.xlsx`.

## Integracao DirectD

Endpoint:

`GET {DIRECTD_BASE_URL}/api/CadastroPessoaJuridica?CNPJ={cnpj}&Token={token}`

Regras:

- nao logar token;
- configurar timeout de requisicao;
- em falha HTTP, retornar erro claro para o front;
- em resposta sem `retorno` valido, retornar erro claro para o front;
- nao armazenar permanentemente a resposta.

## Geracao da planilha

O backend deve reaproveitar o conversor atual:

- montar uma linha com `extract_row(payload)`;
- gerar XLSX com `write_xlsx` ou funcao equivalente em memoria;
- manter as colunas sem regra comercial como vazias.

Se o writer atual exigir caminho em disco, a implementacao pode adicionar uma funcao pequena para escrever em buffer de memoria sem alterar o mapeamento.

## Fluxo no navegador

1. Usuario abre `/`.
2. Browser solicita Basic Auth.
3. Usuario informa CNPJ.
4. Usuario clica em `Gerar planilha`.
5. Front envia requisicao para `POST /api/planilha`.
6. Enquanto aguarda, exibe carregamento e bloqueia duplo clique.
7. Se sucesso:
   - cria um download automatico com o blob recebido;
   - mostra mensagem de sucesso;
   - habilita botao `Baixar novamente`.
8. Se erro:
   - mostra mensagem clara;
   - nao baixa arquivo;
   - permite nova tentativa.

## Erros esperados

- CNPJ invalido: `400`.
- Credenciais Basic Auth invalidas: `401`.
- Variavel obrigatoria ausente: `500` com mensagem de configuracao sem expor segredo.
- Timeout DirectD: erro claro para usuario.
- Falha HTTP DirectD: erro claro para usuario.
- Resposta DirectD sem dados uteis: erro claro para usuario.

## Seguranca

- Basic Auth deve proteger pagina e API.
- Senha deve ser comparada de forma segura, evitando comparacao ingenua quando possivel.
- Token DirectD deve vir de ambiente e nunca ser exibido ao usuario.
- Nao salvar `.env` com valores reais no repositorio.
- Nao exibir JSON bruto ao usuario.

## Validacao

Testes automatizados devem cobrir:

- normalizacao e validacao de CNPJ;
- Basic Auth aceitando credenciais corretas;
- Basic Auth rejeitando credenciais incorretas;
- rota de planilha retornando XLSX quando DirectD responde com JSON valido;
- erro para CNPJ invalido;
- erro para token ausente;
- erro para falha/timeout DirectD;
- front HTML contendo formulario, botao de gerar e botao de baixar novamente.

## Entrega

Ao final, a aplicacao deve poder ser executada localmente com um comando documentado, usando variaveis de ambiente. O usuario deve conseguir acessar a URL local, informar um CNPJ, receber o download automatico e usar o botao de baixar novamente.
