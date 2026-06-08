# Design: Conversor JSON CNPJ para planilha

## Contexto

O projeto deve transformar retornos JSON de consultas de CNPJ da DirectD em uma planilha de importacao de oportunidades. Nesta primeira entrega, a aplicacao nao chamara a API externa; ela usara os arquivos JSON locais ja presentes em `cnpjs/cnpjs` como fonte de dados.

A integracao real com a DirectD fica para uma fase posterior. A documentacao do endpoint de CNPJ indica o formato `GET https://apiv3.directd.com.br/api/CadastroPessoaJuridica?CNPJ={0}&Token=Seu_Token`, com resposta JSON contendo `metaDados` e `retorno`.

## Objetivo

Gerar uma planilha `.xlsx` a partir dos arquivos JSON locais, mantendo somente os campos representados no layout aprovado. Cada arquivo JSON deve gerar uma linha na planilha.

## Fora de escopo

- Chamar a API DirectD em tempo real.
- Criar frontend.
- Implementar Basic Auth.
- Persistir token em `.env`.
- Gerar ou validar oportunidades em sistemas externos.
- Mapear campos que nao existam no layout aprovado.

## Entrada

- Diretorio fonte: `cnpjs/cnpjs`.
- Formato esperado: arquivos `.json`.
- Estrutura esperada por arquivo:
  - `metaDados`: metadados da consulta.
  - `retorno`: dados da pessoa juridica.

O conversor deve ignorar `metaDados` para a planilha, exceto se futuramente houver uma regra explicita para usa-lo.

## Saida

Uma planilha `.xlsx` com uma aba de importacao e exatamente estas colunas, nesta ordem:

1. Pessoa
2. Telefone1
3. Telefone2
4. email Prin
5. email Secu
6. Cargo
7. CPF
8. Organizacao
9. CNPJ
10. CEP
11. Logradouro
12. Numero
13. Complemento
14. Bairro
15. Cidade
16. UF
17. IE
18. Segmento
19. Valor Estimado
20. *Nome Fantasia
21. *CNAE
22. *Vendedor
23. Titulo da oportunidade
24. Status
25. Data Ganho/Perdido
26. *QUANTIDADE DE FUNCIONARIOS

## Mapeamento

| Coluna | Origem no JSON | Regra |
| --- | --- | --- |
| Pessoa | `retorno.razaoSocial` | Usar valor direto. |
| Telefone1 | `retorno.telefones[0].telefoneComDDD` | Usar primeiro telefone, se existir. |
| Telefone2 | `retorno.telefones[1].telefoneComDDD` | Usar segundo telefone, se existir. |
| email Prin | `retorno.emails[0].enderecoEmail` | Usar primeiro e-mail, se existir. |
| email Secu | `retorno.emails[1].enderecoEmail` | Usar segundo e-mail, se existir. |
| Cargo | `retorno.socios[0].cargo` | Usar primeiro socio, se existir. |
| CPF | `retorno.socios[0].documento` | Usar documento do primeiro socio, se existir. |
| Organizacao | `retorno.razaoSocial` | Usar valor direto. |
| CNPJ | `retorno.cnpj` | Usar valor direto. |
| CEP | `retorno.enderecos[0].cep` | Usar primeiro endereco, se existir. |
| Logradouro | `retorno.enderecos[0].logradouro` | Usar primeiro endereco, se existir. |
| Numero | `retorno.enderecos[0].numero` | Usar primeiro endereco, se existir. |
| Complemento | `retorno.enderecos[0].complemento` | Usar primeiro endereco, se existir. |
| Bairro | `retorno.enderecos[0].bairro` | Usar primeiro endereco, se existir. |
| Cidade | `retorno.enderecos[0].cidade` | Usar primeiro endereco, se existir. |
| UF | `retorno.enderecos[0].uf` | Usar primeiro endereco, se existir. |
| IE | vazio | Sem origem definida nesta fase. |
| Segmento | vazio | Sem origem definida nesta fase. |
| Valor Estimado | vazio | Sem origem definida nesta fase. |
| *Nome Fantasia | `retorno.nomeFantasia` | Usar valor direto. |
| *CNAE | `retorno.cnaeCodigo` e `retorno.cnaeDescricao` | Preferir `codigo - descricao`; se faltar codigo, usar descricao. |
| *Vendedor | vazio | Sem origem definida nesta fase. |
| Titulo da oportunidade | vazio | Sem origem definida nesta fase. |
| Status | `retorno.situacaoCadastral` | Usar valor direto. |
| Data Ganho/Perdido | vazio | Sem origem definida nesta fase. |
| *QUANTIDADE DE FUNCIONARIOS | `retorno.quantidadeFuncionarios` | Usar valor direto. |

## Tratamento de listas

- Telefones: usar somente os dois primeiros itens da lista.
- E-mails: usar somente os dois primeiros itens da lista.
- Enderecos: usar somente o primeiro item da lista.
- Socios: usar somente o primeiro item da lista.
- CNAEs secundarios: descartar nesta fase, pois nao ha coluna aprovada para eles.

## Tratamento de dados ausentes

Quando uma chave, lista ou item nao existir, a celula correspondente deve ficar vazia. O conversor nao deve falhar por listas vazias ou campos ausentes.

## Erros e validacao

O conversor deve:

- validar se o diretorio fonte existe;
- processar apenas arquivos `.json`;
- reportar arquivos JSON invalidos sem interromper todo o lote;
- gerar a planilha mesmo que alguns arquivos estejam invalidos;
- informar quantidade de arquivos processados, linhas geradas e arquivos com erro.

## Testes esperados

A implementacao deve ter validacao automatizada ou script de teste cobrindo:

- um JSON completo;
- um JSON sem telefones;
- um JSON sem e-mails;
- um JSON sem socios;
- um JSON com mais de um endereco;
- um arquivo JSON invalido.

## Proximas fases

Depois que a conversao local estiver validada:

1. adicionar `.env` com `DIRECTD_TOKEN`;
2. implementar cliente DirectD;
3. buscar JSON por CNPJ na API real;
4. opcionalmente salvar/cachear o JSON recebido;
5. integrar a geracao de planilha ao fluxo de consulta real.
