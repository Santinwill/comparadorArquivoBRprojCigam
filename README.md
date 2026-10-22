# Comparador Base Fiscal x CIGAM

Aplicação em Python para relacionar e comparar os registros da **Base Fiscal gerada pelo BR Proj** com os lançamentos de entrada do **CIGAM**.

O processamento gera um relatório Excel com:

- divergências fiscais confirmadas;
- diferenças pendentes de validação fiscal;
- itens sem associação segura;
- notas não encontradas;
- itens da Base Fiscal que não foram utilizados;
- rastreabilidade das linhas originais da Base Fiscal e do CIGAM.

> **Situação atual das regras fiscais:** a interpretação dos campos configurados como `REVIEW` permanece provisória. A classificação de documentos de serviço e os campos que podem ter granularidades diferentes entre Base Fiscal e CIGAM aguardam validação da área fiscal.

## Estrutura final do projeto

```text
comparadorArquivoBRprojCigam/
├── input/
│   ├── Base_Fiscal.xlsx
│   └── NF_FISCAL_ENTRADA.xlsx
│
├── output/
│   └── Fiscal_Comparison.xlsx
│
├── src/
│   ├── comparator.py
│   ├── comparison_rules.py
│   ├── excel_exporter.py
│   ├── matcher.py
│   └── normalizer.py
│
├── .gitignore
├── main.py
├── README.md
└── requirements.txt
```

### Arquivos versionados no Git

Na pasta `input/`, somente os arquivos de exemplo devem ser versionados:

```text
input/Base_Fiscal_exemplo.xlsx
input/NF_FISCAL_ENTRADA_exemplo.xlsx
```

Os arquivos reais ficam apenas no ambiente local:

```text
input/Base_Fiscal.xlsx
input/NF_FISCAL_ENTRADA.xlsx
```

A pasta `output/` é preservada no Git por meio de `output/.gitkeep`, mas o relatório gerado não deve ser versionado.

## Requisitos

- Python 3;
- `openpyxl`.

Instale as dependências com:

```powershell
py -m pip install -r requirements.txt
```

O `requirements.txt` deve conter:

```text
openpyxl
```

## Arquivos de entrada

Antes da execução, copie os arquivos reais para a pasta `input/` usando exatamente estes nomes:

```text
input/Base_Fiscal.xlsx
input/NF_FISCAL_ENTRADA.xlsx
```

Os dois arquivos devem possuir uma aba chamada:

```text
Export
```

Feche os arquivos de entrada e o relatório anterior no Excel antes da execução. Caso estejam abertos, o Windows poderá bloquear o acesso e gerar `PermissionError`.

## Execução

O ponto oficial de execução é o arquivo `main.py`.

Na raiz do projeto, execute:

```powershell
py main.py
```

Não é necessário executar os módulos internos individualmente.

Para executar especificamente o exportador como módulo, use:

```powershell
py -m src.excel_exporter
```

## Fluxo da aplicação

```text
main.py
  ↓
src.matcher
  ↓
src.comparator
  ↓
src.excel_exporter
  ↓
output/Fiscal_Comparison.xlsx
```

O processo completo:

1. carrega os arquivos de entrada;
2. normaliza os identificadores usados no matching;
3. agrupa os registros por nota;
4. realiza o matching um-para-um dos itens;
5. compara os campos fiscais;
6. separa divergências confirmadas e revisão pendente;
7. gera o relatório Excel final.

## Módulos principais

### `main.py`

Ponto único de execução. Trata mensagens de sucesso e erros comuns, como:

- arquivo de entrada não encontrado;
- arquivo aberto ou bloqueado;
- erro de validação;
- erro inesperado.

O processo retorna:

```text
0 = execução concluída com sucesso
1 = execução encerrada com erro
```

### `src/normalizer.py`

Normaliza diferenças estruturais usadas na identificação dos registros:

- unidade de negócio;
- número da nota;
- CNPJ/CPF;
- data de emissão, desconsiderando o horário;
- classificação fiscal/NCM;
- quantidade usada na chave de matching.

A normalização não deve esconder divergências fiscais reais.

### `src/matcher.py`

Executa o matching um-para-um. Cada linha da Base Fiscal e cada linha do CIGAM pode ser utilizada no máximo uma vez.

### `src/comparison_rules.py`

Centraliza os campos classificados como:

- `COMPARE`: comparação automática;
- `REVIEW`: revisão fiscal pendente;
- `IGNORE`: campo explicitamente ignorado.

### `src/comparator.py`

Compara os pares associados, preserva os valores originais e separa:

- divergências confirmadas;
- diferenças em revisão;
- pares sem diferenças.

### `src/excel_exporter.py`

Cria e formata o relatório final em:

```text
output/Fiscal_Comparison.xlsx
```

## Identificação da nota

Uma nota é identificada pela combinação:

```text
CD_UNI_NEG
+ NF
+ CNPJ_CPF
+ DATA_EMISSAO
```

O horário de `DATA_EMISSAO` é desconsiderado durante o matching.

Quando nenhuma nota possui a mesma combinação completa, o registro recebe:

```text
NOTA_NAO_ENCONTRADA
```

## Matching dos itens

O matching é executado dentro de cada nota. Cada linha da Base Fiscal e cada linha do CIGAM pode ser utilizada no máximo uma vez.

### `MATCH_EXATO`

O NCM normalizado e a quantidade identificam um único registro em cada origem.

```text
NCM normalizado + QTD_ENTRADA
```

### `MATCH_DESEMPATADO_POR_VALOR`

Quando NCM e quantidade possuem mais de um candidato, o valor é usado como critério auxiliar de desempate.

```text
NCM normalizado + QTD_ENTRADA + VLR_ENTRADA
```

O valor continua sendo auditado posteriormente pelo comparador.

### `MATCH_DUPLICADO_SEQUENCIAL`

Existem registros repetidos com o mesmo NCM, quantidade e valor. A associação ocorre sequencialmente, sem reutilizar linhas.

### `MATCH_SEM_NCM`

Quantidade e valor identificam um único par sem utilizar o NCM:

```text
QTD_ENTRADA + VLR_ENTRADA
```

O NCM continua sendo auditado e pode ser marcado como divergente.

### `MATCH_DUPLICADO_SEM_NCM`

Existem múltiplos registros com a mesma quantidade e valor. A associação ocorre sequencialmente sem utilizar o NCM.

### `MATCH_POR_ELIMINACAO`

Depois das demais etapas, resta exatamente um item da Base Fiscal e um item do CIGAM na mesma nota.

Esse matching possui confiança baixa. No relatório:

- a linha recebe fundo laranja;
- divergências confirmadas permanecem vermelhas;
- campos em revisão permanecem amarelos.

### `MATCH_NAO_RESOLVIDO`

A nota foi localizada, mas o item não apresentou associação segura pelas regras disponíveis.

O projeto não cria automaticamente fatores de conversão de unidade ou quantidade sem uma regra de negócio validada.

## Campos em comparação automática

Atualmente estão configurados como `COMPARE`:

```text
CD_CLASS_FISCAL
QTD_ENTRADA
VLR_ENTRADA
VLR_FRETE
VLR_DESCONTO
BASE_ICMS
VLR_ICMS
BASE_PIS_COFINS
VLR_PIS
VLR_COFINS
VLR_IPI
VLR_PIS_RET
VLR_COFINS_RET
VLR_CSLL_RET
VLR_IRRF_RET
VLR_ISS_RET
VLR_INSS_RET
VLR_DIF_ALIQUOTA
```

Uma diferença nesses campos classifica o par como:

```text
DIVERGENTE
```

## Campos em revisão fiscal

Atualmente estão configurados como `REVIEW`:

```text
CTS_ICMS
ICMS_NAO_TRI
ICMS_OUTROS
ICMS_NAO_DEST
CTS_PIS
CTS_COFINS
PIS_NAO_DEST
COFINS_NAO_DEST
CTS_IPI
IPI_NAO_TRI
IPI_OUTROS
```

Esses campos aparecem em amarelo no relatório, mas não são classificados automaticamente como erro fiscal.

Quando um par possui somente diferenças desse grupo, recebe:

```text
REVISAO_PENDENTE
```

## Comparação numérica

A comparação numérica:

- não utiliza `round()`;
- não utiliza tolerância;
- não utiliza `math.isclose()`;
- não considera valores aproximados como iguais;
- preserva diferenças de precisão para auditoria.

Exemplo de diferença preservada:

```text
300.02999999999997 != 300.03
```

Valores equivalentes apenas por tipo Python são tratados como iguais:

```text
300 == 300.0
```

## Abas do relatório

### `Resumo`

Apresenta:

- totais de registros;
- pares associados;
- divergências confirmadas;
- registros somente para revisão;
- matches não resolvidos;
- notas não encontradas;
- itens da Base não utilizados;
- contagem por campo;
- legenda de cores.

### `Divergencias`

Pares com pelo menos uma divergência em campo configurado como `COMPARE`.

### `Revisao`

Pares sem divergência confirmada, mas com diferenças nos campos configurados como `REVIEW`.

### `Nao_Resolvidos`

Registros do CIGAM cuja nota foi localizada, mas cujo item não teve associação segura.

### `Notas_Nao_Encontradas`

Registros do CIGAM sem nota correspondente pela chave completa.

### `Base_Nao_Utilizada`

Itens da Base Fiscal que não foram utilizados em nenhuma associação.

## Coluna `MOTIVO_STATUS`

As abas de exceção possuem uma explicação operacional.

### `MATCH_NAO_RESOLVIDO`

```text
A nota foi localizada na Base Fiscal, mas o item não apresentou
uma associação segura pelas regras de matching disponíveis.
```

### `NOTA_NAO_ENCONTRADA`

```text
Não foi localizada na Base Fiscal uma nota com a mesma combinação
de CD_UNI_NEG, NF, CNPJ_CPF e DATA_EMISSAO.
```

### `ITEM_BASE_NAO_UTILIZADO`

```text
O item existe na Base Fiscal, mas não foi utilizado em nenhuma
associação com os registros do CIGAM.
```

## Legenda de cores

```text
Vermelho  → divergência confirmada
Amarelo   → diferença pendente de revisão fiscal
Laranja   → match por eliminação com confiança baixa
Verde     → registro sem divergência
Azul      → cabeçalhos e títulos
```

## Formatação do Excel

O relatório aplica:

- datas no formato `dd/mm/yyyy`;
- CNPJ, NF, unidade de negócio, chaves e códigos fiscais como texto;
- preservação de zeros à esquerda;
- inteiros sem separador decimal;
- decimais com quantidade variável de casas;
- filtros em todas as abas de dados;
- congelamento dos cabeçalhos;
- largura automática das colunas;
- destaque integral dos registros `MATCH_POR_ELIMINACAO`.

A formatação não altera o resultado da comparação.

## Resultado da execução utilizada na validação

```text
Registros da Base Fiscal             1770
Registros do CIGAM                    605
Pares associados                      447
Divergências confirmadas              254
Somente revisão                       193
Sem divergências                        0
Matches não resolvidos                  7
Notas não encontradas                 151
Itens da Base não utilizados           14
```

A classificação dos registros do CIGAM fecha em:

```text
447 pares associados
+ 7 matches não resolvidos
+ 151 notas não encontradas
= 605 registros
```

Esses totais pertencem aos arquivos usados na validação e podem mudar em execuções futuras.

## Validação de sintaxe

Para validar todos os módulos principais:

```powershell
py -m py_compile main.py
py -m py_compile src\normalizer.py
py -m py_compile src\matcher.py
py -m py_compile src\comparison_rules.py
py -m py_compile src\comparator.py
py -m py_compile src\excel_exporter.py
```

Se os comandos não mostrarem mensagens, a sintaxe está correta.

## Problemas comuns

### `PermissionError`

Feche:

- `Base_Fiscal.xlsx`;
- `NF_FISCAL_ENTRADA.xlsx`;
- `Fiscal_Comparison.xlsx`.

Depois execute novamente:

```powershell
py main.py
```

### Arquivo de entrada não encontrado

Confirme se os arquivos reais existem com os nomes esperados:

```text
input/Base_Fiscal.xlsx
input/NF_FISCAL_ENTRADA.xlsx
```

### Erro de importação

Execute o projeto pela raiz:

```powershell
py main.py
```

Não execute diretamente:

```powershell
py src\excel_exporter.py
```

Para execução modular, use:

```powershell
py -m src.excel_exporter
```

## Limitações e pendências

- A interpretação dos campos em `REVIEW` ainda aguarda validação fiscal.
- A classificação fiscal de documentos de serviço aguarda regra da área fiscal.
- Não existe regra automática validada para fatores de conversão de unidade ou quantidade.
- A precisão visual do Excel está sujeita ao limite numérico do próprio Excel; a comparação acontece no Python antes da exportação.
- Novas regras devem ser cadastradas em `src/comparison_rules.py` e validadas antes de serem promovidas para `COMPARE`.
