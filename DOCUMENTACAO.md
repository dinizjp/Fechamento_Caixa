# Documentação — Relatório Consolidado (fechamento.py)

## Visão Geral

Aplicação Streamlit que conecta ao banco SQL Server, executa 4 queries, processa os dados e gera um arquivo Excel com 6 abas. O objetivo é consolidar o fechamento de caixa das lojas por período, cruzando vendas em dinheiro, transferências para tesouraria, depósitos e saídas.

---

## Dicionários de Configuração

### `id_empresa_mapping`

Traduz o **ID numérico da empresa** (como está no banco) para o **nome legível** da loja. Usado exclusivamente para montar as linhas da aba **Resultado**.

```python
id_empresa_mapping = {
    58: 'Araguaína II',
    66: 'Balsas II',
    55: 'Araguaína I',
    53: 'Imperatriz II',
    51: 'Imperatriz I',
    65: 'Araguaína IV',
    52: 'Imperatriz III',
    57: 'Araguaína III',
    50: 'Balsas I',
    56: 'Gurupi I',
    61: 'Colinas',
    60: 'Estreito',
    46: 'Formosa I',
    59: 'Guaraí'
}
```

Cada empresa aqui gera **uma linha** na aba Resultado. Os IDs numéricos são usados como critério nos SUMIFS para buscar valores nas abas FechamentoCaixa, Relatorio e Contas a Pagar.

---

### `mapping_dict_plan`

Categoriza os **nomes de conta** da view `view_Contas_a_Pagar` em dois grupos:

| Categoria | Significado |
|---|---|
| `"Saídas"` | Transferências reais para a tesouraria — entram no cálculo da aba Resultado |
| `"Outras saidas"` | Saídas que não devem ser contabilizadas no resumo (consumo de açaí, vales, etc.) |

Esse mapeamento gera a coluna **"De Para"** na aba **Contas a Pagar**, e o SUMIFS da aba Resultado filtra apenas as linhas onde "De Para" = `"Saídas"`.

---

## Funções Python

### `remove_currency(val)`

Limpa e converte valores monetários para `float`. Trata os seguintes casos:

- `None` / `NaN` → retorna `None`
- Já é `int` ou `float` → retorna como `float` direto
- String com `"R$"` → remove o símbolo
- String com vírgula como separador decimal (formato BR: `1.234,56`) → converte para `1234.56`
- Qualquer erro → retorna `None`

Usada nas colunas `Valor` de `df1` (Relatorio) e `df2` (Contas a Pagar).

---

### `query_to_df(cursor, sql, params)`

Executa uma query SQL parametrizada e retorna um `DataFrame` pandas.

- `cursor`: objeto de conexão ODBC
- `sql`: string com a query (usa `?` como placeholder)
- `params`: tupla com os valores dos parâmetros (normalmente `(start_date_str, end_date_str)`)

Usa `cursor.description` para extrair os nomes das colunas direto do banco.

---

### `format_worksheet_as_table(worksheet, df, table_name)`

Formata uma aba do Excel como tabela nativa do Excel (com filtros e estilo). Faz duas coisas:

1. **Ajusta a largura de cada coluna** automaticamente com base no maior valor presente nos dados ou no nome do cabeçalho (o que for maior), mais 2 caracteres de margem.
2. **Cria uma tabela Excel** (`add_table`) com o range correto e os cabeçalhos definidos.

O cálculo de largura usa `.astype(str).str.len().max()` — converte tudo para string e pega o comprimento máximo de forma vetorizada.

---

## Queries SQL

### Query 1 — Transferências (`df1` → aba Relatorio)

Busca todas as transferências financeiras do período na view `Pesquisa_Transferencias_Busca`, filtrando apenas contas que o **usuário 1** tem permissão de visualizar (tabela `Financeiro_Contas_Acessos`).

- Filtro de data: campo `emissao` entre `start_date` e `end_date`
- Ordenação: por `emissao` decrescente

---

### Query 2 — Contas a Pagar (`df2` → aba Contas a Pagar)

Busca lançamentos de contas a pagar da view `view_Contas_a_Pagar`.

- `ID_Situacao IN (0,1)`: apenas pendentes e pagos
- Filtro de empresa: apenas as lojas do `id_empresa_mapping` (mais ID 64, que não está no mapping mas está nas queries)
- Filtro de data: campo `emissao`
- Colunas retornadas: ID_Empresa, Plano de Contas, Conta, Centro Custo, emissao, pagamento, Descrição Lançamento, Valor

---

### Query 3 — Fechamento de Caixa (`df3` → aba FechamentoCaixa)

A query mais complexa. Busca o resumo do fechamento de caixa por empresa/caixa/data, cruzando três tabelas:

- `View_FechamentoCaixa_Resumo`: dados principais do caixa
- `Pesquisa_Fechamento_Caixas`: informações de abertura/fechamento
- `Fechamento_Caixa_Conferencia_Sangrias`: apuração do gerente (subquery)
- `Financeiro_Transferencias`: valor transferido para tesouraria (subquery)

Colunas retornadas e o que significam:

| Coluna | Significado |
|---|---|
| `ID_Empresa` | ID numérico da loja |
| `ID_Caixa` | ID do caixa |
| `Data_Abertura_Str` | Data de abertura em texto (YYYY-MM-DD) |
| `Data_Abertura` | Data/hora de abertura |
| `Data_fechamento` | Data/hora de fechamento |
| `Usuário` | Operador do caixa |
| `Suprimento` | Valor de crédito lançado no caixa (fundo de troco inicial) |
| `Vendas_dinheiro` | Total de vendas pagas em dinheiro |
| `Total_Ent_Dinh` | Total de entradas em dinheiro (inclui suprimento + vendas) |
| `Transf_Tesour` | Valor transferido para a tesouraria naquele caixa |
| `Ap_Ger_Nao_Trans` | Apurado pelo gerente que **não** foi transferido (ficou no caixa) |
| `Apur_Ger_total` | Total apurado pelo gerente nas sangrias |
| `SaldoFinal` | Apurado pelo gerente - Total entradas dinheiro (diferença de caixa) |
| `Vale` | `'Vale'` se o saldo final for ≤ -R$3,00; caso contrário `'Nao'` |

Filtros:
- Apenas lojas do `id_empresa_mapping` (+ 64)
- `ID_Origem_Caixa = 1` (apenas caixas de loja, não caixas internos)
- Data de abertura entre `start_date` e `end_date`

---

### Query 4 — Vendas Trocadas (`df4` → aba vendas trocadas)

Busca a apuração detalhada de conferência de caixa da view `Pesquisa_Resumo_Conferencia_Apuracao`, cruzada com `Fechamento_Caixas`.

Retorna todos os campos da view mais `DataAbertura` e `DataFechamento`. Inclui colunas como `Apurado_Sistema`, `Apurado_Operador`, `Apurado_Gerente`, `Diferenca_Operador`, `Diferenca_Gerente` — usadas para identificar divergências de conferência entre o que o sistema registrou e o que foi contado fisicamente.

Filtro de data: `DataFechamento` entre `start_date` e `end_date`.

---

## Tratamento dos Dados (pós-query)

1. **Remoção de horário**: colunas `datetime` em `df1`, `df2`, `df3` são convertidas para `date` (apenas a data, sem horário).
2. **Limpeza de valor**: coluna `Valor` de `df1` e `df2` passa pela função `remove_currency`.
3. **De Para em df2**: coluna `Conta` é usada para criar a coluna `De Para` via `mapping_dict_plan`.
4. **df4 numérico**: colunas de apuração são convertidas para `float` via `pd.to_numeric(..., errors='coerce')` para evitar erros com valores nulos.
5. **Coluna de fórmula em df3**: antes de exportar, é adicionada a coluna `conferência fundo de troco` com fórmulas Excel (`=IF(G-K=0,0,G-K-L)`), que calcula a diferença entre suprimento, transferência e apurado não transferido.

---

## Abas do Excel

### Aba 1 — Resultado

Aba construída manualmente com `xlsxwriter` (não usa `df.to_excel`). Tem uma linha por empresa do `id_empresa_mapping`.

| Coluna | Letra | Fórmula / Valor |
|---|---|---|
| ID Empresa | A | Valor fixo (ID numérico) |
| Nome empresa | B | Valor fixo (nome da loja) |
| Data emissão | C | Data inicial do filtro (dd/mm/yyyy) |
| Vendas em dinheiro | D | `=SUMIFS(FechamentoCaixa!H:H, FechamentoCaixa!A:A, A{r}, FechamentoCaixa!D:D, C{r})` |
| Valor transferência | E | `=SUMIFS(FechamentoCaixa!J:J, FechamentoCaixa!A:A, A{r}, FechamentoCaixa!D:D, C{r})` |
| Diferença | F | `=D{r}-E{r}` |
| Depósitos | G | `=SUMIFS(Relatorio!H:H, Relatorio!A:A, A{r}, Relatorio!J:J, C{r}, Relatorio!D:D, "FINANCEIRO PARA FINANCEIRO")` |
| Saídas | H | `=SUMIFS('Contas a Pagar'!H:H, 'Contas a Pagar'!A:A, A{r}, 'Contas a Pagar'!E:E, C{r}, 'Contas a Pagar'!I:I, "Saídas")` |
| Transf x Deposito | I | `=E{r} - G{r}` |
| Falta depositar | J | `=E{r} - H{r} - G{r}` |

**Explicação das fórmulas:**

- **Vendas em dinheiro**: soma a coluna `Vendas_dinheiro` (col H) da aba FechamentoCaixa filtrando por `ID_Empresa` (col A) e `Data_Abertura_Str` (col D)
- **Valor transferência**: soma a coluna `Transf_Tesour` (col J) com os mesmos filtros
- **Diferença**: quanto das vendas em dinheiro não foi transferido para a tesouraria
- **Depósitos**: soma os valores da aba Relatorio filtrando por empresa, data e tipo "FINANCEIRO PARA FINANCEIRO" (depósitos bancários efetivos)
- **Saídas**: soma os valores da aba Contas a Pagar filtrando por empresa, data de emissão (col E) e categoria "Saídas" (col I = De Para)
- **Transf x Deposito**: diferença entre o que foi transferido para a tesouraria e o que foi depositado no banco
- **Falta depositar**: quanto ainda falta ser depositado (transferência - saídas - depósitos)

---

### Aba 2 — Relatorio

Dados brutos da view `Pesquisa_Transferencias_Busca`. Contém todas as transferências financeiras do período entre contas autorizadas. É a fonte dos **Depósitos** na aba Resultado (filtra por tipo "FINANCEIRO PARA FINANCEIRO").

---

### Aba 3 — Contas a Pagar

Dados da view `view_Contas_a_Pagar` com uma coluna extra **"De Para"** que categoriza cada lançamento em `"Saídas"` ou `"Outras saidas"` com base no nome da conta (`mapping_dict_plan`). É a fonte das **Saídas** na aba Resultado.

---

### Aba 4 — FechamentoCaixa

Dados do fechamento de caixa por empresa/caixa/dia. É a fonte de **Vendas em dinheiro** e **Valor transferência** na aba Resultado. Contém também a coluna calculada `conferência fundo de troco`:

```
=IF(G-K=0, 0, G-K-L)
```
Onde G = Suprimento, K = Transf_Tesour, L = Ap_Ger_Nao_Trans. Verifica se o fundo de troco está correto.

---

### Aba 5 — vendas trocadas

Apuração detalhada de conferência de caixa. Mostra as divergências entre o que o sistema registrou (`Apurado_Sistema`), o que o operador contou (`Apurado_Operador`) e o que o gerente conferiu (`Apurado_Gerente`), com as respectivas diferenças. Usada para identificar caixas com problemas de conferência.

---

### Aba 6 — De para

Tabela auxiliar de referência que lista todos os mapeamentos do `mapping_dict_plan`: nome da conta → categoria. Serve como dicionário de consulta dentro do próprio Excel para entender de onde vêm as categorizações da aba Contas a Pagar.
