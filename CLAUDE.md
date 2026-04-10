# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```sh
streamlit run fechamento.py
```

## Architecture

Single-file Streamlit app (`fechamento.py`). No test suite, no build step.

**Flow:**
1. User selects a date range in the UI
2. Clicks "Gerar Relatório Consolidado"
3. App connects to SQL Server via pyodbc using credentials from `.streamlit/secrets.toml`
4. Executes 4 SQL queries in sequence → 4 DataFrames
5. Displays DataFrames on screen
6. Generates a downloadable `.xlsx` with 6 tabs

## Key Data Structures

- **`id_empresa_mapping`** — maps numeric store IDs (as in the DB) to readable store names. Each entry becomes one row in the Excel "Resultado" tab. If a new store is added, add it here AND to the `IN (...)` clauses of all 4 SQL queries.
- **`mapping_dict_plan`** — maps account names from `view_Contas_a_Pagar` to either `"Saídas"` (counted in the summary) or `"Outras saidas"` (excluded). This populates the "De Para" column in the "Contas a Pagar" tab.

## Excel Output Structure

The Excel file has 6 tabs:

| Tab | Source | Notes |
|---|---|---|
| Resultado | Built manually with xlsxwriter | Contains SUMIFS formulas referencing other tabs |
| Relatorio | `df1` | Raw transfers from `Pesquisa_Transferencias_Busca` |
| Contas a Pagar | `df2` | From `view_Contas_a_Pagar`, adds "De Para" column |
| FechamentoCaixa | `df3` | From `View_FechamentoCaixa_Resumo`, adds `conferência fundo de troco` formula column |
| vendas trocadas | `df4` | From `Pesquisa_Resumo_Conferencia_Apuracao` |
| De para | `mapping_dict_plan` | Reference table for the Contas a Pagar categorization |

The "Resultado" tab is the summary — its SUMIFS pull from the other tabs by column letter, so **column order in each tab must not change**.

## DB Credentials

Stored in `.streamlit/secrets.toml` (not committed). Structure:

```toml
[mssql]
server="..."
database="..."
username="..."
password="..."
```

Connection uses ODBC Driver 17 for SQL Server with `TrustServerCertificate=yes`.
