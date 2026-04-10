# Fechamento_Caixa

Aplicação Streamlit que conecta ao banco SQL Server, executa consultas de fechamento de caixa e gera um relatório Excel consolidado por período.

## Dependências

```
pip install -r requirements.txt
```

## Configuração

Crie o arquivo `.streamlit/secrets.toml` com as credenciais do banco:

```toml
[mssql]
server="..."
database="..."
username="..."
password="..."
```

## Como usar

```sh
streamlit run fechamento.py
```

Selecione o período desejado e clique em **"Gerar Relatório Consolidado"**. O Excel será disponibilizado para download com as seguintes abas:

| Aba | Conteúdo |
|---|---|
| Resultado | Resumo por empresa e por dia: vendas, transferências, depósitos, parquinho, saídas |
| Relatorio | Transferências financeiras do período (`Pesquisa_Transferencias_Busca`) |
| Contas a Pagar | Lançamentos a pagar com categorização "De Para" |
| FechamentoCaixa | Resumo de fechamento por caixa/dia com conferência de fundo de troco |
| vendas trocadas | Apuração detalhada de conferência (sistema × operador × gerente) |
| De para | Tabela de referência para categorização das contas |

## Adicionar nova loja

1. Incluir o ID e nome em `id_empresa_mapping` no `fechamento.py`
2. Adicionar o ID na cláusula `IN (...)` das 4 queries SQL
