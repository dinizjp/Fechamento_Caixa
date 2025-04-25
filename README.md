# Fechamento_Caixa

## Descrição
O projeto **Fechamento_Caixa** é um sistema que realiza consultas e geração de relatórios relacionados ao fechamento de caixa, transferências, contas a pagar, e conferências de sangrias e vendas. Ele utiliza conexão com banco de dados SQL Server para extrair, processar e exibir informações relevantes por meio de uma interface web criada com Streamlit.

## Sumário
- [Dependências](#dependências)
- [Instalação](#instalação)
- [Uso](#uso)
- [Estrutura de Pastas](#estrutura-de-pastas)

## Dependências
As dependências do projeto estão listadas no arquivo `requirements.txt`:
- streamlit
- pandas
- pyodbc
- python-dotenv
- xlsxwriter

## Instalação
Para configurar o ambiente, execute os seguintes comandos:
```sh
pip install -r requirements.txt
```

## Uso
Para iniciar a aplicação, execute:
```sh
streamlit run fechamento.py
```
Ao abrir a interface web, clique no botão **"Gerar Relatório Consolidado"** para estabelecer conexão com o banco de dados e executar as consultas que irão gerar os relatórios apresentados na aplicação.

## Estrutura de Pastas
```
Fechamento_Caixa/
├── fechamento.py
├── requirements.txt
└── .streamlit/
    └── secrets.toml
```