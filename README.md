# Fechamento_Caixa

## Descrição do Projeto

O projeto **Fechamento_Caixa** consiste em uma aplicação que gera relatórios consolidados de fechamento de caixa a partir de uma base de dados SQL Server. Utilizando Streamlit, a aplicação oferece uma interface amigável para selecionar o período de análise, conectar-se ao banco de dados, executar consultas específicas e apresentar os resultados em forma de tabelas. As consultas envolvem dados de transferências, contas a pagar, fechamento de caixa e outras operações financeiras, permitindo ao usuário obter uma visão consolidada e detalhada do fluxo financeiro em diferentes empresas e períodos.

## Sumário

- [Dependências](#dependências)
- [Instalação](#instalação)
- [Uso](#uso)
- [Estrutura de Pastas](#estrutura-de-pastas)

## Dependências

As bibliotecas necessárias para executar o projeto estão listadas no arquivo `requirements.txt`. São elas:

- `streamlit`  
- `pandas`
- `pyodbc`
- `python-dotenv`
- `xlsxwriter`

Além disso, recomenda-se a instalação do driver ODBC para conexão com SQL Server, listado em `packages.txt`:  
- `msodbcsql17`

## Instalação

Para instalar as dependências do projeto, execute o seguinte comando na raiz do seu ambiente Python:

```bash
pip install -r requirements.txt
```

Certifique-se de que o driver ODBC (`msodbcsql17`) também esteja instalado conforme orientações do arquivo `packages.txt` para conexão com banco de dados.

## Uso

1. Configure as credenciais de conexão ao banco de dados no arquivo `.streamlit/secrets.toml`, com o seguinte formato:

```toml
[mssql]
server = "SEU_SERVIDOR"
database = "SEU_BANCO_DE_DADOS"
username = "SEU_USUARIO"
password = "SUA_SENHA"
```

2. Execute a aplicação com o comando:

```bash
streamlit run fechamento.py
```

3. Acesse a interface no navegador, geralmente disponível em `http://localhost:8501`.

4. Selecione o período desejado e clique em "Gerar Relatório Consolidado" para obter os dados carregados a partir das consultas ao banco.

## Estrutura de Pastas

A estrutura do projeto é a seguinte:

```
Fechamento_Caixa/
├── fechamento.py
├── packages.txt
└── requirements.txt
```

### Descrição dos Arquivos

- **fechamento.py**: Script principal que contém a lógica da aplicação Streamlit, conexão com banco, consultas SQL e geração de relatórios visuais.
- **packages.txt**: Lista de pacotes adicionais necessários para instalação do driver ODBC.
- **requirements.txt**: Lista de bibliotecas Python necessárias para o funcionamento do projeto.

---

## Observações finais

Este projeto foi elaborado para facilitar a análise de fechamento de caixa e movimentações financeiras de diversas empresas. A interface permite uma interação dinâmica com os dados extraídos do banco de dados SQL Server, proporcionando uma visão consolidada e detalhada.