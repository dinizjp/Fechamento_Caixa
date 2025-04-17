```markdown
# Fechamento_Caixa

## Objetivo do Projeto

O projeto **Fechamento_Caixa** tem como objetivo principal gerar relatórios consolidados de fechamento de caixa a partir de uma base de dados. A aplicação oferece uma interface amigável, utilizando Streamlit, permitindo ao usuário interagir com os dados e realizar consultas de acordo com um intervalo de datas específico. O sistema conecta-se a um banco de dados MSSQL para extrair, processar e apresentar informações relevantes sobre as movimentações financeiras.

## Dependências

O projeto requer várias dependências para funcionar corretamente. Você pode instalá-las utilizando o arquivo `requirements.txt` que contém todas as bibliotecas necessárias.

### Dependências:

- `streamlit`: Para criar a interface web da aplicação.
- `pandas`: Para manipulação e análise de dados.
- `pyodbc`: Para conexão com o banco de dados SQL Server.
- `python-dotenv`: Para carregar variáveis de ambiente de um arquivo `.env`.
- `xlsxwriter`: Para exportação de dados para arquivos Excel.

### Como instalar as dependências

1. Certifique-se de ter o `pip` instalado em seu ambiente Python.
2. Execute o seguinte comando na raiz do seu projeto:
   ```bash
   pip install -r requirements.txt
   ```

## Como Executar e Testar

Para executar a aplicação, utilize o Streamlit. Siga os passos abaixo:

1. **Configure suas credenciais** no arquivo `.streamlit/secrets.toml` com as seguintes variáveis:
   ```toml
   [mssql]
   server = "SEU_SERVIDOR"
   database = "SEU_BANCO_DE_DADOS"
   username = "SEU_USUARIO"
   password = "SUA_SENHA"
   ```

2. **Execute a aplicação** com o seguinte comando:
   ```bash
   streamlit run fechamento.py
   ```

3. **Acesse a interface** no navegador, normalmente disponível em `http://localhost:8501`.

## Estrutura de Pastas

O projeto possui a seguinte estrutura de arquivos:

```
Fechamento_Caixa/
├── fechamento.py
├── packages.txt
└── requirements.txt
```

### Descrição dos arquivos principais

- **fechamento.py**: Este é o arquivo principal onde a aplicação Streamlit está implementada. Ele contém a lógica para a conexão ao banco de dados, execução de queries SQL, e geração dos relatórios visuais.

- **packages.txt**: Este arquivo lista as dependências de pacotes que podem ser necessárias para instalação de drivers, como `msodbcsql17` para conexão com o SQL Server.

- **requirements.txt**: Este arquivo contém uma lista de todas as bibliotecas Python necessárias para o projeto.

---

Agradecemos por utilizar o **Fechamento_Caixa**! Se tiver dúvidas ou sugestões, fique à vontade para contribuir.
```
