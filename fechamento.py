import streamlit as st
import pandas as pd
import pyodbc
import io
from datetime import datetime
from xlsxwriter.utility import xl_col_to_name  # Para converter índice de coluna para notação (A, B, ...)

# Configurar a página
st.set_page_config(page_title='Relatório Consolidado', layout='wide')

# Recuperar as credenciais dos segredos (.streamlit/secrets.toml)
server   = st.secrets["mssql"]["server"]
database = st.secrets["mssql"]["database"]
username = st.secrets["mssql"]["username"]
password = st.secrets["mssql"]["password"]

if not all([server, database, username, password]):
    st.error("Verifique se todas as variáveis estão definidas corretamente no secrets.toml.")
    st.stop()
else:
    st.write(f"Servidor: {server}")
    st.write(f"Banco de Dados: {database}")
    st.write(f"Usuário: {username}")

st.title("Relatório Consolidado")

# Seleção do período
start_date = st.date_input("Data de Início", value=datetime(2025, 1, 1))
end_date   = st.date_input("Data de Fim", value=datetime.now())
if start_date > end_date:
    st.error("A data de início não pode ser maior que a data de fim.")
    st.stop()
start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str   = end_date.strftime('%Y-%m-%d')
st.write(f"Período selecionado: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")

# String de conexão (usa ODBC Driver 17 para SQL Server)
connection_string = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
)

# Função para remover "R$" e converter para float
def remove_currency(val):
    try:
        if pd.isnull(val):
            return None
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip()
        if "R$" in s:
            s = s.replace("R$", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return None

# Mapeamentos
mapping_dict_plan = {
    "-": "Outras saidas",
    "66 - CAIXA TESOURARIA | FORMOSA": "Saídas",
    # ... demais itens ...
}
id_empresa_mapping = {
    58: 'Araguaína II', 66: 'Balsas II', 55: 'Araguaína I',
    # ... demais itens ...
}

if st.button("Gerar Relatório Consolidado"):
    try:
        st.info("Estabelecendo conexão com o banco de dados...")
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        st.success("Conexão estabelecida!")
        
        # Query 1
        sql_query1 = f"""
        SELECT *
        FROM Pesquisa_Transferencias_Busca
        WHERE (
              [ID Conta Origem] IN (
                  SELECT ID_Conta FROM Financeiro_Contas_Acessos
                  WHERE ID_Usuario = 1 AND ISNULL(Visualizar, 'N') = 'S')
              OR
              [ID Conta Destino] IN (
                  SELECT ID_Conta FROM Financeiro_Contas_Acessos
                  WHERE ID_Usuario = 1 AND ISNULL(Visualizar, 'N') = 'S')
        )
        AND CONVERT(DATE, emissao) >= '{start_date_str}'
        AND CONVERT(DATE, emissao) <= '{end_date_str}'
        ORDER BY emissao DESC;"""
        st.info("Executando Query 1...")
        cursor.execute(sql_query1)
        df1 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        st.success("Query 1 executada!")
        
        # Query 2
        sql_query2 = f"""
        SELECT ID_Empresa, [Plano de Contas], Conta, [Centro Custo], emissao, pagamento, [Descrição Lançamento], Valor
        FROM view_Contas_a_Pagar
        WHERE ID_Situacao IN (0,1)
          AND CONVERT(DATE, emissao) BETWEEN '{start_date_str}' AND '{end_date_str}'
          AND ID_Empresa IN ({','.join(map(str, id_empresa_mapping.keys()))});"""
        st.info("Executando Query 2...")
        cursor.execute(sql_query2)
        df2 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        df2["De Para"] = df2["Conta"].str.strip().map(mapping_dict_plan).fillna("")
        df2['Valor'] = df2['Valor'].apply(remove_currency)
        st.success("Query 2 executada!")
        
        # Query 3 corrigida (subqueries agora usam SUM)
        sql_query3 = f"""
        SELECT 
            r.ID_Empresa,
            r.ID_Caixa,
            SUBSTRING(CONVERT(varchar, r.[Data Abertura], 120), 1, 10) AS Data_Abertura_Str,
            r.[Data Abertura],
            r.[Data fechamento],
            r.[Usuário],
            CONVERT(float, r.Lancamento_Credito) AS Suprimento,
            CONVERT(float, r.Vendas_dinheiro) AS Vendas_dinheiro,
            CONVERT(float, r.Total_Entradas_Dinheiro) AS Total_Ent_Dinh,
            CONVERT(float, ISNULL((SELECT SUM(ISNULL(t.valor,0)) FROM Financeiro_Transferencias t
                                    WHERE t.ID_Empresa = r.ID_Empresa AND t.ID_Caixa = r.ID_Caixa),0)) AS Transf_Tesour,
            CONVERT(float, ISNULL(((SELECT SUM(fg.apurado_gerente) FROM Fechamento_Caixa_Conferencia_Sangrias fg
                                    WHERE fg.ID_Empresa = r.ID_Empresa AND fg.ID_Caixa = r.ID_Caixa)
                                  - (SELECT SUM(ISNULL(t2.valor,0)) FROM Financeiro_Transferencias t2
                                     WHERE t2.ID_Empresa = r.ID_Empresa AND t2.ID_Caixa = r.ID_Caixa)),0)) AS Ap_Ger_Nao_Trans,
            CONVERT(float,(SELECT SUM(fg2.apurado_gerente) FROM Fechamento_Caixa_Conferencia_Sangrias fg2
                            WHERE fg2.ID_Empresa = r.ID_Empresa AND fg2.ID_Caixa = r.ID_Caixa)) AS Apur_Ger_total,
            CONVERT(float,((SELECT SUM(fg3.apurado_gerente) FROM Fechamento_Caixa_Conferencia_Sangrias fg3
                             WHERE fg3.ID_Empresa = r.ID_Empresa AND fg3.ID_Caixa = r.ID_Caixa)
                           - r.Total_Entradas_Dinheiro)) AS SaldoFinal,
            CASE WHEN ((SELECT ISNULL(SUM(fg4.apurado_gerente),0) FROM Fechamento_Caixa_Conferencia_Sangrias fg4
                         WHERE fg4.ID_Empresa = r.ID_Empresa AND fg4.ID_Caixa = r.ID_Caixa)
                       - r.Total_Entradas_Dinheiro) <= -3 THEN 'Vale' ELSE 'Nao' END AS Vale
        FROM View_FechamentoCaixa_Resumo r
        INNER JOIN Pesquisa_Fechamento_Caixas c
          ON r.ID_Caixa = c.[ID Caixa] AND r.ID_Empresa = c.[ID Empresa]
        WHERE r.ID_Empresa IN ({','.join(map(str, id_empresa_mapping.keys()))})
          AND CONVERT(date, r.[Data Abertura]) BETWEEN '{start_date_str}' AND '{end_date_str}'
          AND c.[ID Origem Caixa] = 1
        ORDER BY r.ID_Empresa, r.ID_Caixa, r.[Data Abertura];"""
        st.info("Executando Query 3...")
        cursor.execute(sql_query3)
        df3 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        st.success("Query 3 executada!")
        
        # Query 4 (vendas trocadas)
        sql_query4 = f"""
        SELECT pr.*, fc.DataAbertura, fc.DataFechamento
        FROM Pesquisa_Resumo_Conferencia_Apuracao pr
        INNER JOIN Fechamento_Caixas fc
          ON fc.ID_Caixa = pr.ID_Caixa AND fc.ID_Empresa = pr.ID_Empresa AND fc.ID_Origem_Caixa = pr.ID_Origem_Caixa
        WHERE fc.ID_Empresa IN ({','.join(map(str, id_empresa_mapping.keys()))})
          AND fc.ID_Origem_Caixa = 1
          AND fc.DataFechamento BETWEEN '{start_date_str}' AND '{end_date_str}'
        ORDER BY fc.ID_Caixa;"""
        st.info("Executando Query 4...")
        cursor.execute(sql_query4)
        df4 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        st.success("Query 4 executada!")
        
        # Tratamento e geração do Excel (idêntico ao anterior)
        # ... (restante do código permanece igual) ...

    except pyodbc.Error as e:
        st.error(f"Erro relacionado ao ODBC: {e}")
    except Exception as ex:
        st.error(f"Erro inesperado: {ex}")
    finally:
        if 'conn' in locals():
            conn.close()
            st.info("Conexão fechada.")
