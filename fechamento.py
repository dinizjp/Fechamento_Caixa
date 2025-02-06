import streamlit as st
import pandas as pd
import pyodbc
import os
import io
from dotenv import load_dotenv
from datetime import datetime

# Configurar a página
st.set_page_config(page_title='Relatório Consolidado', layout='wide')

# Carregar variáveis do arquivo .env
load_dotenv()  # Se necessário, especifique o caminho para o .env

# Recuperar as credenciais do .env
server = st.secrets["db"]["DB_SERVER"]
database = st.secrets["db"]["DB_NAME"]
username = st.secrets["db"]["DB_USER"]
password = st.secrets["db"]["DB_PASSWORD"]


# Verificar se todas as variáveis de ambiente estão definidas
if not all([server, database, username, password]):
    st.error("Verifique se todas as variáveis de ambiente estão definidas corretamente no arquivo .env.")
    st.stop()
else:
    st.write(f"Servidor: {server}")
    st.write(f"Banco de Dados: {database}")
    st.write(f"Usuário: {username}")

st.title("Relatório Consolidado")

# Seleção do período com dois widgets separados
start_date = st.date_input("Data de Início", value=datetime(2025, 1, 1))
end_date = st.date_input("Data de Fim", value=datetime.now())

if start_date > end_date:
    st.error("A data de início não pode ser maior que a data de fim.")
    st.stop()

# Converter datas para o formato SQL (YYYY-MM-DD)
start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str = end_date.strftime('%Y-%m-%d')
st.write(f"Período selecionado: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")

# String de conexão
connection_string = (
    'DRIVER={ODBC Driver 18 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
    'Encrypt=yes;'
    'TrustServerCertificate=yes;'
)

if st.button("Gerar Relatório Consolidado"):
    try:
        st.info("Estabelecendo conexão com o banco de dados...")
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        st.success("Conexão estabelecida com sucesso!")
        
        ## --- QUERY 1: Pesquisa_Transferencias_Busca ---
        sql_query1 = f"""
        SELECT *
        FROM Pesquisa_Transferencias_Busca
        WHERE (
              [ID Conta Origem] IN (
                  SELECT ID_Conta 
                  FROM Financeiro_Contas_Acessos
                  WHERE ID_Usuario = 1 
                    AND ISNULL(Visualizar, 'N') = 'S'
              )
              OR
              [ID Conta Destino] IN (
                  SELECT ID_Conta 
                  FROM Financeiro_Contas_Acessos
                  WHERE ID_Usuario = 1 
                    AND ISNULL(Visualizar, 'N') = 'S'
              )
        )
        AND CONVERT(DATE, emissao) >= '{start_date_str}'
        AND CONVERT(DATE, emissao) <= '{end_date_str}'
        ORDER BY emissao DESC;
        """
        st.info("Executando Query 1...")
        cursor.execute(sql_query1)
        columns1 = [col[0] for col in cursor.description]
        rows1 = cursor.fetchall()
        data1 = [dict(zip(columns1, row)) for row in rows1]
        df1 = pd.DataFrame(data1)
        st.success("Query 1 executada com sucesso!")
        
        ## --- QUERY 2: view_Contas_a_Pagar (somente colunas selecionadas) ---
        sql_query2 = f"""
        SELECT 
             ID_Empresa,
             [Plano de Contas],
             [Centro Custo],
             emissao,
             pagamento,
             [Descrição Lançamento],
             Valor
        FROM view_Contas_a_Pagar
        WHERE ID_Situacao IN (0,1)
          AND CONVERT(DATE, emissao) >= '{start_date_str}'
          AND CONVERT(DATE, emissao) <= '{end_date_str}'
          AND ID_Empresa IN (55, 58, 57, 65, 50, 66, 64, 61, 60, 46, 59, 56, 51, 53, 52);
        """
        st.info("Executando Query 2...")
        cursor.execute(sql_query2)
        columns2 = [col[0] for col in cursor.description]
        rows2 = cursor.fetchall()
        data2 = [dict(zip(columns2, row)) for row in rows2]
        df2 = pd.DataFrame(data2)
        df2['Valor'] = df2['Valor'].astype(float)
        st.success("Query 2 executada com sucesso!")
        
        ## --- QUERY 3: Fechamento de Caixa ---
        sql_query3 = f"""
        SELECT 
            r.ID_Empresa,
            r.ID_Caixa,
            SUBSTRING(CONVERT(varchar, [Data Abertura], 120), 1, 10) AS Data_Abertura_Str,
            [Data Abertura],
            [Data fechamento],
            [Usuário],
            CONVERT(float, Lancamento_Credito) AS Suprimento,
            CONVERT(float, Vendas_dinheiro) AS Vendas_dinheiro, 
            CONVERT(float, Total_Entradas_Dinheiro) AS Total_Ent_Dinh,
            CONVERT(float, ISNULL(
                 (SELECT ISNULL(valor, 0.00)
                  FROM Financeiro_Transferencias t 
                  WHERE t.ID_Empresa = r.ID_Empresa 
                    AND t.ID_Caixa = r.ID_Caixa), 0.00)
                 ) AS Transf_Tesour, 
            CONVERT(float, ISNULL(
                 ((SELECT SUM(apurado_gerente)
                   FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                   WHERE FG.ID_Empresa = r.ID_Empresa 
                     AND FG.ID_Caixa = r.ID_Caixa)
                 - (SELECT ISNULL(valor, 0.00)
                    FROM Financeiro_Transferencias t 
                    WHERE t.ID_Empresa = r.ID_Empresa 
                      AND t.ID_Caixa = r.ID_Caixa)
                 ), 0.00)
                 ) AS Ap_Ger_Nao_Trans,
            CONVERT(float, (SELECT SUM(apurado_gerente)
                 FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                 WHERE FG.ID_Empresa = r.ID_Empresa 
                   AND FG.ID_Caixa = r.ID_Caixa)
                 ) AS Apur_Ger_total,
            CONVERT(float, 
                 ((SELECT SUM(apurado_gerente)
                   FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                   WHERE FG.ID_Empresa = r.ID_Empresa 
                     AND FG.ID_Caixa = r.ID_Caixa)
                 - Total_Entradas_Dinheiro)
                 ) AS SaldoFinal,
            CASE 
               WHEN ((SELECT ISNULL(SUM(apurado_gerente), 0.00)
                      FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                      WHERE FG.ID_Empresa = r.ID_Empresa 
                        AND FG.ID_Caixa = r.ID_Caixa)
                    - ISNULL(Total_Entradas_Dinheiro, 0.00)) <= -3.00 
               THEN 'Vale' 
               ELSE 'Nao' 
            END AS Vale
        FROM 
            View_FechamentoCaixa_Resumo r
        INNER JOIN 
            Pesquisa_Fechamento_Caixas c 
            ON r.ID_Caixa = [ID Caixa] 
           AND r.ID_Empresa = [ID Empresa]
        WHERE  
            r.ID_Empresa IN (55, 58, 57, 65, 50, 66, 64, 61, 60, 46, 59, 56, 51, 53, 52)  
            AND SUBSTRING(CONVERT(varchar, [Data Abertura], 120), 1, 10) >= '{start_date_str}'
            AND SUBSTRING(CONVERT(varchar, [Data Abertura], 120), 1, 10) <= '{end_date_str}'
            AND [ID_Origem_Caixa] = 1 
        ORDER BY 
            r.ID_Empresa, r.ID_Caixa, SUBSTRING(CONVERT(varchar, [Data Abertura], 120), 1, 10);
        """
        st.info("Executando Query 3...")
        cursor.execute(sql_query3)
        columns3 = [col[0] for col in cursor.description]
        rows3 = cursor.fetchall()
        data3 = [dict(zip(columns3, row)) for row in rows3]
        df3 = pd.DataFrame(data3)
        st.success("Query 3 executada com sucesso!")
        
        # Exibir os DataFrames (opcional)
        st.subheader("Pesquisa_Transferencias_Busca")
        st.dataframe(df1)
        
        st.subheader("view_Contas_a_Pagar")
        st.dataframe(df2)
        
        st.subheader("Fechamento de Caixa")
        st.dataframe(df3)
        
        ## --- GERAR ARQUIVO EXCEL COM 3 ABAS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df1.to_excel(writer, index=False, sheet_name='Relatorio')
            df2.to_excel(writer, index=False, sheet_name='Contas a Pagar')
            df3.to_excel(writer, index=False, sheet_name='FechamentoCaixa')
        output.seek(0)
        
        st.download_button(
            label="Baixar Relatório Consolidado (Excel)",
            data=output,
            file_name=f"relatorio_consolidado_{start_date_str}_a_{end_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except pyodbc.Error as e:
        st.error(f"Erro relacionado ao ODBC: {e}")
    except Exception as ex:
        st.error(f"Erro inesperado: {ex}")
    finally:
        if 'conn' in locals() and conn:
            conn.close()
            st.info("Conexão fechada.")