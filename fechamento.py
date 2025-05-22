import streamlit as st
import pandas as pd
import pyodbc
import io
from datetime import datetime
from xlsxwriter.utility import xl_col_to_name  # Para converter índice de coluna para notação (A, B, ...)

# Configurar a página
st.set_page_config(page_title='Relatório Consolidado', layout='wide')

# Recuperar as credenciais dos segredos (.streamlit/secrets.toml)
server = st.secrets["mssql"]["server"]
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
end_date = st.date_input("Data de Fim", value=datetime.now())
if start_date > end_date:
    st.error("A data de início não pode ser maior que a data de fim.")
    st.stop()
start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str = end_date.strftime('%Y-%m-%d')
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

# Dicionários de mapeamento
mapping_dict_plan = {
    "-": "Outras saidas",
    "66 - CAIXA TESOURARIA | FORMOSA": "Saídas",
    "47 - CAIXA TESOURARIA | BALSAS": "Saídas",
    "85 - CAIXA TESOURARIA- BAIXAS DOS COLABORADORES (CONSUMO AÇAI) VALE": "Outras saidas",
    "7 - CAIXA TESOURARIA | ITZ PEDRO NEIVA": "Saídas",
    "8 - CAIXA TESOURARIA | ITZ BS": "Saídas",
    "71 - CAIXA TESOURARIA  | MOMÊ": "Saídas",
    "60 - CAIXA TESOURARIA | CENTRAL": "Saídas",
    "86 - CAIXA TESOURARIA MOMÊ VIA LAGO": "Saídas",
    "13 - CAIXA TESOURARIA | GURUPI": "Saídas",
    "5 - CAIXA TESOURARIA | ARN JB": "Saídas",
    "9 - CAIXA TESOURARIA | ARN MN": "Saídas",
    "88 - CAIXA TESOURARIA | GUARAI": "Saídas",
    "100 - CAIXA TESOURARIA BAIXA CONSUMO AÇAI | T S MOURA": "Outras saidas",
    "92 - CAIXA TESOURARIA | COLINAS": "Saídas",
    "90 - CAIXA TESOURARIA | ESTREITO": "Saídas",
    "99 - CAIXA TESOURARIA CANAÃ": "Saídas",
    "102 - CAIXA TESOURARIA VIA LAGO KIDS": "Saídas",
    "105 - CAIXA TESOURARIA BALSAS 02": "Saídas",
    "101 - CAIXA TESOURARIA HUB | THIAGO": "Saídas"
}

id_empresa_mapping = {
    58: 'Araguaína II', 66: 'Balsas II', 55: 'Araguaína I',
    53: 'Imperatriz II', 51: 'Imperatriz I', 65: 'Araguaína IV',
    52: 'Imperatriz III', 57: 'Araguaína III', 50: 'Balsas I',
    56: 'Gurupi I', 61: 'Colinas', 60: 'Estreito',
    46: 'Formosa I', 59: 'Guaraí'
}

if st.button("Gerar Relatório Consolidado"):
    try:
        st.info("Estabelecendo conexão com o banco de dados...")
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        st.success("Conexão estabelecida!")

        # --- Query 1 ---
        sql_query1 = f"""
        SELECT *
        FROM Pesquisa_Transferencias_Busca
        WHERE (
          [ID Conta Origem] IN (SELECT ID_Conta FROM Financeiro_Contas_Acessos WHERE ID_Usuario=1 AND ISNULL(Visualizar,'N')='S')
          OR [ID Conta Destino] IN (SELECT ID_Conta FROM Financeiro_Contas_Acessos WHERE ID_Usuario=1 AND ISNULL(Visualizar,'N')='S')
        )
        AND CONVERT(DATE, emissao) BETWEEN '{start_date_str}' AND '{end_date_str}'
        ORDER BY emissao DESC;"""
        cursor.execute(sql_query1)
        df1 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        st.success("Query 1 executada!")

        # --- Query 2 ---
        sql_query2 = f"""
        SELECT ID_Empresa,[Plano de Contas],Conta,[Centro Custo],emissao,pagamento,[Descrição Lançamento],Valor
        FROM view_Contas_a_Pagar
        WHERE ID_Situacao IN (0,1)
          AND CONVERT(DATE, emissao) BETWEEN '{start_date_str}' AND '{end_date_str}'
          AND ID_Empresa IN ({','.join(map(str,id_empresa_mapping.keys()))});"""
        cursor.execute(sql_query2)
        df2 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        if 'Conta' in df2.columns:
            df2['Conta'] = df2['Conta'].str.strip()
            df2['De Para'] = df2['Conta'].map(mapping_dict_plan).fillna("")
        if 'Valor' in df2.columns:
            df2['Valor'] = df2['Valor'].apply(remove_currency)
        st.success("Query 2 executada!")

        
        # Query 3: Fechamento de Caixa (aba "FechamentoCaixa")
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
                  WHERE t.ID_Empresa = r.ID_Empresa AND t.ID_Caixa = r.ID_Caixa), 0.00)
                 ) AS Transf_Tesour, 
            CONVERT(float, ISNULL(
                 ((SELECT SUM(apurado_gerente)
                   FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                   WHERE FG.ID_Empresa = r.ID_Empresa AND FG.ID_Caixa = r.ID_Caixa)
                 - (SELECT ISNULL(valor, 0.00)
                    FROM Financeiro_Transferencias t 
                    WHERE t.ID_Empresa = r.ID_Empresa AND t.ID_Caixa = r.ID_Caixa)
                 ), 0.00)
                 ) AS Ap_Ger_Nao_Trans,
            CONVERT(float, (SELECT SUM(apurado_gerente)
                 FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                 WHERE FG.ID_Empresa = r.ID_Empresa AND FG.ID_Caixa = r.ID_Caixa)
                 ) AS Apur_Ger_total,
            CONVERT(float, 
                 ((SELECT SUM(apurado_gerente)
                   FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                   WHERE FG.ID_Empresa = r.ID_Empresa AND FG.ID_Caixa = r.ID_Caixa)
                 - Total_Entradas_Dinheiro)
                 ) AS SaldoFinal,
            CASE 
               WHEN ((SELECT ISNULL(SUM(apurado_gerente), 0.00)
                      FROM Fechamento_Caixa_Conferencia_Sangrias FG 
                      WHERE FG.ID_Empresa = r.ID_Empresa AND FG.ID_Caixa = r.ID_Caixa)
                    - ISNULL(Total_Entradas_Dinheiro, 0.00)) <= -3.00 
               THEN 'Vale' 
               ELSE 'Nao' 
            END AS Vale
        FROM 
            View_FechamentoCaixa_Resumo r
        INNER JOIN 
            Pesquisa_Fechamento_Caixas c 
            ON r.ID_Caixa = [ID Caixa] AND r.ID_Empresa = [ID Empresa]
        WHERE  
            r.ID_Empresa IN (55,58,57,65,50,66,64,61,60,46,59,56,51,53,52)  
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
        st.success("Query 3 executada!")


        
        # --- Query 4 ---
        sql_query4 = f"""
        SELECT pr.*,fc.DataAbertura,fc.DataFechamento
        FROM Pesquisa_Resumo_Conferencia_Apuracao pr
        INNER JOIN Fechamento_Caixas fc ON fc.ID_Caixa=pr.ID_Caixa AND fc.ID_Empresa=pr.ID_Empresa AND fc.ID_Origem_Caixa=pr.ID_Origem_Caixa
        WHERE fc.ID_Empresa IN ({','.join(map(str,id_empresa_mapping.keys()))})
          AND fc.ID_Origem_Caixa=1
          AND fc.DataFechamento BETWEEN '{start_date_str}' AND '{end_date_str}'
        ORDER BY fc.ID_Caixa;"""
        cursor.execute(sql_query4)
        df4 = pd.DataFrame([dict(zip([c[0] for c in cursor.description], row)) for row in cursor.fetchall()])
        st.success("Query 4 executada!")

        # --- Tratamento de datas ---
        for df in [df1, df2, df3]:
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.date

        # Gerar Excel com abas e download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            def format_worksheet_as_table(ws, df, name):
                for i, col in enumerate(df.columns):
                    ws.set_column(i, i, max(df[col].astype(str).map(len).max(), len(col)) + 2)
                nrows, ncols = df.shape
                ws.add_table(f"A1:{xl_col_to_name(ncols-1)}{nrows+1}", {'columns':[{'header':c} for c in df.columns]})

            df1.to_excel(writer, sheet_name='Relatorio', index=False)
            format_worksheet_as_table(writer.sheets['Relatorio'], df1, 'TableRelatorio')
            df2.to_excel(writer, sheet_name='Contas a Pagar', index=False)
            format_worksheet_as_table(writer.sheets['Contas a Pagar'], df2, 'TableContasAPagar')
            df3.to_excel(writer, sheet_name='FechamentoCaixa', index=False)
            format_worksheet_as_table(writer.sheets['FechamentoCaixa'], df3, 'TableFechamentoCaixa')
            df4.to_excel(writer, sheet_name='vendas trocadas', index=False)
            format_worksheet_as_table(writer.sheets['vendas trocadas'], df4, 'TableVendasTrocadas')

        output.seek(0)
        st.download_button("Baixar Relatório Final (Excel)", data=output, file_name=f"relatorio_consolidado_{start_date_str}_a_{end_date_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except pyodbc.Error as e:
        st.error(f"Erro relacionado ao ODBC: {e}")
    except Exception as ex:
        st.error(f"Erro inesperado: {ex}")
    finally:
        if 'conn' in locals() and conn:
            conn.close()
            st.info("Conexão fechada.")
