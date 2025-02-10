import streamlit as st
import pandas as pd
import pyodbc
import os
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

# Novo dicionário para mapeamento da coluna "Conta" (para a aba "Contas a Pagar")
mapping_dict_plan = {
    " - ": "Outras saidas",
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

# Dicionário para mapeamento de ID_Empresa para Nome empresa
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

if st.button("Gerar Relatório Consolidado"):
    try:
        st.info("Estabelecendo conexão com o banco de dados...")
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        st.success("Conexão estabelecida!")
        
        ### --- EXECUÇÃO DAS 3 QUERIES ORIGINAIS ---
        # Query 1: Pesquisa_Transferencias_Busca (aba "Relatorio")
        sql_query1 = f"""
        SELECT *
        FROM Pesquisa_Transferencias_Busca
        WHERE (
              [ID Conta Origem] IN (
                  SELECT ID_Conta 
                  FROM Financeiro_Contas_Acessos
                  WHERE ID_Usuario = 1 AND ISNULL(Visualizar, 'N') = 'S'
              )
              OR
              [ID Conta Destino] IN (
                  SELECT ID_Conta 
                  FROM Financeiro_Contas_Acessos
                  WHERE ID_Usuario = 1 AND ISNULL(Visualizar, 'N') = 'S'
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
        st.success("Query 1 executada!")
        
        # Query 2: view_Contas_a_Pagar (aba "Contas a Pagar")
        sql_query2 = f"""
        SELECT 
             ID_Empresa,
             [Plano de Contas],
             Conta,
             [Centro Custo],
             emissao,
             pagamento,
             [Descrição Lançamento],
             Valor
        FROM view_Contas_a_Pagar
        WHERE ID_Situacao IN (0,1)
          AND CONVERT(DATE, emissao) >= '{start_date_str}'
          AND CONVERT(DATE, emissao) <= '{end_date_str}'
          AND ID_Empresa IN (55,58,57,65,50,66,64,61,60,46,59,56,51,53,52);
        """
        st.info("Executando Query 2...")
        cursor.execute(sql_query2)
        columns2 = [col[0] for col in cursor.description]
        rows2 = cursor.fetchall()
        data2 = [dict(zip(columns2, row)) for row in rows2]
        df2 = pd.DataFrame(data2)
        # Criar a coluna "De Para" a partir do mapeamento da coluna "Conta"
        if 'Conta' in df2.columns:
            df2["De Para"] = df2["Conta"].map(mapping_dict_plan).fillna("")
        # Aplicar a função remove_currency na coluna "Valor", se existir
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
        
        ### --- EXECUÇÃO DA NOVA QUERY (Query 4: vendas trocadas) ---
        sql_query4 = f"""
        SELECT 
            pr.*,
            fc.DataAbertura,
            fc.DataFechamento
        FROM Pesquisa_Resumo_Conferencia_Apuracao pr
        INNER JOIN Fechamento_Caixas fc ON
            fc.ID_Caixa = pr.ID_Caixa AND
            fc.ID_Empresa = pr.ID_Empresa AND
            fc.ID_Origem_Caixa = pr.ID_Origem_Caixa 
        WHERE 
            fc.ID_Empresa IN (55, 58, 57, 65, 50, 66, 64, 61, 60, 46, 59, 56, 51, 53, 52) AND 
            fc.ID_Origem_Caixa = 1 AND
            fc.DataFechamento >= '{start_date_str}' AND
            fc.DataFechamento <= '{end_date_str}'
        ORDER BY fc.ID_Caixa;
        """
        st.info("Executando Query 4 (vendas trocadas)...")
        cursor.execute(sql_query4)
        columns4 = [col[0] for col in cursor.description]
        rows4 = cursor.fetchall()
        data4 = [dict(zip(columns4, row)) for row in rows4]
        df4 = pd.DataFrame(data4)
        st.success("Query 4 executada!")
        
        # --- TRATAMENTO DOS DADOS ---
        # Remover horário das datas de df1, df2 e df3
        for df in [df1, df2, df3]:
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.date

        # Se a planilha "Relatorio" (df1) tiver a coluna "Valor", aplicar o tratamento
        if 'Valor' in df1.columns:
            df1['Valor'] = df1['Valor'].apply(remove_currency)
            
        # Reordenar df1 para que "ID Empresa" seja a primeira coluna
        if 'ID Empresa' in df1.columns:
            df1 = df1[['ID Empresa'] + [col for col in df1.columns if col != 'ID Empresa']]

        if 'ID_Empresa' in df4.columns:
            df4 = df4[['ID_Empresa'] + [col for col in df4.columns if col != 'ID_Empresa']]
        
        # Tratamento específico para df4:
        cols_to_float = ['Apurado_Sistema', 'Apurado_Operador', 'Apurado_Gerente', 'Diferenca_Operador', 'Diferenca_Gerente']
        for col in cols_to_float:
            if col in df4.columns:
                df4[col] = pd.to_numeric(df4[col], errors='coerce')
        if 'DataFechamento' in df4.columns:
            df4['DataFechamento'] = pd.to_datetime(df4['DataFechamento']).dt.date
        
        # Exibir os DataFrames para conferência (opcional)
        st.subheader("Pesquisa_Transferencias_Busca (Relatorio)")
        st.dataframe(df1)
        st.subheader("view_Contas_a_Pagar")
        st.dataframe(df2)
        st.subheader("Fechamento de Caixa")
        st.dataframe(df3)
        st.subheader("vendas trocadas")
        st.dataframe(df4)
        
        ### --- GERAR ARQUIVO EXCEL COM VÁRIAS ABAS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Função auxiliar para ajustar largura das colunas e adicionar tabela (com filtros)
            def format_worksheet_as_table(worksheet, df, table_name):
                for i, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(i, i, max_len + 2)
                n_rows, n_cols = df.shape
                table_range = f"A1:{xl_col_to_name(n_cols-1)}{n_rows+1}"
                worksheet.add_table(table_range, {
                    'name': table_name,
                    'columns': [{'header': col} for col in df.columns]
                })
            
            # --- ABA RESULTADO ---
            worksheet_result = workbook.add_worksheet('Resultado')
            result_headers = ["ID Empresa", "Nome empresa", "Data emissão", "Vendas em dinheiro", 
                              "Valor transferência", "Depósitos", "Saídas", "Transf x Deposito", "Falta depositar"]
            for col_num, header in enumerate(result_headers):
                worksheet_result.write(0, col_num, header)
            mapping_items = list(id_empresa_mapping.items())
            for i, (id_emp, nome_emp) in enumerate(mapping_items):
                excel_row = i + 2  # Linha no Excel (a primeira linha é o cabeçalho)
                worksheet_result.write(excel_row-1, 0, id_emp)
                worksheet_result.write(excel_row-1, 1, nome_emp)
                worksheet_result.write(excel_row-1, 2, start_date.strftime("%d/%m/%Y"))
                formula_vendas = f"=SUMIFS(FechamentoCaixa!H:H,FechamentoCaixa!A:A,Resultado!A{excel_row},FechamentoCaixa!D:D,Resultado!C{excel_row})"
                worksheet_result.write_formula(excel_row-1, 3, formula_vendas)
                formula_transf = f"=SUMIFS(FechamentoCaixa!J:J,FechamentoCaixa!A:A,Resultado!A{excel_row},FechamentoCaixa!D:D,Resultado!C{excel_row})"
                worksheet_result.write_formula(excel_row-1, 4, formula_transf)
                formula_depositos = f'=SUMIFS(Relatorio!H:H,Relatorio!A:A,Resultado!A{excel_row},Relatorio!J:J,Resultado!C{excel_row},Relatorio!D:D,"FINANCEIRO PARA FINANCEIRO")'
                worksheet_result.write_formula(excel_row-1, 5, formula_depositos)
                formula_saidas = f'=SUMIFS(\'Contas a Pagar\'!H:H,\'Contas a Pagar\'!A:A,Resultado!A{excel_row},\'Contas a Pagar\'!E:E,Resultado!C{excel_row},\'Contas a Pagar\'!I:I,"Saídas")'
                worksheet_result.write_formula(excel_row-1, 6, formula_saidas)
                formula_transf_deposito = f"=E{excel_row} - F{excel_row}"
                worksheet_result.write_formula(excel_row-1, 7, formula_transf_deposito)
                formula_falta_depositar = f"=E{excel_row} - G{excel_row} - F{excel_row}"
                worksheet_result.write_formula(excel_row-1, 8, formula_falta_depositar)
            
            n_rows_result = len(mapping_items) + 1
            n_cols_result = len(result_headers)
            col_widths = []
            for j in range(n_cols_result):
                header_len = len(result_headers[j])
                if j == 0:
                    max_data_len = max(len(str(id_emp)) for id_emp, _ in mapping_items) if mapping_items else header_len
                    col_widths.append(max(header_len, max_data_len) + 2)
                elif j == 1:
                    max_data_len = max(len(str(nome_emp)) for _, nome_emp in mapping_items) if mapping_items else header_len
                    col_widths.append(max(header_len, max_data_len) + 2)
                elif j == 2:
                    col_widths.append(max(header_len, 10) + 2)
                else:
                    col_widths.append(header_len + 5)
            for j, width in enumerate(col_widths):
                worksheet_result.set_column(j, j, width)
            table_range_result = f"A1:{xl_col_to_name(n_cols_result-1)}{n_rows_result}"
            worksheet_result.add_table(table_range_result, {
                'name': 'TableResultado',
                'columns': [{'header': header} for header in result_headers]
            })
            
            # --- Abas dos DataFrames ---
            df1.to_excel(writer, index=False, sheet_name='Relatorio')
            ws_relatorio = writer.sheets["Relatorio"]
            format_worksheet_as_table(ws_relatorio, df1, "TableRelatorio")
            
            df2.to_excel(writer, index=False, sheet_name='Contas a Pagar')
            ws_contas = writer.sheets["Contas a Pagar"]
            format_worksheet_as_table(ws_contas, df2, "TableContasAPagar")
            
            # Adicionar coluna de conferência em df3
            nova_coluna = []
            for i in range(df3.shape[0]):
                excel_row = i + 2  # Considerando que a linha 1 é o cabeçalho
                formula_conf = f"=IF(G{excel_row}-K{excel_row}=0,0,G{excel_row}-K{excel_row}-L{excel_row})"
                nova_coluna.append(formula_conf)
            df3["conferência fundo de troco"] = nova_coluna
            
            df3.to_excel(writer, index=False, sheet_name='FechamentoCaixa')
            ws_fechamento = writer.sheets["FechamentoCaixa"]
            format_worksheet_as_table(ws_fechamento, df3, "TableFechamentoCaixa")
            
            df4.to_excel(writer, index=False, sheet_name='vendas trocadas')
            ws_vendas = writer.sheets["vendas trocadas"]
            format_worksheet_as_table(ws_vendas, df4, "TableVendasTrocadas")
            
            mapping_df = pd.DataFrame(list(mapping_dict_plan.items()), columns=["Conta", "De Para"])
            mapping_df.to_excel(writer, index=False, sheet_name="De para")
            ws_depara = writer.sheets["De para"]
            format_worksheet_as_table(ws_depara, mapping_df, "TableDePara")
            
        output.seek(0)
        
        st.download_button(
            label="Baixar Relatório Final (Excel)",
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