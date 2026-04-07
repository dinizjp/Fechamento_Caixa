import streamlit as st
import pandas as pd
import pyodbc
import io
from datetime import datetime
from xlsxwriter.utility import xl_col_to_name

st.set_page_config(page_title='Relatório Consolidado', layout='wide')

server   = st.secrets["mssql"]["server"]
database = st.secrets["mssql"]["database"]
username = st.secrets["mssql"]["username"]
password = st.secrets["mssql"]["password"]

if not all([server, database, username, password]):
    st.error("Verifique se todas as variáveis estão definidas corretamente no secrets.toml.")
    st.stop()

st.title("Relatório Consolidado")

start_date = st.date_input("Data de Início", value=datetime.now())
end_date   = st.date_input("Data de Fim",    value=datetime.now())

if start_date > end_date:
    st.error("A data de início não pode ser maior que a data de fim.")
    st.stop()

start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str   = end_date.strftime('%Y-%m-%d')
st.write(f"Período selecionado: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")

connection_string = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password};"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
)

def remove_currency(val):
    try:
        if pd.isnull(val):
            return None
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip().replace("R$", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return None

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

def query_to_df(cursor, sql, params):
    cursor.execute(sql, params)
    columns = [col[0] for col in cursor.description]
    return pd.DataFrame.from_records(cursor.fetchall(), columns=columns)

def format_worksheet_as_table(worksheet, df, table_name):
    for i, col in enumerate(df.columns):
        if df.empty:
            max_len = len(str(col))
        else:
            max_data = df[col].astype(str).str.len().max()
            max_len = max(int(max_data) if pd.notna(max_data) else 0, len(str(col)))
        worksheet.set_column(i, i, max_len + 2)
    n_rows, n_cols = df.shape
    table_range = f"A1:{xl_col_to_name(n_cols - 1)}{n_rows + 1}"
    worksheet.add_table(table_range, {
        'name': table_name,
        'columns': [{'header': col} for col in df.columns]
    })

if st.button("Gerar Relatório Consolidado"):
    try:
        st.info("Estabelecendo conexão com o banco de dados...")
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        st.success("Conexão estabelecida!")

        params = (start_date_str, end_date_str)

        sql_query1 = """
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
        AND CONVERT(DATE, emissao) >= ?
        AND CONVERT(DATE, emissao) <= ?
        ORDER BY emissao DESC;
        """

        sql_query2 = """
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
          AND CONVERT(DATE, emissao) >= ?
          AND CONVERT(DATE, emissao) <= ?
          AND ID_Empresa IN (55,58,57,65,50,66,64,61,60,46,59,56,51,53,52);
        """

        sql_query3 = """
        SELECT
            r.ID_Empresa,
            r.ID_Caixa,
            SUBSTRING(CONVERT(varchar, [Data_Abertura], 120), 1, 10) AS Data_Abertura_Str,
            [Data_Abertura],
            [Data_fechamento],
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
            AND SUBSTRING(CONVERT(varchar, [Data_Abertura], 120), 1, 10) >= ?
            AND SUBSTRING(CONVERT(varchar, [Data_Abertura], 120), 1, 10) <= ?
            AND [ID_Origem_Caixa] = 1
        ORDER BY
            r.ID_Empresa, r.ID_Caixa, SUBSTRING(CONVERT(varchar, [Data_Abertura], 120), 1, 10);
        """

        sql_query4 = """
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
            fc.DataFechamento >= ? AND
            fc.DataFechamento <= ?
        ORDER BY fc.ID_Caixa;
        """

        st.info("Executando Query 1...")
        df1 = query_to_df(cursor, sql_query1, params)
        st.success("Query 1 executada!")

        st.info("Executando Query 2...")
        df2 = query_to_df(cursor, sql_query2, params)
        st.success("Query 2 executada!")

        st.info("Executando Query 3...")
        df3 = query_to_df(cursor, sql_query3, params)
        st.success("Query 3 executada!")

        st.info("Executando Query 4 (vendas trocadas)...")
        df4 = query_to_df(cursor, sql_query4, params)
        st.success("Query 4 executada!")

        # --- TRATAMENTO DOS DADOS ---

        # Remover horário das colunas datetime em df1, df2, df3
        for df in [df1, df2, df3]:
            for col in df.select_dtypes(include=['datetime64']).columns:
                df[col] = df[col].dt.date

        if 'Valor' in df1.columns:
            df1['Valor'] = df1['Valor'].apply(remove_currency)

        if 'ID Empresa' in df1.columns:
            df1 = df1[['ID Empresa'] + [c for c in df1.columns if c != 'ID Empresa']]

        if 'Conta' in df2.columns:
            df2["Conta"] = df2["Conta"].str.strip()
            df2["De Para"] = df2["Conta"].map(mapping_dict_plan).fillna("")
        if 'Valor' in df2.columns:
            df2['Valor'] = df2['Valor'].apply(remove_currency)

        if 'ID_Empresa' in df4.columns:
            df4 = df4[['ID_Empresa'] + [c for c in df4.columns if c != 'ID_Empresa']]

        cols_to_float = ['Apurado_Sistema', 'Apurado_Operador', 'Apurado_Gerente', 'Diferenca_Operador', 'Diferenca_Gerente']
        for col in cols_to_float:
            if col in df4.columns:
                df4[col] = pd.to_numeric(df4[col], errors='coerce')
        if 'DataFechamento' in df4.columns:
            df4['DataFechamento'] = pd.to_datetime(df4['DataFechamento']).dt.date

        st.subheader("Pesquisa_Transferencias_Busca (Relatorio)")
        st.dataframe(df1)
        st.subheader("view_Contas_a_Pagar")
        st.dataframe(df2)
        st.subheader("Fechamento de Caixa")
        st.dataframe(df3)
        st.subheader("vendas trocadas")
        st.dataframe(df4)

        # --- GERAR ARQUIVO EXCEL ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book

            # Aba Resultado
            worksheet_result = workbook.add_worksheet('Resultado')
            result_headers = ["ID Empresa", "Nome empresa", "Data emissão", "Vendas em dinheiro",
                              "Valor transferência", "Diferença", "Depósitos", "Saídas", "Transf x Deposito", "Falta depositar"]
            for col_num, header in enumerate(result_headers):
                worksheet_result.write(0, col_num, header)

            mapping_items = list(id_empresa_mapping.items())
            for i, (id_emp, nome_emp) in enumerate(mapping_items):
                r = i + 2  # linha Excel (1-indexed, linha 1 é cabeçalho)
                worksheet_result.write(r - 1, 0, id_emp)
                worksheet_result.write(r - 1, 1, nome_emp)
                worksheet_result.write(r - 1, 2, start_date.strftime("%d/%m/%Y"))
                worksheet_result.write_formula(r - 1, 3, f"=SUMIFS(FechamentoCaixa!H:H,FechamentoCaixa!A:A,Resultado!A{r},FechamentoCaixa!D:D,Resultado!C{r})")
                worksheet_result.write_formula(r - 1, 4, f"=SUMIFS(FechamentoCaixa!J:J,FechamentoCaixa!A:A,Resultado!A{r},FechamentoCaixa!D:D,Resultado!C{r})")
                worksheet_result.write_formula(r - 1, 5, f"=D{r}-E{r}")
                worksheet_result.write_formula(r - 1, 6, f'=SUMIFS(Relatorio!H:H,Relatorio!A:A,Resultado!A{r},Relatorio!J:J,Resultado!C{r},Relatorio!D:D,"FINANCEIRO PARA FINANCEIRO")')
                worksheet_result.write_formula(r - 1, 7, f'=SUMIFS(\'Contas a Pagar\'!H:H,\'Contas a Pagar\'!A:A,Resultado!A{r},\'Contas a Pagar\'!E:E,Resultado!C{r},\'Contas a Pagar\'!I:I,"Saídas")')
                worksheet_result.write_formula(r - 1, 8, f"=E{r} - G{r}")
                worksheet_result.write_formula(r - 1, 9, f"=E{r} - H{r} - G{r}")

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

            n_rows_result = len(mapping_items) + 1
            worksheet_result.add_table(f"A1:{xl_col_to_name(n_cols_result - 1)}{n_rows_result}", {
                'name': 'TableResultado',
                'columns': [{'header': h} for h in result_headers]
            })

            # Abas dos DataFrames
            df1.to_excel(writer, index=False, sheet_name='Relatorio')
            format_worksheet_as_table(writer.sheets["Relatorio"], df1, "TableRelatorio")

            df2.to_excel(writer, index=False, sheet_name='Contas a Pagar')
            format_worksheet_as_table(writer.sheets["Contas a Pagar"], df2, "TableContasAPagar")

            df3["conferência fundo de troco"] = [
                f"=IF(G{i+2}-K{i+2}=0,0,G{i+2}-K{i+2}-L{i+2})"
                for i in range(df3.shape[0])
            ]
            df3.to_excel(writer, index=False, sheet_name='FechamentoCaixa')
            format_worksheet_as_table(writer.sheets["FechamentoCaixa"], df3, "TableFechamentoCaixa")

            df4.to_excel(writer, index=False, sheet_name='vendas trocadas')
            format_worksheet_as_table(writer.sheets["vendas trocadas"], df4, "TableVendasTrocadas")

            mapping_df = pd.DataFrame(list(mapping_dict_plan.items()), columns=["Conta", "De Para"])
            mapping_df.to_excel(writer, index=False, sheet_name="De para")
            format_worksheet_as_table(writer.sheets["De para"], mapping_df, "TableDePara")

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
