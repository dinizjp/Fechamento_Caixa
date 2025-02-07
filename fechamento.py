import streamlit as st
import pandas as pd
import pyodbc
import os
import io
from datetime import datetime

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

# Seleção do período (widgets separados)
start_date = st.date_input("Data de Início", value=datetime(2025, 1, 1))
end_date   = st.date_input("Data de Fim", value=datetime.now())
if start_date > end_date:
    st.error("A data de início não pode ser maior que a data de fim.")
    st.stop()
start_date_str = start_date.strftime('%Y-%m-%d')
end_date_str   = end_date.strftime('%Y-%m-%d')
st.write(f"Período selecionado: {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")

# String de conexão (usa ODBC Driver 17 para testes)
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

# Dicionário para mapeamento da coluna "Plano de Contas" (para a aba "Contas a Pagar")
mapping_dict_plan = {
    "OUTRAS PERCAS": "outras percas",
    "COMPRA DE INSUMOS | COMPLEMENTOS": "Saidas",
    "COMPRAS LOCAIS ": "Saidas",
    "DIARIAS ATENDENTES": "Saidas",
    "MANUTENÇÃO DE IMÓVEL": "Saidas",
    "VIGILANTES": "Saidas",
    "MATERIAIS DE USO E CONSUMO": "Saidas",
    "ADIANTAMENTO DE LUCRO | ANA PAULA": "Saidas",
    "MATERIAL DE ESCRITÓRIO": "Saidas",
    "COMPUTADORES E PERIFÉRICOS": "Saidas",
    "AGUA E SANEAMENTO": "Saidas",
    "COMPRA DE MERCADORIAS": "Saidas",
    "COMISSOES TERCEIROS": "Saidas",
    "TAXAS ADM - VENDAS APP TERCEIROS": "Saidas",
    "COMPRA DE INSUMOS | EMBALAGENS": "Saidas",
    "MATERIAL DE HIGIENE E LIMPEZA": "Saidas",
    "COMPRA DE PERIFERICOS": "Saidas",
    "TAXI E CONDUCOES": "Saidas",
    "REFEIÇÕES E LANCHES": "Saidas",
    "HONORÁRIOS ADVOCATÍCIOS": "Saidas",
    "SALARIO PESSOAL": "Saidas",
    "CONSULTORIA FINANCEIRA": "Saidas",
    "BONIFICAÇÃO": "Saidas",
    "DIÁRIA CHAPA": "Saidas",
    "DEPOSITO BANCARIO": "Saidas",
    "SERVIÇOS DE LIMPEZA E CONSERVACAO": "Saidas",
    "CAPTACAO MOEDAS | TROCO": "Saidas",
    "RESCISOES TRABALHISTAS": "Saidas",
    "ADIANTAMENTO DE LUCRO | PEDRO HENRIQUE": "Saidas",
    "DIARIAS DELIVERY | MOTOBOY": "Saidas",
    "FRETES E CARRETOS": "Saidas",
    "SERVIÇOS PRESTADOS POR TERCEIROS - PJ": "Saidas",
    "INSTALAÇÕES | MATERIAL BÁSICO": "Saidas",
    "DESPESAS DIVERSAS COM PESSOAL": "Saidas",
    "OUTRAS DESPESAS": "Saidas",
    "PUBLICIDADE E PROPAGANDA": "Saidas",
    "FORMULÁRIOS E IMPRESSOS": "Saidas",
    "EXAMES MEDICOS TRABALHISTAS": "Saidas",
    "MANUTENÇÃO COMPUTADORES": "Saidas",
    "ORNAMENTACOES": "Saidas",
    "MANUTENÇÃO MÓVEIS": "Saidas",
    "ALUGUEIS IMOVEIS": "Saidas",
    "MANUTENÇÃO ELETRICA": "Saidas",
    "OBRAS | MATERIAL BASICO": "Saidas",
    "MANUTENÇÃO MAQ. E EQUIPAMENTOS": "Saidas",
    "BENS MÓVEIS - OPERACIONAL": "Saidas",
    "DIARIAS MOTORISTAS": "Saidas",
    "COPIAS E AUTENTICACOES": "Saidas",
    "COMUNICAÇÃO VISUAL": "Saidas",
    "COMBUSTIVEL DIVERSOS": "Saidas",
    "COMPRA DE INSUMOS | ACAI": "Saidas",
    "PREST. SERVIÇOS DE TERCEIROS - MARKETING": "Saidas",
    "MANUTENÇÕES - PLAYGROUND": "Saidas",
    "SERVIÇOS PRESTADOS POR TERCEIROS - PF": "Saidas",
    "FALTA DE CAIXA": "Saidas"
}

# Dicionário para mapeamento de ID_Empresa para Nome empresa (14 itens)
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
        
        ### --- EXECUÇÃO DAS 3 QUERIES ---
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
        # Criar a coluna "De Para" a partir do mapeamento da coluna "Plano de Contas"
        if 'Plano de Contas' in df2.columns:
            df2["De Para"] = df2["Plano de Contas"].map(mapping_dict_plan).fillna("")
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
        
        ### --- TRATAMENTO DOS DADOS: REMOVER HORÁRIO (mantendo somente a data) ---
        for df in [df1, df2, df3]:
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.date
        
        # Se a planilha "Relatorio" (df1) tiver a coluna "Valor", aplicar o tratamento
        if 'Valor' in df1.columns:
            df1['Valor'] = df1['Valor'].apply(remove_currency)
        
        # Exibir os DataFrames para conferência (opcional)
        st.subheader("Pesquisa_Transferencias_Busca (Relatorio)")
        st.dataframe(df1)
        st.subheader("view_Contas_a_Pagar")
        st.dataframe(df2)
        st.subheader("Fechamento de Caixa")
        st.dataframe(df3)
        
        ### --- GERAR ARQUIVO EXCEL COM 5 ABAS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            # Criar a aba "Resultado" primeiro (para que ela fique como primeira planilha)
            worksheet_result = workbook.add_worksheet('Resultado')
            # Cabeçalhos da aba Resultado (9 colunas, sem "Suprimento")
            result_headers = ["ID Empresa", "Nome empresa", "Data emissão", "Vendas em dinheiro", 
                              "Valor transferência", "Depósitos", "Saídas", "Transf x Deposito", "Falta depositar"]
            for col_num, header in enumerate(result_headers):
                worksheet_result.write(0, col_num, header)
            # Preencher a aba "Resultado" usando os 14 itens do dicionário id_empresa_mapping
            mapping_items = list(id_empresa_mapping.items())
            for i, (id_emp, nome_emp) in enumerate(mapping_items):
                excel_row = i + 2  # Linha no Excel (linha 1 é o cabeçalho)
                worksheet_result.write(excel_row-1, 0, id_emp)                            # Coluna A: ID Empresa
                worksheet_result.write(excel_row-1, 1, nome_emp)                            # Coluna B: Nome empresa
                worksheet_result.write(excel_row-1, 2, start_date.strftime("%d/%m/%Y"))       # Coluna C: Data emissão (usa data de início)
                # Inserir as fórmulas para as demais colunas:
                # Vendas em dinheiro (coluna D)
                formula_vendas = f"=SUMIFS(FechamentoCaixa!H:H,FechamentoCaixa!A:A,Resultado!A{excel_row},FechamentoCaixa!D:D,Resultado!C{excel_row})"
                worksheet_result.write_formula(excel_row-1, 3, formula_vendas)
                # Valor transferência (coluna E)
                formula_transf = f"=SUMIFS(FechamentoCaixa!J:J,FechamentoCaixa!A:A,Resultado!A{excel_row},FechamentoCaixa!D:D,Resultado!C{excel_row})"
                worksheet_result.write_formula(excel_row-1, 4, formula_transf)
                # Depósitos (coluna F)
                formula_depositos = f'=SUMIFS(Relatorio!G:G,Relatorio!L:L,Resultado!A{excel_row},Relatorio!I:I,Resultado!C{excel_row},Relatorio!C:C,"FINANCEIRO PARA FINANCEIRO")'
                worksheet_result.write_formula(excel_row-1, 5, formula_depositos)
                # Saídas (coluna G)
                formula_saidas = f'=SUMIFS(\'Contas a Pagar\'!G:G,\'Contas a Pagar\'!A:A,Resultado!A{excel_row},\'Contas a Pagar\'!E:E,Resultado!C{excel_row},\'Contas a Pagar\'!H:H,"Saidas")'
                worksheet_result.write_formula(excel_row-1, 6, formula_saidas)
                # Transf x Deposito (coluna H): = Valor transferência - Depósitos
                formula_transf_deposito = f"=E{excel_row} - F{excel_row}"
                worksheet_result.write_formula(excel_row-1, 7, formula_transf_deposito)
                # Falta depositar (coluna I): = Valor transferência - Saídas - Depósitos
                formula_falta_depositar = f"=E{excel_row} - G{excel_row} - F{excel_row}"
                worksheet_result.write_formula(excel_row-1, 8, formula_falta_depositar)
            
            # Em seguida, escrever as demais abas
            df1.to_excel(writer, index=False, sheet_name='Relatorio')
            df2.to_excel(writer, index=False, sheet_name='Contas a Pagar')
            df3.to_excel(writer, index=False, sheet_name='FechamentoCaixa')
            # Escrever a aba "De para" com o mapeamento do plano de contas
            mapping_df = pd.DataFrame(list(mapping_dict_plan.items()), columns=["Plano de Contas", "De Para"])
            mapping_df.to_excel(writer, index=False, sheet_name="De para")
            
            # Adicionar coluna "conferência fundo de troco" na aba "FechamentoCaixa"
            worksheet_fc = writer.sheets["FechamentoCaixa"]
            n_cols_fc = len(df3.columns)  # Número atual de colunas
            worksheet_fc.write(0, n_cols_fc, "conferência fundo de troco")
            n_rows_fc = df3.shape[0]
            for r in range(1, n_rows_fc + 1):
                excel_row = r + 1  # Linha no Excel (primeira linha é o cabeçalho)
                # Supondo:
                # - Suprimento está na coluna G (índice 6)
                # - Ap_Ger_Nao_Trans está na coluna K (índice 10)
                # - Apur_Ger_total está na coluna L (índice 11)
                formula_conf = f"=IF(G{excel_row}-K{excel_row}=0,0,G{excel_row}-K{excel_row}-L{excel_row})"
                worksheet_fc.write_formula(r, n_cols_fc, formula_conf)
            
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