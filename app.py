#codigo totalmente funcional

# app.py — Streamlit: processador com logo, resumo e bordas
import io
import re
import requests
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, Side

# ---------- CONFIG ----------
LOGO_URL = "https://media.licdn.com/dms/image/v2/C4D0BAQFynSl_Yj90cQ/company-logo_200_200/company-logo_200_200/0/1630472942468/vicente_monteiro_advogados_logo?e=2147483647&v=beta&t=HUb5xhVbshv-LGKYpmpkkuJfUxX30S5oMjefFv7jM4s"
VALOR_HORA_DEFAULT = 462.62
# ----------------------------

st.set_page_config(page_title="Gerar planilha por cliente", layout="wide")
st.title("📊 Automatizador — Gerar planilha final por cliente")

# Upload do arquivo
uploaded_file = st.file_uploader("📤 Envie a planilha BRUTA (xlsx/xls)", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write(f"✅ Planilha carregada com {len(df)} linhas.")

    # --- Concatenar as duas primeiras colunas como CLIENTES ---
    coluna1, coluna2 = df.columns[0], df.columns[1]
    df["CLIENTES"] = df[[coluna1, coluna2]].fillna("").agg(" ".join, axis=1).str.strip()

    # --- Definir colunas na ordem correta ---
    colunas_para_manter = [
        "CLIENTES",
        "Duração",
        "Data de início",
        "Executante",
        "Descrição",
        "Vínculos com processo / Número de CNJ",
        "Contrário principal / Nome/razão social",
        "Vínculos com processo / Pasta"
    ]
    df_final = df[[col for col in colunas_para_manter if col in df.columns]].copy()

    # --- Formatar duração e datas ---
    if "Duração" in df_final.columns:
        df_final["Duração"] = df_final["Duração"].apply(lambda x: str(x).split()[-1])
    if "Data de início" in df_final.columns:
        df_final["Data de início"] = pd.to_datetime(df_final["Data de início"], errors="coerce").dt.strftime("%Y-%m-%d")

    # --- Função para somar durações ---
    def somar_duracoes(series):
        tempos = pd.to_timedelta(series, errors='coerce')
        total_segundos = tempos.dt.total_seconds().sum()
        horas = int(total_segundos // 3600)
        minutos = int((total_segundos % 3600) // 60)
        segundos = int(total_segundos % 60)
        return f"{horas:02}:{minutos:02}:{segundos:02}", total_segundos / 3600  # HH:MM:SS e decimal

    # --- Baixar logo ---
    response = requests.get(LOGO_URL)

    # --- Criar workbook ---
    wb = Workbook()
    wb.remove(wb.active)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    contador_tabela = 1  # Variável global para numerar tabelas

    # --- Função para criar aba ---
    def criar_aba(nome, df_dados, valor_hora):
        global contador_tabela
        ws = wb.create_sheet(title=nome[:31])

        # Logo
        logo_stream = BytesIO(response.content)
        img = Image(logo_stream)
        img.width = 150
        img.height = 150
        ws.add_image(img, "A1")

        # Resumo de horas
        start_row_horas = 9
        horas_totais_str, horas_totais_decimal = somar_duracoes(df_dados["Duração"])
        ws[f"A{start_row_horas}"] = "HORAS TOTAIS"
        ws[f"B{start_row_horas}"] = horas_totais_str
        ws[f"A{start_row_horas+1}"] = "VALOR HORA"
        ws[f"B{start_row_horas+1}"] = valor_hora
        ws[f"B{start_row_horas+1}"].number_format = "#.##0,00_);(#.##0,00)"
        ws[f"A{start_row_horas+2}"] = "TOTAL MENSAL"
        ws[f"B{start_row_horas+2}"] = f"=B{start_row_horas+3}*B{start_row_horas+1}"
        ws[f"B{start_row_horas+2}"].number_format = "#.##0,00_);(#.##0,00)"
        ws[f"B{start_row_horas+3}"] = horas_totais_decimal
        ws.row_dimensions[start_row_horas+3].hidden = True

        # Bordas resumo
        for row in range(start_row_horas, start_row_horas + 3):
            for col in range(1, 3):
                ws.cell(row=row, column=col).border = thin_border

        # Dados do cliente
        start_row_tabela = start_row_horas + 5
        for r_idx, row in enumerate(dataframe_to_rows(df_dados, index=False, header=True), start=start_row_tabela):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)
                ws.cell(row=r_idx, column=c_idx).border = thin_border

        # Criar tabela
        max_row = ws.max_row
        max_col = ws.max_column
        last_col_letter = get_column_letter(max_col)
        table_ref = f"A{start_row_tabela}:{last_col_letter}{max_row}"
        tabela_nome = f"TABELA_{contador_tabela}"
        contador_tabela += 1
        tab = Table(displayName=tabela_nome, ref=table_ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        tab.tableStyleInfo = style
        ws.add_table(tab)

    # --- Criar aba GERAL ---
    criar_aba("GERAL", df_final, VALOR_HORA_DEFAULT)

    # --- Criar abas por cliente ---
    for cliente in df_final["CLIENTES"].unique():
        df_cliente = df_final[df_final["CLIENTES"] == cliente]
        criar_aba(str(cliente), df_cliente, VALOR_HORA_DEFAULT)

    # --- Salvar em memória e download ---
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    st.success(f"Planilha gerada com {len(df_final['CLIENTES'].unique())+1} abas (GERAL + clientes).")
    st.download_button(
        "⤓ Baixar planilha final",
        data=bio,
        file_name="Planilha_Final_Clientes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )