import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os

from components.creating_pivot_table_prompt_divergence import CreatingPivotTablePromptDivergence
from components.creating_pivot_table_ai_vs_hu_divergence import CreatingPivotTableAiVsHuDivergence
from components.creating_pivot_each_comp_ia_hu_divergence import CreatingPivotEachCompIaHuDivergence
from components.creating_pivot_calculating_divergences import CreatingPivotCalculatingDivergences
from utils.load_excel_data import load_excel_data
from utils.sheets_helpers import create_spreadsheet_structure, get_sheet_id
from utils.config import *


st.set_page_config(page_title="Gerador de Tabela Dinâmica - Carta Argumentativa", layout="wide")
st.title("📊 Gerador de Tabela Dinâmica - Carta Argumentativa")

# --- Upload do arquivo Excel ---
st.sidebar.header("📁 Upload de Arquivos")
excel_file = st.sidebar.file_uploader("Arquivo Excel (.xlsx)", type=["xlsx"])

# Verifica se o arquivo de credenciais existe na raiz
CREDENTIALS_PATH = "credentials.json"
if os.path.exists(CREDENTIALS_PATH):
    creds = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
else:
    st.warning("⚠️ O arquivo 'credentials.json' não foi encontrado. Utilizando as credenciais do secrets.")
    CREDENTIALS_PATH = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    creds = service_account.Credentials.from_service_account_info(CREDENTIALS_PATH, scopes=SCOPES)

if excel_file:
    ARQUIVO_ORIGEM = excel_file
    ABA1 = 'Worksheet'
    ABA2 = 'Worksheet2'

    COLUNAS_ESPERADAS = [
        'Redação ID', 'Nome do Prompt', 'Prompt', 'Texto da Redação', 'Tema', 
        'Competência 1 - IA', 'Competência 1 - Humano', 'Divergencia Competência 1', 'Modulo Divergencia Competência 1',
        'Competência 2 - IA', 'Competência 2 - humano', 'Divergencia Competência 2', 'Modulo Divergencia Competência 2',
        'Competência 3 - IA', 'Competência 3 - Humano', 'Divergencia Competência 3', 'Modulo Divergencia Competência 3',
        'Competência 4 - IA', 'Competência 4 - Humano', 'Divergencia Competência 4', 'Modulo Divergencia Competência 4',
        'Competência 5 - IA', 'Competência 5 - Humano', 'Divergencia Competência 5', 'Modulo Divergencia Competência 5',
        'Nota - IA', 'Nota - Humano', 'Divergencia Nota', 'Modulo Divergencia Nota',
        'Feedback Competência 1', 'Feedback Competência 2', 'Feedback Competência 3', 'Feedback Competência 4',
        'Feedback Competência 5', 'Feedback Geral'
    ]

    with st.spinner("🔄 Lendo dados do Excel..."):
        df_combinado = load_excel_data(excel_file)

    # Inicializa os serviços do Google
    sheets_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    with st.spinner("📁 Criando planilha no Google Sheets..."):
        sheet_titles = [
            DATA_SHEET_TITLE,
            PIVOT_PROMPTS_SHEET_TITLE,
            PIVOT_IA_HU_SHEET_TITLE,
            PIVOT_DIV_COMP_SHEET_TITLE
        ]
        spreadsheet_id = create_spreadsheet_structure(sheets_service, SPREADSHEET_TITLE, sheet_titles)

        drive_service.permissions().create(
            fileId=spreadsheet_id,
            body={'role': 'writer', 'type': 'anyone'},
            fields='id'
        ).execute()

        values = [list(df_combinado.columns)] + df_combinado.values.tolist()
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f'{DATA_SHEET_TITLE}!A1',
            valueInputOption='RAW',
            body={'values': values}
        ).execute()

        spreadsheet_info = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet_info['sheets']
        sheet_id_dados = get_sheet_id(sheets, DATA_SHEET_TITLE)
        sheet_id_pivot_prompts = get_sheet_id(sheets, PIVOT_PROMPTS_SHEET_TITLE)
        sheet_id_pivot_ia_hu = get_sheet_id(sheets, PIVOT_IA_HU_SHEET_TITLE)
        sheet_id_pivot_div_comp = get_sheet_id(sheets, PIVOT_DIV_COMP_SHEET_TITLE)

    with st.spinner("📊 Criando Tabela Dinâmica - Divergência entre PROMPTS"):
        CreatingPivotTablePromptDivergence(sheet_id_dados, values, df_combinado, sheet_id_pivot_prompts, sheets_service, spreadsheet_id).creating_pivot_table_prompt_divergence()
    
    with st.spinner("📊 Criando Tabela Dinâmica - Divergência entre IA e HU"):
        CreatingPivotTableAiVsHuDivergence(sheet_id_dados, values, df_combinado, sheet_id_pivot_ia_hu, sheets_service, spreadsheet_id).creating_pivot_table_ai_hu_divergence()

    with st.spinner("📊 Criando Tabela Dinâmica - Divergência de cada Competência entre IA e HU"):
        CreatingPivotEachCompIaHuDivergence(sheet_id_dados, values, df_combinado, sheet_id_pivot_div_comp, sheets_service, spreadsheet_id).creating_pivot_each_comp_ia_hu_divergence()

    with st.spinner("🧮 Calculando divergências..."):
        CreatingPivotCalculatingDivergences(df_combinado, sheets_service, spreadsheet_id, DIVERGENCIA_SHEET_TITLE, DIVERGENCIA_TOTAL_SHEET_TITLE).creating_pivot_calculating_divergences()

    st.success("✅ Processo concluído com sucesso!")
    st.markdown(f"🔗 **[Acesse a planilha gerada aqui](https://docs.google.com/spreadsheets/d/{spreadsheet_id})**")

else:
    st.warning("⚠️ Envie o Excel para continuar.")
