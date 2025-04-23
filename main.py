import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os

# --- CONFIGURAÇÕES ---

# Arquivo Excel de origem
ARQUIVO_ORIGEM = 'base_de_dados.xlsx'
ABA1 = 'Worksheet'
ABA2 = 'Worksheet2'

# Colunas esperadas do arquivo de origem
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

# Arquivo de credenciais do Google Cloud
SERVICE_ACCOUNT_FILE = 'credentials.json'

# Escopos para Sheets + Drive (tornar público)
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SPREADSHEET_TITLE = 'Planilha com Tabela Dinâmica'
DATA_SHEET_TITLE = 'Dados'
PIVOT_SHEET_TITLE = 'Tabela Dinâmica'

# --- JUNTA AS ABAS DO EXCEL A PARTIR DA SEGUNDA LINHA ---
df1 = pd.read_excel(ARQUIVO_ORIGEM, sheet_name=ABA1, skiprows=0)
df2 = pd.read_excel(ARQUIVO_ORIGEM, sheet_name=ABA2, skiprows=1, header=None)  # Lê a partir da segunda linha
df2.columns = df1.columns  # Usa as mesmas colunas da primeira aba
df_combinado = pd.concat([df1, df2], ignore_index=True)

# Garante que todas as colunas esperadas estão presentes
df_combinado = df_combinado[COLUNAS_ESPERADAS]
df_combinado = df_combinado.fillna("")

# --- AUTENTICAÇÃO ---
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

sheets_service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

# --- CRIA PLANILHA ---
spreadsheet_body = {
    'properties': {'title': SPREADSHEET_TITLE},
    'sheets': [
        {'properties': {'title': DATA_SHEET_TITLE}},
        {'properties': {'title': PIVOT_SHEET_TITLE}},
    ]
}

spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
spreadsheet_id = spreadsheet['spreadsheetId']
print(f'✅ Planilha criada: https://docs.google.com/spreadsheets/d/{spreadsheet_id}')

# --- TORNA PLANILHA PÚBLICA ---
drive_service.permissions().create(
    fileId=spreadsheet_id,
    body={'role': 'writer', 'type': 'anyone'},
    fields='id'
).execute()

# --- ENVIA DADOS PARA ABA "DADOS" ---
values = [list(df_combinado.columns)] + df_combinado.values.tolist()

sheets_service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id,
    range=f'{DATA_SHEET_TITLE}!A1',
    valueInputOption='RAW',
    body={'values': values}
).execute()

# --- OBTÉM SHEET IDs ---
spreadsheet_info = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
sheets = spreadsheet_info['sheets']
sheet_id_dados = next(s['properties']['sheetId'] for s in sheets if s['properties']['title'] == DATA_SHEET_TITLE)
sheet_id_pivot = next(s['properties']['sheetId'] for s in sheets if s['properties']['title'] == PIVOT_SHEET_TITLE)

# --- CRIA TABELA DINÂMICA ---
requests = [
    {
        "updateCells": {
            "rows": [
                {
                    "values": [
                        {
                            "pivotTable": {
                                "source": {
                                    "sheetId": sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(values),
                                    "endColumnIndex": len(values[0])
                                },
                                "rows": [
                                    {
                                        "sourceColumnOffset": 0,
                                        "showTotals": True,
                                        "sortOrder": "ASCENDING"
                                    }
                                ],
                                "columns": [
                                    {
                                        "sourceColumnOffset": 1,
                                        "showTotals": False,
                                        "sortOrder": "ASCENDING"
                                    }
                                ],
                                "values": [
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 2},
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 3},
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 4},
                                    {"summarizeFunction": "SUM", "sourceColumnOffset": 5}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }
                    ]
                }
            ],
            "start": {
                "sheetId": sheet_id_pivot,
                "rowIndex": 2,
                "columnIndex": 0
            },
            "fields": "pivotTable"
        }
    }
]

sheets_service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={"requests": requests}
).execute()

print("✅ Tabela dinâmica criada com sucesso!")

# --- CÁLCULO DE DIVERGÊNCIA ENTRE PROMPTS POR REDAÇÃO ID ---
prompts = df_combinado['Nome do Prompt'].unique()
df_dict = {}
for prompt in prompts:
    df_dict[prompt] = df_combinado[df_combinado['Nome do Prompt'] == prompt].set_index('Redação ID')

PROMPT_GPT = prompts[0]
PROMPT_NILMA = prompts[1]
ids_comuns = df_dict[PROMPT_GPT].index.intersection(df_dict[PROMPT_NILMA].index)

df_div = df_dict[PROMPT_GPT].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]] - \
         df_dict[PROMPT_NILMA].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]]

df_div = df_div.reset_index()
df_div = df_div.round(2)

# --- ADICIONA ABA DE DIVERGÊNCIA POR COMPETÊNCIA ---
DIVERGENCIA_SHEET_TITLE = 'Divergência'

sheets_service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={ 
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": DIVERGENCIA_SHEET_TITLE
                    }
                }
            }
        ]
    }
).execute()

valores_divergencia = [list(df_div.columns)] + df_div.values.tolist()

sheets_service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id,
    range=f'{DIVERGENCIA_SHEET_TITLE}!A1',
    valueInputOption='RAW',
    body={'values': valores_divergencia}
).execute()

print("✅ Aba de divergência por competência criada com sucesso!")

# --- CÁLCULO DE DIVERGÊNCIA TOTAL ENTRE PROMPTS ---
df_total_gpt = df_dict[PROMPT_GPT].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]].sum(axis=1)
df_total_nilma = df_dict[PROMPT_NILMA].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]].sum(axis=1)

df_div_total = pd.DataFrame({
    'Redação ID': ids_comuns,
    f'Soma Total - {PROMPT_GPT}': df_total_gpt,
    f'Soma Total - {PROMPT_NILMA}': df_total_nilma,
    'Divergência (GPT - NILMA)': (df_total_gpt - df_total_nilma).round(2)
}).reset_index(drop=True)

# --- ADICIONA ABA DE DIVERGÊNCIA TOTAL ---
DIVERGENCIA_TOTAL_SHEET_TITLE = 'Divergência Total'

sheets_service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id,
    body={ 
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": DIVERGENCIA_TOTAL_SHEET_TITLE
                    }
                }
            }
        ]
    }
).execute()

valores_div_total = [list(df_div_total.columns)] + df_div_total.values.tolist()

sheets_service.spreadsheets().values().update(
    spreadsheetId=spreadsheet_id,
    range=f'{DIVERGENCIA_TOTAL_SHEET_TITLE}!A1',
    valueInputOption='RAW',
    body={'values': valores_div_total}
).execute()

print("✅ Aba de divergência total criada com sucesso!")
