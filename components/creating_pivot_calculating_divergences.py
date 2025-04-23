import pandas as pd


class CreatingPivotCalculatingDivergences:
    def __init__(self, df_combinado, sheets_service, spreadsheet_id, DIVERGENCIA_SHEET_TITLE, DIVERGENCIA_TOTAL_SHEET_TITLE) :
        self.df_combinado = df_combinado
        self.sheets_service = sheets_service
        self.spreadsheet_id = spreadsheet_id
        self.DIVERGENCIA_SHEET_TITLE = DIVERGENCIA_SHEET_TITLE
        self.DIVERGENCIA_TOTAL_SHEET_TITLE = DIVERGENCIA_TOTAL_SHEET_TITLE


    def creating_pivot_calculating_divergences(self):
        prompts = self.df_combinado['Nome do Prompt'].unique()
        df_dict = {p: self.df_combinado[self.df_combinado['Nome do Prompt'] == p].set_index('Redação ID') for p in prompts}

        PROMPT_GPT, PROMPT_NILMA = prompts[0], prompts[1]
        ids_comuns = df_dict[PROMPT_GPT].index.intersection(df_dict[PROMPT_NILMA].index)

        df_div = df_dict[PROMPT_GPT].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]] - \
                    df_dict[PROMPT_NILMA].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]]
        df_div = df_div.reset_index().round(2)

        self.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": self.DIVERGENCIA_SHEET_TITLE}}}]}
        ).execute()

        self.sheets_service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f'{self.DIVERGENCIA_SHEET_TITLE}!A1',
            valueInputOption='RAW',
            body={'values': [list(df_div.columns)] + df_div.values.tolist()}
        ).execute()

        df_total_gpt = df_dict[PROMPT_GPT].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]].sum(axis=1)
        df_total_nilma = df_dict[PROMPT_NILMA].loc[ids_comuns, ["Competência 1 - IA", "Competência 2 - IA", "Competência 3 - IA", "Competência 4 - IA"]].sum(axis=1)

        df_div_total = pd.DataFrame({
            'Redação ID': ids_comuns,
            f'Soma Total - {PROMPT_GPT}': df_total_gpt,
            f'Soma Total - {PROMPT_NILMA}': df_total_nilma,
            'Divergência (GPT - NILMA)': (df_total_gpt - df_total_nilma).round(2)
        }).reset_index(drop=True)

        self.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": self.DIVERGENCIA_TOTAL_SHEET_TITLE}}}]}
        ).execute()

        self.sheets_service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=f'{self.DIVERGENCIA_TOTAL_SHEET_TITLE}!A1',
            valueInputOption='RAW',
            body={'values': [list(df_div_total.columns)] + df_div_total.values.tolist()}
        ).execute()
