
class CreatingPivotTableAiVsHuDivergence:
    def __init__(self, sheet_id_dados, values, df_combinado, sheet_id_pivot_ia_hu, sheets_service, spreadsheet_id) :
        self.sheet_id_dados = sheet_id_dados
        self.values = values
        self.df_combinado = df_combinado
        self.sheet_id_pivot_ia_hu = sheet_id_pivot_ia_hu
        self.sheets_service = sheets_service
        self.spreadsheet_id = spreadsheet_id

    def creating_pivot_table_ai_hu_divergence(self):
        requests = [{
            "updateCells": {
                "rows": [{
                    "values": [{
                        "pivotTable": {
                            "source": {
                                "sheetId": self.sheet_id_dados,
                                "startRowIndex": 0,
                                "startColumnIndex": 0,
                                "endRowIndex": len(self.values),
                                "endColumnIndex": len(self.values[0])
                            },
                            "rows": [{
                                "sourceColumnOffset": 0,
                                "showTotals": True,
                                "sortOrder": "ASCENDING"
                            }],
                            "columns": [{
                                "sourceColumnOffset": 1,
                                "showTotals": False,
                                "sortOrder": "ASCENDING"
                            }],
                            "values": [
                                {"summarizeFunction": "SUM", "sourceColumnOffset": 8},
                                {"summarizeFunction": "SUM", "sourceColumnOffset": 12},
                                {"summarizeFunction": "SUM", "sourceColumnOffset": 16},
                                {"summarizeFunction": "SUM", "sourceColumnOffset": 20}
                            ],
                            "criteria": {
                                "1": {
                                    "visibleValues": list(self.df_combinado['Nome do Prompt'].unique())
                                }
                            }
                        }
                    }]
                }],
                "start": {"sheetId": self.sheet_id_pivot_ia_hu, "rowIndex": 2, "columnIndex": 0},
                "fields": "pivotTable"
            }
        }]

        self.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": requests}
        ).execute()
