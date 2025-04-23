class CreatingPivotEachCompIaHuDivergence:
    def __init__(self, sheet_id_dados, values, df_combinado, sheet_id_pivot_div_comp, sheets_service, spreadsheet_id) :
        self.sheet_id_dados = sheet_id_dados
        self.values = values
        self.df_combinado = df_combinado
        self.sheet_id_pivot_div_comp = sheet_id_pivot_div_comp
        self.sheets_service = sheets_service
        self.spreadsheet_id = spreadsheet_id

    def creating_pivot_each_comp_ia_hu_divergence(self):
        requests = [
            # Primeira Tabela - Módulo Divergência Competência 1
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": self.sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(self.df_combinado),
                                    "endColumnIndex": len(self.df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 8,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 8},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 8}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(self.df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": self.sheet_id_pivot_div_comp, "rowIndex": 2, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            },

            # Segunda Tabela - Módulo Divergência Competência 2
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": self.sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(self.df_combinado),
                                    "endColumnIndex": len(self.df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 12,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 12},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 12}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(self.df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": self.sheet_id_pivot_div_comp, "rowIndex": 2, "columnIndex": 4},
                    "fields": "pivotTable"
                }
            },

            # Terceira Tabela - Módulo Divergência Competência 3
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": self.sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(self.df_combinado),
                                    "endColumnIndex": len(self.df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 16,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 16},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 16}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(self.df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": self.sheet_id_pivot_div_comp, "rowIndex": 11, "columnIndex": 0},
                    "fields": "pivotTable"
                }
            },

            # Quarta Tabela - Módulo Divergência Competência 4
            {
                "updateCells": {
                    "rows": [{
                        "values": [{
                            "pivotTable": {
                                "source": {
                                    "sheetId": self.sheet_id_dados,
                                    "startRowIndex": 0,
                                    "startColumnIndex": 0,
                                    "endRowIndex": len(self.df_combinado),
                                    "endColumnIndex": len(self.df_combinado.columns)
                                },
                                "rows": [{
                                    "sourceColumnOffset": 20,  # Coluna usada como linha
                                    "showTotals": True,
                                    "sortOrder": "ASCENDING"
                                }],
                                "columns": [],
                                "values": [
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 20},
                                    {"summarizeFunction": "COUNT", "sourceColumnOffset": 20}
                                ],
                                "criteria": {
                                    "1": {
                                        "visibleValues": list(self.df_combinado['Nome do Prompt'].unique())
                                    }
                                }
                            }
                        }]
                    }],
                    "start": {"sheetId": self.sheet_id_pivot_div_comp, "rowIndex": 11, "columnIndex": 4},
                    "fields": "pivotTable"
                }
            }
        ]

        self.sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body={"requests": requests}
        ).execute()
        