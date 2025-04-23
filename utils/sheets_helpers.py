def get_sheet_id(sheets, title):
    """
    Retorna o ID da aba de planilha com o título especificado.

    Parâmetros:
    - sheets: lista de abas retornada por sheets_service.spreadsheets().get().
    - title: título da aba desejada.

    Retorna:
    - sheetId correspondente à aba.
    """
    return next(s['properties']['sheetId'] for s in sheets if s['properties']['title'] == title)

def create_spreadsheet_structure(sheets_service, spreadsheet_title, sheet_titles):
    spreadsheet_body = {
        'properties': {'title': spreadsheet_title},
        'sheets': [{'properties': {'title': title}} for title in sheet_titles]
    }
    spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
    return spreadsheet['spreadsheetId']
