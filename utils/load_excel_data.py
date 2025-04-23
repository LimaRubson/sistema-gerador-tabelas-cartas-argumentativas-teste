import pandas as pd

def load_excel_data(file, sheet1='Worksheet', sheet2='Worksheet2'):
    """
    Carrega e combina os dados das duas abas do arquivo Excel.

    Parâmetros:
    - file: arquivo Excel carregado via st.file_uploader
    - sheet1: nome da primeira aba (default 'Worksheet')
    - sheet2: nome da segunda aba (default 'Worksheet2')

    Retorna:
    - df_combinado: DataFrame contendo os dados combinados e normalizados
    """
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

    df1 = pd.read_excel(file, sheet_name=sheet1, skiprows=0)
    df2 = pd.read_excel(file, sheet_name=sheet2, skiprows=1, header=None)
    df2.columns = df1.columns  # Usa os mesmos nomes da primeira aba

    df_combinado = pd.concat([df1, df2], ignore_index=True)
    df_combinado = df_combinado[COLUNAS_ESPERADAS].fillna("")

    return df_combinado
