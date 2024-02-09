import pandas as pd
import os

def createInsert(df, nomeTabela):
    colunas = df.columns
    query = "INSERT INTO " + nomeTabela + " ("

    for coluna in colunas:
        query += coluna
        query += ', '

    query = query[:-2]
    query += ')'
    query += '\nVALUES\n'

    return query

#lower == 1 deixa tudo em lower
#Lê todos as linhas e cria todos os valores a serem inseridos
def createValues(df, lower=0):
    valuesLine = '('
    for index in range(len(df)):
        raw_row = df.iloc[index]
        for tabIndex in range(len(raw_row)):
            raw_data = raw_row.iloc[tabIndex]
            if isinstance(raw_data, str):
                if(lower == 1):
                    raw_data = raw_data.lower()
            raw_data = str(raw_data)
            valuesLine = valuesLine + "'" + raw_data + "'" + ", "
        valuesLine = valuesLine[:-2]
        valuesLine += '),'
        valuesLine += '\n('
    valuesLine = valuesLine[:-3]
    valuesLine += ';'
    return valuesLine

#concatena todos os textos e retorna a query
def returnQuery(df, lower, nomeTabela):
    text = ''
    text += createInsert(df, nomeTabela)
    text += createValues(df, lower)
    return text
#Transoforma o xlsm para varios arquivos CSV para posteriormente serem lidos
#Facilita a minha vida e já tenha uma pasta chamada data pf
#retorna 1 se tudo estiver certo
def xlsmToCsv(path, diretorioSaida):
    if os.listdir(diretorioSaida):
        print(f"The folder '{diretorioSaida}' is not empty.")
        return 0

    excel_file = path
    all_sheets = pd.read_excel(excel_file, sheet_name=None)
    sheets = all_sheets.keys()

    #concatena as strings para escrever tudo em um diretorio só
    diretorioFinal = diretorioSaida + '%s.csv'

    for sheet_name in sheets:
        sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
        sheet.to_csv(diretorioFinal % sheet_name, index=False)

    return 1

def leTodosCSV(folder_path):
    files = sorted(os.listdir(folder_path))
    diretorios = []
    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        diretorios.append(file_path)
    return diretorios

def criaTxt(lower, pathXlsm, pathDiretorios = "data/"):
    xlsmToCsv(pathXlsm, pathDiretorios)
    diretorios = leTodosCSV(pathDiretorios)
    finalQuery = ''

    for diretorio in diretorios:

        df = pd.read_csv(diretorio, encoding= 'utf-8')

        last_part = diretorio.split('/')[-1]
        nomeTabela = last_part.split('.')[0]
        finalQuery = finalQuery + returnQuery(df, lower, nomeTabela) + '\n\n'

    file_path = 'query.txt'
    with open(file_path, 'w') as file:
        file.write(finalQuery)

    print(f"query Criada com sucesso {file_path}")