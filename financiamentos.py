import openpyxl
from datetime import datetime

# Carregar a planilha
file_path = 'Financiamento (respostas).xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Formatando a data
def formatar_data(data):
    if isinstance(data, datetime):
        return data.strftime('%d/%m/%Y')
    elif isinstance(data, float) or isinstance(data, int):  # Caso Excel armazene data como número
        data_str = str(int(data))
        if len(data_str) == 8:
            try:
                dia = int(data_str[:2])
                mes = int(data_str[2:4])
                ano = int(data_str[4:])
                data_convertida = datetime(ano, mes, dia)
                return data_convertida.strftime('%d/%m/%Y')
            except ValueError:
                return data_str
    elif isinstance(data, str):
        try:
            data_convertida = datetime.strptime(data, '%Y/%m/%d')
            return data_convertida.strftime('%d/%m/%Y')
        except ValueError:
            return data
    return data

# Função para extrair os dados de uma linha específica
def get_data_by_row(sheet, row_index):
    """
    Função para retornar os dados de uma linha específica.
    :param sheet: A planilha carregada.
    :param row_index: O índice da linha (começando de 2 para ignorar o cabeçalho).
    :return: Dicionário com os dados da linha.
    """
    row = sheet[row_index]
    ficha = {
        "DADOS PESSOAIS": {
            "NOME": row[1].value,
            "RG": int(row[2].value),
            "SSP": row[3].value,
            "DATA DE EXPEDIÇÃO": formatar_data(row[4].value),
            "CPF": int(row[5].value),
            "NASCIMENTO": formatar_data(row[6].value),
            "NATURALIDADE": row[7].value,
            "ESTADO CIVIL": row[8].value,
            "NOME PAI": row[9].value,
            "NOME MÃE": row[10].value,
            "ENDEREÇO": row[13].value,
            "N°": int(row[14].value),
            "BAIRRO": row[15].value,
            "CEP": row[17].value,
            "CIDADE": row[18].value,
            "ESTADO": row[19].value,
            "FONE CELULAR": row[11].value,
            "E-MAIL": row[12].value,
        },
        "DADOS PROFISSIONAIS": {
            "NOME DA EMPRESA": row[21].value,
            "ENDEREÇO": row[22].value,
            "FONE": row[23].value,
            "CARGO": row[24].value,
            "RENDA": row[25].value,
        },
        "OUTRAS INFORMAÇÕES": {
            "POSSUI CNH?": row[26].value,
            "MODELO DA MOTO": row[27].value,
            "ENTRADA": row[29].value,
        },
    }
    return ficha

# Função para extrair os dados de todas as linhas
def extract_financing_data(sheet):
    financing_data = []

    for row_index in range (2,3):  # Linha onde serão extraídos os dados (range entre a linha selecionada e a próxima).
        ficha = get_data_by_row(sheet, row_index)
        financing_data.append(ficha)

    return financing_data

# Extrair todos os dados
formatted_data = extract_financing_data(sheet)

# Exibir os dados extraídos no formato solicitado
for ficha in formatted_data:
    print("FICHA PARA FINANCIAMENTO")
    print("\nDADOS PESSOAIS")
    for key, value in ficha["DADOS PESSOAIS"].items():
        print(f"- {key}: {value}")

    print("\nDADOS PROFISSIONAIS")
    for key, value in ficha["DADOS PROFISSIONAIS"].items():
        print(f"- {key}: {value}")

    print("\nOUTRAS INFORMAÇÕES")
    for key, value in ficha["OUTRAS INFORMAÇÕES"].items():
        print(f"- {key}: {value}")

    print("\n" + "="*20 + "\n")