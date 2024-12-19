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
        if len (data_str) == 8:
            try:
                dia = int(data_str [:2])
                mes = int(data_str [2:4])
                ano = int(data_str [4:])
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

# Função para extrair os dados e formatar conforme o modelo solicitado
def extract_financing_data(sheet):
    financing_data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):  # Começa na linha 2 para ignorar o cabeçalho
        ficha = {
            "DADOS PESSOAIS": {
                "NOME": row[1],
                "RG": int(row[2]),
                "SSP": row[3],
                "DATA DE EXPEDIÇÃO": formatar_data(row[4]),
                "CPF": int(row[5]),
                "NASCIMENTO": formatar_data(row[6]),
                "NATURALIDADE": row[7],
                "ESTADO CIVIL": row[8],
                "NOME PAI": row[9],
                "NOME MÃE": row[10],
                "ENDEREÇO": row[13],
                "N°": int(row[14]),
                "BAIRRO": row[15],
                "CEP": row[17],
                "CIDADE": row[18],
                "ESTADO": row[19],
                "FONE CELULAR": row[11],
                "E-MAIL": row[12],
            },
            "DADOS PROFISSIONAIS": {
                "NOME DA EMPRESA": row[21],
                "ENDEREÇO": row[22],
                "FONE": row[23],
                "CARGO": row[24],
                "RENDA": row[25],
            },
            "OUTRAS INFORMAÇÕES": {
                "POSSUI CNH?": row[26],
                "MODELO DA MOTO": row[27],
                "ENTRADA": row[29],
            },
        }
        financing_data.append(ficha)

    return financing_data

# Extrair os dados
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