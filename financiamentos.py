import openpyxl
from datetime import datetime
import gspread
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials

# Carregar a planilha pelo link
# Definir o escopo de acesso (Google Sheets e Google Drive)
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Autenticar usando o arquivo JSON de credenciais
creds = Credentials.from_service_account_file("lugar/onde/o-arquivo/esta.json", scopes = SCOPES)
client = gspread.authorize(creds)

# Abrir a planilha pelo nome ou pela URL
sheet = client.open_by_key("ID-da-planilha").sheet1

# Puxar os dados da planilha
data = sheet.get_all_records()

# Formatando a data
def formatar_data(data):
    if isinstance(data, datetime):
        return data.strftime('%d/%m/%Y')
    elif isinstance(data, float) or isinstance(data, int) or isinstance(data, str):  # Caso Excel armazene data como número
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
    # Obter os valores da linha
    row = sheet.row_values(row_index)

    # Garantir que a linha tenha 30 colunas (completar com None se necessário)
    total_columns = 30  # Número total de colunas esperadas
    row.extend([None] * (total_columns - len(row)))  # Completar com None

    ficha = {
        "DADOS PESSOAIS": {
            "NOME": row[1],
            "RG": int(row[2]) if row[2] else "-",
            "SSP": row[3],
            "DATA DE EXPEDIÇÃO": formatar_data(row[4]) if row[4] else "-",
            "CPF": str(row[5]) if row[5] else "-",
            "NASCIMENTO": formatar_data(row[6]) if row[6] else "-",
            "NATURALIDADE": row[7],
            "ESTADO CIVIL": row[8],
            "NOME PAI": row[9] if row[9] else "-",
            "NOME MÃE": row[10],
            "ENDEREÇO": row[13],
            "N°": int(row[14]) if row[14] else "-",
            "BAIRRO": row[15],
            "COMPLEMENTO": row[16] if row[16] else "-",
            "CEP": row[17],
            "CIDADE": row[18],
            "ESTADO": row[19],
            "FONE CELULAR": row[11],
            "E-MAIL": row[12],
        },
        "DADOS PROFISSIONAIS": {
            "PROFISSÃO": row[20],
            "NOME DA EMPRESA": row[21],
            "ENDEREÇO": row[22],
            "FONE": row[23],
            "CARGO": row[24],
            "RENDA": row[25],
        },
        "OUTRAS INFORMAÇÕES": {
            "POSSUI CNH?": row[26],
            "MODELO DA MOTO": row[27],
            "ENTRADA": float(row[29]) if row[29] else 0.0,
        },
    }
    return ficha

# Função para extrair os dados de todas as linhas
def extract_financing_data(sheet):
    financing_data = []

    for row_index in range (2, 3):  # Linha onde serão extraídos os dados (range entre a linha selecionada e a próxima).
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
