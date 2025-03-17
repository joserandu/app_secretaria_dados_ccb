import os.path
import numpy as np
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import pandas as pd
from datetime import datetime
import ids
import requests
from io import BytesIO

url = f"https://onedrive.live.com/download?id={ids.id}&resid={ids.resid}&authkey={ids.authkey}"

response = requests.get(url)
response.raise_for_status()  # Garante que a requisição foi bem-sucedida

# Carregar aba "Chamada2024"
df = pd.read_excel(BytesIO(response.content), engine="openpyxl", sheet_name="Chamada2025")

url_escrita = f"https://docs.google.com/spreadsheets/d/{ids.sheet_id_escrita}/edit?gid=0#gid=0"


def contar_alunos():
    """Conta o número de alunos na planilha na coluna C a partir da linha 8.
           - Na primeira célula vazia, o sistema não conta mais os alunos que venham depois."""
    contagem = 0

    for nome in df.iloc[6:, 2]:  # Começa na linha 8 (índice 7) \\ troquei pra 6 pq são 60 nomes até agora
        if pd.isna(nome):  # Para ao encontrar a primeira célula vazia
            break
        contagem += 1

    return contagem


def conta_dias_aulas(procurado):
    """Conta o número de dias em que o parâmetro (que obrigatóriamente nessa versão é x) aparece em cada coluna
    após a coluna H."""
    i = 0
    for coluna in df.columns[7:]:  # Colunas de H em diante
        if df[coluna].eq(procurado).any():
            i += 1
    return i


n_alunos = contar_alunos()
n_aulas = conta_dias_aulas('x')

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    print(contar_alunos())
    # Parte de configurações da API do sheets

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        n_alunos = contar_alunos()

        """
        print(f"Tamanho do df: {df.shape}")  # (n_linhas, n_colunas)
        print(f"Valores esperados para i: {list(range(6, n_alunos + 3))}")
        print(f"Valores esperados para j: {list(range(6 + n_aulas - 2, 6 + n_aulas + 1))}")
        """

        # Lógica para leitura e inscrição das planilhas
        for p in range(-3, 0):

            presencas = []
            dias_da_semana = []
            datas_str = []
            periodos = []
            nomes = []

            for i in range(6, n_alunos+6):  # 6 ao n_a + 3

                j = 7 + n_aulas + p

                presenca = df.iloc[i, j]
                nome = df.iloc[i, 2]
                data_periodo = df.iloc[2, j]

                # print(f"Tipo de data_periodo: {type(data_periodo)} | Valor: {data_periodo}")  # Depuração

                # Se já for datetime, formatamos corretamente antes de transformar em string
                if isinstance(data_periodo, datetime):
                    data_periodo = data_periodo.strftime("%d/%m/%Y")  # Agora está no padrão brasileiro

                try:
                    data_periodo = str(data_periodo)
                    if "-" in data_periodo:
                        data_periodo = data_periodo.replace("-", "/")
                except Exception as e:
                    print(f"Deu erro aqui: {e}")

                if not isinstance(data_periodo, str):
                    raise ValueError("A data informada não é uma string válida.")

                try:
                    # Pegando só a parte da data
                    # Se data_periodo já for um objeto datetime, apenas formatamos
                    if isinstance(data_periodo, datetime):
                        data = data_periodo  # Já é um datetime válido
                    else:
                        # Pegando só a parte da data caso venha com um período junto
                        if " " in data_periodo:
                            data_str = data_periodo.split(" ")[0]  # "01/02/2025"
                        else:
                            data_str = data_periodo

                        # Garantir que dia e mês tenham dois dígitos
                        partes_data = data_str.split("/")

                        if len(partes_data) == 3:
                            dia, mes, ano = partes_data
                            data_str_formatada = f"{int(dia):02d}/{int(mes):02d}/{ano}"
                        else:
                            raise ValueError(f"Formato de data inválido: {data_str}")

                        # Converter para datetime
                        data = datetime.strptime(data_str_formatada, "%d/%m/%Y")

                    # Agora já temos `data` corretamente, podemos pegar o dia da semana
                    dia_da_semana = data.strftime("%A")

                    # Pegando o período (removendo parênteses)
                    if isinstance(data_periodo, str) and "(" in data_periodo:
                        periodo = data_periodo.split("(")[-1].strip(") ")
                    else:
                        periodo = "manhã"

                    print(f"{i} | {presenca} | {dia_da_semana} | {data.strftime('%d/%m/%Y')} | {periodo} | {nome}")

                except Exception as e:
                    print(f"Erro {e}: Verifique se a data está no formato: dd/mm/aaaa (periodo) ou somente dd/mm/aaaa")

                    break

                presencas.append(presenca)
                dias_da_semana.append(dia_da_semana)
                datas_str.append(data_str)
                periodos.append(periodo)
                nomes.append(nome)

            # print(list(enumerate(nomes)))

            valores_add = [

            ]

            for i in range(0, len(presencas)):  # -3 porque está na terceira linha da planilha
                valores_add.append([
                    "" if isinstance(presencas[i], float) and np.isnan(presencas[i]) else presencas[i],
                    "" if isinstance(dias_da_semana[i], float) and np.isnan(dias_da_semana[i]) else dias_da_semana[i],
                    "" if isinstance(datas_str[i], float) and np.isnan(datas_str[i]) else datas_str[i],
                    "" if isinstance(periodos[i], float) and np.isnan(periodos[i]) else periodos[i],
                    "" if isinstance(nomes[i], float) and np.isnan(nomes[i]) else nomes[i]
                ])

            body = {'values': valores_add}

            # Descobrir a primeira linha vazia
            range_coluna_b = "B:B"  # Toda a coluna B
            result = sheet.values().get(spreadsheetId=ids.SAMPLE_SPREADSHEET_ID, range=range_coluna_b).execute()
            valores = result.get("values", [])

            primeira_linha_vazia = len(valores) + 1  # A próxima linha após a última preenchida

            print(f"A primeira linha vazia na coluna B é: {primeira_linha_vazia}")

            result = (
                sheet.values()
                .update(spreadsheetId=ids.SAMPLE_SPREADSHEET_ID,
                        range=f"A{primeira_linha_vazia}",
                        valueInputOption="USER_ENTERED",
                        body=body)
                .execute()
            )

    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()
