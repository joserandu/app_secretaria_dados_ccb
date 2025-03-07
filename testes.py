import pandas as pd
import ids
import requests
from io import BytesIO


url = f"https://onedrive.live.com/download?resid={ids.resid}&authkey={ids.authkey}"

response = requests.get(url)
response.raise_for_status()  # Garante que a requisição foi bem-sucedida

# Carregar os dados no Pandas
df = pd.read_excel(BytesIO(response.content), engine="openpyxl", sheet_name="Chamada2025")

# Carregar os dados no Pandas
df2 = pd.read_excel(BytesIO(response.content), engine="openpyxl", sheet_name="Cadastro")


url_escrita = f"https://docs.google.com/spreadsheets/d/{ids.sheet_id_escrita}/edit?gid=0#gid=0"
