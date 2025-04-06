from kivy.app import App
from kivy.lang import Builder  # GUI
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib
import urllib.parse
from selenium.webdriver.common.by import By
import ids
import requests
from io import BytesIO

url = f"https://onedrive.live.com/download?resid={ids.resid}&authkey={ids.authkey}"

response = requests.get(url)
# print(response.status_code)  # Verifica o status HTTP
# print(response.headers['Content-Type'])  # Verifica o tipo de conteúdo da resposta
response.raise_for_status()  # Garante que a requisição foi bem-sucedida

# Carregar os dados no Pandas
df = pd.read_excel(BytesIO(response.content), engine="openpyxl", sheet_name="Chamada2025")

response = requests.get(url)
response.raise_for_status()  # Garante que a requisição foi bem-sucedida

df2 = pd.read_excel(BytesIO(response.content), engine="openpyxl", sheet_name="Cadastro")

url_escrita = f"https://docs.google.com/spreadsheets/d/{ids.sheet_id_escrita}/edit?gid=0#gid=0"

with BytesIO(response.content) as file:
    xls = pd.ExcelFile(file, engine="openpyxl")
    # print("Abas disponíveis:", xls.sheet_names)


class Aluno:

    @staticmethod
    def contar_alunos():
        """Conta o número de alunos na planilha na coluna C a partir da linha 8.
               - Na primeira célula vazia, o sistema não conta mais os alunos que venham depois."""

        i = 7
        n = 0
        for nome in df.iloc[i:, 2]:
            if pd.notna(nome):
                i += 1
                n += 1
        return n

    @staticmethod
    def conta_dias_aulas(procurado):
        """Conta o número de dias em que o parâmetro (que obrigatóriamente nessa versão é x) aparece em cada coluna
           após a coluna H."""
        i = 0
        for coluna in df.columns[7:]:  # Colunas de H em diante
            if df[coluna].eq(procurado).any():
                i += 1
        return i

    @staticmethod
    def contar_faltas_seguidas(n_aulas):
        """Vai contar quantas células vazias seguidas têm antes da última coluna e associar o número de faltas com o
            nome do aluno.
                   - O método enumerate vai numerar a partir da linha 7 os nomes dos alunos.
                   - A estrutura de repetição interna será responsável por contar as faltas seguidas da direita para
                     esquerda.
                   - O retorno é uma lista do número de faltas de todos os alunos."""
        lista_alunos_faltas = []

        for i, nome in enumerate(df.iloc[6:, 2], start=6):
            if pd.notna(nome):
                faltas = 0
                encontrou_presenca = False

                for aula in reversed(df.iloc[i, 7:7 + n_aulas]):
                    if pd.notna(aula):
                        encontrou_presenca = True
                        break
                    else:
                        if not encontrou_presenca:
                            faltas += 1

                desistencia = df.iloc[i, 6]

                # Adiciona o resultado para o aluno
                lista_alunos_faltas.append({'nome': nome, 'n_faltas': faltas, 'telefone': '', 'desistencia': desistencia})
                # É colocado o telefone em outra função.

        return lista_alunos_faltas

    @staticmethod
    def listar_faltantes(alunos_faltas):
        """Será criada uma lista de dicionários com o nome do aluno, número de faltas e número de telefone para todos
           aqueles alunos que utrapassaram o limite de faltas. (4 faltas consecutivas no critério atual)."""

        lista_faltantes = []

        for aluno in alunos_faltas:
            if aluno['n_faltas'] > 3:
                nome = aluno['nome']
                faltas = aluno['n_faltas']
                desistencia = aluno['desistencia']

                if pd.isna(aluno['desistencia']):
                    lista_faltantes.append({'nome': nome, 'n_faltas': faltas, 'desistencia': desistencia})

        return lista_faltantes

    @staticmethod
    def adicionar_telefone(lista_alunos):
        for aluno in lista_alunos:
            aluno_nome = aluno['nome']
            telefone = None

            for i, nome in df2.iloc[:, 19].items():
                if nome == aluno_nome:
                    telefone = df2.iloc[i, 77]
                    break

            aluno['telefone'] = telefone

        return lista_alunos

    @staticmethod
    def armazenar_faltantes(alunos_faltantes):
        """É armazenado na variavel histórico os alunos que receberam mensagem. É uma função incompleta que eu deixei por
        questão de escalabilidade."""

        historico = []

        for aluno in alunos_faltantes:
            historico.append({aluno['nome'], aluno['n_faltas'], aluno['telefone']})  # colocar o telefone

        return historico

    @staticmethod
    def enviar_mensagem(alunos_faltantes):
        """Aqui será realizado o loop para o envio das mensagens"""

        # Inicializa o navegador
        navegador = webdriver.Chrome()

        # Abre a página do WhatsApp Web
        navegador.get("https://web.whatsapp.com/")

        # Verificação se estamos na página do WhatsApp, id="side" é o da barra lateral de conversas
        while len(navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)

        for aluno in alunos_faltantes:

            telefone = aluno.get('telefone')

            if not telefone:
                # Se o telefone é None, string vazia, ou 0, ignorar e passar para o próximo aluno
                continue

            try:

                # Define o número de telefone e a mensagem
                telefone = aluno['telefone']

                nome = aluno['nome'].split(" ")[0]
                mensagem = f"""
Sentimos sua falta!

Bom dia/Boa tarde, {nome},

Esperamos que você esteja bem.

Notamos que sua frequência no cursinho tem sido baixa e gostaríamos de saber se há algo com o qual  possamos ajudar. Entendemos que imprevistos acontecem e estamos aqui para oferecer suporte, seja ele acadêmico ou até mesmo pessoal.

Sua presença é muito importante para nós, pois acreditamos em você e na sua capacidade de conseguir alcançar seus sonhos e objetivos. 

Se houver algum problema ou dificuldade que você esteja enfrentando, por favor, não deixe de nos contatar. Estamos disponíveis para conversar e encontrar soluções que possam facilitar sua participação nas aulas.

Aguardamos o seu retorno e desejamos que você possa retomar suas atividades conosco o mais breve possível.

Um abraço,

Cursinho Comunitário Bonsucesso.
                """

                # Codifica a mensagem para URL encoding
                texto = urllib.parse.quote(mensagem)
                link = f"{telefone}?text={texto}"

                # Abrir o link com a mensagem
                navegador.get(link)

                while len(navegador.find_elements(By.ID, "action-button")) < 1:
                    time.sleep(1)

                navegador.find_element(By.ID, "action-button").click()

                while len(navegador.find_elements(By.LINK_TEXT, "usar o WhatsApp Web")) < 1:
                    time.sleep(1)

                navegador.find_element(By.LINK_TEXT, "usar o WhatsApp Web").click()

                # Verificação se a conversa foi aberta, id="main" é o da área principal de conversas
                while len(navegador.find_elements(By.ID, "main")) < 1:
                    time.sleep(1)

                # Enviar a mensagem
                campo_de_mensagem = navegador.find_element(By.XPATH,
                                                           '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/'
                                                           'div/div/p')
                campo_de_mensagem.send_keys(Keys.ENTER)

                # Mantém o navegador aberto por um tempo para garantir que a mensagem seja enviada
                time.sleep(15)

            except Exception as e:
                # print(f"Erro ao enviar mensagem para o telefone {telefone}: {e}")
                pass

        # Fecha o navegador
        navegador.quit()


# Main
def main():

    # Número de dias de aula:
    n_aulas = Aluno.conta_dias_aulas("x")

    # Total de alunos
    # n_alunos = Aluno.contar_alunos()

    lista_alunos = Aluno.contar_faltas_seguidas(n_aulas)
    # alunos_faltas = Aluno.adicionar_telefone(lista_alunos)

    alunos_faltantes = Aluno.listar_faltantes(lista_alunos)
    alunos_faltantes = Aluno.adicionar_telefone(alunos_faltantes)

    # print(alunos_faltantes)
    Aluno.enviar_mensagem(alunos_faltantes)


GUI = Builder.load_file('interface.kv')


class AplicativoSecretaria(App):
    def build(self):
        return GUI

    def disparar_main(self):
        main()


AplicativoSecretaria().run()

"""
# Código para chamada no google

# Carregar aba "Chamada2024"
url = f"https://docs.google.com/spreadsheets/d/{ids.sheet_id}/export?format=csv&gid={ids.gid}"
df = pd.read_csv(url)

# Carregar aba "Cadastro"
url2 = f"https://docs.google.com/spreadsheets/d/{ids.sheet_id}/export?format=csv&gid={ids.gid2}"
df2 = pd.read_csv(url2)
"""
