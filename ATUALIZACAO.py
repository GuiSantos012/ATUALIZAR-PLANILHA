import time
import os
import win32com.client
import pandas as pd
import urllib.parse
from PIL import ImageGrab


# inicia o excel
File = win32com.client.Dispatch("Excel.Application")

# 1-visivel/0-invisivel
File.Visible = 1

# Abre o Arquivo
Workbook = File.Workbooks.open(r"caminho do arquivo")

# Espera a tabela Atualizar
time.sleep(15)

# atualizar os Dados
Workbook.RefreshAll()
time.sleep(15)

Workbook.RefreshAll()
time.sleep(15)

# Lista de abas para capturar
abas = [
    "***, ***, ***, ***"
]

# Caminho da pasta para salvar os prints
caminho_pasta = r"C:\Users\Guilherme\...\PRINTS"
os.makedirs(caminho_pasta, exist_ok=True)

# Intervalo para capturar
intervalo = "B3:P26"

# Loop pelas abas
for aba in abas:
    try:
        ws = Workbook.Sheets(aba)
        ws.Activate()
        print(f"Capturando print da aba: {aba}...")

        # Copia o intervalo como imagem
        rng = ws.Range(intervalo)
        rng.CopyPicture(Appearance=1, Format=2)
        time.sleep(2)

        # Captura imagem da área de transferência
        img = ImageGrab.grabclipboard()

        if img:
            # Tira os caracteres invalidos
            nome_arquivo = aba.replace(".", "").replace("/", "-").replace("\\", "-")
            caminho_imagem = os.path.join(caminho_pasta, f"{nome_arquivo}.png")

            # Substitui o arquivo
            if os.path.exists(caminho_imagem):
                os.remove(caminho_imagem)
                print(f"Arquivo existente removido: {caminho_imagem}")

            # Salva o novo print
            img.save(caminho_imagem)
            print(f"Novo print salvo: {caminho_imagem}")

        else:
            print(f"Nenhuma imagem copiada na aba: {aba}. Verifique o intervalo.")

    except Exception as e:
        print(f"Erro ao capturar a aba '{aba}': {e}")

time.sleep(30)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# caminho da planilha:
CAMINHO_PLANILHA = r"C:\Users\Guilherme\...\CONTATOS.xlsx"

# Inicia Chrome (mantem o perfil ja logado)
chrome_options = Options()
chrome_options.add_argument("--user-data-dir=C:/ChromeProfile")  # mantém sessão
chrome_options.add_argument("--profile-directory=Default")

service = Service()
navegador = webdriver.Chrome(service=service, options=chrome_options)
navegador.get("https://web.whatsapp.com")

time.sleep(10)

# Carregar planilha
tabela = pd.read_excel(CAMINHO_PLANILHA)

for linha in tabela.index:

    nome = tabela.loc[linha, "NOME"]
    mensagem = tabela.loc[linha, "MENSAGEM"]
    arquivo = tabela.loc[linha, "ARQUIVO"]
    telefone = tabela.loc[linha, "CONTATO"]

    texto = mensagem.replace("fulano", nome)
    texto_url = urllib.parse.quote(texto)

    link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto_url}"
    navegador.get(link)

    time.sleep(8)

    # botão de enviar mensagem
    try:
        botao_enviar = navegador.find_element(
            By.XPATH,
            '//*[@id="main"]/footer/div[1]/div/span/div/div/div/div[4]/div/span/button/div/div/div[1]/span'
        )
        botao_enviar.click()
        time.sleep(2)
    except:
        print(f"Erro ao enviar texto para {nome}")

    # Enviar ARQUIVO
    if arquivo != "N":

        caminho_arquivo = os.path.abspath(f"PRINTS/{arquivo}")

        # Clicar no botão +
        botao_clip = navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div/div/div[1]/div/span/button/div/div/div[1]/span')
        botao_clip.click()

        time.sleep(1)

        # Input da imagem
        input_file = navegador.find_element(By.XPATH, '//*[@id="main"]/div[2]/div/div/div[2]/div[2]/div[3]/div/div/div/div/div/div/div[1]/div[1]/div[1]/div/input')
        input_file.send_keys(caminho_arquivo)

        time.sleep(2)

        # Botão de enviar arquivo
        botao_enviar_arquivo = navegador.find_element(
            By.XPATH,
            '//*[@id="app"]/div/div/div[3]/div/div[3]/div[2]/div/span/div/div/div/div[2]/div/div[2]/div[2]/span/div/div/span'
        )
        botao_enviar_arquivo.click()

        time.sleep(3)

