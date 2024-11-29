import pandas as pd
import os
import locale
import logging
import configparser
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep

# Configurações de logging
log_path = os.path.join(os.path.expanduser(
    "~"), "Desktop", "log_erro_aniversariantes.log")
logging.basicConfig(filename=log_path, level=logging.ERROR,
                    format='%(asctime)s %(levelname)s %(message)s', filemode='a')
logger = logging.getLogger()
# Define a localidade para português do Brasil
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

# Caminho da pasta "Anivesarios" no Desktop
pastaAniversarios = os.path.join(
    os.path.expanduser("~"), "Desktop", "Aniversarios")

# Caminho do arquivo de configuração do perfil do Chrome "config.ini"
configPath = os.path.join(pastaAniversarios, "config.ini")
if not os.path.exists(configPath):
    logging.error(
        f'Arquivo "config.ini" não foi encontrado na pasta "Aniversarios"\n')
    raise FileNotFoundError(
        f'O arquivo "config.ini" é obrigatório e deve estar na pasta correta(Desktop >> Aniversarios).')
else:
    # Carregando o arquivo de configuração
    config = configparser.ConfigParser()
    config.read(configPath)

# Definição do perfil do Chrome a ser utilizado #
try:
    profilePath = config['chrome']['profile_path']
except Exception as e:
    logging.error(
        "A seção 'chrome' ou a chave 'profile_path' não foi encontrada no arquivo config.ini.\n")
    raise KeyError(
        "A configuração para 'profile_path' é obrigatória no arquivo config.ini.")

# Definição da planilha de aniversariantes #
try:
    # Define o caminho da planilha na pasta "Aniversarios" na área de trabalho
    caminho_planilha_aniversariantes = os.path.join(
        pastaAniversarios, "aniversariantes.xlsx")
    # Carrega a planilha Excel dos aniversariantes
    df_aniversariantes = pd.read_excel(caminho_planilha_aniversariantes)
except Exception as e:
    logging.error(
        f'Arquivo "aniversariantes.xlsx não encontrado, verifique e tente novamente.\n"')
    raise e

# Definição da planilha contendo o telefone do diretor #
try:
    # Define o caminho da planilha na mesma pasta, mas em uma planilha diferente
    caminho_planilha_diretor = os.path.join(pastaAniversarios, "diretor.xlsx")
    # Carrega a planilha Excel do diretor
    df_diretor = pd.read_excel(caminho_planilha_diretor)
except Exception as e:
    logging.error(
        f'Não foi possivel localizar o arquivo "Diretor.xlsx", tente novamente.\n')
    raise e

# Obtém a data de hoje como um objeto datetime
hoje = datetime.today()


# Função para formatar a data e dia da semana em português
def formatar_data(data):
    return data.strftime("%d/%m"), data.strftime("%A")


# Função para gerar a mensagem para os aniversariantes
def gerar_mensagem(aniversariantes, data, dia_semana, saudacao=False):
    mensagem = ''
    if aniversariantes.empty:
        return f'Sem aniversariantes para data de hoje: {data}\n'
    else:
        if saudacao:
            mensagem += f'Ola Diretor!\n'
        mensagem += f"🎉Aniversariantes de {data}, {dia_semana}🎉\n"
        for _, row in aniversariantes.iterrows():
            mensagem += f"🎈{row['Nome de guerra']} (contato: {row['Telefone']}; lotação: {row['Lotacao']})\n\n"
        return mensagem


# Função para formatar a mensagem, substituindo o '\n' pelo simbolo de quebra de linha equivalente ao Whatsapp
def formataMensagem(mensagem):
    mensagemFormatada = mensagem.replace("\n", "%0A")

    return mensagemFormatada


# Filtra os aniversariantes de hoje
aniversariantes_hoje = df_aniversariantes[(
    df_aniversariantes['Dia'] == hoje.day) & (df_aniversariantes['Mês'] == hoje.month)]

# Formata a data de hoje
data_hoje, dia_semana_hoje = formatar_data(hoje)

# Gera a mensagem de hoje
mensagem_hoje = gerar_mensagem(
    aniversariantes_hoje, data_hoje, dia_semana_hoje, saudacao=True)

# Caso seja sexta-feira, inclui aniversariantes de sábado e domingo
if dia_semana_hoje.lower() == 'sexta-feira':
    sabado = hoje + timedelta(days=1)
    domingo = hoje + timedelta(days=2)

    aniversariantes_sabado = df_aniversariantes[(df_aniversariantes['Dia'] == sabado.day) & (
        df_aniversariantes['Mês'] == sabado.month)]
    aniversariantes_domingo = df_aniversariantes[(df_aniversariantes['Dia'] == domingo.day) & (
        df_aniversariantes['Mês'] == domingo.month)]

    if not aniversariantes_sabado.empty:
        data_sabado, _ = formatar_data(sabado)
        mensagem_hoje += gerar_mensagem(aniversariantes_sabado,
                                        data_sabado, "sábado")

    if not aniversariantes_domingo.empty:
        data_domingo, _ = formatar_data(domingo)
        mensagem_hoje += gerar_mensagem(aniversariantes_domingo,
                                        data_domingo, "domingo")

    if aniversariantes_hoje.empty and aniversariantes_sabado.empty and aniversariantes_domingo.empty:
        mensagem_hoje = f'\nSEM ANIVERSARIANTES {data_hoje}'
# Define o caminho do arquivo de texto na área de trabalho
caminho_arquivo = os.path.join(os.path.expanduser(
    "~"), "Desktop", "aniversariantes.txt")

# Adiciona a mensagem ao arquivo, sem sobrescrever
with open(caminho_arquivo, 'a', encoding='utf-8') as file:
    file.write(f"\n{'-'*50}\n")  # Separador visual
    file.write(mensagem_hoje)


# ------------------------- Envio automatico da mensagem ------------------------- #
try:
    if aniversariantes_hoje.empty and aniversariantes_sabado.empty and aniversariantes_domingo.empty:
        logging.info(
            f'Nem um aniversariante encontrado para as datas: {data_hoje, data_sabado, data_domingo}.\n')
    else:
        try:
            os.system('taskkill /f /im chrome.exe')
            sleep(2.5)
        except Exception as e:
            logging.info(
                f'Não foi possível finalizar as intancias do Chrome ou não existiam intancias abertas.\n')
            raise e
        # Configs do ChromeDriver
        try:
            chromeOptions = Options()
            chromeOptions.add_argument(f'user-data-dir={profilePath}')
            driver = webdriver.Chrome(options=chromeOptions)
            driver.maximize_window()
            driver.set_page_load_timeout(300)
            chromeOptions.add_argument('--no-sandbox')
            chromeOptions.add_argument('--disable-dev-shm-usage')
            chromeOptions.add_argument('--disable-gpu')
            chromeOptions.add_argument('--disable-extensions')
            chromeOptions.add_argument('--proxy-server="direct://"')
            chromeOptions.add_argument('--proxy-bypass-list=*')
            chromeOptions.add_argument('--start-maximized')
        except Exception as e:
            logging.error(f'Ero ao configurar perfil de usuário: {str(e)}')
            raise e
        driver.get('https://web.whatsapp.com')
        sleep(5)
        mensagem_hoje_formatada = formataMensagem(mensagem_hoje)
        telDiretor = df_diretor.loc[0, 'Diretor']
        url_whatsapp = f"https://web.whatsapp.com/send?phone={telDiretor}&text={mensagem_hoje_formatada}"
        driver.get(url_whatsapp)
        sleep(5)
        try:
            enviar = driver.find_element(
                "xpath", '//button[@aria-label="Enviar"]')
            enviar.click()
            logging.info(
                f'Automação concluida com sucesso. Mensagem enviada ao Diretor, no numero: {telDiretor}.\nConfira o arquivo de log na área de trabalho para mais informações.\n')
        except Exception as e:
            logging.error(
                f'Falha ao encontrar o botão de enviar mensagem. ERRO:{str(e)}')
            raise e
        sleep(5)
        driver.quit()
except Exception as e:
    logging.error(f'Erro no envio da mensagem.\nErro registrado: {str(e)}')
