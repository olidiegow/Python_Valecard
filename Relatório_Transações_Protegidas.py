import pandas as pd
import win32com.client as win32
import time
import datetime
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import os
import shutil
import openpyxl

# Formatos de Data
hoje = datetime.date.today()
hoje_formatado = hoje.strftime('%d/%m/%Y')
amanha = hoje + datetime.timedelta(days=1)
amanha_formatado = amanha.strftime('%d/%m/%Y')
ontem = hoje - datetime.timedelta(days=1)
ontem_formatado = ontem.strftime('%d/%m/%Y')
ultima_semana = hoje - datetime.timedelta(days=7)
semana_formatada = ultima_semana.strftime('%d/%m/%Y')


# Inicializa o navegador Chrome
navegador = webdriver.Chrome()
navegador.maximize_window()

# Abre a página de login
navegador.get('https://siag.valecard.com.br/frota/pages/start.jsf')

# credenciais de usuário
usuario = navegador.find_element(By.XPATH, '//*[@id="wrap-geral"]/div[2]/div/div/ul/li[2]/input')
senha = navegador.find_element(By.XPATH, '//*[@id="wrap-geral"]/div[2]/div/div/ul/li[3]/input')
dominio = navegador.find_element(By.XPATH, '//*[@id="wrap-geral"]/div[2]/div/div/ul/li[4]/select')
usuario.send_keys('935119')
senha.send_keys('10213')
dominio.send_keys('cliente')

# Submeter as Credenciais
botao_login = navegador.find_element(By.XPATH, '//*[@id="wrap-geral"]/div[2]/div/div/ul/li[5]/input')
botao_login.click()
time.sleep(5)

formulario = navegador.find_element(By.ID, "MENU_FORM_HADOUKEN")

botao_abastecimento = navegador.find_element(By.ID, 'MENU_FORM_HADOUKEN:j_id29')
botao_abastecimento.click()
time.sleep(3)

# Acesar a aba Relatórios
aba_relatorios = 'https://siag.valecard.com.br/frota/pages/reports.jsf'
navegador.get(aba_relatorios)
time.sleep(3)

# Escolher Relatorio
transacao_negada = navegador.find_element(By.XPATH, '//*[@id="form:panel_body"]/table/tbody/tr/td[2]/select')
transacao_negada.send_keys('Transação Negada', Keys.ENTER)
time.sleep(5)

# Definir Parametros do Relatório
data_inicio = navegador.find_elements(By.XPATH, '//*[@id="form:c1InputDate"]')
for element in data_inicio:
    element.send_keys(ontem_formatado)
time.sleep(5)

data_fim = navegador.find_elements(By.XPATH, '//*[@id="form:c2InputDate"]')
for element in data_fim:
    element.send_keys(ontem_formatado)
time.sleep(5)

# Gerar Relatório
botao_pesquisar = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:j_id994"]')))
navegador.execute_script(botao_pesquisar.get_attribute('onclick'))
time.sleep(20)

# Exportar Relatório
botao_csv = WebDriverWait(navegador, 25).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:panel_body"]/div[3]/div[2]/input[2]')))
action = ActionChains(navegador)
action.click(botao_csv).perform()
time.sleep(10)

pasta_download = 'C:\\Users\\diego.oliveira\\Downloads'
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]

planilha = openpyxl.load_workbook('C:\\temp\\Base_CC.xlsx')
aba = planilha.active

for coluna in aba.iter_rows(min_row=2, values_only=True):
    ccusto = coluna[0]
    email_frota = coluna[2]
    email_frota1 = coluna[3]
    email_frota2 = coluna[4]
    email_frota3 = coluna[5]
    email_gerente = coluna[6]
    email_operacao = coluna[7]

    dataframe = pd.read_csv(ultimo_download, encoding='latin1', delimiter=';')
    centro_custo = dataframe['Centro de Custo']

    filtro = dataframe['Centro de Custo'].str.contains(ccusto)
    filtrados = dataframe[filtro]
    dados_copiados = filtrados.copy()

    # Criar uma instância do objeto Outlook
    outlook = win32.Dispatch('Outlook.Application')

    tabela_formatada = dados_copiados.to_html(index=False)
    dados_copiados.to_csv('dados_filtrados.csv', index=False)
    texto_email = "Olá," \
                  " " \
                  "\n\nSegue abaixo a tabela com as transações protegidas:\n\n" \
                  " " \
                  " "
    corpo_email = f"{texto_email}{tabela_formatada}"

    # Criar um objeto de e-mail
    email = outlook.CreateItem(0)
    email.Subject = f"{'Relatório de Transações Bloqueadas de: '}{ontem_formatado}"
    email.HTMLBody = f"<html><body>{corpo_email}</body></html>"
    email.SentOnBehalfOfName = 'diego.oliveira@endicon.com.br'
    email.To = "diego.oliveira@endicon.com.br"
    #email.CC = f"{email_frota}{email_frota2}{email_frota1}{email_frota3}{email_gerente}"
    email.Send()





