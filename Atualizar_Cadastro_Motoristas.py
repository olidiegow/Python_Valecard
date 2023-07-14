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
aba_motoristas = 'https://siag.valecard.com.br/frota/pages/driver.jsf'
navegador.get(aba_motoristas)
time.sleep(3)

abrir_pesquisa = navegador.find_element(By.XPATH, '//*[@id="form:j_id268"]')
navegador.execute_script(abrir_pesquisa.get_attribute('onclick'))
time.sleep(5)

workbook = openpyxl.load_workbook('C:\\temp\\atualizar_cadastro.xlsx')
planilha = workbook.active


for coluna in planilha.iter_rows(min_row=2, values_only=True):
    motorista = coluna[0]
    filial = coluna[1]
    #polo = coluna[3]
    centro_custo = coluna[2]

# Limpar dados de pesquisa

    # Matricula
    matricula = navegador.find_element(By.XPATH, '//*[@id="form:matriculaId"]')
    matricula.clear()
    matricula.send_keys(Keys.TAB)
    time.sleep(1)

    # CPF
    cpf = navegador.find_element(By.XPATH, '//*[@id="form:cpf"]')
    cpf.clear()
    cpf.send_keys(Keys.TAB)
    time.sleep(1)

    # Nome motorista
    nome = navegador.find_element(By.XPATH, '//*[@id="form:panel_body"]/table[1]/tbody/tr[4]/td[2]/input')
    nome.clear()
    nome.send_keys(Keys.TAB)
    nome.send_keys(Keys.ENTER)
    time.sleep(1)

    navegador.get(aba_motoristas)
    time.sleep(3)

    abrir_pesquisa = WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:j_id268"]')))
    action = ActionChains(navegador)
    action.click(abrir_pesquisa).perform()
    time.sleep(2)

    pesquisar_motorista = navegador.find_element(By.XPATH, '//*[@id="form:matriculaId"]')
    pesquisar_motorista.send_keys(motorista)
    pesquisar_motorista.    send_keys(Keys.TAB)
    time.sleep(2)

    abrir_pesquisa = WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:j_id268"]')))
    action = ActionChains(navegador)
    action.click(abrir_pesquisa).perform()
    time.sleep(2)

    editar_motorista = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:table:0:j_id288"]/a')))
    action = ActionChains(navegador)
    action.click(editar_motorista).perform()
    time.sleep(2)

    escolher_filial = navegador.find_element(By.XPATH, '//*[@id="form:j_id203"]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/select')
    select_filial = Select(escolher_filial)
    select_filial.select_by_visible_text(filial)
    time.sleep(2)

    status_motorista = navegador.find_element(By.XPATH, '//*[@id="form:editStatus"]')
    select_status = Select(status_motorista)
    select_status.select_by_visible_text('Ativo')
    time.sleep(2)

    escolher_centro = navegador.find_element(By.XPATH, '//*[@id="form:j_id203"]/table/tbody/tr/td/table/tbody/tr[1]/td[4]/select')
    select_centro = Select(escolher_centro)
    select_centro.select_by_visible_text(centro_custo)
    time.sleep(2)

    #escolher_polo = navegador.find_element_by_xpath('//*[@id="form:j_id203"]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/select')
    #select = Select(escolher_polo)
    #select.select_by_visible_text(polo)
    #time.sleep(2)


    salvar_alteracoes = WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:j_id203"]/table/tbody/tr/td/div/input[1]')))
    action = ActionChains(navegador)
    action.click(salvar_alteracoes).perform()
    time.sleep(5)
    #action.click(salvar_alteracoes).perform()
    #time.sleep(2)

    voltar = navegador.find_element(By.XPATH, '//*[@id="form:j_id203"]/table/tbody/tr/td/div/input[2]')
    voltar.click()
    time.sleep(2)

    print(motorista, "atualizado com sucesso para: ", centro_custo)

    navegador.get(aba_motoristas)
    time.sleep(3)
    navegador.refresh()



















