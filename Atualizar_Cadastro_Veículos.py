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

# Acessar aba abastecimento
formulario = navegador.find_element(By.ID, "MENU_FORM_HADOUKEN")
botao_abastecimento = navegador.find_element(By.ID, 'MENU_FORM_HADOUKEN:j_id29')
botao_abastecimento.click()
time.sleep(3)

# Abrir cadastro de veículos
aba_veiculos = 'https://siag.valecard.com.br/frota/pages/maintenanceVehicle.jsf'
navegador.get(aba_veiculos)
time.sleep(3)

# Acessar planilha de cadastro
workbook = openpyxl.load_workbook('C:\\temp\\atualizar_cadastro.xlsx')
planilha = workbook.active

for coluna in planilha.iter_rows(min_row=2, values_only=True):
    veiculo = coluna[0]
    filial = coluna[1]
    polo = coluna[2]
    centro_custo = coluna[3]

    pesquisar_placa = navegador.find_element(By.XPATH, '//*[@id="form:j_id211:vehicleFinderId"]')
    pesquisar_placa.send_keys(veiculo)
    pesquisar_placa.send_keys(Keys.TAB)
    time.sleep(3)

    abrir_pesquisa = navegador.find_element(By.XPATH, '//*[@id="form:j_id637"]')
    navegador.execute_script(abrir_pesquisa.get_attribute('onclick'))
    time.sleep(2)

    editar_veiculo = WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:panelTable:0:j_id639"]/input')))
    action = ActionChains(navegador)
    action.click(editar_veiculo).perform()
    time.sleep(2)

    escolher_filial = navegador.find_element(By.XPATH, '//*[@id="form:j_id675"]/table/tbody/tr/td/fieldset/table/tbody/tr/td[1]/table/tbody/tr[1]/td[2]/select')
    select = Select(escolher_filial)
    select.select_by_visible_text(filial)
    time.sleep(2)

    escolher_centro = navegador.find_element(By.XPATH, '//*[@id="form:j_id675"]/table/tbody/tr/td/fieldset/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/select')
    select = Select(escolher_centro)
    select.select_by_visible_text(centro_custo)
    time.sleep(2)

    escolher_polo = navegador.find_element(By.XPATH, '//*[@id="form:j_id675"]/table/tbody/tr/td/fieldset/table/tbody/tr/td[2]/table/tbody/tr[1]/td[2]/select')
    select = Select(escolher_polo)
    select.select_by_visible_text(polo)
    time.sleep(2)

    salvar_alteracoes = navegador.find_element(By.XPATH, '//*[@id="form:j_id675"]/table/tbody/tr/td/div/input')
    salvar_alteracoes.click()
    time.sleep(2)

    print(veiculo, " Atualizado com sucesso")

    pesquisar_placa = navegador.find_element(By.XPATH, '//*[@id="form:j_id211:vehicleFinderId"]')
    pesquisar_placa.clear()
    time.sleep(2)






