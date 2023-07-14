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

# Formatos de Data
hoje = datetime.date.today()
hoje_formatado = hoje.strftime('%d/%m/%Y')
amanha = hoje + datetime.timedelta(days=1)
amanha_formatado = amanha.strftime('%d/%m/%Y')
ontem = hoje - datetime.timedelta(days=1)
ontem_formatado = ontem.strftime('%d/%m/%Y')
ultima_semana = hoje - datetime.timedelta(days=7)
semana_formatada = ultima_semana.strftime('%d/%m/%Y')
fim_semana = hoje - datetime.timedelta(days=3)
fds_formatado = fim_semana.strftime('%d/%m/%Y')
parametro_data = fds_formatado

# Inicializa o navegador Chrome
navegador = webdriver.Chrome()
navegador.maximize_window()

# Abre a página de login
navegador.get('https://siag.valecard.com.br/frota/pages/start.jsf')

# credenciais de usuário
usuario = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[2]/input')
senha = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[3]/input')
dominio = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[4]/select')
usuario.send_keys('935119')
senha.send_keys('10212')
dominio.send_keys('cliente')

# Submeter as Credenciais
botao_login = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[5]/input')
botao_login.click()
time.sleep(5)

formulario = navegador.find_element_by_id("MENU_FORM_HADOUKEN")

botao_manutencao = navegador.find_element_by_id('MENU_FORM_HADOUKEN:j_id32')
botao_manutencao.click()
time.sleep(3)

# Acesar a aba Relatórios
aba_relatorios = 'https://siag.valecard.com.br/frota/pages/reportsMaintenance.jsf'
navegador.get(aba_relatorios)
time.sleep(3)

# Escolher Relatorio
geral_manutencao = navegador.find_element_by_xpath('//*[@id="form:panel_body"]/table/tbody/tr/td[2]/select')
select = Select(geral_manutencao)
select.select_by_visible_text('Geral Manutenção')
time.sleep(2)

# Definir Parametros do Relatório
data_inicio = navegador.find_elements_by_xpath('//*[@id="form:c1InputDate"]')
for element in data_inicio:
    element.send_keys(fds_formatado)
time.sleep(1)

data_fim = navegador.find_elements_by_xpath('//*[@id="form:c2InputDate"]')
for element in data_fim:
    element.send_keys(ontem_formatado)
time.sleep(1)

# Gerar Relatório
gerar_relatorio = navegador.find_element_by_name('form:j_id855')
navegador.execute_script(gerar_relatorio.get_attribute('onclick'))
time.sleep(10)


exportar = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:panel_body"]/div[3]/div/input[1]')))
action = ActionChains(navegador)
action.click(exportar).perform()
time.sleep(10)


pasta_download = 'C:\\Users\\diego.oliveira\\Downloads'
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_manutencao = 'Z:\\07.PowerBI\\Dados\\Frota\\2.Manutenção\\1.Valecard\\1.Manutenções'

print(ultimo_download)

shutil.copy(ultimo_download, destino_manutencao)
time.sleep(2)

navegador.quit()
