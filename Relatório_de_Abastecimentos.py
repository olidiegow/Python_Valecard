import time
import datetime
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import shutil
from selenium.webdriver import ActionChains

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
usuario = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[2]/input')
senha = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[3]/input')
dominio = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[4]/select')
usuario.send_keys('935119')
senha.send_keys('10212')
dominio.send_keys('cliente')

# Submeter as Credenciais
botao_login = navegador.find_element_by_xpath('//*[@id="wrap-geral"]/div[2]/div/div/ul/li[5]/input')
botao_login.click()
time.sleep(1)

# Acesar a aba Relatórios
aba_relatorios = 'https://siag.valecard.com.br/frota/pages/reportsCustomClient.jsf'
navegador.get(aba_relatorios)
time.sleep(5)

# Escolher Relatorio
abastecimento_detalhado = navegador.find_element_by_xpath('//*[@id="form:panel_body"]/table/tbody/tr/td[2]/select')
abastecimento_detalhado.send_keys('Abastecimentos DETALHADO', Keys.ENTER)
time.sleep(10)

# Definir Parametros do Relatório
data_inicio = navegador.find_elements_by_xpath('//*[@id="form:c1InputDate"]')
for element in data_inicio:
    element.send_keys(ontem_formatado)
time.sleep(5)

data_fim = navegador.find_elements_by_xpath('//*[@id="form:c2InputDate"]')
for element in data_fim:
    element.send_keys(ontem_formatado)
time.sleep(5)

# Gerar Relatório
botao_pesquisar = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.ID, 'form:j_id119')))
navegador.execute_script(botao_pesquisar.get_attribute('onclick'))
time.sleep(20)

#navegador.execute_script(botao_pesquisar.get_attribute('onclick'))
#time.sleep(20)

# Exportar Relatório
botao_excel = WebDriverWait(navegador, 25).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form:panel_body"]/div[3]/div[2]/input[2]')))
action = ActionChains(navegador)
action.click(botao_excel).perform()
time.sleep(10)

pasta_download = 'C:\\Users\\diego.oliveira\\Downloads'
arquivos = [(os.path.join(pasta_download, f), os.path.getmtime(os.path.join(pasta_download, f)))
            for f in os.listdir(pasta_download) if os.path.isfile(os.path.join(pasta_download, f))]
ultimo_download = max(arquivos, key=lambda x: x[1])[0]
destino_abastecimento = 'Z:\\07.PowerBI\\Dados\\Frota\\1.Abastecimento\\1.Transações'

print(ultimo_download)
shutil.copy(ultimo_download, destino_abastecimento)
navegador.quit()
