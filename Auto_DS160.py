# AA00B26AAF

# TINTI
# 1990
# NICE

# Importação das bibliotecas de automação
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import pyautogui as pg

import time

# Importação das bibliotecas de Excel
import pandas as pd
from openpyxl import load_workbook

# Definindo o service para gerenciar o Driver
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service = servico)

# Abrindo a página inicial
driver.get('https://ceac.state.gov/GenNIV/default.aspx')

# Maximizando a janela
driver.maximize_window()

# Aguardando 30 segundos para o preenchimento do código e, possivelmente, do país
time.sleep(15
)

# Carregando a planilha Excel do formulário preenchido 
wb = load_workbook('Formulario-Visto-EUA_2022.xlsx', data_only=True)
sh = wb['Table 1']

# Definição das variáveis que serão usadas nos testes
aid = 'AA00B26AAF'
surname = 'TINTI'
year_of_birth = '1990'
mother_mother = 'NICE'

# O código comentado abaixo será utilizado apenas para novas aplicações
'''
# Clicar em I Agree, logo após preencher o país e código na primeira página
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[4]/div/div[3]/span[6]/span/span/input').click()
		print('Clicou em I Agree')
	except:
		print('Tentando clicar em I Agree')
		time.sleep(1)
		continue
	break

# Tentando escrever a palavra de segurança NICE
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[4]/div/div[6]/div[3]/input').send_keys("NICE")
		print('Esreveu NICE')
	except:
		print('Tentando escrever NICE')
		time.sleep(1)
		continue
	break

# Tentando clicar em Continue para direcionamento para a página do formulário
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[4]/div/div[7]/fieldset/div/span[1]/input').click()
		print('Clicou em Continue')
	except:
		print('Tentando clicar em Continue')
		time.sleep(1)
		continue
	break'''

# Clica em Retrieve Application (o código deve ter sido preenchido)
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[4]/div/div[2]/div[5]/div[5]/div[2]/div/a').click()
		print('Clicou em Retrieve Application')
	except:
		print('Tentando clicar em Retrieve Application')
		time.sleep(1)
		continue
	break

# Preenche o campo solicitado com o AID
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div/div[2]/input').send_keys(aid)
		print('Preencheu com o AID')
	except:
		print('Tentando preencher com o AID')
		time.sleep(1)
		continue
	break

# Clica pela segunda vez em Retrieve Application
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div/div[3]/table/tbody/tr/td[1]/span/input').click()
		print('Clicou em Retrieve Application')
	except:
		print('Tentando clicar em Retrieve Application')
		time.sleep(1)
		continue
	break

# Insere as 5 letras do sobrenome (segurança)
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[1]/table/tbody/tr[2]/td[1]/input').send_keys(surname)
		print('Inseriu 5 letras surname')
	except:
		print('Tentando inserir 5 letras surname')
		time.sleep(1)
		continue
	break

# Insere o ano de nascimento (segurança)
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[1]/table/tbody/tr[2]/td[2]/input').send_keys(year_of_birth)
		print('Inseriu Year of birth')
	except:
		print('Tentando inserir Year of birth')
		time.sleep(1)
		continue
	break

# insere a palavra de segurança NICE (segurança)
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[1]/table/tbody/tr[3]/td/div/input').send_keys(mother_mother)
		print('Inseriu mother name')
	except:
		print('Tentando inserir mother name')
		time.sleep(1)
		continue
	break

# Clica pela terceira vez em Retrieve Application, dessa vez para de fato acessar o formulário de onde ele havia sido salvo pela ultima vez
while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[2]/table/tbody/tr/td[1]/span/input').click()
		print('Clicou em Retrieve Application')
	except:
		print('Tentando clicar em Retrieve Application')
		time.sleep(1)
		continue
	break

# ========================================= PERSONAL INFO 1 =============================================================================================================

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[1]/div[1]/div[1]/input').send_keys(sh["B4"].value)
		print('Preencheu o sobrenome')
	except:
		print('Tentando preencher o sobrenome')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[1]/div[1]/div[2]/input').send_keys(sh["Q4"].value)
		print('Preencheu o nome')
	except:
		print('Tentando preencher o nome')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[1]/div[1]/div[3]/input').send_keys(sh["B6"].value)
		print('Preencheu o nome completo')
	except:
		print('Tentando preencher o nome completo')
		time.sleep(1)
		continue
	break

def set_usou_outro_nome():
    mark = 'X'
    if str(mark) in str(sh["B182"].value):
        usou_outro_nome = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/div[2]/fieldset[1]/div[1]/div/div[2]/div/div/span/span/table/tbody/tr/td[1]/span/input'
    else:
        usou_outro_nome = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/div[2]/fieldset[1]/div[1]/div/div[2]/div/div/span/span/table/tbody/tr/td[2]/input'
    return usou_outro_nome

while True:
	try:
		driver.find_element(By.XPATH, set_usou_outro_nome()).click()
		print('Radio button - usou outro nome')
	except:
		print('Tentando Radio button - usou outro nome')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/div[3]/fieldset[1]/div[1]/div/div[2]/div/div/span/span/table/tbody/tr/td[2]/input').click()
		print('Radio button - telecode nome')
	except:
		print('Tentando Radio button - telecode nome')
		time.sleep(1)
		continue
	break

def set_sexo():
    mark = 'X'
    if str(mark) in str(sh["B11"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[1]/div/span/span/table/tbody/tr/td[1]/input'
    elif str(mark) in str(sh["M11"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[1]/div/span/span/table/tbody/tr/td[2]/input'
    return sexo

while True:
	try:
		driver.find_element(By.XPATH, set_sexo()).click()
		print('Radio button - sexo')
	except:
		print('Tentando Radio button - sexo')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select').click()
		print('Clicou em Marital Status')
	except:
		print('Tentando clicar em Marital Status')
		time.sleep(1)
		continue
	break

def set_marital_status():
    mark = 'X'
    if str(mark) in str(sh["B14"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select/option[5]'
    elif str(mark) in str(sh["E14"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select/option[2]'
    if str(mark) in str(sh["I14"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select/option[3]'
    elif str(mark) in str(sh["M14"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select/option[7]'
    if str(mark) in str(sh["X14"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select/option[8]'
    elif str(mark) in str(sh["AE14"].value):
        sexo = '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[2]/div/div[2]/div/select/option[6]'
    return sexo

while True:
	try:
		driver.find_element(By.XPATH, set_marital_status()).click()
		print('Combobox - marital status')
	except:
		print('Combobox - marital status')
		time.sleep(1)
		continue
	break

birth_date = str(sh["M16"].value)
print(birth_date)

birth_list = birth_date.split('/')
print(birth_list)

birth_day = birth_list[0]
birth_month = birth_list[1]
birth_year = birth_list[2]

print(birth_day)
print(birth_month)
print(birth_year)

def convert_month(birth_month):
    match birth_month:
        case '1':
            return 'JAN'
        case '2':
            return 'FEB'
        case '3':
            return 'MAR'
        case '4':
            return 'APR'
        case '5':
            return 'MAY'
        case '6':
            return 'JUN'
        case '7':
            return 'JUL'
        case '8':
            return 'AUG'
        case '9':
            return 'SEP'
        case '10':
            return 'OCT'
        case '11':
            return 'NOV'
        case '12':
            return 'DEC'

print(convert_month(birth_month))

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[1]/select[1]').click()
		print('Clicou em dia de nascimento')
	except:
		print('Tentando clicar em dia de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[1]/select[1]').send_keys(birth_day)
		print('Inseriu dia de nascimento')
	except:
		print('Tentando inserir dia de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[1]/select[2]').click()
		print('Clicou em mês de nascimento')
	except:
		print('Tentando clicar em mês de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[1]/select[2]').send_keys(convert_month(birth_month))
		print('Inseriu mês de nascimento')
	except:
		print('Tentando inserir mês de nascimento')
		time.sleep(1)
		continue
	break

pg.press('enter')

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[1]/input').click()
		print('Clicou em ano de nascimento')
	except:
		print('Tentando clicar em ano de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[1]/input').send_keys(birth_year)
		print('Inseriu ano de nascimento')
	except:
		print('Tentando inserir ano de nascimento')
		time.sleep(1)
		continue
	break

local = str(sh["B16"].value)

lista_local = local.split('/')

cidade = lista_local[0]
estado = lista_local[1]
pais = lista_local[2]

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[2]/input').send_keys(cidade)
		print('Inseriu cidade de nascimento')
	except:
		print('Tentando inserir cidade de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/div[3]/input').send_keys(estado)
		print('Inseriu estado de nascimento')
	except:
		print('Tentando inserir estado de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/table/tbody/tr/td/div[3]/fieldset[3]/div[1]/div/div/select').send_keys('Brazil')
		print('Inseriu país de nascimento')
	except:
		print('Tentando inserir país de nascimento')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/fieldset/div/span[2]/input').click()
		print('Salvou')
	except:
		print('Tentando salvar')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div[2]/div/fieldset/div/fieldset/table/tbody/tr/td[1]/div/span/input').click()
		print('Clicou em Continue Application')
	except:
		print('Tentando clicar em Continue Application')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/fieldset/div/span[3]/input').click()
		print('Clicou em Next: Personal 2')
	except:
		print('Tentando clicar em Next: Personal 2')
		time.sleep(1)
		continue
	break

# ==================================== SAVE SCREEN ====================================================================================================================

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div[2]/div/fieldset/div/fieldset/table/tbody/tr/td[1]/div/span/input').click()
		print('Clicou em Continue Application')
	except:
		print('Tentando clicar em Continue Application')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div[2]/div/fieldset/div/fieldset/table/tbody/tr/td[1]/div/span/input').click()
		print('Clicou em Continue Application')
	except:
		print('Tentando clicar em Continue Application')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div[2]/div/fieldset/div/fieldset/table/tbody/tr/td[1]/div/span/input').click()
		print('Clicou em Continue Application')
	except:
		print('Tentando clicar em Continue Application')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div[2]/div/fieldset/div/fieldset/table/tbody/tr/td[1]/div/span/input').click()
		print('Clicou em Continue Application')
	except:
		print('Tentando clicar em Continue Application')
		time.sleep(1)
		continue
	break

# =================================== PERSONAL INFO 2 =====================================================================================================================

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break

'''print(sh["Q3"].value)
print(type(sh["Q3"].value))

def set_pagamento_impostos():
    mark = 'X'
    if str(mark) in str(sh["B182"].value):
        pagamento_impostos = True
    else:
        pagamento_impostos = False
    return pagamento_impostos

print(set_pagamento_impostos())'''

# servico = Service(ChromeDriverManager().install())
# driver = webdriver.Chrome(service = servico)

'''driver.get('https://ceac.state.gov/GenNIV/default.aspx')

driver.maximize_window()

while True:
	try:
		driver.find_element(By.XPATH, '#').click()
		print('')
	except:
		print('')
		time.sleep(1)
		continue
	break'''