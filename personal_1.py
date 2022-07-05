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
time.sleep(15)

# Carregando a planilha Excel do formulário preenchido 
wb = load_workbook('Formulario-Visto-EUA_2022.xlsx', data_only=True)
sh = wb['Table 1']

# Definição das variáveis que serão usadas nos testes
aid = 'AA00B26AAF'
surname = 'A'
year_of_birth = '1990'
mother_mother = 'NICE'

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