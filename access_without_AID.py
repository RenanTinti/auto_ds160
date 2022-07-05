# AA00B26AAF

# TINTI
# 1990
# NICE

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

import time

import pandas as pd
from openpyxl import load_workbook

wb = load_workbook('Formulario-Visto-EUA_2022.xlsx', data_only=True)

sh = wb['Table 1']

aid = 'AA00B26AAF'
surname = 'TINTI'
year_of_birth = '1990'
mother_mother = 'NICE'

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service = servico)

driver.get('https://ceac.state.gov/GenNIV/default.aspx')

driver.maximize_window()

time.sleep(30)

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[4]/div/div[2]/div[5]/div[5]/div[2]/div/a').click()
		print('Clicou em Retrieve Application')
	except:
		print('Tentando clicar em Retrieve Application')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div/div[2]/input').send_keys(aid)
		print('Preencheu com o AID')
	except:
		print('Tentando preencher com o AID')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div/div[3]/table/tbody/tr/td[1]/span/input').click()
		print('Clicou em Retrieve Application')
	except:
		print('Tentando clicar em Retrieve Application')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[1]/table/tbody/tr[2]/td[1]/input').send_keys(surname)
		print('Inseriu 5 letras surname')
	except:
		print('Tentando inserir 5 letras surname')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[1]/table/tbody/tr[2]/td[2]/input').send_keys(year_of_birth)
		print('Inseriu Year of birth')
	except:
		print('Tentando inserir Year of birth')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[1]/table/tbody/tr[3]/td/div/input').send_keys(mother_mother)
		print('Inseriu mother name')
	except:
		print('Tentando inserir mother name')
		time.sleep(1)
		continue
	break

while True:
	try:
		driver.find_element(By.XPATH, '/html/body/form/div[3]/div[5]/div/div[2]/div[1]/div[2]/fieldset/div[2]/div[2]/table/tbody/tr/td[1]/span/input').click()
		print('Clicou em Retrieve Application')
	except:
		print('Tentando clicar em Retrieve Application')
		time.sleep(1)
		continue
	break