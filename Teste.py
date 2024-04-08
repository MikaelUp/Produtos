# Importação das bibliotecas necessárias
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
import os

# Inicialização do WebDriver para acessar a página 1 do site Petlove
driver = webdriver.Chrome()
driver.get('https://www.petlove.com.br/cachorro/racoes?page=1')

# Coleta dos títulos dos produtos da página 1
titulos = driver.find_elements(By.XPATH,"//h2[@class='product-name card-list-name']")

# Coleta dos preços dos produtos da página 1
precos = driver.find_elements(By.XPATH,"//span[@class='catalog-card-prices__price-subscriber']")

# Definição da empresa
empresas = 'PETLOVE'

# Criação do arquivo Excel e da planilha "PETLOVE"
workbook = openpyxl.Workbook()
sheet_Sheet = workbook['Sheet']
sheet_Sheet.title = 'PETLOVE'
sheet_Sheet['A1'].value = 'Nome do Produto'
sheet_Sheet['B1'].value = 'Preço'
sheet_Sheet['C1'].value = 'Empresa'

# Preenchimento da planilha "PETLOVE" com os dados coletados da página 1
for titulo, preco in zip(titulos, precos):
    sheet_Sheet.append([titulo.text,preco.text,empresas])

# Abertura de uma nova instância do WebDriver para acessar a página 2 do site Petlove
driver = webdriver.Chrome()
driver.get('https://www.petlove.com.br/cachorro/racoes?page=2')

# Coleta dos títulos e preços dos produtos da página 2
titulos = driver.find_elements(By.XPATH,"//h2[@class='product-name card-list-name']")
precos = driver.find_elements(By.XPATH,"//span[@class='catalog-card-prices__price-subscriber']")

# Preenchimento da planilha "PETLOVE" com os dados coletados da página 2
for titulo, preco in zip(titulos, precos):
    sheet_Sheet.append([titulo.text,preco.text,empresas])

# Abertura de uma nova instância do WebDriver para acessar a página 1 do site Cobasi
driver = webdriver.Chrome()
driver.get('https://www.cobasi.com.br/pesquisa?terms=racoes&page=1')

# Coleta dos títulos e preços dos produtos da página 1 do site Cobasi
titulos = driver.find_elements(By.XPATH,"//h3[@class='styles__Title-sc-1ac06td-4 dPsqyZ']")
precos = driver.find_elements(By.XPATH,"//span[@class='card-price']")

# Criação da planilha "COBASI" e preenchimento com os dados coletados da página 1
sheet_Sheet2 = workbook.create_sheet('COBASI')
sheet_Sheet2['A1'].value = 'Nome do Produto'
sheet_Sheet2['B1'].value = 'Preço'
sheet_Sheet2['C1'].value = 'Empresa'
empresas2 = 'COBASI'

for titulo, preco in zip(titulos[1::2], precos[0::2]):
    sheet_Sheet2.append([titulo.text,preco.text,empresas2])

# Abertura de uma nova instância do WebDriver para acessar a página 2 do site Cobasi
driver = webdriver.Chrome()
driver.get('https://www.cobasi.com.br/pesquisa?terms=racoes&page=2')

# Coleta dos títulos e preços dos produtos da página 2 do site Cobasi
titulos = driver.find_elements(By.XPATH,"//h3[@class='styles__Title-sc-1ac06td-4 dPsqyZ']")
precos = driver.find_elements(By.XPATH,"//span[@class='card-price']")

# Preenchimento da planilha "COBASI" com os dados coletados da página 2
sheet_Sheet2 = workbook['COBASI']
for titulo, preco in zip(titulos[1::2], precos[0::2]):
    sheet_Sheet2.append([titulo.text,preco.text,empresas2])

# Salvamento do arquivo Excel
workbook.save(r'C:\Users\RH1\Desktop\VS\Estudos\Teste\RAÇÕES.xlsx')


# Abre o arquivo Excel com o aplicativo associado
os.system(f'start excel "{r'C:\Users\RH1\Desktop\VS\Estudos\Teste\RAÇÕES.xlsx'}"')
print('TESTANDO')