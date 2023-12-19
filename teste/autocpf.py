import os
import re
import time
import PyPDF2
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import shutil
from selenium.webdriver.support.ui import Select
from docx import Document
from docx.shared import Pt
import pyautogui
import pandas as pd

# Configurar o navegado
driver = webdriver.Chrome()

#Realizar login no site
driver.get('https://sistemas.sefaz.am.gov.br/gcc/menuservidor/idam/cadastroProdutorRuralAction.action')  # Substitua pelo URL do site
usuario = driver.find_element(By.ID, 'username')  # Substitua pelo campo de usuário
senha = driver.find_element(By.ID, 'password')  # Substitua pelo campo de senha
botao_login = driver.find_element(By.XPATH, '//*[@id="fm1"]/fieldset/div[3]/div/div[4]/input[4]')  # Substitua pelo botão de login

usuario.send_keys('03483401253')
senha.send_keys('carteira23')
botao_login.click()

planilha = pd.read_excel('cpfs.xlsx')

for index, row in planilha.iterrows():
    cpf = row['CPF']  # Pegar o CPF da linha atual
    while True:
        elemento = driver.find_element(By.XPATH,'//*[@id="oCMenu___GCC1006"]')

        # Faça algo com o elemento, por exemplo, clique nele
        elemento.click()

        # Localize a lista suspensa usando o XPath fornecido (a <div> que representa o menu)
        dropdown_div = driver.find_element(By.XPATH, '//*[@id="oCMenu___GCC1008"]')

        # Clique na <div> para expandir o menu
        dropdown_div.click()
        caixacpf = driver.find_element(By.XPATH,'//*[@id="formCceaPessoaFisica_pessoaFisicaTOA_pessoaFisica_cpfFormatado"]')
        caixacpf.click()
        caixacpf.send_keys(cpf)
        time.sleep(2)

        consultar = driver.find_element(By.XPATH,'//*[@id="formCceaPessoaFisica_pessoaFisica!pesquisarPessoaFisica"]')
        consultar.click()

        try:
            lapis = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH,'//*[@id="tbPessoaFisica"]/tbody/tr/td[6]/a')))
            lapis.click()
        except Exception as e:
            print(f"Erro ao clicar em 'abadeclaracao' para o CPF {cpf}")
                            # Atualize a coluna 'Processado' para marcar a linha como processada
            planilha.at[index, 'vish'] = 'deu merda'
                # Salve a planilha atualizada no mesmo arquivo Excel
            planilha.to_excel('cpfs.xlsx', index=False)
            break  # Saia do loop se o usuário não quiser continuar

        nome = driver.find_element(By.XPATH,'//*[@id="formPessoaFisica_pessoaFisicaTOA_pessoaFisica_pfNome"]')
        nome_da_pagina = driver.execute_script("return arguments[0].value;", nome)

        print(nome_da_pagina, cpf)

        doc = Document (r"I:\ARQUIVO DIGITAL CPR\fabio jr\autodec\teste\modelo.docx")
        for paragraph in doc.paragraphs:
            text = paragraph.text
            if '(nome)' in text:
                text = text.replace('(nome)', nome_da_pagina)
            if '(cpf)' in text:
                text = text.replace('(cpf)', cpf)
            paragraph.clear()
            run = paragraph.add_run(text)
            run.font.size = Pt(12)  # Defina o tamanho da fonte conforme necessário
        
            # Salve o documento com um nome específico
        nome_do_arquivo = nome_da_pagina  # Solicitar um nome para o arquivo
        nome_do_arquivo = nome_do_arquivo + ".docx"  # Adicione a extensão .docx
        doc.save(nome_do_arquivo)  # Salve o documento com o nome fornecido


        break

        
# Feche o navegador
driver.quit()
    