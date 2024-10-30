from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import openpyxl
from tkinter import *
import chromedriver_autoinstaller
import pandas as pd

# Definir o caminho para o chromedriver
chromedriver_autoinstaller.install()
    
driver = webdriver.Chrome()

#buscar elemento data/hora
dt = driver.find_element(By.XPATH, '//*[@id="wob_dts"]').get_attribute('value')
#buscarelemento temperatura
temp = driver.find_element(By.XPATH, '//*[@id="wob_tm"]').get_attribute('value')
#buscar elemento umidade
umi = driver.find_element(By.XPATH, '//*[@id="wob_wc"]/div[1]/div[2]/div[2]').get_attribute('value')

arquivo = 'Dados_clima.xlsx'
planilha_nome = 'Lista'

class Aplicacao:
    
    def __init__(self):
        
        #criacão da janela na interface
        self.layout = Tk()
        self.layout.title("Captador de temperatura")
        self.layout.geometry("350x60")
        self.tela = Frame(self.layout)
        self.descricao = Label(self.tela, text="Para gerar arquivos em formato .csv")
        
        #vincular o botão à funcão que cria arquivo e insere dados
        self.exportar = Button(self.tela, text="Clique aqui", command=self.executar)
        
        #alocacão dos elementos na tela
        self.tela.pack()
        self.descricao.pack()
        self.exportar.pack()
        
        
        mainloop()
        
    # para realizar a captacão das informacões no site com o selenium
    def importar(self):

        #navegar até o site
        driver.get('https://www.google.com.br/')
        
        #achar elemento de input e escrever clima
        driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys('clima')
        
        #clicar no botão de pesquisar
        driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]').send_keys(Keys.ENTER)
        
        print(f"Hoje {dt} a temperatura está {temp} e a umidade do ar é de {umi}%!")

        #fechar o navegador
        driver.quit()
            
    def executar(self):
        # executar codigo para o botão funcionar
        tbl = load_workbook(arquivo)
        
        pln = tbl[planilha_nome]
        
        pln.cell(row=2, column=1).value = dt
        
        pln.cell(row=2, column=2).value = temp
        
        pln.cell(row=2, column=3).value = umi
        
#return dados
arquivo.save()    
    
print('Planilha atualizada com sucesso!')    
        

        
tl = Aplicacao()


        