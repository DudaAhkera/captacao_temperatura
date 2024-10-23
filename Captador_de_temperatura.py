from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import openpyxl
import datetime
from tkinter import *
import mysql.connector

class Aplicacao:
    def __init__(self):
        
        #criacão da janela na interface
        self.layout = Tk()
        self.layout.title("Captador de temperatura")
        self.layout.geometry("350x60")
        self.tela = Frame(self.layout)
        self.descricao = Label(self.tela, text="Para gerar arquivos em formato .csv")
        self.exportar = Button(self.tela, text="Clique aqui")
        
        #alocacão dos elementos na tela
        self.tela.pack()
        self.descricao.pack()
        self.exportar.pack()
        
        
        mainloop()
        
    # para realizar a captacão das informacões no site com o selenium
    def importar(self):
        driver = webdriver.Chrome(executable_path='/caminho/para/chromedriver')
        
        #navegar até o site
        driver.get('https://www.climatempo.com.br/previsao-do-tempo/15-dias/cidade/558/saopaulo-sp')
        
        #esperar o site carregar
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="//*[@id="first-block-of-days"]/div[4]/section[1]"]' ))
        )
        
        #extrair uma tabela
        linhas = tabela.find_elements(By.TAG_NAME, "tr")
        
        #extrair os dados da tabela
        dados = []
        for linha in linhas:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            dados.append([coluna.text for coluna in colunas])
            
        #fechar o navegador
        driver.quit()
        
        return dados
        
    def criar_tabela():
        #criar um novo arquivo
        wb = openpyxl.Workbook()
        ws = wb.active
        
        #adicionar dados à planilha
        # ********* ADICIONAR CAMINHO DOS DADOS
        dados = [] 
        for linha in dados:
            ws.append(linha)
        
        #salvar o arquivo excel   
        wb.save('dados_extraidos.xlsx')
        
    def inserir_dados():
        #conectar ao banco de dados mysql
        conexao = mysql.connector.connect(
            host = '172.0.0.1',
            user = 'root',
            password = 'Zulenice20@',
            db = "dados_temp_sp"
        )
        
        cursor = conexao.cursor()
        
        #criar uma tabela no mysql
        cursor.execute("""
            CREATE TABLE IF NOT EXIST dados_temp_sp (
                data_hora DATETIME,
                temperatura FLOAT,
                umidade_do_ar FLOAT
            )
                       """)

        #inserir no banco os dados extraídos
        dados = []
        for linha in dados:
            cursor.execute("""
                INSERT INTO dados_temp_sp (coluna1, coluna2, coluna3)
                VALUES (%s, %s, %s)
                           """, linha)
        
        #confirmar a insercão
        conexao.commit()
        
        #fechar a conexão
        cursor.close()
        conexao.close()
        
tl = Aplicacao()


        