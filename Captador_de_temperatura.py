from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from openpyxl import load_workbook, Workbook
from tkinter import *
import chromedriver_autoinstaller

# Definir o caminho para o chromedriver
chromedriver_autoinstaller.install()

arquivo = 'Dados_clima.xlsx'
planilha_nome = 'Lista'

class Aplicacao:
    
    def __init__(self):
        
        #criacão da janela na interface
        self.layout = Tk()
        self.layout.title("Captador de temperatura")
        self.layout.geometry("350x60")
        self.tela = Frame(self.layout)
        self.descricao = Label(self.tela, text="Para gerar arquivos em formato .xlsx")
        
        #vincular o botão à funcão que cria arquivo e insere dados
        self.exportar = Button(self.tela, text="Clique aqui", command=self.capturar)
        
        #alocacão dos elementos na tela
        self.tela.pack()
        self.descricao.pack()
        self.exportar.pack()
        
        self.dt = None
        self.temp = None
        self.umi = None
        
        mainloop()
        
    def criar_arquivo(self):
        # Criar um novo arquivo Excel e uma planilha
        wb = Workbook()
        ws = wb.active
        ws.title = planilha_nome
        ws.append(["Data/Hora", "Temperatura", "Umidade"])  # Cabeçalhos
        wb.save(arquivo) 
        
           
    # para realizar a captacão das informacões no site com o selenium
    def importar(self):
        
        driver = webdriver.Chrome()

        #navegar até o site
        driver.get('https://www.google.com.br/')
        
        #aguardar página carregar
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="APjFqb"]')))
        
        #achar elemento de input e escrever clima
        driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys('clima')
        
        #clicar no botão de pesquisar
        driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]').send_keys(Keys.ENTER)
        
        
        #buscar elemento data/hora
        self.dt = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="wob_dts"]'))).text
        #buscarelemento temperatura
        self.temp = driver.find_element(By.XPATH, '//*[@id="wob_tm"]').text
        #buscar elemento umidade
        self.umi = driver.find_element(By.XPATH, '//*[@id="wob_hm"]').text
        
        print(f"Hoje {self.dt} a temperatura está {self.temp}ºC e a umidade do ar é de {self.umi}!")

        #fechar o navegador
        driver.quit()
        
            
    def executar(self):
        # executar codigo para o botão funcionar
        tbl = load_workbook(arquivo)
        lista_arq = tbl[planilha_nome]
        
        proxima_linha = lista_arq.max_row + 1
        lista_arq.cell(row=proxima_linha, column=1, value=self.dt)  # Data e Hora
        lista_arq.cell(row=proxima_linha, column=2, value=self.temp)  # Temperatura
        lista_arq.cell(row=proxima_linha, column=3, value=self.umi)  # Umidade
        
        #return dados
        tbl.save(arquivo)    
        
        print('Planilha atualizada com sucesso!') 
    
    def capturar(self):
        self.importar()
        self.executar()
        

        
tl = Aplicacao()


        