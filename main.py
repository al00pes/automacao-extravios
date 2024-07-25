from datetime import datetime
from time import strftime
from selenium import webdriver
import time
import pandas as pd
from anticaptchaofficial.imagecaptcha import * #importando a biblioteca do anticaptcha
from selenium.webdriver.common.by import By
#import pyperclip3
from selenium.webdriver import ActionChains
from selenium.webdriver import Keys
from bs4 import BeautifulSoup
import sys
import credenciais


#iniciando a classe sistema

class sistema():

    def __init__(self):
        self.site = '##########'
        self.driver = webdriver.Chrome (executable_path='U:\\RELATORIOS\\Extravios_automacao\\chrome_version_112\\chromedriver.exe') #caminho onde se encontra o webdriver
        self.driver.maximize_window() # Maximiza o navegador

    #abrindo o site do JMS
    def abrir_site(self):
        self.driver.get(self.site)
        time.sleep(1)

    #Mudando o idioma para português
    def idioma(self):
        #clica para expandir as opções
        self.driver.find_element("xpath",'//*[@id="gateway"]/div[1]/div[1]/div/div[1]/div[1]/input').click()
        time.sleep(2)
        #Clica na opçao "Portugal"
        self.driver.find_element("xpath", '/html/body/div[3]/div[1]/div[1]/ul/li[3]').click()
        time.sleep(2)

    #informando login e senha
    def login_senha(self):
        login = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[1]/div/div/input')
        login.click()
        time.sleep(2)
        login.send_keys(credenciais.login_jms)
        time.sleep(1)
        senha = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[2]/div/div[1]/input')
        senha.click()
        time.sleep(1)
        senha.send_keys(credenciais.senha_jms)
        time.sleep(5)
        conectar = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/button')
        conectar.click()
        time.sleep(20)

      # Essa API está comentando porque não está sendo mais utilizada para fazer a quebra do captcha
     # A quebra do captcha será feita de forma manual e o programa irá esperar 20 segundos para que possa ser feito
    ''' #Resolvendo o captcha com o anticaptcha
    def captcha(self):
        with open('captcha.jpeg','wb') as file:
            captcha_image = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[3]/div/div[2]/img')
            #captcha_image = ("C:\\Users\\arthur.lopes\\Documents\\Python\\captcha.jpeg")
            file.write(captcha_image.screenshot_as_png)
            # inicio da API AntiCaptcha
            solver = imagecaptcha()
            solver.set_verbose(1)
            solver.set_key(credenciais.chave_api)
            time.sleep(3)
            captcha_text = solver.solve_and_return_solution("captcha.jpeg")

            #Condição para a solução do captcha

            if captcha_text !=0: # se for diferente de zero.
                print(captcha_text)
            else:
                print("Task finishes with error"+ solver.solver_error) #Se for igual, irá aparecer um erro

            # Fim da API
            
            campo_captcha = self.driver.find_element("xpath",'//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[3]/div/div[1]/input')
            campo_captcha.click() #seleciona o campo
            time.sleep(2)
            campo_captcha.send_keys(captcha_text)
            time.sleep(5)
        #clicar em "conecte-se"
            conectar = self.driver.find_element("xpath",'//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/button')
            conectar.click()
            time.sleep(15)'''

    # Acessando o local da pesquisa
    def acessando_arbitragem(self):
        #clicando em " qualidade de serviço)
        qualidade = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[2]/div[2]/div[1]/div[2]/div[9]/span')
        qualidade.click()
        time.sleep(3)
        # clicando em "Arbitragem"
        arbitragem = self.driver.find_element("xpath",'//*[@id="gateway"]/div[1]/div[1]/div[1]/div/div[1]/div[4]/ul[1]/li[3]')
        arbitragem.click()
        time.sleep(3)
        #clicando em " Declaração de arbitragem"
        declaracao = self.driver.find_element("xpath",'//*[@id="gateway"]/div[1]/div[1]/div[1]/div/div[1]/div[4]/ul[2]/li[1]/div/div')
        declaracao.click()
        time.sleep(7)

    #Digitando o Id e obter o valor
    def pedido_valor(self):

        #carregando a planilha onde está as informações dos pacotes
        extravios = "U:\\RELATORIOS\\Extravios_automacao\\extravios.xlsx"
        df = pd.read_excel(extravios)
        idpedido = self.driver.find_element("xpath",'//*[@id="__qiankun_subapp_wrapper_for_vue_service_quality_index__"]/div/div/div/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/span[2]/div[1]/input')
        idpedido.click()
        time.sleep(2)
        lista_pedido = [] #Lista de IDPEDIDO
        lista_valor = [] #Lista de valor_nota

        # loop para digitar o idpedido e obter o valor da nota_valor
        for index,row in df.iterrows():
            # pegar o ID da planilha
            idpedido.send_keys(str(row['ID']))
            time.sleep(2)
            numero = self.driver.find_element("xpath",'//*[@id="__qiankun_subapp_wrapper_for_vue_service_quality_index__"]/div/div/div/div[1]/div/div/div[1]/div[1]/div[2]/div[4]/span[2]/div[1]/input')
            valornota = numero.get_property('value') #obter o valor da caixa de texto
            numeropedido = idpedido.get_property('value')
            print(numeropedido)
            print(valornota)
            lista_pedido.append(numeropedido) # adiciona o valor na lista
            lista_valor.append(valornota) # adiciona o valor na lista

            time.sleep(1)
            idpedido.clear()

        #Cria um dataframe
        df = pd.DataFrame(zip(lista_pedido,lista_valor),columns=['idpedido','valor_nota']) # Transforma a lista em dataframe
        print(df) # exibe o dataframe
        #Criando variavel para salvar com a data e horario.
        #data_hoje = time.strftime("%Y-%m-%d_%H:%M:%S")
        #agora_string = data = datetime.datetime.now().strftime("%Y-%m-%d_%H:%M:%S")

        #Irá salvar o arquivo em excel com a data e hora atual
        df.to_excel("U:\\RELATORIOS\\Extravios_automacao\\relatorio.xlsx", index=False) # Salva o dataframe em excel.
        time.sleep(5)
        self.driver.close()
        quit()






sistema = sistema()
sistema.abrir_site()
sistema.idioma()
sistema.login_senha()
#sistema.captcha()
sistema.acessando_arbitragem()
sistema.pedido_valor()
