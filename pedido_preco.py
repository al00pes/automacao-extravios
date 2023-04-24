from selenium import webdriver
import time
import pandas as pd
from anticaptchaofficial.imagecaptcha import * #importando a biblioteca do anticaptcha
import pyperclip
import clipboard
import tkinter as tk



#iniciando a classe JMS
class jms():
    def __init__(self):
        self.site = 'https://jmsbr.jtjms-br.com/'
        self.driver = webdriver.Chrome (executable_path='C:\\chromedriver\\chromedriver.exe')
        self.driver.maximize_window()

    #abrindo o site do JMS
    def abrir_site(self):
        self.driver.get(self.site)
        time.sleep(1)

    #Mudando o idioma para português
    def idioma(self):
        self.driver.find_element("xpath",'//*[@id="gateway"]/div[1]/div[1]/div/div[1]/div[1]/input').click()
        time.sleep(1)
        self.driver.find_element("xpath", '/html/body/div[2]/div[1]/div[1]/ul/li[3]').click()
        time.sleep(2)

    #informando login e senha
    def login_senha(self):
        login = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[1]/div/div/input')
        login.click()
        time.sleep(2)
        login.send_keys('00838020')
        time.sleep(1)
        senha = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[2]/div/div[1]/input')
        senha.click()
        time.sleep(1)
        senha.send_keys('Tecnologia@01')
        time.sleep(2)

    #Resolvendo o captcha com o anticaptcha
    def captcha(self):
        with open('captcha.jpeg','wb') as file:
            captcha_image = self.driver.find_element("xpath", '//*[@id="gateway"]/div[1]/div[1]/div/div[3]/div/form/div[3]/div/div[2]/img')
            #captcha_image = ("C:\\Users\\arthur.lopes\\Documents\\Python\\captcha.jpeg")
            file.write(captcha_image.screenshot_as_png)
            # inicio da API AntiCaptcha
            solver = imagecaptcha()
            solver.set_verbose(1)
            solver.set_key("1b3315c3238d6eca0d77771872d9b402")
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
            time.sleep(15)

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
        time.sleep(5)

    #Digitando o Id e obter o valor
    def pedido_valor(self):

        #carregando a planilha
        extravios = "extravios_01.xlsx"
        df = pd.read_excel(extravios)
        print(df)
        idpedido = self.driver.find_element("xpath", '//*[@id="__qiankun_subapp_wrapper_for_vue_service_quality_index__"]/div/div/div/div[1]/div/div/div[1]/div[1]/div[2]/div[2]/span[2]/div[1]/input')
        value = self.driver.find_element("xpath",'//*[@id="__qiankun_subapp_wrapper_for_vue_service_quality_index__"]/div/div/div/div[1]/div/div/div[1]/div[1]/div[2]/div[4]/span[2]/div[1]/input').text

        for index,row in df.iterrows():
            idpedido.send_keys(str(row['ID'])) # INSERINDO ID PEDIDO
            time.sleep(3)
            print(row)
            #time.sleep(3)
            #root = tk.Tk()
            #text = root.clipboard_get(str(value))
            print(text)
            time.sleep(3)
            idpedido.clear()
            time.sleep(2)



sistema = jms()
sistema.abrir_site()
sistema.idioma()
sistema.login_senha()
sistema.captcha()
sistema.acessando_arbitragem()
sistema.pedido_valor()
