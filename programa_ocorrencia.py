from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from dateutil.relativedelta import relativedelta
from selenium.webdriver.support.ui import Select
while True:
    login_1=input("Digite o seu Login: ")
    senha_1=input("Digite a sua Senha: ")
    certeza=input("Tem certeza que o login e senha estão corretos ? (s - sim/ n- não): ")
    if certeza.lower()=='s' or certeza.lower()=='sim':
        print()
        break
    else:
        print("\nDigite Novamente! \n")
navegador = webdriver.Chrome()
navegador.get("https://certpessoas.fgv.br/ct/")
nome = navegador.find_element(By.NAME, "UserName")  
senha = navegador.find_element(By.NAME, "Password")
navegador.implicitly_wait(2)
nome.send_keys(login_1)#teu login 
senha.send_keys(senha_1)#tua senha
navegador.implicitly_wait(2)
senha.send_keys(Keys.RETURN)
navegador.implicitly_wait(2)
ok=input("Colocou o aceitar nos cooks ?")
