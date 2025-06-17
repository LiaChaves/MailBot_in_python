from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time
import csv
import math



options = Options()
options.add_argument("--user-data-dir=C:\\Users\\seu_usuario\\AppData\\Local\\Google\\Chrome\\User Data")#Diretorio de aonde tem os usuarios que deseja usar ou criar
options.add_argument("--user-data-dir=C:\\Pessoa 1")  # Ao nominar um novo nome, um novo usuario é criado no chrome, indico não salvar o gmail como usuario principal, pois o nome irá se torna outro
options.add_argument("--window-size=1200,768")  # Opcional: maximiza a janela 
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" #aonde está o executavel do seu chrome

# Inicialize o driver
driver = webdriver.Chrome(options=options)

driver.get("https://mail.google.com/mail/u/0/")
print("Chrome Browser Invoked")
time.sleep(4)

clientes = []
with open('clientes.csv', mode='r', encoding='utf-8') as csvfile: #chama o cvs e determina como cada linha e coluna sera usada
    reader = csv.DictReader(csvfile)
    for row in reader:
        clientes.append({
            'Nome': row['nome'],
            'E-mail': row['e-mail'],
            'Empresa': row['empresa']
        })

print(f"Total de clientes encontrados: {len(clientes)}") 
#rota para pegar os pdfs
apresentacao = r"D:\\Seu usuario\\Downloads\\Filosofia\\Posso comer manga e tomar leite?.pdf"
escopo = r"D:\\Seu usuario\\Downloads\\Conhecimentos\\Como dar cambalhota.pdf"
colaboradores = r"D:\\Seu usuario\\Downloads\\Receitas\\Receita de Goibada.pdf"
#aqui você pode usar pdfs da sua escolha para adicionar no e-mail antes do envio(mais a baixo tem o codigo do envio)
# Loop para enviar os e-mails
index = 0
while index < len(clientes):
    cliente = clientes[index]
    print(f"\nProcessando cliente {index+1}/{len(clientes)}: {cliente['Nome']}")
    #ao adicionar na listagem geralmente ponho o primeiro nome ou organizo o csv para retirar todas as str após o primeiro espaço
    #caso tenha mais de um nome use "SeuCSV['Nome'] = SeuCSV['Nome'].str.split().str[0].str.strip().replace(".", "").replace("(", "").replace(")", "").replace(",", "")#mantem só a primeira palavra de um str"
    try:
        # Clica no botão "Escrever"
        try:
            WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='T-I T-I-KE L3']"))#caso não funcione atualize o xpath
            ).click()
            print("Abrindo caixa de e-mail...")
        except Exception as e:
            driver.refresh()
        time.sleep(1)
        
        # Preenche o destinatário
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@class='agP aFw']"))#caso não funcione atualize o xpath
        ).send_keys(cliente['E-mail'])
        
        # Preenche o assunto
        driver.find_element(By.XPATH, "//input[@class='aoT']").send_keys(
            f"Oie {cliente['Nome'] vamos criar uma festa?}"#caso não funcione atualize o xpath
        )
        #envia os pdf's
        time.sleep(1)
        # Preenche o corpo do e-mail
        corpo = driver.find_element(By.XPATH, "//br[@clear='all']")
        corpo.send_keys(f"Eu Lilian convido você e a {cliente['Empresa']} para uma festa muito legal, segue uns pdf's com os assuntos da festa")#use \n para dar espaço de uma linha
        time.sleep(1)
        
        # Localiza o input de arquivo oculto
        input_arquivo = driver.find_element(By.XPATH, "//input[@type='file']") #deixa assim para reconhecer qualquer arquivo, mas pode ser alterado para um arquivo especifico 
        input_arquivo.send_keys(apresentacao)
        
        input_arquivo = driver.find_element(By.XPATH, "//input[@type='file']")
        input_arquivo.send_keys(colaboradores)
        
        input_arquivo = driver.find_element(By.XPATH, "//input[@type='file']")
        input_arquivo.send_keys(escopo)
        print("Corpo do e-mail preenchido")

        time.sleep(6)
        
        # Envia o e-mail
        driver.find_element(By.XPATH, "//div[@class='T-I J-J5-Ji aoO v7 T-I-atl L3']").click() #caso não funcione atualize o xpath
        print("E-mail enviado!")
        time.sleep(1)  # Aguarda o envio
        
        index += 1  # Avança para o próximo cliente

    except Exception as e:
        print(f"Erro no cliente {index+1}: {str(e)}")
        #driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
        #driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform() #caso ocorra um bug e ele abra 2 abas de envio de e-mail, ele fecha as duas e reinicia o processo
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        print(f"Cliente {index} : {cliente['Nome']} erro ao processar.")
        print("Reiniciando o processo...")
        index += 1 #avança para o proximo para evitar erro no looping
        driver.refresh()
        time.sleep(5)

print("\nTodos os e-mails foram enviados com sucesso!")
driver.quit()
