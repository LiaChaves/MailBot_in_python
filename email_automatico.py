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
options.add_argument("--user-data-dir=C:\\Users\\DanCh\\AppData\\Local\\Google\\Chrome\\User Data")
options.add_argument("--user-data-dir=C:\\Pessoa 1")  # Perfil "Pessoal"Pessoa 1 = Lian Pessoa 2 = Luiza Pessoa 3 = Thayna 
options.add_argument("--window-size=1200,768")  # Opcional: maximiza a janela
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"

# Inicialize o driver
driver = webdriver.Chrome(options=options)

driver.get("https://mail.google.com/mail/u/0/")
print("Chrome Browser Invoked")
time.sleep(4)

clientes = []
with open('clientes.csv', mode='r', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        clientes.append({
            'Nome': row['nome'],
            'E-mail': row['e-mail'],
            'Empresa': row['empresa']
        })

print(f"Total de clientes encontrados: {len(clientes)}")
#rota para pegar os pdfs
apresentacao = r"D:\\Lian Chaves\\Downloads\\reforma\\Apresentação Institucional EverySys 2025.pdf"
escopo = r"D:\\Lian Chaves\\Downloads\\reforma\\ESCOPO REFORMA TRIBUTÁRIA.pdf"
colaboradores = r"D:\\Lian Chaves\\Downloads\\reforma\\Time Reforma Tributária.pdf"
# Loop para enviar os e-mails
index = 0
while index < len(clientes):
    cliente = clientes[index]
    print(f"\nProcessando cliente {index+1}/{len(clientes)}: {cliente['Nome']}")
    
    try:
        # Clica no botão "Escrever"
        try:
            WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@class='T-I T-I-KE L3']"))
            ).click()
            print("Abrindo caixa de e-mail...")
        except Exception as e:
            driver.refresh()
        time.sleep(1)
        
        # Preenche o destinatário
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@class='agP aFw']"))
        ).send_keys(cliente['E-mail'])
        
        # Preenche o assunto
        driver.find_element(By.XPATH, "//input[@class='aoT']").send_keys(
            f"Webinar | Reforma Tributária – Convite Especial {cliente['Nome']}"
        )
        #envia os pdf's
        time.sleep(1)
        # Preenche o corpo do e-mail
        corpo = driver.find_element(By.XPATH, "//br[@clear='all']")
        corpo.send_keys(f"A Every System convida a {cliente['Empresa']} para um webinar exclusivo sobre Reforma Tributária, a realizar-se no dia 01/07/2025 às 14h-BRT.\n\nO evento abordará todas as atualizações recentes da reforma, além de apresentar o escopo dos serviços de implementação e acompanhamento contínuo, conduzidos por uma equipa especializada.\n\nDurante o encontro, será possível entender como ocorrerá o processo de mudança fiscal dentro do ERP, além de se familiarizar com rotinas de cadastro e ajustes dos impostos. O objetivo é traduzir e desmistificar as alterações no sistema, garantindo uma operação segura e eficiente antes que as novas regras impactem seu negócio.\n\nEm anexo, segue o material com o escopo completo dos serviços e o currículo dos profissionais envolvidos. As vagas são limitadas! Caso haja interesse, garanta sua participação imediatamente preenchendo o forms disponibilizado no link abaixo. Se não puder participar na data marcada, é possível agendar uma apresentação exclusiva (sujeito à disponibilidade da equipe).\nhttps://docs.google.com/forms/d/e/1FAIpQLSfr54UeI-OQrwKZuVhcWAbzBMGJXowtYRXvN-K798DeOadCyA/viewform?usp=sf_link\n\nAguardamos sua confirmação.As vagas se esgotam rápido!")
        time.sleep(1)
        
        # Localiza o input de arquivo oculto
        input_arquivo = driver.find_element(By.XPATH, "//input[@type='file']")
        input_arquivo.send_keys(apresentacao)
        
        input_arquivo = driver.find_element(By.XPATH, "//input[@type='file']")
        input_arquivo.send_keys(colaboradores)
        
        input_arquivo = driver.find_element(By.XPATH, "//input[@type='file']")
        input_arquivo.send_keys(escopo)
        print("Corpo do e-mail preenchido")

        time.sleep(6)
        
        # Envia o e-mail
        driver.find_element(By.XPATH, "//div[@class='T-I J-J5-Ji aoO v7 T-I-atl L3']").click()
        print("E-mail enviado!")
        time.sleep(1)  # Aguarda o envio
        
        index += 1  # Avança para o próximo cliente

    except Exception as e:
        print(f"Erro no cliente {index+1}: {str(e)}")
        #driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
        #driver.find_element_by_tag_name('body').send_keys(Keys.ESCAPE)
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        print(f"Cliente {index} : {cliente['Nome']} erro ao processar.")
        print("Reiniciando o processo...")
        index += 1
        driver.refresh()
        time.sleep(5)

print("\nTodos os e-mails foram enviados com sucesso!")
driver.quit()