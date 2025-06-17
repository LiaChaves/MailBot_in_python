# 📧 MailBot - Documentação do Script de Automação de E-mails

## Visão Geral

Este script Python automatiza o envio de e-mails personalizados através do Gmail usando Selenium.  
Ele lê dados de clientes de um arquivo CSV, envia e-mails com anexos e personaliza as mensagens com informações específicas de cada cliente.

---

## Fluxograma do Processo

> ![deepseek_mermaid_20250617_f2d691](https://github.com/user-attachments/assets/574e2bf3-8671-4636-a4f8-35b7d8b79bfb)

---

## Pré-requisitos

- Python 3.7+
- Google Chrome instalado
- Bibliotecas Python:

```bash
pip install selenium webdriver-manager
```

---

## Configuração do Ambiente Virtual (.venv)
## Por que usar um ambiente virtual?
Ambientes virtuais isolam as dependências do seu projeto, evitando conflitos entre pacotes de diferentes projetos.

Passo a passo para criar e usar um ambiente virtual:

```bash
# Instalar o módulo virtualenv
pip install virtualenv

# Criar o ambiente virtual
python -m venv .venv

# Ativar o ambiente virtual

# Windows (PowerShell)
.\.venv\Scripts\Activate

# Linux/MacOS
source .venv/bin/activate

# Instalar dependências
pip install selenium webdriver-manager
```

---

## Estrutura do Projeto

```text
MailBot_in_python/
├── .venv/                  # Ambiente virtual
├── email_automatico.py     # Script principal
└── clientes.csv            # Arquivo de dados
```

---

## Estrutura do Código

### 1. Importações

```python
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
```

### 2. Configuração do Navegador

```python
options = Options()
options.add_argument("--user-data-dir=C:\\Pessoa 1")  # Perfil do Chrome
options.add_argument("--window-size=1200,768")         # Tamanho da janela
options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
driver = webdriver.Chrome(options=options)
```

### 3. Acesso ao Gmail

```python
driver.get("https://mail.google.com/mail/u/0/")
print("Chrome Browser Invoked")
time.sleep(4)  # Aguarda carregamento inicial
```

### 4. Leitura de Dados dos Clientes

```python
clientes = []
with open('clientes.csv', mode='r', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        clientes.append({
            'Nome': row['nome'],
            'E-mail': row['e-mail'],
            'Empresa': row['empresa']
        })
```

### 5. Definição de Anexos

```python
apresentacao = r"D:\\Caminho\\Arquivo1.pdf"
escopo = r"D:\\Caminho\\Arquivo2.pdf"
colaboradores = r"D:\\Caminho\\Arquivo3.pdf"
```

### 6. Composição do E-mail

```python
# Destinatário
elemento.send_keys(cliente['E-mail'])

# Assunto
elemento.send_keys(f"Oie {cliente['Nome']} vamos criar uma festa?")

# Corpo da mensagem
corpo.send_keys(f"Eu Lilian convido você e a {cliente['Empresa']}...")

# Anexos
input_arquivo.send_keys(apresentacao)
input_arquivo.send_keys(escopo)
input_arquivo.send_keys(colaboradores)
```

### 7. Envio e Tratamento de Erros

```python
# Envio
driver.find_element(By.XPATH, "//div[@class='T-I J-J5-Ji aoO v7 T-I-atl L3']").click()

# Tratamento de erros
except Exception as e:
    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    driver.refresh()
```

### 8. Finalização

```python
driver.quit()
```

---

## Configuração Inicial

### Primeira Execução (Login)

```python
time.sleep(30)  # Aumente para 20-30 segundos na primeira execução
# Faça login manualmente
# Após login, reduza para:
time.sleep(4)
```

---

## Arquivo CSV de Exemplo

```csv
nome,e-mail,empresa
João Silva,joao@empresa.com,Empresa A
Maria Souza,maria@outro.com,Empresa B
```

---

## Personalização

### Mensagens

```python
elemento.send_keys(f"Novo assunto para {cliente['Nome']}")
corpo.send_keys(f"Novo corpo para {cliente['Empresa']}")
```

### Anexos

```python
novo_anexo = r"D:\\NovoCaminho\\Arquivo.pdf"
```

### Novos Campos

#### CSV:

```csv
nome,e-mail,empresa,telefone
```

#### Código:

```python
'Telefone': row['telefone']
corpo.send_keys(f"Contato: {cliente['Telefone']}")
```

---

## Solução de Problemas

| Problema               | Solução                                      |
|------------------------|----------------------------------------------|
| XPaths desatualizados  | Atualizar usando DevTools (F12)              |
| Bloqueio do Gmail      | Limitar a 350–500 e-mails por dia            |
| Erros de login         | Verificar caminho do perfil do Chrome        |
| Versão incompatível    | Usar webdriver-manager                       |
| Perfil não encontrado  | Criar perfil manualmente                     |

---

## Limitações e Considerações

- Não exceder limites de envio do Gmail  
- XPaths podem mudar com atualizações  
- Usar apenas para fins legítimos  
- Não compartilhar perfis de usuário  

---

## Licença

MIT License – Disponível gratuitamente no GitHub.  
Criado e disponibilizado por **Lilian Chaves**.
