#Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException #Exceção ao não econtrar elemento na página

import time
import openpyxl #manipular excel

from decouple import config #para acesso ao arquivo .env com email senha

import smtplib #Email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

devices = [] #Lista dos celulares

def mail_config():
    #CONFIGURAÇÕES DE EMAIL
    global server, smtp_username, smtp_password, email

    email = config('EMAIL')
    smtp_server =  'smtp.gmail.com' #PARA GMAIL
    smtp_port = 587
    smtp_username = email
    smtp_password = config('PASSWORD')

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_username, smtp_password)
    
def send_mail(): #Envia email com planilha anexada
        #configuração da mensagem
    
        message = MIMEMultipart()
        message['From'] = email
        message['To'] = email
        message['Subject'] = 'Preço dos celulares importados - (Ustore)'

        #conteúdo da mensagem
        messageBody = "Segue em anexo a planilha:"
        message.attach(MIMEText(messageBody, 'plain'))
    
        #anexo do arquivo da planilha
        attachmentFile = "./Dispositivos.xlsx"
        with open(attachmentFile, 'rb') as file:
            attachment = MIMEApplication(file.read(), _subtype='xlsx')
            attachment.add_header('content-disposition', 'attachment', filename='planilha.xlsx')
            message.attach(attachment)

        #enviar o email
        server.sendmail(smtp_username, email, message.as_string())
        server.quit()

def webdriver_config(): #Inicialização do selenium
    #CONFIGURAÇÃO DO SELENIUM COM CHROME
    global driver
    chrome_driver_path = r'./chromedriver.exe' #webdriver do chrome
    service = Service(chrome_driver_path)      #caminho do webdriver
    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--start-maximized") 
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get("https://telefonesimportados.netlify.app/") #url do site

def get_all_devices_on_page(): #Pega as informações dos celulares da página atual
    #identifica o conteúdo a ser analisado
    elements = driver.find_elements(By.CLASS_NAME, 'col-md-3')

    #pega o texto do elemento selecionado e separa por espaço
    for element in elements:
        text = element.text
        lines = text.split('\n')

        if len(lines) >= 3:
            if '$' in lines[1]:
                name = lines[0]
                values = lines[1].split()
                price = values[0]

        #adiciona nome e valor dos dispositivos ao array
        devices.append((name, price))
    
def create_and_save_sheet(): #Cria e salva planilha com nome e preço de todos celulares
            #Criar planinlha
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            # adicionar dados
            i = 2
            sheet["A1"] = "Dispositivo"
            sheet["B1"] = "Preço"

            for device in devices:
                sheet[f"A{i}"] = device[0]
                sheet[f"B{i}"] = device[1]
                i+=1
                
            #Salvar planilha
            workbook.save('Dispositivos.xlsx')
            workbook.close()

def start():
    
    print('=====================\n\033[35mExecução iniciada\033[90m')


    print('1. \033[34mConfigurando selenium...\033[90m')
    webdriver_config() #inicialização do webdrier
    print('     > \033[32mSelenium configurado')
    print('2. \033[34mConfigurando Email...\033[90m')
    mail_config()
    print('     > \033[32mEmail configurado')

    print('3. \033[34mNavegando pelas páginas...\033[90m')
    index = 1
    time.sleep(1)
    while True: #Loop para paginação
        #Verifica se existe o botão next page
        try:
            #Caso exista:
            next_page = driver.find_element(By.CSS_SELECTOR, '[aria-label="Next"]')
            # time.sleep(1)     #Desacelerar, caso necessário
            get_all_devices_on_page()        #Pega todos dispositivos da página
            next_page.click()    #Avança para a próxima
            print(f'    > \033[32mPágina {index} concluída\033[90m')
            index+=1

        except NoSuchElementException:
            #Caso não exista: = Ultima página
            get_all_devices_on_page()        #Chama função novamente para pegar dispositivos da ultima página
            print(f'    > \033[32mPágina {index} concluída\033[90m')
            print('4. \033[34mCriando Planilha...\033[90m')
            create_and_save_sheet()       #Cria e salva planilha com os nomes e valores dos celulares
            print('     > \033[32mPlanilha Salva')
            print('5. \033[34mEnviando Email...\033[90m')
            send_mail()         #envia planilha por email
            print('     > \033[32mEmail enviado')
            
            break     #Finaliza loop de paginação
    driver.quit() #Finaliza execução

    print('\n\033[35mExecução finalizada\n\033[90m=====================')

start()