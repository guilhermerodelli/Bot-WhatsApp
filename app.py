"""
AUTOMATIZANDO MENSAGENS NO WPP
quais tec. preciso para resolver está demanda?
     -teclado (Pyautogui)
     -acesso ao site (webbrowser)
     -automatizar digitação (link wpp)
     
automatizar leitura de dados (opempyxl)

"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

#esperar 20seg para poder fazer acesso na web caso não estaje conectado
webbrowser.open(https://web.whatsapp.com%27/)
sleep(20)

#Ler planilhas e guardar infomações sobre nome e telefone.
workbook = openpyxl.load_workbook('planilha de contato1.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row= 2):
        nome = linha[0].value
        telefone = linha[1].value
        mensagem = f'Olá {nome} essa é uma mensagem automática de uma automatozação com PYTHON, se chegou até você é por que deu certo kkkkkk. '
        try:
            link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
            webbrowser.open(link_mensagem_wpp)
            sleep(10)

            seta = pyautogui.locateCenterOnScreen('seta.png')
            sleep(5)

            pyautogui.click(seta[0],seta[1])
            sleep(5)

                #sleep é a pausa de cada função

            pyautogui.hotkey('ctrl','w')
            sleep(2)

        except Exception as e:
            print(f'Não foi possível enviar mensagem para {nome}. erro:{str(e)}')
            with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{telefone},{nome}\n')



#Criar links personalizados edo wpp e enviar mensagens para cada cliente com base nos dados