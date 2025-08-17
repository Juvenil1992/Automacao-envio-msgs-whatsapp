from urllib.parse import quote
import openpyxl
import webbrowser
import time
import pyautogui as pg

webbrowser.open('https://web.whatsapp.com/')
time.sleep(20)

#Ler Planilha e Infomações de contatos

workbook = openpyxl.load_workbook('clientes.xltx') #aqui você coloca o nome da planilha salva com os contatos.

pagina_clientes = workbook['Planilha1'] #qual aba está as informações.


for linha in pagina_clientes.iter_rows(min_row=2):

    nome = linha[0].value
    telefone = linha[1].value


    mensagem = f'Olá {nome}, Testando Automação Teste 2.'



    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        time.sleep(12)
        seta = pg.locateCenterOnScreen('seta_enviar.png',confidence= 0.8)
        time.sleep(2)
        pg.click(seta[0],seta[1])
        time.sleep(2)
        pg.hotkey('ctrl','w')
        time.sleep(2)
    except:
        print(f'Não Foi Possivel Manda Mensagem para {nome}')
        with open('Erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome} , {telefone}')