#vai lembrar os cliente de cortar o cabelo a cada 10 dias
#pegando os dados de uma lista
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open("https://web.whatsapp.com")
sleep(20)

#ler planilha
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #acessando o nome, telefone e vencimento
    nome = linha[1].value
    telefone = linha[2].value
    vencimento = linha[3].value
    mensagem = f"""
    #- Olá {nome}
    #- Apenas testando bot
    #- Teste de Data {vencimento}
                
            #####Bot Em construção por Dvison#######            """
    #entrando no whatsapp e dando o clique na seta de enviar mensagem
    try:
        link_mensagem = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
        webbrowser.open(link_mensagem)
        sleep(40)
        seta = pyautogui.locateCenterOnScreen("seta.png")
        sleep(5)
        clicar = pyautogui.click(seta)
        sleep(2)
        pyautogui.hotkey("ctrl", "w")

    except:
        print('não foi possivel enviar para {}'.format(nome))
        with open("erros.csv", "a", newline="", encoding="utf-8") as arquivo:
            arquivo.write(f"{nome}, {telefone}")