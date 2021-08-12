import pyautogui
import pyperclip
import logging
import time
import pandas as pd

from pprint import pprint

# set up de logging
logging.basicConfig(
    filename='logs.log',
    level=logging.DEBUG,
    format='%(levelname)s %(asctime)s: %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S'
)


# set up
pyautogui.PAUSE = 2

# passo 1 - entrar no sistema
pyautogui.press('win')
pyautogui.write('ch')
pyautogui.press('enter')

link = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
pyperclip.copy(link)

pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

try:
    # passo 2 - Navegar no sistema
    time.sleep(45)
    pyautogui.click(x=319, y=281, clicks=2)
    time.sleep(30)
    # passo 3 - Baixar o arquivo de vendas;
    pyautogui.click(x=347, y=289)  # no arquivo
    pyautogui.click(x=823, y=196)  # nos pontos
    pyautogui.click(x=551, y=597)  # no nome fazer download
    # passo 4 - Calcular o Faturamento e a quantidade de produtos vendidos;
    path_file = r'C:\Users\shrks\Downloads\Vendas - Dez.xlsx'
    tabela = pd.read_excel(path_file)
    faturamento = tabela['Valor Final'].sum()
    quantidade = tabela['Quantidade'].sum()
    # passo 5 - Enviar E-mail para a diretoria;
    link_email = 'https://mail.google.com/'
    pyperclip.copy(link_email)
    pyautogui.hotkey('ctrl', 't')
    pyautogui.hotkey(
        'ctrl', 'v'
    )
    pyautogui.press('enter')
    time.sleep(120)
    # escrever e-mail

    email_de_destino = 'jhonathas.cesar.f.da.silva@gmail.com'
    assunto_do_email = 'Relatório de Vendas'

    mensagem_do_email = f'''
	Faturamento: {faturamento:,.2f};
	Quantidade: {quantidade:,};
	'''
    pyautogui.click(x=86, y=186)
    pyperclip.copy(email_de_destino)
    pyautogui.hotkey(
        'ctrl', 'v'
    )
    pyautogui.press('tab')
    pyperclip.copy(assunto_do_email)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    pyperclip.copy(mensagem_do_email)
    pyautogui.hotkey('ctrl', 'v')
    # Envia email
    pyautogui.hotkey('ctrl', 'enter')
except Exception as e:
    logging.info('Erro: {0}'.format(e))
