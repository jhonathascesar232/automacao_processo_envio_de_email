import pyautogui
import pyperclip
import time
import logging
import pandas as pd
from pprint import pprint

logging.basicConfig(
    filename='logs.log',
    level=logging.DEBUG,
    format='%(levelname)s %(asctime)s: %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S'
)
pyautogui.PAUSE = 5
# Tempo
MINUTO = 60
# time.sleep(5)
# logging.info(pyautogui.position())
path_file = r'C:\Users\shrks\Downloads\Vendas - Dez.xlsx'
tabela = pd.read_excel(path_file)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

# print(f'{tabela}\nFaturamento: {faturamento}\nQuantidade: {quantidade}')
# logging.info('\n{0}'.format(tabela.head()))

# passo 5 - Enviar E-mail para a diretoria;
# print('iniciando t 10')
# time.sleep(10)
# print('fim t 10')
# link_email = 'https://mail.google.com/'
# pyperclip.copy(link_email)
# pyautogui.hotkey('ctrl', 't')
# pyautogui.hotkey('ctrl', 'v')
# pyautogui.press('enter')
# print('iniciando t 2M')
# time.sleep(MINUTO*2)
# print('fim t 2M')
time.sleep(3)

email_de_destino = 'jhonathas.cesar.f.da.silva@gmail.com'
assunto_do_email = 'Relatório de Vendas'

mensagem_do_email = f'''
Faturamento: {faturamento};
Quantidade: {quantidade};
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
