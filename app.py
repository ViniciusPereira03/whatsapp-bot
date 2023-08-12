import pywhatkit
import keyboard
import time
import platform
import pandas as pd
from datetime import datetime

system_info = platform.uname()

planilha = pd.read_excel('contatos.xlsx')

for index, row in planilha.iterrows():
    nome = row['Nome']
    celular = '+' + str(row['Celular'])
    mensagem = 'Olá ' + nome + ', ' + row['Mensagem']
    
    pywhatkit.sendwhatmsg(celular, mensagem, datetime.now().hour, datetime.now().minute + 1)
    del celular
    time.sleep(30)

    if system_info.system == 'Windows':
        keyboard.press_and_release('ctrl + w')
    elif system_info.system == 'Darwin':
        keyboard.press_and_release('command + w')
        keyboard.press_and_release('enter')
