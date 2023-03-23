import pyautogui
import pyperclip
import time
pyautogui.PAUSE = 4

# Passo 1: Entrar no link
pyautogui.hotkey("ctrl", "t")
pyperclip.copy('https://drive.google.com/drive/my-drive') #você deve estar conectado com a sua conta google no navegador
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=427, y=555, clicks=2)
time.sleep(4)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=512, y=357)
pyautogui.click(x=1154, y=189)
pyautogui.click(x=1062, y=545)
time.sleep(5)

# Passo 4: Calcular os indicadores
import pandas as pd

tabela = pd.read_excel(r"/home/ariel/Downloads/Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.write('https://mail.google.com/mail/u/0/#inbox')
pyautogui.press("enter")
time.sleep(12)

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=101, y=208)
time.sleep(8)

pyautogui.write("ariellelopes888@gmail.com")
    
pyautogui.press("tab") # seleciona o email
# escreve outro email
# tab
# escreve outro email
# tab
pyautogui.press("tab") # pula pro campo de assunto
pyperclip.copy("ok")
pyautogui.write("ok")
pyautogui.hotkey("ctrl", "v") # escrever o assunto
pyautogui.press("tab") #pular pro corpo do email

texto = f"""
Paçoca, avião, sofázão
Este e-mail foi enviado automaticamente através de um código

Os dados abaixo são fictícios
O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Teste ok"""

pyperclip.copy(texto)
pyautogui.write(texto)
pyautogui.hotkey("ctrl", "v")

# clicar no botão enviar

# apertar Ctrl Enter
pyautogui.hotkey("ctrl", "enter")

## para descobrir a posição do cursor
time.sleep(10)
pyautogui.position()

# Como instalar
#!pip install pyautogui
#!pip install pyperclip