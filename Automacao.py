import pyautogui
import pyperclip
import time
pyautogui.PAUSE = 4

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy('https://drive.google.com/drive/my-drive')
#pyautogui.write('https://www.facebook.com/campaign/landing.php?&campaign_id=1661784635&extra_1=s|c|320269343424|b|facebook|&placement=&creative=320269343424&keyword=facebook&partner_id=googlesem&extra_2=campaignid%3D1661784635%26adgroupid%3D63686354323%26matchtype%3Db%26network%3Dg%26source%3Dnotmobile%26search_or_content%3Ds%26device%3Dc%26devicemodel%3D%26adposition%3D%26target%3D%26targetid%3Dkwd-11403291%26loc_physical_ms%3D1031846%26loc_interest_ms%3D%26feeditemid%3D%26param1%3D%26param2%3D&gclid=CjwKCAiA4KaRBhBdEiwAZi1zzhGSobmiHB0v9fukd0I-CuX3_9WbYtB7P5J4RN_TFD3y1eElahdHFhoCbb4QAvD_BwE')
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
#pyautogui.write('https://mail.google.com/mail/u/0/#inbox')
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
pyperclip.copy("AGORA FOOOOI")
#pyautogui.write("AGORA FOOOI")
pyautogui.hotkey("ctrl", "v") # escrever o assunto
pyautogui.press("tab") #pular pro corpo do email

texto = f"""
Paçoca, avião, sofázão
Este e-mail foi enviado automaticamente através de um código

Os dados abaixo são fictícios
O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

bjs"""

pyperclip.copy(texto)
#pyautogui.write(texto)
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