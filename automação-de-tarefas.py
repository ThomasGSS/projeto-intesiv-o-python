import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

#executar o codigo no jupyter.

# 1 - baixando o arquivo (tabela de vendas)

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)


pyautogui.click(x=1023, y=710, clicks=2)
time.sleep(2)


pyautogui.click(x=1100, y=904)
pyautogui.click(x=3288, y=411)
pyautogui.click(x=2716, y=1523) 
time.sleep(5) 

# 2 - acessando a tabela

import pandas as pd

tabela = pd.read_excel(r"C:\Users\thomass\Downloads\Vendas - Dez.xlsx")
display(tabela)

# 3 calculando os valores

faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)

# 4 mandando o email pro "chefe" com os valores

pyautogui.hotkey("ctrl", "t")

pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)


pyautogui.click(x=240, y=415)


pyautogui.write("thomasblood12@gmail.com")
pyautogui.press("tab") 

pyautogui.press("tab") 
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab")

texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Thomas gabriel
"""



pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")


