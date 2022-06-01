import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 0.5

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

pyautogui.click(x=1070, y=725, clicks=2)
time.sleep(2)

pyautogui.click(x=1044, y=869)
pyautogui.click(x=3292, y=399)
pyautogui.click(x=2831, y=1539)
time.sleep(5)

import pandas as pd

tabela = pd.read_excel(r"C:\Users\Admin\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
  pyautogui.hotkey("ctrl", "t")
  pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
  pyautogui.hotkey("ctrl", "v")
  pyautogui.press("enter")
  time.sleep(5)

  pyautogui.click(x=256, y=431)
  time.sleep(2)

  pyautogui.write("jpsilvaterra.js@gmail.com")
  pyautogui.press("tab")
  pyautogui.press("tab")

  pyperclip.copy("Relatório de Vendas")
  pyautogui.hotkey("ctrl", "v")
  pyautogui.press("tab")

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R$: {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,} 

Abs
João Pedro S. Terra"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

  pyautogui.hotkey("ctrl", "enter")
