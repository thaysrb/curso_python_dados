# Importar as bibliotecas
import pyautogui # faz a automação do mouse e do teclado
import time # controla o tempo do nosso programa
import pyperclip # ela permite a gente copiar e colar com o python
import pandas as pd

pyautogui.PAUSE = 1
pyautogui.alert('O programa vai começar não use nada do seu computador')

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

# Passo 1: Entrar no sistema (link do google drive)
pyautogui.hotkey('ctrl', 't')
link = 'https://drive.google.com/drive/shared-with-me'
pyautogui.write('https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym')
pyautogui.press("enter")
# da ctrl+c e ctrl+v
pyperclip.copy(link)
#cola o link
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")
# esperar um pouquinho
time.sleep(5)

# Passo 2: Entrar na pasta da aula 1
pyautogui.click(335, 295, clicks=2) 
# Para clicar -> pyautogui.click

# Passo 3: Fazer o download da Base de vendas
pyautogui.click(371, 402)
pyautogui.click(1161, 155)
pyautogui.click(1014, 534)
time.sleep(10)

# Passo 4: Calcular os indicadores (Faturamento e a Quantidade de Produtos)

df = pd.read_excel(r"D:\Downloads\Vendas - Dez.xlsx")
#display(df)
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()


# Passo 5: Entrar no e-mail
# email da diretoria: thays.rodrigues.barboza@gmail.com

pyautogui.hotkey('ctrl', 't')
#link = 'https://mail.google.com/'
pyautogui.write('https://mail.google.com/')
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(64, 183)
pyautogui.write("thays.rodrigues.barboza@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyperclip.copy("assunto")

pyautogui.press("tab")
texto_email = f"""
Prezados, bom dia

O faturamento de ontem foi: R$ {faturamento:,.2f}
A quantidade de produtos foi: {qtde_produtos,}

Att,
Thays"""
pyperclip.copy(texto_email)
pyautogui.hotkey("ctrl", 'v')
pyautogui.hotkey("ctrl", "enter")
assunto = "Relatório de vendas de ontem"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
# esperar um pouquinho
time.sleep(5)





#pyautogui.press("win")
#pyautogui.write("chrome")
#pyautogui.press("enter")

#encontrar a posição de um "botão"
#time.sleep(5)
#pyautogui.position()