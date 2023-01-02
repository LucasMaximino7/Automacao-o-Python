import pandas as pd
import pyautogui
import pyperclip
import time
import clipboard

tabela = pd.read_excel(r"C:\Users\lucas_mn_oliveira\Desktop\Projeto.xlsx")
tabela_atualizada = pd.read_excel(r"C:\Users\lucas_mn_oliveira\Desktop\Projeto_NovoHorarioDeVOo.xlsx")

###vento navegantes
pyautogui.hotkey("ctrl","t")
pyautogui.write("previsao do tempo para segunda-feira navegantes sc")
pyautogui.hotkey("enter")
time.sleep(1.5)
pyautogui.doubleClick(x=398, y=298)
time.sleep(1)

pyautogui.hotkey("ctrl","c")
vento_navegantes = clipboard.paste()
pyautogui.hotkey("ctrl","w")


vento_navegantes.replace(" ","")

vento_navegantes = int(vento_navegantes)


###vento salvador

pyautogui.hotkey("ctrl","t")
pyautogui.write(r"previsao do tempo para terca-feira Salvador BA")
pyautogui.hotkey("enter")
time.sleep(1.9)
pyautogui.doubleClick(x=402, y=298)
time.sleep(1)

pyautogui.hotkey("ctrl","c")
vento_salvador = clipboard.paste()
pyautogui.hotkey("ctrl","w")


vento_salvador.replace(" ","")

vento_salvador = int(vento_salvador)

###se for menor permanece

if vento_navegantes < 37 and vento_salvador < 37 :
   
    pyautogui.hotkey("ctrl","t")
    pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
    pyautogui.hotkey("enter")
    time.sleep(1.5)
    pyautogui.click(x=164, y=165)
    time.sleep(1)
    pyautogui.write("mateusbaader@gmail.com")
    pyautogui.hotkey("enter")
    time.sleep(1)
    pyautogui.write("lucas_mn_oliveira@estudante.sesisenai.org.br")
    pyautogui.hotkey("enter")
    time.sleep(1)
    pyautogui.write("fabiano.luiz@edu.sc.senai.br")
    pyautogui.hotkey("enter")
    time.sleep(1)
   
    pyautogui.hotkey("tab")
    pyautogui.write(r"Companhia aéria AirBus - Comunicado ")
    pyautogui.hotkey("tab")
   
    hora_saida = tabela["hora_saida"]
   
    texto = f"""
    Olá caro cliente, viemos informar por este email que seu voo decola no horário previsto às {hora_saida[0]}
   
    """
   
    pyperclip.copy(texto)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.hotkey("enter")
   
    pyautogui.hotkey("win","r")
    time.sleep(1)
    pyautogui.write(r"C:\Users\lucas_mn_oliveira\Documents")
    time.sleep(1)
    pyautogui.hotkey("enter")
    time.sleep(0.5)
    pyautogui.click(x=915, y=461)
    pyautogui.write("d")
    pyautogui.hotkey("ctrl","c")
    time.sleep(0.5)
    pyautogui.click(x=273, y=743)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl","v")
    time.sleep(5)
    pyautogui.hotkey("ctrl","enter")
    time.sleep(2)
    pyautogui.hotkey("ctrl","w")
   
   
   
if vento_salvador > 37 or vento_salvador > 37:
   
    pyautogui.hotkey("ctrl","t")
    pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
    pyautogui.hotkey("enter")
    time.sleep(1.5)
    pyautogui.click(x=164, y=165)
    time.sleep(1)  
    pyautogui.write("mateusbaader@gmail.com")
    pyautogui.hotkey("enter")
    time.sleep(1)  
    pyautogui.write("lucas_mn_oliveira@estudante.sesisenai.org.br")
    pyautogui.hotkey("enter")
    time.sleep(1)    
    pyautogui.write("fabiano.luiz@edu.sc.senai.br")
    pyautogui.hotkey("enter")
    time.sleep(1)
    pyautogui.hotkey("tab")
    pyautogui.write(r"Companhia aéria AirBus - Comunicado ")
    pyautogui.hotkey("tab")
   
    hora_saida = tabela_atualizada["hora_saida"]
   
    texto = f"""
    Olá caro cliente, viemos informar que houve um imprevisto em seu voo, o horário de saída será alterado para {hora_saida[0]}
   
    """
   
    pyautogui.hotkey("win","r")
    time.sleep(1)
    pyautogui.write(r"C:\Users\lucas_mn_oliveira\Documents")
    time.sleep(1)
    pyautogui.hotkey("enter")
    time.sleep(0.5)
    pyautogui.click(x=915, y=461)
    pyautogui.write("d")
    pyautogui.hotkey("ctrl","c")
    time.sleep(0.5)
    pyautogui.click(x=273, y=743)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl","v")
    time.sleep(5)
    pyautogui.hotkey("ctrl","enter")
    time.sleep(2)
    pyautogui.hotkey("ctrl","w")