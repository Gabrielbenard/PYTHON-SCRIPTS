import pyautogui
import pyautogui as pa

opcao = pa.confirm("Clique no botão desejado", buttons=["Excel", "Word", "Notepad"])




if opcao =="Excel":
    print("Escolhido foi Excel")

    pa.hotkey("win","r")
    pa.sleep(2)
    pa.typewrite("Excel")
    pa.press("Enter")

    pa.sleep(3)
    pa.click(x=280, y=260)
    pa.typewrite("Escolhi o Excel")

elif opcao =="Word":
    print("Escolhido foi Word")

    pa.hotkey("win", "r")
    pa.sleep(2)
    pa.typewrite("winword")
    pa.press("Enter")

    pa.sleep(3)
    pa.click(x=280, y=260)
    pa.typewrite("Escolhi o Word")

elif opcao =="Notepad":
    print("Escolhido foi Notepad")

    pa.hotkey("win", "r")
    pa.sleep(2)
    pa.typewrite("Notepad")
    pa.press("Enter")
    pa.sleep(3)
    pa.typewrite("Escolhi o notepad")

else:
    print("Nenhuma opção foi escolhida")

