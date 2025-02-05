import pyautogui
import time

# Espera unos segundos para que puedas cambiar a la ventana de Word
time.sleep(5)

# Abre el archivo "lorem datos.docx"
pyautogui.hotkey('win', 's')  # Abre la búsqueda de Windows
time.sleep(1)
pyautogui.typewrite("lorem datos.docx", interval=0.1)  # Escribe el nombre del archivo
time.sleep(1)
pyautogui.press('enter')  # Abre el archivo
time.sleep(1)
pyautogui.hotkey('win', 'left')

# Espera a que se abra el archivo
time.sleep(5)

# Abre un nuevo documento de Word
pyautogui.hotkey('win', 's')  # Abre la búsqueda de Windows
time.sleep(1)
pyautogui.typewrite("Word", interval=0.1)  # Escribe "Word"
time.sleep(1)
pyautogui.press('enter')  # Abre un nuevo documento de Word
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey('win', 'right') 
# Espera a que se abra el nuevo archivo
time.sleep(5)


