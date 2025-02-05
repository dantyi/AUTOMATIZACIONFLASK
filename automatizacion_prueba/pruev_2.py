import pyautogui
import time
import docx

def type_text(text):
    pyautogui.typewrite(text, interval=0.05)
time.sleep(5)
pyautogui.hotkey('win', 's') 
time.sleep(1)
pyautogui.typewrite("lorem datos.docx", interval=0.1)
time.sleep(1)
pyautogui.press('enter') 
time.sleep(5)
pyautogui.hotkey('win', 'left')
time.sleep(5)
doc = docx.Document('lorem datos.docx')
text_to_type = "\n".join([para.text for para in doc.paragraphs])
pyautogui.hotkey('win', 's')
time.sleep(1)
pyautogui.typewrite("Word", interval=0.1)
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey('win', 'right') 
type_text(text_to_type)
