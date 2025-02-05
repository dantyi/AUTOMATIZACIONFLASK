import os
import time
import pyautogui
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'supersecretkey'

# Configuración de uploads
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf', 'docx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Crear directorio de uploads
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def automatizacion_pyautogui(filename):
    try:
        _, extension = os.path.splitext(filename)
        extension = extension.lower()[1:]  # Eliminar el punto y convertir a minúsculas

        pyautogui.hotkey('win', 'r')
        pyautogui.write('notepad')
        pyautogui.press('enter')
        time.sleep(1)
        
        # Mensaje base
        mensaje = f"Archivo subido: {filename}\nHora: {time.ctime()}\nTipo: "
        
        # Clasificación por tipo de archivo
        if extension == 'pdf':
            mensaje += "PDF\nAcción: Abrir en visor PDF"
            # Simular apertura de Adobe Acrobat
            pyautogui.hotkey('win', 'r')
            pyautogui.write('AcroRd32')
            pyautogui.press('enter')
            time.sleep(2)
            
        elif extension == 'docx':
            mensaje += "DOCX\nAcción: Abrir en Word"
            # Simular apertura de Word
            pyautogui.hotkey('win', 'r')
            pyautogui.write('winword')
            pyautogui.press('enter')
            time.sleep(2)
            
        elif extension in ['png', 'jpg', 'jpeg']:
            mensaje += "IMAGEN\nAcción: Abrir en visor de imágenes"
            # Simular apertura de visor de fotos
            pyautogui.hotkey('win', 'r')
            pyautogui.write('mspaint')
            pyautogui.press('enter')
            time.sleep(2)
            
        else:
            mensaje += "Otro tipo de archivo"
        
        # Escribir en el Bloc de Notas
        pyautogui.write(mensaje)
        
        # Tomar screenshot
        screenshot = pyautogui.screenshot()
        screenshot_name = f"screenshot_{extension}_{int(time.time())}.png"
        screenshot_path = os.path.join(app.config['UPLOAD_FOLDER'], screenshot_name)
        screenshot.save(screenshot_path)
        
        # Cerrar aplicaciones abiertas
        for _ in range(2):  # Cerrar tanto el bloc como la aplicación abierta
            pyautogui.hotkey('alt', 'f4')
            pyautogui.press('enter')
            time.sleep(1)
        
        return screenshot_name
        
    except Exception as e:
        print(f"Error en automatización: {str(e)}")
        return None

@app.route('/')
def index():
    # Mostrar últimos screenshots
    screenshots = [f for f in os.listdir(UPLOAD_FOLDER) if f.startswith('screenshot')]
    return render_template('index.html', screenshots=screenshots)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No se seleccionó ningún archivo')
        return redirect(url_for('index'))
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No se seleccionó ningún archivo')
        return redirect(url_for('index'))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        screenshot_name = automatizacion_pyautogui(filename)
        if screenshot_name:
            flash(f'Archivo subido y automatización completada! Screenshot: {screenshot_name}', 'success')
        else:
            flash('Archivo subido pero falló la automatización', 'warning')
        
        return redirect(url_for('index'))
    else:
        flash('Tipo de archivo no permitido', 'error')
        return redirect(url_for('index'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)