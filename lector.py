import cv2
from pyzbar.pyzbar import decode, ZBarSymbol
import win32com.client
import pyperclip
import keyboard
from datetime import datetime

MATRICULAS_SERVICIO = ["0", "68342", "68343", "68344", "68345", "68346", "68347", "68348", "68349", "68350"]

# Inicializa Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = 1

# Abre el archivo de Excel
workbook = excel.Workbooks.Open(r"C:\Users\Jaime\OneDrive\Documentos\Escuela\Servicio\Prueba.xlsm")
worksheet = workbook.Sheets(1)

# Inicializa la captura de video
captura = cv2.VideoCapture(0)
captura.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
captura.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)

while captura.isOpened():
    ret, frame = captura.read()
    if not ret:
        break

    # Procesamiento de la imagen
    scaled_frame = cv2.resize(frame, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
    gray = cv2.cvtColor(scaled_frame, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    processed = clahe.apply(gray)
    processed = cv2.GaussianBlur(processed, (5, 5), 0)

    try:
        codigo_barras = decode(processed, symbols=[ZBarSymbol.CODE39])

        for codigo in codigo_barras:
            x, y, w, h = codigo.rect
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
            texto_decodificado = codigo.data.decode("utf-8")
            print(f"Texto detectado: {texto_decodificado}")

            if excel.Visible:
                if texto_decodificado in MATRICULAS_SERVICIO:
                    excel.Run("Botón1_Haga_clic_en")
                else:
                    # Insertar el texto en la celda actual
                    selected_cell = worksheet.Application.ActiveCell
                    selected_cell.Value = texto_decodificado
                    print(f"Texto '{texto_decodificado}' pegado en {selected_cell.Address}")
                    
                    # Insertar la hora en la celda de al lado
                    adjacent_cell = selected_cell.Offset(-1, 0)
                    current_time = datetime.now().strftime("%H:%M:%S")
                    adjacent_cell.Value = current_time
                    print(f"Hora '{current_time}' añadida en {adjacent_cell.Address}")
            else:
                pyperclip.copy(texto_decodificado)
                print(f"Texto copiado al portapapeles: {texto_decodificado}")
                keyboard.press_and_release('ctrl+v')
                print("Texto pegado en la ubicación actual del cursor")

    except Exception as e:
        print(f"Error al decodificar: {e}")

    cv2.imshow("Captura", frame)
    if cv2.waitKey(1) & 0xFF == 27:
        break

captura.release()
cv2.destroyAllWindows()
