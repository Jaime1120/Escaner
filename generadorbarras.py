import tkinter as tk
from tkinter import filedialog, messagebox
import barcode
from barcode.writer import SVGWriter
import os

# Función para generar códigos de barras en formato SVG
def generar_codigos():
    codigos = entrada_codigo.get("1.0", tk.END).strip().split("\n")
    if not codigos or all(codigo.strip() == "" for codigo in codigos):
        messagebox.showerror("Error", "El campo de códigos no puede estar vacío.")
        return
    
    directorio = filedialog.askdirectory()
    if not directorio:
        return
    
    for codigo in codigos:
        codigo = codigo.strip()
        if not codigo:
            continue
        try:
            code39 = barcode.get_barcode_class('code39')
            barcode_obj = code39(codigo, writer=SVGWriter())
            file_path = os.path.join(directorio, f"{codigo}")
            barcode_obj.save(file_path)
            messagebox.showinfo("Éxito", f"Código de barras guardado en: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar el código de barras para {codigo}: {e}")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Generador de Código de Barras CODE 39")
root.geometry("400x300")

etiqueta = tk.Label(root, text="Ingrese los códigos (uno por línea):")
etiqueta.pack(pady=10)

entrada_codigo = tk.Text(root, width=40, height=10)
entrada_codigo.pack(pady=5)

boton_generar = tk.Button(root, text="Generar Códigos de Barras", command=generar_codigos)
boton_generar.pack(pady=20)

root.mainloop()
