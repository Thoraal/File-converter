import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from docx2pdf import convert
import fitz
import os
from PIL import Image

destino_path = ""

def convertir():
    global archivo_path_label, destino_path, mensaje_label

    archivo_path = filedialog.askopenfilename(filetypes=[("Todos los archivos", "*.*"), ("Archivos Word", "*.docx"), ("Archivos PDF", "*.pdf")])
    if not archivo_path:
        return

    archivo_nombre = os.path.basename(archivo_path)
    archivo_prefijo = archivo_nombre[:4]  # Obtener las primeras 4 letras del nombre del archivo

    archivo_path_label.config(text=archivo_nombre)

    destino_path = os.path.dirname(archivo_path)

    if archivo_path.endswith(".docx"):
        pdf_path = f"{destino_path}/{archivo_prefijo}_documento.pdf"
        if os.path.exists(pdf_path):
            if not messagebox.askyesno("Advertencia", "Ya existe un archivo PDF con el mismo nombre. ¿Desea sobrescribirlo?"):
                return
        try:
            convert(archivo_path, pdf_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error al convertir a PDF: {e}")
            return
    else:
        pdf_path = archivo_path

    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)
    progress_bar["maximum"] = total_pages
    for i in range(total_pages):
        page = pdf_document[i]
        image_path = f"{destino_path}/{archivo_prefijo}_pagina_{i+1}.png"
        if os.path.exists(image_path):
            if not messagebox.askyesno("Advertencia", f"Ya existe un archivo PNG con el mismo nombre ({os.path.basename(image_path)}). ¿Desea sobrescribirlo?"):
                continue
        image = page.get_pixmap()
        image = Image.frombytes("RGB", (image.width, image.height), image.samples)
        image = image.resize((int(image.width * 600 / 72), int(image.height * 600 / 72)), resample=Image.LANCZOS)
        image.save(image_path, "PNG")
        progress_bar["value"] = i + 1
        root.update_idletasks()
    pdf_document.close()

    open_directory_btn.config(state="normal")
    mensaje_label.config(text="La conversión se realizó con éxito.")

def abrir_directorio():
    global destino_path
    os.startfile(destino_path)

root = tk.Tk()
root.title("Toral")
root.geometry("240x210")

btn_convertir = tk.Button(root, text="Convertir a PDF y PNG", command=convertir)
btn_convertir.pack(pady=10)

progress_bar = Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress_bar.pack(pady=10)

archivo_path_label = tk.Label(root, text="Ningún archivo seleccionado")
archivo_path_label.pack(pady=10)

open_directory_btn = tk.Button(root, text="Abrir directorio", command=abrir_directorio, state="disabled")
open_directory_btn.pack(pady=10)

mensaje_label = tk.Label(root, text="")
mensaje_label.pack(pady=10)

root.mainloop()
