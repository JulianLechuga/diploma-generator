import os
import sys
import tkinter as tk
from tkinter import messagebox
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def leer_nombres(archivo_nombres):
    try:
        with open(archivo_nombres, 'r', encoding='utf-8') as file:
            nombres = [line.strip() for line in file.readlines()]
        return nombres
    except Exception as e:
        messagebox.showerror("Error", f"Error al leer el archivo de nombres: {e}")
        return []

def crear_certificados(nombres, plantilla_path):
    try:
        for nombre in nombres:
            doc = Document(plantilla_path)

            for paragraph in doc.paragraphs:
                if '[NOMBRE]' in paragraph.text:
                    for run in paragraph.runs:
                        if '[NOMBRE]' in run.text:
                            run.text = run.text.replace('[NOMBRE]', nombre)
                            # Aplicar la fuente Edwardian Script ITC
                            run.font.name = 'Edwardian Script ITC'
                            r = run._element
                            r.rPr.rFonts.set(qn('w:eastAsia'), 'Edwardian Script ITC')
                            run.font.size = Pt(28)  # Cambia el tamaño de la fuente si es necesario

            doc.save(f'Diploma_{nombre.replace(" ", "_")}.docx')
        messagebox.showinfo("Éxito", "Documentos creados exitosamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al crear los certificados: {e}")

def main():
    if len(sys.argv) != 2:
        messagebox.showerror("Error", "Arrastre y suelte un archivo de texto sobre este ejecutable.")
        return

    archivo_nombres = sys.argv[1]
    plantilla_path = 'Plantilla.docx'
    
    if not os.path.exists(plantilla_path):
        messagebox.showerror("Error", f"No se encontró la plantilla '{plantilla_path}'. Asegúrese de que está en la misma carpeta que el ejecutable.")
        return

    if not os.path.exists(archivo_nombres):
        messagebox.showerror("Error", f"No se encontró el archivo de nombres '{archivo_nombres}'.")
        return

    nombres = leer_nombres(archivo_nombres)
    if nombres:
        if plantilla_path:
            crear_certificados(nombres, plantilla_path)
        else:
            messagebox.showerror("Error", f"Asegúrese que exista un archivo llamado Plantilla.docx sobre el cúal se pueda trabajar.")
            return

if __name__ == "__main__":
    main()
