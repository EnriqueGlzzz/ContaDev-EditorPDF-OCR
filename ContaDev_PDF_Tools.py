import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pytesseract
import subprocess
import time
import os
import sys

# --- UTILIDADES ---
def resolver_ruta(ruta_relativa):
    """ Obtiene la ruta absoluta para que funcione en VS Code y como .exe """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

# --- PASO 1: OCR ---
def ejecutar_ocr():
    """ Convierte PDF de imagen a PDF buscable """
    # Conexión con el motor instalado
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    
    ruta_entrada = filedialog.askopenfilename(title="Selecciona el PDF (Imagen)", filetypes=[("Archivos PDF", "*.pdf")])
    if not ruta_entrada: return

    ruta_salida = filedialog.asksaveasfilename(defaultextension=".pdf", title="Guardar PDF con OCR como...")
    if not ruta_salida: return

    try:
        messagebox.showinfo("Procesando", "El OCR está trabajando. Esto puede tardar unos segundos...")
        # Ejecuta ocrmypdf (debe estar instalado en Windows)
        subprocess.run(["ocrmypdf", "--skip-text", ruta_entrada, ruta_salida], check=True)
        messagebox.showinfo("Éxito", "OCR completado. Ahora puedes usar este archivo en el PASO 2.")
    except Exception as e:
        messagebox.showerror("Error de OCR", f"Asegúrate de tener instalado Tesseract y ocrmypdf.\nDetalle: {e}")

# --- PASO 2: EDICIÓN ---
def procesar_pdf():
    buscar = entrada_buscar.get()
    reemplazar = entrada_reemplazar.get()
    
    try:
        h_inicio = int(entrada_inicio.get()) - 1
        h_fin = int(entrada_fin.get())
    except ValueError:
        messagebox.showwarning("Atención", "Ingresa números válidos para las hojas.")
        return

    if not buscar:
        messagebox.showwarning("Atención", "Escribe el texto a buscar.")
        return

    ruta_entrada = filedialog.askopenfilename(title="Selecciona el PDF para editar", filetypes=[("Archivos PDF", "*.pdf")])
    if not ruta_entrada: return

    ruta_salida = filedialog.asksaveasfilename(defaultextension=".pdf", title="Guardar resultado como...")
    if not ruta_salida: return

    try:
        doc = fitz.open(ruta_entrada)
        total_pags = len(doc)
        
        if h_inicio < 0 or h_fin > total_pags or h_inicio >= h_fin:
            messagebox.showerror("Error", f"Rango inválido. El PDF tiene {total_pags} páginas.")
            return

        cambios = 0
        for i in range(h_inicio, h_fin):
            pagina = doc[i]
            instancias = pagina.search_for(buscar)
            for inst in instancias:
                pagina.add_redact_annot(inst, fill=(1, 1, 1))
                pagina.apply_redactions()
                pagina.insert_text((inst.x0, inst.y1 - 2), reemplazar, fontsize=10, color=(0, 0, 0))
                cambios += 1

        doc.save(ruta_salida)
        doc.close()
        messagebox.showinfo("Éxito", f"Cambios realizados: {cambios}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un problema: {e}")

# --- PANTALLA DE INICIO (SPLASH) ---
def mostrar_splash():
    splash = tk.Tk()
    splash.overrideredirect(True) # Elimina barra de título
    
    # Centrar ventana
    ancho_v, alto_v = 500, 300
    pos_x = (splash.winfo_screenwidth() // 2) - (ancho_v // 2)
    pos_y = (splash.winfo_screenheight() // 2) - (alto_v // 2)
    splash.geometry(f"{ancho_v}x{alto_v}+{pos_x}+{pos_y}")

    try:
        # CORRECCIÓN: Usar el nombre exacto del archivo que tienes en la carpeta
        ruta_logo = resolver_ruta("Minegocio.png") 
        
        if os.path.exists(ruta_logo):
            img = Image.open(ruta_logo)
            img = img.resize((500, 300), Image.Resampling.LANCZOS)
            foto = ImageTk.PhotoImage(img)
            label = tk.Label(splash, image=foto, bg="white")
            label.image = foto 
            label.pack()
        else:
            # Texto de respaldo si la imagen no se encuentra
            tk.Label(splash, text="ContaDev\nCargando...", font=("Arial", 20), bg="#2c3e50", fg="white").pack(expand=True, fill="both")
    except Exception as e:
        tk.Label(splash, text="ContaDev", font=("Arial", 20), bg="#2c3e50", fg="white").pack(expand=True, fill="both")

    splash.update()
    time.sleep(3) # Tiempo que se muestra el logo
    splash.destroy()
# --- INTERFAZ PRINCIPAL ---
if __name__ == "__main__":
    mostrar_splash()

    root = tk.Tk()
    root.title("ContaDev V2 - OCR + Editor")
    root.geometry("450x600")
    root.configure(bg="#f0f0f0")

    # Sección OCR
    tk.Label(root, text="PASO 1: Preparar PDF (OCR)", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=(20,5))
    btn_ocr = tk.Button(root, text="Detectar Texto de Imágenes", command=ejecutar_ocr, bg="#3498db", fg="white", width=30)
    btn_ocr.pack(pady=5)

    tk.Label(root, text="------------------------------------------", bg="#f0f0f0").pack()

    # Sección Edición
    tk.Label(root, text="PASO 2: Buscar y Reemplazar", font=("Arial", 10, "bold"), bg="#f0f0f0").pack(pady=5)
    tk.Label(root, text="Texto a buscar:").pack()
    entrada_buscar = tk.Entry(root, width=35); entrada_buscar.pack()
    tk.Label(root, text="Reemplazar con:").pack(pady=(10,0))
    entrada_reemplazar = tk.Entry(root, width=35); entrada_reemplazar.pack()

    marco_hojas = tk.LabelFrame(root, text=" Rango de Hojas ", bg="#f0f0f0", padx=10, pady=10)
    marco_hojas.pack(pady=15)
    tk.Label(marco_hojas, text="Desde:", bg="#f0f0f0").grid(row=0, column=0)
    entrada_inicio = tk.Entry(marco_hojas, width=8); entrada_inicio.grid(row=0, column=1); entrada_inicio.insert(0, "1")
    tk.Label(marco_hojas, text="Hasta:", bg="#f0f0f0").grid(row=1, column=0)
    entrada_fin = tk.Entry(marco_hojas, width=8); entrada_fin.grid(row=1, column=1); entrada_fin.insert(0, "1")

    btn_edit = tk.Button(root, text="Procesar Cambios", command=procesar_pdf, bg="#2ecc71", fg="white", font=("Arial", 11, "bold"), padx=20, pady=10)
    btn_edit.pack(pady=20)

    root.mainloop()