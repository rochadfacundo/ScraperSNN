import tkinter as tk
import threading
import sys
import os
import tempfile
import shutil
from tkinter import messagebox
from PIL import Image, ImageTk  # Pillow

# Rutas relativas
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__))))
from scraping import buscar_datos
from utils.excel import exportar_a_excel

# 🟢 Ruta del ícono compatible con PyInstaller
def obtener_ruta_icono():
    if hasattr(sys, '_MEIPASS'):
        origen = os.path.join(sys._MEIPASS, "assets", "logo.ico")
        destino = os.path.join(tempfile.gettempdir(), "logo.ico")
        shutil.copy(origen, destino)
        return destino
    return os.path.join("assets", "logo.ico")

# Inicializar ventana
root = tk.Tk()
root.title("Scraper SSN - Búsqueda por Matrícula")
root.geometry("550x470")
root.iconbitmap(obtener_ruta_icono())

# ✅ Logo superior
logo_path = os.path.join("assets", "logo.png")
if os.path.exists(logo_path):
    img = Image.open(logo_path)
    img = img.resize((100, 100))
    logo_img = ImageTk.PhotoImage(img)
    logo_label = tk.Label(root, image=logo_img)
    logo_label.pack(pady=10)

# 🔎 Entrada y botón
tk.Label(root, text="Matrícula del Productor:", font=("Arial", 12)).pack()
entry = tk.Entry(root, font=("Arial", 12), justify="center")
entry.pack(pady=5)

btn = tk.Button(
    root,
    text="Buscar",
    command=lambda: threading.Thread(target=ejecutar_busqueda, daemon=True).start(),
    font=("Arial", 12, "bold"),
    bg="#94c484",
    fg="white",
    activebackground="#7aa76f",
    relief="raised",
    bd=3,
    cursor="hand2",
    padx=10,
    pady=5
)
btn.pack(pady=10)

# 📝 Resultados
resultado = tk.Text(root, height=15, font=("Consolas", 10), state="disabled")
resultado.pack(padx=10, pady=10, fill="both", expand=True)

# 🔄 Búsqueda y exportación
def ejecutar_busqueda():
    matricula = entry.get()
    if not matricula.strip():
        messagebox.showwarning("Matrícula vacía", "Por favor ingresá un número de matrícula.")
        return
    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    resultado.insert(tk.END, "🔄 Buscando datos...\n")
    resultado.config(state="disabled")
    buscar_en_segundo_plano(matricula)

def buscar_en_segundo_plano(matricula):
    datos = buscar_datos(matricula)
    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    for k, v in datos.items():
        resultado.insert(tk.END, f"{k}: {v}\n")
    resultado.config(state="disabled")
    exportar_a_excel(datos)

root.mainloop()
