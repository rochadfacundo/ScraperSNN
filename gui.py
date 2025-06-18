import tkinter as tk
import threading
import sys
import os
import tempfile
import shutil
from tkinter import messagebox
from PIL import Image, ImageTk

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__))))
from scraping import buscar_datos
from utils.excel import exportar_a_excel

def obtener_ruta_icono():
    if hasattr(sys, '_MEIPASS'):
        origen = os.path.join(sys._MEIPASS, "assets", "logo.ico")
        destino = os.path.join(tempfile.gettempdir(), "logo.ico")
        shutil.copy(origen, destino)
        return destino
    return os.path.join("assets", "logo.ico")

root = tk.Tk()
root.title("Scraper SSN - Búsqueda por Matrícula")
root.geometry("550x520")
root.iconbitmap(obtener_ruta_icono())

# Logo
logo_path = os.path.join("assets", "logo.png")
if os.path.exists(logo_path):
    img = Image.open(logo_path)
    img = img.resize((100, 100))
    logo_img = ImageTk.PhotoImage(img)
    logo_label = tk.Label(root, image=logo_img)
    logo_label.pack(pady=10)

# Modo de búsqueda
modo_busqueda = tk.StringVar(value="una")
frame_radio = tk.Frame(root)
frame_radio.pack()

tk.Radiobutton(frame_radio, text="Buscar una matrícula", variable=modo_busqueda, value="una", command=lambda: alternar_inputs()).grid(row=0, column=0, padx=10)
tk.Radiobutton(frame_radio, text="Buscar un rango", variable=modo_busqueda, value="rango", command=lambda: alternar_inputs()).grid(row=0, column=1, padx=10)

# Inputs
frame_inputs = tk.Frame(root)
frame_inputs.pack(pady=10)

entry_una = tk.Entry(frame_inputs, font=("Arial", 12), justify="center", width=30)
entry_desde = tk.Entry(frame_inputs, font=("Arial", 12), justify="center", width=12)
entry_hasta = tk.Entry(frame_inputs, font=("Arial", 12), justify="center", width=12)

entry_una.grid(row=0, column=0, columnspan=2, pady=5)

def alternar_inputs():
    for widget in frame_inputs.winfo_children():
        widget.grid_remove()
    if modo_busqueda.get() == "una":
        entry_una.grid(row=0, column=0, columnspan=2, pady=5)
    else:
        tk.Label(frame_inputs, text="Desde:", font=("Arial", 10)).grid(row=0, column=0, padx=5, pady=5)
        entry_desde.grid(row=0, column=1, padx=5)
        tk.Label(frame_inputs, text="Hasta:", font=("Arial", 10)).grid(row=1, column=0, padx=5, pady=5)
        entry_hasta.grid(row=1, column=1, padx=5)

# Botón
btn = tk.Button(
    root,
    text="Buscar",
    command=lambda: ejecutar_busqueda(),
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

# Área de resultados
resultado = tk.Text(root, height=12, font=("Consolas", 10), state="disabled")
resultado.pack(padx=10, pady=10, fill="both", expand=True)

# Lógica
def ejecutar_busqueda():
    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    resultado.insert(tk.END, "🔄 Buscando datos...\n")
    resultado.config(state="disabled")

    if modo_busqueda.get() == "una":
        matricula = entry_una.get().strip()
        if not matricula:
            messagebox.showwarning("Campo vacío", "Ingresá una matrícula.")
            limpiar_estado()
            return
        threading.Thread(target=lambda: buscar_y_exportar(matricula, True), daemon=True).start()
    else:
        desde = entry_desde.get().strip()
        hasta = entry_hasta.get().strip()
        if not (desde.isdigit() and hasta.isdigit()):
            messagebox.showerror("Error", "Ingresá solo números válidos.")
            limpiar_estado()
            return
        inicio, fin = int(desde), int(hasta)
        if inicio > fin:
            messagebox.showerror("Error", "El valor inicial debe ser menor o igual al final.")
            limpiar_estado()
            return
        threading.Thread(target=lambda: buscar_rango(inicio, fin), daemon=True).start()

def limpiar_estado():
    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    resultado.config(state="disabled")

def buscar_y_exportar(matricula: str):
    datos = buscar_datos(matricula)

    # Verificar si todos los valores son "--"
    if all(valor == "--" for valor in datos.values()):
        resultado.config(state="normal")
        resultado.insert(tk.END, f"⚠️ Matrícula {matricula} sin datos útiles.\n")
        resultado.see(tk.END)
        resultado.config(state="disabled")
        return

    exportar_a_excel(datos, matricula)
    resultado.config(state="normal")
    resultado.insert(tk.END, f"✅ Matrícula {matricula} extraída.\n")
    resultado.see(tk.END)
    resultado.config(state="disabled")


def buscar_rango(inicio: int, fin: int):
    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    resultado.insert(tk.END, "🔄 Buscando datos...\n")
    resultado.config(state="disabled")

    for matricula in range(inicio, fin + 1):
        buscar_y_exportar(str(matricula))
    limpiar_estado()
    resultado.config(state="normal")
    resultado.insert(tk.END, f"✅ Extracción finalizada: {fin - inicio + 1} matrículas.\n")
    resultado.config(state="disabled")

root.mainloop()
