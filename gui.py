import tkinter as tk
import threading
import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__))))
from tkinter import messagebox
from scraping import buscar_datos
from utils.excel import exportar_a_excel

# Interfaz
def ejecutar_busqueda():
    matricula = entry.get()
    if not matricula.strip():
        messagebox.showwarning("Matrícula vacía", "Por favor ingresá un número de matrícula.")
        return
    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    resultado.insert(tk.END, "🔄 Buscando datos...\n")
    resultado.config(state="disabled")
    threading.Thread(target=buscar_en_segundo_plano, args=(matricula,), daemon=True).start()

def buscar_en_segundo_plano(matricula):
    datos = buscar_datos(matricula)

    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    for k, v in datos.items():
        resultado.insert(tk.END, f"{k}: {v}\n")
    resultado.config(state="disabled")

    # 🔽 Exportar a Excel
    exportar_a_excel(datos)


root = tk.Tk()
root.title("Scraper SSN - Búsqueda por Matrícula")
root.geometry("500x400")

tk.Label(root, text="Matrícula del Productor:", font=("Arial", 12)).pack(pady=10)
entry = tk.Entry(root, font=("Arial", 12), justify="center")
entry.pack(pady=5)

btn = tk.Button(root, text="Buscar", command=ejecutar_busqueda, font=("Arial", 12))
btn.pack(pady=10)

resultado = tk.Text(root, height=15, font=("Consolas", 10), state="disabled")
resultado.pack(padx=10, pady=10, fill="both", expand=True)

root.mainloop()
