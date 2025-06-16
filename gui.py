import os
import time
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

# ----------------- Lógica de scraping -----------------
def get_dato(soup, label):
    el = soup.find('span', string=lambda s: s and label in s)
    if el and el.parent:
        return el.parent.get_text(strip=True).replace(label + ":", "").strip()
    return "--"

def buscar_datos(matricula):
    try:
        # Configuración de Chrome (headless)
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')

        chromedriver_path = os.path.join(os.getcwd(), "chromedriver.exe")
        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service, options=options)

        # Paso 1: Ir al sitio
        driver.get("https://www.ssn.gob.ar/STORAGE/REGISTROS/PRODUCTORES/PRODUCTORESACTIVOSFILTRO.ASP")

        # Paso 2: Completar matrícula
        input_matricula = driver.find_element(By.ID, "matricula")
        input_matricula.send_keys(matricula)

        # Paso 3: Hacer clic en "Buscar"
        btn_buscar = driver.find_element(By.NAME, "Submit")
        btn_buscar.click()

        # Paso 4: Esperar y cambiar a nueva pestaña
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[1])

        html = driver.page_source
        driver.quit()

        soup = BeautifulSoup(html, "html.parser")

        datos = {
            "Nombre": get_dato(soup, "Nombre"),
            "Documento": get_dato(soup, "Documento"),
            "CUIT": get_dato(soup, "CUIT"),
            "Domicilio": get_dato(soup, "Domicilio"),
            "Provincia": get_dato(soup, "Provincia"),
            "Teléfonos": get_dato(soup, "Teléfonos"),
            "Email": get_dato(soup, "E-mail"),
        }

        return datos

    except Exception as e:
        return {"Error": f"Ocurrió un error: {str(e)}"}

# ----------------- Interfaz gráfica -----------------
def ejecutar_busqueda():
    matricula = entry.get()
    if not matricula.strip():
        messagebox.showwarning("Matrícula vacía", "Por favor ingresá un número de matrícula.")
        return

    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    resultado.insert(tk.END, "Buscando datos...\n")
    resultado.config(state="disabled")
    root.update()

    datos = buscar_datos(matricula)

    resultado.config(state="normal")
    resultado.delete("1.0", tk.END)
    for k, v in datos.items():
        resultado.insert(tk.END, f"{k}: {v}\n")
    resultado.config(state="disabled")

# Crear ventana
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
