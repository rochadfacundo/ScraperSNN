import os
import re
import sys
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def get_dato(soup, label):
    el = soup.find('span', string=lambda s: s and label in s)
    if el and el.parent:
        text = el.parent.get_text(separator=' ', strip=True)
        cleaned = re.sub(r'\s+', ' ', text).replace(f"{label}:", "").strip()
        if label in ["Documento", "CUIT"]:
            return re.sub(r"[^\d]", "", cleaned)
        return cleaned
    return "--"

def get_dato_fuera_de_span(soup, label):
    p_tags = soup.find_all('p', class_='margen_inf')
    for p in p_tags:
        span = p.find('span', class_='destacadopdtores')
        if span and label in span.get_text(strip=True):
            full_text = p.get_text(strip=True)
            return full_text.replace(span.get_text(strip=True), "").strip()
    return "--"

def buscar_datos(matricula, reintentos=4):
    try:
        for intento in range(reintentos + 1):
            print(f"🟡 Intento {intento + 1} para matrícula {matricula}")

            options = Options()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')

            if getattr(sys, 'frozen', False):
                chromedriver_path = os.path.join(sys._MEIPASS, "chromedriver.exe")
            else:
                chromedriver_path = os.path.join(os.getcwd(), "chromedriver.exe")

            service = Service(executable_path=chromedriver_path)
            driver = webdriver.Chrome(service=service, options=options)

            try:
                driver.get("https://www.ssn.gob.ar/STORAGE/REGISTROS/PRODUCTORES/PRODUCTORESACTIVOSFILTRO.ASP")
                driver.find_element(By.ID, "matricula").send_keys(matricula)
                driver.find_element(By.NAME, "Submit").click()
                print("📤 Formulario enviado")

                WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
                driver.switch_to.window(driver.window_handles[1])
                print("🧭 Segunda pestaña abierta")

                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Nombre')]"))
                )
                print("📄 Página cargada")

                html = driver.page_source
                soup = BeautifulSoup(html, "html.parser")
                driver.quit()

                return {
                    "Nombre": get_dato(soup, "Nombre"),
                    "Documento": get_dato(soup, "Documento"),
                    "CUIT": get_dato(soup, "CUIT"),
                    "Domicilio": get_dato(soup, "Domicilio"),
                    "Provincia": get_dato_fuera_de_span(soup, "Provincia"),
                    "Teléfonos": get_dato(soup, "Teléfonos"),
                    "Email": get_dato(soup, "E-mail"),
                    "Ramo": get_dato(soup, "Ramo"),
                    "Nro. de Resolución": get_dato_fuera_de_span(soup, "Nro. de Resolución"),
                    "Fº de Resolución": get_dato_fuera_de_span(soup, "Fº de Resolución")
                }

            except Exception as e:
                print(f"❌ Fallo intento {intento + 1} para matrícula {matricula}: {e}")
                driver.quit()
                continue

        return {"Error": f"No se encontraron datos para la matrícula {matricula} luego de {reintentos + 1} intentos."}

    except Exception as e:
        return {"Error": f"Ocurrió un error con matrícula {matricula}: {str(e)}"}
