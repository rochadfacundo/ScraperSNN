# 🕵️‍♂️ ScraperSNN

Scraper automático que consulta el padrón de productores de seguros habilitados en el sitio oficial de la Superintendencia de Seguros de la Nación (SSN) de Argentina, y exporta los datos a un archivo Excel con formato personalizado.

---

## Características

- Ingreso de número de matrícula y búsqueda automática en la web de la SSN
- Extracción de datos clave: nombre, documento, CUIT, domicilio, provincia, teléfonos, email
- Exportación a Excel.
---

## Tecnologías usadas

- Python 3.10+
- Selenium
- BeautifulSoup
- OpenPyXL
- Tkinter (GUI)
- Pillow

---

## Instalación

```bash
# Clonar el repositorio
git clone https://github.com/tuusuario/scrap-polizas.git
cd scrap-polizas

# Crear entorno virtual
python -m venv env
env\Scripts\activate  # En Windows

# Instalar dependencias
pip install -r requirements.txt
