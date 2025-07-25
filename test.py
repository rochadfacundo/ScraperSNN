import requests
import json
import os

def descargar_datos_georef(destino="data/georef"):
    os.makedirs(destino, exist_ok=True)

    urls = {
        "localidades": "https://infra.datos.gob.ar/georef/localidades.json",
        "provincias": "https://infra.datos.gob.ar/georef/provincias.json"
    }

    for nombre, url in urls.items():
        print(f"Descargando {nombre}...")
        response = requests.get(url)
        if response.status_code == 200:
            path = os.path.join(destino, f"{nombre}.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(response.json(), f, ensure_ascii=False, indent=2)
            print(f"✔ Guardado en {path}")
        else:
            print(f"❌ Error al descargar {nombre}: {response.status_code}")

# Ejemplo de uso
descargar_datos_georef()
