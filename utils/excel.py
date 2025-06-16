import os
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# Estilos globales
header_font = Font(name="Cascadia Code", bold=True, size=16)
data_font = Font(name="Cascadia Code", size=12)
center_align = Alignment(horizontal="center", vertical="center")
header_fill = PatternFill(start_color="94c484", end_color="94c484", fill_type="solid")
borde_suave = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)
borde_grueso = Border(
    left=Side(style='medium', color='000000'),
    right=Side(style='medium', color='000000'),
    top=Side(style='medium', color='000000'),
    bottom=Side(style='medium', color='000000')
)

def exportar_a_excel(datos: dict):
    # 🗂 Ruta segura a "Documentos/Datos Matriculados"
    carpeta_destino = Path.home() / "Documents" / "Datos de matriculas"
    carpeta_destino.mkdir(parents=True, exist_ok=True)

    fecha_str = datetime.now().strftime("%d-%m-%Y")
    archivo = carpeta_destino / f"datos {fecha_str}.xlsx"

    # Crear o cargar archivo
    if archivo.exists():
        wb = load_workbook(archivo)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultado"

        # Cabeceras
        for col_index, key in enumerate(datos.keys(), 1):
            cell = ws.cell(row=1, column=col_index, value=key)
            cell.font = header_font
            cell.alignment = center_align
            cell.fill = header_fill
            cell.border = borde_grueso
        ws.row_dimensions[1].height = 30

    # Fila disponible
    next_row = ws.max_row + 1

    # Carga de datos
    for col_index, value in enumerate(datos.values(), 1):
        cell = ws.cell(row=next_row, column=col_index, value=value)
        cell.font = data_font
        cell.alignment = center_align
        cell.border = borde_suave

    ws.row_dimensions[next_row].height = 30

    # Autoajuste de columnas (solo si es primera vez)
    if next_row == 2:
        for col_index, (key, value) in enumerate(zip(datos.keys(), datos.values()), 1):
            max_length = max(len(str(key)), len(str(value)))
            width = max_length * 1.4 + 4
            ws.column_dimensions[get_column_letter(col_index)].width = width

    wb.save(archivo)
