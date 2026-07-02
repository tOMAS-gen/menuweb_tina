#!/usr/bin/env python3
"""Verifica que menu.xlsx tenga las 13 hojas esperadas, con encabezados
correctos y valores puntuales que coinciden con los precios vigentes."""

import os

import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HEADERS_ROW = ["nombre", "descripcion", "disponible", "precio", "precio_2", "imagen"]
HEADERS_CAFE = ["nombre", "disponible", "precio_s", "precio_m", "precio_l", "precio_xl"]

EXPECTED_SHEETS = [
    "Café Tradicional", "Café de Especialidad", "Materos", "Frozzen",
    "Té en Hebras", "Frapucchino", "Bebidas", "Pastelería Propia",
    "Antojos", "Focaccias Rellenas", "Pastelería Moderna", "Desayunos",
    "Promos",
]


def main():
    path = os.path.join(BASE_DIR, "menu.xlsx")
    wb = openpyxl.load_workbook(path, data_only=True)

    assert wb.sheetnames == EXPECTED_SHEETS, (
        f"Hojas inesperadas.\nEsperado: {EXPECTED_SHEETS}\nObtenido: {wb.sheetnames}"
    )

    cafe = wb["Café Tradicional"]
    header_cafe = [c.value for c in next(cafe.iter_rows(min_row=1, max_row=1))]
    assert header_cafe == HEADERS_CAFE, f"Encabezado incorrecto en 'Café Tradicional': {header_cafe}"
    fila2 = [c.value for c in next(cafe.iter_rows(min_row=2, max_row=2))]
    assert fila2[0] == "Expresso" and fila2[1] == "si", fila2
    assert fila2[2] == 3300 and fila2[3] == 4200 and fila2[4] == 4800 and fila2[5] == 5500, fila2

    for hoja in EXPECTED_SHEETS[1:]:
        ws = wb[hoja]
        header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        assert header_row == HEADERS_ROW, f"Encabezado incorrecto en '{hoja}': {header_row}"
        assert ws.max_row >= 2, f"La hoja '{hoja}' no tiene filas de datos"

    bebidas = wb["Bebidas"]
    filas = list(bebidas.iter_rows(min_row=2, values_only=True))
    jugo = next(f for f in filas if f[0] == "Jugo natural de naranja")
    assert jugo[3] == 6000 and jugo[4] == 7600, jugo

    moderna = wb["Pastelería Moderna"]
    fila2 = [c.value for c in next(moderna.iter_rows(min_row=2, max_row=2))]
    assert fila2[0] == "Hawaiana", fila2
    assert fila2[5] == "torta-hawaiana.png", fila2

    promos = wb["Promos"]
    nombres_promos = [row[0].value for row in promos.iter_rows(min_row=2)]
    assert len(nombres_promos) == 13, f"Se esperaban 13 promos, hay {len(nombres_promos)}"

    print("OK: menu.xlsx tiene las 13 hojas, encabezados y valores esperados.")


if __name__ == "__main__":
    main()
