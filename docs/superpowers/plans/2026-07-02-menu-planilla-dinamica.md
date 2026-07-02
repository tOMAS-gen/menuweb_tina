# Menú desde planilla Excel multi-hoja (v2) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** El menú de `menuweb_tina/index.html` deja de tener ítems hardcodeados y pasa a leer nombre, precio(s), descripción, disponibilidad e imagen desde `menu.xlsx`, un libro Excel con una hoja por sector. Agregar una fila agrega el ítem al sitio; marcar `disponible=no` lo muestra tachado como agotado, sin sacarlo del menú.

**Architecture:** `menu.xlsx` es la fuente de verdad. `precios.js` lo lee client-side con SheetJS (vendorizado, sin CDN), arma el HTML de cada sección a partir de las filas de su hoja y lo inyecta en contenedores vacíos (`data-sector="..."`) que reemplazan el markup hoy hardcodeado en `index.html`. Si la lectura del `.xlsx` falla, se usa `assets/menu-fallback.json`, generado por el mismo script que genera `menu.xlsx`.

**Tech Stack:** HTML/CSS/JS estático (Tailwind CDN + CSS propio). SheetJS (`xlsx.full.min.js`, vendorizado) para leer `.xlsx` en el navegador. `openpyxl` (Python 3, ya instalado localmente) para generar `menu.xlsx` desde un script de migración.

## Global Constraints

- No hay framework de testing en este repo (sitio estático). La verificación de JS/HTML es manual en navegador, documentada paso a paso en cada tarea. La verificación de los scripts Python sí es automatizable con `python3`.
- Repo git propio en `/Users/tomi/Desktop/en_trabajo/tina/menuweb_tina`, conectado a `origin` = `https://github.com/tOMAS-gen/menuweb_tina.git`. Cada tarea termina con su propio commit local. **No hacer `git push` salvo pedido explícito.** Mensajes de commit sin `Co-Authored-By` ni menciones a Claude/Anthropic.
- Hay un cambio sin commitear en `index.html` (ajuste de posición del `.tea-tail-svg`, líneas ~82 y ~153 y ~500) que no tiene relación con este feature — no tocarlo ni descartarlo; si una tarea edita líneas cercanas, preservar ese cambio tal como está en el working tree al momento de empezar.
- El menú sigue funcionando sin conexión a internet una vez vendorizado SheetJS.
- La sección `id="hebras"` (panel educativo de tés) no se toca.
- Esquema de columnas: hojas normales `nombre | descripcion | disponible | precio | precio_2 | imagen`; hoja "Café Tradicional" `nombre | disponible | precio_s | precio_m | precio_l | precio_xl`. `disponible` es `si`/`no`, default `si`. `imagen` sólo se usa en la hoja "Pastelería Moderna". `precio_2` sólo en ítems que hoy muestran 2 precios sin etiqueta (Jugo natural de naranja, Limonada, Medialuna).
- No se implementa columna de descuento (descartada explícitamente).

---

## Mapeo sector → hoja (referencia para todas las tareas)

| `id` de `<section>` | Nombre de hoja en `menu.xlsx` | Tipo de fila |
|---|---|---|
| `cafe` | Café Tradicional | `size` (S/M/L/XL) |
| `especialidad` | Café de Especialidad | `row` |
| `materos` | Materos | `row` |
| `frozen` | Frozzen | `row` |
| `tes` | Té en Hebras | `row` |
| `frapuccino` | Frapucchino | `row` |
| `bebidas` | Bebidas | `row` |
| `pasteleria` | Pastelería Propia | `row` |
| `antojos` | Antojos | `row` |
| `focaccias` | Focaccias Rellenas | `row` |
| `moderna` | Pastelería Moderna | `pastry` (con imagen) |
| `desayunos` | Desayunos | `row` |
| `promos` | Promos | `promo` |

---

### Task 1: Script generador de `menu.xlsx` y `assets/menu-fallback.json`

**Files:**
- Create: `scripts/build_menu_xlsx.py`
- Create: `scripts/verify_menu_xlsx.py`
- Produces (al ejecutar el script): `menu.xlsx`, `assets/menu-fallback.json`

**Interfaces:**
- Produces: `menu.xlsx` con 13 hojas (tabla de mapeo arriba). Hojas normales: columnas `nombre, descripcion, disponible, precio, precio_2, imagen`. Hoja "Café Tradicional": columnas `nombre, disponible, precio_s, precio_m, precio_l, precio_xl`. `assets/menu-fallback.json`: mismo contenido en JSON, `{ "<nombre de hoja>": [ {...fila...}, ... ], ... }`.
- Consumes: nada (datos hardcodeados en el script, extraídos de `index.html` y del `PRECIOS_FALLBACK` vigente en `precios.js`, que es lo que hoy alimenta el sitio realmente — no hay `precios.csv` en este repo).

- [ ] **Step 1: Crear la carpeta de scripts**

```bash
mkdir -p /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/scripts
```

- [ ] **Step 2: Escribir `scripts/build_menu_xlsx.py`**

```python
#!/usr/bin/env python3
"""Genera menu.xlsx y assets/menu-fallback.json desde los datos del menú.
Editar este archivo y volver a correrlo es la forma de regenerar ambos
artefactos desde una única fuente si hace falta reconstruirlos."""

import json
import os

import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

HEADERS_ROW = ["nombre", "descripcion", "disponible", "precio", "precio_2", "imagen"]
HEADERS_CAFE = ["nombre", "disponible", "precio_s", "precio_m", "precio_l", "precio_xl"]

CAFE_TRADICIONAL = [
    {"nombre": "Expresso", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Doble Expresso", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Café con leche", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Cortado", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Lágrima", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Expresso Panna", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Ice Coffee", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
    {"nombre": "Expresso Tonic", "precio_s": 3300, "precio_m": 4200, "precio_l": 4800, "precio_xl": 5500},
]

SECTORES = {
    "Café de Especialidad": [
        {"nombre": "Capuccino", "descripcion": "Clásico / Frambuesa", "precio": 6400},
        {"nombre": "Café Moccha Vienés", "precio": 6400},
        {"nombre": "Caramel Macchiato", "precio": 6400},
        {"nombre": "Macchiato Avellanas", "precio": 6400},
        {"nombre": "Café Irlandés", "precio": 6400},
        {"nombre": "Latte Vainilla", "precio": 6400},
        {"nombre": "Affogato", "precio": 6400},
        {"nombre": "Iced Coffee", "descripcion": "Pistacho / Coco", "precio": 6400},
        {"nombre": "Café Tina", "precio": 6400},
        {"nombre": "Café Bombón", "precio": 6400},
        {"nombre": "Café caribeño", "precio": 6400},
        {"nombre": "Submarino", "precio": 6400},
        {"nombre": "Bayleis", "precio": 6400},
    ],
    "Materos": [
        {"nombre": "Jaguar / Pampa / Paradise", "descripcion": "mate + blend yerba a elección + termo 80°", "precio": 4600},
        {"nombre": "Opción Tereré", "descripcion": "mate + blend yerba a elección + termo 80°", "precio": 4900},
    ],
    "Frozzen": [
        {"nombre": "Banana con dulce de leche", "precio": 6600},
        {"nombre": "Arándanos", "precio": 6600},
        {"nombre": "Frutos rojos", "precio": 6600},
        {"nombre": "Frutillas con crema", "precio": 6600},
        {"nombre": "Mousse de limón", "precio": 6600},
    ],
    "Té en Hebras": [
        {"nombre": "Tetera para compartir", "precio": 6000},
        {"nombre": "Tetera individual", "precio": 3300},
        {"nombre": "Té / Yerbeado", "precio": 3300},
        {"nombre": "Matcha latte", "descripcion": "té verde en polvo con leche", "precio": 6400},
    ],
    "Frapucchino": [
        {"nombre": "Café bombón", "precio": 6600},
        {"nombre": "Mocha Vainilla Latte", "precio": 6600},
        {"nombre": "Caramel Avellanas", "precio": 6600},
        {"nombre": "Frambuesa", "precio": 6600},
    ],
    "Bebidas": [
        {"nombre": "Agua / Agua Saborizada", "precio": 4100},
        {"nombre": "Gaseosas", "descripcion": "Línea Pepsi", "precio": 4100},
        {"nombre": "Jugo natural de naranja", "precio": 6000, "precio_2": 7600},
        {"nombre": "Pomelada", "precio": 8700},
        {"nombre": "Limonada", "descripcion": "Clásica / Menta jengibre / Arándanos / Maracuyá / Cereza", "precio": 8700, "precio_2": 9600},
    ],
    "Pastelería Propia": [
        {"nombre": "Brownie con frutos rojos", "precio": 7100},
        {"nombre": "Crumble", "precio": 7100},
        {"nombre": "Cheesecake", "precio": 7100},
        {"nombre": "Balcarce", "precio": 7100},
        {"nombre": "Selva Negra", "precio": 7100},
        {"nombre": "Lemon Pie", "precio": 7100},
        {"nombre": "Limón Cocado", "precio": 7100},
        {"nombre": "Tiramisú", "precio": 7100},
    ],
    "Antojos": [
        {"nombre": "Alfajores Premium", "descripcion": "Pistacho / Red Velvet / Chocotorta", "precio": 6600},
        {"nombre": "Alfajor Maicena / Miel", "precio": 4200},
        {"nombre": "Financieros", "descripcion": "4 unidades", "precio": 2800},
        {"nombre": "Madeleine", "descripcion": "3 unidades", "precio": 2800},
        {"nombre": "Cookies", "precio": 2800},
        {"nombre": "Macarons", "precio": 7000},
        {"nombre": "Budín", "precio": 3500},
        {"nombre": "Scones", "descripcion": "3 unidades", "precio": 3600},
        {"nombre": "Tortita", "precio": 850},
        {"nombre": "Tostado jamón y queso", "precio": 9200},
        {"nombre": "Medialuna", "descripcion": "opc. jamón y queso", "precio": 1400, "precio_2": 3300},
        {"nombre": "Tostón con omelette", "precio": 9200},
        {"nombre": "Bruschetta Tina", "precio": 9200},
        {"nombre": "Galletas Bretón", "descripcion": "3 unidades", "precio": 2100},
    ],
    "Focaccias Rellenas": [
        {"nombre": "Serrano", "descripcion": "rúcula, jamón crudo y nueces", "precio": 10800},
        {"nombre": "Azul", "descripcion": "queso untable, queso roquefort y tomate hidratado", "precio": 10800},
        {"nombre": "Gourmet", "descripcion": "peras, queso azul, rúcula, nueces y miel", "precio": 10800},
        {"nombre": "Capresse", "descripcion": "tomate, queso y albahaca", "precio": 10800},
        {"nombre": "Tradicional", "descripcion": "jamón, queso, lechuga y tomate", "precio": 10800},
        {"nombre": "Fugazza", "descripcion": "queso y cebolla caramelizada", "precio": 10800},
    ],
    "Pastelería Moderna": [
        {"nombre": "Hawaiana", "descripcion": "mousse de coco y corazón de ananá", "precio": 7500, "imagen": "torta-hawaiana.png"},
        {"nombre": "Agus Tina", "descripcion": "mousse de chocolate, centro de naranja, terciopelo de chocolate", "precio": 7500, "imagen": "torta-agus-tina.png"},
        {"nombre": "Valen Tina", "descripcion": "profiterol relleno de cheesecake de pistacho, centro de praline de pistacho", "precio": 7500, "imagen": "torta-valen-tina.png"},
        {"nombre": "Capricho", "descripcion": "cremoso de chocolate, centro de maracuyá y frambuesa, terciopelo de chocolate blanco", "precio": 7500, "imagen": "torta-capricho.png"},
        {"nombre": "Brownie boom", "descripcion": "húmedo de chocolate, pasta de maní y dulce de leche, bañado en chocolate y crocante de maní", "precio": 7500, "imagen": "torta-brownie-boom.png"},
        {"nombre": "Banofee", "descripcion": "base de chocolate, ganache de maní, bananas caramelizadas, salsa toffee y mousse de chocolate", "precio": 7500, "imagen": "torta-banofee.png"},
    ],
    "Desayunos": [
        {"nombre": "Americano", "descripcion": "Cafetería + jugo exprimido + pan de campo con palta y huevo", "precio": 14400},
        {"nombre": "Tentación", "descripcion": "Café con leche + jugo exprimido + opción de pastelería", "precio": 13800},
        {"nombre": "Criollo", "descripcion": "Cafetería + jugo exprimido + tostadas con queso, manteca o dulce", "precio": 10800},
        {"nombre": "Alfajor", "descripcion": "Cafetería + jugo exprimido + alfajor", "precio": 10700},
    ],
    "Promos": [
        {"nombre": "Clásico", "descripcion": "Infusión cafetería + 2 medialunas o tostadas + jugo", "precio": 7800},
        {"nombre": "Minis", "descripcion": "Chocolatada / Jugo mediano + cookies / sandwich", "precio": 9000},
        {"nombre": "Criollo", "descripcion": "Infusión cafetería + tostada con manteca o queso y dulce + jugo", "precio": 8300},
        {"nombre": "Alfajor", "descripcion": "Infusión cafetería + alfajor a elección + jugo", "precio": 8000},
        {"nombre": "Tentación", "descripcion": "Café especial + porción de torta + jugo", "precio": 14500},
        {"nombre": "Saludable", "descripcion": "Infusión cafetería + yogurt con granola + frutillas y arándanos + jugo", "precio": 13800},
        {"nombre": "Americano", "descripcion": "Infusión cafetería + tostada con huevo y palta + jugo", "precio": 12800},
        {"nombre": "Stella", "descripcion": "Cerveza + focaccia a elección", "precio": 16000},
        {"nombre": "Carlitos", "descripcion": "Infusión cafetería + tostada con jamón y queso + jugo", "precio": 16600},
        {"nombre": "Gran Americano", "descripcion": "Infusión cafetería + jamón, queso y huevo + mix de frutas + tostadas + jugo", "precio": 17400},
        {"nombre": "Trío Macarons", "descripcion": "Té en hebras / Infusión cafetería + jugo + trío de macarons", "precio": 12800},
        {"nombre": "Tina", "descripcion": "Tetera té en hebras + 2 porciones budín + 2 scones + madeleine + financieros + 2 jugos (para compartir)", "precio": 17400},
        {"nombre": "Degustación", "descripcion": "jarra de limonada a elección + degustación de focaccia (para compartir)", "precio": 37200},
    ],
}


def build_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(title="Café Tradicional")
    ws.append(HEADERS_CAFE)
    for item in CAFE_TRADICIONAL:
        ws.append([item.get(col, "si" if col == "disponible" else None) for col in HEADERS_CAFE])

    for hoja, items in SECTORES.items():
        ws = wb.create_sheet(title=hoja)
        ws.append(HEADERS_ROW)
        for item in items:
            ws.append([item.get(col, "si" if col == "disponible" else None) for col in HEADERS_ROW])

    return wb


def build_fallback():
    fallback = {"Café Tradicional": [dict(item, disponible=item.get("disponible", "si")) for item in CAFE_TRADICIONAL]}
    for hoja, items in SECTORES.items():
        fallback[hoja] = [dict(item, disponible=item.get("disponible", "si")) for item in items]
    return fallback


def main():
    wb = build_workbook()
    xlsx_path = os.path.join(BASE_DIR, "menu.xlsx")
    wb.save(xlsx_path)
    print(f"Escrito {xlsx_path}")

    fallback_path = os.path.join(BASE_DIR, "assets", "menu-fallback.json")
    with open(fallback_path, "w", encoding="utf-8") as f:
        json.dump(build_fallback(), f, ensure_ascii=False, indent=2)
    print(f"Escrito {fallback_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Ejecutar el script**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
python3 scripts/build_menu_xlsx.py
```

Expected:
```
Escrito /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/menu.xlsx
Escrito /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/assets/menu-fallback.json
```

- [ ] **Step 4: Escribir `scripts/verify_menu_xlsx.py`**

```python
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
```

- [ ] **Step 5: Correr la verificación**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
python3 scripts/verify_menu_xlsx.py
```

Expected:
```
OK: menu.xlsx tiene las 13 hojas, encabezados y valores esperados.
```

- [ ] **Step 6: Verificar el fallback JSON a simple vista**

```bash
python3 -c "import json; d = json.load(open('/Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/assets/menu-fallback.json')); print(list(d.keys())); print(len(d['Promos'])); print(d['Pastelería Moderna'][0]); print(d['Bebidas'][2])"
```

Expected: lista de 13 nombres de hoja, `13`, un dict con `nombre: 'Hawaiana'` e `imagen: 'torta-hawaiana.png'`, y un dict `Jugo natural de naranja` con `precio: 6000, precio_2: 7600`.

- [ ] **Step 7: Commit**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
git add scripts/build_menu_xlsx.py scripts/verify_menu_xlsx.py menu.xlsx assets/menu-fallback.json
git commit -m "Generar menu.xlsx y respaldo JSON desde script de migración"
```

---

### Task 2: Vendorizar SheetJS y reescribir `precios.js`

**Files:**
- Create: `assets/xlsx.full.min.js`
- Modify: `precios.js` (reemplazo completo)

**Interfaces:**
- Consumes: `menu.xlsx` y `assets/menu-fallback.json` (Task 1). `assets/tortas/<archivo>` para las imágenes de Pastelería Moderna (ya existen con nombres descriptivos, no hace falta crearlas).
- Produces: al cargar la página, cada `[data-sector="<id>"]` (Task 3 los crea en `index.html`) queda con el HTML de los ítems de su hoja. Ítems con `disponible=no` llevan la clase `agotado` en su bloque contenedor (Task 3 agrega el CSS que la interpreta). Expone `window.PRECIOS_CARGADOS = true|false`.

- [ ] **Step 1: Descargar SheetJS vendorizado**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
curl -sL -o assets/xlsx.full.min.js https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
```

- [ ] **Step 2: Verificar la descarga**

```bash
head -c 200 /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/assets/xlsx.full.min.js
wc -c /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/assets/xlsx.full.min.js
```

Expected: arranca con un comentario tipo `/*! xlsx.js` y pesa varios cientos de KB.

- [ ] **Step 3: Reemplazar `precios.js` completo**

```javascript
// ═══════════════════════════════════════════════════════
//  MENÚ — Cafetería Tina
//  Editá menu.xlsx (una hoja por sector) y recargá la página.
// ═══════════════════════════════════════════════════════

const SECTORES = [
  { id: 'cafe', hoja: 'Café Tradicional', kind: 'size' },
  { id: 'especialidad', hoja: 'Café de Especialidad', kind: 'row' },
  { id: 'materos', hoja: 'Materos', kind: 'row' },
  { id: 'frozen', hoja: 'Frozzen', kind: 'row' },
  { id: 'tes', hoja: 'Té en Hebras', kind: 'row' },
  { id: 'frapuccino', hoja: 'Frapucchino', kind: 'row' },
  { id: 'bebidas', hoja: 'Bebidas', kind: 'row' },
  { id: 'pasteleria', hoja: 'Pastelería Propia', kind: 'row' },
  { id: 'antojos', hoja: 'Antojos', kind: 'row' },
  { id: 'focaccias', hoja: 'Focaccias Rellenas', kind: 'row' },
  { id: 'moderna', hoja: 'Pastelería Moderna', kind: 'pastry' },
  { id: 'desayunos', hoja: 'Desayunos', kind: 'row' },
  { id: 'promos', hoja: 'Promos', kind: 'promo' },
];

function formatPrecio(n) {
  return '$' + Number(n).toLocaleString('es-AR');
}

function esc(str) {
  const div = document.createElement('div');
  div.textContent = str == null ? '' : String(str);
  return div.innerHTML;
}

function esDisponible(raw) {
  const v = (raw == null ? 'si' : String(raw)).trim().toLowerCase();
  return v !== 'no';
}

function normalizarItem(raw) {
  const num = v => (v === undefined || v === null || v === '' ? null : Number(v));
  const str = v => (v === undefined || v === null || v === '' ? null : String(v).trim());
  return {
    nombre: str(raw.nombre),
    descripcion: str(raw.descripcion),
    disponible: esDisponible(raw.disponible),
    precio: num(raw.precio),
    precio_2: num(raw.precio_2),
    precio_s: num(raw.precio_s),
    precio_m: num(raw.precio_m),
    precio_l: num(raw.precio_l),
    precio_xl: num(raw.precio_xl),
    imagen: str(raw.imagen),
  };
}

function agotadoLabel() {
  return '<span class="agotado-label">agotado</span>';
}

function renderSizeRows(items) {
  return items.map(it => `
    <div class="size-row${it.disponible ? '' : ' agotado'}">
      <strong>${esc(it.nombre)}</strong>
      <span>${it.precio_s != null ? formatPrecio(it.precio_s) : ''}</span>
      <span>${it.precio_m != null ? formatPrecio(it.precio_m) : ''}</span>
      <span>${it.precio_l != null ? formatPrecio(it.precio_l) : ''}</span>
      <span>${it.precio_xl != null ? formatPrecio(it.precio_xl) : ''}</span>
      ${it.disponible ? '' : agotadoLabel()}
    </div>`).join('');
}

function renderMenuRows(items) {
  return items.map(it => `
    <div class="menu-row${it.disponible ? '' : ' agotado'}">
      <div class="min-w-0"><strong>${esc(it.nombre)}</strong>${it.descripcion ? `<small class="menu-desc">${esc(it.descripcion)}</small>` : ''}</div>
      <div class="prices">
        <span class="menu-price">${it.precio != null ? formatPrecio(it.precio) : ''}</span>
        ${it.precio_2 != null ? `<span class="menu-price">${formatPrecio(it.precio_2)}</span>` : ''}
      </div>
      ${it.disponible ? '' : agotadoLabel()}
    </div>`).join('');
}

function renderPastryRows(items) {
  return items.map(it => `
    <div class="pastry-card${it.disponible ? '' : ' agotado'}">
      <div class="pastry-bg"></div>
      <div class="pastry-circle"><img src="assets/tortas/${esc(it.imagen)}" alt="${esc(it.nombre)}" loading="lazy"></div>
      <div class="pastry-text"><strong class="pastry-item-name">${esc(it.nombre)}</strong><p class="pastry-desc">${esc(it.descripcion || '')}</p></div>
      <div class="pastry-badge">${it.precio != null ? formatPrecio(it.precio) : ''}</div>
      ${it.disponible ? '' : agotadoLabel()}
    </div>`).join('');
}

function renderPromoRows(items) {
  return items.map(it => {
    const compartir = /\(para compartir\)\s*$/i.test(it.descripcion || '');
    const desc = compartir ? it.descripcion.replace(/\s*\(para compartir\)\s*$/i, '') : (it.descripcion || '');
    const card = `
      <article class="promo-card${it.disponible ? '' : ' agotado'}">
        <div><h3>${esc(it.nombre)}</h3><p>${esc(desc)}</p></div>
        <strong>${it.precio != null ? formatPrecio(it.precio) : ''}</strong>
        ${it.disponible ? '' : agotadoLabel()}
      </article>`;
    return compartir
      ? `<div class="promo-share"><div class="promo-tag">para compartir</div>${card}</div>`
      : card;
  }).join('');
}

const RENDERERS = { size: renderSizeRows, row: renderMenuRows, pastry: renderPastryRows, promo: renderPromoRows };

function renderSector(id, kind, itemsRaw) {
  const container = document.querySelector(`[data-sector="${id}"]`);
  if (!container) return;
  const items = itemsRaw.map(normalizarItem).filter(it => it.nombre);
  container.innerHTML = RENDERERS[kind](items);
}

function renderDesdeWorkbook(workbook) {
  SECTORES.forEach(({ id, hoja, kind }) => {
    const sheet = workbook.Sheets[hoja];
    if (!sheet) {
      console.error(`menu.xlsx no tiene la hoja "${hoja}"`);
      return;
    }
    renderSector(id, kind, XLSX.utils.sheet_to_json(sheet, { defval: null }));
  });
}

function renderDesdeFallback(fallback) {
  SECTORES.forEach(({ id, hoja, kind }) => {
    const items = fallback[hoja];
    if (!items) {
      console.error(`El respaldo no tiene la hoja "${hoja}"`);
      return;
    }
    renderSector(id, kind, items);
  });
}

async function initMenu() {
  try {
    const res = await fetch('menu.xlsx');
    if (!res.ok) throw new Error('fetch de menu.xlsx falló: ' + res.status);
    const buf = await res.arrayBuffer();
    const workbook = XLSX.read(buf, { type: 'array' });
    renderDesdeWorkbook(workbook);
    window.PRECIOS_CARGADOS = true;
  } catch (err) {
    console.error('No se pudo leer menu.xlsx, uso el respaldo assets/menu-fallback.json.', err);
    try {
      const res = await fetch('assets/menu-fallback.json');
      const fallback = await res.json();
      renderDesdeFallback(fallback);
      window.PRECIOS_CARGADOS = true;
    } catch (err2) {
      console.error('Tampoco se pudo leer el respaldo. El menú queda vacío.', err2);
      window.PRECIOS_CARGADOS = false;
    }
  }
}

document.addEventListener('DOMContentLoaded', initMenu);
```

- [ ] **Step 4: Verificación manual de sintaxis JS**

```bash
node --check /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/precios.js
```

Expected: sin salida (exit code 0). Si no hay `node` instalado, este chequeo se saltea y se confirma en Task 4 (navegador).

- [ ] **Step 5: Commit**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
git add assets/xlsx.full.min.js precios.js
git commit -m "Leer menu.xlsx con SheetJS y renderizar los sectores dinámicamente"
```

---

### Task 3: Reestructurar `index.html`

**Files:**
- Modify: `index.html:324-356` (sección `cafe`)
- Modify: `index.html:359-419` (sección `especialidad`)
- Modify: `index.html:422-438` (sección `materos`)
- Modify: `index.html:441-469` (sección `frozen`)
- Modify: `index.html:472-496` (sección `tes`)
- Modify: `index.html:550-574` (sección `frapuccino`)
- Modify: `index.html:577-611` (sección `bebidas`)
- Modify: `index.html:614-630` (sección `pasteleria`)
- Modify: `index.html:633-700` (sección `antojos`)
- Modify: `index.html:703-735` (sección `focaccias`)
- Modify: `index.html:738-751` (sección `moderna`)
- Modify: `index.html:754-775` (sección `desayunos`)
- Modify: `index.html:778-845` (sección `promos`)
- Modify: `index.html:80-81` (agregar CSS de `.agotado` / `.agotado-label`)
- Modify: `index.html:918` (agregar `<script>` de SheetJS antes de `precios.js`)

**Interfaces:**
- Consumes: `precios.js` (Task 2) — busca `[data-sector="<id>"]` y le setea `innerHTML`; agrega clase `agotado` a los bloques de ítems no disponibles.
- Produces: 13 contenedores `data-sector` vacíos, uno por sector de la tabla de mapeo. `grep -c 'data-sector=' index.html` debe dar `13`. `grep -c 'data-precio=' index.html` debe dar `0`.

- [ ] **Step 1: Agregar el CSS de `.agotado`**

Reemplazar (línea 80-81):
```css
    .prices{display:flex;gap:22px;justify-content:flex-end;white-space:nowrap}
    .menu-price{font-size:.75em;font-style:italic}
```
por:
```css
    .prices{display:flex;gap:22px;justify-content:flex-end;white-space:nowrap}
    .menu-price{font-size:.75em;font-style:italic}
    .agotado .menu-price,.agotado .pastry-badge,.agotado>strong,.size-row.agotado span{text-decoration:line-through;opacity:.5}
    .agotado-label{font-size:.65em;color:#a33;text-transform:uppercase;letter-spacing:.05em;margin-left:8px;font-style:normal}
```

- [ ] **Step 2: Sección `cafe` — dejar sólo el encabezado S/M/L/XL y un contenedor vacío**

El bloque actual (`index.html:324-356`) tiene, después de `.size-head`, 8 divs `.size-row` idénticos en estructura (Expresso, Doble Expresso, Café con leche, Cortado, Lágrima, Expresso Panna, Ice Coffee, Expresso Tonic — ya transcriptos en `CAFE_TRADICIONAL` del Task 1). Reemplazar todo el bloque completo por:
```html
      <section id="cafe" class="menu-section layout-left">
        <img class="section-icon" src="assets/iconos/icono-cafe-tradicional.svg" alt="" loading="lazy">
        <div class="section-copy">
          <h2 class="section-title">Caf&#233; tradicional</h2>
          <div class="menu-table">
            <div class="size-head"><span></span><span>S</span><span>M</span><span>L</span><span>XL</span></div>
            <div data-sector="cafe"></div>
          </div>
        </div>
      </section>
```

- [ ] **Step 3: Secciones `menu-row` genéricas — dejar sólo `.section-bar` y un contenedor vacío**

Aplicar el mismo patrón (mismo `<div class="menu-table">`, cambia sólo el `id` del `data-sector`) a `especialidad`, `frozen`, `tes`, `frapuccino`, `bebidas`, `pasteleria`, `antojos`, `focaccias`. Ejemplo con `especialidad` (líneas 359-419) — reemplazar todo el contenido de `<div class="menu-table">` por:
```html
          <div class="menu-table">
            <div class="section-bar"></div>
            <div data-sector="especialidad"></div>
          </div>
```

Repetir para:
- `frozen` → `<div data-sector="frozen"></div>`
- `tes` → `<div data-sector="tes"></div>`
- `frapuccino` → `<div data-sector="frapuccino"></div>`
- `bebidas` → `<div data-sector="bebidas"></div>`
- `pasteleria` → `<div data-sector="pasteleria"></div>`
- `antojos` → `<div data-sector="antojos"></div>`
- `focaccias` → `<div data-sector="focaccias"></div>`

- [ ] **Step 4: Sección `materos` — quitar la nota estática, ahora es descripción por ítem**

Reemplazar (líneas 422-438):
```html
      <section id="materos" class="menu-section layout-left">
        <img class="section-icon" src="assets/iconos/icono-materos.svg" alt="" loading="lazy">
        <div class="section-copy">
          <h2 class="section-title">Materos</h2>
          <div class="menu-table">
            <div class="menu-note-row">mate + blend yerba a elecci&#243;n + termo 80&#176;</div>
            <div class="menu-row">
              <div class="min-w-0"><strong>Jaguar / Pampa / Paradise</strong></div>
              <div class="prices"><span class="menu-price" data-precio="materos_clasico"></span></div>
            </div>
            <div class="menu-row">
              <div class="min-w-0"><strong>Opci&#243;n Terer&#233;</strong></div>
              <div class="prices"><span class="menu-price" data-precio="materos_terere"></span></div>
            </div>
          </div>
        </div>
      </section>
```
por:
```html
      <section id="materos" class="menu-section layout-left">
        <img class="section-icon" src="assets/iconos/icono-materos.svg" alt="" loading="lazy">
        <div class="section-copy">
          <h2 class="section-title">Materos</h2>
          <div class="menu-table">
            <div data-sector="materos"></div>
          </div>
        </div>
      </section>
```

- [ ] **Step 5: Sección `moderna` — contenedor vacío en `.pastry-grid`**

Reemplazar (líneas 738-751) por:
```html
      <section id="moderna" class="menu-section layout-left modern-section">
        <img class="section-icon" src="assets/iconos/icono-pasteleria-moderna.svg" alt="" loading="lazy">
        <div class="section-copy">
          <h2 class="section-title">Paster&#237;a moderna</h2>
          <div class="pastry-grid" data-sector="moderna"></div>
        </div>
      </section>
```

- [ ] **Step 6: Sección `desayunos` — mantener el ícono Sin TACC, vaciar la tabla**

Reemplazar (líneas 754-775) por:
```html
      <section id="desayunos" class="breakfast-block">
        <img src="assets/iconos/icono-desayunos.svg" alt="" loading="lazy">
        <div class="breakfast-table">
          <img class="sin-tacc" src="assets/iconos/icono-sin-tacc.svg" alt="Sin TACC">
          <div data-sector="desayunos"></div>
        </div>
      </section>
```

- [ ] **Step 7: Sección `promos` — contenedor vacío + precio del combo "tazón" fijo en HTML**

El bloque actual (`index.html:778-845`) tiene, después del `<h2>promos</h2>`, un `<div class="promo-grid">` con 13 `.promo-card` (Clásico, Minis, Criollo, Alfajor, Tentación, Saludable, Americano, Stella, Carlitos, Gran Americano, Trío Macarons, Tina, Degustación — ya transcriptos 1:1 en `SECTORES["Promos"]` del Task 1; las dos últimas envueltas en `.promo-share` con la etiqueta "para compartir", que ahora arma `renderPromoRows` en Task 2 detectando el sufijo `(para compartir)` en la descripción). Reemplazar el `<section>` completo por:
```html
      <section id="promos" class="promos-panel">
        <svg class="promo-bubble-svg" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 248.82 95.52">
          <path d="M243.23,0H5.59C2.5,0,0,2.5,0,5.59v46.97h0v42.97l29.4-23.69h213.83c3.09,0,5.59-2.5,5.59-5.59V5.59c0-3.09-2.5-5.59-5.59-5.59Z" style="fill:#614e4c"/>
          <text transform="translate(16.62 30)" style="fill:#fffbf4;font-family:AsapCondensed-Medium,'Asap Condensed';font-size:23.82px;font-weight:500;">agrandá tus promos</text>
          <text transform="translate(16.62 58)" style="fill:#fffbf4;font-family:AsapCondensed-Medium,'Asap Condensed';font-size:23.82px;font-weight:500;">a tazón por <tspan>$1.500</tspan></text>
        </svg>
        <h2>promos</h2>
        <div class="promo-grid" data-sector="promos"></div>
      </section>
```

(El precio del "tazón extra" no es un ítem de catálogo — nota fija de upsell, queda hardcodeada. Si cambia, se edita ese número a mano. Sólo se le quita el atributo `data-precio`, que ya no lo llena nadie — el texto `$1.500` queda escrito directo en el SVG.)

- [ ] **Step 8: Agregar el script de SheetJS antes de `precios.js`**

Reemplazar (línea 918):
```html
  <script src="precios.js"></script>
```
por:
```html
  <script src="assets/xlsx.full.min.js"></script>
  <script src="precios.js"></script>
```

- [ ] **Step 9: Verificar que no quedan `data-precio` sueltos ni ítems hardcodeados**

```bash
grep -n "data-precio" /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/index.html
```

Expected: sin resultados.

```bash
grep -c 'data-sector=' /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina/index.html
```

Expected: `13`.

- [ ] **Step 10: Commit**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
git add index.html
git commit -m "Reemplazar ítems hardcodeados por contenedores data-sector"
```

---

### Task 4: Verificación manual end-to-end en el navegador

**Files:**
- No se crean ni modifican archivos. Esta tarea sólo verifica el resultado de las Tasks 1-3.

**Interfaces:**
- Consumes: sitio completo (`index.html`, `precios.js`, `menu.xlsx`, `assets/xlsx.full.min.js`, `assets/menu-fallback.json`, `assets/tortas/*`).

- [ ] **Step 1: Levantar el servidor local**

```bash
cd /Users/tomi/Desktop/en_trabajo/tina/menuweb_tina
python3 -m http.server 8000 &
sleep 1
open http://localhost:8000
```

- [ ] **Step 2: Chequear la consola del navegador**

Abrir las herramientas de desarrollador y confirmar:
- Sin errores en rojo.
- `window.PRECIOS_CARGADOS` es `true`.

- [ ] **Step 3: Revisar visualmente las 13 secciones**

Recorrer la página y confirmar, por sector:
- Café Tradicional: tabla S/M/L/XL con las 8 filas, precios $3.300/$4.200/$4.800/$5.500.
- Bebidas: "Jugo natural de naranja" y "Limonada" muestran 2 precios lado a lado (sin etiqueta), igual que antes.
- Antojos: "Medialuna" muestra 2 precios lado a lado.
- Pastelería Moderna: 6 fotos (`torta-*.png`) dentro de los círculos, con el badge de precio.
- Promos: "Tina" y "Degustación" con la etiqueta "para compartir"; el combo "agrandá tus promos a tazón" sigue mostrando "$1.500".

- [ ] **Step 4: Probar "agregar un ítem sin tocar código"**

1. Abrir `menu.xlsx`.
2. En la hoja "Café Tradicional", agregar una fila: `nombre="Cortado doble"`, `disponible=si`, `precio_s=3300`, `precio_m=4200`, `precio_l=4800`, `precio_xl=5500`.
3. Guardar, recargar `http://localhost:8000`.
4. Confirmar que "Cortado doble" aparece con sus 4 precios.
5. Deshacer el cambio (borrar la fila) y guardar.

- [ ] **Step 5: Probar "cambiar el nombre de un ítem"**

1. En la hoja "Antojos", cambiar `nombre` de "Cookies" a "Cookies de avena".
2. Guardar, recargar.
3. Confirmar que el sitio ahora dice "Cookies de avena".
4. Deshacer el cambio y guardar.

- [ ] **Step 6: Probar "marcar no disponible"**

1. En la hoja "Focaccias Rellenas", poner `disponible=no` en la fila "Serrano".
2. Guardar, recargar.
3. Confirmar que "Serrano" sigue apareciendo en la sección, con el precio tachado y la etiqueta "agotado", sin desaparecer.
4. Volver `disponible` a `si` y guardar.

- [ ] **Step 7: Probar el fallback**

1. `mv menu.xlsx menu.xlsx.bak`.
2. Recargar la página.
3. Confirmar en consola el mensaje `No se pudo leer menu.xlsx, uso el respaldo...` y que el menú se ve igual (viene de `assets/menu-fallback.json`).
4. `mv menu.xlsx.bak menu.xlsx`.

- [ ] **Step 8: Apagar el servidor**

```bash
kill %1
```

Sin commit — esta tarea no modifica archivos versionados.
