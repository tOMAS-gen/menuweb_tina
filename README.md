# Menú web Tina

Menú digital de **Tina — café, té y pastelería**. Página estática (HTML +
Tailwind CDN) que lee sus ítems (nombre, precio, descripción, disponibilidad,
imagen) desde una planilla de Google Sheets, una hoja por sector.

## Estructura

```
menuweb_tina/
├── index.html                  Página del menú (única vista)
├── precios.js                  Lee la planilla y arma el HTML de cada sector
├── menu.xlsx                   Copia local de referencia (no la lee el sitio)
├── scripts/
│   ├── build_menu_xlsx.py      Genera menu.xlsx + assets/menu-fallback.json
│   └── verify_menu_xlsx.py     Verifica la estructura de menu.xlsx
└── assets/
    ├── xlsx.full.min.js        SheetJS vendorizado (lee .xlsx en el navegador)
    ├── menu-fallback.json      Respaldo si falla la lectura de Google Sheets
    ├── logo/                   Logo del menú y fondo del hero
    ├── iconos/                 Íconos de cada sección + sin TACC + reloj de té
    └── tortas/                 Fotos de la pastelería moderna
```

## Ver localmente

Servir la carpeta (necesita HTTP, `file://` no alcanza porque el navegador
tiene que poder pedir `assets/xlsx.full.min.js` y hacer `fetch` a Google
Sheets):

```bash
python3 -m http.server 8000
# luego abrir http://localhost:8000
```

## Actualizar el menú

Editar la planilla de Google Sheets (una hoja por sector, ID configurado en
`GOOGLE_SHEET_ID` al principio de `precios.js`) y recargar la página — no
hace falta tocar código. La planilla tiene que seguir compartida como
"cualquiera con el enlace puede ver": el sitio la lee sin iniciar sesión.

Columnas por hoja: `nombre | descripcion | disponible | compartir | precio | precio_2 | imagen`
(hoja "Café Tradicional": `nombre | disponible | precio_s | precio_m | precio_l | precio_xl`).
`disponible = no` tacha el precio y muestra "agotado" sin sacar el ítem del
menú. `compartir = si` (sólo tiene sentido en la hoja "Promos") le agrega la
etiqueta "para compartir". La hoja "Promos" tiene además, en columnas H/I
separadas por una columna en blanco, la constante del cartel "agrandá tus
promos a tazón".

Si falla la lectura de Google Sheets (sin internet, planilla borrada, etc.),
el sitio cae a `assets/menu-fallback.json`. Para regenerar ese respaldo (y
`menu.xlsx`, la copia de referencia) tras cambiar los datos base en
`scripts/build_menu_xlsx.py`:

```bash
python3 scripts/build_menu_xlsx.py
python3 scripts/verify_menu_xlsx.py
```

## Despliegue (Docker)

`Dockerfile` sirve el sitio con nginx (no hace falta Node ni build: es HTML
estático que lee Google Sheets del lado del navegador). En cada push a
`main`, `.github/workflows/docker-publish.yml` construye la imagen y la
publica en `ghcr.io/tomas-gen/menuweb_tina:latest`.

Probar la imagen local:

```bash
docker build -t menuweb-tina .
docker run --rm -p 8081:80 menuweb-tina
# abrir http://localhost:8081
```

En el servidor (Watchtower detecta la imagen nueva sola), ver
`docker-compose.example.yml` — la URL prevista es `tina.menuweb.ar`.
