# Menú desde planilla Excel multi-hoja (v2 — sobre menuweb_tina)

Fecha: 2026-07-02

Reemplaza al spec homónimo escrito en `../../../tina/docs/superpowers/specs/` (carpeta de
trabajo anterior, ya no vigente). Arquitectura igual, tres cambios de fondo
respecto de esa versión: carpeta de trabajo (`menuweb_tina/`, repo propio
conectado a GitHub `tOMAS-gen/menuweb_tina`), esquema de precios (2 columnas
genéricas en vez de filas separadas por variante) y columna nueva
`disponible`.

## Problema

`precios.js` sólo trae un objeto clave→precio de respaldo (no hay
`precios.csv` en este repo, así que hoy corre siempre con
`PRECIOS_FALLBACK`). Los nombres de cada ítem y su sección están
hardcodeados en `index.html`. Para agregar un ítem hay que tocar HTML.

## Objetivo

El cliente edita un único archivo Excel (`menu.xlsx`) con una hoja por
sector. Agrega una fila → recarga el navegador → el ítem aparece. Cambia un
nombre → se actualiza. Marca un ítem "no disponible" → se ve tachado como
agotado, sin desaparecer del menú.

## Alcance

Igual que la v1: se migran a la planilla las 13 secciones que son catálogo
de ítems con precio (`cafe`, `especialidad`, `materos`, `frozen`, `tes`,
`frapuccino`, `bebidas`, `pasteleria`, `antojos`, `focaccias`, `moderna`,
`desayunos`, `promos`). Queda fuera `hebras` (panel educativo de tés, sin
precios por ítem).

## Formato del archivo

`menu.xlsx`, una hoja por sector, mismo nombre visible que el `<h2>` de esa
sección.

**Columnas comunes a todas las hojas:**

```
nombre | descripcion | disponible | precio | precio_2 | imagen
```

- `nombre` — obligatorio. Editable: cambiarlo cambia el nombre en el sitio.
- `descripcion` — opcional, texto chico bajo el nombre.
- `disponible` — `si` / `no` (por defecto `si`). En `no`, el ítem se sigue
  mostrando en su lugar pero con precio tachado y una etiqueta "agotado" —
  no desaparece del menú.
- `precio` — el precio que **siempre** está completo. Para ítems con una
  sola variante (peso/tamaño/relleno), es el precio base o el más chico.
- `precio_2` — opcional. Sólo se completa en los ítems que hoy muestran dos
  precios lado a lado sin etiqueta (Jugo natural, Limonada, Medialuna). El
  sitio los muestra igual que ahora: dos precios uno al lado del otro, sin
  texto "chico/grande" — así es como ya se ve hoy, no se agrega una
  etiqueta nueva.
- `imagen` — sólo se usa en la hoja **"Pastelería Moderna"** (única sección
  con foto por ítem). Nombre de archivo dentro de `assets/tortas/` (ya
  existen con nombres descriptivos: `torta-hawaiana.png`, etc. — no hace
  falta renombrar nada esta vez). El cliente normalmente no toca esta
  columna: la imagen la sube y referencia el administrador del sitio, como
  hoy.

**Caso especial — hoja "Café Tradicional":** en vez de `precio`/`precio_2`
usa 4 columnas propias, porque son 4 tamaños reales, no una variante
opcional:

```
nombre | disponible | precio_s | precio_m | precio_l | precio_xl
```

(sin `descripcion` ni `imagen`: ningún ítem de café tradicional los usa hoy).

**Explícitamente fuera de alcance:** columna de descuento — se descarta,
no se implementa.

## Lectura en el sitio

- SheetJS vendorizado (`assets/xlsx.full.min.js`), sin CDN, sitio sigue
  funcionando sin internet en `localhost`.
- `precios.js` pasa a: `fetch('menu.xlsx')` binario → `XLSX.read(...)` →
  recorre cada hoja del libro.
- Cada `<section>` relevante queda con un contenedor vacío
  (`data-sector="cafe"`, etc.) en vez de ítems hardcodeados. El JS genera
  el markup de cada ítem iterando las filas de su hoja, respetando el
  markup específico que ya tiene cada sección (`.size-row` en café,
  `.menu-row` en las hojas simples, `.pastry-card` en Pastelería Moderna,
  `.promo-card`/`.promo-share` en Promos).
- Ítem con `disponible=no`: se agrega una clase CSS (`agotado`) al bloque
  del ítem — precio tachado + etiqueta "agotado" vía CSS, sin sacar el
  ítem del DOM.
- Si falla la lectura de `menu.xlsx`: se loguea el error y se usa
  `assets/menu-fallback.json`, generado por el mismo script que genera
  `menu.xlsx` (una sola fuente de datos, dos salidas — igual que v1).

## Migración de datos existentes

Los ítems y precios reales de `menuweb_tina/index.html` +
`menuweb_tina/precios.js` (que hoy corre siempre con `PRECIOS_FALLBACK`,
por no haber `precios.csv` en este repo) se vuelcan como filas iniciales en
`menu.xlsx`. Único precio que difiere de la v1: `cafe_s` es `$3.300` acá
(la v1 tomó `$5.000` del `precios.csv` de la carpeta vieja, que no existe
en este repo — se usa el valor vigente en `menuweb_tina`).

Dos descripciones que hoy son "inline" (misma línea que el nombre: 
"Gaseosas — Línea Pepsi" y "Medialuna — opc. jamón y queso", clase CSS
`.menu-desc-inline`) pasan a mostrarse en línea aparte (clase `.menu-desc`,
igual que el resto) — cambio visual menor, documentado acá para que no
sorprenda al revisar el resultado.

## Fuera de alcance

- Columna de descuento.
- Editor visual o validación de la planilla.
- Subida de imágenes nuevas por el cliente (sigue siendo tarea manual del
  administrador).
- Sección `hebras` (panel educativo): no se toca.
- Renombrar assets de Pastelería Moderna: ya tienen nombres descriptivos en
  este repo (`torta-*.png`), no hace falta la tarea de renombrado de la v1.
