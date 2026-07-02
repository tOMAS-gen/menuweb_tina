// ═══════════════════════════════════════════════════════
//  MENÚ — Cafetería Tina
//  Editá la planilla de Google Sheets (una hoja por sector) y recargá la
//  página. La planilla tiene que seguir compartida como "cualquiera con
//  el enlace puede ver" para que el sitio la pueda leer sin iniciar sesión.
// ═══════════════════════════════════════════════════════

// ID de la planilla en Google Sheets (de la URL: /spreadsheets/d/<ID>/edit...).
const GOOGLE_SHEET_ID = '1JbJnb5L7dE3Kcq6a6WUC36q9b-78H4od';
const GOOGLE_SHEET_XLSX_URL = `https://docs.google.com/spreadsheets/d/${GOOGLE_SHEET_ID}/export?format=xlsx`;

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
  return div.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function esSiONo(raw, porDefecto) {
  if (raw == null || raw === '') return porDefecto;
  const v = String(raw).trim().toLowerCase();
  return v === 'si' || v === 'sí';
}

function normalizarItem(raw) {
  const num = v => {
    if (v === undefined || v === null || v === '') return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  };
  const str = v => (v === undefined || v === null || v === '' ? null : String(v).trim());
  return {
    nombre: str(raw.nombre),
    descripcion: str(raw.descripcion),
    disponible: esSiONo(raw.disponible, true),
    compartir: esSiONo(raw.compartir, false),
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
    const card = `
      <article class="promo-card${it.disponible ? '' : ' agotado'}">
        <div><h3>${esc(it.nombre)}</h3><p>${esc(it.descripcion || '')}</p></div>
        <strong>${it.precio != null ? formatPrecio(it.precio) : ''}</strong>
        ${it.disponible ? '' : agotadoLabel()}
      </article>`;
    return it.compartir
      ? `<div class="promo-share"><div class="promo-tag">para compartir</div>${card}</div>`
      : card;
  }).join('');
}

// Lee la constante "nombre_constante"/"precio_constante" de la hoja Promos.
// Vive en columnas aparte de la tabla de ítems (separadas por una columna en
// blanco), por eso se busca por nombre de columna y no por posición fija:
// no importa si está en la fila 2 o en cualquier otra, ni si se reordenan
// las columnas de la tabla principal.
function leerConstantePromos(filas2D) {
  if (!filas2D || !filas2D.length) return null;
  const headers = filas2D[0];
  const idxNombre = headers.indexOf('nombre_constante');
  const idxPrecio = headers.indexOf('precio_constante');
  if (idxNombre === -1 || idxPrecio === -1) return null;
  for (let i = 1; i < filas2D.length; i++) {
    const fila = filas2D[i];
    if (fila && fila[idxNombre] != null && fila[idxPrecio] != null) {
      return { nombre: String(fila[idxNombre]).trim(), precio: Number(fila[idxPrecio]) };
    }
  }
  return null;
}

function aplicarTazonExtra(constante) {
  const el = document.getElementById('promo-tazon-extra');
  if (!el || !constante || !Number.isFinite(constante.precio)) return;
  el.textContent = formatPrecio(constante.precio);
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
    if (hoja === 'Promos') {
      const filas2D = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
      aplicarTazonExtra(leerConstantePromos(filas2D));
    }
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
    if (hoja === 'Promos') aplicarTazonExtra(fallback['Promos_tazon_extra']);
  });
}

function ocultarPantallaCarga() {
  const el = document.getElementById('loading-screen');
  if (!el) return;
  el.classList.add('oculto');
  setTimeout(() => el.remove(), 500);
}

async function initMenu() {
  try {
    const res = await fetch(GOOGLE_SHEET_XLSX_URL, { cache: 'no-store' });
    if (!res.ok) throw new Error('fetch de Google Sheets falló: ' + res.status);
    const buf = await res.arrayBuffer();
    const workbook = XLSX.read(buf, { type: 'array' });
    renderDesdeWorkbook(workbook);
    window.PRECIOS_CARGADOS = true;
  } catch (err) {
    console.error('No se pudo leer la planilla de Google Sheets, uso el respaldo assets/menu-fallback.json.', err);
    try {
      const res = await fetch('assets/menu-fallback.json');
      const fallback = await res.json();
      renderDesdeFallback(fallback);
      window.PRECIOS_CARGADOS = true;
    } catch (err2) {
      console.error('Tampoco se pudo leer el respaldo. El menú queda vacío.', err2);
      window.PRECIOS_CARGADOS = false;
    }
  } finally {
    ocultarPantallaCarga();
  }
}

document.addEventListener('DOMContentLoaded', initMenu);
