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
