// ═══════════════════════════════════════════════════════
//  PRECIOS — Cafetería Tina
//  Editá precios.csv y recargá la página.
//  Para que los cambios del CSV se vean, usá start.command.
// ═══════════════════════════════════════════════════════

const PRECIOS_FALLBACK = {
  cafe_s: 3300, cafe_m: 4200, cafe_l: 4800, cafe_xl: 5500,
  especialidad: 6400,
  materos_clasico: 4600, materos_terere: 4900,
  frozen: 6600,
  te_compartir: 6000, te_individual: 3300, te_yerbeado: 3300, te_matcha_latte: 6400,
  frapu: 6600,
  bebida_agua: 4100, bebida_gaseosa: 4100, jugo_peq: 6000, jugo_gr: 7600,
  pomelada: 8700, limonada_peq: 8700, limonada_gr: 9600,
  pasteleria: 7100,
  alfajor_premium: 6600, alfajor_maicena: 4200, financieros: 2800,
  madeleine: 2800, cookies: 2800, macarons: 7000, budin: 3500,
  scones: 3600, tortita: 850, tostado: 9200, medialuna: 1400,
  medialuna_relleno: 3300, toston: 9200, bruschetta: 9200, galletas_breton: 2100,
  focaccia: 10800,
  moderna: 7500,
  desayuno_americano: 14400, desayuno_tentacion: 13800,
  desayuno_criollo: 10800, desayuno_alfajor: 10700,
  promo_clasico: 7800, promo_minis: 9000, promo_criollo: 8300,
  promo_alfajor: 8000, promo_tentacion: 14500, promo_saludable: 13800,
  promo_americano: 12800, promo_stella: 16000, promo_carlitos: 16600,
  promo_gran_americano: 17400, promo_trio_macarons: 12800,
  promo_tina: 17400, promo_degustacion: 37200, promo_tazon_extra: 1500,
};

function formatPrecio(n) {
  return '$' + Number(n).toLocaleString('es-AR');
}

function aplicarPrecios(obj) {
  document.querySelectorAll('[data-precio]').forEach(el => {
    const key = el.dataset.precio;
    if (obj[key] !== undefined) el.textContent = formatPrecio(obj[key]);
  });
  window.PRECIOS = obj;
}

function parseCSV(text) {
  const lines = text.trim().split('\n').slice(1);
  const obj = {};
  lines.forEach(line => {
    const parts = line.split(',');
    const clave = parts[0];
    const precio = parts[3];
    if (clave && precio) obj[clave.trim()] = parseInt(precio.trim(), 10);
  });
  return obj;
}

async function initPrecios() {
  try {
    const res = await fetch('precios.csv');
    if (!res.ok) throw new Error('fetch failed');
    const text = await res.text();
    aplicarPrecios(parseCSV(text));
  } catch {
    aplicarPrecios(PRECIOS_FALLBACK);
  }
}

document.addEventListener('DOMContentLoaded', initPrecios);
