// lib/rotuloPlano.js — usado por lib/googleDrive.js (analizarPdf)
// --- Lectura del RÓTULO de cada plano (N° de lámina, nombre del plano, bloque) ---
// Se lee el texto real del PDF (AutoCAD lo exporta como texto) con su posición
// en la hoja. Se buscan las etiquetas del rótulo ("N° DE LÁMINA:", "PLANO:",
// "BLOQUE:") y se toma el valor escrito debajo (o a la derecha) de cada una.
// Si el PDF no trae texto (rótulo dibujado como líneas), devuelve nulos.

const MAX_PAGINAS_ROTULO = 80;

const RE_ETIQUETA = /:\s*$/; // las etiquetas del rótulo terminan en ":"
const RE_LAMINA = /L[AÁ]M(?:INA)?\b/i;
const RE_PLANO = /^\s*PLANO\s*:?\s*$/i;
const RE_BLOQUE = /^\s*BLOQUE\s*:?\s*$/i;
const RE_CODIGO_LAMINA = /^[A-ZÑ]{1,5}\d{0,2}\s*-\s*\d{1,3}[A-Z]?$/i;

function limpiar(t) {
  return String(t || "").replace(/\s+/g, " ").trim();
}

function itemsVisibles(textContent, viewport) {
  const out = [];
  for (const it of textContent.items || []) {
    const s = limpiar(it.str);
    if (!s) continue;
    const t = it.transform || [1, 0, 0, 1, 0, 0];
    const h = Math.hypot(t[2], t[3]);
    if (!(h > 0.5)) continue; // texto oculto (alto 0): atributos de bloque invisibles, se ignora
    const [x, y] = viewport.convertToViewportPoint(t[4], t[5]);
    out.push({ s, x, y, h });
  }
  return out;
}

function esEtiqueta(it) {
  return RE_ETIQUETA.test(it.s) && it.s.length <= 30;
}

// Valor de una etiqueta: textos que están DEBAJO de ella (en la misma columna)
// hasta la siguiente etiqueta; si no hay nada debajo, a su DERECHA en la misma línea.
function valorDeEtiqueta(etiqueta, items, etiquetas) {
  const columna = (it) => it.x >= etiqueta.x - 60 && it.x <= etiqueta.x + 320;
  const siguientes = etiquetas
    .filter((e) => e !== etiqueta && e.y > etiqueta.y + 2 && Math.abs(e.x - etiqueta.x) < 160)
    .map((e) => e.y);
  const limite = siguientes.length ? Math.min(...siguientes) : Infinity;
  const maxDist = etiqueta.h * 25; // no ir demasiado lejos si es la última etiqueta
  let cand = items.filter(
    (it) => !esEtiqueta(it) && it !== etiqueta && it.y > etiqueta.y + 1 && it.y < limite && it.y - etiqueta.y < maxDist && columna(it)
  );
  if (cand.length === 0) {
    cand = items.filter(
      (it) => !esEtiqueta(it) && it !== etiqueta && Math.abs(it.y - etiqueta.y) < etiqueta.h * 1.2 && it.x > etiqueta.x + 5 && it.x - etiqueta.x < 600
    );
  }
  return cand.sort((a, b) => a.y - b.y || a.x - b.x);
}

function leerRotuloDePagina(items) {
  const etiquetas = items.filter(esEtiqueta);
  const res = { lamina: null, plano: null, bloque: null };

  // N° de lámina: la etiqueta de lámina más abajo/derecha (la del rótulo), y
  // de sus valores el que tenga forma de código (PG-02, A-01…), el de letra más grande.
  const etLamina = etiquetas.filter((e) => RE_LAMINA.test(e.s)).sort((a, b) => b.y + b.x - (a.y + a.x))[0];
  if (etLamina) {
    const vals = valorDeEtiqueta(etLamina, items, etiquetas);
    const codigos = vals.filter((v) => RE_CODIGO_LAMINA.test(v.s)).sort((a, b) => b.h - a.h);
    if (codigos.length) res.lamina = codigos[0].s.replace(/\s+/g, "").toUpperCase();
    else {
      const corto = vals.filter((v) => v.s.length <= 12 && /\d/.test(v.s)).sort((a, b) => b.h - a.h)[0];
      if (corto) res.lamina = corto.s.replace(/\s+/g, "").toUpperCase();
    }
  }

  const etPlano = etiquetas.find((e) => RE_PLANO.test(e.s));
  if (etPlano) {
    const vals = valorDeEtiqueta(etPlano, items, etiquetas);
    const texto = limpiar(vals.map((v) => v.s).join(" "));
    if (texto && texto.length <= 160) res.plano = texto.toUpperCase();
  }

  const etBloque = etiquetas.find((e) => RE_BLOQUE.test(e.s));
  if (etBloque) {
    const vals = valorDeEtiqueta(etBloque, items, etiquetas);
    const texto = limpiar(vals.map((v) => v.s).join(" "));
    if (texto && !/^[-–—_.\s]+$/.test(texto) && texto.length <= 40) res.bloque = texto.toUpperCase();
  }

  return res;
}

/**
 * Lee el rótulo de cada página de un PDF.
 * @param {Uint8Array} bytes
 * @returns {Promise<Array<{p:number, lamina:string|null, plano:string|null, bloque:string|null, mm:[number,number]}>|null>}
 */
export async function leerRotulosPdf(bytes) {
  try {
    const { getDocumentProxy } = await import("unpdf");
    const pdf = await getDocumentProxy(bytes);
    const total = Math.min(pdf.numPages, MAX_PAGINAS_ROTULO);
    const rotulos = [];
    for (let p = 1; p <= total; p++) {
      const page = await pdf.getPage(p);
      const viewport = page.getViewport({ scale: 1 });
      const tc = await page.getTextContent();
      const r = leerRotuloDePagina(itemsVisibles(tc, viewport));
      const mmA = Math.round((viewport.width * 25.4) / 72);
      const mmB = Math.round((viewport.height * 25.4) / 72);
      rotulos.push({ p, ...r, mm: [Math.min(mmA, mmB), Math.max(mmA, mmB)] });
      page.cleanup();
    }
    try { await pdf.destroy(); } catch {}
    return rotulos;
  } catch (err) {
    console.error("[leerRotulosPdf] No se pudo leer el texto del PDF:", (err && err.message) || err);
    return null;
  }
}
