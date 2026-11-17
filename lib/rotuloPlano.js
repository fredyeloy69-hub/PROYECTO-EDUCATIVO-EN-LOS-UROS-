// lib/rotuloPlano.js — usado por lib/googleDrive.js (analizarPdf)
// --- Lectura del RÓTULO de cada plano (N° de lámina, nombre del plano, bloque) ---
// Se lee el texto real del PDF (AutoCAD lo exporta como texto) con su posición
// en la hoja. Se buscan las etiquetas del rótulo ("N° DE LÁMINA:", "PLANO:",
// "BLOQUE:") y se toma el valor escrito debajo (o a la derecha) de cada una.
// Si el PDF no trae texto (rótulo dibujado como líneas), devuelve nulos.

const MAX_PAGINAS_ROTULO = 80;
// Sube este número cuando cambie la forma de leer el rótulo: los PDFs leídos
// con una versión anterior se vuelven a leer en la próxima exportación.
export const VERSION_ROTULO = 2;

const RE_ETIQUETA = /:\s*$/; // las etiquetas del rótulo terminan en ":"
const RE_LAMINA = /L[AÁ]M(?:INA)?\b/i;
const RE_PLANO = /^\s*(?:NOMBRE\s+DEL\s+|T[IÍ]TULO\s+DEL\s+)?PLANO\s*:?\s*$/i;
const RE_CLAVE = /^\s*CLAVE\s*:?\s*$/i;
const RE_ESPECIALIDAD = /^\s*ESPECIALIDAD\s*:?\s*$/i;
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
  const res = { lamina: null, plano: null, bloque: null, clave: null, especialidad: null };

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

  // Sin etiqueta legible (a veces "No. DE LÁMINA" está DIBUJADA con líneas y no
  // es texto): se busca en la zona del rótulo el texto MÁS GRANDE. Si es un
  // código (PG-02) se usa tal cual; si es un número ("03") con letras grandes
  // encima o al lado ("D"), se arma "D-03".
  if (!res.lamina) res.lamina = laminaPorTamano(items, etiquetas);

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

  const etEsp = etiquetas.find((e) => RE_ESPECIALIDAD.test(e.s));
  if (etEsp) {
    const vals = valorDeEtiqueta(etEsp, items, etiquetas);
    const texto = limpiar(vals.map((v) => v.s).join(" "));
    if (texto && !/^[-–—_.\s]+$/.test(texto) && texto.length <= 60) res.especialidad = texto.toUpperCase();
  }

  const etClave = etiquetas.find((e) => RE_CLAVE.test(e.s));
  if (etClave) {
    // El valor más cercano a la etiqueta (no los textos grandes de la lámina que estén al lado)
    const vals = valorDeEtiqueta(etClave, items, etiquetas)
      .filter((v) => v.h <= etClave.h * 4)
      .sort((a, b) => Math.hypot(a.x - etClave.x, a.y - etClave.y) - Math.hypot(b.x - etClave.x, b.y - etClave.y));
    const codigo = vals.find((v) => RE_CODIGO_LAMINA.test(v.s)) || vals[0];
    if (codigo && codigo.s.length <= 20) res.clave = codigo.s.replace(/\s+/g, "").toUpperCase();
  }

  return res;
}

function zonaRotulo(items, etiquetas, ancho, alto) {
  // Etiquetas cerca del borde derecho (rótulo vertical) o del borde inferior
  // (rótulo horizontal). Se toma la que tenga más etiquetas.
  const der = etiquetas.filter((e) => e.x > ancho * 0.72);
  const inf = etiquetas.filter((e) => e.y > alto * 0.72);
  if (der.length >= 3 && der.length >= inf.length) {
    const x0 = Math.min(...der.map((e) => e.x)) - 30;
    return (it) => it.x >= x0;
  }
  if (inf.length >= 3) {
    const y0 = Math.min(...inf.map((e) => e.y)) - 30;
    return (it) => it.y >= y0;
  }
  return null;
}

function laminaPorTamano(items, etiquetas) {
  const ancho = Math.max(...items.map((i) => i.x), 1);
  const alto = Math.max(...items.map((i) => i.y), 1);
  const enZona = zonaRotulo(items, etiquetas, ancho, alto);
  if (!enZona) return null;
  const hEtiq = etiquetas.length ? etiquetas.map((e) => e.h).sort((a, b) => a - b)[Math.floor(etiquetas.length / 2)] : 8;
  const grandes = items
    .filter((it) => enZona(it) && !esEtiqueta(it) && it.h >= hEtiq * 2.5 && it.s.length <= 12)
    .sort((a, b) => b.h - a.h);
  if (grandes.length === 0) return null;
  const top = grandes[0];
  if (RE_CODIGO_LAMINA.test(top.s)) return top.s.replace(/\s+/g, "").toUpperCase();
  if (/^\d{1,3}$/.test(top.s)) {
    // letras grandes justo encima o a la izquierda del número
    const letra = grandes.find(
      (g) =>
        g !== top &&
        /^[A-ZÑ]{1,4}\.?-?$/i.test(g.s) &&
        Math.abs(g.x - top.x) < top.h * 2.5 &&
        top.y - g.y > -top.h * 0.5 &&
        top.y - g.y < top.h * 2.2
    );
    if (letra) return `${letra.s.replace(/[.\-]/g, "").toUpperCase()}-${top.s}`;
    return top.s;
  }
  return null;
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
