import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Tab,
  Table,
  TableRow,
  TableCell,
  WidthType,
  TableLayoutType,
  BorderStyle,
  ShadingType,
  VerticalAlign,
  AlignmentType,
  HeightRule,
  TabStopType,
} from "docx";
import { clasificarHojaMm } from "@/lib/formatosHoja";

// =====================================================================
//  EXPORTES A WORD (.docx) — Visor Uros
//  1) generarIndiceGeneralWord: índice de la documentación, separado por
//     especialidad, con la jerarquía y tabulación de las carpetas de Drive.
//  2) generarListaPlanosWord: "RELACIÓN DE PLANOS" por especialidad, tabla
//     de 5 columnas (ORDEN · CÓDIGO · NOMBRE DE PLANO · N° DE LÁMINA · FORMATO)
//     con el mismo formato de las tablas de los informes.
// =====================================================================

const CODIGO_PROYECTO = "UROS";
const FUENTE = "Arial";

// Hoja A4 vertical con los mismos márgenes de los informes de la SGED.
const PAGINA_A4 = {
  size: { width: 11906, height: 16838 },
  margin: { top: 2160, bottom: 1440, left: 1559, right: 1275 },
};

const NIVEL_MAXIMO_INDICE = 8;

const PALABRAS_CONTENEDOR = [
  "PDF", "EDITABLE", "EDITABLES", "EXCEL", "PLANILLA", "PLANILLAS",
  "CAD", "DWG", "WORD", "DOC", "DOCX", "RVT", "REVIT",
];

// ---------------------------------------------------------------------
//  Utilidades comunes (mismas reglas que exportarExcel.js)
// ---------------------------------------------------------------------

function comparaNatural(a, b) {
  return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
}

function extraerNumeroYDescripcion(nombre) {
  const texto = (nombre || "").trim();
  const match = texto.match(/^(\d+(?:\.\d+)*)[\.\_\-\s]+(.+)$/);
  if (match && match[2].trim()) {
    return { numero: match[1], descripcion: match[2].trim() };
  }
  return { numero: "", descripcion: texto };
}

function partesRuta(carpeta) {
  return (carpeta.ruta || carpeta.nombre || "").split(" / ").map((s) => s.trim()).filter(Boolean);
}

function esCarpetaContenedor(nombre) {
  const upper = (nombre || "").toUpperCase();
  return PALABRAS_CONTENEDOR.some((palabra) => upper.includes(palabra));
}

function nombreArchivoSeguro(texto) {
  return (texto || "").replace(/[^a-zA-Z0-9]+/g, "_");
}

async function descargarDocx(doc, nombreArchivo) {
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = nombreArchivo;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function hoyISO() {
  return new Date().toISOString().slice(0, 10);
}

// =====================================================================
//  1) ÍNDICE GENERAL EN WORD
// =====================================================================

// Árbol completo de carpetas del área (igual que el Índice General en Excel).
function construirArbolCarpetas(carpetasDelArea, nivelMaximo) {
  const raiz = {};
  for (const c of carpetasDelArea) {
    const segmentos = partesRuta(c).slice(1, 1 + nivelMaximo); // quita el área
    let nodo = raiz;
    for (const seg of segmentos) {
      if (!nodo[seg]) nodo[seg] = {};
      nodo = nodo[seg];
    }
  }
  return raiz;
}

// Mismo recorrido que la Lista General en Excel: una fila por cada carpeta
// que trae ÍTEM en su nombre ("2.3.1 PLANOS" → 2.3.1). Las carpetas sin
// ítem (PDF, EDITABLE, etc.) no salen, pero sus hijas sí, sin hundirse de nivel.
function generarFilasConItem(nodo, nivel, resultado) {
  const claves = Object.keys(nodo).sort(comparaNatural);
  for (const clave of claves) {
    const { numero, descripcion } = extraerNumeroYDescripcion(clave);
    const tieneHijos = Object.keys(nodo[clave]).length > 0;
    if (numero !== "") {
      resultado.push({ numero, nombre: descripcion, nivel, tieneHijos });
      generarFilasConItem(nodo[clave], nivel + 1, resultado);
    } else {
      generarFilasConItem(nodo[clave], nivel, resultado);
    }
  }
}

// Ancho del número en cada nivel, calculado según el número MÁS LARGO de
// ese nivel en todo el documento (ej. nivel 3 "10.3.2.1", nivel 4 "10") +
// un espacio fijo. Así el texto queda pegado al número, sin huecos, y todos
// los ítems del mismo nivel quedan alineados. Cada nivel empieza donde
// empieza el TEXTO del nivel anterior (como una lista multinivel de Word) y
// el texto largo queda con sangría francesa (no vuelve debajo del número).
const ESPACIO_NUMERO_TEXTO = 170; // twips (~0.3 cm) entre el número y el texto

// Ancho aproximado en twips de un texto en Arial (1 pt = 20 twips).
function anchoTextoArial(texto, tamMedioPuntos, negrita) {
  let em = 0;
  for (const ch of texto) {
    if (/[0-9]/.test(ch)) em += 0.556;
    else if (ch === "." || ch === ",") em += 0.278;
    else if (ch === " ") em += 0.278;
    else em += 0.667;
  }
  return Math.ceil(em * (tamMedioPuntos / 2) * 20 * (negrita ? 1.06 : 1));
}

// Calcula la sangría de CADA ítem. El ancho del número se mide por GRUPO DE
// HERMANOS (los ítems que cuelgan del mismo padre): así "1, 2, 3…" de un
// grupo no heredan el ancho de un "10.3.2.1" de otra rama. Cada grupo empieza
// justo donde empieza el texto de su padre.
function calcularSangrias(items) {
  const padre = new Array(items.length).fill(-1);
  const pila = []; // índices de los ancestros abiertos
  items.forEach((it, i) => {
    while (pila.length && items[pila[pila.length - 1]].nivel >= it.nivel) pila.pop();
    padre[i] = pila.length ? pila[pila.length - 1] : -1;
    pila.push(i);
  });

  const anchoGrupo = {}; // índice del padre → ancho del número
  items.forEach((it, i) => {
    const esp = it.nivel === 1;
    const num = esp ? `${it.numero}.` : it.numero;
    const w = anchoTextoArial(num, esp ? 22 : 20, true) + ESPACIO_NUMERO_TEXTO;
    anchoGrupo[padre[i]] = Math.max(anchoGrupo[padre[i]] || 0, w);
  });

  const textoIzq = new Array(items.length).fill(0);
  return items.map((it, i) => {
    const inicio = padre[i] >= 0 ? textoIzq[padre[i]] : 0;
    const ancho = anchoGrupo[padre[i]];
    textoIzq[i] = inicio + ancho;
    return { left: inicio + ancho, hanging: ancho };
  });
}

function parrafoIndice(item, sangria) {
  const { left, hanging } = sangria;
  const esEspecialidad = item.nivel === 1;
  const negrita = esEspecialidad || item.tieneHijos;
  const numero = esEspecialidad ? `${item.numero}.` : item.numero;
  const texto = esEspecialidad ? item.nombre.toUpperCase() : item.nombre;
  const tam = esEspecialidad ? 22 : 20; // 11 pt / 10 pt

  return new Paragraph({
    indent: { left, hanging },
    tabStops: [{ type: TabStopType.LEFT, position: left }],
    spacing: {
      before: esEspecialidad ? 240 : 40,
      after: esEspecialidad ? 80 : 40,
      line: 264,
    },
    keepNext: esEspecialidad, // el título de la especialidad no queda solo al final de una página
    children: [
      new TextRun({ text: numero, bold: negrita, size: tam, font: FUENTE }),
      new TextRun({ children: [new Tab(), texto], bold: negrita, size: tam, font: FUENTE }),
    ],
  });
}

export function construirIndiceGeneralDoc(areaNombre, carpetasDelArea) {
  const arbol = construirArbolCarpetas(carpetasDelArea, NIVEL_MAXIMO_INDICE);
  const items = [];
  generarFilasConItem(arbol, 1, items);
  // Negrita solo si de verdad tiene sub-ítems listados debajo (una carpeta
  // que solo tiene "PDF"/"EDITABLE" adentro cuenta como ítem final).
  items.forEach((it, i) => {
    it.tieneHijos = i + 1 < items.length && items[i + 1].nivel > it.nivel;
  });

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
      children: [new TextRun({ text: "ÍNDICE GENERAL DE LA DOCUMENTACIÓN", bold: true, size: 24, font: FUENTE })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: areaNombre.toUpperCase(), bold: true, size: 20, font: FUENTE })],
    }),
  ];

  if (items.length === 0) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "No se encontraron carpetas con número de ítem en su nombre (ej. \"2.3 PLANOS\").",
            italics: true,
            size: 20,
            font: FUENTE,
          }),
        ],
      })
    );
  } else {
    const sangrias = calcularSangrias(items);
    items.forEach((item, i) => children.push(parrafoIndice(item, sangrias[i])));
  }

  const doc = new Document({
    creator: "Visor Uros",
    title: `Índice General — ${areaNombre}`,
    styles: { default: { document: { run: { font: FUENTE, size: 20 } } } },
    sections: [{ properties: { page: PAGINA_A4 }, children }],
  });
  return { doc, totalItems: items.length };
}

export async function generarIndiceGeneralWord(areaNombre, carpetasDelArea) {
  const { doc } = construirIndiceGeneralDoc(areaNombre, carpetasDelArea);
  await descargarDocx(doc, `Indice_General_${CODIGO_PROYECTO}_${nombreArchivoSeguro(areaNombre)}_${hoyISO()}.docx`);
}

// =====================================================================
//  2) RELACIÓN DE PLANOS EN WORD
// =====================================================================

const FORMATOS = ["A4", "A3", "A2", "A1", "A0"];

// Formato por el TAMAÑO REAL de cada hoja del PDF (mismo criterio que el
// reporte de Material de Impresión). Devuelve { A1: 3, A3: 1, ... } o null
// si el PDF todavía no fue analizado.
function formatosRealesDelPdf(archivo) {
  if (!archivo.formatos || archivo.paginas == null) return null;
  const formatos = { A4: 0, A3: 0, A2: 0, A1: 0, A0: 0 };
  const tam = Array.isArray(archivo.tamanos) ? archivo.tamanos : null;
  const sumaTam = tam ? tam.reduce((s, t) => s + (t.n || 0), 0) : 0;

  if (tam && sumaTam === archivo.paginas) {
    for (const t of tam) formatos[clasificarHojaMm(t.mm[0], t.mm[1]).formato] += t.n;
    return formatos;
  }
  for (const f of FORMATOS) formatos[f] = (archivo.formatos && archivo.formatos[f]) || 0;
  for (const x of archivo.aprox || []) {
    const [c, l] = String(x.mm).split("x").map(Number);
    if (!(c > 0) || !(l > 0)) continue;
    if (formatos[x.formato] != null) formatos[x.formato] = Math.max(0, formatos[x.formato] - x.n);
    formatos[clasificarHojaMm(c, l).formato] += x.n;
  }
  return formatos;
}

function formatoDeTexto(texto) {
  const m = (texto || "").match(/(?:^|[^A-Z0-9])A-?([0-4])(?![0-9A-Z])/i);
  return m ? `A${m[1]}` : null;
}

// "A1" → "A-1" (así se escribe en las tablas de los informes)
function formatoParaTabla(f) {
  return f ? f.replace(/^A(\d)$/, "A-$1") : "";
}

// Formato del plano: 1º tamaño real del PDF; 2º si aún no se analizó, el
// nombre del archivo; 3º el nombre de las carpetas. Si no hay nada, vacío —
// nunca se adivina.
function formatoDelPlano(archivo, partesCarpeta, restoNombre) {
  const reales = formatosRealesDelPdf(archivo);
  if (reales) {
    const presentes = FORMATOS.filter((f) => reales[f] > 0).sort((a, b) => reales[b] - reales[a]);
    if (presentes.length > 0) {
      return { texto: presentes.map(formatoParaTabla).join(" / "), fuente: "real", mixto: presentes.length > 1 };
    }
  }
  // Se busca en el nombre SIN el código de lámina, para no confundir la
  // lámina "A-1" con el formato A1.
  const porNombre = formatoDeTexto(restoNombre);
  if (porNombre) return { texto: formatoParaTabla(porNombre), fuente: "nombre", mixto: false };
  for (const seg of [...partesCarpeta].reverse()) {
    const porCarpeta = formatoDeTexto(seg);
    if (porCarpeta && /^\s*A-?[0-4]\s*$/i.test(seg)) return { texto: formatoParaTabla(porCarpeta), fuente: "carpeta", mixto: false };
  }
  return { texto: "", fuente: null, mixto: false };
}

// Respaldo cuando el rótulo no se pudo leer: saca la lámina del NOMBRE DEL
// ARCHIVO, esté donde esté ("PU-01 PLANO DE UBICACION", "CORTES GENERALES
// MAYO PG-03", "PLANO-A-01-PLANIMETRIA"…) y deja el resto como nombre.
const RE_LAMINA_EN_NOMBRE = /(^|[^A-Z0-9Ñ])([A-ZÑ]{1,4}\d{0,2})\s*-\s*(\d{1,3})(?![0-9])/gi;

function separarLaminaYNombre(nombreArchivo) {
  let base = (nombreArchivo || "").replace(/\.pdf$/i, "").replace(/_/g, " ").replace(/\s+/g, " ").trim();
  let lamina = "";
  let prefijo = "";
  const coincidencias = [...base.matchAll(RE_LAMINA_EN_NOMBRE)];
  if (coincidencias.length > 0) {
    const m = coincidencias[coincidencias.length - 1]; // el código suele ir al final o al inicio; se toma el último
    prefijo = m[2].toUpperCase();
    lamina = `${prefijo}-${m[3]}`;
    base = (base.slice(0, m.index) + m[1] + base.slice(m.index + m[0].length)).trim();
  }
  const restoNombre = base;
  const nombre = base
    .replace(/^\d+(?:\.\d+)*[\s.\-]+/, "") // numeración inicial "3." / "05 -"
    .replace(/(?<=[A-ZÑ]{2})-(?=[A-ZÑ]{2})/gi, " ") // "GENERAL-MAYO" → "GENERAL MAYO"
    .replace(/\s*-{2,}\s*/g, " ") // "PLANO--PLANIMETRIA" (quedó al quitar la lámina)
    .replace(/^PLANO\s*-+\s*/i, "PLANO ")
    .replace(/[\s\-(]*\bA-?[0-4]\b\)?\s*$/i, "") // formato al final ("... A2")
    .replace(/^[\s\-.,:]+|[\s\-.,:]+$/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
  return { lamina, prefijo, nombre, restoNombre };
}

// --- Qué carpetas entran a la Relación de Planos ---
// Solo carpetas cuyo nombre dice PLANO, PLANOS, PLANIMETRÍA, PLANIMETRÍAS
// (con o sin tilde). NO entran "PLAN DE …", "PLANILLA", etc.
const RE_CARPETA_PLANOS = /(^|[^A-ZÑÁÉÍÓÚ])(PLANOS?|PLANIMETR[IÍ]AS?)(?![A-ZÑÁÉÍÓÚ])/i;
// Subcarpetas o archivos de DOCUMENTOS que se descartan aunque estén dentro
// de una carpeta de planos (memorias, metrados, anexos, renders, fotos…).
const RE_DOCUMENTO =
  /MEMORIA|ESPECIFICACI|METRADO|PRESUPUEST|(^|[^A-Z])ANEXOS?(?![A-Z])|INFORME|FOTOGR[AÁ]F|(^|[^A-Z])FOTOS?(?![A-Z])|CERTIFICAD|(^|[^A-Z])ACTAS?(?![A-Z])|CRONOGRAMA|(^|[^A-Z])COSTOS?(?![A-Z])|AN[AÁ]LISIS|RENDER|MODELADO|(^|[^A-Z0-9])3D(?![A-Z])/i;

// Índice de la carpeta de planos dentro de la ruta, o -1 si la carpeta NO
// debe entrar (no hay carpeta de planos, o está dentro de una de documentos).
function indiceCarpetaPlanos(partes) {
  let idx = -1;
  for (let i = 1; i < partes.length; i++) {
    if (RE_CARPETA_PLANOS.test(partes[i]) && !RE_DOCUMENTO.test(partes[i])) {
      idx = i;
      break;
    }
  }
  if (idx < 0) return -1;
  for (let i = idx + 1; i < partes.length; i++) {
    if (RE_DOCUMENTO.test(partes[i])) return -1; // ej. "8.2 PLANOS / MEMORIA DESCRIPTIVA"
  }
  return idx;
}

// También lo usa page.js para medir SOLO los PDFs de estas carpetas.
export function esCarpetaDeListaPlanos(ruta) {
  return indiceCarpetaPlanos((ruta || "").split(" / ").map((x) => x.trim()).filter(Boolean)) >= 0;
}

// Código de agrupación (columna CÓDIGO), en este orden:
//  1) el BLOQUE escrito en el rótulo ("A1", "BLOQUE B" → "B")
//  2) una subcarpeta llamada "BLOQUE …" dentro de la carpeta de planos
//  3) las letras de la lámina ("PG-02" → "PG", "PU-01" → "PU")
function codigoDelPlano(bloqueRotulo, partes, idxPlanos, lamina) {
  if (bloqueRotulo) return bloqueRotulo.replace(/^BLOQUE\s*[:\-]?\s*/i, "").trim().toUpperCase();
  const subBloque = partes.slice(idxPlanos + 1).reverse().find((s) => /^\s*(?:\d+(?:\.\d+)*[.\-_\s]+)?BLOQUE\s+\S/i.test(s));
  if (subBloque) return subBloque.replace(/^\s*(?:\d+(?:\.\d+)*[.\-_\s]+)?BLOQUE\s+/i, "").trim().toUpperCase();
  const m = (lamina || "").match(/^([A-ZÑ]+\d{0,2})-/i);
  return m ? m[1].toUpperCase() : "";
}

// Título de la tabla con el CÓDIGO de la carpeta de Drive de donde sale la
// lista, para ubicarla: "10.3. PLANOS DE LA ESPECIALIDAD ARQUITECTURA".
// Si esa carpeta no tiene número, se antepone la carpeta numerada más
// cercana: "5.2. EXPEDIENTE › ANTEPROYECTO".
function tituloConCodigo(segmentos) {
  const limpios = segmentos.map((x) => x.trim()).filter(Boolean);
  if (limpios.length === 0) return "PLANOS";
  const ultimo = limpios[limpios.length - 1];
  if (extraerNumeroYDescripcion(ultimo).numero) return ultimo.toUpperCase();
  for (let k = limpios.length - 2; k >= 0; k--) {
    if (extraerNumeroYDescripcion(limpios[k]).numero) {
      return [limpios[k], ...limpios.slice(k + 1)].join(" › ").toUpperCase();
    }
  }
  return ultimo.toUpperCase();
}

/**
 * Recorre las carpetas del área y arma la relación de planos por especialidad.
 * Especialidad = carpeta que CONTIENE a la carpeta de planos (ej.
 * "02 ARQUITECTURA / 2.5 PLANOS" → ARQUITECTURA). Solo se listan PDFs (un PDF
 * = un plano); los DWG/editables no se cuentan para no duplicar.
 */
export function armarRelacionPlanos(carpetasDelArea) {
  const grupos = {}; // clave (ruta de la especialidad) → { titulo, planos: [] }
  const avisos = [];
  let omitidasA4 = 0;

  for (const c of carpetasDelArea) {
    const partes = partesRuta(c);
    const idx = indiceCarpetaPlanos(partes);
    if (idx < 0) continue;

    // Si la propia carpeta de 1er nivel se llama "PLANOS ...", la especialidad
    // es la subcarpeta siguiente (ej. "10 PLANOS / 10.3. PLANOS DE ARQUITECTURA").
    let segmentosClave;
    let idxPlanos = idx;
    if (idx === 1) {
      const siguiente = partes[2] && !esCarpetaContenedor(partes[2]) ? partes[2] : null;
      segmentosClave = siguiente ? partes.slice(1, 3) : [partes[1]];
      if (siguiente) idxPlanos = 2;
    } else {
      segmentosClave = partes.slice(1, idx);
    }
    const clave = segmentosClave.join(" / ");
    const titulo = tituloConCodigo(segmentosClave);

    const pdfs = (c.archivos || []).filter(
      (a) => (a.tipo === "pdf" || /\.pdf$/i.test(a.nombre || "")) && !RE_DOCUMENTO.test(a.nombre || "")
    );
    if (pdfs.length === 0) continue;

    if (!grupos[clave]) grupos[clave] = { clave, titulo, planos: [] };
    const subcarpetaClave = partes
      .slice(idxPlanos + 1)
      .filter((s) => !esCarpetaContenedor(s) && !/^\s*A-?[0-4]\s*$/i.test(s))
      .join(" / ");

    for (const archivo of pdfs) {
      const desdeNombre = separarLaminaYNombre(archivo.nombre);
      const rotulos = Array.isArray(archivo.rotulos) && archivo.rotulos.length > 0 ? archivo.rotulos : null;
      const formatoArchivo = formatoDelPlano(archivo, partes, desdeNombre.restoNombre);
      const ordenBase = `${subcarpetaClave ? "1" : "0"} / ${subcarpetaClave} / ${archivo.nombre || ""}`;

      // Hojas a listar: con rótulo leído, UNA FILA POR HOJA del PDF (así un PDF
      // que junta varios detalles sale lámina por lámina); sin rótulo, una fila.
      const hojas = rotulos || [null];
      hojas.forEach((r, k) => {
        const unaSola = hojas.length === 1;
        const lamina = (r && r.lamina) || (unaSola ? desdeNombre.lamina : "");
        const nombre =
          (r && r.plano) ||
          (unaSola ? desdeNombre.nombre : "") ||
          (archivo.nombre || "").replace(/\.pdf$/i, "").toUpperCase();
        let formato = formatoArchivo.texto;
        if (r && Array.isArray(r.mm) && r.mm[0] > 0) {
          formato = formatoParaTabla(clasificarHojaMm(r.mm[0], r.mm[1]).formato);
        } else if (!unaSola) {
          formato = "";
        }
        // Hoja A4 sin N° de lámina = documento (memoria, imagen suelta…), no plano.
        if (!lamina && /^A-4$/.test(formato)) {
          omitidasA4 += 1;
          return;
        }
        grupos[clave].planos.push({
          orden: `${ordenBase} / ${String(k).padStart(3, "0")}`,
          codigo: codigoDelPlano(r && r.bloque, partes, idxPlanos, lamina),
          nombre,
          lamina,
          formato,
        });
        const ref = unaSola ? archivo.nombre : `${archivo.nombre} (hoja ${r ? r.p : k + 1})`;
        if (!lamina) avisos.push(`Sin N° de lámina (${titulo}): ${ref} — no se leyó en el rótulo ni en el nombre`);
        if (!formato) avisos.push(`Sin formato (${titulo}): ${ref} — el PDF aún no se pudo medir y el nombre no trae A0–A4`);
      });

      if (!rotulos && formatoArchivo.mixto) {
        avisos.push(`Varios formatos en un solo PDF (${titulo}): ${archivo.nombre} → ${formatoArchivo.texto}`);
      }
      if (!rotulos && archivo.paginas > 1) {
        avisos.push(`PDF con ${archivo.paginas} hojas sin rótulo legible (${titulo}): ${archivo.nombre} — se listó como 1 plano`);
      }
    }
  }

  if (omitidasA4 > 0) {
    avisos.unshift(`Se omitieron ${omitidasA4} hoja(s) A4 sin N° de lámina (documentos, no planos).`);
  }

  const especialidades = Object.values(grupos)
    .sort((a, b) => comparaNatural(a.clave, b.clave))
    .map((g) => ({ ...g, planos: g.planos.sort((a, b) => comparaNatural(a.orden, b.orden)) }))
    .filter((g) => g.planos.length > 0);

  return { especialidades, avisos };
}

// --- Tabla con el mismo formato de los informes ---
const ANCHOS = [866, 1192, 4163, 1413, 1151]; // ORDEN · CÓDIGO · NOMBRE · LÁMINA · FORMATO
const ANCHO_TABLA = ANCHOS.reduce((s, x) => s + x, 0);
const GRIS_ENCABEZADO = "D8D8D8";
const BORDE = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const BORDES_TABLA = { top: BORDE, bottom: BORDE, left: BORDE, right: BORDE, insideHorizontal: BORDE, insideVertical: BORDE };
const MARGEN_CELDA = { top: 30, bottom: 30, left: 70, right: 70 };

function celda(texto, ancho, { negrita = false, centrado = true, fondo = null, columnSpan, rowSpan } = {}) {
  return new TableCell({
    width: { size: ancho, type: WidthType.DXA },
    columnSpan,
    rowSpan,
    verticalAlign: VerticalAlign.CENTER,
    margins: MARGEN_CELDA,
    shading: fondo ? { type: ShadingType.CLEAR, color: "auto", fill: fondo } : undefined,
    children: [
      new Paragraph({
        alignment: centrado ? AlignmentType.CENTER : AlignmentType.LEFT,
        spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: texto || "", bold: negrita, size: 16, font: FUENTE })],
      }),
    ],
  });
}

function tablaPlanos(especialidad) {
  const filas = [
    // Título de la tabla (se repite si la tabla pasa a otra página)
    new TableRow({
      tableHeader: true,
      cantSplit: true,
      height: { value: 330, rule: HeightRule.ATLEAST },
      children: [
        celda(`RELACIÓN DE PLANOS – ${especialidad.titulo}`, ANCHO_TABLA, {
          negrita: true,
          centrado: true,
          fondo: GRIS_ENCABEZADO,
          columnSpan: 5,
        }),
      ],
    }),
    new TableRow({
      tableHeader: true,
      cantSplit: true,
      height: { value: 450, rule: HeightRule.ATLEAST },
      children: ["ORDEN", "CÓDIGO", "NOMBRE DE PLANO", "N° DE LÁMINA", "FORMATO"].map((t, i) =>
        celda(t, ANCHOS[i], { negrita: true, fondo: GRIS_ENCABEZADO })
      ),
    }),
  ];

  // CÓDIGO combinado verticalmente para todas las láminas seguidas del mismo grupo.
  const planos = especialidad.planos;
  const spans = new Array(planos.length).fill(0); // 0 = celda absorbida por la de arriba
  for (let i = 0; i < planos.length; ) {
    let j = i + 1;
    if (planos[i].codigo) {
      while (j < planos.length && planos[j].codigo === planos[i].codigo) j++;
    }
    spans[i] = j - i;
    i = j;
  }

  planos.forEach((p, i) => {
    const celdas = [celda(String(i + 1), ANCHOS[0], { negrita: true })];
    if (spans[i] > 0) {
      celdas.push(celda(p.codigo, ANCHOS[1], { negrita: true, rowSpan: spans[i] > 1 ? spans[i] : undefined }));
    }
    celdas.push(celda(p.nombre, ANCHOS[2], { centrado: false }));
    celdas.push(celda(p.lamina, ANCHOS[3]));
    celdas.push(celda(p.formato, ANCHOS[4]));
    filas.push(new TableRow({ cantSplit: true, height: { value: 280, rule: HeightRule.ATLEAST }, children: celdas }));
  });

  return new Table({
    width: { size: ANCHO_TABLA, type: WidthType.DXA },
    columnWidths: ANCHOS,
    layout: TableLayoutType.FIXED,
    alignment: AlignmentType.CENTER,
    borders: BORDES_TABLA,
    rows: filas,
  });
}

export function construirListaPlanosDoc(areaNombre, carpetasDelArea) {
  const { especialidades, avisos } = armarRelacionPlanos(carpetasDelArea);

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
      children: [new TextRun({ text: "RELACIÓN DE PLANOS", bold: true, size: 24, font: FUENTE })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: areaNombre.toUpperCase(), bold: true, size: 20, font: FUENTE })],
    }),
  ];

  if (especialidades.length === 0) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "No se encontraron PDFs dentro de carpetas de PLANOS en esta carpeta madre.",
            italics: true,
            size: 20,
            font: FUENTE,
          }),
        ],
      })
    );
  }

  especialidades.forEach((esp, i) => {
    children.push(tablaPlanos(esp));
    // Espacio entre tablas (un párrafo vacío también evita que Word las una)
    children.push(new Paragraph({ spacing: { before: 0, after: i < especialidades.length - 1 ? 240 : 0 }, children: [] }));
  });

  const doc = new Document({
    creator: "Visor Uros",
    title: `Relación de Planos — ${areaNombre}`,
    styles: { default: { document: { run: { font: FUENTE, size: 20 } } } },
    sections: [{ properties: { page: PAGINA_A4 }, children }],
  });

  const totalPlanos = especialidades.reduce((s, e) => s + e.planos.length, 0);
  return { doc, avisos, totalPlanos, totalEspecialidades: especialidades.length };
}

/**
 * Descarga la relación de planos en Word y devuelve un resumen
 * ({ totalPlanos, totalEspecialidades, avisos }) para mostrarlo en pantalla.
 */
export async function generarListaPlanosWord(areaNombre, carpetasDelArea) {
  const { doc, avisos, totalPlanos, totalEspecialidades } = construirListaPlanosDoc(areaNombre, carpetasDelArea);
  await descargarDocx(doc, `Relacion_de_Planos_${CODIGO_PROYECTO}_${nombreArchivoSeguro(areaNombre)}_${hoyISO()}.docx`);
  return { avisos, totalPlanos, totalEspecialidades };
}
