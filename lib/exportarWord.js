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

// Ancho reservado para el número en cada nivel (en twips, 567 = 1 cm).
// Cada nivel empieza donde empieza el TEXTO del nivel anterior — igual que
// una lista multinivel de Word — y el texto largo queda alineado con sangría
// francesa (no vuelve debajo del número).
const ANCHO_NUMERO_NIVEL = [567, 680, 794, 907, 1021, 1134, 1134, 1134];

function sangriaNivel(nivel) {
  let inicio = 0;
  for (let n = 1; n < nivel; n++) inicio += ANCHO_NUMERO_NIVEL[Math.min(n - 1, ANCHO_NUMERO_NIVEL.length - 1)];
  const ancho = ANCHO_NUMERO_NIVEL[Math.min(nivel - 1, ANCHO_NUMERO_NIVEL.length - 1)];
  return { left: inicio + ancho, hanging: ancho };
}

function parrafoIndice(item) {
  const { left, hanging } = sangriaNivel(item.nivel);
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
    for (const item of items) children.push(parrafoIndice(item));
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

// Separa "PU-01 PLANO DE UBICACION A2.pdf" en lámina "PU-01" y nombre
// "PLANO DE UBICACION". Si el archivo no empieza con un código de lámina,
// la lámina queda vacía y el nombre es el nombre completo.
function separarLaminaYNombre(nombreArchivo) {
  let base = (nombreArchivo || "").replace(/\.pdf$/i, "").replace(/_/g, " ").replace(/\s+/g, " ").trim();
  let lamina = "";
  let prefijo = "";
  const m = base.match(/^([A-ZÑ]{1,4}\d{0,2})\s*([-.]?)\s*(\d{1,3})(?=[\s\-.,]|$)/i);
  if (m) {
    prefijo = m[1].toUpperCase();
    lamina = `${prefijo}-${m[3]}`;
    base = base.slice(m[0].length);
  }
  const restoNombre = base;
  let nombre = base
    .replace(/^[\s\-.,:]+/, "")
    .replace(/[\s\-(]*\bA-?[0-4]\b\)?\s*$/i, "") // quita el formato si viene al final ("... A2")
    .replace(/[\s\-.,]+$/, "")
    .trim()
    .toUpperCase();
  return { lamina, prefijo, nombre, restoNombre };
}

// Dentro de una ruta, ubica la carpeta de PLANOS (la que lleva "PLANO" en el nombre).
function indiceCarpetaPlanos(partes) {
  for (let i = 1; i < partes.length; i++) {
    if (/PLANO/i.test(partes[i])) return i;
  }
  return -1;
}

// Código de agrupación: si el plano está en una subcarpeta propia dentro de
// la carpeta de planos (ej. "BLOQUE A1"), se usa esa subcarpeta ("A1");
// si no, las letras del código de lámina ("PU-01" → "PU").
function codigoDelPlano(partes, idxPlanos, prefijoLamina) {
  const sub = partes
    .slice(idxPlanos + 1)
    .filter((s) => !esCarpetaContenedor(s) && !/^\s*A-?[0-4]\s*$/i.test(s));
  if (sub.length > 0) {
    const { descripcion } = extraerNumeroYDescripcion(sub[sub.length - 1]);
    return descripcion.replace(/^BLOQUE\s+/i, "").toUpperCase();
  }
  return prefijoLamina || "";
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

  for (const c of carpetasDelArea) {
    const partes = partesRuta(c);
    const idx = indiceCarpetaPlanos(partes);
    if (idx < 0) continue;

    // Si la propia carpeta de 1er nivel se llama "PLANOS ...", la especialidad
    // es la subcarpeta siguiente (ej. "05 PLANOS / ARQUITECTURA").
    let clave;
    let tituloSegmento;
    let idxPlanos = idx;
    if (idx === 1) {
      const siguiente = partes[2] && !esCarpetaContenedor(partes[2]) ? partes[2] : null;
      clave = siguiente ? partes.slice(1, 3).join(" / ") : partes[1];
      tituloSegmento = siguiente || partes[1];
      if (siguiente) idxPlanos = 2;
    } else {
      clave = partes.slice(1, idx).join(" / ");
      tituloSegmento = partes[idx - 1];
    }
    const titulo = extraerNumeroYDescripcion(tituloSegmento).descripcion.replace(/^ESPECIALIDAD\s+DE\s+/i, "").toUpperCase();

    const pdfs = (c.archivos || []).filter((a) => a.tipo === "pdf" || /\.pdf$/i.test(a.nombre || ""));
    if (pdfs.length === 0) continue;

    if (!grupos[clave]) grupos[clave] = { clave, titulo, planos: [] };
    const subcarpetaClave = partes
      .slice(idxPlanos + 1)
      .filter((s) => !esCarpetaContenedor(s) && !/^\s*A-?[0-4]\s*$/i.test(s))
      .join(" / ");

    for (const archivo of pdfs) {
      const { lamina, prefijo, nombre, restoNombre } = separarLaminaYNombre(archivo.nombre);
      const formato = formatoDelPlano(archivo, partes, restoNombre);
      const codigo = codigoDelPlano(partes, idxPlanos, prefijo);
      grupos[clave].planos.push({
        // Primero los planos sueltos de la carpeta de planos (ubicación,
        // generales…), luego los de cada subcarpeta (bloques, detalles…),
        // cada grupo en el orden natural de Drive.
        orden: `${subcarpetaClave ? "1" : "0"} / ${subcarpetaClave} / ${archivo.nombre || ""}`,
        codigo,
        nombre: nombre || (archivo.nombre || "").replace(/\.pdf$/i, "").toUpperCase(),
        lamina,
        formato: formato.texto,
      });
      if (!formato.texto) {
        avisos.push(`Sin formato (${titulo}): ${archivo.nombre} — el PDF aún no se pudo medir y el nombre no trae A0–A4`);
      } else if (formato.mixto) {
        avisos.push(`Varios formatos en un solo PDF (${titulo}): ${archivo.nombre} → ${formato.texto}`);
      }
      if (archivo.paginas > 1) {
        avisos.push(`PDF con ${archivo.paginas} hojas (${titulo}): ${archivo.nombre} — se listó como 1 plano`);
      }
    }
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
