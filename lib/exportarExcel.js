import ExcelJS from "exceljs";
import { LOGO_PUNO_BASE64 } from "@/lib/logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_COLOR = {
  completa: "FF2A9D8F",
  incompleta: "FFF39C12",
  vacia: "FFC0392B",
};

const AZUL_PETROLEO = "FF4A0E17";
const VERDE_AZULADO = "FF7A1F2B";

const CODIGO_PROYECTO = "ACASO";

function sufijoFiltroArchivo(tipoFiltro) {
  const mapa = {
    todas: "Todo",
    completas: "Solo_Completas",
    incompletas: "Solo_Incompletas",
    vacias: "Solo_Vacias",
    incompletas_vacias: "Incompletas_Vacias",
  };
  return mapa[tipoFiltro] || "Todo";
}

/**
 * Genera y descarga un Excel (.xlsx) con formato real — encabezado institucional,
 * logo, nombre del proyecto, jerarquía de carpetas, colores por estado, bordes y notas al pie visibles.
 *
 * @param {string} areaNombre - nombre del área a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 * @param {string} tipoFiltro - filtro aplicado ("todas" | "completas" | "incompletas" | "vacias" | "incompletas_vacias"), solo afecta el nombre del archivo
 */
export async function generarReporteExcelPorArea(areaNombre, carpetasDelArea, tipoFiltro = "todas") {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Acaso I-2";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Reporte";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 6 },
    { width: 55 },
    { width: 16 },
    { width: 55 },
  ];

  // --- Insertar logo institucional ---
  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({
        base64: base64Data,
        extension: "png",
      });
      sheet.addImage(imageId, {
        tl: { col: 0.2, row: 0.2 },
        ext: { width: 45, height: 50 },
      });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  // --- Encabezado institucional ---
  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:D3");
  sheet.getCell("A3").value = "PROYECTO: MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN ACCASO DISTRITO DE PILCUYO DE LA PROVINCIA DE EL COLLAO DEL DEPARTAMENTO DE PUNO";
  sheet.getCell("A3").font = { bold: true, size: 10, color: { argb: "FF7A1F2B" } };
  sheet.getCell("A3").alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getRow(3).height = 30;

  sheet.mergeCells("A4:D4");
  const celdaTitulo = sheet.getCell("A4");
  celdaTitulo.value = `REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
  sheet.getRow(4).height = 22;

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

  sheet.mergeCells("A5:B5");
  sheet.getCell("A5").value = `Generado: ${new Date().toLocaleString("es-PE")}`;
  sheet.getCell("A5").font = { italic: true, size: 9, color: { argb: "FF666666" } };

  sheet.mergeCells("C5:D5");
  sheet.getCell("C5").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
  sheet.getCell("C5").font = { italic: true, size: 9, color: { argb: "FF666666" } };
  sheet.getCell("C5").alignment = { horizontal: "right" };
  sheet.getRow(5).height = 16;

  sheet.addRow([]); // fila 6 en blanco, de separación

  // --- Encabezado de columnas de la tabla (fila 7) ---
  const FILA_ENCABEZADO = 7;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Agrupar por especialidad ---
  const grupos = {};
  const ordenGrupos = [];
  for (const c of carpetasDelArea) {
    const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
    const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
    if (!grupos[especialidad]) {
      grupos[especialidad] = [];
      ordenGrupos.push(especialidad);
    }
    grupos[especialidad].push(c);
  }

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }
  ordenGrupos.sort(comparaNatural);

  let contadorFila = 1;

  function mezclarConBlanco(hexRgb, factor) {
    const r = parseInt(hexRgb.slice(2, 4), 16);
    const g = parseInt(hexRgb.slice(4, 6), 16);
    const b = parseInt(hexRgb.slice(6, 8), 16);
    const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
    return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
  }

  const FACTORES_HOJA = [0.72, 0.80, 0.86, 0.90, 0.93];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min((nivelVisual || 1) - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(VERDE_AZULADO, factor);
  }

  function contarEstados(items) {
    let completas = 0, incompletas = 0, vacias = 0;
    for (const c of items) {
      if (c.estado === "completa") completas++;
      else if (c.estado === "incompleta") incompletas++;
      else if (c.estado === "vacia") vacias++;
    }
    return { completas, incompletas, vacias };
  }

  function textoResumen(items) {
    const { completas, incompletas, vacias } = contarEstados(items);
    const total = items.length;
    return `  (Total: ${total} — ${completas} completas · ${incompletas} incompletas · ${vacias} vacías)`;
  }

  function agregarSubEncabezado(nombre, nivel, items) {
    // El resumen numérico (completas/incompletas/vacías) solo se muestra en
    // los títulos de nivel 1 y 2 (carpetas mayores) — de ahí para abajo se
    // omite para no saturar los sub-encabezados.
    const resumen = items && nivel <= 2 ? textoResumen(items) : "";
    const fila = sheet.addRow([`➤  ${nombre}${resumen}`]);
    sheet.mergeCells(`A${fila.number}:D${fila.number}`);
    const celda = fila.getCell(1);
    const esNivel1 = nivel === 1;
    const factoresPorNivel = [0, 0.45, 0.65, 0.82];
    const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
    celda.font = { bold: true, size: esNivel1 ? 13 : 11, color: { argb: esNivel1 ? "FFFFFFFF" : "FF1A3840" } };
    celda.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: mezclarConBlanco(VERDE_AZULADO, factor) },
    };
    celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
    fila.height = esNivel1 ? 20 : 17;
  }

  function agregarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const fila = sheet.addRow([contadorFila++, nombreMostrado, ESTADO_LABEL[estado] || estado.toUpperCase(), c.detalle || "-"]);
    const fondo = colorHoja(nivelVisual);

    fila.getCell(1).alignment = { horizontal: "center", vertical: "top" };
    fila.getCell(2).alignment = { wrapText: true, vertical: "top", indent: Math.max(0, (nivelVisual || 1) - 1) * 2 };
    fila.getCell(3).font = { bold: true, color: { argb: ESTADO_COLOR[estado] || "FF666666" } };
    fila.getCell(3).alignment = { horizontal: "center", vertical: "top" };
    fila.getCell(4).alignment = { wrapText: true, vertical: "top" };

    fila.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      cell.border = {
        top: { style: "hair", color: { argb: "FFDDDDDD" } },
        bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
        left: { style: "hair", color: { argb: "FFDDDDDD" } },
        right: { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });
  }

  const NIVEL_MAX_INTERMEDIO = 4;
  function agruparRecursivo(items, nivelIdxRuta, nivelVisual) {
    if (nivelIdxRuta > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length > nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    const entradas = Object.keys(subgrupos).map((key) => ({
      key,
      nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
      esGrupo: !key.startsWith("__directo__"),
    }));
    entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

    for (const entrada of entradas) {
      if (entrada.esGrupo) {
        agregarSubEncabezado(entrada.key, nivelVisual, subgrupos[entrada.key]);
        agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    agregarSubEncabezado(especialidad.toUpperCase(), 1, grupos[especialidad]);
    agruparRecursivo(grupos[especialidad], 2, 1);
  }

  // --- PIE DE PÁGINA VISIBLE AL FINAL DE LA TABLA (Celdas reales con diseño) ---
  sheet.addRow([]); // Espacio de separación

  const filaPie1 = sheet.addRow(["Documento generado automáticamente desde el Sistema de Control de Proyectos - Acaso I-2"]);
  sheet.mergeCells(`A${filaPie1.number}:D${filaPie1.number}`);
  filaPie1.getCell(1).font = { italic: true, size: 9, color: { argb: "FF555555" }, bold: true };
  filaPie1.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  filaPie1.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAEDED" } };
  filaPie1.height = 18;

  const filaPie2 = sheet.addRow([`Fecha y hora de emisión: ${new Date().toLocaleDateString("es-PE")} ${new Date().toLocaleTimeString("es-PE")}`]);
  sheet.mergeCells(`A${filaPie2.number}:D${filaPie2.number}`);
  filaPie2.getCell(1).font = { italic: true, size: 8.5, color: { argb: "FF7F8C8D" } };
  filaPie2.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  filaPie2.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEAEDED" } };
  filaPie2.height = 16;

  // Aplicar bordes finos a las celdas del pie para que formen un bloque estético
  [filaPie1, filaPie2].forEach(f => {
    f.eachCell(cell => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFBDC3C7" } },
        bottom: { style: "thin", color: { argb: "FFBDC3C7" } },
        left: { style: "thin", color: { argb: "FFBDC3C7" } },
        right: { style: "thin", color: { argb: "FFBDC3C7" } }
      };
    });
  });

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  // --- PIE DE PÁGINA NATIVO DE EXCEL (Para impresión) ---
  sheet.headerFooter.oddFooter = "&L&I[Sistema Acaso I-2] Reporte institucional&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_${CODIGO_PROYECTO}_${sufijoFiltroArchivo(tipoFiltro)}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// --- Palabras que indican una carpeta "contenedor de tipo de archivo" (PDF,
// Editable, CAD, Excel, etc.) — cuando el último nivel de una carpeta es uno
// de estos, la lista de separadores usa el penúltimo nivel en su lugar,
// porque el nombre del contenedor no aporta información real a un separador.
const PALABRAS_CONTENEDOR = [
  "PDF", "EDITABLE", "EDITABLES", "EXCEL", "PLANILLA", "PLANILLAS",
  "CAD", "DWG", "WORD", "DOC", "DOCX", "RVT", "REVIT",
];

function esCarpetaContenedor(nombre) {
  const upper = (nombre || "").toUpperCase();
  return PALABRAS_CONTENEDOR.some((palabra) => upper.includes(palabra));
}

function comparaNaturalSep(a, b) {
  return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
}

// Para cada carpeta, calcula los segmentos de ruta "efectivos" para el árbol
// de separadores: se descarta el nombre del área (primer segmento) y, si el
// último nivel es un contenedor de tipo de archivo (PDF/EDITABLE/CAD/etc.),
// se usa el penúltimo nivel en su lugar.
function rutaEfectivaSeparador(carpeta) {
  const partes = (carpeta.ruta || carpeta.nombre || "").split(" / ").filter(Boolean);
  let segmentos = partes.slice(1); // quita el nombre del área
  if (segmentos.length === 0) return [];
  const ultimo = segmentos[segmentos.length - 1];
  if (segmentos.length > 1 && esCarpetaContenedor(ultimo)) {
    segmentos = segmentos.slice(0, -1);
  }
  return segmentos;
}

// Construye un árbol anidado a partir de las rutas efectivas de todas las
// carpetas de un área. Carpetas distintas que colapsan a la misma ruta
// efectiva (ej. "PDF" y "EDITABLE" bajo la misma carpeta madre) se
// deduplican automáticamente al insertarse en el mismo nodo del árbol.
function construirArbolSeparadores(carpetasDelArea) {
  const raiz = {};
  for (const c of carpetasDelArea) {
    const segmentos = rutaEfectivaSeparador(c);
    let nodo = raiz;
    for (const seg of segmentos) {
      if (!nodo[seg]) nodo[seg] = {};
      nodo = nodo[seg];
    }
  }
  return raiz;
}

// Extrae el número que YA viene en el nombre real de la carpeta en Drive
// (ej. "1.20_FICHA TÉCNICA" → item "1.20", descripción "FICHA TÉCNICA").
// Si el nombre no empieza con un número reconocible, el ítem queda vacío
// y la descripción es el nombre completo tal cual.
function extraerNumeroYDescripcion(nombre) {
  const texto = (nombre || "").trim();
  const match = texto.match(/^(\d+(?:\.\d+)*)[\.\_\-\s]+(.+)$/);
  if (match && match[2].trim()) {
    return { numero: match[1], descripcion: match[2].trim() };
  }
  return { numero: "", descripcion: texto };
}

// Recorre el árbol en el mismo orden en que aparece en Drive (orden natural,
// respetando números). Por cada nodo extrae el número y la descripción
// reales desde su propio nombre.
//
// Reglas:
// - Si una carpeta NO tiene número reconocible en su nombre, se OMITE de la
//   lista (no genera fila) — pero sus hijos SÍ se recorren igual, respetando
//   la jerarquía (el nivel no avanza para ellos, ya que ocupan visualmente
//   el lugar de su padre omitido).
// - Cada fila que sí se agrega guarda su `nivel` (para el degradado de color)
//   y si es `esUltimoNivel` (no tiene hijos) — el último nivel no se sombrea.
function generarFilasSeparadores(nodo, nivel, resultado) {
  const claves = Object.keys(nodo).sort(comparaNaturalSep);
  for (const clave of claves) {
    const { numero, descripcion } = extraerNumeroYDescripcion(clave);
    const tieneHijos = Object.keys(nodo[clave]).length > 0;
    if (numero !== "") {
      resultado.push({ numero, nombre: descripcion, nivel, esUltimoNivel: !tieneHijos });
      generarFilasSeparadores(nodo[clave], nivel + 1, resultado);
    } else {
      // Carpeta sin número: se omite la fila, pero sus hijos se mantienen
      // en el mismo nivel visual (no se "hunde" un nivel por un padre que no existe en la lista).
      generarFilasSeparadores(nodo[clave], nivel, resultado);
    }
  }
}

// Recorre el árbol igual que generarFilasSeparadores, pero para la "Lista
// General": en vez de CONFIAR en el número que ya trae el nombre de la
// carpeta (y saltarse las que no lo tienen), genera su PROPIA numeración
// según la posición real de cada carpeta en el árbol — así nunca se omite
// ninguna, sin importar cómo esté nombrada.
function generarFilasGeneral(nodo, nivel, codigoPadre, resultado) {
  const claves = Object.keys(nodo).sort(comparaNaturalSep);
  claves.forEach((clave, i) => {
    const { descripcion } = extraerNumeroYDescripcion(clave);
    const codigo = codigoPadre ? `${codigoPadre}.${i + 1}` : String(i + 1);
    const tieneHijos = Object.keys(nodo[clave]).length > 0;
    resultado.push({ numero: codigo, nombre: descripcion || clave, nivel, esUltimoNivel: !tieneHijos });
    generarFilasGeneral(nodo[clave], nivel + 1, codigo, resultado);
  });
}

function mezclarColorConBlanco(hexRgb, factor) {
  const r = parseInt(hexRgb.slice(2, 4), 16);
  const g = parseInt(hexRgb.slice(4, 6), 16);
  const b = parseInt(hexRgb.slice(6, 8), 16);
  const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
  return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
}

const DORADO_FASE2 = "FFB8860B"; // acento distinto (dorado), exclusivo de este documento "fase 2"
const AMARILLO_EDITABLE = "FFFFF2CC";
const BORDE_GRIS = { style: "thin", color: { argb: "FFBBBBBB" } };
const FACTORES_NIVEL_SEPARADOR = [0, 0.55, 0.68, 0.78, 0.86, 0.91, 0.95];
function colorNivelSeparador(nivel) {
  const factor = FACTORES_NIVEL_SEPARADOR[Math.min(Math.max(nivel - 1, 0), FACTORES_NIVEL_SEPARADOR.length - 1)];
  return mezclarColorConBlanco(DORADO_FASE2, factor);
}

/**
 * Genera y descarga un Excel simple de dos columnas (N° / Descripción),
 * pensado para imprimir los separadores físicos de un expediente técnico —
 * NO es un reporte de avance, es un índice numerado plano de las secciones
 * reales del área, en el mismo orden en que están en Drive.
 *
 * @param {string} areaNombre - nombre del área (carpeta madre) a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarListaSeparadoresExcel(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor ACASO I-2";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Separadores";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 14 },
    { width: 75 },
    { width: 18 },
  ];

  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({ base64: base64Data, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 45, height: 50 } });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:C1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:C2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:C3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `ÍNDICE DE SEPARADORES — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
  sheet.getRow(3).height = 22;

  sheet.mergeCells("A4:C4");
  sheet.getCell("A4").value = `Generado: ${new Date().toLocaleString("es-PE")}  ·  la columna "N° DE ARCHIVADORES" queda en blanco para completar a mano`;
  sheet.getCell("A4").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("A4").alignment = { horizontal: "center" };
  sheet.getRow(4).height = 14;

  sheet.addRow([]); // fila 5 en blanco

  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN", "N° DE ARCHIVADORES"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Construir y recorrer el árbol jerárquico ---
  const arbol = construirArbolSeparadores(carpetasDelArea);
  const items = [];
  generarFilasSeparadores(arbol, 1, items);

  for (const item of items) {
    const fila = sheet.addRow([item.numero, item.nombre, null]);
    fila.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    fila.getCell(1).numFmt = "@"; // forzar texto, para que "1.10" no se lea como número
    fila.getCell(2).alignment = { horizontal: "left", vertical: "middle", wrapText: true };

    if (item.nivel === 1) {
      // Título de primer nivel: fondo sólido dorado, texto blanco en negrita.
      fila.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      });
    } else if (item.nivel === 2) {
      // Nivel 2: sombreado tenue, parejo para TODOS los ítems de este nivel
      // (tengan o no subcarpetas propias) — el sombreado no pasa de aquí.
      const fondo = colorNivelSeparador(item.nivel);
      fila.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      });
      if (!item.esUltimoNivel) fila.getCell(2).font = { bold: true };
    } else if (!item.esUltimoNivel) {
      // Nivel 3 en adelante: sin sombreado — solo negrita si tiene hijos.
      fila.getCell(2).font = { bold: true };
    }

    // La columna de archivadores se completa a mano solo en los ítems
    // finales (sin subcarpetas propias) — son los que corresponden a un
    // separador físico real.
    if (item.esUltimoNivel) {
      fila.getCell(3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMARILLO_EDITABLE } };
    }

    fila.getCell(3).border = { top: BORDE_GRIS, bottom: BORDE_GRIS, left: BORDE_GRIS, right: BORDE_GRIS };

    fila.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        ...cell.border,
        top: cell.border?.top || { style: "hair", color: { argb: "FFDDDDDD" } },
        bottom: cell.border?.bottom || { style: "hair", color: { argb: "FFDDDDDD" } },
        left: cell.border?.left || { style: "hair", color: { argb: "FFDDDDDD" } },
        right: cell.border?.right || { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  sheet.headerFooter.oddFooter = "&L&I[Sistema ACASO I-2] Índice de separadores&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Indice_General_${CODIGO_PROYECTO}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// --- Reporte de material de impresión (hojas A4/A3/A2/A1/A0 necesarias) ---

const FORMATOS = ["A4", "A3", "A2", "A1", "A0"];

// Un archivo cuenta como "plano a clasificar" si es .dwg, o si es un PDF que
// vive dentro de una carpeta cuya ruta contiene la palabra "PLANO" (ej.
// "05 PLANOS", "PLANOS Y DETALLES"). Todo lo demás (PDF/Word/Excel fuera de
// carpetas de planos) se cuenta como documento A4.
function esCarpetaDePlanos(ruta) {
  return /PLANO/i.test(ruta || "");
}

// Busca un token suelto "A0".."A4" en el nombre del archivo (ej. "DIE-03
// A2.pdf" → "A2"). Usa límites de palabra para no confundir "A2" dentro de
// "CASA21" o "PLANTA4" con el formato de hoja.
function extraerFormatoDeNombre(nombre) {
  const match = (nombre || "").match(/\bA([0-4])\b/i);
  return match ? `A${match[1]}` : null;
}

// Detecta el formato de un plano: primero intenta con el nombre del propio
// archivo; si no lo trae, revisa el nombre de las carpetas que lo contienen
// (de la más cercana hacia arriba) — porque a veces, en vez de nombrar cada
// PDF, el usuario nombra la CARPETA "A0", "A1", etc. y ahí adentro pega los
// archivos sin ponerle el formato a cada uno.
function detectarFormatoPlano(archivo, rutaCarpeta) {
  const porNombre = extraerFormatoDeNombre(archivo.nombre);
  if (porNombre) return porNombre;
  const segmentos = (rutaCarpeta || "").split(" / ").filter(Boolean).reverse();
  for (const seg of segmentos) {
    const porCarpeta = extraerFormatoDeNombre(seg);
    if (porCarpeta) return porCarpeta;
  }
  return null;
}

/**
 * Genera y descarga un Excel (.xlsx) con la cantidad de HOJAS (no de
 * archivos) de cada formato (A4, A3, A2, A1, A0) que hacen falta imprimir,
 * por especialidad y en total — para cotizar el material a comprar.
 *
 * SOLO se cuentan archivos PDF (lo que realmente se manda a imprimir; DWG,
 * Word y Excel son formatos de trabajo, no el entregable final), y se suman
 * las PÁGINAS reales de cada PDF (guardadas por syncEngine.js al contar el
 * PDF durante la sincronización) — un archivo puede tener 1 página o 200.
 *
 * El formato de cada plano se detecta leyendo el nombre del archivo (ej.
 * "DIE-03 A2.pdf" → A2). Nada se adivina: los planos sin formato en el
 * nombre, las carpetas sin ningún PDF, y los PDFs que no se pudieron abrir
 * para contar sus páginas quedan listados en la hoja "Observaciones" para
 * revisar a mano.
 *
 * @param {string} areaNombre - nombre del área a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área), cada
 *   una con su array `archivos` ([{ nombre, tipo, paginas }, ...]) tal como
 *   lo guarda syncEngine.js
 */
export async function generarListaMaterialImpresion(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Acaso I-2";
  workbook.created = new Date();

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  // --- Recorrer todas las carpetas del área. SOLO se leen archivos PDF —
  // es lo único que realmente se manda a imprimir; DWG/Word/Excel son
  // formatos de trabajo, no el entregable final. Se suman las PÁGINAS
  // reales de cada PDF (un archivo puede tener 1 o 200 hojas), no la
  // cantidad de archivos.
  //
  // Se agrupa hasta el SEGUNDO nivel: especialidad (ej. "03. OBRAS
  // PROVISIONALES + ESTRUCTURAS") y, dentro de ella, cada subcarpeta real
  // (ej. "3.1 OBRAS PROVISIONALES", "3.2 ESTRUCTURAS") por separado, en vez
  // de mezclarlas todas en una sola fila de la especialidad. ---
  const VACIO = () => ({ A4: 0, A3: 0, A2: 0, A1: 0, A0: 0, sinIdentificar: 0 });
  const porEspecialidad = {}; // { especialidad: { subtotal, subcarpetas: { nombre: {...} } } }
  const totalGeneral = VACIO();
  const observaciones = []; // { especialidad, carpeta, archivo, motivo }

  function obtenerBucket(especialidad, subcarpeta) {
    if (!porEspecialidad[especialidad]) {
      porEspecialidad[especialidad] = { subtotal: VACIO(), subcarpetas: {} };
    }
    const esp = porEspecialidad[especialidad];
    if (!esp.subcarpetas[subcarpeta]) {
      esp.subcarpetas[subcarpeta] = VACIO();
    }
    return esp.subcarpetas[subcarpeta];
  }

  function sumar(especialidad, subcarpeta, formato, cantidad) {
    const bucket = obtenerBucket(especialidad, subcarpeta);
    bucket[formato] += cantidad;
    porEspecialidad[especialidad].subtotal[formato] += cantidad;
    totalGeneral[formato] += cantidad;
  }

  function marcarSinIdentificar(especialidad, subcarpeta, cantidad) {
    sumar(especialidad, subcarpeta, "sinIdentificar", cantidad);
  }

  for (const c of carpetasDelArea) {
    const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
    const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
    const subcarpeta = partes.length > 2 ? partes[2] : "(archivos directos)";
    const esPlanos = esCarpetaDePlanos(c.ruta);
    const nombreCarpeta = c.nombre || partes[partes.length - 1] || "(carpeta)";
    const rutaCarpeta = c.ruta || nombreCarpeta;
    const archivosPdf = (c.archivos || []).filter((a) => a.tipo === "pdf");

    if (archivosPdf.length === 0) {
      // No hay nada que imprimir en esta carpeta — no cuenta como error, pero
      // se avisa porque significa que a este ítem todavía le falta su PDF.
      observaciones.push({
        especialidad,
        ruta: rutaCarpeta,
        carpeta: nombreCarpeta,
        archivo: "—",
        motivo: "Esta carpeta no tiene ningún PDF — no se pudo calcular material a imprimir",
      });
      continue;
    }

    for (const archivo of archivosPdf) {
      if (esPlanos) {
        const formato = detectarFormatoPlano(archivo, rutaCarpeta);
        if (!formato) {
          marcarSinIdentificar(especialidad, subcarpeta, archivo.paginas || 1);
          observaciones.push({
            especialidad,
            ruta: rutaCarpeta,
            carpeta: nombreCarpeta,
            archivo: archivo.nombre,
            motivo:
              archivo.paginas != null
                ? `PDF de plano sin formato de hoja (A0-A4) reconocible en el nombre ni en la carpeta — se contaron sus ${archivo.paginas} página(s) como "sin identificar"`
                : `PDF de plano sin formato de hoja (A0-A4) reconocible en el nombre ni en la carpeta, y no se pudo abrir para contar sus páginas — se contó como 1`,
          });
          continue;
        }
        if (archivo.paginas == null) {
          // Se detectó el formato por el nombre, pero no se pudo abrir el PDF
          // para confirmar cuántas hojas trae (protegido, corrupto, etc.) —
          // se asume 1 hoja (lo más común en un plano) y se deja constancia.
          sumar(especialidad, subcarpeta, formato, 1);
          observaciones.push({
            especialidad,
            ruta: rutaCarpeta,
            carpeta: nombreCarpeta,
            archivo: archivo.nombre,
            motivo: `No se pudo abrir el PDF para confirmar cuántas hojas tiene — se contó como 1 hoja ${formato}`,
          });
        } else {
          sumar(especialidad, subcarpeta, formato, archivo.paginas);
        }
      } else {
        // Documento normal (memoria, especificaciones, presupuesto, etc.) — cada
        // página del PDF es una hoja A4.
        if (archivo.paginas == null) {
          sumar(especialidad, subcarpeta, "A4", 1);
          observaciones.push({
            especialidad,
            ruta: rutaCarpeta,
            carpeta: nombreCarpeta,
            archivo: archivo.nombre,
            motivo: "No se pudo abrir el PDF para contar sus páginas — se contó como 1 hoja A4",
          });
        } else {
          sumar(especialidad, subcarpeta, "A4", archivo.paginas);
        }
      }
    }
  }

  // --- Hoja 1: resumen por especialidad + total general ---
  const sheet = workbook.addWorksheet("Material de Impresión", {
    pageSetup: { orientation: "landscape", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 42 }, // especialidad
    { width: 12 }, // A4
    { width: 12 }, // A3
    { width: 12 }, // A2
    { width: 12 }, // A1
    { width: 12 }, // A0
    { width: 16 }, // sin identificar
  ];

  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,") ? LOGO_PUNO_BASE64.split("base64,")[1] : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({ base64: base64Data, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 45, height: 50 } });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:G1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:G2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:G3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `MATERIAL DE IMPRESIÓN — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
  sheet.getRow(3).height = 24;

  sheet.mergeCells("A4:G4");
  sheet.getCell("A4").value =
    `Generado: ${new Date().toLocaleString("es-PE")}  ·  cantidad de HOJAS reales (no de archivos) a imprimir, según las páginas de cada PDF, desglosada por formato  ·  ver hoja "Observaciones" para los casos que requieren revisión manual`;
  sheet.getCell("A4").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("A4").alignment = { horizontal: "center", wrapText: true };
  sheet.getRow(4).height = 26;

  sheet.addRow([]);

  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["ESPECIALIDAD", "HOJAS A4", "HOJAS A3", "HOJAS A2", "HOJAS A1", "HOJAS A0", "SIN IDENTIFICAR"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  const especialidadesOrdenadas = Object.keys(porEspecialidad).sort(comparaNatural);
  for (const esp of especialidadesOrdenadas) {
    const datosEsp = porEspecialidad[esp];
    const v = datosEsp.subtotal;

    // Fila de la especialidad (nivel 1) — en negrita, con fondo tenue, y el
    // SUBTOTAL de todas sus subcarpetas.
    const filaEsp = sheet.addRow([esp, v.A4, v.A3, v.A2, v.A1, v.A0, v.sinIdentificar]);
    filaEsp.eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEFE0C0" } };
    });
    filaEsp.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    for (let col = 2; col <= 7; col++) {
      filaEsp.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
    }
    filaEsp.getCell(7).font = { bold: true, color: { argb: v.sinIdentificar > 0 ? "FFC0392B" : "FF999999" } };

    // Filas de cada subcarpeta (nivel 2) — con sangría, para que se note que
    // están adentro de la especialidad de arriba.
    const subcarpetasOrdenadas = Object.keys(datosEsp.subcarpetas).sort(comparaNatural);
    for (const sub of subcarpetasOrdenadas) {
      const s = datosEsp.subcarpetas[sub];
      const filaSub = sheet.addRow([sub, s.A4, s.A3, s.A2, s.A1, s.A0, s.sinIdentificar]);
      filaSub.getCell(1).alignment = { horizontal: "left", vertical: "middle", indent: 1 };
      for (let col = 2; col <= 7; col++) {
        filaSub.getCell(col).alignment = { horizontal: "center", vertical: "middle" };
      }
      filaSub.getCell(7).font = { color: { argb: s.sinIdentificar > 0 ? "FFC0392B" : "FF999999" }, bold: s.sinIdentificar > 0 };
      filaSub.eachCell((cell) => {
        cell.border = {
          top: { style: "hair", color: { argb: "FFDDDDDD" } },
          bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
          left: { style: "hair", color: { argb: "FFDDDDDD" } },
          right: { style: "hair", color: { argb: "FFDDDDDD" } },
        };
      });
    }
  }

  const filaTotal = sheet.addRow([
    "TOTAL GENERAL",
    totalGeneral.A4,
    totalGeneral.A3,
    totalGeneral.A2,
    totalGeneral.A1,
    totalGeneral.A0,
    totalGeneral.sinIdentificar,
  ]);
  filaTotal.eachCell((cell) => {
    cell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
  filaTotal.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
  filaTotal.height = 22;

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  // --- Hoja 2: observaciones — carpetas sin PDF, planos sin formato en el
  // nombre, y PDFs que no se pudieron abrir para contar sus páginas. Nunca se
  // adivina: todo lo que no se pudo calcular con certeza queda listado aquí
  // para revisar a mano antes de mandar el documento a cotizar.
  //
  // Se ordena por ESPECIALIDAD y luego por RUTA completa (no solo el nombre
  // de la carpeta) — así las observaciones quedan en el mismo orden en que
  // aparecen las carpetas en el Drive, respetando la jerarquía, en vez de
  // salir mezcladas. ---
  if (observaciones.length > 0) {
    const sheet2 = workbook.addWorksheet("Observaciones");
    sheet2.columns = [{ width: 30 }, { width: 32 }, { width: 45 }, { width: 60 }];
    sheet2.addRow(["ESPECIALIDAD", "CARPETA", "ARCHIVO", "OBSERVACIÓN"]).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_PETROLEO } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });
    const observacionesOrdenadas = [...observaciones].sort((a, b) => {
      const porEspecialidad = comparaNatural(a.especialidad, b.especialidad);
      if (porEspecialidad !== 0) return porEspecialidad;
      const porRuta = comparaNatural(a.ruta, b.ruta);
      if (porRuta !== 0) return porRuta;
      return comparaNatural(a.archivo, b.archivo);
    });
    for (const item of observacionesOrdenadas) {
      sheet2.addRow([item.especialidad, item.carpeta, item.archivo, item.motivo]);
    }
    sheet2.views = [{ state: "frozen", ySplit: 1 }];
  }

  const buffer2 = await workbook.xlsx.writeBuffer();
  const blob2 = new Blob([buffer2], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url2 = URL.createObjectURL(blob2);
  const link2 = document.createElement("a");
  link2.href = url2;
  link2.download = `Material_Impresion_${CODIGO_PROYECTO}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link2);
  link2.click();
  document.body.removeChild(link2);
  URL.revokeObjectURL(url2);
}

/**
 * Genera y descarga un "Índice General" completo — mismo estilo visual que
 * generarListaSeparadoresExcel, pero con una diferencia clave: en vez de
 * confiar en el número que ya trae escrito el nombre de cada carpeta (y
 * saltarse las que no lo traen), genera SU PROPIA numeración según la
 * posición real de cada carpeta dentro del árbol de Drive — así nunca omite
 * ninguna carpeta, sin importar cómo esté nombrada.
 *
 * @param {string} areaNombre - nombre del área (carpeta madre) a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarListaGeneralExcel(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor ACASO I-2";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Indice General";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [{ width: 14 }, { width: 75 }, { width: 18 }];

  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({ base64: base64Data, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 45, height: 50 } });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:C1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:C2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:C3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `ÍNDICE GENERAL DE LA DOCUMENTACIÓN — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
  sheet.getRow(3).height = 22;

  sheet.mergeCells("A4:C4");
  sheet.getCell("A4").value =
    `Generado: ${new Date().toLocaleString("es-PE")}  ·  numeración generada según la posición real de cada carpeta — no se omite ninguna, aunque su nombre no traiga número`;
  sheet.getCell("A4").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("A4").alignment = { horizontal: "center" };
  sheet.getRow(4).height = 14;

  sheet.addRow([]); // fila 5 en blanco

  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN", "N° DE ARCHIVADORES"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Construir el árbol (mismo agrupado que Separadores) y recorrerlo con
  // numeración propia, sin saltarse nada ---
  const arbol = construirArbolSeparadores(carpetasDelArea);
  const items = [];
  generarFilasGeneral(arbol, 1, "", items);

  for (const item of items) {
    const fila = sheet.addRow([item.numero, item.nombre, null]);
    fila.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    fila.getCell(1).numFmt = "@"; // forzar texto, para que "1.10" no se lea como número
    fila.getCell(2).alignment = { horizontal: "left", vertical: "middle", wrapText: true };

    if (item.nivel === 1) {
      fila.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: DORADO_FASE2 } };
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      });
    } else if (item.nivel === 2) {
      const fondo = colorNivelSeparador(item.nivel);
      fila.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      });
      if (!item.esUltimoNivel) fila.getCell(2).font = { bold: true };
    } else if (!item.esUltimoNivel) {
      fila.getCell(2).font = { bold: true };
    }

    if (item.esUltimoNivel) {
      fila.getCell(3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMARILLO_EDITABLE } };
    }

    fila.getCell(3).border = { top: BORDE_GRIS, bottom: BORDE_GRIS, left: BORDE_GRIS, right: BORDE_GRIS };

    fila.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        ...cell.border,
        top: cell.border?.top || { style: "hair", color: { argb: "FFDDDDDD" } },
        bottom: cell.border?.bottom || { style: "hair", color: { argb: "FFDDDDDD" } },
        left: cell.border?.left || { style: "hair", color: { argb: "FFDDDDDD" } },
        right: cell.border?.right || { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  sheet.headerFooter.oddFooter = "&L&I[Sistema ACASO I-2] Índice general&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Indice_General_${CODIGO_PROYECTO}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
