import ExcelJS from "exceljs";
import { LOGO_PUNO_BASE64 } from "./logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_COLOR = {
  completa: "FF2ECC71",
  incompleta: "FFF39C12",
  vacia: "FFE74C3C",
};

const AZUL_OSCURO = "FF141E3C";
const NARANJA = "FFE67E22";

function comparaNatural(a, b) {
  return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
}

function mezclarConBlanco(hexRgb, factor) {
  const r = parseInt(hexRgb.slice(2, 4), 16);
  const g = parseInt(hexRgb.slice(4, 6), 16);
  const b = parseInt(hexRgb.slice(6, 8), 16);
  const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
  return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
}

const FACTORES_HOJA = [0.72, 0.78, 0.83, 0.87, 0.90, 0.93, 0.96];
function colorHoja(nivelVisual) {
  const factor = FACTORES_HOJA[Math.min((nivelVisual || 1) - 1, FACTORES_HOJA.length - 1)];
  return mezclarConBlanco(NARANJA, factor);
}

function insertarLogoYEncabezados(sheet, workbook, logoBase64, proyectoNombre, subtituloTexto, usuarioFirma) {
  // Inserción del Logo
  if (logoBase64) {
    try {
      const base64Data = logoBase64.includes('base64,') ? logoBase64.split('base64,')[1] : logoBase64;
      const imageId = workbook.addImage({
        base64: base64Data,
        extension: 'png',
      });
      sheet.addImage(imageId, {
        tl: { col: 0.2, row: 0.2 },
        ext: { width: 70, height: 70 }
      });
    } catch (e) {
      console.warn("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 9.5 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 15;

  sheet.mergeCells("A3:D3");
  const celdaProyecto = sheet.getCell("A3");
  celdaProyecto.value = `PROYECTO: ${proyectoNombre.toUpperCase()}`;
  celdaProyecto.font = { bold: true, size: 9.5, color: { argb: "FF4A2D0F" } };
  celdaProyecto.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getRow(3).height = 28; 

  sheet.mergeCells("A4:D4");
  const celdaTitulo = sheet.getCell("A4");
  celdaTitulo.value = subtituloTexto;
  celdaTitulo.font = { bold: true, size: 11, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
  sheet.getRow(4).height = 22;

  sheet.mergeCells("A5:B5");
  sheet.getCell("A5").value = `Generado por: ${usuarioFirma} — ${new Date().toLocaleString("es-PE")}`;
  sheet.getCell("A5").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
}

/**
 * Genera y descarga un reporte en Excel (.xlsx) para un área específica.
 */
export async function generarReporteExcelPorArea(
  areaNombre,
  carpetasDelArea,
  opts = {}
) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Reporte";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { 
      orientation: "portrait", 
      fitToPage: true, 
      fitToWidth: 1,
      margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
    },
  });

  sheet.columns = [
    { width: 6 },
    { width: 55 },
    { width: 16 },
    { width: 55 },
  ];

  insertarLogoYEncabezados(sheet, workbook, logoBase64, proyectoNombre, `REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`, usuarioFirma);

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

  sheet.mergeCells("C5:D5");
  sheet.getCell("C5").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
  sheet.getCell("C5").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("C5").alignment = { horizontal: "right" };

  sheet.addRow(); // fila 6 en blanco

  const FILA_ENCABEZADO = 7;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN DE CARPETA", "ESTADO", "DETALLE / OBSERVACIÓN"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

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
  ordenGrupos.sort(comparaNatural);

  let contadorFila = 1;

  function agregarSubEncabezado(nombre, nivel) {
    const fila = sheet.addRow([`>  ${nombre}`]);
    sheet.mergeCells(`A${fila.number}:D${fila.number}`);
    const celda = fila.getCell(1);
    const esNivel1 = nivel === 1;
    const factoresPorNivel = [0, 0.40, 0.55, 0.68, 0.78, 0.86, 0.92]; 
    const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
    celda.font = { bold: true, size: esNivel1 ? 11 : 9.5, color: { argb: esNivel1 ? "FFFFFFFF" : "FF4A2D0F" } };
    celda.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: mezclarConBlanco(NARANJA, factor) },
    };
    celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
    fila.height = esNivel1 ? 19 : 16;
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

  const NIVEL_MAX_INTERMEDIO = 6;
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
      const clave = partes.length >= nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
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
        agregarSubEncabezado(entrada.key, nivelVisual);
        agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    agregarSubEncabezado(especialidad.toUpperCase(), 1);
    agruparRecursivo(grupos[especialidad], 2, 2);
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Genera y descarga un reporte Excel Consolidado Global con pestañas por área.
 */
export async function generarReporteExcelConsolidado(
  todasLasCarpetas,
  opts = {}
) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  // Agrupar por área
  const areasMap = {};
  for (const c of todasLasCarpetas) {
    const area = c.area || "Sin área";
    if (!areasMap[area]) areasMap[area] = [];
    areasMap[area].push(c);
  }
  const ordenAreas = Object.keys(areasMap).sort((a, b) => a.localeCompare(b));

  for (const areaNombre of ordenAreas) {
    const carpetasDelArea = areasMap[areaNombre];
    const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Área";
    const sheet = workbook.addWorksheet(nombreHoja, {
      pageSetup: { 
        orientation: "portrait", 
        fitToPage: true, 
        fitToWidth: 1,
        margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
      },
    });

    sheet.columns = [
      { width: 6 },
      { width: 55 },
      { width: 16 },
      { width: 55 },
    ];

    insertarLogoYEncabezados(sheet, workbook, logoBase64, proyectoNombre, `CONSOLIDADO — ÁREA: ${areaNombre.toUpperCase()}`, usuarioFirma);

    const total = carpetasDelArea.length;
    const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
    const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
    const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

    sheet.mergeCells("C5:D5");
    sheet.getCell("C5").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
    sheet.getCell("C5").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
    sheet.getCell("C5").alignment = { horizontal: "right" };

    sheet.addRow(); // fila 6 en blanco

    const FILA_ENCABEZADO = 7;
    const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
    filaEncabezado.values = ["N°", "DESCRIPCIÓN DE CARPETA", "ESTADO", "DETALLE / OBSERVACIÓN"];
    filaEncabezado.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
    });
    filaEncabezado.height = 18;

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
    ordenGrupos.sort(comparaNatural);

    let contadorFila = 1;

    function agregarSubEncabezado(nombre, nivel) {
      const fila = sheet.addRow([`>  ${nombre}`]);
      sheet.mergeCells(`A${fila.number}:D${fila.number}`);
      const celda = fila.getCell(1);
      const esNivel1 = nivel === 1;
      const factoresPorNivel = [0, 0.40, 0.55, 0.68, 0.78, 0.86, 0.92]; 
      const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
      celda.font = { bold: true, size: esNivel1 ? 11 : 9.5, color: { argb: esNivel1 ? "FFFFFFFF" : "FF4A2D0F" } };
      celda.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: mezclarConBlanco(NARANJA, factor) },
      };
      celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
      fila.height = esNivel1 ? 19 : 16;
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

    const NIVEL_MAX_INTERMEDIO = 6;
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
        const clave = partes.length >= nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
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
          agregarSubEncabezado(entrada.key, nivelVisual);
          agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
        } else {
          const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
          ordenado.forEach((c) => agregarFila(c, nivelVisual));
        }
      }
    }

    for (const especialidad of ordenGrupos) {
      agregarSubEncabezado(especialidad.toUpperCase(), 1);
      agruparRecursivo(grupos[especialidad], 2, 2);
    }

    sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_Consolidado_Global_${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
