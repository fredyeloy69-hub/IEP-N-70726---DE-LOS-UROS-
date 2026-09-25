import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { LOGO_PUNO_BASE64 } from "./logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_RGB = {
  completa: [42, 157, 143],
  incompleta: [243, 156, 18],
  vacia: [192, 57, 43],
};

const TEAL_HEADER = [122, 31, 43];
const TEAL = [122, 31, 43];
const TEAL_CLARO = [70, 185, 175];
const VERDE = [42, 157, 143];
const NARANJA = [243, 156, 18];
const ROJO = [192, 57, 43];
const AZUL_PETROLEO = [74, 14, 23];

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

function contarEstadosPDF(items) {
  let completas = 0, incompletas = 0, vacias = 0;
  for (const c of items) {
    if (c.estado === "completa") completas++;
    else if (c.estado === "incompleta") incompletas++;
    else if (c.estado === "vacia") vacias++;
  }
  return { completas, incompletas, vacias };
}

function textoResumenPDF(items) {
  const { completas, incompletas, vacias } = contarEstadosPDF(items);
  const total = items.length;
  return `  (Total: ${total} — ${completas} completas · ${incompletas} incompletas · ${vacias} vacías)`;
}

function dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX) {
  let y = 12;
  if (logoBase64) {
    const anchoLogo = 13;
    const altoLogo = 14.6;
    doc.addImage(logoBase64, "PNG", marginX, y - 3, anchoLogo, altoLogo);
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA", pageWidth / 2, y, { align: "center" });
  y += 5;
  doc.text("SUB GERENCIA DE ESTUDIOS DEFINITIVOS", pageWidth / 2, y, { align: "center" });
  y += 6;
  doc.setFontSize(8);
  const lineasProyecto = doc.splitTextToSize(proyectoNombre, pageWidth - marginX * 2 - 20);
  doc.text(lineasProyecto, pageWidth / 2, y, { align: "center" });
  return y + (lineasProyecto.length * 3.8) + 4;
}

function dibujarBarra(doc, x, y, ancho, alto, pct, colorRgb, colorFondoRgb) {
  const fondo = colorFondoRgb || [20, 37, 41];
  doc.setFillColor(...fondo);
  doc.roundedRect(x, y, ancho, alto, alto / 2, alto / 2, "F");
  const anchoLleno = Math.max(alto, (ancho * Math.min(100, Math.max(0, pct))) / 100);
  doc.setFillColor(...colorRgb);
  doc.roundedRect(x, y, anchoLleno, alto, alto / 2, alto / 2, "F");
}

function dibujarTarjeta(doc, x, y, ancho, alto, valor, etiqueta, colorRgb) {
  doc.setDrawColor(...colorRgb);
  doc.setLineWidth(0.6);
  doc.roundedRect(x, y, ancho, alto, 2, 2, "S");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(20);
  doc.setTextColor(...colorRgb);
  doc.text(String(valor), x + ancho / 2, y + alto / 2 - 1, { align: "center" });
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.setTextColor(90, 90, 90);
  doc.text(etiqueta, x + ancho / 2, y + alto - 5, { align: "center" });
  doc.setTextColor(0, 0, 0);
}

// Función auxiliar para aplicar el pie de página institucional estandarizado en todas las páginas
function agregarPieDePaginaGlobal(doc, usuarioFirma, pageWidth, pageHeight, marginX) {
  const totalPaginas = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPaginas; p++) {
    doc.setPage(p);
    
    // Línea sutil de separación para el pie de página
    doc.setDrawColor(210, 210, 210);
    doc.setLineWidth(0.3);
    doc.line(marginX, pageHeight - 12, pageWidth - marginX, pageHeight - 12);

    doc.setFont("helvetica", "italic");
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(`Sistema de Control de Proyectos — Generado por: ${usuarioFirma}`, marginX, pageHeight - 8);
    doc.text(`Página ${p} de ${totalPaginas}`, pageWidth - marginX, pageHeight - 8, { align: "right" });
    doc.setTextColor(0, 0, 0);
  }
}

export function generarReportePorArea(areaNombre, carpetasDelArea, opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN ACCASO DISTRITO DE PILCUYO DE LA PROVINCIA DE EL COLLAO DEL DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Accaso I-2",
    tipoFiltro = "todas",
  } = opts;

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  const anchoUtilTabla = pageWidth - marginX * 2;

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
  ordenGrupos.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;
  let necesarios = 0;
  let completados = 0;
  for (const c of carpetasDelArea) {
    necesarios += c.archivosNecesarios || 0;
    completados += c.archivosCompletados || 0;
  }
  const pctArchivos = necesarios > 0 ? Math.round((completados / necesarios) * 100) : 0;

  const body = [];
  let contadorFila = 1;

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  function mezclarConBlanco(rgb, factor) {
    return rgb.map((c) => Math.round(c + (255 - c) * factor));
  }

  const FACTORES_HOJA = [0.72, 0.78, 0.83, 0.87, 0.90, 0.93, 0.96];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(TEAL_HEADER, factor);
  }

  function empujarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const sangriaIzq = 2 + Math.max(0, nivelVisual - 1) * 4.5;
    const fondo = colorHoja(nivelVisual);
    body.push([
      { content: String(contadorFila++), styles: { fillColor: fondo, halign: "center" } },
      { content: nombreMostrado, styles: { fillColor: fondo, cellPadding: { top: 2, right: 2, bottom: 2, left: sangriaIzq } } },
      {
        content: ESTADO_LABEL[estado] || estado.toUpperCase(),
        styles: { fillColor: fondo, textColor: ESTADO_RGB[estado] || [100, 100, 100], fontStyle: "bold", halign: "center" },
      },
      { content: c.detalle || "-", styles: { fillColor: fondo } },
    ]);
  }

  function empujarSubEncabezado(nombre, nivelVisual, items) {
    // Misma escala que colorHoja() para las filas de datos — así una carpeta
    // que termina siendo encabezado de grupo y una que termina siendo fila
    // plana, al mismo nivel de profundidad real, quedan del mismo color.
    const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
    const color = mezclarConBlanco(TEAL_HEADER, factor);
    const sangriaIzq = 2 + (nivelVisual - 1) * 4.5;
    const resumen = items && nivelVisual <= 2 ? textoResumenPDF(items) : "";
    const textoSub = `>  ${nombre}${resumen}`;
    const lineasSub = doc.splitTextToSize(textoSub, anchoUtilTabla - sangriaIzq - 4);

    body.push([
      {
        content: lineasSub.join("\n"),
        colSpan: 4,
        styles: {
          fillColor: color,
          textColor: [26, 56, 64],
          fontStyle: "bold",
          fontSize: 9.0,
          halign: "left",
          overflow: "linebreak",
          cellPadding: { top: 2.5, right: 2, bottom: 2.5, left: sangriaIzq },
        },
      },
    ]);
  }

  const NIVEL_MAX_INTERMEDIO = 6;
  function agruparRecursivo(items, nivelIdx, nivelVisual) {
    if (nivelIdx > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length > nivelIdx + 1 ? partes[nivelIdx] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
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
        empujarSubEncabezado(entrada.key, nivelVisual, subgrupos[entrada.key]);
        agruparRecursivo(subgrupos[entrada.key], nivelIdx + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => empujarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    const textoEsp = `${especialidad.toUpperCase()}${textoResumenPDF(grupos[especialidad])}`;
    const lineasEsp = doc.splitTextToSize(textoEsp, anchoUtilTabla - 6);
    body.push([
      { 
        content: lineasEsp.join("\n"), 
        colSpan: 4, 
        styles: { 
          fillColor: TEAL_HEADER, 
          textColor: [255, 255, 255], 
          fontStyle: "bold", 
          fontSize: 10.5, 
          halign: "left", 
          overflow: "linebreak",
          cellPadding: 3 
        } 
      },
    ]);
    agruparRecursivo(grupos[especialidad], 2, 1);
  }

  function dibujarEncabezado() {
    let y = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);
    doc.setFillColor(...AZUL_PETROLEO);
    doc.rect(marginX, y, pageWidth - marginX * 2, 9, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10.5);
    doc.text(`REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, y + 6.0, { align: "center" });
    doc.setTextColor(0, 0, 0);
    y += 15;
    const fecha = new Date().toLocaleString("es-PE");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text(`Generado: ${fecha}`, marginX, y);
    doc.text(`Total: ${total}  ·  Completas: ${completas}  ·  Pendientes: ${total - completas}`, pageWidth - marginX, y, { align: "right" });
    return y + 8;
  }

  // PORTADA — resumen gráfico
  let y = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);
  doc.setFillColor(...AZUL_PETROLEO);
  doc.rect(marginX, y, pageWidth - marginX * 2, 11, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text(`RESUMEN DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, y + 7.5, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 18;

  const fechaGenerado = new Date().toLocaleString("es-PE");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(100, 100, 100);
  doc.text(`Generado: ${fechaGenerado}`, pageWidth / 2, y, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 10;

  const pctCarpetas = total > 0 ? Math.round((completas / total) * 100) : 0;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE POR CARPETAS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL);
  doc.text(`${pctCarpetas}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 4;
  dibujarBarra(doc, marginX, y, pageWidth - marginX * 2, 6, pctCarpetas, TEAL);
  y += 12;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE POR ARCHIVOS (más preciso)", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL_CLARO);
  doc.text(`${pctArchivos}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 4;
  dibujarBarra(doc, marginX, y, pageWidth - marginX * 2, 6, pctArchivos, TEAL_CLARO);
  y += 16;

  const anchoTarjeta = (pageWidth - marginX * 2 - 3 * 4) / 4;
  const altoTarjeta = 22;
  dibujarTarjeta(doc, marginX, y, anchoTarjeta, altoTarjeta, total, "CARPETAS", TEAL);
  dibujarTarjeta(doc, marginX + (anchoTarjeta + 4) * 1, y, anchoTarjeta, altoTarjeta, completas, "COMPLETAS", VERDE);
  dibujarTarjeta(doc, marginX + (anchoTarjeta + 4) * 2, y, anchoTarjeta, altoTarjeta, incompletas, "INCOMPLETAS", NARANJA);
  dibujarTarjeta(doc, marginX + (anchoTarjeta + 4) * 3, y, anchoTarjeta, altoTarjeta, vacias, "VACÍAS", ROJO);
  y += altoTarjeta + 12;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(`AVANCE POR ESPECIALIDAD (${ordenGrupos.length})`, marginX, y);
  y += 7;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  for (const especialidad of ordenGrupos) {
    if (y > pageHeight - 20) {
      doc.addPage();
      y = 18;
    }
    const items = grupos[especialidad];
    let espNecesarios = 0;
    let espCompletados = 0;
    for (const c of items) {
      espNecesarios += c.archivosNecesarios || 0;
      espCompletados += c.archivosCompletados || 0;
    }
    const pctEsp = espNecesarios > 0 ? Math.round((espCompletados / espNecesarios) * 100) : 0;
    const colorEsp = pctEsp >= 100 ? VERDE : pctEsp >= 50 ? TEAL_CLARO : NARANJA;

    const etiqueta = doc.splitTextToSize(especialidad, 70)[0];
    doc.setTextColor(60, 60, 60);
    doc.text(etiqueta, marginX, y + 3);
    doc.setTextColor(0, 0, 0);

    const barraX = marginX + 74;
    const barraAncho = pageWidth - marginX - 22 - barraX;
    dibujarBarra(doc, barraX, y, barraAncho, 4.5, pctEsp, colorEsp);

    doc.setFont("helvetica", "bold");
    doc.setTextColor(...colorEsp);
    doc.text(`${pctEsp}%`, pageWidth - marginX, y + 3.3, { align: "right" });
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "normal");
    y += 8;
  }

  doc.addPage();
  const primeraPaginaTabla = doc.internal.getCurrentPageInfo().pageNumber;
  const startY = dibujarEncabezado();

  autoTable(doc, {
    startY,
    margin: { left: marginX, right: marginX, top: startY, bottom: 18 },
    head: [["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"]],
    body,
    styles: { fontSize: 8.5, cellPadding: 2, valign: "middle", overflow: "linebreak" },
    headStyles: { fillColor: AZUL_PETROLEO, textColor: 255, fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 12, halign: "center" },
      1: { cellWidth: 89 },
      2: { cellWidth: 26, halign: "center" },
      3: { cellWidth: 55 },
    },
    didDrawPage: () => {
      if (doc.internal.getCurrentPageInfo().pageNumber > primeraPaginaTabla) {
        dibujarEncabezado();
      }
    },
  });

  agregarPieDePaginaGlobal(doc, usuarioFirma, pageWidth, pageHeight, marginX);

  const nombreArchivo = `Reporte_${CODIGO_PROYECTO}_${sufijoFiltroArchivo(tipoFiltro)}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date().toISOString().slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);
}

// ============================================================
// NUEVA FUNCIÓN: REPORTE POR ESTADO ESPECÍFICO (Completas, Incompletas o Vacías)
// ============================================================
export function generarReportePorEstadoEspecifico(estadoObjetivo, carpetasDelArea, areaNombre = "General", opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN ACCASO DISTRITO DE PILCUYO DE LA PROVINCIA DE EL COLLAO DEL DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Accaso I-2",
  } = opts;

  const elementosFiltrados = carpetasDelArea.filter(c => (c.estado || "incompleta").toLowerCase() === estadoObjetivo.toLowerCase());

  if (elementosFiltrados.length === 0) {
    alert(`No hay registros con el estado "${ESTADO_LABEL[estadoObjetivo] || estadoObjetivo.toUpperCase()}" para exportar.`);
    return;
  }

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  const anchoUtilTabla = pageWidth - marginX * 2;

  const total = elementosFiltrados.length;
  const etiquetaEstadoTexto = ESTADO_LABEL[estadoObjetivo] || estadoObjetivo.toUpperCase();

  const body = [];
  let contadorFila = 1;

  elementosFiltrados.forEach((c) => {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const areaTexto = c.area || areaNombre;
    
    body.push([
      { content: String(contadorFila++), styles: { halign: "center" } },
      { content: areaTexto },
      { content: nombreMostrado },
      {
        content: ESTADO_LABEL[estado] || estado.toUpperCase(),
        styles: { textColor: ESTADO_RGB[estado] || [100, 100, 100], fontStyle: "bold", halign: "center" },
      },
      { content: c.detalle || "-" },
    ]);
  });

  function dibujarEncabezadoEstado() {
    let y = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);
    doc.setFillColor(...AZUL_PETROLEO);
    doc.rect(marginX, y, pageWidth - marginX * 2, 9, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10.5);
    doc.text(`REPORTE DE CARPETAS: ${etiquetaEstadoTexto} (${areaNombre.toUpperCase()})`, pageWidth / 2, y + 6.0, { align: "center" });
    doc.setTextColor(0, 0, 0);
    y += 15;
    const fecha = new Date().toLocaleString("es-PE");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text(`Generado: ${fecha}`, marginX, y);
    doc.text(`Total elementos: ${total}`, pageWidth - marginX, y, { align: "right" });
    return y + 8;
  }

  // --- PORTADA ESPECÍFICA PARA EL ESTADO ---
  let y = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);
  doc.setFillColor(...AZUL_PETROLEO);
  doc.rect(marginX, y, pageWidth - marginX * 2, 11, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text(`REPORTE EXCLUSIVO — ${etiquetaEstadoTexto}`, pageWidth / 2, y + 7.5, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 18;

  const fechaGenerado = new Date().toLocaleString("es-PE");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(100, 100, 100);
  doc.text(`Área: ${areaNombre} | Filtro de Estado: ${etiquetaEstadoTexto} | Generado: ${fechaGenerado}`, pageWidth / 2, y, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 14;

  const colorTarjeta = ESTADO_RGB[estadoObjetivo] || TEAL;
  const anchoTarjeta = 80;
  const altoTarjeta = 26;
  const xTarjeta = (pageWidth - anchoTarjeta) / 2;
  
  dibujarTarjeta(doc, xTarjeta, y, anchoTarjeta, altoTarjeta, total, `TOTAL DE CARPETAS ${etiquetaEstadoTexto}`, colorTarjeta);
  y += altoTarjeta + 20;

  doc.addPage();
  const primeraPaginaTabla = doc.internal.getCurrentPageInfo().pageNumber;
  const startY = dibujarEncabezadoEstado();

  autoTable(doc, {
    startY,
    margin: { left: marginX, right: marginX, top: startY, bottom: 18 },
    head: [["N°", "ÁREA", "DESCRIPCIÓN / CARPETA", "ESTADO", "DETALLE"]],
    body,
    styles: { fontSize: 8, cellPadding: 2, valign: "middle", overflow: "linebreak" },
    headStyles: { fillColor: AZUL_PETROLEO, textColor: 255, fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 10, halign: "center" },
      1: { cellWidth: 38 },
      2: { cellWidth: 55 },
      3: { cellWidth: 24, halign: "center" },
      4: { cellWidth: 55 },
    },
    didDrawPage: () => {
      if (doc.internal.getCurrentPageInfo().pageNumber > primeraPaginaTabla) {
        dibujarEncabezadoEstado();
      }
    },
  });

  agregarPieDePaginaGlobal(doc, usuarioFirma, pageWidth, pageHeight, marginX);

  const nombreArchivo = `Reporte_${CODIGO_PROYECTO}_${estadoObjetivo}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date().toISOString().slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);
}

// ============================================================
// REPORTE CONSOLIDADO GLOBAL — CADA ÁREA CON SU PROPIA PORTADA Y GRÁFICOS
// ============================================================
export function generarReporteConsolidadoGlobal(datosEntrada, opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN ACCASO DISTRITO DE PILCUYO DE LA PROVINCIA DE EL COLLAO DEL DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Accaso I-2",
    tipoFiltro = "todas",
  } = opts;

  let todasLasCarpetas = [];
  if (Array.isArray(datosEntrada)) {
    if (datosEntrada.length > 0 && (datosEntrada[0].carpetas || datosEntrada[0].items)) {
      for (const areaObj of datosEntrada) {
        const nombreArea = areaObj.nombre || areaObj.area || "Sin área";
        const subCarpetas = areaObj.carpetas || areaObj.items || [];
        for (const c of subCarpetas) {
          todasLasCarpetas.push({ ...c, area: c.area || nombreArea });
        }
      }
    } else {
      todasLasCarpetas = datosEntrada.map(c => ({ ...c, area: c.area || (c.ruta ? c.ruta.split(" / ")[0] : "Sin área") }));
    }
  } else if (datosEntrada && typeof datosEntrada === "object") {
    const entries = Object.entries(datosEntrada);
    for (const [key, val] of entries) {
      if (Array.isArray(val)) {
        for (const c of val) {
          todasLasCarpetas.push({ ...c, area: c.area || key });
        }
      } else if (val && typeof val === "object") {
        const subArr = val.carpetas || val.items || Object.values(val).find(Array.isArray) || [];
        for (const c of subArr) {
          todasLasCarpetas.push({ ...c, area: c.area || key });
        }
      }
    }
  }

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  const anchoUtilTabla = pageWidth - marginX * 2;

  const areasMap = {};
  for (const c of todasLasCarpetas) {
    const area = c.area || "Sin área";
    if (!areasMap[area]) areasMap[area] = [];
    areasMap[area].push(c);
  }
  const ordenAreas = Object.keys(areasMap).sort((a, b) => a.localeCompare(b));

  const totalCarpetas = todasLasCarpetas.length;
  const totalCompletas = todasLasCarpetas.filter((c) => c.estado === "completa").length;
  let totalNec = 0;
  let totalComp = 0;
  for (const c of todasLasCarpetas) {
    totalNec += c.archivosNecesarios || 0;
    totalComp += c.archivosCompletados || 0;
  }
  const pctGlobalArchivos = totalNec > 0 ? Math.round((totalComp / totalNec) * 100) : 0;
  const pctCarpetasGlobal = totalCarpetas > 0 ? Math.round((totalCompletas / totalCarpetas) * 100) : 0;

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  function mezclarConBlanco(rgb, factor) {
    return rgb.map((c) => Math.round(c + (255 - c) * factor));
  }

  const FACTORES_HOJA = [0.72, 0.78, 0.83, 0.87, 0.90, 0.93, 0.96];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(TEAL_HEADER, factor);
  }

  // ==========================================
  // 1. PORTADA GLOBAL DEL PROYECTO
  // ==========================================
  let y = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);
  doc.setFillColor(...AZUL_PETROLEO);
  doc.rect(marginX, y, pageWidth - marginX * 2, 11, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("CONSOLIDADO GENERAL DE AVANCE DEL PROYECTO", pageWidth / 2, y + 7.5, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 18;

  const fechaGenerado = new Date().toLocaleString("es-PE");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(100, 100, 100);
  doc.text(`Generado: ${fechaGenerado}`, pageWidth / 2, y, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 12;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE GLOBAL POR CARPETAS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL);
  doc.text(`${pctCarpetasGlobal}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 6;
  dibujarBarra(doc, marginX, y, pageWidth - marginX * 2, 6, pctCarpetasGlobal, TEAL);
  y += 14;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE GLOBAL POR ARCHIVOS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL_CLARO);
  doc.text(`${pctGlobalArchivos}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 6;
  dibujarBarra(doc, marginX, y, pageWidth - marginX * 2, 6, pctGlobalArchivos, TEAL_CLARO);
  y += 18;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(`DESGLOSE POR ÁREAS (${ordenAreas.length})`, marginX, y);
  y += 8;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  for (const areaNombre of ordenAreas) {
    if (y > pageHeight - 25) {
      doc.addPage();
      y = 18;
    }
    const itemsArea = areasMap[areaNombre];
    let areaNec = 0;
    let areaComp = 0;
    for (const c of itemsArea) {
      areaNec += c.archivosNecesarios || 0;
      areaComp += c.archivosCompletados || 0;
    }
    const pctArea = areaNec > 0 ? Math.round((areaComp / areaNec) * 100) : 0;
    const colorArea = pctArea >= 100 ? VERDE : pctArea >= 50 ? TEAL_CLARO : NARANJA;

    const etiqueta = doc.splitTextToSize(areaNombre, 70)[0];
    doc.setTextColor(60, 60, 60);
    doc.text(etiqueta, marginX, y + 3);
    doc.setTextColor(0, 0, 0);

    const barraX = marginX + 74;
    const barraAncho = pageWidth - marginX - 22 - barraX;
    dibujarBarra(doc, barraX, y, barraAncho, 4.5, pctArea, colorArea);

    doc.setFont("helvetica", "bold");
    doc.setTextColor(...colorArea);
    doc.text(`${pctArea}%`, pageWidth - marginX, y + 3.3, { align: "right" });
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "normal");
    y += 8;
  }

  // ==========================================
  // 2. SECCIÓN INDIVIDUAL POR CADA ÁREA / CARPETA MADRE
  // ==========================================
  for (const areaNombre of ordenAreas) {
    const carpetasArea = areasMap[areaNombre];
    const totalArea = carpetasArea.length;
    const completasArea = carpetasArea.filter((c) => c.estado === "completa").length;
    const incompletasArea = carpetasArea.filter((c) => c.estado === "incompleta").length;
    const vaciasArea = carpetasArea.filter((c) => c.estado === "vacia").length;
    let necArea = 0;
    let compArea = 0;
    for (const c of carpetasArea) {
      necArea += c.archivosNecesarios || 0;
      compArea += c.archivosCompletados || 0;
    }
    const pctArchArea = necArea > 0 ? Math.round((compArea / necArea) * 100) : 0;
    const pctCarpArea = totalArea > 0 ? Math.round((completasArea / totalArea) * 100) : 0;

    doc.addPage();
    let yArea = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);

    doc.setFillColor(...AZUL_PETROLEO);
    doc.rect(marginX, yArea, pageWidth - marginX * 2, 11, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text(`RESUMEN DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, yArea + 7.5, { align: "center" });
    doc.setTextColor(0, 0, 0);
    yArea += 18;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(100, 100, 100);
    doc.text(`Generado: ${fechaGenerado}`, pageWidth / 2, yArea, { align: "center" });
    doc.setTextColor(0, 0, 0);
    yArea += 10;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("AVANCE POR CARPETAS", marginX, yArea);
    doc.setFontSize(16);
    doc.setTextColor(...TEAL);
    doc.text(`${pctCarpArea}%`, pageWidth - marginX, yArea, { align: "right" });
    doc.setTextColor(0, 0, 0);
    yArea += 4;
    dibujarBarra(doc, marginX, yArea, pageWidth - marginX * 2, 6, pctCarpArea, TEAL);
    yArea += 12;

    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("AVANCE POR ARCHIVOS (más preciso)", marginX, yArea);
    doc.setFontSize(16);
    doc.setTextColor(...TEAL_CLARO);
    doc.text(`${pctArchArea}%`, pageWidth - marginX, yArea, { align: "right" });
    doc.setTextColor(0, 0, 0);
    yArea += 4;
    dibujarBarra(doc, marginX, yArea, pageWidth - marginX * 2, 6, pctArchArea, TEAL_CLARO);
    yArea += 16;

    const anchoTarjeta = (pageWidth - marginX * 2 - 3 * 4) / 4;
    const altoTarjeta = 22;
    dibujarTarjeta(doc, marginX, yArea, anchoTarjeta, altoTarjeta, totalArea, "CARPETAS", TEAL);
    dibujarTarjeta(doc, marginX + (anchoTarjeta + 4) * 1, yArea, anchoTarjeta, altoTarjeta, completasArea, "COMPLETAS", VERDE);
    dibujarTarjeta(doc, marginX + (anchoTarjeta + 4) * 2, yArea, anchoTarjeta, altoTarjeta, incompletasArea, "INCOMPLETAS", NARANJA);
    dibujarTarjeta(doc, marginX + (anchoTarjeta + 4) * 3, yArea, anchoTarjeta, altoTarjeta, vaciasArea, "VACÍAS", ROJO);

    doc.addPage();
    
    function dibujarEncabezadoAreaTabla() {
      let yT = dibujarBloqueInstitucional(doc, logoBase64, proyectoNombre, pageWidth, marginX);
      doc.setFillColor(...AZUL_PETROLEO);
      doc.rect(marginX, yT, pageWidth - marginX * 2, 9, "F");
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(10.5);
      doc.text(`REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, yT + 6.0, { align: "center" });
      doc.setTextColor(0, 0, 0);
      yT += 15;
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      doc.text(`Generado: ${fechaGenerado}`, marginX, yT);
      doc.text(`Total: ${totalArea}  ·  Completas: ${completasArea}  ·  Pendientes: ${totalArea - completasArea}`, pageWidth - marginX, yT, { align: "right" });
      return yT + 8;
    }

    const startYAreaTabla = dibujarEncabezadoAreaTabla();

    const gruposEsp = {};
    const ordenEsp = [];
    for (const c of carpetasArea) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
      if (!gruposEsp[especialidad]) {
        gruposEsp[especialidad] = [];
        ordenEsp.push(especialidad);
      }
      gruposEsp[especialidad].push(c);
    }
    ordenEsp.sort((a, b) => comparaNatural(a, b));

    const bodyArea = [];
    let contadorFilaArea = 1;

    function empujarFilaArea(c, nivelVisual) {
      const estado = c.estado || "incompleta";
      const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
      const sangriaIzq = 2 + Math.max(0, nivelVisual - 1) * 4.5;
      const fondo = colorHoja(nivelVisual);
      bodyArea.push([
        { content: String(contadorFilaArea++), styles: { fillColor: fondo, halign: "center" } },
        { content: nombreMostrado, styles: { fillColor: fondo, cellPadding: { top: 2, right: 2, bottom: 2, left: sangriaIzq } } },
        {
          content: ESTADO_LABEL[estado] || estado.toUpperCase(),
          styles: { fillColor: fondo, textColor: ESTADO_RGB[estado] || [100, 100, 100], fontStyle: "bold", halign: "center" },
        },
        { content: c.detalle || "-", styles: { fillColor: fondo } },
      ]);
    }

    function empujarSubEncabezadoArea(nombre, nivelVisual, items) {
      // Misma escala que colorHoja() para las filas de datos.
      const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
      const color = mezclarConBlanco(TEAL_HEADER, factor);
      const sangriaIzq = 2 + (nivelVisual - 1) * 4.5;
      const resumen = items && nivelVisual <= 2 ? textoResumenPDF(items) : "";
      const textoSub = `>  ${nombre}${resumen}`;
      const lineasSub = doc.splitTextToSize(textoSub, anchoUtilTabla - sangriaIzq - 4);
      bodyArea.push([
        {
          content: lineasSub.join("\n"),
          colSpan: 4,
          styles: {
            fillColor: color,
            textColor: [26, 56, 64],
            fontStyle: "bold",
            fontSize: 9.0,
            halign: "left",
            overflow: "linebreak",
            cellPadding: { top: 2.5, right: 2, bottom: 2.5, left: sangriaIzq },
          },
        },
      ]);
    }

    function agruparRecursivoArea(items, nivelIdx, nivelVisual) {
      if (nivelIdx > 6) {
        const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => empujarFilaArea(c, nivelVisual));
        return;
      }
      const subgrupos = {};
      let hayNivelMasProfundo = false;
      for (const c of items) {
        const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
        const clave = partes.length > nivelIdx + 1 ? partes[nivelIdx] : null;
        if (clave) hayNivelMasProfundo = true;
        const key = clave || `__directo__${c.nombre || c.id}`;
        if (!subgrupos[key]) subgrupos[key] = [];
        subgrupos[key].push(c);
      }
      if (!hayNivelMasProfundo) {
        const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => empujarFilaArea(c, nivelVisual));
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
          empujarSubEncabezadoArea(entrada.key, nivelVisual, subgrupos[entrada.key]);
          agruparRecursivoArea(subgrupos[entrada.key], nivelIdx + 1, nivelVisual + 1);
        } else {
          const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
          ordenado.forEach((c) => empujarFilaArea(c, nivelVisual));
        }
      }
    }

    for (const esp of ordenEsp) {
      const textoEsp = `»  ${esp.toUpperCase()}${textoResumenPDF(gruposEsp[esp])}`;
      const lineasEsp = doc.splitTextToSize(textoEsp, anchoUtilTabla - 6);
      bodyArea.push([
        {
          content: lineasEsp.join("\n"),
          colSpan: 4,
          styles: {
            fillColor: TEAL_HEADER,
            textColor: [255, 255, 255],
            fontStyle: "bold",
            fontSize: 9.5,
            halign: "left",
            overflow: "linebreak",
            cellPadding: 2.5,
          },
        },
      ]);
      agruparRecursivoArea(gruposEsp[esp], 2, 1);
    }

    const primeraPaginaAreaTabla = doc.internal.getCurrentPageInfo().pageNumber;
    autoTable(doc, {
      startY: startYAreaTabla,
      margin: { left: marginX, right: marginX, top: startYAreaTabla, bottom: 18 },
      head: [["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"]],
      body: bodyArea,
      styles: { fontSize: 8.5, cellPadding: 2, valign: "middle", overflow: "linebreak" },
      headStyles: { fillColor: AZUL_PETROLEO, textColor: 255, fontStyle: "bold" },
      columnStyles: {
        0: { cellWidth: 12, halign: "center" },
        1: { cellWidth: 89 },
        2: { cellWidth: 26, halign: "center" },
        3: { cellWidth: 55 },
      },
      didDrawPage: () => {
        if (doc.internal.getCurrentPageInfo().pageNumber > primeraPaginaAreaTabla) {
          dibujarEncabezadoAreaTabla();
        }
      },
    });
  }

  agregarPieDePaginaGlobal(doc, usuarioFirma, pageWidth, pageHeight, marginX);

  const sufijoGlobal = tipoFiltro === "todas" ? "Completo" : sufijoFiltroArchivo(tipoFiltro);
  const nombreArchivo = `Reporte_${CODIGO_PROYECTO}_${sufijoGlobal}_Consolidado_${new Date().toISOString().slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);
}
