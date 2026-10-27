// Clasifica el tamaño REAL de una hoja de PDF en formato A4/A3/A2/A1/A0.
// No depende del nombre del archivo: lee el tamaño físico de cada página.
//
// REGLA:
//  - Se compara la hoja (lado corto × lado largo, en mm) con las medidas
//    estándar de abajo y se elige la MÁS CERCANA.
//  - Tolerancia: ±20 mm en cada lado — las medidas reales en Drive varían entre 5 y 20 mm (cambiar TOLERANCIA_MM si hace falta).
//    Si la hoja está dentro de esa tolerancia de una medida estándar, queda en
//    ese formato sin avisar, tanto si es un poco más grande como más chica.
//  - También se acepta sin avisar una hoja A reducida/ampliada de forma pareja
//    (misma proporción 1 : 1.414, hasta ≈ ±15 % de escala).
//  - Si se pasa de eso, igual se cuenta en el formato más cercano, pero se
//    avisa en la hoja "Observaciones" para que se revise.
//  - Medidas menores o de otros estándares (A5, A6, A7, DL, Carta) cuentan
//    como A4, que es la hoja más pequeña que se imprime.

// Tolerancia por lado, en milímetros
export const TOLERANCIA_MM = 20;

// [nombre, lado corto (mm), lado largo (mm), formato en que se cuenta]
const MEDIDAS_REFERENCIA = [
  ["A0", 841, 1189, "A0"],
  ["A1", 594, 841, "A1"],
  ["A2", 420, 594, "A2"],
  ["A3", 297, 420, "A3"],
  ["A4", 210, 297, "A4"],
  ["A5", 148, 210, "A4"],
  ["A6", 105, 148, "A4"],
  ["A7", 74, 105, "A4"],
  ["DL", 110, 220, "A4"],
  ["Carta", 215.9, 279.4, "A4"],
];

// Se mantiene por compatibilidad con otras partes del código
export const FORMATOS_ISO = MEDIDAS_REFERENCIA.slice(0, 5).map(([n, c, l]) => [n, c, l]);

const PT_A_MM = 25.4 / 72; // 1 punto PDF = 1/72 de pulgada

/**
 * @param {number} anchoPt - ancho de la página en puntos PDF
 * @param {number} altoPt - alto de la página en puntos PDF
 * @returns {{formato: string, exacto: boolean, mm: [number, number]}}
 *   `exacto: true` = dentro de la tolerancia, no hace falta avisar.
 *   Funciona igual para hojas verticales u horizontales.
 */
export function clasificarHoja(anchoPt, altoPt) {
  return clasificarHojaMm(anchoPt * PT_A_MM, altoPt * PT_A_MM);
}

/** Igual que clasificarHoja pero recibiendo las medidas directamente en milímetros. */
export function clasificarHojaMm(anchoMm, altoMm) {
  const corto = Math.min(anchoMm, altoMm);
  const largo = Math.max(anchoMm, altoMm);
  const mm = [Math.round(corto), Math.round(largo)];

  let mejor = null;
  for (const ref of MEDIDAS_REFERENCIA) {
    const dc = corto - ref[1];
    const dl = largo - ref[2];
    const distancia = Math.sqrt(dc * dc + dl * dl);
    if (!mejor || distancia < mejor.distancia) {
      mejor = { ref, distancia, maxLado: Math.max(Math.abs(dc), Math.abs(dl)) };
    }
  }

  const formato = mejor.ref[3];

  // Hoja de la serie A REDUCIDA o AMPLIADA de forma pareja (muy común al
  // exportar o "ajustar a la página": un A3 que sale de 262×371 o de 313×443).
  // Conserva la proporción de hoja A (1 : 1.414) y su tamaño está cerca del
  // formato → se cuenta en ese formato sin avisar.
  const refFormato = MEDIDAS_REFERENCIA.find((r) => r[0] === formato);
  const proporcionA = Math.abs(largo / corto - Math.SQRT2) / Math.SQRT2 <= 0.04;
  const razonArea = (corto * largo) / (refFormato[1] * refFormato[2]);
  const escalaCercana = razonArea >= 0.75 && razonArea <= 1.33; // ≈ ±15 % de escala lineal

  return {
    formato,
    exacto: mejor.maxLado <= TOLERANCIA_MM || (proporcionA && escalaCercana),
    mm,
  };
}
