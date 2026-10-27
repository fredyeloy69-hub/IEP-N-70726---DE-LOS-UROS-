// Clasifica el tamaño REAL de una hoja de PDF en formato A4/A3/A2/A1/A0.
// No depende del nombre del archivo: lee el tamaño físico de cada página.

// [formato, lado corto (mm), lado largo (mm)] — medidas ISO 216 (serie A)
export const FORMATOS_ISO = [
  ["A4", 210, 297],
  ["A3", 297, 420],
  ["A2", 420, 594],
  ["A1", 594, 841],
  ["A0", 841, 1189],
];

const PT_A_MM = 25.4 / 72; // 1 punto PDF = 1/72 de pulgada
const TOL_EXACTO = 0.025; // ±2.5 % en ambos lados = formato exacto
const TOL_CABE = 0.03; // 3 % de holgura al decidir en qué hoja "cabe"

/**
 * @param {number} anchoPt - ancho de la página en puntos PDF
 * @param {number} altoPt - alto de la página en puntos PDF
 * @returns {{formato: string, exacto: boolean, mm: [number, number]}}
 *   Funciona igual para hojas verticales u horizontales (se ordenan los lados).
 *   Si el tamaño no es de la serie A (ej. Carta, 600×900), se asigna la hoja A
 *   MÁS PEQUEÑA en la que cabe — que es la que realmente hay que comprar — y
 *   se marca `exacto: false` para dejarlo avisado en "Observaciones".
 */
export function clasificarHoja(anchoPt, altoPt) {
  const w = anchoPt * PT_A_MM;
  const h = altoPt * PT_A_MM;
  const corto = Math.min(w, h);
  const largo = Math.max(w, h);
  const mm = [Math.round(corto), Math.round(largo)];

  for (const [formato, c, l] of FORMATOS_ISO) {
    if (Math.abs(corto - c) / c <= TOL_EXACTO && Math.abs(largo - l) / l <= TOL_EXACTO) {
      return { formato, exacto: true, mm };
    }
  }
  for (const [formato, c, l] of FORMATOS_ISO) {
    if (corto <= c * (1 + TOL_CABE) && largo <= l * (1 + TOL_CABE)) {
      return { formato, exacto: false, mm };
    }
  }
  return { formato: "A0", exacto: false, mm }; // más grande que A0 (rollos, etc.)
}
