// Paleta de colores de la interfaz, personalizable por cada usuario.
//
// Cómo funciona: toda la pantalla usa la familia "indigo" de Tailwind como color de acento (botones, menú activo,
// enlaces, fondos suaves...), y en Tailwind 4 esos colores son variables CSS (--color-indigo-50 ... --color-indigo-950).
// Cambiar la paleta es reemplazar esas 11 variables en <html>: toda la interfaz cambia de una vez, sin tocar cada pantalla.
//
// Qué NO cambia: los colores de estado (verde = bien, ámbar = atención, rojo = error) y los grises, ni los DOCUMENTOS que genera
// la app (PDF de órdenes y expedientes, orden impresa, correos): ver estiloDocumento().

export const PASOS = [50, 100, 200, 300, 400, 500, 600, 700, 800, 900, 950]
const nombreVar = (paso) => `--color-indigo-${paso}`

// ---------------------------------------------------------------- matemática de color (OKLab / sRGB) ----------
const aLineal = (c) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4)
const aGamma = (c) => (c <= 0.0031308 ? 12.92 * c : 1.055 * c ** (1 / 2.4) - 0.055)

export function hexARgbLineal(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim())
  if (!m) return null
  const n = parseInt(m[1], 16)
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => aLineal(v / 255))
}
export function rgbLinealAHex(rgb) {
  return '#' + rgb.map((v) => Math.round(Math.min(1, Math.max(0, aGamma(Math.min(1, Math.max(0, v))))) * 255).toString(16).padStart(2, '0')).join('')
}
export function rgbLinealAOklab([r, g, b]) {
  const l = Math.cbrt(0.4122214708 * r + 0.5363325363 * g + 0.0514459929 * b)
  const m = Math.cbrt(0.2119034982 * r + 0.6806995451 * g + 0.1073969566 * b)
  const s = Math.cbrt(0.0883024619 * r + 0.2817188376 * g + 0.6299787005 * b)
  return [0.2104542553 * l + 0.793617785 * m - 0.0040720468 * s, 1.9779984951 * l - 2.428592205 * m + 0.4505937099 * s, 0.0259040371 * l + 0.7827717662 * m - 0.808675766 * s]
}
export function oklchARgbLineal(L, C, h) {
  const a = C * Math.cos((h * Math.PI) / 180), b = C * Math.sin((h * Math.PI) / 180)
  const l = (L + 0.3963377774 * a + 0.2158037573 * b) ** 3
  const m = (L - 0.1055613458 * a - 0.0638541728 * b) ** 3
  const s = (L - 0.0894841775 * a - 1.291485548 * b) ** 3
  return [4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s, -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s, -0.0041960863 * l - 0.7034186147 * m + 1.707614701 * s]
}
export function hexAOklch(hex) {
  const rgb = hexARgbLineal(hex)
  if (!rgb) return null
  const [L, a, b] = rgbLinealAOklab(rgb)
  return { L, C: Math.hypot(a, b), h: (((Math.atan2(b, a) * 180) / Math.PI) + 360) % 360 }
}
const enGamut = (rgb) => rgb.every((v) => v >= -0.0005 && v <= 1.0005)
// la mayor saturación que se puede mostrar en una pantalla normal (sRGB) con esa luminosidad y tono
export function croma_maximo(L, h) {
  let lo = 0, hi = 0.4
  for (let i = 0; i < 22; i++) { const mid = (lo + hi) / 2; if (enGamut(oklchARgbLineal(L, mid, h))) lo = mid; else hi = mid }
  return lo
}
export const luminancia = ([r, g, b]) => 0.2126 * Math.min(1, Math.max(0, r)) + 0.7152 * Math.min(1, Math.max(0, g)) + 0.0722 * Math.min(1, Math.max(0, b))
export const contraste = (y1, y2) => (Math.max(y1, y2) + 0.05) / (Math.min(y1, y2) + 0.05)

// ---------------------------------------------------------------- generación de la escala (50 ... 950) ----------
// Parte de un TONO (0–360) y una INTENSIDAD (0–1). La luminosidad de cada paso es fija, salvo el 600 (el del botón principal), que se
// busca para que el texto blanco encima se lea bien (contraste mínimo 5:1, por encima del 4,5:1 de la norma de accesibilidad WCAG AA).
const L_FIJOS = { 50: 0.975, 100: 0.945, 200: 0.895, 300: 0.82 }
const FACTOR_CROMA = { 50: 0.17, 100: 0.3, 200: 0.5, 300: 0.7, 400: 0.88, 500: 0.97, 600: 1, 700: 0.96, 800: 0.85, 900: 0.7, 950: 0.5 }
export const CONTRASTE_MINIMO_BOTON = 5

export function generarEscala({ h, s }) {
  const tono = ((Number(h) % 360) + 360) % 360
  const intensidad = Math.min(1, Math.max(0.05, Number(s)))
  const color = (L, paso) => oklchARgbLineal(L, croma_maximo(L, tono) * intensidad * FACTOR_CROMA[paso], tono)
  // paso 600: el más claro que todavía cumple el contraste con blanco
  let L600 = 0.36
  for (let L = 0.62; L >= 0.36; L -= 0.005) { if (contraste(1, luminancia(color(L, 600))) >= CONTRASTE_MINIMO_BOTON) { L600 = L; break } }
  const L = { ...L_FIJOS, 600: L600, 500: Math.min(L600 + 0.07, 0.66), 700: L600 - 0.06, 800: L600 - 0.115, 900: L600 - 0.165, 950: Math.max(0.2, L600 - 0.25) }
  L[400] = Math.max(0.7, L[500] + 0.07)
  const escala = {}
  PASOS.forEach((p) => {
    const rgb = color(L[p], p)
    escala[p] = { oklch: `oklch(${L[p].toFixed(3)} ${Math.hypot(...rgbLinealAOklab(rgb).slice(1)).toFixed(3)} ${tono.toFixed(1)})`, rgb, hex: rgbLinealAHex(rgb), L: L[p] }
  })
  return escala
}

// ---------------------------------------------------------------- paletas disponibles ----------
// "indigo" es el diseño original: no se cambia nada. Las demás se generan a partir de su tono.
export const PALETAS = [
  { id: 'indigo', nombre: 'Índigo (original)', original: true },
  { id: 'azul', nombre: 'Azul', h: 258, s: 0.95 },
  { id: 'cielo', nombre: 'Cielo', h: 235, s: 0.95 },
  { id: 'turquesa', nombre: 'Turquesa', h: 192, s: 0.95 },
  { id: 'esmeralda', nombre: 'Esmeralda', h: 160, s: 0.95 },
  { id: 'violeta', nombre: 'Violeta', h: 302, s: 0.95 },
  { id: 'rosa', nombre: 'Rosa', h: 352, s: 0.95 },
  { id: 'naranja', nombre: 'Naranja', h: 48, s: 1 },
  { id: 'grafito', nombre: 'Grafito', h: 265, s: 0.1 },
]
export const TEMA_ORIGINAL = { paleta: 'indigo' }
const HEX = /^#[0-9a-f]{6}$/i

// Valida lo que llega de la cuenta del usuario (puede venir vacío, viejo o alterado): devuelve un tema válido o null.
export function sanitizarTema(crudo) {
  if (!crudo || typeof crudo !== 'object') return null
  if (crudo.paleta === 'personalizado') return HEX.test(String(crudo.color || '')) ? { paleta: 'personalizado', color: String(crudo.color).toLowerCase() } : null
  return PALETAS.some((p) => p.id === crudo.paleta) ? { paleta: crudo.paleta } : null
}
export function parametrosDe(tema) {
  const t = sanitizarTema(tema)
  if (!t || t.paleta === 'indigo') return null
  if (t.paleta === 'personalizado') {
    const c = hexAOklch(t.color)
    // un gris (sin tono propio) se trata como "grafito"
    if (c.C < 0.012) return { h: 265, s: 0.1 }
    const tope = croma_maximo(c.L, c.h)
    return { h: c.h, s: Math.min(1, Math.max(0.15, tope > 0 ? c.C / tope : 1)) }
  }
  const p = PALETAS.find((x) => x.id === t.paleta)
  return { h: p.h, s: p.s }
}
// { '--color-indigo-50': 'oklch(...)', ... } o null si es la paleta original
export function variablesDe(tema) {
  const par = parametrosDe(tema)
  if (!par) return null
  const e = generarEscala(par)
  return Object.fromEntries(PASOS.map((p) => [nombreVar(p), e[p].oklch]))
}

// ---------------------------------------------------------------- aplicar en pantalla ----------
let ORIGINAL = null // valores de fábrica de las 11 variables, leídos de la hoja de estilos real antes de aplicar nada
export function leerOriginal(raiz = typeof document !== 'undefined' ? document.documentElement : null) {
  if (ORIGINAL) return ORIGINAL
  if (!raiz || typeof getComputedStyle === 'undefined') return {}
  // se leen SIN las variables que haya puesto un tema (si ya hubiera alguna), para tomar el valor real de fábrica
  const guardadas = {}
  PASOS.forEach((p) => { const v = raiz.style.getPropertyValue(nombreVar(p)); if (v) { guardadas[p] = v; raiz.style.removeProperty(nombreVar(p)) } })
  const cs = getComputedStyle(raiz)
  const leidas = {}
  PASOS.forEach((p) => { const v = cs.getPropertyValue(nombreVar(p)).trim(); if (v) leidas[p] = v })
  Object.entries(guardadas).forEach(([p, v]) => raiz.style.setProperty(nombreVar(p), v))
  if (Object.keys(leidas).length === PASOS.length) ORIGINAL = leidas
  return leidas
}

// ¿esta instalación permite cambiar la paleta? (si la hoja de estilos no usa variables, cambiarlas no tendría efecto)
export function soportaTema(clase = 'bg-indigo-600') {
  if (typeof document === 'undefined' || typeof getComputedStyle === 'undefined') return false
  const sonda = document.createElement('span')
  sonda.className = clase
  sonda.style.cssText = 'position:absolute;left:-9999px;top:-9999px;'
  document.body.appendChild(sonda)
  try {
    const antes = getComputedStyle(sonda).backgroundColor
    sonda.style.setProperty(nombreVar(600), 'rgb(1, 2, 3)')
    const despues = getComputedStyle(sonda).backgroundColor
    return despues !== antes && /rgb\(\s*1,\s*2,\s*3\s*\)/.test(despues)
  } finally { sonda.remove() }
}

export function aplicarTema(tema, raiz = typeof document !== 'undefined' ? document.documentElement : null) {
  if (!raiz) return
  leerOriginal(raiz) // se guarda lo de fábrica antes de cambiar nada
  const vars = variablesDe(tema)
  PASOS.forEach((p) => raiz.style.removeProperty(nombreVar(p)))
  if (vars) Object.entries(vars).forEach(([k, v]) => raiz.style.setProperty(k, v))
  raiz.dataset.paleta = sanitizarTema(tema)?.paleta || 'indigo'
}

// Estilo para fijar un bloque a los colores ORIGINALES, aunque el usuario tenga otra paleta. Se pone en el contenedor de lo que
// representa un DOCUMENTO de la app (la orden impresa, por ejemplo): las variables que se declaran ahí mandan sobre las de <html>.
export function estiloDocumento() {
  const o = leerOriginal()
  return Object.fromEntries(PASOS.filter((p) => o[p]).map((p) => [nombreVar(p), o[p]]))
}