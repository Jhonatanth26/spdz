import { PDFDocument, rgb, StandardFonts } from 'pdf-lib'

// Genera el PDF de la Orden de Servicio/Trabajo desde los datos de la solicitud —
// se usa cuando el sistema contable (Zeus) no genera este tipo de orden.
// Sigue el formato de la orden de trabajo de la empresa: encabezado con el recuadro del número, datos del proveedor,
// detalle del trabajo, tabla de ítems (con su AIU por ítem), totales, "SON", forma de pago, firmas y observaciones.
// Devuelve los bytes del PDF (Uint8Array), listos para subir a Storage y luego firmar.
//
// items: [{ nombre, cantidad, unidad, valorUnitario, total (costo directo), aiuPct: {a,u,i}, admin, util, imprev, ivaUtil, totalItem }]
const NEGRO = rgb(0.08, 0.08, 0.1)
const GRIS = rgb(0.42, 0.42, 0.47)
const LINEA = rgb(0.25, 0.25, 0.28)
const FONDO = rgb(0.95, 0.95, 0.96)

export async function generarOrdenServicioPDF({
  solicitud, empresa, proveedor, proveedorNombre, area, solicitanteNombre, items = [],
  costoDirecto = 0, administracion = 0, utilidad = 0, imprevistos = 0, ivaUtilidad = 0, total = 0, planesPago = [],
}) {
  const pdfDoc = await PDFDocument.create()
  const fR = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fB = await pdfDoc.embedFont(StandardFonts.HelveticaBold)

  const W = 612, H = 792, M = 45
  const ANCHO = W - M * 2
  const PIE = 50 // espacio reservado abajo para el pie de página
  let pagina
  let y

  const nueva = () => { pagina = pdfDoc.addPage([W, H]); y = H - M }
  nueva()
  const hay = (alto) => { if (y - alto < PIE) { nueva(); return true } return false }

  // ---------- utilidades de texto ----------
  // Helvetica solo admite el alfabeto latino básico: cualquier otro símbolo (o salto de línea) se reemplaza para no romper el PDF
  const limpio = (s) => String(s ?? '').replace(/[\r\n\t]+/g, ' ').replace(/[^\x20-\x7E\u00A0-\u00FF\u2013\u2014\u2018\u2019\u201C\u201D\u2022\u2026\u20AC]/g, '?')
  const ancho = (s, size, bold) => (bold ? fB : fR).widthOfTextAtSize(limpio(s), size)
  const txt = (s, x, yy, size, o = {}) => pagina.drawText(limpio(s), { x, y: yy, size, font: o.bold ? fB : fR, color: o.color || NEGRO })
  const txtDer = (s, xDer, yy, size, o = {}) => txt(s, xDer - ancho(s, size, o.bold), yy, size, o)
  const txtCentro = (s, xC, yy, size, o = {}) => txt(s, xC - ancho(s, size, o.bold) / 2, yy, size, o)
  // envuelve un texto (respetando los saltos de línea que traiga) dentro de un ancho máximo; parte las palabras que no caben
  const envolver = (str, maxW, size, bold) => {
    const salida = []
    String(str ?? '').split(/\r?\n/).forEach((parrafo) => {
      let actual = ''
      const poner = (palabra) => {
        const prueba = actual ? `${actual} ${palabra}` : palabra
        if (ancho(prueba, size, bold) <= maxW) { actual = prueba; return }
        if (actual) { salida.push(actual); actual = '' }
        if (ancho(palabra, size, bold) <= maxW) { actual = palabra; return }
        let trozo = ''
        for (const ch of palabra) { if (ancho(trozo + ch, size, bold) > maxW && trozo) { salida.push(trozo); trozo = '' } trozo += ch }
        actual = trozo
      }
      parrafo.split(/\s+/).filter(Boolean).forEach(poner)
      salida.push(actual)
    })
    return salida.length ? salida : ['']
  }
  const caja = (x, yTop, w, h, o = {}) => pagina.drawRectangle({ x, y: yTop - h, width: w, height: h, borderColor: LINEA, borderWidth: o.grosor || 0.8, color: o.fondo })
  const linea = (x1, y1, x2, y2, grosor = 0.8) => pagina.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: grosor, color: LINEA })
  const fmt = (v) => '$ ' + Math.round(parseFloat(v) || 0).toLocaleString('es-CO')
  const dato = (v) => (v === undefined || v === null || String(v).trim() === '' ? '—' : String(v))

  // ---------- ENCABEZADO ----------
  txt(empresa?.nombre || 'Empresa', M, y - 13, 15, { bold: true })
  txt(`Nit: ${dato(empresa?.nit)}`, M, y - 29, 9, { color: GRIS })
  const bw = 180, bx = W - M - bw
  caja(bx, y, bw, 46)
  linea(bx, y - 21, bx + bw, y - 21)
  txtCentro('ORDEN DE SERVICIO / TRABAJO', bx + bw / 2, y - 14, 9, { bold: true })
  txtCentro(`No ${solicitud.folio}`, bx + bw / 2, y - 38, 13, { bold: true })
  y -= 62

  const fecha = new Date().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' })
  txt('FECHA:', M, y - 9, 9, { bold: true })
  txt(fecha, M + 44, y - 9, 9)
  y -= 22

  // ---------- DATOS DEL PROVEEDOR Y DEL SOLICITANTE (cuadrícula con borde) ----------
  const cols = [92, 175, 68, ANCHO - 92 - 175 - 68]
  const filasDatos = [
    ['Proveedor:', dato(proveedor?.nombre || proveedorNombre), 'NIT:', dato(proveedor?.nit)],
    ['Dirección:', dato(proveedor?.direccion), 'Teléfono:', dato(proveedor?.telefono)],
    ['Oficina solicitante:', dato(area), 'Responsable:', dato(solicitanteNombre)],
  ]
  filasDatos.forEach((f) => {
    const lineas = [envolver(f[1], cols[1] - 10, 9, false), envolver(f[3], cols[3] - 10, 9, false)]
    const alto = Math.max(20, Math.max(lineas[0].length, lineas[1].length) * 11.5 + 9)
    hay(alto)
    let x = M
    f.forEach((celda, i) => {
      caja(x, y, cols[i], alto, { fondo: i % 2 === 0 ? FONDO : undefined })
      if (i % 2 === 0) txt(celda, x + 5, y - 6 - 9 * 0.9, 9, { bold: true })
      else lineas[(i - 1) / 2].forEach((l, k) => txt(l, x + 5, y - 6 - 9 * 0.9 - k * 11.5, 9))
      x += cols[i]
    })
    y -= alto
  })
  y -= 12

  // ---------- DETALLE DEL TRABAJO AUTORIZADO ----------
  {
    let lineas = envolver(solicitud.objetivo, ANCHO - 16, 9, false)
    if (lineas.length > 16) lineas = [...lineas.slice(0, 15), lineas[15].slice(0, 90) + '…']
    const alto = Math.max(50, 8 + 12 + lineas.length * 11.5 + 8)
    hay(alto)
    caja(M, y, ANCHO, alto)
    txt('DETALLE DEL TRABAJO AUTORIZADO:', M + 8, y - 8 - 9 * 0.9, 9, { bold: true })
    lineas.forEach((l, k) => txt(l, M + 8, y - 8 - 12 - 9 * 0.9 - k * 11.5, 9))
    y -= alto + 12
  }

  // ---------- ÍTEMS (con su AIU por ítem) ----------
  const cw = [172, 44, 76, 78, 66, ANCHO - 172 - 44 - 76 - 78 - 66]
  const xs = cw.reduce((acc, w, i) => { acc.push(i === 0 ? M : acc[i - 1] + cw[i - 1]); return acc }, [])
  const derecha = (i) => xs[i] + cw[i] - 5
  let topSegmento = null
  const cerrarSegmento = () => {
    if (topSegmento === null) return
    for (let i = 1; i < cw.length; i++) linea(xs[i], topSegmento, xs[i], y, 0.5)
    caja(M, topSegmento, ANCHO, topSegmento - y, { grosor: 0.8 })
    topSegmento = null
  }
  const encabezadoTabla = () => {
    topSegmento = y
    pagina.drawRectangle({ x: M, y: y - 22, width: ANCHO, height: 22, color: FONDO })
    txt('DESCRIPCIÓN', xs[0] + 5, y - 14, 7.5, { bold: true })
    txtCentro('CANT.', xs[1] + cw[1] / 2, y - 14, 7.5, { bold: true })
    txtDer('VR. UNITARIO', derecha(2), y - 14, 7.5, { bold: true })
    txtDer('COSTO DIRECTO', derecha(3), y - 14, 7.5, { bold: true })
    txtCentro('AIU (A/U/I %)', xs[4] + cw[4] / 2, y - 14, 7.5, { bold: true })
    txtDer('TOTAL ÍTEM', derecha(5), y - 14, 7.5, { bold: true })
    y -= 22
    linea(M, y, M + ANCHO, y, 0.8)
  }
  hay(22 + 40)
  encabezadoTabla()
  const aiuTexto = (a) => [a?.a, a?.u, a?.i].map((v) => (parseFloat(v) ? String(parseFloat(v)) : '0')).join(' / ')
  items.forEach((it) => {
    const lDesc = envolver(it.nombre, cw[0] - 10, 8.5, false)
    const alto = Math.max(30, lDesc.length * 10.5 + 10)
    if (y - alto < PIE) { cerrarSegmento(); nueva(); encabezadoTabla() }
    const base = y - 6 - 8.5 * 0.9
    lDesc.forEach((l, k) => txt(l, xs[0] + 5, base - k * 10.5, 8.5))
    txtCentro(String(it.cantidad ?? ''), xs[1] + cw[1] / 2, base, 8.5)
    txtCentro(it.unidad || '', xs[1] + cw[1] / 2, base - 10.5, 6.5, { color: GRIS })
    txtDer(fmt(it.valorUnitario), derecha(2), base, 8.5)
    txtDer(fmt(it.total), derecha(3), base, 8.5)
    txtCentro(aiuTexto(it.aiuPct), xs[4] + cw[4] / 2, base, 8)
    const valorAiu = (parseFloat(it.admin) || 0) + (parseFloat(it.util) || 0) + (parseFloat(it.imprev) || 0) + (parseFloat(it.ivaUtil) || 0)
    txtCentro(fmt(valorAiu), xs[4] + cw[4] / 2, base - 10.5, 7, { color: GRIS })
    txtDer(fmt(it.totalItem ?? it.total), derecha(5), base, 8.5, { bold: true })
    y -= alto
    linea(M, y, M + ANCHO, y, 0.4)
  })
  cerrarSegmento()
  y -= 14

  // ---------- TOTALES ----------
  const hayAiu = (administracion + utilidad + imprevistos) > 0
  const filasTotales = [['SUBTOTAL (COSTO DIRECTO):', costoDirecto]]
  if (hayAiu) filasTotales.push(['ADMINISTRACIÓN:', administracion], ['UTILIDAD:', utilidad], ['IMPREVISTOS:', imprevistos], ['IVA SOBRE LA UTILIDAD (19%):', ivaUtilidad])
  hay(filasTotales.length * 13 + 30 + 40)
  const anchoTot = 250
  filasTotales.forEach(([et, v]) => { txt(et, M, y - 9, 9, { bold: true }); txtDer(fmt(v), M + anchoTot, y - 9, 9); y -= 13 })
  linea(M, y - 2, M + anchoTot, y - 2, 0.6)
  txt('TOTAL', M, y - 15, 11, { bold: true }); txtDer(fmt(total), M + anchoTot, y - 15, 11, { bold: true })
  y -= 26

  // ---------- SON: (monto en letras) ----------
  {
    const lineas = envolver(`SON: ${numeroALetras_(total)}`, ANCHO - 12, 9, true)
    const alto = lineas.length * 12 + 10
    hay(alto)
    linea(M, y, M + ANCHO, y, 0.6)
    lineas.forEach((l, k) => txt(l, M + 4, y - 5 - 9 * 0.9 - k * 12, 9, { bold: true }))
    y -= alto
    linea(M, y, M + ANCHO, y, 0.6)
    y -= 14
  }

  // ---------- FORMA DE PAGO / PLAZO DE ENTREGA / FIRMAS / OBSERVACIONES (se mantienen juntos) ----------
  const lineasPago = []
  planesPago.forEach((pl) => {
    const p = pl.pagos
    const corto = (t, n) => (t.length > n ? t.slice(0, n - 1).trimEnd() + '…' : t)
    const pref = planesPago.length > 1 ? `${pl.numero ? pl.numero + '. ' : ''}${corto(String(pl.nombre || ''), 50)}:` : ''
    const tramos = p.tipoPago === 'contado'
      ? [`Pago único ${fmt(p.pagoUnico?.valor)} — ${p.pagoUnico?.fecha || 'sin fecha'}`]
      : [
          `Anticipo ${fmt(p.anticipo?.valor)} — ${p.anticipo?.fecha || 'sin fecha'}`,
          ...(p.intermedio?.activo ? [`Intermedio ${fmt(p.intermedio?.valor)} — ${p.intermedio?.fecha || 'sin fecha'}`] : []),
          `Pago final ${fmt(p.final?.valor)} — ${p.final?.fecha || 'sin fecha'}`,
        ]
    if (pref) lineasPago.push(pref)
    tramos.forEach((t) => lineasPago.push(t))
  })
  const wPago = Math.round(ANCHO * 0.6), wPlazo = ANCHO - wPago - 8
  const lPago = lineasPago.length ? lineasPago.flatMap((l) => envolver(l, wPago - 16, 8.5, false)) : ['Sin plan de pagos definido.']
  const lPlazo = [solicitud.folio, ...(solicitud.fechaEstimada ? [`Entrega estimada: ${solicitud.fechaEstimada}`] : [])].flatMap((l) => envolver(l, wPlazo - 16, 8.5, false))
  const altoPago = Math.max(46, 8 + 12 + Math.max(lPago.length, lPlazo.length) * 10.5 + 8)
  const altoFirmas = 118 // deja espacio bajo "AUTORIZA" para las 3 líneas de la firma digital
  const altoObs = 46
  hay(altoPago + 10 + altoFirmas + 10 + altoObs)

  caja(M, y, wPago, altoPago)
  txt('FORMA DE PAGO:', M + 8, y - 8 - 9 * 0.9, 9, { bold: true })
  lPago.forEach((l, k) => txt(l, M + 8, y - 8 - 12 - 8.5 * 0.9 - k * 10.5, 8.5))
  caja(M + wPago + 8, y, wPlazo, altoPago)
  txt('PLAZO ENTREGA TRABAJOS:', M + wPago + 16, y - 8 - 9 * 0.9, 9, { bold: true })
  lPlazo.forEach((l, k) => txt(l, M + wPago + 16, y - 8 - 12 - 8.5 * 0.9 - k * 10.5, 8.5))
  y -= altoPago + 10

  // firmas: en AUTORIZA se estampa la firma digital de Dirección Financiera (queda anotada su posición en el PDF)
  const wFirma = (ANCHO - 10) / 2
  const yLineaFirma = y - 62
  ;[['AUTORIZA', M], ['ACEPTADO', M + wFirma + 10]].forEach(([et, x]) => {
    caja(x, y, wFirma, altoFirmas)
    linea(x + 22, yLineaFirma, x + wFirma - 22, yLineaFirma, 0.8)
    txtCentro(et, x + wFirma / 2, yLineaFirma - 12, 9, { bold: true })
  })
  pdfDoc.setKeywords([`firma-pag=${pdfDoc.getPageCount()}`, `firma-cx=${(M + wFirma / 2).toFixed(1)}`, `firma-yl=${yLineaFirma.toFixed(1)}`])
  y -= altoFirmas + 10

  caja(M, y, ANCHO, altoObs)
  txt('OBSERVACIONES:', M + 8, y - 8 - 9 * 0.9, 9, { bold: true })
  y -= altoObs

  // ---------- PIE DE CADA PÁGINA ----------
  const paginas = pdfDoc.getPages()
  paginas.forEach((pg, idx) => {
    pg.drawText('Documento generado automáticamente por el Sistema de Gestión de Compras.', { x: M, y: 30, size: 7, font: fR, color: GRIS })
    const t = `Página ${idx + 1} de ${paginas.length}`
    pg.drawText(t, { x: W - M - fR.widthOfTextAtSize(t, 7), y: 30, size: 7, font: fR, color: GRIS })
  })

  return pdfDoc.save()
}

// convierte un número a su forma escrita en español, para el "SON: ..." (ej. 7591300 -> "SIETE MILLONES QUINIENTOS NOVENTA Y UN MIL TRESCIENTOS PESOS M/CTE")
function numeroALetras_(n) {
  const UNIDADES = ['', 'UN', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE']
  const DIEZ_A_DIECINUEVE = ['DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE']
  const DECENAS = ['', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA']
  const CENTENAS = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS']
  const VEINTI = ['', 'VEINTIÚN', 'VEINTIDÓS', 'VEINTITRÉS', 'VEINTICUATRO', 'VEINTICINCO', 'VEINTISÉIS', 'VEINTISIETE', 'VEINTIOCHO', 'VEINTINUEVE']

  function trescientos(num) {
    if (num === 0) return ''
    if (num === 100) return 'CIEN'
    let s = ''
    const c = Math.floor(num / 100), resto = num % 100
    if (c > 0) s += CENTENAS[c] + ' '
    if (resto >= 10 && resto <= 19) { s += DIEZ_A_DIECINUEVE[resto - 10] }
    else {
      const d = Math.floor(resto / 10), u = resto % 10
      if (d === 2 && u > 0) s += VEINTI[u]
      else { if (d > 0) s += DECENAS[d]; if (d > 0 && u > 0) s += ' Y '; if (u > 0) s += UNIDADES[u] }
    }
    return s.trim()
  }

  const entero = Math.round(Math.abs(n || 0))
  if (entero === 0) return 'CERO PESOS M/CTE'
  const millones = Math.floor(entero / 1000000)
  const miles = Math.floor((entero % 1000000) / 1000)
  const resto = entero % 1000
  const partes = []
  if (millones > 0) partes.push(millones === 1 ? 'UN MILLÓN' : trescientos(millones) + ' MILLONES')
  if (miles > 0) partes.push(miles === 1 ? 'MIL' : trescientos(miles) + ' MIL')
  if (resto > 0) partes.push(trescientos(resto))
  return (partes.join(' ').trim() || 'CERO') + ' PESOS M/CTE'
}