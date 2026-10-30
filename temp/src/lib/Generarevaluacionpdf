import { PDFDocument, rgb, StandardFonts } from 'pdf-lib'

// Genera el PDF con el resultado de la evaluación de un proveedor: datos del proveedor, matriz de resultados
// (cada criterio con su calificación, peso y aporte), resultado final y clasificación.
// Devuelve los bytes del PDF (Uint8Array), listos para subir a Storage y enviar como enlace por correo.
//
// filas: [{ grupo, criterio, calificacion, peso (texto, ej. "15%"), aporte (texto, ej. "7.5%") }]
const NEGRO = rgb(0.08, 0.08, 0.1)
const GRIS = rgb(0.42, 0.42, 0.47)
const LINEA = rgb(0.25, 0.25, 0.28)
const FONDO = rgb(0.95, 0.95, 0.96)
const COLORES = { 'Confiable': rgb(0.07, 0.5, 0.27), 'Medio Confiable': rgb(0.72, 0.45, 0.04), 'Poco Confiable': rgb(0.75, 0.15, 0.15) }

export async function generarEvaluacionPDF({
  empresa, proveedor, nit, folio, tipo, fechaEvaluacion, evaluador, cargoEvaluador,
  filas = [], resultado = '', clasificacion = '', accion = '', seguimiento = '', observaciones = '', logoBytes = null,
}) {
  const pdfDoc = await PDFDocument.create()
  const fR = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fB = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
  const W = 612, H = 792, M = 45
  const ANCHO = W - M * 2
  const PIE = 62
  let pagina
  let y
  const nueva = () => { pagina = pdfDoc.addPage([W, H]); y = H - M }
  nueva()
  const hay = (alto) => { if (y - alto < PIE) { nueva(); return true } return false }

  const limpio = (s) => String(s ?? '').replace(/[\r\n\t]+/g, ' ').replace(/[^\x20-\x7E -ÿ–—‘’“”•…€]/g, '?')
  const ancho = (s, size, bold) => (bold ? fB : fR).widthOfTextAtSize(limpio(s), size)
  const txt = (s, x, yy, size, o = {}) => pagina.drawText(limpio(s), { x, y: yy, size, font: o.bold ? fB : fR, color: o.color || NEGRO })
  const txtCentro = (s, xC, yy, size, o = {}) => txt(s, xC - ancho(s, size, o.bold) / 2, yy, size, o)
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
  const caja = (x, yTop, w, h, o = {}) => pagina.drawRectangle({ x, y: yTop - h, width: w, height: h, borderColor: LINEA, borderWidth: 0.8, color: o.fondo })
  const linea = (x1, y1, x2, y2, grosor = 0.8) => pagina.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: grosor, color: LINEA })

  // ---------- ENCABEZADO ----------
  let xTitulo = M
  if (logoBytes) {
    try {
      let img
      try { img = await pdfDoc.embedPng(logoBytes) } catch { img = await pdfDoc.embedJpg(logoBytes) }
      const esc = Math.min(110 / img.width, 44 / img.height)
      pagina.drawImage(img, { x: M, y: y - img.height * esc, width: img.width * esc, height: img.height * esc })
      xTitulo = M + img.width * esc + 14
    } catch { /* sin logo */ }
  }
  txt('RESULTADO DE LA EVALUACIÓN DE PROVEEDOR', xTitulo, y - 14, 14, { bold: true })
  txt(empresa || '', xTitulo, y - 30, 10, { color: GRIS })
  y -= 54
  linea(M, y, M + ANCHO, y, 1.2)
  y -= 14

  // ---------- DATOS ----------
  const datos = [
    ['Proveedor:', proveedor || '', 'NIT:', nit || ''],
    ['Solicitud:', folio || '', 'Tipo:', tipo || ''],
    ['Fecha de evaluación:', fechaEvaluacion || '', 'Evaluado por:', [evaluador, cargoEvaluador].filter(Boolean).join(' - ')],
  ]
  const wL = 92, wV = (ANCHO - wL * 2) / 2
  datos.forEach((fila) => {
    const lineas = fila.map((v, k) => envolver(v, (k % 2 === 0 ? wL : wV) - 10, 9, k % 2 === 0))
    const alto = Math.max(...lineas.map((l) => l.length)) * 11 + 8
    hay(alto)
    let x = M
    fila.forEach((_, k) => {
      const w = k % 2 === 0 ? wL : wV
      caja(x, y, w, alto, { fondo: k % 2 === 0 ? FONDO : undefined })
      lineas[k].forEach((l, i) => txt(l, x + 5, y - 12 - i * 11, 9, { bold: k % 2 === 0 }))
      x += w
    })
    y -= alto
  })
  y -= 14

  // ---------- MATRIZ DE RESULTADOS ----------
  txt('MATRIZ DE RESULTADOS', M, y - 9, 10, { bold: true })
  y -= 18
  const cols = [{ t: 'Grupo', w: 70 }, { t: 'Criterio', w: ANCHO - 70 - 100 - 50 - 60 }, { t: 'Calificación', w: 100 }, { t: 'Peso', w: 50 }, { t: 'Aporte', w: 60 }]
  const encabezadoTabla = () => {
    hay(22)
    let x = M
    cols.forEach((c) => { caja(x, y, c.w, 18, { fondo: FONDO }); txt(c.t, x + 5, y - 12, 8.5, { bold: true }); x += c.w })
    y -= 18
  }
  encabezadoTabla()
  filas.forEach((f) => {
    const celdas = [f.grupo, f.criterio, f.calificacion, f.peso, f.aporte]
    const lineas = celdas.map((v, k) => envolver(v, cols[k].w - 10, 8.5, false))
    const alto = Math.max(...lineas.map((l) => l.length)) * 10.5 + 8
    if (hay(alto + 4)) encabezadoTabla()
    let x = M
    cols.forEach((c, k) => {
      caja(x, y, c.w, alto)
      lineas[k].forEach((l, i) => txt(l, x + 5, y - 11 - i * 10.5, 8.5))
      x += c.w
    })
    y -= alto
  })
  y -= 16

  // ---------- RESULTADO FINAL ----------
  hay(70)
  const color = COLORES[clasificacion] || NEGRO
  caja(M, y, ANCHO, 62, { fondo: FONDO })
  txt('RESULTADO FINAL', M + 12, y - 18, 9, { bold: true, color: GRIS })
  txt(`${resultado}`, M + 12, y - 42, 22, { bold: true, color })
  txt(clasificacion, M + 140, y - 42, 18, { bold: true, color })
  txt('Poco Confiable <= 50%  |  Medio Confiable 51%-79%  |  Confiable >= 80%', M + 12, y - 56, 7, { color: GRIS })
  y -= 62 + 14

  const bloque = (titulo, texto) => {
    if (!String(texto || '').trim()) return
    const lineas = envolver(texto, ANCHO - 16, 9, false)
    hay(22 + lineas.length * 11)
    txt(titulo, M, y - 9, 9, { bold: true })
    y -= 16
    lineas.forEach((l) => { txt(l, M + 8, y - 8, 9); y -= 11 })
    y -= 8
  }
  bloque('Acción:', accion)
  bloque('Seguimiento:', seguimiento)
  bloque('Observaciones:', observaciones)

  hay(40)
  y -= 14
  linea(M, y, M + 200, y, 0.8)
  txt(evaluador || '', M, y - 12, 9, { bold: true })
  txt(`${cargoEvaluador || ''}${cargoEvaluador ? ' - ' : ''}${empresa || ''}`, M, y - 23, 8.5, { color: GRIS })

  const paginas = pdfDoc.getPages()
  paginas.forEach((pg, idx) => {
    pg.drawText('Documento generado automáticamente por el Sistema de Gestión de Compras.', { x: M, y: 30, size: 7, font: fR, color: GRIS })
    const t = `Página ${idx + 1} de ${paginas.length}`
    pg.drawText(t, { x: W - M - fR.widthOfTextAtSize(t, 7), y: 30, size: 7, font: fR, color: GRIS })
  })
  return pdfDoc.save()
}
