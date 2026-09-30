import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

const NEGRO = rgb(0.08, 0.08, 0.1)
const GRIS = rgb(0.42, 0.42, 0.47)

// Página(s) de resumen de la solicitud: sus datos generales y la línea de tiempo completa (cada cambio de estado, cada
// aprobación o rechazo, con quién y cuándo). Es lo primero que trae el expediente unificado — antes de esto, el PDF
// solo traía los archivos adjuntos, pero no la traza de la solicitud en sí. Todo el texto ya viene armado desde la
// app (con las etiquetas y formatos que ya usa el resto del sistema); aquí solo se diagrama.
// timeline: [{ fecha, texto }] — ya ordenada cronológicamente.
export async function generarResumenSolicitudPDF({ folio, tipo, area, empresa, solicitante, fechaCreacion, fechaEstimada, objetivo, justificacion, estadoActual, total, timeline = [] }) {
  const pdfDoc = await PDFDocument.create()
  const fR = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fB = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
  const W = 612, H = 792, M = 50, PIE = 40
  let pagina, y
  const nueva = () => { pagina = pdfDoc.addPage([W, H]); y = H - M }
  nueva()

  const limpio = (s) => String(s ?? '').replace(/[\r\n\t]+/g, ' ').replace(/[^\x20-\x7E\u00A0-\u00FF\u2013\u2014\u2018\u2019\u201C\u201D\u2022\u2026\u20AC]/g, '?')
  const ancho = (s, size, bold) => (bold ? fB : fR).widthOfTextAtSize(limpio(s), size)
  const txt = (s, x, yy, size, o = {}) => pagina.drawText(limpio(s), { x, y: yy, size, font: o.bold ? fB : fR, color: o.color || NEGRO })
  const partir = (str, maxW, size, bold) => {
    const salida = []
    String(str ?? '').split(/\r?\n/).forEach((parrafo) => {
      let actual = ''
      String(parrafo).split(/\s+/).filter(Boolean).forEach((palabra) => {
        const prueba = actual ? `${actual} ${palabra}` : palabra
        if (ancho(prueba, size, bold) <= maxW) { actual = prueba; return }
        if (actual) salida.push(actual)
        actual = ancho(palabra, size, bold) <= maxW ? palabra : palabra.slice(0, Math.max(1, Math.floor(maxW / (size * 0.55))))
      })
      salida.push(actual)
    })
    return salida.length ? salida : ['']
  }
  const salto = (alto) => { if (y - alto < PIE) nueva() }

  txt('Resumen de la solicitud', M, y - 18, 17, { bold: true })
  y -= 42
  txt(`Solicitud ${folio || ''}${tipo ? ` — ${tipo}` : ''}`, M, y, 11, { color: GRIS })
  y -= 22

  const filasDatos = [
    ['Empresa', empresa || '—'], ['Área', area || '—'], ['Solicitante', solicitante || '—'],
    ['Fecha de creación', fechaCreacion || '—'], ['Fecha estimada', fechaEstimada || '—'],
    ['Estado actual', estadoActual || '—'], ['Total (con lo aplicado a hoy)', total || '—'],
  ]
  filasDatos.forEach(([et, val]) => {
    salto(13)
    txt(et + ':', M, y, 9, { bold: true, color: GRIS })
    txt(String(val), M + 150, y, 9)
    y -= 13
  })
  y -= 6
  if (objetivo) {
    const l = partir(`Objetivo: ${objetivo}`, W - 2 * M, 9, false)
    salto(l.length * 12)
    l.forEach((linea) => { txt(linea, M, y, 9); y -= 12 })
    y -= 4
  }
  if (justificacion) {
    const l = partir(`Justificación: ${justificacion}`, W - 2 * M, 9, false)
    salto(l.length * 12)
    l.forEach((linea) => { txt(linea, M, y, 9); y -= 12 })
    y -= 4
  }
  y -= 10
  salto(20)
  txt('Línea de tiempo', M, y, 12, { bold: true })
  y -= 6
  pagina.drawLine({ start: { x: M, y }, end: { x: W - M, y }, thickness: 0.6, color: rgb(0.8, 0.8, 0.83) })
  y -= 16

  if (!timeline.length) {
    txt('Esta solicitud todavía no tiene eventos registrados.', M, y, 9, { color: GRIS })
    y -= 13
  }
  const anchoFecha = 118
  timeline.forEach((ev) => {
    const lineas = partir(ev.texto, W - 2 * M - anchoFecha, 9, false)
    const alto = Math.max(13, lineas.length * 12)
    salto(alto)
    txt(ev.fecha || '—', M, y, 8.5, { color: GRIS })
    lineas.forEach((linea, k) => txt(linea, M + anchoFecha, y - k * 12, 9))
    y -= alto
  })

  const paginas = pdfDoc.getPages()
  paginas.forEach((pg, idx) => {
    pg.drawText('Resumen generado automáticamente por el Sistema de Gestión de Compras.', { x: M, y: 24, size: 7, font: fR, color: GRIS })
    const t = `Página ${idx + 1} de ${paginas.length}`
    pg.drawText(t, { x: W - M - fR.widthOfTextAtSize(t, 7), y: 24, size: 7, font: fR, color: GRIS })
  })
  return pdfDoc.save()
}

// Une los archivos de una solicitud (cotizaciones adjuntas, órdenes, soportes de recepción, factura, actas...) en un
// solo PDF, con una portada a modo de índice. Pensado para tener el expediente completo listo en el momento de una
// auditoría, sin tener que abrir archivo por archivo.
//
// items: [{ titulo, tipo: 'pdf' | 'imagen', url }] o [{ titulo, tipo, bytes }] — con "bytes" (Uint8Array/ArrayBuffer)
// se usa el contenido directamente (para páginas ya generadas por la propia app, como el resumen); con "url" se
// descarga primero (por ejemplo una URL firmada de Storage).
// Un archivo que no se pueda leer (borrado, dañado, formato no soportado) no detiene el proceso: queda anotado en el
// índice como "no se pudo incluir" y el resto se sigue procesando.
// Devuelve los bytes del PDF (Uint8Array) y, aparte, la lista de resultados (qué se incluyó y qué no).
export async function unificarArchivosPDF({ folio, titulo, items = [] }) {
  const pdfDoc = await PDFDocument.create()
  const fR = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fB = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
  const W = 612, H = 792, M = 50
  const ROJO = rgb(0.75, 0.2, 0.2)

  const portada = pdfDoc.addPage([W, H])

  // ---------- procesa cada archivo, anotando en qué página quedó (o si falló) ----------
  const resultados = []
  for (const it of items) {
    const antes = pdfDoc.getPageCount()
    let ok = false
    let error = null
    try {
      const bytes = it.bytes ? it.bytes : await fetch(it.url).then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status}`)
        return r.arrayBuffer()
      })
      if (it.tipo === 'pdf') {
        const src = await PDFDocument.load(bytes, { ignoreEncryption: true })
        const paginas = await pdfDoc.copyPages(src, src.getPageIndices())
        paginas.forEach((p) => pdfDoc.addPage(p))
      } else {
        let img
        try {
          img = await pdfDoc.embedJpg(bytes)
        } catch {
          img = await pdfDoc.embedPng(bytes)
        }
        const pagina = pdfDoc.addPage([W, H])
        const escala = Math.min((W - 2 * M) / img.width, (H - 2 * M) / img.height, 1)
        const dims = img.scale(escala)
        pagina.drawImage(img, { x: (W - dims.width) / 2, y: (H - dims.height) / 2, width: dims.width, height: dims.height })
      }
      ok = pdfDoc.getPageCount() > antes
      if (!ok) error = 'El archivo no tiene páginas.'
    } catch (e) {
      error = e?.message || 'No se pudo leer el archivo.'
    }
    resultados.push({ titulo: it.titulo, ok, error, paginaDesde: ok ? antes + 1 : null })
  }

  // ---------- portada: título + índice con la página donde quedó cada archivo ----------
  let y = H - M
  portada.drawText(titulo || 'Expediente de auditoría', { x: M, y: y - 18, size: 17, font: fB, color: NEGRO })
  y -= 42
  portada.drawText(`Solicitud ${folio || ''}`, { x: M, y, size: 11, font: fR, color: GRIS })
  y -= 15
  portada.drawText(`Generado: ${new Date().toLocaleString('es-CO')}`, { x: M, y, size: 9, font: fR, color: GRIS })
  y -= 15
  const incluidos = resultados.filter((r) => r.ok).length
  portada.drawText(`${incluidos} de ${resultados.length} archivo(s) incluido(s)`, { x: M, y, size: 9, font: fR, color: GRIS })
  y -= 26
  portada.drawText('Contenido', { x: M, y, size: 11, font: fB, color: NEGRO })
  y -= 6
  portada.drawLine({ start: { x: M, y }, end: { x: W - M, y }, thickness: 0.6, color: rgb(0.8, 0.8, 0.83) })
  y -= 16

  const anchoTexto = W - 2 * M - 60
  const partir = (txt, size, font) => {
    const palabras = String(txt).split(/\s+/)
    const lineas = []
    let actual = ''
    palabras.forEach((p) => {
      const prueba = actual ? `${actual} ${p}` : p
      if (font.widthOfTextAtSize(prueba, size) <= anchoTexto) actual = prueba
      else { if (actual) lineas.push(actual); actual = p }
    })
    if (actual) lineas.push(actual)
    return lineas.length ? lineas : ['']
  }

  if (!resultados.length) {
    portada.drawText('Esta solicitud no tiene archivos adjuntos.', { x: M, y, size: 9.5, font: fR, color: GRIS })
  }
  // el índice vive solo en la portada (una página): si no alcanza el espacio, se corta con una nota — el archivo
  // igual queda incluido más adelante en el PDF, solo que no se lista aquí uno por uno
  for (let i = 0; i < resultados.length; i++) {
    const r = resultados[i]
    const lineas = partir(r.titulo, 9.5, fR)
    const alto = lineas.length * 12
    if (y - alto < M + 20) {
      const restantes = resultados.length - i
      const nota = `... y ${restantes} archivo(s) más (incluidos en el PDF, sin listar aquí por espacio).`
      portada.drawText(nota, { x: M, y, size: 9, font: fR, color: GRIS })
      break
    }
    const etiqueta = r.ok ? `p. ${r.paginaDesde}` : 'no se pudo incluir'
    lineas.forEach((linea, k) => {
      if (k === 0) portada.drawText(`${i + 1}. ${linea}`, { x: M, y, size: 9.5, font: fR, color: r.ok ? NEGRO : ROJO })
      else portada.drawText(linea, { x: M + 14, y, size: 9.5, font: fR, color: r.ok ? NEGRO : ROJO })
      if (k === 0) portada.drawText(etiqueta, { x: W - M - fR.widthOfTextAtSize(etiqueta, 9.5), y, size: 9.5, font: fR, color: r.ok ? GRIS : ROJO })
      y -= 12
    })
  }

  const bytes = await pdfDoc.save()
  return { bytes, resultados }
}