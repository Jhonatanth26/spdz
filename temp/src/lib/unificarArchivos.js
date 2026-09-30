import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

// Une los archivos de una solicitud (cotizaciones adjuntas, órdenes, soportes de recepción, factura, actas...) en un
// solo PDF, con una portada a modo de índice. Pensado para tener el expediente completo listo en el momento de una
// auditoría, sin tener que abrir archivo por archivo.
//
// items: [{ titulo, url, tipo: 'pdf' | 'imagen' }] — url debe ser accesible (una URL firmada de Storage, por ejemplo).
// Un archivo que no se pueda leer (borrado, dañado, formato no soportado) no detiene el proceso: queda anotado en el
// índice como "no se pudo incluir" y el resto se sigue procesando.
// Devuelve los bytes del PDF (Uint8Array) y, aparte, la lista de resultados (qué se incluyó y qué no).
export async function unificarArchivosPDF({ folio, titulo, items = [] }) {
  const pdfDoc = await PDFDocument.create()
  const fR = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fB = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
  const W = 612, H = 792, M = 50
  const GRIS = rgb(0.42, 0.42, 0.47)
  const NEGRO = rgb(0.08, 0.08, 0.1)
  const ROJO = rgb(0.75, 0.2, 0.2)

  const portada = pdfDoc.addPage([W, H])

  // ---------- procesa cada archivo, anotando en qué página quedó (o si falló) ----------
  const resultados = []
  for (const it of items) {
    const antes = pdfDoc.getPageCount()
    let ok = false
    let error = null
    try {
      const bytes = await fetch(it.url).then((r) => {
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