import { PDFDocument, rgb } from 'pdf-lib'

// Toma el PDF original (por su URL en Storage) y estampa la firma en TODAS sus páginas
// (foto de firma arriba, y debajo el texto "Firmado digitalmente por..."), devolviendo un PDF nuevo (Blob).
export async function firmarPDF(urlPdfOriginal, urlFirmaFoto, nombreFirmante, cargoFirmante, empresaNombre) {
  const pdfBytes = await fetch(urlPdfOriginal).then((r) => r.arrayBuffer())
  const pdfDoc = await PDFDocument.load(pdfBytes)
  const paginas = pdfDoc.getPages()

  let imagen = null
  if (urlFirmaFoto) {
    const imgBytes = await fetch(urlFirmaFoto).then((r) => r.arrayBuffer())
    try {
      imagen = await pdfDoc.embedPng(imgBytes)
    } catch {
      imagen = await pdfDoc.embedJpg(imgBytes)
    }
  }

  const fecha = new Date().toLocaleString('es-CO')
  const lineas = [
    `Firmado digitalmente por ${nombreFirmante}`,
    [cargoFirmante, empresaNombre].filter(Boolean).join(' · '),
    fecha,
  ].filter(Boolean)

  paginas.forEach((pagina, idx) => {
    const { width } = pagina.getSize()
    const columnaX = width - 210 // misma columna para la imagen y el texto, alineados a la derecha

    let yTexto = 68 // arranca justo aquí; si hay imagen, se recalcula más abajo de ella

    if (imagen) {
      const dims = imagen.scale(0.18)
      const yImagen = 68
      pagina.drawImage(imagen, {
        x: columnaX,
        y: yImagen,
        width: dims.width,
        height: dims.height,
      })
      yTexto = yImagen - 12 // el texto queda justo debajo de la imagen
    }

    lineas.forEach((linea, i) => {
      pagina.drawText(linea, {
        x: columnaX,
        y: yTexto - i * 11,
        size: 8,
        color: rgb(0.35, 0.35, 0.35),
      })
    })

    // contador de páginas (ej. 1/15, 2/15...)
    pagina.drawText(`${idx + 1}/${paginas.length}`, {
      x: width - 45,
      y: 20,
      size: 8,
      color: rgb(0.35, 0.35, 0.35),
    })
  })

  const nuevosBytes = await pdfDoc.save()
  return new Blob([nuevosBytes], { type: 'application/pdf' })
}