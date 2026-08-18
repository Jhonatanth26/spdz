import { PDFDocument, rgb } from 'pdf-lib'

// Toma el PDF original (por su URL en Storage) y estampa la firma en TODAS sus páginas
// (foto de firma + nombre, cargo y empresa de quien firma), devolviendo un PDF nuevo (Blob).
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

  paginas.forEach((pagina) => {
    const { width } = pagina.getSize()

    if (imagen) {
      const dims = imagen.scale(0.18)
      pagina.drawImage(imagen, {
        x: width - dims.width - 50,
        y: 78,
        width: dims.width,
        height: dims.height,
      })
    }

    lineas.forEach((linea, i) => {
      pagina.drawText(linea, {
        x: 40,
        y: 42 - i * 11,
        size: 8,
        color: rgb(0.35, 0.35, 0.35),
      })
    })
  })

  const nuevosBytes = await pdfDoc.save()
  return new Blob([nuevosBytes], { type: 'application/pdf' })
}