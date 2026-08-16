import { PDFDocument, rgb } from 'pdf-lib'

// Toma el PDF original (por su URL en Storage), le estampa la foto de firma
// de quien firma en la última página, y devuelve un nuevo PDF (como Blob) ya firmado.
export async function firmarPDF(urlPdfOriginal, urlFirmaFoto, nombreFirmante) {
  const pdfBytes = await fetch(urlPdfOriginal).then((r) => r.arrayBuffer())
  const pdfDoc = await PDFDocument.load(pdfBytes)
  const paginas = pdfDoc.getPages()
  const ultima = paginas[paginas.length - 1]
  const { width } = ultima.getSize()

  if (urlFirmaFoto) {
    const imgBytes = await fetch(urlFirmaFoto).then((r) => r.arrayBuffer())
    let imagen
    try {
      imagen = await pdfDoc.embedPng(imgBytes)
    } catch {
      imagen = await pdfDoc.embedJpg(imgBytes)
    }
    const dims = imagen.scale(0.2)
    ultima.drawImage(imagen, {
      x: width - dims.width - 50,
      y: 70,
      width: dims.width,
      height: dims.height,
    })
  }

  const fecha = new Date().toLocaleString('es-CO')
  ultima.drawText(`Firmado digitalmente por ${nombreFirmante} - ${fecha}`, {
    x: 40,
    y: 40,
    size: 8,
    color: rgb(0.35, 0.35, 0.35),
  })

  const nuevosBytes = await pdfDoc.save()
  return new Blob([nuevosBytes], { type: 'application/pdf' })
}