import { PDFDocument, rgb, StandardFonts } from 'pdf-lib'

// Toma el PDF original (por su URL en Storage) y estampa la firma digital, devolviendo un PDF nuevo (Blob).
//  - Órdenes que genera el sistema (servicio/trabajo): el PDF trae anotada en sus metadatos la posición del recuadro
//    "AUTORIZA" (palabras clave firma-pag / firma-cx / firma-yl), y la firma se estampa ahí: la foto sobre la línea y,
//    debajo de "AUTORIZA", el texto "Firmado digitalmente por...".
//  - Órdenes subidas por Compras (PDF de otro sistema): se estampa como siempre, abajo a la derecha de TODAS las páginas,
//    con el contador de páginas.
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

  // ¿el PDF trae anotada la posición de la firma? (lo genera el sistema para las órdenes de servicio/trabajo)
  const claves = pdfDoc.getKeywords() || ''
  const numero = (re) => { const m = claves.match(re); return m ? parseFloat(m[1]) : null }
  const paginaFija = numero(/firma-pag=(\d+)/)
  const centroX = numero(/firma-cx=([\d.]+)/)
  const yLinea = numero(/firma-yl=([\d.]+)/)

  if (paginaFija !== null && centroX !== null && yLinea !== null) {
    const fuente = await pdfDoc.embedFont(StandardFonts.Helvetica)
    const pagina = paginas[Math.min(Math.max(paginaFija, 1), paginas.length) - 1]
    if (imagen) {
      // la foto se ajusta al espacio sobre la línea de firma (sin salirse del recuadro)
      const escala = Math.min(1, 150 / imagen.width, 52 / imagen.height)
      const dims = imagen.scale(escala)
      pagina.drawImage(imagen, { x: centroX - dims.width / 2, y: yLinea + 4, width: dims.width, height: dims.height })
    }
    lineas.forEach((linea, i) => {
      let tamano = 8
      while (fuente.widthOfTextAtSize(linea, tamano) > 210 && tamano > 5.5) tamano -= 0.5 // que una firma con nombre largo no se salga del recuadro
      const ancho = fuente.widthOfTextAtSize(linea, tamano)
      pagina.drawText(linea, { x: centroX - ancho / 2, y: yLinea - 27 - i * 11, size: tamano, font: fuente, color: rgb(0.35, 0.35, 0.35) })
    })
  } else {
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
  }

  const nuevosBytes = await pdfDoc.save()
  return new Blob([nuevosBytes], { type: 'application/pdf' })
}