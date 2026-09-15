import { PDFDocument, rgb, StandardFonts } from 'pdf-lib'

// Genera el PDF de la Orden de Servicio/Trabajo desde los datos de la solicitud —
// se usa cuando el sistema contable (Zeus) no genera este tipo de orden.
// Devuelve los bytes del PDF (Uint8Array), listos para subir a Storage y luego firmar.
export async function generarOrdenServicioPDF({ solicitud, empresa, proveedorNombre, items, costoDirecto, administracion, utilidad, imprevistos, ivaUtilidad, total, aiuPcts }) {
  const pdfDoc = await PDFDocument.create()
  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold)

  const anchoPagina = 612 // carta
  const altoPagina = 792
  const margen = 45
  const margenDerecho = anchoPagina - margen
  const anchoUtil = anchoPagina - margen * 2

  let pagina = pdfDoc.addPage([anchoPagina, altoPagina])
  let y = altoPagina - margen

  const nuevaPagina = () => {
    pagina = pdfDoc.addPage([anchoPagina, altoPagina])
    y = altoPagina - margen
  }
  const saltoSiHaceFalta = (alturaNecesaria) => {
    if (y - alturaNecesaria < margen + 40) nuevaPagina()
  }

  const texto = (str, x, tamano, opts = {}) => {
    pagina.drawText(str || '', { x, y, size: tamano, font: opts.bold ? fontBold : fontRegular, color: opts.color || rgb(0.12, 0.16, 0.22) })
  }
  // dibuja el texto terminando exactamente en xDerecha, sin importar cuántos caracteres tenga —
  // así un número o una etiqueta larga nunca se encima con la columna vecina
  const textoDerecha = (str, xDerecha, tamano, opts = {}) => {
    const font = opts.bold ? fontBold : fontRegular
    const ancho = font.widthOfTextAtSize(str || '', tamano)
    pagina.drawText(str || '', { x: xDerecha - ancho, y, size: tamano, font, color: opts.color || rgb(0.12, 0.16, 0.22) })
  }

  // envuelve texto largo dentro de un ancho máximo, devolviendo las líneas
  const envolver = (str, anchoMax, tamano, font) => {
    const palabras = (str || '').split(' ')
    const lineas = []
    let actual = ''
    palabras.forEach((p) => {
      const prueba = actual ? `${actual} ${p}` : p
      if (font.widthOfTextAtSize(prueba, tamano) > anchoMax && actual) {
        lineas.push(actual)
        actual = p
      } else {
        actual = prueba
      }
    })
    if (actual) lineas.push(actual)
    return lineas
  }

  // ---------- ENCABEZADO ----------
  texto(empresa?.nombre || 'Empresa', margen, 16, { bold: true, color: rgb(0.05, 0.05, 0.15) })
  textoDerecha(solicitud.folio, margenDerecho, 13, { bold: true })
  y -= 22
  texto('ORDEN DE SERVICIO / TRABAJO', margen, 13, { bold: true, color: rgb(0.31, 0.27, 0.9) })
  y -= 18
  texto(`Fecha de generación: ${hoy_()}`, margen, 9, { color: rgb(0.4, 0.4, 0.45) })
  y -= 24

  // línea divisoria
  pagina.drawLine({ start: { x: margen, y }, end: { x: margenDerecho, y }, thickness: 0.75, color: rgb(0.85, 0.85, 0.9) })
  y -= 18

  // ---------- DATOS ----------
  const xEtiquetaDatos = margen + 85
  texto('Proveedor:', margen, 10, { bold: true }); texto(proveedorNombre || '—', xEtiquetaDatos, 10)
  y -= 16
  texto('Solicitante:', margen, 10, { bold: true }); texto(solicitud.firmas?.solicitante?.nombre || '—', xEtiquetaDatos, 10)
  y -= 16
  texto('Objetivo:', margen, 10, { bold: true })
  const lineasObjetivo = envolver(solicitud.objetivo, anchoUtil - 85, 10, fontRegular)
  lineasObjetivo.slice(0, 3).forEach((linea, i) => { if (i === 0) texto(linea, xEtiquetaDatos, 10); else { y -= 13; texto(linea, xEtiquetaDatos, 10) } })
  y -= 22

  pagina.drawLine({ start: { x: margen, y }, end: { x: margenDerecho, y }, thickness: 0.75, color: rgb(0.85, 0.85, 0.9) })
  y -= 20

  // ---------- TABLA DE ÍTEMS ----------
  // columnas: Ítem (izq.) | Cant. | Valor unit. | Total — las 3 últimas alineadas a la derecha,
  // cada una termina exactamente en su propio borde, así nunca se encimen entre sí
  const colItem = margen
  const xCant = margen + 330
  const xValorUnit = margen + 430
  const xTotal = margenDerecho
  texto('Ítem', colItem, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  textoDerecha('Cant.', xCant, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  textoDerecha('Valor unit.', xValorUnit, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  textoDerecha('Total', xTotal, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  y -= 14
  pagina.drawLine({ start: { x: margen, y }, end: { x: margenDerecho, y }, thickness: 0.5, color: rgb(0.9, 0.9, 0.93) })
  y -= 14

  items.forEach((it) => {
    const lineasNombre = envolver(it.nombre, 270, 9, fontRegular)
    saltoSiHaceFalta(lineasNombre.length * 12 + 8)
    lineasNombre.forEach((linea, i) => {
      texto(linea, colItem, 9)
      if (i === 0) {
        textoDerecha(`${it.cantidad} ${it.unidad}`, xCant, 9)
        textoDerecha(fmt_(it.valorUnitario), xValorUnit, 9)
        textoDerecha(fmt_(it.total), xTotal, 9, { bold: true })
      }
      y -= 12
    })
    y -= 4
  })

  y -= 10
  pagina.drawLine({ start: { x: margen + 250, y }, end: { x: margenDerecho, y }, thickness: 0.5, color: rgb(0.9, 0.9, 0.93) })
  y -= 16

  // ---------- TOTALES (Costo Directo + AIU) ----------
  // etiqueta a la izquierda de esta zona, valor alineado al margen derecho real de la página —
  // con esto el valor nunca se encima con la etiqueta, sin importar cuántos dígitos tenga el %
  saltoSiHaceFalta(120)
  const xEtiquetaTotal = margen + 250
  const filaTotal = (etiqueta, valor, opts = {}) => { texto(etiqueta, xEtiquetaTotal, opts.size || 10, opts); textoDerecha(valor, margenDerecho, opts.size || 10, opts); y -= opts.salto || 15 }
  filaTotal('Costo Directo', fmt_(costoDirecto))
  filaTotal(`Administración (${aiuPcts?.administracionPct || 0}%)`, fmt_(administracion))
  filaTotal(`Utilidad (${aiuPcts?.utilidadPct || 0}%)`, fmt_(utilidad))
  filaTotal(`Imprevistos (${aiuPcts?.imprevistosPct || 0}%)`, fmt_(imprevistos))
  filaTotal('IVA sobre la Utilidad (19%)', fmt_(ivaUtilidad))
  y -= 2
  filaTotal('Total', fmt_(total), { bold: true, size: 11, salto: 26 })

  // ---------- CONDICIONES DE PAGO ----------
  if (solicitud.tipo === 'servicio' && solicitud.pagosConfirmados) {
    saltoSiHaceFalta(70)
    texto('Condiciones de pago', margen, 10, { bold: true })
    y -= 15
    const p = solicitud.pagos
    if (p.tipoPago === 'contado') {
      texto(`Pago único (de contado): ${fmt_(p.pagoUnico.valor)}`, margen, 9)
      textoDerecha(p.pagoUnico.fecha || 'sin fecha', margenDerecho, 9)
      y -= 13
    } else {
      texto(`Anticipo: ${fmt_(p.anticipo.valor)}`, margen, 9); textoDerecha(p.anticipo.fecha || 'sin fecha', margenDerecho, 9)
      y -= 13
      if (p.intermedio?.activo) { texto(`Intermedio: ${fmt_(p.intermedio.valor)}`, margen, 9); textoDerecha(p.intermedio.fecha || 'sin fecha', margenDerecho, 9); y -= 13 }
      texto(`Pago final: ${fmt_(p.final.valor)}`, margen, 9); textoDerecha(p.final.fecha || 'sin fecha', margenDerecho, 9)
    }
    y -= 20
  }

  // ---------- PIE / GENERADO POR EL SISTEMA ----------
  const paginas = pdfDoc.getPages()
  paginas.forEach((pg, idx) => {
    pg.drawText('Documento generado automáticamente por el Sistema de Gestión de Compras.', {
      x: margen, y: 30, size: 7, font: fontRegular, color: rgb(0.55, 0.55, 0.6),
    })
    pg.drawText(`${idx + 1}/${paginas.length}`, { x: margenDerecho - 25, y: 30, size: 7, font: fontRegular, color: rgb(0.55, 0.55, 0.6) })
  })

  return pdfDoc.save()
}

function fmt_(v) {
  const n = parseFloat(v) || 0
  return '$ ' + Math.round(n).toLocaleString('es-CO')
}
function hoy_() {
  return new Date().toLocaleDateString('es-CO')
}