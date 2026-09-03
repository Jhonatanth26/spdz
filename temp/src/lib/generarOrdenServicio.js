import { PDFDocument, rgb, StandardFonts } from 'pdf-lib'

// Genera el PDF de la Orden de Servicio/Trabajo desde los datos de la solicitud —
// se usa cuando el sistema contable (Zeus) no genera este tipo de orden.
// Devuelve los bytes del PDF (Uint8Array), listos para subir a Storage y luego firmar.
export async function generarOrdenServicioPDF({ solicitud, empresa, proveedorNombre, items, total, subtotal, iva }) {
  const pdfDoc = await PDFDocument.create()
  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica)
  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold)

  const anchoPagina = 612 // carta
  const altoPagina = 792
  const margen = 45
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
  y -= 22
  texto('ORDEN DE SERVICIO / TRABAJO', margen, 13, { bold: true, color: rgb(0.31, 0.27, 0.9) })
  texto(solicitud.folio, anchoPagina - margen - 90, 13, { bold: true })
  y -= 18
  texto(`Fecha de generación: ${hoy_()}`, margen, 9, { color: rgb(0.4, 0.4, 0.45) })
  y -= 24

  // línea divisoria
  pagina.drawLine({ start: { x: margen, y }, end: { x: anchoPagina - margen, y }, thickness: 0.75, color: rgb(0.85, 0.85, 0.9) })
  y -= 18

  // ---------- DATOS ----------
  texto('Proveedor:', margen, 10, { bold: true }); texto(proveedorNombre || '—', margen + 70, 10)
  y -= 16
  texto('Solicitante:', margen, 10, { bold: true }); texto(solicitud.firmas?.solicitante?.nombre || '—', margen + 70, 10)
  y -= 16
  texto('Objetivo:', margen, 10, { bold: true })
  const lineasObjetivo = envolver(solicitud.objetivo, anchoUtil - 70, 10, fontRegular)
  lineasObjetivo.slice(0, 3).forEach((linea, i) => { if (i === 0) texto(linea, margen + 70, 10); else { y -= 13; texto(linea, margen + 70, 10) } })
  y -= 22

  pagina.drawLine({ start: { x: margen, y }, end: { x: anchoPagina - margen, y }, thickness: 0.75, color: rgb(0.85, 0.85, 0.9) })
  y -= 20

  // ---------- TABLA DE ÍTEMS ----------
  const colItem = margen, colCant = margen + 300, colValor = margen + 370, colTotal = anchoPagina - margen - 80
  texto('Ítem', colItem, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  texto('Cant.', colCant, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  texto('Valor unit.', colValor, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  texto('Total', colTotal, 9, { bold: true, color: rgb(0.4, 0.4, 0.45) })
  y -= 14
  pagina.drawLine({ start: { x: margen, y }, end: { x: anchoPagina - margen, y }, thickness: 0.5, color: rgb(0.9, 0.9, 0.93) })
  y -= 14

  items.forEach((it) => {
    const lineasNombre = envolver(it.nombre, 250, 9, fontRegular)
    saltoSiHaceFalta(lineasNombre.length * 12 + 8)
    lineasNombre.forEach((linea, i) => {
      texto(linea, colItem, 9)
      if (i === 0) {
        texto(`${it.cantidad} ${it.unidad}`, colCant, 9)
        texto(fmt_(it.valorUnitario), colValor, 9)
        texto(fmt_(it.total), colTotal, 9, { bold: true })
      }
      y -= 12
    })
    y -= 4
  })

  y -= 10
  pagina.drawLine({ start: { x: colValor - 10, y }, end: { x: anchoPagina - margen, y }, thickness: 0.5, color: rgb(0.9, 0.9, 0.93) })
  y -= 16

  // ---------- TOTALES ----------
  texto('Subtotal', colValor, 10); texto(fmt_(subtotal), colTotal, 10)
  y -= 15
  texto('IVA', colValor, 10); texto(fmt_(iva), colTotal, 10)
  y -= 17
  texto('Total', colValor, 11, { bold: true }); texto(fmt_(total), colTotal, 11, { bold: true })
  y -= 26

  // ---------- CONDICIONES DE PAGO ----------
  if (solicitud.tipo === 'servicio' && solicitud.pagosConfirmados) {
    saltoSiHaceFalta(70)
    texto('Condiciones de pago', margen, 10, { bold: true })
    y -= 15
    const p = solicitud.pagos
    if (p.tipoPago === 'contado') {
      texto(`Pago único (de contado): ${fmt_(p.pagoUnico.valor)} — ${p.pagoUnico.fecha || 'sin fecha'}`, margen, 9)
      y -= 13
    } else {
      texto(`Anticipo: ${fmt_(p.anticipo.valor)} — ${p.anticipo.fecha || 'sin fecha'}`, margen, 9)
      y -= 13
      if (p.intermedio?.activo) { texto(`Intermedio: ${fmt_(p.intermedio.valor)} — ${p.intermedio.fecha || 'sin fecha'}`, margen, 9); y -= 13 }
      texto(`Pago final: ${fmt_(p.final.valor)} — ${p.final.fecha || 'sin fecha'}`, margen, 9)
    }
    y -= 20
  }

  // ---------- PIE / GENERADO POR EL SISTEMA ----------
  const paginas = pdfDoc.getPages()
  paginas.forEach((pg, idx) => {
    pg.drawText('Documento generado automáticamente por el Sistema de Gestión de Compras.', {
      x: margen, y: 30, size: 7, font: fontRegular, color: rgb(0.55, 0.55, 0.6),
    })
    pg.drawText(`${idx + 1}/${paginas.length}`, { x: anchoPagina - margen - 25, y: 30, size: 7, font: fontRegular, color: rgb(0.55, 0.55, 0.6) })
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