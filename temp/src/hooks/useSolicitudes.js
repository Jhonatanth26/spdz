import { useState, useEffect, useCallback } from 'react'
import { supabase } from '../lib/supabaseClient'

// Convierte una fila de Supabase (columnas sueltas + detalle jsonb)
// al objeto "solicitud" completo que usa el resto de la app.
const desdeFila = (fila) => ({
  id: fila.id,
  folio: fila.folio,
  tipo: fila.tipo,
  empresaId: fila.empresa_id,
  areaId: fila.area_id,
  departamentoId: fila.departamento_id,
  conceptoGastoId: fila.concepto_gasto_id,
  solicitanteId: fila.solicitante_id,
  fechaCreacion: fila.fecha_creacion,
  fechaEstimada: fila.fecha_estimada,
  objetivo: fila.objetivo,
  justificacion: fila.justificacion,
  status: fila.status,
  ...(fila.detalle || {}), // items, firmas, revisionCompras, pagosSugeridos, pagos,
  // pagosConfirmados, ocEnviada, recepcion, historialEstados, notificaciones
})

// Hace lo inverso: separa el objeto de la app en columnas + el bloque "detalle".
const haciaFila = (s) => ({
  folio: s.folio,
  tipo: s.tipo,
  empresa_id: s.empresaId,
  area_id: s.areaId,
  departamento_id: s.departamentoId || null,
  concepto_gasto_id: s.conceptoGastoId || null,
  solicitante_id: s.solicitanteId,
  fecha_creacion: s.fechaCreacion,
  fecha_estimada: s.fechaEstimada || null,
  objetivo: s.objetivo,
  justificacion: s.justificacion,
  status: s.status,
  detalle: {
    items: s.items,
    firmas: s.firmas,
    revisionCompras: s.revisionCompras,
    pagosSugeridos: s.pagosSugeridos,
    pagos: s.pagos,
    pagosConfirmados: s.pagosConfirmados,
    ocEnviada: s.ocEnviada,
    recepcion: s.recepcion,
    historialEstados: s.historialEstados,
    notificaciones: s.notificaciones,
    prioridad: s.prioridad,
    evaluacionProveedor: s.evaluacionProveedor,
    aiu: s.aiu,
    centroCosto: s.centroCosto,
    ultimoRechazo: s.ultimoRechazo,
    pagosConfirmadosPor: s.pagosConfirmadosPor,
    presupuestoAlEnviar: s.presupuestoAlEnviar,
  },
})

export function useSolicitudes() {
  const [solicitudes, setSolicitudes] = useState([])
  const [cargando, setCargando] = useState(true)

  const recargar = useCallback(async () => {
    const { data, error } = await supabase
      .from('solicitudes')
      .select('*')
      .order('created_at', { ascending: false })
    if (error) console.error('Error cargando solicitudes:', error.message)
    else setSolicitudes((data || []).map(desdeFila))
    setCargando(false)
  }, [])

  useEffect(() => { recargar() }, [recargar])

  const crear = async (nueva) => {
    // se pide de vuelta el id real (lo genera la base de datos): lo necesitan las notificaciones, que apuntan a la solicitud
    const { data, error } = await supabase.from('solicitudes').insert(haciaFila(nueva)).select('id').single()
    if (error) { console.error('Error creando solicitud:', error.message); return { error } }
    await recargar()
    return { id: data?.id }
  }

  const actualizar = async (sol) => {
    // ".select('id')" devuelve las filas que REALMENTE se modificaron. Cuando las reglas de seguridad de la base rechazan una
    // actualización no hay ningún error: la fila simplemente no se toca. Sin esta comprobación la app mostraba "aprobado" y al
    // recargar volvía el estado anterior, sin que nadie supiera por qué.
    const { data, error } = await supabase.from('solicitudes').update(haciaFila(sol)).eq('id', sol.id).select('id')
    if (error) { console.error('Error actualizando solicitud:', error.message); return error }
    if (!data || data.length === 0) {
      const sinPermiso = new Error('la base de datos no guardó el cambio: tu usuario no tiene permiso para modificar esta solicitud (regla de seguridad). Avísale al administrador.')
      console.error('Actualización rechazada (0 filas modificadas) en la solicitud', sol.id)
      await recargar() // la pantalla vuelve a lo que de verdad hay guardado
      return sinPermiso
    }
    await recargar()
  }

  const eliminar = async (id) => {
    const { data, error } = await supabase.from('solicitudes').delete().eq('id', id).select('id')
    if (error) { console.error('Error eliminando solicitud:', error.message); return error }
    if (!data || data.length === 0) { await recargar(); return new Error('la base de datos no eliminó la solicitud: tu usuario no tiene permiso (regla de seguridad).') }
    await recargar()
  }

  return { solicitudes, cargando, crear, actualizar, eliminar, recargar }
}