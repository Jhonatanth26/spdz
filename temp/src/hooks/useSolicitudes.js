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
  departamentoId: fila.departamento_id,
  centroCostoId: fila.centro_costo_id,
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
  departamento_id: s.departamentoId || null,
  centro_costo_id: s.centroCostoId || null,
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
    const { error } = await supabase.from('solicitudes').insert(haciaFila(nueva))
    if (error) { console.error('Error creando solicitud:', error.message); return error }
    await recargar()
  }

  const actualizar = async (sol) => {
    const { error } = await supabase.from('solicitudes').update(haciaFila(sol)).eq('id', sol.id)
    if (error) { console.error('Error actualizando solicitud:', error.message); return error }
    await recargar()
  }

  return { solicitudes, cargando, crear, actualizar, recargar }
}