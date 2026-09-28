import { useState, useEffect, useCallback, useRef } from 'react'
import { supabase } from '../lib/supabaseClient'

// Notificaciones dentro de la app (campanita) — separadas del log interno de cada
// solicitud. Se refrescan solas cada 15s (más simple y confiable que configurar
// Supabase Realtime para este caso de uso).
export function useNotificaciones(usuarioId) {
  const [notificaciones, setNotificaciones] = useState([])
  const [cargando, setCargando] = useState(true)
  const timerRef = useRef(null)

  const cargar = useCallback(async () => {
    if (!usuarioId) { setNotificaciones([]); setCargando(false); return }
    const { data, error } = await supabase
      .from('notificaciones_usuario')
      .select('*')
      .eq('usuario_id', usuarioId)
      .order('creado_en', { ascending: false })
      .limit(50)
    if (error) { console.error('Error cargando notificaciones:', error.message); setCargando(false); return }
    setNotificaciones((data || []).map((r) => ({
      id: r.id, usuarioId: r.usuario_id, mensaje: r.mensaje, solicitudId: r.solicitud_id,
      leida: r.leida, creadoEn: r.creado_en,
    })))
    setCargando(false)
  }, [usuarioId])

  useEffect(() => {
    cargar()
    timerRef.current = setInterval(cargar, 15000)
    return () => clearInterval(timerRef.current)
  }, [cargar])

  // crea una notificación PARA OTRO usuario (ej. al aprobar, se le crea al siguiente responsable)
  const crear = async (destinatarioId, mensaje, solicitudId = null) => {
    if (!destinatarioId) return
    const { error } = await supabase.from('notificaciones_usuario').insert({
      usuario_id: destinatarioId, mensaje, solicitud_id: solicitudId, leida: false,
    })
    if (error) console.error('Error creando notificación:', error.message)
  }

  const marcarLeida = async (id) => {
    setNotificaciones((prev) => prev.map((n) => (n.id === id ? { ...n, leida: true } : n)))
    const { error } = await supabase.from('notificaciones_usuario').update({ leida: true }).eq('id', id)
    if (error) console.error('Error marcando notificación como leída:', error.message)
  }

  const marcarTodasLeidas = async () => {
    const sinLeer = notificaciones.filter((n) => !n.leida).map((n) => n.id)
    if (!sinLeer.length) return
    setNotificaciones((prev) => prev.map((n) => ({ ...n, leida: true })))
    const { error } = await supabase.from('notificaciones_usuario').update({ leida: true }).in('id', sinLeer)
    if (error) console.error('Error marcando todas como leídas:', error.message)
  }

  const eliminar = async (id) => {
    setNotificaciones((prev) => prev.filter((n) => n.id !== id))
    const { error } = await supabase.from('notificaciones_usuario').delete().eq('id', id)
    if (error) console.error('Error eliminando notificación:', error.message)
  }

  return { notificaciones, cargando, crear, marcarLeida, marcarTodasLeidas, eliminar, recargar: cargar }
}