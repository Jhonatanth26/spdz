import { useState, useEffect } from 'react'
import { supabase } from '../lib/supabaseClient'

// Maneja la sesión de Supabase Auth y trae el perfil (nombre, rol, área...)
// desde la tabla "usuarios" una vez que la persona inicia sesión.
export function useAuth() {
  const [session, setSession] = useState(null)
  const [perfil, setPerfil] = useState(null)
  const [cargando, setCargando] = useState(true)

  const cargarPerfil = async (authUserId) => {
    const { data, error } = await supabase
      .from('usuarios')
      .select('*, area:areas(nombre)')
      .eq('auth_user_id', authUserId)
      .single()

    if (error) {
      console.error('No se encontró un perfil en "usuarios" para este login:', error.message)
      setPerfil(null)
    } else {
      setPerfil(data)
    }
  }

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session)
      if (session) cargarPerfil(session.user.id).finally(() => setCargando(false))
      else setCargando(false)
    })

    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session)
      if (session) cargarPerfil(session.user.id)
      else setPerfil(null)
    })

    return () => listener.subscription.unsubscribe()
  }, [])

  const iniciarSesion = async (email, password) => {
    const { error } = await supabase.auth.signInWithPassword({ email, password })
    return error
  }

  const cerrarSesion = () => supabase.auth.signOut()

  // Actualiza campos del propio perfil (ej. la foto de firma en "Mi perfil")
  // y refresca el perfil en memoria para que se refleje de inmediato.
  const actualizarPerfil = async (cambios) => {
    if (!session) return
    const { error } = await supabase
      .from('usuarios')
      .update(cambios)
      .eq('auth_user_id', session.user.id)
    if (error) { console.error('Error actualizando el perfil:', error.message); return error }
    await cargarPerfil(session.user.id)
  }

  return { session, perfil, cargando, iniciarSesion, cerrarSesion, actualizarPerfil }
}