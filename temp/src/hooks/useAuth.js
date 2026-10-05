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
      // Solo "*": la app no usa el nombre del área del perfil (usa area_id). Pedirlo con un "join" (areas(nombre)) se
      // rompe en cuanto existe más de una relación entre "usuarios" y "areas" — por ejemplo areas.director_id — y
      // entonces nadie puede iniciar sesión, porque el perfil no se logra cargar.
      .select('*')
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

  // Cambia la contraseña de quien tiene la sesión iniciada (desde "Mi perfil"). Devuelve el error, si lo hubo.
  const cambiarContrasena = async (nueva) => {
    const { error } = await supabase.auth.updateUser({ password: nueva })
    return error
  }

  // Guarda la paleta de colores de quien tiene la sesión iniciada. Va en los datos de la propia cuenta (user_metadata), así que
  // la acompaña a cualquier computador o celular y no hace falta ninguna columna ni ninguna política en la base de datos.
  const guardarTema = async (tema) => {
    const { error } = await supabase.auth.updateUser({ data: { tema } })
    if (error) console.error('No se pudo guardar la paleta de colores:', error.message)
    return error
  }

  return { session, perfil, cargando, iniciarSesion, cerrarSesion, actualizarPerfil, cambiarContrasena, guardarTema }
}