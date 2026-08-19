import { useState, useEffect, useCallback } from 'react'
import { supabase } from '../lib/supabaseClient'

// Hook genérico para leer y guardar cualquier tabla de catálogo en Supabase.
//
// tabla: nombre de la tabla en Supabase (ej. "empresas")
// desdeDb: convierte una fila de la BD (snake_case) al formato que usa la app (camelCase)
// haciaDb: hace lo inverso, antes de guardar
// orderBy: columna por la que ordenar (opcional)
export function useSupabaseTable(tabla, { desdeDb = (r) => r, haciaDb = (r) => r, orderBy } = {}) {
  const [datos, setDatos] = useState([])
  const [cargando, setCargando] = useState(true)
  const [error, setError] = useState(null)

  const recargar = useCallback(async () => {
    setCargando(true)
    let query = supabase.from(tabla).select('*')
    if (orderBy) query = query.order(orderBy)
    const { data, error } = await query
    if (error) {
      console.error(`Error cargando "${tabla}":`, error.message)
      setError(error.message)
    } else {
      setError(null)
      setDatos((data || []).map(desdeDb))
    }
    setCargando(false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tabla])

  useEffect(() => { recargar() }, [recargar])

  // crea o actualiza un registro. Si trae un id que ya existe localmente, actualiza; si no, inserta.
  // Devuelve el registro guardado (con su id real de la base de datos) o un error.
  const guardar = async (registro) => {
    const esNuevo = !datos.some((d) => d.id === registro.id)
    const payload = haciaDb(registro)

    if (esNuevo) {
      const { id, ...sinId } = payload // dejamos que Supabase genere el UUID
      const { data, error } = await supabase.from(tabla).insert(sinId).select().single()
      if (error) { console.error(`Error creando en "${tabla}":`, error.message); return error }
      await recargar()
      return desdeDb(data)
    } else {
      const { id, ...cambios } = payload
      const { data, error } = await supabase.from(tabla).update(cambios).eq('id', registro.id).select().single()
      if (error) { console.error(`Error actualizando "${tabla}":`, error.message); return error }
      await recargar()
      return desdeDb(data)
    }
  }

  // importar varias filas de una vez (usado por "Importar CSV")
  const guardarVarios = async (registros) => {
    const payload = registros.map((r) => { const { id, ...sinId } = haciaDb(r); return sinId })
    const { error } = await supabase.from(tabla).insert(payload)
    if (error) { console.error(`Error importando a "${tabla}":`, error.message); return error }
    await recargar()
  }

  const eliminar = async (id) => {
    const { error } = await supabase.from(tabla).delete().eq('id', id)
    if (error) { console.error(`Error eliminando de "${tabla}":`, error.message); return error }
    await recargar()
  }

  return { datos, cargando, error, guardar, guardarVarios, eliminar, recargar }
}