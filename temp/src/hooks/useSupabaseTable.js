import { useState, useEffect, useCallback, useRef } from 'react'
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
  // "cargando" solo vale TRUE la primera vez. Las recargas posteriores (después de guardar, editar o eliminar) se
  // hacen "en silencio": la app usa "cargando" para reemplazar TODA la pantalla por "Cargando...", y eso desmontaba
  // el catálogo en el que estabas y te devolvía al primero cada vez que guardabas algo.
  const yaCargoUnaVez = useRef(false)

  const recargar = useCallback(async () => {
    if (!yaCargoUnaVez.current) setCargando(true)
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
    yaCargoUnaVez.current = true
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

  // guarda muchos registros de una vez (actualiza los que ya existen y crea los nuevos) y recarga UNA sola vez al final,
  // en vez de una recarga por cada registro. Devuelve, por cada registro, { registro, error } (error = null si salió bien).
  const guardarLote = async (registros, { enParalelo = 8 } = {}) => {
    const resultados = new Array(registros.length)
    for (let i = 0; i < registros.length; i += enParalelo) {
      await Promise.all(registros.slice(i, i + enParalelo).map(async (registro, k) => {
        const esNuevo = !datos.some((d) => d.id === registro.id)
        const { id, ...resto } = haciaDb(registro)
        const { error } = esNuevo
          ? await supabase.from(tabla).insert(resto)
          : await supabase.from(tabla).update(resto).eq('id', registro.id)
        if (error) console.error(`Error guardando en "${tabla}":`, error.message)
        resultados[i + k] = { registro, error: error || null }
      }))
    }
    await recargar()
    return resultados
  }

  const eliminar = async (id) => {
    const { error } = await supabase.from(tabla).delete().eq('id', id)
    if (error) { console.error(`Error eliminando de "${tabla}":`, error.message); return error }
    await recargar()
  }

  return { datos, cargando, error, guardar, guardarVarios, guardarLote, eliminar, recargar }
}