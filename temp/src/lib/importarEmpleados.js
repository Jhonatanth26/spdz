import { supabase } from './supabaseClient'

// Carga masiva de empleados: cada persona queda con su perfil en la tabla "usuarios" Y con su acceso creado en
// Supabase Authentication, usando su DOCUMENTO como contraseña inicial. Todo lo que toca Authentication lo hace la
// Edge Function "importar-empleados" (solo la puede usar un Administrador); aquí solo se arma y valida la lista.

export const COLUMNAS_EMPLEADOS = ['Nombre', 'Documento', 'Correo', 'Cargo', 'Área', 'Rol', 'Áreas adicionales']
export const CLAVE_MINIMA = 6 // largo mínimo de contraseña que acepta Supabase por defecto

const sinTildes = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase()

// El documento es la contraseña: se limpian puntos, comas, espacios y guiones (1.017.234.567 -> 1017234567)
export const documentoComoClave = (doc) => String(doc ?? '').trim().replace(/[.,\s\-]/g, '')

// Busca una columna por cualquiera de sus nombres posibles, sin importar mayúsculas ni tildes
function valorDe(fila, ...nombres) {
  const claves = Object.keys(fila)
  for (const n of nombres) {
    const k = claves.find((c) => sinTildes(c) === sinTildes(n))
    if (k !== undefined) return String(fila[k] ?? '').trim()
  }
  return ''
}

// filasCrudas: lo que sale de leer el Excel/CSV (un objeto por fila, con los encabezados del archivo)
// opciones: { areas: [{id, nombre}], roles: [string], correosExistentes: Set<string> (en minúsculas) }
// Devuelve una fila validada por cada fila del archivo, sin tocar nada todavía.
export function validarEmpleados(filasCrudas, { areas = [], roles = [], correosExistentes = new Set() } = {}) {
  const areaPorNombre = new Map(areas.map((a) => [sinTildes(a.nombre), a]))
  const rolPorNombre = new Map(roles.map((r) => [sinTildes(r), r]))
  const correoVisto = new Map() // correo -> número de la primera fila donde apareció
  const docVisto = new Map()

  return filasCrudas.map((f, i) => {
    const fila = i + 2 // fila real en el Excel (la 1 es el encabezado)
    const errores = []
    const nombre = valorDe(f, 'Nombre', 'Nombre completo', 'Empleado', 'Nombres y apellidos')
    const documento = documentoComoClave(valorDe(f, 'Documento', 'Cédula', 'Cedula', 'Identificación', 'Identificacion', 'No. documento', 'Numero de documento', 'Número de documento'))
    const correo = valorDe(f, 'Correo', 'Email', 'E-mail', 'Correo electrónico', 'Correo electronico').toLowerCase()
    const cargo = valorDe(f, 'Cargo')
    const areaTxt = valorDe(f, 'Área', 'Area', 'Área principal', 'Area principal')
    const rolTxt = valorDe(f, 'Rol')
    const adicionalesTxt = valorDe(f, 'Áreas adicionales', 'Areas adicionales', 'Áreas adicionales a cargo')

    if (!nombre) errores.push('Falta el nombre')
    if (!correo) errores.push('Falta el correo')
    else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) errores.push('Correo inválido')
    else if (correoVisto.has(correo)) errores.push(`Correo repetido en el archivo (también en la fila ${correoVisto.get(correo)})`)
    else correoVisto.set(correo, fila)

    if (!documento) errores.push('Falta el documento')
    else if (documento.length < CLAVE_MINIMA) errores.push(`El documento tiene menos de ${CLAVE_MINIMA} caracteres (no alcanza para ser contraseña)`)
    else if (docVisto.has(documento)) errores.push(`Documento repetido en el archivo (también en la fila ${docVisto.get(documento)})`)
    else docVisto.set(documento, fila)

    let areaId = null
    if (areaTxt) {
      const a = areaPorNombre.get(sinTildes(areaTxt))
      if (a) areaId = a.id
      else errores.push(`El área «${areaTxt}» no existe en el catálogo`)
    }

    let rol = null // null = no se indicó: a una persona nueva se le pone Solicitante, y a una existente no se le toca el rol
    if (rolTxt) {
      rol = rolPorNombre.get(sinTildes(rolTxt)) || null
      if (!rol) errores.push(`El rol «${rolTxt}» no es válido`)
    }

    const areasAdicionales = []
    if (adicionalesTxt) {
      adicionalesTxt.split(/[;,|]/).map((x) => x.trim()).filter(Boolean).forEach((txt) => {
        const a = areaPorNombre.get(sinTildes(txt))
        if (a) areasAdicionales.push(a.id)
        else errores.push(`El área adicional «${txt}» no existe en el catálogo`)
      })
    }

    const existe = correo && correosExistentes.has(correo)
    return {
      fila, nombre, documento, correo, cargo, areaId, areaTxt, rol, rolTxt, areasAdicionales,
      existe: !!existe, errores,
      estado: errores.length ? 'error' : existe ? 'actualiza' : 'nuevo',
    }
  })
}

// Manda los empleados válidos a la Edge Function, por lotes (para no pasarse del tiempo máximo de una llamada).
// Devuelve un resultado por empleado: { correo, ok, mensaje }
export async function importarEmpleadosEnLotes(empleados, { tamLote = 20, onProgreso } = {}) {
  const resultados = []
  for (let i = 0; i < empleados.length; i += tamLote) {
    const lote = empleados.slice(i, i + tamLote).map((e) => ({
      nombre: e.nombre, documento: e.documento, correo: e.correo, cargo: e.cargo,
      areaId: e.areaId, areasAdicionales: e.areasAdicionales, rol: e.rol,
    }))
    const { data, error } = await supabase.functions.invoke('importar-empleados', { body: { empleados: lote } })
    let mensaje = null
    if (error) {
      mensaje = error.message
      try { const cuerpo = await error.context?.json?.(); if (cuerpo?.error) mensaje = cuerpo.error } catch { /* se queda el mensaje genérico */ }
    } else if (data?.error) mensaje = data.error
    if (mensaje) lote.forEach((e) => resultados.push({ correo: e.correo, ok: false, mensaje }))
    else resultados.push(...(data?.resultados || []))
    onProgreso?.(Math.min(i + tamLote, empleados.length), empleados.length)
  }
  return resultados
}