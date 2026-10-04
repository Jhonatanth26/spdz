// Exportar las áreas (con su director y presupuesto) a Excel, editarlas allá y volver a subirlas para actualizar todo
// de una vez. Aquí solo está la lógica de armar las hojas y de validar el archivo que se vuelve a subir; la lectura y
// escritura del archivo (Excel/CSV) la hace la pantalla.

// Roles de quienes pueden ser el director responsable de un área. Además de los directores de área, la cabeza de una
// dirección puede tener otro rol: el gerente general (Gerencia) dirige el área "Gerencia General", y quien lleva la
// Dirección Financiera tiene el rol Dirección Financiera.
export const ROLES_ASIGNABLES_DIRECTOR = ['Director de Área', 'Jefe de Área y Director', 'Dirección Financiera', 'Gerencia']
export const ROLES_DIRECTOR = ROLES_ASIGNABLES_DIRECTOR
export const COLUMNAS_AREAS = ['Área', 'Presupuesto mensual', 'Director responsable']
// palabras que, escritas en la celda del director, significan "quitar el director de esta área"
const PALABRAS_QUITAR = /^(quitar|sin director|ninguno|ninguna|n\/a|-)$/i

const norm = (s) => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase()

function valorDe(fila, ...nombres) {
  const claves = Object.keys(fila)
  for (const n of nombres) {
    const k = claves.find((c) => norm(c) === norm(n))
    if (k !== undefined) return String(fila[k] ?? '').trim()
  }
  return ''
}

// ---------- lo que se exporta ----------
export function filasExportarAreas(areas, usuarios) {
  const porId = new Map(usuarios.map((u) => [u.id, u]))
  return [COLUMNAS_AREAS, ...areas.map((a) => [a.nombre, Number(a.presupuesto) || 0, porId.get(a.directorId)?.nombre || ''])]
}

// lista de apoyo: a quién se puede poner como director (con su correo, para no confundir homónimos) y qué dirige hoy
export function filasDirectoresValidos(areas, usuarios) {
  const directores = usuarios.filter((u) => ROLES_DIRECTOR.includes(u.rol)).sort((a, b) => String(a.nombre).localeCompare(String(b.nombre), 'es'))
  return [['Nombre', 'Correo', 'Rol', 'Dirige hoy'], ...directores.map((u) => [u.nombre, u.email || '', u.rol, areas.filter((a) => a.directorId === u.id).map((a) => a.nombre).join('; ')])]
}

export const filasInstruccionesAreas = [
  ['Cómo actualizar las áreas'],
  ['• Cambia solo las columnas "Presupuesto mensual" y "Director responsable". El nombre del área NO se cambia aquí: sirve para saber a qué área corresponde cada fila.'],
  ['• Director responsable: escribe el nombre tal cual aparece en la hoja "Directores válidos" (puedes copiarlo y pegarlo). Solo se pueden asignar personas con rol de Director de Área, Jefe de Área y Director, Dirección Financiera o Gerencia.'],
  ['• Si hay dos directores con el mismo nombre, agrega una columna llamada "Correo del director" y pon ahí el correo.'],
  ['• Para QUITAR el director de un área, escribe QUITAR en esa celda.'],
  ['• Si dejas una celda vacía, ese dato NO se cambia.'],
  ['• Una misma persona puede ser director de varias áreas.'],
  ['• Esta carga no crea áreas nuevas. Si el nombre del área no existe, la fila saldrá con error (para crear áreas nuevas usa "Importar CSV").'],
]

// ---------- presupuesto: "50000000", "50.000.000", "$ 50,000,000", "1.500,50" ----------
export function leerPresupuesto(txt) {
  const t = String(txt ?? '').replace(/[$\s]/g, '').replace(/cop/i, '')
  if (!t) return { vacio: true }
  let n
  if (/^\d+$/.test(t)) n = Number(t)
  else if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) n = Number(t.replace(/\./g, '').replace(',', '.'))      // 50.000.000 ó 1.500,50
  else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) n = Number(t.replace(/,/g, ''))                           // 50,000,000 ó 1,500.50
  else if (/^\d+[.,]\d{1,2}$/.test(t)) n = Number(t.replace(',', '.'))                                      // 1500,5
  else return { error: true }
  return Number.isFinite(n) && n >= 0 ? { valor: n } : { error: true }
}

// ---------- lo que se vuelve a subir ----------
// Devuelve una fila validada por cada fila del archivo: qué cambiaría o por qué no se puede. No toca nada todavía.
export function validarAreas(filasCrudas, { areas = [], usuarios = [] } = {}) {
  const directores = usuarios.filter((u) => ROLES_DIRECTOR.includes(u.rol))
  const areaPorNombre = new Map(areas.map((a) => [norm(a.nombre), a]))
  const dirPorId = new Map(usuarios.map((u) => [u.id, u]))
  const vistas = new Map()

  return filasCrudas.map((f, i) => {
    const fila = i + 2 // fila real en el Excel (la 1 es el encabezado)
    const errores = []
    const nombre = valorDe(f, 'Área', 'Area', 'Nombre', 'Nombre del área', 'Nombre del area')
    const presTxt = valorDe(f, 'Presupuesto mensual', 'Presupuesto')
    const dirTxt = valorDe(f, 'Director responsable', 'Director', 'Director del área', 'Director del area')
    const correoTxt = valorDe(f, 'Correo del director', 'Correo director', 'Correo').toLowerCase()

    // ¿qué área es?
    let area = null
    if (!nombre) errores.push('Falta el nombre del área')
    else {
      area = areaPorNombre.get(norm(nombre)) || null
      if (!area) errores.push(`El área «${nombre}» no existe en el catálogo (esta carga no crea áreas nuevas)`)
      else if (vistas.has(area.id)) errores.push(`El área está repetida en el archivo (ya aparece en la fila ${vistas.get(area.id)})`)
      else vistas.set(area.id, fila)
    }

    // presupuesto
    let presupuesto
    if (presTxt) {
      const p = leerPresupuesto(presTxt)
      if (p.error) errores.push(`Presupuesto no válido: «${presTxt}»`)
      else if (!p.vacio) presupuesto = p.valor
    }

    // director: undefined = no cambia; '' = quitar; id = asignar
    let directorId
    if (dirTxt && PALABRAS_QUITAR.test(dirTxt) && !correoTxt) directorId = ''
    else if (dirTxt || correoTxt) {
      const porCorreo = correoTxt ? usuarios.find((u) => String(u.email || '').toLowerCase() === correoTxt) : null
      if (correoTxt && !porCorreo) errores.push(`No hay ningún usuario con el correo «${correoTxt}»`)
      else if (correoTxt && !ROLES_DIRECTOR.includes(porCorreo.rol)) errores.push(`${porCorreo.nombre} no puede ser director de un área (su rol es ${porCorreo.rol})`)
      else if (correoTxt && dirTxt && norm(porCorreo.nombre) !== norm(dirTxt)) errores.push(`El nombre «${dirTxt}» y el correo «${correoTxt}» no son de la misma persona (el correo es de ${porCorreo.nombre})`)
      else if (correoTxt) directorId = porCorreo.id
      else {
        const coinciden = directores.filter((u) => norm(u.nombre) === norm(dirTxt))
        if (coinciden.length === 1) directorId = coinciden[0].id
        else if (coinciden.length > 1) errores.push(`Hay ${coinciden.length} directores llamados «${dirTxt}»: pon el correo en la columna "Correo del director"`)
        else {
          const noDirector = usuarios.find((u) => norm(u.nombre) === norm(dirTxt))
          errores.push(noDirector
            ? `${noDirector.nombre} existe pero no puede ser director de un área (su rol es ${noDirector.rol})`
            : `No hay ningún director llamado «${dirTxt}» (revisa la hoja "Directores válidos")`)
        }
      }
    }

    // ¿qué cambiaría?
    const cambios = []
    if (area && !errores.length) {
      if (presupuesto !== undefined && presupuesto !== (Number(area.presupuesto) || 0)) cambios.push(`Presupuesto: ${(Number(area.presupuesto) || 0).toLocaleString('es-CO')} → ${presupuesto.toLocaleString('es-CO')}`)
      if (directorId !== undefined && directorId !== (area.directorId || '')) {
        const antes = dirPorId.get(area.directorId)?.nombre || 'sin director'
        const despues = directorId ? dirPorId.get(directorId)?.nombre : 'sin director'
        cambios.push(`Director: ${antes} → ${despues}`)
      }
    }
    return { fila, nombre, area, presupuesto, directorId, cambios, errores, estado: errores.length ? 'error' : cambios.length ? 'actualiza' : 'sin_cambios' }
  })
}

// los registros listos para guardar: el área tal como está, con solo lo que cambió
export function registrosParaGuardarAreas(filasValidadas) {
  return filasValidadas.filter((f) => f.estado === 'actualiza').map((f) => ({
    ...f.area,
    ...(f.presupuesto !== undefined ? { presupuesto: f.presupuesto } : {}),
    ...(f.directorId !== undefined ? { directorId: f.directorId } : {}),
  }))
}