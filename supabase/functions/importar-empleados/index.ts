// @ts-nocheck
// Crea (o actualiza) empleados en bloque: su acceso en Supabase Authentication — con el DOCUMENTO como contraseña —
// y su perfil en la tabla "usuarios", ya vinculados entre sí. Solo lo puede usar un usuario con rol Administrador.
// El documento llega aquí únicamente para fijar la contraseña inicial: NO se guarda en ninguna tabla.
// No importa ninguna librería externa: habla directamente con la API de Supabase usando fetch.

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
}
const ROLES = ["Solicitante", "Jefe de Área", "Director de Área", "Jefe de Área y Director", "Compras", "Dirección Financiera", "Contabilidad", "Gerencia", "Administrador"]
const MAX_POR_LLAMADA = 50

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")
const ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY")

const json = (cuerpo, status = 200) =>
  new Response(JSON.stringify(cuerpo), { status, headers: { ...corsHeaders, "Content-Type": "application/json" } })

async function llamar(ruta, { metodo = "GET", llave = SERVICE_KEY, token, cuerpo, extra = {} } = {}) {
  const res = await fetch(`${SUPABASE_URL}${ruta}`, {
    method: metodo,
    headers: {
      apikey: llave,
      Authorization: `Bearer ${token ?? llave}`,
      ...(cuerpo !== undefined ? { "Content-Type": "application/json" } : {}),
      ...extra,
    },
    body: cuerpo !== undefined ? JSON.stringify(cuerpo) : undefined,
  })
  const texto = await res.text()
  let datos = null
  try { datos = texto ? JSON.parse(texto) : null } catch { datos = texto }
  if (!res.ok) {
    const msg = (datos && typeof datos === "object" && (datos.msg || datos.message || datos.error_description || datos.error))
      || (typeof datos === "string" && datos) || `Error ${res.status}`
    throw new Error(String(msg))
  }
  return datos
}

const SESION_INVALIDA = "Sesión no válida. Vuelve a iniciar sesión."

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders })

  try {
    // 1) quién llama: debe tener sesión y ser Administrador
    const token = (req.headers.get("Authorization") ?? "").replace(/^Bearer\s+/i, "").trim()
    if (!token) return json({ error: SESION_INVALIDA }, 401)
    let quien
    try { quien = await llamar("/auth/v1/user", { llave: ANON_KEY, token }) } catch { return json({ error: SESION_INVALIDA }, 401) }
    if (!quien?.id) return json({ error: SESION_INVALIDA }, 401)

    const perfiles = await llamar(`/rest/v1/usuarios?select=rol&auth_user_id=eq.${encodeURIComponent(quien.id)}&limit=1`)
    if (perfiles?.[0]?.rol !== "Administrador") return json({ error: "Solo un Administrador puede importar empleados." }, 403)

    let entrada
    try { entrada = await req.json() } catch { return json({ error: "La petición no trae un cuerpo válido." }, 400) }
    const empleados = entrada?.empleados
    if (!Array.isArray(empleados) || !empleados.length) return json({ error: "No llegó ningún empleado." }, 400)
    if (empleados.length > MAX_POR_LLAMADA) return json({ error: `Máximo ${MAX_POR_LLAMADA} empleados por llamada.` }, 400)

    // 2) accesos que ya existen en Authentication (por correo), para no duplicarlos
    const authPorCorreo = new Map()
    for (let pagina = 1; ; pagina++) {
      const r = await llamar(`/auth/v1/admin/users?page=${pagina}&per_page=1000`)
      const lista = r?.users ?? []
      lista.forEach((u) => { if (u.email) authPorCorreo.set(u.email.toLowerCase(), u.id) })
      if (lista.length < 1000) break
    }

    // 3) uno por uno — si uno falla, los demás siguen
    const resultados = []
    for (const e of empleados) {
      const correo = String(e.correo ?? "").trim().toLowerCase()
      try {
        const nombre = String(e.nombre ?? "").trim()
        const clave = String(e.documento ?? "").trim()
        if (!correo || !nombre) throw new Error("Falta el nombre o el correo.")
        if (e.rol && !ROLES.includes(e.rol)) throw new Error(`Rol no válido: ${e.rol}`)

        // acceso en Authentication
        let authId = authPorCorreo.get(correo)
        let creoAcceso = false
        if (!authId) {
          if (clave.length < 6) throw new Error("El documento (contraseña) debe tener al menos 6 caracteres.")
          const creado = await llamar("/auth/v1/admin/users", { metodo: "POST", cuerpo: { email: correo, password: clave, email_confirm: true } })
          authId = creado?.user?.id ?? creado?.id
          if (!authId) throw new Error("Supabase no devolvió el id del acceso creado.")
          creoAcceso = true
          authPorCorreo.set(correo, authId)
        }

        // perfil en "usuarios": se busca por correo sin distinguir mayúsculas (si ya existe, se actualiza)
        const campos = { nombre, cargo: e.cargo || null, auth_user_id: authId }
        if (e.areaId) campos.area_id = e.areaId
        if (Array.isArray(e.areasAdicionales)) campos.areas_adicionales = e.areasAdicionales
        const candidatos = await llamar(`/rest/v1/usuarios?select=id,email&email=ilike.${encodeURIComponent(correo)}`)
        const existente = (candidatos ?? []).find((f) => String(f.email ?? "").toLowerCase() === correo)

        if (existente) {
          if (e.rol) campos.rol = e.rol // si el archivo no trae rol, el que ya tenía no se toca
          await llamar(`/rest/v1/usuarios?id=eq.${encodeURIComponent(existente.id)}`, { metodo: "PATCH", cuerpo: campos, extra: { Prefer: "return=minimal" } })
          resultados.push({ correo, ok: true, mensaje: creoAcceso ? "Perfil actualizado y acceso creado." : "Perfil actualizado (ya tenía acceso: no se cambió su contraseña)." })
        } else {
          await llamar("/rest/v1/usuarios", { metodo: "POST", cuerpo: { ...campos, email: correo, rol: e.rol || "Solicitante" }, extra: { Prefer: "return=minimal" } })
          resultados.push({ correo, ok: true, mensaje: creoAcceso ? "Creado con acceso." : "Perfil creado (su acceso ya existía en Authentication)." })
        }
      } catch (err) {
        resultados.push({ correo, ok: false, mensaje: err?.message ?? String(err) })
      }
    }

    return json({ resultados })
  } catch (err) {
    console.error("Error importando empleados:", err)
    return json({ error: err?.message ?? String(err) }, 500)
  }
})
