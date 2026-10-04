import { serve } from "https://deno.land/std@0.168.0/http/server.ts"
import { createClient } from "https://esm.sh/@supabase/supabase-js@2"

// Crea (o actualiza) empleados en bloque: su acceso en Supabase Authentication — con el DOCUMENTO como contraseña —
// y su perfil en la tabla "usuarios", ya vinculados entre sí. Solo lo puede usar un usuario con rol Administrador.
// El documento llega aquí únicamente para fijar la contraseña inicial: NO se guarda en ninguna tabla.

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
}
const ROLES = ["Solicitante", "Jefe de Área", "Director de Área", "Jefe de Área y Director", "Compras", "Dirección Financiera", "Contabilidad", "Gerencia", "Administrador"]
const MAX_POR_LLAMADA = 50

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
const ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY")!

const json = (cuerpo: unknown, status = 200) =>
  new Response(JSON.stringify(cuerpo), { status, headers: { ...corsHeaders, "Content-Type": "application/json" } })

serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders })

  try {
    // 1) quién llama: debe tener sesión y ser Administrador
    const quien = createClient(SUPABASE_URL, ANON_KEY, { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } } })
    const { data: { user }, error: errSesion } = await quien.auth.getUser()
    if (errSesion || !user) return json({ error: "Sesión no válida. Vuelve a iniciar sesión." }, 401)

    const admin = createClient(SUPABASE_URL, SERVICE_KEY, { auth: { autoRefreshToken: false, persistSession: false } })
    const { data: perfil } = await admin.from("usuarios").select("rol").eq("auth_user_id", user.id).maybeSingle()
    if (perfil?.rol !== "Administrador") return json({ error: "Solo un Administrador puede importar empleados." }, 403)

    const { empleados } = await req.json()
    if (!Array.isArray(empleados) || !empleados.length) return json({ error: "No llegó ningún empleado." }, 400)
    if (empleados.length > MAX_POR_LLAMADA) return json({ error: `Máximo ${MAX_POR_LLAMADA} empleados por llamada.` }, 400)

    // 2) accesos que ya existen en Authentication (por correo), para no duplicarlos
    const authPorCorreo = new Map<string, string>()
    for (let pagina = 1; ; pagina++) {
      const { data, error } = await admin.auth.admin.listUsers({ page: pagina, perPage: 1000 })
      if (error) throw error
      data.users.forEach((u) => { if (u.email) authPorCorreo.set(u.email.toLowerCase(), u.id) })
      if (data.users.length < 1000) break
    }

    // 3) uno por uno — si uno falla, los demás siguen
    const resultados: { correo: string; ok: boolean; mensaje: string }[] = []
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
          const { data, error } = await admin.auth.admin.createUser({ email: correo, password: clave, email_confirm: true })
          if (error) throw error
          authId = data.user.id
          creoAcceso = true
          authPorCorreo.set(correo, authId)
        }

        // perfil en "usuarios" (se busca por correo: si ya existe, se actualiza)
        const campos: Record<string, unknown> = { nombre, cargo: e.cargo || null, auth_user_id: authId }
        if (e.areaId) campos.area_id = e.areaId
        if (Array.isArray(e.areasAdicionales)) campos.areas_adicionales = e.areasAdicionales
        const { data: existente, error: errBusca } = await admin.from("usuarios").select("id").eq("email", correo).maybeSingle()
        if (errBusca) throw errBusca

        if (existente) {
          if (e.rol) campos.rol = e.rol // si el archivo no trae rol, el que ya tenía no se toca
          const { error } = await admin.from("usuarios").update(campos).eq("id", existente.id)
          if (error) throw error
          resultados.push({ correo, ok: true, mensaje: creoAcceso ? "Perfil actualizado y acceso creado." : "Perfil actualizado (ya tenía acceso: no se cambió su contraseña)." })
        } else {
          const { error } = await admin.from("usuarios").insert({ ...campos, email: correo, rol: e.rol || "Solicitante" })
          if (error) throw error
          resultados.push({ correo, ok: true, mensaje: creoAcceso ? "Creado con acceso." : "Perfil creado (su acceso ya existía en Authentication)." })
        }
      } catch (err) {
        resultados.push({ correo, ok: false, mensaje: (err as Error)?.message ?? String(err) })
      }
    }

    return json({ resultados })
  } catch (err) {
    console.error("Error importando empleados:", err)
    return json({ error: (err as Error)?.message ?? String(err) }, 500)
  }
})