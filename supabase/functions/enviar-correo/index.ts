import { serve } from "https://deno.land/std@0.168.0/http/server.ts"
import { SMTPClient } from "https://deno.land/x/denomailer/mod.ts"

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
}

const GMAIL_USER = Deno.env.get("spdique25@gmail.com")!
const GMAIL_APP_PASSWORD = Deno.env.get("qeix bvwo mtzd lanx")!

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders })
  }

  try {
    const { to, subject, html } = await req.json()
    const destinatarios = Array.isArray(to) ? to.filter(Boolean) : [to].filter(Boolean)

    if (!destinatarios.length) {
      return new Response(JSON.stringify({ error: "Sin destinatario" }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      })
    }

    const client = new SMTPClient({
      connection: {
        hostname: "smtp.gmail.com",
        port: 465,
        tls: true,
        auth: {
          username: GMAIL_USER,
          password: GMAIL_APP_PASSWORD,
        },
      },
    })

    await client.send({
      from: `Gestión de Compras <${GMAIL_USER}>`,
      to: destinatarios,
      subject,
      html,
    })

    await client.close()

    return new Response(JSON.stringify({ ok: true }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    })
  } catch (err) {
    console.error("Error enviando correo:", err)
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    })
  }
})