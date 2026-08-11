import { supabase } from './supabaseClient'

// Llama a la Edge Function "enviar-correo" (la que ya desplegaste en Supabase).
// Si algo falla (sin correo del destinatario, función caída, etc.) solo lo
// registra en consola — nunca debe romper el flujo de la solicitud.
export async function enviarCorreo(to, subject, html) {
  if (!to) { console.warn('enviarCorreo: no hay destinatario, se omite el envío.'); return }
  const { error } = await supabase.functions.invoke('enviar-correo', {
    body: { to, subject, html },
  })
  if (error) console.error('Error enviando correo real:', error.message)
}