import { supabase } from './supabaseClient'
const MODO_DEMO = import.meta.env.VITE_MODO_DEMO === 'true'

// Llama a la Edge Function "enviar-correo" (la que ya desplegaste en Supabase).
// "to" puede ser un correo (string) o varios (array de strings).
// Si algo falla (sin destinatario, función caída, etc.) solo lo
// registra en consola — nunca debe romper el flujo de la solicitud.
export async function enviarCorreo(to, subject, html) {
  const sinDestinatario = !to || (Array.isArray(to) && to.length === 0)
  if (sinDestinatario) { console.warn('enviarCorreo: no hay destinatario, se omite el envío.'); return }
  // en la versión de prueba los correos solo se simulan: nunca se le escribe a nadie de verdad
  if (MODO_DEMO) { console.info('[DEMO] Correo simulado a', to, '—', subject); return }
  const { error } = await supabase.functions.invoke('enviar-correo', {
    body: { to, subject, html },
  })
  if (error) console.error('Error enviando correo real:', error.message)
}