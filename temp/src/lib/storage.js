import { supabase } from './supabaseClient'

// Sube un archivo al bucket "adjuntos" de Supabase Storage y devuelve su URL pública.
// carpeta: subcarpeta dentro del bucket, para organizar (ej. "logos", "firmas", "cotizaciones").
export async function subirArchivo(file, carpeta = 'general') {
  const nombreLimpio = file.name.replace(/[^a-zA-Z0-9._-]/g, '_')
  const ruta = `${carpeta}/${Date.now()}_${nombreLimpio}`

  const { error } = await supabase.storage.from('adjuntos').upload(ruta, file, { upsert: false })
  if (error) {
    console.error('Error subiendo archivo a Storage:', error.message)
    return null
  }

  const { data } = supabase.storage.from('adjuntos').getPublicUrl(ruta)
  return data.publicUrl
}