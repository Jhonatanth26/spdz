import { supabase } from './supabaseClient'

// Sube un archivo al bucket privado "adjuntos" y devuelve la RUTA guardada
// (ya no una URL pública — el bucket ahora es privado por seguridad).
export async function subirArchivo(file, carpeta = 'general') {
  const nombreLimpio = file.name.replace(/[^a-zA-Z0-9._-]/g, '_')
  const ruta = `${carpeta}/${Date.now()}_${nombreLimpio}`

  const { error } = await supabase.storage.from('adjuntos').upload(ruta, file, { upsert: false })
  if (error) {
    console.error('Error subiendo archivo a Storage:', error.message)
    return null
  }
  return ruta
}

// Genera una URL temporal (por defecto 1 hora) para ver/descargar un archivo privado.
// Se debe llamar cada vez que se necesite mostrar o abrir el archivo — no se guarda
// la URL firmada en la base de datos porque expira.
export async function obtenerUrlFirmada(ruta, expiraSegundos = 3600) {
  if (!ruta) return null
  const { data, error } = await supabase.storage.from('adjuntos').createSignedUrl(ruta, expiraSegundos)
  if (error) {
    console.error('Error generando URL firmada:', error.message)
    return null
  }
  return data.signedUrl
}

// Límite de tamaño para archivos que se suben desde la app (10 MB)
export const TAMANO_MAXIMO_MB = 10
export function archivoDentroDelLimite(file) {
  return file.size <= TAMANO_MAXIMO_MB * 1024 * 1024
}