// Prueba temporal — pégala dentro de tu componente App, dentro de un useEffect,
// solo para confirmar que la conexión funciona. Bórrala después.

import { useEffect } from 'react'
import { supabase } from './lib/supabaseClient'

useEffect(() => {
  supabase
    .from('empresas')
    .select('*')
    .then(({ data, error }) => {
      if (error) console.error('Error conectando a Supabase:', error.message)
      else console.log('Conexión exitosa. Empresas encontradas:', data)
    })
}, [])