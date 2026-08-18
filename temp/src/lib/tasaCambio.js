// Trae la tasa de cambio del día (moneda → COP) usando un servicio gratuito,
// sin necesidad de clave ni backend propio.
export async function obtenerTasaCambioCOP(moneda) {
  if (!moneda || moneda === 'COP') return 1
  try {
    const resp = await fetch('https://open.er-api.com/v6/latest/USD')
    const data = await resp.json()
    if (data.result !== 'success' || !data.rates?.COP) return null

    if (moneda === 'USD') return data.rates.COP
    if (!data.rates[moneda]) return null
    // cuántos COP equivalen a 1 unidad de esa moneda (vía cruce con USD)
    return data.rates.COP / data.rates[moneda]
  } catch (err) {
    console.error('Error obteniendo tasa de cambio:', err.message)
    return null
  }
}