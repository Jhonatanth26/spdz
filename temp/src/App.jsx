import React, { useState, useMemo, useEffect, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { useAuth } from "./hooks/useAuth";
import { useSupabaseTable } from "./hooks/useSupabaseTable";
import { useSolicitudes } from "./hooks/useSolicitudes";
import { useNotificaciones } from "./hooks/useNotificaciones";
import { subirArchivo, obtenerUrlFirmada, archivoDentroDelLimite, TAMANO_MAXIMO_MB, subirArchivoPublico, subirBytes } from "./lib/storage";
import { obtenerTasaCambioCOP } from "./lib/tasaCambio";
import { firmarPDF } from "./lib/firmarPdf";
import { generarOrdenServicioPDF } from "./lib/generarOrdenServicio";
import { enviarCorreo } from "./lib/correo";
import LoginReal from "./LoginReal";
import {
  ShoppingCart, Wrench, Building2, CheckCircle2, XCircle, Clock,
  FileText, TrendingUp, ChevronRight, Plus, Trash2, Pencil,
  Calendar, Award, ArrowLeft, LayoutDashboard, ListChecks, BarChart3,
  DollarSign, PackageCheck, CalendarClock, Boxes, Users, Truck,
  Settings, Target, ClipboardList, Lock, LogOut, History, PenTool, ShieldCheck,
  Paperclip, Mail, Camera, Timer, Layers, MessageSquare, UserCircle, Send, CheckSquare, PanelLeftClose, PanelLeftOpen, Upload as UploadIcon, Bell,
} from "lucide-react";
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, Tooltip,
  ResponsiveContainer, CartesianGrid, Legend,
} from "recharts";

/* ---------------------------------------------------------
   CONFIG / DATOS BASE
--------------------------------------------------------- */
const EMPRESAS_INIT = [
  { id: "emp1", nombre: "Empresa 1", nit: "900111222-1", logoUrl: null },
  { id: "emp2", nombre: "Empresa 2", nit: "900333444-2", logoUrl: null },
];

const AREAS_INIT = [
  { id: "compras", nombre: "Compras", presupuesto: 40000000 },
  { id: "produccion", nombre: "Producción", presupuesto: 60000000 },
  { id: "logistica", nombre: "Logística", presupuesto: 25000000 },
  { id: "mercadeo", nombre: "Mercadeo", presupuesto: 15000000 },
  { id: "sistemas", nombre: "Sistemas", presupuesto: 20000000 },
];

const CENTROS_COSTO_INIT = [
  { id: "cc1", nombre: "CC-100 Administración" },
  { id: "cc2", nombre: "CC-200 Planta de producción" },
  { id: "cc3", nombre: "CC-300 Logística y distribución" },
  { id: "cc4", nombre: "CC-400 Tecnología" },
];

const CONCEPTOS_GASTO_INIT = [
  { id: "cg1", nombre: "Materia prima" },
  { id: "cg2", nombre: "Mantenimiento" },
  { id: "cg3", nombre: "Servicios generales" },
  { id: "cg4", nombre: "Tecnología / software" },
  { id: "cg5", nombre: "Papelería y oficina" },
];

// clave de demo para todos: "1234"
const USUARIOS_INIT = [
  { id: "u1", nombre: "Jhonatan Thomas", usuario: "jthomas", clave: "1234", email: "jhonatan.thomas@modaoxford.com", cargo: "Coordinador de Procesos y Planeación", areaId: "produccion", rol: "Solicitante", firmaFotoUrl: null },
  { id: "u2", nombre: "Laura Restrepo", usuario: "lrestrepo", clave: "1234", email: "laura.restrepo@modaoxford.com", cargo: "Jefe de Producción", areaId: "produccion", rol: "Jefe de Área", firmaFotoUrl: null },
  { id: "u3", nombre: "Carlos Vélez", usuario: "cvelez", clave: "1234", email: "carlos.velez@modaoxford.com", cargo: "Jefe de Sistemas", areaId: "sistemas", rol: "Jefe de Área", firmaFotoUrl: null },
  { id: "u4", nombre: "María Fernanda Ríos", usuario: "mrios", clave: "1234", email: "maria.rios@modaoxford.com", cargo: "Directora Financiera", areaId: "compras", rol: "Dirección Financiera", firmaFotoUrl: null },
  { id: "u5", nombre: "Andrés Gómez", usuario: "agomez", clave: "1234", email: "andres.gomez@modaoxford.com", cargo: "Gerente General", areaId: "compras", rol: "Gerencia", firmaFotoUrl: null },
  { id: "u6", nombre: "Paula Zapata", usuario: "pzapata", clave: "1234", email: "paula.zapata@modaoxford.com", cargo: "Analista de Compras", areaId: "compras", rol: "Compras", firmaFotoUrl: null },
];

const PROVEEDORES_INIT = [
  { id: "p1", nombre: "Distribuidora del Norte", nit: "800123456-1", contacto: "3001234567", email: "ventas@distribuidoranorte.com" },
  { id: "p2", nombre: "Suministros Andinos", nit: "800654321-2", contacto: "3019876543", email: "contacto@suministrosandinos.com" },
  { id: "p3", nombre: "InsuQuímicos SAS", nit: "800789456-3", contacto: "3025551234", email: "pedidos@insuquimicos.com" },
];

const ITEMS_CATALOGO_INIT = [
  { id: "i1", nombre: "Sal industrial", unidadDefault: "libra", categoria: "Insumos" },
  { id: "i2", nombre: "Mantenimiento anual servidores", unidadDefault: "servicio", categoria: "Tecnología" },
  { id: "i3", nombre: "Resma de papel carta", unidadDefault: "paquete", categoria: "Papelería" },
];

// histórico de compras ya existentes en el sistema (solo visible para Compras)
const HISTORICO_INIT = [
  { id: "h1", itemNombre: "Sal industrial", fecha: "2026-04-12", proveedor: "Suministros Andinos", precioUnitario: 1850, cantidad: 40, unidad: "libra" },
  { id: "h2", itemNombre: "Sal industrial", fecha: "2026-01-08", proveedor: "Distribuidora del Norte", precioUnitario: 4100, cantidad: 20, unidad: "kilo" },
];

const UMBRAL_DIRECCION = 500000;
const UMBRAL_GERENCIA = 100000000;
const UNIDADES = ["unidad", "libra", "kilo", "gramo", "litro", "mililitro", "metro", "caja", "paquete", "hora", "servicio"];
const IVA_OPCIONES = [0, 5, 19];
const MONEDAS = ["COP", "USD", "EUR", "MXN"];
const COLORS = ["#4f46e5", "#f59e0b", "#10b981", "#ef4444", "#0ea5e9", "#a855f7"];
const ROLES = ["Solicitante", "Jefe de Área", "Director de Área", "Jefe de Área y Director", "Dirección Financiera", "Gerencia", "Compras", "Administrador"];

const PASOS = [
  { key: "solicitud", label: "Solicitud creada" },
  { key: "aprobacion_jefe", label: "Aprobación jefe de área" },
  { key: "aprobacion_director", label: "Aprobación director de área" },
  { key: "cotizando", label: "Revisión y cotizaciones (compras)" },
  { key: "comparativo", label: "Cuadro comparativo" },
  { key: "aprobacion_financiera", label: "Dirección financiera" },
  { key: "aprobacion_gerencia", label: "Gerencia" },
  { key: "orden", label: "Orden generada" },
  { key: "oc_enviada", label: "Enviada al proveedor" },
  { key: "recepcion", label: "Recepción / Ejecución" },
  { key: "completada", label: "Completada" },
];

const fmt = (n) => (n || 0).toLocaleString("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });
// etiqueta de un concepto de gasto: Grupo – Código – Cuenta – Centro de costo (el centro de costo va anexado al concepto)
// colores para distinguir cada ítem (mismos en la tabla de ítems, cotizaciones, comparativo, AIU y plan de pagos)
const TINTES_ITEM = [
  { fondo: "bg-slate-50", borde: "border-slate-300", punto: "bg-slate-500" },
  { fondo: "bg-indigo-50/60", borde: "border-indigo-200", punto: "bg-indigo-500" },
  { fondo: "bg-amber-50/60", borde: "border-amber-200", punto: "bg-amber-500" },
  { fondo: "bg-emerald-50/60", borde: "border-emerald-200", punto: "bg-emerald-500" },
  { fondo: "bg-rose-50/50", borde: "border-rose-200", punto: "bg-rose-500" },
  { fondo: "bg-sky-50/60", borde: "border-sky-200", punto: "bg-sky-500" },
];
const tinteItem = (idx) => TINTES_ITEM[(idx || 0) % TINTES_ITEM.length];
// El AIU (costos indirectos) lo define Compras al cotizar. No se pide ni se muestra en lo que llena el solicitante.
// En el detalle de la solicitud se ve (en solo lectura para quien no es Compras) cuando ya está definido: desde la
// etapa de cotización, o antes si Compras ya lo cargó (ej. una solicitud devuelta que vuelve al jefe o al director).
const ETAPAS_SIN_AIU = ["solicitud", "aprobacion_jefe", "aprobacion_director"];
const tieneAiuDefinido = (s) => tieneAiuValores(s.aiu) || (s.items || []).some((it) => tieneAiuValores(it.aiu) || (it.cotizaciones || []).some((c) => tieneAiuValores(c.aiu)));
const aiuVisible = (s) => s.tipo === "servicio" && (s.items || []).some((it) => (it.cotizaciones || []).length > 0) && (!ETAPAS_SIN_AIU.includes(s.status) || tieneAiuDefinido(s));
const labelConcepto = (c) => [c.grupo, c.codigo, c.nombre, c.centroCosto].filter(Boolean).join(" – ");

// convierte un número a su forma escrita en español, para el "Son: ..." de las órdenes (ej. 7591300 -> "SIETE MILLONES QUINIENTOS NOVENTA Y UN MIL TRESCIENTOS")
function numeroALetras(n) {
  const UNIDADES = ["", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
  const DIEZ_A_DIECINUEVE = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
  const DECENAS = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
  const CENTENAS = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];

  function trescientos(num) {
    if (num === 0) return "";
    if (num === 100) return "CIEN";
    let s = "";
    const c = Math.floor(num / 100), resto = num % 100;
    if (c > 0) s += CENTENAS[c] + " ";
    if (resto >= 10 && resto <= 19) { s += DIEZ_A_DIECINUEVE[resto - 10]; }
    else {
      const d = Math.floor(resto / 10), u = resto % 10;
      if (d === 2 && u > 0) s += "VEINTI" + UNIDADES[u].toLowerCase().charAt(0).toUpperCase() + UNIDADES[u].toLowerCase().slice(1);
      else { if (d > 0) s += DECENAS[d]; if (d > 0 && u > 0) s += " Y "; if (u > 0) s += UNIDADES[u]; }
    }
    return s.trim();
  }

  let entero = Math.round(Math.abs(n || 0));
  if (entero === 0) return "CERO PESOS M/CTE";
  const millones = Math.floor(entero / 1000000);
  const miles = Math.floor((entero % 1000000) / 1000);
  const resto = entero % 1000;
  let partes = [];
  if (millones > 0) partes.push((millones === 1 ? "UN MILLÓN" : trescientos(millones) + " MILLONES"));
  if (miles > 0) partes.push((miles === 1 ? "MIL" : trescientos(miles) + " MIL"));
  if (resto > 0) partes.push(trescientos(resto));
  return (partes.join(" ").trim() || "CERO") + " PESOS M/CTE";
}
// hace crecer un <textarea> automáticamente según el contenido que se escribe
const autoResize = (e) => { e.target.style.height = "auto"; e.target.style.height = e.target.scrollHeight + "px"; };
const hoy = () => new Date().toISOString().slice(0, 10);
const ahoraISO = () => new Date().toISOString();
let idCounter = 4000;
const nextId = () => (idCounter++).toString();

/* ---------------------------------------------------------
   LÓGICA DE NEGOCIO — dinero
--------------------------------------------------------- */
// precio final efectivo de la cotización (en su moneda original), aplicando el descuento si existe;
// si no hay descuento, usa el precio final negociado manualmente, o el precio inicial si no hay ninguno
function precioFinalEfectivo(cot) {
  const inicial = parseFloat(cot.precioUnitario) || 0;
  const base = parseFloat(cot.precioFinal) || inicial; // parte del precio final negociado si existe, si no del inicial
  const descuento = parseFloat(cot.descuentoValor);
  if (descuento > 0) {
    const descontado = cot.descuentoTipo === "porcentaje" ? base * (1 - descuento / 100) : base - descuento;
    return Math.max(0, descontado);
  }
  return base;
}
// convierte el precio final efectivo a COP, según la moneda y tasa de cambio registradas
function precioEnCOP(cot) {
  const efectivo = precioFinalEfectivo(cot);
  const tasa = cot.moneda && cot.moneda !== "COP" ? (parseFloat(cot.tasaCambio) || 1) : 1;
  return efectivo * tasa;
}
// precio equivalente por unidad solicitada, ya en COP
function precioEquivalente(cot) {
  const factor = parseFloat(cot.factorConversion) || 1;
  return precioEnCOP(cot) / factor;
}
function desgloseCotizacion(cot, cantidadSolicitada) {
  const subtotal = precioEquivalente(cot) * (parseFloat(cantidadSolicitada) || 0);
  const ivaPct = parseFloat(cot.ivaPct ?? 19) || 0;
  const iva = subtotal * (ivaPct / 100);
  return { subtotal, iva, total: subtotal + iva };
}
// un ítem con UNA sola cotización necesita el comentario de por qué solo se cotizó con un proveedor
function cotizacionUnicaSinJustificar(item) {
  return (item.cotizaciones || []).length === 1 && !((item.cotizaciones[0].justificacionUnico || "").trim());
}
// Reglas de precio de una cotización (dos topes):
//  1) el precio inicial cotizado no puede superar el precio estimado que puso el solicitante;
//  2) el precio final negociado no puede superar el precio inicial cotizado por el proveedor.
// (1) se compara en la misma base: COP por unidad del ítem; (2) va en la misma moneda y unidad de la cotización.
function erroresPrecioCotizacion(item, cot) {
  const errores = [];
  const inicial = parseFloat(cot.precioUnitario) || 0;
  const finalNeg = parseFloat(cot.precioFinal) || 0;
  const factor = parseFloat(cot.factorConversion) || 1;
  const tasaCot = cot.moneda && cot.moneda !== "COP" ? (parseFloat(cot.tasaCambio) || 1) : 1;
  const tasaEst = item.moneda && item.moneda !== "COP" ? (parseFloat(item.tasaCambio) || 1) : 1;
  const estimadoCOP = (parseFloat(item.precioEstimado) || 0) * tasaEst;
  const inicialCOP = (inicial * tasaCot) / factor;
  if (estimadoCOP > 0 && inicial > 0 && inicialCOP - estimadoCOP > 0.5) errores.push({ tipo: "estimado", texto: `El precio inicial (${fmt(inicialCOP)} por ${item.unidad}) supera el precio estimado por el solicitante (${fmt(estimadoCOP)}).` });
  if (finalNeg > 0 && inicial > 0 && finalNeg - inicial > 0.0001) errores.push({ tipo: "final", texto: "El precio final negociado no puede ser mayor que el precio inicial cotizado por el proveedor." });
  return errores;
}
function itemConPrecioInvalido(item) { return (item.cotizaciones || []).some((c) => erroresPrecioCotizacion(item, c).length > 0); }
function primerErrorPrecio(item) {
  for (const c of item.cotizaciones || []) { const e = erroresPrecioCotizacion(item, c); if (e.length) return e[0].texto; }
  return null;
}
function tieneAiuValores(aiu) { return !!aiu && (parseFloat(aiu.administracionPct) || parseFloat(aiu.utilidadPct) || parseFloat(aiu.imprevistosPct)); }
// total real de una cotización de servicio: usa el AIU propio de esa cotización si lo tiene, si no
// el del ítem como respaldo — es el mismo criterio que usa el cálculo real de la solicitud
function totalConAiuCotizacion(desglose, cotizacion, itemAiu) {
  const aiu = tieneAiuValores(cotizacion?.aiu) ? cotizacion.aiu : (itemAiu || {});
  const adm = desglose.subtotal * (parseFloat(aiu.administracionPct) || 0) / 100;
  const util = desglose.subtotal * (parseFloat(aiu.utilidadPct) || 0) / 100;
  const imprev = desglose.subtotal * (parseFloat(aiu.imprevistosPct) || 0) / 100;
  return desglose.subtotal + adm + util + imprev + util * 0.19;
}
// sinIva/itemAiu: cuando se pasan (órdenes de servicio), la calificación por precio usa el total
// CON AIU de cada proveedor, no solo el Costo Directo — si no, se recomendaría mal cuando dos
// proveedores tienen Costo Directo parecido pero AIU muy distinto
function calcularScores(cotizaciones, cantidadSolicitada, sinIva, itemAiu) {
  if (!cotizaciones.length) return [];
  const desgloses = cotizaciones.map((c) => desgloseCotizacion(c, cantidadSolicitada));
  const totales = desgloses.map((d, i) => (sinIva ? totalConAiuCotizacion(d, cotizaciones[i], itemAiu) : d.total));
  const entregas = cotizaciones.map((c) => parseFloat(c.diasEntrega) || 0);
  const condiciones = cotizaciones.map((c) => parseFloat(c.condicionesScore) || 0);
  const minTotal = Math.min(...totales), maxTotal = Math.max(...totales);
  const minEnt = Math.min(...entregas), maxEnt = Math.max(...entregas);
  const minCond = Math.min(...condiciones), maxCond = Math.max(...condiciones);
  return cotizaciones.map((c, i) => {
    const precioScore = maxTotal === minTotal ? 1 : (maxTotal - totales[i]) / (maxTotal - minTotal);
    const entregaScore = maxEnt === minEnt ? 1 : (maxEnt - entregas[i]) / (maxEnt - minEnt);
    const condScore = maxCond === minCond ? 1 : (condiciones[i] - minCond) / (maxCond - minCond);
    return { ...c, ...desgloses[i], totalConAiu: sinIva ? totales[i] : undefined, score: precioScore * 0.6 + entregaScore * 0.25 + condScore * 0.15 };
  });
}
function mejorCotizacionIdx(cotizaciones, cantidadSolicitada, sinIva, itemAiu) {
  const scored = calcularScores(cotizaciones, cantidadSolicitada, sinIva, itemAiu);
  if (!scored.length) return -1;
  let best = 0;
  scored.forEach((s, i) => { if (s.score > scored[best].score) best = i; });
  return best;
}
function desgloseItem(item) {
  if (!item.cotizaciones.length) {
    const inicial = parseFloat(item.precioEstimado) || 0;
    const descuento = parseFloat(item.descuentoValor);
    const precioConDescuento = descuento > 0
      ? Math.max(0, item.descuentoTipo === "porcentaje" ? inicial * (1 - descuento / 100) : inicial - descuento)
      : inicial;
    const tasa = item.moneda && item.moneda !== "COP" ? (parseFloat(item.tasaCambio) || 1) : 1;
    const precioEnCop = precioConDescuento * tasa;
    const subtotal = precioEnCop * (parseFloat(item.cantidad) || 0);
    const ivaPct = parseFloat(item.ivaEstimado ?? 19) || 0;
    const iva = subtotal * (ivaPct / 100);
    return { subtotal, iva, total: subtotal + iva };
  }
  const idx = item.cotizacionSeleccionada ?? mejorCotizacionIdx(item.cotizaciones, item.cantidad);
  const cot = item.cotizaciones[idx];
  return cot ? desgloseCotizacion(cot, item.cantidad) : { subtotal: 0, iva: 0, total: 0 };
}
function desgloseSolicitud(solicitud) {
  if (solicitud.tipo === "servicio") return desgloseSolicitudServicio(solicitud);
  return solicitud.items.reduce((acc, item) => {
    const d = desgloseItem(item);
    return { subtotal: acc.subtotal + d.subtotal, iva: acc.iva + d.iva, total: acc.total + d.total };
  }, { subtotal: 0, iva: 0, total: 0 });
}
// órdenes de servicio/trabajo: Costo Directo + AIU (Administración, Imprevistos, Utilidad) + IVA solo sobre la Utilidad —
// no llevan IVA por ítem, ese campo queda deshabilitado para este tipo de solicitud.
// El AIU es por ítem (independiente del proveedor que finalmente se adjudique) — el solicitante lo estima
// al crear la solicitud; una vez hay cotización seleccionada, el AIU propio de esa cotización lo reemplaza.
function desgloseSolicitudServicio(solicitud) {
  let costoDirecto = 0, administracion = 0, utilidad = 0, imprevistos = 0;
  solicitud.items.forEach((item) => {
    const dItem = desgloseItem(item);
    costoDirecto += dItem.subtotal;
    let aiu = item.aiu || solicitud.aiu || {};
    if (item.cotizaciones?.length) {
      const sel = item.cotizacionSeleccionada ?? mejorCotizacionIdx(item.cotizaciones, item.cantidad);
      const cot = item.cotizaciones[sel];
      if (cot?.aiu && (parseFloat(cot.aiu.administracionPct) || parseFloat(cot.aiu.utilidadPct) || parseFloat(cot.aiu.imprevistosPct))) aiu = cot.aiu;
    }
    administracion += dItem.subtotal * (parseFloat(aiu.administracionPct) || 0) / 100;
    utilidad += dItem.subtotal * (parseFloat(aiu.utilidadPct) || 0) / 100;
    imprevistos += dItem.subtotal * (parseFloat(aiu.imprevistosPct) || 0) / 100;
  });
  const ivaUtilidad = utilidad * 0.19;
  const total = costoDirecto + administracion + utilidad + imprevistos + ivaUtilidad;
  return { costoDirecto, administracion, utilidad, imprevistos, ivaUtilidad, total, subtotal: costoDirecto, iva: ivaUtilidad };
}
function totalSolicitud(s) { return desgloseSolicitud(s).total; }
function requiereDireccion(m) { return m >= UMBRAL_DIRECCION; }
function requiereGerencia(m) { return m >= UMBRAL_GERENCIA; }
function totalPagado(pagos) {
  if (pagos?.tipoPago === "contado") return parseFloat(pagos?.pagoUnico?.valor) || 0;
  return (parseFloat(pagos?.anticipo?.valor) || 0) + (pagos?.intermedio?.activo ? (parseFloat(pagos.intermedio.valor) || 0) : 0) + (parseFloat(pagos?.final?.valor) || 0);
}
// arma la lista de tramos de pago (uno solo si es "de contado", o hasta 3 si es plan por etapas)
function tramosDePago(pagos) {
  if (pagos?.tipoPago === "contado") return [{ tipo: "Pago único", ...pagos.pagoUnico }];
  return [
    { tipo: "Anticipo", ...pagos.anticipo },
    ...(pagos.intermedio?.activo ? [{ tipo: "Intermedio", ...pagos.intermedio }] : []),
    { tipo: "Final", ...pagos.final },
  ];
}
// ---- Plan de pagos POR ÍTEM (compras y servicios) ----
// Cada ítem tiene su plan sugerido por el solicitante (item.pagosSugeridos), su plan oficial (item.pagos)
// y su confirmación (item.pagosConfirmados). Las solicitudes antiguas guardaban un solo plan general en la
// solicitud (s.pagos): mientras tengan un único ítem se sigue leyendo de ahí, para no perder nada.
function planTieneValores(p) {
  return !!p && (parseFloat(p.pagoUnico?.valor) > 0 || parseFloat(p.anticipo?.valor) > 0 || parseFloat(p.intermedio?.valor) > 0 || parseFloat(p.final?.valor) > 0);
}
function planOficialItem(s, it) {
  if (planTieneValores(it.pagos)) return { ...planPagosVacio(), ...it.pagos };
  if (s.items.length === 1 && planTieneValores(s.pagos)) return { ...planPagosVacio(), ...s.pagos };
  return planPagosVacio();
}
function planSugeridoItem(s, it) {
  if (planTieneValores(it.pagosSugeridos)) return { ...planPagosVacio(), ...it.pagosSugeridos };
  if (s.items.length === 1 && planTieneValores(s.pagosSugeridos)) return { ...planPagosVacio(), ...s.pagosSugeridos };
  return planPagosVacio();
}
function planConfirmadoItem(s, it) { return it.pagosConfirmados ?? (s.items.length === 1 ? !!s.pagosConfirmados : false); }
function planesTodosConfirmados(s) { return s.items.every((it) => planConfirmadoItem(s, it)); }
// tramos de pago de los ítems con plan confirmado (marcando de qué ítem viene cada uno cuando hay varios)
function tramosDePagoSolicitud(s) {
  return s.items.flatMap((it) => (planConfirmadoItem(s, it) ? tramosDePago(planOficialItem(s, it)).map((t) => ({ ...t, itemNombre: s.items.length > 1 ? it.nombre : undefined })) : []));
}
// filas para el calendario y el dashboard: un pago por cada tramo de los planes confirmados; lo que no tiene
// plan confirmado usa la fecha estimada de entrega con su valor (un solo renglón para esos ítems)
function filasPagosSolicitud(s, prov) {
  const filas = [];
  const dias = (fecha) => Math.round((new Date(fecha + "T00:00:00") - new Date(hoy() + "T00:00:00")) / 86400000);
  const sinIva = s.tipo === "servicio";
  tramosDePagoSolicitud(s).forEach((t, i) => {
    if (!(parseFloat(t.valor) > 0) || !t.fecha) return;
    filas.push({ id: `${s.id}-${t.itemNombre || ""}-${t.tipo}-${i}`, solicitudId: s.id, folio: s.folio, proveedor: prov, tipo: t.itemNombre ? `${t.tipo} (${t.itemNombre})` : t.tipo, valor: parseFloat(t.valor), fecha: t.fecha, dias: dias(t.fecha), pagado: !!t.pagado });
  });
  const sinPlan = s.items.filter((it) => !planConfirmadoItem(s, it));
  const totalSinPlan = sinPlan.reduce((acc, it) => acc + totalItemConAiu(it, sinIva), 0);
  if (sinPlan.length && totalSinPlan > 0 && s.fechaEstimada) {
    filas.push({ id: `${s.id}-estimado`, solicitudId: s.id, folio: s.folio, proveedor: prov, tipo: "Sin plan de pagos (fecha de entrega est.)", valor: totalSinPlan, fecha: s.fechaEstimada, dias: dias(s.fechaEstimada), pagado: false });
  }
  return filas;
}
// cuando cambia el total (ej. se ajustó el AIU), reescala proporcionalmente los valores ya puestos
// en el plan para que la suma siga cuadrando exacto con el nuevo total, sin perder las proporciones
function reescalarPlanPago(pagos, totalNuevo) {
  if (pagos.tipoPago === "contado") return { ...pagos, pagoUnico: { ...pagos.pagoUnico, valor: totalNuevo } };
  const programado = totalPagado(pagos);
  if (!(programado > 0)) return pagos;
  const factor = totalNuevo / programado;
  if (!isFinite(factor) || factor <= 0) return pagos;
  return {
    ...pagos,
    anticipo: { ...pagos.anticipo, valor: pagos.anticipo.valor ? Math.round(parseFloat(pagos.anticipo.valor) * factor) : pagos.anticipo.valor },
    intermedio: { ...pagos.intermedio, valor: pagos.intermedio.valor ? Math.round(parseFloat(pagos.intermedio.valor) * factor) : pagos.intermedio.valor },
    final: { ...pagos.final, valor: pagos.final.valor ? Math.round(parseFloat(pagos.final.valor) * factor) : pagos.final.valor },
  };
}
// total real de UN ítem (Costo Directo + su propio AIU, con el mismo respaldo cotización→ítem que
// usa el cálculo general) — es la base contra la que debe cuadrar el plan de pagos de ese ítem
function totalItemConAiu(item, sinIva) {
  const d = desgloseItem(item);
  if (!sinIva) return d.total;
  let aiu = item.aiu || {};
  if (item.cotizaciones?.length) {
    const sel = item.cotizacionSeleccionada ?? mejorCotizacionIdx(item.cotizaciones, item.cantidad, true, item.aiu);
    const cot = item.cotizaciones[sel];
    if (tieneAiuValores(cot?.aiu)) aiu = cot.aiu;
  }
  const adm = d.subtotal * (parseFloat(aiu.administracionPct) || 0) / 100;
  const util = d.subtotal * (parseFloat(aiu.utilidadPct) || 0) / 100;
  const imprev = d.subtotal * (parseFloat(aiu.imprevistosPct) || 0) / 100;
  return d.subtotal + adm + util + imprev + util * 0.19;
}
// valida que las fechas del plan de pagos queden en orden creciente: anticipo ≤ intermedio (si aplica) ≤ final
// devuelve un mensaje de error, o null si está bien
function validarOrdenFechas(pagos, campo, nuevaFecha) {
  const anticipo = campo === "anticipo" ? nuevaFecha : pagos.anticipo.fecha;
  const intermedio = campo === "intermedio" ? nuevaFecha : pagos.intermedio.fecha;
  const final = campo === "final" ? nuevaFecha : pagos.final.fecha;
  const hayIntermedio = pagos.intermedio.activo || campo === "intermedio";
  if (anticipo && final && anticipo > final) return "La fecha del anticipo no puede ser posterior a la del pago final.";
  if (hayIntermedio && anticipo && intermedio && anticipo > intermedio) return "La fecha del anticipo no puede ser posterior a la del pago intermedio.";
  if (hayIntermedio && intermedio && final && intermedio > final) return "La fecha del pago intermedio no puede ser posterior a la del pago final.";
  return null;
}
// evita porcentajes negativos o absurdamente altos (ej. AIU) — deja vacío o cualquier número entre 0 y 100
function pctValido(val) {
  if (val === "") return "";
  const n = parseFloat(val);
  if (isNaN(n)) return "";
  if (n < 0) return 0;
  if (n > 100) return 100;
  return val;
}
// duración legible entre dos timestamps ISO
function duracion(iniISO, finISO) {
  if (!iniISO || !finISO) return "—";
  const ms = new Date(finISO) - new Date(iniISO);
  if (ms < 0) return "—";
  const mins = Math.floor(ms / 60000);
  if (mins < 60) return `${mins} min`;
  const horas = Math.floor(mins / 60);
  if (horas < 24) return `${horas} h ${mins % 60} min`;
  const dias = Math.floor(horas / 24);
  return `${dias} d ${horas % 24} h`;
}

/* ---------------------------------------------------------
   PERMISOS — configurables desde la app (Catálogo → Permisos),
   ya no están fijos en el código. Administrador siempre tiene
   todo, sin excepción, para que nunca se pueda quedar sin acceso.
--------------------------------------------------------- */
const PERMISOS_DISPONIBLES = [
  { key: "aprobar_jefe", label: "Aprobar como jefe de área" },
  { key: "aprobar_director", label: "Aprobar como director de área" },
  { key: "aprobar_financiera", label: "Aprobar como Dirección Financiera" },
  { key: "aprobar_gerencia", label: "Aprobar como Gerencia" },
  { key: "gestionar_cotizaciones", label: "Gestionar cotizaciones y cuadro comparativo" },
  { key: "editar_pagos", label: "Editar y confirmar el plan de pagos (por ítem, en compras y servicios)" },
  { key: "ver_catalogos", label: "Ver el Catálogo" },
  { key: "ver_todas_solicitudes", label: "Ver todas las solicitudes (no solo las propias)" },
  { key: "reabrir_solicitudes", label: "Reabrir solicitudes rechazadas" },
  { key: "ver_historico", label: "Ver histórico de compras" },
  { key: "ver_mis_pendientes", label: "Ver la pantalla \"Mis pendientes\"" },
  { key: "ver_calendario_pagos", label: "Ver el Calendario de pagos" },
  { key: "ver_ordenes_enviadas", label: "Ver el reporte de Órdenes enviadas a proveedores" },
  { key: "ver_plan_inversion", label: "Ver el Plan de inversión" },
  { key: "ver_evaluaciones_proveedores", label: "Ver el reporte de Evaluación de proveedores" },
];

// mapa en memoria { [rol]: { [permiso]: true } } — se sincroniza cada vez que se
// cargan/actualizan los permisos desde Supabase (ver App()). Así todas las
// funciones de abajo consultan el permiso sin que el resto del código cambie.
let __permisosPorRol = {};
function construirMapaPermisos(lista) {
  const mapa = {};
  (lista || []).forEach((p) => {
    if (!mapa[p.rol]) mapa[p.rol] = {};
    mapa[p.rol][p.permiso] = !!p.activo;
  });
  return mapa;
}
function tienePermiso(rol, permiso) {
  if (rol === "Administrador") return true;
  return !!__permisosPorRol[rol]?.[permiso];
}

const tieneAreaACargo = (u, areaId) => u.areaId === areaId || (u.areasAdicionales || []).includes(areaId);
const puedeAprobarJefe = (u, s) => u.rol === "Administrador" || (tienePermiso(u.rol, "aprobar_jefe") && tieneAreaACargo(u, s.areaId));
const puedeAprobarDirector = (u, s) => u.rol === "Administrador" || (tienePermiso(u.rol, "aprobar_director") && tieneAreaACargo(u, s.areaId));
const puedeGestionarCotizaciones = (u) => tienePermiso(u.rol, "gestionar_cotizaciones");
const puedeAprobarFinanciera = (u) => tienePermiso(u.rol, "aprobar_financiera");
const puedeAprobarGerencia = (u) => tienePermiso(u.rol, "aprobar_gerencia");
const puedeReabrir = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "reabrir_solicitudes");
// determina en qué paso quedó marcada como rechazada, para poder reabrirla justo ahí
function pasoDelRechazo(solicitud) {
  if (solicitud.firmas?.gerencia?.aprobado === false) return { status: "aprobacion_gerencia", campo: "gerencia" };
  if (solicitud.firmas?.financiera?.aprobado === false) return { status: "aprobacion_financiera", campo: "financiera" };
  if (solicitud.firmas?.director?.aprobado === false) return { status: "aprobacion_director", campo: "director" };
  if (solicitud.firmas?.jefe?.aprobado === false) return { status: "aprobacion_jefe", campo: "jefe" };
  if (solicitud.revisionCompras?.estado === "rechazada") return { status: "cotizando", campo: null, revision: true };
  return { status: "aprobacion_jefe", campo: null };
}
const puedeVerCatalogos = (u) => tienePermiso(u.rol, "ver_catalogos");
const puedeVerTodasSolicitudes = (u) => tienePermiso(u.rol, "ver_todas_solicitudes");
const puedeEditarPagos = (u) => tienePermiso(u.rol, "editar_pagos");
const puedeVerMisPendientes = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_mis_pendientes");
const puedeVerCalendarioPagos = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_calendario_pagos");
const puedeVerOrdenesEnviadas = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_ordenes_enviadas");
const puedeVerPlanInversion = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_plan_inversion");
const puedeVerEvaluaciones = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_evaluaciones_proveedores");

/* ---------------------------------------------------------
   EVALUACIÓN DE PROVEEDORES — formato oficial (Registro Selección y Evaluación de Proveedores)
--------------------------------------------------------- */
const CRITERIOS_EVALUACION = [
  { key: "estandaresCalidad", componente: "Calidad", subcomponente: "Estándares", texto: "Cumplimiento con estándares de calidad establecidos por la empresa (certificados de conformidad)", peso: 0.06 },
  { key: "condicionesTecnicas", componente: "Calidad", subcomponente: "Estándares", texto: "Cumplimiento con las condiciones técnicas requeridas", peso: 0.06 },
  { key: "atencion", componente: "Calidad", subcomponente: "Servicio al cliente", texto: "Atención", peso: 0.06 },
  { key: "tiempoEntrega", componente: "Calidad", subcomponente: "Servicio al cliente", texto: "Tiempo de entrega (una vez recibida la solicitud, entrega rápidamente el producto o servicio)", peso: 0.08 },
  { key: "servicioPostventa", componente: "Calidad", subcomponente: "Servicio al cliente", texto: "Servicio postventa", peso: 0.06 },
  { key: "stockDisponible", componente: "Calidad", subcomponente: "Servicio al cliente", texto: "Mantiene producto en stock o disponible para el servicio", peso: 0.04 },
  { key: "sgc", componente: "Calidad", subcomponente: "Servicio al cliente", texto: "Posee sistema de Gestión de la Calidad (certificado, en proceso o no tiene)", peso: 0.04 },
  { key: "atencionQuejas", componente: "Calidad", subcomponente: "Servicio al cliente", texto: "Atiende oportunamente las quejas y solicitudes", peso: 0.08 },
  { key: "sgSst", componente: "HSE", subcomponente: "Seguridad y salud en el trabajo", texto: "Cuenta con Sistema de Gestión en Seguridad y Salud en el Trabajo (SG-SST) implementado y funcionando", peso: 0.10 },
  { key: "envioSgSst", componente: "HSE", subcomponente: "Seguridad y salud en el trabajo", texto: "Envía oportunamente los requerimientos de SG-SST (procedimientos, certificados, fichas MSDS, etc.)", peso: 0.09 },
  { key: "politicasHseq", componente: "HSE", subcomponente: "Seguridad y salud en el trabajo", texto: "Cumple oportunamente las políticas de HSEQ (utilización EPP, inducción, procedimientos)", peso: 0.07 },
  { key: "licenciaAmbiental", componente: "HSE", subcomponente: "Gestión ambiental", texto: "Dispone de una Licencia Ambiental (si aplica)", peso: 0.09 },
  { key: "programaPostconsumo", componente: "HSE", subcomponente: "Gestión ambiental", texto: "Cuenta y brinda un programa de postconsumo (tóners, cartuchos, baterías, aceite usado, llantas, pilas, residuos electrónicos, residuos de iluminación)", peso: 0.08 },
  { key: "programaResiduos", componente: "HSE", subcomponente: "Gestión ambiental", texto: "Dispone de un programa de disposición de residuos (si aplica)", peso: 0.09 },
];
const DOCUMENTOS_EVALUACION = [
  { key: "rut", label: "RUT de la empresa" },
  { key: "camaraComercio", label: "Certificado de Cámara de Comercio" },
  { key: "cedulaRL", label: "Fotocopia de la cédula del representante legal" },
  { key: "referenciasComerciales", label: "Referencias comerciales (2)" },
  { key: "certificadosHSEQ", label: "Certificados de calidad, seguridad y salud en el trabajo y medio ambiente" },
];

function evaluacionProveedorVacia() {
  return {
    proveedorId: null, proveedorNombre: "", tipoProveedor: "",
    fechaSeleccion: "", fechaEvaluacion: "",
    nit: "", cc: "", representanteLegal: "", telefono: "", fax: "", email: "",
    ciudad: "", direccion: "", serviciosPresta: "", descripcion: "", marca: "",
    documentos: {},
    criterios: {},
    observaciones: "",
    aprobadoPorCargo: "",
    firmaRealizada: { nombre: null, cargo: null, empresa: null, fecha: null, fotoUrl: null },
    completada: false, fechaCompletado: "",
  };
}

// puntaje 0-1 (suma de calificación/10 × peso de cada criterio con valor)
function puntajeEvaluacion(criterios) {
  let total = 0;
  CRITERIOS_EVALUACION.forEach((c) => { const v = parseFloat(criterios?.[c.key]); if (v > 0) total += (v / 10) * c.peso; });
  return total;
}
function clasificacionConfianza(pct) {
  if (pct >= 80) return { texto: "Confiable", tone: "green" };
  if (pct >= 51) return { texto: "Medio Confiable", tone: "amber" };
  return { texto: "Poco Confiable", tone: "red" };
}
// true si se calificaron los 14 criterios (obligatorio para poder completar la solicitud)
function evaluacionProveedorCompleta(ev) {
  if (!ev) return false;
  return CRITERIOS_EVALUACION.every((c) => parseFloat(ev.criterios?.[c.key]) > 0);
}

const puedeVerHistorico = (u) => tienePermiso(u.rol, "ver_historico");

/* ---------------------------------------------------------
   DATOS SEMILLA
--------------------------------------------------------- */
function planPagosVacio() { return { tipoPago: "plan", pagoUnico: { valor: "", fecha: "" }, anticipo: { valor: "", fecha: "" }, intermedio: { activo: false, valor: "", fecha: "" }, final: { valor: "", fecha: "" } }; }

function datosSemilla() {
  const s1 = {
    id: nextId(), folio: "SOL-1001", tipo: "compra", empresaId: "emp1", areaId: "produccion",
    solicitanteId: "u1", fechaCreacion: "2026-07-20", fechaEstimada: "2026-08-05",
    objetivo: "Garantizar el abastecimiento de sal industrial para el proceso de tinturado.",
    justificacion: "El inventario actual cubre solo 5 días de producción; se requiere reposición para no detener la línea.",
    conceptoGastoId: "cg1",
    status: "comparativo",
    revisionCompras: { estado: "aprobada", observacion: "Cantidades correctas.", usuario: "Paula Zapata", fecha: "2026-07-21" },
    items: [{
      id: nextId(), itemCatalogoId: "i1", nombre: "Sal industrial", cantidad: 50, unidad: "libra",
      precioEstimado: 2000, ivaEstimado: 19,
      cotizaciones: [
        { proveedorId: "p1", unidadCotizada: "kilo", factorConversion: 2.2, precioUnitario: 4200, precioFinal: 4100, diasEntrega: 3, condicionesScore: 7, ivaPct: 19, archivoNombre: "cot_norte.pdf" },
        { proveedorId: "p2", unidadCotizada: "libra", factorConversion: 1, precioUnitario: 1900, precioFinal: "", diasEntrega: 5, condicionesScore: 8, ivaPct: 19, archivoNombre: "cot_andinos.pdf" },
        { proveedorId: "p3", unidadCotizada: "kilo", factorConversion: 2.2, precioUnitario: 4400, precioFinal: "", diasEntrega: 2, condicionesScore: 6, ivaPct: 19, archivoNombre: "" },
      ],
      cotizacionSeleccionada: null, observacionSeleccion: "",
    }],
    firmas: {
      solicitante: { nombre: "Jhonatan Thomas", fecha: "2026-07-20", fotoUrl: null },
      jefe: { aprobado: true, nombre: "Laura Restrepo", fecha: "2026-07-21", observacion: "De acuerdo, es insumo crítico.", fotoUrl: null },
      financiera: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
      gerencia: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
    },
    pagosSugeridos: { anticipo: { valor: 40000, fecha: "2026-08-01" }, intermedio: { activo: false, valor: "", fecha: "" }, final: { valor: 60000, fecha: "2026-08-10" } },
    pagos: planPagosVacio(),
    pagosConfirmados: false,
    ocEnviada: { ordenesProveedor: [] },
    prioridad: null,
    evaluacionProveedor: evaluacionProveedorVacia(),
    recepcion: { archivos: [], comentario: "", recibidoSatisfaccion: false, usuario: "", fecha: "" },
    historialEstados: [
      { status: "solicitud", fecha: "2026-07-20T09:00:00.000Z" },
      { status: "aprobacion_jefe", fecha: "2026-07-20T09:05:00.000Z" },
      { status: "cotizando", fecha: "2026-07-21T14:00:00.000Z" },
      { status: "comparativo", fecha: ahoraISO() },
    ],
    notificaciones: [],
  };
  const s2 = {
    id: nextId(), folio: "SOL-1002", tipo: "servicio", empresaId: "emp2", areaId: "sistemas",
    solicitanteId: "u1", fechaCreacion: "2026-07-22", fechaEstimada: "2026-09-15",
    objetivo: "Mantener la disponibilidad y seguridad de la infraestructura de servidores.",
    justificacion: "El contrato de mantenimiento anterior venció; sin este servicio se pierde soporte y garantía del proveedor.",
    conceptoGastoId: "cg4",
    status: "aprobacion_jefe",
    revisionCompras: { estado: "no_aplica", observacion: "", usuario: "", fecha: "" },
    items: [{
      id: nextId(), itemCatalogoId: "i2", nombre: "Mantenimiento anual servidores", cantidad: 1, unidad: "servicio",
      precioEstimado: 9500000, ivaEstimado: 19, cotizaciones: [], cotizacionSeleccionada: null, observacionSeleccion: "",
    }],
    firmas: {
      solicitante: { nombre: "Jhonatan Thomas", fecha: "2026-07-22", fotoUrl: null },
      jefe: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
      financiera: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
      gerencia: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
    },
    pagosSugeridos: planPagosVacio(),
    pagos: planPagosVacio(),
    pagosConfirmados: false,
    ocEnviada: { ordenesProveedor: [] },
    prioridad: null,
    evaluacionProveedor: evaluacionProveedorVacia(),
    recepcion: { archivos: [], comentario: "", recibidoSatisfaccion: false, usuario: "", fecha: "" },
    historialEstados: [
      { status: "solicitud", fecha: "2026-07-22T10:00:00.000Z" },
      { status: "aprobacion_jefe", fecha: "2026-07-22T10:02:00.000Z" },
    ],
    notificaciones: [{ fecha: "2026-07-22T10:02:00.000Z", mensaje: "Correo simulado a Laura Restrepo: tienes una nueva solicitud SOL-1002 pendiente de aprobación." }],
  };
  return [s1, s2];
}

/* ---------------------------------------------------------
   UI GENÉRICOS
--------------------------------------------------------- */
function Badge({ children, tone = "slate", title }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700 border-slate-200", green: "bg-emerald-50 text-emerald-700 border-emerald-200",
    amber: "bg-amber-50 text-amber-700 border-amber-200", red: "bg-rose-50 text-rose-700 border-rose-200", blue: "bg-blue-50 text-blue-700 border-blue-200",
  };
  return <span title={title} className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium border cursor-default ${tones[tone]}`}>{children}</span>;
}

function Stepper({ status }) {
  const idx = PASOS.findIndex((p) => p.key === status);
  return (
    <div className="flex flex-wrap gap-2">
      {PASOS.map((p, i) => {
        const done = i < idx, current = i === idx;
        return (
          <div key={p.key} className={`flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium border
            ${current ? "bg-indigo-600 text-white border-indigo-600" : done ? "bg-indigo-50 text-indigo-700 border-indigo-200" : "bg-slate-50 text-slate-400 border-slate-200"}`}>
            {done ? <CheckCircle2 size={13} /> : current ? <Clock size={13} /> : null}{p.label}
          </div>
        );
      })}
    </div>
  );
}

function FirmaBlock({ rol, firma }) {
  if (!firma?.nombre) return (
    <div className="border border-dashed border-slate-200 rounded-lg p-3 text-center text-xs text-slate-400">
      <PenTool size={14} className="mx-auto mb-1" /> Firma {rol} pendiente
    </div>
  );
  return (
    <div className="border border-slate-200 rounded-lg p-3">
      <div className="text-[11px] text-slate-400 mb-1">Firma {rol}</div>
      {firma.fotoUrl ? (
        <ImagenPrivada path={firma.fotoUrl} alt={`firma ${rol}`} className="h-10 object-contain mb-1" />
      ) : (
        <div className="font-serif italic text-slate-700 text-base border-b border-slate-300 pb-1 mb-1">{firma.nombre}</div>
      )}
      <div className="text-[11px] text-slate-400">{firma.nombre}{firma.cargo && ` · ${firma.cargo}`}{firma.empresa && ` · ${firma.empresa}`} · {firma.fecha}{firma.aprobado === false ? " · Rechazado" : firma.aprobado ? " · Aprobado" : ""}</div>
      {firma.observacion && <div className="text-xs text-slate-500 mt-1 italic">"{firma.observacion}"</div>}
    </div>
  );
}

// input de archivo: sube de verdad a Supabase Storage y guarda la URL pública resultante
// muestra una imagen guardada en el bucket privado, resolviendo su URL temporal al montar
function ImagenPrivada({ path, alt, className }) {
  const [url, setUrl] = useState(null);
  useEffect(() => {
    let vivo = true;
    if (path) obtenerUrlFirmada(path).then((u) => { if (vivo) setUrl(u); });
    else setUrl(null);
    return () => { vivo = false; };
  }, [path]);
  if (!path) return null;
  return url ? <img src={url} alt={alt || ""} className={className} /> : <div className={`${className} bg-slate-100 animate-pulse rounded`} />;
}

// enlace que resuelve la URL temporal justo al hacer clic (no queda expuesta en el HTML)
function EnlacePrivado({ path, children, className, title }) {
  const [cargando, setCargando] = useState(false);
  const abrir = async () => {
    if (!path) return;
    setCargando(true);
    const url = await obtenerUrlFirmada(path);
    setCargando(false);
    if (url) window.open(url, "_blank", "noopener,noreferrer");
    else alert("No se pudo abrir el archivo.");
  };
  return <button type="button" onClick={abrir} disabled={cargando} title={title} className={className}>{cargando ? "..." : children}</button>;
}

function AdjuntarArchivo({ nombre, onSeleccionar, label, small, carpeta, soloPdf }) {
  const [subiendo, setSubiendo] = useState(false);
  const manejar = async (file) => {
    if (soloPdf && file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) { alert("Este archivo debe ser un PDF."); return; }
    if (!archivoDentroDelLimite(file)) { alert(`El archivo pesa más de ${TAMANO_MAXIMO_MB} MB. Sube uno más liviano.`); return; }
    setSubiendo(true);
    const ruta = await subirArchivo(file, carpeta || "adjuntos");
    setSubiendo(false);
    if (ruta) onSeleccionar(ruta);
  };
  // "nombre" ahora es la ruta guardada en Storage; mostramos solo el nombre del archivo (sin el prefijo de fecha)
  const mostrar = nombre ? decodeURIComponent(nombre.split("/").pop().replace(/^\d+_/, "")) : null;
  return (
    <span className="inline-flex items-center gap-1.5">
      <label className={`inline-flex items-center gap-1.5 border border-dashed border-slate-300 rounded-md px-2 py-1 cursor-pointer text-slate-500 hover:border-indigo-400 hover:text-indigo-600 ${small ? "text-[11px]" : "text-xs"}`}>
        <Paperclip size={small ? 11 : 13} />
        {subiendo ? <span>Subiendo...</span> : mostrar ? <span className="truncate max-w-[120px]">{mostrar}</span> : <span>{label || "Adjuntar archivo"}</span>}
        <input type="file" accept={soloPdf ? ".pdf" : ".pdf,image/*"} className="hidden" disabled={subiendo} onChange={(e) => e.target.files[0] && manejar(e.target.files[0])} />
      </label>
      {nombre && !subiendo && <EnlacePrivado path={nombre} className={`text-indigo-600 underline ${small ? "text-[11px]" : "text-xs"}`}>ver</EnlacePrivado>}
    </span>
  );
}

function CrudTable({ titulo, icon: Icon, columnas, datos, onGuardar, onEliminar, plantilla, currentUser }) {
  const [editId, setEditId] = useState(null);
  const [form, setForm] = useState(plantilla);
  const [creando, setCreando] = useState(false);
  const [mensajeImport, setMensajeImport] = useState("");
  const [seleccionados, setSeleccionados] = useState([]);
  const iniciarEdicion = (fila) => { setEditId(fila.id); setForm(fila); setCreando(false); };
  const iniciarCreacion = () => { setEditId(null); setForm(plantilla); setCreando(true); };
  const primerCampoVacio = !String(form[columnas[0]?.key] || "").trim() || columnas.some((c) => c.requerido && !String(form[c.key] || "").trim());
  const guardar = () => { if (primerCampoVacio) return; onGuardar(editId ? { ...form, id: editId } : { ...form, id: nextId() }); setEditId(null); setCreando(false); setForm(plantilla); };
  const cancelar = () => { setEditId(null); setCreando(false); setForm(plantilla); };
  const Campo = (c) => {
    const bloqueado = c.soloAdmin && currentUser?.rol !== "Administrador";
    if (bloqueado) {
      const valorMostrado = c.type === "select" ? (c.options.find((o) => o.value === form[c.key])?.label || "—") : (form[c.key] || "—");
      return <div title="Solo un Administrador puede cambiar este campo" className="border border-slate-100 bg-slate-50 rounded-md px-2 py-1 text-xs w-full text-slate-400">{valorMostrado}</div>;
    }
    if (c.type === "multiselect") {
      const valores = form[c.key] || [];
      const toggle = (v) => setForm({ ...form, [c.key]: valores.includes(v) ? valores.filter((x) => x !== v) : [...valores, v] });
      return (
        <div className="border border-slate-200 rounded-md px-2 py-1 max-h-24 overflow-y-auto space-y-0.5">
          {c.options.map((o) => (
            <label key={o.value} className="flex items-center gap-1.5 text-xs">
              <input type="checkbox" checked={valores.includes(o.value)} onChange={() => toggle(o.value)} /> {o.label}
            </label>
          ))}
        </div>
      );
    }
    return c.type === "select" ? (
      <select value={form[c.key] || ""} onChange={(e) => setForm({ ...form, [c.key]: e.target.value })} className="border border-slate-200 rounded-md px-2 py-1 text-xs w-full">
        <option value="">—</option>{c.options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
      </select>
    ) : (
      <input type={c.type || "text"} value={form[c.key] || ""} onChange={(e) => setForm({ ...form, [c.key]: e.target.value })} className="border border-slate-200 rounded-md px-2 py-1 text-xs w-full" />
    );
  };

  const todosSeleccionados = datos.length > 0 && seleccionados.length === datos.length;
  const alternarTodos = () => setSeleccionados(todosSeleccionados ? [] : datos.map((f) => f.id));
  const alternarUno = (id) => setSeleccionados((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const eliminarSeleccionados = () => {
    if (!seleccionados.length) return;
    if (!window.confirm(`¿Eliminar ${seleccionados.length} registro(s) seleccionado(s)? Esta acción no se puede deshacer.`)) return;
    seleccionados.forEach((id) => onEliminar(id));
    setSeleccionados([]);
  };

  const descargarPlantilla = () => {
    const encabezado = columnas.map((c) => c.label).join(",");
    // fila de ejemplo: para columnas de selección, muestra el nombre esperado (ej. una empresa real)
    // en vez de dejarlo en blanco, para que quede claro que se escribe el nombre, no el ID
    const ejemplo = columnas.map((c) => (c.type === "select" && c.options?.length ? c.options[0].label : "")).join(",");
    const csv = "\uFEFF" + encabezado + "\n" + ejemplo + "\n";
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `plantilla_${titulo.toLowerCase().replace(/\s+/g, "_")}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  const importarCSV = (file) => {
    // el archivo se lee como bytes para decidir la codificación: los CSV guardados desde Excel en Windows
    // suelen venir en ANSI (Windows-1252) y, leídos como UTF-8, dañan las tildes y las Ñ ("NAVIDEÑOS" → "NAVIDE�OS")
    const lector = new FileReader();
    lector.onerror = () => { setMensajeImport("No se pudo leer el archivo CSV."); setTimeout(() => setMensajeImport(""), 4000); };
    lector.onload = () => {
      const bytes = new Uint8Array(lector.result);
      let texto;
      try { texto = new TextDecoder("utf-8", { fatal: true }).decode(bytes); }
      catch { texto = new TextDecoder("windows-1252").decode(bytes); }
      texto = texto.replace(/^\uFEFF/, "");
      Papa.parse(texto, {
        header: true, skipEmptyLines: true, transformHeader: (h) => h.trim(),
        complete: (res) => {
          let importados = 0, duplicados = 0;
          const sinCoincidir = [];
          // un registro idéntico a uno que ya existe (o repetido en el mismo archivo) no se vuelve a cargar
          const firma = (o) => columnas.map((c) => String(o[c.key] ?? "").trim().toLowerCase()).join("|");
          const existentes = new Set(datos.map(firma));
          res.data.forEach((fila) => {
            const nueva = { id: nextId() };
            columnas.forEach((c) => {
              // el encabezado puede ser el nombre técnico (empresaId) o el título visible (Empresa)
              const valorCrudo = (fila[c.key] ?? fila[c.label] ?? "").toString().trim();
              if (c.type === "select" && valorCrudo) {
                // en un CSV es más práctico escribir el nombre (ej. "SP Dique") que el ID interno
                const porNombre = c.options.find((o) => o.label.trim().toLowerCase() === valorCrudo.toLowerCase());
                if (porNombre) nueva[c.key] = porNombre.value;
                else if (c.options.some((o) => o.value === valorCrudo)) nueva[c.key] = valorCrudo;
                else { nueva[c.key] = ""; sinCoincidir.push(`"${valorCrudo}" (${c.label})`); }
              } else {
                nueva[c.key] = valorCrudo;
              }
            });
            if (!Object.values(nueva).some((v) => v && v !== nueva.id)) return;
            const f = firma(nueva);
            if (existentes.has(f)) { duplicados++; return; }
            existentes.add(f);
            onGuardar(nueva); importados++;
          });
          const avisoSinCoincidir = sinCoincidir.length ? ` ${sinCoincidir.length} valor(es) no coincidieron con ninguna opción y quedaron vacíos: ${[...new Set(sinCoincidir)].slice(0, 5).join(", ")}${sinCoincidir.length > 5 ? "..." : ""}.` : "";
          const avisoDuplicados = duplicados ? ` ${duplicados} registro(s) repetido(s) se omitieron.` : "";
          setMensajeImport(`${importados} registro(s) importado(s) correctamente.${avisoDuplicados}${avisoSinCoincidir}`);
          setTimeout(() => setMensajeImport(""), 9000);
        },
        error: () => { setMensajeImport("No se pudo leer el archivo CSV."); setTimeout(() => setMensajeImport(""), 4000); },
      });
    };
    lector.readAsArrayBuffer(file);
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
      <div className="px-5 py-3 border-b border-slate-100 flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-2 font-medium text-slate-700"><Icon size={16} /> {titulo}</div>
        <div className="flex items-center gap-2">
          {seleccionados.length > 0 && (
            <button onClick={eliminarSeleccionados} className="text-xs text-rose-600 font-medium flex items-center gap-1 bg-rose-50 border border-rose-200 rounded-md px-2 py-1"><Trash2 size={12} /> Eliminar {seleccionados.length} seleccionado{seleccionados.length > 1 ? "s" : ""}</button>
          )}
          <button onClick={descargarPlantilla} className="text-xs text-slate-500 font-medium flex items-center gap-1"><FileText size={12} /> Plantilla CSV</button>
          <label className="text-xs text-indigo-600 font-medium flex items-center gap-1 cursor-pointer"><UploadIcon size={12} /> Importar CSV
            <input type="file" accept=".csv" className="hidden" onChange={(e) => e.target.files[0] && importarCSV(e.target.files[0])} />
          </label>
          {!creando && <button onClick={iniciarCreacion} className="text-xs text-indigo-600 font-medium flex items-center gap-1"><Plus size={13} /> Agregar</button>}
        </div>
      </div>
      {mensajeImport && <div className="px-5 py-1.5 text-[11px] text-emerald-600 bg-emerald-50 border-b border-emerald-100">{mensajeImport}</div>}
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-slate-500 text-xs"><tr>
          <th className="px-4 py-2 w-8"><input type="checkbox" checked={todosSeleccionados} onChange={alternarTodos} /></th>
          {columnas.map((c) => <th key={c.key} className="text-left px-4 py-2 font-medium">{c.label}</th>)}<th></th></tr></thead>
        <tbody>
          {creando && (
            <tr className="border-t border-slate-100 bg-indigo-50/40">
              <td className="px-4 py-1.5"></td>
              {columnas.map((c) => <td key={c.key} className="px-4 py-1.5">{Campo(c)}</td>)}
              <td className="px-4 py-1.5 text-right whitespace-nowrap"><button onClick={guardar} disabled={primerCampoVacio} className="text-emerald-600 text-xs font-medium mr-2 disabled:opacity-40 disabled:cursor-not-allowed">Guardar</button><button onClick={cancelar} className="text-slate-400 text-xs">Cancelar</button></td>
            </tr>
          )}
          {datos.map((fila) => editId === fila.id ? (
            <tr key={fila.id} className="border-t border-slate-100 bg-indigo-50/40">
              <td className="px-4 py-1.5"></td>
              {columnas.map((c) => <td key={c.key} className="px-4 py-1.5">{Campo(c)}</td>)}
              <td className="px-4 py-1.5 text-right whitespace-nowrap"><button onClick={guardar} disabled={primerCampoVacio} className="text-emerald-600 text-xs font-medium mr-2 disabled:opacity-40 disabled:cursor-not-allowed">Guardar</button><button onClick={cancelar} className="text-slate-400 text-xs">Cancelar</button></td>
            </tr>
          ) : (
            <tr key={fila.id} className={`border-t border-slate-100 ${seleccionados.includes(fila.id) ? "bg-indigo-50/30" : ""}`}>
              <td className="px-4 py-2"><input type="checkbox" checked={seleccionados.includes(fila.id)} onChange={() => alternarUno(fila.id)} /></td>
              {columnas.map((c) => <td key={c.key} className="px-4 py-2 text-slate-600">{c.type === "select" ? (c.options.find((o) => o.value === fila[c.key])?.label || "—") : c.type === "multiselect" ? ((fila[c.key] || []).map((v) => c.options.find((o) => o.value === v)?.label).filter(Boolean).join(", ") || "—") : (fila[c.key] || "—")}</td>)}
              <td className="px-4 py-2 text-right whitespace-nowrap"><button onClick={() => iniciarEdicion(fila)} className="text-slate-400 hover:text-indigo-600 p-1"><Pencil size={13} /></button><button onClick={() => onEliminar(fila.id)} className="text-slate-400 hover:text-rose-500 p-1"><Trash2 size={13} /></button></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ---------------------------------------------------------
   LOGIN: ahora se usa LoginReal.jsx (Supabase Auth) — ver App()
--------------------------------------------------------- */

/* ---------------------------------------------------------
   MI PERFIL (firma tipo foto)
--------------------------------------------------------- */
function PerfilUsuario({ currentUser, onGuardar }) {
  const [preview, setPreview] = useState(currentUser.firmaFotoUrl);
  const [subiendo, setSubiendo] = useState(false);
  const cargarFoto = async (file) => {
    if (!archivoDentroDelLimite(file)) { alert(`El archivo pesa más de ${TAMANO_MAXIMO_MB} MB. Sube uno más liviano.`); return; }
    setSubiendo(true);
    const ruta = await subirArchivo(file, "firmas");
    setSubiendo(false);
    if (ruta) setPreview(ruta);
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-6 max-w-md">
      <div className="flex items-center gap-2 mb-4"><UserCircle size={18} className="text-indigo-600" /><h2 className="text-lg font-semibold text-slate-800">Mi perfil</h2></div>
      <div className="text-sm text-slate-600 mb-1"><b>{currentUser.nombre}</b></div>
      <div className="text-xs text-slate-400 mb-4">{currentUser.cargo} · {currentUser.rol}</div>
      <label className="text-xs font-medium text-slate-500">Firma (foto)</label>
      <div className="border border-dashed border-slate-300 rounded-lg p-4 mt-1 text-center">
        {preview ? <ImagenPrivada path={preview} alt="firma" className="h-20 mx-auto object-contain mb-2" /> : <div className="text-xs text-slate-400 mb-2">Sin firma cargada. Sube una foto de tu firma en papel.</div>}
        <label className="inline-flex items-center gap-1.5 text-xs text-indigo-600 font-medium cursor-pointer"><Camera size={13} /> {subiendo ? "Subiendo..." : preview ? "Cambiar foto" : "Subir foto"}
          <input type="file" accept="image/*" className="hidden" disabled={subiendo} onChange={(e) => e.target.files[0] && cargarFoto(e.target.files[0])} />
        </label>
      </div>
      <button onClick={() => onGuardar({ ...currentUser, firmaFotoUrl: preview })} className="mt-4 bg-indigo-600 text-white rounded-lg px-4 py-2 text-sm font-medium">Guardar perfil</button>
    </div>
  );
}

/* ---------------------------------------------------------
   DASHBOARD
--------------------------------------------------------- */
function Dashboard({ areas, solicitudes, proveedores, currentUser, onAbrir, onVerCalendario }) {
  const gastoPorArea = useMemo(() => {
    const map = {}; areas.forEach((a) => (map[a.id] = 0));
    solicitudes.forEach((s) => { if (!["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)) map[s.areaId] = (map[s.areaId] || 0) + totalSolicitud(s); });
    return map;
  }, [areas, solicitudes]);

  // mismo cálculo que "Calendario de pagos", pero solo los 5 más próximos sin pagar
  const proximosPagos = useMemo(() => {
    if (!puedeVerCalendarioPagos(currentUser)) return [];
    const filas = [];
    solicitudes.forEach((s) => {
      if (["completada", "rechazada"].includes(s.status)) return;
      const prov = proveedoresAdjudicados(s, proveedores);
      filas.push(...filasPagosSolicitud(s, prov).filter((f) => !f.pagado));
    });
    filas.sort((a, b) => a.fecha.localeCompare(b.fecha));
    return filas.slice(0, 5);
  }, [solicitudes, proveedores, currentUser]);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="flex items-center gap-2 text-slate-500 text-sm mb-1"><FileText size={16} /> Solicitudes activas</div><div className="text-2xl font-semibold text-slate-800">{solicitudes.filter((s) => s.status !== "completada" && s.status !== "rechazada").length}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="flex items-center gap-2 text-slate-500 text-sm mb-1"><Clock size={16} /> Pendientes de aprobación</div><div className="text-2xl font-semibold text-slate-800">{solicitudes.filter((s) => ["aprobacion_jefe", "aprobacion_financiera", "aprobacion_gerencia"].includes(s.status)).length}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="flex items-center gap-2 text-slate-500 text-sm mb-1"><DollarSign size={16} /> Comprometido este mes (con IVA)</div><div className="text-2xl font-semibold text-slate-800">{fmt(Object.values(gastoPorArea).reduce((a, b) => a + b, 0))}</div></div>
      </div>

      {puedeVerCalendarioPagos(currentUser) && (
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between">
            <div className="font-medium text-slate-700 flex items-center gap-2"><CalendarClock size={16} /> Próximos pagos</div>
            {onVerCalendario && <button onClick={onVerCalendario} className="text-xs text-indigo-600 font-medium">Ver calendario completo →</button>}
          </div>
          {proximosPagos.length === 0 ? (
            <div className="px-5 py-6 text-sm text-slate-400 text-center">No hay pagos programados por vencer.</div>
          ) : (
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-slate-500 text-xs"><tr><th className="text-left px-5 py-2 font-medium">Fecha</th><th className="text-left px-5 py-2 font-medium">Días</th><th className="text-left px-5 py-2 font-medium">Consecutivo</th><th className="text-left px-5 py-2 font-medium">Proveedor</th><th className="text-left px-5 py-2 font-medium">Tipo</th><th className="text-right px-5 py-2 font-medium">Valor</th></tr></thead>
              <tbody>{proximosPagos.map((f) => (
                <tr key={f.id} onClick={() => onAbrir?.(f.solicitudId)} className="border-t border-slate-100 hover:bg-slate-50 cursor-pointer">
                  <td className="px-5 py-2.5 whitespace-nowrap">{f.fecha}</td>
                  <td className={`px-5 py-2.5 whitespace-nowrap font-medium ${f.dias < 0 ? "text-rose-600" : f.dias <= 3 ? "text-amber-600" : "text-slate-500"}`}>{f.dias < 0 ? `Vencido ${Math.abs(f.dias)}d` : f.dias === 0 ? "Hoy" : `${f.dias}d`}</td>
                  <td className="px-5 py-2.5 font-medium text-slate-700">{f.folio}</td>
                  <td className="px-5 py-2.5 text-slate-600 max-w-[200px] truncate" title={f.proveedor}>{f.proveedor}</td>
                  <td className="px-5 py-2.5 text-slate-600">{f.tipo}</td>
                  <td className="px-5 py-2.5 text-right font-medium">{fmt(f.valor)}</td>
                </tr>
              ))}</tbody>
            </table>
          )}
        </div>
      )}

      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-100 font-medium text-slate-700">Presupuesto mensual por área</div>
        <table className="w-full text-sm">
          <thead className="bg-slate-50 text-slate-500"><tr><th className="text-left px-5 py-2 font-medium">Área</th><th className="text-right px-5 py-2 font-medium">Presupuesto</th><th className="text-right px-5 py-2 font-medium">Comprometido</th><th className="text-right px-5 py-2 font-medium">Disponible</th><th className="px-5 py-2 font-medium">% Uso</th></tr></thead>
          <tbody>{areas.map((a) => { const gastado = gastoPorArea[a.id] || 0, disponible = a.presupuesto - gastado, pct = Math.min(100, (gastado / a.presupuesto) * 100);
            return (<tr key={a.id} className="border-t border-slate-100"><td className="px-5 py-3 text-slate-700 font-medium">{a.nombre}</td><td className="px-5 py-3 text-right text-slate-600">{fmt(a.presupuesto)}</td><td className="px-5 py-3 text-right text-slate-600">{fmt(gastado)}</td><td className={`px-5 py-3 text-right font-medium ${disponible < 0 ? "text-rose-600" : "text-emerald-600"}`}>{fmt(disponible)}</td><td className="px-5 py-3"><div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden"><div className={`h-2 rounded-full ${pct > 90 ? "bg-rose-500" : pct > 70 ? "bg-amber-500" : "bg-emerald-500"}`} style={{ width: `${pct}%` }} /></div></td></tr>); })}</tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   ESTADÍSTICAS (con filtro por empresa)
--------------------------------------------------------- */
/* ---------------------------------------------------------
   REPORTE: promedio de evaluaciones de un proveedor en un rango de fechas
   (para auditorías ISO 9001)
--------------------------------------------------------- */
/* ---------------------------------------------------------
   CALENDARIO DE PAGOS — solo Dirección Financiera y Compras
--------------------------------------------------------- */
/* ---------------------------------------------------------
   ÓRDENES ENVIADAS A PROVEEDORES — consolidado para contabilidad
--------------------------------------------------------- */
function ReporteOrdenesEnviadas({ solicitudes, proveedores, empresas, onAbrir }) {
  const filas = [];
  solicitudes.forEach((s) => {
    if (!["oc_enviada", "recepcion", "completada"].includes(s.status)) return;
    (s.ocEnviada?.ordenesProveedor || []).forEach((o) => {
      if (!o.archivoFirmadoUrl) return;
      filas.push({
        id: `${s.id}-${o.proveedorId || o.proveedorNombre}`,
        solicitudId: s.id,
        folio: s.folio,
        tipo: s.tipo === "compra" ? "Solicitud de compra" : "Orden de servicio/trabajo",
        empresa: empresas.find((e) => e.id === s.empresaId)?.nombre || "",
        proveedor: o.proveedorNombre,
        fecha: o.fecha,
        total: totalSolicitud(s),
        archivo: o.archivoFirmadoUrl,
      });
    });
  });
  filas.sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")); // más reciente primero

  const descargar = () => {
    const encabezado = ["Fecha envío", "Consecutivo", "Tipo", "Empresa", "Proveedor", "Total solicitud"];
    const cuerpo = filas.map((f) => [f.fecha, f.folio, f.tipo, f.empresa, f.proveedor, f.total]);
    const hoja = XLSX.utils.aoa_to_sheet([encabezado, ...cuerpo]);
    hoja["!cols"] = [{ wch: 14 }, { wch: 14 }, { wch: 22 }, { wch: 18 }, { wch: 30 }, { wch: 16 }];
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Órdenes enviadas");
    XLSX.writeFile(libro, `Ordenes_enviadas_${hoy()}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Órdenes enviadas a proveedores</h2>
          <p className="text-xs text-slate-400 mt-1">Consolidado de todas las órdenes ya firmadas y enviadas — para registrar en el sistema contable.</p>
        </div>
        <button onClick={descargar} disabled={!filas.length} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40 flex items-center gap-1"><FileText size={13} /> Descargar Excel</button>
      </div>

      {filas.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">Todavía no hay órdenes enviadas al proveedor.</div>
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs"><tr>
              <th className="text-left px-4 py-2 font-medium">Fecha envío</th>
              <th className="text-left px-4 py-2 font-medium">Consecutivo</th>
              <th className="text-left px-4 py-2 font-medium">Tipo</th>
              <th className="text-left px-4 py-2 font-medium">Empresa</th>
              <th className="text-left px-4 py-2 font-medium">Proveedor</th>
              <th className="text-right px-4 py-2 font-medium">Total solicitud</th>
              <th className="text-center px-4 py-2 font-medium">PDF</th>
            </tr></thead>
            <tbody>{filas.map((f) => (
              <tr key={f.id} className="border-t border-slate-100">
                <td className="px-4 py-2 whitespace-nowrap">{f.fecha}</td>
                <td className="px-4 py-2 font-medium text-slate-700 cursor-pointer hover:text-indigo-600" onClick={() => onAbrir?.(f.solicitudId)}>{f.folio}</td>
                <td className="px-4 py-2 text-slate-600">{f.tipo}</td>
                <td className="px-4 py-2 text-slate-600">{f.empresa}</td>
                <td className="px-4 py-2 text-slate-600">{f.proveedor}</td>
                <td className="px-4 py-2 text-right font-medium">{fmt(f.total)}</td>
                <td className="px-4 py-2 text-center"><EnlacePrivado path={f.archivo} className="text-indigo-600 underline text-xs">Descargar</EnlacePrivado></td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function CalendarioPagos({ solicitudes, proveedores, onAbrir }) {
  const [seleccionados, setSeleccionados] = useState([]);
  const [vista, setVista] = useState("lista");
  const [mesActual, setMesActual] = useState(() => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), 1); });
  const [diaExpandido, setDiaExpandido] = useState(null);

  // arma una fila por cada pago confirmado (anticipo / intermedio / final) de cada solicitud tipo servicio;
  // las que no tienen plan de pagos (compras, o servicios sin confirmar) usan su fecha estimada de entrega
  const filas = [];
  solicitudes.forEach((s) => {
    if (["completada", "rechazada"].includes(s.status)) return;
    const prov = proveedoresAdjudicados(s, proveedores);
    filas.push(...filasPagosSolicitud(s, prov));
  });
  filas.sort((a, b) => a.fecha.localeCompare(b.fecha)); // más antiguo primero

  const alternar = (id) => setSeleccionados((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const todosSeleccionados = filas.length > 0 && seleccionados.length === filas.length;
  const alternarTodos = () => setSeleccionados(todosSeleccionados ? [] : filas.map((f) => f.id));
  const totalSeleccionado = filas.filter((f) => seleccionados.includes(f.id)).reduce((acc, f) => acc + f.valor, 0);

  const textoDias = (dias, pagado) => {
    if (pagado) return "Pagado";
    if (dias < 0) return `Vencido hace ${Math.abs(dias)} día${Math.abs(dias) === 1 ? "" : "s"}`;
    if (dias === 0) return "Hoy";
    return `Faltan ${dias} día${dias === 1 ? "" : "s"}`;
  };
  const colorDias = (dias, pagado) => pagado ? "text-slate-400" : dias < 0 ? "text-rose-600" : dias <= 3 ? "text-amber-600" : "text-slate-600";

  const descargar = () => {
    const encabezado = ["Fecha", "Días", "Consecutivo", "Proveedor", "Tipo de pago", "Valor", "Pagado"];
    const cuerpo = filas.map((f) => [f.fecha, f.pagado ? "" : f.dias, f.folio, f.proveedor, f.tipo, f.valor, f.pagado ? "Sí" : "No"]);
    const hoja = XLSX.utils.aoa_to_sheet([encabezado, ...cuerpo]);
    hoja["!cols"] = [{ wch: 14 }, { wch: 8 }, { wch: 14 }, { wch: 30 }, { wch: 14 }, { wch: 16 }, { wch: 10 }];
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Pagos");
    XLSX.writeFile(libro, `Calendario_pagos_${hoy()}.xlsx`);
  };

  // ---- vista de calendario (cuadrícula del mes) ----
  const nombreMes = mesActual.toLocaleDateString("es-CO", { month: "long", year: "numeric" });
  const primerDiaSemana = (mesActual.getDay() + 6) % 7; // que la semana empiece en lunes
  const diasDelMes = new Date(mesActual.getFullYear(), mesActual.getMonth() + 1, 0).getDate();
  const fechaISO = (dia) => `${mesActual.getFullYear()}-${String(mesActual.getMonth() + 1).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;
  const pagosPorDia = {};
  filas.forEach((f) => { (pagosPorDia[f.fecha] = pagosPorDia[f.fecha] || []).push(f); });
  const celdas = [...Array(primerDiaSemana).fill(null), ...Array.from({ length: diasDelMes }, (_, i) => i + 1)];

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Calendario de pagos</h2>
          <p className="text-xs text-slate-400 mt-1">Pagos con plan confirmado, y solicitudes sin plan de pagos (usan su fecha estimada de entrega) — ordenados del más antiguo al más reciente.</p>
        </div>
        <div className="flex items-center gap-2">
          <div className="flex border border-slate-200 rounded-md overflow-hidden">
            <button onClick={() => setVista("lista")} className={`text-xs px-3 py-1.5 font-medium ${vista === "lista" ? "bg-indigo-600 text-white" : "bg-white text-slate-600"}`}>Lista</button>
            <button onClick={() => setVista("calendario")} className={`text-xs px-3 py-1.5 font-medium ${vista === "calendario" ? "bg-indigo-600 text-white" : "bg-white text-slate-600"}`}>Calendario</button>
          </div>
          <button onClick={descargar} disabled={!filas.length} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40 flex items-center gap-1"><FileText size={13} /> Descargar Excel</button>
        </div>
      </div>

      {seleccionados.length > 0 && (
        <div className="bg-indigo-50 border border-indigo-200 rounded-xl px-4 py-3 flex items-center justify-between">
          <span className="text-sm text-slate-700">{seleccionados.length} pago(s) seleccionado(s)</span>
          <span className="text-lg font-semibold text-indigo-700">{fmt(totalSeleccionado)}</span>
        </div>
      )}

      {filas.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">No hay pagos programados con plan confirmado todavía.</div>
      ) : vista === "calendario" ? (
        <div className="bg-white rounded-xl border border-slate-200 p-4">
          <div className="flex items-center justify-between mb-3">
            <button onClick={() => setMesActual(new Date(mesActual.getFullYear(), mesActual.getMonth() - 1, 1))} className="p-1.5 rounded-md hover:bg-slate-100 text-slate-500"><ChevronRight size={16} className="rotate-180" /></button>
            <div className="text-sm font-medium text-slate-700 capitalize">{nombreMes}</div>
            <button onClick={() => setMesActual(new Date(mesActual.getFullYear(), mesActual.getMonth() + 1, 1))} className="p-1.5 rounded-md hover:bg-slate-100 text-slate-500"><ChevronRight size={16} /></button>
          </div>
          <div className="grid grid-cols-7 gap-1 text-center text-[10px] font-medium text-slate-400 mb-1">
            {["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"].map((d) => <div key={d}>{d}</div>)}
          </div>
          <div className="grid grid-cols-7 gap-1">
            {celdas.map((dia, idx) => {
              if (!dia) return <div key={idx} />;
              const fechaDia = fechaISO(dia);
              const pagosDia = pagosPorDia[fechaDia] || [];
              const totalDia = pagosDia.reduce((a, f) => a + f.valor, 0);
              const esHoy = fechaDia === hoy();
              const hayVencido = pagosDia.some((f) => !f.pagado && f.dias < 0);
              return (
                <div
                  key={idx}
                  onClick={() => pagosDia.length > 0 && setDiaExpandido(diaExpandido === fechaDia ? null : fechaDia)}
                  className={`min-h-[76px] rounded-md border p-1 text-left ${pagosDia.length > 0 ? "cursor-pointer hover:border-indigo-300" : ""} ${diaExpandido === fechaDia ? "border-indigo-500 ring-1 ring-indigo-300" : esHoy ? "border-indigo-400 bg-indigo-50/40" : "border-slate-100"}`}
                >
                  <div className={`text-[11px] mb-1 ${esHoy ? "font-semibold text-indigo-600" : "text-slate-400"}`}>{dia}</div>
                  {pagosDia.slice(0, 2).map((f) => (
                    <button key={f.id} onClick={(e) => { e.stopPropagation(); onAbrir?.(f.solicitudId); }} title={`${f.folio} — ${f.tipo} — ${fmt(f.valor)}`} className={`block w-full text-left text-[10px] truncate rounded px-1 py-0.5 mb-0.5 ${f.pagado ? "bg-slate-100 text-slate-400" : hayVencido ? "bg-rose-100 text-rose-700" : "bg-indigo-100 text-indigo-700"}`}>
                      {f.folio}
                    </button>
                  ))}
                  {pagosDia.length > 2 && <div className="text-[9px] text-slate-400">+{pagosDia.length - 2} más</div>}
                  {pagosDia.length > 0 && <div className="text-[9px] font-medium text-slate-500 mt-0.5">{fmt(totalDia)}</div>}
                </div>
              );
            })}
          </div>

          {diaExpandido && pagosPorDia[diaExpandido] && (
            <div className="mt-4 border-t border-slate-100 pt-3">
              <div className="flex items-center justify-between mb-2">
                <div className="text-sm font-medium text-slate-700">Pagos del {diaExpandido}</div>
                <button onClick={() => setDiaExpandido(null)} className="text-slate-400 hover:text-slate-600 text-xs">✕ Cerrar</button>
              </div>
              <div className="space-y-1.5">
                {pagosPorDia[diaExpandido].map((f) => (
                  <div key={f.id} onClick={() => onAbrir?.(f.solicitudId)} className={`flex items-center justify-between border rounded-md px-3 py-2 cursor-pointer hover:bg-slate-50 ${f.pagado ? "opacity-50 border-slate-100" : "border-slate-200"}`}>
                    <div className="flex items-center gap-3">
                      <span className="text-sm font-medium text-indigo-600">{f.folio}</span>
                      <span className="text-xs text-slate-500">{f.proveedor}</span>
                      <span className="text-xs text-slate-400">{f.tipo}</span>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className={`text-xs font-medium ${colorDias(f.dias, f.pagado)}`}>{textoDias(f.dias, f.pagado)}</span>
                      <span className="text-sm font-medium text-slate-700">{fmt(f.valor)}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-slate-200 overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs"><tr>
              <th className="px-4 py-2 w-8"><input type="checkbox" checked={todosSeleccionados} onChange={alternarTodos} /></th>
              <th className="text-left px-4 py-2 font-medium">Fecha</th>
              <th className="text-left px-4 py-2 font-medium">Días</th>
              <th className="text-left px-4 py-2 font-medium">Consecutivo</th>
              <th className="text-left px-4 py-2 font-medium">Proveedor</th>
              <th className="text-left px-4 py-2 font-medium">Tipo de pago</th>
              <th className="text-right px-4 py-2 font-medium">Valor</th>
              <th className="text-center px-4 py-2 font-medium">Pagado</th>
            </tr></thead>
            <tbody>{filas.map((f) => (
              <tr key={f.id} className={`border-t border-slate-100 ${f.pagado ? "opacity-50" : ""} ${seleccionados.includes(f.id) ? "bg-indigo-50/40" : ""}`}>
                <td className="px-4 py-2" onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={seleccionados.includes(f.id)} onChange={() => alternar(f.id)} /></td>
                <td className="px-4 py-2 whitespace-nowrap">{f.fecha}</td>
                <td className={`px-4 py-2 whitespace-nowrap font-medium ${colorDias(f.dias, f.pagado)}`}>{textoDias(f.dias, f.pagado)}</td>
                <td className="px-4 py-2 font-medium text-slate-700 cursor-pointer hover:text-indigo-600" onClick={() => onAbrir?.(f.solicitudId)}>{f.folio}</td>
                <td className="px-4 py-2 text-slate-600">{f.proveedor}</td>
                <td className="px-4 py-2 text-slate-600">{f.tipo}</td>
                <td className="px-4 py-2 text-right font-medium">{fmt(f.valor)}</td>
                <td className="px-4 py-2 text-center">{f.pagado ? "✓" : "—"}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function ReporteEvaluacionesProveedores({ solicitudes, proveedores, onAbrir }) {
  const [vista, setVista] = useState("individual");
  const [filtro, setFiltro] = useState("");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");
  // rango de fechas propio del ranking general (independiente del filtro de la vista por proveedor)
  const [rDesde, setRDesde] = useState("");
  const [rHasta, setRHasta] = useState("");

  const claveProveedor = (ev) => ev.proveedorId ? `id:${ev.proveedorId}` : `nombre:${(ev.proveedorNombre || "").trim().toLowerCase()}`;

  const todasCompletadas = solicitudes.filter((s) => s.evaluacionProveedor?.completada);

  // ranking general: agrupa TODAS las evaluaciones por proveedor (sin filtro de fecha/proveedor),
  // calcula su promedio, y una tendencia comparando la mitad más reciente contra la más antigua
  const rankingProveedores = useMemo(() => {
    const grupos = {};
    todasCompletadas.forEach((s) => {
      const ev = s.evaluacionProveedor;
      if (!ev.proveedorId && !ev.proveedorNombre) return;
      if (rDesde && (ev.fechaCompletado || "") < rDesde) return;
      if (rHasta && (ev.fechaCompletado || "") > rHasta) return;
      const key = claveProveedor(ev);
      if (!grupos[key]) grupos[key] = { key, nombre: ev.proveedorNombre || proveedores.find((p) => p.id === ev.proveedorId)?.nombre || "Sin nombre", evaluaciones: [] };
      grupos[key].evaluaciones.push({ pct: puntajeEvaluacion(ev.criterios) * 100, fecha: ev.fechaCompletado || "" });
    });
    return Object.values(grupos).map((g) => {
      const ordenadas = [...g.evaluaciones].sort((a, b) => a.fecha.localeCompare(b.fecha));
      const promedio = ordenadas.reduce((a, e) => a + e.pct, 0) / ordenadas.length;
      let tendencia = null;
      if (ordenadas.length >= 2) {
        const mitad = Math.floor(ordenadas.length / 2);
        const promAntiguo = ordenadas.slice(0, mitad).reduce((a, e) => a + e.pct, 0) / mitad;
        const promReciente = ordenadas.slice(mitad).reduce((a, e) => a + e.pct, 0) / (ordenadas.length - mitad);
        const dif = promReciente - promAntiguo;
        tendencia = Math.abs(dif) < 3 ? "estable" : dif > 0 ? "subiendo" : "bajando";
      }
      return { ...g, cantidad: ordenadas.length, promedio, ultima: ordenadas[ordenadas.length - 1]?.fecha || "", tendencia };
    }).sort((a, b) => b.promedio - a.promedio);
  }, [todasCompletadas, proveedores, rDesde, rHasta]);
  // opciones del selector: se arman a partir de las evaluaciones mismas (por ID si lo tienen, si no por nombre)
  // así no se pierden proveedores cuyo registro no quedó vinculado por ID (datos de antes de esa mejora)
  const opciones = [];
  const vistos = new Set();
  todasCompletadas.forEach((s) => {
    const ev = s.evaluacionProveedor;
    if (!ev.proveedorId && !ev.proveedorNombre) return;
    const key = claveProveedor(ev);
    if (!vistos.has(key)) { vistos.add(key); opciones.push({ key, label: ev.proveedorNombre || proveedores.find((p) => p.id === ev.proveedorId)?.nombre || "Sin nombre" }); }
  });
  opciones.sort((a, b) => a.label.localeCompare(b.label));

  const evaluaciones = todasCompletadas.filter((s) => {
    const ev = s.evaluacionProveedor;
    if (filtro && claveProveedor(ev) !== filtro) return false;
    if (desde && ev.fechaCompletado < desde) return false;
    if (hasta && ev.fechaCompletado > hasta) return false;
    return true;
  });

  const pendientes = solicitudes.filter((s) => s.status === "recepcion" && !evaluacionProveedorCompleta(s.evaluacionProveedor));

  const promedioPct = evaluaciones.length
    ? evaluaciones.reduce((acc, s) => acc + puntajeEvaluacion(s.evaluacionProveedor.criterios), 0) / evaluaciones.length * 100
    : null;
  const clas = promedioPct !== null ? clasificacionConfianza(promedioPct) : null;

  const promediosPorCriterio = CRITERIOS_EVALUACION.map((c) => {
    const valores = evaluaciones.map((s) => parseFloat(s.evaluacionProveedor.criterios?.[c.key])).filter((v) => v > 0);
    return { ...c, promedio: valores.length ? valores.reduce((a, b) => a + b, 0) / valores.length : null };
  });

  const descargarReporte = () => {
    const filas = [];
    filas.push(["REPORTE PROMEDIO DE EVALUACIÓN DE PROVEEDOR — ISO 9001"]);
    filas.push([]);
    filas.push(["Proveedor:", filtro ? opciones.find((o) => o.key === filtro)?.label : "Todos"]);
    filas.push(["Rango de fechas:", desde || "sin límite", "a", hasta || "sin límite"]);
    filas.push(["Número de evaluaciones incluidas:", evaluaciones.length]);
    filas.push(["Resultado promedio (%):", promedioPct !== null ? promedioPct.toFixed(1) : "—"]);
    filas.push(["Clasificación:", clas ? clas.texto : "—"]);
    filas.push([]);
    filas.push(["CRITERIO", "PROMEDIO (1-10)"]);
    promediosPorCriterio.forEach((c) => filas.push([c.texto, c.promedio !== null ? c.promedio.toFixed(1) : "—"]));
    filas.push([]);
    filas.push(["DETALLE POR SOLICITUD"]);
    filas.push(["Consecutivo", "Fecha evaluación", "Resultado (%)", "Clasificación"]);
    evaluaciones.forEach((s) => {
      const pct = puntajeEvaluacion(s.evaluacionProveedor.criterios) * 100;
      filas.push([s.folio, s.evaluacionProveedor.fechaCompletado, pct.toFixed(1), clasificacionConfianza(pct).texto]);
    });
    const hoja = XLSX.utils.aoa_to_sheet(filas);
    hoja["!cols"] = [{ wch: 45 }, { wch: 18 }, { wch: 16 }, { wch: 18 }];
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Reporte");
    XLSX.writeFile(libro, `Reporte_Evaluacion_Proveedor_${hoy()}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Evaluación de proveedores</h2>
          <p className="text-xs text-slate-400 mt-1">{vista === "individual" ? "Promedio de resultados de un proveedor en un rango de fechas — útil para auditorías ISO 9001." : "Comparativo de desempeño entre todos los proveedores evaluados."}</p>
        </div>
        <div className="flex border border-slate-200 rounded-md overflow-hidden shrink-0">
          <button onClick={() => setVista("individual")} className={`text-xs px-3 py-1.5 font-medium ${vista === "individual" ? "bg-indigo-600 text-white" : "bg-white text-slate-600"}`}>Por proveedor</button>
          <button onClick={() => setVista("ranking")} className={`text-xs px-3 py-1.5 font-medium ${vista === "ranking" ? "bg-indigo-600 text-white" : "bg-white text-slate-600"}`}>Ranking general</button>
        </div>
      </div>

      {/* PENDIENTES DE EVALUACIÓN — se ve en ambas vistas */}
      {pendientes.length > 0 && (
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="font-medium text-slate-700 mb-3 flex items-center gap-2"><Clock size={15} /> Pendientes de evaluación ({pendientes.length})</div>
          <table className="w-full text-sm">
            <thead className="text-slate-400 text-xs border-b border-slate-100"><tr><th className="text-left py-1">Consecutivo</th><th className="text-left py-1">Objetivo</th><th className="text-left py-1">Proveedor adjudicado</th></tr></thead>
            <tbody>{pendientes.map((s) => (
              <tr key={s.id} onClick={() => onAbrir?.(s.id)} className="border-t border-slate-50 hover:bg-slate-50 cursor-pointer">
                <td className="py-1.5 font-medium text-slate-700">{s.folio}</td>
                <td className="py-1.5 text-slate-600 max-w-[300px] truncate" title={s.objetivo}>{s.objetivo}</td>
                <td className="py-1.5 text-slate-600">{s.evaluacionProveedor?.proveedorNombre || "—"}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}

      {vista === "ranking" ? (
        <>
        <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
          <div><label className="text-[11px] font-medium text-slate-500">Desde</label><input type="date" value={rDesde} onChange={(e) => setRDesde(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" /></div>
          <div><label className="text-[11px] font-medium text-slate-500">Hasta</label><input type="date" value={rHasta} onChange={(e) => setRHasta(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" /></div>
          {(rDesde || rHasta) && <button onClick={() => { setRDesde(""); setRHasta(""); }} className="text-xs text-slate-500 underline pb-2">Quitar filtro</button>}
          <div className="text-[11px] text-slate-400 pb-2">El ranking, los promedios y la tendencia se calculan solo con las evaluaciones de este periodo.</div>
        </div>
        {rankingProveedores.length === 0 ? (
          <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">{rDesde || rHasta ? "No hay evaluaciones completas en ese rango de fechas." : "Todavía no hay evaluaciones completas para comparar."}</div>
        ) : (
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-slate-50 text-slate-500 text-xs"><tr>
                <th className="text-left px-4 py-2">#</th>
                <th className="text-left px-4 py-2">Proveedor</th>
                <th className="text-right px-4 py-2">Evaluaciones</th>
                <th className="text-right px-4 py-2">Promedio</th>
                <th className="text-center px-4 py-2">Clasificación</th>
                <th className="text-center px-4 py-2">Tendencia</th>
                <th className="text-right px-4 py-2">Última evaluación</th>
                <th className="text-center px-4 py-2">Preferido</th>
              </tr></thead>
              <tbody>{rankingProveedores.map((g, i) => { const c = clasificacionConfianza(g.promedio);
                return (
                  <tr key={g.key} onClick={() => { setFiltro(g.key); setVista("individual"); }} className="border-t border-slate-100 hover:bg-slate-50 cursor-pointer">
                    <td className="px-4 py-2.5 text-slate-400">{i + 1}</td>
                    <td className="px-4 py-2.5 font-medium text-slate-700">{g.nombre}</td>
                    <td className="px-4 py-2.5 text-right text-slate-600">{g.cantidad}</td>
                    <td className="px-4 py-2.5 text-right font-medium text-slate-700">{g.promedio.toFixed(1)}%</td>
                    <td className="px-4 py-2.5 text-center"><Badge tone={c.tone}>{c.texto}</Badge></td>
                    <td className="px-4 py-2.5 text-center">
                      {g.tendencia === "subiendo" && <span className="text-emerald-600 text-xs font-medium">↑ Mejorando</span>}
                      {g.tendencia === "bajando" && <span className="text-rose-600 text-xs font-medium">↓ Empeorando</span>}
                      {g.tendencia === "estable" && <span className="text-slate-400 text-xs">→ Estable</span>}
                      {g.tendencia === null && <span className="text-slate-300 text-xs">—</span>}
                    </td>
                    <td className="px-4 py-2.5 text-right text-slate-500">{g.ultima || "—"}</td>
                    <td className="px-4 py-2.5 text-center">{g.promedio >= 80 && g.cantidad >= 2 && <span title="Buen desempeño consistente" className="text-amber-500">★</span>}</td>
                  </tr>
                ); })}
              </tbody>
            </table>
          </div>
        )}
        </>
      ) : (
      <>
      {/* 1. FILTRO POR PERIODO */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
        <div>
          <label className="text-[11px] font-medium text-slate-500">Proveedor</label>
          <select value={filtro} onChange={(e) => setFiltro(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm min-w-[200px]">
            <option value="">Todos los evaluados</option>
            {opciones.map((o) => <option key={o.key} value={o.key}>{o.label}</option>)}
          </select>
        </div>
        <div><label className="text-[11px] font-medium text-slate-500">Desde</label><input type="date" value={desde} onChange={(e) => setDesde(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" /></div>
        <div><label className="text-[11px] font-medium text-slate-500">Hasta</label><input type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" /></div>
        <button onClick={descargarReporte} disabled={!evaluaciones.length} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40 flex items-center gap-1"><FileText size={13} /> Descargar Excel (resumen)</button>
      </div>

      {evaluaciones.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">No hay evaluaciones completas que coincidan con estos filtros.</div>
      ) : (
        <>
          {/* 3. PROMEDIO GENERAL + GRÁFICO POR CRITERIO */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 flex items-center justify-between flex-wrap gap-3">
            <div>
              <div className="text-xs text-slate-400">Resultado promedio ({evaluaciones.length} evaluación{evaluaciones.length > 1 ? "es" : ""})</div>
              <div className="text-2xl font-semibold text-slate-800">{promedioPct.toFixed(1)}%</div>
            </div>
            <Badge tone={clas.tone}>{clas.texto}</Badge>
          </div>

          <div className="bg-white rounded-xl border border-slate-200 p-5">
            <div className="font-medium text-slate-700 mb-3">Promedio por criterio</div>
            <ResponsiveContainer width="100%" height={420}>
              <BarChart data={promediosPorCriterio.map((c) => ({ nombre: c.texto.length > 42 ? c.texto.slice(0, 40) + "…" : c.texto, textoCompleto: c.texto, valor: c.promedio || 0 }))} layout="vertical" margin={{ left: 10, right: 20 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                <XAxis type="number" domain={[0, 10]} />
                <YAxis type="category" dataKey="nombre" width={260} tick={{ fontSize: 10 }} />
                <Tooltip formatter={(v) => v.toFixed(1)} labelFormatter={(_, p) => p?.[0]?.payload?.textoCompleto || ""} />
                <Bar dataKey="valor" fill="#4f46e5" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* 4. EVALUACIONES YA REALIZADAS */}
          <div className="bg-white rounded-xl border border-slate-200 p-5">
            <div className="font-medium text-slate-700 mb-3">Evaluaciones ya realizadas — clic en una fila para ver el detalle completo</div>
            <table className="w-full text-sm">
              <thead className="text-slate-400 text-xs border-b border-slate-100"><tr><th className="text-left py-1">Consecutivo</th><th className="text-left py-1">Proveedor</th><th className="text-left py-1">Fecha evaluación</th><th className="text-right py-1">Resultado</th><th className="text-right py-1">Clasificación</th><th></th></tr></thead>
              <tbody>{evaluaciones.map((s) => { const pct = puntajeEvaluacion(s.evaluacionProveedor.criterios) * 100; const c = clasificacionConfianza(pct); return (
                <tr key={s.id} onClick={() => onAbrir?.(s.id)} className="border-t border-slate-50 hover:bg-slate-50 cursor-pointer">
                  <td className="py-1.5 font-medium text-slate-700">{s.folio}</td>
                  <td className="py-1.5 text-slate-600">{s.evaluacionProveedor.proveedorNombre}</td>
                  <td className="py-1.5">{s.evaluacionProveedor.fechaCompletado}</td>
                  <td className="py-1.5 text-right">{pct.toFixed(1)}%</td>
                  <td className="py-1.5 text-right"><Badge tone={c.tone}>{c.texto}</Badge></td>
                  <td className="py-1.5 text-right"><button onClick={(e) => { e.stopPropagation(); descargarExcelEvaluacion(s.evaluacionProveedor, s); }} title="Descargar esta evaluación" className="text-slate-400 hover:text-indigo-600 p-1"><FileText size={14} /></button></td>
                </tr>
              ); })}</tbody>
            </table>
          </div>
        </>
      )}
      </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   PLAN DE INVERSIÓN — cronograma editable por proyecto, con
   distribución por periodos (mes/semana), exportable a Excel y PDF.
--------------------------------------------------------- */
const NOMBRES_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
// reconoce el mes de un texto libre como "Abril" o "abril 2026" (sin acentos, sin importar mayúsculas)
function indiceMesDesdeTexto(texto) {
  const limpio = (texto || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  return NOMBRES_MESES.findIndex((m) => limpio.startsWith(m));
}
// interpreta un rango de días como "14-18" → {desde: 14, hasta: 18}; si no tiene guion, usa todo el mes
function parsearRangoDias(rango) {
  const m = (rango || "").match(/(\d{1,2})\s*-\s*(\d{1,2})/);
  if (m) return { desde: parseInt(m[1]), hasta: parseInt(m[2]) };
  return { desde: 1, hasta: 31 };
}
// reparte los pagos del plan de una solicitud entre los periodos del plan de inversión, según la
// fecha de cada pago (anticipo/intermedio/final, o pago único). Si no existe un periodo para el
// mes de un pago, lo crea automáticamente. Devuelve los periodos (con los nuevos agregados) y los
// valores a poner en la fila de ese proyecto.
function distribuirPagosEnPeriodos(solicitud, periodosExistentes) {
  // pagos de cada ítem: su plan confirmado, o el que sugirió el solicitante si aún no se confirma; un ítem
  // sin ningún plan usa la fecha estimada de entrega y su valor total como un solo pago
  const sinIvaPlan = solicitud.tipo === "servicio";
  const tramos = [];
  solicitud.items.forEach((it) => {
    const plan = planConfirmadoItem(solicitud, it) ? planOficialItem(solicitud, it) : planSugeridoItem(solicitud, it);
    const tr = tramosDePago(plan).filter((t) => parseFloat(t.valor) > 0 && t.fecha);
    if (tr.length) tramos.push(...tr);
    else { const totalIt = totalItemConAiu(it, sinIvaPlan); if (totalIt > 0 && solicitud.fechaEstimada) tramos.push({ valor: totalIt, fecha: solicitud.fechaEstimada }); }
  });
  let periodos = [...periodosExistentes];
  const valores = {};
  tramos.forEach((t) => {
    const [anio, mes, dia] = t.fecha.split("-").map(Number);
    const diaNum = dia;
    let periodo = periodos.find((pe) => {
      const idxMes = indiceMesDesdeTexto(pe.mes);
      if (idxMes !== mes - 1) return false;
      const { desde, hasta } = parsearRangoDias(pe.rango);
      return diaNum >= desde && diaNum <= hasta;
    });
    if (!periodo) {
      // no hay ningún periodo para ese mes/rango todavía — se crea uno nuevo automáticamente
      const nombreMes = NOMBRES_MESES[mes - 1];
      periodo = { id: nextId(), mes: nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1), etiqueta: "Semana", rango: `${dia}` };
      periodos.push(periodo);
    }
    valores[periodo.id] = (parseFloat(valores[periodo.id]) || 0) + parseFloat(t.valor);
  });
  return { periodos, valores };
}

// arma una fecha aproximada (el primer día del rango) para un periodo, usada para poder ordenar
// los pagos cronológicamente al reconstruir el plan de pagos de la solicitud
function fechaAproximadaPeriodo(periodo, anio) {
  const idxMes = indiceMesDesdeTexto(periodo.mes);
  if (idxMes < 0) return null;
  const { desde } = parsearRangoDias(periodo.rango);
  const mesStr = String(idxMes + 1).padStart(2, "0");
  const diaStr = String(Math.min(Math.max(desde, 1), 28)).padStart(2, "0"); // 28 para no salirse de ningún mes
  return `${anio}-${mesStr}-${diaStr}`;
}

// operación inversa a distribuirPagosEnPeriodos: toma los valores que quedaron en la fila del plan
// de inversión (después de ajustes manuales) y arma un plan de pagos válido para la solicitud —
// hasta 3 pagos caben como anticipo/intermedio/final; con 1 solo, queda como pago único
function reconstruirPlanDesdePeriodos(proyecto, periodos, anio) {
  const conValor = periodos
    .map((pe) => ({ periodo: pe, valor: parseFloat(proyecto.valores?.[pe.id]) || 0, fecha: fechaAproximadaPeriodo(pe, anio) }))
    .filter((x) => x.valor > 0)
    .sort((a, b) => (a.fecha || "").localeCompare(b.fecha || ""));
  if (!conValor.length) return { error: "No hay ningún valor puesto en las columnas de esta fila." };
  if (conValor.length > 3) return { error: "Esta fila tiene valores en más de 3 periodos — el plan de pagos de una solicitud admite máximo 3 (anticipo, intermedio y final). Consolida los valores en 3 periodos o menos antes de actualizar." };
  if (conValor.some((x) => !x.fecha)) return { error: "Alguno de los periodos con valor no tiene un mes reconocible — revisa el nombre del mes en el encabezado de esa columna." };

  let pagos = planPagosVacio();
  if (conValor.length === 1) {
    pagos = { ...pagos, tipoPago: "contado", pagoUnico: { valor: conValor[0].valor, fecha: conValor[0].fecha } };
  } else if (conValor.length === 2) {
    pagos = { ...pagos, tipoPago: "plan", anticipo: { valor: conValor[0].valor, fecha: conValor[0].fecha }, final: { valor: conValor[1].valor, fecha: conValor[1].fecha } };
  } else {
    pagos = { ...pagos, tipoPago: "plan", anticipo: { valor: conValor[0].valor, fecha: conValor[0].fecha }, intermedio: { activo: true, valor: conValor[1].valor, fecha: conValor[1].fecha }, final: { valor: conValor[2].valor, fecha: conValor[2].fecha } };
  }
  return { pagos };
}

function periodoVacio() { return { id: nextId(), mes: "", etiqueta: "Semana", rango: "" }; }
function proyectoVacio(item) { return { id: nextId(), item, nombre: "", valores: {}, destacado: false, solicitudId: null }; }

function inversionProyecto(proyecto) {
  return Object.values(proyecto.valores || {}).reduce((acc, v) => acc + (parseFloat(v) || 0), 0);
}
function totalPorPeriodo(proyectos, periodoId) {
  return proyectos.reduce((acc, p) => acc + (parseFloat(p.valores?.[periodoId]) || 0), 0);
}
function totalGeneralPlan(proyectos) {
  return proyectos.reduce((acc, p) => acc + inversionProyecto(p), 0);
}

function PlanInversion({ empresas, currentUser, solicitudes, onAbrir, onActualizarSolicitud }) {
  const { datos: planes, cargando, guardar: guardarPlanDB, eliminar: eliminarPlanDB } = useSupabaseTable('planes_inversion', {
    desdeDb: (r) => ({ id: r.id, titulo: r.titulo, empresaId: r.empresa_id, anio: r.anio, periodos: r.datos?.periodos || [], proyectos: r.datos?.proyectos || [] }),
    haciaDb: (r) => ({ id: r.id, titulo: r.titulo, empresa_id: r.empresaId, anio: r.anio, datos: { periodos: r.periodos || [], proyectos: r.proyectos || [] } }),
    orderBy: 'anio',
  });
  const [planActivoId, setPlanActivoId] = useState(null);
  const [exportandoPDF, setExportandoPDF] = useState(false);
  const [seleccionados, setSeleccionados] = useState([]);
  // buffer local: la pantalla se actualiza al instante con cada tecla; el guardado en Supabase
  // ocurre en segundo plano con un pequeño retraso, para no esperar el viaje de ida y vuelta
  const [planLocal, setPlanLocal] = useState(null);
  const guardarTimer = useRef(null);

  useEffect(() => {
    if (!planActivoId && planes.length) setPlanActivoId(planes[0].id);
  }, [planes, planActivoId]);

  // sincroniza el buffer local cuando cambia de plan, o cuando llega la primera carga desde la BD
  useEffect(() => {
    const planDB = planes.find((p) => p.id === planActivoId);
    if (planDB && (!planLocal || planLocal.id !== planDB.id)) { setPlanLocal(planDB); setSeleccionados([]); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [planActivoId, planes]);

  useEffect(() => {
    if (!exportandoPDF) return;
    const t = setTimeout(() => window.print(), 150);
    const limpiar = () => setExportandoPDF(false);
    window.addEventListener("afterprint", limpiar);
    return () => { clearTimeout(t); window.removeEventListener("afterprint", limpiar); };
  }, [exportandoPDF]);

  const plan = planLocal;
  const puedeEditar = ["Administrador", "Gerencia", "Dirección Financiera"].includes(currentUser.rol);

  // actualiza la pantalla al instante, y guarda en Supabase con un pequeño retraso (no en cada
  // tecla) para no saturar la base de datos en una tabla con tantos campos
  const actualizarPlan = (cambios) => {
    const actualizado = { ...plan, ...cambios };
    setPlanLocal(actualizado);
    if (guardarTimer.current) clearTimeout(guardarTimer.current);
    guardarTimer.current = setTimeout(() => guardarPlanDB(actualizado), 600);
  };

  const crearPlan = async () => {
    const nuevo = { id: nextId(), titulo: "Cronograma y Plan de Inversión", empresaId: empresas[0]?.id || null, anio: new Date().getFullYear(), periodos: [], proyectos: [] };
    const creado = await guardarPlanDB(nuevo);
    if (creado?.id) { setPlanActivoId(creado.id); setPlanLocal(creado); }
  };

  const agregarPeriodo = () => actualizarPlan({ periodos: [...(plan.periodos || []), periodoVacio()] });
  const quitarPeriodo = (id) => {
    const nuevosProyectos = plan.proyectos.map((p) => { const v = { ...p.valores }; delete v[id]; return { ...p, valores: v }; });
    actualizarPlan({ periodos: plan.periodos.filter((pe) => pe.id !== id), proyectos: nuevosProyectos });
  };
  const editarPeriodo = (id, campo, val) => actualizarPlan({ periodos: plan.periodos.map((pe) => (pe.id === id ? { ...pe, [campo]: val } : pe)) });

  const agregarProyecto = () => actualizarPlan({ proyectos: [...(plan.proyectos || []), proyectoVacio((plan.proyectos?.length || 0) + 1)] });
  const quitarProyecto = (id) => actualizarPlan({ proyectos: plan.proyectos.filter((p) => p.id !== id) });
  const alternarSeleccion = (id) => setSeleccionados((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const todosSeleccionados = plan?.proyectos.length > 0 && seleccionados.length === plan.proyectos.length;
  const alternarTodos = () => setSeleccionados(todosSeleccionados ? [] : plan.proyectos.map((p) => p.id));
  const borrarSeleccionados = () => {
    if (!seleccionados.length) return;
    if (!confirm(`¿Borrar ${seleccionados.length} proyecto(s) del plan? Esto no se puede deshacer.`)) return;
    actualizarPlan({ proyectos: plan.proyectos.filter((p) => !seleccionados.includes(p.id)) });
    setSeleccionados([]);
  };
  const editarProyecto = (id, campo, val) => actualizarPlan({ proyectos: plan.proyectos.map((p) => (p.id === id ? { ...p, [campo]: val } : p)) });
  const editarValor = (proyectoId, periodoId, val) => actualizarPlan({ proyectos: plan.proyectos.map((p) => (p.id === proyectoId ? { ...p, valores: { ...p.valores, [periodoId]: val } } : p)) });

  // vincula la fila con una solicitud real: trae el objetivo, y reparte cada pago de su plan de
  // pagos en el periodo (mes) que le corresponda según la fecha — creando periodos nuevos si hace falta
  const vincularSolicitud = (proyectoId, solicitudId) => {
    if (!solicitudId) { actualizarPlan({ proyectos: plan.proyectos.map((p) => (p.id === proyectoId ? { ...p, solicitudId: null } : p)) }); return; }
    const solicitud = solicitudes.find((s) => s.id === solicitudId);
    if (!solicitud) return;
    const { periodos: periodosActualizados, valores } = distribuirPagosEnPeriodos(solicitud, plan.periodos);
    actualizarPlan({
      periodos: periodosActualizados,
      proyectos: plan.proyectos.map((p) => (p.id === proyectoId ? { ...p, solicitudId, nombre: `${solicitud.folio} — ${solicitud.objetivo}`, valores } : p)),
    });
  };

  // arma automáticamente una fila por cada solicitud real de la empresa/año del plan que todavía
  // no esté vinculada — no duplica las que ya existen, y no toca las filas manuales que haya
  const generarDesdeSolicitudes = () => {
    const yaVinculadas = new Set(plan.proyectos.map((p) => p.solicitudId).filter(Boolean));
    const anioTieneAlgunPago = (s) => {
      const fechas = s.items.flatMap((it) => tramosDePago(planConfirmadoItem(s, it) ? planOficialItem(s, it) : planSugeridoItem(s, it)).filter((t) => parseFloat(t.valor) > 0 && t.fecha).map((t) => t.fecha));
      if (fechas.length) return fechas.some((f) => f.startsWith(String(plan.anio)));
      return (s.fechaEstimada || s.fechaCreacion || "").startsWith(String(plan.anio));
    };
    const candidatas = solicitudes.filter((s) =>
      !["rechazada"].includes(s.status) &&
      (!plan.empresaId || s.empresaId === plan.empresaId) &&
      anioTieneAlgunPago(s) &&
      !yaVinculadas.has(s.id)
    );
    if (!candidatas.length) { alert("No hay solicitudes nuevas por agregar (de esta empresa y año) que no estén ya en el plan."); return; }
    let periodos = [...plan.periodos];
    let siguienteItem = (plan.proyectos.reduce((max, p) => Math.max(max, parseFloat(p.item) || 0), 0)) + 1;
    const nuevosProyectos = candidatas.map((s) => {
      const r = distribuirPagosEnPeriodos(s, periodos);
      periodos = r.periodos;
      return { id: nextId(), item: siguienteItem++, nombre: `${s.folio} — ${s.objetivo}`, valores: r.valores, destacado: false, solicitudId: s.id };
    });
    actualizarPlan({ periodos, proyectos: [...plan.proyectos, ...nuevosProyectos] });
  };

  // empuja los valores (ya ajustados a mano) de una fila vinculada de vuelta al plan de pagos
  // real de la solicitud — deja el plan como "confirmado" ya que viene de un ajuste deliberado
  const [actualizando, setActualizando] = useState(null);
  const actualizarPlanDePagos = async (proyecto) => {
    const solicitud = solicitudes.find((s) => s.id === proyecto.solicitudId);
    if (!solicitud) return;
    if (solicitud.items.length > 1) { alert("Esta solicitud tiene varios ítems y cada uno maneja su propio plan de pagos. Ajusta el plan de cada ítem dentro de la solicitud (enlace «Ver solicitud vinculada»)."); return; }
    const { pagos, error } = reconstruirPlanDesdePeriodos(proyecto, plan.periodos, plan.anio);
    if (error) { alert(error); return; }
    if (!confirm(`Esto va a reemplazar el plan de pagos de la solicitud ${solicitud.folio} con los valores puestos en esta fila. ¿Confirmas?`)) return;
    setActualizando(proyecto.id);
    // el plan de pagos ahora vive en cada ítem: aquí solo se puede empujar de vuelta cuando la solicitud tiene un único ítem
    await onActualizarSolicitud({ ...solicitud, items: solicitud.items.map((it) => ({ ...it, pagos, pagosConfirmados: true, pagosConfirmadosPor: { nombre: currentUser.nombre, rol: currentUser.rol } })) });
    setActualizando(null);
    alert(`Plan de pagos de ${solicitud.folio} actualizado.`);
  };

  const descargarExcel = () => {
    if (!plan) return;
    const filas = [];
    filas.push([plan.titulo?.toUpperCase() || ""]);
    filas.push([]);
    filas.push(["ITEM", "PROYECTO", "INVERSIÓN", ...plan.periodos.map((pe) => pe.mes)]);
    filas.push(["", "", "", ...plan.periodos.map((pe) => pe.etiqueta)]);
    filas.push(["", "", "", ...plan.periodos.map((pe) => pe.rango)]);
    plan.proyectos.forEach((p) => {
      filas.push([p.item, p.nombre, inversionProyecto(p), ...plan.periodos.map((pe) => parseFloat(p.valores?.[pe.id]) || "")]);
    });
    filas.push(["", "TOTAL INVERSIÓN", totalGeneralPlan(plan.proyectos), ...plan.periodos.map((pe) => totalPorPeriodo(plan.proyectos, pe.id))]);
    const hoja = XLSX.utils.aoa_to_sheet(filas);
    hoja["!cols"] = [{ wch: 6 }, { wch: 45 }, { wch: 16 }, ...plan.periodos.map(() => ({ wch: 14 }))];
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Plan de inversión");
    XLSX.writeFile(libro, `${(plan.titulo || "Plan_inversion").replace(/[^a-zA-Z0-9]/g, "_")}_${plan.anio}.xlsx`);
  };

  if (cargando) return <div className="text-sm text-slate-400">Cargando...</div>;

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Plan de inversión</h2>
          <p className="text-xs text-slate-400 mt-1">Cronograma editable por proyecto, con distribución en el tiempo — exportable a Excel y PDF.</p>
        </div>
        <div className="flex items-center gap-2 no-print">
          {planes.length > 0 && (
            <select value={planActivoId || ""} onChange={(e) => setPlanActivoId(e.target.value)} className="border border-slate-200 rounded-lg px-2 py-1.5 text-sm">
              {planes.map((p) => <option key={p.id} value={p.id}>{p.titulo} ({p.anio})</option>)}
            </select>
          )}
          {puedeEditar && <button onClick={crearPlan} className="text-xs bg-white border border-slate-200 text-slate-600 px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><Plus size={13} /> Nuevo plan</button>}
          {plan && puedeEditar && <button onClick={generarDesdeSolicitudes} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><TrendingUp size={13} /> Generar desde solicitudes</button>}
          {plan && puedeEditar && seleccionados.length > 0 && <button onClick={borrarSeleccionados} className="text-xs bg-rose-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><Trash2 size={13} /> Borrar seleccionadas ({seleccionados.length})</button>}
          {plan && <button onClick={descargarExcel} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><FileText size={13} /> Excel</button>}
          {plan && <button onClick={() => setExportandoPDF(true)} className="text-xs bg-slate-800 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><FileText size={13} /> PDF</button>}
        </div>
      </div>

      {!plan ? (
        <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">
          {puedeEditar ? <>No hay ningún plan de inversión todavía. <button onClick={crearPlan} className="text-indigo-600 underline">Crear el primero</button></> : "No hay ningún plan de inversión todavía."}
        </div>
      ) : (
        <>
          {puedeEditar && (
            <div className="bg-white rounded-xl border border-slate-200 p-4 grid grid-cols-1 md:grid-cols-3 gap-3 no-print">
              <div><label className="text-[11px] font-medium text-slate-500">Título</label><input value={plan.titulo} onChange={(e) => actualizarPlan({ titulo: e.target.value })} className="w-full mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" /></div>
              <div><label className="text-[11px] font-medium text-slate-500">Empresa</label><select value={plan.empresaId || ""} onChange={(e) => actualizarPlan({ empresaId: e.target.value })} className="w-full mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm"><option value="">—</option>{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></div>
              <div><label className="text-[11px] font-medium text-slate-500">Año</label><input type="number" value={plan.anio} onChange={(e) => actualizarPlan({ anio: e.target.value })} className="w-full mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" /></div>
            </div>
          )}

          <style>{`
            @media print {
              .print-wrapper-oculto { display: block !important; }
              body * { visibility: hidden; }
              #plan-imprimible, #plan-imprimible * { visibility: visible; }
              #plan-imprimible { position: absolute; left: 0; top: 0; width: 100%; border: none !important; }
              #plan-imprimible .no-print { display: none !important; }
              @page { size: landscape; margin: 10mm; }
            }
          `}</style>
          <div id="plan-imprimible" className="bg-white rounded-xl border border-slate-200 p-4 overflow-x-auto">
            <div className="text-center font-semibold text-slate-800 mb-3 text-sm">{plan.titulo} {plan.anio} {empresas.find((e) => e.id === plan.empresaId)?.nombre || ""}</div>
            <table className="w-full text-xs border-collapse min-w-[900px]">
              <thead>
                <tr className="bg-slate-50">
                  {puedeEditar && <th rowSpan={3} className="no-print border border-slate-200 px-1 py-1.5 w-6"><input type="checkbox" checked={todosSeleccionados} onChange={alternarTodos} /></th>}
                  <th rowSpan={3} className="border border-slate-200 px-2 py-1.5 w-10">ITEM</th>
                  <th rowSpan={3} className="border border-slate-200 px-2 py-1.5 min-w-[340px]">PROYECTO</th>
                  <th rowSpan={3} className="border border-slate-200 px-2 py-1.5 w-28">INVERSIÓN</th>
                  {plan.periodos.map((pe) => (
                    <th key={pe.id} className="border border-slate-200 px-1 py-1 min-w-[72px] max-w-[80px] relative group">
                      {puedeEditar ? <input value={pe.mes} onChange={(e) => editarPeriodo(pe.id, "mes", e.target.value)} placeholder="Mes" className="w-full text-center text-xs font-semibold border-0 bg-transparent focus:bg-white" /> : pe.mes}
                      {puedeEditar && <button onClick={() => quitarPeriodo(pe.id)} className="no-print absolute -top-1 -right-1 text-rose-400 hover:text-rose-600 bg-white rounded-full opacity-0 group-hover:opacity-100"><Trash2 size={11} /></button>}
                    </th>
                  ))}
                  {puedeEditar && <th rowSpan={3} className="no-print border border-slate-200 px-2 py-1.5 w-10"><button onClick={agregarPeriodo} title="Agregar periodo" className="text-indigo-600"><Plus size={14} /></button></th>}
                </tr>
                <tr className="bg-slate-50">
                  {plan.periodos.map((pe) => (
                    <th key={pe.id} className="border border-slate-200 px-1 py-1">
                      {puedeEditar ? <input value={pe.etiqueta} onChange={(e) => editarPeriodo(pe.id, "etiqueta", e.target.value)} className="w-full text-center text-[11px] border-0 bg-transparent focus:bg-white" /> : pe.etiqueta}
                    </th>
                  ))}
                </tr>
                <tr className="bg-slate-50">
                  {plan.periodos.map((pe) => (
                    <th key={pe.id} className="border border-slate-200 px-1 py-1">
                      {puedeEditar ? <input value={pe.rango} onChange={(e) => editarPeriodo(pe.id, "rango", e.target.value)} placeholder="Rango" className="w-full text-center text-[11px] font-normal border-0 bg-transparent focus:bg-white" /> : pe.rango}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {plan.proyectos.map((p) => (
                  <tr key={p.id} className={`${p.destacado ? "bg-emerald-100/70" : ""} ${seleccionados.includes(p.id) ? "outline outline-2 outline-indigo-300" : ""}`}>
                    {puedeEditar && <td className="no-print border border-slate-200 px-1 py-1.5 text-center align-top"><input type="checkbox" checked={seleccionados.includes(p.id)} onChange={() => alternarSeleccion(p.id)} /></td>}
                    <td className="border border-slate-200 px-2 py-1.5 text-center align-top">
                      <div className="flex items-center gap-1 justify-center">
                        {p.item}
                        {puedeEditar && <button onClick={() => quitarProyecto(p.id)} className="no-print text-slate-300 hover:text-rose-500"><Trash2 size={11} /></button>}
                      </div>
                      {puedeEditar && <label className="no-print flex items-center justify-center gap-1 mt-1 text-[9px] text-slate-400"><input type="checkbox" checked={!!p.destacado} onChange={(e) => editarProyecto(p.id, "destacado", e.target.checked)} /> resaltar</label>}
                    </td>
                    <td className="border border-slate-200 px-2 py-1.5 align-top">
                      {puedeEditar && (
                        <select value={p.solicitudId || ""} onChange={(e) => vincularSolicitud(p.id, e.target.value)} className="no-print w-full text-[10px] border border-slate-200 rounded px-1 py-0.5 mb-1 text-slate-500">
                          <option value="">— Escribir manual, o vincular con una solicitud —</option>
                          {solicitudes.filter((s) => !["rechazada"].includes(s.status)).map((s) => <option key={s.id} value={s.id}>{s.folio} — {s.objetivo?.slice(0, 40)}</option>)}
                        </select>
                      )}
                      {p.solicitudId && onAbrir && <button onClick={() => onAbrir(p.solicitudId)} className="no-print text-[10px] text-indigo-600 underline mb-1 block">Ver solicitud vinculada →</button>}
                      {p.solicitudId && <div className="no-print text-[9px] text-slate-400 mb-1">Los valores por periodo se trajeron del plan de pagos de esa solicitud — puedes ajustarlos aquí y luego actualizarla.</div>}
                      {p.solicitudId && puedeEditar && solicitudes.find((x) => x.id === p.solicitudId)?.items.length === 1 && (
                        <button onClick={() => actualizarPlanDePagos(p)} disabled={actualizando === p.id} className="no-print text-[10px] bg-amber-500 text-white px-2 py-1 rounded font-medium mb-1 flex items-center gap-1 disabled:opacity-50">
                          <CalendarClock size={11} /> {actualizando === p.id ? "Actualizando..." : "Actualizar plan de pagos"}
                        </button>
                      )}
                      {puedeEditar ? <textarea value={p.nombre} onChange={(e) => editarProyecto(p.id, "nombre", e.target.value)} rows={p.nombre?.length > 80 ? 4 : 1} className="w-full text-xs border-0 bg-transparent resize-y focus:bg-white" /> : <span className="whitespace-pre-wrap">{p.nombre}</span>}
                    </td>
                    <td className="border border-slate-200 px-2 py-1.5 text-right align-top font-medium">{fmt(inversionProyecto(p))}</td>
                    {plan.periodos.map((pe) => (
                      <td key={pe.id} className="border border-slate-200 px-1 py-1.5 text-right align-top">
                        {puedeEditar ? <InputMiles value={p.valores?.[pe.id] || ""} onChange={(v) => editarValor(p.id, pe.id, v)} className="w-full text-right text-xs border-0 bg-transparent focus:bg-white px-1" /> : (parseFloat(p.valores?.[pe.id]) > 0 ? fmt(p.valores[pe.id]) : "")}
                      </td>
                    ))}
                    {puedeEditar && <td className="no-print border border-slate-200"></td>}
                  </tr>
                ))}
                <tr className="bg-slate-100 font-semibold">
                  {puedeEditar && <td className="no-print border border-slate-200"></td>}
                  <td className="border border-slate-200 px-2 py-1.5 text-center">{plan.proyectos.length}</td>
                  <td className="border border-slate-200 px-2 py-1.5">TOTAL INVERSIÓN</td>
                  <td className="border border-slate-200 px-2 py-1.5 text-right">{fmt(totalGeneralPlan(plan.proyectos))}</td>
                  {plan.periodos.map((pe) => (
                    <td key={pe.id} className="border border-slate-200 px-2 py-1.5 text-right">{fmt(totalPorPeriodo(plan.proyectos, pe.id))}</td>
                  ))}
                  {puedeEditar && <td className="no-print border border-slate-200"></td>}
                </tr>
              </tbody>
            </table>
            {puedeEditar && <button onClick={agregarProyecto} className="no-print mt-3 text-xs text-indigo-600 font-medium flex items-center gap-1"><Plus size={13} /> Agregar proyecto</button>}
          </div>
        </>
      )}
    </div>
  );
}


function Estadisticas({ solicitudes, areas, empresas, proveedores }) {
  const [filtroEmpresa, setFiltroEmpresa] = useState("todas");
  const [fDesde, setFDesde] = useState("");
  const [fHasta, setFHasta] = useState("");
  const base = solicitudes.filter((s) =>
    (filtroEmpresa === "todas" || s.empresaId === filtroEmpresa) &&
    (!fDesde || s.fechaCreacion >= fDesde) &&
    (!fHasta || s.fechaCreacion <= fHasta)
  );

  const porArea = areas.map((a) => ({ nombre: a.nombre, monto: base.filter((s) => s.areaId === a.id && !["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)).reduce((acc, s) => acc + totalSolicitud(s), 0) }));
  const porTipo = [{ nombre: "Compra", value: base.filter((s) => s.tipo === "compra").length }, { nombre: "Servicio", value: base.filter((s) => s.tipo === "servicio").length }];
  const porEstado = [
    ...PASOS.filter((p) => p.key !== "recepcion").map((p) => ({ nombre: p.label, value: base.filter((s) => s.status === p.key).length })),
    { nombre: "Recepción / Ejecución (pendiente)", value: base.filter((s) => s.status === "recepcion" && !s.recepcion?.recibidoSatisfaccion).length },
    { nombre: "Recibida (falta evaluación)", value: base.filter((s) => s.status === "recepcion" && s.recepcion?.recibidoSatisfaccion).length },
  ].filter((e) => e.value > 0);
  const porEmpresaComparativo = empresas.map((e) => ({ nombre: e.nombre, monto: solicitudes.filter((s) => s.empresaId === e.id && !["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)).reduce((acc, s) => acc + totalSolicitud(s), 0) }));
  const proveedorMonto = {};
  base.forEach((s) => s.items.forEach((it) => {
    if (!it.cotizaciones.length) return;
    const idx = it.cotizacionSeleccionada ?? mejorCotizacionIdx(it.cotizaciones, it.cantidad);
    const cot = it.cotizaciones[idx]; if (!cot) return;
    const prov = proveedores.find((p) => p.id === cot.proveedorId)?.nombre || cot.proveedorNombre || "—";
    proveedorMonto[prov] = (proveedorMonto[prov] || 0) + desgloseCotizacion(cot, it.cantidad).total;
  }));
  const porProveedor = Object.entries(proveedorMonto).map(([nombre, monto]) => ({ nombre, monto })).sort((a, b) => b.monto - a.monto);
  const totalGeneral = base.reduce((acc, s) => acc + totalSolicitud(s), 0);
  const ivaGeneral = base.reduce((acc, s) => acc + desgloseSolicitud(s).iva, 0);

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div className="flex items-end gap-3">
          <div>
            <div className="text-[11px] font-medium text-slate-500 mb-1">Filtrar por empresa</div>
            <select value={filtroEmpresa} onChange={(e) => setFiltroEmpresa(e.target.value)} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm">
              <option value="todas">Todas las empresas</option>{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}
            </select>
          </div>
          <div>
            <div className="text-[11px] font-medium text-slate-500 mb-1">Desde</div>
            <input type="date" value={fDesde} onChange={(e) => setFDesde(e.target.value)} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm" />
          </div>
          <div>
            <div className="text-[11px] font-medium text-slate-500 mb-1">Hasta</div>
            <input type="date" value={fHasta} onChange={(e) => setFHasta(e.target.value)} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm" />
          </div>
          {(fDesde || fHasta) && <button onClick={() => { setFDesde(""); setFHasta(""); }} className="text-xs text-slate-500 underline mb-1.5">Limpiar fechas</button>}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="text-xs text-slate-500 mb-1">Total solicitudes</div><div className="text-2xl font-semibold text-slate-800">{base.length}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="text-xs text-slate-500 mb-1">Completadas</div><div className="text-2xl font-semibold text-emerald-600">{base.filter((s) => s.status === "completada").length}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="text-xs text-slate-500 mb-1">Valor total gestionado</div><div className="text-xl font-semibold text-slate-800">{fmt(totalGeneral)}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="text-xs text-slate-500 mb-1">IVA total</div><div className="text-xl font-semibold text-slate-800">{fmt(ivaGeneral)}</div></div>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="font-medium text-slate-700 mb-3 text-sm">Monto comprometido por área</div>
          <ResponsiveContainer width="100%" height={240}><BarChart data={porArea}><CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" /><XAxis dataKey="nombre" tick={{ fontSize: 11 }} /><YAxis tick={{ fontSize: 10 }} tickFormatter={(v) => `${(v / 1e6).toFixed(0)}M`} /><Tooltip formatter={(v) => fmt(v)} /><Bar dataKey="monto" fill="#4f46e5" radius={[4, 4, 0, 0]} /></BarChart></ResponsiveContainer>
        </div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="font-medium text-slate-700