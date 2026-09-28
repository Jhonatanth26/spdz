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
// Ítems plegables: con varios ítems las pantallas se vuelven larguísimas. Cada ítem tiene una flecha para expandir o
// contraer y, contraído, deja una línea de resumen. El contenido SIGUE montado (solo se oculta), así que nada de lo
// que se haya escrito se pierde al contraer.
function useAbiertosItems(items, abiertoInicial) {
  const inicial = abiertoInicial ?? (items || []).length <= 2; // con 1 o 2 ítems arrancan abiertos; con más, contraídos
  const [mapa, setMapa] = useState({});
  return {
    abierto: (id) => mapa[id] ?? inicial,
    alternar: (id) => setMapa((m) => ({ ...m, [id]: !(m[id] ?? inicial) })),
    abrir: (id) => setMapa((m) => ({ ...m, [id]: true })),
    todos: (v) => setMapa(Object.fromEntries((items || []).map((it) => [it.id, v]))),
  };
}
function ItemColapsable({ abierto, onToggle, numero, titulo, resumen, insignia, tinte, children }) {
  const t = tinte || TINTES_ITEM[0];
  return (
    <div className={`border ${t.borde} ${t.fondo} rounded-lg`}>
      <button type="button" onClick={onToggle} aria-expanded={abierto} className="w-full flex items-center gap-2 px-3 py-2 text-left">
        <ChevronRight size={16} className={`shrink-0 text-slate-400 transition-transform ${abierto ? "rotate-90" : ""}`} />
        <span className={`inline-block w-2 h-2 rounded-full shrink-0 ${t.punto}`} />
        <span className="text-sm font-medium text-slate-700 truncate">{numero}. {titulo || "Ítem sin nombre"}</span>
        <span className="ml-auto flex items-center gap-2 text-xs text-slate-500 shrink-0">{resumen}{insignia}</span>
      </button>
      <div className={abierto ? "px-2 pb-2" : "hidden"}>{children}</div>
    </div>
  );
}
function ControlExpandirTodo({ n, onTodos }) {
  if (n <= 2) return null;
  return (
    <div className="flex items-center gap-2 text-[11px] text-slate-500">
      <button type="button" onClick={() => onTodos(true)} className="hover:text-indigo-600 underline">Expandir todo</button>
      <span>·</span>
      <button type="button" onClick={() => onTodos(false)} className="hover:text-indigo-600 underline">Contraer todo</button>
    </div>
  );
}
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
// Cotización que se usa para dar valor a un ítem: la que Compras eligió y, mientras no haya elegido, la RECOMENDADA
// por el cuadro comparativo (misma cuenta que se ve en pantalla: en servicios incluye el AIU de cada proveedor).
// Es la ÚNICA regla: resumen, aprobaciones, pagos, órdenes y reportes la leen de aquí, para que la misma solicitud
// nunca tenga dos cifras distintas.
function idxCotizacionActiva(item, sinIva) {
  if (!item.cotizaciones?.length) return -1;
  return item.cotizacionSeleccionada ?? mejorCotizacionIdx(item.cotizaciones, item.cantidad, sinIva, item.aiu);
}
function desgloseItem(item, sinIva) {
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
  const idx = idxCotizacionActiva(item, sinIva);
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
    const dItem = desgloseItem(item, true);
    costoDirecto += dItem.subtotal;
    let aiu = item.aiu || solicitud.aiu || {};
    if (item.cotizaciones?.length) {
      const sel = idxCotizacionActiva(item, true);
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
  const d = desgloseItem(item, sinIva);
  if (!sinIva) return d.total;
  let aiu = item.aiu || {};
  if (item.cotizaciones?.length) {
    const sel = idxCotizacionActiva(item, true);
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
    const idx = idxCotizacionActiva(it, s.tipo === "servicio");
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
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="font-medium text-slate-700 mb-3 text-sm">Solicitudes por tipo</div>
          <ResponsiveContainer width="100%" height={240}><PieChart><Pie data={porTipo} dataKey="value" nameKey="nombre" cx="50%" cy="50%" outerRadius={80} label>{porTipo.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}</Pie><Tooltip /><Legend /></PieChart></ResponsiveContainer>
        </div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="font-medium text-slate-700 mb-3 text-sm">Solicitudes por estado del flujo</div>
          <ResponsiveContainer width="100%" height={240}><BarChart data={porEstado} layout="vertical" margin={{ left: 40 }}><CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" /><XAxis type="number" tick={{ fontSize: 10 }} allowDecimals={false} /><YAxis dataKey="nombre" type="category" tick={{ fontSize: 10 }} width={140} /><Tooltip /><Bar dataKey="value" fill="#0ea5e9" radius={[0, 4, 4, 0]} /></BarChart></ResponsiveContainer>
        </div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="font-medium text-slate-700 mb-3 text-sm">Comparativo entre empresas (siempre totales)</div>
          <ResponsiveContainer width="100%" height={240}><PieChart><Pie data={porEmpresaComparativo} dataKey="monto" nameKey="nombre" cx="50%" cy="50%" outerRadius={80} label={(e) => e.nombre}>{porEmpresaComparativo.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}</Pie><Tooltip formatter={(v) => fmt(v)} /></PieChart></ResponsiveContainer>
        </div>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="px-5 py-3 border-b border-slate-100 font-medium text-slate-700 text-sm">Monto adjudicado por proveedor</div>
        <table className="w-full text-sm"><thead className="bg-slate-50 text-slate-500 text-xs"><tr><th className="text-left px-5 py-2">Proveedor</th><th className="text-right px-5 py-2">Monto adjudicado</th></tr></thead>
          <tbody>{porProveedor.length ? porProveedor.map((p) => (<tr key={p.nombre} className="border-t border-slate-100"><td className="px-5 py-2 text-slate-700">{p.nombre}</td><td className="px-5 py-2 text-right text-slate-600">{fmt(p.monto)}</td></tr>)) : <tr><td colSpan={2} className="px-5 py-4 text-slate-400 text-center text-xs">Aún no hay cotizaciones seleccionadas</td></tr>}</tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   HISTÓRICO DE COMPRAS (solo visible para el rol Compras)
--------------------------------------------------------- */
function HistoricoCompras({ nombreItem, historico, setHistorico }) {
  const [mostrarForm, setMostrarForm] = useState(false);
  const [form, setForm] = useState({ fecha: "", proveedor: "", precioUnitario: "", cantidad: "", unidad: "unidad" });
  const registros = historico.filter((h) => h.itemNombre.toLowerCase() === (nombreItem || "").toLowerCase());
  const agregar = () => {
    if (!form.fecha || !form.proveedor || !form.precioUnitario) return;
    setHistorico([...historico, { id: nextId(), itemNombre: nombreItem, ...form }]);
    setForm({ fecha: "", proveedor: "", precioUnitario: "", cantidad: "", unidad: "unidad" }); setMostrarForm(false);
  };
  if (!nombreItem) return null;
  return (
    <div className="bg-slate-50 rounded-lg p-2.5 border border-slate-200 mt-1.5">
      <div className="flex items-center justify-between mb-1">
        <div className="text-[11px] font-medium text-slate-500 flex items-center gap-1"><History size={12} /> Histórico de compras — {nombreItem}</div>
        {!mostrarForm && <button onClick={() => setMostrarForm(true)} className="text-[11px] text-indigo-600 font-medium flex items-center gap-1"><Plus size={11} /> Agregar registro</button>}
      </div>
      {registros.length > 0 ? (
        <table className="w-full text-[11px]">
          <thead className="text-slate-400"><tr><th className="text-left py-0.5">Fecha</th><th className="text-left py-0.5">Proveedor</th><th className="text-right py-0.5">Precio unit.</th><th className="text-right py-0.5">Cant.</th></tr></thead>
          <tbody>{registros.map((h) => (<tr key={h.id} className="border-t border-slate-200"><td className="py-1">{h.fecha}</td><td className="py-1">{h.proveedor}</td><td className="py-1 text-right">{fmt(h.precioUnitario)}</td><td className="py-1 text-right">{h.cantidad} {h.unidad}</td></tr>))}</tbody>
        </table>
      ) : !mostrarForm && <div className="text-[11px] text-slate-400">Sin histórico registrado en el sistema para este ítem.</div>}
      {mostrarForm && (
        <div className="grid grid-cols-5 gap-1 mt-1.5">
          <input type="date" value={form.fecha} onChange={(e) => setForm({ ...form, fecha: e.target.value })} className="border border-slate-200 rounded-md px-1.5 py-1 text-[11px]" />
          <input placeholder="Proveedor" value={form.proveedor} onChange={(e) => setForm({ ...form, proveedor: e.target.value })} className="border border-slate-200 rounded-md px-1.5 py-1 text-[11px]" />
          <input type="number" placeholder="Precio unit." value={form.precioUnitario} onChange={(e) => setForm({ ...form, precioUnitario: e.target.value })} className="border border-slate-200 rounded-md px-1.5 py-1 text-[11px]" />
          <input type="number" placeholder="Cant." value={form.cantidad} onChange={(e) => setForm({ ...form, cantidad: e.target.value })} className="border border-slate-200 rounded-md px-1.5 py-1 text-[11px]" />
          <div className="flex gap-1"><button onClick={agregar} className="bg-indigo-600 text-white text-[11px] px-2 rounded-md flex-1">Guardar</button><button onClick={() => setMostrarForm(false)} className="text-slate-400 text-[11px]">×</button></div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   NUEVA SOLICITUD
--------------------------------------------------------- */
// campo de búsqueda con autocompletado para elegir un ítem del catálogo (o dejarlo libre si no coincide con nada)
function AutocompletarItem({ itemsCatalogo, valorTexto, catalogoId, onElegir, onEscribir }) {
  const [abierto, setAbierto] = useState(false);
  const [indiceActivo, setIndiceActivo] = useState(0);
  const contenedorRef = useRef(null);
  const listaRef = useRef(null);

  useEffect(() => {
    const cerrarSiClicFuera = (e) => { if (contenedorRef.current && !contenedorRef.current.contains(e.target)) setAbierto(false); };
    document.addEventListener("mousedown", cerrarSiClicFuera);
    return () => document.removeEventListener("mousedown", cerrarSiClicFuera);
  }, []);

  const coincidencias = valorTexto.trim()
    ? itemsCatalogo.filter((c) => c.nombre.toLowerCase().includes(valorTexto.trim().toLowerCase())).slice(0, 30)
    : itemsCatalogo.slice(0, 30);

  useEffect(() => { setIndiceActivo(0); }, [valorTexto, abierto]);
  useEffect(() => {
    if (abierto && listaRef.current) {
      const activo = listaRef.current.children[indiceActivo];
      if (activo) activo.scrollIntoView({ block: "nearest" });
    }
  }, [indiceActivo, abierto]);

  const manejarTeclas = (e) => {
    if (!abierto || !coincidencias.length) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setIndiceActivo((i) => Math.min(i + 1, coincidencias.length - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setIndiceActivo((i) => Math.max(i - 1, 0)); }
    else if (e.key === "Enter") { e.preventDefault(); const elegido = coincidencias[indiceActivo]; if (elegido) { onElegir(elegido); setAbierto(false); } }
    else if (e.key === "Escape") setAbierto(false);
  };

  return (
    <div className="relative" ref={contenedorRef}>
      <input
        value={valorTexto}
        onChange={(e) => { onEscribir(e.target.value); setAbierto(true); }}
        onFocus={() => setAbierto(true)}
        onKeyDown={manejarTeclas}
        placeholder="Descripción del ítem — escribe para buscar en el catálogo"
        className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm"
      />
      {abierto && coincidencias.length > 0 && (
        <div ref={listaRef} className="absolute z-20 mt-1 w-full max-h-56 overflow-y-auto bg-white border border-slate-200 rounded-md shadow-lg">
          {coincidencias.map((c, idx) => (
            <button
              key={c.id}
              type="button"
              onMouseEnter={() => setIndiceActivo(idx)}
              onClick={() => { onElegir(c); setAbierto(false); }}
              className={`w-full text-left px-3 py-1.5 text-xs ${idx === indiceActivo ? "bg-indigo-50 text-indigo-700" : catalogoId === c.id ? "bg-indigo-50 text-indigo-700 font-medium" : "text-slate-700"}`}
            >
              {c.nombre} <span className="text-slate-400">({c.unidadDefault})</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// campo de búsqueda con autocompletado para elegir un proveedor del catálogo (o dejarlo libre si no coincide con nada)
// input numérico que muestra separador de miles mientras se escribe (ej. 150.000),
// pero por dentro sigue guardando solo el número plano para los cálculos.
// campo de fecha que no permite elegir (ni escribir) una fecha anterior a hoy
function InputFecha({ value, onChange, disabled, className }) {
  const manejarCambio = (e) => {
    const val = e.target.value;
    if (val && val < hoy()) { alert("No se puede seleccionar una fecha anterior a hoy."); return; }
    onChange(val);
  };
  return <input type="date" min={hoy()} disabled={disabled} value={value} onChange={manejarCambio} className={className} />;
}

function InputMiles({ value, onChange, className, placeholder, disabled }) {
  const formatear = (v) => (v || v === 0) && v !== "" ? Number(v).toLocaleString("es-CO") : "";
  const [texto, setTexto] = useState(formatear(value));
  useEffect(() => { setTexto(formatear(value)); }, [value]); // eslint-disable-line

  const manejarCambio = (e) => {
    const crudo = e.target.value.replace(/[^\d]/g, "");
    setTexto(crudo ? Number(crudo).toLocaleString("es-CO") : "");
    onChange(crudo);
  };

  return <input type="text" inputMode="numeric" value={texto} onChange={manejarCambio} placeholder={placeholder} className={className} disabled={disabled} />;
}

// Lista desplegable con búsqueda: se escribe y la lista se va filtrando (sin importar tildes ni mayúsculas;
// cada palabra escrita debe aparecer en la opción, en cualquier orden). Pensada para catálogos largos,
// como el plan de cuentas (Grupo – Código – Cuenta – Centro de costo).
const normalizarBusqueda = (t) => (t || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
function SelectorBuscable({ opciones, value, onChange, placeholder, maxVisibles = 80, disabled }) {
  const [abierto, setAbierto] = useState(false);
  const [texto, setTexto] = useState("");
  const [activo, setActivo] = useState(0);
  const contRef = useRef(null);
  const listaRef = useRef(null);
  const seleccionada = opciones.find((o) => o.value === value);

  const filtradas = useMemo(() => {
    const palabras = normalizarBusqueda(texto).split(/\s+/).filter(Boolean);
    if (!palabras.length) return opciones;
    return opciones.filter((o) => { const n = normalizarBusqueda(o.label); return palabras.every((p) => n.includes(p)); });
  }, [opciones, texto]);
  const visibles = filtradas.slice(0, maxVisibles);

  useEffect(() => {
    const cerrarSiFuera = (e) => { if (contRef.current && !contRef.current.contains(e.target)) { setAbierto(false); setTexto(""); } };
    document.addEventListener("mousedown", cerrarSiFuera);
    return () => document.removeEventListener("mousedown", cerrarSiFuera);
  }, []);
  useEffect(() => { setActivo(0); }, [texto]);
  useEffect(() => { if (abierto) listaRef.current?.children[activo]?.scrollIntoView?.({ block: "nearest" }); }, [activo, abierto]);

  const elegir = (o) => { onChange(o ? o.value : ""); setAbierto(false); setTexto(""); };
  const teclado = (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); setAbierto(true); setActivo((i) => Math.min(visibles.length - 1, i + 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setActivo((i) => Math.max(0, i - 1)); }
    else if (e.key === "Enter") { if (abierto && visibles[activo]) { e.preventDefault(); elegir(visibles[activo]); } }
    else if (e.key === "Escape") { setAbierto(false); setTexto(""); }
  };

  return (
    <div ref={contRef} className="relative mt-1">
      <input
        type="text"
        disabled={disabled}
        value={abierto ? texto : (seleccionada?.label || "")}
        onChange={(e) => { setTexto(e.target.value); setAbierto(true); }}
        onFocus={() => { setAbierto(true); setTexto(""); }}
        onClick={() => setAbierto(true)}
        onKeyDown={teclado}
        placeholder={seleccionada ? seleccionada.label : placeholder}
        autoComplete="off"
        className="w-full border border-slate-200 rounded-lg px-3 py-2 pr-8 text-sm placeholder:text-slate-400"
      />
      {value && !disabled && (
        <button type="button" title="Quitar selección" onMouseDown={(e) => { e.preventDefault(); elegir(null); }} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><XCircle size={15} /></button>
      )}
      {abierto && (
        <div className="absolute z-30 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-lg">
          <div ref={listaRef} className="max-h-64 overflow-y-auto">
            {visibles.length === 0 ? (
              <div className="px-3 py-3 text-xs text-slate-400">Sin resultados para "{texto}". Prueba con menos palabras.</div>
            ) : visibles.map((o, i) => (
              <div key={o.value} onMouseDown={(e) => { e.preventDefault(); elegir(o); }} onMouseEnter={() => setActivo(i)} className={`px-3 py-2 text-sm cursor-pointer ${i === activo ? "bg-indigo-50 text-indigo-700" : "text-slate-700"} ${o.value === value ? "font-medium" : ""}`}>{o.label}</div>
            ))}
          </div>
          {filtradas.length > visibles.length && <div className="px-3 py-1.5 text-[11px] text-slate-400 border-t border-slate-100 bg-slate-50 rounded-b-lg">Mostrando {visibles.length} de {filtradas.length} — sigue escribiendo para filtrar.</div>}
        </div>
      )}
    </div>
  );
}

function AutocompletarProveedor({ proveedores, valorTexto, proveedorId, onElegir, onEscribir, className }) {
  const [abierto, setAbierto] = useState(false);
  const [indiceActivo, setIndiceActivo] = useState(0);
  const contenedorRef = useRef(null);
  const listaRef = useRef(null);

  useEffect(() => {
    const cerrarSiClicFuera = (e) => { if (contenedorRef.current && !contenedorRef.current.contains(e.target)) setAbierto(false); };
    document.addEventListener("mousedown", cerrarSiClicFuera);
    return () => document.removeEventListener("mousedown", cerrarSiClicFuera);
  }, []);

  const coincidencias = valorTexto.trim()
    ? proveedores.filter((p) => p.nombre.toLowerCase().includes(valorTexto.trim().toLowerCase())).slice(0, 20)
    : proveedores.slice(0, 20);

  useEffect(() => { setIndiceActivo(0); }, [valorTexto, abierto]);
  useEffect(() => {
    if (abierto && listaRef.current) {
      const activo = listaRef.current.children[indiceActivo];
      if (activo) activo.scrollIntoView({ block: "nearest" });
    }
  }, [indiceActivo, abierto]);

  const manejarTeclas = (e) => {
    if (!abierto || !coincidencias.length) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setIndiceActivo((i) => Math.min(i + 1, coincidencias.length - 1)); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setIndiceActivo((i) => Math.max(i - 1, 0)); }
    else if (e.key === "Enter") { e.preventDefault(); const elegido = coincidencias[indiceActivo]; if (elegido) { onElegir(elegido); setAbierto(false); } }
    else if (e.key === "Escape") setAbierto(false);
  };

  return (
    <div className={`relative ${className || ""}`} ref={contenedorRef}>
      <input
        value={valorTexto}
        onChange={(e) => { onEscribir(e.target.value); setAbierto(true); }}
        onFocus={() => setAbierto(true)}
        onKeyDown={manejarTeclas}
        placeholder="Proveedor — escribe para buscar o crear uno nuevo"
        className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs"
      />
      {abierto && coincidencias.length > 0 && (
        <div ref={listaRef} className="absolute z-20 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-slate-200 rounded-md shadow-lg">
          {coincidencias.map((p, idx) => (
            <button
              key={p.id}
              type="button"
              onMouseEnter={() => setIndiceActivo(idx)}
              onClick={() => { onElegir(p); setAbierto(false); }}
              className={`w-full text-left px-3 py-1.5 text-xs ${idx === indiceActivo ? "bg-indigo-50 text-indigo-700" : proveedorId === p.id ? "bg-indigo-50 text-indigo-700 font-medium" : "text-slate-700"}`}
            >
              {p.nombre}
            </button>
          ))}
        </div>
      )}
      {abierto && valorTexto.trim() && !coincidencias.some((p) => p.nombre.toLowerCase() === valorTexto.trim().toLowerCase()) && (
        <div className="absolute z-20 mt-1 w-full bg-white border border-slate-200 rounded-md shadow-lg px-3 py-1.5 text-[11px] text-emerald-600">
          + Se creará "{valorTexto.trim()}" como proveedor nuevo al guardar
        </div>
      )}
    </div>
  );
}


// una sola cotización (mismo proveedor, mismo archivo) que cubre varios ítems a la vez —
// cada ítem seleccionado puede tener su propio precio dentro del mismo documento.
// Útil cuando un proveedor cotiza varios ítems juntos en un solo PDF.
function CotizacionGeneralForm({ items, proveedores, guardarProveedor, onAplicar, onCerrar }) {
  const [seleccionados, setSeleccionados] = useState(items.map((i) => i.id)); // por defecto: todos (modo "cotización general")
  const [proveedorId, setProveedorId] = useState("");
  const [proveedorNombre, setProveedorNombre] = useState("");
  const [proveedorEmailNuevo, setProveedorEmailNuevo] = useState("");
  const [proveedorNitNuevo, setProveedorNitNuevo] = useState("");
  const [archivoNombre, setArchivoNombre] = useState(null);
  const [moneda, setMoneda] = useState("COP");
  const [tasaCambio, setTasaCambio] = useState(1);
  const [precios, setPrecios] = useState({});
  const [guardando, setGuardando] = useState(false);

  const toggleItem = (id) => setSeleccionados((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const todosSeleccionados = seleccionados.length === items.length;
  const alternarTodos = () => setSeleccionados(todosSeleccionados ? [] : items.map((i) => i.id));

  const listo = seleccionados.length > 0 && (proveedorId || proveedorNombre.trim()) && seleccionados.every((id) => parseFloat(precios[id]) > 0);

  const aplicar = async () => {
    if (!listo) return;
    setGuardando(true);
    let idFinal = proveedorId;
    if (!idFinal && proveedorNombre.trim() && guardarProveedor) {
      const existente = proveedores.find((p) => p.nombre.trim().toLowerCase() === proveedorNombre.trim().toLowerCase());
      if (existente) idFinal = existente.id;
      else {
        const creado = await guardarProveedor({ nombre: proveedorNombre.trim(), nit: proveedorNitNuevo.trim(), actividadEconomica: "", contacto: "", email: proveedorEmailNuevo.trim() });
        if (creado?.id) idFinal = creado.id;
      }
    }
    const base = {
      proveedorId: idFinal || "",
      proveedorNombre: idFinal ? "" : proveedorNombre.trim(),
      precioFinal: "",
      moneda, tasaCambio,
      descuentoTipo: "porcentaje", descuentoValor: "",
      diasEntrega: "", condicionesScore: 5,
      ivaPct: 19,
      archivoNombre,
    };
    const preciosLimpios = {};
    seleccionados.forEach((id) => { preciosLimpios[id] = precios[id]; });
    setGuardando(false);
    onAplicar(preciosLimpios, base);
  };

  return (
    <div className="border border-indigo-200 bg-indigo-50/40 rounded-lg p-3 mb-3 space-y-3">
      <div className="flex items-center justify-between">
        <div className="text-sm font-medium text-slate-700">Cotización general — un mismo proveedor/archivo para varios ítems</div>
        <button onClick={onCerrar} className="text-slate-400 hover:text-slate-600 text-xs">✕ Cerrar</button>
      </div>

      <div>
        <label className="text-[11px] text-slate-500 block mb-1">Ítems que cubre esta cotización</label>
        <label className="flex items-center gap-1.5 text-xs mb-1"><input type="checkbox" checked={todosSeleccionados} onChange={alternarTodos} /> Todos (cotización general para toda la solicitud)</label>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-1">
          {items.map((it, idx) => (
            <label key={it.id} className="flex items-center gap-1.5 text-xs bg-white border border-slate-200 rounded-md px-2 py-1">
              <input type="checkbox" checked={seleccionados.includes(it.id)} onChange={() => toggleItem(it.id)} />
              {idx + 1}. {it.nombre || "(sin nombre)"}
            </label>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2">
        <div>
          <label className="text-[11px] text-slate-500 block mb-1">Proveedor</label>
          <AutocompletarProveedor
            proveedores={proveedores}
            valorTexto={proveedorId ? (proveedores.find((p) => p.id === proveedorId)?.nombre || "") : proveedorNombre}
            proveedorId={proveedorId}
            onElegir={(p) => { setProveedorId(p.id); setProveedorNombre(""); }}
            onEscribir={(texto) => { setProveedorNombre(texto); setProveedorId(""); }}
          />
        </div>
        <div>
          <label className="text-[11px] text-slate-500 block mb-1">Moneda</label>
          <select value={moneda} onChange={(e) => setMoneda(e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs">{MONEDAS.map((m) => <option key={m} value={m}>{m}</option>)}</select>
        </div>
      </div>
      {!proveedorId && proveedorNombre.trim() && (
        <div className="flex gap-2 max-w-xs">
          <input type="text" placeholder="NIT / RUT (opcional)" value={proveedorNitNuevo} onChange={(e) => setProveedorNitNuevo(e.target.value)} className="w-1/2 border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
          <input type="email" placeholder="Correo del proveedor (opcional)" value={proveedorEmailNuevo} onChange={(e) => setProveedorEmailNuevo(e.target.value)} className="w-1/2 border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
        </div>
      )}
      {moneda !== "COP" && (
        <input type="number" placeholder={`Tasa ${moneda}→COP`} value={tasaCambio} onChange={(e) => setTasaCambio(e.target.value)} className="w-40 border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
      )}

      <AdjuntarArchivo nombre={archivoNombre} label="Adjuntar el documento de cotización (PDF/foto)" onSeleccionar={setArchivoNombre} carpeta="cotizaciones-generales" />

      {seleccionados.length > 0 && (
        <div>
          <label className="text-[11px] text-slate-500 block mb-1">Precio por ítem (según lo que dice el documento)</label>
          <div className="space-y-1">
            {items.filter((it) => seleccionados.includes(it.id)).map((it, idx) => (
              <div key={it.id} className="flex items-center gap-2">
                <span className="text-xs text-slate-600 flex-1 truncate">{items.findIndex((x) => x.id === it.id) + 1}. {it.nombre || "(sin nombre)"}</span>
                <InputMiles placeholder="Precio" value={precios[it.id] || ""} onChange={(v) => setPrecios((prev) => ({ ...prev, [it.id]: v }))} className="w-32 border border-slate-200 rounded-md px-2 py-1 text-xs" />
              </div>
            ))}
          </div>
        </div>
      )}

      <button onClick={aplicar} disabled={!listo || guardando} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40">{guardando ? "Aplicando..." : `Aplicar a ${seleccionados.length} ítem(s)`}</button>
      {!listo && !guardando && (
        <div className="text-[11px] text-amber-600">
          Falta:{" "}
          {[
            !seleccionados.length && "seleccionar al menos un ítem",
            !(proveedorId || proveedorNombre.trim()) && "el proveedor",
            seleccionados.some((id) => !(parseFloat(precios[id]) > 0)) && "el precio de uno o más ítems seleccionados",
          ].filter(Boolean).join(", ")}.
        </div>
      )}
    </div>
  );
}

function NuevaSolicitud({ areas, departamentos, empresas, itemsCatalogo, guardarItemCatalogo, proveedores, guardarProveedor, conceptosGasto, usuarios, currentUser, solicitudes, onCrear, onCancel }) {
  const [tipo, setTipo] = useState("compra");
  const [empresaId, setEmpresaId] = useState(empresas[0]?.id || "");
  const [areaId, setAreaId] = useState(currentUser.areaId || areas[0].id);
  const [departamentoId, setDepartamentoId] = useState("");
  const [conceptoGastoId, setConceptoGastoId] = useState("");
  const opcionesConcepto = useMemo(() => conceptosGasto
    .filter((c) => !c.empresaId || c.empresaId === empresaId)
    .sort((a, b) => ((a.grupo || a.codigo) ? 0 : 1) - ((b.grupo || b.codigo) ? 0 : 1) || (a.grupo || "").localeCompare(b.grupo || "") || (a.codigo || "").localeCompare(b.codigo || ""))
    .map((c) => ({ value: c.id, label: labelConcepto(c) })), [conceptosGasto, empresaId]);
  const [fechaEstimada, setFechaEstimada] = useState("");
  const [objetivo, setObjetivo] = useState("");
  const [justificacion, setJustificacion] = useState("");
  const [items, setItems] = useState([{ id: nextId(), itemCatalogoId: "", nombre: "", cantidad: 1, unidad: "unidad", precioEstimado: "", moneda: "COP", tasaCambio: 1, descuentoTipo: "porcentaje", descuentoValor: "", ivaEstimado: 19, aiu: { administracionPct: "", utilidadPct: "", imprevistosPct: "" }, cotizaciones: [] }]);
  const [aiu, setAiu] = useState({ administracionPct: "", utilidadPct: "", imprevistosPct: "" });

  const resumenItemForm = (it) => {
    const d = desgloseItem(it, tipo === "servicio");
    return <span>{it.cantidad} {it.unidad}{parseFloat(it.precioEstimado) > 0 ? ` · ${fmt(tipo === "servicio" ? d.subtotal : d.total)}` : " · sin precio"}</span>;
  };
  const insigniaItemForm = (it) => (
    <>
      {!(it.nombre || "").trim() && <Badge tone="amber">Sin nombre</Badge>}
      {itemConPrecioInvalido(it) && <Badge tone="red">Revisar precio</Badge>}
      {cotizacionUnicaSinJustificar(it) && <Badge tone="amber">Falta comentario</Badge>}
      {(it.cotizaciones || []).length > 0 && <span className="flex items-center gap-0.5 text-slate-500" title="Cotizaciones adjuntas"><Paperclip size={11} />{it.cotizaciones.length}</span>}
      {planTieneValores(it.pagosSugeridos) && <span className="text-indigo-600" title="Tiene plan de pagos sugerido"><CalendarClock size={12} /></span>}
    </>
  );
  const abItems = useAbiertosItems(items, true); // en el formulario todo arranca abierto; al agregar un ítem, los anteriores se contraen
  const addItem = () => { abItems.todos(false); setItems([...items, { id: nextId(), itemCatalogoId: "", nombre: "", cantidad: 1, unidad: "unidad", precioEstimado: "", moneda: "COP", tasaCambio: 1, descuentoTipo: "porcentaje", descuentoValor: "", ivaEstimado: 19, aiu: { administracionPct: "", utilidadPct: "", imprevistosPct: "" }, cotizaciones: [] }]); };
  const removeItem = (id) => setItems(items.filter((i) => i.id !== id));
  const updateItem = (id, field, val) => setItems((prev) => prev.map((i) => (i.id === id ? { ...i, [field]: val } : i)));
  const [cargandoTasaItem, setCargandoTasaItem] = useState(null);
  const actualizarTasaItem = async (id, moneda) => {
    if (!moneda || moneda === "COP") return;
    setCargandoTasaItem(id);
    const tasa = await obtenerTasaCambioCOP(moneda);
    setCargandoTasaItem(null);
    if (tasa) updateItem(id, "tasaCambio", tasa.toFixed(2));
    else alert("No se pudo obtener la tasa de cambio automática. Ingrésala manualmente.");
  };
  const cambiarMonedaItem = (id, moneda) => { updateItem(id, "moneda", moneda); if (moneda !== "COP") actualizarTasaItem(id, moneda); };
  const setCotizacionesItem = (itemId, cots) => setItems(items.map((i) => (i.id === itemId ? { ...i, cotizaciones: cots } : i)));

  const [mostrarCotGeneral, setMostrarCotGeneral] = useState(false);
  const totalGeneral = tipo === "servicio" ? desgloseSolicitud({ tipo, items, aiu }) : items.reduce((acc, it) => { const d = desgloseItem(it); return { subtotal: acc.subtotal + d.subtotal, iva: acc.iva + d.iva, total: acc.total + d.total }; }, { subtotal: 0, iva: 0, total: 0 });
  // presupuesto disponible del área elegida — mismo cálculo que usa el Dashboard: solicitudes que
  // ya están comprometiendo presupuesto (no descartadas ni todavía sin aprobar el primer paso)
  const areaSel = areas.find((a) => a.id === areaId);
  const comprometidoArea = useMemo(() => {
    return solicitudes.filter((s) => s.areaId === areaId && !["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)).reduce((acc, s) => acc + totalSolicitud(s), 0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [solicitudes, areaId]);
  const disponibleArea = (areaSel?.presupuesto || 0) - comprometidoArea;
  const seSalDelPresupuesto = areaSel?.presupuesto > 0 && totalGeneral.total > disponibleArea;
  // aplica una misma cotización (proveedor + archivo) a varios ítems seleccionados de una sola vez,
  // cada uno con su propio precio dentro del mismo documento
  const aplicarCotizacionGeneral = (precios, cotizacionBase) => {
    setItems((prev) => prev.map((i) => {
      if (!(i.id in precios)) return i;
      if (i.cotizaciones.length >= 3) return i;
      return { ...i, cotizaciones: [...i.cotizaciones, { ...cotizacionBase, precioUnitario: precios[i.id], unidadCotizada: i.unidad, factorConversion: 1 }] };
    }));
    setMostrarCotGeneral(false);
  };

  const submit = () => {
    const sinNombre = items.findIndex((i) => !i.nombre.trim());
    if (sinNombre >= 0) { abItems.abrir(items[sinNombre].id); alert(`Falta el nombre del ítem ${sinNombre + 1}.`); return; }
    if (!items.length || !objetivo.trim() || !justificacion.trim()) return;
    if (!conceptoGastoId) { alert("Falta elegir el concepto de gasto."); return; }
    const idxUnica = items.findIndex((i) => cotizacionUnicaSinJustificar(i));
    if (idxUnica >= 0) { abItems.abrir(items[idxUnica].id); alert(`Ítem ${idxUnica + 1}: adjuntaste una sola cotización, escribe el comentario de por qué solo se cotizó con un proveedor.`); return; }
    const idxPrecio = items.findIndex((i) => itemConPrecioInvalido(i));
    if (idxPrecio >= 0) { abItems.abrir(items[idxPrecio].id); alert(`Ítem ${idxPrecio + 1}: ${primerErrorPrecio(items[idxPrecio])}`); return; }
    // si el solicitante empezó a llenar el plan de pagos sugerido de un ítem, debe cuadrar exacto con el total
    // de ESE ítem — si lo dejó vacío no pasa nada, es opcional
    for (let k = 0; k < items.length; k++) {
      const ps = items[k].pagosSugeridos;
      if (!planTieneValores(ps)) continue;
      const totalIt = totalItemConAiu(items[k], tipo === "servicio");
      const restanteIt = totalIt - totalPagado(ps);
      if (Math.abs(restanteIt) > 0.5) {
        abItems.abrir(items[k].id);
        alert(`Ítem ${k + 1}: el plan de pagos sugerido ${restanteIt > 0 ? `no cubre el total (faltan ${fmt(restanteIt)})` : `supera el total en ${fmt(-restanteIt)}`}. Ajústalo o déjalo vacío si no quieres sugerir uno.`);
        return;
      }
      const faltaFechaIt = ps.tipoPago === "contado" ? !ps.pagoUnico.fecha : (!ps.anticipo.fecha || !ps.final.fecha || (ps.intermedio.activo && !ps.intermedio.fecha));
      if (faltaFechaIt) { abItems.abrir(items[k].id); alert(`Ítem ${k + 1}: falta poner la fecha del plan de pagos sugerido.`); return; }
    }
    // los ítems escritos a mano (sin elegir del catálogo) también quedan guardados ahí, para no perder esa información
    items.forEach((it) => {
      if (!it.itemCatalogoId && it.nombre.trim()) {
        const yaExiste = itemsCatalogo.some((c) => c.nombre.trim().toLowerCase() === it.nombre.trim().toLowerCase());
        if (!yaExiste) guardarItemCatalogo({ nombre: it.nombre.trim(), unidadDefault: it.unidad, categoria: "" });
      }
    });
    const jefe = usuarios.find((u) => tieneAreaACargo(u, areaId) && ["Jefe de Área", "Jefe de Área y Director"].includes(u.rol));
    const director = usuarios.find((u) => tieneAreaACargo(u, areaId) && ["Director de Área", "Jefe de Área y Director"].includes(u.rol));
    const folio = "SOL-" + (1000 + Math.floor(Math.random() * 8999));
    // si quien crea la solicitud es el propio jefe del área seleccionada, queda auto-aprobada en ese paso
    // (no tiene sentido que se apruebe a sí mismo con un clic aparte) — pero igual pasa por Director de Área,
    // salvo que la misma persona también tenga el rol combinado, en cuyo caso se salta los dos pasos
    const esJefeDeSuPropiaArea = ["Jefe de Área", "Jefe de Área y Director"].includes(currentUser.rol) && tieneAreaACargo(currentUser, areaId);
    const esAmbosRoles = currentUser.rol === "Jefe de Área y Director" && tieneAreaACargo(currentUser, areaId);
    const statusInicial = esAmbosRoles ? "cotizando" : esJefeDeSuPropiaArea ? "aprobacion_director" : "aprobacion_jefe";
    onCrear({
      id: nextId(), folio,
      tipo, empresaId, areaId, departamentoId: departamentoId || null, centroCosto: conceptosGasto.find((c) => c.id === conceptoGastoId)?.centroCosto || "", conceptoGastoId, solicitanteId: currentUser.id,
      fechaCreacion: hoy(), fechaEstimada, objetivo, justificacion,
      status: statusInicial,
      presupuestoAlEnviar: { presupuesto: areaSel?.presupuesto || 0, comprometido: comprometidoArea, disponible: disponibleArea, total: totalGeneral.total },
      aiu: tipo === "servicio" ? aiu : { administracionPct: "", utilidadPct: "", imprevistosPct: "" },
      revisionCompras: tipo === "compra" ? { estado: "pendiente", observacion: "", usuario: "", fecha: "" } : { estado: "no_aplica", observacion: "", usuario: "", fecha: "" },
      items: items.map((i) => ({ ...i, ivaEstimado: tipo === "servicio" ? 0 : i.ivaEstimado, cotizacionSeleccionada: null, observacionSeleccion: "" })),
      firmas: {
        solicitante: { nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresas.find((e) => e.id === empresaId)?.nombre || "", fecha: hoy(), fotoUrl: currentUser.firmaFotoUrl || null },
        jefe: esJefeDeSuPropiaArea
          ? { aprobado: true, nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresas.find((e) => e.id === empresaId)?.nombre || "", fecha: hoy(), observacion: "Creada y aprobada por el mismo jefe de área.", fotoUrl: currentUser.firmaFotoUrl || null }
          : { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
        director: esAmbosRoles
          ? { aprobado: true, nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresas.find((e) => e.id === empresaId)?.nombre || "", fecha: hoy(), observacion: "Aprobado junto con el paso de jefe de área (mismo responsable).", fotoUrl: currentUser.firmaFotoUrl || null }
          : { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
        financiera: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
        gerencia: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
      },
      pagosSugeridos: planPagosVacio(), pagos: planPagosVacio(), pagosConfirmados: false,
      ocEnviada: { ordenesProveedor: [] },
    prioridad: null,
    evaluacionProveedor: evaluacionProveedorVacia(),
      recepcion: { archivos: [], comentario: "", recibidoSatisfaccion: false, usuario: "", fecha: "" },
      historialEstados: [{ status: "solicitud", fecha: ahoraISO() }, { status: statusInicial, fecha: ahoraISO() }],
      notificaciones: esAmbosRoles
        ? [{ fecha: ahoraISO(), mensaje: `Solicitud creada y auto-aprobada por ${currentUser.nombre} (jefe de área y director) — lista para cotizar.` }]
        : esJefeDeSuPropiaArea
        ? [{ fecha: ahoraISO(), mensaje: director?.email ? `Solicitud auto-aprobada por ${currentUser.nombre} (jefe de área). Correo enviado a ${director.nombre} (${director.email}) para su aprobación.` : `Solicitud auto-aprobada por ${currentUser.nombre} (jefe de área). No hay un director de área con correo configurado para notificar.` }]
        : [{ fecha: ahoraISO(), mensaje: jefe?.email ? `Correo enviado a ${jefe.nombre} (${jefe.email})` : "Solicitud creada. No hay un jefe de área con correo configurado para notificar." }],
    });
    if (!esJefeDeSuPropiaArea && jefe?.email) {
      enviarCorreo(
        jefe.email,
        `Nueva solicitud pendiente: ${folio}`,
        `<p>Hola ${jefe.nombre},</p><p><b>${currentUser.nombre}</b> creó la solicitud <b>${folio}</b> (${tipo === "compra" ? "Solicitud de compra" : "Orden de servicio/trabajo"}) y quedó pendiente de tu aprobación.</p><p><b>Objetivo:</b> ${objetivo}</p>`
      );
    } else if (esJefeDeSuPropiaArea && !esAmbosRoles && director?.email) {
      enviarCorreo(
        director.email,
        `Solicitud pendiente de tu aprobación: ${folio}`,
        `<p>Hola ${director.nombre},</p><p>La solicitud <b>${folio}</b> fue creada y auto-aprobada por ${currentUser.nombre} (jefe de área) y quedó pendiente de tu aprobación como Director de Área.</p><p><b>Objetivo:</b> ${objetivo}</p>`
      );
    }
  };

  return (
    <div className="flex gap-5 items-start max-w-6xl">
    <div className="bg-white rounded-xl border border-slate-200 p-6 flex-1 min-w-0">
      <h2 className="text-lg font-semibold text-slate-800 mb-5">Nueva solicitud</h2>
      <div className="grid grid-cols-2 gap-4 mb-5">
        <div>
          <label className="text-xs font-medium text-slate-500">Tipo de solicitud</label>
          <div className="flex gap-2 mt-1">
            <button onClick={() => setTipo("compra")} className={`flex-1 flex items-center justify-center gap-2 px-3 py-2 rounded-lg border text-sm font-medium ${tipo === "compra" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}><ShoppingCart size={15} /> Solicitud de compra</button>
            <button onClick={() => setTipo("servicio")} className={`flex-1 flex items-center justify-center gap-2 px-3 py-2 rounded-lg border text-sm font-medium ${tipo === "servicio" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}><Wrench size={15} /> Orden de servicio/trabajo</button>
          </div>
        </div>
        <div><label className="text-xs font-medium text-slate-500">Empresa</label><select value={empresaId} onChange={(e) => { setEmpresaId(e.target.value); setConceptoGastoId(""); }} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Área solicitante</label><select value={areaId} onChange={(e) => setAreaId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{areas.map((a) => <option key={a.id} value={a.id}>{a.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Departamento que reporta</label><select value={departamentoId} onChange={(e) => setDepartamentoId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm"><option value="">— Sin especificar —</option>{departamentos.map((d) => <option key={d.id} value={d.id}>{d.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Solicitante</label><div className="w-full mt-1 border border-slate-100 bg-slate-50 rounded-lg px-3 py-2 text-sm text-slate-500">{currentUser.nombre} (firma automática)</div></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500">Concepto de gasto <span className="text-slate-400 font-normal">(Grupo – Código – Cuenta – Centro de costo)</span></label><SelectorBuscable opciones={opcionesConcepto} value={conceptoGastoId} onChange={setConceptoGastoId} placeholder="Escribe para buscar: grupo, código, cuenta o centro de costo..." /></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500">{tipo === "compra" ? "Fecha estimada de entrega" : "Fecha estimada de terminación"}</label><InputFecha value={fechaEstimada} onChange={setFechaEstimada} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm" /></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><Target size={12} /> Objetivo</label><textarea value={objetivo} onChange={(e) => { setObjetivo(e.target.value); autoResize(e); }} rows={2} placeholder="¿Qué se busca lograr con esta solicitud?" className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none overflow-hidden" /></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><ClipboardList size={12} /> Justificación</label><textarea value={justificacion} onChange={(e) => { setJustificacion(e.target.value); autoResize(e); }} rows={2} placeholder="¿Por qué es necesaria?" className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none overflow-hidden" /></div>
      </div>

      <div className="mb-2 flex items-center justify-between flex-wrap gap-2">
        <label className="text-xs font-medium text-slate-500">Ítems solicitados</label>
        <div className="flex items-center gap-3">
          <button onClick={() => setMostrarCotGeneral(true)} className="text-xs text-slate-600 font-medium flex items-center gap-1 border border-slate-200 rounded-md px-2 py-1"><FileText size={13} /> Cotización general</button>
          <button onClick={addItem} className="text-xs text-indigo-600 font-medium flex items-center gap-1"><Plus size={13} /> Agregar ítem</button>
        </div>
      </div>

      {mostrarCotGeneral && (
        <CotizacionGeneralForm items={items} proveedores={proveedores} guardarProveedor={guardarProveedor} onAplicar={aplicarCotizacionGeneral} onCerrar={() => setMostrarCotGeneral(false)} />
      )}

      <div className="space-y-2 mb-5">
        <ControlExpandirTodo n={items.length} onTodos={abItems.todos} />
        {items.map((it, idx) => (
          <ItemColapsable key={it.id} abierto={abItems.abierto(it.id)} onToggle={() => abItems.alternar(it.id)} numero={idx + 1} titulo={it.nombre} tinte={tinteItem(idx)} resumen={resumenItemForm(it)} insignia={insigniaItemForm(it)}>
          <div className="space-y-1.5">
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-slate-400 shrink-0 w-5 text-right">{idx + 1}.</span>
              <div className="flex-1">
                <AutocompletarItem
                  itemsCatalogo={itemsCatalogo}
                  valorTexto={it.nombre}
                  catalogoId={it.itemCatalogoId}
                  onElegir={(cat) => setItems((prev) => prev.map((i) => (i.id === it.id ? { ...i, itemCatalogoId: cat.id, nombre: cat.nombre, unidad: cat.unidadDefault } : i)))}
                  onEscribir={(texto) => setItems((prev) => prev.map((i) => (i.id === it.id ? { ...i, itemCatalogoId: "", nombre: texto } : i)))}
                />
              </div>
            </div>
            <div className="flex gap-2 items-start flex-wrap">
              <input type="number" min="0" placeholder="Cant." value={it.cantidad} onChange={(e) => updateItem(it.id, "cantidad", e.target.value)} className="w-16 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              <select value={it.unidad} onChange={(e) => updateItem(it.id, "unidad", e.target.value)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm">{UNIDADES.map((u) => <option key={u} value={u}>{u}</option>)}</select>
              <InputMiles placeholder="Precio est." value={it.precioEstimado} onChange={(v) => updateItem(it.id, "precioEstimado", v)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              <select value={it.moneda} onChange={(e) => cambiarMonedaItem(it.id, e.target.value)} className="w-20 border border-slate-200 rounded-md px-2 py-1.5 text-sm">{MONEDAS.map((m) => <option key={m} value={m}>{m}</option>)}</select>
              {it.moneda !== "COP" && (
                <div className="flex items-center gap-1">
                  <input type="number" min="0" step="0.01" placeholder={`Tasa ${it.moneda}→COP`} title={`¿Cuántos COP equivalen a 1 ${it.moneda}?`} value={it.tasaCambio} onChange={(e) => updateItem(it.id, "tasaCambio", e.target.value)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
                  <button type="button" onClick={() => actualizarTasaItem(it.id, it.moneda)} disabled={cargandoTasaItem === it.id} title="Actualizar tasa del día" className="text-slate-400 hover:text-indigo-600 disabled:opacity-50 shrink-0">{cargandoTasaItem === it.id ? "..." : "↻"}</button>
                </div>
              )}
              <select value={it.descuentoTipo} onChange={(e) => updateItem(it.id, "descuentoTipo", e.target.value)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm">
                <option value="porcentaje">Desc. %</option>
                <option value="valor">Desc. $</option>
              </select>
              {it.descuentoTipo === "valor" ? (
                <InputMiles placeholder="Descuento en $" value={it.descuentoValor} onChange={(v) => updateItem(it.id, "descuentoValor", v)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              ) : (
                <input type="number" min="0" placeholder="Descuento en %" value={it.descuentoValor} onChange={(e) => updateItem(it.id, "descuentoValor", e.target.value)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              )}
              {tipo !== "servicio" && <select value={it.ivaEstimado} onChange={(e) => updateItem(it.id, "ivaEstimado", e.target.value)} className="w-20 border border-slate-200 rounded-md px-2 py-1.5 text-sm">{IVA_OPCIONES.map((v) => <option key={v} value={v}>IVA {v}%</option>)}</select>}
              {items.length > 1 && <button onClick={() => removeItem(it.id)} className="text-slate-400 hover:text-rose-500 p-1.5"><Trash2 size={15} /></button>}
            </div>
            {parseFloat(it.precioEstimado) > 0 && (() => {
              const d = desgloseItem(it, tipo === "servicio");
              const inicial = parseFloat(it.precioEstimado) || 0;
              const descuento = parseFloat(it.descuentoValor) || 0;
              const conDescuento = descuento > 0 ? Math.max(0, it.descuentoTipo === "porcentaje" ? inicial * (1 - descuento / 100) : inicial - descuento) : inicial;
              const ahorro = inicial - conDescuento;
              const simbolo = it.moneda && it.moneda !== "COP" ? `${it.moneda} ` : "$";
              return (
                <div className="text-[11px] text-slate-500 pl-1 space-y-0.5">
                  {ahorro > 0 && <div>Precio inicial: {simbolo}{inicial.toLocaleString("es-CO")} · Descuento: -{simbolo}{ahorro.toLocaleString("es-CO", { maximumFractionDigits: 2 })} · Precio con descuento: <b>{simbolo}{conDescuento.toLocaleString("es-CO", { maximumFractionDigits: 2 })}</b></div>}
                  {tipo === "servicio" ? <div>Subtotal (Costo Directo de este ítem): <b>{fmt(d.subtotal)}</b></div> : <div>Subtotal: {fmt(d.subtotal)} · IVA: {fmt(d.iva)} · <b>Total: {fmt(d.total)}</b></div>}
                </div>
              );
            })()}
            {/* el solicitante puede adjuntar hasta 3 cotizaciones desde ya, opcional */}
            <CotizacionForm item={it} proveedores={proveedores} guardarProveedor={guardarProveedor} onGuardar={(_, cots) => setCotizacionesItem(it.id, cots)} compacto opcionalTitulo="Adjuntar cotizaciones (opcional, máx. 3)" sinIva={tipo === "servicio"} />
            <div className="border-t border-slate-200/70 pt-1.5">
              {totalItemConAiu(it, tipo === "servicio") > 0
                ? <PlanPagoCotizacion pagos={it.pagosSugeridos} total={totalItemConAiu(it, tipo === "servicio")} onChange={(pg) => updateItem(it.id, "pagosSugeridos", pg)} etiqueta="+ Sugerir plan de pagos para este ítem (opcional — pasa como valor inicial a Compras y Dirección Financiera)" titulo="Plan de pagos sugerido de este ítem" />
                : <div className="text-[11px] text-amber-600">Pon un precio estimado para poder sugerir un plan de pagos.</div>}
            </div>
          </div>
          </ItemColapsable>
        ))}
      </div>

      <div className="flex gap-2 justify-end">
        <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm text-slate-500 border border-slate-200">Cancelar</button>
        <button onClick={submit} className="px-4 py-2 rounded-lg text-sm bg-indigo-600 text-white font-medium">Enviar solicitud</button>
      </div>
    </div>

    {/* Resumen a la derecha — aprovecha el espacio en blanco */}
    <div className="w-72 shrink-0 sticky top-4 hidden lg:block">
      <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
        <div className="font-medium text-slate-700 text-sm">Resumen</div>
        <div className="text-xs text-slate-500 space-y-1">
          <div className="flex justify-between"><span>Ítems</span><span className="font-medium text-slate-700">{items.length}</span></div>
          <div className="flex justify-between"><span>Empresa</span><span className="font-medium text-slate-700 text-right">{empresas.find((e) => e.id === empresaId)?.nombre || "—"}</span></div>
          <div className="flex justify-between"><span>Área</span><span className="font-medium text-slate-700 text-right">{areas.find((a) => a.id === areaId)?.nombre || "—"}</span></div>
        </div>
        <div className="border-t border-slate-100 pt-3 space-y-1 text-xs">
          {items.map((it, idx) => {
            const d = desgloseItem(it, tipo === "servicio");
            if (!(d.subtotal > 0)) return null;
            return <div key={it.id} className="flex justify-between text-slate-500"><span className="truncate pr-2">{idx + 1}. {it.nombre || "(sin nombre)"}</span><span className="shrink-0 text-slate-700">{fmt(tipo === "servicio" ? d.subtotal : d.total)}</span></div>;
          })}
        </div>
        {tipo === "servicio" ? (
          <div className="border-t border-slate-200 pt-3 space-y-1">
            <div className="flex justify-between text-xs text-slate-500"><span>Costo Directo</span><span>{fmt(totalGeneral.costoDirecto)}</span></div>
            <div className="text-[10px] text-slate-400">Los costos indirectos (AIU) los define Compras al cotizar.</div>
            <div className="flex justify-between text-sm font-semibold text-slate-800 pt-1"><span>Total</span><span>{fmt(totalGeneral.total)}</span></div>
          </div>
        ) : (
          <div className="border-t border-slate-200 pt-3 space-y-1">
            <div className="flex justify-between text-xs text-slate-500"><span>Subtotal</span><span>{fmt(totalGeneral.subtotal)}</span></div>
            <div className="flex justify-between text-xs text-slate-500"><span>IVA</span><span>{fmt(totalGeneral.iva)}</span></div>
            <div className="flex justify-between text-sm font-semibold text-slate-800 pt-1"><span>Total</span><span>{fmt(totalGeneral.total)}</span></div>
          </div>
        )}
        {requiereGerencia(totalGeneral.total) && <div className="text-[11px] text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2 py-1.5">Este monto requerirá aprobación de Gerencia.</div>}
        {!requiereGerencia(totalGeneral.total) && requiereDireccion(totalGeneral.total) && <div className="text-[11px] text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2 py-1.5">Este monto requerirá aprobación de Dirección Financiera.</div>}

        {areaSel && (
          <div className="border-t border-slate-200 pt-3 space-y-1.5">
            <div className="text-xs font-medium text-slate-500 flex items-center gap-1"><DollarSign size={12} /> Presupuesto de {areaSel.nombre}</div>
            {areaSel.presupuesto > 0 ? (
              <>
                <div className="flex justify-between text-xs text-slate-500"><span>Presupuesto</span><span>{fmt(areaSel.presupuesto)}</span></div>
                <div className="flex justify-between text-xs text-slate-500"><span>Ya comprometido</span><span>{fmt(comprometidoArea)}</span></div>
                <div className="flex justify-between text-xs text-slate-500"><span>Disponible</span><span className={disponibleArea < 0 ? "text-rose-600 font-medium" : "text-emerald-600 font-medium"}>{fmt(disponibleArea)}</span></div>
                <div className="w-full bg-slate-100 rounded-full h-1.5 overflow-hidden"><div className={`h-1.5 rounded-full ${seSalDelPresupuesto ? "bg-rose-500" : "bg-emerald-500"}`} style={{ width: `${Math.min(100, (comprometidoArea / areaSel.presupuesto) * 100)}%` }} /></div>
                {totalGeneral.total > 0 && (
                  seSalDelPresupuesto
                    ? <div className="text-[11px] text-rose-600 bg-rose-50 border border-rose-200 rounded-md px-2 py-1.5">⚠ Esta solicitud supera el presupuesto disponible del área por {fmt(totalGeneral.total - disponibleArea)}.</div>
                    : <div className="text-[11px] text-emerald-600">✓ Esta solicitud cabe dentro del presupuesto disponible.</div>
                )}
              </>
            ) : (
              <div className="text-[11px] text-slate-400">Esta área no tiene un presupuesto mensual configurado.</div>
            )}
          </div>
        )}
      </div>
    </div>
    </div>
  );
}

/* ---------------------------------------------------------
   COTIZACIONES Y COMPARATIVO
--------------------------------------------------------- */
// plan de pagos opcional propio de una cotización/proveedor específico — cubre tanto "por cotización"
// como "por ítem" (ya que cada cotización pertenece a un ítem). Si no se activa, no aplica nada
// especial: se sigue usando el plan general de la solicitud como siempre.
function PlanPagoCotizacion({ pagos, total, onChange, etiqueta, titulo }) {
  const [abierto, setAbierto] = useState(!!pagos);
  const p = { ...planPagosVacio(), ...pagos };
  const restante = total - totalPagado(p);
  const faltaFecha = p.tipoPago === "contado" ? !p.pagoUnico.fecha : (!p.anticipo.fecha || !p.final.fecha || (p.intermedio.activo && !p.final.fecha));

  const set = (campo, sub, val) => {
    if (sub === "fecha" && val) { const error = validarOrdenFechas(p, campo, val); if (error) { alert(error); return; } }
    onChange({ ...p, [campo]: { ...p[campo], [sub]: val } });
  };
  const setTipoPago = (tipo) => onChange({ ...p, tipoPago: tipo, pagoUnico: tipo === "contado" ? { ...p.pagoUnico, valor: total } : p.pagoUnico });

  if (!abierto) {
    return <button type="button" onClick={() => setAbierto(true)} className="text-[11px] text-indigo-600 underline">{etiqueta || "+ Definir plan de pagos propio de este proveedor (opcional)"}</button>;
  }

  return (
    <div className="border border-indigo-100 bg-indigo-50/30 rounded-lg p-2 space-y-1.5">
      <div className="flex items-center justify-between">
        <span className="text-[11px] font-medium text-slate-600">{titulo || "Plan de pagos de este proveedor"}</span>
        <button type="button" onClick={() => { onChange(null); setAbierto(false); }} className="text-[10px] text-slate-400 hover:text-rose-500">Quitar</button>
      </div>
      <div className="flex gap-1.5">
        <button type="button" onClick={() => setTipoPago("plan")} className={`px-2 py-0.5 rounded text-[10px] font-medium border ${p.tipoPago !== "contado" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>Por etapas</button>
        <button type="button" onClick={() => setTipoPago("contado")} className={`px-2 py-0.5 rounded text-[10px] font-medium border ${p.tipoPago === "contado" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>Pago único</button>
      </div>
      {p.tipoPago === "contado" ? (
        <div className="flex gap-2 items-center">
          <span className="text-[10px] text-slate-500">Valor (= total): {fmt(total)}</span>
          <InputFecha value={p.pagoUnico.fecha} onChange={(v) => set("pagoUnico", "fecha", v)} className="border border-slate-200 rounded-md px-2 py-1 text-xs" />
        </div>
      ) : (
        <div className="grid grid-cols-3 gap-1.5">
          <div><InputMiles placeholder="Anticipo" value={p.anticipo.valor} onChange={(v) => set("anticipo", "valor", v)} className="w-full border border-slate-200 rounded-md px-1.5 py-1 text-[11px] mb-1" /><InputFecha value={p.anticipo.fecha} onChange={(v) => set("anticipo", "fecha", v)} className="w-full border border-slate-200 rounded-md px-1.5 py-1 text-[11px]" /></div>
          <div><label className="text-[10px] flex items-center gap-1"><input type="checkbox" checked={p.intermedio.activo} onChange={(e) => set("intermedio", "activo", e.target.checked)} /> Interm.</label><InputMiles placeholder="Valor" disabled={!p.intermedio.activo} value={p.intermedio.valor} onChange={(v) => set("intermedio", "valor", v)} className="w-full border border-slate-200 rounded-md px-1.5 py-1 text-[11px] mb-1 disabled:bg-slate-100" /><InputFecha disabled={!p.intermedio.activo} value={p.intermedio.fecha} onChange={(v) => set("intermedio", "fecha", v)} className="w-full border border-slate-200 rounded-md px-1.5 py-1 text-[11px] disabled:bg-slate-100" /></div>
          <div><InputMiles placeholder="Final" value={p.final.valor} onChange={(v) => set("final", "valor", v)} className="w-full border border-slate-200 rounded-md px-1.5 py-1 text-[11px] mb-1" /><InputFecha value={p.final.fecha} onChange={(v) => set("final", "fecha", v)} className="w-full border border-slate-200 rounded-md px-1.5 py-1 text-[11px]" /></div>
        </div>
      )}
      {totalPagado(p) > 0 && (
        <div className={`text-[10px] ${Math.abs(restante) > 0.5 ? "text-amber-600" : "text-emerald-600"}`}>
          {Math.abs(restante) > 0.5 ? `Falta cuadrar: ${fmt(Math.abs(restante))}` : "✓ Cuadra con el total"}
        </div>
      )}
    </div>
  );
}

function CotizacionForm({ item, proveedores, guardarProveedor, onGuardar, compacto, opcionalTitulo, sinIva, tinte }) {
  const [abierto, setAbierto] = useState(!compacto);
  const [cots, setCots] = useState(item.cotizaciones.length ? item.cotizaciones : []);
  const [guardadoMsg, setGuardadoMsg] = useState(false);
  const update = (i, field, val) => setCots((prev) => prev.map((c, idx) => (idx === i ? { ...c, [field]: val } : c)));
  const updateAiu = (i, campo, val) => setCots((prev) => prev.map((c, idx) => (idx === i ? { ...c, aiu: { ...(c.aiu || {}), [campo]: pctValido(val) } } : c)));
  const [cargandoTasa, setCargandoTasa] = useState(null);
  const actualizarTasaAutomatica = async (i, moneda) => {
    if (!moneda || moneda === "COP") return;
    setCargandoTasa(i);
    const tasa = await obtenerTasaCambioCOP(moneda);
    setCargandoTasa(null);
    if (tasa) update(i, "tasaCambio", tasa.toFixed(2));
    else alert("No se pudo obtener la tasa de cambio automática. Ingrésala manualmente.");
  };
  const cambiarMoneda = (i, moneda) => { update(i, "moneda", moneda); if (moneda !== "COP") actualizarTasaAutomatica(i, moneda); };
  const addCot = () => cots.length < 3 && setCots([...cots, { proveedorId: "", proveedorNombre: "", unidadCotizada: item.unidad, factorConversion: 1, precioUnitario: item.precioEstimado || "", precioFinal: "", moneda: item.moneda || "COP", tasaCambio: item.tasaCambio || 1, descuentoTipo: "porcentaje", descuentoValor: "", diasEntrega: "", condicionesScore: 5, ivaPct: sinIva ? 0 : (item.ivaEstimado ?? 19), aiu: { administracionPct: "", utilidadPct: "", imprevistosPct: "" }, pagos: null, archivoNombre: "" }]);
  const removeCot = (i) => setCots(cots.filter((_, idx) => idx !== i));

  // guarda automáticamente lo que ya se alcanzó a escribir (incluido el archivo adjunto), sin depender
  // de que se le dé clic al botón — así nada se pierde si alguien solo adjunta el archivo y no le da "Guardar"
  useEffect(() => {
    const validas = cots.filter((c) => (c.proveedorId || c.proveedorNombre) && c.precioUnitario);
    onGuardar(item.id, validas);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cots]);

  const guardar = async () => {
    if (cots.some((c) => c.precioUnitario && !(c.proveedorId || (c.proveedorNombre || "").trim()))) { alert("Toda cotización debe tener el nombre del proveedor. Escríbelo o elimina esa fila."); return; }
    const conPrecioInvalido = cots.find((c) => c.precioUnitario && erroresPrecioCotizacion(item, c).length > 0);
    if (conPrecioInvalido) { alert(erroresPrecioCotizacion(item, conPrecioInvalido)[0].texto + " Corrígelo antes de guardar."); return; }
    const validasPrevias = cots.filter((c) => (c.proveedorId || c.proveedorNombre) && c.precioUnitario);
    if (validasPrevias.length === 1 && !(validasPrevias[0].justificacionUnico || "").trim()) { alert("Solo hay una cotización para este ítem: escribe el comentario de por qué solo se cotizó con un proveedor."); return; }
    let listaCots = [...cots];
    // los proveedores escritos a mano (sin elegirlos del catálogo) también quedan guardados ahí,
    // y la cotización queda vinculada a su ID real (no solo al nombre) para que el envío de correo no falle
    if (guardarProveedor) {
      for (let idx = 0; idx < listaCots.length; idx++) {
        const c = listaCots[idx];
        const esValida = (c.proveedorId || c.proveedorNombre) && c.precioUnitario;
        if (!esValida || c.proveedorId || !c.proveedorNombre?.trim()) continue;
        const nombreLimpio = c.proveedorNombre.trim();
        const emailNuevo = (c.proveedorEmailNuevo || "").trim();
        const nitNuevo = (c.proveedorNitNuevo || "").trim();
        const existente = proveedores.find((p) => p.nombre.trim().toLowerCase() === nombreLimpio.toLowerCase());
        if (!existente) {
          const creado = await guardarProveedor({ nombre: nombreLimpio, nit: nitNuevo, actividadEconomica: "", contacto: "", email: emailNuevo });
          if (creado?.id) listaCots[idx] = { ...c, proveedorId: creado.id, proveedorNombre: "" };
        } else {
          if (emailNuevo && !existente.email) await guardarProveedor({ ...existente, email: emailNuevo });
          listaCots[idx] = { ...c, proveedorId: existente.id, proveedorNombre: "" };
        }
      }
    }
    setCots(listaCots);
    const validas = listaCots.filter((c) => (c.proveedorId || c.proveedorNombre) && c.precioUnitario);
    onGuardar(item.id, validas);
    setGuardadoMsg(true); setTimeout(() => setGuardadoMsg(false), 2500);
  };

  if (compacto && !abierto) return <button onClick={() => setAbierto(true)} className="text-[11px] text-indigo-600 font-medium flex items-center gap-1"><Paperclip size={11} /> {opcionalTitulo || "Adjuntar cotizaciones"}</button>;

  return (
    <div className={`${tinte?.fondo || "bg-slate-50"} rounded-lg p-3 border ${tinte?.borde || "border-slate-200"}`}>
      <div className="flex items-center justify-between mb-2">
        <div className="text-sm font-medium text-slate-700">
          {item.nombre} <span className="text-slate-400 font-normal">({item.cantidad} {item.unidad})</span>
          {parseFloat(item.precioEstimado) > 0 && (
            <span className="text-[11px] text-slate-400 font-normal ml-2">— precio solicitado: {item.moneda && item.moneda !== "COP" ? `${item.moneda} ${Number(item.precioEstimado).toLocaleString("es-CO")}` : fmt(item.precioEstimado)}{item.descuentoValor > 0 && ` (con ${item.descuentoTipo === "porcentaje" ? `${item.descuentoValor}% dcto.` : `dcto. de ${fmt(item.descuentoValor)}`})`}</span>
          )}
        </div>
        {cots.length < 3 && <button onClick={addCot} className="text-xs text-indigo-600 flex items-center gap-1"><Plus size={12} /> Cotización</button>}
      </div>
      <div className="space-y-3">
        {cots.map((c, i) => { const d = desgloseCotizacion(c, item.cantidad); return (
          <div key={i} className="space-y-2 border-b border-slate-200 pb-3 last:border-0 last:pb-0">
            <div className="grid grid-cols-8 gap-1.5 items-end">
              <div className="col-span-2 flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Proveedor</label>
                <AutocompletarProveedor
                  proveedores={proveedores}
                  valorTexto={c.proveedorId ? (proveedores.find((p) => p.id === c.proveedorId)?.nombre || "") : (c.proveedorNombre || "")}
                  proveedorId={c.proveedorId}
                  onElegir={(p) => { update(i, "proveedorId", p.id); update(i, "proveedorNombre", ""); }}
                  onEscribir={(texto) => { update(i, "proveedorNombre", texto); update(i, "proveedorId", ""); }}
                />
              </div>
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Unidad cotizada</label>
                <select value={c.unidadCotizada} onChange={(e) => update(i, "unidadCotizada", e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">{UNIDADES.map((u) => <option key={u} value={u}>{u}</option>)}</select>
              </div>
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">1 {c.unidadCotizada} equivale a</label>
                <div className="flex items-center gap-1">
                  <input type="number" value={c.factorConversion} onChange={(e) => update(i, "factorConversion", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                  <span className="text-[10px] text-slate-400 whitespace-nowrap">{item.unidad}</span>
                </div>
              </div>
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Moneda</label>
                <select value={c.moneda || "COP"} onChange={(e) => cambiarMoneda(i, e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">{MONEDAS.map((m) => <option key={m} value={m}>{m}</option>)}</select>
              </div>
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Precio inicial</label>
                <InputMiles value={c.precioUnitario} onChange={(v) => update(i, "precioUnitario", v)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
              </div>
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Precio final neg.</label>
                <InputMiles value={c.precioFinal} onChange={(v) => update(i, "precioFinal", v)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
              </div>
              {!sinIva && (
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">IVA</label>
                <select value={c.ivaPct} onChange={(e) => update(i, "ivaPct", e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">{IVA_OPCIONES.map((v) => <option key={v} value={v}>IVA {v}%</option>)}</select>
              </div>
              )}
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Días entrega</label>
                <div className="flex gap-1"><input type="number" value={c.diasEntrega} onChange={(e) => update(i, "diasEntrega", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs" />{cots.length > 1 && <button onClick={() => removeCot(i)} className="text-slate-400 hover:text-rose-500 shrink-0"><Trash2 size={13} /></button>}</div>
              </div>
            </div>

            {erroresPrecioCotizacion(item, c).map((e, k) => (
              <div key={k} className="text-[11px] text-rose-600 bg-rose-50 border border-rose-200 rounded-md px-2 py-1.5">
                ⚠ {e.texto}{e.tipo === "estimado" && !compacto ? " Si el precio real es mayor, devuelve la solicitud (Rechazar) con el motivo para que el solicitante actualice el precio estimado." : ""}
              </div>
            ))}
            {!c.proveedorId && c.proveedorNombre?.trim() && (
              <div className="flex gap-2 max-w-xs">
                <div className="flex flex-col gap-0.5 w-1/2">
                  <label className="text-[10px] text-slate-400">NIT / RUT (opcional)</label>
                  <input type="text" value={c.proveedorNitNuevo || ""} onChange={(e) => update(i, "proveedorNitNuevo", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                </div>
                <div className="flex flex-col gap-0.5 w-1/2">
                  <label className="text-[10px] text-slate-400">Correo del proveedor (opcional)</label>
                  <input type="email" value={c.proveedorEmailNuevo || ""} onChange={(e) => update(i, "proveedorEmailNuevo", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                </div>
              </div>
            )}

            <div className="grid grid-cols-8 gap-1.5 items-end">
              {c.moneda && c.moneda !== "COP" && (
                <div className="col-span-2 flex flex-col gap-0.5">
                  <label className="text-[10px] text-slate-400" title={`¿Cuántos COP equivalen a 1 ${c.moneda}?`}>Tasa {c.moneda}→COP</label>
                  <div className="flex items-center gap-1">
                    <input type="number" value={c.tasaCambio} onChange={(e) => update(i, "tasaCambio", e.target.value)} className="flex-1 border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                    <button type="button" onClick={() => actualizarTasaAutomatica(i, c.moneda)} disabled={cargandoTasa === i} title="Actualizar tasa del día" className="text-slate-400 hover:text-indigo-600 disabled:opacity-50 shrink-0">{cargandoTasa === i ? "..." : "↻"}</button>
                  </div>
                </div>
              )}
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Tipo descuento</label>
                <select value={c.descuentoTipo || "porcentaje"} onChange={(e) => update(i, "descuentoTipo", e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">
                  <option value="porcentaje">Desc. %</option>
                  <option value="valor">Desc. $</option>
                </select>
              </div>
              <div className="col-span-2 flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">{c.descuentoTipo === "valor" ? "Descuento en $" : "Descuento en %"}</label>
                {c.descuentoTipo === "valor" ? (
                  <InputMiles value={c.descuentoValor} onChange={(v) => update(i, "descuentoValor", v)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                ) : (
                  <input type="number" value={c.descuentoValor} onChange={(e) => update(i, "descuentoValor", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                )}
              </div>
            </div>

            {sinIva && !compacto && (
              <div className="grid grid-cols-3 gap-2">
                <div className="flex flex-col gap-0.5">
                  <label className="text-[10px] text-slate-400">Admón. % (este proveedor)</label>
                  <input type="number" min="0" max="100" step="0.1" placeholder="0" value={c.aiu?.administracionPct || ""} onChange={(e) => updateAiu(i, "administracionPct", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                </div>
                <div className="flex flex-col gap-0.5">
                  <label className="text-[10px] text-slate-400">Utilidad %</label>
                  <input type="number" min="0" max="100" step="0.1" placeholder="0" value={c.aiu?.utilidadPct || ""} onChange={(e) => updateAiu(i, "utilidadPct", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                </div>
                <div className="flex flex-col gap-0.5">
                  <label className="text-[10px] text-slate-400">Imprevistos %</label>
                  <input type="number" min="0" max="100" step="0.1" placeholder="0" value={c.aiu?.imprevistosPct || ""} onChange={(e) => updateAiu(i, "imprevistosPct", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
                </div>
              </div>
            )}

            {sinIva && c.precioUnitario && (() => {
              const aiuC = c.aiu || {};
              const admC = d.subtotal * (parseFloat(aiuC.administracionPct) || 0) / 100;
              const utilC = d.subtotal * (parseFloat(aiuC.utilidadPct) || 0) / 100;
              const imprevC = d.subtotal * (parseFloat(aiuC.imprevistosPct) || 0) / 100;
              const totalCotizacion = compacto ? d.subtotal : d.subtotal + admC + utilC + imprevC + utilC * 0.19;
              return <PlanPagoCotizacion pagos={c.pagos} total={totalCotizacion} onChange={(pagos) => update(i, "pagos", pagos)} />;
            })()}

            <div className="flex items-start justify-between gap-3 flex-wrap">
              <AdjuntarArchivo small nombre={c.archivoNombre} label="Adjuntar cotización (PDF/foto)" onSeleccionar={(n) => update(i, "archivoNombre", n)} />
              {(c.proveedorId || c.proveedorNombre) && c.precioUnitario && (() => {
                const factor = parseFloat(c.factorConversion) || 1;
                const precioPorUnidad = precioEquivalente(c);
                const aiuC = c.aiu || {};
                const admC = d.subtotal * (parseFloat(aiuC.administracionPct) || 0) / 100;
                const utilC = d.subtotal * (parseFloat(aiuC.utilidadPct) || 0) / 100;
                const imprevC = d.subtotal * (parseFloat(aiuC.imprevistosPct) || 0) / 100;
                const ivaUtilC = utilC * 0.19;
                const totalConAiu = d.subtotal + admC + utilC + imprevC + ivaUtilC;
                return (
                  <div className="bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2 text-xs text-slate-600 space-y-0.5">
                    {factor !== 1 && <div>1 {c.unidadCotizada} = {factor} {item.unidad} → {fmt(precioFinalEfectivo(c) * (c.moneda && c.moneda !== "COP" ? (parseFloat(c.tasaCambio) || 1) : 1))} ÷ {factor} = <b>{fmt(precioPorUnidad)}</b> por {item.unidad}</div>}
                    <div>{item.cantidad} {item.unidad} × {fmt(precioPorUnidad)} = Costo Directo <b>{fmt(d.subtotal)}</b></div>
                    {!sinIva && <div>+ IVA {c.ivaPct}%: <b>{fmt(d.iva)}</b></div>}
                    {sinIva && !compacto && (admC > 0 || utilC > 0 || imprevC > 0) && <div>+ AIU: <b>{fmt(admC + utilC + imprevC + ivaUtilC)}</b></div>}
                    <div className="text-emerald-700 font-semibold text-sm pt-0.5 border-t border-emerald-200 mt-1">= Total: {fmt(sinIva ? (compacto ? d.subtotal : totalConAiu) : d.total)}</div>
                  </div>
                );
              })()}
            </div>
          </div>
        );})}
        {!cots.length && <div className="text-[11px] text-slate-400">Sin cotizaciones aún. Usa "+ Cotización" para agregar hasta 3.</div>}
      </div>
      {cots.some((c) => c.precioUnitario && !(c.proveedorId || (c.proveedorNombre || "").trim())) && (
        <div className="text-[11px] text-rose-600 mt-2">⚠ Toda cotización debe tener el nombre del proveedor — las filas sin proveedor no se guardan.</div>
      )}
      {(() => {
        const validas = cots.filter((c) => (c.proveedorId || c.proveedorNombre) && c.precioUnitario);
        if (validas.length !== 1) return null;
        return (
          <div className="mt-2 bg-amber-50 border border-amber-200 rounded-lg p-2.5">
            <label className="text-[11px] font-medium text-amber-700">Solo hay una cotización para este ítem — explica por qué (obligatorio)</label>
            <textarea value={validas[0].justificacionUnico || ""} onChange={(e) => update(cots.indexOf(validas[0]), "justificacionUnico", e.target.value)} rows={2} className="w-full mt-1 border border-amber-200 rounded-md px-2 py-1.5 text-xs resize-none bg-white" placeholder="Ej. proveedor único en la zona, marca exclusiva, urgencia..." />
          </div>
        );
      })()}
      <div className="text-[11px] text-slate-400 mt-1">Factor = a cuántas {item.unidad} equivale 1 unidad cotizada por el proveedor. El precio final negociado (si existe) es el que se usa para calcular el total.</div>
      <div className="flex items-center gap-2 mt-2">
        <button onClick={guardar} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium">Guardar cotizaciones</button>
        {guardadoMsg && <span className="text-[11px] text-emerald-600 flex items-center gap-1"><CheckCircle2 size={12} /> Cotizaciones guardadas correctamente</span>}
      </div>
    </div>
  );
}

function ComparativoTabla({ item, numero, proveedores, onSeleccionar, seleccionada, soloLectura, sinIva, tinte, ocultarAiu, abierto = true, onToggle }) {
  const scored = calcularScores(item.cotizaciones, item.cantidad, sinIva, item.aiu);
  const bestIdx = mejorCotizacionIdx(item.cotizaciones, item.cantidad, sinIva, item.aiu);
  const elegidaIdx = seleccionada ?? bestIdx;
  const elegida = scored[elegidaIdx];
  const [pendienteIdx, setPendienteIdx] = useState(null);
  const [obsTemp, setObsTemp] = useState("");
  const nombreProv = (c) => proveedores.find((p) => p.id === c.proveedorId)?.nombre || c.proveedorNombre || "—";
  const tieneAiu = (aiu) => !!aiu && (parseFloat(aiu.administracionPct) || parseFloat(aiu.utilidadPct) || parseFloat(aiu.imprevistosPct));
  // de dónde sale el AIU que se está aplicando: de esta cotización puntual, o del ítem en general (respaldo)
  const origenAiu = (c) => (tieneAiu(c.aiu) ? "cotizacion" : tieneAiu(item.aiu) ? "item" : null);
  const aiuAplicado = (c) => (tieneAiu(c.aiu) ? c.aiu : item.aiu || {});
  const detalleAiu = (c) => { const a = aiuAplicado(c); return `Admón. ${a.administracionPct || 0}% · Utilidad ${a.utilidadPct || 0}% · Imprevistos ${a.imprevistosPct || 0}%`; };
  // total incluyendo el AIU propio de esa cotización (cada proveedor puede tener % distintos); si la
  // cotización no tiene AIU puesto, usa el del ítem como respaldo — igual que en el cálculo real
  const totalConAiu = (c) => {
    const aiu = aiuAplicado(c);
    const adm = c.subtotal * (parseFloat(aiu.administracionPct) || 0) / 100;
    const util = c.subtotal * (parseFloat(aiu.utilidadPct) || 0) / 100;
    const imprev = c.subtotal * (parseFloat(aiu.imprevistosPct) || 0) / 100;
    return c.subtotal + adm + util + imprev + util * 0.19;
  };

  const clickElegir = (i) => {
    if (soloLectura) return;
    if (i === bestIdx) { onSeleccionar(item.id, i, ""); setPendienteIdx(null); return; }
    setPendienteIdx(i); setObsTemp(item.observacionSeleccion || "");
  };
  const confirmarNoSugerido = () => { if (!obsTemp.trim()) return; onSeleccionar(item.id, pendienteIdx, obsTemp.trim()); setPendienteIdx(null); };

  return (
    <div className={`border ${tinte?.borde || "border-slate-200"} rounded-lg overflow-hidden`}>
      <div className={`px-3 py-2 ${tinte?.fondo || "bg-slate-50"} text-sm font-medium text-slate-700 flex items-center justify-between gap-2`}>
        <button type="button" onClick={onToggle} aria-expanded={abierto} className="flex items-center gap-2 text-left min-w-0">
          {onToggle && <ChevronRight size={16} className={`shrink-0 text-slate-400 transition-transform ${abierto ? "rotate-90" : ""}`} />}
          <span className="truncate">{numero ? `${numero}. ` : ""}{item.nombre} — {item.cantidad} {item.unidad}</span>
        </button>
        <span className="flex items-center gap-2 shrink-0 text-xs font-normal text-slate-500">
          {!abierto && elegida && <span>{nombreProv(elegida)} · {fmt(sinIva ? (ocultarAiu ? elegida.subtotal : totalConAiu(elegida)) : elegida.total)}</span>}
          {soloLectura && <span className="text-[11px] text-slate-400 flex items-center gap-1"><Lock size={11} /> Bloqueado (orden ya generada)</span>}
        </span>
      </div>
      <div className={abierto ? "" : "hidden"}>
      <table className="w-full text-xs">
        <thead className="bg-white text-slate-500 border-b border-slate-100"><tr><th className="text-left px-3 py-2">Proveedor</th><th className="text-right px-3 py-2">Precio inicial</th><th className="text-right px-3 py-2">Descuento</th><th className="text-right px-3 py-2">Precio final</th><th className="text-right px-3 py-2">Cant.</th><th className="text-right px-3 py-2">{sinIva ? "Costo Directo" : "Total (COP)"}</th>{sinIva && !ocultarAiu && <th className="text-right px-3 py-2">Total c/AIU</th>}<th className="text-right px-3 py-2">Entrega</th><th className="text-right px-3 py-2">Score</th><th className="px-3 py-2"></th></tr></thead>
        <tbody>{scored.map((c, i) => (
          <tr key={i} className={`border-t border-slate-100 ${i === bestIdx ? "bg-emerald-50/60" : ""}`}>
            <td className="px-3 py-2 font-medium text-slate-700 flex items-center gap-1">{i === bestIdx && <Award size={13} className="text-emerald-600" />} {nombreProv(c)} {c.archivoNombre && <EnlacePrivado path={c.archivoNombre} className="text-slate-400 hover:text-indigo-600" title="Ver cotización adjunta"><Paperclip size={11} /></EnlacePrivado>}</td>
            <td className="px-3 py-2 text-right">{c.precioUnitario ? `${c.moneda && c.moneda !== "COP" ? c.moneda + " " : ""}${Number(c.precioUnitario).toLocaleString("es-CO")}` : "—"}</td>
            <td className="px-3 py-2 text-right">{c.descuentoValor ? (c.descuentoTipo === "valor" ? `-${Number(c.descuentoValor).toLocaleString("es-CO")}` : `-${c.descuentoValor}%`) : "—"}</td>
            <td className="px-3 py-2 text-right">{c.moneda && c.moneda !== "COP" ? `${c.moneda} ${precioFinalEfectivo(c).toLocaleString("es-CO")}` : fmt(precioFinalEfectivo(c))}</td>
            <td className="px-3 py-2 text-right text-slate-400">× {item.cantidad}</td>
            <td className="px-3 py-2 text-right font-medium">{fmt(sinIva ? c.subtotal : c.total)}</td>
            {sinIva && !ocultarAiu && (
              <td className="px-3 py-2 text-right">
                <div className="font-semibold text-slate-700">{fmt(totalConAiu(c))}</div>
                {origenAiu(c) && (
                  <div title={detalleAiu(c)} className={`inline-block mt-0.5 text-[9px] px-1.5 py-0.5 rounded-full font-medium cursor-help ${origenAiu(c) === "cotizacion" ? "bg-indigo-100 text-indigo-700" : "bg-slate-100 text-slate-500"}`}>
                    AIU de {origenAiu(c) === "cotizacion" ? "esta cotización" : "el ítem"}
                  </div>
                )}
              </td>
            )}
            <td className="px-3 py-2 text-right">{c.diasEntrega} días</td>
            <td className="px-3 py-2 text-right font-medium">{(c.score * 100).toFixed(0)}%</td>
            <td className="px-3 py-2 text-right"><button disabled={soloLectura} onClick={() => clickElegir(i)} className={`text-[11px] px-2 py-1 rounded-md border font-medium disabled:opacity-40 ${elegidaIdx === i ? "bg-indigo-600 text-white border-indigo-600" : "border-slate-200 text-slate-600"}`}>{elegidaIdx === i ? "Seleccionada" : "Elegir"}</button></td>
          </tr>
        ))}</tbody>
      </table>
      {pendienteIdx !== null && (
        <div className="px-3 py-2 bg-amber-50 border-t border-amber-200">
          <div className="text-[11px] text-amber-700 mb-1 flex items-center gap-1"><ShieldCheck size={12} /> Estás eligiendo un proveedor distinto al sugerido por el sistema. Justifica esta decisión (obligatorio):</div>
          <textarea value={obsTemp} onChange={(e) => setObsTemp(e.target.value)} rows={2} className="w-full border border-amber-300 rounded-md px-2 py-1.5 text-xs resize-none" placeholder="Ej. mejor plazo de pago, relación histórica con el proveedor, disponibilidad inmediata..." />
          <div className="flex gap-2 justify-end mt-1">
            <button onClick={() => setPendienteIdx(null)} className="text-[11px] text-slate-500">Cancelar</button>
            <button onClick={confirmarNoSugerido} disabled={!obsTemp.trim()} className="text-[11px] bg-amber-600 text-white px-2 py-1 rounded-md disabled:opacity-40">Confirmar selección</button>
          </div>
        </div>
      )}
      {elegida && (
        <div className="px-3 py-2 bg-slate-50 border-t border-slate-100 text-xs text-slate-600 flex justify-end gap-4">
          {sinIva && ocultarAiu ? <span>Costo Directo: <b>{fmt(elegida.subtotal)}</b></span> : sinIva ? <><span>Costo Directo: <b>{fmt(elegida.subtotal)}</b></span><span>Total con AIU: <b>{fmt(totalConAiu(elegida))}</b> {origenAiu(elegida) && <span className="text-[10px] text-slate-400">({detalleAiu(elegida)} — de {origenAiu(elegida) === "cotizacion" ? "esta cotización" : "el ítem"})</span>}</span></> : (<><span>Subtotal: <b>{fmt(elegida.subtotal)}</b></span><span>IVA: <b>{fmt(elegida.iva)}</b></span><span>Total: <b>{fmt(elegida.total)}</b></span></>)}
        </div>
      )}
      {elegida?.pagos && (
        <div className="px-3 py-2 border-t border-slate-100 text-[11px] text-slate-600 bg-indigo-50/30">
          <span className="font-medium">Plan de pagos propio de {nombreProv(elegida)}: </span>
          {elegida.pagos.tipoPago === "contado"
            ? `Pago único ${fmt(elegida.pagos.pagoUnico?.valor)} — ${elegida.pagos.pagoUnico?.fecha || "sin fecha"}`
            : `Anticipo ${fmt(elegida.pagos.anticipo?.valor)} (${elegida.pagos.anticipo?.fecha || "sin fecha"})${elegida.pagos.intermedio?.activo ? `, Intermedio ${fmt(elegida.pagos.intermedio.valor)} (${elegida.pagos.intermedio.fecha || "sin fecha"})` : ""}, Final ${fmt(elegida.pagos.final?.valor)} (${elegida.pagos.final?.fecha || "sin fecha"})`}
        </div>
      )}
      {item.cotizaciones.length === 1 && item.cotizaciones[0].justificacionUnico && <div className="px-3 py-2 border-t border-slate-100 text-[11px] text-slate-600 bg-slate-50">Cotización única — motivo: "{item.cotizaciones[0].justificacionUnico}"</div>}
      {item.observacionSeleccion && <div className="px-3 py-2 border-t border-slate-100 text-[11px] text-amber-700 italic bg-amber-50/50">Justificación de selección no sugerida: "{item.observacionSeleccion}"</div>}
      </div>
    </div>
  );
}

// Compras (u otro rol con permiso) valida y ajusta los % de AIU que el solicitante estimó,
// una vez el proveedor ya entregó su cotización real — solo aplica a órdenes de servicio/trabajo.
function AiuEditor({ solicitud, onGuardarItems, editable, proveedores = [] }) {
  const ab = useAbiertosItems(solicitud.items);
  const setItemAiu = (itemId, campo, val) => {
    onGuardarItems(solicitud.items.map((it) => (it.id === itemId ? { ...it, aiu: { ...(it.aiu || {}), [campo]: pctValido(val) } } : it)));
  };
  const d = desgloseSolicitud(solicitud);

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4">
      <div className="font-medium text-slate-700 flex items-center gap-2"><DollarSign size={16} /> Costos indirectos (AIU) por ítem {!editable && <span className="text-[11px] text-slate-400 font-normal">(solo lectura)</span>}</div>
      <div className="text-[11px] text-slate-400">Cada ítem tiene su propio AIU, independiente del proveedor que finalmente se adjudique. Si el ítem ya tiene una cotización seleccionada, el AIU de esa cotización tiene prioridad sobre el que se ve aquí.</div>
      <div className="space-y-3">
        <ControlExpandirTodo n={solicitud.items.length} onTodos={ab.todos} />
        {solicitud.items.map((it, idx) => {
          const aiu = it.aiu || {};
          const tieneAiuCot = (a) => !!a && (parseFloat(a.administracionPct) || parseFloat(a.utilidadPct) || parseFloat(a.imprevistosPct));
          let cotConAiu = null;
          if (it.cotizaciones?.length) {
            const sel = idxCotizacionActiva(it, true);
            const cot = it.cotizaciones[sel];
            if (tieneAiuCot(cot?.aiu)) cotConAiu = cot;
          }
          const aiuMostrado = cotConAiu ? cotConAiu.aiu : aiu;
          const bloqueado = !editable || !!cotConAiu; // si el AIU viene de la cotización, se edita allá, no aquí
          return (
            <ItemColapsable key={it.id} abierto={ab.abierto(it.id)} onToggle={() => ab.alternar(it.id)} numero={idx + 1} titulo={it.nombre} tinte={tinteItem(idx)} resumen={<span>AIU {aiuMostrado.administracionPct || 0} / {aiuMostrado.utilidadPct || 0} / {aiuMostrado.imprevistosPct || 0} %</span>}>
            <div className="p-1">
              {cotConAiu ? (
                <div className="text-[11px] text-indigo-600 bg-indigo-50 rounded-md px-2 py-1.5 mb-1">
                  AIU definido en la cotización de <b>{proveedores.find((pv) => pv.id === cotConAiu.proveedorId)?.nombre || cotConAiu.proveedorNombre || "el proveedor"}</b>{editable ? " — para cambiarlo, edítalo en el formulario de cotización." : "."}
                </div>
              ) : null}
              <div className="grid grid-cols-3 gap-2">
                <div><label className="text-[10px] text-slate-400 block mb-0.5">Admón. %</label><input disabled={bloqueado} type="number" min="0" max="100" step="0.1" value={aiuMostrado.administracionPct || ""} onChange={(e) => setItemAiu(it.id, "administracionPct", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /></div>
                <div><label className="text-[10px] text-slate-400 block mb-0.5">Utilidad %</label><input disabled={bloqueado} type="number" min="0" max="100" step="0.1" value={aiuMostrado.utilidadPct || ""} onChange={(e) => setItemAiu(it.id, "utilidadPct", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /></div>
                <div><label className="text-[10px] text-slate-400 block mb-0.5">Imprevistos %</label><input disabled={bloqueado} type="number" min="0" max="100" step="0.1" value={aiuMostrado.imprevistosPct || ""} onChange={(e) => setItemAiu(it.id, "imprevistosPct", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /></div>
              </div>
            </div>
            </ItemColapsable>
          );
        })}
      </div>
      <div className="text-[11px] text-slate-400">El IVA (19%) se calcula automáticamente solo sobre la Utilidad.</div>
      <div className="space-y-0.5 text-xs text-slate-500 max-w-xs">
        <div className="flex justify-between"><span>Costo Directo</span><span>{fmt(d.costoDirecto)}</span></div>
        <div className="flex justify-between"><span>Administración</span><span>{fmt(d.administracion)}</span></div>
        <div className="flex justify-between"><span>Utilidad</span><span>{fmt(d.utilidad)}</span></div>
        <div className="flex justify-between"><span>Imprevistos</span><span>{fmt(d.imprevistos)}</span></div>
        <div className="flex justify-between"><span>IVA sobre Utilidad</span><span>{fmt(d.ivaUtilidad)}</span></div>
        <div className="flex justify-between font-medium text-slate-700 border-t border-slate-100 pt-1 mt-1"><span>Total</span><span>{fmt(d.total)}</span></div>
      </div>
    </div>
  );
}

function ResumenTotales({ solicitud }) {
  const d = desgloseSolicitud(solicitud);
  // los costos indirectos (AIU) aparecen solo cuando ya hay cotizaciones cargadas
  if (solicitud.tipo === "servicio" && !aiuVisible(solicitud)) {
    return (
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-3 text-sm">Totales de la solicitud</div>
        <div className="flex flex-col items-end gap-1 text-sm max-w-xs ml-auto">
          <div className="flex justify-between w-full"><span className="text-slate-500">Costo Directo</span><span className="text-slate-700">{fmt(d.costoDirecto)}</span></div>
          <div className="text-[11px] text-slate-400 w-full text-right">Los costos indirectos (AIU) se calculan cuando Compras ingrese las cotizaciones.</div>
        </div>
      </div>
    );
  }
  if (solicitud.tipo === "servicio") {
    return (
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-3 text-sm">Totales de la solicitud (Costo Directo + AIU)</div>
        <div className="flex flex-col items-end gap-1 text-sm max-w-xs ml-auto">
          <div className="flex justify-between w-full"><span className="text-slate-500">Costo Directo</span><span className="text-slate-700">{fmt(d.costoDirecto)}</span></div>
          <div className="flex justify-between w-full"><span className="text-slate-500">Administración ({solicitud.aiu?.administracionPct || 0}%)</span><span className="text-slate-700">{fmt(d.administracion)}</span></div>
          <div className="flex justify-between w-full"><span className="text-slate-500">Utilidad ({solicitud.aiu?.utilidadPct || 0}%)</span><span className="text-slate-700">{fmt(d.utilidad)}</span></div>
          <div className="flex justify-between w-full"><span className="text-slate-500">Imprevistos ({solicitud.aiu?.imprevistosPct || 0}%)</span><span className="text-slate-700">{fmt(d.imprevistos)}</span></div>
          <div className="flex justify-between w-full"><span className="text-slate-500">IVA sobre la Utilidad (19%)</span><span className="text-slate-700">{fmt(d.ivaUtilidad)}</span></div>
          <div className="flex justify-between w-full border-t border-slate-200 pt-1 mt-1"><span className="font-medium text-slate-800">Total solicitud</span><span className="font-semibold text-slate-900">{fmt(d.total)}</span></div>
        </div>
      </div>
    );
  }
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="font-medium text-slate-700 mb-3 text-sm">Totales de la solicitud</div>
      <div className="flex flex-col items-end gap-1 text-sm max-w-xs ml-auto">
        <div className="flex justify-between w-full"><span className="text-slate-500">Subtotal</span><span className="text-slate-700">{fmt(d.subtotal)}</span></div>
        <div className="flex justify-between w-full"><span className="text-slate-500">Total IVA</span><span className="text-slate-700">{fmt(d.iva)}</span></div>
        <div className="flex justify-between w-full border-t border-slate-200 pt-1 mt-1"><span className="font-medium text-slate-800">Total solicitud</span><span className="font-semibold text-slate-900">{fmt(d.total)}</span></div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   REVISIÓN DE COMPRAS (solo compra) — histórico + aprobar/rechazar/modificar
--------------------------------------------------------- */
function RevisionCompras({ solicitud, historico, setHistorico, currentUser, onGuardarItems, onDecision }) {
  const [observacion, setObservacion] = useState("");
  const esCompras = currentUser.rol === "Compras" || currentUser.rol === "Administrador";
  if (solicitud.tipo !== "compra" || solicitud.revisionCompras.estado === "no_aplica") return null;

  if (solicitud.revisionCompras.estado !== "pendiente") {
    return (
      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-2 flex items-center gap-2"><CheckSquare size={16} /> Revisión de Compras</div>
        <Badge tone={solicitud.revisionCompras.estado === "aprobada" ? "green" : "red"}>{solicitud.revisionCompras.estado === "aprobada" ? "Aprobada" : "Rechazada"} por {solicitud.revisionCompras.usuario} · {solicitud.revisionCompras.fecha}</Badge>
        {solicitud.revisionCompras.observacion && <div className="text-xs text-slate-500 italic mt-2">"{solicitud.revisionCompras.observacion}"</div>}
      </div>
    );
  }

  if (!esCompras) return (
    <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">Pendiente de revisión por el área de Compras (validación de cantidades e histórico) antes de solicitar cotizaciones.</div>
  );

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
      <div className="font-medium text-slate-700 flex items-center gap-2"><CheckSquare size={16} /> Revisión de Compras</div>
      <div className="text-xs text-slate-400">Solo visible para Compras: histórico de precios y opción de ajustar cantidades antes de continuar.</div>
      {solicitud.items.map((it) => (
        <div key={it.id} className="border border-slate-200 rounded-lg p-2.5">
          <div className="flex items-center gap-2 mb-1">
            <span className="text-sm font-medium text-slate-700">{it.nombre}</span>
            <input type="number" value={it.cantidad} onChange={(e) => onGuardarItems(solicitud.items.map((x) => (x.id === it.id ? { ...x, cantidad: e.target.value } : x)))} className="w-20 border border-slate-200 rounded-md px-2 py-1 text-xs" />
            <span className="text-xs text-slate-400">{it.unidad}</span>
          </div>
          <HistoricoCompras nombreItem={it.nombre} historico={historico} setHistorico={setHistorico} />
        </div>
      ))}
      <div><label className="text-xs font-medium text-slate-500">Observación (obligatoria si rechaza)</label><textarea value={observacion} onChange={(e) => setObservacion(e.target.value)} rows={2} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" /></div>
      <div className="flex gap-2 justify-end">
        <button onClick={() => observacion.trim() && onDecision("rechazada", observacion)} className="px-3 py-1.5 rounded-lg text-xs text-rose-600 border border-rose-200">Rechazar</button>
        <button onClick={() => onDecision("aprobada", observacion)} className="px-3 py-1.5 rounded-lg text-xs bg-indigo-600 text-white font-medium">Aprobar y continuar a cotización</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   PAGOS: sugeridos por solicitante + confirmados por Dirección Financiera
--------------------------------------------------------- */
/* ---------------------------------------------------------
   PLAN DE PAGOS POR ÍTEM — compras y servicios. Cada ítem tiene su propio
   plan y su propia confirmación (o se confirman todos a la vez). Quién puede
   editar y confirmar lo define el permiso "editar_pagos" (Catálogo → Permisos).
   Si el solicitante sugirió un plan para el ítem, entra como valor inicial.
--------------------------------------------------------- */
function planFaltaFecha(p) {
  return p.tipoPago === "contado" ? !p.pagoUnico.fecha : (!p.anticipo.fecha || !p.final.fecha || (p.intermedio.activo && !p.intermedio.fecha));
}
// plan con el que arranca un ítem: el oficial si ya existe; si no, lo que sugirió el solicitante (reescalado al
// total actual del ítem, por si cambió con el AIU); si tampoco hay sugerencia, vacío
function planInicialItem(s, it, totalItem) {
  const hayOficial = planTieneValores(it.pagos) || (s.items.length === 1 && planTieneValores(s.pagos));
  if (hayOficial) return planOficialItem(s, it);
  const sug = planSugeridoItem(s, it);
  return planTieneValores(sug) ? reescalarPlanPago(sug, totalItem) : planPagosVacio();
}

function ItemPlanPago({ item, numero, totalItem, planInicial, esSugerido, confirmado, opcional, editable, puedeReabrirPlan, onGuardar, onConfirmar, onEditarDeNuevo, abierto = true, onToggle }) {
  const [pagos, setPagos] = useState(planInicial);
  const editadoLocal = useRef(false);
  const claveInicial = JSON.stringify(planInicial);
  // si el solicitante cambia su sugerencia mientras nadie ha tocado el plan oficial, se refresca el valor inicial
  useEffect(() => { if (!editadoLocal.current) setPagos(planInicial); // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [item.id, esSugerido ? claveInicial : "oficial"]);
  const guardar = (copy) => { editadoLocal.current = true; setPagos(copy); onGuardar(copy); };
  const pagado = totalPagado(pagos);
  const restante = totalItem - pagado;
  const hayPlan = planTieneValores(pagos);
  const descuadrado = confirmado && Math.abs(restante) > 0.5;
  const faltaFecha = planFaltaFecha(pagos);

  const set = (campo, sub, val) => {
    if (sub === "fecha" && val) { const error = validarOrdenFechas(pagos, campo, val); if (error) { alert(error); return; } }
    guardar({ ...pagos, [campo]: { ...pagos[campo], [sub]: val } });
  };
  const setTipoPago = (tipo) => guardar({ ...pagos, tipoPago: tipo, pagoUnico: tipo === "contado" ? { ...pagos.pagoUnico, valor: totalItem } : pagos.pagoUnico });

  // si cambia el total del ítem (ej. se ajustó el AIU) el plan se reescala solo para seguir cuadrando —
  // solo lo hace quien puede editar: mirar un plan nunca debe modificar nada
  useEffect(() => {
    if (!editable || !hayPlan) return;
    if (Math.abs(totalItem - totalPagado(pagos)) < 0.5) return;
    guardar(reescalarPlanPago(pagos, totalItem));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [totalItem]);

  const confirmar = () => {
    if (!hayPlan) { alert("Este ítem todavía no tiene plan de pagos."); return; }
    if (Math.abs(restante) > 0.5) { alert(restante > 0 ? `Faltan ${fmt(restante)} por programar en este ítem.` : `El plan supera el total del ítem en ${fmt(-restante)}.`); return; }
    if (faltaFecha) { alert("Falta poner la fecha de uno o más pagos de este ítem."); return; }
    onConfirmar(pagos);
  };

  return (
    <div className={`border ${tinteItem(numero - 1).borde} ${tinteItem(numero - 1).fondo} rounded-lg p-3`}>
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <button type="button" onClick={onToggle} aria-expanded={abierto} className="flex items-center gap-2 text-sm font-medium text-slate-700 text-left min-w-0">
          <ChevronRight size={16} className={`shrink-0 text-slate-400 transition-transform ${abierto ? "rotate-90" : ""}`} />
          <span className={`inline-block w-2 h-2 rounded-full shrink-0 ${tinteItem(numero - 1).punto}`} />
          <span className="truncate">{numero}. {item.nombre}</span>
        </button>
        <div className="flex items-center gap-2 shrink-0">
          <span className="text-xs text-slate-500">{fmt(totalItem)}</span>
          {confirmado
            ? <Badge tone={descuadrado ? "red" : "green"}>Confirmado{item.pagosConfirmadosPor ? ` por ${item.pagosConfirmadosPor.nombre}` : ""}</Badge>
            : <Badge tone={opcional ? "slate" : "amber"}>{opcional ? "Sin confirmar (opcional)" : "Pendiente"}</Badge>}
        </div>
      </div>
      <div className={abierto ? "mt-2" : "hidden"}>
      {confirmado && puedeReabrirPlan && <button onClick={onEditarDeNuevo} className="text-[11px] text-indigo-600 underline mb-2 block">Editar de nuevo</button>}
      {esSugerido && !confirmado && hayPlan && <div className="text-[11px] text-indigo-600 mb-2">Valores iniciales tomados del plan que sugirió el solicitante.</div>}
      {!hayPlan && !editable ? (
        <div className="text-[11px] text-slate-400">Sin plan de pagos definido para este ítem.</div>
      ) : (
        <>
          <div className="flex gap-1.5 mb-2">
            <button type="button" disabled={!editable} onClick={() => setTipoPago("plan")} className={`px-2 py-1 rounded text-[11px] font-medium border disabled:opacity-40 ${pagos.tipoPago !== "contado" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>Por etapas</button>
            <button type="button" disabled={!editable} onClick={() => setTipoPago("contado")} className={`px-2 py-1 rounded text-[11px] font-medium border disabled:opacity-40 ${pagos.tipoPago === "contado" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>Pago único</button>
          </div>
          {pagos.tipoPago === "contado" ? (
            <div className="max-w-xs">
              <div className="text-[11px] text-slate-400 mb-1">Valor (= total del ítem)</div>
              <div className="w-full border border-slate-200 bg-slate-50 rounded-md px-2 py-1.5 text-sm mb-1 text-slate-600">{fmt(totalItem)}</div>
              <InputFecha disabled={!editable} value={pagos.pagoUnico.fecha} onChange={(v) => set("pagoUnico", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
              <div><label className="text-[10px] text-slate-400 block mb-0.5">Anticipo</label><InputMiles disabled={!editable} placeholder="Valor" value={pagos.anticipo.valor} onChange={(v) => set("anticipo", "valor", v)} className="w-full mb-1 border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /><InputFecha disabled={!editable} value={pagos.anticipo.fecha} onChange={(v) => set("anticipo", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /></div>
              <div><label className="text-[10px] text-slate-400 flex items-center gap-1 mb-0.5"><input disabled={!editable} type="checkbox" checked={pagos.intermedio.activo} onChange={(e) => set("intermedio", "activo", e.target.checked)} /> Intermedio</label><InputMiles disabled={!editable || !pagos.intermedio.activo} placeholder="Valor" value={pagos.intermedio.valor} onChange={(v) => set("intermedio", "valor", v)} className="w-full mb-1 border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /><InputFecha disabled={!editable || !pagos.intermedio.activo} value={pagos.intermedio.fecha} onChange={(v) => set("intermedio", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /></div>
              <div><label className="text-[10px] text-slate-400 block mb-0.5">Pago final</label><InputMiles disabled={!editable} placeholder="Valor" value={pagos.final.valor} onChange={(v) => set("final", "valor", v)} className="w-full mb-1 border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /><InputFecha disabled={!editable} value={pagos.final.fecha} onChange={(v) => set("final", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" /></div>
            </div>
          )}
          <div className="flex items-center justify-between mt-2 gap-2 flex-wrap">
            <div className="text-[11px] text-slate-500">Total ítem: <b>{fmt(totalItem)}</b> · Programado: <b>{fmt(pagado)}</b> · Restante: <b className={restante > 0.5 ? "text-amber-600" : restante < -0.5 ? "text-rose-600" : "text-emerald-600"}>{fmt(restante)}</b></div>
            {editable && !confirmado && <button onClick={confirmar} className="text-[11px] bg-emerald-600 text-white px-2 py-1 rounded-md font-medium">Confirmar este ítem</button>}
          </div>
        </>
      )}
      </div>
    </div>
  );
}

function PagosPorItem({ solicitud, currentUser, onGuardarItems }) {
  const ab = useAbiertosItems(solicitud.items);
  const sinIva = solicitud.tipo === "servicio";
  const opcional = solicitud.tipo === "compra"; // en solicitudes de compra el plan no bloquea la aprobación
  const ocYaEnviada = ["orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status);
  const sinPrecio = !(totalSolicitud(solicitud) > 0);
  const tienePermiso = puedeEditarPagos(currentUser);
  const puedeEditarRol = tienePermiso && !sinPrecio;

  const filas = solicitud.items.map((it) => {
    const totalItem = totalItemConAiu(it, sinIva);
    const conf = planConfirmadoItem(solicitud, it);
    const inicial = planInicialItem(solicitud, it, totalItem);
    const hayOficial = planTieneValores(it.pagos) || (solicitud.items.length === 1 && planTieneValores(solicitud.pagos));
    return { it, totalItem, conf, inicial, esSugerido: !hayOficial && planTieneValores(inicial), descuadrado: conf && Math.abs(totalItem - totalPagado(planOficialItem(solicitud, it))) > 0.5 };
  });
  // una vez generada la orden el plan queda fijo, salvo un plan confirmado que no cuadra (válvula de corrección)
  const puedeTocar = (f) => puedeEditarRol && (!ocYaEnviada || f.descuadrado);
  const conPlan = filas.filter((f) => f.conf || planTieneValores(f.inicial));
  const todosConfirmados = opcional ? (conPlan.length > 0 && conPlan.every((f) => f.conf)) : filas.every((f) => f.conf);
  const hayAlgoPorConfirmar = filas.some((f) => !f.conf && (!opcional || planTieneValores(f.inicial)));

  const guardarItem = (itemId, pagos) => onGuardarItems(solicitud.items.map((it) => (it.id === itemId ? { ...it, pagos } : it)));
  const confirmarItem = (itemId, pagos) => onGuardarItems(solicitud.items.map((it) => (it.id === itemId ? { ...it, pagos, pagosConfirmados: true, pagosConfirmadosPor: { nombre: currentUser.nombre, rol: currentUser.rol } } : it)));
  const editarDeNuevoItem = (itemId) => onGuardarItems(solicitud.items.map((it) => (it.id === itemId ? { ...it, pagosConfirmados: false } : it)));

  const confirmarTodos = () => {
    const porConfirmar = filas.filter((f) => !f.conf && (!opcional || planTieneValores(f.inicial)));
    const problema = porConfirmar.find((f) => !planTieneValores(f.inicial) || Math.abs(f.totalItem - totalPagado(f.inicial)) > 0.5 || planFaltaFecha(f.inicial));
    if (problema) { alert(`El ítem "${problema.it.nombre}" todavía no está listo para confirmar (revisa el valor y las fechas).`); return; }
    if (!porConfirmar.length) return;
    onGuardarItems(solicitud.items.map((it) => {
      const f = porConfirmar.find((x) => x.it.id === it.id);
      return f ? { ...it, pagos: f.inicial, pagosConfirmados: true, pagosConfirmadosPor: { nombre: currentUser.nombre, rol: currentUser.rol } } : it;
    }));
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="font-medium text-slate-700 flex items-center gap-2"><CalendarClock size={16} /> Plan de pagos por ítem</div>
        {todosConfirmados ? <Badge tone="green">Todos los ítems confirmados</Badge> : (puedeEditarRol && !ocYaEnviada && hayAlgoPorConfirmar && <button onClick={confirmarTodos} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium">Confirmar todos los ítems</button>)}
      </div>
      {opcional && <div className="text-[11px] text-slate-400">En las solicitudes de compra el plan de pagos es opcional: no bloquea la aprobación.</div>}
      {ocYaEnviada && <div className="text-[11px] text-slate-400 bg-slate-50 border border-slate-200 rounded-md px-3 py-2">🔒 La orden ya fue generada — las condiciones de pago quedaron fijas.</div>}
      {sinPrecio && !ocYaEnviada && <div className="text-[11px] text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-3 py-2">⚠ Todavía no hay precios cargados — se habilita en cuanto Compras cotice.</div>}
      {!tienePermiso && !ocYaEnviada && !sinPrecio && <div className="text-[11px] text-slate-400">Tu rol no tiene el permiso para editar y confirmar el plan de pagos. El administrador puede activarlo en Catálogo → Permisos.</div>}
      <ControlExpandirTodo n={solicitud.items.length} onTodos={ab.todos} />
      <div className="space-y-3">
        {filas.map((f, idx) => (
          <ItemPlanPago
            key={f.it.id}
            abierto={ab.abierto(f.it.id)}
            onToggle={() => ab.alternar(f.it.id)}
            item={f.it}
            numero={idx + 1}
            totalItem={f.totalItem}
            planInicial={f.inicial}
            esSugerido={f.esSugerido}
            confirmado={f.conf}
            opcional={opcional}
            editable={puedeTocar(f) && !f.conf}
            puedeReabrirPlan={puedeTocar(f)}
            onGuardar={(pagos) => guardarItem(f.it.id, pagos)}
            onConfirmar={(pagos) => confirmarItem(f.it.id, pagos)}
            onEditarDeNuevo={() => editarDeNuevoItem(f.it.id)}
          />
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   ORDEN ENVIADA AL PROVEEDOR
--------------------------------------------------------- */
function OcEnviadaPanel({ solicitud, proveedores, empresa, currentUser, onGuardar }) {
  const [firmandoIdx, setFirmandoIdx] = useState(null);
  const [generandoIdx, setGenerandoIdx] = useState(null);
  if (solicitud.status !== "orden") return null;
  // gestionar/generar la orden es tarea de Compras; firmarla es tarea de Dirección Financiera —
  // ambos roles necesitan ver este panel (cada quien solo puede tocar lo que le corresponde)
  if (!puedeGestionarCotizaciones(currentUser) && !puedeAprobarFinanciera(currentUser)) return null;
  const puedeGenerar = puedeGestionarCotizaciones(currentUser);

  const necesarios = proveedoresAdjudicadosDetalle(solicitud, proveedores);
  const ordenes = necesarios.map((n) => {
    const existente = (solicitud.ocEnviada.ordenesProveedor || []).find((o) => mismoProveedor(o, n));
    return existente || { ...n, archivoOriginalUrl: "", archivoFirmadoUrl: "", fecha: "", usuario: "" };
  });

  const actualizarOrden = (idx, cambios) => {
    const copia = ordenes.map((o, i) => (i === idx ? { ...o, ...cambios } : o));
    onGuardar({ ordenesProveedor: copia });
  };

  // el sistema contable (Zeus) no genera órdenes de servicio/trabajo, así que para ese tipo
  // el propio sistema arma el PDF a partir de los ítems adjudicados a este proveedor.
  const generarOrdenAutomatica = async (idx) => {
    const orden = ordenes[idx];
    if (orden.archivoFirmadoUrl) return; // ya firmada — no se puede regenerar (invalidaría la firma)
    setGenerandoIdx(idx);
    try {
      const itemsDelProveedor = solicitud.items.filter((it) => {
        if (!it.cotizaciones.length) return false;
        const sel = idxCotizacionActiva(it, true);
        const cot = it.cotizaciones[sel];
        return cot && mismoProveedor({ proveedorId: cot.proveedorId, proveedorNombre: cot.proveedorNombre }, orden);
      });
      const baseItems = itemsDelProveedor.length ? itemsDelProveedor : solicitud.items;
      const itemsParaPdf = baseItems.map((it) => { const d = desgloseItem(it, true); const unitario = parseFloat(it.cantidad) > 0 ? d.subtotal / parseFloat(it.cantidad) : 0; return { nombre: it.nombre, cantidad: it.cantidad, unidad: it.unidad, valorUnitario: unitario, total: d.subtotal }; });
      // usa el AIU de la cotización de este proveedor específicamente (cada uno puede tener % distintos)
      let costoDirecto = 0, administracion = 0, utilidad = 0, imprevistos = 0;
      baseItems.forEach((it) => {
        const dIt = desgloseItem(it, true);
        costoDirecto += dIt.subtotal;
        let aiu = it.aiu || solicitud.aiu || {};
        if (it.cotizaciones?.length) {
          const sel = idxCotizacionActiva(it, true);
          const cot = it.cotizaciones[sel];
          if (cot?.aiu && (parseFloat(cot.aiu.administracionPct) || parseFloat(cot.aiu.utilidadPct) || parseFloat(cot.aiu.imprevistosPct))) aiu = cot.aiu;
        }
        administracion += dIt.subtotal * (parseFloat(aiu.administracionPct) || 0) / 100;
        utilidad += dIt.subtotal * (parseFloat(aiu.utilidadPct) || 0) / 100;
        imprevistos += dIt.subtotal * (parseFloat(aiu.imprevistosPct) || 0) / 100;
      });
      const ivaUtilidad = utilidad * 0.19;
      const total = costoDirecto + administracion + utilidad + imprevistos + ivaUtilidad;
      const primerCot = baseItems[0]?.cotizaciones?.[idxCotizacionActiva(baseItems[0], true)];
      const aiuPcts = primerCot?.aiu || solicitud.aiu || {};
      const totales = { costoDirecto, administracion, utilidad, imprevistos, ivaUtilidad, total, aiuPcts };
      const planesPago = baseItems.filter((it) => planConfirmadoItem(solicitud, it)).map((it) => ({ nombre: it.nombre, pagos: planOficialItem(solicitud, it) })).filter((pl) => planTieneValores(pl.pagos));
      const bytes = await generarOrdenServicioPDF({ solicitud, empresa, proveedorNombre: orden.proveedorNombre, items: itemsParaPdf, planesPago, ...totales });
      const ruta = await subirBytes(bytes, `Orden_Servicio_${solicitud.folio}_${orden.proveedorNombre.replace(/[^a-zA-Z0-9]/g, "_")}.pdf`, "ordenes-originales");
      if (ruta) actualizarOrden(idx, { archivoOriginalUrl: ruta, archivoFirmadoUrl: "", fecha: "", usuario: "", cargadaPor: currentUser.nombre, cargadaEn: ahoraISO() });
      else alert("No se pudo generar el documento. Intenta de nuevo.");
    } catch (e) {
      console.error("Error generando la orden de servicio:", e);
      alert("No se pudo generar el documento.");
    }
    setGenerandoIdx(null);
  };

  const firmarOrden = async (idx) => {
    const orden = ordenes[idx];
    if (!orden.archivoOriginalUrl) return;
    setFirmandoIdx(idx);
    try {
      const urlOriginalFirmada = await obtenerUrlFirmada(orden.archivoOriginalUrl);
      const urlFirmaFotoFirmada = currentUser.firmaFotoUrl ? await obtenerUrlFirmada(currentUser.firmaFotoUrl) : null;
      if (!urlOriginalFirmada) throw new Error("No se pudo acceder al documento original.");
      const blob = await firmarPDF(urlOriginalFirmada, urlFirmaFotoFirmada, currentUser.nombre, currentUser.cargo, empresa?.nombre);
      const archivo = new File([blob], `OC_${solicitud.folio}_${orden.proveedorNombre.replace(/[^a-zA-Z0-9]/g, "_")}.pdf`, { type: "application/pdf" });
      const ruta = await subirArchivo(archivo, "ordenes-firmadas");
      if (ruta) actualizarOrden(idx, { archivoFirmadoUrl: ruta, fecha: hoy(), usuario: currentUser.nombre });
      else alert("No se pudo guardar el documento firmado. Intenta de nuevo.");
    } catch (e) {
      console.error("Error firmando el PDF:", e);
      alert("No se pudo firmar el documento. Verifica que el archivo cargado sea un PDF válido (no una imagen).");
    }
    setFirmandoIdx(null);
  };

  if (!necesarios.length) return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 text-xs text-amber-600">No se detectó ningún proveedor adjudicado todavía — revisa el cuadro comparativo antes de continuar.</div>
  );

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4">
      <div className="font-medium text-slate-700 flex items-center gap-2"><Send size={16} /> Firma y envío de la orden al proveedor</div>
      <div className="text-xs text-slate-500">Se detectaron <b>{ordenes.length}</b> proveedor(es) adjudicado(s) en esta solicitud. {solicitud.tipo === "servicio" ? <>Como el sistema contable no genera órdenes de servicio/trabajo, el propio sistema arma el documento por ti.</> : <>Sube aquí la orden generada en el sistema contable de cada uno.</>} El documento se envía a <b>Dirección Financiera</b> para su firma digital. Al marcar como "Enviada al proveedor", el documento ya firmado se envía automáticamente por correo a cada proveedor <b>que tenga correo registrado en el Catálogo</b> — si alguno no lo tiene, tendrás que enviársela tú manualmente.</div>

      {ordenes.map((o, i) => (
        <div key={i} className="border border-slate-200 rounded-lg p-3 space-y-2">
          <div className="text-sm font-medium text-slate-700 flex items-center gap-2"><Truck size={13} /> {o.proveedorNombre}</div>
          {solicitud.tipo === "servicio" ? (
            <div className="flex items-center gap-2 flex-wrap">
              {o.archivoFirmadoUrl ? (
                <span className="text-[11px] text-slate-400 flex items-center gap-1"><Lock size={12} /> Ya firmada — no se puede volver a generar.</span>
              ) : puedeGenerar ? (
                <>
                  <button onClick={() => generarOrdenAutomatica(i)} disabled={generandoIdx === i} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-50 flex items-center gap-1"><FileText size={12} /> {generandoIdx === i ? "Generando..." : o.archivoOriginalUrl ? "Volver a generar la orden" : "Generar orden de servicio (automático)"}</button>
                </>
              ) : (
                <span className="text-[11px] text-amber-600">Pendiente de que Compras genere el documento.</span>
              )}
            </div>
          ) : puedeGenerar ? (
            <AdjuntarArchivo nombre={o.archivoOriginalUrl} label={`Adjuntar solicitud de compra para ${o.proveedorNombre} (solo PDF)`} onSeleccionar={(url) => actualizarOrden(i, { archivoOriginalUrl: url, archivoFirmadoUrl: "", fecha: "", usuario: "", cargadaPor: currentUser.nombre, cargadaEn: ahoraISO() })} carpeta="ordenes-originales" soloPdf />
          ) : (
            !o.archivoOriginalUrl && <span className="text-[11px] text-amber-600">Pendiente de que Compras suba la orden del sistema contable.</span>
          )}
          {o.archivoOriginalUrl && !o.archivoFirmadoUrl && (
            <div className="bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2 text-xs text-emerald-700 flex items-start gap-2">
              <CheckCircle2 size={15} className="mt-0.5 shrink-0" />
              <div>
                <div className="font-medium">Transacción registrada: {solicitud.tipo === "servicio" ? "orden de servicio generada" : "solicitud de compra cargada"} correctamente{o.cargadaPor ? ` por ${o.cargadaPor}` : ""}{o.cargadaEn ? ` el ${new Date(o.cargadaEn).toLocaleString("es-CO", { dateStyle: "short", timeStyle: "short" })}` : ""}.</div>
                <div className="text-emerald-600">Queda pendiente de la firma de Dirección Financiera. <EnlacePrivado path={o.archivoOriginalUrl} className="underline">Ver documento</EnlacePrivado></div>
              </div>
            </div>
          )}
          {o.archivoOriginalUrl && (
            o.archivoFirmadoUrl ? (
              <div className="text-xs text-emerald-700 flex items-center gap-2"><CheckCircle2 size={13} /> Firmada por {o.usuario} el {o.fecha}. <EnlacePrivado path={o.archivoFirmadoUrl} className="underline">Ver PDF firmado</EnlacePrivado></div>
            ) : puedeAprobarFinanciera(currentUser) ? (
              <div>
                <div className="text-[11px] text-slate-500 mb-1">{currentUser.firmaFotoUrl ? "Se estampará tu firma guardada en \"Mi perfil\" en cada hoja del documento, junto a tu nombre, cargo y empresa." : "Sin foto de firma guardada — se firmará cada hoja solo con nombre, cargo y empresa."}</div>
                <button onClick={() => firmarOrden(i)} disabled={firmandoIdx === i} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-50 flex items-center gap-1"><PenTool size={12} /> {firmandoIdx === i ? "Firmando..." : "Firmar como Dirección Financiera"}</button>
              </div>
            ) : (
              <div className="text-xs text-amber-600">Pendiente de firma de Dirección Financiera.</div>
            )
          )}
        </div>
      ))}
    </div>
  );
}

// permite a Compras reenviar por correo una orden ya firmada (ej. si el proveedor la perdió o no llegó)
function ReenviarOrdenesPanel({ solicitud, proveedores, guardarProveedor, empresa, currentUser }) {
  const [seleccionadas, setSeleccionadas] = useState([]);
  const [correosManual, setCorreosManual] = useState({}); // { idx: { c1: "", c2: "" } }
  const [enviando, setEnviando] = useState(false);
  const [mensaje, setMensaje] = useState("");
  if (!puedeGestionarCotizaciones(currentUser)) return null;
  if (!["oc_enviada", "recepcion", "completada"].includes(solicitud.status)) return null;

  const ordenesFirmadas = (solicitud.ocEnviada.ordenesProveedor || []).filter((o) => o.archivoFirmadoUrl);
  if (!ordenesFirmadas.length) return null;

  const alternar = (idx) => setSeleccionadas((prev) => prev.includes(idx) ? prev.filter((x) => x !== idx) : [...prev, idx]);
  const setManual = (idx, campo, val) => setCorreosManual((prev) => ({ ...prev, [idx]: { ...prev[idx], [campo]: val } }));

  const reenviar = async () => {
    setEnviando(true);
    let enviados = 0, sinCorreo = 0;
    for (const idx of seleccionadas) {
      const orden = ordenesFirmadas[idx];
      const prov = buscarProveedorDeOrden(orden, proveedores);
      const manual1 = (correosManual[idx]?.c1 || "").trim();
      const manual2 = (correosManual[idx]?.c2 || "").trim();
      const correos = correosDe(prov).length ? correosDe(prov) : [manual1, manual2].filter(Boolean);
      if (!correos.length) { sinCorreo++; continue; }
      const url = await obtenerUrlFirmada(orden.archivoFirmadoUrl, 604800);
      if (url) {
        await enviarCorreo(
          correos,
          `Reenvío — Solicitud de compra / orden de servicio ${solicitud.folio}`,
          `<p>Hola ${prov?.nombre || orden.proveedorNombre},</p><p>Te reenviamos el enlace de la solicitud de compra / orden de servicio <b>${solicitud.folio}</b> a nombre de ${empresa?.nombre || ""}.</p><p><a href="${url}">Ver / descargar la orden firmada</a></p><p>Este enlace estará disponible por 7 días.</p>`
        );
        enviados++;
        // si los correos se escribieron a mano y el proveedor existe sin correo, los guardamos para la próxima vez
        if (!correosDe(prov).length && guardarProveedor && prov) {
          await guardarProveedor({ ...prov, email: manual1, email2: manual2 });
        }
      }
    }
    setEnviando(false);
    setSeleccionadas([]);
    setMensaje(`${enviados} correo(s) reenviado(s)${sinCorreo ? `. ${sinCorreo} sin correo (escríbelo abajo para poder enviarlo).` : "."}`);
    setTimeout(() => setMensaje(""), 5000);
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-2">
      <div className="font-medium text-slate-700 flex items-center gap-2"><Send size={16} /> Reenviar orden(es) firmada(s) al proveedor</div>
      <div className="text-xs text-slate-400">Solo visible para Compras. Útil si el proveedor no recibió el correo o lo perdió.</div>
      {ordenesFirmadas.map((o, idx) => {
        const prov = buscarProveedorDeOrden(o, proveedores);
        return (
          <div key={idx} className="flex items-center gap-2 text-sm text-slate-700 border border-slate-200 rounded-md px-3 py-2">
            <input type="checkbox" checked={seleccionadas.includes(idx)} onChange={() => alternar(idx)} />
            <span>{o.proveedorNombre}</span>
            {correosDe(prov).length ? (
              <span className="text-[11px] text-slate-400 ml-auto">{correosDe(prov).join(" · ")}</span>
            ) : (
              <div className="ml-auto flex gap-1.5">
                <input
                  type="email"
                  placeholder="Correo del proveedor"
                  value={correosManual[idx]?.c1 || ""}
                  onChange={(e) => setManual(idx, "c1", e.target.value)}
                  onClick={(e) => e.stopPropagation()}
                  className="border border-slate-200 rounded-md px-2 py-1 text-xs w-48"
                />
                <input
                  type="email"
                  placeholder="Correo adicional (opcional)"
                  value={correosManual[idx]?.c2 || ""}
                  onChange={(e) => setManual(idx, "c2", e.target.value)}
                  onClick={(e) => e.stopPropagation()}
                  className="border border-slate-200 rounded-md px-2 py-1 text-xs w-48"
                />
              </div>
            )}
          </div>
        );
      })}
      <button onClick={reenviar} disabled={!seleccionadas.length || enviando} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40">{enviando ? "Enviando..." : `Reenviar (${seleccionadas.length})`}</button>
      {mensaje && <div className="text-[11px] text-emerald-600">{mensaje}</div>}
    </div>
  );
}

/* ---------------------------------------------------------
   RECEPCIÓN
--------------------------------------------------------- */
/* ---------------------------------------------------------
   EVALUACIÓN POSTERIOR A LA RECEPCIÓN (obligatoria para completar)
--------------------------------------------------------- */
function CalificacionSelect({ value, onChange, disabled }) {
  const val = value || 1;
  const pct = ((val - 1) / 9) * 100;
  return (
    <div className="flex items-center gap-2 w-full max-w-[190px] shrink-0">
      <input
        type="range"
        min="1"
        max="10"
        step="1"
        value={val}
        disabled={disabled}
        onChange={(e) => onChange(Number(e.target.value))}
        title={value ? `Calificación: ${value}` : "Desliza para calificar"}
        className="flex-1 h-1.5 rounded-full appearance-none cursor-pointer disabled:cursor-not-allowed accent-indigo-600"
        style={{ background: `linear-gradient(to right, #4f46e5 ${value ? pct : 0}%, #e2e8f0 ${value ? pct : 0}%)` }}
      />
      <span className={`text-xs font-semibold w-4 text-center shrink-0 ${value ? "text-indigo-600" : "text-slate-300"}`}>{value || "—"}</span>
    </div>
  );
}

// genera y descarga el Excel del "Registro Selección y Evaluación de Proveedores" con los datos ya diligenciados
function descargarExcelEvaluacion(ev, solicitud) {
  const filas = [];
  filas.push(["REGISTRO SELECCIÓN Y EVALUACIÓN DE PROVEEDORES"]);
  filas.push([]);
  filas.push(["PROVEEDOR (Razón Social):", ev.proveedorNombre]);
  filas.push(["TIPO DE PROVEEDOR:", ev.tipoProveedor === "compra" ? "Compra" : ev.tipoProveedor === "servicio" ? "Servicio" : "Trabajo"]);
  filas.push(["FECHA DE SELECCIÓN:", ev.fechaSeleccion, "NIT:", ev.nit, "CC:", ev.cc]);
  filas.push(["FECHA DE EVALUACIÓN:", ev.fechaEvaluacion]);
  filas.push(["CIUDAD:", ev.ciudad, "DIRECCIÓN:", ev.direccion]);
  filas.push(["REPRESENTANTE LEGAL:", ev.representanteLegal, "TELÉFONO:", ev.telefono]);
  filas.push(["FAX:", ev.fax, "EMAIL:", ev.email]);
  filas.push([]);
  filas.push(["SERVICIOS QUE PRESTA", ev.serviciosPresta]);
  filas.push(["DESCRIPCIÓN", ev.descripcion]);
  filas.push(["MARCA", ev.marca]);
  filas.push([]);
  filas.push(["FAVOR ANEXAR LOS SIGUIENTES DOCUMENTOS", "SI", "NO", "NO APLICA"]);
  DOCUMENTOS_EVALUACION.forEach((d) => {
    const v = ev.documentos?.[d.key];
    filas.push([d.label, v === "si" ? "X" : "", v === "no" ? "X" : "", v === "no_aplica" ? "X" : ""]);
  });
  filas.push([]);
  filas.push(["COMPONENTE", "SUBCOMPONENTE", "CRITERIO", "CALIFICACIÓN (1-10)", "PONDERACIÓN", "%"]);
  CRITERIOS_EVALUACION.forEach((c) => {
    const cal = parseFloat(ev.criterios?.[c.key]) || 0;
    filas.push([c.componente, c.subcomponente, c.texto, cal, c.peso, cal ? ((cal / 10) * c.peso) : 0]);
  });
  const pct = puntajeEvaluacion(ev.criterios) * 100;
  const clas = clasificacionConfianza(pct);
  filas.push([]);
  filas.push(["RESULTADO (%)", pct.toFixed(1)]);
  filas.push(["CLASIFICACIÓN", clas.texto]);
  filas.push(["Rangos:", "Poco Confiable ≤ 50%", "Medio Confiable 51%-79%", "Confiable ≥ 80%"]);
  filas.push([]);
  filas.push(["OBSERVACIONES:", ev.observaciones || ""]);
  filas.push([]);
  filas.push(["Aprobado por (cargo):", ev.aprobadoPorCargo || ""]);
  filas.push(["Realizado por:", ev.firmaRealizada?.nombre || ""]);
  filas.push(["Cargo:", ev.firmaRealizada?.cargo || ""]);
  filas.push(["Empresa:", ev.firmaRealizada?.empresa || ""]);
  filas.push(["Fecha firma:", ev.firmaRealizada?.fecha || ""]);
  filas.push(["Solicitud:", solicitud.folio]);

  const hoja = XLSX.utils.aoa_to_sheet(filas);
  hoja["!cols"] = [{ wch: 32 }, { wch: 30 }, { wch: 45 }, { wch: 16 }, { wch: 12 }, { wch: 10 }];
  const libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja, "Evaluación");
  XLSX.writeFile(libro, `Evaluacion_Proveedor_${(ev.proveedorNombre || "").replace(/[^a-zA-Z0-9]/g, "_")}_${solicitud.folio}.xlsx`);
}

// permite corregir/agregar el plan de pagos sugerido después de reabrir una solicitud rechazada
// (el mismo Sí/No y campos que existen al crearla, pero editable desde el detalle)
function EvaluacionPanel({ solicitud, empresa, proveedores, currentUser, onGuardar }) {
  const [reevaluando, setReevaluando] = useState(false);
  const [ev, setEv] = useState({ ...evaluacionProveedorVacia(), ...solicitud.evaluacionProveedor });
  // si la solicitud cambia por fuera (ej. otra persona la actualizó), se resincroniza el estado local
  useEffect(() => { setEv({ ...evaluacionProveedorVacia(), ...solicitud.evaluacionProveedor }); }, [solicitud.id, solicitud.evaluacionProveedor?.completada, solicitud.evaluacionProveedor?.fechaCompletado]);

  const esCompras = currentUser.rol === "Compras" || currentUser.rol === "Administrador";
  const yaFinalizada = solicitud.status !== "recepcion";

  // primera vez que se abre: se pre-llena automáticamente con los datos del proveedor adjudicado y de la solicitud
  const proveedorSugerido = proveedores.find((p) => p.id === ev.proveedorId) || proveedores.find((p) => proveedoresAdjudicadosDetalle(solicitud, proveedores).some((n) => mismoProveedor({ proveedorId: p.id }, n)));
  // escribe local al instante (fluido) y guarda en segundo plano, sin bloquear la escritura
  const set = (campo, val) => { const copy = { ...ev, [campo]: val }; setEv(copy); onGuardar(copy); };
  const setDoc = (key, val) => { const copy = { ...ev, documentos: { ...ev.documentos, [key]: val } }; setEv(copy); onGuardar(copy); };
  const setCriterio = (key, val) => { const copy = { ...ev, criterios: { ...ev.criterios, [key]: val } }; setEv(copy); onGuardar(copy); };

  // rellena todos los campos con los datos de un proveedor del catálogo — se usa tanto en la
  // detección automática como cuando Compras lo elige manualmente (respaldo si la detección falla)
  const aplicarProveedor = (proveedor) => {
    const fechaComparativo = solicitud.historialEstados?.find((h) => h.status === "comparativo")?.fecha;
    const copy = {
      ...ev,
      proveedorId: proveedor.id,
      proveedorNombre: proveedor.nombre,
      tipoProveedor: proveedor.tipoProveedor || solicitud.tipo,
      nit: proveedor.nit || "",
      ciudad: proveedor.ciudad || "",
      direccion: proveedor.direccion || "",
      telefono: proveedor.telefono || "",
      representanteLegal: proveedor.representanteLegal || "",
      email: proveedor.email || "",
      fechaSeleccion: ev.fechaSeleccion || (fechaComparativo ? fechaComparativo.slice(0, 10) : solicitud.fechaCreacion) || "",
      fechaEvaluacion: ev.fechaEvaluacion || hoy(),
      serviciosPresta: ev.serviciosPresta || proveedor.actividadEconomica || "",
      descripcion: ev.descripcion || solicitud.items.map((it) => it.nombre).join(", "),
    };
    setEv(copy);
    onGuardar(copy);
  };

  useEffect(() => {
    if (esCompras && !ev.completada && !ev.proveedorId && proveedorSugerido) aplicarProveedor(proveedorSugerido);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [proveedorSugerido?.id, esCompras]);

  // la evaluación es responsabilidad exclusiva de Compras — nadie más la ve, ni en solo lectura
  if (!esCompras) return null;

  const pct = puntajeEvaluacion(ev.criterios) * 100;
  const clas = clasificacionConfianza(pct);
  // ya completada (solicitud finalizada) y sin haber pedido reevaluar → solo lectura
  const disabled = !esCompras || (yaFinalizada && !reevaluando);
  const esRegistroViejoVacio = yaFinalizada && !ev.completada; // solicitudes completadas antes de este formato, sin diligenciar

  const guardarEvaluacion = () => {
    onGuardar({
      ...ev,
      completada: true,
      fechaCompletado: hoy(),
      fechaEvaluacion: ev.fechaEvaluacion || hoy(),
      firmaRealizada: { nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresa?.nombre || "", fecha: hoy(), fotoUrl: currentUser.firmaFotoUrl || null },
    });
    setReevaluando(false);
  };

  const grupos = [];
  CRITERIOS_EVALUACION.forEach((c) => {
    let g = grupos.find((x) => x.componente === c.componente && x.subcomponente === c.subcomponente);
    if (!g) { g = { componente: c.componente, subcomponente: c.subcomponente, items: [] }; grupos.push(g); }
    g.items.push(c);
  });

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="font-medium text-slate-700 flex items-center gap-2">
          <Award size={16} /> Evaluación de proveedor {!yaFinalizada && <span className="text-[11px] text-amber-600 font-normal">(obligatoria para completar — la hace Compras)</span>}
          {yaFinalizada && ev.completada && !reevaluando && <Badge tone="green">Diligenciada</Badge>}
        </div>
        <div className="flex items-center gap-2">
          {ev.completada && <button onClick={() => descargarExcelEvaluacion(ev, solicitud)} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><FileText size={13} /> Descargar Excel</button>}
          {yaFinalizada && esCompras && !reevaluando && <button onClick={() => setReevaluando(true)} className="text-xs text-indigo-600 underline">{ev.completada ? "Reevaluar" : "Diligenciar ahora"}</button>}
        </div>
      </div>

      {esRegistroViejoVacio && !reevaluando && (
        <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
          Esta solicitud se completó antes de este formato de evaluación, así que quedó sin diligenciar. Usa "Diligenciar ahora" para llenarla con la información disponible.
        </div>
      )}

      {!disabled && (
        <div>
          <label className="text-xs text-slate-500 block mb-1">Proveedor a evaluar (elígelo si no se detectó solo)</label>
          <select disabled={disabled} value={ev.proveedorId || ""} onChange={(e) => { const p = proveedores.find((x) => x.id === e.target.value); if (p) aplicarProveedor(p); }} className="w-full max-w-sm border border-slate-200 rounded-md px-2 py-1.5 text-sm">
            <option value="">— Elegir del catálogo —</option>
            {proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
          </select>
        </div>
      )}

      {/* DATOS DEL PROVEEDOR */}
      <div className="border border-slate-200 rounded-lg p-3 space-y-2">
        <div className="text-sm font-medium text-slate-700">Datos del proveedor</div>
        <div className="grid grid-cols-2 md:grid-cols-3 gap-2 text-xs">
          <div><label className="text-slate-500 block mb-0.5">Razón social</label><input disabled={disabled} value={ev.proveedorNombre} onChange={(e) => set("proveedorNombre", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Tipo de proveedor</label><select disabled={disabled} value={ev.tipoProveedor} onChange={(e) => set("tipoProveedor", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50"><option value="">—</option><option value="compra">Compra</option><option value="servicio">Servicio</option><option value="trabajo">Trabajo</option></select></div>
          <div><label className="text-slate-500 block mb-0.5">NIT</label><input disabled={disabled} value={ev.nit} onChange={(e) => set("nit", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Fecha de selección</label><InputFecha disabled={disabled} value={ev.fechaSeleccion} onChange={(v) => set("fechaSeleccion", v)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Fecha de evaluación</label><InputFecha disabled={disabled} value={ev.fechaEvaluacion} onChange={(v) => set("fechaEvaluacion", v)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Ciudad</label><input disabled={disabled} value={ev.ciudad} onChange={(e) => set("ciudad", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div className="col-span-2"><label className="text-slate-500 block mb-0.5">Dirección</label><input disabled={disabled} value={ev.direccion} onChange={(e) => set("direccion", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Representante legal</label><input disabled={disabled} value={ev.representanteLegal} onChange={(e) => set("representanteLegal", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Teléfono</label><input disabled={disabled} value={ev.telefono} onChange={(e) => set("telefono", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div><label className="text-slate-500 block mb-0.5">Email</label><input disabled={disabled} value={ev.email} onChange={(e) => set("email", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
          <div className="col-span-2 md:col-span-3"><label className="text-slate-500 block mb-0.5">Servicios que presta / descripción</label><input disabled={disabled} value={ev.serviciosPresta} onChange={(e) => set("serviciosPresta", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50" /></div>
        </div>
      </div>

      {/* DOCUMENTOS */}
      <div className="border border-slate-200 rounded-lg p-3 space-y-2">
        <div className="text-sm font-medium text-slate-700 mb-1">Documentos anexos</div>
        {DOCUMENTOS_EVALUACION.map((d) => (
          <div key={d.key} className="flex items-center justify-between text-xs gap-2 flex-wrap">
            <span className="text-slate-600 flex-1">{d.label}</span>
            <div className="flex gap-1.5 shrink-0">
              {[{ v: "si", l: "Sí" }, { v: "no", l: "No" }, { v: "no_aplica", l: "No aplica" }].map((op) => (
                <button
                  key={op.v}
                  type="button"
                  disabled={disabled}
                  onClick={() => setDoc(d.key, ev.documentos?.[d.key] === op.v ? "" : op.v)}
                  className={`px-2 py-1 rounded-md text-[11px] font-medium border disabled:opacity-40 disabled:cursor-not-allowed ${
                    ev.documentos?.[d.key] === op.v ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200 hover:border-indigo-300"
                  }`}
                >
                  {op.l}
                </button>
              ))}
            </div>
          </div>
        ))}
      </div>

      {/* CRITERIOS PONDERADOS */}
      <div className="border border-slate-200 rounded-lg p-3 space-y-3">
        <div className="text-sm font-medium text-slate-700">Criterios de selección y evaluación (1-10)</div>
        {grupos.map((g, gi) => (
          <div key={gi}>
            <div className="text-[11px] font-medium text-slate-500 mb-1">{g.componente} — {g.subcomponente}</div>
            <div className="space-y-1">
              {g.items.map((c) => (
                <div key={c.key} className="flex items-center justify-between gap-2 text-xs">
                  <span className="text-slate-600 flex-1">{c.texto} <span className="text-slate-400">(peso {(c.peso * 100).toFixed(0)}%)</span></span>
                  <CalificacionSelect value={ev.criterios?.[c.key]} onChange={(v) => setCriterio(c.key, v)} disabled={disabled} />
                </div>
              ))}
            </div>
          </div>
        ))}
        <div className="border-t border-slate-200 pt-2 flex items-center justify-between">
          <div className="text-sm text-slate-600">Resultado: <b className="text-slate-800">{pct.toFixed(1)}%</b></div>
          <Badge tone={clas.tone}>{clas.texto}</Badge>
        </div>
        <div className="text-[10px] text-slate-400">Poco Confiable ≤ 50% · Medio Confiable 51%-79% · Confiable ≥ 80%</div>
      </div>

      <div>
        <label className="text-xs text-slate-500 block mb-1">Observaciones (opcional)</label>
        <textarea value={ev.observaciones} onChange={(e) => set("observaciones", e.target.value)} disabled={disabled} rows={2} className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none disabled:bg-slate-50" />
      </div>

      <div>
        <label className="text-xs text-slate-500 block mb-0.5">Aprobado por (cargo)</label>
        <input disabled={disabled} value={ev.aprobadoPorCargo} onChange={(e) => set("aprobadoPorCargo", e.target.value)} className="w-full max-w-xs border border-slate-200 rounded-md px-2 py-1 text-xs disabled:bg-slate-50" />
      </div>

      <div>
        <div className="text-xs text-slate-500 mb-1">Firma de quien realiza la evaluación</div>
        <FirmaBlock rol="evaluación" firma={ev.firmaRealizada} />
        {!disabled && <div className="text-[10px] text-slate-400 mt-1">Se registra automáticamente con tu nombre, cargo y firma guardada en "Mi perfil" al guardar.</div>}
      </div>

      {!disabled && !evaluacionProveedorCompleta(ev) && <div className="text-[11px] text-amber-600">Falta calificar todos los criterios (14) para poder {yaFinalizada ? "guardar" : "completar la solicitud"}.</div>}
      {!disabled && evaluacionProveedorCompleta(ev) && (
        <div className="flex items-center gap-2">
          <button onClick={guardarEvaluacion} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium">Guardar evaluación</button>
          {reevaluando && <button onClick={() => setReevaluando(false)} className="text-xs text-slate-500">Cancelar</button>}
        </div>
      )}
    </div>
  );
}


function RecepcionPanel({ solicitud, currentUser, usuarios, onGuardar, crearNotificacion }) {
  const [r, setR] = useState({ ...solicitud.recepcion, archivos: solicitud.recepcion.archivos || (solicitud.recepcion.archivoNombre ? [solicitud.recepcion.archivoNombre] : []) });
  const [enviado, setEnviado] = useState(false);
  const set = (fields) => { const copy = { ...r, ...fields, usuario: currentUser.nombre, fecha: hoy() }; setR(copy); onGuardar(copy); setEnviado(false); };
  const agregarArchivo = (url) => set({ archivos: [...r.archivos, url] });
  const quitarArchivo = (i) => set({ archivos: r.archivos.filter((_, idx) => idx !== i) });
  // estadoRecepcion: null (pendiente) | "satisfaccion" | "observaciones" — cualquiera de las dos últimas cuenta
  // como "recibido" para el resto del flujo (avanzar, evaluación); solo cambia si quedó con observaciones o no.
  const estado = r.recibidoSatisfaccion ? (r.tipoRecepcion || "satisfaccion") : null;
  const elegir = (tipo) => set({ recibidoSatisfaccion: true, tipoRecepcion: tipo });

  const enviarACompras = () => {
    const equipoCompras = usuarios.filter((u) => ["Compras", "Administrador"].includes(u.rol) && u.email);
    if (equipoCompras.length) {
      enviarCorreo(
        equipoCompras.map((u) => u.email),
        `Recepción con observaciones: ${solicitud.folio}`,
        `<p>Hola,</p><p>${currentUser.nombre} registró la recepción de la solicitud <b>${solicitud.folio}</b> con observaciones:</p><p><i>"${r.comentario}"</i></p><p>Queda pendiente de revisión.</p>`
      );
    }
    usuarios.filter((u) => ["Compras", "Administrador"].includes(u.rol)).forEach((u) => crearNotificacion?.(u.id, `${currentUser.nombre} registró la recepción de ${solicitud.folio} con observaciones: "${r.comentario}"`, solicitud.id));
    setEnviado(true);
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
      <div className="font-medium text-slate-700 flex items-center gap-2">
        <PackageCheck size={16} /> Recepción
        {estado === "satisfaccion" && <Badge tone="green">Recibida a satisfacción</Badge>}
        {estado === "observaciones" && <Badge tone="amber">Recibida con observaciones</Badge>}
        {!estado && <Badge tone="amber">Pendiente de recepción</Badge>}
      </div>
      <div>
        <label className="text-xs font-medium text-slate-500 mb-1 block">Soportes de recepción (puedes adjuntar varios)</label>
        <div className="space-y-1.5">
          {r.archivos.map((url, i) => (
            <div key={i} className="flex items-center gap-2">
              <AdjuntarArchivo nombre={url} onSeleccionar={() => {}} />
              <button onClick={() => quitarArchivo(i)} className="text-slate-400 hover:text-rose-500"><Trash2 size={13} /></button>
            </div>
          ))}
          <AdjuntarArchivo nombre={null} label={r.archivos.length ? "Adjuntar otro archivo (PDF/foto)" : "Adjuntar soporte de recepción (PDF/foto)"} onSeleccionar={agregarArchivo} />
        </div>
      </div>
      <div><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><MessageSquare size={12} /> Comentarios {estado === "observaciones" ? "(describe las observaciones)" : "(opcional)"}</label><textarea value={r.comentario} onChange={(e) => set({ comentario: e.target.value })} rows={2} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" /></div>
      <div className="flex gap-4">
        <label className="flex items-center gap-2 text-sm text-slate-700"><input type="radio" name={`recepcion-${solicitud.id}`} checked={estado === "satisfaccion"} onChange={() => elegir("satisfaccion")} /> Recibido a satisfacción</label>
        <label className="flex items-center gap-2 text-sm text-slate-700"><input type="radio" name={`recepcion-${solicitud.id}`} checked={estado === "observaciones"} onChange={() => elegir("observaciones")} /> Recibido con observaciones</label>
      </div>
      {estado === "observaciones" && !r.comentario.trim() && <div className="text-[11px] text-amber-600">Describe en los comentarios cuáles fueron las observaciones.</div>}
      {estado === "observaciones" && r.comentario.trim() && (
        <div className="flex items-center gap-2">
          <button onClick={enviarACompras} className="text-xs bg-amber-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><Send size={13} /> Enviar a Compras</button>
          {enviado && <span className="text-[11px] text-emerald-600 flex items-center gap-1"><CheckCircle2 size={12} /> Enviado</span>}
        </div>
      )}
      {!estado && <div className="text-[11px] text-amber-600">Marca cómo se recibió para que Compras pueda hacer la evaluación y finalizar la solicitud.</div>}
      {estado === "satisfaccion" && <div className="text-[11px] text-emerald-600">✓ Recibida a satisfacción — falta que Compras complete la evaluación del proveedor para finalizar.</div>}
      {estado === "observaciones" && <div className="text-[11px] text-amber-600">✓ Recibida con observaciones — falta que Compras complete la evaluación del proveedor para finalizar.</div>}
    </div>
  );
}

/* ---------------------------------------------------------
   ORDEN DE COMPRA / TRABAJO — documento consolidado
--------------------------------------------------------- */
function OrdenDocumento({ solicitud, empresa, area, departamento, solicitante, proveedores, conceptosGasto }) {
  const d = desgloseSolicitud(solicitud);
  const exportarPDF = () => window.print();
  const conceptoGasto = conceptosGasto.find((c) => c.id === solicitud.conceptoGastoId);
  const nombreProv = (c) => proveedores.find((p) => p.id === c.proveedorId)?.nombre || c.proveedorNombre || "—";
  const pagoActivo = solicitud.tipo === "servicio";
  const mostrarAiu = aiuVisible(solicitud);
  // proveedor principal para el encabezado tipo "orden de trabajo" — si hay más de uno adjudicado,
  // se usa el primero y se listan los demás dentro de la tabla de ítems más abajo
  const proveedoresPrincipales = proveedoresAdjudicadosDetalle(solicitud, proveedores);
  const proveedorPrincipalCat = proveedoresPrincipales[0] ? proveedores.find((p) => p.id === proveedoresPrincipales[0].proveedorId) : null;
  const nombreProveedorPrincipal = proveedoresPrincipales.length > 1 ? proveedoresPrincipales.map((p) => p.proveedorNombre).join(" / ") : (proveedoresPrincipales[0]?.proveedorNombre || "—");
  const nitProveedorPrincipal = proveedorPrincipalCat?.nit || "—";
  const direccionProveedorPrincipal = proveedorPrincipalCat?.direccion || "—";
  const telefonoProveedorPrincipal = proveedorPrincipalCat?.telefono || "—";

  return (
    <div>
      <style>{`
        @media print {
          .print-wrapper-oculto { display: block !important; }
          body * { visibility: hidden; }
          #orden-imprimible, #orden-imprimible * { visibility: visible; }
          #orden-imprimible { position: absolute; left: 0; top: 0; width: 100%; border: none !important; }
          #orden-imprimible .no-print { display: none !important; }
          #orden-imprimible .salto-pagina { page-break-before: always; }
        }
      `}</style>
    <div id="orden-imprimible" className="bg-white border-2 border-slate-800 p-6 space-y-4 text-sm">
      {/* ENCABEZADO ESTILO CARTA MEMBRETADA */}
      <div className="flex justify-between items-start gap-4 pb-2">
        <div className="flex items-start gap-3">
          {empresa?.logoUrl && <img src={empresa.logoUrl} alt={empresa.nombre} className="h-14 max-w-[120px] object-contain" />}
          <div>
            <div className="text-lg font-bold text-slate-800 uppercase">{empresa?.nombre}</div>
            <div className="text-xs text-slate-600">Nit: {empresa?.nit || "—"}</div>
          </div>
        </div>
        <div className="border-2 border-slate-800 text-center shrink-0">
          <div className="px-3 py-1 border-b-2 border-slate-800 text-xs font-semibold uppercase">{solicitud.tipo === "compra" ? "Solicitud de Compra" : "Orden de Trabajo"}</div>
          <div className="px-3 py-1 text-sm font-bold">No {solicitud.folio}</div>
        </div>
      </div>

      {/* DATOS EN RECUADRO, ESTILO FORMULARIO */}
      <table className="w-full text-xs border-collapse border border-slate-400">
        <tbody>
          <tr>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50 w-32">FECHA:</td>
            <td className="border border-slate-400 px-2 py-1" colSpan={3}>{solicitud.fechaCreacion}</td>
          </tr>
          <tr>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50">Proveedor:</td>
            <td className="border border-slate-400 px-2 py-1">{nombreProveedorPrincipal}</td>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50 w-24">NIT:</td>
            <td className="border border-slate-400 px-2 py-1">{nitProveedorPrincipal}</td>
          </tr>
          <tr>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50">Dirección:</td>
            <td className="border border-slate-400 px-2 py-1">{direccionProveedorPrincipal}</td>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50">Teléfono:</td>
            <td className="border border-slate-400 px-2 py-1">{telefonoProveedorPrincipal}</td>
          </tr>
          <tr>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50">Oficina Solicitante:</td>
            <td className="border border-slate-400 px-2 py-1">{area?.nombre}{departamento && ` — ${departamento.nombre}`}</td>
            <td className="border border-slate-400 px-2 py-1 font-semibold bg-slate-50">Responsable:</td>
            <td className="border border-slate-400 px-2 py-1">{solicitante?.nombre}</td>
          </tr>
        </tbody>
      </table>

      <div className="border border-slate-400 p-2">
        <div className="text-xs font-semibold mb-1">DETALLE {solicitud.tipo === "compra" ? "DE LA COMPRA AUTORIZADA" : "DEL TRABAJO AUTORIZADO"}:</div>
        <div className="text-xs text-slate-700">{solicitud.objetivo}</div>
      </div>

      <div className="grid grid-cols-2 gap-2 text-xs text-slate-600">
        <div><b>Centro de costo:</b> {solicitud.centroCosto || conceptoGasto?.centroCosto || "—"}</div><div><b>Concepto de gasto:</b> {conceptoGasto ? labelConcepto(conceptoGasto) : "—"}</div>
      </div>

      {/* ÍTEMS Y PROVEEDOR ADJUDICADO */}
      <div>
        <div className="text-xs font-medium text-slate-500 mb-1">Ítems adjudicados</div>
        <table className="w-full text-xs">
          <thead className="text-slate-400 border-b border-slate-200">
            <tr>
              <th className="text-left py-1 px-2">Ítem</th>
              <th className="text-right py-1 px-2">Cant.</th>
              <th className="text-left py-1 px-2">Proveedor</th>
              {mostrarAiu ? (
                <>
                  <th className="text-right py-1 px-2">Costo Directo</th>
                  <th className="text-right py-1 px-2">AIU</th>
                  <th className="text-right py-1 px-2">Total</th>
                </>
              ) : (
                <th className="text-right py-1 px-2">Total</th>
              )}
            </tr>
          </thead>
          <tbody>{solicitud.items.map((it) => { const idx = idxCotizacionActiva(it, pagoActivo); const cot = it.cotizaciones[idx]; const dd = desgloseItem(it, pagoActivo);
            const aiuUsado = tieneAiuValores(cot?.aiu) ? cot.aiu : (it.aiu || {});
            const aiuTexto = `A${aiuUsado.administracionPct || 0} U${aiuUsado.utilidadPct || 0} I${aiuUsado.imprevistosPct || 0}`;
            return (
              <tr key={it.id} className="border-t border-slate-100">
                <td className="py-1.5 px-2">{it.nombre}</td>
                <td className="py-1.5 px-2 text-right whitespace-nowrap">{it.cantidad} {it.unidad}</td>
                <td className="py-1.5 px-2">{cot ? nombreProv(cot) : "—"}</td>
                {mostrarAiu ? (
                  <>
                    <td className="py-1.5 px-2 text-right">{fmt(dd.subtotal)}</td>
                    <td className="py-1.5 px-2 text-right whitespace-nowrap">{aiuTexto}%</td>
                    <td className="py-1.5 px-2 text-right">{fmt(totalConAiuCotizacion(dd, cot, it.aiu))}</td>
                  </>
                ) : (
                  <td className="py-1.5 px-2 text-right">{fmt(dd.total)}</td>
                )}
              </tr>
            );
          })}</tbody>
        </table>
      </div>

      {/* HISTÓRICO DE COTIZACIONES POR ÍTEM — en compra se ven las hasta 3 recibidas; en servicio
          solo el proveedor adjudicado (sin precio inicial ni score, que no aportan al documento final) */}
      {solicitud.items.some((it) => it.cotizaciones.length > 0) && (
        <div>
          <div className="text-xs font-medium text-slate-500 mb-1">{pagoActivo ? "Proveedor adjudicado por ítem" : "Histórico de cotizaciones recibidas"}</div>
          {solicitud.items.filter((it) => it.cotizaciones.length > 0).map((it) => {
            const scored = calcularScores(it.cotizaciones, it.cantidad, pagoActivo, it.aiu);
            const bestIdx = mejorCotizacionIdx(it.cotizaciones, it.cantidad, pagoActivo, it.aiu);
            const elegidaIdx = it.cotizacionSeleccionada ?? bestIdx;
            const filas = pagoActivo ? [scored[elegidaIdx]].filter(Boolean) : scored;
            return (
              <div key={it.id} className="mb-2">
                <div className="text-[11px] font-medium text-slate-600">{it.nombre} ({it.cantidad} {it.unidad})</div>
                <table className="w-full text-[11px] mb-1">
                  <thead className="text-slate-400 border-b border-slate-100">
                    <tr>
                      <th className="text-left py-0.5 px-1.5">Proveedor</th>
                      {!pagoActivo && <th className="text-right py-0.5 px-1.5">Precio inicial</th>}
                      <th className="text-right py-0.5 px-1.5">Precio final neg.</th>
                      <th className="text-right py-0.5 px-1.5">Total</th>
                      <th className="text-right py-0.5 px-1.5">Entrega</th>
                      {!pagoActivo && <th className="text-right py-0.5 px-1.5">Score</th>}
                      {!pagoActivo && <th className="text-center py-0.5 px-1.5">Elegida</th>}
                    </tr>
                  </thead>
                  <tbody>{filas.map((c, i) => (
                    <tr key={i} className="border-t border-slate-50">
                      <td className="py-0.5 px-1.5">{nombreProv(c)}</td>
                      {!pagoActivo && <td className="py-0.5 px-1.5 text-right">{fmt(c.precioUnitario)}</td>}
                      <td className="py-0.5 px-1.5 text-right">{c.precioFinal ? fmt(c.precioFinal) : "—"}</td>
                      <td className="py-0.5 px-1.5 text-right">{fmt(pagoActivo ? totalConAiuCotizacion(c, c, it.aiu) : c.total)}</td>
                      <td className="py-0.5 px-1.5 text-right">{c.diasEntrega} días</td>
                      {!pagoActivo && <td className="py-0.5 px-1.5 text-right">{(c.score * 100).toFixed(0)}%</td>}
                      {!pagoActivo && <td className="py-0.5 px-1.5 text-center">{i === elegidaIdx ? "✓" : ""}</td>}
                    </tr>
                  ))}</tbody>
                </table>
                {!pagoActivo && it.observacionSeleccion && <div className="text-[11px] text-amber-700 italic">Justificación de selección no sugerida: "{it.observacionSeleccion}"</div>}
              </div>
            );
          })}
        </div>
      )}

      {/* TOTALES */}
      {mostrarAiu ? (
        <div className="flex justify-end text-sm border-t border-slate-200 pt-2">
          <div className="text-right space-y-0.5">
            <div className="text-xs text-slate-500">Costo Directo: {fmt(d.costoDirecto)}</div>
            <div className="text-xs text-slate-500">Administración ({solicitud.aiu?.administracionPct || 0}%): {fmt(d.administracion)} · Utilidad ({solicitud.aiu?.utilidadPct || 0}%): {fmt(d.utilidad)}</div>
            <div className="text-xs text-slate-500">Imprevistos ({solicitud.aiu?.imprevistosPct || 0}%): {fmt(d.imprevistos)} · IVA sobre Utilidad (19%): {fmt(d.ivaUtilidad)}</div>
            <div><span className="text-slate-500">Total: </span><b className="text-slate-800">{fmt(d.total)}</b></div>
          </div>
        </div>
      ) : solicitud.tipo === "servicio" ? (
        <div className="flex justify-end text-sm border-t border-slate-200 pt-2">
          <div className="text-right">
            <div className="text-xs text-slate-500">Costo Directo: {fmt(d.costoDirecto)}</div>
            <div><span className="text-slate-500">Total: </span><b className="text-slate-800">{fmt(d.total)}</b></div>
          </div>
        </div>
      ) : (
        <div className="flex justify-end text-sm border-t border-slate-200 pt-2">
          <div className="text-right">
            <div className="text-xs text-slate-500">Subtotal: {fmt(d.subtotal)} · IVA: {fmt(d.iva)}</div>
            <div><span className="text-slate-500">Total: </span><b className="text-slate-800">{fmt(d.total)}</b></div>
          </div>
        </div>
      )}

      {/* SON: (monto en letras) */}
      <div className="text-xs font-semibold border-t border-b border-slate-300 py-1.5">SON: {numeroALetras(d.total)}</div>

      {/* FORMA DE PAGO Y PLAZO DE ENTREGA */}
      <table className="w-full text-xs border-collapse border border-slate-400">
        <tbody>
          <tr>
            <td className="border border-slate-400 px-2 py-2 align-top w-1/2">
              <div className="font-semibold mb-1">FORMA DE PAGO:</div>
              <div className="text-slate-600">{(() => { const tr = tramosDePagoSolicitud(solicitud).filter((t) => parseFloat(t.valor) > 0); return tr.length ? tr.map((t) => `${t.tipo}${t.itemNombre ? ` (${t.itemNombre})` : ""}: ${fmt(t.valor)} — ${t.fecha || "sin fecha"}`).join(" · ") : "Sin definir"; })()}</div>
            </td>
            <td className="border border-slate-400 px-2 py-2 align-top w-1/2">
              <div className="font-semibold mb-1">PLAZO ENTREGA {solicitud.tipo === "compra" ? "MATERIALES" : "TRABAJOS"}:</div>
              <div className="text-slate-600">{solicitud.folio}{solicitud.fechaEstimada && ` — ${solicitud.fechaEstimada}`}</div>
            </td>
          </tr>
        </tbody>
      </table>

      {/* AUTORIZA / ACEPTADO */}
      <table className="w-full text-xs border-collapse border border-slate-400 mt-6">
        <tbody>
          <tr>
            <td className="border border-slate-400 px-2 pt-8 pb-2 text-center w-1/2">
              <div className="border-t border-slate-800 pt-1 mx-4">AUTORIZA</div>
            </td>
            <td className="border border-slate-400 px-2 pt-8 pb-2 text-center w-1/2">
              <div className="border-t border-slate-800 pt-1 mx-4">ACEPTADO</div>
            </td>
          </tr>
        </tbody>
      </table>

      <div className="border border-slate-400 p-2">
        <div className="text-xs font-semibold mb-1">OBSERVACIONES:</div>
        <div className="text-xs text-slate-600 min-h-[24px]">{solicitud.recepcion?.comentario || ""}</div>
      </div>

      {/* REVISIÓN DE COMPRAS */}
      {solicitud.tipo === "compra" && solicitud.revisionCompras.estado !== "no_aplica" && (
        <div className="text-xs"><b className="text-slate-500">Revisión de Compras:</b> {solicitud.revisionCompras.estado} — {solicitud.revisionCompras.usuario || "—"} ({solicitud.revisionCompras.fecha || "—"}){solicitud.revisionCompras.observacion && ` · "${solicitud.revisionCompras.observacion}"`}</div>
      )}

      {/* FIRMAS */}
      <div className="salto-pagina">
        <div className="text-xs font-medium text-slate-500 mb-2">Historial de firmas y aprobaciones (registro digital interno)</div>
        <div className="grid grid-cols-2 gap-3 text-xs">
          {[["Solicitante", solicitud.firmas.solicitante], ["Jefe de área", solicitud.firmas.jefe], ["Director de área", solicitud.firmas.director], ["Dirección financiera", solicitud.firmas.financiera], ["Gerencia", solicitud.firmas.gerencia]].map(([rol, f]) => (
            <div key={rol} className="border border-slate-200 rounded-md p-2">
              <div className="text-[11px] text-slate-400">{rol}</div>
              {f?.nombre ? (
                <>
                  {f.fotoUrl && <ImagenPrivada path={f.fotoUrl} className="h-8 object-contain my-1" alt="" />}
                  <div className="font-medium text-slate-700">{f.nombre}</div>
                  <div className="text-[11px] text-slate-400">{f.fecha}{f.aprobado === false ? " · Rechazado" : f.aprobado ? " · Aprobado" : ""}</div>
                  {f.observacion && <div className="text-[11px] text-slate-500 italic mt-0.5">"{f.observacion}"</div>}
                </>
              ) : <div className="text-slate-400">Pendiente</div>}
            </div>
          ))}
        </div>
      </div>

      {/* PLAN DE PAGOS (por ítem) */}
      {(() => {
        const filas = solicitud.items.map((it, idx) => {
          const conf = planConfirmadoItem(solicitud, it);
          const plan = conf ? planOficialItem(solicitud, it) : planSugeridoItem(solicitud, it);
          return { it, idx, conf, tramos: tramosDePago(plan).filter((t) => parseFloat(t.valor) > 0) };
        }).filter((f) => f.tramos.length);
        if (!filas.length) return null;
        return (
          <div>
            <div className="text-xs font-medium text-slate-500 mb-1">Plan de pagos por ítem</div>
            {filas.map((f) => (
              <div key={f.it.id} className="mb-1.5">
                <div className="text-[11px] font-medium text-slate-600">{f.idx + 1}. {f.it.nombre} {f.conf ? "(confirmado)" : "(sugerido por el solicitante — sin confirmar)"}</div>
                <table className="w-full text-[11px]">
                  <tbody>{f.tramos.map((t, i) => (
                    <tr key={i} className="border-t border-slate-50"><td className="py-0.5 px-1.5">{t.tipo}</td><td className="py-0.5 px-1.5 text-right">{fmt(t.valor)}</td><td className="py-0.5 px-1.5 text-right">{t.fecha || "—"}</td><td className="py-0.5 px-1.5 text-center">{t.pagado ? "✓" : ""}</td></tr>
                  ))}</tbody>
                </table>
              </div>
            ))}
          </div>
        );
      })()}

      {/* OC ENVIADA Y RECEPCIÓN */}
      {((solicitud.ocEnviada.ordenesProveedor && solicitud.ocEnviada.ordenesProveedor.length) || (solicitud.recepcion.archivos && solicitud.recepcion.archivos.length) || solicitud.recepcion.comentario) && (
        <div className="grid grid-cols-2 gap-3 text-xs">
          <div><div className="font-medium text-slate-500 mb-0.5">Órdenes enviadas al proveedor</div>
            {(solicitud.ocEnviada.ordenesProveedor || []).length ? (solicitud.ocEnviada.ordenesProveedor || []).map((o, i) => (
              <div key={i} className="text-slate-600">{o.proveedorNombre}: {o.archivoFirmadoUrl ? `firmada por ${o.usuario} · ${o.fecha}` : "sin firmar"}</div>
            )) : <div className="text-slate-600">—</div>}
          </div>
          <div><div className="font-medium text-slate-500 mb-0.5">Recepción</div><div className="text-slate-600">{solicitud.recepcion.recibidoSatisfaccion ? (solicitud.recepcion.tipoRecepcion === "observaciones" ? "Recibido con observaciones" : "Recibido a satisfacción") : "Pendiente"}{solicitud.recepcion.archivos?.length > 0 && ` · ${solicitud.recepcion.archivos.length} archivo(s) adjunto(s)`}{solicitud.recepcion.comentario && <div className="italic">"{solicitud.recepcion.comentario}"</div>}</div></div>
        </div>
      )}

      {/* EVALUACIÓN */}
      {solicitud.evaluacionProveedor && Object.keys(solicitud.evaluacionProveedor.criterios || {}).length > 0 && (
        <div>
          <div className="text-xs font-medium text-slate-500 mb-1">Evaluación de proveedor{solicitud.evaluacionProveedor.proveedorNombre ? ` — ${solicitud.evaluacionProveedor.proveedorNombre}` : ""}</div>
          <div className="text-xs text-slate-600">
            Resultado: <b>{(puntajeEvaluacion(solicitud.evaluacionProveedor.criterios) * 100).toFixed(1)}%</b> — {clasificacionConfianza(puntajeEvaluacion(solicitud.evaluacionProveedor.criterios) * 100).texto}
            {solicitud.evaluacionProveedor.observaciones && <div className="italic mt-0.5">"{solicitud.evaluacionProveedor.observaciones}"</div>}
          </div>
        </div>
      )}

      {/* HISTORIAL DE ESTADOS / LEAD TIME */}
      <div>
        <div className="text-xs font-medium text-slate-500 mb-1">Historial del proceso (lead time)</div>
        <table className="w-full text-[11px]">
          <thead className="text-slate-400 border-b border-slate-100"><tr><th className="text-left py-0.5">Etapa</th><th className="text-left py-0.5">Fecha/hora</th><th className="text-right py-0.5">Duración</th></tr></thead>
          <tbody>{solicitud.historialEstados.map((h, i) => (
            <tr key={i} className="border-t border-slate-50"><td className="py-0.5">{PASOS.find((p) => p.key === h.status)?.label || h.status}</td><td className="py-0.5">{new Date(h.fecha).toLocaleString("es-CO")}</td><td className="py-0.5 text-right">{i === 0 ? "—" : duracion(solicitud.historialEstados[i - 1].fecha, h.fecha)}</td></tr>
          ))}</tbody>
        </table>
      </div>

      {/* NOTIFICACIONES */}
      {solicitud.notificaciones.length > 0 && (
        <div>
          <div className="text-xs font-medium text-slate-500 mb-1">Notificaciones enviadas</div>
          {solicitud.notificaciones.map((n, i) => <div key={i} className="text-[11px] text-slate-500 border-t border-slate-50 pt-0.5">{new Date(n.fecha).toLocaleString("es-CO")} — {n.mensaje}</div>)}
        </div>
      )}

      <div className="text-[11px] text-slate-400 border-t border-slate-200 pt-2">Documento generado automáticamente por el sistema de Gestión de Compras, incluye el histórico completo de cotizaciones, aprobaciones y transacciones de la solicitud.</div>
    </div>
    </div>
  );
}

/* ---------------------------------------------------------
   TIEMPO DEL PROCESO (lead time)
--------------------------------------------------------- */
function TiempoProceso({ historial }) {
  if (!historial?.length) return null;
  const inicio = historial[0].fecha, ultimo = historial[historial.length - 1].fecha;
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="font-medium text-slate-700 mb-2 flex items-center gap-2"><Timer size={16} /> Tiempo del proceso (lead time)</div>
      <table className="w-full text-xs mb-2">
        <thead className="text-slate-400"><tr><th className="text-left py-1">Etapa</th><th className="text-left py-1">Fecha/hora</th><th className="text-right py-1">Duración desde etapa anterior</th></tr></thead>
        <tbody>{historial.map((h, i) => (
          <tr key={i} className="border-t border-slate-100"><td className="py-1">{PASOS.find((p) => p.key === h.status)?.label || h.status}</td><td className="py-1">{new Date(h.fecha).toLocaleString("es-CO")}</td><td className="py-1 text-right">{i === 0 ? "—" : duracion(historial[i - 1].fecha, h.fecha)}</td></tr>
        ))}</tbody>
      </table>
      <div className="text-xs text-slate-500">Tiempo total transcurrido: <b className="text-slate-700">{duracion(inicio, ultimo)}</b></div>
    </div>
  );
}

/* ---------------------------------------------------------
   NOTIFICACIONES (simuladas — requieren backend real en producción)
--------------------------------------------------------- */
function NotificacionesPanel({ notificaciones }) {
  const [abierto, setAbierto] = useState(false);
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4">
      <button onClick={() => setAbierto(!abierto)} className="w-full flex items-center justify-between text-sm font-medium text-slate-700">
        <span className="flex items-center gap-2"><Mail size={15} /> Notificaciones por correo ({notificaciones.length})</span>
        <ChevronRight size={15} className={`transition-transform ${abierto ? "rotate-90" : ""}`} />
      </button>
      {abierto && (
        <div className="mt-2 space-y-1.5">
          <div className="text-[11px] text-slate-400">Simulación dentro del prototipo — en producción esto se envía con un servicio real (ej. Supabase Edge Function + Resend/SendGrid).</div>
          {notificaciones.map((n, i) => <div key={i} className="text-xs text-slate-600 border-t border-slate-100 pt-1.5">{new Date(n.fecha).toLocaleString("es-CO")} — {n.mensaje}</div>)}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   DETALLE DE SOLICITUD
--------------------------------------------------------- */
function accionLabel(solicitud, total) {
  switch (solicitud.status) {
    case "aprobacion_jefe": return "Aprobar como jefe de área";
    case "aprobacion_director": return "Aprobar como director de área";
    case "cotizando": return "Generar cuadro comparativo";
    case "comparativo": return requiereDireccion(total) ? "Enviar a Dirección Financiera" : requiereGerencia(total) ? "Enviar a Gerencia" : "Generar orden";
    case "aprobacion_financiera": return "Aprobar como Dirección Financiera";
    case "aprobacion_gerencia": return "Aprobar como Gerencia";
    case "orden": return "Marcar como enviada al proveedor";
    case "oc_enviada": return "Confirmar recepción / iniciar ejecución";
    case "recepcion": return "Marcar como completada";
    default: return "Avanzar";
  }
}

function SolicitudDetalle({ solicitud, areas, departamentos, empresas, usuarios, proveedores, guardarProveedor, itemsCatalogo, conceptosGasto, historico, setHistorico, currentUser, onUpdate, onEliminar, onVolver, crearNotificacion, guardarItemCatalogo, solicitudes }) {
  const [observacion, setObservacion] = useState("");
  const [prioridadSel, setPrioridadSel] = useState(solicitud.prioridad || "Medio");
  // Compras corrige el nombre que escribió el solicitante y queda guardado en el catálogo de ítems
  const puedeCorregirItems = puedeGestionarCotizaciones(currentUser) && ["cotizando", "comparativo"].includes(solicitud.status);
  const corregirNombreItem = async (it, nuevoNombre) => {
    const nombre = (nuevoNombre || "").trim();
    if (!nombre || nombre === it.nombre) return;
    const existente = itemsCatalogo.find((c) => c.nombre.trim().toLowerCase() === nombre.toLowerCase());
    const actual = it.itemCatalogoId ? itemsCatalogo.find((c) => c.id === it.itemCatalogoId) : null;
    let catId = it.itemCatalogoId || null;
    if (existente) catId = existente.id; // ya existe con el nombre correcto en el catálogo: se vincula
    else if (actual) await guardarItemCatalogo?.({ ...actual, nombre }); // se corrige la entrada del catálogo que había creado el nombre mal escrito
    else { const creado = await guardarItemCatalogo?.({ nombre, unidadDefault: it.unidad, categoria: "" }); if (creado?.id) catId = creado.id; }
    patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, nombre, itemCatalogoId: catId || x.itemCatalogoId } : x)) });
    mostrarToast("Ítem corregido y guardado en el catálogo");
  };
  const abCot = useAbiertosItems(solicitud.items);   // cotizaciones de Compras
  const abComp = useAbiertosItems(solicitud.items);  // cuadro comparativo
  const abAdj = useAbiertosItems(solicitud.items);   // cotizaciones que adjunta el solicitante al corregir
  const abSug = useAbiertosItems(solicitud.items);   // plan de pagos sugerido al corregir
  const [rechazando, setRechazando] = useState(false);
  const [motivoRechazo, setMotivoRechazo] = useState("");
  const [toast, setToast] = useState(null);
  const mostrarToast = (mensaje) => { setToast(mensaje); setTimeout(() => setToast(null), 3500); };
  const area = areas.find((a) => a.id === solicitud.areaId);
  const departamento = departamentos.find((d) => d.id === solicitud.departamentoId);
  const empresa = empresas.find((e) => e.id === solicitud.empresaId);
  const solicitante = usuarios.find((u) => u.id === solicitud.solicitanteId);
  const total = totalSolicitud(solicitud);
  const areaPresup = areas.find((a) => a.id === solicitud.areaId);
  const comprometidoAreaDet = (solicitudes || []).filter((x) => x.areaId === solicitud.areaId && x.id !== solicitud.id && !["solicitud", "aprobacion_jefe", "rechazada"].includes(x.status)).reduce((acc, x) => acc + totalSolicitud(x), 0);
  const todasCotizadas = solicitud.items.every((i) => i.cotizaciones.length > 0 && i.cotizaciones.every((c) => c.proveedorId || (c.proveedorNombre || "").trim()) && !cotizacionUnicaSinJustificar(i) && !itemConPrecioInvalido(i));
  const comparativoBloqueado = ["orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status);
  const patch = (fields) => onUpdate({ ...solicitud, ...fields });

  const reabrirSolicitud = () => {
    const { status: destino, campo, revision } = pasoDelRechazo(solicitud);
    const cambios = {
      status: destino,
      historialEstados: empujarHistorial(destino),
      notificaciones: notificar(`Solicitud reabierta por ${currentUser.nombre} (${currentUser.rol}) para corregir y volver a enviar.`),
    };
    // guarda el motivo del rechazo aparte ANTES de reiniciar la firma — si no, se pierde y nadie
    // puede ver después por qué se devolvió (la firma queda limpia para la próxima aprobación)
    if (campo && solicitud.firmas?.[campo]) {
      cambios.ultimoRechazo = { nombre: solicitud.firmas[campo].nombre, observacion: solicitud.firmas[campo].observacion, fecha: solicitud.firmas[campo].fecha };
    }
    if (campo) cambios.firmas = { ...solicitud.firmas, [campo]: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null } };
    if (revision) cambios.revisionCompras = { estado: "pendiente", observacion: "", usuario: "", fecha: "" };
    patch(cambios);
  };

  // el solicitante confirma que ya corrigió y avisa por correo a quien le toca aprobar ahora
  const reenviarParaAprobacion = () => {
    const responsable = usuarios.find((u) => solicitud.status === "aprobacion_jefe"
      ? (["Jefe de Área", "Jefe de Área y Director"].includes(u.rol) && tieneAreaACargo(u, solicitud.areaId))
      : (["Director de Área", "Jefe de Área y Director"].includes(u.rol) && tieneAreaACargo(u, solicitud.areaId)));
    if (responsable?.email) {
      enviarCorreo(
        responsable.email,
        `Solicitud corregida y lista para tu aprobación: ${solicitud.folio}`,
        `<p>Hola ${responsable.nombre},</p><p>${currentUser.nombre} corrigió la solicitud <b>${solicitud.folio}</b> y quedó lista de nuevo para tu aprobación.</p><p><b>Objetivo:</b> ${solicitud.objetivo}</p>`
      );
    }
    patch({ notificaciones: notificar(responsable?.email ? `${currentUser.nombre} reenvió la solicitud corregida. Correo enviado a ${responsable.nombre} (${responsable.email}).` : `${currentUser.nombre} reenvió la solicitud corregida. No hay un responsable con correo configurado para notificar.`) });
    crearNotificacion?.(responsable?.id, `${currentUser.nombre} corrigió la solicitud ${solicitud.folio} y quedó lista de nuevo para tu aprobación.`, solicitud.id);
    alert(responsable?.email ? `Se avisó a ${responsable.nombre} por correo.` : "Se registró el reenvío, pero no hay un responsable con correo configurado para notificar.");
  };
  const empujarHistorial = (status) => [...solicitud.historialEstados, { status, fecha: ahoraISO() }];
  const notificar = (mensaje) => [...solicitud.notificaciones, { fecha: ahoraISO(), mensaje }];

  const guardarCotizaciones = (itemId, cots) => patch({ items: solicitud.items.map((i) => (i.id === itemId ? { ...i, cotizaciones: cots } : i)) });
  const [mostrarCotGeneralCompras, setMostrarCotGeneralCompras] = useState(false);
  // aplica una misma cotización (proveedor + archivo) a varios ítems a la vez, cada uno con su propio precio
  const aplicarCotizacionGeneralCompras = (precios, cotizacionBase) => {
    patch({
      items: solicitud.items.map((i) => {
        if (!(i.id in precios) || i.cotizaciones.length >= 3) return i;
        return { ...i, cotizaciones: [...i.cotizaciones, { ...cotizacionBase, precioUnitario: precios[i.id], unidadCotizada: i.unidad, factorConversion: 1 }] };
      }),
    });
    setMostrarCotGeneralCompras(false);
  };
  const seleccionarCotizacion = (itemId, idx, obs) => patch({ items: solicitud.items.map((i) => (i.id === itemId ? { ...i, cotizacionSeleccionada: idx, observacionSeleccion: obs } : i)) });
  const guardarItemsRevision = (items) => patch({ items });
  const decidirRevisionCompras = (estado, obs) => {
    if (estado === "rechazada") { patch({ status: "rechazada", revisionCompras: { estado, observacion: obs, usuario: currentUser.nombre, fecha: hoy() }, notificaciones: notificar(`Correo simulado a ${solicitante?.nombre}: tu solicitud ${solicitud.folio} fue rechazada por Compras.`) }); return; }
    patch({ revisionCompras: { estado, observacion: obs, usuario: currentUser.nombre, fecha: hoy() } });
  };

  const puedeActuar = () => {
    switch (solicitud.status) {
      case "aprobacion_jefe": return puedeAprobarJefe(currentUser, solicitud);
      case "aprobacion_director": return puedeAprobarDirector(currentUser, solicitud);
      case "cotizando": return puedeGestionarCotizaciones(currentUser) && (solicitud.tipo !== "compra" || solicitud.revisionCompras.estado === "aprobada");
      case "comparativo": return puedeGestionarCotizaciones(currentUser);
      case "aprobacion_financiera": return puedeAprobarFinanciera(currentUser);
      case "aprobacion_gerencia": return puedeAprobarGerencia(currentUser);
      case "orden": case "oc_enviada": return puedeGestionarCotizaciones(currentUser) || currentUser.rol === "Solicitante";
      case "recepcion": return puedeGestionarCotizaciones(currentUser); // el cierre final (marcar completada) lo hace Compras
      default: return true;
    }
  };

  const firmar = () => ({ aprobado: true, nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresa?.nombre || "", fecha: hoy(), observacion, fotoUrl: currentUser.firmaFotoUrl || null });
  const avanzar = () => {
    const s = solicitud.status;
    // notifica a TODOS los usuarios de un rol (para Dirección Financiera/Gerencia, que suelen ser varios)
    const notificarRol = (rol, mensaje) => usuarios.filter((u) => u.rol === rol).forEach((u) => crearNotificacion?.(u.id, mensaje, solicitud.id));
    if (s === "aprobacion_jefe") {
      if (currentUser.rol === "Jefe de Área y Director") {
        // la misma persona hace ambos roles: se aprueban los dos pasos de una vez, sin duplicar el clic
        patch({ status: "cotizando", prioridad: prioridadSel, firmas: { ...solicitud.firmas, jefe: firmar(), director: { ...firmar(), observacion: "Aprobado junto con el paso de jefe de área (mismo responsable)." } }, historialEstados: empujarHistorial("cotizando"), notificaciones: notificar(`Correo simulado a Compras: solicitud ${solicitud.folio} aprobada (jefe y director), lista para cotizar.`) });
        mostrarToast("✓ Solicitud aprobada como jefe de área y director");
        notificarRol("Compras", `La solicitud ${solicitud.folio} ya está aprobada y lista para cotizar.`);
      } else {
        const director = usuarios.find((u) => u.areaId === solicitud.areaId && ["Director de Área", "Jefe de Área y Director"].includes(u.rol));
        patch({ status: "aprobacion_director", prioridad: prioridadSel, firmas: { ...solicitud.firmas, jefe: firmar() }, historialEstados: empujarHistorial("aprobacion_director"), notificaciones: notificar(director?.email ? `Correo enviado a ${director.nombre} (${director.email}): solicitud ${solicitud.folio} pendiente de tu aprobación.` : `Solicitud aprobada por el jefe de área. No hay un director de área con correo configurado para notificar.`) });
        mostrarToast("✓ Solicitud aprobada como jefe de área");
        crearNotificacion?.(director?.id, `La solicitud ${solicitud.folio} está pendiente de tu aprobación.`, solicitud.id);
      }
    }
    else if (s === "aprobacion_director") { patch({ status: "cotizando", firmas: { ...solicitud.firmas, director: firmar() }, historialEstados: empujarHistorial("cotizando"), notificaciones: notificar(`Correo simulado a Compras: solicitud ${solicitud.folio} aprobada, lista para cotizar.`) }); mostrarToast("✓ Solicitud aprobada como director de área"); notificarRol("Compras", `La solicitud ${solicitud.folio} ya está aprobada y lista para cotizar.`); }
    else if (s === "cotizando" && todasCotizadas) patch({ status: "comparativo", historialEstados: empujarHistorial("comparativo") });
    else if (s === "comparativo") {
      const next = requiereDireccion(total) ? "aprobacion_financiera" : requiereGerencia(total) ? "aprobacion_gerencia" : "orden";
      patch({ status: next, historialEstados: empujarHistorial(next), notificaciones: notificar(`Correo simulado: solicitud ${solicitud.folio} avanza a ${PASOS.find((p) => p.key === next)?.label}.`) });
      if (next === "aprobacion_financiera") notificarRol("Dirección Financiera", `La solicitud ${solicitud.folio} está pendiente de tu aprobación.`);
      else if (next === "aprobacion_gerencia") notificarRol("Gerencia", `La solicitud ${solicitud.folio} está pendiente de tu aprobación.`);
      else notificarRol("Compras", `La solicitud ${solicitud.folio} ya tiene orden generada.`);
    }
    else if (s === "aprobacion_financiera") {
      const pagosOk = planesTodosConfirmados(solicitud);
      if (solicitud.tipo === "servicio" && !pagosOk) { alert("Falta confirmar el plan de pagos (de cada ítem) antes de aprobar y continuar."); return; }
      const next = requiereGerencia(total) ? "aprobacion_gerencia" : "orden";
      patch({ status: next, firmas: { ...solicitud.firmas, financiera: firmar() }, historialEstados: empujarHistorial(next) });
      mostrarToast("✓ Solicitud aprobada por Dirección Financiera");
      if (next === "aprobacion_gerencia") notificarRol("Gerencia", `La solicitud ${solicitud.folio} está pendiente de tu aprobación.`);
      else notificarRol("Compras", `La solicitud ${solicitud.folio} ya tiene orden generada.`);
    }
    else if (s === "aprobacion_gerencia") { patch({ status: "orden", firmas: { ...solicitud.firmas, gerencia: firmar() }, historialEstados: empujarHistorial("orden") }); mostrarToast("✓ Solicitud aprobada por Gerencia"); notificarRol("Compras", `La solicitud ${solicitud.folio} ya tiene orden generada.`); }
    else if (s === "orden") {
      if (!todasOrdenesFirmadas(solicitud, proveedores)) return;
      const ordenesConCorreo = [];
      const ordenesSinCorreo = [];
      (solicitud.ocEnviada.ordenesProveedor || []).forEach((orden) => {
        if (!orden.archivoFirmadoUrl) return;
        const prov = buscarProveedorDeOrden(orden, proveedores);
        if (correosDe(prov).length) ordenesConCorreo.push({ orden, prov });
        else ordenesSinCorreo.push(orden);
      });
      const detalleProveedores = ordenesConCorreo.length ? ` Se envió a: ${ordenesConCorreo.map((o) => `${o.prov.nombre} (${correosDe(o.prov).join(", ")})`).join(", ")}.` : "";
      const avisoSinCorreo = ordenesSinCorreo.length ? ` ⚠ Sin correo registrado, NO se envió a: ${ordenesSinCorreo.map((o) => o.proveedorNombre).join(", ")} — usa "Reenviar orden(es) firmada(s)" para escribirlo y enviarlo.` : "";
      patch({ status: "oc_enviada", historialEstados: empujarHistorial("oc_enviada"), notificaciones: notificar(`Correo enviado a ${solicitante?.nombre} (${solicitante?.email || "sin correo"}) con copia de la orden.${detalleProveedores}${avisoSinCorreo}`) });
      if (solicitante?.email) {
        // arma los enlaces firmados de cada orden ya firmada, para que el solicitante también reciba su copia
        Promise.all(
          (solicitud.ocEnviada.ordenesProveedor || []).filter((o) => o.archivoFirmadoUrl).map(async (o) => ({ nombre: o.proveedorNombre, url: await obtenerUrlFirmada(o.archivoFirmadoUrl, 604800) }))
        ).then((enlaces) => {
          const listaEnlaces = enlaces.filter((e) => e.url).map((e) => `<li><a href="${e.url}">${e.nombre} — ver / descargar orden firmada</a></li>`).join("");
          enviarCorreo(
            solicitante.email,
            `Tu orden ${solicitud.folio} fue enviada al proveedor`,
            `<p>Hola ${solicitante.nombre},</p><p>La orden <b>${solicitud.folio}</b> ya fue enviada al proveedor y quedó lista para recepción. Adjuntamos tu copia:</p><ul>${listaEnlaces}</ul><p>Estos enlaces estarán disponibles por 7 días.</p>`
          );
        });
      }
      // envía la orden firmada por correo a cada proveedor que sí tenga correo registrado (hasta 2 correos por proveedor)
      ordenesConCorreo.forEach(({ orden, prov }) => {
        obtenerUrlFirmada(orden.archivoFirmadoUrl, 604800).then((url) => {
          if (!url) return;
          enviarCorreo(
            correosDe(prov),
            `Solicitud de compra / orden de servicio ${solicitud.folio}`,
            `<p>Hola ${prov.nombre},</p><p>Adjuntamos el enlace de la solicitud de compra / orden de servicio <b>${solicitud.folio}</b> a nombre de ${empresa?.nombre || ""}.</p><p><a href="${url}">Ver / descargar la orden firmada</a></p><p>Este enlace estará disponible por 7 días.</p>`
          );
        });
      });
    }
    else if (s === "oc_enviada") patch({ status: "recepcion", historialEstados: empujarHistorial("recepcion") });
    else if (s === "recepcion") {
      if (!solicitud.recepcion.recibidoSatisfaccion) return;
      if (!evaluacionCompleta(solicitud)) return;
      if (solicitud.tipo === "servicio") {
        const pagado = solicitud.items.reduce((acc, it) => acc + totalPagado(planOficialItem(solicitud, it)), 0);
        if (pagado < total - 0.5) return;
      }
      patch({ status: "completada", historialEstados: empujarHistorial("completada"), notificaciones: notificar(`Correo enviado a ${solicitante?.nombre} (${solicitante?.email || "sin correo"}): tu solicitud ${solicitud.folio} fue completada.`) });
      if (solicitante?.email) {
        enviarCorreo(
          solicitante.email,
          `Tu solicitud ${solicitud.folio} fue completada`,
          `<p>Hola ${solicitante.nombre},</p><p>Tu solicitud <b>${solicitud.folio}</b> quedó completada. Puedes ver el detalle completo y exportarla a PDF desde la aplicación.</p>`
        );
      }
    }
    setObservacion("");
  };
  const rechazar = () => {
    if (!motivoRechazo.trim()) { alert("Escribe el motivo del rechazo — es lo que verá el solicitante para poder corregir la solicitud."); return; }
    const campo = solicitud.status === "aprobacion_jefe" ? "jefe" : solicitud.status === "aprobacion_director" ? "director" : solicitud.status === "aprobacion_financiera" ? "financiera" : solicitud.status === "aprobacion_gerencia" ? "gerencia" : null;
    patch({ status: "rechazada", ultimoRechazo: { nombre: currentUser.nombre, rol: currentUser.rol, observacion: motivoRechazo, fecha: hoy() }, firmas: campo ? { ...solicitud.firmas, [campo]: { aprobado: false, nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresa?.nombre || "", fecha: hoy(), observacion: motivoRechazo, fotoUrl: currentUser.firmaFotoUrl || null } } : solicitud.firmas, historialEstados: empujarHistorial("rechazada"), notificaciones: notificar(`Correo enviado a ${solicitante?.nombre} (${solicitante?.email || "sin correo"}): tu solicitud ${solicitud.folio} fue rechazada.`) });
    if (solicitante?.email) {
      enviarCorreo(
        solicitante.email,
        `Tu solicitud ${solicitud.folio} fue rechazada`,
        `<p>Hola ${solicitante.nombre},</p><p>Tu solicitud <b>${solicitud.folio}</b> fue rechazada por ${currentUser.nombre} (${currentUser.rol}).</p>${motivoRechazo ? `<p><b>Motivo:</b> ${motivoRechazo}</p>` : ""}`
      );
    }
    crearNotificacion?.(solicitante?.id, `Tu solicitud ${solicitud.folio} fue rechazada por ${currentUser.nombre}.${motivoRechazo ? ` Motivo: "${motivoRechazo}"` : ""}`, solicitud.id);
    setMotivoRechazo(""); setRechazando(false);
  };

  const mostrarObservacion = ["aprobacion_jefe", "aprobacion_director", "aprobacion_financiera", "aprobacion_gerencia"].includes(solicitud.status);
  const autorizado = puedeActuar();
  // solo se avisa "tu rol no tiene permiso" a quien de verdad podría tener algo que ver con este paso
  // (ej. un jefe de área de otra área) — no a roles que estructuralmente nunca actúan en este paso
  const ROLES_RELEVANTES_POR_PASO = {
    aprobacion_jefe: ["Jefe de Área", "Jefe de Área y Director"],
    aprobacion_director: ["Director de Área", "Jefe de Área y Director"],
    aprobacion_financiera: ["Dirección Financiera"],
    aprobacion_gerencia: ["Gerencia"],
  };
  const pasoLeConcierne = (ROLES_RELEVANTES_POR_PASO[solicitud.status] || []).includes(currentUser.rol);

  return (
    <div className="space-y-5">
      {toast && (
        <div className="fixed top-4 right-4 z-50 bg-emerald-600 text-white text-sm font-medium px-4 py-2.5 rounded-lg shadow-lg flex items-center gap-2">
          <CheckCircle2 size={16} /> {String(toast).replace(/^✓\s*/, "")}
        </div>
      )}
      <button onClick={onVolver} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-700"><ArrowLeft size={15} /> Volver a solicitudes</button>

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="flex items-start justify-between flex-wrap gap-3">
          {empresa?.logoUrl && <img src={empresa.logoUrl} alt={empresa.nombre} className="h-10 max-w-[100px] object-contain order-first" />}
          <div>
            <div className="flex items-center gap-2 flex-wrap"><h2 className="text-lg font-semibold text-slate-800">{solicitud.folio}</h2><Badge tone={solicitud.tipo === "compra" ? "blue" : "amber"}>{solicitud.tipo === "compra" ? <ShoppingCart size={12} /> : <Wrench size={12} />} {solicitud.tipo === "compra" ? "Solicitud de compra" : "Orden de servicio/trabajo"}</Badge>{solicitud.prioridad && <Badge tone={solicitud.prioridad === "Alto" ? "red" : solicitud.prioridad === "Medio" ? "amber" : "slate"}>Prioridad {solicitud.prioridad}</Badge>}{["recepcion", "completada"].includes(solicitud.status) && solicitud.recepcion?.recibidoSatisfaccion && <Badge tone={solicitud.recepcion.tipoRecepcion === "observaciones" ? "amber" : "green"}>{solicitud.recepcion.tipoRecepcion === "observaciones" ? "Recibida con observaciones" : "Recibida"}</Badge>}{solicitud.status === "rechazada" && <Badge tone="red">Rechazada</Badge>}<button onClick={() => window.print()} className="text-xs bg-slate-800 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1 no-print"><FileText size={13} /> Exportar solicitud completa a PDF</button>{currentUser.rol === "Administrador" && <button onClick={() => onEliminar(solicitud.id, solicitud.folio)} className="text-xs bg-rose-50 text-rose-600 border border-rose-200 px-3 py-1.5 rounded-md font-medium flex items-center gap-1 no-print"><Trash2 size={13} /> Eliminar solicitud</button>}</div>
            <div className="text-sm text-slate-500 mt-1 flex items-center gap-3 flex-wrap"><span className="flex items-center gap-1"><Building2 size={13} /> {empresa?.nombre}</span><span>Área: {area?.nombre}{departamento && ` · Depto: ${departamento.nombre}`}</span><span>Solicitante: {solicitante?.nombre}</span><span className="flex items-center gap-1"><Calendar size={13} /> Est.: {solicitud.fechaEstimada || "—"}</span></div>
          </div>
          <div className="text-right">
            <div className="text-xs text-slate-400">Total solicitud (con IVA)</div>
            <div className="text-xl font-semibold text-slate-800">{fmt(total)}</div>
            {requiereDireccion(total) && <Badge tone="amber">Requiere Dirección Financiera</Badge>}
            {requiereGerencia(total) && <div className="mt-1"><Badge tone="red">Requiere Gerencia</Badge></div>}
          </div>
        </div>
        <div className="mt-4 pt-4 border-t border-slate-100 grid grid-cols-1 md:grid-cols-2 gap-4">
          <div><div className="text-xs font-medium text-slate-400 flex items-center gap-1 mb-1"><Target size={12} /> Objetivo</div><div className="text-sm text-slate-600">{solicitud.objetivo}</div></div>
          <div><div className="text-xs font-medium text-slate-400 flex items-center gap-1 mb-1"><ClipboardList size={12} /> Justificación</div><div className="text-sm text-slate-600">{solicitud.justificacion}</div></div>
        </div>
        <div className="mt-4 pt-4 border-t border-slate-100"><Stepper status={solicitud.status} /></div>
      </div>

      {areaPresup?.presupuesto > 0 && solicitud.status !== "rechazada" && (() => {
        const restante = areaPresup.presupuesto - comprometidoAreaDet - total;
        return (
          <div className={`rounded-xl border p-4 ${restante < 0 ? "bg-rose-50 border-rose-200" : "bg-white border-slate-200"}`}>
            <div className="text-sm font-medium text-slate-700 flex items-center gap-2 mb-2"><DollarSign size={15} /> Presupuesto de {areaPresup.nombre}</div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div><div className="text-slate-400">Presupuesto</div><div className="font-medium text-slate-700">{fmt(areaPresup.presupuesto)}</div></div>
              <div><div className="text-slate-400">Comprometido (otras solicitudes)</div><div className="font-medium text-slate-700">{fmt(comprometidoAreaDet)}</div></div>
              <div><div className="text-slate-400">Esta solicitud</div><div className="font-medium text-slate-700">{fmt(total)}</div></div>
              <div><div className="text-slate-400">Disponible después</div><div className={`font-semibold ${restante < 0 ? "text-rose-600" : "text-emerald-600"}`}>{fmt(restante)}</div></div>
            </div>
            {restante < 0 && <div className="text-[11px] text-rose-600 mt-2">⚠ Esta solicitud supera el presupuesto disponible del área por {fmt(-restante)}.</div>}
            {solicitud.presupuestoAlEnviar && <div className="text-[11px] text-slate-400 mt-2">Al momento de enviarse, el área tenía {fmt(solicitud.presupuestoAlEnviar.disponible)} disponibles.</div>}
          </div>
        );
      })()}

      {(() => {
        // busca la firma que quedó marcada como rechazada (jefe/director/financiera/gerencia) para
        // mostrar el motivo — funciona tanto mientras el estado sigue en "Rechazada" como después
        // de reabrirla, mientras se está corrigiendo (para que no se pierda el motivo original)
        const campoRechazo = ["gerencia", "financiera", "director", "jefe"].find((c) => solicitud.firmas?.[c]?.aprobado === false);
        const firmaRechazo = campoRechazo ? solicitud.firmas[campoRechazo] : (solicitud.ultimoRechazo || null);
        const enCorreccion = solicitud.status !== "rechazada" && solicitud.historialEstados?.some((h) => h.status === "rechazada");
        if (solicitud.status === "rechazada" && (puedeReabrir(currentUser) || currentUser.id === solicitud.solicitanteId)) {
          return (
            <div className="bg-rose-50 border border-rose-200 rounded-xl p-4 flex items-center justify-between gap-3 flex-wrap">
              <div className="text-sm text-rose-700">
                Esta solicitud fue rechazada por {firmaRechazo?.nombre || "—"}.
                {firmaRechazo?.observacion && <div className="mt-1 italic">"{firmaRechazo.observacion}"</div>}
                {" "}Si el motivo fue un error que ya se corrigió (ej. en los precios estimados o en el plan de pagos), puedes reabrirla — volverá al paso donde fue rechazada.
              </div>
              <button onClick={reabrirSolicitud} className="text-xs bg-rose-600 text-white px-3 py-1.5 rounded-md font-medium shrink-0">Reabrir para corregir</button>
            </div>
          );
        }
        if (enCorreccion && (firmaRechazo || solicitud.ultimoRechazo)) {
          const motivo = firmaRechazo || solicitud.ultimoRechazo;
          return (
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-4">
              <div className="text-sm text-amber-700">Esta solicitud fue devuelta por {motivo.nombre} para corregirla:</div>
              <div className="text-sm text-amber-800 italic mt-1">"{motivo.observacion}"</div>
            </div>
          );
        }
        return null;
      })()}

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-3">Ítems solicitados</div>
        {(currentUser.id === solicitud.solicitanteId || puedeReabrir(currentUser)) && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && solicitud.items.every((it) => !(it.cotizaciones?.length > 0)) && (
          <div className="text-xs text-indigo-600 bg-indigo-50 border border-indigo-200 rounded-md px-3 py-2 mb-3">Puedes corregir el precio estimado de cada ítem mientras la solicitud esté en este paso.</div>
        )}
        <div className="overflow-x-auto">
        <table className="w-full text-sm mb-2 min-w-[880px]">
          <thead className="text-slate-400 text-xs">
            <tr>
              <th className="text-left py-1.5 pr-2 w-8">#</th>
              <th className="text-left py-1.5 pr-3">Ítem</th>
              <th className="text-right py-1.5 px-2 whitespace-nowrap w-20">Cantidad</th>
              <th className="text-right py-1.5 px-2 whitespace-nowrap w-20">Unidad</th>
              <th className="text-right py-1.5 px-2 whitespace-nowrap w-28">Valor unitario</th>
              <th className="text-right py-1.5 px-2 whitespace-nowrap w-28">Total ítem</th>
              <th className="text-left py-1.5 px-2 whitespace-nowrap w-40">Cotizaciones</th>
              <th className="w-10"></th>
            </tr>
          </thead>
          <tbody>{solicitud.items.map((it, idx) => {
          const cat = it.itemCatalogoId ? itemsCatalogo.find((c) => c.id === it.itemCatalogoId) : null;
          const desactualizado = cat && (cat.nombre !== it.nombre || cat.unidadDefault !== it.unidad) && !["completada", "rechazada"].includes(solicitud.status);
          const d = desgloseItem(it, solicitud.tipo === "servicio");
          const unitario = parseFloat(it.cantidad) > 0 ? d.subtotal / parseFloat(it.cantidad) : 0;
          const cotConArchivo = (it.cotizaciones || []).filter((c) => c.archivoNombre);
          const puedeAjustarEstimado = (currentUser.id === solicitud.solicitanteId || puedeReabrir(currentUser)) && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && (it.cotizaciones?.length > 0);
          const puedeEditarPrecio = (currentUser.id === solicitud.solicitanteId || puedeReabrir(currentUser)) && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && !(it.cotizaciones?.length > 0);
          return (
            <tr key={it.id} className={`border-t border-slate-100 align-top ${tinteItem(idx).fondo}`}>
              <td className="py-2 pr-2 px-2 text-slate-500 whitespace-nowrap"><span className={`inline-block w-2 h-2 rounded-full mr-1.5 ${tinteItem(idx).punto}`} />{idx + 1}</td>
              <td className="py-2 pr-3">
                {puedeCorregirItems ? (
                  <>
                    <input key={it.nombre} defaultValue={it.nombre} onBlur={(e) => corregirNombreItem(it, e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1 text-sm bg-white" />
                    <div className="text-[10px] text-slate-400 mt-0.5">Compras: corrige el nombre si hace falta — se guarda en el catálogo.</div>
                  </>
                ) : it.nombre}
              </td>
              <td className="py-2 px-2 text-right whitespace-nowrap">{it.cantidad}</td>
              <td className="py-2 px-2 text-right whitespace-nowrap">{it.unidad}</td>
              <td className="py-2 px-2 text-right whitespace-nowrap">
                {puedeEditarPrecio ? (
                  <InputMiles value={it.precioEstimado} onChange={(v) => patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, precioEstimado: v } : x)) })} className="w-24 border border-slate-200 rounded-md px-2 py-1 text-xs text-right" />
                ) : (
                  <>
                    {unitario > 0 ? fmt(unitario) : "—"}
                    {puedeAjustarEstimado && (
                      <div className="mt-1">
                        <div className="text-[10px] text-slate-400">Precio estimado</div>
                        <InputMiles value={it.precioEstimado} onChange={(v) => patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, precioEstimado: v } : x)) })} className="w-24 border border-slate-200 rounded-md px-2 py-1 text-xs text-right" />
                      </div>
                    )}
                  </>
                )}
              </td>
              <td className="py-2 px-2 text-right font-medium whitespace-nowrap">{d.subtotal > 0 ? fmt(d.subtotal) : "—"}</td>
              <td className="py-2 px-2">
                {cotConArchivo.length ? (
                  <div className="flex flex-col gap-1">
                    {cotConArchivo.map((c, ci) => (
                      <EnlacePrivado key={ci} path={c.archivoNombre} className="text-indigo-600 hover:text-indigo-800 flex items-center gap-1 text-xs whitespace-nowrap">
                        <FileText size={12} /> {proveedores.find((p) => p.id === c.proveedorId)?.nombre || c.proveedorNombre || "Proveedor"}
                      </EnlacePrivado>
                    ))}
                  </div>
                ) : <span className="text-slate-300 text-xs">Sin archivo</span>}
              </td>
              <td className="py-2 pl-2 text-right">
                {desactualizado && (
                  <button
                    onClick={() => patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, nombre: cat.nombre, unidad: cat.unidadDefault } : x)) })}
                    title={`El catálogo tiene: "${cat.nombre}" (${cat.unidadDefault})`}
                    className="text-[11px] text-amber-600 underline whitespace-nowrap"
                  >
                    ⚠ Actualizar
                  </button>
                )}
              </td>
            </tr>
          );
          })}</tbody>
        </table>
        </div>
      </div>

      {(currentUser.id === solicitud.solicitanteId || puedeReabrir(currentUser)) && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-4">
          <div className="font-medium text-slate-700">Adjuntar cotizaciones (opcional)</div>
          <div className="text-xs text-slate-400">Si ya tienes una cotización de algún proveedor para un ítem, puedes adjuntarla aquí — le ahorra trabajo a Compras más adelante.</div>
          <ControlExpandirTodo n={solicitud.items.length} onTodos={abAdj.todos} />
          {solicitud.items.map((it, idx) => (
            <ItemColapsable key={it.id} abierto={abAdj.abierto(it.id)} onToggle={() => abAdj.alternar(it.id)} numero={idx + 1} titulo={it.nombre} tinte={tinteItem(idx)} resumen={<span>{(it.cotizaciones || []).length} cotización(es)</span>}>
              <CotizacionForm item={it} proveedores={proveedores} guardarProveedor={guardarProveedor} onGuardar={(_, cots) => patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, cotizaciones: cots } : x)) })} compacto sinIva={solicitud.tipo === "servicio"} />
            </ItemColapsable>
          ))}
        </div>
      )}

      {(currentUser.id === solicitud.solicitanteId || puedeReabrir(currentUser)) && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="font-medium text-slate-700 flex items-center gap-2"><CalendarClock size={16} /> Plan de pagos sugerido, por ítem (opcional)</div>
          <div className="text-xs text-slate-400">Pasa como valor inicial al plan de pagos que revisan Compras y Dirección Financiera.</div>
          <ControlExpandirTodo n={solicitud.items.length} onTodos={abSug.todos} />
          {solicitud.items.map((it, idx) => {
            const totalIt = totalItemConAiu(it, solicitud.tipo === "servicio");
            const sug = planSugeridoItem(solicitud, it);
            return (
              <ItemColapsable key={it.id} abierto={abSug.abierto(it.id)} onToggle={() => abSug.alternar(it.id)} numero={idx + 1} titulo={it.nombre} tinte={tinteItem(idx)} resumen={<span>Total {fmt(totalIt)}{planTieneValores(sug) ? " · con plan" : ""}</span>}>
              <div className="space-y-1.5 p-1">
                {totalIt > 0
                  ? <PlanPagoCotizacion key={it.id} pagos={planTieneValores(sug) ? sug : null} total={totalIt} onChange={(pg) => patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, pagosSugeridos: pg } : x)) })} etiqueta="+ Sugerir plan de pagos para este ítem" titulo="Plan de pagos sugerido de este ítem" />
                  : <div className="text-[11px] text-amber-600">Pon un precio estimado para poder sugerir un plan de pagos.</div>}
              </div>
              </ItemColapsable>
            );
          })}
        </div>
      )}

      {currentUser.id === solicitud.solicitanteId && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && solicitud.historialEstados.some((h) => h.status === "rechazada") && (
        <div className="bg-indigo-50 border border-indigo-200 rounded-xl p-4 flex items-center justify-between gap-3 flex-wrap">
          <div className="text-sm text-indigo-700">Cuando termines de corregir, avísale a quien debe aprobarla — le llega un correo directo.</div>
          <button onClick={reenviarParaAprobacion} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium shrink-0 flex items-center gap-1"><Send size={13} /> Reenviar para aprobación</button>
        </div>
      )}

      {solicitud.status === "cotizando" && puedeVerHistorico(currentUser) && (
        <RevisionCompras solicitud={solicitud} historico={historico} setHistorico={setHistorico} currentUser={currentUser} onGuardarItems={guardarItemsRevision} onDecision={decidirRevisionCompras} />
      )}
      {solicitud.status === "cotizando" && !puedeVerHistorico(currentUser) && solicitud.tipo === "compra" && solicitud.revisionCompras.estado === "pendiente" && (
        <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">Pendiente de revisión por el área de Compras antes de continuar con las cotizaciones.</div>
      )}

      {!comparativoBloqueado && ["cotizando", "comparativo", "aprobacion_financiera", "aprobacion_gerencia"].includes(solicitud.status) && (solicitud.tipo !== "compra" || solicitud.revisionCompras.estado === "aprobada") && puedeGestionarCotizaciones(currentUser) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div className="font-medium text-slate-700">Cargar hasta 3 cotizaciones por ítem (Compras)</div>
            <button onClick={() => setMostrarCotGeneralCompras(true)} className="text-xs text-slate-600 font-medium flex items-center gap-1 border border-slate-200 rounded-md px-2 py-1"><FileText size={13} /> Cotización general</button>
          </div>
          <div className="text-[11px] text-slate-400">Si por error solo guardaste 1 o 2, puedes seguir agregando hasta 3 aquí mismo, incluso después de generar el cuadro comparativo — hasta que se cree la orden.</div>
          {mostrarCotGeneralCompras && (
            <CotizacionGeneralForm items={solicitud.items} proveedores={proveedores} guardarProveedor={guardarProveedor} onAplicar={aplicarCotizacionGeneralCompras} onCerrar={() => setMostrarCotGeneralCompras(false)} />
          )}
          <ControlExpandirTodo n={solicitud.items.length} onTodos={abCot.todos} />
          {solicitud.items.map((it, idxIt) => (
            <ItemColapsable key={it.id} abierto={abCot.abierto(it.id)} onToggle={() => abCot.alternar(it.id)} numero={idxIt + 1} titulo={it.nombre} tinte={tinteItem(idxIt)}
              resumen={<span>{(it.cotizaciones || []).length} cotización(es)</span>}
              insignia={itemConPrecioInvalido(it) ? <Badge tone="red">Revisar precio</Badge> : cotizacionUnicaSinJustificar(it) ? <Badge tone="amber">Falta comentario</Badge> : null}>
              <CotizacionForm item={it} proveedores={proveedores} guardarProveedor={guardarProveedor} onGuardar={guardarCotizaciones} sinIva={solicitud.tipo === "servicio"} />
            </ItemColapsable>
          ))}
        </div>
      )}

      {["aprobacion_jefe", "aprobacion_director", "comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && solicitud.items.some((i) => i.cotizaciones.length > 0) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="font-medium text-slate-700 flex items-center gap-2"><TrendingUp size={16} /> Cuadro comparativo (sugerencia automática)</div>
          {["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && <div className="text-xs text-slate-400">Cotizaciones que el solicitante adjuntó al crear la solicitud — Compras podrá completar y ajustar esto más adelante.</div>}
          <ControlExpandirTodo n={solicitud.items.filter((i) => i.cotizaciones.length > 0).length} onTodos={abComp.todos} />
          {solicitud.items.filter((i) => i.cotizaciones.length > 0).map((it) => <ComparativoTabla key={it.id} abierto={abComp.abierto(it.id)} onToggle={() => abComp.alternar(it.id)} item={it} tinte={tinteItem(solicitud.items.findIndex((x) => x.id === it.id))} numero={solicitud.items.findIndex((x) => x.id === it.id) + 1} proveedores={proveedores} onSeleccionar={seleccionarCotizacion} seleccionada={it.cotizacionSeleccionada} soloLectura={comparativoBloqueado || ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status)} sinIva={solicitud.tipo === "servicio"} ocultarAiu={solicitud.tipo === "servicio" && !aiuVisible(solicitud)} />)}
        </div>
      )}

      {["aprobacion_jefe", "aprobacion_director", "cotizando", "comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && (
        aiuVisible(solicitud)
          ? <AiuEditor proveedores={proveedores} solicitud={solicitud} onGuardarItems={(items) => patch({ items })} editable={puedeGestionarCotizaciones(currentUser) && !["oc_enviada", "recepcion", "completada"].includes(solicitud.status)} />
          : <ResumenTotales solicitud={solicitud} />
      )}

      {["aprobacion_jefe", "aprobacion_director", "cotizando", "comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && (
        <PagosPorItem solicitud={solicitud} currentUser={currentUser} onGuardarItems={(items) => patch({ items })} />
      )}

      <OcEnviadaPanel solicitud={solicitud} proveedores={proveedores} empresa={empresa} currentUser={currentUser} onGuardar={(oc) => patch({ ocEnviada: oc })} />

      {(solicitud.ocEnviada?.ordenesProveedor || []).some((o) => o.archivoFirmadoUrl) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-2">
          <div className="font-medium text-slate-700 flex items-center gap-2"><FileText size={16} /> Copia de la orden enviada al proveedor</div>
          {(solicitud.ocEnviada.ordenesProveedor || []).filter((o) => o.archivoFirmadoUrl).map((o, i) => (
            <div key={i} className="text-sm text-slate-600 flex items-center gap-2">
              <Truck size={13} className="text-slate-400" /> {o.proveedorNombre} — <EnlacePrivado path={o.archivoFirmadoUrl} className="text-indigo-600 underline">Ver / descargar PDF firmado</EnlacePrivado>
            </div>
          ))}
        </div>
      )}

      <ReenviarOrdenesPanel solicitud={solicitud} proveedores={proveedores} guardarProveedor={guardarProveedor} empresa={empresa} currentUser={currentUser} />

      {["recepcion", "completada"].includes(solicitud.status) && <RecepcionPanel solicitud={solicitud} currentUser={currentUser} usuarios={usuarios} onGuardar={(r) => patch({ recepcion: r })} crearNotificacion={crearNotificacion} />}

      {["recepcion", "completada"].includes(solicitud.status) && (
        <EvaluacionPanel
          solicitud={solicitud}
          empresa={empresa}
          proveedores={proveedores}
          currentUser={currentUser}
          onGuardar={(ev) => patch({ evaluacionProveedor: ev })}
        />
      )}

      <div className="print-wrapper-oculto" style={{ display: "none" }}>
        <OrdenDocumento solicitud={solicitud} empresa={empresa} area={area} departamento={departamento} solicitante={solicitante} proveedores={proveedores} conceptosGasto={conceptosGasto} />
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-3 flex items-center gap-2"><PenTool size={15} /> Firmas</div>
        <div className="grid grid-cols-1 md:grid-cols-5 gap-3">
          <FirmaBlock rol="solicitante" firma={solicitud.firmas.solicitante} />
          <FirmaBlock rol="jefe de área" firma={solicitud.firmas.jefe} />
          <FirmaBlock rol="director de área" firma={solicitud.firmas.director} />
          <FirmaBlock rol="dirección financiera" firma={solicitud.firmas.financiera} />
          <FirmaBlock rol="gerencia" firma={solicitud.firmas.gerencia} />
        </div>
      </div>

      <TiempoProceso historial={solicitud.historialEstados} />
      <NotificacionesPanel notificaciones={solicitud.notificaciones} />

      {!autorizado && pasoLeConcierne && solicitud.status !== "completada" && solicitud.status !== "rechazada" && (
        <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 flex items-center gap-2"><ShieldCheck size={14} /> Tu rol ({currentUser.rol}) no tiene permiso para actuar sobre este paso del flujo.</div>
      )}

      {solicitud.status !== "completada" && solicitud.status !== "rechazada" && autorizado && (
        <div className="space-y-2">
          {solicitud.status === "aprobacion_jefe" && (
            <div>
              <label className="text-xs font-medium text-slate-500">Prioridad</label>
              <div className="flex gap-2 mt-1">
                {["Alto", "Medio", "Bajo"].map((p) => (
                  <button key={p} type="button" onClick={() => setPrioridadSel(p)} className={`px-3 py-1.5 rounded-lg text-xs font-medium border ${prioridadSel === p ? (p === "Alto" ? "bg-rose-600 text-white border-rose-600" : p === "Medio" ? "bg-amber-500 text-white border-amber-500" : "bg-slate-500 text-white border-slate-500") : "bg-white text-slate-600 border-slate-200"}`}>{p}</button>
                ))}
              </div>
            </div>
          )}
          {mostrarObservacion && (<div><label className="text-xs font-medium text-slate-500">Observación de aprobación (opcional)</label><textarea value={observacion} onChange={(e) => setObservacion(e.target.value)} rows={2} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" placeholder="Comentarios sobre esta aprobación..." /></div>)}
          {(() => {
            const motivos = [];
            if (solicitud.status === "cotizando" && !solicitud.items.every((i) => i.cotizaciones.length > 0)) motivos.push("Falta cargar al menos una cotización para cada ítem.");
            if (solicitud.status === "cotizando") solicitud.items.forEach((i, idx) => { if (i.cotizaciones.length > 0 && cotizacionUnicaSinJustificar(i)) motivos.push(`Ítem ${idx + 1}: falta el comentario de por qué solo se cotizó con un proveedor.`); });
            if (solicitud.status === "cotizando") solicitud.items.forEach((i, idx) => { (i.cotizaciones || []).forEach((c) => erroresPrecioCotizacion(i, c).forEach((e) => motivos.push(`Ítem ${idx + 1}: ${e.texto}`))); });
            if (solicitud.status === "aprobacion_financiera" && solicitud.tipo === "servicio" && !planesTodosConfirmados(solicitud)) motivos.push("Falta confirmar el plan de pagos de cada ítem.");
            if (solicitud.status === "orden" && !todasOrdenesFirmadas(solicitud, proveedores)) motivos.push("Falta que Dirección Financiera firme la orden de uno o más proveedores.");
            if (solicitud.status === "recepcion") {
              if (!solicitud.recepcion.recibidoSatisfaccion) motivos.push("Falta marcar cómo se recibió (a satisfacción o con observaciones) en el panel de Recepción.");
              if (!evaluacionCompleta(solicitud)) motivos.push("Falta completar la evaluación del proveedor (14 criterios).");
            }
            return motivos.length > 0 && (
              <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 space-y-0.5">
                <div className="font-medium">Falta esto para poder continuar:</div>
                {motivos.map((m, i) => <div key={i}>• {m}</div>)}
              </div>
            );
          })()}
          {rechazando && (
            <div className="bg-rose-50 border border-rose-200 rounded-lg p-3 space-y-2">
              <label className="text-xs font-medium text-rose-700">Motivo del rechazo (obligatorio)</label>
              <div className="text-[11px] text-rose-600">La solicitud se devuelve como <b>rechazada</b>. El solicitante verá este motivo para poder corregirla y volver a enviarla.</div>
              <textarea value={motivoRechazo} onChange={(e) => setMotivoRechazo(e.target.value)} rows={3} className="w-full border border-rose-200 rounded-lg px-3 py-2 text-sm resize-none bg-white" placeholder="Explica qué debe corregir el solicitante..." />
              <div className="flex gap-2 justify-end">
                <button onClick={() => { setRechazando(false); setMotivoRechazo(""); }} className="text-xs text-slate-500 px-3 py-1.5">Cancelar</button>
                <button onClick={rechazar} disabled={!motivoRechazo.trim()} className="text-xs bg-rose-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40">Confirmar rechazo</button>
              </div>
            </div>
          )}
          <div className="flex gap-2 justify-end">
            <button onClick={() => setRechazando((v) => !v)} className="px-4 py-2 rounded-lg text-sm text-rose-600 border border-rose-200 flex items-center gap-1"><XCircle size={15} /> Rechazar</button>
            <button onClick={avanzar} disabled={(solicitud.status === "cotizando" && !todasCotizadas) || (solicitud.status === "aprobacion_financiera" && solicitud.tipo === "servicio" && !planesTodosConfirmados(solicitud)) || (solicitud.status === "orden" && !todasOrdenesFirmadas(solicitud, proveedores)) || (solicitud.status === "recepcion" && (!solicitud.recepcion.recibidoSatisfaccion || !evaluacionCompleta(solicitud)))} className="px-4 py-2 rounded-lg text-sm bg-indigo-600 text-white font-medium disabled:opacity-40 flex items-center gap-1">{accionLabel(solicitud, total)} <ChevronRight size={15} /></button>
          </div>
        </div>
      )}
    </div>
  );
}

// nombre(s) del/los proveedor(es) adjudicado(s) en una solicitud (según la cotización seleccionada por ítem)
function proveedoresAdjudicados(s, proveedores) {
  const nombres = [...new Set(s.items.map((it) => {
    if (!it.cotizaciones.length) return null;
    const idx = idxCotizacionActiva(it, s.tipo === "servicio");
    const cot = it.cotizaciones[idx];
    if (!cot) return null;
    return proveedores.find((p) => p.id === cot.proveedorId)?.nombre || cot.proveedorNombre || null;
  }).filter(Boolean))];
  return nombres.length ? nombres.join(", ") : "—";
}

// lista de proveedores distintos adjudicados en una solicitud, con id (si es del catálogo) y nombre
function proveedoresAdjudicadosDetalle(s, proveedores) {
  const vistos = new Set();
  const lista = [];
  s.items.forEach((it) => {
    if (!it.cotizaciones.length) return;
    const idx = idxCotizacionActiva(it, s.tipo === "servicio");
    const cot = it.cotizaciones[idx];
    if (!cot) return;
    const prov = proveedores.find((p) => p.id === cot.proveedorId);
    const clave = prov?.id || cot.proveedorNombre || "sin-proveedor";
    if (vistos.has(clave)) return;
    vistos.add(clave);
    lista.push({ proveedorId: prov?.id || null, proveedorNombre: prov?.nombre || cot.proveedorNombre || "Proveedor sin definir" });
  });
  return lista;
}

// compara si dos referencias de proveedor (id o nombre libre) son la misma
function mismoProveedor(a, b) {
  if (a.proveedorId || b.proveedorId) return a.proveedorId === b.proveedorId;
  return a.proveedorNombre === b.proveedorNombre;
}

// busca el proveedor real de una orden: primero por ID, y si no hay o no aparece, por nombre
// (sin distinguir mayúsculas/espacios) — cubre datos guardados antes de vincular por ID.
function buscarProveedorDeOrden(orden, proveedores) {
  if (orden.proveedorId) {
    const porId = proveedores.find((p) => p.id === orden.proveedorId);
    if (porId) return porId;
  }
  if (orden.proveedorNombre) {
    const nombreNorm = orden.proveedorNombre.trim().toLowerCase();
    return proveedores.find((p) => p.nombre.trim().toLowerCase() === nombreNorm);
  }
  return null;
}

// junta los hasta 2 correos registrados de un proveedor (email principal + adicional)
function correosDe(prov) {
  return [prov?.email, prov?.email2].map((e) => (e || "").trim()).filter(Boolean);
}

// true si ya existe una orden firmada para cada proveedor adjudicado de la solicitud
// etiqueta de estado a mostrar en listados: distingue "Recibida" dentro del paso de Recepción/Ejecución
function estadoMostrado(s) {
  if (s.status === "rechazada") return "Rechazada";
  if (s.status === "recepcion" && s.recepcion?.recibidoSatisfaccion) return "Recibida";
  return PASOS.find((p) => p.key === s.status)?.label || s.status;
}

function todasOrdenesFirmadas(solicitud, proveedores) {
  const necesarios = proveedoresAdjudicadosDetalle(solicitud, proveedores);
  const ordenes = solicitud.ocEnviada.ordenesProveedor || [];
  return necesarios.length > 0 && necesarios.every((n) => ordenes.some((o) => mismoProveedor(o, n) && o.archivoFirmadoUrl));
}

// true si la evaluación del proveedor (formato oficial, 14 criterios) ya está completa
function evaluacionCompleta(solicitud) {
  return evaluacionProveedorCompleta(solicitud.evaluacionProveedor);
}

// true si esta solicitud está esperando una acción del usuario actual, según su rol y el paso en que está
function requiereMiAccion(currentUser, s, proveedores) {
  if (["completada", "rechazada"].includes(s.status)) return false;
  switch (s.status) {
    case "aprobacion_jefe": return puedeAprobarJefe(currentUser, s);
    case "aprobacion_director": return puedeAprobarDirector(currentUser, s);
    case "cotizando": return puedeGestionarCotizaciones(currentUser) && (s.tipo !== "compra" || s.revisionCompras.estado === "aprobada");
    case "comparativo": return puedeGestionarCotizaciones(currentUser);
    case "aprobacion_financiera": return puedeAprobarFinanciera(currentUser);
    case "aprobacion_gerencia": return puedeAprobarGerencia(currentUser);
    case "orden": return puedeAprobarFinanciera(currentUser) && !todasOrdenesFirmadas(s, proveedores);
    case "recepcion": return puedeGestionarCotizaciones(currentUser) && (!s.recepcion.recibidoSatisfaccion || !evaluacionCompleta(s));
    default: return false;
  }
}

function VistaSolicitudes({ solicitudes, areas, empresas, usuarios, proveedores, currentUser, onAbrir, onExportar, onEliminarSeleccionadas, titulo }) {
  const [fArea, setFArea] = useState("todas");
  const [fEmpresa, setFEmpresa] = useState("todas");
  const [fEstado, setFEstado] = useState("todos");
  const [fTipo, setFTipo] = useState("todos");
  const [fCreadoPor, setFCreadoPor] = useState("todos");
  const [fDesde, setFDesde] = useState("");
  const [fHasta, setFHasta] = useState("");
  const [busqueda, setBusqueda] = useState("");

  const creadores = usuarios.filter((u) => solicitudes.some((s) => s.solicitanteId === u.id));
  const texto = busqueda.trim().toLowerCase();
  const filtradas = solicitudes.filter((s) =>
    (fArea === "todas" || s.areaId === fArea) &&
    (fEmpresa === "todas" || s.empresaId === fEmpresa) &&
    (fEstado === "todos" || s.status === fEstado) &&
    (fTipo === "todos" || s.tipo === fTipo) &&
    (fCreadoPor === "todos" || s.solicitanteId === fCreadoPor) &&
    (!fDesde || s.fechaCreacion >= fDesde) &&
    (!fHasta || s.fechaCreacion <= fHasta) &&
    (!texto ||
      s.folio.toLowerCase().includes(texto) ||
      (s.objetivo || "").toLowerCase().includes(texto) ||
      s.items.some((it) => it.nombre.toLowerCase().includes(texto)) ||
      proveedoresAdjudicados(s, proveedores).toLowerCase().includes(texto))
  );
  const hayFiltros = fArea !== "todas" || fEmpresa !== "todas" || fEstado !== "todos" || fTipo !== "todos" || fCreadoPor !== "todos" || fDesde || fHasta || busqueda;
  const limpiarFiltros = () => { setFArea("todas"); setFEmpresa("todas"); setFEstado("todos"); setFTipo("todos"); setFCreadoPor("todos"); setFDesde(""); setFHasta(""); setBusqueda(""); };

  const exportarExcel = () => {
    const filas = filtradas.map((s) => {
      const area = areas.find((a) => a.id === s.areaId);
      const empresa = empresas.find((e) => e.id === s.empresaId);
      const solicitante = usuarios.find((u) => u.id === s.solicitanteId);
      const paso = PASOS.find((p) => p.key === s.status);
      return {
        "Consecutivo": s.folio,
        "Tipo": s.tipo === "compra" ? "Solicitud de compra" : "Orden de servicio/trabajo",
        "Área": area?.nombre || "",
        "Empresa": empresa?.nombre || "",
        "Solicitante": solicitante?.nombre || "",
        "Fecha de registro": s.fechaCreacion,
        "Fecha estimada": s.fechaEstimada || "",
        "Objetivo": s.objetivo,
        "Justificación": s.justificacion,
        "Proveedor adjudicado": proveedoresAdjudicados(s, proveedores),
        "Total (IVA incl.)": totalSolicitud(s),
        "Estado": s.status === "rechazada" ? "Rechazada" : (paso?.label || s.status),
      };
    });
    const hoja = XLSX.utils.json_to_sheet(filas);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Solicitudes");
    XLSX.writeFile(libro, `solicitudes_${hoy()}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <h2 className="text-lg font-semibold text-slate-800">{titulo}</h2>
        <button onClick={exportarExcel} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><FileText size={13} /> Descargar Excel</button>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
        <div className="flex-1 min-w-[220px]">
          <label className="text-[11px] font-medium text-slate-500">Buscar</label>
          <div className="relative mt-1">
            <input value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Consecutivo, objetivo, ítem o proveedor..." className="w-full border border-slate-200 rounded-lg pl-3 pr-8 py-1.5 text-sm" />
            {busqueda && <button onClick={() => setBusqueda("")} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs">✕</button>}
          </div>
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">Creado por</label>
          <select value={fCreadoPor} onChange={(e) => setFCreadoPor(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm"><option value="todos">Todos</option>{creadores.map((u) => <option key={u.id} value={u.id}>{u.nombre}</option>)}</select>
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">Área</label>
          <select value={fArea} onChange={(e) => setFArea(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm"><option value="todas">Todas</option>{areas.map((a) => <option key={a.id} value={a.id}>{a.nombre}</option>)}</select>
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">Empresa</label>
          <select value={fEmpresa} onChange={(e) => setFEmpresa(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm"><option value="todas">Todas</option>{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select>
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">Estado</label>
          <select value={fEstado} onChange={(e) => setFEstado(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm"><option value="todos">Todos</option>{PASOS.map((p) => <option key={p.key} value={p.key}>{p.label}</option>)}<option value="rechazada">Rechazada</option></select>
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">Tipo</label>
          <select value={fTipo} onChange={(e) => setFTipo(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm"><option value="todos">Todos</option><option value="compra">Solicitud de compra</option><option value="servicio">Orden de servicio/trabajo</option></select>
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">Fecha de registro — desde</label>
          <input type="date" value={fDesde} onChange={(e) => setFDesde(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
        <div>
          <label className="text-[11px] font-medium text-slate-500">hasta</label>
          <input type="date" value={fHasta} onChange={(e) => setFHasta(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
        {hayFiltros && <button onClick={limpiarFiltros} className="text-xs text-slate-500 underline mb-1.5">Limpiar filtros</button>}
        <div className="text-xs text-slate-400 ml-auto mb-1.5">{filtradas.length} de {solicitudes.length} solicitudes</div>
      </div>
      <ListaSolicitudes solicitudes={filtradas} areas={areas} empresas={empresas} proveedores={proveedores} currentUser={currentUser} onAbrir={onAbrir} onExportar={onExportar} onEliminarSeleccionadas={onEliminarSeleccionadas} />
    </div>
  );
}

function ListaSolicitudes({ solicitudes, areas, empresas, proveedores, currentUser, onAbrir, onExportar, onEliminarSeleccionadas }) {
  const [enviandoId, setEnviandoId] = useState(null);
  const [seleccionadas, setSeleccionadas] = useState([]);
  const [porPagina, setPorPagina] = useState(25);
  const [paginaActual, setPaginaActual] = useState(1);
  const esAdmin = currentUser?.rol === "Administrador";

  // paginaSegura ya se encarga de que nunca quede "atascado" en una página que dejó de existir
  // (ej. si el filtro reduce la lista) — no hace falta un efecto aparte que reinicie el estado.

  const totalPaginas = Math.max(1, Math.ceil(solicitudes.length / porPagina));
  const paginaSegura = Math.min(paginaActual, totalPaginas);
  const desde = (paginaSegura - 1) * porPagina;
  const solicitudesPagina = solicitudes.slice(desde, desde + porPagina);

  const alternar = (id) => setSeleccionadas((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const todasSeleccionadas = solicitudes.length > 0 && seleccionadas.length === solicitudes.length;
  const alternarTodas = () => setSeleccionadas(todasSeleccionadas ? [] : solicitudes.map((s) => s.id));

  const eliminarSeleccion = () => {
    if (!seleccionadas.length) return;
    if (!window.confirm(`¿Eliminar ${seleccionadas.length} solicitud(es) por completo? Esta acción no se puede deshacer.`)) return;
    onEliminarSeleccionadas(seleccionadas);
    setSeleccionadas([]);
  };

  const reenviarTodas = async (e, s, empresa) => {
    e.stopPropagation();
    const ordenesFirmadas = (s.ocEnviada?.ordenesProveedor || []).filter((o) => o.archivoFirmadoUrl);
    if (!ordenesFirmadas.length) return;
    setEnviandoId(s.id);
    let enviados = 0, sinCorreo = 0;
    for (const orden of ordenesFirmadas) {
      const prov = buscarProveedorDeOrden(orden, proveedores);
      if (!correosDe(prov).length) { sinCorreo++; continue; }
      const url = await obtenerUrlFirmada(orden.archivoFirmadoUrl, 604800);
      if (url) {
        await enviarCorreo(
          correosDe(prov),
          `Reenvío — Solicitud de compra / orden de servicio ${s.folio}`,
          `<p>Hola ${prov.nombre},</p><p>Te reenviamos el enlace de la solicitud de compra / orden de servicio <b>${s.folio}</b> a nombre de ${empresa?.nombre || ""}.</p><p><a href="${url}">Ver / descargar la orden firmada</a></p><p>Este enlace estará disponible por 7 días.</p>`
        );
        enviados++;
      }
    }
    setEnviandoId(null);
    alert(`${enviados} correo(s) reenviado(s) al proveedor.${sinCorreo ? ` ${sinCorreo} proveedor(es) sin correo registrado — entra a la solicitud para revisarlo.` : ""}`);
  };

  return (
    <div className="space-y-2">
      {esAdmin && seleccionadas.length > 0 && (
        <div className="flex items-center justify-between bg-rose-50 border border-rose-200 rounded-lg px-4 py-2">
          <span className="text-sm text-rose-700">{seleccionadas.length} solicitud(es) seleccionada(s)</span>
          <button onClick={eliminarSeleccion} className="text-xs bg-rose-600 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1"><Trash2 size={13} /> Eliminar seleccionadas</button>
        </div>
      )}
    <div className="bg-white rounded-xl border border-slate-200 overflow-auto max-h-[70vh]">
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-slate-500 sticky top-0 z-10"><tr>
          {esAdmin && <th className="px-4 py-2 w-8"><input type="checkbox" checked={todasSeleccionadas} onChange={alternarTodas} /></th>}
          <th className="text-left px-4 py-2 font-medium">Consecutivo</th><th className="text-left px-4 py-2 font-medium">Tipo</th><th className="text-left px-4 py-2 font-medium">Prioridad</th><th className="text-left px-4 py-2 font-medium">Área</th><th className="text-left px-4 py-2 font-medium">Empresa</th><th className="text-left px-4 py-2 font-medium">Fecha de registro</th><th className="text-left px-4 py-2 font-medium">Objetivo</th><th className="text-left px-4 py-2 font-medium">Proveedor adjudicado</th><th className="text-right px-4 py-2 font-medium">Total (IVA incl.)</th><th className="text-left px-4 py-2 font-medium">Estado</th><th></th><th></th></tr></thead>
        <tbody>{solicitudesPagina.map((s) => { const area = areas.find((a) => a.id === s.areaId), empresa = empresas.find((e) => e.id === s.empresaId), paso = PASOS.find((p) => p.key === s.status);
          const puedeReenviar = currentUser && puedeGestionarCotizaciones(currentUser) && (s.ocEnviada?.ordenesProveedor || []).some((o) => o.archivoFirmadoUrl);
          return (<tr key={s.id} className={`border-t border-slate-100 hover:bg-slate-50 cursor-pointer ${seleccionadas.includes(s.id) ? "bg-rose-50/40" : ""}`} onClick={() => onAbrir(s.id)}>
            {esAdmin && <td className="px-4 py-2.5" onClick={(e) => e.stopPropagation()}><input type="checkbox" checked={seleccionadas.includes(s.id)} onChange={() => alternar(s.id)} /></td>}
            <td className="px-4 py-2.5 font-medium text-slate-700">{s.folio}</td>
            <td className="px-4 py-2.5"><Badge tone={s.tipo === "compra" ? "blue" : "amber"}>{s.tipo === "compra" ? "Compra" : "Servicio"}</Badge></td>
            <td className="px-4 py-2.5">{s.prioridad ? <Badge tone={s.prioridad === "Alto" ? "red" : s.prioridad === "Medio" ? "amber" : "slate"}>{s.prioridad}</Badge> : <span className="text-slate-300 text-xs">—</span>}</td>
            <td className="px-4 py-2.5 text-slate-600">{area?.nombre}</td>
            <td className="px-4 py-2.5 text-slate-600">{empresa?.nombre}</td>
            <td className="px-4 py-2.5 text-slate-600 whitespace-nowrap">{s.fechaCreacion}</td>
            <td className="px-4 py-2.5 text-slate-600 max-w-[220px] truncate" title={s.objetivo}>{s.objetivo}</td>
            <td className="px-4 py-2.5 text-slate-600 max-w-[160px] truncate" title={proveedoresAdjudicados(s, proveedores)}>{proveedoresAdjudicados(s, proveedores)}</td>
            <td className="px-4 py-2.5 text-right text-slate-600">{fmt(totalSolicitud(s))}</td>
            <td className="px-4 py-2.5"><Badge tone={s.status === "completada" ? "green" : s.status === "rechazada" ? "red" : (s.status === "recepcion" && s.recepcion?.recibidoSatisfaccion) ? "blue" : "slate"}>{estadoMostrado(s)}</Badge></td>
            {puedeReenviar && <td className="px-4 py-2.5 text-right"><button title="Reenviar orden firmada al proveedor" disabled={enviandoId === s.id} onClick={(e) => reenviarTodas(e, s, empresa)} className="text-slate-400 hover:text-indigo-600 p-1 disabled:opacity-40"><Send size={15} /></button></td>}
            <td className="px-4 py-2.5 text-right"><button title="Exportar a PDF" onClick={(e) => { e.stopPropagation(); onExportar(s); }} className="text-slate-400 hover:text-indigo-600 p-1"><FileText size={15} /></button></td>
            <td className="px-4 py-2.5 text-right"><ChevronRight size={15} className="text-slate-300" /></td></tr>); })}</tbody>
      </table>
    </div>
    {solicitudes.length > 0 && (
      <div className="flex items-center justify-end gap-4 text-xs text-slate-500 px-1">
        <div className="flex items-center gap-1.5">
          <span>Filas por página:</span>
          <select value={porPagina} onChange={(e) => { setPorPagina(Number(e.target.value)); setPaginaActual(1); }} className="border border-slate-200 rounded-md px-1.5 py-1 text-xs">
            {[10, 25, 50, 100].map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
        </div>
        <span>{desde + 1}–{Math.min(desde + porPagina, solicitudes.length)} de {solicitudes.length}</span>
        <div className="flex items-center gap-1">
          <button onClick={() => setPaginaActual((p) => Math.max(1, p - 1))} disabled={paginaSegura <= 1} className="p-1 rounded disabled:opacity-30 text-slate-500 hover:bg-slate-100"><ChevronRight size={15} className="rotate-180" /></button>
          <button onClick={() => setPaginaActual((p) => Math.min(totalPaginas, p + 1))} disabled={paginaSegura >= totalPaginas} className="p-1 rounded disabled:opacity-30 text-slate-500 hover:bg-slate-100"><ChevronRight size={15} /></button>
        </div>
      </div>
    )}
    </div>
  );
}

/* ---------------------------------------------------------
   LOGOS DE EMPRESAS
--------------------------------------------------------- */
function EmpresasLogos({ empresas, onGuardar }) {
  const cargarLogo = async (empresa, file) => {
    if (!archivoDentroDelLimite(file)) { alert(`El archivo pesa más de ${TAMANO_MAXIMO_MB} MB. Sube uno más liviano.`); return; }
    const url = await subirArchivoPublico(file, "logos");
    if (url) onGuardar({ ...empresa, logoUrl: url });
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="font-medium text-slate-700 mb-3 flex items-center gap-2"><Building2 size={16} /> Logos por empresa</div>
      <div className="text-xs text-slate-400 mb-3">El logo se toma automáticamente según la empresa seleccionada en cada solicitud (se usa en el documento de la orden y en el encabezado).</div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        {empresas.map((e) => (
          <div key={e.id} className="border border-slate-200 rounded-lg p-3 flex items-center gap-3">
            <div className="w-16 h-16 border border-dashed border-slate-300 rounded-lg flex items-center justify-center overflow-hidden bg-slate-50 shrink-0">
              {e.logoUrl ? <img src={e.logoUrl} alt={e.nombre} className="max-w-full max-h-full object-contain" /> : <Building2 size={20} className="text-slate-300" />}
            </div>
            <div className="flex-1">
              <div className="text-sm font-medium text-slate-700">{e.nombre}</div>
              <label className="text-[11px] text-indigo-600 font-medium cursor-pointer inline-flex items-center gap-1 mt-1"><Camera size={11} /> {e.logoUrl ? "Cambiar logo" : "Subir logo"}
                <input type="file" accept="image/*" className="hidden" onChange={(ev) => ev.target.files[0] && cargarLogo(e, ev.target.files[0])} />
              </label>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   CATÁLOGOS
--------------------------------------------------------- */
function Catalogos({
  currentUser,
  solicitudes,
  empresas, guardarEmpresa, eliminarEmpresa,
  areas, guardarArea, eliminarArea,
  departamentos, guardarDepartamento, eliminarDepartamento,
  proveedores, guardarProveedor, eliminarProveedor,
  usuarios, guardarUsuario, eliminarUsuario,
  itemsCatalogo, guardarItemCatalogo, eliminarItemCatalogo,
  conceptosGasto, guardarConceptoGasto, eliminarConceptoGasto,
  permisos, togglePermiso,
}) {
  const [sub, setSub] = useState("empresas");
  const tabs = [
    { key: "empresas", label: "Empresas", icon: Building2 }, { key: "areas", label: "Áreas", icon: Layers }, { key: "departamentos", label: "Departamentos", icon: Layers }, { key: "proveedores", label: "Proveedores", icon: Truck },
    { key: "usuarios", label: "Usuarios y roles", icon: Users }, { key: "items", label: "Ítems", icon: Boxes },
    { key: "conceptos", label: "Conceptos de gasto", icon: ClipboardList },
    ...(currentUser?.rol === "Administrador" ? [{ key: "permisos", label: "Permisos", icon: Lock }] : []),
  ];

  // no se puede borrar un proveedor o un ítem que ya está referenciado en alguna solicitud existente
  const proveedorEnUso = (id) => solicitudes.some((s) => s.items.some((it) => it.cotizaciones.some((c) => c.proveedorId === id)));
  const eliminarProveedorSeguro = (id) => {
    if (proveedorEnUso(id)) { alert("Este proveedor tiene cotizaciones registradas en solicitudes existentes y no se puede eliminar (para no romper ese historial). Puedes editarlo, pero no borrarlo."); return; }
    eliminarProveedor(id);
  };
  const itemEnUso = (id) => solicitudes.some((s) => s.items.some((it) => it.itemCatalogoId === id));
  const eliminarItemCatalogoSeguro = (id) => {
    if (itemEnUso(id)) { alert("Este ítem está siendo usado en solicitudes existentes y no se puede eliminar."); return; }
    eliminarItemCatalogo(id);
  };

  return (
    <div className="space-y-4">
      <div className="flex gap-2 flex-wrap">{tabs.map((t) => (<button key={t.key} onClick={() => setSub(t.key)} className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium border ${sub === t.key ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}><t.icon size={14} /> {t.label}</button>))}</div>
      {sub === "empresas" && (
        <>
          <EmpresasLogos empresas={empresas} onGuardar={guardarEmpresa} />
          <CrudTable titulo="Empresas parametrizadas" icon={Building2} columnas={[{ key: "nombre", label: "Nombre" }, { key: "nit", label: "NIT" }]} datos={empresas} onGuardar={guardarEmpresa} onEliminar={eliminarEmpresa} plantilla={{ nombre: "", nit: "" }} />
        </>
      )}
      {sub === "areas" && (
        <CrudTable titulo="Áreas" icon={Layers}
          columnas={[{ key: "nombre", label: "Nombre" }, { key: "presupuesto", label: "Presupuesto mensual", type: "number" }]}
          datos={areas} onGuardar={guardarArea} onEliminar={eliminarArea} plantilla={{ nombre: "", presupuesto: 0 }} />
      )}
      {sub === "departamentos" && (
        <CrudTable titulo="Departamentos" icon={Layers}
          columnas={[{ key: "nombre", label: "Nombre" }]}
          datos={departamentos} onGuardar={guardarDepartamento} onEliminar={eliminarDepartamento} plantilla={{ nombre: "" }} />
      )}
      {sub === "proveedores" && <CrudTable titulo="Proveedores" icon={Truck}
        columnas={[
          { key: "nombre", label: "Razón social" },
          { key: "tipoProveedor", label: "Tipo de proveedor", type: "select", options: [{ value: "compra", label: "Compra" }, { value: "servicio", label: "Servicio" }, { value: "trabajo", label: "Trabajo" }] },
          { key: "nit", label: "NIT" },
          { key: "ciudad", label: "Ciudad" },
          { key: "direccion", label: "Dirección" },
          { key: "telefono", label: "Teléfono" },
          { key: "representanteLegal", label: "Representante legal" },
          { key: "email", label: "Correo (obligatorio)", requerido: true },
          { key: "email2", label: "Correo adicional (opcional)" },
        ]}
        datos={proveedores} onGuardar={guardarProveedor} onEliminar={eliminarProveedorSeguro}
        plantilla={{ nombre: "", tipoProveedor: "", nit: "", ciudad: "", direccion: "", telefono: "", representanteLegal: "", email: "", email2: "" }} />}
      {sub === "usuarios" && (
        <>
          <div className="text-xs text-slate-400 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2">
            El rol determina qué puede aprobar cada usuario: <b>Jefe de Área</b> aprueba solicitudes de su misma área, <b>Dirección Financiera</b> y <b>Gerencia</b> aprueban según el monto, <b>Compras</b> gestiona cotizaciones, histórico y pagos.
            <br /><b>Importante:</b> editar o agregar una fila aquí solo cambia sus datos de perfil (nombre, cargo, área, rol). Para que una persona pueda <i>iniciar sesión</i>, primero debes crearla en Supabase → Authentication → Users con el mismo correo, y vincular su ID ahí.
          </div>
          <CrudTable titulo="Usuarios y roles" icon={Users} currentUser={currentUser}
            columnas={[
              { key: "nombre", label: "Nombre" },
              { key: "email", label: "Correo electrónico" },
              { key: "cargo", label: "Cargo" },
              { key: "areaId", label: "Área principal", type: "select", options: areas.map((a) => ({ value: a.id, label: a.nombre })) },
              { key: "areasAdicionales", label: "Áreas adicionales a cargo (Director)", type: "multiselect", options: areas.map((a) => ({ value: a.id, label: a.nombre })) },
              { key: "rol", label: "Rol", type: "select", options: ROLES.map((r) => ({ value: r, label: r })), soloAdmin: true },
            ]}
            datos={usuarios} onGuardar={guardarUsuario} onEliminar={eliminarUsuario} plantilla={{ nombre: "", email: "", cargo: "", areaId: "", areasAdicionales: [], rol: "Solicitante" }} />
        </>
      )}
      {sub === "items" && <CrudTable titulo="Catálogo de ítems" icon={Boxes} columnas={[{ key: "nombre", label: "Nombre" }, { key: "unidadDefault", label: "Unidad", type: "select", options: UNIDADES.map((u) => ({ value: u, label: u })) }, { key: "categoria", label: "Categoría" }]} datos={itemsCatalogo} onGuardar={guardarItemCatalogo} onEliminar={eliminarItemCatalogoSeguro} plantilla={{ nombre: "", unidadDefault: "unidad", categoria: "" }} />}
      {sub === "conceptos" && <CrudTable titulo="Conceptos de gasto (plan de cuentas)" icon={ClipboardList} columnas={[{ key: "empresaId", label: "Empresa", type: "select", options: empresas.map((e) => ({ value: e.id, label: e.nombre })), requerido: true }, { key: "grupo", label: "Grupo" }, { key: "codigo", label: "Código" }, { key: "nombre", label: "Cuenta" }, { key: "centroCosto", label: "Centro de costo" }]} datos={conceptosGasto} onGuardar={guardarConceptoGasto} onEliminar={eliminarConceptoGasto} plantilla={{ empresaId: "", grupo: "", codigo: "", nombre: "", centroCosto: "" }} />}
      {sub === "permisos" && currentUser?.rol === "Administrador" && (
        <div className="bg-white rounded-xl border border-slate-200 overflow-x-auto">
          <div className="px-5 py-3 border-b border-slate-100">
            <div className="flex items-center gap-2 font-medium text-slate-700"><Lock size={16} /> Permisos por rol</div>
            <div className="text-xs text-slate-400 mt-1">Marca o desmarca lo que puede hacer cada rol — el cambio aplica de inmediato a todos, sin necesidad de tocar código. Administrador siempre tiene todo, por seguridad no aparece en esta lista.</div>
          </div>
          <table className="w-full text-sm">
            <thead className="bg-slate-50 text-slate-500 text-xs"><tr>
              <th className="text-left px-4 py-2 font-medium sticky left-0 bg-slate-50">Permiso</th>
              {ROLES.filter((r) => r !== "Administrador" && r !== "Solicitante").map((r) => <th key={r} className="text-center px-3 py-2 font-medium whitespace-nowrap">{r}</th>)}
            </tr></thead>
            <tbody>
              {PERMISOS_DISPONIBLES.map((p) => (
                <tr key={p.key} className="border-t border-slate-100">
                  <td className="px-4 py-2 text-slate-600 sticky left-0 bg-white">{p.label}</td>
                  {ROLES.filter((r) => r !== "Administrador" && r !== "Solicitante").map((r) => {
                    const activo = !!permisos.find((x) => x.rol === r && x.permiso === p.key)?.activo;
                    return (
                      <td key={r} className="text-center px-3 py-2">
                        <input type="checkbox" checked={activo} onChange={() => togglePermiso(r, p.key)} />
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   APP PRINCIPAL
--------------------------------------------------------- */
// Campanita de notificaciones dentro de la app — se refresca sola, con contador de no leídas,
// desplegable con las últimas, clic para ir directo a la solicitud y marcarla leída de una vez.
function NotificacionesBell({ notificaciones, onMarcarLeida, onMarcarTodasLeidas, onAbrir }) {
  const [abierto, setAbierto] = useState(false);
  const ref = useRef(null);
  const sinLeer = notificaciones.filter((n) => !n.leida).length;

  useEffect(() => {
    const cerrarSiFuera = (e) => { if (ref.current && !ref.current.contains(e.target)) setAbierto(false); };
    document.addEventListener("mousedown", cerrarSiFuera);
    return () => document.removeEventListener("mousedown", cerrarSiFuera);
  }, []);

  const clickNotificacion = (n) => {
    if (!n.leida) onMarcarLeida(n.id);
    if (n.solicitudId) onAbrir(n.solicitudId);
    setAbierto(false);
  };

  const hace = (fechaISO) => {
    const dif = Date.now() - new Date(fechaISO).getTime();
    const min = Math.floor(dif / 60000);
    if (min < 1) return "Ahora";
    if (min < 60) return `Hace ${min} min`;
    const horas = Math.floor(min / 60);
    if (horas < 24) return `Hace ${horas} h`;
    return `Hace ${Math.floor(horas / 24)} d`;
  };

  return (
    <div ref={ref} className="relative z-40">
      <button onClick={() => setAbierto((v) => !v)} className="relative bg-white border border-slate-200 rounded-full p-2.5 shadow-sm hover:bg-slate-50">
        <Bell size={18} className="text-slate-600" />
        {sinLeer > 0 && <span className="absolute -top-1 -right-1 bg-rose-600 text-white text-[10px] font-semibold rounded-full min-w-[16px] h-4 px-1 flex items-center justify-center">{sinLeer > 9 ? "9+" : sinLeer}</span>}
      </button>
      {abierto && (
        <div className="absolute right-0 mt-2 w-80 bg-white border border-slate-200 rounded-xl shadow-lg overflow-hidden">
          <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
            <span className="text-sm font-medium text-slate-700">Notificaciones</span>
            {sinLeer > 0 && <button onClick={onMarcarTodasLeidas} className="text-[11px] text-indigo-600 font-medium">Marcar todas leídas</button>}
          </div>
          <div className="max-h-96 overflow-y-auto">
            {notificaciones.length === 0 ? (
              <div className="px-4 py-8 text-center text-xs text-slate-400">No tienes notificaciones todavía.</div>
            ) : (
              notificaciones.map((n) => (
                <div key={n.id} onClick={() => clickNotificacion(n)} className={`px-4 py-2.5 border-b border-slate-50 cursor-pointer hover:bg-slate-50 ${!n.leida ? "bg-indigo-50/40" : ""}`}>
                  <div className="flex items-start gap-2">
                    {!n.leida && <span className="w-1.5 h-1.5 rounded-full bg-indigo-600 mt-1.5 shrink-0" />}
                    <div className="min-w-0">
                      <div className={`text-xs ${!n.leida ? "text-slate-700 font-medium" : "text-slate-500"}`}>{n.mensaje}</div>
                      <div className="text-[10px] text-slate-400 mt-0.5">{hace(n.creadoEn)}</div>
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  // --- Catálogos leídos/guardados en Supabase (áreas, departamentos, empresas, proveedores, ítems, centros de costo, conceptos de gasto, usuarios) ---
  const { datos: areas, cargando: cargandoAreas, guardar: guardarArea, eliminar: eliminarArea } = useSupabaseTable('areas', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, presupuesto: r.presupuesto_mensual }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, presupuesto_mensual: Number(r.presupuesto) || 0 }),
    orderBy: 'nombre',
  });
  const { datos: departamentos, cargando: cargandoDepartamentos, guardar: guardarDepartamento, eliminar: eliminarDepartamento } = useSupabaseTable('departamentos', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, areaId: r.area_id }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, area_id: r.areaId }),
    orderBy: 'nombre',
  });
  const { datos: empresas, cargando: cargandoEmpresas, guardar: guardarEmpresa, eliminar: eliminarEmpresa } = useSupabaseTable('empresas', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, nit: r.nit, logoUrl: r.logo_url }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, nit: r.nit, logo_url: r.logoUrl }),
    orderBy: 'nombre',
  });
  const { datos: proveedores, cargando: cargandoProveedores, guardar: guardarProveedor, eliminar: eliminarProveedor, guardarVarios: importarProveedores } = useSupabaseTable('proveedores', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, tipoProveedor: r.tipo_proveedor, nit: r.nit, ciudad: r.ciudad, direccion: r.direccion, telefono: r.telefono, representanteLegal: r.representante_legal, actividadEconomica: r.actividad_economica, contacto: r.contacto, email: r.email, email2: r.email2 }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, tipo_proveedor: r.tipoProveedor, nit: r.nit, ciudad: r.ciudad, direccion: r.direccion, telefono: r.telefono, representante_legal: r.representanteLegal, actividad_economica: r.actividadEconomica, contacto: r.contacto, email: r.email, email2: r.email2 }),
    orderBy: 'nombre',
  });
  const { datos: usuarios, cargando: cargandoUsuarios, guardar: guardarUsuario, eliminar: eliminarUsuario, guardarVarios: importarUsuarios } = useSupabaseTable('usuarios', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, email: r.email, cargo: r.cargo, areaId: r.area_id, areasAdicionales: r.areas_adicionales || [], rol: r.rol, firmaFotoUrl: r.firma_foto_url }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, email: r.email, cargo: r.cargo, area_id: r.areaId, areas_adicionales: r.areasAdicionales || [], rol: r.rol }),
    orderBy: 'nombre',
  });
  const { datos: permisos, cargando: cargandoPermisos, guardar: guardarPermiso } = useSupabaseTable('permisos', {
    desdeDb: (r) => ({ id: r.id, rol: r.rol, permiso: r.permiso, activo: r.activo }),
    haciaDb: (r) => ({ id: r.id, rol: r.rol, permiso: r.permiso, activo: r.activo }),
    orderBy: 'rol',
  });
  // sincroniza el mapa en memoria que usan todas las funciones "puedeXxx" cada vez que
  // cambian los permisos — así se reflejan de inmediato en toda la app, sin recargar.
  useEffect(() => { __permisosPorRol = construirMapaPermisos(permisos); }, [permisos]);
  const togglePermiso = (rol, permisoKey) => {
    const actual = permisos.find((p) => p.rol === rol && p.permiso === permisoKey);
    if (actual) guardarPermiso({ ...actual, activo: !actual.activo });
    else guardarPermiso({ id: nextId(), rol, permiso: permisoKey, activo: true });
  };
  const { datos: itemsCatalogo, cargando: cargandoItems, guardar: guardarItemCatalogo, eliminar: eliminarItemCatalogo, guardarVarios: importarItems } = useSupabaseTable('items_catalogo', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, unidadDefault: r.unidad_default, categoria: r.categoria }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, unidad_default: r.unidadDefault, categoria: r.categoria }),
    orderBy: 'nombre',
  });
  const { datos: conceptosGasto, cargando: cargandoConceptos, guardar: guardarConceptoGasto, eliminar: eliminarConceptoGasto, guardarVarios: importarConceptos } = useSupabaseTable('conceptos_gasto', {
    orderBy: 'nombre',
    desdeDb: (r) => ({ id: r.id, empresaId: r.empresa_id, grupo: r.grupo || "", codigo: r.codigo || "", nombre: r.nombre, centroCosto: r.centro_costo || "" }),
    haciaDb: (r) => ({ id: r.id, empresa_id: r.empresaId || null, grupo: r.grupo || "", codigo: r.codigo || "", nombre: r.nombre, centro_costo: r.centroCosto || "" }),
  });
  const cargandoCatalogos = cargandoAreas || cargandoDepartamentos || cargandoEmpresas || cargandoProveedores || cargandoUsuarios || cargandoItems || cargandoConceptos;

  const [historico, setHistorico] = useState(HISTORICO_INIT);
  const { solicitudes, cargando: cargandoSolicitudes, crear: crearSolicitudDB, actualizar: actualizarSolicitudDB, eliminar: eliminarSolicitudDB } = useSolicitudes();
  const [tab, setTab] = useState("solicitudes");
  const [abierta, setAbierta] = useState(null);
  const [creando, setCreando] = useState(false);
  const [perfil, setPerfil] = useState(false);
  const [menuExpandido, setMenuExpandido] = useState(true);
  const [gruposAbiertos, setGruposAbiertos] = useState({ misSolicitudes: true, misPendientes: true, gestion: true });
  const toggleGrupo = (g) => setGruposAbiertos((prev) => ({ ...prev, [g]: !prev[g] }));
  const [exportando, setExportando] = useState(null);

  useEffect(() => {
    if (!exportando) return;
    const t = setTimeout(() => window.print(), 150);
    const limpiar = () => setExportando(null);
    window.addEventListener("afterprint", limpiar);
    return () => { clearTimeout(t); window.removeEventListener("afterprint", limpiar); };
  }, [exportando]);

  // --- Sesión real con Supabase Auth ---
  const { perfil: perfilAuth, cargando: cargandoSesion, iniciarSesion, cerrarSesion, actualizarPerfil } = useAuth();
  const { notificaciones: notisUsuario, crear: crearNotiUsuario, marcarLeida: marcarNotiLeida, marcarTodasLeidas: marcarTodasNotisLeidas } = useNotificaciones(perfilAuth?.id);

  if (cargandoSesion) {
    return <div className="min-h-screen flex items-center justify-center text-slate-400 text-sm">Cargando...</div>;
  }
  if (!perfilAuth) return <LoginReal onIniciarSesion={iniciarSesion} />;
  if (cargandoCatalogos || cargandoSolicitudes) {
    return <div className="min-h-screen flex items-center justify-center text-slate-400 text-sm">Cargando catálogos y solicitudes...</div>;
  }

  // Adapta el perfil que viene de Supabase (snake_case) a la forma que usa el resto de la app (camelCase).
  const currentUser = {
    id: perfilAuth.id,
    nombre: perfilAuth.nombre,
    email: perfilAuth.email,
    cargo: perfilAuth.cargo,
    rol: perfilAuth.rol,
    firmaFotoUrl: perfilAuth.firma_foto_url,
    areaId: perfilAuth.area_id,
    areasAdicionales: perfilAuth.areas_adicionales || [],
  };

  const crearSolicitud = async (nueva) => { await crearSolicitudDB(nueva); setCreando(false); setTab("solicitudes"); };
  const actualizarSolicitud = async (upd) => { await actualizarSolicitudDB(upd); };

  const eliminarSolicitud = async (id, folio) => {
    if (currentUser.rol !== "Administrador") return;
    if (!window.confirm(`¿Eliminar por completo la solicitud ${folio}? Esta acción no se puede deshacer — se borra todo su historial, cotizaciones, firmas y evaluación.`)) return;
    await eliminarSolicitudDB(id);
    setAbierta(null);
  };
  const eliminarSolicitudesSeleccionadas = async (ids) => {
    if (currentUser.rol !== "Administrador") return;
    for (const id of ids) await eliminarSolicitudDB(id);
  };
  // La foto de firma del perfil, por ahora, solo se guarda en memoria durante la sesión.
  // Falta conectar esto a un "update" real sobre la tabla usuarios (próximo módulo a migrar).
  const guardarPerfil = async (u) => { await actualizarPerfil({ firma_foto_url: u.firmaFotoUrl }); setPerfil(false); };
  const solicitudAbierta = solicitudes.find((s) => s.id === abierta);
  const solicitudesVisibles = puedeVerTodasSolicitudes(currentUser) ? solicitudes : solicitudes.filter((s) => s.solicitanteId === currentUser.id);
  const solicitudesPorFirmar = solicitudes.filter((s) => s.status === "orden" && !todasOrdenesFirmadas(s, proveedores));
  const solicitudesMisPendientes = solicitudes.filter((s) => requiereMiAccion(currentUser, s, proveedores));

  const NavBtn = ({ id, icon: Icon, label, badge }) => (
    <button title={label} onClick={() => { setTab(id); setAbierta(null); setCreando(false); setPerfil(false); }} className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium w-full text-left ${!menuExpandido ? "justify-center px-2" : ""} ${tab === id && !abierta && !creando && !perfil ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"}`}>
      <Icon size={16} className="shrink-0" /> {menuExpandido && <span className="flex-1">{label}</span>}
      {!!badge && <span className={`text-[10px] font-bold rounded-full px-1.5 py-0.5 ${tab === id ? "bg-white text-indigo-600" : "bg-rose-500 text-white"}`}>{badge}</span>}
    </button>
  );

  return (
    <div className="min-h-screen bg-slate-50 flex text-slate-800" style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      <aside className={`${menuExpandido ? "w-56" : "w-16"} bg-white border-r border-slate-200 p-3 flex flex-col gap-1 shrink-0 transition-all duration-200`}>
        <div className={`flex items-center gap-2 mb-4 ${menuExpandido ? "px-1 justify-between" : "justify-center"}`}>
          {menuExpandido && (
            <div className="flex items-center gap-2 min-w-0">
              <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center text-white font-bold text-sm shrink-0">GC</div>
              <div className="min-w-0"><div className="text-sm font-semibold text-slate-800 leading-tight truncate">Gestión de Compras</div><div className="text-[11px] text-slate-400 leading-tight">Multiempresa</div></div>
            </div>
          )}
          <button title={menuExpandido ? "Contraer menú" : "Expandir menú"} onClick={() => setMenuExpandido(!menuExpandido)} className="text-slate-400 hover:text-indigo-600 hover:bg-slate-100 rounded-lg p-1.5 shrink-0">
            {menuExpandido ? <PanelLeftClose size={16} /> : <PanelLeftOpen size={16} />}
          </button>
        </div>

        {menuExpandido ? (
          <button onClick={() => toggleGrupo("misSolicitudes")} className="flex items-center justify-between px-2 mt-1 mb-0.5 w-full text-[10px] font-semibold text-slate-400 uppercase tracking-wide hover:text-slate-600">
            <span>Solicitudes</span><ChevronRight size={11} className={`transition-transform ${gruposAbiertos.misSolicitudes ? "rotate-90" : ""}`} />
          </button>
        ) : <div className="mt-1" />}
        {(gruposAbiertos.misSolicitudes || !menuExpandido) && (
          <>
            <button title="Nueva solicitud" onClick={() => { setCreando(true); setAbierta(null); setPerfil(false); }} className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium w-full text-left mb-1 ${!menuExpandido ? "justify-center px-2" : ""} ${creando ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"}`}>
              <Plus size={16} className="shrink-0" /> {menuExpandido && "Nueva solicitud"}
            </button>
            <NavBtn id="solicitudes" icon={ListChecks} label={puedeVerTodasSolicitudes(currentUser) ? "Solicitudes" : "Mis solicitudes"} />
          </>
        )}

        {(puedeVerMisPendientes(currentUser) || puedeAprobarFinanciera(currentUser)) && (
          <>
            {menuExpandido ? (
              <button onClick={() => toggleGrupo("misPendientes")} className="flex items-center justify-between px-2 mt-3 mb-0.5 w-full text-[10px] font-semibold text-slate-400 uppercase tracking-wide hover:text-slate-600">
                <span>Pendientes</span><ChevronRight size={11} className={`transition-transform ${gruposAbiertos.misPendientes ? "rotate-90" : ""}`} />
              </button>
            ) : <div className="mt-3" />}
            {(gruposAbiertos.misPendientes || !menuExpandido) && (
              <>
                {puedeVerMisPendientes(currentUser) && <NavBtn id="misPendientes" icon={Clock} label="Pendientes" badge={solicitudesMisPendientes.length} />}
                {puedeAprobarFinanciera(currentUser) && <NavBtn id="porFirmar" icon={PenTool} label="Órdenes por firmar" badge={solicitudesPorFirmar.length} />}
              </>
            )}
          </>
        )}

        {menuExpandido ? (
          <button onClick={() => toggleGrupo("gestion")} className="flex items-center justify-between px-2 mt-3 mb-0.5 w-full text-[10px] font-semibold text-slate-400 uppercase tracking-wide hover:text-slate-600">
            <span>Gestión de la información</span><ChevronRight size={11} className={`transition-transform ${gruposAbiertos.gestion ? "rotate-90" : ""}`} />
          </button>
        ) : <div className="mt-3" />}
        {(gruposAbiertos.gestion || !menuExpandido) && (
          <>
            <NavBtn id="dashboard" icon={LayoutDashboard} label="Dashboard" />
            <NavBtn id="estadisticas" icon={BarChart3} label="Estadísticas" />
            {puedeVerEvaluaciones(currentUser) && <NavBtn id="evalProveedores" icon={Award} label="Evaluación proveedores" />}
            {puedeVerCalendarioPagos(currentUser) && <NavBtn id="calendarioPagos" icon={CalendarClock} label="Calendario de pagos" />}
            {puedeVerOrdenesEnviadas(currentUser) && <NavBtn id="ordenesEnviadas" icon={FileText} label="Órdenes enviadas" />}
            {puedeVerPlanInversion(currentUser) && <NavBtn id="planInversion" icon={TrendingUp} label="Plan de inversión" />}
            {puedeVerCatalogos(currentUser) && <NavBtn id="catalogos" icon={Settings} label="Catálogo" />}
          </>
        )}

        <div className="mt-auto pt-4 border-t border-slate-100">
          <button title="Mi perfil" onClick={() => { setPerfil(true); setAbierta(null); setCreando(false); }} className={`flex items-center gap-2 px-2 py-1.5 rounded-lg text-xs text-slate-500 hover:bg-slate-100 w-full mb-1 ${!menuExpandido ? "justify-center" : ""}`}><UserCircle size={13} className="shrink-0" /> {menuExpandido && "Mi perfil"}</button>
          {menuExpandido && (
            <div className="px-2 mb-2">
              <div className="text-sm font-medium text-slate-700 truncate">{currentUser.nombre}</div>
              <div className="text-[11px] text-slate-400 truncate" title={[areas.find((a) => a.id === currentUser.areaId)?.nombre, ...(currentUser.areasAdicionales || []).map((id) => areas.find((a) => a.id === id)?.nombre)].filter(Boolean).join(", ")}>
                {currentUser.rol} · {[areas.find((a) => a.id === currentUser.areaId)?.nombre, ...(currentUser.areasAdicionales || []).map((id) => areas.find((a) => a.id === id)?.nombre)].filter(Boolean).join(", ")}
              </div>
            </div>
          )}
          <button title="Cerrar sesión" onClick={() => { cerrarSesion(); setAbierta(null); setCreando(false); }} className={`flex items-center gap-2 px-2 py-1.5 rounded-lg text-xs text-slate-500 hover:bg-slate-100 w-full ${!menuExpandido ? "justify-center" : ""}`}><LogOut size={13} className="shrink-0" /> {menuExpandido && "Cerrar sesión"}</button>
          {menuExpandido && <div className="text-[11px] text-slate-400 px-2 leading-relaxed mt-2">Umbral Dir. Financiera: {fmt(UMBRAL_DIRECCION)}<br />Umbral Gerencia: {fmt(UMBRAL_GERENCIA)}</div>}
        </div>
      </aside>

      <div className="flex-1 flex flex-col overflow-hidden">
        <div className="flex justify-end px-4 pt-3 shrink-0">
          <NotificacionesBell notificaciones={notisUsuario} onMarcarLeida={marcarNotiLeida} onMarcarTodasLeidas={marcarTodasNotisLeidas} onAbrir={(id) => { setAbierta(id); setCreando(false); setPerfil(false); }} />
        </div>

      <main className="flex-1 p-6 pt-2 overflow-auto">
        {creando ? (
          <NuevaSolicitud areas={areas} departamentos={departamentos} empresas={empresas} itemsCatalogo={itemsCatalogo} guardarItemCatalogo={guardarItemCatalogo} proveedores={proveedores} guardarProveedor={guardarProveedor} conceptosGasto={conceptosGasto} usuarios={usuarios} currentUser={currentUser} solicitudes={solicitudes} onCrear={crearSolicitud} onCancel={() => setCreando(false)} />
        ) : perfil ? (
          <PerfilUsuario currentUser={currentUser} onGuardar={guardarPerfil} />
        ) : solicitudAbierta ? (
          <SolicitudDetalle solicitudes={solicitudes} guardarItemCatalogo={guardarItemCatalogo} solicitud={solicitudAbierta} areas={areas} departamentos={departamentos} empresas={empresas} usuarios={usuarios} proveedores={proveedores} guardarProveedor={guardarProveedor} itemsCatalogo={itemsCatalogo} conceptosGasto={conceptosGasto} historico={historico} setHistorico={setHistorico} currentUser={currentUser} onUpdate={actualizarSolicitud} onEliminar={eliminarSolicitud} onVolver={() => setAbierta(null)} crearNotificacion={crearNotiUsuario} />
        ) : tab === "dashboard" ? (
          <Dashboard areas={areas} solicitudes={solicitudesVisibles} proveedores={proveedores} currentUser={currentUser} onAbrir={setAbierta} onVerCalendario={() => setTab("calendarioPagos")} />
        ) : tab === "misPendientes" && puedeVerMisPendientes(currentUser) ? (
          <div className="space-y-4">
            <div>
              <h2 className="text-lg font-semibold text-slate-800">Mis pendientes</h2>
              <p className="text-xs text-slate-400 mt-1">Solicitudes que están esperando una acción tuya en este momento, según tu rol.</p>
            </div>
            {solicitudesMisPendientes.length ? (
              <ListaSolicitudes solicitudes={solicitudesMisPendientes} areas={areas} empresas={empresas} proveedores={proveedores} currentUser={currentUser} onAbrir={setAbierta} onExportar={setExportando} onEliminarSeleccionadas={eliminarSolicitudesSeleccionadas} />
            ) : (
              <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">No tienes solicitudes pendientes de tu acción en este momento.</div>
            )}
          </div>
        ) : tab === "porFirmar" && puedeAprobarFinanciera(currentUser) ? (
          <div className="space-y-4">
            <div>
              <h2 className="text-lg font-semibold text-slate-800">Órdenes por firmar</h2>
              <p className="text-xs text-slate-400 mt-1">Solicitudes en el paso "Orden generada" que tienen al menos una orden de proveedor pendiente de tu firma.</p>
            </div>
            {solicitudesPorFirmar.length ? (
              <ListaSolicitudes solicitudes={solicitudesPorFirmar} areas={areas} empresas={empresas} proveedores={proveedores} currentUser={currentUser} onAbrir={setAbierta} onExportar={setExportando} onEliminarSeleccionadas={eliminarSolicitudesSeleccionadas} />
            ) : (
              <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">No hay órdenes pendientes de firma en este momento.</div>
            )}
          </div>
        ) : tab === "estadisticas" ? (
          <Estadisticas solicitudes={solicitudesVisibles} areas={areas} empresas={empresas} proveedores={proveedores} />
        ) : tab === "evalProveedores" && puedeVerEvaluaciones(currentUser) ? (
          <ReporteEvaluacionesProveedores solicitudes={solicitudes} proveedores={proveedores} onAbrir={setAbierta} />
        ) : tab === "calendarioPagos" && puedeVerCalendarioPagos(currentUser) ? (
          <CalendarioPagos solicitudes={solicitudes} proveedores={proveedores} onAbrir={setAbierta} />
        ) : tab === "ordenesEnviadas" && puedeVerOrdenesEnviadas(currentUser) ? (
          <ReporteOrdenesEnviadas solicitudes={solicitudes} proveedores={proveedores} empresas={empresas} onAbrir={setAbierta} />
        ) : tab === "planInversion" && puedeVerPlanInversion(currentUser) ? (
          <PlanInversion empresas={empresas} currentUser={currentUser} solicitudes={solicitudes} proveedores={proveedores} onAbrir={setAbierta} onActualizarSolicitud={actualizarSolicitudDB} />
        ) : tab === "catalogos" && puedeVerCatalogos(currentUser) ? (
          <Catalogos
            currentUser={currentUser}
            solicitudes={solicitudes}
            empresas={empresas} guardarEmpresa={guardarEmpresa} eliminarEmpresa={eliminarEmpresa}
            areas={areas} guardarArea={guardarArea} eliminarArea={eliminarArea}
            departamentos={departamentos} guardarDepartamento={guardarDepartamento} eliminarDepartamento={eliminarDepartamento}
            proveedores={proveedores} guardarProveedor={guardarProveedor} eliminarProveedor={eliminarProveedor}
            usuarios={usuarios} guardarUsuario={guardarUsuario} eliminarUsuario={eliminarUsuario}
            itemsCatalogo={itemsCatalogo} guardarItemCatalogo={guardarItemCatalogo} eliminarItemCatalogo={eliminarItemCatalogo}
            conceptosGasto={conceptosGasto} guardarConceptoGasto={guardarConceptoGasto} eliminarConceptoGasto={eliminarConceptoGasto}
            permisos={permisos} togglePermiso={togglePermiso}
          />
        ) : (
          <VistaSolicitudes solicitudes={solicitudesVisibles} areas={areas} empresas={empresas} usuarios={usuarios} proveedores={proveedores} currentUser={currentUser} onAbrir={setAbierta} onExportar={setExportando} onEliminarSeleccionadas={eliminarSolicitudesSeleccionadas} titulo={puedeVerTodasSolicitudes(currentUser) ? "Solicitudes" : "Mis solicitudes"} />
        )}
      </main>
      </div>

      {exportando && !solicitudAbierta && (
        <div className="print-wrapper-oculto" style={{ display: "none" }}>
          <OrdenDocumento
            solicitud={exportando}
            empresa={empresas.find((e) => e.id === exportando.empresaId)}
            area={areas.find((a) => a.id === exportando.areaId)}
            departamento={departamentos.find((d) => d.id === exportando.departamentoId)}
            solicitante={usuarios.find((u) => u.id === exportando.solicitanteId)}
            proveedores={proveedores}
            conceptosGasto={conceptosGasto}
          />
        </div>
      )}
    </div>
  );
}