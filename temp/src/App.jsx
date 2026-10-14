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
import { unificarArchivosPDF, generarResumenSolicitudPDF } from "./lib/unificarArchivos";
import { COLUMNAS_EMPLEADOS, validarEmpleados, importarEmpleadosEnLotes } from "./lib/importarEmpleados";
import { PALETAS, TEMA_ORIGINAL, aplicarTema, sanitizarTema, soportaTema, leerOriginal, estiloDocumento, generarEscala, parametrosDe } from "./lib/tema";
import { filasExportarAreas, filasDirectoresValidos, filasInstruccionesAreas, validarAreas, registrosParaGuardarAreas } from "./lib/areasExcel";
import { enviarCorreo } from "./lib/correo";
import LoginReal from "./LoginReal";
import { LOGO_FULL } from "./lib/logo";
import { supabase } from "./lib/supabaseClient";
import {
  ShoppingCart, Wrench, Building2, CheckCircle2, XCircle, Clock,
  FileText, TrendingUp, ChevronRight, Plus, Trash2, Pencil,
  Calendar, Award, ArrowLeft, LayoutDashboard, ListChecks, BarChart3,
  DollarSign, PackageCheck, CalendarClock, Boxes, Users, Truck,
  Settings, Target, ClipboardList, Lock, LogOut, History, PenTool, ShieldCheck,
  Paperclip, Mail, Camera, Timer, AlertTriangle, Info, Palette, Layers, MessageSquare, UserCircle, Send, CheckSquare, PanelLeftClose, PanelLeftOpen, Upload as UploadIcon, Bell, Archive,
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

// los umbrales se editan en Catálogo → Parámetros (tabla "parametros"); estos son los valores por defecto si la tabla aún no existe
let UMBRAL_DIRECCION = 500000;
let UMBRAL_GERENCIA = 100000000;
const UNIDADES = ["unidad", "libra", "kilo", "gramo", "litro", "mililitro", "metro", "caja", "paquete", "hora", "servicio"];
const IVA_OPCIONES = [0, 5, 19];
const MONEDAS = ["COP", "USD", "EUR", "MXN"];
// el primero es el color de acento: sigue la paleta de cada usuario (el resto son colores de estado y no cambian)
const COLORS = ["var(--color-indigo-600)", "#f59e0b", "#10b981", "#ef4444", "#0ea5e9", "#a855f7"];
const ROLES = ["Solicitante", "Jefe de Área", "Director de Área", "Jefe de Área y Director", "Dirección Financiera", "Contabilidad", "Gerencia", "Compras", "Administrador"];

const PASOS = [
  { key: "solicitud", label: "Solicitud creada" },
  { key: "aprobacion_jefe", label: "Aprobación jefe de área (flujo anterior)", oculto: true }, // ya no existe en el flujo; solo sirve para leer historiales viejos
  { key: "aprobacion_director", label: "Aprobación director de área" },
  { key: "aprobacion_gerente", label: "Aprobación gerente general (paso retirado)", oculto: true }, // ya no existe en el flujo; solo sirve para leer historiales
  { key: "cotizando", label: "Revisión y cotizaciones (compras)" },
  { key: "comparativo", label: "Cuadro comparativo" },
  { key: "aprobacion_financiera", label: "Dirección financiera" },
  { key: "aprobacion_gerencia", label: "Gerencia" },
  { key: "orden", label: "Orden generada" },
  { key: "oc_enviada", label: "Enviada al proveedor" },
  { key: "recepcion", label: "Recepción / Ejecución" },
  { key: "completada", label: "Completada" },
];

// pasos que realmente existen hoy (sin el paso del jefe de área, que ya no aprueba nadie)
const PASOS_ACTIVOS = PASOS.filter((p) => !p.oculto);
// una solicitud que quedó esperando al jefe de área (flujo anterior) pasa a esperar al director
// Estados que ya no existen o quedaron mal, llevados al paso que realmente les corresponde:
//  - esperando al jefe de área (flujo anterior): pasa a esperar al director, que es quien aprueba;
//  - esperando al director, pero con la firma del director ya puesta: ya fue aprobada (p. ej. una solicitud que Compras devolvió y
//    que cayó de nuevo en la primera aprobación) — sigue con Compras, no se le pide al director que la apruebe otra vez;
//  - en el paso retirado del gerente general: ya tenía la aprobación del director, así que sigue a Compras.
const normalizarEstadoLegado = (s) => {
  if (s.status === "aprobacion_gerente") return { ...s, status: "cotizando" };
  if (s.status === "aprobacion_jefe" || s.status === "aprobacion_director") {
    if (s.firmas?.director?.aprobado === true) return { ...s, status: "cotizando" };
    if (s.status === "aprobacion_jefe") return { ...s, status: "aprobacion_director" };
  }
  return s;
};
// etapas en las que la solicitud todavía NO compromete el presupuesto del área (falta la aprobación del director)
const ESTADOS_SIN_COMPROMISO = ["solicitud", "aprobacion_jefe", "aprobacion_director", "rechazada"];

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
// fases (secciones) del detalle de la solicitud que se pueden expandir o contraer
const FASES_DETALLE = ["presupuesto", "items", "adjuntar", "sugerido", "revision", "cotizaciones", "comparativo", "aiu", "pagos", "orden", "copia", "reenviar", "estadoOrden", "recepcion", "evaluacion", "firmas"].map((id) => ({ id }));
const ETAPAS_SIN_AIU = ["solicitud", "aprobacion_jefe", "aprobacion_director"];
const tieneAiuDefinido = (s) => tieneAiuValores(s.aiu) || (s.items || []).some((it) => tieneAiuValores(it.aiu) || (it.cotizaciones || []).some((c) => tieneAiuValores(c.aiu)));
const aiuVisible = (s) => s.tipo === "servicio" && (s.items || []).some((it) => (it.cotizaciones || []).length > 0) && (!ETAPAS_SIN_AIU.includes(s.status) || tieneAiuDefinido(s));
// Avisos (toasts) globales: el mismo aviso verde de arriba a la derecha para toda la app — crear una solicitud,
// cada cambio del flujo y las notificaciones que llegan. tipo: "ok" (verde), "alerta" (rojo: rechazos y errores),
// "info" (índigo: notificaciones). Se cierran solos, o al hacer clic (y si traen onClick, lo ejecutan).
const ToastContext = React.createContext({ mostrar: () => {} });
const useToast = () => React.useContext(ToastContext);
function ToastHost({ toasts, onCerrar }) {
  const colores = { ok: "bg-emerald-600", alerta: "bg-rose-600", info: "bg-indigo-600" };
  return (
    <div className="fixed top-16 right-4 z-50 flex flex-col gap-2 items-end pointer-events-none">
      {toasts.map((t) => {
        const Icono = t.tipo === "info" ? Bell : t.tipo === "alerta" ? XCircle : CheckCircle2;
        return (
          <div key={t.id} role="status" onClick={() => { t.onClick?.(); onCerrar(t.id); }} className={`pointer-events-auto cursor-pointer ${colores[t.tipo] || colores.ok} text-white text-sm font-medium px-4 py-2.5 rounded-lg shadow-lg flex items-center gap-2 max-w-sm`}>
            <Icono size={16} className="shrink-0" /> <span>{t.mensaje}</span>
          </div>
        );
      })}
    </div>
  );
}

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
// Una FASE (sección) de la solicitud: tarjeta con flecha para expandir o contraer. Contraída conserva el título y la
// descripción (si la tiene) y un resumen corto; el contenido SIGUE montado y solo se oculta, así que nada se pierde.
// Se puede controlar desde afuera (abierto / onToggle) o dejar que maneje su propio estado.
function SeccionColapsable({ titulo, icono: Icono, descripcion, resumen, insignia, acciones, abierto, onToggle, defaultAbierto = true, claseCaja = "bg-white rounded-xl border border-slate-200", children }) {
  const [interno, setInterno] = useState(defaultAbierto);
  const controlado = abierto !== undefined;
  const esAbierto = controlado ? abierto : interno;
  const alternar = controlado && onToggle ? onToggle : () => setInterno((v) => !v);
  return (
    <div className={claseCaja}>
      <div className="flex items-start gap-3 px-5 py-3.5">
        <button type="button" data-fase="true" onClick={alternar} aria-expanded={esAbierto} className="flex-1 min-w-0 flex items-start gap-2 text-left">
          <ChevronRight size={16} className={`mt-0.5 shrink-0 text-slate-400 transition-transform ${esAbierto ? "rotate-90" : ""}`} />
          <span className="min-w-0 flex-1">
            <span className="font-medium text-slate-700 flex items-center gap-2 flex-wrap">{Icono && <Icono size={16} className="shrink-0" />}{titulo}{insignia}</span>
            {descripcion && <span className="block text-xs text-slate-400 mt-0.5 font-normal">{descripcion}</span>}
          </span>
        </button>
        {(resumen || acciones) && <div className="flex items-center gap-2 shrink-0 text-xs text-slate-500">{resumen}{acciones}</div>}
      </div>
      <div className={esAbierto ? "px-5 pb-5 space-y-3" : "hidden"}>{children}</div>
    </div>
  );
}
function ControlExpandirTodo({ n, onTodos, etiquetas = { abrir: "Expandir todo", cerrar: "Contraer todo" } }) {
  if (n <= 2) return null;
  return (
    <div className="flex items-center gap-2 text-[11px] text-slate-500">
      <button type="button" onClick={() => onTodos(true)} className="hover:text-indigo-600 underline">{etiquetas.abrir}</button>
      <span>·</span>
      <button type="button" onClick={() => onTodos(false)} className="hover:text-indigo-600 underline">{etiquetas.cerrar}</button>
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
  const VEINTI = ["", "VEINTIÚN", "VEINTIDÓS", "VEINTITRÉS", "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"];

  function trescientos(num) {
    if (num === 0) return "";
    if (num === 100) return "CIEN";
    let s = "";
    const c = Math.floor(num / 100), resto = num % 100;
    if (c > 0) s += CENTENAS[c] + " ";
    if (resto >= 10 && resto <= 19) { s += DIEZ_A_DIECINUEVE[resto - 10]; }
    else {
      const d = Math.floor(resto / 10), u = resto % 10;
      if (d === 2 && u > 0) s += VEINTI[u];
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
// días transcurridos desde que se creó la solicitud: si ya terminó (completada o rechazada), hasta esa fecha;
// si sigue en curso, hasta ahora. Sirve para mostrar cuánto lleva cada solicitud y para ordenar por las más demoradas.
function diasTranscurridos(s) {
  const inicio = s.historialEstados?.[0]?.fecha || s.fechaCreacion;
  if (!inicio) return 0;
  const FIN_ETAPAS = ["completada", "rechazada"];
  const fin = FIN_ETAPAS.includes(s.status)
    ? [...(s.historialEstados || [])].reverse().find((h) => h.status === s.status)?.fecha || ahoraISO()
    : ahoraISO();
  const ms = new Date(fin) - new Date(inicio);
  return Math.max(0, Math.round(ms / 86400000));
}
// color según cuánto lleva: solo se marca como demorada una solicitud todavía en curso (una ya cerrada no "se demora" más)
function toneDemora(s) {
  if (["completada", "rechazada"].includes(s.status)) return "slate";
  const d = diasTranscurridos(s);
  if (d >= 15) return "red";
  if (d >= 7) return "amber";
  return "slate";
}
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
  if (finalNeg > 0 && inicial > 0 && finalNeg - inicial > 0.0001) errores.push({ tipo: "final", texto: "El precio final negociado no puede ser mayor que el precio inicial cotizado por el proveedor." });
  // el descuento nunca puede igualar o superar el valor inicial: dejaría el precio final en $0 y los totales en blanco
  const desc = parseFloat(cot.descuentoValor) || 0;
  if (inicial > 0 && desc > 0) {
    if (cot.descuentoTipo === "valor" ? desc >= inicial : desc >= 100) {
      errores.push({ tipo: "descuento", texto: cot.descuentoTipo === "valor"
        ? "El descuento no puede ser igual o superior al valor inicial (el precio final quedaría en $0)."
        : "El descuento no puede ser igual o superior al valor inicial: en «Desc. %» escribe un porcentaje menor a 100. Si es un valor en pesos, cambia el tipo a «Desc. $»." });
    }
  }
  if (desc < 0) errores.push({ tipo: "descuento", texto: "El descuento no puede ser negativo." });
  return errores;
}
function itemConPrecioInvalido(item) { return (item.cotizaciones || []).some((c) => erroresPrecioCotizacion(item, c).length > 0); }
function primerErrorPrecio(item) {
  for (const c of item.cotizaciones || []) { const e = erroresPrecioCotizacion(item, c); if (e.length) return e[0].texto; }
  return null;
}
const idCotizacion = () => "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
function tieneAiuValores(aiu) { return !!aiu && !!(parseFloat(aiu.administracionPct) || parseFloat(aiu.utilidadPct) || parseFloat(aiu.imprevistosPct)); }
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
    // el descuento se negocia en la COTIZACIÓN: el precio que estima el solicitante no lleva descuento
    const precioConDescuento = parseFloat(item.precioEstimado) || 0;
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
  if (pagos?.tipoPago === "contado") return [{ tipo: pagos.anticipado ? "Pago anticipado" : "Pago único", ...pagos.pagoUnico }];
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
// ¿falta alguna fecha? solo cuentan los pagos que tienen valor (un plan con solo "pago final" no necesita fecha de anticipo)
function planFaltaFechaPlan(p) {
  if (p?.tipoPago === "contado") return !p.pagoUnico?.fecha;
  const con = (x) => parseFloat(x?.valor) > 0;
  return (con(p?.anticipo) && !p.anticipo.fecha) || (con(p?.final) && !p.final.fecha) || (!!p?.intermedio?.activo && !p.intermedio.fecha);
}
// si solo se llenó el "pago final" y cubre todo el valor, en realidad es un pago único: se cambia solo (con su fecha)
// al cambiar anticipo o intermedio, el pago final se recalcula solo (total − lo demás) para que el plan siempre cuadre
function autoFinalPlanPago(p, total, campo) {
  if (!p || p.tipoPago === "contado" || campo === "final" || !(total > 0)) return p;
  const v = (x) => parseFloat(x?.valor) || 0;
  const resto = Math.round((total - v(p.anticipo) - (p.intermedio?.activo ? v(p.intermedio) : 0)) * 100) / 100;
  if (!(v(p.anticipo) > 0 || (p.intermedio?.activo && v(p.intermedio) > 0))) return p;
  return { ...p, final: { ...p.final, valor: resto > 0 ? resto : "" } };
}
function normalizarPlanPago(p, total) {
  if (!p || p.tipoPago === "contado") return p;
  const v = (x) => parseFloat(x?.valor) || 0;
  if (v(p.anticipo) > 0 || (p.intermedio?.activo && v(p.intermedio) > 0)) return p;
  if (!(v(p.final) > 0) || !(total > 0) || Math.abs(v(p.final) - total) > 0.5) return p;
  return { ...p, tipoPago: "contado", anticipado: false, pagoUnico: { valor: total, fecha: p.final.fecha || "" }, anticipo: { valor: "", fecha: "" }, intermedio: { ...(p.intermedio || {}), activo: false, valor: "", fecha: "" }, final: { valor: "", fecha: "" } };
}
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

// ¿el proveedor del ítem exige pago anticipado? (el elegido; si todavía no hay elegido, cualquiera de los cotizados)
function exigeAnticipoPago(it) {
  const cots = it.cotizaciones || [];
  if (!cots.length) return false;
  // la que aplica: la elegida; si todavía no hay elegida, la que el sistema recomienda (así una segunda cotización SIN anticipo
  // más barata o mejor puntuada no queda amarrada al pago anticipado de otra que no va a ganar)
  const idx = it.cotizacionSeleccionada ?? (cots.length > 1 ? mejorCotizacionIdx(cots, it.cantidad, false, it.aiu) : 0);
  return !!cots[idx]?.anticipoObligatorio;
}
// PAGO ANTICIPADO OBLIGATORIO: el proveedor no despacha hasta recibir el pago. El sistema arma solo el plan de pagos: un único pago
// por el VALOR FINAL NEGOCIADO completo (el total del ítem), antes de despachar. No hay valores que digitar; solo falta la fecha, que
// se toma de un plan que ya existiera (el que sugirió el solicitante, por ejemplo).
function planPagoAnticipado(base, total) {
  const fecha = base?.pagoUnico?.fecha || base?.anticipo?.fecha || "";
  return { ...planPagosVacio(), tipoPago: "contado", anticipado: true, pagoUnico: { valor: total, fecha } };
}
// ¿cuadra el plan de pagos con el total del ítem? Estado y números listos para mostrar
//   sin_plan: nada programado · cuadra · falta: se programó menos del total · sobra: se programó más del total
function estadoCuadre(pagos, total) {
  const hayPlan = planTieneValores(pagos);
  const programado = hayPlan ? totalPagado(pagos) : 0;
  const diferencia = Math.round((total - programado) * 100) / 100;
  const tramos = hayPlan ? tramosDePago(pagos) : [];
  let estado = "sin_plan";
  if (hayPlan) estado = Math.abs(diferencia) <= 0.5 ? "cuadra" : diferencia > 0 ? "falta" : "sobra";
  return { estado, hayPlan, programado, diferencia, total, porcentaje: total > 0 ? Math.round((programado / total) * 100) : 0, tramos, sinFecha: tramos.filter((t) => !t.fecha && (parseFloat(t.valor) > 0 || t.tipo === "Pago único" || t.tipo === "Pago anticipado")).map((t) => t.tipo) };
}
// qué sigue cuando ya están confirmados los planes de pago de todos los ítems, según el paso en que está la solicitud
function siguientePasoPlanes(solicitud, currentUser) {
  const total = totalSolicitud(solicitud);
  switch (solicitud.status) {
    case "aprobacion_director": return { accion: accionLabel(solicitud, total), rol: null, responsable: "el director del área", yo: puedeAprobarDirector(currentUser, solicitud) };
    case "cotizando": case "comparativo": case "orden": return { accion: accionLabel(solicitud, total), rol: "Compras", responsable: "Compras", yo: puedeGestionarCotizaciones(currentUser) };
    case "aprobacion_financiera": return { accion: accionLabel(solicitud, total), rol: "Dirección Financiera", responsable: "Dirección Financiera", yo: puedeAprobarFinanciera(currentUser) };
    case "aprobacion_gerencia": return { accion: accionLabel(solicitud, total), rol: "Gerencia", responsable: "Gerencia", yo: puedeAprobarGerencia(currentUser) };
    default: return { accion: null, rol: null, responsable: null, yo: false }; // enviada, recepción, completada: no hay un paso que dependa de esto
  }
}
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
  { key: "ver_plan_inversion", label: "Ver el Cronograma de pagos" },
  { key: "editar_plan_inversion", label: "Editar el Cronograma de pagos (fechas y montos)" },
  { key: "editar_pagos_finales", label: "Modificar fechas y montos finales de pago después de emitida la orden (Dirección Financiera / Contabilidad)" },
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
// Director responsable de cada área (se asigna en Catálogo → Áreas). Se guarda en memoria como los permisos, para que
// todas las funciones que deciden "quién aprueba" lo consulten sin tener que pasar la lista de áreas a todas partes.
let __directorPorArea = {};
const construirMapaDirectores = (areas) => Object.fromEntries((areas || []).filter((a) => a.directorId).map((a) => [a.id, a.directorId]));
const ROLES_DIRECTOR = ["Director de Área", "Jefe de Área y Director"];
// quiénes pueden ser el director responsable de un área: además de los directores de área, la cabeza de una dirección puede tener
// otro rol (el gerente general tiene rol Gerencia; quien lleva la Dirección Financiera, rol Dirección Financiera)
const ROLES_ASIGNABLES_DIRECTOR = ["Director de Área", "Jefe de Área y Director", "Dirección Financiera", "Gerencia"];
// se muestra en la pantalla para poder confirmar qué versión de la app está publicada
const MODO_DEMO = import.meta.env.VITE_MODO_DEMO === 'true'; // solo se activa en el proyecto de demostración
const VERSION_APP = "novedades v9";
// ¿esta persona es el director de esa área? Con director asignado en el área: solo esa persona. Sin asignar (áreas que
// todavía no lo tienen): el criterio de antes — alguien con rol de director que tenga esa área a cargo.
const esDirectorDeArea = (u, areaId) => (__directorPorArea[areaId] ? __directorPorArea[areaId] === u.id : ROLES_DIRECTOR.includes(u.rol) && tieneAreaACargo(u, areaId));
const directorDelArea = (areaId, usuarios) => (__directorPorArea[areaId] ? usuarios.find((u) => u.id === __directorPorArea[areaId]) || null : usuarios.find((u) => esDirectorDeArea(u, areaId)) || null);
const puedeAprobarDirector = (u, s) => u.rol === "Administrador" || (tienePermiso(u.rol, "aprobar_director") && esDirectorDeArea(u, s.areaId));
// mapeo de la tabla "areas": director_id solo viaja si hay director o si la columna ya existe (así no falla antes de correr el SQL)
const areaDesdeDb = (r) => ({ id: r.id, nombre: r.nombre, presupuesto: r.presupuesto_mensual, directorId: r.director_id || "", _hayColDirector: r.director_id !== undefined });
const areaHaciaDb = (r) => ({ id: r.id, nombre: r.nombre, presupuesto_mensual: Number(r.presupuesto) || 0, ...((r.directorId || r._hayColDirector) ? { director_id: r.directorId || null } : {}) });
const puedeGestionarCotizaciones = (u) => tienePermiso(u.rol, "gestionar_cotizaciones");
const puedeAprobarFinanciera = (u) => tienePermiso(u.rol, "aprobar_financiera");
const puedeAprobarGerencia = (u) => tienePermiso(u.rol, "aprobar_gerencia");
const puedeReabrir = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "reabrir_solicitudes");
// determina en qué paso quedó marcada como rechazada, para poder reabrirla justo ahí
function pasoDelRechazo(solicitud) {
  // paso donde se rechazó: lo guarda el propio rechazo (pasoOrigen); en rechazos anteriores se deduce del historial (el estado
  // inmediatamente anterior a la última "rechazada")
  const hist = solicitud.historialEstados || [];
  let origen = solicitud.ultimoRechazo?.pasoOrigen || null;
  if (!origen) { const i = hist.map((h) => h.status).lastIndexOf("rechazada"); if (i > 0) origen = hist[i - 1].status; }
  if (origen === "aprobacion_jefe") origen = "aprobacion_director"; // el paso del jefe de área ya no existe
  // una firma marcada "no aprobado" dice quién rechazó: ese paso se reabre y esa firma se reinicia
  if (solicitud.firmas?.gerencia?.aprobado === false) return { status: "aprobacion_gerencia", campo: "gerencia" };
  if (solicitud.firmas?.financiera?.aprobado === false) return { status: "aprobacion_financiera", campo: "financiera" };
  if (solicitud.firmas?.director?.aprobado === false) return { status: "aprobacion_director", campo: "director" };
  if (solicitud.firmas?.jefe?.aprobado === false) return { status: "aprobacion_director", campo: "jefe" }; // rechazo hecho en el flujo anterior
  // rechazó Compras (en su revisión, o en cotización / comparativo): vuelve a Compras. Las aprobaciones que ya tenía NO se repiten.
  if (solicitud.revisionCompras?.estado === "rechazada" || origen === "cotizando" || origen === "comparativo") return { status: "cotizando", campo: null, revision: true };
  // rechazada en otro paso conocido (orden, enviada, recepción...): se reabre en ese mismo paso
  if (origen && PASOS.some((p) => p.key === origen) && !["solicitud", "rechazada", "completada"].includes(origen)) return { status: origen, campo: null };
  // sin datos de dónde se rechazó: si ya había pasado la aprobación del director, vuelve a Compras; si no, al director
  if (solicitud.firmas?.director?.aprobado === true || solicitud.revisionCompras?.estado === "aprobada") return { status: "cotizando", campo: null, revision: true };
  return { status: "aprobacion_director", campo: null };
}
const puedeVerCatalogos = (u) => tienePermiso(u.rol, "ver_catalogos");
const puedeVerTodasSolicitudes = (u) => tienePermiso(u.rol, "ver_todas_solicitudes");
const puedeEditarPagos = (u) => tienePermiso(u.rol, "editar_pagos");
const puedeVerMisPendientes = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_mis_pendientes");
const puedeVerCalendarioPagos = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_calendario_pagos");
const puedeVerOrdenesEnviadas = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_ordenes_enviadas");
const puedeVerPlanInversion = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "ver_plan_inversion");
const puedeEditarPlanInversion = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "editar_plan_inversion");
const puedeEditarPagosFinales = (u) => u.rol === "Administrador" || tienePermiso(u.rol, "editar_pagos_finales");
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
      { status: "aprobacion_director", fecha: "2026-07-20T09:05:00.000Z" },
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
    status: "aprobacion_director",
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
      { status: "aprobacion_director", fecha: "2026-07-22T10:02:00.000Z" },
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
  const idx = PASOS_ACTIVOS.findIndex((p) => p.key === status);
  return (
    <div className="flex flex-wrap gap-2">
      {PASOS_ACTIVOS.map((p, i) => {
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
   PALETA DE COLORES (por usuario)
   Cambia el color de acento de TODA la interfaz para quien la elige. No toca los colores de estado (verde / ámbar / rojo) ni los
   documentos que genera la app (PDF, orden impresa, correos): esos conservan siempre el estilo original.
--------------------------------------------------------- */
function SelectorPaleta({ tema, onElegir, compacto }) {
  const actual = sanitizarTema(tema) || TEMA_ORIGINAL;
  const conFondoTinte = actual.fondo !== "neutro";
  // al cambiar de paleta se conserva la preferencia de fondo y bordes (con tinte / neutros)
  const elegir = (t) => onElegir(conFondoTinte ? t : { ...t, fondo: "neutro" });
  const [soportado, setSoportado] = useState(true);
  useEffect(() => { setSoportado(soportaTema()); }, []);
  const [colorLibre, setColorLibre] = useState(actual.paleta === "personalizado" ? actual.color : "#0f766e");
  // color del botón (paso 600) de cada paleta, para dibujar su muestra
  const muestra = (p) => (p.original ? (leerOriginal()["--color-indigo-600"] || "#4f46e5") : generarEscala({ h: p.h, s: p.s })[600].oklch);
  const muestraLibre = () => generarEscala(parametrosDe({ paleta: "personalizado", color: colorLibre }))[600].oklch;
  return (
    <div className="space-y-3">
      <div className={`grid gap-2 ${compacto ? "grid-cols-3" : "grid-cols-3 sm:grid-cols-5"}`}>
        {PALETAS.map((p) => {
          const sel = actual.paleta === p.id;
          return (
            <button key={p.id} type="button" onClick={() => elegir({ paleta: p.id })} aria-pressed={sel} title={p.nombre}
              className={`flex flex-col items-center gap-1 rounded-lg border p-2 text-[11px] ${sel ? "border-slate-700 bg-slate-50 font-semibold text-slate-800" : "border-slate-200 text-slate-600 hover:border-slate-400"}`}>
              <span className="relative w-8 h-8 rounded-full border border-black/10 flex items-center justify-center" style={{ background: muestra(p) }}>{sel && <CheckCircle2 size={16} className="text-white" />}</span>
              <span className="leading-tight text-center">{p.nombre}</span>
            </button>
          );
        })}
      </div>
      <div className={`flex items-center gap-2 rounded-lg border p-2 ${actual.paleta === "personalizado" ? "border-slate-700 bg-slate-50" : "border-slate-200"}`}>
        <label className="flex items-center gap-2 text-[11px] text-slate-600 cursor-pointer">
          <input type="color" value={colorLibre} onChange={(e) => { setColorLibre(e.target.value); elegir({ paleta: "personalizado", color: e.target.value }); }} className="w-8 h-8 rounded border border-slate-200 bg-white p-0.5 cursor-pointer" />
          <span><b>Personalizado</b>{!compacto && <span className="text-slate-400"> — elige cualquier color; el sistema arma el resto de la paleta</span>}</span>
        </label>
        {actual.paleta === "personalizado" && <span className="ml-auto inline-block w-5 h-5 rounded-full border border-black/10 shrink-0" style={{ background: muestraLibre() }} title="Color del botón principal con este color" />}
      </div>
      <label className="flex items-start gap-2 text-[11px] text-slate-600 cursor-pointer">
        <input type="checkbox" checked={conFondoTinte} onChange={(e) => onElegir(e.target.checked ? { ...actual, fondo: undefined } : { ...actual, fondo: "neutro" })} className="mt-0.5" />
        <span><b>Fondo y bordes con el color de la paleta</b>{!compacto && <span className="text-slate-400"> — el fondo de la página y los bordes toman un tinte suave de tu paleta (si lo quitas, quedan grises)</span>}</span>
      </label>
      {!soportado && <div className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-2 py-1.5">Esta instalación no permite cambiar la paleta (los estilos no usan variables de color). Avísale al administrador.</div>}
      {!compacto && (
        <>
          <div className="rounded-lg border border-slate-200 bg-slate-50 p-3 space-y-2">
            <div className="text-[10px] uppercase tracking-wide text-slate-400">Así se verá</div>
            <div className="flex flex-wrap items-center gap-2">
              <button type="button" className="bg-indigo-600 text-white rounded-lg px-3 py-1.5 text-xs font-medium">Botón principal</button>
              <button type="button" className="border border-indigo-200 text-indigo-700 bg-indigo-50 rounded-lg px-3 py-1.5 text-xs">Secundario</button>
              <span className="text-xs text-indigo-600 underline">Enlace</span>
              <span className="text-[10px] rounded-full px-2 py-0.5 bg-indigo-100 text-indigo-700">Etiqueta</span>
            </div>
            <div className="h-2 rounded-full bg-indigo-100"><div className="h-2 w-2/3 rounded-full bg-indigo-500" /></div>
            <div className="rounded-md border border-slate-200 bg-white px-2 py-1.5 text-[11px] text-slate-500">Una tarjeta con su borde, sobre el fondo de la página</div>
          </div>
          <div className="text-[11px] text-slate-400 leading-relaxed">La paleta cambia solo <b>tu</b> pantalla y se guarda en tu cuenta. Los colores de estado (verde, ámbar, rojo) no cambian, y los <b>documentos</b> que genera la app (PDF de órdenes y expedientes, orden impresa, correos) conservan siempre el mismo estilo.</div>
        </>
      )}
      {actual.paleta !== "indigo" && <button type="button" onClick={() => elegir({ paleta: "indigo" })} className="text-[11px] text-slate-500 underline">Volver a la paleta original</button>}
    </div>
  );
}

// botón de acceso rápido, junto a la campana de notificaciones
function BotonPaleta({ tema, onElegir }) {
  const [abierto, setAbierto] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!abierto) return;
    const fuera = (e) => { if (ref.current && !ref.current.contains(e.target)) setAbierto(false); };
    document.addEventListener("mousedown", fuera);
    return () => document.removeEventListener("mousedown", fuera);
  }, [abierto]);
  return (
    <div ref={ref} className="relative z-40">
      <button type="button" title="Paleta de colores" aria-label="Paleta de colores" onClick={() => setAbierto((v) => !v)} className="bg-white border border-slate-200 rounded-full p-2.5 shadow-sm hover:bg-slate-50"><Palette size={18} className="text-slate-600" /></button>
      {abierto && (
        <div className="absolute right-0 mt-2 w-72 bg-white border border-slate-200 rounded-xl shadow-lg p-3 space-y-2">
          <div className="text-xs font-semibold text-slate-700 flex items-center gap-1.5"><Palette size={13} /> Paleta de colores</div>
          <SelectorPaleta tema={tema} onElegir={onElegir} compacto />
        </div>
      )}
    </div>
  );
}


/* ---------------------------------------------------------
   MI PERFIL (firma tipo foto)
--------------------------------------------------------- */
function PerfilUsuario({ currentUser, onGuardar, onCambiarContrasena, tema, onElegirTema }) {
  const [preview, setPreview] = useState(currentUser.firmaFotoUrl);
  const [subiendo, setSubiendo] = useState(false);
  const [clave, setClave] = useState("");
  const [clave2, setClave2] = useState("");
  const [cambiandoClave, setCambiandoClave] = useState(false);
  const [msgClave, setMsgClave] = useState(null); // { ok, texto }
  const cambiarClave = async () => {
    if (clave.length < 6) { setMsgClave({ ok: false, texto: "La contraseña debe tener al menos 6 caracteres." }); return; }
    if (clave !== clave2) { setMsgClave({ ok: false, texto: "Las dos contraseñas no coinciden." }); return; }
    setCambiandoClave(true);
    const error = await onCambiarContrasena(clave);
    setCambiandoClave(false);
    if (error) setMsgClave({ ok: false, texto: `No se pudo cambiar la contraseña: ${error.message}` });
    else { setClave(""); setClave2(""); setMsgClave({ ok: true, texto: "Contraseña actualizada. Úsala la próxima vez que inicies sesión." }); }
  };
  const cargarFoto = async (file) => {
    if (!archivoDentroDelLimite(file)) { alert(`El archivo pesa más de ${TAMANO_MAXIMO_MB} MB. Sube uno más liviano.`); return; }
    setSubiendo(true);
    const ruta = await subirArchivo(file, "firmas");
    setSubiendo(false);
    if (ruta) setPreview(ruta);
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-6 max-w-lg">
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
      {onElegirTema && (
        <div className="mt-6 pt-5 border-t border-slate-100 space-y-2">
          <div className="flex items-center gap-1.5 text-sm font-medium text-slate-700"><Palette size={14} /> Paleta de colores</div>
          <SelectorPaleta tema={tema} onElegir={onElegirTema} />
        </div>
      )}
      {onCambiarContrasena && (
        <div className="mt-6 pt-5 border-t border-slate-100 space-y-2">
          <div className="flex items-center gap-1.5 text-sm font-medium text-slate-700"><Lock size={14} /> Cambiar contraseña</div>
          <div className="text-[11px] text-slate-400">Si tu contraseña inicial es tu documento, te recomendamos cambiarla por una que solo tú conozcas.</div>
          <input type="password" autoComplete="new-password" value={clave} onChange={(e) => setClave(e.target.value)} placeholder="Nueva contraseña (mín. 6 caracteres)" className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
          <input type="password" autoComplete="new-password" value={clave2} onChange={(e) => setClave2(e.target.value)} placeholder="Repite la nueva contraseña" className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm" />
          {msgClave && <div className={`text-xs ${msgClave.ok ? "text-emerald-600" : "text-rose-600"}`}>{msgClave.texto}</div>}
          <button onClick={cambiarClave} disabled={cambiandoClave || !clave} className="bg-slate-700 text-white rounded-lg px-4 py-2 text-sm font-medium disabled:opacity-50">{cambiandoClave ? "Guardando..." : "Cambiar contraseña"}</button>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   DASHBOARD
--------------------------------------------------------- */
function Dashboard({ areas, solicitudes, proveedores, currentUser, onAbrir, onVerCalendario }) {
  const gastoPorArea = useMemo(() => {
    const map = {}; areas.forEach((a) => (map[a.id] = 0));
    solicitudes.forEach((s) => { if (!ESTADOS_SIN_COMPROMISO.includes(s.status)) map[s.areaId] = (map[s.areaId] || 0) + totalSolicitud(s); });
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
            {onVerCalendario && <button onClick={onVerCalendario} className="text-xs text-indigo-600 font-medium">Ver cronograma completo →</button>}
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
                <Bar dataKey="valor" fill="var(--color-indigo-600)" radius={[0, 4, 4, 0]} />
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
   ÓRDENES ENVIADAS A PROVEEDORES — consolidado para contabilidad
--------------------------------------------------------- */
function ReporteOrdenesEnviadas({ solicitudes, proveedores, empresas, onAbrir }) {
  const ESTADO_ETIQUETA = { cancelada: "Cancelada", cambiada: "Cambiada" };
  const [fEmpresa, setFEmpresa] = useState("");
  const [fTipo, setFTipo] = useState("");
  const [fDesde, setFDesde] = useState("");
  const [fHasta, setFHasta] = useState("");
  const [fTexto, setFTexto] = useState("");
  const filasTodas = [];
  solicitudes.forEach((s) => {
    if (!["oc_enviada", "recepcion", "completada"].includes(s.status)) return;
    // facturas / cuentas de cobro y comprobante de recepción (remisión, acta o correo) viven en la recepción de
    // la SOLICITUD (no por proveedor) — con varios proveedores adjudicados, esto aplica al conjunto de la solicitud
    const facturas = s.recepcion?.facturas || (s.recepcion?.facturaUrl ? [{ url: s.recepcion.facturaUrl }] : []);
    const tieneComprobanteRecepcion = (s.recepcion?.archivos?.length > 0) || (s.recepcion?.actas?.length > 0);
    (s.ocEnviada?.ordenesProveedor || []).forEach((o) => {
      if (!o.archivoFirmadoUrl) return;
      const prov = buscarProveedorDeOrden(o, proveedores);
      const cotsAnticipo = s.items.flatMap((it) => (it.cotizaciones || []).filter((c) => c.anticipoObligatorio && mismoProveedor({ proveedorId: c.proveedorId, proveedorNombre: c.proveedorNombre }, o)));
      const itemsDelProveedor = cotsAnticipo;
      filasTodas.push({
        id: `${s.id}-${o.proveedorId || o.proveedorNombre}`,
        solicitudId: s.id,
        folio: s.folio,
        tipo: s.tipo === "compra" ? "Orden de compra (OC)" : "Orden de servicio/trabajo (OS/OT)",
        tipoClave: s.tipo,
        empresaId: s.empresaId,
        empresa: empresas.find((e) => e.id === s.empresaId)?.nombre || "",
        proveedor: o.proveedorNombre,
        fecha: o.fecha,
        total: totalSolicitud(s),
        archivo: o.archivoFirmadoUrl,
        estadoOrden: o.estadoOrden && o.estadoOrden !== "enviada" ? o.estadoOrden : "enviada",
        notaEstadoOrden: o.notaEstadoOrden || "",
        tieneFactura: facturas.length > 0,
        urlsFactura: facturas.map((f) => f.url).filter(Boolean),
        actas: s.recepcion?.actas || [],
        comentarios: comentariosDeOrden(s),
        etiquetaFactura: prov?.personaNatural === "si" ? "Cuenta de cobro" : "Factura",
        tieneComprobanteRecepcion,
        anticipoObligatorio: itemsDelProveedor.length > 0,
      });
    });
  });
  filasTodas.sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")); // más reciente primero
  const sinTild = (t) => String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  const filas = filasTodas.filter((f) => (!fEmpresa || f.empresaId === fEmpresa) && (!fTipo || f.tipoClave === fTipo) && (!fDesde || (f.fecha || "") >= fDesde) && (!fHasta || (f.fecha || "") <= fHasta)
    && (!fTexto.trim() || sinTild(`${f.folio} ${f.proveedor}`).includes(sinTild(fTexto.trim()))));
  const hayFiltro = fEmpresa || fTipo || fDesde || fHasta || fTexto.trim();

  const descargar = () => {
    const encabezado = ["Fecha envío", "Consecutivo", "Tipo", "Empresa", "Proveedor", "Total solicitud", "Estado de la orden", "Factura / cuenta de cobro", "Acta de Recibo", "Pago anticipado obligatorio", "Comentarios"];
    const siNo = (b) => (b ? "Sí" : "Falta");
    const cuerpo = filas.map((f) => [f.fecha, f.folio, f.tipo, f.empresa, f.proveedor, f.total, ESTADO_ETIQUETA[f.estadoOrden] || "Enviada", siNo(f.tieneFactura), siNo(f.tieneComprobanteRecepcion), f.anticipoObligatorio ? "Sí" : "", f.comentarios]);
    const hoja = XLSX.utils.aoa_to_sheet([encabezado, ...cuerpo]);
    hoja["!cols"] = [{ wch: 14 }, { wch: 14 }, { wch: 22 }, { wch: 18 }, { wch: 30 }, { wch: 16 }, { wch: 16 }, { wch: 20 }, { wch: 20 }, { wch: 18 }];
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Órdenes enviadas");
    XLSX.writeFile(libro, `Ordenes_enviadas_${hoy()}.xlsx`);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-end flex-wrap gap-2">
        <button onClick={descargar} disabled={!filas.length} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40 flex items-center gap-1"><FileText size={13} /> Descargar Excel</button>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-3 flex flex-wrap gap-3 items-end no-print">
        <div><label className="text-[11px] font-medium text-slate-500">Empresa</label>
          <select value={fEmpresa} onChange={(e) => setFEmpresa(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm bg-white"><option value="">Todas</option>{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></div>
        <div><label className="text-[11px] font-medium text-slate-500">Tipo</label>
          <select value={fTipo} onChange={(e) => setFTipo(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm bg-white"><option value="">Todos</option><option value="compra">OC (orden de compra)</option><option value="servicio">OS/OT (servicio / trabajo)</option></select></div>
        <div><label className="text-[11px] font-medium text-slate-500">Desde</label><input type="date" value={fDesde} onChange={(e) => setFDesde(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm bg-white" /></div>
        <div><label className="text-[11px] font-medium text-slate-500">Hasta</label><input type="date" value={fHasta} onChange={(e) => setFHasta(e.target.value)} className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm bg-white" /></div>
        <div><label className="text-[11px] font-medium text-slate-500">Consecutivo o proveedor</label><input value={fTexto} onChange={(e) => setFTexto(e.target.value)} placeholder="Buscar..." className="block mt-1 border border-slate-200 rounded-lg px-2 py-1.5 text-sm bg-white w-44" /></div>
        {hayFiltro && <button onClick={() => { setFEmpresa(""); setFTipo(""); setFDesde(""); setFHasta(""); setFTexto(""); }} className="text-xs text-slate-500 underline pb-2">Quitar filtros</button>}
        <div className="text-[11px] text-slate-400 pb-2 ml-auto">{filas.length} de {filasTodas.length} órdenes</div>
      </div>

      {filas.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">{filasTodas.length ? "Ninguna orden coincide con los filtros." : "Todavía no hay órdenes enviadas al proveedor."}</div>
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
              <th className="text-left px-4 py-2 font-medium">Estado de la orden</th>
              <th className="text-center px-4 py-2 font-medium">Factura</th>
              <th className="text-left px-4 py-2 font-medium">Comentarios</th>
              <th className="text-center px-4 py-2 font-medium">Acta de Recibo</th>
              <th className="text-center px-4 py-2 font-medium">OC/OS/OT</th>
            </tr></thead>
            <tbody>{filas.map((f) => (
              <tr key={f.id} className="border-t border-slate-100">
                <td className="px-4 py-2 whitespace-nowrap">{f.fecha}</td>
                <td className="px-4 py-2 font-medium text-slate-700 cursor-pointer hover:text-indigo-600" onClick={() => onAbrir?.(f.solicitudId)}>{f.folio}</td>
                <td className="px-4 py-2 text-slate-600">{f.tipo}</td>
                <td className="px-4 py-2 text-slate-600">{f.empresa}</td>
                <td className="px-4 py-2 text-slate-600">{f.proveedor} {f.anticipoObligatorio && <span title="Pago anticipado obligatorio: se paga el valor final negociado antes de despachar" className="text-amber-600 inline-block align-middle"><Timer size={12} /></span>}</td>
                <td className="px-4 py-2 text-right font-medium">{fmt(f.total)}</td>
                <td className="px-4 py-2">{f.estadoOrden === "enviada" ? <Badge tone="green">Enviada</Badge> : <span title={f.notaEstadoOrden}><Badge tone={f.estadoOrden === "cancelada" ? "red" : "amber"}>{ESTADO_ETIQUETA[f.estadoOrden]}</Badge></span>}</td>
                <td className="px-4 py-2 text-center">{f.tieneFactura ? <div className="flex flex-col items-center gap-0.5">{f.urlsFactura.map((u, k) => <EnlacePrivado key={k} path={u} className="text-indigo-600 underline text-xs">{f.etiquetaFactura}{f.urlsFactura.length > 1 ? ` ${k + 1}` : ""}</EnlacePrivado>)}</div> : <Badge tone="red">Falta {f.etiquetaFactura.toLowerCase()}</Badge>}</td>
                <td className="px-4 py-2 text-slate-600 max-w-[220px] truncate" title={f.comentarios}>{f.comentarios || <span className="text-slate-300 text-xs">—</span>}</td>
                <td className="px-4 py-2 text-center">{f.tieneComprobanteRecepcion ? (f.actas.length ? <div className="flex flex-col items-center gap-0.5">{f.actas.map((u, k) => <EnlacePrivado key={k} path={u} className="text-indigo-600 underline text-xs">Acta de recibo{f.actas.length > 1 ? ` ${k + 1}` : ""}</EnlacePrivado>)}</div> : <Badge tone="green">Sí</Badge>) : <Badge tone="red">Falta</Badge>}</td>
                <td className="px-4 py-2 text-center"><EnlacePrivado path={f.archivo} className="text-indigo-600 underline text-xs">Descargar</EnlacePrivado></td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   CRONOGRAMA DE PAGOS — cronograma editable por proyecto, con
   distribución por periodos (mes/semana), exportable a Excel y PDF.
--------------------------------------------------------- */
const NOMBRES_MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
// reconoce el mes de un texto libre como "Abril" o "abril 2026" (sin acentos, sin importar mayúsculas)
function indiceMesDesdeTexto(texto) {
  const limpio = (texto || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
  return NOMBRES_MESES.findIndex((m) => limpio.startsWith(m.slice(0, 3))); // reconoce "Abril" y también "Abr"
}
// muestra el mes con tres letras ("Abril" → "Abr", "Abril 2026" → "Abr 2026"); si no es un mes reconocible, lo deja igual
function abrevMes(texto) {
  const m = String(texto || "").match(/^(\s*)([A-Za-zÁÉÍÓÚáéíóúñÑ]+)(.*)$/);
  if (!m || indiceMesDesdeTexto(m[2]) < 0) return texto || "";
  return m[2].charAt(0).toUpperCase() + m[2].slice(1, 3).toLowerCase() + m[3];
}
// interpreta un rango de días como "14-18" → {desde: 14, hasta: 18}; si no tiene guion, usa todo el mes
function parsearRangoDias(rango) {
  const m = (rango || "").match(/(\d{1,2})\s*-\s*(\d{1,2})/);
  if (m) return { desde: parseInt(m[1]), hasta: parseInt(m[2]) };
  return { desde: 1, hasta: 31 };
}
// reparte los pagos del plan de una solicitud entre los periodos del cronograma de pagos, según la
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
      periodo = { id: nextId(), mes: abrevMes(nombreMes), etiqueta: "Semana", rango: `${dia}` };
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

// los cronogramas guardados antes con el nombre "Plan de inversión" se muestran como "Cronograma de Pagos"
function tituloCronograma(t) { return !t || /inversi/i.test(t) ? "Cronograma de Pagos" : t; }
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
  const puedeEditar = puedeEditarPlanInversion(currentUser);

  // actualiza la pantalla al instante, y guarda en Supabase con un pequeño retraso (no en cada
  // tecla) para no saturar la base de datos en una tabla con tantos campos
  const actualizarPlan = (cambios) => {
    const actualizado = { ...plan, ...cambios };
    setPlanLocal(actualizado);
    if (guardarTimer.current) clearTimeout(guardarTimer.current);
    guardarTimer.current = setTimeout(() => guardarPlanDB(actualizado), 600);
  };

  const crearPlan = async () => {
    const nuevo = { id: nextId(), titulo: "Cronograma de Pagos", empresaId: empresas[0]?.id || null, anio: new Date().getFullYear(), periodos: [], proyectos: [] };
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
  const alternarSeleccion = (id) => setSe