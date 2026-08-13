import React, { useState, useMemo, useEffect } from "react";
import Papa from "papaparse";
import { useAuth } from "./hooks/useAuth";
import { useSupabaseTable } from "./hooks/useSupabaseTable";
import { useSolicitudes } from "./hooks/useSolicitudes";
import { subirArchivo } from "./lib/storage";
import { enviarCorreo } from "./lib/correo";
import LoginReal from "./LoginReal";
import {
  ShoppingCart, Wrench, Building2, CheckCircle2, XCircle, Clock,
  FileText, TrendingUp, ChevronRight, Plus, Trash2, Pencil,
  Calendar, Award, ArrowLeft, LayoutDashboard, ListChecks, BarChart3,
  DollarSign, PackageCheck, CalendarClock, Boxes, Users, Truck,
  Settings, Target, ClipboardList, Lock, LogOut, History, PenTool, ShieldCheck,
  Paperclip, Mail, Camera, Timer, Layers, MessageSquare, UserCircle, Send, CheckSquare, PanelLeftClose, PanelLeftOpen, Upload as UploadIcon,
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

const UMBRAL_DIRECCION = 3000000;
const UMBRAL_GERENCIA = 12000000;
const UNIDADES = ["unidad", "libra", "kilo", "gramo", "litro", "mililitro", "metro", "caja", "paquete", "hora", "servicio"];
const IVA_OPCIONES = [0, 5, 19];
const COLORS = ["#4f46e5", "#f59e0b", "#10b981", "#ef4444", "#0ea5e9", "#a855f7"];
const ROLES = ["Solicitante", "Jefe de Área", "Dirección Financiera", "Gerencia", "Compras", "Administrador"];

const PASOS = [
  { key: "solicitud", label: "Solicitud creada" },
  { key: "aprobacion_jefe", label: "Aprobación jefe de área" },
  { key: "cotizando", label: "Revisión y cotizaciones (compras)" },
  { key: "comparativo", label: "Cuadro comparativo" },
  { key: "aprobacion_financiera", label: "Dirección financiera" },
  { key: "aprobacion_gerencia", label: "Gerencia" },
  { key: "orden", label: "Orden generada" },
  { key: "oc_enviada", label: "OC enviada al proveedor" },
  { key: "recepcion", label: "Recepción / Ejecución" },
  { key: "completada", label: "Completada" },
];

const fmt = (n) => (n || 0).toLocaleString("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });
const hoy = () => new Date().toISOString().slice(0, 10);
const ahoraISO = () => new Date().toISOString();
let idCounter = 4000;
const nextId = () => (idCounter++).toString();

/* ---------------------------------------------------------
   LÓGICA DE NEGOCIO — dinero
--------------------------------------------------------- */
// usa precio final negociado si existe; si no, el precio inicial cotizado
function precioEquivalente(cot) {
  const factor = parseFloat(cot.factorConversion) || 1;
  const precio = parseFloat(cot.precioFinal || cot.precioUnitario) || 0;
  return precio / factor;
}
function desgloseCotizacion(cot, cantidadSolicitada) {
  const subtotal = precioEquivalente(cot) * (parseFloat(cantidadSolicitada) || 0);
  const ivaPct = parseFloat(cot.ivaPct ?? 19) || 0;
  const iva = subtotal * (ivaPct / 100);
  return { subtotal, iva, total: subtotal + iva };
}
function calcularScores(cotizaciones, cantidadSolicitada) {
  if (!cotizaciones.length) return [];
  const desgloses = cotizaciones.map((c) => desgloseCotizacion(c, cantidadSolicitada));
  const totales = desgloses.map((d) => d.total);
  const entregas = cotizaciones.map((c) => parseFloat(c.diasEntrega) || 0);
  const condiciones = cotizaciones.map((c) => parseFloat(c.condicionesScore) || 0);
  const minTotal = Math.min(...totales), maxTotal = Math.max(...totales);
  const minEnt = Math.min(...entregas), maxEnt = Math.max(...entregas);
  const minCond = Math.min(...condiciones), maxCond = Math.max(...condiciones);
  return cotizaciones.map((c, i) => {
    const precioScore = maxTotal === minTotal ? 1 : (maxTotal - totales[i]) / (maxTotal - minTotal);
    const entregaScore = maxEnt === minEnt ? 1 : (maxEnt - entregas[i]) / (maxEnt - minEnt);
    const condScore = maxCond === minCond ? 1 : (condiciones[i] - minCond) / (maxCond - minCond);
    return { ...c, ...desgloses[i], score: precioScore * 0.6 + entregaScore * 0.25 + condScore * 0.15 };
  });
}
function mejorCotizacionIdx(cotizaciones, cantidadSolicitada) {
  const scored = calcularScores(cotizaciones, cantidadSolicitada);
  if (!scored.length) return -1;
  let best = 0;
  scored.forEach((s, i) => { if (s.score > scored[best].score) best = i; });
  return best;
}
function desgloseItem(item) {
  if (!item.cotizaciones.length) {
    const subtotal = (parseFloat(item.precioEstimado) || 0) * (parseFloat(item.cantidad) || 0);
    const ivaPct = parseFloat(item.ivaEstimado ?? 19) || 0;
    const iva = subtotal * (ivaPct / 100);
    return { subtotal, iva, total: subtotal + iva };
  }
  const idx = item.cotizacionSeleccionada ?? mejorCotizacionIdx(item.cotizaciones, item.cantidad);
  const cot = item.cotizaciones[idx];
  return cot ? desgloseCotizacion(cot, item.cantidad) : { subtotal: 0, iva: 0, total: 0 };
}
function desgloseSolicitud(solicitud) {
  return solicitud.items.reduce((acc, item) => {
    const d = desgloseItem(item);
    return { subtotal: acc.subtotal + d.subtotal, iva: acc.iva + d.iva, total: acc.total + d.total };
  }, { subtotal: 0, iva: 0, total: 0 });
}
function totalSolicitud(s) { return desgloseSolicitud(s).total; }
function requiereDireccion(m) { return m >= UMBRAL_DIRECCION; }
function requiereGerencia(m) { return m >= UMBRAL_GERENCIA; }
function totalPagado(pagos) {
  return (parseFloat(pagos?.anticipo?.valor) || 0) + (pagos?.intermedio?.activo ? (parseFloat(pagos.intermedio.valor) || 0) : 0) + (parseFloat(pagos?.final?.valor) || 0);
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
   PERMISOS
--------------------------------------------------------- */
const puedeAprobarJefe = (u, s) => u.rol === "Administrador" || (u.rol === "Jefe de Área" && u.areaId === s.areaId);
const puedeGestionarCotizaciones = (u) => u.rol === "Administrador" || u.rol === "Compras";
const puedeAprobarFinanciera = (u) => u.rol === "Administrador" || u.rol === "Dirección Financiera";
const puedeAprobarGerencia = (u) => u.rol === "Administrador" || u.rol === "Gerencia";
const puedeVerCatalogos = (u) => u.rol === "Administrador" || ["Compras", "Dirección Financiera", "Gerencia"].includes(u.rol);
const puedeVerTodasSolicitudes = (u) => u.rol === "Administrador" || u.rol !== "Solicitante";
const puedeEditarPagos = (u) => u.rol === "Administrador" || u.rol === "Dirección Financiera";
const puedeVerHistorico = (u) => u.rol === "Administrador" || u.rol === "Compras";

/* ---------------------------------------------------------
   DATOS SEMILLA
--------------------------------------------------------- */
function planPagosVacio() { return { anticipo: { valor: "", fecha: "" }, intermedio: { activo: false, valor: "", fecha: "" }, final: { valor: "", fecha: "" } }; }

function datosSemilla() {
  const s1 = {
    id: nextId(), folio: "SOL-1001", tipo: "compra", empresaId: "emp1", areaId: "produccion",
    solicitanteId: "u1", fechaCreacion: "2026-07-20", fechaEstimada: "2026-08-05",
    objetivo: "Garantizar el abastecimiento de sal industrial para el proceso de tinturado.",
    justificacion: "El inventario actual cubre solo 5 días de producción; se requiere reposición para no detener la línea.",
    centroCostoId: "cc2", conceptoGastoId: "cg1",
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
    ocEnviada: { archivoNombre: "", fecha: "", usuario: "" },
    recepcion: { archivoNombre: "", comentario: "", recibidoSatisfaccion: false, usuario: "", fecha: "" },
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
    centroCostoId: "cc4", conceptoGastoId: "cg4",
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
    ocEnviada: { archivoNombre: "", fecha: "", usuario: "" },
    recepcion: { archivoNombre: "", comentario: "", recibidoSatisfaccion: false, usuario: "", fecha: "" },
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
function Badge({ children, tone = "slate" }) {
  const tones = {
    slate: "bg-slate-100 text-slate-700 border-slate-200", green: "bg-emerald-50 text-emerald-700 border-emerald-200",
    amber: "bg-amber-50 text-amber-700 border-amber-200", red: "bg-rose-50 text-rose-700 border-rose-200", blue: "bg-blue-50 text-blue-700 border-blue-200",
  };
  return <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium border ${tones[tone]}`}>{children}</span>;
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
        <img src={firma.fotoUrl} alt={`firma ${rol}`} className="h-10 object-contain mb-1" />
      ) : (
        <div className="font-serif italic text-slate-700 text-base border-b border-slate-300 pb-1 mb-1">{firma.nombre}</div>
      )}
      <div className="text-[11px] text-slate-400">{firma.nombre} · {firma.fecha}{firma.aprobado === false ? " · Rechazado" : firma.aprobado ? " · Aprobado" : ""}</div>
      {firma.observacion && <div className="text-xs text-slate-500 mt-1 italic">"{firma.observacion}"</div>}
    </div>
  );
}

// input de archivo: sube de verdad a Supabase Storage y guarda la URL pública resultante
function AdjuntarArchivo({ nombre, onSeleccionar, label, small, carpeta }) {
  const [subiendo, setSubiendo] = useState(false);
  const manejar = async (file) => {
    setSubiendo(true);
    const url = await subirArchivo(file, carpeta || "adjuntos");
    setSubiendo(false);
    if (url) onSeleccionar(url);
  };
  // si "nombre" es una URL de Storage, mostramos solo el nombre del archivo (sin el prefijo de fecha)
  const mostrar = nombre ? decodeURIComponent(nombre.split("/").pop().replace(/^\d+_/, "")) : null;
  return (
    <span className="inline-flex items-center gap-1.5">
      <label className={`inline-flex items-center gap-1.5 border border-dashed border-slate-300 rounded-md px-2 py-1 cursor-pointer text-slate-500 hover:border-indigo-400 hover:text-indigo-600 ${small ? "text-[11px]" : "text-xs"}`}>
        <Paperclip size={small ? 11 : 13} />
        {subiendo ? <span>Subiendo...</span> : mostrar ? <span className="truncate max-w-[120px]">{mostrar}</span> : <span>{label || "Adjuntar archivo"}</span>}
        <input type="file" accept=".pdf,image/*" className="hidden" disabled={subiendo} onChange={(e) => e.target.files[0] && manejar(e.target.files[0])} />
      </label>
      {nombre && !subiendo && <a href={nombre} target="_blank" rel="noreferrer" className={`text-indigo-600 underline ${small ? "text-[11px]" : "text-xs"}`}>ver</a>}
    </span>
  );
}

function CrudTable({ titulo, icon: Icon, columnas, datos, onGuardar, onEliminar, plantilla }) {
  const [editId, setEditId] = useState(null);
  const [form, setForm] = useState(plantilla);
  const [creando, setCreando] = useState(false);
  const [mensajeImport, setMensajeImport] = useState("");
  const iniciarEdicion = (fila) => { setEditId(fila.id); setForm(fila); setCreando(false); };
  const iniciarCreacion = () => { setEditId(null); setForm(plantilla); setCreando(true); };
  const guardar = () => { onGuardar(editId ? { ...form, id: editId } : { ...form, id: nextId() }); setEditId(null); setCreando(false); setForm(plantilla); };
  const cancelar = () => { setEditId(null); setCreando(false); setForm(plantilla); };
  const Campo = (c) => c.type === "select" ? (
    <select value={form[c.key] || ""} onChange={(e) => setForm({ ...form, [c.key]: e.target.value })} className="border border-slate-200 rounded-md px-2 py-1 text-xs w-full">
      <option value="">—</option>{c.options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  ) : (
    <input type={c.type || "text"} value={form[c.key] || ""} onChange={(e) => setForm({ ...form, [c.key]: e.target.value })} className="border border-slate-200 rounded-md px-2 py-1 text-xs w-full" />
  );

  const descargarPlantilla = () => {
    const csv = columnas.map((c) => c.key).join(",") + "\n";
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `plantilla_${titulo.toLowerCase().replace(/\s+/g, "_")}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  const importarCSV = (file) => {
    Papa.parse(file, {
      header: true, skipEmptyLines: true,
      complete: (res) => {
        let importados = 0;
        res.data.forEach((fila) => {
          const nueva = { id: nextId() };
          columnas.forEach((c) => { nueva[c.key] = (fila[c.key] ?? "").toString().trim(); });
          if (Object.values(nueva).some((v) => v && v !== nueva.id)) { onGuardar(nueva); importados++; }
        });
        setMensajeImport(`${importados} registro(s) importado(s) correctamente.`);
        setTimeout(() => setMensajeImport(""), 4000);
      },
      error: () => { setMensajeImport("No se pudo leer el archivo CSV."); setTimeout(() => setMensajeImport(""), 4000); },
    });
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
      <div className="px-5 py-3 border-b border-slate-100 flex items-center justify-between flex-wrap gap-2">
        <div className="flex items-center gap-2 font-medium text-slate-700"><Icon size={16} /> {titulo}</div>
        <div className="flex items-center gap-2">
          <button onClick={descargarPlantilla} className="text-xs text-slate-500 font-medium flex items-center gap-1"><FileText size={12} /> Plantilla CSV</button>
          <label className="text-xs text-indigo-600 font-medium flex items-center gap-1 cursor-pointer"><UploadIcon size={12} /> Importar CSV
            <input type="file" accept=".csv" className="hidden" onChange={(e) => e.target.files[0] && importarCSV(e.target.files[0])} />
          </label>
          {!creando && <button onClick={iniciarCreacion} className="text-xs text-indigo-600 font-medium flex items-center gap-1"><Plus size={13} /> Agregar</button>}
        </div>
      </div>
      {mensajeImport && <div className="px-5 py-1.5 text-[11px] text-emerald-600 bg-emerald-50 border-b border-emerald-100">{mensajeImport}</div>}
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-slate-500 text-xs"><tr>{columnas.map((c) => <th key={c.key} className="text-left px-4 py-2 font-medium">{c.label}</th>)}<th></th></tr></thead>
        <tbody>
          {creando && (
            <tr className="border-t border-slate-100 bg-indigo-50/40">
              {columnas.map((c) => <td key={c.key} className="px-4 py-1.5">{Campo(c)}</td>)}
              <td className="px-4 py-1.5 text-right whitespace-nowrap"><button onClick={guardar} className="text-emerald-600 text-xs font-medium mr-2">Guardar</button><button onClick={cancelar} className="text-slate-400 text-xs">Cancelar</button></td>
            </tr>
          )}
          {datos.map((fila) => editId === fila.id ? (
            <tr key={fila.id} className="border-t border-slate-100 bg-indigo-50/40">
              {columnas.map((c) => <td key={c.key} className="px-4 py-1.5">{Campo(c)}</td>)}
              <td className="px-4 py-1.5 text-right whitespace-nowrap"><button onClick={guardar} className="text-emerald-600 text-xs font-medium mr-2">Guardar</button><button onClick={cancelar} className="text-slate-400 text-xs">Cancelar</button></td>
            </tr>
          ) : (
            <tr key={fila.id} className="border-t border-slate-100">
              {columnas.map((c) => <td key={c.key} className="px-4 py-2 text-slate-600">{c.type === "select" ? (c.options.find((o) => o.value === fila[c.key])?.label || "—") : (fila[c.key] || "—")}</td>)}
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
    setSubiendo(true);
    const url = await subirArchivo(file, "firmas");
    setSubiendo(false);
    if (url) setPreview(url);
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-6 max-w-md">
      <div className="flex items-center gap-2 mb-4"><UserCircle size={18} className="text-indigo-600" /><h2 className="text-lg font-semibold text-slate-800">Mi perfil</h2></div>
      <div className="text-sm text-slate-600 mb-1"><b>{currentUser.nombre}</b></div>
      <div className="text-xs text-slate-400 mb-4">{currentUser.cargo} · {currentUser.rol}</div>
      <label className="text-xs font-medium text-slate-500">Firma (foto)</label>
      <div className="border border-dashed border-slate-300 rounded-lg p-4 mt-1 text-center">
        {preview ? <img src={preview} alt="firma" className="h-20 mx-auto object-contain mb-2" /> : <div className="text-xs text-slate-400 mb-2">Sin firma cargada. Sube una foto de tu firma en papel.</div>}
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
function Dashboard({ areas, solicitudes }) {
  const gastoPorArea = useMemo(() => {
    const map = {}; areas.forEach((a) => (map[a.id] = 0));
    solicitudes.forEach((s) => { if (!["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)) map[s.areaId] = (map[s.areaId] || 0) + totalSolicitud(s); });
    return map;
  }, [areas, solicitudes]);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="flex items-center gap-2 text-slate-500 text-sm mb-1"><FileText size={16} /> Solicitudes activas</div><div className="text-2xl font-semibold text-slate-800">{solicitudes.filter((s) => s.status !== "completada" && s.status !== "rechazada").length}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="flex items-center gap-2 text-slate-500 text-sm mb-1"><Clock size={16} /> Pendientes de aprobación</div><div className="text-2xl font-semibold text-slate-800">{solicitudes.filter((s) => ["aprobacion_jefe", "aprobacion_financiera", "aprobacion_gerencia"].includes(s.status)).length}</div></div>
        <div className="bg-white rounded-xl border border-slate-200 p-5"><div className="flex items-center gap-2 text-slate-500 text-sm mb-1"><DollarSign size={16} /> Comprometido este mes (con IVA)</div><div className="text-2xl font-semibold text-slate-800">{fmt(Object.values(gastoPorArea).reduce((a, b) => a + b, 0))}</div></div>
      </div>
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
function Estadisticas({ solicitudes, areas, empresas, proveedores }) {
  const [filtroEmpresa, setFiltroEmpresa] = useState("todas");
  const base = filtroEmpresa === "todas" ? solicitudes : solicitudes.filter((s) => s.empresaId === filtroEmpresa);

  const porArea = areas.map((a) => ({ nombre: a.nombre, monto: base.filter((s) => s.areaId === a.id && !["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)).reduce((acc, s) => acc + totalSolicitud(s), 0) }));
  const porTipo = [{ nombre: "Compra", value: base.filter((s) => s.tipo === "compra").length }, { nombre: "Servicio", value: base.filter((s) => s.tipo === "servicio").length }];
  const porEstado = PASOS.map((p) => ({ nombre: p.label, value: base.filter((s) => s.status === p.key).length })).filter((e) => e.value > 0);
  const porEmpresaComparativo = empresas.map((e) => ({ nombre: e.nombre, monto: solicitudes.filter((s) => s.empresaId === e.id && !["solicitud", "aprobacion_jefe", "rechazada"].includes(s.status)).reduce((acc, s) => acc + totalSolicitud(s), 0) }));
  const proveedorMonto = {};
  base.forEach((s) => s.items.forEach((it) => {
    if (!it.cotizaciones.length) return;
    const idx = it.cotizacionSeleccionada ?? mejorCotizacionIdx(it.cotizaciones, it.cantidad);
    const cot = it.cotizaciones[idx]; if (!cot) return;
    const prov = proveedores.find((p) => p.id === cot.proveedorId)?.nombre || "—";
    proveedorMonto[prov] = (proveedorMonto[prov] || 0) + desgloseCotizacion(cot, it.cantidad).total;
  }));
  const porProveedor = Object.entries(proveedorMonto).map(([nombre, monto]) => ({ nombre, monto })).sort((a, b) => b.monto - a.monto);
  const totalGeneral = base.reduce((acc, s) => acc + totalSolicitud(s), 0);
  const ivaGeneral = base.reduce((acc, s) => acc + desgloseSolicitud(s).iva, 0);

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div className="text-sm font-medium text-slate-700">Filtrar por empresa</div>
        <select value={filtroEmpresa} onChange={(e) => setFiltroEmpresa(e.target.value)} className="border border-slate-200 rounded-lg px-3 py-1.5 text-sm">
          <option value="todas">Todas las empresas</option>{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}
        </select>
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
function NuevaSolicitud({ areas, empresas, itemsCatalogo, centrosCosto, conceptosGasto, usuarios, currentUser, onCrear, onCancel }) {
  const [tipo, setTipo] = useState("compra");
  const [empresaId, setEmpresaId] = useState(empresas[0]?.id || "");
  const [areaId, setAreaId] = useState(currentUser.areaId || areas[0].id);
  const [centroCostoId, setCentroCostoId] = useState(centrosCosto[0]?.id || "");
  const [conceptoGastoId, setConceptoGastoId] = useState(conceptosGasto[0]?.id || "");
  const [fechaEstimada, setFechaEstimada] = useState("");
  const [objetivo, setObjetivo] = useState("");
  const [justificacion, setJustificacion] = useState("");
  const [items, setItems] = useState([{ id: nextId(), itemCatalogoId: "", nombre: "", cantidad: 1, unidad: "unidad", precioEstimado: "", ivaEstimado: 19, cotizaciones: [] }]);
  const [pagosSugeridos, setPagosSugeridos] = useState(planPagosVacio());

  const addItem = () => setItems([...items, { id: nextId(), itemCatalogoId: "", nombre: "", cantidad: 1, unidad: "unidad", precioEstimado: "", ivaEstimado: 19, cotizaciones: [] }]);
  const removeItem = (id) => setItems(items.filter((i) => i.id !== id));
  const updateItem = (id, field, val) => setItems(items.map((i) => (i.id === id ? { ...i, [field]: val } : i)));
  const elegirDelCatalogo = (id, catalogoId) => {
    const cat = itemsCatalogo.find((c) => c.id === catalogoId);
    setItems(items.map((i) => (i.id === id ? (cat ? { ...i, itemCatalogoId: catalogoId, nombre: cat.nombre, unidad: cat.unidadDefault } : { ...i, itemCatalogoId: "" }) : i)));
  };
  const setCotizacionesItem = (itemId, cots) => setItems(items.map((i) => (i.id === itemId ? { ...i, cotizaciones: cots } : i)));

  const submit = () => {
    if (!items.length || items.some((i) => !i.nombre.trim()) || !objetivo.trim() || !justificacion.trim()) return;
    const jefe = usuarios.find((u) => u.areaId === areaId && u.rol === "Jefe de Área");
    const folio = "SOL-" + (1000 + Math.floor(Math.random() * 8999));
    onCrear({
      id: nextId(), folio,
      tipo, empresaId, areaId, centroCostoId, conceptoGastoId, solicitanteId: currentUser.id,
      fechaCreacion: hoy(), fechaEstimada, objetivo, justificacion,
      status: "aprobacion_jefe",
      revisionCompras: tipo === "compra" ? { estado: "pendiente", observacion: "", usuario: "", fecha: "" } : { estado: "no_aplica", observacion: "", usuario: "", fecha: "" },
      items: items.map((i) => ({ ...i, cotizacionSeleccionada: null, observacionSeleccion: "" })),
      firmas: {
        solicitante: { nombre: currentUser.nombre, fecha: hoy(), fotoUrl: currentUser.firmaFotoUrl || null },
        jefe: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
        financiera: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
        gerencia: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null },
      },
      pagosSugeridos, pagos: planPagosVacio(), pagosConfirmados: false,
      ocEnviada: { archivoNombre: "", fecha: "", usuario: "" },
      recepcion: { archivoNombre: "", comentario: "", recibidoSatisfaccion: false, usuario: "", fecha: "" },
      historialEstados: [{ status: "solicitud", fecha: ahoraISO() }, { status: "aprobacion_jefe", fecha: ahoraISO() }],
      notificaciones: [{ fecha: ahoraISO(), mensaje: jefe?.email ? `Correo enviado a ${jefe.nombre} (${jefe.email})` : "Solicitud creada. No hay un jefe de área con correo configurado para notificar." }],
    });
    if (jefe?.email) {
      enviarCorreo(
        jefe.email,
        `Nueva solicitud pendiente: ${folio}`,
        `<p>Hola ${jefe.nombre},</p><p><b>${currentUser.nombre}</b> creó la solicitud <b>${folio}</b> (${tipo === "compra" ? "Orden de compra" : "Orden de servicio"}) y quedó pendiente de tu aprobación.</p><p><b>Objetivo:</b> ${objetivo}</p>`
      );
    }
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-6 max-w-3xl">
      <h2 className="text-lg font-semibold text-slate-800 mb-5">Nueva solicitud</h2>
      <div className="grid grid-cols-2 gap-4 mb-5">
        <div>
          <label className="text-xs font-medium text-slate-500">Tipo de solicitud</label>
          <div className="flex gap-2 mt-1">
            <button onClick={() => setTipo("compra")} className={`flex-1 flex items-center justify-center gap-2 px-3 py-2 rounded-lg border text-sm font-medium ${tipo === "compra" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}><ShoppingCart size={15} /> Orden de compra</button>
            <button onClick={() => setTipo("servicio")} className={`flex-1 flex items-center justify-center gap-2 px-3 py-2 rounded-lg border text-sm font-medium ${tipo === "servicio" ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}><Wrench size={15} /> Orden de servicio</button>
          </div>
        </div>
        <div><label className="text-xs font-medium text-slate-500">Empresa</label><select value={empresaId} onChange={(e) => setEmpresaId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Área solicitante</label><select value={areaId} onChange={(e) => setAreaId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{areas.map((a) => <option key={a.id} value={a.id}>{a.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Solicitante</label><div className="w-full mt-1 border border-slate-100 bg-slate-50 rounded-lg px-3 py-2 text-sm text-slate-500">{currentUser.nombre} (firma automática)</div></div>
        <div><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><Layers size={12} /> Centro de costo</label><select value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{centrosCosto.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Concepto de gasto</label><select value={conceptoGastoId} onChange={(e) => setConceptoGastoId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{conceptosGasto.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500">{tipo === "compra" ? "Fecha estimada de entrega" : "Fecha estimada de terminación"}</label><input type="date" value={fechaEstimada} onChange={(e) => setFechaEstimada(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm" /></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><Target size={12} /> Objetivo</label><textarea value={objetivo} onChange={(e) => setObjetivo(e.target.value)} rows={2} placeholder="¿Qué se busca lograr con esta solicitud?" className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" /></div>
        <div className="col-span-2"><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><ClipboardList size={12} /> Justificación</label><textarea value={justificacion} onChange={(e) => setJustificacion(e.target.value)} rows={2} placeholder="¿Por qué es necesaria?" className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" /></div>
      </div>

      <div className="mb-2 flex items-center justify-between"><label className="text-xs font-medium text-slate-500">Ítems solicitados</label><button onClick={addItem} className="text-xs text-indigo-600 font-medium flex items-center gap-1"><Plus size={13} /> Agregar ítem</button></div>
      <div className="space-y-2 mb-5">
        {items.map((it) => (
          <div key={it.id} className="bg-slate-50 rounded-lg p-2 space-y-1.5">
            <select value={it.itemCatalogoId} onChange={(e) => elegirDelCatalogo(it.id, e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs">
              <option value="">Ítem libre (escribir abajo)</option>{itemsCatalogo.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
            </select>
            <div className="flex gap-2 items-start">
              <input placeholder="Descripción del ítem" value={it.nombre} onChange={(e) => updateItem(it.id, "nombre", e.target.value)} className="flex-[2] border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              <input type="number" min="0" placeholder="Cant." value={it.cantidad} onChange={(e) => updateItem(it.id, "cantidad", e.target.value)} className="w-16 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              <select value={it.unidad} onChange={(e) => updateItem(it.id, "unidad", e.target.value)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm">{UNIDADES.map((u) => <option key={u} value={u}>{u}</option>)}</select>
              <input type="number" min="0" placeholder="Precio est." value={it.precioEstimado} onChange={(e) => updateItem(it.id, "precioEstimado", e.target.value)} className="w-24 border border-slate-200 rounded-md px-2 py-1.5 text-sm" />
              <select value={it.ivaEstimado} onChange={(e) => updateItem(it.id, "ivaEstimado", e.target.value)} className="w-20 border border-slate-200 rounded-md px-2 py-1.5 text-sm">{IVA_OPCIONES.map((v) => <option key={v} value={v}>IVA {v}%</option>)}</select>
              {items.length > 1 && <button onClick={() => removeItem(it.id)} className="text-slate-400 hover:text-rose-500 p-1.5"><Trash2 size={15} /></button>}
            </div>
            {/* el solicitante puede adjuntar hasta 3 cotizaciones desde ya, opcional */}
            <CotizacionForm item={it} proveedores={[]} onGuardar={(_, cots) => setCotizacionesItem(it.id, cots)} compacto opcionalTitulo="Adjuntar cotizaciones (opcional, máx. 3)" />
          </div>
        ))}
      </div>

      <div className="mb-5 bg-slate-50 rounded-lg p-3 border border-slate-200">
        <div className="text-xs font-medium text-slate-500 mb-2 flex items-center gap-1"><CalendarClock size={13} /> Plan de pagos sugerido (opcional — Dirección Financiera lo confirmará o ajustará)</div>
        <div className="grid grid-cols-3 gap-2">
          <div><input type="number" placeholder="Anticipo" value={pagosSugeridos.anticipo.valor} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, anticipo: { ...pagosSugeridos.anticipo, valor: e.target.value } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1" /><input type="date" value={pagosSugeridos.anticipo.fecha} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, anticipo: { ...pagosSugeridos.anticipo, fecha: e.target.value } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs" /></div>
          <div><label className="text-[11px] flex items-center gap-1 mb-1"><input type="checkbox" checked={pagosSugeridos.intermedio.activo} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, intermedio: { ...pagosSugeridos.intermedio, activo: e.target.checked } })} /> Intermedio</label><input type="number" placeholder="Valor" disabled={!pagosSugeridos.intermedio.activo} value={pagosSugeridos.intermedio.valor} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, intermedio: { ...pagosSugeridos.intermedio, valor: e.target.value } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><input type="date" disabled={!pagosSugeridos.intermedio.activo} value={pagosSugeridos.intermedio.fecha} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, intermedio: { ...pagosSugeridos.intermedio, fecha: e.target.value } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
          <div><input type="number" placeholder="Pago final" value={pagosSugeridos.final.valor} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, final: { ...pagosSugeridos.final, valor: e.target.value } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1" /><input type="date" value={pagosSugeridos.final.fecha} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, final: { ...pagosSugeridos.final, fecha: e.target.value } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs" /></div>
        </div>
      </div>

      <div className="flex gap-2 justify-end">
        <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm text-slate-500 border border-slate-200">Cancelar</button>
        <button onClick={submit} className="px-4 py-2 rounded-lg text-sm bg-indigo-600 text-white font-medium">Enviar solicitud</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   COTIZACIONES Y COMPARATIVO
--------------------------------------------------------- */
function CotizacionForm({ item, proveedores, onGuardar, compacto, opcionalTitulo }) {
  const [abierto, setAbierto] = useState(!compacto);
  const [cots, setCots] = useState(item.cotizaciones.length ? item.cotizaciones : []);
  const [guardadoMsg, setGuardadoMsg] = useState(false);
  const update = (i, field, val) => { const copy = [...cots]; copy[i] = { ...copy[i], [field]: val }; setCots(copy); };
  const addCot = () => cots.length < 3 && setCots([...cots, { proveedorId: "", proveedorNombre: "", unidadCotizada: item.unidad, factorConversion: 1, precioUnitario: "", precioFinal: "", diasEntrega: "", condicionesScore: 5, ivaPct: item.ivaEstimado ?? 19, archivoNombre: "" }]);
  const removeCot = (i) => setCots(cots.filter((_, idx) => idx !== i));
  const guardar = () => { onGuardar(item.id, cots.filter((c) => (c.proveedorId || c.proveedorNombre) && c.precioUnitario)); setGuardadoMsg(true); setTimeout(() => setGuardadoMsg(false), 2500); };

  if (compacto && !abierto) return <button onClick={() => setAbierto(true)} className="text-[11px] text-indigo-600 font-medium flex items-center gap-1"><Paperclip size={11} /> {opcionalTitulo || "Adjuntar cotizaciones"}</button>;

  return (
    <div className="bg-slate-50 rounded-lg p-3 border border-slate-200">
      <div className="flex items-center justify-between mb-2"><div className="text-sm font-medium text-slate-700">{item.nombre} <span className="text-slate-400 font-normal">({item.cantidad} {item.unidad})</span></div>{cots.length < 3 && <button onClick={addCot} className="text-xs text-indigo-600 flex items-center gap-1"><Plus size={12} /> Cotización</button>}</div>
      <div className="space-y-2">
        {cots.map((c, i) => { const d = desgloseCotizacion(c, item.cantidad); return (
          <div key={i} className="space-y-1 border-b border-slate-200 pb-2 last:border-0 last:pb-0">
            <div className="grid grid-cols-8 gap-1.5 items-center">
              {proveedores.length > 0 ? (
                <select value={c.proveedorId} onChange={(e) => update(i, "proveedorId", e.target.value)} className="col-span-2 border border-slate-200 rounded-md px-2 py-1.5 text-xs"><option value="">Proveedor...</option>{proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}</select>
              ) : (
                <input placeholder="Proveedor" value={c.proveedorNombre || ""} onChange={(e) => update(i, "proveedorNombre", e.target.value)} className="col-span-2 border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
              )}
              <select value={c.unidadCotizada} onChange={(e) => update(i, "unidadCotizada", e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">{UNIDADES.map((u) => <option key={u} value={u}>{u}</option>)}</select>
              <input type="number" placeholder="Factor" title={`¿A cuántas ${item.unidad} equivale 1 ${c.unidadCotizada}?`} value={c.factorConversion} onChange={(e) => update(i, "factorConversion", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
              <input type="number" placeholder="Precio inicial" value={c.precioUnitario} onChange={(e) => update(i, "precioUnitario", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
              <input type="number" placeholder="Precio final neg." value={c.precioFinal} onChange={(e) => update(i, "precioFinal", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
              <select value={c.ivaPct} onChange={(e) => update(i, "ivaPct", e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">{IVA_OPCIONES.map((v) => <option key={v} value={v}>IVA {v}%</option>)}</select>
              <div className="flex gap-1"><input type="number" placeholder="Días" value={c.diasEntrega} onChange={(e) => update(i, "diasEntrega", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs" />{cots.length > 1 && <button onClick={() => removeCot(i)} className="text-slate-400 hover:text-rose-500"><Trash2 size={13} /></button>}</div>
            </div>
            <div className="flex items-center justify-between">
              <AdjuntarArchivo small nombre={c.archivoNombre} label="Adjuntar cotización (PDF/foto)" onSeleccionar={(n) => update(i, "archivoNombre", n)} />
              {(c.proveedorId || c.proveedorNombre) && c.precioUnitario && <div className="text-[11px] text-slate-500">Subtotal: {fmt(d.subtotal)} · IVA: {fmt(d.iva)} · <b>Total: {fmt(d.total)}</b></div>}
            </div>
          </div>
        );})}
        {!cots.length && <div className="text-[11px] text-slate-400">Sin cotizaciones aún. Usa "+ Cotización" para agregar hasta 3.</div>}
      </div>
      <div className="text-[11px] text-slate-400 mt-1">Factor = a cuántas {item.unidad} equivale 1 unidad cotizada por el proveedor. El precio final negociado (si existe) es el que se usa para calcular el total.</div>
      <div className="flex items-center gap-2 mt-2">
        <button onClick={guardar} className="text-xs bg-indigo-600 text-white px-3 py-1.5 rounded-md font-medium">Guardar cotizaciones</button>
        {guardadoMsg && <span className="text-[11px] text-emerald-600 flex items-center gap-1"><CheckCircle2 size={12} /> Cotizaciones guardadas correctamente</span>}
      </div>
    </div>
  );
}

function ComparativoTabla({ item, proveedores, onSeleccionar, seleccionada, soloLectura }) {
  const scored = calcularScores(item.cotizaciones, item.cantidad);
  const bestIdx = mejorCotizacionIdx(item.cotizaciones, item.cantidad);
  const elegidaIdx = seleccionada ?? bestIdx;
  const elegida = scored[elegidaIdx];
  const [pendienteIdx, setPendienteIdx] = useState(null);
  const [obsTemp, setObsTemp] = useState("");
  const nombreProv = (c) => proveedores.find((p) => p.id === c.proveedorId)?.nombre || c.proveedorNombre || "—";

  const clickElegir = (i) => {
    if (soloLectura) return;
    if (i === bestIdx) { onSeleccionar(item.id, i, ""); setPendienteIdx(null); return; }
    setPendienteIdx(i); setObsTemp(item.observacionSeleccion || "");
  };
  const confirmarNoSugerido = () => { if (!obsTemp.trim()) return; onSeleccionar(item.id, pendienteIdx, obsTemp.trim()); setPendienteIdx(null); };

  return (
    <div className="border border-slate-200 rounded-lg overflow-hidden">
      <div className="px-3 py-2 bg-slate-50 text-sm font-medium text-slate-700 flex items-center justify-between">
        <span>{item.nombre} — {item.cantidad} {item.unidad}</span>
        {soloLectura && <span className="text-[11px] text-slate-400 flex items-center gap-1"><Lock size={11} /> Bloqueado (orden ya generada)</span>}
      </div>
      <table className="w-full text-xs">
        <thead className="bg-white text-slate-500 border-b border-slate-100"><tr><th className="text-left px-3 py-2">Proveedor</th><th className="text-right px-3 py-2">Precio inicial</th><th className="text-right px-3 py-2">Precio final neg.</th><th className="text-right px-3 py-2">Total</th><th className="text-right px-3 py-2">Entrega</th><th className="text-right px-3 py-2">Score</th><th className="px-3 py-2"></th></tr></thead>
        <tbody>{scored.map((c, i) => (
          <tr key={i} className={`border-t border-slate-100 ${i === bestIdx ? "bg-emerald-50/60" : ""}`}>
            <td className="px-3 py-2 font-medium text-slate-700 flex items-center gap-1">{i === bestIdx && <Award size={13} className="text-emerald-600" />} {nombreProv(c)} {c.archivoNombre && <Paperclip size={11} className="text-slate-400" />}</td>
            <td className="px-3 py-2 text-right">{fmt(c.precioUnitario)}</td>
            <td className="px-3 py-2 text-right">{c.precioFinal ? fmt(c.precioFinal) : "—"}</td>
            <td className="px-3 py-2 text-right font-medium">{fmt(c.total)}</td>
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
      {elegida && <div className="px-3 py-2 bg-slate-50 border-t border-slate-100 text-xs text-slate-600 flex justify-end gap-4"><span>Subtotal: <b>{fmt(elegida.subtotal)}</b></span><span>IVA: <b>{fmt(elegida.iva)}</b></span><span>Total: <b>{fmt(elegida.total)}</b></span></div>}
      {item.observacionSeleccion && <div className="px-3 py-2 border-t border-slate-100 text-[11px] text-amber-700 italic bg-amber-50/50">Justificación de selección no sugerida: "{item.observacionSeleccion}"</div>}
    </div>
  );
}

function ResumenTotales({ solicitud }) {
  const d = desgloseSolicitud(solicitud);
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
function PagosEstructurados({ solicitud, total, currentUser, onProgramar, onConfirmar }) {
  const [pagos, setPagos] = useState(solicitud.pagos);
  const editable = puedeEditarPagos(currentUser) && !solicitud.pagosConfirmados;
  const pagado = totalPagado(pagos);
  const restante = total - pagado;
  const sug = solicitud.pagosSugeridos;
  const hasSugerencia = parseFloat(sug?.anticipo?.valor) > 0 || parseFloat(sug?.final?.valor) > 0;

  const set = (campo, sub, val) => { const copy = { ...pagos, [campo]: { ...pagos[campo], [sub]: val } }; setPagos(copy); onProgramar(copy); };
  const usarSugerencia = () => { setPagos(sug); onProgramar(sug); };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="flex items-center justify-between mb-1">
        <div className="flex items-center gap-2 text-slate-700 font-medium"><CalendarClock size={16} /> Plan de pagos (máx. 3)</div>
        {solicitud.pagosConfirmados ? <Badge tone="green">Confirmado por Dirección Financiera</Badge> : <Badge tone="amber">Pendiente de confirmación</Badge>}
      </div>
      {hasSugerencia && !solicitud.pagosConfirmados && (
        <div className="text-[11px] text-slate-500 bg-slate-50 border border-slate-200 rounded-md px-2 py-1.5 mb-2 flex items-center justify-between">
          <span>El solicitante sugirió: anticipo {fmt(sug.anticipo.valor)} ({sug.anticipo.fecha || "sin fecha"}){sug.intermedio.activo ? `, intermedio ${fmt(sug.intermedio.valor)}` : ""}, final {fmt(sug.final.valor)} ({sug.final.fecha || "sin fecha"})</span>
          {puedeEditarPagos(currentUser) && <button onClick={usarSugerencia} className="text-indigo-600 font-medium ml-2 shrink-0">Usar sugerencia</button>}
        </div>
      )}
      {!editable && !solicitud.pagosConfirmados && <div className="text-[11px] text-slate-400 mb-3">Solo Dirección Financiera puede editar y confirmar este plan.</div>}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-2">
        <div className="border border-slate-200 rounded-lg p-3">
          <div className="text-xs font-medium text-slate-500 mb-2">Anticipo</div>
          <input disabled={!editable} type="number" placeholder="Valor anticipo" value={pagos.anticipo.valor} onChange={(e) => set("anticipo", "valor", e.target.value)} className="w-full mb-1.5 border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <input disabled={!editable} type="date" value={pagos.anticipo.fecha} onChange={(e) => set("anticipo", "fecha", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <div className="text-[11px] text-slate-400 mt-1">Fecha primer pago</div>
        </div>
        <div className="border border-slate-200 rounded-lg p-3">
          <label className="text-xs font-medium text-slate-500 mb-2 flex items-center gap-1.5"><input disabled={!editable} type="checkbox" checked={pagos.intermedio.activo} onChange={(e) => set("intermedio", "activo", e.target.checked)} /> Pago intermedio (opcional)</label>
          <input disabled={!editable || !pagos.intermedio.activo} type="number" placeholder="Valor intermedio" value={pagos.intermedio.valor} onChange={(e) => set("intermedio", "valor", e.target.value)} className="w-full mb-1.5 border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <input disabled={!editable || !pagos.intermedio.activo} type="date" value={pagos.intermedio.fecha} onChange={(e) => set("intermedio", "fecha", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
        </div>
        <div className="border border-slate-200 rounded-lg p-3">
          <div className="text-xs font-medium text-slate-500 mb-2">Pago final</div>
          <input disabled={!editable} type="number" placeholder="Valor pago final" value={pagos.final.valor} onChange={(e) => set("final", "valor", e.target.value)} className="w-full mb-1.5 border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <input disabled={!editable} type="date" value={pagos.final.fecha} onChange={(e) => set("final", "fecha", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <div className="text-[11px] text-slate-400 mt-1">Fecha pago final</div>
        </div>
      </div>
      <div className="flex items-center justify-between mt-3">
        <div className="text-sm text-slate-500">Total orden: <b className="text-slate-700">{fmt(total)}</b> · Programado: <b className="text-slate-700">{fmt(pagado)}</b> · Restante: <b className={restante > 0.5 ? "text-amber-600" : restante < -0.5 ? "text-rose-600" : "text-emerald-600"}>{fmt(restante)}</b></div>
        {editable && <button onClick={onConfirmar} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium">Confirmar plan de pagos</button>}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   ORDEN ENVIADA AL PROVEEDOR
--------------------------------------------------------- */
function OcEnviadaPanel({ solicitud, currentUser, onGuardar }) {
  const [archivo, setArchivo] = useState(solicitud.ocEnviada.archivoNombre);
  if (solicitud.status !== "orden") return null;
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="font-medium text-slate-700 mb-2 flex items-center gap-2"><Send size={16} /> Envío de la orden al proveedor</div>
      <div className="text-xs text-slate-500 mb-2">Adjunta la orden de compra/servicio enviada para poder marcarla como "OC enviada".</div>
      <AdjuntarArchivo nombre={archivo} label="Adjuntar OC enviada (PDF/foto)" onSeleccionar={(n) => { setArchivo(n); onGuardar({ archivoNombre: n, fecha: hoy(), usuario: currentUser.nombre }); }} />
    </div>
  );
}

/* ---------------------------------------------------------
   RECEPCIÓN
--------------------------------------------------------- */
function RecepcionPanel({ solicitud, currentUser, onGuardar }) {
  const [r, setR] = useState(solicitud.recepcion);
  const set = (fields) => { const copy = { ...r, ...fields, usuario: currentUser.nombre, fecha: hoy() }; setR(copy); onGuardar(copy); };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
      <div className="font-medium text-slate-700 flex items-center gap-2"><PackageCheck size={16} /> Recepción</div>
      <AdjuntarArchivo nombre={r.archivoNombre} label="Adjuntar soporte de recepción (PDF/foto)" onSeleccionar={(n) => set({ archivoNombre: n })} />
      <div><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><MessageSquare size={12} /> Comentarios (opcional)</label><textarea value={r.comentario} onChange={(e) => set({ comentario: e.target.value })} rows={2} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" /></div>
      <label className="flex items-center gap-2 text-sm text-slate-700"><input type="checkbox" checked={r.recibidoSatisfaccion} onChange={(e) => set({ recibidoSatisfaccion: e.target.checked })} /> Recibo a satisfacción</label>
      {!r.recibidoSatisfaccion && <div className="text-[11px] text-amber-600">Marca "Recibo a satisfacción" para poder completar la solicitud.</div>}
    </div>
  );
}

/* ---------------------------------------------------------
   ORDEN DE COMPRA / TRABAJO — documento consolidado
--------------------------------------------------------- */
function OrdenDocumento({ solicitud, empresa, area, solicitante, proveedores, centrosCosto, conceptosGasto }) {
  const d = desgloseSolicitud(solicitud);
  const exportarPDF = () => window.print();
  const centroCosto = centrosCosto.find((c) => c.id === solicitud.centroCostoId);
  const conceptoGasto = conceptosGasto.find((c) => c.id === solicitud.conceptoGastoId);
  const nombreProv = (c) => proveedores.find((p) => p.id === c.proveedorId)?.nombre || c.proveedorNombre || "—";
  const pagoActivo = solicitud.tipo === "servicio";

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
    <div id="orden-imprimible" className="bg-white rounded-xl border-2 border-slate-300 p-6 space-y-5">
      {/* ENCABEZADO */}
      <div className="flex justify-between items-start border-b border-slate-200 pb-3">
        <div className="flex items-center gap-3">
          {empresa?.logoUrl && <img src={empresa.logoUrl} alt={empresa.nombre} className="h-12 max-w-[120px] object-contain" />}
          <div>
            <div className="text-base font-semibold text-slate-800">{solicitud.tipo === "compra" ? "Orden de Compra" : "Orden de Trabajo / Servicio"}</div>
            <div className="text-xs text-slate-400">{solicitud.folio} · {empresa?.nombre}</div>
          </div>
        </div>
        <Badge tone="blue">{PASOS.find((p) => p.key === solicitud.status)?.label}</Badge>
      </div>

      <div className="grid grid-cols-2 gap-2 text-xs text-slate-600">
        <div><b>Área:</b> {area?.nombre}</div><div><b>Solicitante:</b> {solicitante?.nombre}</div>
        <div><b>Centro de costo:</b> {centroCosto?.nombre || "—"}</div><div><b>Concepto de gasto:</b> {conceptoGasto?.nombre || "—"}</div>
        <div><b>Fecha creación:</b> {solicitud.fechaCreacion}</div><div><b>Fecha estimada:</b> {solicitud.fechaEstimada || "—"}</div>
      </div>

      {/* OBJETIVO Y JUSTIFICACIÓN */}
      <div className="grid grid-cols-2 gap-3 text-xs">
        <div><div className="font-medium text-slate-500 mb-0.5">Objetivo</div><div className="text-slate-600">{solicitud.objetivo}</div></div>
        <div><div className="font-medium text-slate-500 mb-0.5">Justificación</div><div className="text-slate-600">{solicitud.justificacion}</div></div>
      </div>

      {/* ÍTEMS Y PROVEEDOR ADJUDICADO */}
      <div>
        <div className="text-xs font-medium text-slate-500 mb-1">Ítems adjudicados</div>
        <table className="w-full text-xs">
          <thead className="text-slate-400 border-b border-slate-200"><tr><th className="text-left py-1">Ítem</th><th className="text-right py-1">Cant.</th><th className="text-left py-1">Proveedor</th><th className="text-right py-1">Total</th></tr></thead>
          <tbody>{solicitud.items.map((it) => { const idx = it.cotizacionSeleccionada ?? mejorCotizacionIdx(it.cotizaciones, it.cantidad); const cot = it.cotizaciones[idx]; const dd = desgloseItem(it);
            return <tr key={it.id} className="border-t border-slate-100"><td className="py-1.5">{it.nombre}</td><td className="py-1.5 text-right">{it.cantidad} {it.unidad}</td><td className="py-1.5">{cot ? nombreProv(cot) : "—"}</td><td className="py-1.5 text-right">{fmt(dd.total)}</td></tr>; })}</tbody>
        </table>
      </div>

      {/* HISTÓRICO DE COTIZACIONES POR ÍTEM (las 3, no solo la elegida) */}
      {solicitud.items.some((it) => it.cotizaciones.length > 0) && (
        <div>
          <div className="text-xs font-medium text-slate-500 mb-1">Histórico de cotizaciones recibidas</div>
          {solicitud.items.filter((it) => it.cotizaciones.length > 0).map((it) => {
            const scored = calcularScores(it.cotizaciones, it.cantidad);
            const bestIdx = mejorCotizacionIdx(it.cotizaciones, it.cantidad);
            const elegidaIdx = it.cotizacionSeleccionada ?? bestIdx;
            return (
              <div key={it.id} className="mb-2">
                <div className="text-[11px] font-medium text-slate-600">{it.nombre} ({it.cantidad} {it.unidad})</div>
                <table className="w-full text-[11px] mb-1">
                  <thead className="text-slate-400 border-b border-slate-100"><tr><th className="text-left py-0.5">Proveedor</th><th className="text-right py-0.5">Precio inicial</th><th className="text-right py-0.5">Precio final neg.</th><th className="text-right py-0.5">Total</th><th className="text-right py-0.5">Entrega</th><th className="text-right py-0.5">Score</th><th className="text-center py-0.5">Elegida</th></tr></thead>
                  <tbody>{scored.map((c, i) => (
                    <tr key={i} className="border-t border-slate-50"><td className="py-0.5">{nombreProv(c)}</td><td className="py-0.5 text-right">{fmt(c.precioUnitario)}</td><td className="py-0.5 text-right">{c.precioFinal ? fmt(c.precioFinal) : "—"}</td><td className="py-0.5 text-right">{fmt(c.total)}</td><td className="py-0.5 text-right">{c.diasEntrega} días</td><td className="py-0.5 text-right">{(c.score * 100).toFixed(0)}%</td><td className="py-0.5 text-center">{i === elegidaIdx ? "✓" : ""}</td></tr>
                  ))}</tbody>
                </table>
                {it.observacionSeleccion && <div className="text-[11px] text-amber-700 italic">Justificación de selección no sugerida: "{it.observacionSeleccion}"</div>}
              </div>
            );
          })}
        </div>
      )}

      {/* TOTALES */}
      <div className="flex justify-end text-sm border-t border-slate-200 pt-2">
        <div className="text-right">
          <div className="text-xs text-slate-500">Subtotal: {fmt(d.subtotal)} · IVA: {fmt(d.iva)}</div>
          <div><span className="text-slate-500">Total: </span><b className="text-slate-800">{fmt(d.total)}</b></div>
        </div>
      </div>

      {/* REVISIÓN DE COMPRAS */}
      {solicitud.tipo === "compra" && solicitud.revisionCompras.estado !== "no_aplica" && (
        <div className="text-xs"><b className="text-slate-500">Revisión de Compras:</b> {solicitud.revisionCompras.estado} — {solicitud.revisionCompras.usuario || "—"} ({solicitud.revisionCompras.fecha || "—"}){solicitud.revisionCompras.observacion && ` · "${solicitud.revisionCompras.observacion}"`}</div>
      )}

      {/* FIRMAS */}
      <div className="salto-pagina">
        <div className="text-xs font-medium text-slate-500 mb-2">Firmas y aprobaciones</div>
        <div className="grid grid-cols-2 gap-3 text-xs">
          {[["Solicitante", solicitud.firmas.solicitante], ["Jefe de área", solicitud.firmas.jefe], ["Dirección financiera", solicitud.firmas.financiera], ["Gerencia", solicitud.firmas.gerencia]].map(([rol, f]) => (
            <div key={rol} className="border border-slate-200 rounded-md p-2">
              <div className="text-[11px] text-slate-400">{rol}</div>
              {f?.nombre ? (
                <>
                  {f.fotoUrl && <img src={f.fotoUrl} className="h-8 object-contain my-1" alt="" />}
                  <div className="font-medium text-slate-700">{f.nombre}</div>
                  <div className="text-[11px] text-slate-400">{f.fecha}{f.aprobado === false ? " · Rechazado" : f.aprobado ? " · Aprobado" : ""}</div>
                  {f.observacion && <div className="text-[11px] text-slate-500 italic mt-0.5">"{f.observacion}"</div>}
                </>
              ) : <div className="text-slate-400">Pendiente</div>}
            </div>
          ))}
        </div>
      </div>

      {/* PLAN DE PAGOS (solo servicio) */}
      {pagoActivo && (
        <div>
          <div className="text-xs font-medium text-slate-500 mb-1">Plan de pagos {solicitud.pagosConfirmados ? "(confirmado por Dirección Financiera)" : "(sin confirmar)"}</div>
          <table className="w-full text-[11px]">
            <thead className="text-slate-400 border-b border-slate-100"><tr><th className="text-left py-0.5">Pago</th><th className="text-right py-0.5">Valor</th><th className="text-right py-0.5">Fecha</th><th className="text-center py-0.5">Pagado</th></tr></thead>
            <tbody>
              <tr className="border-t border-slate-50"><td className="py-0.5">Anticipo</td><td className="py-0.5 text-right">{fmt(solicitud.pagos.anticipo.valor)}</td><td className="py-0.5 text-right">{solicitud.pagos.anticipo.fecha || "—"}</td><td className="py-0.5 text-center">{solicitud.pagos.anticipo.pagado ? "✓" : ""}</td></tr>
              {solicitud.pagos.intermedio.activo && <tr className="border-t border-slate-50"><td className="py-0.5">Intermedio</td><td className="py-0.5 text-right">{fmt(solicitud.pagos.intermedio.valor)}</td><td className="py-0.5 text-right">{solicitud.pagos.intermedio.fecha || "—"}</td><td className="py-0.5 text-center">{solicitud.pagos.intermedio.pagado ? "✓" : ""}</td></tr>}
              <tr className="border-t border-slate-50"><td className="py-0.5">Final</td><td className="py-0.5 text-right">{fmt(solicitud.pagos.final.valor)}</td><td className="py-0.5 text-right">{solicitud.pagos.final.fecha || "—"}</td><td className="py-0.5 text-center">{solicitud.pagos.final.pagado ? "✓" : ""}</td></tr>
            </tbody>
          </table>
        </div>
      )}

      {/* OC ENVIADA Y RECEPCIÓN */}
      {(solicitud.ocEnviada.archivoNombre || solicitud.recepcion.archivoNombre || solicitud.recepcion.comentario) && (
        <div className="grid grid-cols-2 gap-3 text-xs">
          <div><div className="font-medium text-slate-500 mb-0.5">OC enviada al proveedor</div><div className="text-slate-600">{solicitud.ocEnviada.archivoNombre ? `${solicitud.ocEnviada.archivoNombre} · ${solicitud.ocEnviada.fecha}` : "—"}</div></div>
          <div><div className="font-medium text-slate-500 mb-0.5">Recepción</div><div className="text-slate-600">{solicitud.recepcion.recibidoSatisfaccion ? "Recibido a satisfacción" : "Pendiente"}{solicitud.recepcion.archivoNombre && ` · ${solicitud.recepcion.archivoNombre}`}{solicitud.recepcion.comentario && <div className="italic">"{solicitud.recepcion.comentario}"</div>}</div></div>
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
    case "cotizando": return "Generar cuadro comparativo";
    case "comparativo": return requiereDireccion(total) ? "Enviar a Dirección Financiera" : requiereGerencia(total) ? "Enviar a Gerencia" : "Generar orden";
    case "aprobacion_financiera": return requiereGerencia(total) ? "Aprobar y enviar a Gerencia" : "Aprobar y generar orden";
    case "aprobacion_gerencia": return "Aprobar y generar orden";
    case "orden": return "Marcar OC enviada al proveedor";
    case "oc_enviada": return "Confirmar recepción / iniciar ejecución";
    case "recepcion": return "Marcar como completada";
    default: return "Avanzar";
  }
}

function SolicitudDetalle({ solicitud, areas, empresas, usuarios, proveedores, centrosCosto, conceptosGasto, historico, setHistorico, currentUser, onUpdate, onVolver }) {
  const [observacion, setObservacion] = useState("");
  const area = areas.find((a) => a.id === solicitud.areaId);
  const empresa = empresas.find((e) => e.id === solicitud.empresaId);
  const solicitante = usuarios.find((u) => u.id === solicitud.solicitanteId);
  const total = totalSolicitud(solicitud);
  const todasCotizadas = solicitud.items.every((i) => i.cotizaciones.length > 0);
  const comparativoBloqueado = ["orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status);
  const patch = (fields) => onUpdate({ ...solicitud, ...fields });
  const empujarHistorial = (status) => [...solicitud.historialEstados, { status, fecha: ahoraISO() }];
  const notificar = (mensaje) => [...solicitud.notificaciones, { fecha: ahoraISO(), mensaje }];

  const guardarCotizaciones = (itemId, cots) => patch({ items: solicitud.items.map((i) => (i.id === itemId ? { ...i, cotizaciones: cots } : i)) });
  const seleccionarCotizacion = (itemId, idx, obs) => patch({ items: solicitud.items.map((i) => (i.id === itemId ? { ...i, cotizacionSeleccionada: idx, observacionSeleccion: obs } : i)) });
  const guardarItemsRevision = (items) => patch({ items });
  const decidirRevisionCompras = (estado, obs) => {
    if (estado === "rechazada") { patch({ status: "rechazada", revisionCompras: { estado, observacion: obs, usuario: currentUser.nombre, fecha: hoy() }, notificaciones: notificar(`Correo simulado a ${solicitante?.nombre}: tu solicitud ${solicitud.folio} fue rechazada por Compras.`) }); return; }
    patch({ revisionCompras: { estado, observacion: obs, usuario: currentUser.nombre, fecha: hoy() } });
  };

  const puedeActuar = () => {
    switch (solicitud.status) {
      case "aprobacion_jefe": return puedeAprobarJefe(currentUser, solicitud);
      case "cotizando": return puedeGestionarCotizaciones(currentUser) && (solicitud.tipo !== "compra" || solicitud.revisionCompras.estado === "aprobada");
      case "comparativo": return puedeGestionarCotizaciones(currentUser);
      case "aprobacion_financiera": return puedeAprobarFinanciera(currentUser);
      case "aprobacion_gerencia": return puedeAprobarGerencia(currentUser);
      case "orden": case "oc_enviada": case "recepcion": return puedeGestionarCotizaciones(currentUser) || currentUser.rol === "Solicitante";
      default: return true;
    }
  };

  const firmar = () => ({ aprobado: true, nombre: currentUser.nombre, fecha: hoy(), observacion, fotoUrl: currentUser.firmaFotoUrl || null });
  const avanzar = () => {
    const s = solicitud.status;
    if (s === "aprobacion_jefe") patch({ status: "cotizando", firmas: { ...solicitud.firmas, jefe: firmar() }, historialEstados: empujarHistorial("cotizando"), notificaciones: notificar(`Correo simulado a Compras: solicitud ${solicitud.folio} aprobada, lista para cotizar.`) });
    else if (s === "cotizando" && todasCotizadas) patch({ status: "comparativo", historialEstados: empujarHistorial("comparativo") });
    else if (s === "comparativo") { const next = requiereDireccion(total) ? "aprobacion_financiera" : requiereGerencia(total) ? "aprobacion_gerencia" : "orden"; patch({ status: next, historialEstados: empujarHistorial(next), notificaciones: notificar(`Correo simulado: solicitud ${solicitud.folio} avanza a ${PASOS.find((p) => p.key === next)?.label}.`) }); }
    else if (s === "aprobacion_financiera") { const next = requiereGerencia(total) ? "aprobacion_gerencia" : "orden"; patch({ status: next, firmas: { ...solicitud.firmas, financiera: firmar() }, historialEstados: empujarHistorial(next) }); }
    else if (s === "aprobacion_gerencia") patch({ status: "orden", firmas: { ...solicitud.firmas, gerencia: firmar() }, historialEstados: empujarHistorial("orden") });
    else if (s === "orden") {
      if (!solicitud.ocEnviada.archivoNombre) return;
      patch({ status: "oc_enviada", historialEstados: empujarHistorial("oc_enviada"), notificaciones: notificar(`Correo enviado a ${solicitante?.nombre} (${solicitante?.email || "sin correo"}): la orden ${solicitud.folio} fue enviada al proveedor.`) });
      if (solicitante?.email) {
        enviarCorreo(
          solicitante.email,
          `Tu orden ${solicitud.folio} fue enviada al proveedor`,
          `<p>Hola ${solicitante.nombre},</p><p>La orden <b>${solicitud.folio}</b> ya fue enviada al proveedor y quedó lista para recepción.</p>`
        );
      }
    }
    else if (s === "oc_enviada") patch({ status: "recepcion", historialEstados: empujarHistorial("recepcion") });
    else if (s === "recepcion") {
      if (!solicitud.recepcion.recibidoSatisfaccion) return;
      if (solicitud.tipo === "servicio") {
        const pagado = totalPagado(solicitud.pagos);
        if (pagado < total - 0.5) return;
      }
      patch({ status: "completada", historialEstados: empujarHistorial("completada"), notificaciones: notificar(`Correo simulado a ${solicitante?.nombre}: tu solicitud ${solicitud.folio} fue completada.`) });
    }
    setObservacion("");
  };
  const rechazar = () => {
    const campo = solicitud.status === "aprobacion_jefe" ? "jefe" : solicitud.status === "aprobacion_financiera" ? "financiera" : solicitud.status === "aprobacion_gerencia" ? "gerencia" : null;
    patch({ status: "rechazada", firmas: campo ? { ...solicitud.firmas, [campo]: { aprobado: false, nombre: currentUser.nombre, fecha: hoy(), observacion, fotoUrl: currentUser.firmaFotoUrl || null } } : solicitud.firmas, historialEstados: empujarHistorial("rechazada"), notificaciones: notificar(`Correo simulado a ${solicitante?.nombre}: tu solicitud ${solicitud.folio} fue rechazada.`) });
    setObservacion("");
  };

  const mostrarObservacion = ["aprobacion_jefe", "aprobacion_financiera", "aprobacion_gerencia"].includes(solicitud.status);
  const autorizado = puedeActuar();

  return (
    <div className="space-y-5">
      <button onClick={onVolver} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-700"><ArrowLeft size={15} /> Volver a solicitudes</button>

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="flex items-start justify-between flex-wrap gap-3">
          {empresa?.logoUrl && <img src={empresa.logoUrl} alt={empresa.nombre} className="h-10 max-w-[100px] object-contain order-first" />}
          <div>
            <div className="flex items-center gap-2 flex-wrap"><h2 className="text-lg font-semibold text-slate-800">{solicitud.folio}</h2><Badge tone={solicitud.tipo === "compra" ? "blue" : "amber"}>{solicitud.tipo === "compra" ? <ShoppingCart size={12} /> : <Wrench size={12} />} {solicitud.tipo === "compra" ? "Orden de compra" : "Orden de servicio"}</Badge>{solicitud.status === "rechazada" && <Badge tone="red">Rechazada</Badge>}<button onClick={() => window.print()} className="text-xs bg-slate-800 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1 no-print"><FileText size={13} /> Exportar solicitud completa a PDF</button></div>
            <div className="text-sm text-slate-500 mt-1 flex items-center gap-3 flex-wrap"><span className="flex items-center gap-1"><Building2 size={13} /> {empresa?.nombre}</span><span>Área: {area?.nombre}</span><span>Solicitante: {solicitante?.nombre}</span><span className="flex items-center gap-1"><Calendar size={13} /> Est.: {solicitud.fechaEstimada || "—"}</span></div>
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

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-3 flex items-center gap-2"><PenTool size={15} /> Firmas</div>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
          <FirmaBlock rol="solicitante" firma={solicitud.firmas.solicitante} />
          <FirmaBlock rol="jefe de área" firma={solicitud.firmas.jefe} />
          <FirmaBlock rol="dirección financiera" firma={solicitud.firmas.financiera} />
          <FirmaBlock rol="gerencia" firma={solicitud.firmas.gerencia} />
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="font-medium text-slate-700 mb-3">Ítems solicitados</div>
        <table className="w-full text-sm mb-2"><thead className="text-slate-400 text-xs"><tr><th className="text-left py-1">Ítem</th><th className="text-right py-1">Cantidad</th><th className="text-right py-1">Unidad</th></tr></thead><tbody>{solicitud.items.map((it) => (<tr key={it.id} className="border-t border-slate-100"><td className="py-1.5">{it.nombre}</td><td className="py-1.5 text-right">{it.cantidad}</td><td className="py-1.5 text-right">{it.unidad}</td></tr>))}</tbody></table>
      </div>

      {solicitud.status === "cotizando" && puedeVerHistorico(currentUser) && (
        <RevisionCompras solicitud={solicitud} historico={historico} setHistorico={setHistorico} currentUser={currentUser} onGuardarItems={guardarItemsRevision} onDecision={decidirRevisionCompras} />
      )}
      {solicitud.status === "cotizando" && !puedeVerHistorico(currentUser) && solicitud.tipo === "compra" && solicitud.revisionCompras.estado === "pendiente" && (
        <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">Pendiente de revisión por el área de Compras antes de continuar con las cotizaciones.</div>
      )}

      {solicitud.status === "cotizando" && (solicitud.tipo !== "compra" || solicitud.revisionCompras.estado === "aprobada") && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="font-medium text-slate-700">Cargar hasta 3 cotizaciones por ítem (Compras)</div>
          {solicitud.items.map((it) => <CotizacionForm key={it.id} item={it} proveedores={proveedores} onGuardar={guardarCotizaciones} />)}
        </div>
      )}

      {["comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && solicitud.items.some((i) => i.cotizaciones.length > 0) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="font-medium text-slate-700 flex items-center gap-2"><TrendingUp size={16} /> Cuadro comparativo (sugerencia automática)</div>
          {solicitud.items.filter((i) => i.cotizaciones.length > 0).map((it) => <ComparativoTabla key={it.id} item={it} proveedores={proveedores} onSeleccionar={seleccionarCotizacion} seleccionada={it.cotizacionSeleccionada} soloLectura={comparativoBloqueado} />)}
        </div>
      )}

      {["comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && <ResumenTotales solicitud={solicitud} />}

      {solicitud.tipo === "servicio" && ["aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && (
        <PagosEstructurados solicitud={solicitud} total={total} currentUser={currentUser} onProgramar={(pagos) => patch({ pagos })} onConfirmar={() => patch({ pagosConfirmados: true })} />
      )}

      <OcEnviadaPanel solicitud={solicitud} currentUser={currentUser} onGuardar={(oc) => patch({ ocEnviada: oc })} />

      {["recepcion", "completada"].includes(solicitud.status) && <RecepcionPanel solicitud={solicitud} currentUser={currentUser} onGuardar={(r) => patch({ recepcion: r })} />}

      <div className="print-wrapper-oculto" style={{ display: "none" }}>
        <OrdenDocumento solicitud={solicitud} empresa={empresa} area={area} solicitante={solicitante} proveedores={proveedores} centrosCosto={centrosCosto} conceptosGasto={conceptosGasto} />
      </div>

      <TiempoProceso historial={solicitud.historialEstados} />
      <NotificacionesPanel notificaciones={solicitud.notificaciones} />

      {!autorizado && solicitud.status !== "completada" && solicitud.status !== "rechazada" && (
        <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 flex items-center gap-2"><ShieldCheck size={14} /> Tu rol ({currentUser.rol}) no tiene permiso para actuar sobre este paso del flujo.</div>
      )}

      {solicitud.status !== "completada" && solicitud.status !== "rechazada" && autorizado && (
        <div className="space-y-2">
          {mostrarObservacion && (<div><label className="text-xs font-medium text-slate-500">Observación de aprobación (opcional)</label><textarea value={observacion} onChange={(e) => setObservacion(e.target.value)} rows={2} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" placeholder="Comentarios sobre esta aprobación..." /></div>)}
          <div className="flex gap-2 justify-end">
            <button onClick={rechazar} className="px-4 py-2 rounded-lg text-sm text-rose-600 border border-rose-200 flex items-center gap-1"><XCircle size={15} /> Rechazar</button>
            <button onClick={avanzar} disabled={(solicitud.status === "cotizando" && !todasCotizadas) || (solicitud.status === "orden" && !solicitud.ocEnviada.archivoNombre) || (solicitud.status === "recepcion" && !solicitud.recepcion.recibidoSatisfaccion)} className="px-4 py-2 rounded-lg text-sm bg-indigo-600 text-white font-medium disabled:opacity-40 flex items-center gap-1">{accionLabel(solicitud, total)} <ChevronRight size={15} /></button>
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
    const idx = it.cotizacionSeleccionada ?? mejorCotizacionIdx(it.cotizaciones, it.cantidad);
    const cot = it.cotizaciones[idx];
    if (!cot) return null;
    return proveedores.find((p) => p.id === cot.proveedorId)?.nombre || cot.proveedorNombre || null;
  }).filter(Boolean))];
  return nombres.length ? nombres.join(", ") : "—";
}

function VistaSolicitudes({ solicitudes, areas, empresas, usuarios, proveedores, onAbrir, onExportar, titulo }) {
  const [fArea, setFArea] = useState("todas");
  const [fEmpresa, setFEmpresa] = useState("todas");
  const [fEstado, setFEstado] = useState("todos");
  const [fCreadoPor, setFCreadoPor] = useState("todos");
  const [fDesde, setFDesde] = useState("");
  const [fHasta, setFHasta] = useState("");

  const creadores = usuarios.filter((u) => solicitudes.some((s) => s.solicitanteId === u.id));
  const filtradas = solicitudes.filter((s) =>
    (fArea === "todas" || s.areaId === fArea) &&
    (fEmpresa === "todas" || s.empresaId === fEmpresa) &&
    (fEstado === "todos" || s.status === fEstado) &&
    (fCreadoPor === "todos" || s.solicitanteId === fCreadoPor) &&
    (!fDesde || s.fechaCreacion >= fDesde) &&
    (!fHasta || s.fechaCreacion <= fHasta)
  );
  const hayFiltros = fArea !== "todas" || fEmpresa !== "todas" || fEstado !== "todos" || fCreadoPor !== "todos" || fDesde || fHasta;
  const limpiarFiltros = () => { setFArea("todas"); setFEmpresa("todas"); setFEstado("todos"); setFCreadoPor("todos"); setFDesde(""); setFHasta(""); };

  return (
    <div className="space-y-4">
      <h2 className="text-lg font-semibold text-slate-800">{titulo}</h2>
      <div className="bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
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
      <ListaSolicitudes solicitudes={filtradas} areas={areas} empresas={empresas} proveedores={proveedores} onAbrir={onAbrir} onExportar={onExportar} />
    </div>
  );
}

function ListaSolicitudes({ solicitudes, areas, empresas, proveedores, onAbrir, onExportar }) {
  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-x-auto">
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-slate-500"><tr><th className="text-left px-4 py-2 font-medium">Folio</th><th className="text-left px-4 py-2 font-medium">Tipo</th><th className="text-left px-4 py-2 font-medium">Área</th><th className="text-left px-4 py-2 font-medium">Empresa</th><th className="text-left px-4 py-2 font-medium">Fecha de registro</th><th className="text-left px-4 py-2 font-medium">Objetivo</th><th className="text-left px-4 py-2 font-medium">Proveedor adjudicado</th><th className="text-right px-4 py-2 font-medium">Total (IVA incl.)</th><th className="text-left px-4 py-2 font-medium">Estado</th><th></th><th></th></tr></thead>
        <tbody>{solicitudes.map((s) => { const area = areas.find((a) => a.id === s.areaId), empresa = empresas.find((e) => e.id === s.empresaId), paso = PASOS.find((p) => p.key === s.status);
          return (<tr key={s.id} className="border-t border-slate-100 hover:bg-slate-50 cursor-pointer" onClick={() => onAbrir(s.id)}>
            <td className="px-4 py-2.5 font-medium text-slate-700">{s.folio}</td>
            <td className="px-4 py-2.5"><Badge tone={s.tipo === "compra" ? "blue" : "amber"}>{s.tipo === "compra" ? "Compra" : "Servicio"}</Badge></td>
            <td className="px-4 py-2.5 text-slate-600">{area?.nombre}</td>
            <td className="px-4 py-2.5 text-slate-600">{empresa?.nombre}</td>
            <td className="px-4 py-2.5 text-slate-600 whitespace-nowrap">{s.fechaCreacion}</td>
            <td className="px-4 py-2.5 text-slate-600 max-w-[220px] truncate" title={s.objetivo}>{s.objetivo}</td>
            <td className="px-4 py-2.5 text-slate-600 max-w-[160px] truncate" title={proveedoresAdjudicados(s, proveedores)}>{proveedoresAdjudicados(s, proveedores)}</td>
            <td className="px-4 py-2.5 text-right text-slate-600">{fmt(totalSolicitud(s))}</td>
            <td className="px-4 py-2.5"><Badge tone={s.status === "completada" ? "green" : s.status === "rechazada" ? "red" : "slate"}>{s.status === "rechazada" ? "Rechazada" : paso?.label}</Badge></td>
            <td className="px-4 py-2.5 text-right"><button title="Exportar a PDF" onClick={(e) => { e.stopPropagation(); onExportar(s); }} className="text-slate-400 hover:text-indigo-600 p-1"><FileText size={15} /></button></td>
            <td className="px-4 py-2.5 text-right"><ChevronRight size={15} className="text-slate-300" /></td></tr>); })}</tbody>
      </table>
    </div>
  );
}

/* ---------------------------------------------------------
   LOGOS DE EMPRESAS
--------------------------------------------------------- */
function EmpresasLogos({ empresas, onGuardar }) {
  const cargarLogo = async (empresa, file) => {
    const url = await subirArchivo(file, "logos");
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
  empresas, guardarEmpresa, eliminarEmpresa,
  areas,
  proveedores, guardarProveedor, eliminarProveedor,
  usuarios, guardarUsuario, eliminarUsuario,
  itemsCatalogo, guardarItemCatalogo, eliminarItemCatalogo,
  centrosCosto, guardarCentroCosto, eliminarCentroCosto,
  conceptosGasto, guardarConceptoGasto, eliminarConceptoGasto,
}) {
  const [sub, setSub] = useState("empresas");
  const tabs = [
    { key: "empresas", label: "Empresas", icon: Building2 }, { key: "proveedores", label: "Proveedores", icon: Truck },
    { key: "usuarios", label: "Usuarios y roles", icon: Users }, { key: "items", label: "Ítems", icon: Boxes },
    { key: "centros", label: "Centros de costo", icon: Layers }, { key: "conceptos", label: "Conceptos de gasto", icon: ClipboardList },
  ];

  return (
    <div className="space-y-4">
      <div className="flex gap-2 flex-wrap">{tabs.map((t) => (<button key={t.key} onClick={() => setSub(t.key)} className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium border ${sub === t.key ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}><t.icon size={14} /> {t.label}</button>))}</div>
      {sub === "empresas" && (
        <>
          <EmpresasLogos empresas={empresas} onGuardar={guardarEmpresa} />
          <CrudTable titulo="Empresas parametrizadas" icon={Building2} columnas={[{ key: "nombre", label: "Nombre" }, { key: "nit", label: "NIT" }]} datos={empresas} onGuardar={guardarEmpresa} onEliminar={eliminarEmpresa} plantilla={{ nombre: "", nit: "" }} />
        </>
      )}
      {sub === "proveedores" && <CrudTable titulo="Proveedores" icon={Truck} columnas={[{ key: "nombre", label: "Nombre" }, { key: "nit", label: "NIT" }, { key: "contacto", label: "Contacto" }, { key: "email", label: "Correo electrónico" }]} datos={proveedores} onGuardar={guardarProveedor} onEliminar={eliminarProveedor} plantilla={{ nombre: "", nit: "", contacto: "", email: "" }} />}
      {sub === "usuarios" && (
        <>
          <div className="text-xs text-slate-400 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2">
            El rol determina qué puede aprobar cada usuario: <b>Jefe de Área</b> aprueba solicitudes de su misma área, <b>Dirección Financiera</b> y <b>Gerencia</b> aprueban según el monto, <b>Compras</b> gestiona cotizaciones, histórico y pagos.
            <br /><b>Importante:</b> editar o agregar una fila aquí solo cambia sus datos de perfil (nombre, cargo, área, rol). Para que una persona pueda <i>iniciar sesión</i>, primero debes crearla en Supabase → Authentication → Users con el mismo correo, y vincular su ID ahí.
          </div>
          <CrudTable titulo="Usuarios y roles" icon={Users}
            columnas={[{ key: "nombre", label: "Nombre" }, { key: "email", label: "Correo electrónico" }, { key: "cargo", label: "Cargo" }, { key: "areaId", label: "Área", type: "select", options: areas.map((a) => ({ value: a.id, label: a.nombre })) }, { key: "rol", label: "Rol", type: "select", options: ROLES.map((r) => ({ value: r, label: r })) }]}
            datos={usuarios} onGuardar={guardarUsuario} onEliminar={eliminarUsuario} plantilla={{ nombre: "", email: "", cargo: "", areaId: "", rol: "Solicitante" }} />
        </>
      )}
      {sub === "items" && <CrudTable titulo="Catálogo de ítems" icon={Boxes} columnas={[{ key: "nombre", label: "Nombre" }, { key: "unidadDefault", label: "Unidad", type: "select", options: UNIDADES.map((u) => ({ value: u, label: u })) }, { key: "categoria", label: "Categoría" }]} datos={itemsCatalogo} onGuardar={guardarItemCatalogo} onEliminar={eliminarItemCatalogo} plantilla={{ nombre: "", unidadDefault: "unidad", categoria: "" }} />}
      {sub === "centros" && <CrudTable titulo="Centros de costo" icon={Layers} columnas={[{ key: "nombre", label: "Nombre" }]} datos={centrosCosto} onGuardar={guardarCentroCosto} onEliminar={eliminarCentroCosto} plantilla={{ nombre: "" }} />}
      {sub === "conceptos" && <CrudTable titulo="Conceptos de gasto" icon={ClipboardList} columnas={[{ key: "nombre", label: "Nombre" }]} datos={conceptosGasto} onGuardar={guardarConceptoGasto} onEliminar={eliminarConceptoGasto} plantilla={{ nombre: "" }} />}
    </div>
  );
}

/* ---------------------------------------------------------
   APP PRINCIPAL
--------------------------------------------------------- */
export default function App() {
  // --- Catálogos leídos/guardados en Supabase (áreas, empresas, proveedores, ítems, centros de costo, conceptos de gasto, usuarios) ---
  const { datos: areas, cargando: cargandoAreas } = useSupabaseTable('areas', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, presupuesto: r.presupuesto_mensual }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, presupuesto_mensual: r.presupuesto }),
    orderBy: 'nombre',
  });
  const { datos: empresas, cargando: cargandoEmpresas, guardar: guardarEmpresa, eliminar: eliminarEmpresa } = useSupabaseTable('empresas', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, nit: r.nit, logoUrl: r.logo_url }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, nit: r.nit, logo_url: r.logoUrl }),
    orderBy: 'nombre',
  });
  const { datos: proveedores, cargando: cargandoProveedores, guardar: guardarProveedor, eliminar: eliminarProveedor, guardarVarios: importarProveedores } = useSupabaseTable('proveedores', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, nit: r.nit, contacto: r.contacto, email: r.email }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, nit: r.nit, contacto: r.contacto, email: r.email }),
    orderBy: 'nombre',
  });
  const { datos: usuarios, cargando: cargandoUsuarios, guardar: guardarUsuario, eliminar: eliminarUsuario, guardarVarios: importarUsuarios } = useSupabaseTable('usuarios', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, email: r.email, cargo: r.cargo, areaId: r.area_id, rol: r.rol, firmaFotoUrl: r.firma_foto_url }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, email: r.email, cargo: r.cargo, area_id: r.areaId, rol: r.rol }),
    orderBy: 'nombre',
  });
  const { datos: itemsCatalogo, cargando: cargandoItems, guardar: guardarItemCatalogo, eliminar: eliminarItemCatalogo, guardarVarios: importarItems } = useSupabaseTable('items_catalogo', {
    desdeDb: (r) => ({ id: r.id, nombre: r.nombre, unidadDefault: r.unidad_default, categoria: r.categoria }),
    haciaDb: (r) => ({ id: r.id, nombre: r.nombre, unidad_default: r.unidadDefault, categoria: r.categoria }),
    orderBy: 'nombre',
  });
  const { datos: centrosCosto, cargando: cargandoCentros, guardar: guardarCentroCosto, eliminar: eliminarCentroCosto, guardarVarios: importarCentros } = useSupabaseTable('centros_costo', {
    orderBy: 'nombre',
  });
  const { datos: conceptosGasto, cargando: cargandoConceptos, guardar: guardarConceptoGasto, eliminar: eliminarConceptoGasto, guardarVarios: importarConceptos } = useSupabaseTable('conceptos_gasto', {
    orderBy: 'nombre',
  });
  const cargandoCatalogos = cargandoAreas || cargandoEmpresas || cargandoProveedores || cargandoUsuarios || cargandoItems || cargandoCentros || cargandoConceptos;

  const [historico, setHistorico] = useState(HISTORICO_INIT);
  const { solicitudes, cargando: cargandoSolicitudes, crear: crearSolicitudDB, actualizar: actualizarSolicitudDB } = useSolicitudes();
  const [tab, setTab] = useState("solicitudes");
  const [abierta, setAbierta] = useState(null);
  const [creando, setCreando] = useState(false);
  const [perfil, setPerfil] = useState(false);
  const [menuExpandido, setMenuExpandido] = useState(true);
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
  };

  const crearSolicitud = async (nueva) => { await crearSolicitudDB(nueva); setCreando(false); setTab("solicitudes"); };
  const actualizarSolicitud = async (upd) => { await actualizarSolicitudDB(upd); };
  // La foto de firma del perfil, por ahora, solo se guarda en memoria durante la sesión.
  // Falta conectar esto a un "update" real sobre la tabla usuarios (próximo módulo a migrar).
  const guardarPerfil = async (u) => { await actualizarPerfil({ firma_foto_url: u.firmaFotoUrl }); setPerfil(false); };
  const solicitudAbierta = solicitudes.find((s) => s.id === abierta);
  const solicitudesVisibles = puedeVerTodasSolicitudes(currentUser) ? solicitudes : solicitudes.filter((s) => s.solicitanteId === currentUser.id);

  const NavBtn = ({ id, icon: Icon, label }) => (
    <button title={label} onClick={() => { setTab(id); setAbierta(null); setCreando(false); setPerfil(false); }} className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium w-full text-left ${!menuExpandido ? "justify-center px-2" : ""} ${tab === id && !abierta && !creando && !perfil ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"}`}>
      <Icon size={16} className="shrink-0" /> {menuExpandido && label}
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

        <button title="Nueva solicitud" onClick={() => { setCreando(true); setAbierta(null); setPerfil(false); }} className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium w-full text-left mb-1 ${!menuExpandido ? "justify-center px-2" : ""} ${creando ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"}`}>
          <Plus size={16} className="shrink-0" /> {menuExpandido && "Nueva solicitud"}
        </button>
        <NavBtn id="solicitudes" icon={ListChecks} label={puedeVerTodasSolicitudes(currentUser) ? "Solicitudes" : "Mis solicitudes"} />
        <NavBtn id="dashboard" icon={LayoutDashboard} label="Dashboard" />
        <NavBtn id="estadisticas" icon={BarChart3} label="Estadísticas" />
        {puedeVerCatalogos(currentUser) && <NavBtn id="catalogos" icon={Settings} label="Catálogo" />}

        <div className="mt-auto pt-4 border-t border-slate-100">
          <button title="Mi perfil" onClick={() => { setPerfil(true); setAbierta(null); setCreando(false); }} className={`flex items-center gap-2 px-2 py-1.5 rounded-lg text-xs text-slate-500 hover:bg-slate-100 w-full mb-1 ${!menuExpandido ? "justify-center" : ""}`}><UserCircle size={13} className="shrink-0" /> {menuExpandido && "Mi perfil"}</button>
          {menuExpandido && (
            <div className="px-2 mb-2">
              <div className="text-sm font-medium text-slate-700 truncate">{currentUser.nombre}</div>
              <div className="text-[11px] text-slate-400 truncate">{currentUser.rol} · {areas.find((a) => a.id === currentUser.areaId)?.nombre}</div>
            </div>
          )}
          <button title="Cerrar sesión" onClick={() => { cerrarSesion(); setAbierta(null); setCreando(false); }} className={`flex items-center gap-2 px-2 py-1.5 rounded-lg text-xs text-slate-500 hover:bg-slate-100 w-full ${!menuExpandido ? "justify-center" : ""}`}><LogOut size={13} className="shrink-0" /> {menuExpandido && "Cerrar sesión"}</button>
          {menuExpandido && <div className="text-[11px] text-slate-400 px-2 leading-relaxed mt-2">Umbral Dir. Financiera: {fmt(UMBRAL_DIRECCION)}<br />Umbral Gerencia: {fmt(UMBRAL_GERENCIA)}</div>}
        </div>
      </aside>

      <main className="flex-1 p-6 overflow-auto">
        {creando ? (
          <NuevaSolicitud areas={areas} empresas={empresas} itemsCatalogo={itemsCatalogo} centrosCosto={centrosCosto} conceptosGasto={conceptosGasto} usuarios={usuarios} currentUser={currentUser} onCrear={crearSolicitud} onCancel={() => setCreando(false)} />
        ) : perfil ? (
          <PerfilUsuario currentUser={currentUser} onGuardar={guardarPerfil} />
        ) : solicitudAbierta ? (
          <SolicitudDetalle solicitud={solicitudAbierta} areas={areas} empresas={empresas} usuarios={usuarios} proveedores={proveedores} centrosCosto={centrosCosto} conceptosGasto={conceptosGasto} historico={historico} setHistorico={setHistorico} currentUser={currentUser} onUpdate={actualizarSolicitud} onVolver={() => setAbierta(null)} />
        ) : tab === "dashboard" ? (
          <Dashboard areas={areas} solicitudes={solicitudesVisibles} />
        ) : tab === "estadisticas" ? (
          <Estadisticas solicitudes={solicitudesVisibles} areas={areas} empresas={empresas} proveedores={proveedores} />
        ) : tab === "catalogos" && puedeVerCatalogos(currentUser) ? (
          <Catalogos
            empresas={empresas} guardarEmpresa={guardarEmpresa} eliminarEmpresa={eliminarEmpresa}
            areas={areas}
            proveedores={proveedores} guardarProveedor={guardarProveedor} eliminarProveedor={eliminarProveedor}
            usuarios={usuarios} guardarUsuario={guardarUsuario} eliminarUsuario={eliminarUsuario}
            itemsCatalogo={itemsCatalogo} guardarItemCatalogo={guardarItemCatalogo} eliminarItemCatalogo={eliminarItemCatalogo}
            centrosCosto={centrosCosto} guardarCentroCosto={guardarCentroCosto} eliminarCentroCosto={eliminarCentroCosto}
            conceptosGasto={conceptosGasto} guardarConceptoGasto={guardarConceptoGasto} eliminarConceptoGasto={eliminarConceptoGasto}
          />
        ) : (
          <VistaSolicitudes solicitudes={solicitudesVisibles} areas={areas} empresas={empresas} usuarios={usuarios} proveedores={proveedores} onAbrir={setAbierta} onExportar={setExportando} titulo={puedeVerTodasSolicitudes(currentUser) ? "Solicitudes" : "Mis solicitudes"} />
        )}
      </main>

      {exportando && !solicitudAbierta && (
        <div className="print-wrapper-oculto" style={{ display: "none" }}>
          <OrdenDocumento
            solicitud={exportando}
            empresa={empresas.find((e) => e.id === exportando.empresaId)}
            area={areas.find((a) => a.id === exportando.areaId)}
            solicitante={usuarios.find((u) => u.id === exportando.solicitanteId)}
            proveedores={proveedores}
            centrosCosto={centrosCosto}
            conceptosGasto={conceptosGasto}
          />
        </div>
      )}
    </div>
  );
}