import React, { useState, useMemo, useEffect, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { useAuth } from "./hooks/useAuth";
import { useSupabaseTable } from "./hooks/useSupabaseTable";
import { useSolicitudes } from "./hooks/useSolicitudes";
import { subirArchivo, obtenerUrlFirmada, archivoDentroDelLimite, TAMANO_MAXIMO_MB, subirArchivoPublico } from "./lib/storage";
import { obtenerTasaCambioCOP } from "./lib/tasaCambio";
import { firmarPDF } from "./lib/firmarPdf";
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
  { key: "oc_enviada", label: "OC enviada al proveedor" },
  { key: "recepcion", label: "Recepción / Ejecución" },
  { key: "completada", label: "Completada" },
];

const fmt = (n) => (n || 0).toLocaleString("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });
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
  { key: "editar_pagos", label: "Editar y confirmar el plan de pagos" },
  { key: "ver_catalogos", label: "Ver el Catálogo" },
  { key: "ver_todas_solicitudes", label: "Ver todas las solicitudes (no solo las propias)" },
  { key: "reabrir_solicitudes", label: "Reabrir solicitudes rechazadas" },
  { key: "ver_historico", label: "Ver histórico de compras" },
  { key: "ver_mis_pendientes", label: "Ver la pantalla \"Mis pendientes\"" },
  { key: "ver_calendario_pagos", label: "Ver el Calendario de pagos" },
  { key: "ver_ordenes_enviadas", label: "Ver el reporte de Órdenes enviadas a proveedores" },
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

  // arma una fila por cada pago confirmado (anticipo / intermedio / final) de cada solicitud tipo servicio;
  // las que no tienen plan de pagos (compras, o servicios sin confirmar) usan su fecha estimada de entrega
  const filas = [];
  solicitudes.forEach((s) => {
    if (["completada", "rechazada"].includes(s.status)) return;
    const prov = proveedoresAdjudicados(s, proveedores);
    if (s.tipo === "servicio" && s.pagosConfirmados) {
      const tramos = [
        { tipo: "Anticipo", ...s.pagos.anticipo },
        ...(s.pagos.intermedio.activo ? [{ tipo: "Intermedio", ...s.pagos.intermedio }] : []),
        { tipo: "Final", ...s.pagos.final },
      ];
      tramos.forEach((t) => {
        if (!(parseFloat(t.valor) > 0) || !t.fecha) return;
        const dias = Math.round((new Date(t.fecha + "T00:00:00") - new Date(hoy() + "T00:00:00")) / 86400000);
        filas.push({ id: `${s.id}-${t.tipo}`, solicitudId: s.id, folio: s.folio, proveedor: prov, tipo: t.tipo, valor: parseFloat(t.valor), fecha: t.fecha, dias, pagado: !!t.pagado });
      });
    } else if (s.fechaEstimada) {
      const total = totalSolicitud(s);
      if (total > 0) {
        const dias = Math.round((new Date(s.fechaEstimada + "T00:00:00") - new Date(hoy() + "T00:00:00")) / 86400000);
        filas.push({ id: `${s.id}-estimado`, solicitudId: s.id, folio: s.folio, proveedor: prov, tipo: "Sin plan de pagos (fecha de entrega est.)", valor: total, fecha: s.fechaEstimada, dias, pagado: false });
      }
    }
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

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Calendario de pagos</h2>
          <p className="text-xs text-slate-400 mt-1">Pagos con plan confirmado, y solicitudes sin plan de pagos (usan su fecha estimada de entrega) — ordenados del más antiguo al más reciente.</p>
        </div>
        <button onClick={descargar} disabled={!filas.length} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium disabled:opacity-40 flex items-center gap-1"><FileText size={13} /> Descargar Excel</button>
      </div>

      {seleccionados.length > 0 && (
        <div className="bg-indigo-50 border border-indigo-200 rounded-xl px-4 py-3 flex items-center justify-between">
          <span className="text-sm text-slate-700">{seleccionados.length} pago(s) seleccionado(s)</span>
          <span className="text-lg font-semibold text-indigo-700">{fmt(totalSeleccionado)}</span>
        </div>
      )}

      {filas.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-sm text-slate-400">No hay pagos programados con plan confirmado todavía.</div>
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
  const [filtro, setFiltro] = useState("");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");

  const claveProveedor = (ev) => ev.proveedorId ? `id:${ev.proveedorId}` : `nombre:${(ev.proveedorNombre || "").trim().toLowerCase()}`;

  const todasCompletadas = solicitudes.filter((s) => s.evaluacionProveedor?.completada);
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
      <div>
        <h2 className="text-lg font-semibold text-slate-800">Evaluación de proveedores</h2>
        <p className="text-xs text-slate-400 mt-1">Promedio de resultados de un proveedor en un rango de fechas — útil para auditorías ISO 9001.</p>
      </div>

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

      {/* 2. PENDIENTES DE EVALUACIÓN */}
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
        const creado = await guardarProveedor({ nombre: proveedorNombre.trim(), nit: "", actividadEconomica: "", contacto: "", email: proveedorEmailNuevo.trim() });
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
        <input type="email" placeholder="Correo del proveedor (opcional)" value={proveedorEmailNuevo} onChange={(e) => setProveedorEmailNuevo(e.target.value)} className="w-full max-w-xs border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
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

function NuevaSolicitud({ areas, departamentos, empresas, itemsCatalogo, guardarItemCatalogo, proveedores, guardarProveedor, centrosCosto, conceptosGasto, usuarios, currentUser, onCrear, onCancel }) {
  const [tipo, setTipo] = useState("compra");
  const [empresaId, setEmpresaId] = useState(empresas[0]?.id || "");
  const [areaId, setAreaId] = useState(currentUser.areaId || areas[0].id);
  const [departamentoId, setDepartamentoId] = useState("");
  const [centroCostoId, setCentroCostoId] = useState(centrosCosto[0]?.id || "");
  const [conceptoGastoId, setConceptoGastoId] = useState(conceptosGasto[0]?.id || "");
  const [fechaEstimada, setFechaEstimada] = useState("");
  const [objetivo, setObjetivo] = useState("");
  const [justificacion, setJustificacion] = useState("");
  const [items, setItems] = useState([{ id: nextId(), itemCatalogoId: "", nombre: "", cantidad: 1, unidad: "unidad", precioEstimado: "", moneda: "COP", tasaCambio: 1, descuentoTipo: "porcentaje", descuentoValor: "", ivaEstimado: 19, cotizaciones: [] }]);
  const [pagosSugeridos, setPagosSugeridos] = useState(planPagosVacio());
  const [tienePlanPagos, setTienePlanPagos] = useState(false);

  const addItem = () => setItems([...items, { id: nextId(), itemCatalogoId: "", nombre: "", cantidad: 1, unidad: "unidad", precioEstimado: "", moneda: "COP", tasaCambio: 1, descuentoTipo: "porcentaje", descuentoValor: "", ivaEstimado: 19, cotizaciones: [] }]);
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
  const setFechaSugerida = (campo, val) => {
    if (val) {
      const error = validarOrdenFechas(pagosSugeridos, campo, val);
      if (error) { alert(error); return; }
    }
    setPagosSugeridos({ ...pagosSugeridos, [campo]: { ...pagosSugeridos[campo], fecha: val } });
  };
  const totalGeneral = items.reduce((acc, it) => { const d = desgloseItem(it); return { subtotal: acc.subtotal + d.subtotal, iva: acc.iva + d.iva, total: acc.total + d.total }; }, { subtotal: 0, iva: 0, total: 0 });

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
    if (!items.length || items.some((i) => !i.nombre.trim()) || !objetivo.trim() || !justificacion.trim()) return;
    // si el solicitante empezó a llenar el plan de pagos sugerido, debe cuadrar exacto con el total —
    // si lo dejó completamente vacío, no pasa nada, es opcional
    const algoDelPlanLlenado = parseFloat(pagosSugeridos.anticipo.valor) > 0 || parseFloat(pagosSugeridos.intermedio.valor) > 0 || parseFloat(pagosSugeridos.final.valor) > 0;
    if (algoDelPlanLlenado) {
      const restantePlan = totalGeneral.total - totalPagado(pagosSugeridos);
      if (Math.abs(restantePlan) > 0.5) {
        alert(restantePlan > 0 ? `El plan de pagos sugerido no cubre el total: faltan ${fmt(restantePlan)}. Complétalo o déjalo completamente vacío si no quieres sugerir uno.` : `El plan de pagos sugerido supera el total en ${fmt(-restantePlan)}. Ajústalo antes de enviar.`);
        return;
      }
      if (!pagosSugeridos.anticipo.fecha || !pagosSugeridos.final.fecha || (pagosSugeridos.intermedio.activo && !pagosSugeridos.intermedio.fecha)) {
        alert("Falta poner la fecha de uno o más pagos del plan sugerido antes de enviar.");
        return;
      }
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
      tipo, empresaId, areaId, departamentoId: departamentoId || null, centroCostoId, conceptoGastoId, solicitanteId: currentUser.id,
      fechaCreacion: hoy(), fechaEstimada, objetivo, justificacion,
      status: statusInicial,
      revisionCompras: tipo === "compra" ? { estado: "pendiente", observacion: "", usuario: "", fecha: "" } : { estado: "no_aplica", observacion: "", usuario: "", fecha: "" },
      items: items.map((i) => ({ ...i, cotizacionSeleccionada: null, observacionSeleccion: "" })),
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
      pagosSugeridos, pagos: planPagosVacio(), pagosConfirmados: false,
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
        <div><label className="text-xs font-medium text-slate-500">Empresa</label><select value={empresaId} onChange={(e) => setEmpresaId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{empresas.map((e) => <option key={e.id} value={e.id}>{e.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Área solicitante</label><select value={areaId} onChange={(e) => setAreaId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{areas.map((a) => <option key={a.id} value={a.id}>{a.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Departamento que reporta</label><select value={departamentoId} onChange={(e) => setDepartamentoId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm"><option value="">— Sin especificar —</option>{departamentos.map((d) => <option key={d.id} value={d.id}>{d.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Solicitante</label><div className="w-full mt-1 border border-slate-100 bg-slate-50 rounded-lg px-3 py-2 text-sm text-slate-500">{currentUser.nombre} (firma automática)</div></div>
        <div><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><Layers size={12} /> Centro de costo</label><select value={centroCostoId} onChange={(e) => setCentroCostoId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{centrosCosto.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></div>
        <div><label className="text-xs font-medium text-slate-500">Concepto de gasto</label><select value={conceptoGastoId} onChange={(e) => setConceptoGastoId(e.target.value)} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm">{conceptosGasto.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}</select></div>
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
        {items.map((it, idx) => (
          <div key={it.id} className={`${["bg-slate-50", "bg-indigo-50/60", "bg-amber-50/60", "bg-emerald-50/60", "bg-rose-50/50", "bg-sky-50/60"][idx % 6]} rounded-lg p-2 space-y-1.5 border border-black/5`}>
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
              <select value={it.ivaEstimado} onChange={(e) => updateItem(it.id, "ivaEstimado", e.target.value)} className="w-20 border border-slate-200 rounded-md px-2 py-1.5 text-sm">{IVA_OPCIONES.map((v) => <option key={v} value={v}>IVA {v}%</option>)}</select>
              {items.length > 1 && <button onClick={() => removeItem(it.id)} className="text-slate-400 hover:text-rose-500 p-1.5"><Trash2 size={15} /></button>}
            </div>
            {parseFloat(it.precioEstimado) > 0 && (() => {
              const d = desgloseItem(it);
              const inicial = parseFloat(it.precioEstimado) || 0;
              const descuento = parseFloat(it.descuentoValor) || 0;
              const conDescuento = descuento > 0 ? Math.max(0, it.descuentoTipo === "porcentaje" ? inicial * (1 - descuento / 100) : inicial - descuento) : inicial;
              const ahorro = inicial - conDescuento;
              const simbolo = it.moneda && it.moneda !== "COP" ? `${it.moneda} ` : "$";
              return (
                <div className="text-[11px] text-slate-500 pl-1 space-y-0.5">
                  {ahorro > 0 && <div>Precio inicial: {simbolo}{inicial.toLocaleString("es-CO")} · Descuento: -{simbolo}{ahorro.toLocaleString("es-CO", { maximumFractionDigits: 2 })} · Precio con descuento: <b>{simbolo}{conDescuento.toLocaleString("es-CO", { maximumFractionDigits: 2 })}</b></div>}
                  <div>Subtotal: {fmt(d.subtotal)} · IVA: {fmt(d.iva)} · <b>Total: {fmt(d.total)}</b></div>
                </div>
              );
            })()}
            {/* el solicitante puede adjuntar hasta 3 cotizaciones desde ya, opcional */}
            <CotizacionForm item={it} proveedores={proveedores} guardarProveedor={guardarProveedor} onGuardar={(_, cots) => setCotizacionesItem(it.id, cots)} compacto opcionalTitulo="Adjuntar cotizaciones (opcional, máx. 3)" />
          </div>
        ))}
      </div>

      <div className="mb-5 bg-slate-50 rounded-lg p-3 border border-slate-200">
        <div className="text-xs font-medium text-slate-500 mb-2 flex items-center gap-1"><CalendarClock size={13} /> ¿Esta solicitud cuenta con plan de pagos?</div>
        <div className="flex gap-2 mb-3">
          <button type="button" onClick={() => setTienePlanPagos(true)} className={`px-3 py-1.5 rounded-md text-xs font-medium border ${tienePlanPagos ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>Sí</button>
          <button type="button" onClick={() => { setTienePlanPagos(false); setPagosSugeridos(planPagosVacio()); }} className={`px-3 py-1.5 rounded-md text-xs font-medium border ${!tienePlanPagos ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>No</button>
        </div>
        {tienePlanPagos && (
          <>
        {!(totalGeneral.total > 0) && <div className="text-[11px] text-amber-600 mb-2">Pon un precio estimado en al menos un ítem para poder sugerir un plan de pagos.</div>}
        <div className="grid grid-cols-3 gap-2">
          <div><label className="text-[11px] mb-1 invisible block">Anticipo</label><InputMiles disabled={!(totalGeneral.total > 0)} placeholder="Anticipo" value={pagosSugeridos.anticipo.valor} onChange={(v) => setPagosSugeridos({ ...pagosSugeridos, anticipo: { ...pagosSugeridos.anticipo, valor: v } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><InputFecha disabled={!(totalGeneral.total > 0)} value={pagosSugeridos.anticipo.fecha} onChange={(v) => setFechaSugerida("anticipo", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
          <div><label className="text-[11px] flex items-center gap-1 mb-1"><input type="checkbox" disabled={!(totalGeneral.total > 0)} checked={pagosSugeridos.intermedio.activo} onChange={(e) => setPagosSugeridos({ ...pagosSugeridos, intermedio: { ...pagosSugeridos.intermedio, activo: e.target.checked } })} /> Intermedio</label><InputMiles placeholder="Valor" disabled={!(totalGeneral.total > 0) || !pagosSugeridos.intermedio.activo} value={pagosSugeridos.intermedio.valor} onChange={(v) => setPagosSugeridos({ ...pagosSugeridos, intermedio: { ...pagosSugeridos.intermedio, valor: v } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><InputFecha disabled={!(totalGeneral.total > 0) || !pagosSugeridos.intermedio.activo} value={pagosSugeridos.intermedio.fecha} onChange={(v) => setFechaSugerida("intermedio", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
          <div><label className="text-[11px] mb-1 invisible block">Pago final</label><InputMiles disabled={!(totalGeneral.total > 0)} placeholder="Pago final" value={pagosSugeridos.final.valor} onChange={(v) => setPagosSugeridos({ ...pagosSugeridos, final: { ...pagosSugeridos.final, valor: v } })} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><InputFecha disabled={!(totalGeneral.total > 0)} value={pagosSugeridos.final.fecha} onChange={(v) => setFechaSugerida("final", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
        </div>
        {totalGeneral.total > 0 && (parseFloat(pagosSugeridos.anticipo.valor) > 0 || parseFloat(pagosSugeridos.intermedio.valor) > 0 || parseFloat(pagosSugeridos.final.valor) > 0) && (() => {
          const restantePlan = totalGeneral.total - totalPagado(pagosSugeridos);
          const faltaFecha = !pagosSugeridos.anticipo.fecha || !pagosSugeridos.final.fecha || (pagosSugeridos.intermedio.activo && !pagosSugeridos.intermedio.fecha);
          return (
            <>
            <div className={`text-[11px] mt-2 ${Math.abs(restantePlan) > 0.5 ? "text-amber-600" : "text-emerald-600"}`}>
              {Math.abs(restantePlan) > 0.5 ? `Falta cuadrar: ${fmt(Math.abs(restantePlan))} ${restantePlan > 0 ? "por programar" : "de más"}` : "✓ El plan cuadra exacto con el total"}
            </div>
            {faltaFecha && <div className="text-[11px] text-amber-600">Falta poner la fecha de uno o más pagos.</div>}
            </>
          );
        })()}
        </>
        )}
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
            const d = desgloseItem(it);
            if (!(d.subtotal > 0)) return null;
            return <div key={it.id} className="flex justify-between text-slate-500"><span className="truncate pr-2">{idx + 1}. {it.nombre || "(sin nombre)"}</span><span className="shrink-0 text-slate-700">{fmt(d.total)}</span></div>;
          })}
        </div>
        <div className="border-t border-slate-200 pt-3 space-y-1">
          <div className="flex justify-between text-xs text-slate-500"><span>Subtotal</span><span>{fmt(totalGeneral.subtotal)}</span></div>
          <div className="flex justify-between text-xs text-slate-500"><span>IVA</span><span>{fmt(totalGeneral.iva)}</span></div>
          <div className="flex justify-between text-sm font-semibold text-slate-800 pt-1"><span>Total</span><span>{fmt(totalGeneral.total)}</span></div>
        </div>
        {requiereGerencia(totalGeneral.total) && <div className="text-[11px] text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2 py-1.5">Este monto requerirá aprobación de Gerencia.</div>}
        {!requiereGerencia(totalGeneral.total) && requiereDireccion(totalGeneral.total) && <div className="text-[11px] text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-2 py-1.5">Este monto requerirá aprobación de Dirección Financiera.</div>}
      </div>
    </div>
    </div>
  );
}

/* ---------------------------------------------------------
   COTIZACIONES Y COMPARATIVO
--------------------------------------------------------- */
function CotizacionForm({ item, proveedores, guardarProveedor, onGuardar, compacto, opcionalTitulo }) {
  const [abierto, setAbierto] = useState(!compacto);
  const [cots, setCots] = useState(item.cotizaciones.length ? item.cotizaciones : []);
  const [guardadoMsg, setGuardadoMsg] = useState(false);
  const update = (i, field, val) => setCots((prev) => prev.map((c, idx) => (idx === i ? { ...c, [field]: val } : c)));
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
  const addCot = () => cots.length < 3 && setCots([...cots, { proveedorId: "", proveedorNombre: "", unidadCotizada: item.unidad, factorConversion: 1, precioUnitario: item.precioEstimado || "", precioFinal: "", moneda: item.moneda || "COP", tasaCambio: item.tasaCambio || 1, descuentoTipo: "porcentaje", descuentoValor: "", diasEntrega: "", condicionesScore: 5, ivaPct: item.ivaEstimado ?? 19, archivoNombre: "" }]);
  const removeCot = (i) => setCots(cots.filter((_, idx) => idx !== i));

  // guarda automáticamente lo que ya se alcanzó a escribir (incluido el archivo adjunto), sin depender
  // de que se le dé clic al botón — así nada se pierde si alguien solo adjunta el archivo y no le da "Guardar"
  useEffect(() => {
    const validas = cots.filter((c) => (c.proveedorId || c.proveedorNombre) && c.precioUnitario);
    onGuardar(item.id, validas);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cots]);

  const guardar = async () => {
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
        const existente = proveedores.find((p) => p.nombre.trim().toLowerCase() === nombreLimpio.toLowerCase());
        if (!existente) {
          const creado = await guardarProveedor({ nombre: nombreLimpio, nit: "", actividadEconomica: "", contacto: "", email: emailNuevo });
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
    <div className="bg-slate-50 rounded-lg p-3 border border-slate-200">
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
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">IVA</label>
                <select value={c.ivaPct} onChange={(e) => update(i, "ivaPct", e.target.value)} className="border border-slate-200 rounded-md px-1 py-1.5 text-xs">{IVA_OPCIONES.map((v) => <option key={v} value={v}>IVA {v}%</option>)}</select>
              </div>
              <div className="flex flex-col gap-0.5">
                <label className="text-[10px] text-slate-400">Días entrega</label>
                <div className="flex gap-1"><input type="number" value={c.diasEntrega} onChange={(e) => update(i, "diasEntrega", e.target.value)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs" />{cots.length > 1 && <button onClick={() => removeCot(i)} className="text-slate-400 hover:text-rose-500 shrink-0"><Trash2 size={13} /></button>}</div>
              </div>
            </div>

            {!c.proveedorId && c.proveedorNombre?.trim() && (
              <div className="flex flex-col gap-0.5 max-w-xs">
                <label className="text-[10px] text-slate-400">Correo del proveedor (opcional)</label>
                <input type="email" value={c.proveedorEmailNuevo || ""} onChange={(e) => update(i, "proveedorEmailNuevo", e.target.value)} className="border border-slate-200 rounded-md px-2 py-1.5 text-xs" />
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

            <div className="flex items-start justify-between gap-3 flex-wrap">
              <AdjuntarArchivo small nombre={c.archivoNombre} label="Adjuntar cotización (PDF/foto)" onSeleccionar={(n) => update(i, "archivoNombre", n)} />
              {(c.proveedorId || c.proveedorNombre) && c.precioUnitario && (() => {
                const factor = parseFloat(c.factorConversion) || 1;
                const precioPorUnidad = precioEquivalente(c);
                return (
                  <div className="bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2 text-xs text-slate-600 space-y-0.5">
                    {factor !== 1 && <div>1 {c.unidadCotizada} = {factor} {item.unidad} → {fmt(precioFinalEfectivo(c) * (c.moneda && c.moneda !== "COP" ? (parseFloat(c.tasaCambio) || 1) : 1))} ÷ {factor} = <b>{fmt(precioPorUnidad)}</b> por {item.unidad}</div>}
                    <div>{item.cantidad} {item.unidad} × {fmt(precioPorUnidad)} = Subtotal <b>{fmt(d.subtotal)}</b></div>
                    <div>+ IVA {c.ivaPct}%: <b>{fmt(d.iva)}</b></div>
                    <div className="text-emerald-700 font-semibold text-sm pt-0.5 border-t border-emerald-200 mt-1">= Total: {fmt(d.total)}</div>
                  </div>
                );
              })()}
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
        <thead className="bg-white text-slate-500 border-b border-slate-100"><tr><th className="text-left px-3 py-2">Proveedor</th><th className="text-right px-3 py-2">Precio inicial</th><th className="text-right px-3 py-2">Descuento</th><th className="text-right px-3 py-2">Precio final</th><th className="text-right px-3 py-2">Cant.</th><th className="text-right px-3 py-2">Total (COP)</th><th className="text-right px-3 py-2">Entrega</th><th className="text-right px-3 py-2">Score</th><th className="px-3 py-2"></th></tr></thead>
        <tbody>{scored.map((c, i) => (
          <tr key={i} className={`border-t border-slate-100 ${i === bestIdx ? "bg-emerald-50/60" : ""}`}>
            <td className="px-3 py-2 font-medium text-slate-700 flex items-center gap-1">{i === bestIdx && <Award size={13} className="text-emerald-600" />} {nombreProv(c)} {c.archivoNombre && <EnlacePrivado path={c.archivoNombre} className="text-slate-400 hover:text-indigo-600" title="Ver cotización adjunta"><Paperclip size={11} /></EnlacePrivado>}</td>
            <td className="px-3 py-2 text-right">{c.precioUnitario ? `${c.moneda && c.moneda !== "COP" ? c.moneda + " " : ""}${Number(c.precioUnitario).toLocaleString("es-CO")}` : "—"}</td>
            <td className="px-3 py-2 text-right">{c.descuentoValor ? (c.descuentoTipo === "valor" ? `-${Number(c.descuentoValor).toLocaleString("es-CO")}` : `-${c.descuentoValor}%`) : "—"}</td>
            <td className="px-3 py-2 text-right">{c.moneda && c.moneda !== "COP" ? `${c.moneda} ${precioFinalEfectivo(c).toLocaleString("es-CO")}` : fmt(precioFinalEfectivo(c))}</td>
            <td className="px-3 py-2 text-right text-slate-400">× {item.cantidad}</td>
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
function PagosEstructurados({ solicitud, total, currentUser, onProgramar, onConfirmar, onEditarDeNuevo }) {
  const [pagos, setPagos] = useState(solicitud.pagos);
  // una vez la orden ya se envió al proveedor, las condiciones de pago quedan fijas — ya no se pueden tocar
  const ocYaEnviada = ["oc_enviada", "recepcion", "completada"].includes(solicitud.status);
  // sin precios (ni cotización ni estimado), no hay contra qué cuadrar el plan — se habilita cuando Compras cargue precios
  const sinPrecio = !(total > 0);
  const editable = puedeEditarPagos(currentUser) && !solicitud.pagosConfirmados && !ocYaEnviada && !sinPrecio;
  const pagado = totalPagado(pagos);
  const restante = total - pagado;
  const sug = solicitud.pagosSugeridos;
  const hasSugerencia = parseFloat(sug?.anticipo?.valor) > 0 || parseFloat(sug?.final?.valor) > 0;
  const descuadrado = solicitud.pagosConfirmados && Math.abs(restante) > 0.5;

  const set = (campo, sub, val) => {
    if (sub === "fecha" && val) {
      const error = validarOrdenFechas(pagos, campo, val);
      if (error) { alert(error); return; }
    }
    const copy = { ...pagos, [campo]: { ...pagos[campo], [sub]: val } }; setPagos(copy); onProgramar(copy);
  };
  const usarSugerencia = () => { setPagos(sug); onProgramar(sug); };

  const confirmar = () => {
    if (Math.abs(restante) > 0.5) {
      alert(restante > 0 ? `El plan de pagos no cubre el total: faltan ${fmt(restante)} por programar.` : `El plan de pagos supera el total en ${fmt(-restante)}. Ajusta los valores antes de confirmar.`);
      return;
    }
    if (!pagos.anticipo.fecha || !pagos.final.fecha || (pagos.intermedio.activo && !pagos.intermedio.fecha)) {
      alert("Falta poner la fecha de uno o más pagos antes de confirmar.");
      return;
    }
    onConfirmar();
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5">
      <div className="flex items-center justify-between mb-1">
        <div className="flex items-center gap-2 text-slate-700 font-medium"><CalendarClock size={16} /> Plan de pagos (máx. 3)</div>
        <div className="flex items-center gap-2">
          {solicitud.pagosConfirmados ? <Badge tone={descuadrado ? "red" : "green"}>Confirmado por Dirección Financiera</Badge> : <Badge tone="amber">Pendiente de confirmación</Badge>}
          {solicitud.pagosConfirmados && !ocYaEnviada && puedeEditarPagos(currentUser) && <button onClick={onEditarDeNuevo} className="text-[11px] text-indigo-600 underline">Editar de nuevo</button>}
        </div>
      </div>
      {ocYaEnviada && (
        <div className="text-[11px] text-slate-400 bg-slate-50 border border-slate-200 rounded-md px-3 py-2 mb-2">
          🔒 La orden ya fue enviada al proveedor — las condiciones de pago quedaron fijas y no se pueden modificar.
        </div>
      )}
      {sinPrecio && !ocYaEnviada && (
        <div className="text-[11px] text-amber-600 bg-amber-50 border border-amber-200 rounded-md px-3 py-2 mb-2">
          ⚠ Todavía no hay ningún precio (estimado ni cotizado) para esta solicitud, así que no hay contra qué cuadrar el plan de pagos. Se habilita en cuanto Compras cargue al menos una cotización.
        </div>
      )}
      {descuadrado && (
        <div className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-md px-3 py-2 mb-2">
          ⚠ Este plan quedó confirmado con un descuadre de <b>{fmt(Math.abs(restante))}</b> ({restante > 0 ? "falta programar" : "programado de más"}) — probablemente de antes de esta validación. Usa "Editar de nuevo" para corregirlo.
        </div>
      )}
      {hasSugerencia && !solicitud.pagosConfirmados && (
        <div className="text-[11px] text-slate-500 bg-slate-50 border border-slate-200 rounded-md px-2 py-1.5 mb-2 flex items-center justify-between">
          <span>El solicitante sugirió: anticipo {fmt(sug.anticipo.valor)} ({sug.anticipo.fecha || "sin fecha"}){sug.intermedio.activo ? `, intermedio ${fmt(sug.intermedio.valor)}` : ""}, final {fmt(sug.final.valor)} ({sug.final.fecha || "sin fecha"})</span>
          {editable && <button onClick={usarSugerencia} className="text-indigo-600 font-medium ml-2 shrink-0">Usar sugerencia</button>}
        </div>
      )}
      {!editable && !solicitud.pagosConfirmados && <div className="text-[11px] text-slate-400 mb-3">Solo Dirección Financiera puede editar y confirmar este plan.</div>}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mt-2">
        <div className="border border-slate-200 rounded-lg p-3">
          <div className="text-xs font-medium text-slate-500 mb-2">Anticipo</div>
          <InputMiles disabled={!editable} placeholder="Valor anticipo" value={pagos.anticipo.valor} onChange={(v) => set("anticipo", "valor", v)} className="w-full mb-1.5 border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <InputFecha disabled={!editable} value={pagos.anticipo.fecha} onChange={(v) => set("anticipo", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <div className="text-[11px] text-slate-400 mt-1">Fecha primer pago</div>
        </div>
        <div className="border border-slate-200 rounded-lg p-3">
          <label className="text-xs font-medium text-slate-500 mb-2 flex items-center gap-1.5"><input disabled={!editable} type="checkbox" checked={pagos.intermedio.activo} onChange={(e) => set("intermedio", "activo", e.target.checked)} /> Pago intermedio (opcional)</label>
          <InputMiles disabled={!editable || !pagos.intermedio.activo} placeholder="Valor intermedio" value={pagos.intermedio.valor} onChange={(v) => set("intermedio", "valor", v)} className="w-full mb-1.5 border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <InputFecha disabled={!editable || !pagos.intermedio.activo} value={pagos.intermedio.fecha} onChange={(v) => set("intermedio", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
        </div>
        <div className="border border-slate-200 rounded-lg p-3">
          <div className="text-xs font-medium text-slate-500 mb-2">Pago final</div>
          <InputMiles disabled={!editable} placeholder="Valor pago final" value={pagos.final.valor} onChange={(v) => set("final", "valor", v)} className="w-full mb-1.5 border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <InputFecha disabled={!editable} value={pagos.final.fecha} onChange={(v) => set("final", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-sm disabled:bg-slate-50" />
          <div className="text-[11px] text-slate-400 mt-1">Fecha pago final</div>
        </div>
      </div>
      <div className="flex items-center justify-between mt-3">
        <div className="text-sm text-slate-500">Total orden: <b className="text-slate-700">{fmt(total)}</b> · Programado: <b className="text-slate-700">{fmt(pagado)}</b> · Restante: <b className={restante > 0.5 ? "text-amber-600" : restante < -0.5 ? "text-rose-600" : "text-emerald-600"}>{fmt(restante)}</b></div>
        {editable && <button onClick={confirmar} className="text-xs bg-emerald-600 text-white px-3 py-1.5 rounded-md font-medium">Confirmar plan de pagos</button>}
      </div>
      {editable && Math.abs(restante) > 0.5 && <div className="text-[11px] text-amber-600 mt-1">El plan debe cubrir exactamente el total de la orden para poder confirmarse.</div>}
    </div>
  );
}

/* ---------------------------------------------------------
   ORDEN ENVIADA AL PROVEEDOR
--------------------------------------------------------- */
function OcEnviadaPanel({ solicitud, proveedores, empresa, currentUser, onGuardar }) {
  const [firmandoIdx, setFirmandoIdx] = useState(null);
  if (solicitud.status !== "orden") return null;

  const necesarios = proveedoresAdjudicadosDetalle(solicitud, proveedores);
  const ordenes = necesarios.map((n) => {
    const existente = (solicitud.ocEnviada.ordenesProveedor || []).find((o) => mismoProveedor(o, n));
    return existente || { ...n, archivoOriginalUrl: "", archivoFirmadoUrl: "", fecha: "", usuario: "" };
  });

  const actualizarOrden = (idx, cambios) => {
    const copia = ordenes.map((o, i) => (i === idx ? { ...o, ...cambios } : o));
    onGuardar({ ordenesProveedor: copia });
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
      <div className="text-xs text-slate-500">Se detectaron <b>{ordenes.length}</b> proveedor(es) adjudicado(s) en esta solicitud. Sube aquí la orden generada en el sistema contable de cada uno; se envía a <b>Dirección Financiera</b> para su firma digital. Al marcar "OC enviada al proveedor", la orden ya firmada se envía automáticamente por correo a cada proveedor <b>que tenga correo registrado en el Catálogo</b> — si alguno no lo tiene, tendrás que enviársela tú manualmente.</div>

      {ordenes.map((o, i) => (
        <div key={i} className="border border-slate-200 rounded-lg p-3 space-y-2">
          <div className="text-sm font-medium text-slate-700 flex items-center gap-2"><Truck size={13} /> {o.proveedorNombre}</div>
          <AdjuntarArchivo nombre={o.archivoOriginalUrl} label={`Adjuntar OC para ${o.proveedorNombre} (solo PDF)`} onSeleccionar={(url) => actualizarOrden(i, { archivoOriginalUrl: url, archivoFirmadoUrl: "", fecha: "", usuario: "" })} carpeta="ordenes-originales" soloPdf />
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
          `Reenvío — Orden de compra/servicio ${solicitud.folio}`,
          `<p>Hola ${prov?.nombre || orden.proveedorNombre},</p><p>Te reenviamos el enlace de la orden de compra/servicio <b>${solicitud.folio}</b> a nombre de ${empresa?.nombre || ""}.</p><p><a href="${url}">Ver / descargar la orden firmada</a></p><p>Este enlace estará disponible por 7 días.</p>`
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
  return (
    <select value={value || ""} onChange={(e) => onChange(e.target.value ? Number(e.target.value) : null)} disabled={disabled} className="border border-slate-200 rounded-md px-2 py-1 text-xs w-16 disabled:bg-slate-50">
      <option value="">—</option>
      {Array.from({ length: 10 }, (_, i) => i + 1).map((n) => <option key={n} value={n}>{n}</option>)}
    </select>
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
function PagosSugeridosEditor({ solicitud, total, onGuardar }) {
  const sug = solicitud.pagosSugeridos || planPagosVacio();
  const tienePlan = parseFloat(sug.anticipo.valor) > 0 || parseFloat(sug.intermedio.valor) > 0 || parseFloat(sug.final.valor) > 0;
  const [mostrar, setMostrar] = useState(tienePlan);

  const set = (campo, sub, val) => {
    if (sub === "fecha" && val) {
      const error = validarOrdenFechas(sug, campo, val);
      if (error) { alert(error); return; }
    }
    onGuardar({ ...sug, [campo]: { ...sug[campo], [sub]: val } });
  };
  const restante = total - totalPagado(sug);

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
      <div className="text-sm font-medium text-slate-700 flex items-center gap-2"><CalendarClock size={15} /> ¿Esta solicitud cuenta con plan de pagos?</div>
      <div className="flex gap-2">
        <button type="button" onClick={() => setMostrar(true)} className={`px-3 py-1.5 rounded-md text-xs font-medium border ${mostrar ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>Sí</button>
        <button type="button" onClick={() => { setMostrar(false); onGuardar(planPagosVacio()); }} className={`px-3 py-1.5 rounded-md text-xs font-medium border ${!mostrar ? "bg-indigo-600 text-white border-indigo-600" : "bg-white text-slate-600 border-slate-200"}`}>No</button>
      </div>
      {mostrar && (
        <>
          {!(total > 0) && <div className="text-[11px] text-amber-600">Pon un precio estimado en al menos un ítem para poder sugerir un plan de pagos.</div>}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
            <div><label className="text-[11px] mb-1 invisible block">Anticipo</label><InputMiles disabled={!(total > 0)} placeholder="Anticipo" value={sug.anticipo.valor} onChange={(v) => set("anticipo", "valor", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><InputFecha disabled={!(total > 0)} value={sug.anticipo.fecha} onChange={(v) => set("anticipo", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
            <div><label className="text-[11px] flex items-center gap-1 mb-1"><input type="checkbox" disabled={!(total > 0)} checked={sug.intermedio.activo} onChange={(e) => set("intermedio", "activo", e.target.checked)} /> Intermedio</label><InputMiles placeholder="Valor" disabled={!(total > 0) || !sug.intermedio.activo} value={sug.intermedio.valor} onChange={(v) => set("intermedio", "valor", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><InputFecha disabled={!(total > 0) || !sug.intermedio.activo} value={sug.intermedio.fecha} onChange={(v) => set("intermedio", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
            <div><label className="text-[11px] mb-1 invisible block">Pago final</label><InputMiles disabled={!(total > 0)} placeholder="Pago final" value={sug.final.valor} onChange={(v) => set("final", "valor", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs mb-1 disabled:bg-slate-100" /><InputFecha disabled={!(total > 0)} value={sug.final.fecha} onChange={(v) => set("final", "fecha", v)} className="w-full border border-slate-200 rounded-md px-2 py-1.5 text-xs disabled:bg-slate-100" /></div>
          </div>
          {total > 0 && tienePlan && (
            <div className={`text-[11px] ${Math.abs(restante) > 0.5 ? "text-amber-600" : "text-emerald-600"}`}>
              {Math.abs(restante) > 0.5 ? `Falta cuadrar: ${fmt(Math.abs(restante))} ${restante > 0 ? "por programar" : "de más"}` : "✓ El plan cuadra exacto con el total"}
            </div>
          )}
          {total > 0 && tienePlan && (!sug.anticipo.fecha || !sug.final.fecha || (sug.intermedio.activo && !sug.intermedio.fecha)) && (
            <div className="text-[11px] text-amber-600">Falta poner la fecha de uno o más pagos.</div>
          )}
        </>
      )}
    </div>
  );
}

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

  useEffect(() => {
    if (esCompras && !ev.completada && !ev.proveedorId && proveedorSugerido) {
      // fecha en que quedó adjudicado el proveedor (paso "comparativo" del historial), si existe
      const fechaComparativo = solicitud.historialEstados?.find((h) => h.status === "comparativo")?.fecha;
      const copy = {
        ...ev,
        proveedorId: proveedorSugerido.id,
        proveedorNombre: proveedorSugerido.nombre,
        tipoProveedor: proveedorSugerido.tipoProveedor || solicitud.tipo,
        nit: proveedorSugerido.nit || "",
        ciudad: proveedorSugerido.ciudad || "",
        direccion: proveedorSugerido.direccion || "",
        telefono: proveedorSugerido.telefono || "",
        representanteLegal: proveedorSugerido.representanteLegal || "",
        email: proveedorSugerido.email || "",
        fechaSeleccion: ev.fechaSeleccion || (fechaComparativo ? fechaComparativo.slice(0, 10) : solicitud.fechaCreacion) || "",
        fechaEvaluacion: ev.fechaEvaluacion || hoy(),
        serviciosPresta: ev.serviciosPresta || proveedorSugerido.actividadEconomica || "",
        descripcion: ev.descripcion || solicitud.items.map((it) => it.nombre).join(", "),
      };
      setEv(copy);
      onGuardar(copy);
    }
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
      <div className="border border-slate-200 rounded-lg p-3 space-y-1.5">
        <div className="text-sm font-medium text-slate-700 mb-1">Documentos anexos</div>
        {DOCUMENTOS_EVALUACION.map((d) => (
          <div key={d.key} className="flex items-center justify-between text-xs gap-2">
            <span className="text-slate-600">{d.label}</span>
            <select disabled={disabled} value={ev.documentos?.[d.key] || ""} onChange={(e) => setDoc(d.key, e.target.value)} className="border border-slate-200 rounded-md px-2 py-1 disabled:bg-slate-50 shrink-0">
              <option value="">—</option><option value="si">Sí</option><option value="no">No</option><option value="no_aplica">No aplica</option>
            </select>
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


function RecepcionPanel({ solicitud, currentUser, onGuardar }) {
  const [r, setR] = useState({ ...solicitud.recepcion, archivos: solicitud.recepcion.archivos || (solicitud.recepcion.archivoNombre ? [solicitud.recepcion.archivoNombre] : []) });
  const set = (fields) => { const copy = { ...r, ...fields, usuario: currentUser.nombre, fecha: hoy() }; setR(copy); onGuardar(copy); };
  const agregarArchivo = (url) => set({ archivos: [...r.archivos, url] });
  const quitarArchivo = (i) => set({ archivos: r.archivos.filter((_, idx) => idx !== i) });

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
      <div className="font-medium text-slate-700 flex items-center gap-2">
        <PackageCheck size={16} /> Recepción
        {r.recibidoSatisfaccion ? <Badge tone="green">Recibida</Badge> : <Badge tone="amber">Pendiente de recepción</Badge>}
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
      <div><label className="text-xs font-medium text-slate-500 flex items-center gap-1"><MessageSquare size={12} /> Comentarios (opcional)</label><textarea value={r.comentario} onChange={(e) => set({ comentario: e.target.value })} rows={2} className="w-full mt-1 border border-slate-200 rounded-lg px-3 py-2 text-sm resize-none" /></div>
      <label className="flex items-center gap-2 text-sm text-slate-700"><input type="checkbox" checked={r.recibidoSatisfaccion} onChange={(e) => set({ recibidoSatisfaccion: e.target.checked })} /> Recibo a satisfacción</label>
      {!r.recibidoSatisfaccion && <div className="text-[11px] text-amber-600">Marca "Recibo a satisfacción" para que Compras pueda hacer la evaluación y finalizar la solicitud.</div>}
      {r.recibidoSatisfaccion && <div className="text-[11px] text-emerald-600">✓ Recibida — falta que Compras complete la evaluación del proveedor para finalizar.</div>}
    </div>
  );
}

/* ---------------------------------------------------------
   ORDEN DE COMPRA / TRABAJO — documento consolidado
--------------------------------------------------------- */
function OrdenDocumento({ solicitud, empresa, area, departamento, solicitante, proveedores, centrosCosto, conceptosGasto }) {
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
            <div className="text-base font-semibold text-slate-800">{solicitud.tipo === "compra" ? "Solicitud de Compra" : "Orden de servicio/trabajo"}</div>
            <div className="text-xs text-slate-400">{solicitud.folio} · {empresa?.nombre}</div>
          </div>
        </div>
        <Badge tone="blue">{PASOS.find((p) => p.key === solicitud.status)?.label}</Badge>
      </div>

      <div className="grid grid-cols-2 gap-2 text-xs text-slate-600">
        <div><b>Área:</b> {area?.nombre}{departamento && ` — ${departamento.nombre}`}</div><div><b>Solicitante:</b> {solicitante?.nombre}</div>
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

      {/* PLAN DE PAGOS (solo servicio) */}
      {pagoActivo && (() => {
        const usaSugerido = !solicitud.pagosConfirmados && !(solicitud.pagos.anticipo.valor > 0 || solicitud.pagos.final.valor > 0);
        const p = usaSugerido ? solicitud.pagosSugeridos : solicitud.pagos;
        return (
          <div>
            <div className="text-xs font-medium text-slate-500 mb-1">Plan de pagos {solicitud.pagosConfirmados ? "(confirmado por Dirección Financiera)" : usaSugerido ? "(sugerido por el solicitante — pendiente de confirmar)" : "(sin confirmar)"}</div>
            <table className="w-full text-[11px]">
              <thead className="text-slate-400 border-b border-slate-100"><tr><th className="text-left py-0.5">Pago</th><th className="text-right py-0.5">Valor</th><th className="text-right py-0.5">Fecha</th><th className="text-center py-0.5">Pagado</th></tr></thead>
              <tbody>
                <tr className="border-t border-slate-50"><td className="py-0.5">Anticipo</td><td className="py-0.5 text-right">{fmt(p.anticipo.valor)}</td><td className="py-0.5 text-right">{p.anticipo.fecha || "—"}</td><td className="py-0.5 text-center">{p.anticipo.pagado ? "✓" : ""}</td></tr>
                {p.intermedio.activo && <tr className="border-t border-slate-50"><td className="py-0.5">Intermedio</td><td className="py-0.5 text-right">{fmt(p.intermedio.valor)}</td><td className="py-0.5 text-right">{p.intermedio.fecha || "—"}</td><td className="py-0.5 text-center">{p.intermedio.pagado ? "✓" : ""}</td></tr>}
                <tr className="border-t border-slate-50"><td className="py-0.5">Final</td><td className="py-0.5 text-right">{fmt(p.final.valor)}</td><td className="py-0.5 text-right">{p.final.fecha || "—"}</td><td className="py-0.5 text-center">{p.final.pagado ? "✓" : ""}</td></tr>
              </tbody>
            </table>
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
          <div><div className="font-medium text-slate-500 mb-0.5">Recepción</div><div className="text-slate-600">{solicitud.recepcion.recibidoSatisfaccion ? "Recibido a satisfacción" : "Pendiente"}{solicitud.recepcion.archivos?.length > 0 && ` · ${solicitud.recepcion.archivos.length} archivo(s) adjunto(s)`}{solicitud.recepcion.comentario && <div className="italic">"{solicitud.recepcion.comentario}"</div>}</div></div>
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
    case "aprobacion_financiera": return requiereGerencia(total) ? "Aprobar y enviar a Gerencia" : "Aprobar y generar orden";
    case "aprobacion_gerencia": return "Aprobar y generar orden";
    case "orden": return "Marcar OC enviada al proveedor";
    case "oc_enviada": return "Confirmar recepción / iniciar ejecución";
    case "recepcion": return "Marcar como completada";
    default: return "Avanzar";
  }
}

function SolicitudDetalle({ solicitud, areas, departamentos, empresas, usuarios, proveedores, guardarProveedor, itemsCatalogo, centrosCosto, conceptosGasto, historico, setHistorico, currentUser, onUpdate, onEliminar, onVolver }) {
  const [observacion, setObservacion] = useState("");
  const [prioridadSel, setPrioridadSel] = useState(solicitud.prioridad || "Medio");
  const area = areas.find((a) => a.id === solicitud.areaId);
  const departamento = departamentos.find((d) => d.id === solicitud.departamentoId);
  const empresa = empresas.find((e) => e.id === solicitud.empresaId);
  const solicitante = usuarios.find((u) => u.id === solicitud.solicitanteId);
  const total = totalSolicitud(solicitud);
  const todasCotizadas = solicitud.items.every((i) => i.cotizaciones.length > 0);
  const comparativoBloqueado = ["orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status);
  const patch = (fields) => onUpdate({ ...solicitud, ...fields });

  const reabrirSolicitud = () => {
    const { status: destino, campo, revision } = pasoDelRechazo(solicitud);
    const cambios = {
      status: destino,
      historialEstados: empujarHistorial(destino),
      notificaciones: notificar(`Solicitud reabierta por ${currentUser.nombre} (${currentUser.rol}) para corregir y volver a enviar.`),
    };
    if (campo) cambios.firmas = { ...solicitud.firmas, [campo]: { aprobado: null, nombre: null, fecha: null, observacion: "", fotoUrl: null } };
    if (revision) cambios.revisionCompras = { estado: "pendiente", observacion: "", usuario: "", fecha: "" };
    patch(cambios);
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
    if (s === "aprobacion_jefe") {
      if (currentUser.rol === "Jefe de Área y Director") {
        // la misma persona hace ambos roles: se aprueban los dos pasos de una vez, sin duplicar el clic
        patch({ status: "cotizando", prioridad: prioridadSel, firmas: { ...solicitud.firmas, jefe: firmar(), director: { ...firmar(), observacion: "Aprobado junto con el paso de jefe de área (mismo responsable)." } }, historialEstados: empujarHistorial("cotizando"), notificaciones: notificar(`Correo simulado a Compras: solicitud ${solicitud.folio} aprobada (jefe y director), lista para cotizar.`) });
      } else {
        const director = usuarios.find((u) => u.areaId === solicitud.areaId && ["Director de Área", "Jefe de Área y Director"].includes(u.rol));
        patch({ status: "aprobacion_director", prioridad: prioridadSel, firmas: { ...solicitud.firmas, jefe: firmar() }, historialEstados: empujarHistorial("aprobacion_director"), notificaciones: notificar(director?.email ? `Correo enviado a ${director.nombre} (${director.email}): solicitud ${solicitud.folio} pendiente de tu aprobación.` : `Solicitud aprobada por el jefe de área. No hay un director de área con correo configurado para notificar.`) });
      }
    }
    else if (s === "aprobacion_director") patch({ status: "cotizando", firmas: { ...solicitud.firmas, director: firmar() }, historialEstados: empujarHistorial("cotizando"), notificaciones: notificar(`Correo simulado a Compras: solicitud ${solicitud.folio} aprobada, lista para cotizar.`) });
    else if (s === "cotizando" && todasCotizadas) patch({ status: "comparativo", historialEstados: empujarHistorial("comparativo") });
    else if (s === "comparativo") { const next = requiereDireccion(total) ? "aprobacion_financiera" : requiereGerencia(total) ? "aprobacion_gerencia" : "orden"; patch({ status: next, historialEstados: empujarHistorial(next), notificaciones: notificar(`Correo simulado: solicitud ${solicitud.folio} avanza a ${PASOS.find((p) => p.key === next)?.label}.`) }); }
    else if (s === "aprobacion_financiera") {
      if (solicitud.tipo === "servicio" && !solicitud.pagosConfirmados) { alert("Falta confirmar el plan de pagos antes de aprobar y continuar."); return; }
      const next = requiereGerencia(total) ? "aprobacion_gerencia" : "orden";
      patch({ status: next, firmas: { ...solicitud.firmas, financiera: firmar() }, historialEstados: empujarHistorial(next) });
    }
    else if (s === "aprobacion_gerencia") patch({ status: "orden", firmas: { ...solicitud.firmas, gerencia: firmar() }, historialEstados: empujarHistorial("orden") });
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
      patch({ status: "oc_enviada", historialEstados: empujarHistorial("oc_enviada"), notificaciones: notificar(`Correo enviado a ${solicitante?.nombre} (${solicitante?.email || "sin correo"}).${detalleProveedores}${avisoSinCorreo}`) });
      if (solicitante?.email) {
        enviarCorreo(
          solicitante.email,
          `Tu orden ${solicitud.folio} fue enviada al proveedor`,
          `<p>Hola ${solicitante.nombre},</p><p>La orden <b>${solicitud.folio}</b> ya fue enviada al proveedor y quedó lista para recepción.</p>`
        );
      }
      // envía la orden firmada por correo a cada proveedor que sí tenga correo registrado (hasta 2 correos por proveedor)
      ordenesConCorreo.forEach(({ orden, prov }) => {
        obtenerUrlFirmada(orden.archivoFirmadoUrl, 604800).then((url) => {
          if (!url) return;
          enviarCorreo(
            correosDe(prov),
            `Orden de compra/servicio ${solicitud.folio}`,
            `<p>Hola ${prov.nombre},</p><p>Adjuntamos el enlace de la orden de compra/servicio <b>${solicitud.folio}</b> a nombre de ${empresa?.nombre || ""}.</p><p><a href="${url}">Ver / descargar la orden firmada</a></p><p>Este enlace estará disponible por 7 días.</p>`
          );
        });
      });
    }
    else if (s === "oc_enviada") patch({ status: "recepcion", historialEstados: empujarHistorial("recepcion") });
    else if (s === "recepcion") {
      if (!solicitud.recepcion.recibidoSatisfaccion) return;
      if (!evaluacionCompleta(solicitud)) return;
      if (solicitud.tipo === "servicio") {
        const pagado = totalPagado(solicitud.pagos);
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
    const campo = solicitud.status === "aprobacion_jefe" ? "jefe" : solicitud.status === "aprobacion_director" ? "director" : solicitud.status === "aprobacion_financiera" ? "financiera" : solicitud.status === "aprobacion_gerencia" ? "gerencia" : null;
    patch({ status: "rechazada", firmas: campo ? { ...solicitud.firmas, [campo]: { aprobado: false, nombre: currentUser.nombre, cargo: currentUser.cargo || "", empresa: empresa?.nombre || "", fecha: hoy(), observacion, fotoUrl: currentUser.firmaFotoUrl || null } } : solicitud.firmas, historialEstados: empujarHistorial("rechazada"), notificaciones: notificar(`Correo enviado a ${solicitante?.nombre} (${solicitante?.email || "sin correo"}): tu solicitud ${solicitud.folio} fue rechazada.`) });
    if (solicitante?.email) {
      enviarCorreo(
        solicitante.email,
        `Tu solicitud ${solicitud.folio} fue rechazada`,
        `<p>Hola ${solicitante.nombre},</p><p>Tu solicitud <b>${solicitud.folio}</b> fue rechazada por ${currentUser.nombre} (${currentUser.rol}).</p>${observacion ? `<p><b>Motivo:</b> ${observacion}</p>` : ""}`
      );
    }
    setObservacion("");
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
    cotizando: ["Compras"],
    comparativo: ["Compras"],
    orden: ["Dirección Financiera"],
    recepcion: ["Compras"],
  };
  const pasoLeConcierne = (ROLES_RELEVANTES_POR_PASO[solicitud.status] || []).includes(currentUser.rol);

  return (
    <div className="space-y-5">
      <button onClick={onVolver} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-700"><ArrowLeft size={15} /> Volver a solicitudes</button>

      <div className="bg-white rounded-xl border border-slate-200 p-5">
        <div className="flex items-start justify-between flex-wrap gap-3">
          {empresa?.logoUrl && <img src={empresa.logoUrl} alt={empresa.nombre} className="h-10 max-w-[100px] object-contain order-first" />}
          <div>
            <div className="flex items-center gap-2 flex-wrap"><h2 className="text-lg font-semibold text-slate-800">{solicitud.folio}</h2><Badge tone={solicitud.tipo === "compra" ? "blue" : "amber"}>{solicitud.tipo === "compra" ? <ShoppingCart size={12} /> : <Wrench size={12} />} {solicitud.tipo === "compra" ? "Solicitud de compra" : "Orden de servicio/trabajo"}</Badge>{solicitud.prioridad && <Badge tone={solicitud.prioridad === "Alto" ? "red" : solicitud.prioridad === "Medio" ? "amber" : "slate"}>Prioridad {solicitud.prioridad}</Badge>}{["recepcion", "completada"].includes(solicitud.status) && solicitud.recepcion?.recibidoSatisfaccion && <Badge tone="green">Recibida</Badge>}{solicitud.status === "rechazada" && <Badge tone="red">Rechazada</Badge>}<button onClick={() => window.print()} className="text-xs bg-slate-800 text-white px-3 py-1.5 rounded-md font-medium flex items-center gap-1 no-print"><FileText size={13} /> Exportar solicitud completa a PDF</button>{currentUser.rol === "Administrador" && <button onClick={() => onEliminar(solicitud.id, solicitud.folio)} className="text-xs bg-rose-50 text-rose-600 border border-rose-200 px-3 py-1.5 rounded-md font-medium flex items-center gap-1 no-print"><Trash2 size={13} /> Eliminar solicitud</button>}</div>
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

      {solicitud.status === "rechazada" && (puedeReabrir(currentUser) || currentUser.id === solicitud.solicitanteId) && (
        <div className="bg-rose-50 border border-rose-200 rounded-xl p-4 flex items-center justify-between gap-3 flex-wrap">
          <div className="text-sm text-rose-700">Esta solicitud fue rechazada. Si el motivo fue un error que ya se corrigió (ej. en los precios estimados o en el plan de pagos), puedes reabrirla — volverá al paso donde fue rechazada.</div>
          <button onClick={reabrirSolicitud} className="text-xs bg-rose-600 text-white px-3 py-1.5 rounded-md font-medium shrink-0">Reabrir para corregir</button>
        </div>
      )}

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
          const d = desgloseItem(it);
          const unitario = parseFloat(it.cantidad) > 0 ? d.subtotal / parseFloat(it.cantidad) : 0;
          const cotConArchivo = (it.cotizaciones || []).filter((c) => c.archivoNombre);
          const puedeEditarPrecio = (currentUser.id === solicitud.solicitanteId || puedeReabrir(currentUser)) && ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && !(it.cotizaciones?.length > 0);
          return (
            <tr key={it.id} className="border-t border-slate-100 align-top">
              <td className="py-2 pr-2 text-slate-400">{idx + 1}</td>
              <td className="py-2 pr-3">{it.nombre}</td>
              <td className="py-2 px-2 text-right whitespace-nowrap">{it.cantidad}</td>
              <td className="py-2 px-2 text-right whitespace-nowrap">{it.unidad}</td>
              <td className="py-2 px-2 text-right whitespace-nowrap">
                {puedeEditarPrecio ? (
                  <InputMiles value={it.precioEstimado} onChange={(v) => patch({ items: solicitud.items.map((x) => (x.id === it.id ? { ...x, precioEstimado: v } : x)) })} className="w-24 border border-slate-200 rounded-md px-2 py-1 text-xs text-right" />
                ) : (
                  unitario > 0 ? fmt(unitario) : "—"
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
        <PagosSugeridosEditor solicitud={solicitud} total={total} onGuardar={(sug) => patch({ pagosSugeridos: sug })} />
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
          {solicitud.items.map((it) => <CotizacionForm key={it.id} item={it} proveedores={proveedores} guardarProveedor={guardarProveedor} onGuardar={guardarCotizaciones} />)}
        </div>
      )}

      {["aprobacion_jefe", "aprobacion_director", "comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && solicitud.items.some((i) => i.cotizaciones.length > 0) && (
        <div className="bg-white rounded-xl border border-slate-200 p-5 space-y-3">
          <div className="font-medium text-slate-700 flex items-center gap-2"><TrendingUp size={16} /> Cuadro comparativo (sugerencia automática)</div>
          {["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status) && <div className="text-xs text-slate-400">Cotizaciones que el solicitante adjuntó al crear la solicitud — Compras podrá completar y ajustar esto más adelante.</div>}
          {solicitud.items.filter((i) => i.cotizaciones.length > 0).map((it) => <ComparativoTabla key={it.id} item={it} proveedores={proveedores} onSeleccionar={seleccionarCotizacion} seleccionada={it.cotizacionSeleccionada} soloLectura={comparativoBloqueado || ["aprobacion_jefe", "aprobacion_director"].includes(solicitud.status)} />)}
        </div>
      )}

      {["aprobacion_jefe", "aprobacion_director", "comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && <ResumenTotales solicitud={solicitud} />}

      {solicitud.tipo === "servicio" && ["aprobacion_jefe", "aprobacion_director", "cotizando", "comparativo", "aprobacion_financiera", "aprobacion_gerencia", "orden", "oc_enviada", "recepcion", "completada"].includes(solicitud.status) && (
        <PagosEstructurados solicitud={solicitud} total={total} currentUser={currentUser} onProgramar={(pagos) => patch({ pagos })} onConfirmar={() => patch({ pagosConfirmados: true })} onEditarDeNuevo={() => patch({ pagosConfirmados: false })} />
      )}

      <OcEnviadaPanel solicitud={solicitud} proveedores={proveedores} empresa={empresa} currentUser={currentUser} onGuardar={(oc) => patch({ ocEnviada: oc })} />

      <ReenviarOrdenesPanel solicitud={solicitud} proveedores={proveedores} guardarProveedor={guardarProveedor} empresa={empresa} currentUser={currentUser} />

      {["recepcion", "completada"].includes(solicitud.status) && <RecepcionPanel solicitud={solicitud} currentUser={currentUser} onGuardar={(r) => patch({ recepcion: r })} />}

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
        <OrdenDocumento solicitud={solicitud} empresa={empresa} area={area} departamento={departamento} solicitante={solicitante} proveedores={proveedores} centrosCosto={centrosCosto} conceptosGasto={conceptosGasto} />
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
            if (solicitud.status === "cotizando" && !todasCotizadas) motivos.push("Falta cargar al menos una cotización para cada ítem.");
            if (solicitud.status === "aprobacion_financiera" && solicitud.tipo === "servicio" && !solicitud.pagosConfirmados) motivos.push("Falta confirmar el plan de pagos.");
            if (solicitud.status === "orden" && !todasOrdenesFirmadas(solicitud, proveedores)) motivos.push("Falta que Dirección Financiera firme la orden de uno o más proveedores.");
            if (solicitud.status === "recepcion") {
              if (!solicitud.recepcion.recibidoSatisfaccion) motivos.push("Falta marcar \"Recibo a satisfacción\" en el panel de Recepción.");
              if (!evaluacionCompleta(solicitud)) motivos.push("Falta completar la evaluación del proveedor (14 criterios).");
            }
            return motivos.length > 0 && (
              <div className="text-xs text-amber-600 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 space-y-0.5">
                <div className="font-medium">Falta esto para poder continuar:</div>
                {motivos.map((m, i) => <div key={i}>• {m}</div>)}
              </div>
            );
          })()}
          <div className="flex gap-2 justify-end">
            <button onClick={rechazar} className="px-4 py-2 rounded-lg text-sm text-rose-600 border border-rose-200 flex items-center gap-1"><XCircle size={15} /> Rechazar</button>
            <button onClick={avanzar} disabled={(solicitud.status === "cotizando" && !todasCotizadas) || (solicitud.status === "aprobacion_financiera" && solicitud.tipo === "servicio" && !solicitud.pagosConfirmados) || (solicitud.status === "orden" && !todasOrdenesFirmadas(solicitud, proveedores)) || (solicitud.status === "recepcion" && (!solicitud.recepcion.recibidoSatisfaccion || !evaluacionCompleta(solicitud)))} className="px-4 py-2 rounded-lg text-sm bg-indigo-600 text-white font-medium disabled:opacity-40 flex items-center gap-1">{accionLabel(solicitud, total)} <ChevronRight size={15} /></button>
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

// lista de proveedores distintos adjudicados en una solicitud, con id (si es del catálogo) y nombre
function proveedoresAdjudicadosDetalle(s, proveedores) {
  const vistos = new Set();
  const lista = [];
  s.items.forEach((it) => {
    if (!it.cotizaciones.length) return;
    const idx = it.cotizacionSeleccionada ?? mejorCotizacionIdx(it.cotizaciones, it.cantidad);
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
    (fCreadoPor === "todos" || s.solicitanteId === fCreadoPor) &&
    (!fDesde || s.fechaCreacion >= fDesde) &&
    (!fHasta || s.fechaCreacion <= fHasta) &&
    (!texto ||
      s.folio.toLowerCase().includes(texto) ||
      (s.objetivo || "").toLowerCase().includes(texto) ||
      s.items.some((it) => it.nombre.toLowerCase().includes(texto)) ||
      proveedoresAdjudicados(s, proveedores).toLowerCase().includes(texto))
  );
  const hayFiltros = fArea !== "todas" || fEmpresa !== "todas" || fEstado !== "todos" || fCreadoPor !== "todos" || fDesde || fHasta || busqueda;
  const limpiarFiltros = () => { setFArea("todas"); setFEmpresa("todas"); setFEstado("todos"); setFCreadoPor("todos"); setFDesde(""); setFHasta(""); setBusqueda(""); };

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
  const esAdmin = currentUser?.rol === "Administrador";

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
          `Reenvío — Orden de compra/servicio ${s.folio}`,
          `<p>Hola ${prov.nombre},</p><p>Te reenviamos el enlace de la orden de compra/servicio <b>${s.folio}</b> a nombre de ${empresa?.nombre || ""}.</p><p><a href="${url}">Ver / descargar la orden firmada</a></p><p>Este enlace estará disponible por 7 días.</p>`
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
    <div className="bg-white rounded-xl border border-slate-200 overflow-x-auto">
      <table className="w-full text-sm">
        <thead className="bg-slate-50 text-slate-500"><tr>
          {esAdmin && <th className="px-4 py-2 w-8"><input type="checkbox" checked={todasSeleccionadas} onChange={alternarTodas} /></th>}
          <th className="text-left px-4 py-2 font-medium">Consecutivo</th><th className="text-left px-4 py-2 font-medium">Tipo</th><th className="text-left px-4 py-2 font-medium">Prioridad</th><th className="text-left px-4 py-2 font-medium">Área</th><th className="text-left px-4 py-2 font-medium">Empresa</th><th className="text-left px-4 py-2 font-medium">Fecha de registro</th><th className="text-left px-4 py-2 font-medium">Objetivo</th><th className="text-left px-4 py-2 font-medium">Proveedor adjudicado</th><th className="text-right px-4 py-2 font-medium">Total (IVA incl.)</th><th className="text-left px-4 py-2 font-medium">Estado</th><th></th><th></th></tr></thead>
        <tbody>{solicitudes.map((s) => { const area = areas.find((a) => a.id === s.areaId), empresa = empresas.find((e) => e.id === s.empresaId), paso = PASOS.find((p) => p.key === s.status);
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
  centrosCosto, guardarCentroCosto, eliminarCentroCosto,
  conceptosGasto, guardarConceptoGasto, eliminarConceptoGasto,
  permisos, togglePermiso,
}) {
  const [sub, setSub] = useState("empresas");
  const tabs = [
    { key: "empresas", label: "Empresas", icon: Building2 }, { key: "areas", label: "Áreas", icon: Layers }, { key: "departamentos", label: "Departamentos", icon: Layers }, { key: "proveedores", label: "Proveedores", icon: Truck },
    { key: "usuarios", label: "Usuarios y roles", icon: Users }, { key: "items", label: "Ítems", icon: Boxes },
    { key: "centros", label: "Centros de costo", icon: Layers }, { key: "conceptos", label: "Conceptos de gasto", icon: ClipboardList },
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
      {sub === "centros" && <CrudTable titulo="Centros de costo" icon={Layers} columnas={[{ key: "nombre", label: "Nombre" }]} datos={centrosCosto} onGuardar={guardarCentroCosto} onEliminar={eliminarCentroCosto} plantilla={{ nombre: "" }} />}
      {sub === "conceptos" && <CrudTable titulo="Conceptos de gasto" icon={ClipboardList} columnas={[{ key: "nombre", label: "Nombre" }]} datos={conceptosGasto} onGuardar={guardarConceptoGasto} onEliminar={eliminarConceptoGasto} plantilla={{ nombre: "" }} />}
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
  const { datos: centrosCosto, cargando: cargandoCentros, guardar: guardarCentroCosto, eliminar: eliminarCentroCosto, guardarVarios: importarCentros } = useSupabaseTable('centros_costo', {
    orderBy: 'nombre',
  });
  const { datos: conceptosGasto, cargando: cargandoConceptos, guardar: guardarConceptoGasto, eliminar: eliminarConceptoGasto, guardarVarios: importarConceptos } = useSupabaseTable('conceptos_gasto', {
    orderBy: 'nombre',
  });
  const cargandoCatalogos = cargandoAreas || cargandoDepartamentos || cargandoEmpresas || cargandoProveedores || cargandoUsuarios || cargandoItems || cargandoCentros || cargandoConceptos;

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

      <main className="flex-1 p-6 overflow-auto">
        {creando ? (
          <NuevaSolicitud areas={areas} departamentos={departamentos} empresas={empresas} itemsCatalogo={itemsCatalogo} guardarItemCatalogo={guardarItemCatalogo} proveedores={proveedores} guardarProveedor={guardarProveedor} centrosCosto={centrosCosto} conceptosGasto={conceptosGasto} usuarios={usuarios} currentUser={currentUser} onCrear={crearSolicitud} onCancel={() => setCreando(false)} />
        ) : perfil ? (
          <PerfilUsuario currentUser={currentUser} onGuardar={guardarPerfil} />
        ) : solicitudAbierta ? (
          <SolicitudDetalle solicitud={solicitudAbierta} areas={areas} departamentos={departamentos} empresas={empresas} usuarios={usuarios} proveedores={proveedores} guardarProveedor={guardarProveedor} itemsCatalogo={itemsCatalogo} centrosCosto={centrosCosto} conceptosGasto={conceptosGasto} historico={historico} setHistorico={setHistorico} currentUser={currentUser} onUpdate={actualizarSolicitud} onEliminar={eliminarSolicitud} onVolver={() => setAbierta(null)} />
        ) : tab === "dashboard" ? (
          <Dashboard areas={areas} solicitudes={solicitudesVisibles} />
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
            centrosCosto={centrosCosto} guardarCentroCosto={guardarCentroCosto} eliminarCentroCosto={eliminarCentroCosto}
            conceptosGasto={conceptosGasto} guardarConceptoGasto={guardarConceptoGasto} eliminarConceptoGasto={eliminarConceptoGasto}
            permisos={permisos} togglePermiso={togglePermiso}
          />
        ) : (
          <VistaSolicitudes solicitudes={solicitudesVisibles} areas={areas} empresas={empresas} usuarios={usuarios} proveedores={proveedores} currentUser={currentUser} onAbrir={setAbierta} onExportar={setExportando} onEliminarSeleccionadas={eliminarSolicitudesSeleccionadas} titulo={puedeVerTodasSolicitudes(currentUser) ? "Solicitudes" : "Mis solicitudes"} />
        )}
      </main>

      {exportando && !solicitudAbierta && (
        <div className="print-wrapper-oculto" style={{ display: "none" }}>
          <OrdenDocumento
            solicitud={exportando}
            empresa={empresas.find((e) => e.id === exportando.empresaId)}
            area={areas.find((a) => a.id === exportando.areaId)}
            departamento={departamentos.find((d) => d.id === exportando.departamentoId)}
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