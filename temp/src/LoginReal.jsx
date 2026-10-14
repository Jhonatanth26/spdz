import React, { useState, useEffect } from "react";
import { Lock } from "lucide-react";
import { supabase } from "./lib/supabaseClient";
import { LOGO_FULL } from "./lib/logo";

const MODO_DEMO = import.meta.env.VITE_MODO_DEMO === 'true'; // solo se activa en el proyecto de demostración

function LoginReal({ onIniciarSesion }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [empresas, setEmpresas] = useState([]);
  const [registro, setRegistro] = useState(false); // solo en la versión de prueba: crear cuenta propia
  const [nombre, setNombre] = useState("");

  useEffect(() => {
    // los logos son públicos y la tabla "empresas" se puede leer sin sesión iniciada,
    // así que se pueden mostrar aquí mismo antes de ingresar
    supabase.from("empresas").select("nombre, logo_url").then(({ data, error }) => {
      if (!error && data) setEmpresas(data.filter((e) => e.logo_url));
    });
  }, []);

  const crearCuenta = async () => {
    setError("");
    if (!nombre.trim() || !email.trim() || password.length < 6) { setError("Escribe tu nombre, tu correo y una contraseña de mínimo 6 caracteres."); return; }
    setCargando(true);
    const { data, error: err } = await supabase.auth.signUp({ email: email.trim(), password, options: { data: { nombre: nombre.trim() } } });
    setCargando(false);
    if (err) { setError(err.message.includes("registered") ? "Ese correo ya tiene una cuenta de prueba: usa «Ingresar»." : "No se pudo crear la cuenta: " + err.message); return; }
    if (!data.session) setError("Cuenta creada. Revisa tu correo para confirmarla y luego ingresa.");
  };

  const submit = async () => {
    if (registro) return crearCuenta();
    setError(""); setCargando(true);
    const err = await onIniciarSesion(email.trim(), password);
    setCargando(false);
    if (err) setError("Correo o contraseña incorrectos.");
  };

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4" style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      <div className="bg-white border border-slate-200 rounded-2xl p-8 w-full max-w-sm">
        <div className="flex items-center gap-2 mb-6">
          <img src={LOGO_FULL} alt="SiCompras" className="h-9 w-auto" />
        </div>
        {MODO_DEMO && <div className="mb-4 text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">Versión de prueba con datos de ejemplo. No ingreses información real: otras personas también la están probando.</div>}
        {registro && (<>
          <label className="text-xs font-medium text-slate-500">Tu nombre</label>
          <input value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Nombre y apellido" className="w-full mt-1 mb-3 border border-slate-200 rounded-lg px-3 py-2 text-sm" />
        </>)}
        <label className="text-xs font-medium text-slate-500">Correo electrónico</label>
        <input
          type="email"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder="nombre@empresa.com"
          className="w-full mt-1 mb-3 border border-slate-200 rounded-lg px-3 py-2 text-sm"
          onKeyDown={(e) => e.key === "Enter" && submit()}
        />
        <label className="text-xs font-medium text-slate-500">Contraseña</label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          className="w-full mt-1 mb-1 border border-slate-200 rounded-lg px-3 py-2 text-sm"
          onKeyDown={(e) => e.key === "Enter" && submit()}
        />
        {error && <div className="text-xs text-rose-600 mb-2">{error}</div>}
        <button
          onClick={submit}
          disabled={cargando}
          className="w-full mt-3 bg-indigo-600 text-white rounded-lg py-2 text-sm font-medium flex items-center justify-center gap-2 disabled:opacity-50"
        >
          <Lock size={14} /> {cargando ? (registro ? "Creando cuenta..." : "Ingresando...") : (registro ? "Crear cuenta de prueba" : "Ingresar")}
        </button>

        {MODO_DEMO && (
          <button type="button" onClick={() => { setRegistro((v) => !v); setError(""); }} className="w-full mt-3 text-xs text-indigo-600 underline">
            {registro ? "Ya tengo cuenta de prueba: ingresar" : "¿Primera vez? Crear mi cuenta de prueba"}
          </button>
        )}

        {empresas.length > 0 && (
          <div className="mt-6 pt-5 border-t border-slate-100">
            <div className="text-[10px] text-slate-400 text-center mb-2 uppercase tracking-wide">Empresas del grupo</div>
            <div className="flex items-center justify-center gap-4 flex-wrap">
              {empresas.map((e, i) => (
                <img key={i} src={e.logo_url} alt={e.nombre} title={e.nombre} className="h-8 max-w-[90px] object-contain opacity-80" />
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default LoginReal;