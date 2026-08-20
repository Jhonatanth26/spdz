import React, { useState, useEffect } from "react";
import { Lock } from "lucide-react";
import { supabase } from "./lib/supabaseClient";

function LoginReal({ onIniciarSesion }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  const [empresas, setEmpresas] = useState([]);

  useEffect(() => {
    // los logos son públicos y la tabla "empresas" se puede leer sin sesión iniciada,
    // así que se pueden mostrar aquí mismo antes de ingresar
    supabase.from("empresas").select("nombre, logo_url").then(({ data, error }) => {
      if (!error && data) setEmpresas(data.filter((e) => e.logo_url));
    });
  }, []);

  const submit = async () => {
    setError(""); setCargando(true);
    const err = await onIniciarSesion(email.trim(), password);
    setCargando(false);
    if (err) setError("Correo o contraseña incorrectos.");
  };

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4" style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      <div className="bg-white border border-slate-200 rounded-2xl p-8 w-full max-w-sm">
        <div className="flex items-center gap-2 mb-6">
          <div className="w-10 h-10 rounded-lg bg-indigo-600 flex items-center justify-center text-white font-bold">GC</div>
          <div><div className="font-semibold text-slate-800">Gestión de Compras</div><div className="text-xs text-slate-400">Ingresa con tu correo</div></div>
        </div>
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
          <Lock size={14} /> {cargando ? "Ingresando..." : "Ingresar"}
        </button>

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