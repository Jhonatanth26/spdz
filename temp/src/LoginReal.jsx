import React, { useState } from "react";
import { Lock } from "lucide-react";

function LoginReal({ onIniciarSesion }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);

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
      </div>
    </div>
  );
}

export default LoginReal;