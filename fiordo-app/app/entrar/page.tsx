"use client";
import { useState } from "react";

const campo: React.CSSProperties = { font: "inherit", fontSize: 16, padding: "10px 12px", border: "1px solid var(--line)", borderRadius: 10, background: "var(--surface)", color: "inherit" };

export default function Entrar() {
  const [usuario, setUsuario] = useState("");
  const [clave, setClave] = useState("");
  const [error, setError] = useState("");
  const [enviando, setEnviando] = useState(false);

  async function entrar(e: React.FormEvent) {
    e.preventDefault();
    setEnviando(true);
    setError("");
    try {
      const r = await fetch("/api/entrar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ usuario, clave }),
      });
      if (r.ok) window.location.replace("/");
      else setError("Usuario o contraseña incorrectos.");
    } catch {
      setError("No hay conexión. Probá de nuevo.");
    } finally {
      setEnviando(false);
    }
  }

  return (
    <main style={{ minHeight: "100dvh", display: "grid", placeItems: "center", padding: 16 }}>
      <form onSubmit={entrar} className="card" style={{ width: "100%", maxWidth: 340, display: "grid", gap: 12 }}>
        <div style={{ textAlign: "center" }}>
          <img src="/icon-192.png" alt="" width={56} height={56} style={{ borderRadius: 12 }} />
          <h1 style={{ margin: "8px 0 0", fontSize: "1.25rem" }}>Fiordo</h1>
          <p style={{ margin: 0, opacity: 0.7, fontSize: ".9rem" }}>Números del café</p>
        </div>
        <label style={{ display: "grid", gap: 4 }}>
          Usuario
          <input value={usuario} onChange={(e) => setUsuario(e.target.value)} autoCapitalize="none" autoCorrect="off" autoComplete="username" required style={campo} />
        </label>
        <label style={{ display: "grid", gap: 4 }}>
          Contraseña
          <input type="password" value={clave} onChange={(e) => setClave(e.target.value)} autoComplete="current-password" required style={campo} />
        </label>
        {error && <p style={{ color: "#b3261e", margin: 0, fontSize: ".9rem" }}>{error}</p>}
        <button className="btn" type="submit" disabled={enviando}>{enviando ? "Entrando…" : "Entrar"}</button>
      </form>
    </main>
  );
}
