"use client";
import { useCallback, useEffect, useMemo, useState } from "react";
import { Modo, borrarMes, cargar, guardarConfig, guardarMeses } from "@/lib/almacen";
import { incorporar } from "@/lib/lector";
import { calcular } from "@/lib/motor";
import { CONFIG_INICIAL, Config, Datos, Gasto, Parametros, norm } from "@/lib/tipos";
import { GASTOS_EJEMPLO, comprasEjemplo, ventasEjemplo } from "@/lib/ejemplo";
import { nombreMes } from "@/lib/formato";
import Hoy from "@/components/Hoy";
import Logo from "@/components/Logo";
import Cargar, { Lote, LoteCompras } from "@/components/Cargar";
import { incorporarCompras } from "@/lib/compras";
import Productos from "@/components/Productos";
import Gastos from "@/components/Gastos";
import Reportes from "@/components/Reportes";
import { IcoCargar, IcoGastos, IcoHoy, IcoProductos, IcoReportes } from "@/components/Iconos";

type Vista = "hoy" | "cargar" | "productos" | "gastos" | "reportes";
const VISTAS: { id: Vista; nombre: string; Icono: () => React.ReactElement }[] = [
  { id: "hoy", nombre: "Hoy", Icono: IcoHoy },
  { id: "cargar", nombre: "Cargar", Icono: IcoCargar },
  { id: "productos", nombre: "Productos", Icono: IcoProductos },
  { id: "gastos", nombre: "Gastos", Icono: IcoGastos },
  { id: "reportes", nombre: "Reportes", Icono: IcoReportes },
];

const mesActual = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
};

export default function App() {
  const [datos, setDatos] = useState<Datos>({ config: CONFIG_INICIAL, meses: {} });
  const [modo, setModo] = useState<Modo>("local");
  const [listo, setListo] = useState(false);
  const [mes, setMes] = useState(mesActual());
  const [vista, setVista] = useState<Vista>("hoy");
  const [error, setError] = useState("");
  const [anterior, setAnterior] = useState<Datos | null>(null);
  const [cartel, setCartel] = useState("");
  const [instalar, setInstalar] = useState<{ prompt: () => Promise<void> } | null>(null);

  // App instalable: copia para abrir sin conexión y botón "Instalar app" cuando el navegador lo permite
  useEffect(() => {
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js").catch(() => {});
    const alOfrecer = (e: Event) => {
      e.preventDefault();
      setInstalar(e as unknown as { prompt: () => Promise<void> });
    };
    const alInstalar = () => setInstalar(null);
    window.addEventListener("beforeinstallprompt", alOfrecer);
    window.addEventListener("appinstalled", alInstalar);
    return () => {
      window.removeEventListener("beforeinstallprompt", alOfrecer);
      window.removeEventListener("appinstalled", alInstalar);
    };
  }, []);

  useEffect(() => {
    cargar().then(({ datos, modo }) => {
      setDatos(datos);
      setModo(modo);
      const ultimos = Object.keys(datos.meses).sort();
      if (ultimos.length) setMes(ultimos[ultimos.length - 1]);
      setListo(true);
    });
    const leerHash = () => {
      const h = window.location.hash.slice(1) as Vista;
      if (VISTAS.some((v) => v.id === h)) setVista(h);
    };
    leerHash();
    window.addEventListener("hashchange", leerHash);
    return () => window.removeEventListener("hashchange", leerHash);
  }, []);

  const ir = useCallback((v: Vista) => {
    setVista(v);
    history.replaceState(null, "", `#${v}`);
    window.scrollTo({ top: 0 });
  }, []);

  const r = useMemo(() => calcular(datos.meses[mes], mes, datos.config), [datos, mes]);
  const opcionesMes = useMemo(() => [...new Set([mesActual(), mes, ...Object.keys(datos.meses)])].sort().reverse(), [datos.meses, mes]);

  const conError = async (fn: () => Promise<void>) => {
    try {
      setError("");
      await fn();
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const actualizarConfig = (cambio: (c: Config) => Config) => {
    const nuevo = { ...datos, config: cambio(datos.config) };
    setDatos(nuevo);
    conError(() => guardarConfig(nuevo, modo));
  };

  async function guardarVentas(lotes: Lote[], lotesCompras: LoteCompras[] = []) {
    setAnterior(datos);
    let meses = datos.meses;
    for (const l of lotes) meses = incorporar(meses, l.lineas, l.archivo, l.soloCompletar);
    for (const l of lotesCompras) meses = incorporarCompras(meses, l.compras, l.archivo);
    const lineas = lotes.flatMap((l) => l.lineas);
    const cambiados = [...new Set([...lineas.map((l) => l.fecha), ...lotesCompras.flatMap((l) => l.compras.map((c) => c.fecha))].map((f) => f.slice(0, 7)))].sort();
    const nombres = { ...datos.config.nombresArticulos };
    for (const l of lineas) nombres[norm(l.articulo)] ??= l.articulo;
    const nuevo = { config: { ...datos.config, nombresArticulos: nombres }, meses };
    setDatos(nuevo);
    setMes(cambiados[cambiados.length - 1]);
    await guardarMeses(nuevo, cambiados, modo);
    await guardarConfig(nuevo, modo);
  }

  async function deshacerCarga() {
    if (!anterior) return;
    const cambiados = Object.keys(datos.meses).filter((m) => datos.meses[m] !== anterior.meses[m]);
    setDatos(anterior);
    setAnterior(null);
    await conError(async () => {
      for (const m of cambiados) {
        if (anterior.meses[m]) await guardarMeses(anterior, [m], modo);
        else await borrarMes(anterior, m, modo);
      }
    });
  }

  function probarEjemplo() {
    const hoy = new Date();
    const m = mesActual();
    const lineas = ventasEjemplo(m, Math.max(1, hoy.getDate() - 1));
    const config: Config = { ...datos.config, gastosPorMes: { ...datos.config.gastosPorMes, [m]: GASTOS_EJEMPLO } };
    const nuevo = { config, meses: incorporarCompras(incorporar(datos.meses, lineas, "datos de ejemplo"), comprasEjemplo(lineas), "datos de ejemplo") };
    setDatos(nuevo);
    setMes(m);
    conError(async () => {
      await guardarMeses(nuevo, [m], modo);
      await guardarConfig(nuevo, modo);
    });
  }

  if (!listo)
    return <div className="vacio" style={{ minHeight: "100dvh", justifyContent: "center" }}><p className="muted">Cargando…</p></div>;

  return (
    <div className="app">
      <header className="top">
        <div className="top-in">
          <div className="marca">
            <Logo />
            <div>
              <h1>fiordo</h1>
              <span>Números del café · punto de equilibrio</span>
            </div>
          </div>
          {datos.meses[mes]?.archivos.includes("datos de ejemplo") && <span className="chip">Datos de ejemplo</span>}
          {instalar && (
            <button className="btn-instalar" onClick={async () => { await instalar.prompt(); setInstalar(null); }}>Instalar app</button>
          )}
          <label className="sr-only" htmlFor="mes" style={{ position: "absolute", left: -9999 }}>Mes</label>
          <select id="mes" className="mes-sel" value={mes} onChange={(e) => setMes(e.target.value)}>
            {opcionesMes.map((m) => <option key={m} value={m}>{nombreMes(m)}</option>)}
          </select>
        </div>
      </header>

      <nav className="nav" aria-label="Secciones">
        {VISTAS.map(({ id, nombre, Icono }) => (
          <button key={id} aria-current={vista === id ? "page" : undefined} onClick={() => ir(id)}>
            <Icono />
            {nombre}
          </button>
        ))}
      </nav>

      <main className="main">
        {error && <div className="aviso bad" role="alert">{error}</div>}
        {cartel && (vista === "hoy" || vista === "gastos") && (
          <div className="cartel" role="status">
            <span>✓ {cartel}</span>
            <div className="fila" style={{ gap: 8 }}>
              {anterior && <button className="btn sec" onClick={async () => { await deshacerCarga(); setCartel(""); }}>Deshacer</button>}
              <button className="btn sec" aria-label="Cerrar aviso" onClick={() => setCartel("")}>✕</button>
            </div>
          </div>
        )}
        {vista === "hoy" && <Hoy r={r} iibbPct={datos.config.parametros.iibbPct} irA={ir} probarEjemplo={probarEjemplo} />}
        {vista === "cargar" && <Cargar meses={datos.meses} ivaComputable={datos.config.parametros.ivaComputable} guardar={guardarVentas} deshacer={anterior ? deshacerCarga : null} verResultado={(m, v = "hoy") => { setCartel(m); ir(v); }} />}
        {vista === "productos" && (
          <Productos r={r} />
        )}
        {vista === "gastos" && (
          <Gastos
            r={r}
            config={datos.config}
            modo={modo}
            guardarGastos={(gastos: Gasto[], base: boolean) =>
              actualizarConfig((c) => ({ ...c, gastosPorMes: { ...c.gastosPorMes, [mes]: gastos }, gastosBase: base ? gastos : c.gastosBase }))
            }
            guardarParametros={(p: Parametros, nombreLocal: string) => actualizarConfig((c) => ({ ...c, parametros: p, nombreLocal }))}
            clasificarProveedor={(cuit, prov) => actualizarConfig((c) => ({ ...c, proveedores: { ...(c.proveedores ?? {}), [cuit]: prov } }))}
            irACargar={() => ir("cargar")}
            borrarMes={() => {
              const meses = { ...datos.meses };
              delete meses[mes];
              const nuevo = { ...datos, meses };
              setDatos(nuevo);
              conError(() => borrarMes(nuevo, mes, modo));
            }}
          />
        )}
        {vista === "reportes" && <Reportes r={r} nombreLocal={datos.config.nombreLocal} iibbPct={datos.config.parametros.iibbPct} />}
      </main>
    </div>
  );
}
