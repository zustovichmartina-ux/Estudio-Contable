# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from pathlib import Path

from imputacion_bancaria import imputar_extracto

ROOT = Path(__file__).resolve().parent
movs = json.loads((ROOT / "abril2026.json").read_text(encoding="utf-8"))["movimientos"]
padron = json.loads((ROOT / "padron_sunny.json").read_text(encoding="utf-8"))
out = imputar_extracto(movs, padron)
out["archivo"] = "BANCO GALICIA 042026.pdf"
out["cliente"] = padron["cliente"]
(ROOT / "imputados_abril.json").write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
data_js = json.dumps(out, ensure_ascii=False)
padron_js = json.dumps(padron, ensure_ascii=False)

html = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>AE-Studio · Imputación bancaria</title>
<style>
  :root { --bg:#0b1220; --side:#0e1628; --card:#141c2e; --line:#243049; --txt:#e8eef8; --dim:#8b9bb4; --ok:#3dd68c; --mal:#ff6b6b; --warn:#f5a524; --azul:#3b82f6; --lock:#64748b; }
  * { box-sizing:border-box; }
  body { margin:0; font:14px/1.4 Segoe UI, sans-serif; background:var(--bg); color:var(--txt); display:flex; min-height:100vh; }
  aside { width:220px; background:var(--side); border-right:1px solid var(--line); padding:18px 14px; }
  aside h1 { font-size:15px; margin:0 0 18px; letter-spacing:.04em; }
  aside a { display:block; color:var(--dim); text-decoration:none; padding:8px 10px; border-radius:6px; }
  aside a.on, aside a:hover { background:var(--card); color:var(--txt); }
  aside .sec { margin:16px 0 6px; font-size:11px; color:var(--dim); text-transform:uppercase; letter-spacing:.08em; }
  main { flex:1; padding:20px 24px 48px; overflow:auto; }
  .top { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; }
  .kpis { display:grid; grid-template-columns:repeat(6,1fr); gap:10px; margin-bottom:16px; }
  .kpi { background:var(--card); border:1px solid var(--line); padding:12px; }
  .kpi b { display:block; font-size:18px; }
  .kpi span { color:var(--dim); font-size:11px; text-transform:uppercase; }
  .kpi.ok b { color:var(--ok); } .kpi.mal b { color:var(--mal); } .kpi.warn b { color:var(--warn); } .kpi.azul b { color:var(--azul); }
  table { width:100%; border-collapse:collapse; font-size:13px; background:var(--card); }
  th { text-align:left; color:var(--dim); font-weight:600; padding:8px 10px; border-bottom:1px solid var(--line); }
  td { padding:8px 10px; border-bottom:1px solid var(--line); vertical-align:top; }
  td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
  .det { color:var(--dim); font-size:12px; }
  .pill { display:inline-block; padding:2px 8px; border-radius:999px; font-size:11px; }
  .pill.fija { background:#1f2937; color:#cbd5e1; }
  .pill.sugerida { background:#16325c; color:#93c5fd; }
  .pill.pendiente { background:#3a2710; color:var(--warn); }
  select { background:var(--bg); color:var(--txt); border:1px solid var(--line); padding:4px 6px; max-width:280px; }
  select:disabled { opacity:.7; }
  button { background:var(--azul); color:#fff; border:0; padding:8px 14px; border-radius:6px; cursor:pointer; }
  button.ghost { background:transparent; border:1px solid var(--line); color:var(--txt); }
  button:disabled { opacity:.4; cursor:not-allowed; }
  .note { color:var(--dim); font-size:12px; margin-top:8px; }
  .foot { display:flex; gap:8px; margin-top:16px; align-items:center; }
</style>
</head>
<body>
<aside>
  <h1>AE · Estudio</h1>
  <div class="sec">Analisis</div>
  <a href="imputacion.html" class="on">Movimientos</a>
  <a href="#deudores">Deudores</a>
  <a href="#proveedores">Proveedores</a>
  <div class="sec">Salida</div>
  <a href="#tango">Asiento Tango</a>
  <a href="#papeles">Papeles de trabajo</a>
  <div class="sec">Sistema</div>
  <a href="index.html">Buzon OCR</a>
</aside>
<main>
  <div class="top">
    <div>
      <div style="font-size:20px;font-weight:700">SUNNY BEACH SA · Banco Galicia</div>
      <div class="note">Misma grilla del extracto, con imputacion. Generales = cuenta fija. Transferencias = deudor/proveedor.</div>
    </div>
    <button type="button" onclick="location.href='index.html'">Importar extracto</button>
  </div>
  <div class="kpis">
    <div class="kpi"><span>Movimientos</span><b id="k-n">0</b></div>
    <div class="kpi ok"><span>Fijas</span><b id="k-f">0</b></div>
    <div class="kpi azul"><span>Sugeridas padron</span><b id="k-s">0</b></div>
    <div class="kpi warn"><span>A imputar</span><b id="k-p">0</b></div>
    <div class="kpi ok"><span>Creditos</span><b id="k-c">0</b></div>
    <div class="kpi mal"><span>Debitos</span><b id="k-d">0</b></div>
  </div>
  <table>
    <thead>
      <tr>
        <th>Fecha</th><th>Descripcion</th><th>Credito</th><th>Debito</th><th>Saldo</th>
        <th>Imputacion</th><th>Contraparte</th><th>Origen</th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
  <div class="foot">
    <button type="button" disabled title="Plantilla vacia pendiente">Generar asiento Tango</button>
    <button type="button" class="ghost" disabled title="Plantilla vacia pendiente">Armar papeles de trabajo</button>
    <span class="note">Esas dos salidas usan las plantillas vacias que me vas a pasar. Hasta entonces no invento el formato.</span>
  </div>
  <h3 id="deudores" style="margin-top:28px">Deudores (padron)</h3>
  <div id="lst-deu" class="note"></div>
  <h3 id="proveedores">Proveedores (padron)</h3>
  <div id="lst-prv" class="note"></div>
</main>
<script>
const DATA = __DATA__;
const PADRON = __PADRON__;
function ar(n){ if(!n) return ""; const neg=n<0; const [e,d]=Math.abs(n).toFixed(2).split("."); return (neg?"-":"")+e.replace(/\B(?=(\d{3})+(?!\d))/g,".")+","+d; }
function opciones(m){
  const opts=["Gastos bancarios","Deudores por ventas","Proveedores","Cheques a depositar","Valores diferidos","A imputar"];
  (PADRON.deudores||[]).forEach(p=>opts.push("Deudores por ventas · "+p.nombre));
  (PADRON.proveedores||[]).forEach(p=>opts.push("Proveedores · "+p.nombre));
  const cur=m.imputacion||"A imputar";
  if(!opts.includes(cur)) opts.unshift(cur);
  return opts.map(o=>`<option ${o===cur?"selected":""}>${o}</option>`).join("");
}
function pintar(){
  const movs=DATA.movimientos||[];
  const r=DATA.resumen||{};
  document.getElementById("k-n").textContent=r.total||movs.length;
  document.getElementById("k-f").textContent=r.fijas||0;
  document.getElementById("k-s").textContent=r.sugeridas||0;
  document.getElementById("k-p").textContent=r.pendientes||0;
  document.getElementById("k-c").textContent="$ "+ar(movs.reduce((s,m)=>s+(m.credito||0),0));
  document.getElementById("k-d").textContent="$ "+ar(movs.reduce((s,m)=>s+(m.debito||0),0));
  document.getElementById("tbody").innerHTML=movs.map((m,i)=>{
    const lock=m.origen_imputacion==="fija";
    const imp = lock
      ? `<span title="Cuenta fija ${m.cuenta||""}">${m.imputacion}</span>`
      : `<select data-i="${i}">${opciones(m)}</select>`;
    return `<tr>
      <td>${m.fecha}</td>
      <td><b>${m.descripcion||""}</b><div class="det">${m.detalle||""}</div></td>
      <td class="num" style="color:var(--ok)">${m.credito?ar(m.credito):""}</td>
      <td class="num" style="color:var(--mal)">${m.debito?ar(m.debito):""}</td>
      <td class="num">${ar(m.saldo)}</td>
      <td>${imp}</td>
      <td>${m.contraparte||""}</td>
      <td><span class="pill ${m.origen_imputacion}">${m.origen_imputacion}</span></td>
    </tr>`;
  }).join("");
  document.getElementById("lst-deu").textContent=(PADRON.deudores||[]).map(p=>p.nombre+" · "+p.cuit+" · "+p.cuenta).join("  |  ")||"Sin padron";
  document.getElementById("lst-prv").textContent=(PADRON.proveedores||[]).map(p=>p.nombre+" · "+p.cuit+" · "+p.cuenta).join("  |  ")||"Sin padron";
}
pintar();
</script>
</body></html>
"""
html = html.replace("__DATA__", data_js).replace("__PADRON__", padron_js)
(ROOT / "imputacion.html").write_text(html, encoding="utf-8")
print("imputacion.html", (ROOT / "imputacion.html").stat().st_size)
print(out["resumen"])
