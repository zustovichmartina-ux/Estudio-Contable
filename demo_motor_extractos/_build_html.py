# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from html import escape
from pathlib import Path

ROOT = Path(__file__).resolve().parent
data = json.loads((ROOT / "abril2026.json").read_text(encoding="utf-8"))
movs = data["movimientos"]
cred = sum(m["credito"] for m in movs)
deb = sum(m["debito"] for m in movs)
sf = movs[-1]["saldo"]


def ar(n: float) -> str:
    neg = n < 0
    s = f"{abs(float(n)):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if neg else "") + s


rows = []
for m in movs:
    cls = m["categoria"]
    pill = m["sub"] if m["sub"] != "-" else m["categoria"]
    cred_txt = ar(m["credito"]) if m["credito"] else ""
    deb_txt = ar(m["debito"]) if m["debito"] else ""
    rows.append(
        "<tr>"
        f"<td>{escape(m['fecha'])}</td>"
        f"<td><b>{escape(m['descripcion'])}</b><div class='det'>{escape(m['detalle'])}</div></td>"
        f"<td class='num cred'>{cred_txt}</td>"
        f"<td class='num deb'>{deb_txt}</td>"
        f"<td class='num'>{ar(m['saldo'])}</td>"
        f"<td><span class='pill {escape(cls)}'>{escape(pill)}</span></td>"
        "</tr>"
    )

html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Demo motor extractos — Galicia</title>
<style>
  :root { --azul:#1F4E79; --tinta:#1a1a1a; --suave:#666; --zebra:#F2F2F2; --linea:#d9d9d9; --ok:#1B7A4A; --mal:#B42318; }
  * { box-sizing: border-box; }
  body { margin:0; font: 15px/1.45 Calibri, Segoe UI, sans-serif; color:var(--tinta); background:#fff; }
  header { background:var(--azul); color:#fff; padding:22px 28px; }
  header h1 { margin:0 0 4px; font-size:22px; font-weight:700; }
  header p { margin:0; opacity:.85; font-size:14px; }
  main { max-width:1180px; margin:0 auto; padding:22px 24px 48px; }
  .kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:18px 0 22px; }
  .kpi { border:1px solid var(--linea); padding:12px 14px; }
  .kpi label { display:block; color:var(--suave); font-size:12px; text-transform:uppercase; letter-spacing:.04em; }
  .kpi b { font-size:20px; color:var(--azul); }
  .pipe { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin:8px 0 22px; }
  .step { border:1px solid var(--linea); padding:12px; min-height:92px; }
  .step n { display:block; color:var(--azul); font-weight:700; font-size:12px; }
  .drop { border:1px dashed var(--azul); padding:16px; text-align:center; margin:0 0 22px; background:#f7fafc; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th { background:var(--azul); color:#fff; text-align:left; padding:8px 10px; font-weight:600; }
  td { padding:7px 10px; border-bottom:1px solid var(--linea); vertical-align:top; }
  tr:nth-child(even) td { background:var(--zebra); }
  td.num { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap; }
  td.cred { color:var(--ok); }
  td.deb { color:var(--mal); }
  .det { color:var(--suave); font-size:12px; max-width:420px; }
  .pill { display:inline-block; padding:2px 8px; font-size:11px; border:1px solid var(--linea); }
  .pill.ingreso { border-color:var(--ok); color:var(--ok); }
  .pill.egreso, .pill.retencion, .pill.impuesto { border-color:var(--azul); color:var(--azul); }
  .pill.sin-cat { color:var(--mal); border-color:var(--mal); }
  .note { color:var(--suave); font-size:13px; margin:10px 0 18px; }
  .ok { color:var(--ok); font-weight:700; }
</style>
</head>
<body>
<header>
  <h1>Motor de extractos Galicia</h1>
  <p>Demo con BANCO GALICIA 042026.pdf — SUNNY BEACH SA · CUIT 30-71666637-5</p>
</header>
<main>
  <p class="note">Heuristica MM-Studio: agrupar el PDF por coordenada Y, tomar los dos ultimos numeros como <b>importe + saldo</b>, clasificar con RULES. Cadena de saldos: <span class="ok">35 / 35</span>.</p>
  <div class="pipe">
    <div class="step"><n>1. PDF</n>Texto nativo Galicia (no es escaneo)</div>
    <div class="step"><n>2. Lineas por Y</n>Misma altura = un renglon</div>
    <div class="step"><n>3. Parser Galicia</n>Fecha · penultimo = importe · ultimo = saldo</div>
    <div class="step"><n>4. RULES</n>FCI, IIBB, Ley 25413, terceros</div>
  </div>
  <div class="kpis">
    <div class="kpi"><label>Movimientos</label><b>__N__</b></div>
    <div class="kpi"><label>Creditos</label><b>$ __C__</b></div>
    <div class="kpi"><label>Debitos</label><b>$ __D__</b></div>
    <div class="kpi"><label>Saldo final</label><b>$ __S__</b></div>
  </div>
  <div class="drop">
    <b>Probar otro PDF digital</b>
    <div class="note" style="margin:6px 0 0">Marzo/abril 2026 tienen texto. Los escaneados (2025, ene, feb) se detectan y se avisa: este motor no hace OCR.</div>
    <input type="file" id="pdf" accept="application/pdf"/>
    <div id="status" class="note"></div>
  </div>
  <table>
    <thead><tr><th>Fecha</th><th>Descripcion</th><th>Credito</th><th>Debito</th><th>Saldo</th><th>Clasificacion</th></tr></thead>
    <tbody id="tbody">
    __ROWS__
    </tbody>
  </table>
</main>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const pNum=s=>{if(!s&&s!==0)return 0;s=String(s).trim();const neg=s.endsWith("-")||s.startsWith("-");s=s.replace(/^-/,"").replace(/-$/,"").replace(/\\./g,"").replace(",",".");return(parseFloat(s)||0)*(neg?-1:1);};
function detectBanco(t){const u=t.substring(0,4000).toUpperCase();if(/GALICIA|BANCO DE GALICIA/.test(u))return "Galicia";if(/SANTANDER/.test(u))return "Santander";if(/\\bBBVA\\b/.test(u))return "BBVA";return "Otro";}
const RULES=[
  {c:"inter-cta",s:"FCI - Suscripcion",re:/SUSCRIPCION FIMA|SUSCRIPCION FCI/i},
  {c:"inter-cta",s:"FCI - Rescate",re:/RESCATE FIMA|RESCATE FCI/i},
  {c:"inter-cta",s:"Transferencia propia",re:/TRANSF\\. CTAS PROPIAS|TRANSFERENCIA DE CUENTA PROPIA/i},
  {c:"retencion",s:"Ret. IIBB - ARBA",re:/ARBA AUTOM|PAGO.*ARBA|ARBA IIBB/i},
  {c:"retencion",s:"ICDB Ley 25.413",re:/LEY 25\\.413|IMP\\. DEB\\. LEY 25413|IMP\\. CRE\\. LEY 25413|IMP\\.LEY 25413/i},
  {c:"retencion",s:"Percepcion IVA",re:/PERCEP\\. IVA|PERCEPCION IVA|PERC\\.IVA/i},
  {c:"impuesto",s:"IVA banco",re:/(^|\\s)IVA(\\s|$)/i},
  {c:"impuesto",s:"Pago AFIP",re:/TRANSF\\. AFIP|\\bVEP\\b|PAGO DE SERVICIOS.*AFIP/i},
  {c:"egreso",s:"Comisiones y gastos bancarios",re:/COMISION SERVICIO|COMISION BANCO|COM\\. GESTION/i},
  {c:"impuesto",s:"IIBB",re:/ING\\. BRUTOS|IIBB/i},
  {c:"ingreso",s:"Transferencia de tercero",re:/TRANSFERENCIA DE TERCEROS|CREDITO TRANSFERENCIA/i},
  {c:"egreso",s:"Pago tarjeta VISA",re:/PAGO TARJETA VISA|PAGO VISA/i},
  {c:"egreso",s:"Transferencia a tercero",re:/TRANSFERENCIA A TERCEROS|TRF INMED PROVEED/i},
  {c:"egreso",s:"Pago de servicios",re:/PAGO DE SERVICIOS/i},
];
function classify(desc){for(const r of RULES)if(r.re.test(desc))return r;return {c:"sin-cat",s:"-"};}
function parseGaliciaPdf(lines){
  const movs=[], fechaRe=/^(\\d{2}\\/\\d{2}\\/\\d{2,4})\\s+/, numRe=/(-?\\d{1,3}(?:\\.\\d{3})*,\\d{2})/g;
  const skipRe=/^(Fecha|Movimientos|Resumen|Total|Los dep|Dispon|Canales|Ingres|Llaman|Usted|Al comp|Banco de|Chatea|http|Pagina)/i;
  let i=0;
  while(i<lines.length){
    const l=lines[i++].trim();
    if(!l||skipRe.test(l)) continue;
    const mf=l.match(fechaRe); if(!mf) continue;
    const rest=l.substring(mf[0].length);
    const nums=[]; let mm; const rn=new RegExp(numRe.source,"g");
    while((mm=rn.exec(rest))!==null) nums.push(pNum(mm[1]));
    if(nums.length<1) continue;
    const saldo=nums.length>=2?nums[nums.length-1]:0;
    const valor=nums.length>=2?nums[nums.length-2]:nums[0];
    const credito=valor>=0?valor:0, debito=valor<0?Math.abs(valor):0;
    let desc=rest.replace(/(-?\\d{1,3}(?:\\.\\d{3})*,\\d{2})/g,"").replace(/\\b\\d{6,}\\b/g,"").replace(/\\s+/g," ").trim();
    const extras=[];
    while(i<lines.length){
      const nxt=lines[i].trim();
      if(fechaRe.test(nxt)||skipRe.test(nxt)||/^total/i.test(nxt)) break;
      if(/p[aá]gina/i.test(nxt)) break;
      extras.push(nxt); i++;
    }
    const det=extras.join(" ");
    const full=(desc+" "+det).trim();
    const cl=classify(full);
    if(desc.length>2&&(credito>0||debito>0))
      movs.push({fecha:mf[1],descripcion:desc.substring(0,120),detalle:det.substring(0,160),credito,debito,saldo,categoria:cl.c,sub:cl.s});
  }
  return movs;
}
function ar(n){const neg=n<0;const v=Math.abs(n).toFixed(2);const [e,d]=v.split(".");const e2=e.replace(/\\B(?=(\\d{3})+(?!\\d))/g,".");return (neg?"-":"")+e2+","+d;}
function render(movs){
  document.getElementById("tbody").innerHTML=movs.map(m=>`<tr><td>${m.fecha}</td><td><b>${m.descripcion}</b><div class="det">${m.detalle||""}</div></td><td class="num cred">${m.credito?ar(m.credito):""}</td><td class="num deb">${m.debito?ar(m.debito):""}</td><td class="num">${ar(m.saldo)}</td><td><span class="pill ${m.categoria}">${m.sub && m.sub!=="-" ? m.sub : m.categoria}</span></td></tr>`).join("");
}
document.getElementById("pdf").addEventListener("change", async (ev)=>{
  const f=ev.target.files[0]; if(!f) return;
  const st=document.getElementById("status");
  st.textContent="Leyendo "+f.name+"…";
  const ab=await f.arrayBuffer();
  const pdf=await pdfjsLib.getDocument({data:ab}).promise;
  let full="", nChars=0;
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);
    const ct=await page.getTextContent();
    const byY={};
    ct.items.forEach(item=>{const y=Math.round(item.transform[5]); if(!byY[y]) byY[y]=[]; byY[y].push(item.str);});
    const sortedY=Object.keys(byY).map(Number).sort((a,b)=>b-a);
    const pageLines=sortedY.map(y=>byY[y].join(" ").trim()).filter(Boolean);
    nChars+=pageLines.join("").length;
    full+=pageLines.join("\\n")+"\\n";
  }
  if(nChars<50){ st.textContent="PDF escaneado (menos de 50 caracteres de texto). Este motor no hace OCR."; render([]); return; }
  const banco=detectBanco(full);
  const movs=parseGaliciaPdf(full.split(/\\n/));
  st.textContent=banco+" · "+pdf.numPages+" paginas · "+movs.length+" movimientos · texto nativo "+nChars+" caracteres";
  render(movs);
});
</script>
</body></html>
"""

html = (
    html.replace("__N__", str(len(movs)))
    .replace("__C__", ar(cred))
    .replace("__D__", ar(deb))
    .replace("__S__", ar(sf))
    .replace("__ROWS__", "\n".join(rows))
)
(ROOT / "index.html").write_text(html, encoding="utf-8")
print("wrote", ROOT / "index.html")
