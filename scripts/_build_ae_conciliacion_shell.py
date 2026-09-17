# -*- coding: utf-8 -*-
"""Arma el HTML de Conciliación copiando AE_Studio (sin login ni secretos)."""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / "ae_conciliacion_component" / "index.html"
CANDIDATES = [
    Path(r"C:\Users\recep\Downloads\AE_Studio.html"),
    Path(r"C:\Users\recep\Desktop\mm-studio\app\AE_Studio.html"),
    Path(r"C:\Users\recep\Desktop\mm-studio\repo\app\AE_Studio.html"),
]

BRIDGE = r"""
<script>
(function(){
  function sendMessageToStreamlitClient(type, data){
    var out = Object.assign({isStreamlitMessage:true, type:type}, data);
    window.parent.postMessage(out, "*");
  }
  window._stSend = function(value){
    sendMessageToStreamlitClient("streamlit:setComponentValue", {value:value});
  };
  window._stHeight = function(h){
    sendMessageToStreamlitClient("streamlit:setFrameHeight", {height:h});
  };
  sendMessageToStreamlitClient("streamlit:componentReady", {apiVersion:1});
  window.addEventListener("message", function(event){
    if(!event.data) return;
    if(event.data.type === "streamlit:render"){
      bootFromPython(event.data.args || {});
    }
  });
})();

function _sbQ(){
  var q = {
    select: function(){return q;}, insert: function(){return q;}, update: function(){return q;},
    delete: function(){return q;}, upsert: function(){return q;}, eq: function(){return q;},
    in: function(){return q;}, order: function(){return q;},
    then: function(res){ return Promise.resolve({data:[], error:null}).then(res); }
  };
  return q;
}
sb = { from: function(){ return _sbQ(); } };

function fileToB64(file, cb){
  var r = new FileReader();
  r.onload = function(){ cb(String(r.result).split(",")[1] || ""); };
  r.readAsDataURL(file);
}
function sendFile(kind, file){
  if(!file){ notify("No hay archivo"); return; }
  if(!impCliOk()){ notify("Selecciona el cliente primero"); return; }
  loading("Enviando extracto...");
  fileToB64(file, function(b64){
    _stSend({action:"import", kind:kind, name:file.name, b64:b64, ts:Date.now()});
  });
}
function sendAction(payload){
  payload.ts = Date.now();
  _stSend(payload);
}

handlePdfFile = function(file){ sendFile("pdf", file); };
handleXlFile = function(file){ sendFile("excel", file); };
procesarPaste = function(){
  var t = (document.getElementById("paste-area").value||"").trim();
  if(!t){ notify("Pega texto primero"); return; }
  if(!impCliOk()){ notify("Selecciona el cliente primero"); return; }
  loading("Procesando texto...");
  sendAction({action:"import", kind:"text", text:t});
};
confirmarImport = function(){ loading("Guardando..."); sendAction({action:"save"}); };
cancelPrev = function(){ sendAction({action:"cancel_preview"}); };
loadMovs = function(){ refreshDash(); notify("Actualizado"); };
loadUsuarios = function(){
  var rows = window._usuarios || [];
  var tb = document.getElementById("tb-usr");
  if(!tb) return;
  tb.innerHTML = rows.length ? rows.map(function(u){
    var rol = u.es_admin ? "admin" : "contador";
    var tagRol = u.es_admin ? "retencion" : "ingreso";
    return "<tr><td style=\"font-weight:500\">"+(u.nombre||u.usuario||"")+"</td>"+
      "<td style=\"font-size:11px;color:var(--tx2)\">"+(u.usuario||"-")+"</td>"+
      "<td><span class=\"tag "+tagRol+"\">"+rol+"</span></td>"+
      "<td><span class=\"tag "+(u.activo?"ingreso":"egreso")+"\">"+(u.activo?"Activo":"Inactivo")+"</span></td>"+
      "<td style=\"font-size:10px;color:var(--tx3)\">-</td><td>-</td></tr>";
  }).join("") : "<tr><td colspan=\"6\" style=\"text-align:center;padding:20px;color:var(--tx3)\">Sin usuarios</td></tr>";
};
clearClientMovs = function(){
  if(!activeClient || !confirm("Eliminar TODOS los movimientos del cliente?")) return;
  sendAction({action:"clear"});
};
doLogout = function(){};
cambiarPin = function(){ notify("El PIN se cambia en el login de la web del estudio"); };
resetConfig = function(){ notify("La conexion la maneja la web del estudio"); };
saveCliente = function(){ notify("Los clientes se cargan en Clientes de la web"); };
pickClient = function(id){ sendAction({action:"pick_client", id:id}); };
onClientChange = function(){
  var id = document.getElementById("imp-cli").value;
  if(id) sendAction({action:"pick_client", id:id});
};
openClientPick = function(){
  renderCliPicker();
  openMod("mod-client-pick");
};
function renderCliPicker(){
  var box = document.getElementById("cli-picker");
  if(!box) return;
  box.innerHTML = (clientes||[]).map(function(c){
    return '<div class="clii" onclick="closeMod(\'mod-client-pick\');pickClient(\''+c.id+'\')"><span>'+c.nombre+'</span><span class="tbadge">'+(c.cuit||"")+'</span></div>';
  }).join("") || '<div style="font-size:11px;color:var(--tx3)">Sin clientes</div>';
}
renderSbCli = function(){
  var box = document.getElementById("sb-cli");
  if(!box) return;
  box.innerHTML = (clientes||[]).map(function(c){
    var on = activeClient && String(activeClient.id)===String(c.id) ? " active" : "";
    return '<div class="clii'+on+'" onclick="pickClient(\''+c.id+'\')">'+c.nombre+'</div>';
  }).join("");
};
populateImpSel = function(){
  var sel = document.getElementById("imp-cli");
  if(!sel) return;
  var cur = activeClient ? String(activeClient.id) : "";
  sel.innerHTML = '<option value="">Selecciona el cliente antes de subir el archivo</option>' +
    (clientes||[]).map(function(c){
      return '<option value="'+c.id+'"'+(String(c.id)===cur?" selected":"")+'>'+c.nombre+' — '+(c.cuit||"")+'</option>';
    }).join("");
  var badge = document.getElementById("cli-sel-badge");
  if(badge) badge.style.display = cur ? "block" : "none";
};
setBancoActivo = function(b){ sendAction({action:"set_banco", banco:b||""}); };
saveReclass = function(){
  if(!rcId) return;
  sendAction({
    action:"reclass",
    id: rcId,
    categoria: document.getElementById("rc-cat").value,
    sub_categoria: document.getElementById("rc-sub").value || "-",
    gravado: document.getElementById("rc-grav").value,
    observacion: document.getElementById("rc-obs").value,
    cuenta_codigo: document.getElementById("rc-cta").value || ""
  });
  closeMod("mod-reclass");
};

function applyPythonPreview(rows){
  prevRows = rows || [];
  if(!prevRows.length){
    var panel = document.getElementById("prev-panel");
    if(panel) panel.style.display = "none";
    loading(false);
    return;
  }
  showPrev();
  loading(false);
}

function bootFromPython(args){
  var auth = document.getElementById("auth-screen");
  if(auth) auth.classList.remove("show");
  var app = document.getElementById("app");
  if(app) app.classList.add("show");

  clientes = args.clientes || [];
  movs = args.movs || [];
  movsAll = movs.slice();
  activeClient = args.cliente || null;
  planCuentas = args.plan || [];
  proveedores = args.proveedores || [];
  empleados = args.empleados || [];
  window._usuarios = args.usuarios || [];
  user = {id:"web"};
  profile = {nombre: args.usuario || "Estudio", rol:"contador"};
  estUser = {id:"web", nombre: args.usuario || "Estudio"};

  var un = document.getElementById("user-nm");
  if(un) un.textContent = args.usuario || "Estudio";
  var ua = document.getElementById("user-avt");
  if(ua) ua.textContent = (args.usuario || "E").substring(0,1).toUpperCase();

  if(activeClient){
    var pill = document.getElementById("cpill-txt");
    if(pill) pill.textContent = activeClient.nombre;
    var dot = document.getElementById("cpill-dot");
    if(dot) dot.style.background = "var(--gr)";
  }
  renderBancoSel();
  if(args.banco){
    var bs = document.getElementById("bank-sel");
    if(bs) bs.value = args.banco;
    activeBank = args.banco;
  }
  renderSbCli();
  populateImpSel();
  updateBadges();
  var bprov = document.getElementById("b-prov"); if(bprov) bprov.textContent = proveedores.length;
  var bemp = document.getElementById("b-emp"); if(bemp) bemp.textContent = empleados.length;
  applyPythonPreview(args.preview || []);
  if(args.notify) notify(args.notify);
  if(args.view) showView(args.view);
  else renderActiveView();
  loading(false);
  setTimeout(function(){ _stHeight(Math.max(document.documentElement.scrollHeight, 920)); }, 80);
}

applyTheme();
if(window.__AE_ARGS && typeof bootFromPython==='function'){bootFromPython(window.__AE_ARGS);}
</script>
"""


def _strip_secrets(src: str) -> str:
    src = re.sub(r"const SB_URL='[^']*';", "const SB_URL='';", src)
    src = re.sub(r"const SB_KEY='[^']*';", "const SB_KEY='';", src)
    src = re.sub(r"const SB_TECH_EMAIL='[^']*';", "const SB_TECH_EMAIL='';", src)
    src = re.sub(r"const SB_TECH_PASS='[^']*';", "const SB_TECH_PASS='';", src)
    src = src.replace(
        '<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.js"></script>\n',
        "",
    )
    src = src.replace('Guardar en Supabase', 'Guardar movimientos')
    src = src.replace('Se guarda en Supabase para todo el equipo.', 'Los clientes se gestionan en la web del estudio.')
    src = src.replace(
        '<div class="card"><h3>Conexion Supabase</h3>',
        '<div class="card" style="display:none"><h3>Conexion</h3>',
    )
    src = src.replace(
        'applyTheme();\nsb=window.supabase.createClient(SB_URL,SB_KEY);\nbootLogin();',
        "/* boot: Streamlit */",
    )
    src = src.replace('<div class="ls show" id="auth-screen">', '<div class="ls" id="auth-screen">')
    src = src.replace('<div class="app" id="app">', '<div class="app show" id="app">')
    src = src.replace(
        '<span onclick="cambiarPin()" title="Cambiar PIN"',
        '<span style="display:none" onclick="cambiarPin()" title="Cambiar PIN"',
    )
    src = src.replace(
        '<button class="btn sm danger" onclick="doLogout()">Salir</button>',
        "",
    )
    return src


def main() -> None:
    src_path = next((p for p in CANDIDATES if p.is_file()), None)
    if not src_path:
        raise SystemExit("No encuentro AE_Studio.html")
    raw = src_path.read_text(encoding="utf-8")
    raw = _strip_secrets(raw)
    if "eyJ" in raw or "MMstudio" in raw or "cumbgsqbzawhiycimkvt" in raw:
        raise SystemExit("El HTML generado todavía tiene secretos; aborto.")
    raw = raw.replace("</body>", BRIDGE + "\n</body>")
    DEST.parent.mkdir(parents=True, exist_ok=True)
    DEST.write_text(raw, encoding="utf-8")
    print(f"OK {DEST} ({DEST.stat().st_size} bytes) from {src_path}")


if __name__ == "__main__":
    main()
