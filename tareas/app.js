/* Tareas Gerencia Vitis Olei — app instalable que funciona sin conexión.
   Todo se guarda primero en el teléfono (bandeja de salida) y se envía al Excel
   en OneDrive, a través de un flujo de Power Automate, cuando hay red. */
(() => {
'use strict';
const APP_VERSION = "1.0.0";
const DEFAULT_CATS = {
  departamentos:["Administración","CEO","Cosecha","Fertilización","Junta","Mantenimiento","Proveedores","Proyectos","Sanidad","Taller Agrícola","Vivero"],
  estados:["1.Sin avance","2.En proceso","3.Socializar","4.Cumplido","5.Stand-by"],
  avances:[0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1],
  criticidades:["Alta","Media","Baja"],
  responsables:["Jairo"]
};
const CAT_LABELS = {departamentos:"Departamentos",estados:"Estados",criticidades:"Criticidad",responsables:"Líder responsable"};
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const norm = s => String(s||"").normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase().trim();

/* ---------- Almacenamiento local ---------- */
const lsGet = (k, d) => { try { const v = JSON.parse(localStorage.getItem(k)||"null"); return v ?? d; } catch(e){ return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch(e){ return false; } };
const K = {remote:"tg_remote", out:"tg_outbox", cats:"tg_cats", cfg:"tg_cfg", ui:"tg_ui", sync:"tg_last_sync", failed:"tg_failed"};
try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch(e){}

let cfg = lsGet(K.cfg, {flowUrl:"", excelUrl:""});
let remote = lsGet(K.remote, []);
let outbox = lsGet(K.out, []);
let failed = lsGet(K.failed, []);
let cats = {...structuredClone(DEFAULT_CATS), ...lsGet(K.cats, {})};
let lastSync = lsGet(K.sync, null);
let tasks = [];
let flushing = false, pulling = false, lastError = "";
const ui = {dep:"", estado:"", crit:"", resp:"", texto:"", kpi:"abiertas", sort:"fechaPendiente", dir:"asc", ...lsGet(K.ui, {})};
const saveUi = () => lsSet(K.ui, ui);

/* ---------- Fechas ---------- */
const pad = n => String(n).padStart(2,"0");
const iso = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const today = () => { const d=new Date(); d.setHours(0,0,0,0); return d; };
const parseIso = s => { if(!s) return null; const [y,m,d]=String(s).slice(0,10).split("-").map(Number); return y ? new Date(y,m-1,d) : null; };
const fmt = s => { const d=parseIso(s); return d ? `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${String(d.getFullYear()).slice(2)}` : ""; };
const daysSince = s => { const d=parseIso(s); return d ? Math.round((today()-d)/864e5) : null; };
const isDone = t => /cumplido/i.test(t.estado||"");
const isLate = t => !isDone(t) && t.fechaEntrega && parseIso(t.fechaEntrega) < today();
const estadoClass = e => { const m=/^(\d)/.exec(e||""); return "e"+(m?m[1]:"0"); };
/* Excel puede devolver fechas como ISO o como número de serie */
function excelDate(v){
  if (v==null || v==="") return "";
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  if (/^\d+(\.\d+)?$/.test(s)){ const d = new Date(Date.UTC(1899,11,30) + Math.floor(Number(s))*864e5); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth()+1)}-${pad(d.getUTCDate())}`; }
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
  return "";
}
function excelNum(v){ if (v==null || v==="") return null; const n = Number(String(v).replace("%","").replace(",",".")); if (isNaN(n)) return null; return n>1 ? n/100 : n; }

let toastT; function toast(msg){ const el=$("#toast"); el.textContent=msg; el.hidden=false; clearTimeout(toastT); toastT=setTimeout(()=>el.hidden=true,3500); }
function showBanner(msg){ const b=$("#banner"); if(!msg){ b.hidden=true; return; } b.hidden=false; b.textContent=msg; }

/* ---------- Modelo: lo del Excel + cambios pendientes ---------- */
function newId(){ return "t" + Date.now().toString(36) + Math.random().toString(36).slice(2,6); }
function nextNo(){ return tasks.reduce((m,t)=>Math.max(m, Number(t.no)||0), 0) + 1; }
function compose(){
  const map = new Map(remote.map(t=>[t.id,t]));
  for (const op of outbox){
    if (op.kind==="task") map.set(op.id, {id:op.id, ...op.body, pendiente:true});
    else if (op.kind==="del") map.delete(op.id);
  }
  tasks = [...map.values()];
}
function enqueue(op){
  op.ts = Date.now(); op.tries = 0;
  outbox = outbox.filter(o => o.id!==op.id);
  outbox.push(op);
  if (!lsSet(K.out, outbox)) toast("El teléfono no tiene espacio para guardar sin conexión.");
  compose(); render(); flush();
}

/* ---------- Comunicación con el flujo ---------- */
const withTimeout = (p, ms) => Promise.race([p, new Promise((_,rej)=>setTimeout(()=>rej({code:"timeout", net:true}), ms))]);
async function callFlow(payload){
  if (!cfg.flowUrl) throw {code:"no_config"};
  let r;
  try {
    // text/plain evita la verificación previa (CORS) que Power Automate no responde
    r = await withTimeout(fetch(cfg.flowUrl, {method:"POST", headers:{"Content-Type":"text/plain;charset=UTF-8"}, body:JSON.stringify(payload), cache:"no-store"}), 45000);
  } catch(e){ throw {code: e.code||"network", net:true}; }
  const txt = await r.text();
  if (!r.ok) throw {code:"http_"+r.status, status:r.status, detail:txt.slice(0,300)};
  try { return txt ? JSON.parse(txt) : {}; } catch(e){ return {}; }
}
function toPayload(t){
  return {ID:t.id, No:Number(t.no)||null, Departamento:t.departamento||"", Tema:t.tema||"",
    FechaPendiente:t.fechaPendiente||"", FechaEntrega:t.fechaEntrega||"", FechaRevision:t.fechaRevision||"",
    Estado:t.estado||"", Avance:t.avance==null?"":t.avance, Criticidad:t.criticidad||"", Responsable:t.responsable||"",
    Observaciones:t.observaciones||"", Origen:t.origen==="voz"?"Voz":t.origen==="excel"?"Excel":"Manual",
    Actualizado:(t.actualizado||new Date().toISOString()).slice(0,16).replace("T"," ")};
}
const ORIGEN = {voz:"voz", excel:"excel", manual:"manual"};
function fromRow(r, i){
  const g = k => r[k] ?? "";
  const id = String(g("ID")).trim();
  return {
    id: id || ("sinid-" + (g("No")||i) + "-" + norm(g("TEMA")).slice(0,20)),
    sinId: !id,
    no: Number(g("No")) || null,
    departamento: String(g("DEPARTAMENTO")).trim(),
    tema: String(g("TEMA")).trim(),
    fechaPendiente: excelDate(g("FECHA PENDIENTE")),
    fechaEntrega: excelDate(g("FECHA DE ENTREGA")),
    fechaRevision: excelDate(g("FECHA DE REVISIÓN") || g("FECHA DE REVISON")),
    estado: String(g("ESTADO")).trim(),
    avance: excelNum(g("AVANCE")),
    criticidad: String(g("CRITICIDAD")).trim(),
    responsable: String(g("RESPONSABLE")).trim(),
    observaciones: String(g("OBSERVACIONES")).trim(),
    origen: ORIGEN[norm(g("ORIGEN"))] || "excel",
    actualizado: String(g("ACTUALIZADO"))
  };
}
const firstVal = row => { for (const [k,v] of Object.entries(row||{})) if (!k.startsWith("@") && k!=="ItemInternalId" && v!=="" && v!=null) return v; return null; };
const colVals = rows => Array.isArray(rows) ? rows.map(firstVal).filter(v=>v!=null).map(v=>String(v).trim()) : null;

async function flush(){
  if (flushing || !outbox.length || !cfg.flowUrl || navigator.onLine===false){ updateSync(); return; }
  flushing = true; updateSync();
  try {
    while (outbox.length){
      const op = outbox[0];
      try {
        if (op.kind==="task") await callFlow({accion:"guardar", tarea: toPayload({id:op.id, ...op.body})});
        else await callFlow({accion:"eliminar", id: op.id});
      } catch(e){
        if (e.net || e.code==="no_config" || e.status===429 || e.status===408) throw e;   // sin red: se reintenta luego
        op.tries = (op.tries||0)+1; lsSet(K.out, outbox);
        lastError = e.detail || e.code;
        if (op.tries >= 4){ failed.push({...op, error:lastError}); lsSet(K.failed, failed); outbox = outbox.filter(o=>o!==op); lsSet(K.out, outbox); continue; }
        throw e;
      }
      // confirmado: pasa a la copia local del Excel
      if (op.kind==="task"){ const t={id:op.id, ...op.body}; const i=remote.findIndex(x=>x.id===op.id); if(i>=0) remote[i]=t; else remote.push(t); }
      else remote = remote.filter(x=>x.id!==op.id);
      lsSet(K.remote, remote);
      if (outbox[0]===op) outbox.shift(); else outbox = outbox.filter(o=>o!==op);
      lsSet(K.out, outbox);
      lastSync = new Date().toISOString(); lsSet(K.sync, lastSync); lastError = "";
      compose(); render();
    }
  } catch(e) { /* queda en cola */ }
  finally { flushing = false; updateSync(); }
}
async function pull(manual){
  if (pulling || !cfg.flowUrl || navigator.onLine===false){ updateSync(); return false; }
  pulling = true; updateSync();
  try {
    const r = await callFlow({accion:"listar"});
    if (!Array.isArray(r.tareas)) throw {code:"respuesta", detail:"El flujo no devolvió la lista 'tareas'."};
    remote = r.tareas.map(fromRow).filter(t=>t.tema);
    lsSet(K.remote, remote);
    const nc = {
      departamentos: colVals(r.departamentos), estados: colVals(r.estados),
      criticidades: colVals(r.criticidades), responsables: colVals(r.responsables)
    };
    for (const [k,v] of Object.entries(nc)) if (v && v.length) cats[k] = v;
    lsSet(K.cats, cats);
    lastSync = new Date().toISOString(); lsSet(K.sync, lastSync); lastError = "";
    compose(); renderAll();
    if (manual) toast(`Listado actualizado desde el Excel (${remote.length} tareas)`);
    return true;
  } catch(e){
    lastError = e.detail || e.code;
    if (manual) toast(e.net ? "Sin conexión: se actualizará cuando haya señal." : "El Excel no respondió bien: " + lastError);
    return false;
  } finally { pulling = false; updateSync(); }
}
async function syncAll(manual){ await flush(); if (!outbox.length) await pull(manual); }

function updateSync(){
  const el = $("#sync"); if (!el) return;
  const n = outbox.length, off = navigator.onLine===false;
  const hora = lastSync ? new Date(lastSync).toLocaleTimeString("es-CO",{hour:"2-digit",minute:"2-digit"}) : "";
  let cls, txt;
  if (!cfg.flowUrl){ cls = n?"pend":"off"; txt = n ? `${n} cambio${n>1?"s":""} guardado${n>1?"s":""} en el teléfono · falta conectar el Excel` : "Sin conectar al Excel"; }
  else if (off){ cls="off"; txt = n ? `Sin señal · ${n} cambio${n>1?"s":""} guardado${n>1?"s":""} en el teléfono` : `Sin señal · copia del ${lastSync?new Date(lastSync).toLocaleDateString("es-CO",{day:"2-digit",month:"short"})+" "+hora:"teléfono"}`; }
  else if (flushing){ cls="pend"; txt = `Enviando ${n} cambio${n>1?"s":""} al Excel…`; }
  else if (pulling){ cls="pend"; txt = "Actualizando desde el Excel…"; }
  else if (n){ cls="pend"; txt = `${n} cambio${n>1?"s":""} por enviar · reintentando`; }
  else { cls="ok"; txt = "Sincronizado con Excel" + (hora ? " · " + hora : ""); }
  el.className = "sync " + cls; el.textContent = txt;
  if (failed.length) showBanner(`${failed.length} cambio${failed.length>1?"s":""} no se pudo${failed.length>1?"ieron":""} escribir en el Excel (${failed[failed.length-1].error}). Revisa el flujo y pulsa Sincronizar para reintentar.`);
  else if (!cfg.flowUrl) showBanner("Las tareas se guardan en este teléfono. Toca “Conexión” para enlazar tu Excel de OneDrive y sincronizar.");
  else if (lastError && n && !off) showBanner("No se pudo enviar al Excel: " + lastError + ". Se reintentará automáticamente.");
  else showBanner("");
}

async function saveTask(t){ t.actualizado = new Date().toISOString(); const {id, pendiente, sinId, ...body} = t; enqueue({kind:"task", id, body}); }
async function deleteTask(id){ enqueue({kind:"del", id}); }

/* ---------- Formularios ---------- */
function opts(list, val, blank){ return (blank?`<option value="">${blank}</option>`:"") + list.map(v=>`<option value="${esc(v)}"${String(v)===String(val??"")?" selected":""}>${esc(v)}</option>`).join(""); }
function avOpts(val){ return `<option value="">—</option>` + cats.avances.map(v=>`<option value="${v}"${Number(val)===Number(v)?" selected":""}>${Math.round(v*100)}%</option>`).join(""); }
function withCurrent(list, v){ return v && !list.includes(v) ? [...list, v] : list; }
function formHtml(t, pfx){
  return `
  <label class="f full">Tema<textarea class="in" id="${pfx}tema" rows="2" required>${esc(t.tema)}</textarea></label>
  <label class="f">Departamento<select class="in" id="${pfx}departamento">${opts(withCurrent(cats.departamentos,t.departamento), t.departamento, "Sin departamento")}</select></label>
  <label class="f">Responsable<select class="in" id="${pfx}responsable">${opts(withCurrent(cats.responsables,t.responsable), t.responsable, "Sin asignar")}</select></label>
  <label class="f">Fecha pendiente<input class="in" type="date" id="${pfx}fechaPendiente" value="${esc(t.fechaPendiente)}"></label>
  <label class="f">Fecha de entrega<input class="in" type="date" id="${pfx}fechaEntrega" value="${esc(t.fechaEntrega)}"></label>
  <label class="f">Estado<select class="in" id="${pfx}estado">${opts(withCurrent(cats.estados,t.estado), t.estado, "Sin estado")}</select></label>
  <label class="f">Avance<select class="in" id="${pfx}avance">${avOpts(t.avance)}</select></label>
  <label class="f">Criticidad<select class="in" id="${pfx}criticidad">${opts(withCurrent(cats.criticidades,t.criticidad), t.criticidad, "—")}</select></label>
  <label class="f">Fecha de revisión<input class="in" type="date" id="${pfx}fechaRevision" value="${esc(t.fechaRevision)}"></label>
  <label class="f full">Observaciones<textarea class="in" id="${pfx}observaciones" rows="2">${esc(t.observaciones)}</textarea></label>`;
}
const FIELDS = ["tema","departamento","responsable","fechaPendiente","fechaEntrega","estado","avance","criticidad","fechaRevision","observaciones"];
function readForm(pfx, base){
  const t = {...base};
  for (const f of FIELDS){ const el=document.getElementById(pfx+f); if(!el) continue; t[f] = f==="avance" ? (el.value===""?null:Number(el.value)) : el.value.trim(); }
  if (/cumplido/i.test(t.estado) && t.avance==null) t.avance = 1;
  return t;
}
function blankTask(){ return {id:newId(), no:nextNo(), departamento:"", tema:"", fechaPendiente:iso(today()), fechaEntrega:"", fechaRevision:"", estado:"1.Sin avance", avance:null, criticidad:"", responsable:"", observaciones:"", origen:"manual", creado:new Date().toISOString()}; }
function renderNewForm(){
  const f=$("#newForm"); const keep = f.dataset.ready ? readForm("n_", blankTask()) : blankTask();
  f.innerHTML = formHtml(keep,"n_") + `<div class="full row"><button class="btn primary" type="submit">Agregar tarea</button></div>`;
  f.dataset.ready="1";
}
$("#newForm").addEventListener("submit", e => {
  e.preventDefault();
  const t = readForm("n_", blankTask());
  if (!t.tema){ toast("Escribe el tema de la tarea."); return; }
  saveTask(t); toast(`Tarea ${t.no} guardada`); $("#newForm").dataset.ready=""; renderNewForm();
});

/* ---------- Vista ---------- */
function kpiFilter(list){
  switch(ui.kpi){
    case "abiertas": return list.filter(t=>!isDone(t));
    case "vencidas": return list.filter(isLate);
    case "alta": return list.filter(t=>!isDone(t) && t.criticidad==="Alta");
    case "cumplidas": return list.filter(isDone);
    case "sinArea": return list.filter(t=>!t.departamento);
    default: return list;
  }
}
function renderKpis(){
  const c = {abiertas:tasks.filter(t=>!isDone(t)).length, vencidas:tasks.filter(isLate).length, alta:tasks.filter(t=>!isDone(t)&&t.criticidad==="Alta").length, cumplidas:tasks.filter(isDone).length, total:tasks.length, sinArea:tasks.filter(t=>!t.departamento).length};
  const items = [["abiertas","Abiertas",c.abiertas,""],["vencidas","Vencidas (entrega pasada)",c.vencidas,"bad"],["alta","Abiertas criticidad alta",c.alta,"warn"],["cumplidas","Cumplidas",c.cumplidas,"ok"],["todas","Total registradas",c.total,""]];
  if (c.sinArea) items.push(["sinArea","Sin departamento",c.sinArea,"warn"]);
  $("#kpis").innerHTML = items.map(([k,l,v,cl])=>`<button type="button" class="kpi ${cl}" data-kpi="${k}" aria-pressed="${ui.kpi===k}"><b>${v}</b><small>${l}</small></button>`).join("");
}
function renderChips(){
  const base = kpiFilter(tasks);
  const n = d => base.filter(t=>(t.departamento||"")===d).length;
  const deps = [...new Set([...cats.departamentos, ...tasks.map(t=>t.departamento).filter(Boolean)])];
  $("#depChips").innerHTML = `<button type="button" class="chip" data-dep="" aria-pressed="${ui.dep===""}">Todas<span class="n">${base.length}</span></button>` +
    deps.map(d=>`<button type="button" class="chip" data-dep="${esc(d)}" aria-pressed="${ui.dep===d}">${esc(d)}<span class="n">${n(d)}</span></button>`).join("");
}
function renderFilters(){
  $("#fEstado").innerHTML = opts(cats.estados, ui.estado, "Todos los estados");
  $("#fCrit").innerHTML = opts(cats.criticidades, ui.crit, "Toda criticidad");
  $("#fResp").innerHTML = opts(cats.responsables, ui.resp, "Todos los responsables");
  $("#fTexto").value = ui.texto;
}
function filtered(){
  let l = kpiFilter(tasks);
  if (ui.dep) l = l.filter(t=>(t.departamento||"")===ui.dep);
  if (ui.estado) l = l.filter(t=>t.estado===ui.estado);
  if (ui.crit) l = l.filter(t=>t.criticidad===ui.crit);
  if (ui.resp) l = l.filter(t=>t.responsable===ui.resp);
  if (ui.texto){ const q=norm(ui.texto); l = l.filter(t=>norm(t.tema+" "+t.observaciones+" "+t.departamento).includes(q)); }
  const key = ui.sort, dir = ui.dir==="desc"?-1:1;
  const val = t => key==="dias" ? (daysSince(t.fechaPendiente) ?? -1e9) : key==="no"||key==="avance" ? (Number(t[key])||0) : String(t[key]||"￿");
  return [...l].sort((a,b)=>{ const x=val(a), y=val(b); return (x<y?-1:x>y?1:0)*dir || (Number(a.no)-Number(b.no)); });
}
function renderTable(){
  const rows = filtered();
  document.querySelectorAll("thead th").forEach(th=>{ th.classList.toggle("asc", th.dataset.sort===ui.sort && ui.dir==="asc"); th.classList.toggle("desc", th.dataset.sort===ui.sort && ui.dir==="desc"); });
  $("#tbody").innerHTML = rows.map(t=>{
    const d = daysSince(t.fechaPendiente), late = isLate(t);
    return `<tr data-id="${esc(t.id)}" tabindex="0">
      <td class="num c-no">${esc(t.no)}</td>
      <td class="c-dep">${esc(t.departamento) || '<span style="color:var(--warn)">Sin departamento</span>'}</td>
      <td class="tema">${esc(t.tema)}${t.origen==="voz"?'<span class="tag-voz">VOZ</span>':""}${t.pendiente?'<span class="tag-voz" style="color:var(--warn);border-color:var(--warn)">SIN ENVIAR</span>':""}${t.observaciones?`<span class="obs">${esc(t.observaciones)}</span>`:""}</td>
      <td class="date c-fp">${fmt(t.fechaPendiente)}</td>
      <td class="num c-days ${d!=null && !isDone(t) && d>30?"days late":""}">${isDone(t)||d==null?"":d}</td>
      <td class="date c-fe ${late?"days late":""}">${fmt(t.fechaEntrega)}</td>
      <td class="c-est"><span class="pill ${estadoClass(t.estado)}">${esc((t.estado||"Sin estado").replace(/^\d\./,""))}</span></td>
      <td class="c-av">${t.avance!=null?`<span class="bar"><i style="width:${Math.round(t.avance*100)}%"></i></span><span class="num" style="font-family:var(--f-mono);font-size:12px">${Math.round(t.avance*100)}%</span>`:""}</td>
      <td class="c-cri">${t.criticidad?`<span class="crit ${esc(t.criticidad)}">${esc(t.criticidad)}</span>`:""}</td>
      <td class="c-resp">${esc(t.responsable)}</td>
    </tr>`;
  }).join("");
  const empty = $("#empty");
  if (!tasks.length){ empty.hidden=false; empty.innerHTML = cfg.flowUrl ? `<b>Aún no hay tareas</b>Pulsa el botón de sincronizar para traerlas del Excel, o dicta una nota de voz para registrar la primera.` : `<b>Aún no hay tareas en este teléfono</b>Toca “Conexión” para traer tu listado del Excel, o dicta una nota de voz para registrar la primera.`; }
  else if (!rows.length){ empty.hidden=false; empty.innerHTML=`<b>Ninguna tarea con estos filtros</b>Cambia el indicador, el departamento o la búsqueda.`; }
  else empty.hidden=true;
}
function renderCats(){
  $("#cats").innerHTML = Object.entries(CAT_LABELS).map(([k,l])=>`
    <div class="cat panel"><h3>${l}</h3><ul>${cats[k].map(v=>`<li><span>${esc(v)}</span></li>`).join("")}</ul></div>`).join("") +
    `<div class="cat panel"><h3>Avance</h3><p class="hint" style="margin:0">Escala de 10 % a 100 %, igual que en el Excel.</p></div>`;
}
function render(){ renderKpis(); renderChips(); renderTable(); }
function renderAll(){ renderFilters(); renderNewForm(); renderCats(); render(); }

$("#kpis").addEventListener("click", e=>{ const b=e.target.closest("[data-kpi]"); if(!b) return; ui.kpi=b.dataset.kpi; saveUi(); render(); });
$("#depChips").addEventListener("click", e=>{ const b=e.target.closest("[data-dep]"); if(!b) return; ui.dep=b.dataset.dep; saveUi(); render(); });
$("#fEstado").addEventListener("change", e=>{ ui.estado=e.target.value; if(ui.estado) ui.kpi="todas"; saveUi(); render(); });
$("#fCrit").addEventListener("change", e=>{ ui.crit=e.target.value; saveUi(); render(); });
$("#fResp").addEventListener("change", e=>{ ui.resp=e.target.value; saveUi(); render(); });
$("#fTexto").addEventListener("input", e=>{ ui.texto=e.target.value; saveUi(); renderTable(); });
document.querySelector("thead").addEventListener("click", e=>{ const th=e.target.closest("th[data-sort]"); if(!th) return; const k=th.dataset.sort; ui.dir = ui.sort===k && ui.dir==="asc" ? "desc":"asc"; ui.sort=k; saveUi(); renderTable(); });
$("#tbody").addEventListener("click", e=>{ const tr=e.target.closest("tr[data-id]"); if(tr) openEdit(tr.dataset.id); });
$("#tbody").addEventListener("keydown", e=>{ if(e.key==="Enter"){ const tr=e.target.closest("tr[data-id]"); if(tr) openEdit(tr.dataset.id); } });
function showTab(c){ $("#tabTareas").setAttribute("aria-selected", !c); $("#tabCats").setAttribute("aria-selected", c); $("#viewTareas").hidden=c; $("#viewCats").hidden=!c; }
$("#tabTareas").onclick=()=>showTab(false); $("#tabCats").onclick=()=>showTab(true);

/* ---------- Editar ---------- */
let editing = null;
function openEdit(id){
  const t = tasks.find(x=>x.id===id); if(!t) return;
  editing = t;
  $("#mTitle").textContent = `Tarea ${t.no ?? ""}`;
  $("#editForm").innerHTML = formHtml(t,"e_");
  $("#mRo").hidden = !t.sinId; $("#mSave").disabled = !!t.sinId;
  $("#delZone").innerHTML = t.sinId ? "" : `<button class="btn danger" id="mDel" type="button">Eliminar</button>`;
  $("#overlay").hidden = false;
}
function closeEdit(){ $("#overlay").hidden = true; editing = null; }
$("#mClose").onclick = closeEdit; $("#mCancel").onclick = closeEdit;
$("#overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeEdit(); });
document.addEventListener("keydown", e=>{ if(e.key==="Escape"){ if(!$("#overlay").hidden) closeEdit(); if(!$("#cfgOverlay").hidden) $("#cfgOverlay").hidden=true; } });
$("#mSave").onclick = ()=>{
  if(!editing || editing.sinId) return;
  const t = readForm("e_", editing);
  if(!t.tema){ toast("El tema no puede quedar vacío."); return; }
  saveTask(t); toast(`Tarea ${t.no} guardada`); closeEdit();
};
$("#delZone").addEventListener("click", e=>{
  if (e.target.id==="mDel") $("#delZone").innerHTML = `<div class="confirm-del">¿Eliminar esta tarea? <button class="btn danger" id="mDelYes" type="button">Sí, eliminar</button><button class="btn" id="mDelNo" type="button">No</button></div>`;
  else if (e.target.id==="mDelNo") $("#delZone").innerHTML = `<button class="btn danger" id="mDel" type="button">Eliminar</button>`;
  else if (e.target.id==="mDelYes" && editing){ const n=editing.no; deleteTask(editing.id); toast(`Tarea ${n} eliminada`); closeEdit(); }
});

/* ---------- Conexión ---------- */
function openCfg(){ $("#cfgFlow").value = cfg.flowUrl||""; $("#cfgExcel").value = cfg.excelUrl||""; $("#cfgMsg").textContent=""; $("#cfgMsg").className="hint"; $("#cfgOverlay").hidden=false; }
$("#btnCfg").onclick = openCfg;
$("#cfgClose").onclick = ()=>$("#cfgOverlay").hidden=true;
$("#cfgOverlay").addEventListener("click", e=>{ if(e.target.id==="cfgOverlay") $("#cfgOverlay").hidden=true; });
function readCfg(){ return {flowUrl:$("#cfgFlow").value.trim(), excelUrl:$("#cfgExcel").value.trim()}; }
$("#cfgTest").onclick = async ()=>{
  const c = readCfg(); const msg=$("#cfgMsg");
  if (!/^https:\/\//.test(c.flowUrl)){ msg.className="err-msg"; msg.textContent="Pega la URL completa del flujo; empieza por https://"; return; }
  msg.className="hint"; msg.textContent="Probando…";
  const prev = cfg.flowUrl; cfg.flowUrl = c.flowUrl;
  try { const r = await callFlow({accion:"listar"});
    if (!Array.isArray(r.tareas)) throw {code:"respuesta", detail:"el flujo respondió, pero sin la lista 'tareas'. Revisa la acción Respuesta del caso listar."};
    msg.className="ok-msg"; msg.textContent=`Conexión correcta: el Excel tiene ${r.tareas.length} filas en Tabla2.`;
  } catch(e){ msg.className="err-msg"; msg.textContent = e.net ? "No hay conexión a internet para probar." : "El flujo no respondió bien: " + (e.detail||e.code); }
  finally { cfg.flowUrl = prev; }
};
$("#cfgSave").onclick = ()=>{
  cfg = readCfg(); lsSet(K.cfg, cfg); $("#cfgOverlay").hidden=true; toast("Conexión guardada"); updateSync(); syncAll(true);
};
$("#btnSync").onclick = ()=>{
  if (!cfg.flowUrl){ openCfg(); return; }
  if (failed.length){ outbox = [...failed.map(f=>({...f, tries:0})), ...outbox]; failed=[]; lsSet(K.failed, failed); lsSet(K.out, outbox); }
  syncAll(true);
};
$("#btnExcel").onclick = ()=>{
  if (!cfg.excelUrl){ openCfg(); toast("Agrega el enlace de tu Excel en OneDrive."); return; }
  window.open(cfg.excelUrl, "_blank", "noopener");
};

/* ---------- Voz ---------- */
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
let rec = null, recOn = false, baseText = "";
function micHint(msg, alert){ const h=$("#micHint"); h.hidden=false; h.textContent=msg; h.classList.toggle("alert",!!alert); }
function keyboardHint(offline){ micHint(offline ? "Sin señal el dictado del navegador no funciona. Toca el cuadro y usa el micrófono del teclado: con el paquete de español descargado en Gboard o iPhone, dicta sin internet." : "El dictado directo no está disponible aquí. Toca el cuadro y usa el micrófono del teclado (Gboard o iPhone); luego pulsa “Convertir en tareas”.", true); $("#dictado").focus(); }
$("#btnMic").onclick = () => {
  if (recOn && rec){ rec.stop(); return; }
  if (!SR){ keyboardHint(false); return; }
  if (navigator.onLine===false){ keyboardHint(true); return; }
  try {
    rec = new SR(); rec.lang = "es-CO"; rec.continuous = true; rec.interimResults = true;
    baseText = $("#dictado").value ? $("#dictado").value.trim() + " " : "";
    rec.onresult = ev => { let fin="", inter=""; for (let i=0;i<ev.results.length;i++){ const r=ev.results[i]; if(r.isFinal) fin+=r[0].transcript; else inter+=r[0].transcript; } $("#dictado").value = baseText + fin + inter; };
    rec.onstart = () => { recOn=true; $("#btnMic").classList.add("rec"); $("#vozStatus").textContent="Escuchando… toca de nuevo para terminar"; };
    rec.onend = () => { recOn=false; $("#btnMic").classList.remove("rec"); $("#vozStatus").textContent = $("#dictado").value.trim() ? "Listo. Revisa el texto y conviértelo en tareas" : "Toca el micrófono y dicta"; };
    rec.onerror = ev => { recOn=false; $("#btnMic").classList.remove("rec"); if (ev.error==="network") keyboardHint(true); else if (["not-allowed","service-not-allowed","audio-capture"].includes(ev.error)) keyboardHint(false); else if(ev.error!=="no-speech" && ev.error!=="aborted") micHint("El dictado se interrumpió ("+ev.error+"). Sigue con el micrófono del teclado.", true); };
    rec.start();
  } catch(err){ keyboardHint(false); }
};
$("#btnClearVoz").onclick = () => { $("#dictado").value=""; drafts=[]; renderDrafts(); $("#vozStatus").textContent="Toca el micrófono y dicta"; };
try { const d = localStorage.getItem("tg_draft_text"); if (d) $("#dictado").value = d; } catch(e){}
$("#dictado").addEventListener("input", e=>{ try{ localStorage.setItem("tg_draft_text", e.target.value); }catch(_){} });

/* Interpretación de la nota (funciona sin internet) */
const DEP_SYN = [["taller","Taller Agrícola"],["maquinaria","Taller Agrícola"],["tractor","Taller Agrícola"],["equipo","Taller Agrícola"],["fertiliz","Fertilización"],["enmienda","Fertilización"],["abono","Fertilización"],["cosecha","Cosecha"],["cortero","Cosecha"],["racimo","Cosecha"],["vivero","Vivero"],["plantula","Vivero"],["sanidad","Sanidad"],["plaga","Sanidad"],["picudo","Sanidad"],["pudricion","Sanidad"],["enfermedad","Sanidad"],["proveedor","Proveedores"],["rff","Proveedores"],["mantenimiento","Mantenimiento"],["drenaje","Mantenimiento"],["junta","Junta"],["proyecto","Proyectos"],["siembra","Proyectos"],["administra","Administración"],["forecast","Administración"],["nomina","Administración"],["gerente general","CEO"],["ceo","CEO"]];
const MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
const DIAS = ["domingo","lunes","martes","miercoles","jueves","viernes","sabado"];
function stripMeta(s){
  return s.replace(/\b(criticidad|prioridad)\s+(alta|media|baja)\b/gi,"")
    .replace(/\b(es\s+)?urgente\b/gi,"")
    .replace(/\b(para|antes del?|hasta el)\s+(el\s+)?(hoy|mañana|pasado mañana|(lunes|martes|miércoles|miercoles|jueves|viernes|sábado|sabado|domingo)(\s+de la (próxima|proxima|siguiente) semana)?|\d{1,2}\s+de\s+[a-záéíóú]+|\d{1,2}\/\d{1,2}(\/\d{2,4})?|fin de mes|la (próxima|proxima|siguiente) semana)\b/gi,"")
    .replace(/\b(el\s+)?\d{1,2}\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\b/gi,"")
    .replace(/\s{2,}/g," ").replace(/\s+([,.])/g,"$1").replace(/(^[\s,.:;-]+|[\s,.:;-]+$)/g,"");
}
function localParse(text){
  const parts = text.split(/\n+|\b(?:otra tarea|siguiente tarea|nueva tarea|también|tambien)\b[:,]?/i).map(s=>s.trim()).filter(s=>s.length>3);
  return parts.map(p=>{
    const n = norm(p); const t = {tema:p, departamento:"", criticidad:"", fechaEntrega:"", responsable:"", observaciones:""};
    let depWord = "";
    for (const d of cats.departamentos) if (n.includes(norm(d))) { t.departamento=d; depWord=d; break; }
    if (!t.departamento) for (const [k,d] of DEP_SYN) if (n.includes(k) && cats.departamentos.includes(d)) { t.departamento=d; break; }
    if (/urgente|critic|prioridad alta|criticidad alta/.test(n)) t.criticidad="Alta"; else if (/(prioridad|criticidad) baja/.test(n)) t.criticidad="Baja"; else if (/(prioridad|criticidad) media/.test(n)) t.criticidad="Media";
    for (const r of cats.responsables) if (n.includes(norm(r))) { t.responsable=r; break; }
    const td = today(); let m;
    if (/pasado manana/.test(n)) t.fechaEntrega = iso(new Date(td.getTime()+2*864e5));
    else if (/\bmanana\b/.test(n)) t.fechaEntrega = iso(new Date(td.getTime()+864e5));
    else if (/\bpara hoy\b/.test(n)) t.fechaEntrega = iso(td);
    else if ((m = n.match(/(\d{1,2}) de (enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)/))){ let d=new Date(td.getFullYear(), MESES.indexOf(m[2]), Number(m[1])); if (d<td) d.setFullYear(d.getFullYear()+1); t.fechaEntrega=iso(d); }
    else if ((m = n.match(/(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?/))){ const y=m[3]?Number(m[3].length===2?"20"+m[3]:m[3]):td.getFullYear(); t.fechaEntrega=iso(new Date(y,Number(m[2])-1,Number(m[1]))); }
    else if ((m = n.match(/\b(lunes|martes|miercoles|jueves|viernes|sabado|domingo)\b/))){ const w=DIAS.indexOf(m[1]); let add=(w-td.getDay()+7)%7||7; if(/proxima semana|siguiente semana/.test(n)) add+=7; t.fechaEntrega=iso(new Date(td.getTime()+add*864e5)); }
    else if (/fin de mes/.test(n)) t.fechaEntrega=iso(new Date(td.getFullYear(), td.getMonth()+1, 0));
    else if (/proxima semana|siguiente semana/.test(n)) t.fechaEntrega=iso(new Date(td.getTime()+7*864e5));
    let tema = p.replace(/^(agregar|anotar|tarea|registrar|recordar)[:,]?\s*/i,"");
    if (depWord) tema = tema.replace(new RegExp("^\\s*"+depWord.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")+"\\s*[,:.-]?\\s*","i"),"");
    tema = stripMeta(tema) || p;
    t.tema = tema.charAt(0).toUpperCase()+tema.slice(1);
    return t;
  });
}
let drafts = [];
$("#btnParse").onclick = () => {
  const text = $("#dictado").value.trim();
  if (!text){ toast("Primero dicta o escribe la nota."); $("#dictado").focus(); return; }
  if (recOn && rec) rec.stop();
  const list = localParse(text);
  drafts = list.map(t=>({...blankTask(), ...t, origen:"voz"}));
  $("#vozStatus").textContent = `${drafts.length} tarea${drafts.length>1?"s":""} detectada${drafts.length>1?"s":""} — revisa y guarda`;
  renderDrafts();
};
function renderDrafts(){
  $("#drafts").innerHTML = drafts.map((t,i)=>`
    <div class="draft" data-i="${i}">
      <label class="f">Tema<input class="in" data-k="tema" value="${esc(t.tema)}"></label>
      <div class="two">
        <label class="f">Departamento<select class="in" data-k="departamento">${opts(cats.departamentos,t.departamento,"Sin departamento")}</select></label>
        <label class="f">Criticidad<select class="in" data-k="criticidad">${opts(cats.criticidades,t.criticidad,"—")}</select></label>
        <label class="f">Entrega<input class="in" type="date" data-k="fechaEntrega" value="${esc(t.fechaEntrega)}"></label>
        <label class="f">Responsable<select class="in" data-k="responsable">${opts(cats.responsables,t.responsable,"Sin asignar")}</select></label>
      </div>
      <label class="f">Observaciones<input class="in" data-k="observaciones" value="${esc(t.observaciones)}"></label>
      <div class="row"><button class="btn primary" type="button" data-act="save">Guardar</button><button class="btn ghost" type="button" data-act="drop">Descartar</button></div>
    </div>`).join("") + (drafts.length>1?`<button class="btn primary" type="button" id="saveAllDrafts">Guardar las ${drafts.length} tareas</button>`:"");
}
$("#drafts").addEventListener("input", e=>{ const card=e.target.closest(".draft"); const k=e.target.dataset.k; if(card && k) drafts[Number(card.dataset.i)][k]=e.target.value; });
$("#drafts").addEventListener("change", e=>{ const card=e.target.closest(".draft"); const k=e.target.dataset.k; if(card && k) drafts[Number(card.dataset.i)][k]=e.target.value; });
$("#drafts").addEventListener("click", e=>{
  if (e.target.id==="saveAllDrafts"){ saveDrafts([...drafts.keys()]); return; }
  const act=e.target.dataset.act; const card=e.target.closest(".draft"); if(!act||!card) return;
  const i=Number(card.dataset.i);
  if (act==="drop"){ drafts.splice(i,1); renderDrafts(); } else saveDrafts([i]);
});
function saveDrafts(idxs){
  let saved=0;
  for (const i of idxs){ const t={...drafts[i], no:nextNo()}; if(!t.tema) continue; saveTask(t); saved++; }
  drafts = drafts.filter((_,i)=>!idxs.includes(i)); renderDrafts();
  if (saved){ toast(`${saved} tarea${saved>1?"s":""} guardada${saved>1?"s":""}${navigator.onLine===false?" en el teléfono; se enviarán al volver la señal":""}`); if(!drafts.length){ $("#dictado").value=""; try{localStorage.removeItem("tg_draft_text");}catch(e){} $("#vozStatus").textContent="Toca el micrófono y dicta"; } }
}

/* ---------- Arranque ---------- */
compose(); renderAll(); updateSync();
window.addEventListener("online", ()=>{ updateSync(); syncAll(false); });
window.addEventListener("offline", updateSync);
document.addEventListener("visibilitychange", ()=>{ if (!document.hidden) syncAll(false); });
setInterval(()=>{ if (outbox.length) flush(); }, 20000);
setInterval(()=>{ if (!document.hidden && !outbox.length) pull(false); }, 5*60000);
syncAll(false);

if ("serviceWorker" in navigator){
  navigator.serviceWorker.register("sw.js").then(reg=>{
    reg.addEventListener("updatefound", ()=>{ const w=reg.installing; w && w.addEventListener("statechange", ()=>{ if (w.state==="installed" && navigator.serviceWorker.controller) toast("Hay una versión nueva de la app: ciérrala y ábrela de nuevo."); }); });
  }).catch(()=>{});
}
window.__app = {version:APP_VERSION, state:()=>({outbox, remote, tasks, failed})};
})();
