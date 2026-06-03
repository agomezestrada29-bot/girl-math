import { useState, useMemo, useCallback } from "react";

const C = {
  bg:"#0B0B12", surface:"#13131C", card:"#1A1A26", card2:"#1F1F2E",
  border:"#252535", text:"#EEEAF8", muted:"#6E6C80", soft:"#9896AA",
  accent:"#B78CF7", teal:"#2DD4BF", amber:"#F59E0B", red:"#F87171",
  blue:"#60A5FA", pink:"#F472B6", green:"#4ADE80", orange:"#FB923C",
};

const fmt = n => new Intl.NumberFormat("es-CO",{style:"currency",currency:"COP",maximumFractionDigits:0}).format(n||0);
const fmtShort = n => n>=1000000?`$${(n/1000000).toFixed(1)}M`:n>=1000?`$${Math.round(n/1000)}k`:fmt(n);
const today = new Date();
function daysUntil(day) {
  const d = new Date(today.getFullYear(),today.getMonth(),day);
  if (d<today) d.setMonth(d.getMonth()+1);
  return Math.ceil((d-today)/864e5);
}
const uid = () => Math.random().toString(36).slice(2,9);

const COLOR_OPTIONS = [
  {label:"Violeta", value:C.accent}, {label:"Teal",    value:C.teal},
  {label:"Azul",    value:C.blue},   {label:"Rosa",     value:C.pink},
  {label:"Ámbar",   value:C.amber},  {label:"Verde",    value:C.green},
  {label:"Rojo",    value:C.red},    {label:"Naranja",  value:C.orange},
  {label:"Púrpura", value:"#A78BFA"},{label:"Gris",     value:C.soft},
];

const CATS = {
  // categorías que usa tu app "Gestor de Gastos" — mapeadas a bolsillos
  ori:        {label:"Orión",           icon:"🐾", color:C.pink,    bolsilloId:"ori"   },
  carro:      {label:"Carro",           icon:"🚗", color:C.blue,    bolsilloId:"carro" },
  comida:     {label:"Comida",          icon:"🥗", color:C.green,   bolsilloId:"vida"  },
  resto:      {label:"Restaurantes",    icon:"🍽️", color:C.orange,  bolsilloId:"vida"  },
  transp:     {label:"Transporte",      icon:"🚕", color:C.blue,    bolsilloId:"carro" },
  subs:       {label:"Suscripciones",   icon:"📱", color:C.accent,  bolsilloId:"subs"  },
  salud:      {label:"Salud/Bienestar", icon:"💊", color:C.teal,    bolsilloId:"vida"  },
  ocio:       {label:"Ocio/Salidas",    icon:"🎉", color:C.pink,    bolsilloId:"vida"  },
  deuda:      {label:"Deudas/Adulting", icon:"💳", color:C.red,     bolsilloId:"tc"    },
  casa:       {label:"Casa",            icon:"🏠", color:C.teal,    bolsilloId:"vida"  },
  donaciones: {label:"Donaciones",      icon:"💝", color:C.pink,    bolsilloId:"vida"  },
  adulting:   {label:"Adulting",        icon:"📋", color:C.amber,   bolsilloId:"segsoc"},
  otros:      {label:"Otros",           icon:"📦", color:C.soft,    bolsilloId:"vida"  },
};

// mapa de categorías del Excel de "Gestor de Gastos" → ids internos
const CAT_MAP = {
  "Orion":"ori", "Orión":"ori",
  "Transporte":"transp",
  "Ocio":"ocio",
  "Suscripciones":"subs",
  "Salud":"salud",
  "Casa":"casa",
  "Donaciones":"donaciones",
  "Adulting":"adulting",
  "Comida":"comida",
  "Restaurantes":"resto",
  "Carro":"carro",
};

// parsea número serial de Excel (días desde 1900-01-01) a fecha JS
function excelDateToJS(serial) {
  if (!serial || isNaN(serial)) return new Date().toISOString().slice(0,10);
  const ms = (serial - 25569) * 86400 * 1000;
  return new Date(ms).toISOString().slice(0,10);
}

// carga SheetJS dinámicamente si no está disponible
async function loadXLSX() {
  if (window.XLSX) return window.XLSX;
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
    s.onload = () => resolve(window.XLSX);
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// lee el xlsx con SheetJS y devuelve array de gastos normalizados
async function parseGestorExcel(file) {
  const XLSX = await loadXLSX();
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, {type:"array"});
  const ws  = wb.Sheets["Gastos"] || wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});

  // fila 0 = título, fila 1 = encabezados, fila 2+ = datos
  const header = rows[1] || [];
  const iDate  = header.findIndex(h=>String(h).toLowerCase().includes("fecha"));
  const iCat   = header.findIndex(h=>String(h).toLowerCase().includes("categor"));
  const iAmt   = header.findIndex(h=>String(h).toLowerCase().includes("predeterminada") && String(h).toLowerCase().includes("cantidad"));
  const iNote  = header.findIndex(h=>String(h).toLowerCase().includes("comentario"));

  const gastos = [];
  for (let i=2; i<rows.length; i++) {
    const row = rows[i];
    if (!row[iAmt] || isNaN(row[iAmt]) || +row[iAmt]===0) continue;
    const catRaw = String(row[iCat]||"").trim();
    const catId  = CAT_MAP[catRaw] || "otros";
    gastos.push({
      id:    Math.random().toString(36).slice(2,9),
      nombre: String(row[iNote]||catRaw||"Sin descripción").trim(),
      monto:  Math.abs(+row[iAmt]),
      categoria: catId,
      bolsilloId: CATS[catId]?.bolsilloId || "vida",
      fecha: excelDateToJS(row[iDate]),
      importado: true,
    });
  }
  return gastos;
}

const BOLSILLOS_0 = [
  {id:"emerg",  nombre:"Fondo Emergencias", icon:"🛡️", color:C.teal,    saldo:523000, meta:2000000, grupo:"sagrado",   descripcion:"Solo emergencias reales — no tocar",         metaMensual:200000},
  {id:"buceo",  nombre:"IDC — Instructora", icon:"🤿", color:C.blue,    saldo:200000, meta:11000000,grupo:"sagrado",   descripcion:"Meta $11M · Fase 1: $3.2M antes de ago/sep", metaMensual:400000, metaFase1:3200000, fechaFase1:"ago/sep 2025"},
  {id:"viajes", nombre:"Viajes Colombia",   icon:"🗺️", color:C.pink,    saldo:194000, meta:1500000, grupo:"sagrado",   descripcion:"Explorar tu país",                           metaMensual:100000},
  {id:"vida",   nombre:"Vida Diaria",       icon:"☀️", color:C.accent,  saldo:670000, meta:800000,  grupo:"vida",      descripcion:"Comida, transporte, ocio, día a día",        metaMensual:800000},
  {id:"ori",    nombre:"Orión",             icon:"🐾", color:C.pink,    saldo:0,      meta:800000,  grupo:"vida",      descripcion:"Vida, meds, guardería, imprevistos",          metaMensual:800000},
  {id:"carro",  nombre:"Carro",             icon:"🚗", color:C.blue,    saldo:0,      meta:700000,  grupo:"vida",      descripcion:"Seguro, gasolina, imprevistos",               metaMensual:700000},
  {id:"subs",   nombre:"Suscripciones",     icon:"📱", color:C.accent,  saldo:0,      meta:136450,  grupo:"vida",      descripcion:"Spotify, Google, Apple + CapCut (ago)",      metaMensual:136450},
  {id:"segsoc", nombre:"Seguridad Social",  icon:"🏥", color:C.teal,    saldo:300000, meta:507000,  grupo:"provision", descripcion:"Pago obligatorio día 10 cada mes",           metaMensual:507000},
  {id:"tc",     nombre:"TC — Abono real",   icon:"💳", color:C.amber,   saldo:400000, meta:2100000, grupo:"provision", descripcion:"Mínimo $1.5M + $600k abono para bajar $5.9M",metaMensual:2100000, deudaTotal:5900000, minimoTC:1500000, abonoExtra:600000},
  {id:"eraia",  nombre:"Provisión Eraia",   icon:"🏢", color:"#A78BFA", saldo:224000, meta:2665000, grupo:"provision", descripcion:"Deuda con Eraia",                            metaMensual:300000},
  {id:"mama",   nombre:"Regalo Mamá",       icon:"💐", color:C.pink,    saldo:0,      meta:616000,  grupo:"provision", descripcion:"Día de la madre — 28 MAY",                   metaMensual:616000},
  {id:"capcut", nombre:"CapCut (agosto)",   icon:"🎬", color:C.orange,  saldo:0,      meta:30000,   grupo:"provision", descripcion:"Suscripción CapCut — provisionar para agosto",metaMensual:30000},
];

// urgencia: "critica" = provisionar siempre primero | "normal" = cuando haya espacio | "flexible" = solo si sobra
const PAGOS_0 = [
  {id:"imp",   nombre:"Impuesto Vehicular", monto:689500,  diaPago:30,pagado:false,prioridad:1,bolsilloId:"carro",  categoria:"carro", urgencia:"critica",  nota:"Multa si no se paga"},
  {id:"tcmin", nombre:"Mínimo TC",          monto:1500000, diaPago:5, pagado:false,prioridad:2,bolsilloId:"tc",     categoria:"deuda",  urgencia:"critica",  nota:"Deuda total $5.9M — pagar mínimo + extra para bajarla"},
  {id:"soat",  nombre:"SOAT Mayo",          monto:1118166, diaPago:5, pagado:false,prioridad:3,bolsilloId:"carro",  categoria:"carro",  urgencia:"critica",  nota:"Ilegal circular sin SOAT"},
  {id:"mama",  nombre:"Regalo Mamá",        monto:616000,  diaPago:28,pagado:false,prioridad:4,bolsilloId:"mama",   categoria:"otros",  urgencia:"normal",   nota:"Día de la madre"},
  {id:"segsoc",nombre:"Seguridad Social",   monto:507000,  diaPago:10,pagado:false,prioridad:5,bolsilloId:"segsoc", categoria:"salud",  urgencia:"critica",  nota:"Pago obligatorio mensual"},
  {id:"capcut",nombre:"CapCut (agosto)",    monto:30000,   diaPago:1, pagado:false,prioridad:6,bolsilloId:"capcut", categoria:"subs",   urgencia:"normal",   nota:"Provisionar desde ya — vence agosto"},
  {id:"juan",  nombre:"Deuda Juan",         monto:633639,  diaPago:30,pagado:false,prioridad:7,bolsilloId:"eraia",  categoria:"deuda",  urgencia:"flexible", nota:"Puede esperar si hay necesidad"},
  {id:"lauri", nombre:"Deuda Lauri",        monto:171340,  diaPago:30,pagado:false,prioridad:8,bolsilloId:"eraia",  categoria:"deuda",  urgencia:"flexible", nota:"Puede esperar si hay necesidad"},
];

const INGRESOS_0 = [
  {id:"aas",   fuente:"AAS",          monto:1500000,diaEsp:28,recibido:false},
  {id:"gomez", fuente:"Gomez Estrada",monto:1000000,diaEsp:3, recibido:false},
  {id:"coral", fuente:"Casa Coral",   monto:850000, diaEsp:3, recibido:false},
  {id:"nao",   fuente:"NAO",          monto:261000, diaEsp:5, recibido:false},
  {id:"saduma",fuente:"Saduma",       monto:262500, diaEsp:5, recibido:false},
  {id:"eraia", fuente:"Eraia Admin",  monto:210000, diaEsp:5, recibido:false},
];

// urgencia badges config
const URGENCIA = {
  critica:  {label:"🔴 Crítica",  color:C.red,   desc:"Provisionar siempre primero"},
  normal:   {label:"🟡 Normal",   color:C.amber, desc:"Provisionar cuando haya margen"},
  flexible: {label:"🟢 Flexible", color:C.teal,  desc:"Solo si sobra después de lo demás"},
};

// ─── CONSEJERA ───────────────────────────────────────────────────────────────
// Calcula meses restantes hasta una fecha objetivo "mes/año" string
function mesesHasta(fechaStr) {
  if (!fechaStr) return 99;
  const [mesNombre, anio] = fechaStr.split("/");
  const meses = {ene:0,feb:1,mar:2,abr:3,may:4,jun:5,jul:6,ago:7,sep:8,oct:9,nov:10,dic:11};
  const mesNum = meses[mesNombre?.toLowerCase().slice(0,3)] ?? 7;
  const target = new Date(+anio || today.getFullYear(), mesNum, 1);
  const diff = (target.getFullYear()-today.getFullYear())*12 + (target.getMonth()-today.getMonth());
  return Math.max(diff, 1);
}

function generarPlan(monto, bolsillos, pagos) {
  const plan=[]; let resto=monto;
  const push=(nombre,icon,color,tipo,mnt,razon,bolsilloId)=>{
    if(mnt<=0||resto<=0) return;
    const real=Math.min(mnt,resto);
    plan.push({nombre,icon,color,tipo,monto:real,razon,bolsilloId});
    resto-=real;
  };

  // P1 — críticos venciendo ≤7 días
  const criticos = pagos
    .filter(p=>!p.pagado && p.urgencia==="critica" && daysUntil(p.diaPago)<=7)
    .sort((a,b)=>a.prioridad-b.prioridad);
  for(const p of criticos){
    const b=bolsillos.find(x=>x.id===p.bolsilloId);
    const falta=Math.max(p.monto-(b?.saldo||0),0);
    if(falta>0) push(`Completar: ${p.nombre}`,CATS[p.categoria]?.icon||"🔴",C.red,"urgente",falta,`🔴 Crítico · vence en ${daysUntil(p.diaPago)} día(s)${p.nota?` — ${p.nota}`:""}`,p.bolsilloId);
  }

  // P2 — críticos lejanos: provisionar cuota proporcional
  const criticosLejos = pagos
    .filter(p=>!p.pagado && p.urgencia==="critica" && daysUntil(p.diaPago)>7)
    .sort((a,b)=>a.prioridad-b.prioridad);
  for(const p of criticosLejos){
    if(resto<=0) break;
    const b=bolsillos.find(x=>x.id===p.bolsilloId);
    // Para TC: meta = mínimo + abono extra
    const metaReal = b?.minimoTC ? (b.minimoTC + (b.abonoExtra||0)) : p.monto;
    const falta=Math.max(metaReal-(b?.saldo||0),0);
    if(falta>0){
      const aporte=Math.min(falta, Math.round(resto*0.25));
      if(aporte>0){
        const esTC = p.id==="tcmin";
        const razon = esTC
          ? `🔴 Crítico · mínimo $1.5M + $600k abono = menos intereses sobre $5.9M`
          : `🔴 Crítico · faltan ${daysUntil(p.diaPago)} días (${fmt(b?.saldo||0)} provisionado)`;
        push(`Provisionar: ${p.nombre}`,CATS[p.categoria]?.icon||"💳",C.amber,"provision",aporte,razon,p.bolsilloId);
      }
    }
  }

  // P3 — VIDA DIARIA: siempre aparece, llevar al 70% mínimo
  for(const bid of["vida","ori","carro","subs"]){
    const b=bolsillos.find(x=>x.id===bid); if(!b) continue;
    const min=Math.round(b.meta*0.7);
    const falta=Math.max(min-b.saldo,0);
    const pct = bid==="vida"?0.28:bid==="ori"?0.18:bid==="carro"?0.15:0.05;
    const aporte=falta>0 ? Math.min(falta,Math.round(resto*pct)) : 0;
    // Siempre mostrar vida diaria aunque esté cubierta, como recordatorio
    if(b.saldo>=min){
      plan.push({nombre:b.nombre,icon:b.icon,color:b.color+"80",tipo:"vida",monto:0,razon:`✓ Cubierto (${fmt(b.saldo)} / ${fmt(min)} mínimo)`,bolsilloId:bid, informativo:true});
    } else if(aporte>0){
      push(b.nombre,b.icon,b.color,"vida",aporte,`Llevar a mínimo del mes: ${fmt(min)} (tienes ${fmt(b.saldo)})`,bid);
    }
  }

  // P4 — normales
  const normales = pagos.filter(p=>!p.pagado && p.urgencia==="normal").sort((a,b)=>a.prioridad-b.prioridad);
  for(const p of normales){
    if(resto<=0) break;
    const b=bolsillos.find(x=>x.id===p.bolsilloId);
    const falta=Math.max(p.monto-(b?.saldo||0),0);
    if(falta>0){
      const aporte=Math.min(falta, Math.round(resto*0.2));
      if(aporte>0) push(`Provisionar: ${p.nombre}`,CATS[p.categoria]?.icon||"📌",C.amber,"provision",aporte,`🟡 Normal · día ${p.diaPago}${p.nota?` — ${p.nota}`:""}`,p.bolsilloId);
    }
  }

  // P5 — fondo emergencias si está bajo 30%
  const emerg=bolsillos.find(x=>x.id==="emerg");
  if(emerg&&emerg.saldo<emerg.meta*0.3&&resto>50000) push(emerg.nombre,emerg.icon,emerg.color,"ahorro",Math.min(Math.round(resto*0.3),emerg.meta-emerg.saldo),"Colchón bajo — reforzar",`emerg`);

  // P6 — BUCEO fase 1: cuota mensual necesaria para llegar a $3.2M antes de ago/sep
  const buceo = bolsillos.find(x=>x.id==="buceo");
  if(buceo && buceo.metaFase1 && buceo.saldo < buceo.metaFase1 && resto>0){
    const mesesRestantes = mesesHasta(buceo.fechaFase1);
    const faltaFase1 = buceo.metaFase1 - buceo.saldo;
    const cuotaIdeal = Math.ceil(faltaFase1 / mesesRestantes);
    const aporte = Math.min(cuotaIdeal, Math.round(resto*0.15));
    if(aporte>0) push(buceo.nombre,buceo.icon,buceo.color,"sueño",aporte,
      `🤿 Fase 1: necesitas ${fmt(cuotaIdeal)}/mes para llegar a ${fmt(buceo.metaFase1)} en ${buceo.fechaFase1} (faltan ${fmt(faltaFase1)})`,
      "buceo");
  }

  // Viajes
  const viajes=bolsillos.find(x=>x.id==="viajes");
  if(viajes&&viajes.saldo<viajes.meta&&resto>50000){
    const aporte=Math.min(Math.round(resto*0.1),viajes.meta-viajes.saldo);
    if(aporte>0) push(viajes.nombre,viajes.icon,viajes.color,"sueño",aporte,`🗺️ Acumular para viajes por Colombia`,`viajes`);
  }

  // P7 — flexibles
  const flexibles = pagos.filter(p=>!p.pagado && p.urgencia==="flexible").sort((a,b)=>a.prioridad-b.prioridad);
  for(const p of flexibles){
    if(resto<=0) break;
    const b=bolsillos.find(x=>x.id===p.bolsilloId);
    const falta=Math.max(p.monto-(b?.saldo||0),0);
    if(falta>0){
      const aporte=Math.min(falta, Math.round(resto*0.3));
      if(aporte>0) push(`Abonar: ${p.nombre}`,CATS[p.categoria]?.icon||"🟢",C.teal,"flexible",aporte,`🟢 Flexible · lo que se pueda (${fmt(b?.saldo||0)} provisionado)${p.nota?` — ${p.nota}`:""}`,p.bolsilloId);
    }
  }

  if(resto>0) plan.push({nombre:"Sobrante / Libre",icon:"✨",color:C.soft,tipo:"libre",monto:resto,razon:"Para lo que surja o reforzar sueños",bolsilloId:"vida"});
  return plan;
}

function generarAlertas(gastosList, bolsillos, pagos) {
  const alertas=[];
  const diasMes=new Date(today.getFullYear(),today.getMonth()+1,0).getDate();
  const pctMes=today.getDate()/diasMes;

  // alertas de gastos vs bolsillos
  const porCat={};
  gastosList.filter(g=>{const d=new Date(g.fecha);return d.getMonth()===today.getMonth()&&d.getFullYear()===today.getFullYear();})
    .forEach(g=>{porCat[g.categoria]=(porCat[g.categoria]||0)+g.monto;});
  for(const [cat,total] of Object.entries(porCat)){
    const ci=CATS[cat]; if(!ci?.bolsilloId) continue;
    const b=bolsillos.find(x=>x.id===ci.bolsilloId); if(!b) continue;
    const p=total/b.meta;
    if(p>0.9) alertas.push({nivel:"rojo",icon:"🔴",msg:`${ci.label}: casi sin saldo en ${b.nombre} (${fmt(b.meta-total)} restante)`,color:C.red});
    else if(p>pctMes*1.3) alertas.push({nivel:"amarillo",icon:"🟡",msg:`${ci.label}: ritmo acelerado (${fmt(total)} con ${Math.round(pctMes*100)}% del mes)`,color:C.amber});
  }

  // alerta TC: si solo está provisionando el mínimo sin abono extra
  const tcB = bolsillos.find(x=>x.id==="tc");
  if(tcB && tcB.deudaTotal){
    if(tcB.saldo < tcB.minimoTC) alertas.push({nivel:"rojo",icon:"💳",msg:`TC: solo tienes ${fmt(tcB.saldo)} — necesitas ${fmt(tcB.minimoTC)} mínimo. Riesgo de más intereses sobre ${fmt(tcB.deudaTotal)}.`,color:C.red});
    else if(tcB.saldo < tcB.minimoTC + tcB.abonoExtra) alertas.push({nivel:"amarillo",icon:"💳",msg:`TC: mínimo cubierto pero sin abono extra. Agrega ${fmt(tcB.abonoExtra)} más para bajar la deuda de ${fmt(tcB.deudaTotal)}.`,color:C.amber});
  }

  // alerta buceo fase 1
  const buceo = bolsillos.find(x=>x.id==="buceo");
  if(buceo && buceo.metaFase1 && buceo.saldo < buceo.metaFase1){
    const mesesR = mesesHasta(buceo.fechaFase1);
    const cuota = Math.ceil((buceo.metaFase1-buceo.saldo)/mesesR);
    alertas.push({nivel:"info",icon:"🤿",msg:`IDC: necesitas ahorrar ${fmt(cuota)}/mes para llegar a ${fmt(buceo.metaFase1)} en ${buceo.fechaFase1} (tienes ${fmt(buceo.saldo)})`,color:C.blue});
  }

  return alertas;
}

// ─── ESTILOS ─────────────────────────────────────────────────────────────────
const S = {
  card: {background:C.card,border:`1px solid ${C.border}`,borderRadius:16,padding:"14px 16px",marginBottom:9},
  inp:  {width:"100%",background:"#ffffff0c",border:`1px solid ${C.border}`,borderRadius:10,color:C.text,fontSize:13,padding:"9px 13px",outline:"none",fontFamily:"inherit"},
  btn:  (bg,col)=>({background:bg,border:"none",borderRadius:10,color:col||"#0B0B12",fontSize:13,fontWeight:600,padding:"10px 16px",cursor:"pointer"}),
  tab:  (a)=>({background:a?C.accent+"20":"transparent",color:a?C.accent:C.muted,border:"none",borderBottom:a?`2px solid ${C.accent}`:"2px solid transparent",padding:"9px 12px",fontSize:11,fontWeight:a?700:500,cursor:"pointer",whiteSpace:"nowrap"}),
  lbl:  {fontSize:10,color:C.muted,marginBottom:4,letterSpacing:.5},
  sec:  {fontSize:10,color:C.muted,letterSpacing:2,textTransform:"uppercase",margin:"14px 0 8px"},
};

function ProgressBar({saldo,meta,color,height=5}){
  return <div style={{background:"#ffffff0e",borderRadius:999,height,overflow:"hidden",marginTop:6}}><div style={{width:`${Math.min(saldo/meta*100,100)}%`,height:"100%",background:color,borderRadius:999,transition:"width .5s"}}/></div>;
}
function Badge({text,color}){
  return <span style={{fontSize:9,background:color+"22",color,borderRadius:5,padding:"2px 7px",fontWeight:700,marginLeft:5}}>{text}</span>;
}

// ─── MODAL BASE ───────────────────────────────────────────────────────────────
function Modal({title,color,onClose,children}){
  return (
    <div style={{position:"fixed",inset:0,background:"#000000CC",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:C.surface,border:`1px solid ${color||C.border}`,borderRadius:"20px 20px 0 0",padding:"20px 20px 32px",width:"100%",maxWidth:720,maxHeight:"85vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontSize:15,fontWeight:700,color:color||C.text}}>{title}</div>
          <button onClick={onClose} style={{background:"#ffffff0f",border:"none",borderRadius:8,color:C.muted,padding:"5px 10px",cursor:"pointer",fontSize:13}}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

// ─── FORM FIELDS ──────────────────────────────────────────────────────────────
function Field({label,children}){return <div style={{marginBottom:10}}><div style={S.lbl}>{label}</div>{children}</div>;}
function Inp(props){return <input style={S.inp} {...props}/>;}
function Sel({children,...props}){return <select style={{...S.inp,background:C.card}} {...props}>{children}</select>;}

// ════════════════════════════════════════════════════════════════════════════
export default function App(){
  const [tab,        setTab]        = useState("consejera");
  const [bolsillos,  setBolsillos]  = useState(BOLSILLOS_0);
  const [pagos,      setPagos]      = useState(PAGOS_0);
  const [ingresos,   setIngresos]   = useState(INGRESOS_0);
  const [gastosList, setGastosList] = useState([]);
  const [historial,  setHistorial]  = useState([]);
  const [importando, setImportando] = useState(false);
  const [importMsg,  setImportMsg]  = useState(null); // {tipo:'ok'|'err', texto, count}

  const handleImportarExcel = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportando(true); setImportMsg(null);
    try {
      const nuevos = await parseGestorExcel(file);
      if (nuevos.length === 0) {
        setImportMsg({tipo:"err", texto:"No se encontraron gastos en el archivo. ¿Es el Excel de Gestor de Gastos?"});
        setImportando(false); return;
      }
      // agregar gastos sin duplicar (por fecha+monto+nombre)
      setGastosList(prev => {
        const existentes = new Set(prev.map(g=>`${g.fecha}|${g.monto}|${g.nombre}`));
        const filtrados = nuevos.filter(g=>!existentes.has(`${g.fecha}|${g.monto}|${g.nombre}`));
        return [...filtrados, ...prev];
      });
      // descontar de bolsillos automáticamente
      setBolsillos(bs => {
        const descuentos = {};
        nuevos.forEach(g => { descuentos[g.bolsilloId] = (descuentos[g.bolsilloId]||0) + g.monto; });
        return bs.map(b => descuentos[b.id] ? {...b, saldo: Math.max(0, b.saldo - descuentos[b.id])} : b);
      });
      setImportMsg({tipo:"ok", texto:`✓ ${nuevos.length} gastos importados correctamente`, count: nuevos.length});
    } catch(err) {
      setImportMsg({tipo:"err", texto:`Error leyendo el archivo: ${err.message}`});
    }
    setImportando(false);
    e.target.value = "";
  }, []);

  // consejera
  const [montoInput, setMontoInput] = useState("");
  const [fuenteSel,  setFuenteSel]  = useState("");
  const [plan,       setPlan]       = useState(null);
  const [planOk,     setPlanOk]     = useState(false);

  // modales
  const [modal, setModal] = useState(null); // null | {tipo, data}
  const closeModal = () => setModal(null);

  // forms reutilizables
  const [formBolsillo, setFormBolsillo] = useState({nombre:"",icon:"💰",color:C.teal,grupo:"vida",meta:"",metaMensual:"",descripcion:""});
  const [formPago,     setFormPago]     = useState({nombre:"",monto:"",diaPago:"",bolsilloId:"vida",categoria:"otros",urgencia:"normal",nota:""});
  const [formIngreso,  setFormIngreso]  = useState({fuente:"",monto:"",diaEsp:""});
  const [formGasto,    setFormGasto]    = useState({nombre:"",monto:"",categoria:"comida",bolsilloId:"",fecha:today.toISOString().slice(0,10)});

  const alertas = useMemo(()=>generarAlertas(gastosList,bolsillos,pagos),[gastosList,bolsillos,pagos]);
  const totalBolsillos = bolsillos.reduce((s,b)=>s+b.saldo,0);
  const totalPorEntrar = ingresos.filter(i=>!i.recibido).reduce((s,i)=>s+i.monto,0);
  const gastosMes = useMemo(()=>{
    const m={};
    gastosList.filter(g=>{const d=new Date(g.fecha);return d.getMonth()===today.getMonth()&&d.getFullYear()===today.getFullYear();}).forEach(g=>{m[g.categoria]=(m[g.categoria]||0)+g.monto;});
    return m;
  },[gastosList]);

  // ── acciones ──
  const hacerPlan = ()=>{if(!montoInput||+montoInput<=0)return;setPlan(generarPlan(+montoInput,bolsillos,pagos));setPlanOk(false);};
  const aplicarPlan = ()=>{
    if(!plan) return;
    setBolsillos(bs=>bs.map(b=>{const sum=plan.filter(p=>p.bolsilloId===b.id).reduce((s,p)=>s+p.monto,0);return sum>0?{...b,saldo:b.saldo+sum}:b;}));
    if(fuenteSel) setIngresos(arr=>arr.map(i=>i.fuente===fuenteSel?{...i,recibido:true}:i));
    setHistorial(h=>[{id:uid(),fuente:fuenteSel||"Ingreso",monto:+montoInput,plan,fecha:today.toLocaleDateString("es-CO")},...h]);
    setPlanOk(true);
  };

  const guardarBolsillo = ()=>{
    if(!formBolsillo.nombre||!formBolsillo.meta) return;
    if(modal?.data?.id){
      setBolsillos(bs=>bs.map(b=>b.id===modal.data.id?{...b,...formBolsillo,meta:+formBolsillo.meta,metaMensual:+formBolsillo.metaMensual}:b));
    } else {
      setBolsillos(bs=>[...bs,{...formBolsillo,id:uid(),saldo:0,meta:+formBolsillo.meta,metaMensual:+formBolsillo.metaMensual||0}]);
    }
    closeModal();
  };

  const guardarPago = ()=>{
    if(!formPago.nombre||!formPago.monto) return;
    if(modal?.data?.id){
      setPagos(ps=>ps.map(p=>p.id===modal.data.id?{...p,...formPago,monto:+formPago.monto,diaPago:+formPago.diaPago}:p));
    } else {
      const maxPrioridad = pagos.reduce((m,p)=>Math.max(m,p.prioridad),0);
      setPagos(ps=>[...ps,{...formPago,id:uid(),monto:+formPago.monto,diaPago:+formPago.diaPago,pagado:false,prioridad:maxPrioridad+1}]);
    }
    closeModal();
  };

  const guardarIngreso = ()=>{
    if(!formIngreso.fuente||!formIngreso.monto) return;
    if(modal?.data?.id){
      setIngresos(arr=>arr.map(i=>i.id===modal.data.id?{...i,...formIngreso,monto:+formIngreso.monto,diaEsp:+formIngreso.diaEsp}:i));
    } else {
      setIngresos(arr=>[...arr,{...formIngreso,id:uid(),monto:+formIngreso.monto,diaEsp:+formIngreso.diaEsp||15,recibido:false}]);
    }
    closeModal();
  };

  const registrarGasto = ()=>{
    if(!formGasto.nombre||!formGasto.monto) return;
    const bid=formGasto.bolsilloId||CATS[formGasto.categoria]?.bolsilloId||"vida";
    const g={...formGasto,id:uid(),monto:+formGasto.monto,bolsilloId:bid};
    setGastosList(l=>[g,...l]);
    setBolsillos(bs=>bs.map(b=>b.id===bid?{...b,saldo:Math.max(0,b.saldo-g.monto)}:b));
    setFormGasto({nombre:"",monto:"",categoria:"comida",bolsilloId:"",fecha:today.toISOString().slice(0,10)});
    closeModal();
  };

  const abrirEditarBolsillo = b=>{
    setFormBolsillo({nombre:b.nombre,icon:b.icon,color:b.color,grupo:b.grupo,meta:b.meta,metaMensual:b.metaMensual,descripcion:b.descripcion});
    setModal({tipo:"bolsillo",data:b});
  };
  const abrirEditarPago = p=>{
    setFormPago({nombre:p.nombre,monto:p.monto,diaPago:p.diaPago,bolsilloId:p.bolsilloId,categoria:p.categoria,urgencia:p.urgencia||"normal",nota:p.nota||""});
    setModal({tipo:"pago",data:p});
  };
  const abrirEditarIngreso = i=>{
    setFormIngreso({fuente:i.fuente,monto:i.monto,diaEsp:i.diaEsp});
    setModal({tipo:"ingreso",data:i});
  };

  // ── grupos de bolsillos ──
  const grupos = [{key:"sagrado",label:"🔒 Sagrados",color:C.blue},{key:"vida",label:"☀️ Vida & gastos",color:C.accent},{key:"provision",label:"⚡ Provisiones",color:C.amber}];

  return (
    <div style={{minHeight:"100vh",background:C.bg,color:C.text,fontFamily:"'Sora','Segoe UI',sans-serif"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        input,select,textarea{font-family:inherit}
        input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
        select option{background:${C.card};color:${C.text}}
        ::-webkit-scrollbar{width:3px}::-webkit-scrollbar-thumb{background:#ffffff15;border-radius:99px}
        @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        .fu{animation:fadeUp .3s ease}
        button:hover{opacity:.82}
      `}</style>

      {/* ── HEADER ── */}
      <div style={{background:C.surface,borderBottom:`1px solid ${C.border}`,padding:"16px 18px 0"}}>
        <div style={{maxWidth:720,margin:"0 auto"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
            <div>
              <div style={{fontSize:9,letterSpacing:3,color:C.accent,textTransform:"uppercase",marginBottom:4}}>✦ Consejera Financiera</div>
              <div style={{fontSize:20,fontWeight:700,letterSpacing:-.5}}>Plata Inteligente 💜</div>
              <div style={{fontSize:11,color:C.muted,marginTop:2}}>{today.toLocaleDateString("es-CO",{weekday:"long",day:"numeric",month:"long"})}</div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:9,color:C.muted,marginBottom:2}}>En bolsillos</div>
              <div style={{fontSize:18,fontWeight:700,color:C.teal}}>{fmtShort(totalBolsillos)}</div>
              <div style={{fontSize:10,color:C.blue,marginTop:1}}>↓ {fmtShort(totalPorEntrar)} por entrar</div>
            </div>
          </div>
          {alertas.length>0&&(
            <div style={{background:C.red+"10",border:`1px solid ${C.red}25`,borderRadius:10,padding:"7px 13px",marginBottom:12}}>
              {alertas.slice(0,2).map((a,i)=><div key={i} style={{color:a.color,fontSize:12,display:"flex",gap:6,marginBottom:i<1&&alertas.length>1?3:0}}><span>{a.icon}</span><span>{a.msg}</span></div>)}
            </div>
          )}
          <div style={{display:"flex",overflowX:"auto",gap:0}}>
            {[["consejera","💡 Consejera"],["bolsillos","👜 Bolsillos"],["metas","🎯 Metas"],["gastos","💸 Gastos"],["pagos","⚡ Pagos"],["ingresos","💰 Ingresos"],["config","⚙️ Editar"]].map(([k,l])=>(
              <button key={k} onClick={()=>setTab(k)} style={S.tab(tab===k)}>{l}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{maxWidth:720,margin:"0 auto",padding:"18px 16px 100px"}}>

        {/* ════ CONSEJERA ════ */}
        {tab==="consejera"&&(
          <div className="fu">
            <div style={{...S.card,background:`linear-gradient(135deg,${C.accent}12,${C.card})`,border:`1px solid ${C.accent}28`}}>
              <div style={{fontSize:10,color:C.accent,letterSpacing:2,textTransform:"uppercase",marginBottom:10}}>¿Te entró plata? Dime cuánto 💜</div>
              <div style={{fontSize:12,color:C.muted,marginBottom:14,lineHeight:1.7}}>Ingresa el monto, dime de quién, y te digo exactamente cómo distribuirlo.</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                <div>
                  <div style={S.lbl}>Monto recibido</div>
                  <input type="number" placeholder="Ej: 1000000" value={montoInput} onChange={e=>{setMontoInput(e.target.value);setPlan(null);setPlanOk(false);}} style={{...S.inp,fontSize:16,fontWeight:600,color:C.accent}}/>
                </div>
                <div>
                  <div style={S.lbl}>¿De quién?</div>
                  <select value={fuenteSel} onChange={e=>setFuenteSel(e.target.value)} style={{...S.inp,background:C.card}}>
                    <option value="">Seleccionar...</option>
                    {ingresos.map(i=><option key={i.id} value={i.fuente}>{i.fuente} — {fmtShort(i.monto)}</option>)}
                    <option value="otro">Otro ingreso</option>
                  </select>
                </div>
              </div>
              <button onClick={hacerPlan} disabled={!montoInput||+montoInput<=0} style={{...S.btn(montoInput&&+montoInput>0?C.accent:"#ffffff12",montoInput&&+montoInput>0?"#0B0B12":C.muted),width:"100%",padding:"12px",fontSize:14}}>Ver distribución de {montoInput?fmt(+montoInput):""}</button>
            </div>

            {plan&&(
              <div>
                <div style={S.sec}>Plan para {fmt(+montoInput)}</div>
                {plan.map((p,i)=>(
                  <div key={i} className="fu" style={{
                    ...S.card,
                    borderLeft:`3px solid ${p.informativo?p.color+"40":p.color}`,
                    borderRadius:"0 14px 14px 0",
                    opacity:p.informativo?0.6:1,
                    padding:p.informativo?"9px 14px":"14px 16px",
                  }}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div style={{flex:1}}>
                        <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:p.informativo?0:3}}>
                          <span style={{fontSize:p.informativo?13:15}}>{p.icon}</span>
                          <span style={{fontSize:p.informativo?12:13,fontWeight:p.informativo?500:600,color:p.informativo?C.muted:C.text}}>{p.nombre}</span>
                          {!p.informativo&&<Badge text={p.tipo.toUpperCase()} color={p.color}/>}
                        </div>
                        {!p.informativo&&<div style={{fontSize:11,color:C.muted,lineHeight:1.5,marginTop:3}}>{p.razon}</div>}
                        {p.informativo&&<div style={{fontSize:10,color:C.muted,marginLeft:20}}>{p.razon}</div>}
                      </div>
                      {!p.informativo&&<div style={{fontSize:16,fontWeight:700,color:p.color,marginLeft:12,whiteSpace:"nowrap"}}>{fmt(p.monto)}</div>}
                    </div>
                  </div>
                ))}
                <div style={{...S.card,background:planOk?C.teal+"10":C.accent+"0e",border:`1px solid ${planOk?C.teal:C.accent}28`}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:planOk?0:10}}>
                    <div>
                      <div style={{fontSize:13,fontWeight:600,color:planOk?C.teal:C.text}}>{planOk?"✓ ¡Aplicado! Bolsillos actualizados":"¿Aplicar esta distribución?"}</div>
                      {!planOk&&<div style={{fontSize:11,color:C.muted,marginTop:3}}>Los bolsillos se actualizarán automáticamente.</div>}
                    </div>
                    {planOk&&<span style={{fontSize:22}}>🎉</span>}
                  </div>
                  {!planOk&&<button onClick={aplicarPlan} style={{...S.btn(C.accent),width:"100%",padding:"11px",fontSize:13}}>Aplicar distribución</button>}
                </div>
              </div>
            )}

            {historial.length>0&&(
              <div style={{marginTop:20}}>
                <div style={S.sec}>Historial</div>
                {historial.slice(0,4).map(h=>(
                  <div key={h.id} style={{...S.card,padding:"11px 14px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
                      <span style={{fontSize:13,fontWeight:600}}>{h.fuente}</span>
                      <span style={{fontSize:13,fontWeight:700,color:C.teal}}>{fmt(h.monto)}</span>
                    </div>
                    <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                      {(h.plan||[]).map((p,i)=><span key={i} style={{fontSize:10,background:p.color+"18",color:p.color,borderRadius:5,padding:"1px 7px"}}>{p.icon} {fmtShort(p.monto)}</span>)}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* ════ BOLSILLOS ════ */}
        {tab==="bolsillos"&&(
          <div className="fu">
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:14}}>
              {[["Total",totalBolsillos,C.teal],["Sueños",bolsillos.filter(b=>b.grupo==="sagrado").reduce((s,b)=>s+b.saldo,0),C.blue],["Provisiones",bolsillos.filter(b=>b.grupo==="provision").reduce((s,b)=>s+b.saldo,0),C.amber]].map(([l,v,c])=>(
                <div key={l} style={{...S.card,padding:"11px 13px"}}><div style={{fontSize:9,color:C.muted,marginBottom:2}}>{l}</div><div style={{fontSize:14,fontWeight:700,color:c}}>{fmtShort(v)}</div></div>
              ))}
            </div>
            {grupos.map(gr=>(
              <div key={gr.key}>
                <div style={{...S.sec,color:gr.color}}>{gr.label}</div>
                {bolsillos.filter(b=>b.grupo===gr.key).map(b=>{
                  const gastadoMes=gastosList.filter(g=>g.bolsilloId===b.id&&new Date(g.fecha).getMonth()===today.getMonth()).reduce((s,g)=>s+g.monto,0);
                  return (
                    <div key={b.id} style={S.card}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                        <div style={{flex:1}}>
                          <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3,flexWrap:"wrap"}}>
                            <span style={{fontSize:16}}>{b.icon}</span>
                            <span style={{fontSize:13,fontWeight:600}}>{b.nombre}</span>
                            {b.grupo==="sagrado"&&<Badge text="NO TOCAR" color={C.muted}/>}
                            {b.metaMensual>0&&<Badge text={`${fmtShort(b.metaMensual)}/mes`} color={b.color}/>}
                          </div>
                          <div style={{fontSize:11,color:C.muted,marginBottom:4}}>{b.descripcion}</div>
                          <ProgressBar saldo={b.saldo} meta={b.meta} color={b.color}/>
                          <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:C.muted,marginTop:4}}>
                            <span>{fmt(b.saldo)} / {fmt(b.meta)} ({Math.round(b.saldo/b.meta*100)}%)</span>
                            {gastadoMes>0&&<span style={{color:C.orange}}>Gasto mes: {fmt(gastadoMes)}</span>}
                          </div>
                        </div>
                        <div style={{textAlign:"right",marginLeft:12}}>
                          <div style={{fontSize:16,fontWeight:700,color:b.color}}>{fmt(b.saldo)}</div>
                          <div style={{display:"flex",gap:5,marginTop:8,justifyContent:"flex-end"}}>
                            <button onClick={()=>{const m=prompt(`+ a ${b.nombre}:`);if(m&&+m>0)setBolsillos(bs=>bs.map(x=>x.id===b.id?{...x,saldo:x.saldo+ +m}:x));}} style={{background:b.color+"20",border:`1px solid ${b.color}30`,borderRadius:8,color:b.color,fontSize:11,padding:"3px 9px",cursor:"pointer"}}>+</button>
                            {b.grupo!=="sagrado"&&<button onClick={()=>{const m=prompt(`− de ${b.nombre}:`);if(m&&+m>0)setBolsillos(bs=>bs.map(x=>x.id===b.id?{...x,saldo:Math.max(0,x.saldo- +m)}:x));}} style={{background:"#ffffff0a",border:`1px solid ${C.border}`,borderRadius:8,color:C.muted,fontSize:11,padding:"3px 9px",cursor:"pointer"}}>−</button>}
                            <button onClick={()=>abrirEditarBolsillo(b)} style={{background:"#ffffff0a",border:`1px solid ${C.border}`,borderRadius:8,color:C.muted,fontSize:11,padding:"3px 9px",cursor:"pointer"}}>✏️</button>
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            ))}
            <button onClick={()=>{setFormBolsillo({nombre:"",icon:"💰",color:C.teal,grupo:"vida",meta:"",metaMensual:"",descripcion:""});setModal({tipo:"bolsillo",data:null});}} style={{width:"100%",background:"transparent",border:`1px dashed ${C.border}`,borderRadius:14,padding:"12px",color:C.muted,fontSize:13,cursor:"pointer",marginTop:6}}>+ Nuevo bolsillo</button>
          </div>
        )}

        {/* ════ GASTOS ════ */}
        {tab==="gastos"&&(
          <div className="fu">
            {/* ── IMPORTAR DESDE GESTOR DE GASTOS ── */}
            <div style={{...S.card,background:`linear-gradient(135deg,${C.teal}10,${C.card})`,border:`1px solid ${C.teal}30`,marginBottom:14}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div>
                  <div style={{fontSize:13,fontWeight:600,color:C.teal}}>📥 Importar desde Gestor de Gastos</div>
                  <div style={{fontSize:11,color:C.muted,marginTop:2}}>Exporta el Excel desde tu app y súbelo aquí — se mapean solos a los bolsillos.</div>
                </div>
                <label style={{background:importando?C.muted:C.teal,borderRadius:10,padding:"8px 14px",cursor:importando?"default":"pointer",fontSize:12,fontWeight:700,color:"#0B0B12",whiteSpace:"nowrap",flexShrink:0}}>
                  {importando?"Cargando...":"Subir Excel"}
                  <input type="file" accept=".xlsx,.xls" onChange={handleImportarExcel} style={{display:"none"}} disabled={importando}/>
                </label>
              </div>
              {importMsg&&(
                <div style={{background:importMsg.tipo==="ok"?C.teal+"15":C.red+"15",border:`1px solid ${importMsg.tipo==="ok"?C.teal:C.red}30`,borderRadius:9,padding:"8px 12px",fontSize:12,color:importMsg.tipo==="ok"?C.teal:C.red}}>
                  {importMsg.texto}
                  {importMsg.tipo==="ok"&&<button onClick={()=>setImportMsg(null)} style={{background:"none",border:"none",color:C.muted,fontSize:11,cursor:"pointer",marginLeft:8}}>✕</button>}
                </div>
              )}
              <div style={{marginTop:8,display:"flex",gap:6,flexWrap:"wrap"}}>
                {Object.entries(CAT_MAP).map(([k,v])=>(
                  <span key={k} style={{fontSize:9,background:CATS[v]?.color+"18",color:CATS[v]?.color,borderRadius:5,padding:"2px 7px"}}>{CATS[v]?.icon} {k}</span>
                ))}
              </div>
            </div>

            {alertas.length>0&&alertas.map((a,i)=>(
              <div key={i} style={{background:a.color+"10",border:`1px solid ${a.color}28`,borderRadius:12,padding:"8px 13px",marginBottom:7,fontSize:12,color:a.color,display:"flex",gap:8}}><span>{a.icon}</span><span>{a.msg}</span></div>
            ))}
            <div style={S.card}>
              <div style={{...S.sec,margin:"0 0 10px"}}>Gastos este mes — {fmt(Object.values(gastosMes).reduce((s,v)=>s+v,0))}</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(120px,1fr))",gap:8}}>
                {Object.entries(gastosMes).map(([cat,total])=>{
                  const c=CATS[cat]; if(!c) return null;
                  const b=bolsillos.find(x=>x.id===c.bolsilloId);
                  const pct=b?Math.min(total/b.meta*100,100):0;
                  return <div key={cat} style={{background:"#ffffff06",borderRadius:11,padding:"10px 12px"}}>
                    <div style={{fontSize:14,marginBottom:2}}>{c.icon}</div>
                    <div style={{fontSize:10,color:C.muted,marginBottom:2}}>{c.label}</div>
                    <div style={{fontSize:12,fontWeight:600,color:pct>80?C.red:C.text}}>{fmt(total)}</div>
                    {b&&<div style={{background:"#ffffff0c",borderRadius:999,height:3,overflow:"hidden",marginTop:4}}><div style={{width:`${pct}%`,height:"100%",background:pct>80?C.red:c.color}}/></div>}
                  </div>;
                })}
                {Object.keys(gastosMes).length===0&&<div style={{fontSize:12,color:C.muted,gridColumn:"1/-1",padding:"8px 0"}}>Sin gastos este mes aún.</div>}
              </div>
            </div>

            <button onClick={()=>{setFormGasto({nombre:"",monto:"",categoria:"comida",bolsilloId:"",fecha:today.toISOString().slice(0,10)});setModal({tipo:"gasto",data:null});}} style={{...S.btn(C.accent),width:"100%",padding:"12px",fontSize:13,marginBottom:12}}>+ Registrar gasto</button>

            {gastosList.length>0&&(
              <>
                <div style={S.sec}>Todos los gastos</div>
                {gastosList.slice(0,30).map(g=>{
                  const c=CATS[g.categoria]||CATS.otros;
                  const b=bolsillos.find(x=>x.id===g.bolsilloId);
                  return <div key={g.id} style={{...S.card,padding:"11px 14px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div style={{flex:1}}>
                        <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:2}}><span style={{fontSize:14}}>{c.icon}</span><span style={{fontSize:13,fontWeight:500}}>{g.nombre}</span></div>
                        <div style={{fontSize:10,color:C.muted,display:"flex",gap:8}}>
                          <span style={{color:c.color}}>{c.label}</span>
                          {b&&<span>📂 {b.icon} {b.nombre}</span>}
                          <span>{new Date(g.fecha).toLocaleDateString("es-CO",{day:"numeric",month:"short"})}</span>
                        </div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:14,fontWeight:700,color:c.color}}>{fmt(g.monto)}</span>
                        <button onClick={()=>{const b2=bolsillos.find(x=>x.id===g.bolsilloId);if(b2)setBolsillos(bs=>bs.map(x=>x.id===b2.id?{...x,saldo:x.saldo+g.monto}:x));setGastosList(l=>l.filter(x=>x.id!==g.id));}} style={{background:"#ffffff0a",border:"none",borderRadius:7,color:C.muted,fontSize:11,padding:"3px 7px",cursor:"pointer"}}>🗑️</button>
                      </div>
                    </div>
                  </div>;
                })}
              </>
            )}
          </div>
        )}

        {/* ════ PAGOS ════ */}
        {tab==="pagos"&&(
          <div className="fu">
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
              <div style={S.card}><div style={{fontSize:9,color:C.muted,marginBottom:2}}>Pendiente</div><div style={{fontSize:16,fontWeight:700,color:C.red}}>{fmt(pagos.filter(p=>!p.pagado).reduce((s,p)=>s+p.monto,0))}</div></div>
              <div style={S.card}><div style={{fontSize:9,color:C.muted,marginBottom:2}}>Pagado</div><div style={{fontSize:16,fontWeight:700,color:C.teal}}>{fmt(pagos.filter(p=>p.pagado).reduce((s,p)=>s+p.monto,0))}</div></div>
            </div>

            {pagos.filter(p=>!p.pagado).sort((a,b)=>daysUntil(a.diaPago)-daysUntil(b.diaPago)).map(p=>{
              const days=daysUntil(p.diaPago),col=days<=2?C.red:days<=7?C.amber:C.blue;
              const b=bolsillos.find(x=>x.id===p.bolsilloId),provisto=b?Math.min(b.saldo,p.monto):0,falta=Math.max(p.monto-provisto,0);
              const urg=URGENCIA[p.urgencia||"normal"];
              return <div key={p.id} style={{...S.card,borderLeft:`3px solid ${col}`,borderRadius:"0 14px 14px 0"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:4,flexWrap:"wrap"}}>
                      <span style={{fontSize:13,fontWeight:600}}>{p.nombre}</span>
                      <Badge text={days<=1?"¡HOY!":`día ${p.diaPago} · ${days}d`} color={col}/>
                      <Badge text={urg.label} color={urg.color}/>
                    </div>
                    <div style={{background:"#ffffff0a",borderRadius:999,height:4,overflow:"hidden",marginBottom:4}}><div style={{width:`${Math.min(provisto/p.monto*100,100)}%`,height:"100%",background:col}}/></div>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:C.muted}}>
                      <span>{b&&`${b.icon} ${b.nombre}: ${fmt(provisto)}`}</span>
                      {falta>0&&<span style={{color:C.red}}>Falta: {fmt(falta)}</span>}
                    </div>
                    {p.nota&&<div style={{fontSize:10,color:C.muted,marginTop:3,fontStyle:"italic"}}>{p.nota}</div>}
                  </div>
                  <div style={{textAlign:"right",marginLeft:12}}>
                    <div style={{fontSize:15,fontWeight:700,color:col}}>{fmt(p.monto)}</div>
                    <div style={{display:"flex",gap:5,marginTop:7,justifyContent:"flex-end"}}>
                      <button onClick={()=>abrirEditarPago(p)} style={{background:"#ffffff0a",border:"none",borderRadius:8,color:C.muted,fontSize:11,padding:"3px 8px",cursor:"pointer"}}>✏️</button>
                      <button onClick={()=>setPagos(ps=>ps.map(x=>x.id===p.id?{...x,pagado:true}:x))} style={{background:col+"20",border:"none",borderRadius:8,color:col,fontSize:11,padding:"3px 10px",cursor:"pointer"}}>Pagado ✓</button>
                    </div>
                  </div>
                </div>
              </div>;
            })}

            {pagos.filter(p=>p.pagado).length>0&&<>
              <div style={S.sec}>Pagados ✓</div>
              {pagos.filter(p=>p.pagado).map(p=>(
                <div key={p.id} style={{...S.card,opacity:.45,padding:"10px 14px",display:"flex",justifyContent:"space-between"}}>
                  <span style={{fontSize:12,textDecoration:"line-through",color:C.muted}}>{p.nombre}</span>
                  <span style={{fontSize:12,color:C.teal,fontWeight:600}}>{fmt(p.monto)}</span>
                </div>
              ))}
            </>}

            <button onClick={()=>{setFormPago({nombre:"",monto:"",diaPago:"",bolsilloId:"vida",categoria:"otros"});setModal({tipo:"pago",data:null});}} style={{width:"100%",background:"transparent",border:`1px dashed ${C.border}`,borderRadius:14,padding:"11px",color:C.muted,fontSize:13,cursor:"pointer",marginTop:6}}>+ Agregar pago</button>
          </div>
        )}

        {/* ════ INGRESOS ════ */}
        {tab==="ingresos"&&(
          <div className="fu">
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div style={S.sec}>Por recibir este mes</div>
              <div style={{fontSize:14,fontWeight:700,color:C.blue}}>{fmt(totalPorEntrar)}</div>
            </div>
            {ingresos.map(i=>{
              const days=daysUntil(i.diaEsp),col=i.recibido?C.teal:days<=3?C.red:days<=7?C.amber:C.blue;
              return <div key={i.id} style={{...S.card,opacity:i.recibido?.55:1}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3}}>
                      <span style={{fontSize:13,fontWeight:600,textDecoration:i.recibido?"line-through":"none"}}>{i.fuente}</span>
                      <Badge text={i.recibido?"✓ Recibido":`día ${i.diaEsp}`} color={col}/>
                    </div>
                    {!i.recibido&&<div style={{fontSize:11,color:C.muted}}>En {days} días · {fmt(i.monto)}</div>}
                  </div>
                  <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
                    <div style={{fontSize:15,fontWeight:700,color:col}}>{fmt(i.monto)}</div>
                    <div style={{display:"flex",gap:5}}>
                      <button onClick={()=>abrirEditarIngreso(i)} style={{background:"#ffffff0a",border:"none",borderRadius:8,color:C.muted,fontSize:11,padding:"3px 8px",cursor:"pointer"}}>✏️</button>
                      {!i.recibido&&<button onClick={()=>{setMontoInput(String(i.monto));setFuenteSel(i.fuente);setTab("consejera");}} style={{background:C.accent+"20",border:`1px solid ${C.accent}35`,borderRadius:8,color:C.accent,fontSize:11,padding:"3px 10px",cursor:"pointer"}}>Distribuir ↗</button>}
                    </div>
                  </div>
                </div>
              </div>;
            })}
            <button onClick={()=>{setFormIngreso({fuente:"",monto:"",diaEsp:""});setModal({tipo:"ingreso",data:null});}} style={{width:"100%",background:"transparent",border:`1px dashed ${C.border}`,borderRadius:14,padding:"11px",color:C.muted,fontSize:13,cursor:"pointer",marginTop:4}}>+ Agregar ingreso esperado</button>
          </div>
        )}

        {/* ════ ⚙️ EDITAR / CONFIG ════ */}
        {tab==="metas"&&(
          <div className="fu">
            {/* TC — deuda real */}
            <div style={{...S.card,background:"linear-gradient(135deg,#2D0A0A,#1A0808)",border:`1px solid ${C.red}35`,marginBottom:12}}>
              <div style={{fontSize:10,color:C.red,letterSpacing:2,textTransform:"uppercase",marginBottom:10}}>💳 Deuda TC — Plan de liquidación</div>
              {(()=>{
                const b=bolsillos.find(x=>x.id==="tc");
                const deuda=b?.deudaTotal||5900000, minimo=b?.minimoTC||1500000, extra=b?.abonoExtra||600000;
                const abonoReal=minimo+extra;
                const mesesPagandoMinimo=Math.ceil(deuda/(minimo*0.7));
                const mesesPagandoExtra=Math.ceil(deuda/(abonoReal*0.7));
                return <>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                    {[["Deuda total",fmt(deuda),C.red],["Mínimo mensual",fmt(minimo),C.amber],["Abono extra sugerido",fmt(extra),C.orange],["Pago ideal/mes",fmt(abonoReal),C.teal]].map(([l,v,c])=>(
                      <div key={l} style={{background:"#ffffff06",borderRadius:10,padding:"10px 12px"}}>
                        <div style={{fontSize:10,color:C.muted,marginBottom:2}}>{l}</div>
                        <div style={{fontSize:14,fontWeight:700,color:c}}>{v}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{background:"#ffffff06",borderRadius:10,padding:"10px 14px",fontSize:12,color:C.muted,lineHeight:1.7}}>
                    Solo mínimo: liquidarías en ~<b style={{color:C.red}}>{mesesPagandoMinimo} meses</b> pero pagarás mucho en intereses.<br/>
                    Con abono extra ({fmt(abonoReal)}/mes): ~<b style={{color:C.teal}}>{mesesPagandoExtra} meses</b> y ahorras intereses.
                  </div>
                  <ProgressBar saldo={b?.saldo||0} meta={abonoReal} color={C.amber} height={6}/>
                  <div style={{fontSize:10,color:C.muted,marginTop:4}}>Provisionado este mes: {fmt(b?.saldo||0)} / {fmt(abonoReal)}</div>
                </>;
              })()}
            </div>

            {/* IDC — Instructora de Buceo */}
            <div style={{...S.card,background:"linear-gradient(135deg,#0A1A2D,#081525)",border:`1px solid ${C.blue}35`,marginBottom:12}}>
              <div style={{fontSize:10,color:C.blue,letterSpacing:2,textTransform:"uppercase",marginBottom:10}}>🤿 IDC — Instructora de Buceo</div>
              {(()=>{
                const b=bolsillos.find(x=>x.id==="buceo");
                if(!b) return null;
                const fase1=b.metaFase1||3200000, metaTotal=b.meta||11000000;
                const mesesR=mesesHasta(b.fechaFase1||"ago/sep 2025");
                const cuotaMes=Math.ceil(Math.max(fase1-b.saldo,0)/mesesR);
                const pctFase1=Math.min(b.saldo/fase1*100,100);
                return <>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                    {[["Tienes ahorrado",fmt(b.saldo),C.blue],["Fase 1 — arrancar",fmt(fase1),C.accent],["Meta total IDC",fmt(metaTotal),C.teal],["Cuota ideal/mes",fmt(cuotaMes),C.blue]].map(([l,v,c])=>(
                      <div key={l} style={{background:"#ffffff06",borderRadius:10,padding:"10px 12px"}}>
                        <div style={{fontSize:10,color:C.muted,marginBottom:2}}>{l}</div>
                        <div style={{fontSize:14,fontWeight:700,color:c}}>{v}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{marginBottom:8}}>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.muted,marginBottom:4}}>
                      <span>Fase 1 — {b.fechaFase1}</span>
                      <span style={{color:C.blue}}>{Math.round(pctFase1)}% ({fmt(b.saldo)} / {fmt(fase1)})</span>
                    </div>
                    <div style={{background:"#ffffff0e",borderRadius:999,height:8,overflow:"hidden"}}>
                      <div style={{width:`${pctFase1}%`,height:"100%",background:`linear-gradient(90deg,${C.blue},${C.accent})`,borderRadius:999,transition:"width .5s"}}/>
                    </div>
                  </div>
                  <div style={{marginBottom:8}}>
                    <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.muted,marginBottom:4}}>
                      <span>Meta total IDC</span>
                      <span style={{color:C.teal}}>{Math.round(b.saldo/metaTotal*100)}% ({fmt(b.saldo)} / {fmt(metaTotal)})</span>
                    </div>
                    <div style={{background:"#ffffff0e",borderRadius:999,height:5,overflow:"hidden"}}>
                      <div style={{width:`${Math.min(b.saldo/metaTotal*100,100)}%`,height:"100%",background:C.teal,borderRadius:999}}/>
                    </div>
                  </div>
                  <div style={{background:C.blue+"10",border:`1px solid ${C.blue}20`,borderRadius:10,padding:"9px 13px",fontSize:12,color:C.blue,lineHeight:1.6}}>
                    Ahorrando <b>{fmt(cuotaMes)}/mes</b> llegas a la Fase 1 ({fmt(fase1)}) en <b>{mesesR} mes{mesesR!==1?"es":""}</b>. ¡Tú puedes! 🤿
                  </div>
                </>;
              })()}
            </div>

            {/* Resumen de sueños */}
            <div style={S.sec}>🌟 Todos los objetivos</div>
            {bolsillos.filter(b=>b.grupo==="sagrado"||["capcut","mama"].includes(b.id)).map(b=>(
              <div key={b.id} style={S.card}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3}}>
                      <span style={{fontSize:16}}>{b.icon}</span>
                      <span style={{fontSize:13,fontWeight:600}}>{b.nombre}</span>
                      {b.metaFase1&&<Badge text="Tiene fases" color={C.blue}/>}
                    </div>
                    <div style={{fontSize:11,color:C.muted,marginBottom:4}}>{b.descripcion}</div>
                    <ProgressBar saldo={b.saldo} meta={b.metaFase1||b.meta} color={b.color}/>
                    <div style={{fontSize:10,color:C.muted,marginTop:3}}>{fmt(b.saldo)} / {fmt(b.metaFase1||b.meta)} — {Math.round(b.saldo/(b.metaFase1||b.meta)*100)}%</div>
                  </div>
                  <div style={{textAlign:"right",marginLeft:12}}>
                    <div style={{fontSize:15,fontWeight:700,color:b.color}}>{fmt(b.saldo)}</div>
                    {b.metaMensual>0&&<div style={{fontSize:10,color:C.muted,marginTop:3}}>{fmt(b.metaMensual)}/mes ideal</div>}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {tab==="config"&&(
          <div className="fu">
            <div style={{...S.card,background:`linear-gradient(135deg,${C.accent}10,${C.card})`,border:`1px solid ${C.accent}25`,marginBottom:16}}>
              <div style={{fontSize:12,color:C.accent,fontWeight:600,marginBottom:6}}>⚙️ Panel de configuración</div>
              <div style={{fontSize:12,color:C.muted,lineHeight:1.7}}>Aquí puedes editar todo sin ayuda externa — bolsillos, pagos e ingresos. Toca el ítem que quieres cambiar.</div>
            </div>

            {/* Bolsillos */}
            <div style={S.sec}>👜 Bolsillos — editar</div>
            {bolsillos.map(b=>(
              <div key={b.id} style={{...S.card,padding:"11px 14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <div style={{width:8,height:8,borderRadius:"50%",background:b.color,flexShrink:0}}/>
                    <span style={{fontSize:13}}>{b.icon} {b.nombre}</span>
                    <Badge text={b.grupo} color={b.color}/>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:12,fontWeight:600,color:b.color}}>{fmt(b.saldo)}</span>
                    <button onClick={()=>abrirEditarBolsillo(b)} style={{background:C.accent+"20",border:"none",borderRadius:8,color:C.accent,fontSize:11,padding:"4px 10px",cursor:"pointer"}}>Editar</button>
                    <button onClick={()=>{if(window.confirm(`¿Eliminar "${b.nombre}"?`))setBolsillos(bs=>bs.filter(x=>x.id!==b.id));}} style={{background:C.red+"15",border:"none",borderRadius:8,color:C.red,fontSize:11,padding:"4px 10px",cursor:"pointer"}}>🗑️</button>
                  </div>
                </div>
              </div>
            ))}
            <button onClick={()=>{setFormBolsillo({nombre:"",icon:"💰",color:C.teal,grupo:"vida",meta:"",metaMensual:"",descripcion:""});setModal({tipo:"bolsillo",data:null});}} style={{width:"100%",background:"transparent",border:`1px dashed ${C.border}`,borderRadius:12,padding:"10px",color:C.muted,fontSize:13,cursor:"pointer",marginBottom:16}}>+ Nuevo bolsillo</button>

            {/* Pagos */}
            <div style={S.sec}>⚡ Pagos — prioridades y edición</div>
            <div style={{...S.card,background:C.surface,marginBottom:10,padding:"10px 14px"}}>
              <div style={{fontSize:11,color:C.muted,lineHeight:1.7}}>
                Arrastra el número de prioridad o usa ↑↓ para cambiar el orden. La consejera respetará este orden dentro de cada nivel de urgencia.
              </div>
            </div>
            {pagos.sort((a,b)=>a.prioridad-b.prioridad).map((p,idx)=>{
              const urg=URGENCIA[p.urgencia||"normal"];
              return (
                <div key={p.id} style={{...S.card,padding:"10px 14px",borderLeft:`3px solid ${urg.color}`,borderRadius:"0 14px 14px 0"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,flex:1,minWidth:0}}>
                      {/* orden arrows */}
                      <div style={{display:"flex",flexDirection:"column",gap:1}}>
                        <button onClick={()=>{
                          if(idx===0) return;
                          setPagos(ps=>{
                            const sorted=[...ps].sort((a,b)=>a.prioridad-b.prioridad);
                            const prev=sorted[idx-1]; const curr=sorted[idx];
                            return ps.map(x=>x.id===curr.id?{...x,prioridad:prev.prioridad}:x.id===prev.id?{...x,prioridad:curr.prioridad}:x);
                          });
                        }} style={{background:"#ffffff0a",border:"none",borderRadius:4,color:C.muted,fontSize:9,padding:"2px 5px",cursor:"pointer",lineHeight:1}}>▲</button>
                        <button onClick={()=>{
                          if(idx===pagos.length-1) return;
                          setPagos(ps=>{
                            const sorted=[...ps].sort((a,b)=>a.prioridad-b.prioridad);
                            const next=sorted[idx+1]; const curr=sorted[idx];
                            return ps.map(x=>x.id===curr.id?{...x,prioridad:next.prioridad}:x.id===next.id?{...x,prioridad:curr.prioridad}:x);
                          });
                        }} style={{background:"#ffffff0a",border:"none",borderRadius:4,color:C.muted,fontSize:9,padding:"2px 5px",cursor:"pointer",lineHeight:1}}>▼</button>
                      </div>
                      <div style={{width:20,height:20,borderRadius:"50%",background:urg.color+"25",color:urg.color,fontSize:10,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{p.prioridad}</div>
                      <div style={{minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:600,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{p.nombre}</div>
                        <div style={{display:"flex",alignItems:"center",gap:5,marginTop:2,flexWrap:"wrap"}}>
                          <span style={{fontSize:10,background:urg.color+"20",color:urg.color,borderRadius:5,padding:"1px 6px",fontWeight:600}}>{urg.label}</span>
                          <span style={{fontSize:10,color:C.muted}}>día {p.diaPago} · {fmt(p.monto)}</span>
                        </div>
                        {p.nota&&<div style={{fontSize:10,color:C.muted,marginTop:2,fontStyle:"italic"}}>{p.nota}</div>}
                      </div>
                    </div>
                    <div style={{display:"flex",gap:5,flexShrink:0}}>
                      <button onClick={()=>abrirEditarPago(p)} style={{background:C.amber+"18",border:"none",borderRadius:8,color:C.amber,fontSize:11,padding:"4px 9px",cursor:"pointer"}}>Editar</button>
                      <button onClick={()=>{if(window.confirm(`¿Eliminar "${p.nombre}"?`))setPagos(ps=>ps.filter(x=>x.id!==p.id));}} style={{background:C.red+"12",border:"none",borderRadius:8,color:C.red,fontSize:11,padding:"4px 9px",cursor:"pointer"}}>🗑️</button>
                    </div>
                  </div>
                </div>
              );
            })}
            <button onClick={()=>{setFormPago({nombre:"",monto:"",diaPago:"",bolsilloId:"vida",categoria:"otros"});setModal({tipo:"pago",data:null});}} style={{width:"100%",background:"transparent",border:`1px dashed ${C.border}`,borderRadius:12,padding:"10px",color:C.muted,fontSize:13,cursor:"pointer",marginBottom:16}}>+ Nuevo pago</button>

            {/* Ingresos */}
            <div style={S.sec}>💰 Ingresos — editar</div>
            {ingresos.map(i=>(
              <div key={i.id} style={{...S.card,padding:"11px 14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <div style={{fontSize:13,fontWeight:500}}>{i.fuente}</div>
                    <div style={{fontSize:10,color:C.muted}}>Día {i.diaEsp} · {fmt(i.monto)} {i.recibido?"· ✓ recibido":""}</div>
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={()=>abrirEditarIngreso(i)} style={{background:C.blue+"20",border:"none",borderRadius:8,color:C.blue,fontSize:11,padding:"4px 10px",cursor:"pointer"}}>Editar</button>
                    <button onClick={()=>{if(window.confirm(`¿Eliminar "${i.fuente}"?`))setIngresos(arr=>arr.filter(x=>x.id!==i.id));}} style={{background:C.red+"15",border:"none",borderRadius:8,color:C.red,fontSize:11,padding:"4px 10px",cursor:"pointer"}}>🗑️</button>
                  </div>
                </div>
              </div>
            ))}
            <button onClick={()=>{setFormIngreso({fuente:"",monto:"",diaEsp:""});setModal({tipo:"ingreso",data:null});}} style={{width:"100%",background:"transparent",border:`1px dashed ${C.border}`,borderRadius:12,padding:"10px",color:C.muted,fontSize:13,cursor:"pointer",marginBottom:16}}>+ Nuevo ingreso</button>

            {/* Reset */}
            <div style={S.sec}>🔴 Zona peligrosa</div>
            <div style={{...S.card,border:`1px solid ${C.red}25`}}>
              <div style={{fontSize:12,color:C.muted,marginBottom:10}}>Esto borra todos los datos y regresa a los valores iniciales. No se puede deshacer.</div>
              <button onClick={()=>{if(window.confirm("¿Restablecer todo a los datos iniciales?")){setBolsillos(BOLSILLOS_0);setPagos(PAGOS_0);setIngresos(INGRESOS_0);setGastosList([]);setHistorial([]);setPlan(null);setPlanOk(false);setMontoInput("");setFuenteSel("");}}} style={{...S.btn(C.red+"20",C.red),border:`1px solid ${C.red}30`,width:"100%",padding:"10px"}}>Restablecer datos iniciales</button>
            </div>
          </div>
        )}
      </div>

      {/* ════ MODALES ════ */}

      {/* Modal bolsillo */}
      {modal?.tipo==="bolsillo"&&(
        <Modal title={modal.data?"Editar bolsillo":"Nuevo bolsillo"} color={C.accent} onClose={closeModal}>
          <Field label="Nombre"><Inp placeholder="Ej: Vacaciones" value={formBolsillo.nombre} onChange={e=>setFormBolsillo(x=>({...x,nombre:e.target.value}))}/></Field>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Ícono (emoji)"><Inp placeholder="💰" value={formBolsillo.icon} onChange={e=>setFormBolsillo(x=>({...x,icon:e.target.value}))}/></Field>
            <Field label="Color">
              <Sel value={formBolsillo.color} onChange={e=>setFormBolsillo(x=>({...x,color:e.target.value}))}>
                {COLOR_OPTIONS.map(c=><option key={c.value} value={c.value}>{c.label}</option>)}
              </Sel>
            </Field>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Meta total (COP)"><Inp type="number" placeholder="Ej: 2000000" value={formBolsillo.meta} onChange={e=>setFormBolsillo(x=>({...x,meta:e.target.value}))}/></Field>
            <Field label="Cuota mensual (COP)"><Inp type="number" placeholder="Ej: 200000" value={formBolsillo.metaMensual} onChange={e=>setFormBolsillo(x=>({...x,metaMensual:e.target.value}))}/></Field>
          </div>
          <Field label="Tipo">
            <Sel value={formBolsillo.grupo} onChange={e=>setFormBolsillo(x=>({...x,grupo:e.target.value}))}>
              <option value="sagrado">🔒 Sagrado — no tocar</option>
              <option value="vida">☀️ Vida & gastos</option>
              <option value="provision">⚡ Provisión / deuda</option>
            </Sel>
          </Field>
          <Field label="Descripción"><Inp placeholder="Para qué es este bolsillo" value={formBolsillo.descripcion} onChange={e=>setFormBolsillo(x=>({...x,descripcion:e.target.value}))}/></Field>
          {/* preview color */}
          <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
            {COLOR_OPTIONS.map(co=>(
              <button key={co.value} onClick={()=>setFormBolsillo(x=>({...x,color:co.value}))} style={{width:28,height:28,borderRadius:"50%",background:co.value,border:formBolsillo.color===co.value?`3px solid #fff`:`2px solid transparent`,cursor:"pointer"}}/>
            ))}
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={guardarBolsillo} style={{...S.btn(C.accent),flex:1,padding:"11px"}}>Guardar</button>
            <button onClick={closeModal} style={{...S.btn("#ffffff10",C.muted),padding:"11px 16px"}}>Cancelar</button>
          </div>
        </Modal>
      )}

      {/* Modal pago */}
      {modal?.tipo==="pago"&&(
        <Modal title={modal.data?"Editar pago":"Nuevo pago"} color={C.amber} onClose={closeModal}>
          <Field label="Nombre"><Inp placeholder="Ej: Deuda X" value={formPago.nombre} onChange={e=>setFormPago(x=>({...x,nombre:e.target.value}))}/></Field>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Monto (COP)"><Inp type="number" placeholder="Ej: 500000" value={formPago.monto} onChange={e=>setFormPago(x=>({...x,monto:e.target.value}))}/></Field>
            <Field label="Día de pago"><Inp type="number" placeholder="1–31" value={formPago.diaPago} onChange={e=>setFormPago(x=>({...x,diaPago:e.target.value}))}/></Field>
          </div>

          {/* URGENCIA — el corazón del cambio */}
          <Field label="Nivel de urgencia">
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
              {Object.entries(URGENCIA).map(([key,u])=>(
                <button key={key} onClick={()=>setFormPago(x=>({...x,urgencia:key}))} style={{
                  background:formPago.urgencia===key?u.color+"25":"#ffffff08",
                  border:`1px solid ${formPago.urgencia===key?u.color:C.border}`,
                  borderRadius:10, padding:"10px 8px", cursor:"pointer", textAlign:"center",
                  transition:"all .15s",
                }}>
                  <div style={{fontSize:13,marginBottom:3}}>{u.label}</div>
                  <div style={{fontSize:10,color:formPago.urgencia===key?u.color:C.muted,lineHeight:1.4}}>{u.desc}</div>
                </button>
              ))}
            </div>
          </Field>

          <div style={{background:
            formPago.urgencia==="critica"?C.red+"10":
            formPago.urgencia==="flexible"?C.teal+"10":C.amber+"10",
            border:`1px solid ${formPago.urgencia==="critica"?C.red:formPago.urgencia==="flexible"?C.teal:C.amber}20`,
            borderRadius:10,padding:"9px 13px",marginBottom:10,fontSize:11,
            color:formPago.urgencia==="critica"?C.red:formPago.urgencia==="flexible"?C.teal:C.amber,lineHeight:1.6,
          }}>
            {formPago.urgencia==="critica" && "La consejera provisiona esto PRIMERO, antes que vida diaria y sueños."}
            {formPago.urgencia==="normal"  && "La consejera lo provisiona después de urgencias y vida diaria."}
            {formPago.urgencia==="flexible"&& "La consejera solo abona aquí si sobra después de todo lo demás."}
          </div>

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Bolsillo asociado">
              <Sel value={formPago.bolsilloId} onChange={e=>setFormPago(x=>({...x,bolsilloId:e.target.value}))}>
                {bolsillos.map(b=><option key={b.id} value={b.id}>{b.icon} {b.nombre}</option>)}
              </Sel>
            </Field>
            <Field label="Categoría">
              <Sel value={formPago.categoria} onChange={e=>setFormPago(x=>({...x,categoria:e.target.value}))}>
                {Object.entries(CATS).map(([k,v])=><option key={k} value={k}>{v.icon} {v.label}</option>)}
              </Sel>
            </Field>
          </div>
          <Field label="Nota (opcional)"><Inp placeholder="Ej: Puede esperar 2 meses" value={formPago.nota} onChange={e=>setFormPago(x=>({...x,nota:e.target.value}))}/></Field>
          <div style={{display:"flex",gap:8}}>
            <button onClick={guardarPago} style={{...S.btn(C.amber),flex:1,padding:"11px"}}>Guardar</button>
            <button onClick={closeModal} style={{...S.btn("#ffffff10",C.muted),padding:"11px 16px"}}>Cancelar</button>
          </div>
        </Modal>
      )}

      {/* Modal ingreso */}
      {modal?.tipo==="ingreso"&&(
        <Modal title={modal.data?"Editar ingreso":"Nuevo ingreso esperado"} color={C.blue} onClose={closeModal}>
          <Field label="Fuente / Nombre"><Inp placeholder="Ej: Cliente nuevo" value={formIngreso.fuente} onChange={e=>setFormIngreso(x=>({...x,fuente:e.target.value}))}/></Field>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Monto (COP)"><Inp type="number" placeholder="Ej: 1000000" value={formIngreso.monto} onChange={e=>setFormIngreso(x=>({...x,monto:e.target.value}))}/></Field>
            <Field label="Día esperado"><Inp type="number" placeholder="1–31" value={formIngreso.diaEsp} onChange={e=>setFormIngreso(x=>({...x,diaEsp:e.target.value}))}/></Field>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={guardarIngreso} style={{...S.btn(C.blue),flex:1,padding:"11px"}}>Guardar</button>
            <button onClick={closeModal} style={{...S.btn("#ffffff10",C.muted),padding:"11px 16px"}}>Cancelar</button>
          </div>
        </Modal>
      )}

      {/* Modal gasto */}
      {modal?.tipo==="gasto"&&(
        <Modal title="Registrar gasto" color={C.pink} onClose={closeModal}>
          <Field label="¿En qué?"><Inp placeholder="Ej: Meds Ori, gasolina..." value={formGasto.nombre} onChange={e=>setFormGasto(x=>({...x,nombre:e.target.value}))}/></Field>
          <Field label="Monto (COP)"><Inp type="number" placeholder="0" value={formGasto.monto} onChange={e=>setFormGasto(x=>({...x,monto:e.target.value}))}/></Field>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            <Field label="Categoría">
              <Sel value={formGasto.categoria} onChange={e=>{const bid=CATS[e.target.value]?.bolsilloId||"vida";setFormGasto(x=>({...x,categoria:e.target.value,bolsilloId:bid}));}}>
                {Object.entries(CATS).map(([k,v])=><option key={k} value={k}>{v.icon} {v.label}</option>)}
              </Sel>
            </Field>
            <Field label="Bolsillo (descuenta de)">
              <Sel value={formGasto.bolsilloId||CATS[formGasto.categoria]?.bolsilloId||"vida"} onChange={e=>setFormGasto(x=>({...x,bolsilloId:e.target.value}))}>
                {bolsillos.map(b=><option key={b.id} value={b.id}>{b.icon} {b.nombre} — {fmt(b.saldo)}</option>)}
              </Sel>
            </Field>
          </div>
          <Field label="Fecha"><Inp type="date" value={formGasto.fecha} onChange={e=>setFormGasto(x=>({...x,fecha:e.target.value}))}/></Field>
          <div style={{background:C.orange+"12",border:`1px solid ${C.orange}25`,borderRadius:10,padding:"9px 13px",marginBottom:14,fontSize:12,color:C.orange}}>
            💡 Al registrar, se descontará automáticamente del bolsillo seleccionado.
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={registrarGasto} style={{...S.btn(C.pink),flex:1,padding:"11px"}}>Registrar y descontar</button>
            <button onClick={closeModal} style={{...S.btn("#ffffff10",C.muted),padding:"11px 16px"}}>Cancelar</button>
          </div>
        </Modal>
      )}
    </div>
  );
}
