import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ─── SUPABASE ───────────────────────────────────────────────
const SB_URL = "https://goizernbjfthejdekykz.supabase.co";
const SB_KEY = "sb_publishable_QwxzW5eyomjU5WDtW6qERQ_Kaxxgx89";
const ADMIN_PASS = "Budmat2026";
const sbH = { "apikey": SB_KEY, "Authorization": `Bearer ${SB_KEY}`, "Content-Type": "application/json" };
const supabase = createClient(SB_URL, SB_KEY);

async function sbGet(t, qs = "") { const r = await fetch(`${SB_URL}/rest/v1/${t}?select=*${qs}`, { headers: sbH }); if (!r.ok) throw new Error(await r.text()); return r.json(); }
async function sbPost(t, d) { const r = await fetch(`${SB_URL}/rest/v1/${t}`, { method: "POST", headers: { ...sbH, "Prefer": "return=minimal" }, body: JSON.stringify(d) }); if (!r.ok) throw new Error(await r.text()); }
async function sbPatch(t, f, d) { const r = await fetch(`${SB_URL}/rest/v1/${t}?${f}`, { method: "PATCH", headers: { ...sbH, "Prefer": "return=minimal" }, body: JSON.stringify(d) }); if (!r.ok) throw new Error(await r.text()); }
async function sbDelete(t, f) { const r = await fetch(`${SB_URL}/rest/v1/${t}?${f}`, { method: "DELETE", headers: sbH }); if (!r.ok) throw new Error(await r.text()); }
async function sbUpsert(t, d, c = "sp") { const r = await fetch(`${SB_URL}/rest/v1/${t}?on_conflict=${c}`, { method: "POST", headers: { ...sbH, "Prefer": "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify(d) }); if (!r.ok) throw new Error(await r.text()); }

// ─── CONSTANTS ──────────────────────────────────────────────
const SP = ["Andrzej Górny", "Arkadiusz Czerniak"];
const SP_LABEL = { "Andrzej Górny": "Andrzej", "Arkadiusz Czerniak": "Arek" };
const SP_INIT  = { "Andrzej Górny": "AN", "Arkadiusz Czerniak": "AR" };
const SP_COLOR = { "Andrzej Górny": "#185FA5", "Arkadiusz Czerniak": "#0F6E56" };
const MONTHS   = ["Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec","Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień"];
const now = new Date();

function ds(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; }
function todayStr() { const t = new Date(); t.setHours(0,0,0,0); return ds(t); }

// ─── EXCEL PARSE ────────────────────────────────────────────
function parseExcel(file) {
  return new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
        res(rows.map(r => {
          const d = r["Data wprow."]; const dt = d ? (d instanceof Date ? d : new Date(d)) : null;
          return { przypisany: r["Przypisany"]||null, status: r["Status"]||null, status_crm: r["Status CRM"]||null, kontrahent: r["Kontrahent"]||null, model: r["Potrzeba - Model"]||null, data_wprow: dt ? dt.toISOString().split("T")[0] : null, count_override: 1 };
        }).filter(r => r.przypisany));
      } catch(err) { rej(err); }
    };
    fr.onerror = rej; fr.readAsArrayBuffer(file);
  });
}

// ─── KPI ENGINE ─────────────────────────────────────────────
function isOrder(r) { return r.status === "03. Realizacja umów" || (r.status === "04. Zakończona" && r.status_crm === "Sukces"); }

function computeKPIs(leads, manual, targets, month, year) {
  const out = {};
  for (const sp of SP) {
    const spLeads = leads.filter(r => r.przypisany === sp);
    const curLeads = spLeads.filter(r => { if (!r.data_wprow) return false; const d = new Date(r.data_wprow); return d.getMonth()+1 === month && d.getFullYear() === year; });
    const dmsOrders = curLeads.filter(isOrder).reduce((s,r) => s+(r.count_override||1), 0);

    const allManCur = manual.filter(o => { const d=new Date(o.date); return o.sp===sp && d.getMonth()+1===month && d.getFullYear()===year; });
    const overrideRec = allManCur.find(o => o.model==="__TOTAL_OVERRIDE__");
    const manList = allManCur.filter(o => !["__TOTAL_OVERRIDE__","__KOREKTA__"].includes(o.model));
    const effectiveDms = overrideRec ? Math.max(0, parseInt(overrideRec.client)||0) : dmsOrders;
    const totalOrders = effectiveDms + manList.length;
    const convMon = curLeads.length > 0 ? Math.round(totalOrders/curLeads.length*100) : 0;

    const pipelineRows = spLeads.filter(r => {
      if (r.status !== "02. Wybór ofert" || !r.data_wprow) return false;
      const d = new Date(r.data_wprow); const dm = d.getMonth()+1, dy = d.getFullYear();
      for (let i=0; i<=2; i++) { let m=month-i, y=year; if (m<=0){m+=12;y--;} if (dm===m&&dy===y) return true; }
      return false;
    });

    const target = (targets.find(t => t.sp===sp)||{}).target||10;
    out[sp] = { totalOrders, dmsOrders, effectiveDms, hasOverride: !!overrideRec, manOrders: manList.length, curLeads: curLeads.length, pipeline: pipelineRows.length, pipelineRows, convMon, target, planPct: Math.round(totalOrders/target*100) };
  }
  [...SP].sort((a,b) => out[b].totalOrders-out[a].totalOrders||out[b].convMon-out[a].convMon).forEach((sp,i) => out[sp].rank=i+1);
  return out;
}

// ─── GAMIFICATION ENGINE ────────────────────────────────────
function computeGame(leads, manual, kpis, month, year) {
  const td = new Date(); td.setHours(0,0,0,0);
  const TDS = ds(td);
  const daysInMonth = new Date(year, month, 0).getDate();
  const isCur = month===now.getMonth()+1 && year===now.getFullYear();
  const daysLeft = isCur ? daysInMonth - td.getDate() : 0;
  const daysElapsed = isCur ? td.getDate() : daysInMonth;

  const stats = {};
  for (const sp of SP) {
    const spLeads = leads.filter(r => r.przypisany===sp);
    const spMan = manual.filter(o => o.sp===sp);
    const byDate = {};

    spLeads.forEach(r => {
      if (!r.data_wprow||!isOrder(r)) return;
      const d = new Date(r.data_wprow);
      if (d.getMonth()+1===month && d.getFullYear()===year) byDate[r.data_wprow] = (byDate[r.data_wprow]||0)+(r.count_override||1);
    });
    spMan.forEach(o => { const d = new Date(o.date); if (d.getMonth()+1===month&&d.getFullYear()===year) byDate[o.date]=(byDate[o.date]||0)+1; });

    // Streak
    let streak = 0;
    const yd = new Date(td); yd.setDate(td.getDate()-1);
    const YDS = ds(yd);
    let cur = byDate[TDS] ? new Date(td) : byDate[YDS] ? new Date(yd) : null;
    if (cur) { while(true) { const s=ds(cur); const mn=String(month).padStart(2,"0"); if (s<`${year}-${mn}-01`) break; if (byDate[s]) { streak++; cur.setDate(cur.getDate()-1); } else break; } }

    // Last 5 days
    let last5 = 0;
    for (let i=0; i<5; i++) { const d2=new Date(td); d2.setDate(td.getDate()-i); last5+=(byDate[ds(d2)]||0); }

    // Today
    const ordersToday = byDate[TDS]||0;

    // This week
    const wd = td.getDay()===0?7:td.getDay();
    const ws = new Date(td); ws.setDate(td.getDate()-wd+1);
    const WSS = ds(ws);
    let weeklyOrders = 0;
    Object.entries(byDate).forEach(([date,cnt]) => { if (date>=WSS&&date<=TDS) weeklyOrders+=cnt; });

    // Personal best
    const mTotals = {};
    spLeads.forEach(r => { if (!r.data_wprow||!isOrder(r)) return; const d=new Date(r.data_wprow); const k=`${d.getFullYear()}-${d.getMonth()+1}`; mTotals[k]=(mTotals[k]||0)+(r.count_override||1); });
    spMan.forEach(o => { const d=new Date(o.date); const k=`${d.getFullYear()}-${d.getMonth()+1}`; mTotals[k]=(mTotals[k]||0)+1; });
    let pBest=0, pBestLabel=null;
    Object.entries(mTotals).forEach(([k,v]) => { if (v>pBest) { pBest=v; const [y2,m2]=k.split("-"); pBestLabel=`${MONTHS[parseInt(m2)-1]} ${y2}`; } });

    // Pace
    const { totalOrders, target } = kpis[sp];
    const needed = Math.max(0, target-totalOrders);
    const reqPace = daysLeft>0 ? needed/daysLeft : 0;
    const curPace = daysElapsed>0 ? totalOrders/daysElapsed : 0;

    stats[sp] = { streak, last5, ordersToday, weeklyOrders, weeklyTarget: Math.ceil(target/4), pBest, pBestLabel, daysLeft, needed, reqPace, curPace, isOnPace: curPace>=reqPace, badges:[] };
  }

  // Badges
  const leader = SP.find(s => kpis[s].rank===1);
  const chaser = SP.find(s => kpis[s].rank===2);
  const gap = kpis[leader].totalOrders - kpis[chaser].totalOrders;

  // Leader badges
  if (stats[leader].streak>=3) stats[leader].badges.push({ e:"🔥", l:`${stats[leader].streak} dni z rzędu`, c:"#dc2626", b:"#fee2e2" });
  if (stats[leader].isOnPace) stats[leader].badges.push({ e:"✅", l:"Na kursie", c:"#059669", b:"#d1fae5" });
  if (kpis[leader].totalOrders>=kpis[leader].target) stats[leader].badges.push({ e:"🏆", l:"Cel osiągnięty!", c:"#BA7517", b:"#FAEEDA" });

  // Chaser badges — every positive signal visible
  if (stats[chaser].last5 > stats[leader].last5 && stats[chaser].last5>0)
    stats[chaser].badges.push({ e:"⚡", l:`Przyspiesza! (${stats[chaser].last5} vs ${stats[leader].last5} w ostatnich 5 dniach)`, c:"#7c3aed", b:"#ede9fe" });
  if (gap===1)
    stats[chaser].badges.push({ e:"🎯", l:"Jeden krok do lidera!", c:"#dc2626", b:"#fee2e2" });
  if (gap===0)
    stats[chaser].badges.push({ e:"🤝", l:"Remis na szczycie!", c:"#185FA5", b:"#dbeafe" });
  if (stats[chaser].ordersToday > stats[leader].ordersToday)
    stats[chaser].badges.push({ e:"🔥", l:"Dziś rządzi!", c:"#ea580c", b:"#ffedd5" });
  if (kpis[chaser].convMon > kpis[leader].convMon && kpis[chaser].convMon>0)
    stats[chaser].badges.push({ e:"📈", l:"Lepsza konwersja niż lider!", c:"#059669", b:"#d1fae5" });
  if (stats[chaser].streak>=2)
    stats[chaser].badges.push({ e:"🔥", l:`${stats[chaser].streak} dni z rzędu`, c:"#dc2626", b:"#fee2e2" });
  if (stats[chaser].weeklyOrders>=stats[chaser].weeklyTarget && stats[chaser].weeklyTarget>0)
    stats[chaser].badges.push({ e:"🎉", l:"Cel tygodniowy osiągnięty!", c:"#7c3aed", b:"#ede9fe" });
  if (stats[chaser].isOnPace)
    stats[chaser].badges.push({ e:"✅", l:"Na kursie do celu", c:"#059669", b:"#d1fae5" });
  if (kpis[chaser].totalOrders>=stats[chaser].pBest && stats[chaser].pBest>0)
    stats[chaser].badges.push({ e:"🏅", l:"Bije własny rekord!", c:"#BA7517", b:"#FAEEDA" });
  if (kpis[chaser].totalOrders>=kpis[chaser].target)
    stats[chaser].badges.push({ e:"🏆", l:"Cel osiągnięty!", c:"#BA7517", b:"#FAEEDA" });

  return stats;
}

// ─── STYLES ─────────────────────────────────────────────────
const inpM = { width:"100%", padding:"8px 10px", borderRadius:"8px", border:"1px solid #d1d5db", background:"#fff", color:"#111827", fontSize:"14px", fontFamily:"inherit", boxSizing:"border-box" };
const btnP = { padding:"10px 20px", borderRadius:"8px", border:"none", background:"#185FA5", color:"#fff", fontWeight:"500", cursor:"pointer", fontSize:"14px", display:"flex", alignItems:"center", gap:"6px" };
const btnS = { padding:"10px 16px", borderRadius:"8px", border:"1px solid #d1d5db", background:"#fff", color:"#374151", cursor:"pointer", fontSize:"14px" };
const planColor = p => p>=100?"#3B6D11":p>=60?"#185FA5":p>=30?"#BA7517":"#A32D2D";

// ─── MODAL ──────────────────────────────────────────────────
function Modal({ onClose, title, subtitle, children, maxWidth="440px" }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:9999, padding:"20px" }}>
      <div style={{ background:"#fff", border:"1px solid #d1d5db", borderRadius:"12px", padding:"24px", width:"100%", maxWidth, boxShadow:"0 8px 32px rgba(0,0,0,0.2)", maxHeight:"88vh", overflow:"hidden", display:"flex", flexDirection:"column" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:"16px", flexShrink:0 }}>
          <div>
            <p style={{ fontSize:"16px", fontWeight:"500", margin:"0 0 2px", color:"#111827" }}>{title}</p>
            {subtitle&&<p style={{ fontSize:"12px", color:"#6b7280", margin:0 }}>{subtitle}</p>}
          </div>
          <button onClick={onClose} style={{ background:"none", border:"none", cursor:"pointer", color:"#6b7280", fontSize:"20px", padding:0, lineHeight:1 }}><i className="ti ti-x"/></button>
        </div>
        <div style={{ overflowY:"auto", flex:1 }}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return <div style={{ marginBottom:"14px" }}><p style={{ fontSize:"11px", color:"#6b7280", margin:"0 0 5px", textTransform:"uppercase", letterSpacing:"0.3px" }}>{label}</p>{children}</div>;
}

function Badge({ e, l, c, b }) {
  return <span style={{ display:"inline-flex", alignItems:"center", gap:"4px", fontSize:"11px", fontWeight:"500", padding:"4px 10px", borderRadius:"20px", background:b, color:c, whiteSpace:"nowrap" }}>{e} {l}</span>;
}

// ════════════════════════════════════════════════════════════
export default function App() {
  const [page, setPage]       = useState("dashboard");
  const [isAdmin, setIsAdmin] = useState(false);
  const [adminTab, setAdminTab] = useState("upload");
  const [passInput, setPassInput] = useState("");
  const [passErr, setPassErr] = useState(false);

  const [leads, setLeads]   = useState([]);
  const [manual, setManual] = useState([]);
  const [targets, setTargets] = useState([]);
  const [loading, setLoading] = useState(true);
  const [dbErr, setDbErr]   = useState(null);

  const [selMonth, setSelMonth] = useState(now.getMonth()+1);
  const [selYear] = useState(now.getFullYear());
  const [modal, setModal] = useState(null);
  const [pipelineSP, setPipelineSP] = useState(null);

  const [addForm, setAddForm] = useState({ sp: SP[0], count: 1, client: "" });

  // Admin upload
  const [xlsxName, setXlsxName] = useState(null);
  const [preview, setPreview]   = useState(null);
  const [uploading, setUploading] = useState(false);
  const [uploadMsg, setUploadMsg] = useState(null);

  // Admin targets
  const [editTargets, setEditTargets]   = useState(null);
  const [savingTargets, setSavingTargets] = useState(false);

  // Admin DMS order editing
  const [editingCountId, setEditingCountId] = useState(null);
  const [editingCountVal, setEditingCountVal] = useState(1);

  // Admin total editing
  const [editingTotalSP, setEditingTotalSP] = useState(null);
  const [editTotalVal, setEditTotalVal] = useState(0);

  const [records, setRecords]       = useState([]);
  const [editRecords, setEditRecords] = useState(null);
  const [savingRecords, setSavingRecords] = useState(false);

  // Achievements (cele zespołu)
  const [achievements, setAchievements] = useState([]);
  const [newAch, setNewAch] = useState({ name:"", description:"", target:30 });
  const [savingAch, setSavingAch] = useState(false);

  const load = useCallback(async () => {
    setLoading(true); setDbErr(null);
    try {
      const [l,m,t,r] = await Promise.all([sbGet("leads"), sbGet("manual_orders"), sbGet("targets"), sbGet("records")]);
      setLeads(l); setManual(m); setTargets(t); setRecords(r);
    } catch(e) { setDbErr(e.message); }
    // Cele ładujemy osobno — gdyby tabeli brakowało, dashboard i tak działa
    try { setAchievements(await sbGet("achievements")); } catch(e) { setAchievements([]); }
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  useEffect(() => {
    const channel = supabase
      .channel("db-realtime")
      .on("postgres_changes", { event: "*", schema: "public", table: "manual_orders" }, () => load())
      .on("postgres_changes", { event: "*", schema: "public", table: "leads" }, () => load())
      .on("postgres_changes", { event: "*", schema: "public", table: "achievements" }, () => load())
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, [load]);

  useEffect(() => {
    const interval = setInterval(() => { load(); }, 5 * 60 * 1000);
    return () => clearInterval(interval);
  }, [load]);

  function handleLogin(e) {
    e.preventDefault();
    if (passInput===ADMIN_PASS) { setIsAdmin(true); setPage("admin"); setPassErr(false); setPassInput(""); }
    else setPassErr(true);
  }

  async function handleFile(e) {
    const f = e.target.files[0]; if (!f) return;
    setXlsxName(f.name); setUploadMsg(null);
    try { setPreview(await parseExcel(f)); } catch(err) { setUploadMsg({ ok:false, text:"Błąd: "+err.message }); }
    e.target.value="";
  }

  async function confirmUpload() {
    if (!preview) return;
    setUploading(true); setUploadMsg(null);
    try {
      await sbDelete("leads","id=gte.0");
      const B=400;
      for (let i=0;i<preview.length;i+=B) await sbPost("leads",preview.slice(i,i+B));
      setPreview(null); setXlsxName(null);
      setUploadMsg({ ok:true, text:`Wgrano ${preview.length} rekordów — ${new Date().toLocaleTimeString("pl-PL")}` });
      await load();
    } catch(err) { setUploadMsg({ ok:false, text:"Błąd zapisu: "+err.message }); }
    setUploading(false);
  }

  async function saveTargets() {
    setSavingTargets(true);
    try {
      await Promise.all(SP.map(sp => sbUpsert("targets",{ sp, target: editTargets[sp]||10 })));
      await load(); setEditTargets(null);
    } catch(err) { alert("Błąd: "+err.message); }
    setSavingTargets(false);
  }

  async function addOrder() {
    if (!addForm.model.trim()||!addForm.client.trim()) return;
    try {
      await sbPost("manual_orders",{ sp:addForm.sp, model:addForm.model, client:addForm.client, date:addForm.date });
      await load(); setModal(null); setAddForm(f=>({...f,model:"",client:""}));
    } catch(err) { alert("Błąd: "+err.message); }
  }

  async function deleteOrder(id) {
    try { await sbDelete("manual_orders",`id=eq.${id}`); await load(); } catch(err) { alert("Błąd: "+err.message); }
  }

  async function saveLeadCount(id, val) {
    try { await sbPatch("leads",`id=eq.${id}`,{ count_override: Math.max(1,parseInt(val)||1) }); await load(); setEditingCountId(null); }
    catch(err) { alert("Błąd: "+err.message); }
  }

  async function saveRecords() {
    setSavingRecords(true);
    try {
      await Promise.all(SP.map(sp => sbUpsert("records",{ sp, best_count: editRecords[sp]?.count||0, best_month: editRecords[sp]?.month||"" })));
      await load(); setEditRecords(null);
    } catch(err) { alert("Błąd: "+err.message); }
    setSavingRecords(false);
  }

  async function saveMonthTotal(sp, newTotal) {
    const existing = manual.filter(o => {
      const d=new Date(o.date);
      return o.sp===sp && d.getMonth()+1===selMonth && d.getFullYear()===selYear && ["__TOTAL_OVERRIDE__","__KOREKTA__"].includes(o.model);
    });
    try {
      for (const o of existing) await sbDelete("manual_orders",`id=eq.${o.id}`);
      const dateStr = `${selYear}-${String(selMonth).padStart(2,"0")}-01`;
      await sbPost("manual_orders",{ sp, model:"__TOTAL_OVERRIDE__", client: String(Math.max(0,newTotal)), date: dateStr });
      await load(); setEditingTotalSP(null);
    } catch(err) { alert("Błąd: "+err.message); }
  }

  async function addOrders() {
    const count = Math.max(1, parseInt(addForm.count)||1);
    const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}`;
    const client = addForm.client.trim() || "—";
    try {
      for (let i=0; i<count; i++) await sbPost("manual_orders",{ sp: addForm.sp, model:"Zamówienie", client, date: dateStr });
      await load(); setModal(null); setAddForm(f=>({...f, count:1, client:""}));
    } catch(err) { alert("Błąd: "+err.message); }
  }

  // ── Cele / Achievements ──
  async function addAchievement() {
    if (!newAch.name.trim()) return alert("Podaj nazwę celu");
    const target = Math.max(1, parseInt(newAch.target)||1);
    setSavingAch(true);
    try {
      await sbPost("achievements", { name: newAch.name.trim(), description: newAch.description.trim()||null, target, active: true });
      await load(); setNewAch({ name:"", description:"", target:30 });
    } catch(err) { alert("Błąd: "+err.message); }
    setSavingAch(false);
  }
  async function toggleAchievement(id, active) {
    try { await sbPatch("achievements",`id=eq.${id}`,{ active: !active }); await load(); }
    catch(err) { alert("Błąd: "+err.message); }
  }
  async function deleteAchievement(id) {
    try { await sbDelete("achievements",`id=eq.${id}`); await load(); }
    catch(err) { alert("Błąd: "+err.message); }
  }
  // Postęp celu = realne zamówienia (DMS + ręczne) od dnia ustawienia celu, cały dział
  function achProgress(ach) {
    const since = (ach.created_at||"").slice(0,10);
    let count = 0;
    leads.forEach(r => { if (isOrder(r) && r.data_wprow && r.data_wprow >= since) count += (r.count_override||1); });
    manual.forEach(o => { if (["__TOTAL_OVERRIDE__","__KOREKTA__"].includes(o.model)) return; if ((o.date||"") >= since) count += 1; });
    return count;
  }
  const kpis = leads.length>0 ? computeKPIs(leads,manual,targets,selMonth,selYear) : null;
  const game = kpis ? computeGame(leads,manual,kpis,selMonth,selYear) : null;
  const ranked = kpis ? [...SP].sort((a,b)=>kpis[a].rank-kpis[b].rank) : SP;

  const manualThisMonth = manual.filter(o=>{ const d=new Date(o.date); return d.getMonth()+1===selMonth&&d.getFullYear()===selYear&&o.sp!=="Inne"&&!["__TOTAL_OVERRIDE__","__KOREKTA__"].includes(o.model); });
  const inneThisMonth = manual.filter(o=>{ const d=new Date(o.date); return d.getMonth()+1===selMonth&&d.getFullYear()===selYear&&o.sp==="Inne"; });

  const teamOrders  = kpis ? SP.reduce((s,sp)=>s+kpis[sp].totalOrders,0) + inneThisMonth.length : 0;
  const teamTarget  = kpis ? SP.reduce((s,sp)=>s+kpis[sp].target,0) : 0;
  const teamPlan    = teamTarget>0 ? Math.round(teamOrders/teamTarget*100) : 0;
  const teamPipeline = kpis ? SP.reduce((s,sp)=>s+kpis[sp].pipeline,0) : 0;
  const teamNeeded  = kpis ? Math.max(0,teamTarget-teamOrders) : 0;
  const leader = kpis ? SP.find(s=>kpis[s].rank===1) : null;
  const chaser = kpis ? SP.find(s=>kpis[s].rank===2) : null;
  const gap = kpis && leader && chaser ? kpis[leader].totalOrders-kpis[chaser].totalOrders : 0;
  const daysInMonth = new Date(selYear,selMonth,0).getDate();
  const isCurMonth = selMonth===now.getMonth()+1&&selYear===now.getFullYear();
  const daysLeft = isCurMonth ? daysInMonth-now.getDate() : 0;

  // DMS orders for admin (status 03 or 04/Sukces, any date)
  const dmsOrderLeads = leads.filter(isOrder).sort((a,b)=>(b.data_wprow||"").localeCompare(a.data_wprow||""));
  const tMap = Object.fromEntries(targets.map(t=>[t.sp,t.target]));
  const rMap = Object.fromEntries(records.map(r=>[r.sp,{ count: r.best_count, month: r.best_month }]));

  // ══════════════════════════════════════════════════════════
  // LOGIN
  if (page==="login") return (
    <div style={{ minHeight:"500px", display:"flex", alignItems:"center", justifyContent:"center", padding:"20px" }}>
      <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:"12px", padding:"32px", width:"100%", maxWidth:"340px", boxShadow:"0 4px 16px rgba(0,0,0,0.08)" }}>
        <div style={{ width:"40px",height:"40px",background:"#185FA5",borderRadius:"10px",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:"16px" }}><i className="ti ti-settings" style={{ color:"#fff",fontSize:"20px" }}/></div>
        <p style={{ fontSize:"18px",fontWeight:"500",margin:"0 0 4px" }}>Panel administratora</p>
        <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 24px" }}>Ford Budmat Auto · Płock</p>
        <form onSubmit={handleLogin}>
          <Field label="Hasło dostępu"><input type="password" value={passInput} onChange={e=>{setPassInput(e.target.value);setPassErr(false);}} style={inpM} placeholder="••••••••" autoFocus/></Field>
          {passErr&&<p style={{ fontSize:"12px",color:"#dc2626",margin:"-8px 0 12px" }}>Nieprawidłowe hasło</p>}
          <div style={{ display:"flex",gap:"8px",marginTop:"8px" }}>
            <button type="submit" style={{ ...btnP,flex:1,justifyContent:"center" }}>Zaloguj</button>
            <button type="button" onClick={()=>setPage("dashboard")} style={btnS}>Wróć</button>
          </div>
        </form>
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════════
  // ADMIN
  if (page==="admin"&&isAdmin) return (
    <div style={{ fontFamily:"inherit",padding:"20px",maxWidth:"800px",margin:"0 auto" }}>
      <div style={{ display:"flex",alignItems:"center",gap:"10px",marginBottom:"20px",flexWrap:"wrap" }}>
        <div style={{ flex:1 }}>
          <p style={{ fontSize:"17px",fontWeight:"500",margin:"0 0 1px" }}>Panel administratora</p>
          <p style={{ fontSize:"12px",color:"#6b7280",margin:0 }}>Ford Budmat Auto · Płock</p>
        </div>
        <button onClick={()=>setPage("dashboard")} style={{ ...btnS,display:"flex",alignItems:"center",gap:"5px" }}><i className="ti ti-layout-dashboard" style={{ fontSize:"15px" }}/>Dashboard</button>
        <button onClick={()=>{setIsAdmin(false);setPage("dashboard");setPassInput("");}} style={{ ...btnS,color:"#6b7280" }}>Wyloguj</button>
      </div>

      {/* Tabs */}
      <div style={{ display:"flex",gap:"2px",marginBottom:"16px",background:"#f3f4f6",padding:"4px",borderRadius:"8px" }}>
        {[{id:"upload",icon:"ti-upload",label:"Export DMS"},{id:"targets",icon:"ti-target",label:"Plany"},{id:"orders",icon:"ti-clipboard-list",label:"Zamówienia"},{id:"cele",icon:"ti-trophy",label:"Cele"}].map(tab=>(
          <button key={tab.id} onClick={()=>setAdminTab(tab.id)} style={{ flex:1,padding:"8px",borderRadius:"6px",border:"none",cursor:"pointer",fontSize:"13px",fontWeight:adminTab===tab.id?"500":"400",background:adminTab===tab.id?"#fff":"transparent",color:adminTab===tab.id?"#111827":"#6b7280",display:"flex",alignItems:"center",justifyContent:"center",gap:"6px" }}>
            <i className={`ti ${tab.icon}`} style={{ fontSize:"14px" }}/>{tab.label}
          </button>
        ))}
      </div>

      {/* ── UPLOAD ── */}
      {adminTab==="upload"&&(
        <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"20px" }}>
          <div style={{ display:"flex",gap:"16px",marginBottom:"20px",padding:"12px 16px",background:"#f9fafb",borderRadius:"8px" }}>
            <div><p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 2px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Rekordów w bazie</p><p style={{ fontSize:"22px",fontWeight:"500",margin:0 }}>{leads.length}</p></div>
          </div>
          <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 14px" }}>Wgraj export z Autostacji (.xlsx). Istniejące dane zostaną <strong>zastąpione</strong> — zamówienia ręczne i plany pozostają.</p>
          <div style={{ display:"flex",gap:"10px",flexWrap:"wrap",marginBottom:"14px" }}>
            <label style={{ padding:"9px 16px",borderRadius:"8px",border:"1px solid #d1d5db",cursor:"pointer",fontSize:"14px",display:"flex",alignItems:"center",gap:"6px",userSelect:"none" }}>
              <i className="ti ti-file-spreadsheet" style={{ fontSize:"15px" }}/>{xlsxName||"Wybierz plik .xlsx"}
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }}/>
            </label>
            {preview&&<button onClick={confirmUpload} disabled={uploading} style={{ ...btnP,background:"#0F6E56",opacity:uploading?0.7:1 }}><i className="ti ti-check"/>{uploading?"Zapisuję...":`Zatwierdź — ${preview.length} rekordów`}</button>}
            {preview&&<button onClick={()=>{setPreview(null);setXlsxName(null);}} style={btnS}>Anuluj</button>}
          </div>
          {uploadMsg&&<div style={{ padding:"10px 14px",borderRadius:"8px",background:uploadMsg.ok?"#d1fae5":"#fee2e2",color:uploadMsg.ok?"#065f46":"#991b1b",fontSize:"13px",marginBottom:"14px" }}>{uploadMsg.ok?"✓":"⚠"} {uploadMsg.text}</div>}
          {preview&&(
            <div>
              <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 8px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Podgląd — pierwsze 5 wierszy</p>
              <div style={{ border:"1px solid #e5e7eb",borderRadius:"8px",overflow:"hidden" }}>
                {preview.slice(0,5).map((r,i)=>(
                  <div key={i} style={{ display:"flex",gap:"10px",padding:"8px 12px",borderBottom:i<4?"1px solid #e5e7eb":"none",fontSize:"12px",alignItems:"center" }}>
                    <span style={{ color:"#6b7280",minWidth:"90px",flexShrink:0 }}>{r.przypisany||"—"}</span>
                    <span style={{ color:"#6b7280",minWidth:"72px",flexShrink:0 }}>{r.data_wprow||"—"}</span>
                    <span style={{ flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{r.model||"—"}</span>
                    <span style={{ padding:"2px 8px",borderRadius:"10px",fontSize:"11px",flexShrink:0,background:r.status==="03. Realizacja umów"?"#d1fae5":"#f3f4f6",color:r.status==="03. Realizacja umów"?"#065f46":"#6b7280" }}>{r.status?r.status.replace(/^\d+\.\s/,""):"—"}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ── TARGETS ── */}
      {adminTab==="targets"&&(
        <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"20px" }}>
          <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 20px" }}>Miesięczne plany sprzedaży per handlowiec.</p>
          {SP.map(sp=>(
            <div key={sp} style={{ display:"flex",alignItems:"center",gap:"12px",marginBottom:"16px" }}>
              <div style={{ width:"36px",height:"36px",borderRadius:"50%",background:SP_COLOR[sp]+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:"500",color:SP_COLOR[sp],flexShrink:0 }}>{SP_INIT[sp]}</div>
              <span style={{ flex:1,fontSize:"15px",fontWeight:"500" }}>{SP_LABEL[sp]}</span>
              <span style={{ fontSize:"13px",color:"#6b7280",minWidth:"110px" }}>Aktualnie: <strong>{tMap[sp]||10} szt.</strong></span>
              <input type="number" min="0" max="99" value={editTargets?(editTargets[sp]??(tMap[sp]||10)):(tMap[sp]||10)} onChange={e=>setEditTargets(p=>({...(p||tMap),[sp]:+e.target.value}))} style={{ ...inpM,width:"72px",textAlign:"center" }}/>
              <span style={{ fontSize:"13px",color:"#6b7280" }}>szt.</span>
            </div>
          ))}
          <div style={{ display:"flex",gap:"8px",marginTop:"8px" }}>
            <button onClick={saveTargets} disabled={savingTargets||!editTargets} style={{ ...btnP,opacity:(!editTargets||savingTargets)?0.6:1 }}><i className="ti ti-device-floppy"/>{savingTargets?"Zapisuję...":"Zapisz plany"}</button>
            {editTargets&&<button onClick={()=>setEditTargets(null)} style={btnS}>Anuluj</button>}
          </div>

          {/* Records */}
          <div style={{ marginTop:"24px",paddingTop:"20px",borderTop:"1px solid #e5e7eb" }}>
            <p style={{ fontSize:"14px",fontWeight:"500",margin:"0 0 4px" }}>Rekordy wszech czasów</p>
            <p style={{ fontSize:"12px",color:"#6b7280",margin:"0 0 16px" }}>Wpisz ręcznie — uwzględnia miesiące sprzed eksportu (np. styczeń–luty)</p>
            {SP.map(sp=>{
              const cur = rMap[sp]||{ count:0, month:"" };
              const eVal = editRecords?.[sp];
              return (
                <div key={sp} style={{ display:"flex",alignItems:"center",gap:"12px",marginBottom:"14px" }}>
                  <div style={{ width:"36px",height:"36px",borderRadius:"50%",background:SP_COLOR[sp]+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:"500",color:SP_COLOR[sp],flexShrink:0 }}>{SP_INIT[sp]}</div>
                  <span style={{ flex:1,fontSize:"15px",fontWeight:"500" }}>{SP_LABEL[sp]}</span>
                  <span style={{ fontSize:"13px",color:"#6b7280",minWidth:"150px" }}>Aktualnie: <strong>{cur.count} zam.</strong>{cur.month?` · ${cur.month}`:""}</span>
                  <input type="number" min="0" placeholder="szt." value={eVal?.count??cur.count} onChange={e=>setEditRecords(p=>({...(p||{}), [sp]:{ count:+e.target.value, month:(p?.[sp]?.month??cur.month) }}))} style={{ ...inpM,width:"72px",textAlign:"center" }}/>
                  <input type="text" placeholder="np. Styczeń 2025" value={eVal?.month??cur.month} onChange={e=>setEditRecords(p=>({...(p||{}), [sp]:{ count:(p?.[sp]?.count??cur.count), month:e.target.value }}))} style={{ ...inpM,width:"140px" }}/>
                </div>
              );
            })}
            <div style={{ display:"flex",gap:"8px",marginTop:"4px" }}>
              <button onClick={saveRecords} disabled={savingRecords||!editRecords} style={{ ...btnP,opacity:(!editRecords||savingRecords)?0.6:1 }}><i className="ti ti-device-floppy"/>{savingRecords?"Zapisuję...":"Zapisz rekordy"}</button>
              {editRecords&&<button onClick={()=>setEditRecords(null)} style={btnS}>Anuluj</button>}
            </div>
          </div>
        </div>
      )}

      {/* ── ZAMÓWIENIA ── */}
      {adminTab==="orders"&&kpis&&(
        <div style={{ display:"flex",flexDirection:"column",gap:"12px" }}>
          <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 4px" }}>Bieżący miesiąc — {MONTHS[selMonth-1]} {selYear}</p>
          {SP.map(sp=>{
            const k = kpis[sp];
            const spManual = manual.filter(o=>{ const d=new Date(o.date); return o.sp===sp && d.getMonth()+1===selMonth && d.getFullYear()===selYear && o.model!=="__KOREKTA__"; });
            const spKorekta = manual.filter(o=>{ const d=new Date(o.date); return o.sp===sp && d.getMonth()+1===selMonth && d.getFullYear()===selYear && o.model==="__KOREKTA__"; });
            const isEditing = editingTotalSP===sp;
            return (
              <div key={sp} style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"20px",borderLeft:`4px solid ${SP_COLOR[sp]}` }}>
                {/* Header */}
                <div style={{ display:"flex",alignItems:"center",gap:"12px",marginBottom:"14px" }}>
                  <div style={{ width:"36px",height:"36px",borderRadius:"50%",background:SP_COLOR[sp]+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"13px",fontWeight:"500",color:SP_COLOR[sp],flexShrink:0 }}>{SP_INIT[sp]}</div>
                  <div style={{ flex:1 }}>
                    <p style={{ fontSize:"16px",fontWeight:"500",margin:"0 0 2px" }}>{SP_LABEL[sp]}</p>
                    <p style={{ fontSize:"12px",color:"#6b7280",margin:0 }}>{k.dmsOrders} z DMS · {spManual.length} ręczne{spKorekta.length>0?` · ${spKorekta.length} korekta`:""}</p>
                  </div>
                  {/* Total + edit */}
                  {isEditing?(
                    <div style={{ display:"flex",gap:"6px",alignItems:"center" }}>
                      <input type="number" min={k.dmsOrders} value={editTotalVal} onChange={e=>setEditTotalVal(+e.target.value)} style={{ ...inpM,width:"72px",textAlign:"center",fontSize:"20px",fontWeight:"500",padding:"6px 8px" }} autoFocus/>
                      <button onClick={()=>saveMonthTotal(sp,editTotalVal)} style={{ padding:"6px 12px",borderRadius:"8px",border:"none",background:"#185FA5",color:"#fff",cursor:"pointer",fontSize:"13px",fontWeight:"500" }}>Zapisz</button>
                      <button onClick={()=>setEditingTotalSP(null)} style={{ padding:"6px 10px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"13px" }}>✕</button>
                    </div>
                  ):(
                    <div style={{ display:"flex",alignItems:"center",gap:"8px" }}>
                      <span style={{ fontSize:"36px",fontWeight:"500",color:SP_COLOR[sp] }}>{k.totalOrders}</span>
                      <button onClick={()=>{setEditingTotalSP(sp);setEditTotalVal(k.totalOrders);}} style={{ padding:"4px 10px",borderRadius:"6px",border:"1px solid #d1d5db",background:"#fff",cursor:"pointer",fontSize:"12px",color:"#374151",display:"flex",alignItems:"center",gap:"4px" }}>
                        <i className="ti ti-edit" style={{ fontSize:"12px" }}/>Edytuj
                      </button>
                    </div>
                  )}
                </div>

                {/* Manual orders list */}
                {spManual.length>0&&(
                  <div style={{ borderTop:"1px solid #f3f4f6",paddingTop:"10px",marginBottom:"10px" }}>
                    {spManual.map(o=>(
                      <div key={o.id} style={{ display:"flex",alignItems:"center",gap:"10px",padding:"5px 0",fontSize:"13px" }}>
                        <span style={{ flex:2,fontWeight:"500",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{o.model}</span>
                        <span style={{ flex:2,color:"#6b7280",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{o.client}</span>
                        <span style={{ fontSize:"12px",color:"#9ca3af",whiteSpace:"nowrap" }}>{o.date}</span>
                        <button onClick={()=>deleteOrder(o.id)} style={{ background:"none",border:"none",cursor:"pointer",color:"#dc2626",padding:"2px",flexShrink:0 }}><i className="ti ti-trash" style={{ fontSize:"14px" }}/></button>
                      </div>
                    ))}
                    {spKorekta.length>0&&spKorekta.map(o=>(
                      <div key={o.id} style={{ display:"flex",alignItems:"center",gap:"10px",padding:"5px 0",fontSize:"13px" }}>
                        <span style={{ flex:2,color:"#6b7280",fontStyle:"italic" }}>Korekta managera</span>
                        <span style={{ flex:2 }}/>
                        <span style={{ fontSize:"12px",color:"#9ca3af",whiteSpace:"nowrap" }}>{o.date}</span>
                        <button onClick={()=>deleteOrder(o.id)} style={{ background:"none",border:"none",cursor:"pointer",color:"#dc2626",padding:"2px",flexShrink:0 }}><i className="ti ti-trash" style={{ fontSize:"14px" }}/></button>
                      </div>
                    ))}
                  </div>
                )}

                {/* Add button */}
                <button onClick={()=>{setAddForm(f=>({...f,sp}));setModal("add");}} style={{ fontSize:"13px",color:"#185FA5",background:"none",border:"1px dashed #93c5fd",borderRadius:"6px",padding:"6px 12px",cursor:"pointer",width:"100%" }}>
                  + Dodaj zamówienie
                </button>
              </div>
            );
          })}
        </div>
      )}

      {/* ── CELE ── */}
      {adminTab==="cele"&&(
        <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"20px" }}>
          <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 4px" }}>Cele zespołowe. Zliczają wszystkie zamówienia działu (DMS + ręczne) od dnia ustawienia celu — biegną przez kolejne miesiące, aż wyłączysz cel.</p>
          <p style={{ fontSize:"12px",color:"#9ca3af",margin:"0 0 18px" }}>Aktywne cele pokazują się na dużym ekranie na górze dashboardu.</p>

          {/* Nowy cel */}
          <div style={{ background:"#f9fafb",border:"1px solid #e5e7eb",borderRadius:"10px",padding:"16px",marginBottom:"20px" }}>
            <p style={{ fontSize:"13px",fontWeight:"500",margin:"0 0 12px",display:"flex",alignItems:"center",gap:"6px" }}><i className="ti ti-flag" style={{ fontSize:"15px",color:"#185FA5" }}/>Nowy cel</p>
            <div style={{ display:"flex",gap:"10px",flexWrap:"wrap",alignItems:"flex-end" }}>
              <div style={{ flex:2,minWidth:"160px" }}>
                <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 5px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Nazwa</p>
                <input value={newAch.name} onChange={e=>setNewAch(p=>({...p,name:e.target.value}))} placeholder="np. 🍗 30 zamówień = kebab dla działu" style={inpM}/>
              </div>
              <div style={{ width:"90px" }}>
                <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 5px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Ile zam.</p>
                <input type="number" min="1" value={newAch.target} onChange={e=>setNewAch(p=>({...p,target:e.target.value}))} style={{ ...inpM,textAlign:"center" }}/>
              </div>
            </div>
            <div style={{ marginTop:"10px" }}>
              <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 5px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Nagroda / opis (opcjonalnie)</p>
              <input value={newAch.description} onChange={e=>setNewAch(p=>({...p,description:e.target.value}))} placeholder="np. Stawiam wszystkim kebaba!" style={inpM}/>
            </div>
            <button onClick={addAchievement} disabled={savingAch} style={{ ...btnP,marginTop:"14px",opacity:savingAch?0.6:1 }}><i className="ti ti-plus"/>{savingAch?"Dodaję...":"Dodaj cel"}</button>
          </div>

          {/* Lista celów */}
          {achievements.length===0?(
            <p style={{ fontSize:"13px",color:"#9ca3af",textAlign:"center",padding:"24px" }}>Brak celów. Dodaj pierwszy powyżej.</p>
          ):(
            achievements.slice().sort((a,b)=>(b.active-a.active)||((b.created_at||"").localeCompare(a.created_at||""))).map(ach=>{
              const prog = achProgress(ach);
              const done = prog>=ach.target;
              const pct = Math.min(100,Math.round(prog/ach.target*100));
              return (
                <div key={ach.id} style={{ display:"flex",alignItems:"center",gap:"14px",padding:"14px 0",borderBottom:"1px solid #f3f4f6",opacity:ach.active?1:0.5 }}>
                  <div style={{ flex:1,minWidth:0 }}>
                    <p style={{ fontSize:"14px",fontWeight:"500",margin:"0 0 2px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{ach.name}{!ach.active&&<span style={{ fontSize:"11px",color:"#9ca3af",fontWeight:"400" }}> · wyłączony</span>}</p>
                    {ach.description&&<p style={{ fontSize:"12px",color:"#6b7280",margin:"0 0 6px" }}>{ach.description}</p>}
                    <div style={{ display:"flex",alignItems:"center",gap:"8px" }}>
                      <div style={{ flex:1,maxWidth:"240px",height:"5px",background:"#e5e7eb",borderRadius:"3px",overflow:"hidden" }}>
                        <div style={{ height:"100%",width:`${pct}%`,background:done?"#3B6D11":"#185FA5",borderRadius:"3px" }}/>
                      </div>
                      <span style={{ fontSize:"12px",fontWeight:"500",color:done?"#3B6D11":"#374151",whiteSpace:"nowrap" }}>{done?"✓ Osiągnięty!":`${prog} / ${ach.target}`}</span>
                    </div>
                  </div>
                  <button onClick={()=>toggleAchievement(ach.id,ach.active)} style={{ ...btnS,padding:"6px 12px",fontSize:"13px",whiteSpace:"nowrap" }}>{ach.active?"Wyłącz":"Włącz"}</button>
                  <button onClick={()=>deleteAchievement(ach.id)} style={{ background:"none",border:"none",cursor:"pointer",color:"#dc2626",padding:"4px",flexShrink:0 }}><i className="ti ti-trash" style={{ fontSize:"16px" }}/></button>
                </div>
              );
            })
          )}
        </div>
      )}
    </div>
  );

  // ══════════════════════════════════════════════════════════
  // DASHBOARD
  return (
    <div style={{ fontFamily:"inherit" }}>
      <div style={{ padding:"20px",maxWidth:"920px",margin:"0 auto" }}>

        {/* Header */}
        <div style={{ display:"flex",alignItems:"center",gap:"10px",marginBottom:"20px",flexWrap:"wrap" }}>
          <div style={{ flex:1,minWidth:"160px" }}>
            <p style={{ fontSize:"17px",fontWeight:"500",margin:"0 0 1px" }}>Wyniki zespołu</p>
            <p style={{ fontSize:"12px",color:"#6b7280",margin:0 }}>Ford Budmat Auto · Płock</p>
          </div>
          <div style={{ display:"flex",gap:"8px",flexWrap:"wrap",alignItems:"center" }}>
            <select value={selMonth} onChange={e=>setSelMonth(+e.target.value)} style={{ padding:"7px 10px",borderRadius:"8px",border:"1px solid #d1d5db",background:"#fff",color:"#111827",fontSize:"14px",cursor:"pointer" }}>
              {MONTHS.map((m,i)=><option key={i} value={i+1}>{m} {selYear}</option>)}
            </select>
            <button onClick={load} style={{ padding:"7px 10px",borderRadius:"8px",border:"1px solid #d1d5db",cursor:"pointer",background:"#fff",color:"#374151",display:"flex",alignItems:"center" }}><i className="ti ti-refresh" style={{ fontSize:"15px" }}/></button>
            <button onClick={()=>setModal("add")} style={{ ...btnP,padding:"7px 14px" }}><i className="ti ti-plus" style={{ fontSize:"15px" }}/>Zamówienie</button>
            <button onClick={()=>setPage("login")} style={{ padding:"7px 10px",borderRadius:"8px",border:"1px solid #d1d5db",cursor:"pointer",background:"#fff",color:"#6b7280",display:"flex",alignItems:"center" }}><i className="ti ti-settings" style={{ fontSize:"15px" }}/></button>
          </div>
        </div>

        {loading?(
          <div style={{ textAlign:"center",padding:"80px",color:"#6b7280" }}>
            <i className="ti ti-loader-2" style={{ fontSize:"36px",display:"block",marginBottom:"12px",opacity:0.35 }}/>
            <p style={{ margin:0 }}>Łączenie z bazą...</p>
          </div>
        ):dbErr?(
          <div style={{ textAlign:"center",padding:"60px" }}>
            <i className="ti ti-database-off" style={{ fontSize:"36px",display:"block",marginBottom:"12px",color:"#dc2626",opacity:0.6 }}/>
            <p style={{ fontWeight:"500",color:"#dc2626",margin:"0 0 6px" }}>Błąd połączenia</p>
            <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 16px" }}>{dbErr}</p>
            <button onClick={load} style={btnS}>Spróbuj ponownie</button>
          </div>
        ):leads.length===0?(
          <div style={{ textAlign:"center",padding:"64px",border:"1px dashed #e5e7eb",borderRadius:"12px" }}>
            <i className="ti ti-database-import" style={{ fontSize:"40px",display:"block",marginBottom:"12px",opacity:0.3 }}/>
            <p style={{ fontSize:"16px",fontWeight:"500",margin:"0 0 6px" }}>Baza jest pusta</p>
            <p style={{ fontSize:"13px",color:"#6b7280",margin:"0 0 16px" }}>Zaloguj się jako admin i wgraj export</p>
            <button onClick={()=>setPage("login")} style={btnP}>Przejdź do panelu admina</button>
          </div>
        ):(<>

          {/* ── CELE / ACHIEVEMENTS ── */}
          {achievements.filter(a=>a.active).length>0&&(
            <div style={{ display:"flex",flexDirection:"column",gap:"10px",marginBottom:"14px" }}>
              {achievements.filter(a=>a.active).map(ach=>{
                const prog = achProgress(ach);
                const done = prog>=ach.target;
                const pct = Math.min(100,Math.round(prog/ach.target*100));
                return (
                  <div key={ach.id} style={{ background:done?"#ECFccb":"#fff", border:`1px solid ${done?"#84cc16":"#e5e7eb"}`, borderRadius:"12px", padding:"14px 20px", display:"flex", alignItems:"center", gap:"18px", borderLeft:`4px solid ${done?"#3B6D11":"#185FA5"}` }}>
                    <div style={{ flex:1,minWidth:0,textAlign:"center" }}>
                      <p style={{ fontSize:"17px",fontWeight:"600",margin:"0 0 3px",color:done?"#3B6D11":"#111827" }}>{ach.name}</p>
                      {ach.description&&<p style={{ fontSize:"13px",color:done?"#3f6212":"#6b7280",margin:0 }}>{ach.description}</p>}
                    </div>
                    <div style={{ display:"flex",alignItems:"center",gap:"12px",flexShrink:0,minWidth:"280px" }}>
                      {done?(
                        <span style={{ fontSize:"18px",fontWeight:"700",color:"#3B6D11",whiteSpace:"nowrap" }}>✓ OSIĄGNIĘTE!</span>
                      ):(<>
                        <div style={{ flex:1,height:"8px",background:"#e5e7eb",borderRadius:"4px",overflow:"hidden",minWidth:"160px" }}>
                          <div style={{ height:"100%",width:`${pct}%`,background:"#185FA5",borderRadius:"4px",transition:"width 0.4s" }}/>
                        </div>
                        <span style={{ fontSize:"22px",fontWeight:"600",color:"#185FA5",whiteSpace:"nowrap" }}>{prog}<span style={{ fontSize:"14px",color:"#6b7280",fontWeight:"400" }}> / {ach.target}</span></span>
                      </>)}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* ── TEAM SUMMARY ── */}
          <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"20px",marginBottom:"14px" }}>
            <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 16px",textTransform:"uppercase",letterSpacing:"0.4px" }}>Zespół · {MONTHS[selMonth-1]} {selYear}</p>

            <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:0,paddingBottom:"16px",marginBottom:"16px",borderBottom:"1px solid #e5e7eb" }}>
              <div style={{ borderRight:"1px solid #e5e7eb",paddingRight:"20px" }}>
                <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 6px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Zamówienia / Plan</p>
                <div style={{ display:"flex",alignItems:"flex-end",gap:"6px",marginBottom:"8px" }}>
                  <span style={{ fontSize:"44px",fontWeight:"500",lineHeight:1,color:planColor(teamPlan) }}>{teamOrders}</span>
                  <span style={{ fontSize:"14px",color:"#6b7280",paddingBottom:"6px" }}>/ {teamTarget}</span>
                </div>
                <div style={{ display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px" }}>
                  <div style={{ flex:1,height:"5px",background:"#e5e7eb",borderRadius:"3px",overflow:"hidden" }}>
                    <div style={{ height:"100%",width:`${Math.min(teamPlan,100)}%`,background:planColor(teamPlan),borderRadius:"3px" }}/>
                  </div>
                  <span style={{ fontSize:"13px",fontWeight:"500",color:planColor(teamPlan),whiteSpace:"nowrap" }}>{teamPlan}%</span>
                </div>
                {isCurMonth&&daysLeft>0&&(
                  <p style={{ fontSize:"12px",color:teamNeeded===0?"#059669":"#6b7280",margin:0 }}>
                    {teamNeeded===0?"✓ Cel osiągnięty!":`Zostało ${daysLeft} ${daysLeft===1?"dzień":"dni"} · potrzeba ${teamNeeded} ${teamNeeded===1?"zamówienia":"zamówień"}`}
                  </p>
                )}
              </div>
              <div style={{ paddingLeft:"20px" }}>
                <p style={{ fontSize:"11px",color:"#6b7280",margin:"0 0 6px",textTransform:"uppercase",letterSpacing:"0.3px" }}>Pipeline aktywny</p>
                <p style={{ fontSize:"44px",fontWeight:"500",lineHeight:1,margin:"0 0 6px",color:"#BA7517" }}>{teamPipeline}</p>
                <p style={{ fontSize:"12px",color:"#6b7280",margin:0 }}>oferty status 02 · bieżący + 2 miesiące</p>
              </div>
            </div>

          </div>

          {/* ── PERSON CARDS ── */}
          <div style={{ display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:"12px",marginBottom:"14px" }}>
            {ranked.map(sp=>{
              const k=kpis[sp]; const g=game?.[sp]; const c=SP_COLOR[sp];
              const bc=planColor(k.planPct);
              const isChaser = k.rank===2;
              const isLeader = k.rank===1;
              const rec = rMap[sp]||{ count:0, month:"" };
              const dispRecord = Math.max(rec.count, g?.pBest||0);
              const dispRecordLabel = rec.count>=g?.pBest ? rec.month : g?.pBestLabel;
              const nearRecord = dispRecord>0 && k.totalOrders>=dispRecord-1 && k.totalOrders<dispRecord;
              const beatsRecord = dispRecord>0 && k.totalOrders>=dispRecord;

              return (
                <div key={sp} style={{ background: isLeader?"#fffbeb":"#fff", border:`1px solid ${isLeader?"#fbbf24":"#e5e7eb"}`, borderRadius:"12px", padding:"18px", borderTop:`3px solid ${isLeader?"#BA7517":c}` }}>

                  {/* Name + rank */}
                  <div style={{ display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"16px" }}>
                    <div style={{ display:"flex",alignItems:"center",gap:"10px" }}>
                      <div style={{ width:"36px",height:"36px",borderRadius:"50%",background:(isLeader?"#BA7517":c)+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"13px",fontWeight:"500",color:isLeader?"#BA7517":c,flexShrink:0 }}>{SP_INIT[sp]}</div>
                      <p style={{ fontSize:"18px",fontWeight:"500",margin:0 }}>{SP_LABEL[sp]}</p>
                    </div>
                    <span style={{ fontSize:"11px",fontWeight:"600",padding:"4px 10px",borderRadius:"20px",background:isLeader?"#BA7517":"#F1EFE8",color:isLeader?"#fff":"#5F5E5A",display:"flex",alignItems:"center",gap:"4px" }}>
                      {isLeader?"👑":""} {isLeader?"Lider salonu":"#2"}
                    </span>
                  </div>

                  {/* Big number */}
                  <div style={{ display:"flex",alignItems:"flex-end",gap:"8px",marginBottom:"14px" }}>
                    <span style={{ fontSize:"52px",fontWeight:"500",lineHeight:1,color:isLeader?"#BA7517":c }}>{k.totalOrders}</span>
                    <div style={{ paddingBottom:"7px",lineHeight:1.4 }}>
                      <span style={{ fontSize:"13px",color:"#6b7280" }}>/ {k.target}<br/>zamówień</span>
                      {k.hasOverride&&<div style={{ fontSize:"10px",color:"#f59e0b",marginTop:"2px" }}>✏ DMS nadpisany ({k.dmsOrders}→{k.effectiveDms})</div>}
                    </div>
                  </div>

                  {/* Progress */}
                  <div style={{ marginBottom:"14px" }}>
                    <div style={{ display:"flex",justifyContent:"space-between",marginBottom:"5px" }}>
                      <span style={{ fontSize:"12px",color:"#6b7280" }}>Realizacja planu</span>
                      <span style={{ fontSize:"12px",fontWeight:"500",color:bc }}>{k.planPct}%</span>
                    </div>
                    <div style={{ height:"5px",background:"#e5e7eb",borderRadius:"3px",overflow:"hidden" }}>
                      <div style={{ height:"100%",width:`${Math.min(k.planPct,100)}%`,background:isLeader?"#BA7517":bc,borderRadius:"3px" }}/>
                    </div>
                  </div>

                  {/* Stats */}
                  <div style={{ display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"12px" }}>
                    <div style={{ background:isLeader?"#fef3c7":"#f9fafb",borderRadius:"8px",padding:"10px 6px",textAlign:"center" }}>
                      <p style={{ fontSize:"20px",fontWeight:"500",margin:"0 0 3px" }}>{k.convMon}%</p>
                      <p style={{ fontSize:"9px",color:"#6b7280",margin:0,textTransform:"uppercase",letterSpacing:"0.3px",lineHeight:1.4 }}>konwersja<br/>miesięczna</p>
                    </div>
                    <div onClick={k.pipeline>0?()=>{setPipelineSP(sp);setModal("pipeline");}:undefined}
                      style={{ background:isLeader?"#fef3c7":"#f9fafb",borderRadius:"8px",padding:"10px 6px",textAlign:"center",cursor:k.pipeline>0?"pointer":"default",border:k.pipeline>0?"1px solid #e5e7eb":"1px solid transparent" }}>
                      <p style={{ fontSize:"20px",fontWeight:"500",margin:"0 0 3px",color:k.pipeline>0?"#BA7517":"inherit" }}>{k.pipeline}</p>
                      <p style={{ fontSize:"9px",color:"#6b7280",margin:0,textTransform:"uppercase",letterSpacing:"0.3px",lineHeight:1.4 }}>pipeline{k.pipeline>0?" →":""}<br/>&nbsp;</p>
                    </div>
                  </div>

                  {/* Streak + pace */}
                  {g&&(
                    <div style={{ display:"flex",gap:"8px",marginBottom:"10px",flexWrap:"wrap" }}>
                      {g.streak>=2&&<span style={{ fontSize:"11px",display:"flex",alignItems:"center",gap:"3px",color:"#ea580c",fontWeight:"500" }}>🔥 {g.streak} dni</span>}
                      {isCurMonth&&<span style={{ fontSize:"11px",display:"flex",alignItems:"center",gap:"3px",color:g.isOnPace?"#059669":"#dc2626",fontWeight:"500" }}>{g.isOnPace?"▲":"▼"} {g.curPace.toFixed(1)}/dzień</span>}
                      {isCurMonth&&daysLeft>0&&g.needed>0&&<span style={{ fontSize:"11px",color:"#6b7280" }}>{g.needed} do celu · {daysLeft} dni</span>}
                      {isCurMonth&&g.needed===0&&<span style={{ fontSize:"11px",color:"#059669",fontWeight:"500" }}>✓ Cel osiągnięty!</span>}
                    </div>
                  )}

                  {/* Record — leader */}
                  {isLeader&&dispRecord>0&&(
                    <div style={{ padding:"8px 12px",background:beatsRecord?"#d1fae5":nearRecord?"#fef3c7":"#fffbeb",border:`1px solid ${beatsRecord?"#6ee7b7":nearRecord?"#fbbf24":"#fde68a"}`,borderRadius:"8px",marginBottom:"10px",fontSize:"12px" }}>
                      {beatsRecord
                        ? <span style={{ color:"#065f46",fontWeight:"600" }}>🏆 NOWY REKORD SALONU! (poprzedni: {dispRecord} — {dispRecordLabel})</span>
                        : nearRecord
                          ? <span style={{ color:"#92400e",fontWeight:"500" }}>🎯 Jedno do rekordu! Rekord: {dispRecord} zam. — {dispRecordLabel}</span>
                          : <span style={{ color:"#92400e" }}>👑 Rekord salonu: <strong>{dispRecord} zamówień</strong> — {dispRecordLabel}</span>
                      }
                    </div>
                  )}

                  {/* Personal best — chaser */}
                  {isChaser&&dispRecord>0&&(
                    <div style={{ padding:"6px 10px",background:"#f0f9ff",borderRadius:"6px",marginBottom:"10px",fontSize:"11px",color:"#0369a1" }}>
                      🏅 Twój rekord: <strong>{dispRecord} zamówień</strong>{dispRecordLabel?` — ${dispRecordLabel}`:""}
                      {beatsRecord&&<span style={{ color:"#059669",fontWeight:"500" }}> — bijesz go teraz!</span>}
                      {nearRecord&&<span style={{ color:"#ea580c",fontWeight:"500" }}> — jeszcze jedno!</span>}
                    </div>
                  )}

                  {/* Weekly target chaser */}
                  {g&&isChaser&&(
                    <div style={{ padding:"6px 10px",background:"#fafafa",borderRadius:"6px",marginBottom:"10px",fontSize:"11px",color:"#374151",display:"flex",justifyContent:"space-between",alignItems:"center" }}>
                      <span>Cel tygodniowy</span>
                      <span style={{ fontWeight:"500",color:g.weeklyOrders>=g.weeklyTarget?"#059669":"#111827" }}>{g.weeklyOrders} / {g.weeklyTarget} {g.weeklyOrders>=g.weeklyTarget?"✓":""}</span>
                    </div>
                  )}

                  {/* Badges */}
                  {g&&g.badges.length>0&&(
                    <div style={{ display:"flex",flexWrap:"wrap",gap:"5px",marginBottom:"10px" }}>
                      {g.badges.map((b,i)=><Badge key={i} {...b}/>)}
                    </div>
                  )}

                  <p style={{ fontSize:"11px",color:"#6b7280",margin:0 }}>
                    {k.totalOrders} z {k.curLeads} leadów w {MONTHS[selMonth-1].toLowerCase()}
                    {k.manOrders>0&&<span style={{ color:"#059669" }}> · {k.manOrders} ręcznych</span>}
                  </p>
                </div>
              );
            })}
          </div>

          {/* Manual orders */}
          {manualThisMonth.length>0&&(
            <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"16px" }}>
              <p style={{ fontSize:"11px",fontWeight:"500",margin:"0 0 10px",color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.3px" }}>Zamówienia dodane ręcznie</p>
              {manualThisMonth.map(o=>(
                <div key={o.id} style={{ display:"flex",alignItems:"center",gap:"12px",padding:"7px 0",borderBottom:"1px solid #f3f4f6",fontSize:"13px" }}>
                  <span style={{ width:"8px",height:"8px",borderRadius:"50%",background:SP_COLOR[o.sp]||"#888",flexShrink:0 }}/>
                  <span style={{ minWidth:"52px",color:"#6b7280" }}>{SP_LABEL[o.sp]||o.sp}</span>
                  <span style={{ flex:2,color:"#6b7280",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{o.client}</span>
                  <span style={{ fontSize:"12px",color:"#9ca3af",whiteSpace:"nowrap" }}>{o.date}</span>
                  <button onClick={()=>deleteOrder(o.id)} style={{ background:"none",border:"none",cursor:"pointer",color:"#6b7280",padding:"2px 4px",flexShrink:0 }}><i className="ti ti-trash" style={{ fontSize:"14px" }}/></button>
                </div>
              ))}
            </div>
          )}

          {/* Inne */}
          {inneThisMonth.length>0&&(
            <div style={{ background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"16px",marginTop:"12px" }}>
              <div style={{ display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:"10px" }}>
                <p style={{ fontSize:"11px",fontWeight:"500",margin:0,color:"#6b7280",textTransform:"uppercase",letterSpacing:"0.3px" }}>Inne</p>
                <span style={{ fontSize:"20px",fontWeight:"500",color:"#374151" }}>{inneThisMonth.length}</span>
              </div>
              {inneThisMonth.map(o=>(
                <div key={o.id} style={{ display:"flex",alignItems:"center",gap:"12px",padding:"7px 0",borderBottom:"1px solid #f3f4f6",fontSize:"13px" }}>
                  <span style={{ flex:2,color:"#374151",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{o.client||"—"}</span>
                  <span style={{ fontSize:"12px",color:"#9ca3af",whiteSpace:"nowrap" }}>{o.date}</span>
                  <button onClick={()=>deleteOrder(o.id)} style={{ background:"none",border:"none",cursor:"pointer",color:"#6b7280",padding:"2px 4px",flexShrink:0 }}><i className="ti ti-trash" style={{ fontSize:"14px" }}/></button>
                </div>
              ))}
            </div>
          )}
        </>)}
      </div>

      {/* MODAL: Add order */}
      {modal==="add"&&(
        <Modal onClose={()=>setModal(null)} title="Dodaj zamówienie" maxWidth="320px">
          <Field label="Handlowiec">
            <select value={addForm.sp} onChange={e=>setAddForm({...addForm,sp:e.target.value})} style={inpM}>
              {SP.map(s=><option key={s} value={s}>{SP_LABEL[s]}</option>)}
              <option value="Inne">Inne</option>
            </select>
          </Field>
          <Field label="Liczba zamówień">
            <input type="number" min="1" max="20" value={addForm.count} onChange={e=>setAddForm({...addForm,count:+e.target.value})} style={{ ...inpM, fontSize:"24px", fontWeight:"500", textAlign:"center", padding:"12px" }} autoFocus/>
          </Field>
          <Field label="Klient (opcjonalnie)">
            <input value={addForm.client} onChange={e=>setAddForm({...addForm,client:e.target.value})} placeholder="Nazwa firmy lub osoby" style={inpM}/>
          </Field>
          <div style={{ display:"flex",gap:"8px",marginTop:"20px" }}>
            <button onClick={addOrders} style={{ ...btnP,flex:1,justifyContent:"center" }}>Dodaj</button>
            <button onClick={()=>setModal(null)} style={btnS}>Anuluj</button>
          </div>
        </Modal>
      )}

      {/* MODAL: Pipeline */}
      {modal==="pipeline"&&pipelineSP&&kpis&&(
        <Modal onClose={()=>setModal(null)} title={`Pipeline · ${SP_LABEL[pipelineSP]}`} subtitle="Status 02 · bieżący + 2 poprzednie miesiące" maxWidth="540px">
          {kpis[pipelineSP].pipelineRows.length===0?<p style={{ textAlign:"center",color:"#6b7280",padding:"32px" }}>Brak leadów</p>:
            kpis[pipelineSP].pipelineRows.map((r,i)=>{
              const d=r.data_wprow?new Date(r.data_wprow):null;
              const isCur=d&&d.getMonth()+1===selMonth&&d.getFullYear()===selYear;
              return (
                <div key={i} style={{ padding:"10px 0",borderBottom:"1px solid #f3f4f6",display:"flex",gap:"12px",alignItems:"center" }}>
                  <span style={{ fontSize:"11px",color:isCur?"#185FA5":"#9ca3af",whiteSpace:"nowrap",minWidth:"46px" }}>{d?`${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getFullYear()).slice(2)}`:"—"}</span>
                  <div style={{ flex:1,minWidth:0 }}>
                    <p style={{ fontSize:"13px",fontWeight:"500",margin:"0 0 1px",color:"#111827",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap" }}>{r.kontrahent||"—"}</p>
                    <p style={{ fontSize:"11px",color:"#6b7280",margin:0 }}>{r.model||"Model nieznany"}</p>
                  </div>
                  {isCur&&<span style={{ fontSize:"11px",padding:"2px 7px",borderRadius:"20px",background:"#dbeafe",color:"#1e40af",flexShrink:0 }}>bieżący</span>}
                </div>
              );
            })
          }
        </Modal>
      )}
    </div>
  );
}
