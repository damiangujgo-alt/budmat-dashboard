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
    Object.entries(mTotals).forEach(([k,v]) => { if(v>pBest) {pBest=v; pBestLabel=k;} });

    const weeklyTarget = 3;
    stats[sp] = { streak, last5, ordersToday, weeklyOrders, weeklyTarget, pBest, pBestLabel, badges: [] };
    if(streak>=3) stats[sp].badges.push({emoji:"🔥",label:`Seria ${streak} dni`});
    if(ordersToday>=2) stats[sp].badges.push({emoji:"⚡",label:"Dzisiaj rozpęta forma"});
    if(weeklyOrders>=weeklyTarget*1.5) stats[sp].badges.push({emoji:"🚀",label:"Rozjedź tydzień!"});
  }
  return stats;
}

// ─── ACHIEVEMENT PROGRESS ENGINE ────────────────────────────
function computeAchievementProgress(leads, manual, achievement) {
  if (!achievement || !achievement.active) return null;
  
  const createdDate = new Date(achievement.created_at);
  const totalOrders = {};
  
  // Count all orders from all salespeople from created_at onwards
  leads.forEach(r => {
    if (!r.data_wprow || !isOrder(r)) return;
    const d = new Date(r.data_wprow);
    if (d >= createdDate) {
      totalOrders[r.przypisany] = (totalOrders[r.przypisany] || 0) + (r.count_override || 1);
    }
  });
  
  manual.forEach(o => {
    const d = new Date(o.date);
    if (d >= createdDate) {
      totalOrders[o.sp] = (totalOrders[o.sp] || 0) + 1;
    }
  });
  
  const totalCount = Object.values(totalOrders).reduce((a, b) => a + b, 0);
  const progress = Math.min(totalCount, achievement.target);
  const percentage = Math.round((progress / achievement.target) * 100);
  const isCompleted = totalCount >= achievement.target;
  
  return { progress, percentage, isCompleted, totalCount };
}

// ─── UI COMPONENTS ──────────────────────────────────────────
const inpM = { width:"100%", padding:"8px 12px", border:"1px solid #d1d5db", borderRadius:"6px", fontSize:"14px", fontFamily:"inherit", boxSizing:"border-box" };
const btnP = { display:"flex", alignItems:"center", gap:"6px", padding:"10px 16px", background:"#0f6e56", color:"#fff", border:"none", borderRadius:"8px", fontSize:"14px", fontWeight:"500", cursor:"pointer" };
const btnS = { display:"flex", alignItems:"center", gap:"6px", padding:"10px 16px", background:"#f3f4f6", color:"#374151", border:"none", borderRadius:"8px", fontSize:"14px", cursor:"pointer" };

function Badge({emoji, label}) { return <span style={{ display:"inline-flex", alignItems:"center", gap:"4px", padding:"4px 8px", background:"#fef3c7", borderRadius:"12px", fontSize:"11px", fontWeight:"500", color:"#92400e" }}>{emoji} {label}</span>; }
function Modal({onClose, title, subtitle, maxWidth="420px", children}) {
  return (
    <div style={{ position:"fixed", top:0, left:0, width:"100%", height:"100%", background:"rgba(0,0,0,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:1000 }} onClick={onClose}>
      <div style={{ background:"#fff", borderRadius:"12px", padding:"20px", maxWidth, width:"90%", boxShadow:"0 10px 40px rgba(0,0,0,0.2)" }} onClick={e=>e.stopPropagation()}>
        <div style={{ marginBottom:"16px" }}>
          <h2 style={{ margin:"0 0 4px", fontSize:"18px", fontWeight:"600", color:"#111827" }}>{title}</h2>
          {subtitle && <p style={{ margin:0, fontSize:"13px", color:"#6b7280" }}>{subtitle}</p>}
        </div>
        {children}
      </div>
    </div>
  );
}
function Field({label, children}) {
  return (
    <div style={{ marginBottom:"12px" }}>
      <label style={{ display:"block", fontSize:"12px", fontWeight:"500", color:"#374151", marginBottom:"4px", textTransform:"uppercase", letterSpacing:"0.3px" }}>{label}</label>
      {children}
    </div>
  );
}

// ─── MAIN APP ───────────────────────────────────────────────
export default function App() {
  const [leads, setLeads] = useState([]);
  const [manual, setManual] = useState([]);
  const [targets, setTargets] = useState([]);
  const [modal, setModal] = useState(null);
  const [kpis, setKpis] = useState(null);
  const [game, setGame] = useState(null);
  const [selMonth, setSelMonth] = useState(now.getMonth()+1);
  const [selYear, setSelYear] = useState(now.getFullYear());
  const [addForm, setAddForm] = useState({sp: SP[0], count: 1, client: ""});
  const [pipelineSP, setPipelineSP] = useState(null);
  // Achievements disabled temporarily
  const achievements = [];
  const setAchievements = () => {};
  const [showAchievementsModal, setShowAchievementsModal] = useState(false);

  useEffect(() => {
    const loadData = async () => {
      try {
        const [l, m, t, a] = await Promise.all([sbGet("leads"), sbGet("manual_orders"), sbGet("targets"), sbGet("achievements")]);
        setLeads(l);
        setManual(m);
        setTargets(t);
        setAchievements(a.filter(x => x.active));
        const k = computeKPIs(l, m, t, selMonth, selYear);
        const g = computeGame(l, m, k, selMonth, selYear);
        setKpis(k);
        setGame(g);
      } catch(err) { console.error(err); }
    };
    loadData();
    const iv = setInterval(loadData, 5000);
    return () => clearInterval(iv);
  }, [selMonth, selYear]);

  const addOrders = async () => {
    if (addForm.count < 1) return alert("Liczba musi być > 0");
    try {
      for (let i = 0; i < addForm.count; i++) {
        await sbPost("manual_orders", {sp: addForm.sp, client: addForm.client || null, date: todayStr(), model: null});
      }
      setAddForm({sp: SP[0], count: 1, client: ""});
      setModal(null);
      loadData();
    } catch(err) { console.error(err); }
  };

  const deleteOrder = async (id) => {
    try {
      await sbDelete("manual_orders", `id=eq.${id}`);
      loadData();
    } catch(err) { console.error(err); }
  };



  const manualThisMonth = manual.filter(o => { const d = new Date(o.date); return d.getMonth()+1 === selMonth && d.getFullYear() === selYear; });
  const inneThisMonth = leads.filter(r => r.status === "Inne" && r.data_wprow && new Date(r.data_wprow).getMonth()+1 === selMonth && new Date(r.data_wprow).getFullYear() === selYear);

  return (
    <div style={{ minHeight:"100vh", background:"#f9fafb", fontFamily:"system-ui, -apple-system, sans-serif", padding:"20px" }}>
      {/* HEADER */}
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:"20px", gap:"12px", flexWrap:"wrap" }}>
        <div style={{ display:"flex", gap:"8px", alignItems:"center", flex:1, minWidth:"200px" }}>
          <select value={selMonth} onChange={e => setSelMonth(+e.target.value)} style={{ ...inpM, flex:1 }}>
            {MONTHS.map((m, i) => <option key={i} value={i+1}>{m}</option>)}
          </select>
          <select value={selYear} onChange={e => setSelYear(+e.target.value)} style={{ ...inpM, flex:1 }}>
            {[2024, 2025, 2026, 2027].map(y => <option key={y} value={y}>{y}</option>)}
          </select>
        </div>
        <div style={{ display:"flex", gap:"8px" }}>
          <button onClick={() => setModal("add")} style={btnP}>+ Dodaj</button>
          <button onClick={() => setModal("admin")} style={btnS}>⚙️</button>
        </div>
      </div>

      {/* ACHIEVEMENTS SECTION - DISABLED TEMPORARILY */}

      {/* MAIN CONTENT */}
      {kpis && game && (
        <>
          {/* Rankings */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(240px, 1fr))", gap:"12px", marginBottom:"16px" }}>
            {SP.map((sp, idx) => {
              const k = kpis[sp];
              const g = game[sp];
              const isLeader = k.rank === 1;
              const isChaser = k.rank === 2;
              const dispRecord = k.pipelineRows.length > 0 ? Math.max(...[...k.pipelineRows, {data_wprow: null}].map(r => r.data_wprow || "").filter(Boolean).length) : 0;
              const dispRecordLabel = null;
              const beatsRecord = false;
              const nearRecord = false;

              return (
                <div key={sp} style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:"12px", padding:"16px" }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:"12px" }}>
                    <div style={{ flex:1 }}>
                      <p style={{ margin:"0 0 4px", fontSize:"16px", fontWeight:"600", color:"#111827" }}>{SP_LABEL[sp]}</p>
                      <div style={{ display:"flex", gap:"8px", alignItems:"center" }}>
                        <span style={{ fontSize:"28px", fontWeight:"700", color:SP_COLOR[sp] }}>{k.totalOrders}</span>
                        <span style={{ fontSize:"12px", color:"#6b7280" }}>/{k.target}</span>
                      </div>
                    </div>
                    <div style={{ textAlign:"right" }}>
                      <span style={{ display:"inline-block", padding:"4px 8px", background:"#f3f4f6", borderRadius:"6px", fontSize:"12px", fontWeight:"500", color:"#374151" }}>#{k.rank}</span>
                    </div>
                  </div>

                  {isLeader && dispRecord > 0 && (
                    <div style={{ padding:"8px 12px", background:"#d1fae5", border:"1px solid #6ee7b7", borderRadius:"8px", marginBottom:"10px", fontSize:"12px" }}>
                      <span style={{ color:"#065f46", fontWeight:"600" }}>🏆 NOWY REKORD SALONU!</span>
                    </div>
                  )}

                  {g && g.badges.length > 0 && (
                    <div style={{ display:"flex", flexWrap:"wrap", gap:"5px", marginBottom:"10px" }}>
                      {g.badges.map((b, i) => <Badge key={i} {...b} />)}
                    </div>
                  )}

                  <p style={{ fontSize:"11px", color:"#6b7280", margin:0 }}>
                    {k.totalOrders} z {k.curLeads} leadów w {MONTHS[selMonth-1].toLowerCase()}
                    {k.manOrders > 0 && <span style={{ color:"#059669" }}> · {k.manOrders} ręcznych</span>}
                  </p>
                </div>
              );
            })}
          </div>

          {/* Manual orders */}
          {manualThisMonth.length > 0 && (
            <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:"12px", padding:"16px" }}>
              <p style={{ fontSize:"11px", fontWeight:"500", margin:"0 0 10px", color:"#6b7280", textTransform:"uppercase", letterSpacing:"0.3px" }}>Zamówienia dodane ręcznie</p>
              {manualThisMonth.map(o => (
                <div key={o.id} style={{ display:"flex", alignItems:"center", gap:"12px", padding:"7px 0", borderBottom:"1px solid #f3f4f6", fontSize:"13px" }}>
                  <span style={{ width:"8px", height:"8px", borderRadius:"50%", background:SP_COLOR[o.sp] || "#888", flexShrink:0 }}/>
                  <span style={{ minWidth:"52px", color:"#6b7280" }}>{SP_LABEL[o.sp] || o.sp}</span>
                  <span style={{ flex:2, color:"#6b7280", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{o.client}</span>
                  <span style={{ fontSize:"12px", color:"#9ca3af", whiteSpace:"nowrap" }}>{o.date}</span>
                  <button onClick={() => deleteOrder(o.id)} style={{ background:"none", border:"none", cursor:"pointer", color:"#6b7280", padding:"2px 4px", flexShrink:0 }}>🗑️</button>
                </div>
              ))}
            </div>
          )}

          {/* Inne */}
          {inneThisMonth.length > 0 && (
            <div style={{ background:"#fff", border:"1px solid #e5e7eb", borderRadius:"12px", padding:"16px", marginTop:"12px" }}>
              <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:"10px" }}>
                <p style={{ fontSize:"11px", fontWeight:"500", margin:0, color:"#6b7280", textTransform:"uppercase", letterSpacing:"0.3px" }}>Inne</p>
                <span style={{ fontSize:"20px", fontWeight:"500", color:"#374151" }}>{inneThisMonth.length}</span>
              </div>
              {inneThisMonth.map(o => (
                <div key={o.id} style={{ display:"flex", alignItems:"center", gap:"12px", padding:"7px 0", borderBottom:"1px solid #f3f4f6", fontSize:"13px" }}>
                  <span style={{ flex:2, color:"#374151", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{o.client || "—"}</span>
                  <span style={{ fontSize:"12px", color:"#9ca3af", whiteSpace:"nowrap" }}>{o.date}</span>
                  <button onClick={() => deleteOrder(o.id)} style={{ background:"none", border:"none", cursor:"pointer", color:"#6b7280", padding:"2px 4px", flexShrink:0 }}>🗑️</button>
                </div>
              ))}
            </div>
          )}
        </>
      )}

      {/* MODAL: Add order */}
      {modal === "add" && (
        <Modal onClose={() => setModal(null)} title="Dodaj zamówienie" maxWidth="320px">
          <Field label="Handlowiec">
            <select value={addForm.sp} onChange={e => setAddForm({...addForm, sp: e.target.value})} style={inpM}>
              {SP.map(s => <option key={s} value={s}>{SP_LABEL[s]}</option>)}
            </select>
          </Field>
          <Field label="Liczba zamówień">
            <input type="number" min="1" max="20" value={addForm.count} onChange={e => setAddForm({...addForm, count: +e.target.value})} style={{ ...inpM, fontSize:"24px", fontWeight:"500", textAlign:"center", padding:"12px" }} autoFocus/>
          </Field>
          <Field label="Klient (opcjonalnie)">
            <input value={addForm.client} onChange={e => setAddForm({...addForm, client: e.target.value})} placeholder="Nazwa firmy lub osoby" style={inpM}/>
          </Field>
          <div style={{ display:"flex", gap:"8px", marginTop:"20px" }}>
            <button onClick={addOrders} style={{ ...btnP, flex:1, justifyContent:"center" }}>Dodaj</button>
            <button onClick={() => setModal(null)} style={{...btnS, flex:1, justifyContent:"center"}}>Anuluj</button>
          </div>
        </Modal>
      )}

      {/* MODAL: Achievements - DISABLED */}
    </div>
  );
}
