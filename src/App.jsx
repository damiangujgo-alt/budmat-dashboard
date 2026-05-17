import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── SUPABASE CONFIG ────────────────────────────────────────
const SB_URL = "https://goizernbjfthejdekykz.supabase.co";
const SB_KEY = "sb_publishable_QwxzW5eyomjU5WDtW6qERQ_Kaxxgx89";
const ADMIN_PASS = "Budmat2026";

const sbHeaders = {
  "apikey": SB_KEY,
  "Authorization": `Bearer ${SB_KEY}`,
  "Content-Type": "application/json",
};

async function sbGet(table, qs = "") {
  const r = await fetch(`${SB_URL}/rest/v1/${table}?select=*${qs}`, { headers: sbHeaders });
  if (!r.ok) throw new Error(`GET ${table}: ${await r.text()}`);
  return r.json();
}

async function sbPost(table, data) {
  const r = await fetch(`${SB_URL}/rest/v1/${table}`, {
    method: "POST",
    headers: { ...sbHeaders, "Prefer": "return=minimal" },
    body: JSON.stringify(data),
  });
  if (!r.ok) throw new Error(`POST ${table}: ${await r.text()}`);
}

async function sbDelete(table, filter) {
  const r = await fetch(`${SB_URL}/rest/v1/${table}?${filter}`, {
    method: "DELETE",
    headers: sbHeaders,
  });
  if (!r.ok) throw new Error(`DELETE ${table}: ${await r.text()}`);
}

async function sbUpsert(table, data, conflictCol = "sp") {
  const r = await fetch(`${SB_URL}/rest/v1/${table}?on_conflict=${conflictCol}`, {
    method: "POST",
    headers: { ...sbHeaders, "Prefer": "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify(data),
  });
  if (!r.ok) throw new Error(`UPSERT ${table}: ${await r.text()}`);
}

// ─── CONSTANTS ──────────────────────────────────────────────
const SALESPEOPLE = ["Andrzej Górny", "Arkadiusz Czerniak"];
const SP_LABEL    = { "Andrzej Górny": "Andrzej", "Arkadiusz Czerniak": "Arek" };
const SP_INIT     = { "Andrzej Górny": "AN", "Arkadiusz Czerniak": "AR" };
const SP_COLOR    = { "Andrzej Górny": "#185FA5", "Arkadiusz Czerniak": "#0F6E56" };
const MONTHS_PL   = ["Styczeń","Luty","Marzec","Kwiecień","Maj","Czerwiec","Lipiec","Sierpień","Wrzesień","Październik","Listopad","Grudzień"];
const now = new Date();

// ─── EXCEL PARSING ──────────────────────────────────────────
function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
        const mapped = rows.map(r => {
          const d = r["Data wprow."];
          const dt = d ? (d instanceof Date ? d : new Date(d)) : null;
          return {
            przypisany: r["Przypisany"] || null,
            status: r["Status"] || null,
            status_crm: r["Status CRM"] || null,
            kontrahent: r["Kontrahent"] || null,
            model: r["Potrzeba - Model"] || null,
            data_wprow: dt ? dt.toISOString().split("T")[0] : null,
          };
        }).filter(r => r.przypisany);
        resolve(mapped);
      } catch (err) { reject(err); }
    };
    fr.onerror = reject;
    fr.readAsArrayBuffer(file);
  });
}

// ─── KPI ENGINE ─────────────────────────────────────────────
function computeKPIs(leads, manual, targets, month, year) {
  const out = {};
  for (const sp of SALESPEOPLE) {
    const spLeads = leads.filter(r => r.przypisany === sp);
    const curLeads = spLeads.filter(r => {
      if (!r.data_wprow) return false;
      const d = new Date(r.data_wprow);
      return d.getMonth() + 1 === month && d.getFullYear() === year;
    });
    const dmsOrders = curLeads.filter(r =>
      r.status === "03. Realizacja umów" ||
      (r.status === "04. Zakończona" && r.status_crm === "Sukces")
    ).length;
    const manList = manual.filter(o => {
      const d = new Date(o.date);
      return o.sp === sp && d.getMonth() + 1 === month && d.getFullYear() === year;
    });
    const totalOrders = dmsOrders + manList.length;
    const convMon = curLeads.length > 0 ? Math.round(totalOrders / curLeads.length * 100) : 0;

    const pipelineRows = spLeads.filter(r => {
      if (r.status !== "02. Wybór ofert" || !r.data_wprow) return false;
      const d = new Date(r.data_wprow);
      const dm = d.getMonth() + 1, dy = d.getFullYear();
      for (let i = 0; i <= 2; i++) {
        let m = month - i, y = year;
        if (m <= 0) { m += 12; y--; }
        if (dm === m && dy === y) return true;
      }
      return false;
    });

    const target = (targets.find(t => t.sp === sp) || {}).target || 10;
    const planPct = Math.round(totalOrders / target * 100);
    out[sp] = { totalOrders, dmsOrders, manOrders: manList.length, curLeads: curLeads.length, pipeline: pipelineRows.length, pipelineRows, convMon, target, planPct };
  }
  [...SALESPEOPLE]
    .sort((a, b) => out[b].totalOrders - out[a].totalOrders || out[b].convMon - out[a].convMon)
    .forEach((sp, i) => out[sp].rank = i + 1);
  return out;
}

// ─── SHARED STYLES ──────────────────────────────────────────
const inpM = { width: "100%", padding: "8px 10px", borderRadius: "8px", border: "1px solid #d1d5db", background: "#fff", color: "#111827", fontSize: "14px", fontFamily: "inherit", boxSizing: "border-box" };
const btnPrimary = { padding: "10px 20px", borderRadius: "8px", border: "none", background: "#185FA5", color: "#fff", fontWeight: "500", cursor: "pointer", fontSize: "14px", display: "flex", alignItems: "center", gap: "6px" };
const btnSecondary = { padding: "10px 16px", borderRadius: "8px", border: "1px solid #d1d5db", background: "#fff", color: "#374151", cursor: "pointer", fontSize: "14px" };

// ─── MODAL ──────────────────────────────────────────────────
function Modal({ onClose, title, subtitle, children, maxWidth = "440px" }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, padding: "20px" }}>
      <div style={{ background: "#fff", border: "1px solid #d1d5db", borderRadius: "12px", padding: "24px", width: "100%", maxWidth, boxShadow: "0 8px 32px rgba(0,0,0,0.2)", maxHeight: "88vh", overflow: "hidden", display: "flex", flexDirection: "column" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "16px", flexShrink: 0 }}>
          <div>
            <p style={{ fontSize: "16px", fontWeight: "500", margin: "0 0 2px", color: "#111827" }}>{title}</p>
            {subtitle && <p style={{ fontSize: "12px", color: "#6b7280", margin: 0 }}>{subtitle}</p>}
          </div>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: "#6b7280", fontSize: "20px", padding: 0, lineHeight: 1, flexShrink: 0 }}><i className="ti ti-x" /></button>
        </div>
        <div style={{ overflowY: "auto", flex: 1 }}>{children}</div>
      </div>
    </div>
  );
}

// ─── FIELD ──────────────────────────────────────────────────
function Field({ label, children }) {
  return (
    <div style={{ marginBottom: "14px" }}>
      <p style={{ fontSize: "11px", color: "#6b7280", margin: "0 0 5px", textTransform: "uppercase", letterSpacing: "0.3px" }}>{label}</p>
      {children}
    </div>
  );
}

// ════════════════════════════════════════════════════════════
export default function App() {
  // ── state ──
  const [page, setPage]           = useState("dashboard"); // dashboard | login | admin
  const [isAdmin, setIsAdmin]     = useState(false);
  const [adminTab, setAdminTab]   = useState("upload");
  const [passInput, setPassInput] = useState("");
  const [passErr, setPassErr]     = useState(false);

  const [leads, setLeads]         = useState([]);
  const [manual, setManual]       = useState([]);
  const [targets, setTargets]     = useState([]);
  const [loading, setLoading]     = useState(true);
  const [dbError, setDbError]     = useState(null);

  const [selMonth, setSelMonth]   = useState(now.getMonth() + 1);
  const [selYear]                 = useState(now.getFullYear());

  const [modal, setModal]         = useState(null); // "add" | "pipeline"
  const [pipelineSP, setPipelineSP] = useState(null);

  const [addForm, setAddForm]     = useState({
    sp: SALESPEOPLE[0], model: "", client: "",
    date: `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}-${String(now.getDate()).padStart(2,"0")}`
  });

  // admin upload
  const [xlsxName, setXlsxName]       = useState(null);
  const [preview, setPreview]         = useState(null); // parsed rows
  const [uploading, setUploading]     = useState(false);
  const [uploadMsg, setUploadMsg]     = useState(null);

  // admin targets
  const [editTargets, setEditTargets] = useState(null);
  const [savingTargets, setSavingTargets] = useState(false);

  // ── load ──
  const load = useCallback(async () => {
    setLoading(true); setDbError(null);
    try {
      const [l, m, t] = await Promise.all([
        sbGet("leads"),
        sbGet("manual_orders"),
        sbGet("targets"),
      ]);
      setLeads(l); setManual(m); setTargets(t);
    } catch (e) {
      setDbError(e.message);
    }
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  // ── admin login ──
  function handleLogin(e) {
    e.preventDefault();
    if (passInput === ADMIN_PASS) {
      setIsAdmin(true); setPage("admin"); setPassErr(false); setPassInput("");
    } else {
      setPassErr(true);
    }
  }

  function logout() { setIsAdmin(false); setPage("dashboard"); setPassInput(""); }

  // ── upload ──
  async function handleFile(e) {
    const f = e.target.files[0]; if (!f) return;
    setXlsxName(f.name); setUploadMsg(null);
    try { setPreview(await parseExcel(f)); }
    catch (err) { setUploadMsg({ ok: false, text: "Błąd odczytu: " + err.message }); }
    e.target.value = "";
  }

  async function confirmUpload() {
    if (!preview) return;
    setUploading(true); setUploadMsg(null);
    try {
      await sbDelete("leads", "id=gte.0");
      const BATCH = 400;
      for (let i = 0; i < preview.length; i += BATCH) {
        await sbPost("leads", preview.slice(i, i + BATCH));
      }
      setPreview(null); setXlsxName(null);
      setUploadMsg({ ok: true, text: `Wgrano ${preview.length} rekordów — ${new Date().toLocaleTimeString("pl-PL")}` });
      await load();
    } catch (err) {
      setUploadMsg({ ok: false, text: "Błąd zapisu: " + err.message });
    }
    setUploading(false);
  }

  // ── targets ──
  async function saveTargets() {
    setSavingTargets(true);
    try {
      await Promise.all(SALESPEOPLE.map(sp =>
        sbUpsert("targets", { sp, target: editTargets[sp] || 10 })
      ));
      await load(); setEditTargets(null);
    } catch (err) { alert("Błąd: " + err.message); }
    setSavingTargets(false);
  }

  // ── manual orders ──
  async function addOrder() {
    if (!addForm.model.trim() || !addForm.client.trim()) return;
    try {
      await sbPost("manual_orders", { sp: addForm.sp, model: addForm.model, client: addForm.client, date: addForm.date });
      await load(); setModal(null);
      setAddForm(f => ({ ...f, model: "", client: "" }));
    } catch (err) { alert("Błąd: " + err.message); }
  }

  async function deleteOrder(id) {
    try { await sbDelete("manual_orders", `id=eq.${id}`); await load(); }
    catch (err) { alert("Błąd: " + err.message); }
  }

  // ── kpis ──
  const kpis   = leads.length > 0 ? computeKPIs(leads, manual, targets, selMonth, selYear) : null;
  const ranked = kpis ? [...SALESPEOPLE].sort((a, b) => kpis[a].rank - kpis[b].rank) : SALESPEOPLE;

  const teamOrders  = kpis ? SALESPEOPLE.reduce((s, sp) => s + kpis[sp].totalOrders, 0) : 0;
  const teamTarget  = kpis ? SALESPEOPLE.reduce((s, sp) => s + kpis[sp].target, 0) : 0;
  const teamPlan    = teamTarget > 0 ? Math.round(teamOrders / teamTarget * 100) : 0;
  const teamPipeline = kpis ? SALESPEOPLE.reduce((s, sp) => s + kpis[sp].pipeline, 0) : 0;
  const planColor   = (p) => p >= 100 ? "#3B6D11" : p >= 60 ? "#185FA5" : p >= 30 ? "#BA7517" : "#A32D2D";

  const manualThisMonth = manual.filter(o => {
    const d = new Date(o.date);
    return d.getMonth() + 1 === selMonth && d.getFullYear() === selYear;
  });

  // ════════════════════════════════════════════════════════
  // LOGIN PAGE
  // ════════════════════════════════════════════════════════
  if (page === "login") return (
    <div style={{ minHeight: "500px", display: "flex", alignItems: "center", justifyContent: "center", padding: "20px" }}>
      <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "32px", width: "100%", maxWidth: "340px" }}>
        <div style={{ width: "40px", height: "40px", background: "#185FA5", borderRadius: "10px", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: "16px" }}>
          <i className="ti ti-settings" style={{ color: "#fff", fontSize: "20px" }} />
        </div>
        <p style={{ fontSize: "18px", fontWeight: "500", margin: "0 0 4px" }}>Panel administratora</p>
        <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "0 0 24px" }}>Ford Budmat Auto · Płock</p>
        <form onSubmit={handleLogin}>
          <Field label="Hasło dostępu">
            <input type="password" value={passInput} onChange={e => { setPassInput(e.target.value); setPassErr(false); }}
              style={inpM} placeholder="••••••••" autoFocus />
          </Field>
          {passErr && <p style={{ fontSize: "12px", color: "#dc2626", margin: "-8px 0 12px" }}>Nieprawidłowe hasło</p>}
          <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
            <button type="submit" style={{ ...btnPrimary, flex: 1, justifyContent: "center" }}>Zaloguj</button>
            <button type="button" onClick={() => setPage("dashboard")} style={btnSecondary}>Wróć</button>
          </div>
        </form>
      </div>
    </div>
  );

  // ════════════════════════════════════════════════════════
  // ADMIN PAGE
  // ════════════════════════════════════════════════════════
  if (page === "admin" && isAdmin) {
    const tMap = Object.fromEntries(targets.map(t => [t.sp, t.target]));
    const lastUpload = leads.length > 0
      ? new Date(leads.reduce((max, l) => l.uploaded_at > max ? l.uploaded_at : max, leads[0].uploaded_at)).toLocaleString("pl-PL")
      : null;

    return (
      <div style={{ fontFamily: "inherit", padding: "20px", maxWidth: "760px", margin: "0 auto" }}>

        {/* Admin header */}
        <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "20px", flexWrap: "wrap" }}>
          <div style={{ flex: 1 }}>
            <p style={{ fontSize: "17px", fontWeight: "500", margin: "0 0 1px" }}>Panel administratora</p>
            <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: 0 }}>Ford Budmat Auto · Płock</p>
          </div>
          <button onClick={() => setPage("dashboard")} style={{ padding: "7px 14px", borderRadius: "var(--border-radius-md)", border: "0.5px solid var(--color-border-secondary)", cursor: "pointer", fontSize: "14px", background: "var(--color-background-primary)", color: "var(--color-text-primary)", display: "flex", alignItems: "center", gap: "5px" }}>
            <i className="ti ti-layout-dashboard" style={{ fontSize: "15px" }} /> Dashboard
          </button>
          <button onClick={logout} style={{ padding: "7px 12px", borderRadius: "var(--border-radius-md)", border: "0.5px solid var(--color-border-secondary)", cursor: "pointer", fontSize: "14px", background: "var(--color-background-primary)", color: "var(--color-text-secondary)" }}>
            Wyloguj
          </button>
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", gap: "2px", marginBottom: "16px", background: "var(--color-background-secondary)", padding: "4px", borderRadius: "var(--border-radius-md)" }}>
          {[
            { id: "upload", icon: "ti-upload", label: "Export DMS" },
            { id: "targets", icon: "ti-target", label: "Plany" },
            { id: "orders", icon: "ti-clipboard-list", label: "Zamówienia" },
          ].map(tab => (
            <button key={tab.id} onClick={() => setAdminTab(tab.id)} style={{ flex: 1, padding: "8px 10px", borderRadius: "6px", border: "none", cursor: "pointer", fontSize: "13px", fontWeight: adminTab === tab.id ? "500" : "400", background: adminTab === tab.id ? "var(--color-background-primary)" : "transparent", color: adminTab === tab.id ? "var(--color-text-primary)" : "var(--color-text-secondary)", display: "flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
              <i className={`ti ${tab.icon}`} style={{ fontSize: "14px" }} /> {tab.label}
            </button>
          ))}
        </div>

        {/* ── TAB: UPLOAD ── */}
        {adminTab === "upload" && (
          <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "20px" }}>

            <div style={{ display: "flex", gap: "16px", marginBottom: "20px", padding: "12px 16px", background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)" }}>
              <div>
                <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 2px", textTransform: "uppercase", letterSpacing: "0.3px" }}>Rekordów w bazie</p>
                <p style={{ fontSize: "22px", fontWeight: "500", margin: 0 }}>{leads.length}</p>
              </div>
              {lastUpload && (
                <div style={{ borderLeft: "0.5px solid var(--color-border-tertiary)", paddingLeft: "16px" }}>
                  <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 2px", textTransform: "uppercase", letterSpacing: "0.3px" }}>Ostatni upload</p>
                  <p style={{ fontSize: "14px", fontWeight: "500", margin: 0 }}>{lastUpload}</p>
                </div>
              )}
            </div>

            <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "0 0 14px" }}>
              Wgraj export z Autostacji (.xlsx). Istniejące dane zostaną <strong>zastąpione</strong> — ręczne zamówienia i plany pozostają.
            </p>

            <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginBottom: "14px" }}>
              <label style={{ padding: "9px 16px", borderRadius: "var(--border-radius-md)", border: "0.5px solid var(--color-border-secondary)", cursor: "pointer", fontSize: "14px", display: "flex", alignItems: "center", gap: "6px", color: "var(--color-text-primary)", userSelect: "none" }}>
                <i className="ti ti-file-spreadsheet" style={{ fontSize: "15px" }} />
                {xlsxName || "Wybierz plik .xlsx"}
                <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
              </label>

              {preview && (
                <button onClick={confirmUpload} disabled={uploading}
                  style={{ ...btnPrimary, background: "#0F6E56", opacity: uploading ? 0.7 : 1 }}>
                  <i className="ti ti-check" style={{ fontSize: "15px" }} />
                  {uploading ? "Zapisuję..." : `Zatwierdź — ${preview.length} rekordów`}
                </button>
              )}

              {preview && (
                <button onClick={() => { setPreview(null); setXlsxName(null); }}
                  style={{ ...btnSecondary, padding: "9px 14px" }}>
                  Anuluj
                </button>
              )}
            </div>

            {uploadMsg && (
              <div style={{ padding: "10px 14px", borderRadius: "8px", background: uploadMsg.ok ? "#d1fae5" : "#fee2e2", color: uploadMsg.ok ? "#065f46" : "#991b1b", fontSize: "13px", marginBottom: "14px" }}>
                <i className={`ti ${uploadMsg.ok ? "ti-circle-check" : "ti-alert-circle"}`} style={{ marginRight: "6px" }} />
                {uploadMsg.text}
              </div>
            )}

            {preview && (
              <div>
                <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 8px", textTransform: "uppercase", letterSpacing: "0.3px" }}>Podgląd — pierwsze 5 wierszy</p>
                <div style={{ border: "0.5px solid var(--color-border-tertiary)", borderRadius: "8px", overflow: "hidden" }}>
                  {preview.slice(0, 5).map((r, i) => (
                    <div key={i} style={{ display: "flex", gap: "10px", padding: "8px 12px", borderBottom: i < 4 ? "0.5px solid var(--color-border-tertiary)" : "none", fontSize: "12px", alignItems: "center" }}>
                      <span style={{ color: "var(--color-text-secondary)", minWidth: "90px", flexShrink: 0 }}>{r.przypisany || "—"}</span>
                      <span style={{ color: "var(--color-text-secondary)", minWidth: "72px", flexShrink: 0 }}>{r.data_wprow || "—"}</span>
                      <span style={{ flex: 1, color: "var(--color-text-primary)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.model || "—"}</span>
                      <span style={{ padding: "2px 8px", borderRadius: "10px", fontSize: "11px", flexShrink: 0, background: r.status === "03. Realizacja umów" ? "#d1fae5" : r.status === "02. Wybór ofert" ? "#fef3c7" : "var(--color-background-secondary)", color: r.status === "03. Realizacja umów" ? "#065f46" : r.status === "02. Wybór ofert" ? "#92400e" : "var(--color-text-secondary)" }}>
                        {r.status ? r.status.replace(/^\d+\.\s/, "") : "—"}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── TAB: TARGETS ── */}
        {adminTab === "targets" && (
          <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "20px" }}>
            <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "0 0 20px" }}>Miesięczne plany sprzedaży. Zmiana dotyczy wszystkich miesięcy do momentu ponownej edycji.</p>
            {SALESPEOPLE.map(sp => {
              const cur = tMap[sp] || 10;
              const val = editTargets ? (editTargets[sp] ?? cur) : cur;
              return (
                <div key={sp} style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "16px" }}>
                  <div style={{ width: "36px", height: "36px", borderRadius: "50%", background: SP_COLOR[sp] + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "12px", fontWeight: "500", color: SP_COLOR[sp], flexShrink: 0 }}>{SP_INIT[sp]}</div>
                  <span style={{ flex: 1, fontSize: "15px", fontWeight: "500" }}>{SP_LABEL[sp]}</span>
                  <span style={{ fontSize: "13px", color: "var(--color-text-secondary)", minWidth: "100px" }}>Aktualnie: <strong>{cur} szt.</strong></span>
                  <input type="number" min="0" max="99" value={val}
                    onChange={e => setEditTargets(prev => ({ ...(prev || tMap), [sp]: +e.target.value }))}
                    style={{ ...inpM, width: "72px", textAlign: "center" }} />
                  <span style={{ fontSize: "13px", color: "var(--color-text-secondary)" }}>szt.</span>
                </div>
              );
            })}
            <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
              <button onClick={saveTargets} disabled={savingTargets || !editTargets}
                style={{ ...btnPrimary, opacity: (!editTargets || savingTargets) ? 0.6 : 1 }}>
                <i className="ti ti-device-floppy" style={{ fontSize: "15px" }} />
                {savingTargets ? "Zapisuję..." : "Zapisz plany"}
              </button>
              {editTargets && <button onClick={() => setEditTargets(null)} style={btnSecondary}>Anuluj</button>}
            </div>
          </div>
        )}

        {/* ── TAB: ORDERS ── */}
        {adminTab === "orders" && (
          <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "20px" }}>
            <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "0 0 16px" }}>
              Wszystkie zamówienia dodane ręcznie — łącznie {manual.length} szt.
            </p>
            {manual.length === 0 ? (
              <p style={{ color: "var(--color-text-secondary)", fontSize: "14px", textAlign: "center", padding: "32px" }}>Brak zamówień ręcznych.</p>
            ) : [...manual].reverse().map(o => (
              <div key={o.id} style={{ display: "flex", alignItems: "center", gap: "12px", padding: "9px 0", borderBottom: "0.5px solid var(--color-border-tertiary)", fontSize: "13px" }}>
                <span style={{ width: "8px", height: "8px", borderRadius: "50%", background: SP_COLOR[o.sp] || "#888", flexShrink: 0 }} />
                <span style={{ minWidth: "56px", color: "var(--color-text-secondary)", fontWeight: "500" }}>{SP_LABEL[o.sp] || o.sp}</span>
                <span style={{ flex: 2, fontWeight: "500", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{o.model}</span>
                <span style={{ flex: 2, color: "var(--color-text-secondary)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{o.client}</span>
                <span style={{ color: "var(--color-text-secondary)", fontSize: "12px", whiteSpace: "nowrap" }}>{o.date}</span>
                <button onClick={() => deleteOrder(o.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#dc2626", padding: "2px 6px", flexShrink: 0 }}><i className="ti ti-trash" style={{ fontSize: "14px" }} /></button>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  // ════════════════════════════════════════════════════════
  // DASHBOARD PAGE
  // ════════════════════════════════════════════════════════
  return (
    <div style={{ fontFamily: "inherit" }}>
      <div style={{ padding: "20px", maxWidth: "880px", margin: "0 auto" }}>

        {/* Header */}
        <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "20px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: "160px" }}>
            <p style={{ fontSize: "17px", fontWeight: "500", margin: "0 0 1px" }}>Wyniki zespołu</p>
            <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: 0 }}>Ford Budmat Auto · Płock</p>
          </div>
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center" }}>
            <select value={selMonth} onChange={e => setSelMonth(+e.target.value)}
              style={{ padding: "7px 10px", borderRadius: "var(--border-radius-md)", border: "0.5px solid var(--color-border-secondary)", background: "var(--color-background-primary)", color: "var(--color-text-primary)", fontSize: "14px", cursor: "pointer" }}>
              {MONTHS_PL.map((m, i) => <option key={i} value={i+1}>{m} {selYear}</option>)}
            </select>
            <button onClick={load} title="Odśwież dane"
              style={{ padding: "7px 10px", borderRadius: "var(--border-radius-md)", border: "0.5px solid var(--color-border-secondary)", cursor: "pointer", fontSize: "14px", background: "var(--color-background-primary)", color: "var(--color-text-primary)", display: "flex", alignItems: "center" }}>
              <i className="ti ti-refresh" style={{ fontSize: "15px" }} />
            </button>
            <button onClick={() => setModal("add")}
              style={{ ...btnPrimary, padding: "7px 14px" }}>
              <i className="ti ti-plus" style={{ fontSize: "15px" }} /> Zamówienie
            </button>
            <button onClick={() => setPage("login")} title="Panel administratora"
              style={{ padding: "7px 10px", borderRadius: "var(--border-radius-md)", border: "0.5px solid var(--color-border-secondary)", cursor: "pointer", background: "var(--color-background-primary)", color: "var(--color-text-secondary)", display: "flex", alignItems: "center" }}>
              <i className="ti ti-settings" style={{ fontSize: "15px" }} />
            </button>
          </div>
        </div>

        {/* States */}
        {loading ? (
          <div style={{ textAlign: "center", padding: "80px 20px", color: "var(--color-text-secondary)" }}>
            <i className="ti ti-loader-2" style={{ fontSize: "36px", display: "block", marginBottom: "12px", opacity: 0.35 }} />
            <p style={{ margin: 0 }}>Łączenie z bazą...</p>
          </div>
        ) : dbError ? (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <i className="ti ti-database-off" style={{ fontSize: "36px", display: "block", marginBottom: "12px", color: "#dc2626", opacity: 0.6 }} />
            <p style={{ fontWeight: "500", color: "#dc2626", margin: "0 0 6px" }}>Błąd połączenia z bazą</p>
            <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "0 0 16px", maxWidth: "400px", marginLeft: "auto", marginRight: "auto" }}>{dbError}</p>
            <button onClick={load} style={btnSecondary}>Spróbuj ponownie</button>
          </div>
        ) : leads.length === 0 ? (
          <div style={{ textAlign: "center", padding: "64px 20px", border: "0.5px dashed var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)" }}>
            <i className="ti ti-database-import" style={{ fontSize: "40px", display: "block", marginBottom: "12px", opacity: 0.3 }} />
            <p style={{ fontSize: "16px", fontWeight: "500", margin: "0 0 6px" }}>Baza jest pusta</p>
            <p style={{ fontSize: "13px", color: "var(--color-text-secondary)", margin: "0 0 16px" }}>Zaloguj się jako admin i wgraj export z Autostacji</p>
            <button onClick={() => setPage("login")} style={btnPrimary}>Przejdź do panelu admina</button>
          </div>
        ) : (<>

          {/* ── TEAM SUMMARY ── */}
          <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "20px", marginBottom: "14px" }}>
            <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 16px", textTransform: "uppercase", letterSpacing: "0.4px" }}>
              Zespół · {MONTHS_PL[selMonth-1]} {selYear}
            </p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0, paddingBottom: "16px", marginBottom: "16px", borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
              <div style={{ borderRight: "0.5px solid var(--color-border-tertiary)", paddingRight: "20px" }}>
                <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 6px", textTransform: "uppercase", letterSpacing: "0.3px" }}>Zamówienia / Plan</p>
                <div style={{ display: "flex", alignItems: "flex-end", gap: "6px", marginBottom: "8px" }}>
                  <span style={{ fontSize: "44px", fontWeight: "500", lineHeight: 1, color: planColor(teamPlan) }}>{teamOrders}</span>
                  <span style={{ fontSize: "14px", color: "var(--color-text-secondary)", paddingBottom: "6px" }}>/ {teamTarget}</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                  <div style={{ flex: 1, height: "5px", background: "var(--color-background-tertiary)", borderRadius: "3px", overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${Math.min(teamPlan, 100)}%`, background: planColor(teamPlan), borderRadius: "3px" }} />
                  </div>
                  <span style={{ fontSize: "13px", fontWeight: "500", color: planColor(teamPlan), whiteSpace: "nowrap" }}>{teamPlan}%</span>
                </div>
              </div>
              <div style={{ paddingLeft: "20px" }}>
                <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: "0 0 6px", textTransform: "uppercase", letterSpacing: "0.3px" }}>Pipeline aktywny</p>
                <p style={{ fontSize: "44px", fontWeight: "500", lineHeight: 1, margin: "0 0 6px", color: "#BA7517" }}>{teamPipeline}</p>
                <p style={{ fontSize: "12px", color: "var(--color-text-secondary)", margin: 0 }}>oferty status 02 · bieżący + 2 miesiące</p>
              </div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px" }}>
              {ranked.map(sp => {
                const k = kpis[sp]; const c = SP_COLOR[sp];
                return (
                  <div key={sp} style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <div style={{ width: "28px", height: "28px", borderRadius: "50%", background: c + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "11px", fontWeight: "500", color: c, flexShrink: 0 }}>{SP_INIT[sp]}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "3px" }}>
                        <span style={{ fontSize: "13px", fontWeight: "500" }}>{SP_LABEL[sp]}</span>
                        <span style={{ fontSize: "13px", color: "var(--color-text-secondary)" }}>{k.totalOrders} / {k.target}</span>
                      </div>
                      <div style={{ height: "3px", background: "var(--color-background-tertiary)", borderRadius: "2px", overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${Math.min(k.planPct, 100)}%`, background: c, borderRadius: "2px" }} />
                      </div>
                    </div>
                    <span style={{ fontSize: "12px", fontWeight: "500", color: c, minWidth: "34px", textAlign: "right" }}>{k.planPct}%</span>
                  </div>
                );
              })}
            </div>
          </div>

          {/* ── PERSON CARDS ── */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0,1fr))", gap: "12px", marginBottom: "14px" }}>
            {ranked.map(sp => {
              const k = kpis[sp]; const c = SP_COLOR[sp];
              const bc = planColor(k.planPct);
              return (
                <div key={sp} style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "18px", borderTop: `3px solid ${c}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                      <div style={{ width: "36px", height: "36px", borderRadius: "50%", background: c + "18", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: "500", color: c, flexShrink: 0 }}>{SP_INIT[sp]}</div>
                      <p style={{ fontSize: "18px", fontWeight: "500", margin: 0 }}>{SP_LABEL[sp]}</p>
                    </div>
                    <span style={{ fontSize: "11px", fontWeight: "500", padding: "3px 9px", borderRadius: "20px", background: k.rank === 1 ? "#FAEEDA" : "#F1EFE8", color: k.rank === 1 ? "#BA7517" : "#5F5E5A" }}>#{k.rank}</span>
                  </div>

                  <div style={{ display: "flex", alignItems: "flex-end", gap: "8px", marginBottom: "14px" }}>
                    <span style={{ fontSize: "52px", fontWeight: "500", lineHeight: 1, color: c }}>{k.totalOrders}</span>
                    <span style={{ fontSize: "13px", color: "var(--color-text-secondary)", paddingBottom: "7px", lineHeight: 1.4 }}>/ {k.target}<br />zamówień</span>
                  </div>

                  <div style={{ marginBottom: "16px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "5px" }}>
                      <span style={{ fontSize: "12px", color: "var(--color-text-secondary)" }}>Realizacja planu</span>
                      <span style={{ fontSize: "12px", fontWeight: "500", color: bc }}>{k.planPct}%</span>
                    </div>
                    <div style={{ height: "5px", background: "var(--color-background-tertiary)", borderRadius: "3px", overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${Math.min(k.planPct,100)}%`, background: bc, borderRadius: "3px" }} />
                    </div>
                  </div>

                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", marginBottom: "12px" }}>
                    <div style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "10px 6px", textAlign: "center" }}>
                      <p style={{ fontSize: "20px", fontWeight: "500", margin: "0 0 3px" }}>{k.convMon}%</p>
                      <p style={{ fontSize: "9px", color: "var(--color-text-secondary)", margin: 0, textTransform: "uppercase", letterSpacing: "0.3px", lineHeight: 1.4 }}>konwersja<br/>miesięczna</p>
                    </div>
                    <div onClick={k.pipeline > 0 ? () => { setPipelineSP(sp); setModal("pipeline"); } : undefined}
                      style={{ background: "var(--color-background-secondary)", borderRadius: "var(--border-radius-md)", padding: "10px 6px", textAlign: "center", cursor: k.pipeline > 0 ? "pointer" : "default", border: k.pipeline > 0 ? "0.5px solid var(--color-border-secondary)" : "0.5px solid transparent" }}>
                      <p style={{ fontSize: "20px", fontWeight: "500", margin: "0 0 3px", color: k.pipeline > 0 ? "#BA7517" : "inherit" }}>{k.pipeline}</p>
                      <p style={{ fontSize: "9px", color: "var(--color-text-secondary)", margin: 0, textTransform: "uppercase", letterSpacing: "0.3px", lineHeight: 1.4 }}>pipeline{k.pipeline > 0 ? " →" : ""}<br/>&nbsp;</p>
                    </div>
                  </div>

                  <p style={{ fontSize: "11px", color: "var(--color-text-secondary)", margin: 0 }}>
                    {k.totalOrders} z {k.curLeads} leadów w {MONTHS_PL[selMonth-1].toLowerCase()}
                    {k.manOrders > 0 && <span style={{ color: "#3B6D11" }}> · {k.manOrders} ręcznych</span>}
                  </p>
                </div>
              );
            })}
          </div>

          {/* ── MANUAL ORDERS ── */}
          {manualThisMonth.length > 0 && (
            <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: "var(--border-radius-lg)", padding: "16px" }}>
              <p style={{ fontSize: "11px", fontWeight: "500", margin: "0 0 10px", color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.3px" }}>Zamówienia dodane ręcznie</p>
              {manualThisMonth.map(o => (
                <div key={o.id} style={{ display: "flex", alignItems: "center", gap: "12px", padding: "7px 0", borderBottom: "0.5px solid var(--color-border-tertiary)", fontSize: "13px" }}>
                  <span style={{ width: "8px", height: "8px", borderRadius: "50%", background: SP_COLOR[o.sp] || "#888", flexShrink: 0 }} />
                  <span style={{ minWidth: "52px", color: "var(--color-text-secondary)" }}>{SP_LABEL[o.sp] || o.sp}</span>
                  <span style={{ flex: 2, fontWeight: "500", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{o.model}</span>
                  <span style={{ flex: 2, color: "var(--color-text-secondary)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{o.client}</span>
                  <span style={{ fontSize: "12px", color: "var(--color-text-secondary)", whiteSpace: "nowrap" }}>{o.date}</span>
                  <button onClick={() => deleteOrder(o.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "var(--color-text-secondary)", padding: "2px 4px", flexShrink: 0 }}>
                    <i className="ti ti-trash" style={{ fontSize: "14px" }} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </>)}
      </div>

      {/* ── MODAL: Add order ── */}
      {modal === "add" && (
        <Modal onClose={() => setModal(null)} title="Dodaj zamówienie" subtitle="Zamówienie zostanie zapisane w bazie">
          <Field label="Handlowiec">
            <select value={addForm.sp} onChange={e => setAddForm({ ...addForm, sp: e.target.value })} style={inpM}>
              {SALESPEOPLE.map(s => <option key={s} value={s}>{SP_LABEL[s]}</option>)}
            </select>
          </Field>
          <Field label="Model pojazdu">
            <input value={addForm.model} onChange={e => setAddForm({ ...addForm, model: e.target.value })} placeholder="np. Ford Transit Custom" style={inpM} />
          </Field>
          <Field label="Klient">
            <input value={addForm.client} onChange={e => setAddForm({ ...addForm, client: e.target.value })} placeholder="Nazwa firmy lub osoby" style={inpM} />
          </Field>
          <Field label="Data">
            <input type="date" value={addForm.date} onChange={e => setAddForm({ ...addForm, date: e.target.value })} style={inpM} />
          </Field>
          <div style={{ display: "flex", gap: "8px", marginTop: "20px" }}>
            <button onClick={addOrder} style={{ ...btnPrimary, flex: 1, justifyContent: "center" }}>Dodaj</button>
            <button onClick={() => setModal(null)} style={btnSecondary}>Anuluj</button>
          </div>
        </Modal>
      )}

      {/* ── MODAL: Pipeline ── */}
      {modal === "pipeline" && pipelineSP && kpis && (
        <Modal onClose={() => setModal(null)} title={`Pipeline · ${SP_LABEL[pipelineSP]}`} subtitle="Status 02 · bieżący + 2 poprzednie miesiące" maxWidth="540px">
          {kpis[pipelineSP].pipelineRows.length === 0 ? (
            <p style={{ textAlign: "center", color: "#6b7280", padding: "32px" }}>Brak leadów</p>
          ) : kpis[pipelineSP].pipelineRows.map((r, i) => {
            const d = r.data_wprow ? new Date(r.data_wprow) : null;
            const isCur = d && d.getMonth()+1 === selMonth && d.getFullYear() === selYear;
            return (
              <div key={i} style={{ padding: "10px 0", borderBottom: "1px solid #f3f4f6", display: "flex", gap: "12px", alignItems: "center" }}>
                <span style={{ fontSize: "11px", color: isCur ? "#185FA5" : "#9ca3af", whiteSpace: "nowrap", minWidth: "46px" }}>
                  {d ? `${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getFullYear()).slice(2)}` : "—"}
                </span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <p style={{ fontSize: "13px", fontWeight: "500", margin: "0 0 1px", color: "#111827", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.kontrahent || "—"}</p>
                  <p style={{ fontSize: "11px", color: "#6b7280", margin: 0 }}>{r.model || "Model nieznany"}</p>
                </div>
                {isCur && <span style={{ fontSize: "11px", padding: "2px 7px", borderRadius: "20px", background: "#dbeafe", color: "#1e40af", flexShrink: 0 }}>bieżący</span>}
              </div>
            );
          })}
        </Modal>
      )}
    </div>
  );
}
