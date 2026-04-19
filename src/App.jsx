import React, { useEffect, useMemo, useState } from "react";

const STORAGE_KEY = "tmdk_sales_force_plain_cache_v1";
const SETTINGS_KEY = "tmdk_sales_force_plain_settings_v1";
const SESSION_KEY = "tmdk_sales_force_plain_session_v1";
const SALES_TYPES = ["Direct", "Indirect"];
const VISIT_TYPES = ["Routine Visit", "Prospecting", "Follow-up", "Key Account Review", "Order Collection", "Complaint Resolution"];
const PIPELINE_STATUSES = ["Prospecting", "Follow-up", "Negotiation", "Converted", "Dormant"];
const ROLES = ["ASM", "RSM", "NSM", "Sales Operations Admin"];
const ADMIN_ROLES = ["NSM", "Sales Operations Admin"];
const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwQIVyacsmcIhySMROBUXTtOuGpH-GO1scyU5L7za__mFRLKYL0ttK8hJk5EeWZROuo5g/exec";
const DEFAULT_LOGO_URL = "https://drive.google.com/thumbnail?id=15ZeFX5OUlXVKvJt409-WpiYzcevUUG-l&sz=w1000";

function nowISO() {
  return new Date().toISOString();
}
function todayISO() {
  return nowISO().slice(0, 10);
}
function addDays(dateStr, days) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}
function currency(n) {
  return new Intl.NumberFormat("en-NG", {
    style: "currency",
    currency: "NGN",
    maximumFractionDigits: 0,
  }).format(Number(n || 0));
}
function splitCsv(text) {
  return String(text || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);
}
function generateId(prefix, items) {
  const max = (items || []).reduce((acc, item) => {
    const raw = String(item?.id || "");
    const num = parseInt(raw.replace(prefix, ""), 10);
    return Number.isFinite(num) ? Math.max(acc, num) : acc;
  }, 0);
  return `${prefix}${String(max + 1).padStart(3, "0")}`;
}
function downloadCSV(filename, rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const escape = (v) => `"${String(v ?? "").replaceAll('"', '""')}"`;
  const csv = [headers.join(","), ...rows.map((row) => headers.map((header) => escape(row[header])).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
function parsePossibleJson(value) {
  if (typeof value !== "string") return value;
  const trimmed = value.trim();
  if ((trimmed.startsWith("[") && trimmed.endsWith("]")) || (trimmed.startsWith("{") && trimmed.endsWith("}"))) {
    try {
      return JSON.parse(trimmed);
    } catch {
      return value;
    }
  }
  return value;
}
function hydrateRows(rows, entity) {
  return (rows || []).map((row) => {
    const hydrated = { ...row };
    Object.keys(hydrated).forEach((key) => {
      hydrated[key] = parsePossibleJson(hydrated[key]);
    });
    if (entity === "managers" && !Array.isArray(hydrated.states)) hydrated.states = splitCsv(String(hydrated.states || ""));
    ["active", "plannedAtStartOfDay", "orderMade"].forEach((boolKey) => {
      if (boolKey in hydrated) hydrated[boolKey] = hydrated[boolKey] === true || hydrated[boolKey] === "true";
    });
    return hydrated;
  });
}
function normalizeDb(raw) {
  const db = raw || {};
  return {
    managers: hydrateRows(Array.isArray(db.managers) ? db.managers : [], "managers"),
    customers: hydrateRows(Array.isArray(db.customers) ? db.customers : [], "customers"),
    visits: hydrateRows(Array.isArray(db.visits) ? db.visits : [], "visits"),
    auditLog: hydrateRows(Array.isArray(db.auditLog) ? db.auditLog : [], "auditLog"),
    authUsers: hydrateRows(Array.isArray(db.authUsers) ? db.authUsers : [], "authUsers"),
  };
}
function mergeCollections(leftRows, rightRows) {
  const map = new Map();
  [...(leftRows || []), ...(rightRows || [])].forEach((row) => {
    if (!row?.id) return;
    const existing = map.get(row.id);
    if (!existing) {
      map.set(row.id, row);
      return;
    }
    const existingStamp = new Date(existing.updatedAt || 0).getTime();
    const rowStamp = new Date(row.updatedAt || 0).getTime();
    if (rowStamp >= existingStamp) map.set(row.id, row);
  });
  return Array.from(map.values());
}
function mergeDatabases(localDb, remoteDb) {
  const local = normalizeDb(localDb);
  const remote = normalizeDb(remoteDb);
  return {
    managers: mergeCollections(local.managers, remote.managers),
    customers: mergeCollections(local.customers, remote.customers),
    visits: mergeCollections(local.visits, remote.visits),
    auditLog: mergeCollections(local.auditLog, remote.auditLog),
    authUsers: mergeCollections(local.authUsers, remote.authUsers),
  };
}
function withMeta(record, actorEmail, extra = {}) {
  return { ...record, updatedAt: nowISO(), updatedBy: actorEmail || "system", deletedAt: null, deletedBy: null, ...extra };
}
function addAuditEntry(db, actor, action, entityType, entityId, summary) {
  const entry = {
    id: generateId("A", db.auditLog || []),
    action,
    entityType,
    entityId,
    summary,
    actorEmail: actor?.email || "system",
    actorName: actor?.name || "System",
    actorRole: actor?.role || "System",
    createdAt: nowISO(),
    updatedAt: nowISO(),
    updatedBy: actor?.email || "system",
    deletedAt: null,
    deletedBy: null,
  };
  return { ...db, auditLog: [...(db.auditLog || []), entry] };
}
function sanitizeRows(rows) {
  return (rows || []).map((row) => {
    const clean = { ...row };
    Object.keys(clean).forEach((key) => {
      if (Array.isArray(clean[key])) clean[key] = JSON.stringify(clean[key]);
      if (clean[key] === undefined || clean[key] === null) clean[key] = "";
    });
    return clean;
  });
}
function defaultManagerForm() { return { name: "", role: "ASM", region: "", statesText: "", email: "", phone: "", reportsTo: "", active: true }; }
function defaultCustomerForm() { return { name: "", customerType: "Farm", salesType: "Direct", region: "", state: "", city: "", address: "", contactPerson: "", phone: "", email: "", segment: "Prospect", active: true }; }
function defaultVisitForm() { return { visitDate: todayISO(), managerId: "", customerId: "", plannedAtStartOfDay: true, visitType: "Routine Visit", objective: "", outcome: "", comments: "", orderMade: false, orderQtyMt: 0, orderValueNgn: 0, pipelineStatus: "Prospecting", nextActionDate: todayISO(), visitStatus: "Planned", salesType: "Direct" }; }

const seedCache = normalizeDb({
  managers: [
    withMeta({ id: "M001", name: "ASM Kaduna I", role: "ASM", region: "North West", states: ["Kaduna", "Katsina"], email: "asm.kaduna@tmdk.com", phone: "08030000001", reportsTo: "M003", active: true, createdAt: nowISO() }, "system"),
    withMeta({ id: "M002", name: "ASM Kano", role: "ASM", region: "North West", states: ["Kano"], email: "asm.kano@tmdk.com", phone: "08030000002", reportsTo: "M003", active: true, createdAt: nowISO() }, "system"),
    withMeta({ id: "M003", name: "RSM North West", role: "RSM", region: "North West", states: ["Kaduna", "Katsina", "Kano", "Jigawa", "Sokoto", "Kebbi", "Zamfara"], email: "rsm.nw@tmdk.com", phone: "08030000003", reportsTo: "M004", active: true, createdAt: nowISO() }, "system"),
    withMeta({ id: "M004", name: "National Sales Manager", role: "NSM", region: "All Regions", states: [], email: "nsm@tmdk.com", phone: "08030000004", reportsTo: "", active: true, createdAt: nowISO() }, "system"),
    withMeta({ id: "M005", name: "Sales Operations Admin", role: "Sales Operations Admin", region: "All Regions", states: [], email: "salesops@tmdk.com", phone: "08030000005", reportsTo: "M004", active: true, createdAt: nowISO() }, "system"),
  ],
  customers: [
    withMeta({ id: "C001", name: "Gidan Gona Farms", customerType: "Farm", salesType: "Direct", region: "North West", state: "Kaduna", city: "Zaria", address: "Zaria, Kaduna State", contactPerson: "Musa Aliyu", phone: "08040000001", email: "musa@gidangona.com", segment: "Prospect", active: true, createdAt: nowISO() }, "system"),
    withMeta({ id: "C002", name: "Arewa Agro Traders", customerType: "Distributor", salesType: "Indirect", region: "North West", state: "Kano", city: "Kano", address: "Kano State", contactPerson: "Aisha Bello", phone: "08040000002", email: "aisha@arewaagro.com", segment: "Key Account", active: true, createdAt: nowISO() }, "system"),
  ],
  visits: [
    withMeta({ id: "V001", visitDate: todayISO(), managerId: "M001", customerId: "C001", plannedAtStartOfDay: true, visitType: "Routine Visit", objective: "Introduce product and assess feed demand", outcome: "Prospecting", comments: "Customer requested price list.", orderMade: false, orderQtyMt: 0, orderValueNgn: 0, pipelineStatus: "Follow-up", nextActionDate: addDays(todayISO(), 5), visitStatus: "Completed", salesType: "Direct", createdAt: nowISO() }, "system"),
  ],
  auditLog: [],
  authUsers: [],
});

function runSelfTests() {
  console.assert(splitCsv("Kaduna, Kano").length === 2, "splitCsv should split comma text");
  console.assert(generateId("M", [{ id: "M001" }, { id: "M010" }]) === "M011", "generateId should increment properly");
  console.assert(defaultVisitForm().pipelineStatus === "Prospecting", "default visit pipeline default");
  const merged = mergeCollections([{ id: "X1", updatedAt: "2026-01-01T00:00:00.000Z", value: 1 }], [{ id: "X1", updatedAt: "2026-01-02T00:00:00.000Z", value: 2 }]);
  console.assert(merged[0].value === 2, "mergeCollections should keep newest row");
}

function useLocalCache() {
  const [db, setDb] = useState(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? normalizeDb(JSON.parse(raw)) : seedCache;
    } catch {
      return seedCache;
    }
  });
  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(db)); }, [db]);
  useEffect(() => { runSelfTests(); }, []);
  return [db, setDb];
}
function useSettings() {
  const [settings, setSettings] = useState(() => {
    const defaults = {
      appsScriptUrl: DEFAULT_APPS_SCRIPT_URL,
      apiKey: "",
      logoUrl: DEFAULT_LOGO_URL,
      autoPullEnabled: false,
      autoPullSeconds: 60,
      autoPushEnabled: false,
      autoPushSeconds: 120,
    };
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      return raw ? { ...defaults, ...JSON.parse(raw) } : defaults;
    } catch {
      return defaults;
    }
  });
  useEffect(() => { localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); }, [settings]);
  return [settings, setSettings];
}
function useBackendSession() {
  const [session, setSession] = useState(() => {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch {
      return null;
    }
  });
  useEffect(() => {
    if (session) localStorage.setItem(SESSION_KEY, JSON.stringify(session));
    else localStorage.removeItem(SESSION_KEY);
  }, [session]);
  return [session, setSession];
}

async function backendGet(settings, action, session, extraParams = {}) {
  if (!settings.appsScriptUrl) throw new Error("Add your Google Apps Script Web App URL first.");
  const url = new URL(settings.appsScriptUrl);
  url.searchParams.set("action", action);
  if (settings.apiKey) url.searchParams.set("apiKey", settings.apiKey);
  if (session?.token) url.searchParams.set("token", session.token);
  Object.entries(extraParams).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") url.searchParams.set(key, String(value));
  });
  const response = await fetch(url.toString(), { method: "GET" });
  const json = await response.json();
  if (!response.ok || !json?.success) throw new Error(json?.message || `Request failed: ${response.status}`);
  return json;
}
async function backendPost(settings, body) {
  if (!settings.appsScriptUrl) throw new Error("Add your Google Apps Script Web App URL first.");
  const response = await fetch(settings.appsScriptUrl, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ ...body, apiKey: settings.apiKey || body.apiKey || "" }),
  });
  const json = await response.json();
  if (!response.ok || !json?.success) throw new Error(json?.message || `Request failed: ${response.status}`);
  return json;
}
async function backendLogin(settings, email, password) { const result = await backendPost(settings, { action: "login", email, password }); return result.data; }
async function backendLogout(settings, session) { return backendPost(settings, { action: "logout", token: session?.token || "" }); }
async function backendRead(settings, session) { const result = await backendGet(settings, "read", session); return normalizeDb(result.data); }
async function backendMerge(settings, session, db) { return backendPost(settings, { action: "merge", token: session?.token || "", data: { managers: sanitizeRows(db.managers), customers: sanitizeRows(db.customers), visits: sanitizeRows(db.visits), auditLog: sanitizeRows(db.auditLog), authUsers: sanitizeRows(db.authUsers) } }); }
async function backendChangePassword(settings, session, currentPassword, newPassword) { return backendPost(settings, { action: "changePassword", token: session?.token || "", currentPassword, newPassword }); }

function getVisibleManagers(allManagers, currentManager) {
  if (!currentManager) return [];
  if (ADMIN_ROLES.includes(currentManager.role)) return allManagers.filter((m) => !m.deletedAt);
  if (currentManager.role === "RSM") return allManagers.filter((m) => !m.deletedAt && (m.id === currentManager.id || (m.role === "ASM" && m.region === currentManager.region)));
  return allManagers.filter((m) => !m.deletedAt && m.id === currentManager.id);
}

const styles = {
  page: {
    minHeight: "100vh",
    background: "linear-gradient(180deg, #f8fafc 0%, #f0fdf4 100%)",
    padding: 20,
    fontFamily: 'Inter, Arial, sans-serif',
    color: "#0f172a",
  },
  container: {
    maxWidth: 1280,
    margin: "0 auto",
    display: "flex",
    flexDirection: "column",
    gap: 18,
  },
  card: {
    background: "rgba(255,255,255,0.92)",
    border: "1px solid #e2e8f0",
    borderRadius: 20,
    padding: 18,
    boxShadow: "0 10px 30px rgba(15,23,42,0.06)",
    backdropFilter: "blur(8px)",
  },
  button: {
    border: "1px solid #22c55e",
    background: "linear-gradient(135deg, #22c55e 0%, #16a34a 100%)",
    color: "#ffffff",
    borderRadius: 12,
    padding: "10px 16px",
    cursor: "pointer",
    fontWeight: 600,
    boxShadow: "0 8px 20px rgba(34,197,94,0.22)",
  },
  buttonSecondary: {
    border: "1px solid #dbe4ee",
    background: "#ffffff",
    color: "#0f172a",
    borderRadius: 12,
    padding: "10px 16px",
    cursor: "pointer",
    fontWeight: 600,
  },
  input: {
    width: "100%",
    border: "1px solid #dbe4ee",
    borderRadius: 12,
    padding: 12,
    boxSizing: "border-box",
    background: "#ffffff",
    color: "#0f172a",
  },
  textarea: {
    width: "100%",
    border: "1px solid #dbe4ee",
    borderRadius: 12,
    padding: 12,
    boxSizing: "border-box",
    minHeight: 90,
    background: "#ffffff",
    color: "#0f172a",
  },
  select: {
    width: "100%",
    border: "1px solid #dbe4ee",
    borderRadius: 12,
    padding: 12,
    boxSizing: "border-box",
    background: "#ffffff",
    color: "#0f172a",
  },
  grid2: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 },
  grid4: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 },
  tableWrap: { overflowX: "auto" },
  table: { width: "100%", borderCollapse: "collapse" },
  th: {
    textAlign: "left",
    padding: 12,
    borderBottom: "1px solid #edf2f7",
    fontSize: 12,
    letterSpacing: "0.04em",
    textTransform: "uppercase",
    color: "#64748b",
  },
  td: {
    padding: 12,
    borderBottom: "1px solid #edf2f7",
    fontSize: 14,
    color: "#0f172a",
  },
  badge: {
    display: "inline-block",
    background: "#dcfce7",
    color: "#166534",
    padding: "5px 10px",
    borderRadius: 999,
    fontSize: 12,
    fontWeight: 600,
  },
  tabs: { display: "flex", flexWrap: "wrap", gap: 10 },
  tab: (active) => ({
    border: active ? "1px solid #bbf7d0" : "1px solid #dbe4ee",
    background: active ? "linear-gradient(135deg, #22c55e 0%, #16a34a 100%)" : "rgba(255,255,255,0.92)",
    color: active ? "#ffffff" : "#0f172a",
    borderRadius: 12,
    padding: "10px 16px",
    cursor: "pointer",
    fontWeight: 600,
    boxShadow: active ? "0 8px 20px rgba(34,197,94,0.18)" : "none",
  }),
  label: { display: "block", fontSize: 13, marginBottom: 6, color: "#475569", fontWeight: 600 },
  modalBackdrop: {
    position: "fixed",
    inset: 0,
    background: "rgba(15,23,42,0.45)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 16,
    zIndex: 50,
  },
  modal: {
    width: "min(960px, 100%)",
    maxHeight: "90vh",
    overflow: "auto",
    background: "#ffffff",
    borderRadius: 22,
    padding: 22,
    border: "1px solid #e2e8f0",
    boxShadow: "0 18px 50px rgba(15,23,42,0.18)",
  },
};

function Field({ label, children }) {
  return <div><label style={styles.label}>{label}</label>{children}</div>;
}
function Card({ children }) { return <div style={styles.card}>{children}</div>; }
function Button({ children, onClick, secondary = false, disabled = false, type = "button" }) {
  return <button type={type} disabled={disabled} onClick={onClick} style={{ ...(secondary ? styles.buttonSecondary : styles.button), opacity: disabled ? 0.6 : 1 }}>{children}</button>;
}
function Badge({ children }) { return <span style={styles.badge}>{children}</span>; }
function Modal({ open, title, children, onClose }) {
  if (!open) return null;
  return (
    <div style={styles.modalBackdrop} onClick={onClose}>
      <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0 }}>{title}</h3>
          <Button secondary onClick={onClose}>Close</Button>
        </div>
        {children}
      </div>
    </div>
  );
}
function PermissionBanner({ role, title, message }) {
  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 13, color: "#64748b" }}>{title}</div>
          <div style={{ fontSize: 14, fontWeight: 600 }}>{message}</div>
        </div>
        <Badge>{role}</Badge>
      </div>
    </Card>
  );
}

function LogoBlock({ logoUrl, compact = false }) {
  const fallbackUrl = "https://drive.google.com/thumbnail?id=15ZeFX5OUlXVKvJt409-WpiYzcevUUG-l&sz=w1000";
  const [currentLogo, setCurrentLogo] = useState(logoUrl || fallbackUrl);

  useEffect(() => {
    setCurrentLogo(logoUrl || fallbackUrl);
  }, [logoUrl]);

  return (
    <div style={{ display: "flex", alignItems: "center", gap: compact ? 10 : 14 }}>
      <div
        style={{
          width: compact ? 44 : 56,
          height: compact ? 44 : 56,
          borderRadius: compact ? 12 : 16,
          background: "linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%)",
          border: "1px solid #bbf7d0",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          overflow: "hidden",
          boxShadow: "0 8px 20px rgba(34,197,94,0.14)",
        }}
      >
        {currentLogo ? (
          <img
            src={currentLogo}
            alt="TMDK logo"
            style={{ width: "100%", height: "100%", objectFit: "contain", background: "#fff" }}
            onError={() => {
              if (currentLogo !== fallbackUrl) setCurrentLogo(fallbackUrl);
              else setCurrentLogo("");
            }}
          />
        ) : (
          <span style={{ fontWeight: 800, color: "#166534", fontSize: compact ? 16 : 20 }}>T</span>
        )}
      </div>
      <div>
        <div style={{ fontWeight: 800, fontSize: compact ? 18 : 22, letterSpacing: "-0.02em" }}>
          TMDK Sales Force
        </div>
        <div style={{ color: "#64748b", fontSize: compact ? 12 : 13 }}>
          Pipeline, field execution, and manager intelligence
        </div>
      </div>
    </div>
  );
}

function LoginScreen({ settings, setSettings, onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [status, setStatus] = useState("");
  const [loading, setLoading] = useState(false);
  return (
    <div style={styles.page}>
      <div style={{ ...styles.container, maxWidth: 1080, minHeight: "80vh", justifyContent: "center" }}>
        <div style={{ ...styles.grid2, alignItems: "stretch" }}>
          <Card>
            <LogoBlock logoUrl={settings.logoUrl} />
            <div style={{ marginTop: 22, display: "grid", gap: 14 }}>
              <div style={{ display: "grid", gap: 10 }}>
                {[
                  "Role-based visibility for ASM, RSM, NSM, and Sales Operations Admin",
                  "Incremental merge sync with Google Apps Script backend",
                  "Visit planning, pipeline tracking, exports, and audit trail",
                ].map((item) => (
                  <div key={item} style={{ display: "flex", gap: 10, alignItems: "flex-start", color: "#475569" }}>
                    <div style={{ width: 10, height: 10, borderRadius: 999, background: "#22c55e", marginTop: 5 }} />
                    <div>{item}</div>
                  </div>
                ))}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12 }}>
                <div style={{ border: "1px solid #edf2f7", borderRadius: 16, padding: 14, background: "#ffffff" }}>
                  <div style={{ color: "#64748b", fontSize: 12, textTransform: "uppercase", letterSpacing: "0.04em" }}>Live sync</div>
                  <div style={{ fontSize: 22, fontWeight: 800, marginTop: 6 }}>Pull / Push</div>
                </div>
                <div style={{ border: "1px solid #edf2f7", borderRadius: 16, padding: 14, background: "#ffffff" }}>
                  <div style={{ color: "#64748b", fontSize: 12, textTransform: "uppercase", letterSpacing: "0.04em" }}>Currency</div>
                  <div style={{ fontSize: 22, fontWeight: 800, marginTop: 6 }}>NGN</div>
                </div>
              </div>
            </div>
          </Card>
          <Card>
            <div style={{ fontSize: 12, color: "#16a34a", fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase" }}>Secure sign in</div>
            <h2 style={{ marginTop: 10, marginBottom: 8 }}>Access your sales workspace</h2>
            <p style={{ color: "#64748b", marginTop: 0 }}>Sign in with your assigned credentials.</p>
            <div style={{ border: "1px solid #edf2f7", borderRadius: 14, padding: 12, background: "#ffffff", color: "#64748b", fontSize: 13 }}>
              Backend URL and company logo are already configured in the code.
            </div>
            <div style={{ display: "grid", gap: 12, marginTop: 14 }}>
              <Field label="Email"><input style={styles.input} value={email} onChange={(e) => setEmail(e.target.value)} /></Field>
              <Field label="Password"><input style={styles.input} type="password" value={password} onChange={(e) => setPassword(e.target.value)} /></Field>
              {status ? <div style={{ border: "1px solid #ecfdf5", borderRadius: 12, padding: 12, background: "#f8fafc" }}>{status}</div> : null}
              <Button onClick={async () => {
                try {
                  setLoading(true);
                  setStatus("Signing in...");
                  const session = await backendLogin(settings, email, password);
                  setStatus("Login successful.");
                  onLogin(session);
                } catch (error) {
                  setStatus(error instanceof Error ? error.message : "Login failed.");
                } finally {
                  setLoading(false);
                }
              }} disabled={loading}>Sign In</Button>
              <div style={{ color: "#64748b", fontSize: 13 }}>Use your role email and password to continue.</div>
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

function AppHeader({ session, onLogout, onChangePassword, logoUrl }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
      <LogoBlock logoUrl={logoUrl} compact />
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <Card><div style={{ fontSize: 14 }}><strong>{session?.name}</strong><br /><span style={{ color: "#64748b" }}>{session?.email} • {session?.role}</span></div></Card>
        <Button secondary onClick={onChangePassword}>Change Password</Button>
        <Button secondary onClick={onLogout}>Sign Out</Button>
      </div>
    </div>
  );
}

function Filters({ filters, setFilters, managers, customers }) {
  const availableRegions = Array.from(new Set(customers.map((c) => c.region).filter(Boolean)));
  const availableStates = Array.from(new Set(customers.map((c) => c.state).filter(Boolean)));
  const resetFilters = () => setFilters({ from: "", to: "", region: "all", state: "all", managerId: "all", pipeline: "all" });
  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, gap: 12, flexWrap: "wrap" }}>
        <h3 style={{ margin: 0 }}>Dashboard Filters</h3>
        <Button secondary onClick={resetFilters}>Reset Filters</Button>
      </div>
      <div style={styles.grid2}>
        <Field label="Date From"><input style={styles.input} type="date" value={filters.from} onChange={(e) => setFilters((f) => ({ ...f, from: e.target.value }))} /></Field>
        <Field label="Date To"><input style={styles.input} type="date" value={filters.to} onChange={(e) => setFilters((f) => ({ ...f, to: e.target.value }))} /></Field>
        <Field label="Region"><select style={styles.select} value={filters.region} onChange={(e) => setFilters((f) => ({ ...f, region: e.target.value }))}><option value="all">All regions</option>{availableRegions.map((region) => <option key={region} value={region}>{region}</option>)}</select></Field>
        <Field label="State"><select style={styles.select} value={filters.state} onChange={(e) => setFilters((f) => ({ ...f, state: e.target.value }))}><option value="all">All states</option>{availableStates.map((state) => <option key={state} value={state}>{state}</option>)}</select></Field>
        <Field label="Manager"><select style={styles.select} value={filters.managerId} onChange={(e) => setFilters((f) => ({ ...f, managerId: e.target.value }))}><option value="all">All managers</option>{managers.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
        <Field label="Pipeline"><select style={styles.select} value={filters.pipeline} onChange={(e) => setFilters((f) => ({ ...f, pipeline: e.target.value }))}><option value="all">All pipeline</option>{PIPELINE_STATUSES.map((status) => <option key={status} value={status}>{status}</option>)}</select></Field>
      </div>
    </Card>
  );
}

function Dashboard({ visits, customers, managers, currentManager, filters }) {
  const visibleManagers = getVisibleManagers(managers, currentManager);
  const visibleManagerIds = visibleManagers.map((m) => m.id);
  const activeCustomers = customers.filter((c) => !c.deletedAt);
  const activeVisits = visits.filter((v) => !v.deletedAt);
  const enrichedVisits = activeVisits.map((v) => ({ ...v, manager: managers.find((m) => m.id === v.managerId), customer: customers.find((c) => c.id === v.customerId) }));
  const filteredVisits = enrichedVisits.filter((v) => {
    const dateOk = (!filters.from || v.visitDate >= filters.from) && (!filters.to || v.visitDate <= filters.to);
    const regionOk = filters.region === "all" || !filters.region || v.customer?.region === filters.region;
    const stateOk = filters.state === "all" || !filters.state || v.customer?.state === filters.state;
    const managerOk = (filters.managerId === "all" || !filters.managerId || v.managerId === filters.managerId) && visibleManagerIds.includes(v.managerId);
    const pipelineOk = filters.pipeline === "all" || !filters.pipeline || v.pipelineStatus === filters.pipeline;
    return dateOk && regionOk && stateOk && managerOk && pipelineOk;
  });
  const totalVisits = filteredVisits.length;
  const completedVisits = filteredVisits.filter((v) => String(v.visitStatus || "").toLowerCase() === "completed").length;
  const convertedVisits = filteredVisits.filter((v) => v.orderMade || v.pipelineStatus === "Converted").length;
  const totalOrderValue = filteredVisits.reduce((sum, v) => sum + Number(v.orderValueNgn || 0), 0);
  const totalOrderMt = filteredVisits.reduce((sum, v) => sum + Number(v.orderQtyMt || 0), 0);
  const plannedVisits = filteredVisits.filter((v) => v.plannedAtStartOfDay).length;
  const conversionRatio = totalVisits ? ((convertedVisits / totalVisits) * 100).toFixed(1) : "0.0";
  const byPipeline = PIPELINE_STATUSES.map((status) => ({ label: status, value: filteredVisits.filter((v) => v.pipelineStatus === status).length }));
  const byManager = visibleManagerIds.map((id) => {
    const manager = managers.find((m) => m.id === id);
    const rows = filteredVisits.filter((v) => v.managerId === id);
    return { name: manager?.name || id, visits: rows.length, converted: rows.filter((r) => r.orderMade || r.pipelineStatus === "Converted").length, value: rows.reduce((sum, r) => sum + Number(r.orderValueNgn || 0), 0), mt: rows.reduce((sum, r) => sum + Number(r.orderQtyMt || 0), 0) };
  }).filter((row) => row.visits > 0).sort((a, b) => b.visits - a.visits);
  const upcomingFollowUps = filteredVisits.filter((v) => v.nextActionDate >= todayISO() && v.pipelineStatus !== "Converted").sort((a, b) => a.nextActionDate.localeCompare(b.nextActionDate)).slice(0, 5);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={styles.grid4}>
        {[{ label: "Total Visits", value: totalVisits }, { label: "Completed Visits", value: completedVisits }, { label: "Conversion Ratio", value: `${conversionRatio}%` }, { label: "Order Value", value: currency(totalOrderValue) }, { label: "Order Volume", value: `${totalOrderMt.toFixed(2)} MT` }, { label: "Planned Visits", value: plannedVisits }].map((kpi) => <Card key={kpi.label}><div style={{ color: "#64748b", fontSize: 13 }}>{kpi.label}</div><div style={{ fontSize: 24, fontWeight: 700, marginTop: 8 }}>{kpi.value}</div></Card>)}
      </div>
      <div style={styles.grid2}>
        <Card>
          <h3 style={{ marginTop: 0 }}>Pipeline Mix</h3>
          <div style={{ display: "grid", gap: 10 }}>{byPipeline.map((row) => <div key={row.label}><div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}><span>{row.label}</span><span>{row.value}</span></div><div style={{ height: 8, background: "#dcfce7", borderRadius: 999 }}><div style={{ height: 8, background: "#16a34a", borderRadius: 999, width: `${totalVisits ? Math.round((row.value / totalVisits) * 100) : 0}%` }} /></div></div>)}</div>
        </Card>
        <Card>
          <h3 style={{ marginTop: 0 }}>Manager Performance</h3>
          <div style={styles.tableWrap}><table style={styles.table}><thead><tr>{["Manager", "Visits", "Converted", "Conversion %", "Order MT", "Order Value"].map((h) => <th key={h} style={styles.th}>{h}</th>)}</tr></thead><tbody>{byManager.length ? byManager.map((m) => <tr key={m.name}><td style={styles.td}>{m.name}</td><td style={styles.td}>{m.visits}</td><td style={styles.td}>{m.converted}</td><td style={styles.td}>{m.visits ? ((m.converted / m.visits) * 100).toFixed(1) : 0}%</td><td style={styles.td}>{m.mt.toFixed(2)} MT</td><td style={styles.td}>{currency(m.value)}</td></tr>) : <tr><td style={styles.td} colSpan={6}>No records for selected filters.</td></tr>}</tbody></table></div>
        </Card>
      </div>
      <div style={styles.grid2}>
        <Card><h3 style={{ marginTop: 0 }}>Role Access Summary</h3><div style={{ color: "#475569", display: "grid", gap: 8 }}><div><strong>ASM:</strong> sees personal dashboard and activities.</div><div><strong>RSM:</strong> sees personal dashboard plus ASMs in assigned region.</div><div><strong>NSM / Sales Ops Admin:</strong> sees everything.</div></div></Card>
        <Card><h3 style={{ marginTop: 0 }}>Upcoming Follow-ups</h3><div style={{ display: "grid", gap: 8 }}>{upcomingFollowUps.length ? upcomingFollowUps.map((v) => <div key={v.id} style={{ border: "1px solid #ecfdf5", borderRadius: 12, padding: 10 }}><div style={{ fontWeight: 600 }}>{v.customer?.name}</div><div style={{ color: "#64748b", fontSize: 13 }}>{v.manager?.name} • {v.nextActionDate}</div></div>) : <div style={{ color: "#64748b" }}>No upcoming follow-ups.</div>}</div></Card>
      </div>
    </div>
  );
}

function ManagersTab({ managers, authUsers, setDb, currentManager, session }) {
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const canManage = ADMIN_ROLES.includes(session.role);
  const activeManagers = managers.filter((m) => !m.deletedAt);
  const [form, setForm] = useState(defaultManagerForm());
  useEffect(() => {
    setForm(editing ? { name: editing.name, role: editing.role, region: editing.region, statesText: (editing.states || []).join(", "), email: editing.email, phone: editing.phone, reportsTo: editing.reportsTo, active: editing.active } : defaultManagerForm());
  }, [editing, open]);
  const saveManager = () => {
    setDb((db) => {
      const existingAuth = db.authUsers.find((u) => String(u.email).toLowerCase() === String(form.email).toLowerCase() && !u.deletedAt);
      const payload = withMeta({ id: editing?.id || generateId("M", db.managers), name: form.name, role: form.role, region: form.region, states: splitCsv(form.statesText), email: form.email, phone: form.phone, reportsTo: form.reportsTo, active: form.active, createdAt: editing?.createdAt || nowISO() }, session.email);
      let nextDb = { ...db, managers: editing ? db.managers.map((m) => (m.id === editing.id ? payload : m)) : [...db.managers, payload] };
      if (!editing && !existingAuth && form.email) {
        const authUser = withMeta({ id: generateId("U", db.authUsers), email: form.email, password: "1234", managerId: payload.id, role: form.role, name: form.name, active: true, createdAt: nowISO() }, session.email);
        nextDb = { ...nextDb, authUsers: [...nextDb.authUsers, authUser] };
      }
      return addAuditEntry(nextDb, session, editing ? "update" : "create", "manager", payload.id, `${editing ? "Updated" : "Created"} manager ${payload.name}`);
    });
    setOpen(false);
    setEditing(null);
  };
  const removeManager = (id) => {
    if (id === currentManager.id) return window.alert("You cannot delete the user you are currently signed in as.");
    const target = activeManagers.find((m) => m.id === id);
    const directReports = activeManagers.filter((m) => m.reportsTo === id);
    if (directReports.length) return window.alert(`Cannot delete ${target?.name || "this manager"} because they still have direct reports.`);
    if (!window.confirm(`Delete ${target?.name || "this manager"}?`)) return;
    setDb((db) => addAuditEntry({ ...db, managers: db.managers.map((m) => m.id === id ? { ...m, active: false, deletedAt: nowISO(), deletedBy: session.email, updatedAt: nowISO(), updatedBy: session.email } : m), authUsers: db.authUsers.map((u) => u.managerId === id ? { ...u, active: false, deletedAt: nowISO(), deletedBy: session.email, updatedAt: nowISO(), updatedBy: session.email } : u), visits: db.visits.map((v) => v.managerId === id ? { ...v, deletedAt: nowISO(), deletedBy: session.email, updatedAt: nowISO(), updatedBy: session.email } : v) }, session, "delete", "manager", id, `Soft-deleted manager ${target?.name || id}`));
  };
  return (
    <div style={{ display: "grid", gap: 16 }}>
      <PermissionBanner role={session.role} title="Manager permissions" message={canManage ? "You can add, edit, and soft-delete managers." : "You can only view managers."} />
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 12 }}><div><h3 style={{ margin: 0 }}>Managers Master</h3><div style={{ color: "#64748b", fontSize: 14 }}>Manage ASM, RSM, NSM and Sales Operations Admin profiles.</div></div>{canManage ? <Button onClick={() => { setEditing(null); setOpen(true); }}>Add Manager</Button> : null}</div>
        <div style={styles.tableWrap}><table style={styles.table}><thead><tr>{["ID", "Name", "Role", "Region", "States", "Reports To", "Status"].map((h) => <th key={h} style={styles.th}>{h}</th>)}{canManage ? <th style={styles.th}>Actions</th> : null}</tr></thead><tbody>{activeManagers.map((m) => <tr key={m.id}><td style={styles.td}>{m.id}</td><td style={styles.td}>{m.name}</td><td style={styles.td}><Badge>{m.role}</Badge></td><td style={styles.td}>{m.region}</td><td style={styles.td}>{(m.states || []).join(", ")}</td><td style={styles.td}>{activeManagers.find((x) => x.id === m.reportsTo)?.name || "—"}</td><td style={styles.td}>{m.active ? "Active" : "Inactive"}</td>{canManage ? <td style={styles.td}><div style={{ display: "flex", gap: 8 }}><Button secondary onClick={() => { setEditing(m); setOpen(true); }}>Edit</Button><Button secondary onClick={() => removeManager(m.id)}>Delete</Button></div></td> : null}</tr>)}</tbody></table></div>
      </Card>
      <Modal open={open} onClose={() => setOpen(false)} title={editing ? "Edit Manager" : "Add Manager"}>
        <div style={styles.grid2}>
          <Field label="Name"><input style={styles.input} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
          <Field label="Role"><select style={styles.select} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>{ROLES.map((role) => <option key={role} value={role}>{role}</option>)}</select></Field>
          <Field label="Region"><input style={styles.input} value={form.region} onChange={(e) => setForm({ ...form, region: e.target.value })} /></Field>
          <Field label="States Covered (comma separated)"><input style={styles.input} value={form.statesText} onChange={(e) => setForm({ ...form, statesText: e.target.value })} /></Field>
          <Field label="Email"><input style={styles.input} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
          <Field label="Phone"><input style={styles.input} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label="Reports To"><select style={styles.select} value={form.reportsTo || ""} onChange={(e) => setForm({ ...form, reportsTo: e.target.value })}><option value="">None</option>{activeManagers.filter((m) => m.id !== editing?.id).map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
          <Field label="Active"><label><input type="checkbox" checked={form.active} onChange={(e) => setForm({ ...form, active: e.target.checked })} /> Active</label></Field>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}><Button secondary onClick={() => setOpen(false)}>Cancel</Button><Button onClick={saveManager}>Save Manager</Button></div>
      </Modal>
    </div>
  );
}

function CustomersTab({ customers, setDb, session }) {
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState(defaultCustomerForm());
  const [search, setSearch] = useState("");
  const canManage = ["NSM", "Sales Operations Admin", "RSM", "ASM"].includes(session.role);
  const rows = customers.filter((c) => !c.deletedAt && JSON.stringify(c).toLowerCase().includes(search.toLowerCase()));
  useEffect(() => {
    setForm(editing ? { name: editing.name, customerType: editing.customerType, salesType: editing.salesType, region: editing.region, state: editing.state, city: editing.city, address: editing.address, contactPerson: editing.contactPerson, phone: editing.phone, email: editing.email, segment: editing.segment, active: editing.active } : defaultCustomerForm());
  }, [editing, open]);
  const saveCustomer = () => {
    setDb((db) => {
      const payload = withMeta({ id: editing?.id || generateId("C", db.customers), ...form, createdAt: editing?.createdAt || nowISO() }, session.email);
      return addAuditEntry({ ...db, customers: editing ? db.customers.map((c) => (c.id === editing.id ? payload : c)) : [...db.customers, payload] }, session, editing ? "update" : "create", "customer", payload.id, `${editing ? "Updated" : "Created"} customer ${payload.name}`);
    });
    setOpen(false);
    setEditing(null);
  };
  const removeCustomer = (id) => {
    const target = rows.find((c) => c.id === id);
    if (!window.confirm(`Delete ${target?.name || "this customer"}?`)) return;
    setDb((db) => addAuditEntry({ ...db, customers: db.customers.map((c) => c.id === id ? { ...c, active: false, deletedAt: nowISO(), deletedBy: session.email, updatedAt: nowISO(), updatedBy: session.email } : c), visits: db.visits.map((v) => v.customerId === id ? { ...v, deletedAt: nowISO(), deletedBy: session.email, updatedAt: nowISO(), updatedBy: session.email } : v) }, session, "delete", "customer", id, `Soft-deleted customer ${target?.name || id}`));
  };
  return (
    <div style={{ display: "grid", gap: 16 }}>
      <PermissionBanner role={session.role} title="Customer permissions" message={canManage ? "You can add, edit, and soft-delete customers." : "You can only view customers."} />
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 12 }}><div><h3 style={{ margin: 0 }}>Customer Master</h3><div style={{ color: "#64748b", fontSize: 14 }}>Track direct and indirect sales customers.</div></div><div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}><input style={styles.input} placeholder="Search customers" value={search} onChange={(e) => setSearch(e.target.value)} />{canManage ? <Button onClick={() => { setEditing(null); setOpen(true); }}>Add Customer</Button> : null}</div></div>
        <div style={styles.tableWrap}><table style={styles.table}><thead><tr>{["ID", "Name", "Type", "Region", "State", "Sales Type", "Segment", "Status"].map((h) => <th key={h} style={styles.th}>{h}</th>)}{canManage ? <th style={styles.th}>Actions</th> : null}</tr></thead><tbody>{rows.map((c) => <tr key={c.id}><td style={styles.td}>{c.id}</td><td style={styles.td}>{c.name}</td><td style={styles.td}>{c.customerType}</td><td style={styles.td}>{c.region}</td><td style={styles.td}>{c.state}</td><td style={styles.td}>{c.salesType}</td><td style={styles.td}>{c.segment}</td><td style={styles.td}>{c.active ? "Active" : "Inactive"}</td>{canManage ? <td style={styles.td}><div style={{ display: "flex", gap: 8 }}><Button secondary onClick={() => { setEditing(c); setOpen(true); }}>Edit</Button><Button secondary onClick={() => removeCustomer(c.id)}>Delete</Button></div></td> : null}</tr>)}</tbody></table></div>
      </Card>
      <Modal open={open} onClose={() => setOpen(false)} title={editing ? "Edit Customer" : "Add Customer"}>
        <div style={styles.grid2}>
          <Field label="Customer Name"><input style={styles.input} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
          <Field label="Customer Type"><input style={styles.input} value={form.customerType} onChange={(e) => setForm({ ...form, customerType: e.target.value })} /></Field>
          <Field label="Sales Type"><select style={styles.select} value={form.salesType} onChange={(e) => setForm({ ...form, salesType: e.target.value })}>{SALES_TYPES.map((type) => <option key={type} value={type}>{type}</option>)}</select></Field>
          <Field label="Segment"><input style={styles.input} value={form.segment} onChange={(e) => setForm({ ...form, segment: e.target.value })} /></Field>
          <Field label="Region"><input style={styles.input} value={form.region} onChange={(e) => setForm({ ...form, region: e.target.value })} /></Field>
          <Field label="State"><input style={styles.input} value={form.state} onChange={(e) => setForm({ ...form, state: e.target.value })} /></Field>
          <Field label="City"><input style={styles.input} value={form.city} onChange={(e) => setForm({ ...form, city: e.target.value })} /></Field>
          <Field label="Contact Person"><input style={styles.input} value={form.contactPerson} onChange={(e) => setForm({ ...form, contactPerson: e.target.value })} /></Field>
          <Field label="Phone"><input style={styles.input} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
          <Field label="Email"><input style={styles.input} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
          <Field label="Address"><textarea style={styles.textarea} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field>
          <Field label="Active"><label><input type="checkbox" checked={form.active} onChange={(e) => setForm({ ...form, active: e.target.checked })} /> Active</label></Field>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}><Button secondary onClick={() => setOpen(false)}>Cancel</Button><Button onClick={saveCustomer}>Save Customer</Button></div>
      </Modal>
    </div>
  );
}

function VisitsTab({ visits, customers, managers, setDb, currentManager, session }) {
  const [open, setOpen] = useState(false);
  const [bulkOpen, setBulkOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState(defaultVisitForm());
  const [search, setSearch] = useState("");
  const [bulkManagerId, setBulkManagerId] = useState("");
  const [bulkVisitDate, setBulkVisitDate] = useState(todayISO());
  const [bulkVisitType, setBulkVisitType] = useState("Routine Visit");
  const [bulkObjective, setBulkObjective] = useState("Planned route visit");
  const [bulkSelected, setBulkSelected] = useState([]);
  const canManage = ["NSM", "Sales Operations Admin", "RSM", "ASM"].includes(session.role);
  const visibleManagers = getVisibleManagers(managers, currentManager);
  const visibleManagerIds = visibleManagers.map((m) => m.id);
  const activeCustomers = customers.filter((c) => !c.deletedAt);
  const rows = visits.filter((v) => !v.deletedAt).map((v) => ({ ...v, manager: managers.find((m) => m.id === v.managerId), customer: customers.find((c) => c.id === v.customerId) })).filter((v) => visibleManagerIds.includes(v.managerId)).filter((v) => JSON.stringify(v).toLowerCase().includes(search.toLowerCase())).sort((a, b) => b.visitDate.localeCompare(a.visitDate));
  useEffect(() => {
    setForm(editing ? { visitDate: editing.visitDate, managerId: editing.managerId, customerId: editing.customerId, plannedAtStartOfDay: editing.plannedAtStartOfDay, visitType: editing.visitType, objective: editing.objective, outcome: editing.outcome, comments: editing.comments, orderMade: editing.orderMade, orderQtyMt: editing.orderQtyMt || 0, orderValueNgn: editing.orderValueNgn, pipelineStatus: editing.pipelineStatus, nextActionDate: editing.nextActionDate, visitStatus: editing.visitStatus, salesType: editing.salesType } : defaultVisitForm());
  }, [editing, open]);
  const saveVisit = () => {
    setDb((db) => {
      const payload = withMeta({ id: editing?.id || generateId("V", db.visits), ...form, createdAt: editing?.createdAt || nowISO() }, session.email);
      return addAuditEntry({ ...db, visits: editing ? db.visits.map((v) => (v.id === editing.id ? payload : v)) : [...db.visits, payload] }, session, editing ? "update" : "create", "visit", payload.id, `${editing ? "Updated" : "Created"} visit ${payload.id}`);
    });
    setOpen(false);
    setEditing(null);
  };
  const removeVisit = (id) => {
    if (!window.confirm("Delete this visit record?")) return;
    setDb((db) => addAuditEntry({ ...db, visits: db.visits.map((v) => v.id === id ? { ...v, deletedAt: nowISO(), deletedBy: session.email, updatedAt: nowISO(), updatedBy: session.email } : v) }, session, "delete", "visit", id, `Soft-deleted visit ${id}`));
  };
  const toggleBulk = (id) => setBulkSelected((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const bulkSave = () => {
    if (!bulkManagerId || !bulkSelected.length) return window.alert("Select a manager and at least one customer.");
    setDb((db) => {
      const working = [...db.visits];
      const newVisits = bulkSelected.map((customerId) => {
        const nextId = generateId("V", working);
        const customer = db.customers.find((c) => c.id === customerId);
        const visit = withMeta({ id: nextId, visitDate: bulkVisitDate, managerId: bulkManagerId, customerId, plannedAtStartOfDay: true, visitType: bulkVisitType, objective: bulkObjective, outcome: "", comments: "", orderMade: false, orderQtyMt: 0, orderValueNgn: 0, pipelineStatus: "Prospecting", nextActionDate: bulkVisitDate, visitStatus: "Planned", salesType: customer?.salesType || "Direct", createdAt: nowISO() }, session.email);
        working.push(visit);
        return visit;
      });
      return addAuditEntry({ ...db, visits: [...db.visits, ...newVisits] }, session, "bulk_create", "visit", newVisits.map((v) => v.id).join(", "), `Bulk planned ${newVisits.length} visits`);
    });
    setBulkOpen(false);
    setBulkSelected([]);
  };
  return (
    <div style={{ display: "grid", gap: 16 }}>
      <PermissionBanner role={session.role} title="Visit permissions" message={canManage ? "You can create, edit, delete, and bulk-plan visits." : "You can only view visits."} />
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 12 }}><div><h3 style={{ margin: 0 }}>Visits</h3><div style={{ color: "#64748b", fontSize: 14 }}>Plan visits and track conversion pipeline.</div></div><div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}><input style={styles.input} placeholder="Search visits" value={search} onChange={(e) => setSearch(e.target.value)} />{canManage ? <Button secondary onClick={() => setBulkOpen(true)}>Bulk Plan</Button> : null}{canManage ? <Button onClick={() => { setEditing(null); setOpen(true); }}>Add Visit</Button> : null}</div></div>
        <div style={styles.tableWrap}><table style={styles.table}><thead><tr>{["Date", "Manager", "Customer", "Visit Type", "Pipeline", "Order", "Order MT", "Order Value", "Status"].map((h) => <th key={h} style={styles.th}>{h}</th>)}{canManage ? <th style={styles.th}>Actions</th> : null}</tr></thead><tbody>{rows.map((v) => <tr key={v.id}><td style={styles.td}>{v.visitDate}</td><td style={styles.td}>{v.manager?.name}</td><td style={styles.td}>{v.customer?.name}</td><td style={styles.td}>{v.visitType}</td><td style={styles.td}><Badge>{v.pipelineStatus}</Badge></td><td style={styles.td}>{v.orderMade ? "Yes" : "No"}</td><td style={styles.td}>{Number(v.orderQtyMt || 0).toFixed(2)} MT</td><td style={styles.td}>{currency(v.orderValueNgn)}</td><td style={styles.td}>{v.visitStatus}</td>{canManage ? <td style={styles.td}><div style={{ display: "flex", gap: 8 }}><Button secondary onClick={() => { setEditing(v); setOpen(true); }}>Edit</Button><Button secondary onClick={() => removeVisit(v.id)}>Delete</Button></div></td> : null}</tr>)}</tbody></table></div>
      </Card>
      <Modal open={open} onClose={() => setOpen(false)} title={editing ? "Edit Visit" : "Add Visit"}>
        <div style={styles.grid2}>
          <Field label="Visit Date"><input style={styles.input} type="date" value={form.visitDate} onChange={(e) => setForm({ ...form, visitDate: e.target.value })} /></Field>
          <Field label="Manager"><select style={styles.select} value={form.managerId} onChange={(e) => setForm({ ...form, managerId: e.target.value })}><option value="">Select manager</option>{visibleManagers.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
          <Field label="Customer"><select style={styles.select} value={form.customerId} onChange={(e) => { const v = e.target.value; const customer = activeCustomers.find((c) => c.id === v); setForm({ ...form, customerId: v, salesType: customer?.salesType || form.salesType }); }}><option value="">Select customer</option>{activeCustomers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select></Field>
          <Field label="Sales Type"><select style={styles.select} value={form.salesType} onChange={(e) => setForm({ ...form, salesType: e.target.value })}>{SALES_TYPES.map((type) => <option key={type} value={type}>{type}</option>)}</select></Field>
          <Field label="Visit Type"><select style={styles.select} value={form.visitType} onChange={(e) => setForm({ ...form, visitType: e.target.value })}>{VISIT_TYPES.map((type) => <option key={type} value={type}>{type}</option>)}</select></Field>
          <Field label="Planned At Start Of Day"><label><input type="checkbox" checked={form.plannedAtStartOfDay} onChange={(e) => setForm({ ...form, plannedAtStartOfDay: e.target.checked })} /> Planned</label></Field>
          <Field label="Objective"><textarea style={styles.textarea} value={form.objective} onChange={(e) => setForm({ ...form, objective: e.target.value })} /></Field>
          <Field label="Outcome"><input style={styles.input} value={form.outcome} onChange={(e) => setForm({ ...form, outcome: e.target.value })} /></Field>
          <Field label="Pipeline Status"><select style={styles.select} value={form.pipelineStatus} onChange={(e) => setForm({ ...form, pipelineStatus: e.target.value })}>{PIPELINE_STATUSES.map((status) => <option key={status} value={status}>{status}</option>)}</select></Field>
          <Field label="Comments"><textarea style={styles.textarea} value={form.comments} onChange={(e) => setForm({ ...form, comments: e.target.value })} /></Field>
          <Field label="Order Made"><label><input type="checkbox" checked={form.orderMade} onChange={(e) => setForm({ ...form, orderMade: e.target.checked })} /> Order Made</label></Field>
          <Field label="Order Quantity (MT)"><input style={styles.input} type="number" step="0.01" value={form.orderQtyMt} onChange={(e) => setForm({ ...form, orderQtyMt: Number(e.target.value || 0) })} /></Field>
          <Field label="Order Value (NGN)"><input style={styles.input} type="number" value={form.orderValueNgn} onChange={(e) => setForm({ ...form, orderValueNgn: Number(e.target.value || 0) })} /></Field>
          <Field label="Next Action Date"><input style={styles.input} type="date" value={form.nextActionDate} onChange={(e) => setForm({ ...form, nextActionDate: e.target.value })} /></Field>
          <Field label="Visit Status"><input style={styles.input} value={form.visitStatus} onChange={(e) => setForm({ ...form, visitStatus: e.target.value })} /></Field>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}><Button secondary onClick={() => setOpen(false)}>Cancel</Button><Button onClick={saveVisit}>Save Visit</Button></div>
      </Modal>
      <Modal open={bulkOpen} onClose={() => setBulkOpen(false)} title="Bulk Visit Planning">
        <div style={styles.grid2}>
          <Field label="Manager"><select style={styles.select} value={bulkManagerId} onChange={(e) => setBulkManagerId(e.target.value)}><option value="">Select manager</option>{visibleManagers.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
          <Field label="Visit Date"><input style={styles.input} type="date" value={bulkVisitDate} onChange={(e) => setBulkVisitDate(e.target.value)} /></Field>
          <Field label="Visit Type"><select style={styles.select} value={bulkVisitType} onChange={(e) => setBulkVisitType(e.target.value)}>{VISIT_TYPES.map((type) => <option key={type} value={type}>{type}</option>)}</select></Field>
          <Field label="Objective"><textarea style={styles.textarea} value={bulkObjective} onChange={(e) => setBulkObjective(e.target.value)} /></Field>
        </div>
        <div style={{ marginTop: 16, ...styles.tableWrap }}><table style={styles.table}><thead><tr>{["Select", "Customer", "State", "Region", "Sales Type"].map((h) => <th key={h} style={styles.th}>{h}</th>)}</tr></thead><tbody>{activeCustomers.map((c) => <tr key={c.id}><td style={styles.td}><input type="checkbox" checked={bulkSelected.includes(c.id)} onChange={() => toggleBulk(c.id)} /></td><td style={styles.td}>{c.name}</td><td style={styles.td}>{c.state}</td><td style={styles.td}>{c.region}</td><td style={styles.td}>{c.salesType}</td></tr>)}</tbody></table></div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 16 }}><div style={{ color: "#64748b" }}>{bulkSelected.length} customer(s) selected</div><div style={{ display: "flex", gap: 8 }}><Button secondary onClick={() => setBulkOpen(false)}>Cancel</Button><Button onClick={bulkSave}>Create Planned Visits</Button></div></div>
      </Modal>
    </div>
  );
}

function AuditTab({ auditLog, session }) {
  const [search, setSearch] = useState("");
  const rows = auditLog.filter((a) => !a.deletedAt).filter((a) => JSON.stringify(a).toLowerCase().includes(search.toLowerCase())).sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  return (
    <div style={{ display: "grid", gap: 16 }}>
      <PermissionBanner role={session.role} title="Audit trail permissions" message="This audit list is visible only to NSM and Sales Operations Admin." />
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 12 }}><div><h3 style={{ margin: 0 }}>Audit Trail</h3><div style={{ color: "#64748b", fontSize: 14 }}>Every important action is logged.</div></div><input style={styles.input} placeholder="Search audit trail" value={search} onChange={(e) => setSearch(e.target.value)} /></div>
        <div style={styles.tableWrap}><table style={styles.table}><thead><tr>{["Time", "Actor", "Role", "Action", "Entity", "Summary"].map((h) => <th key={h} style={styles.th}>{h}</th>)}</tr></thead><tbody>{rows.map((a) => <tr key={a.id}><td style={styles.td}>{new Date(a.createdAt).toLocaleString()}</td><td style={styles.td}>{a.actorName}</td><td style={styles.td}>{a.actorRole}</td><td style={styles.td}><Badge>{a.action}</Badge></td><td style={styles.td}>{a.entityType} • {a.entityId}</td><td style={styles.td}>{a.summary}</td></tr>)}</tbody></table></div>
      </Card>
    </div>
  );
}

function UsersTab({ authUsers, managers, setDb, session }) {
  const [search, setSearch] = useState("");
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({ email: "", password: "1234", managerId: "", role: "ASM", name: "", active: true });
  const activeUsers = authUsers.filter((u) => !u.deletedAt).filter((u) => JSON.stringify(u).toLowerCase().includes(search.toLowerCase()));
  const activeManagers = managers.filter((m) => !m.deletedAt);
  useEffect(() => {
    setForm(editing ? { email: editing.email, password: editing.password || "1234", managerId: editing.managerId, role: editing.role, name: editing.name, active: editing.active !== false } : { email: "", password: "1234", managerId: "", role: "ASM", name: "", active: true });
  }, [editing, open]);
  const saveUser = () => {
    if (!form.email || !form.password || !form.managerId) return window.alert("Email, password, and manager are required.");
    setDb((db) => {
      const manager = db.managers.find((m) => m.id === form.managerId);
      const payload = withMeta({ id: editing?.id || generateId("U", db.authUsers), email: form.email, password: form.password, managerId: form.managerId, role: form.role, name: form.name || manager?.name || form.email, active: form.active, createdAt: editing?.createdAt || nowISO() }, session.email);
      return addAuditEntry({ ...db, authUsers: editing ? db.authUsers.map((u) => (u.id === editing.id ? payload : u)) : [...db.authUsers, payload] }, session, editing ? "update" : "create", "auth_user", payload.id, `${editing ? "Updated" : "Created"} auth user ${payload.email}`);
    });
    setOpen(false);
    setEditing(null);
  };
  const disableUser = (id) => {
    const target = activeUsers.find((u) => u.id === id);
    if (!window.confirm(`Disable ${target?.email || "this user"}?`)) return;
    setDb((db) => addAuditEntry({ ...db, authUsers: db.authUsers.map((u) => u.id === id ? { ...u, active: false, updatedAt: nowISO(), updatedBy: session.email } : u) }, session, "disable", "auth_user", id, `Disabled auth user ${target?.email || id}`));
  };
  const resetPasswordLocal = (id) => {
    const target = activeUsers.find((u) => u.id === id);
    const nextPassword = window.prompt(`Enter a new password for ${target?.email || "this user"}:`, "1234");
    if (!nextPassword) return;
    setDb((db) => addAuditEntry({ ...db, authUsers: db.authUsers.map((u) => u.id === id ? { ...u, password: nextPassword, updatedAt: nowISO(), updatedBy: session.email } : u) }, session, "reset_password_local", "auth_user", id, `Locally reset password for ${target?.email || id}`));
  };
  return (
    <div style={{ display: "grid", gap: 16 }}>
      <PermissionBanner role={session.role} title="User access management" message="These user changes are staged locally and then merged to backend when you push sync." />
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 12 }}><div><h3 style={{ margin: 0 }}>User Access</h3><div style={{ color: "#64748b", fontSize: 14 }}>Manage login users and map them to managers.</div></div><div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}><input style={styles.input} placeholder="Search users" value={search} onChange={(e) => setSearch(e.target.value)} /> <Button onClick={() => { setEditing(null); setOpen(true); }}>Add User</Button></div></div>
        <div style={styles.tableWrap}><table style={styles.table}><thead><tr>{["Name", "Email", "Role", "Manager", "Status"].map((h) => <th key={h} style={styles.th}>{h}</th>)}<th style={styles.th}>Actions</th></tr></thead><tbody>{activeUsers.map((u) => <tr key={u.id}><td style={styles.td}>{u.name}</td><td style={styles.td}>{u.email}</td><td style={styles.td}><Badge>{u.role}</Badge></td><td style={styles.td}>{activeManagers.find((m) => m.id === u.managerId)?.name || "—"}</td><td style={styles.td}>{u.active ? "Active" : "Inactive"}</td><td style={styles.td}><div style={{ display: "flex", gap: 8 }}><Button secondary onClick={() => { setEditing(u); setOpen(true); }}>Edit</Button><Button secondary onClick={() => resetPasswordLocal(u.id)}>Reset Password</Button><Button secondary onClick={() => disableUser(u.id)}>Disable</Button></div></td></tr>)}</tbody></table></div>
      </Card>
      <Modal open={open} onClose={() => setOpen(false)} title={editing ? "Edit User" : "Add User"}>
        <div style={styles.grid2}>
          <Field label="Name"><input style={styles.input} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
          <Field label="Email"><input style={styles.input} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
          <Field label="Password"><input style={styles.input} value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></Field>
          <Field label="Role"><select style={styles.select} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>{ROLES.map((role) => <option key={role} value={role}>{role}</option>)}</select></Field>
          <Field label="Manager"><select style={styles.select} value={form.managerId} onChange={(e) => setForm({ ...form, managerId: e.target.value })}><option value="">Select manager</option>{activeManagers.map((m) => <option key={m.id} value={m.id}>{m.name}</option>)}</select></Field>
          <Field label="Active"><label><input type="checkbox" checked={form.active} onChange={(e) => setForm({ ...form, active: e.target.checked })} /> Active</label></Field>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}><Button secondary onClick={() => setOpen(false)}>Cancel</Button><Button onClick={saveUser}>Save User</Button></div>
      </Modal>
    </div>
  );
}

function ExportTab({ managers, customers, visits, auditLog, authUsers, settings, setSettings, onPull, onPush, syncStatus, syncing, session }) {
  const visitRows = visits.filter((v) => !v.deletedAt).map((v) => {
    const manager = managers.find((m) => m.id === v.managerId);
    const customer = customers.find((c) => c.id === v.customerId);
    return { Visit_ID: v.id, Visit_Date: v.visitDate, Manager_ID: v.managerId, Manager_Name: manager?.name || "", Manager_Role: manager?.role || "", Customer_ID: v.customerId, Customer_Name: customer?.name || "", Region: customer?.region || "", State: customer?.state || "", Sales_Type: v.salesType, Visit_Type: v.visitType, Planned_At_Start_Of_Day: v.plannedAtStartOfDay, Objective: v.objective, Outcome: v.outcome, Comments: v.comments, Order_Made: v.orderMade, Order_Qty_MT: v.orderQtyMt, Order_Value_NGN: v.orderValueNgn, Pipeline_Status: v.pipelineStatus, Next_Action_Date: v.nextActionDate, Visit_Status: v.visitStatus, Updated_At: v.updatedAt, Updated_By: v.updatedBy };
  });
  const auditRows = auditLog.filter((a) => !a.deletedAt);
  const activeManagers = managers.filter((m) => !m.deletedAt);
  const activeCustomers = customers.filter((c) => !c.deletedAt);
  const activeAuthUsers = authUsers.filter((u) => !u.deletedAt);
  return (
    <div style={{ display: "grid", gap: 16 }}>
      <PermissionBanner role={session.role} title="Export and sync permissions" message="Use this section to sync with Apps Script backend and export current data." />
      <Card>
        <h3 style={{ marginTop: 0 }}>Backend Sync Settings</h3>
        <div style={styles.grid2}>
          <Field label="Backend URL"><input style={styles.input} value={settings.appsScriptUrl} readOnly /></Field>
          <Field label="Logo URL"><input style={styles.input} value={settings.logoUrl || ""} readOnly /></Field>
          <Field label="Auto Pull"><label><input type="checkbox" checked={Boolean(settings.autoPullEnabled)} onChange={(e) => setSettings((s) => ({ ...s, autoPullEnabled: e.target.checked }))} /> Refresh automatically</label></Field>
          <Field label="Auto Pull Seconds"><input style={styles.input} type="number" min="15" value={settings.autoPullSeconds || 60} onChange={(e) => setSettings((s) => ({ ...s, autoPullSeconds: Math.max(15, Number(e.target.value || 60)) }))} /></Field>
          <Field label="Auto Push"><label><input type="checkbox" checked={Boolean(settings.autoPushEnabled)} onChange={(e) => setSettings((s) => ({ ...s, autoPushEnabled: e.target.checked }))} /> Push automatically</label></Field>
          <Field label="Auto Push Seconds"><input style={styles.input} type="number" min="30" value={settings.autoPushSeconds || 120} onChange={(e) => setSettings((s) => ({ ...s, autoPushSeconds: Math.max(30, Number(e.target.value || 120)) }))} /></Field>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 16 }}><Button secondary onClick={onPull} disabled={syncing}>Pull & Merge</Button><Button onClick={onPush} disabled={syncing}>Push Incremental Merge</Button></div>
        <div style={{ marginTop: 12, border: "1px solid #ecfdf5", borderRadius: 12, padding: 12, background: "#f8fafc" }}>{syncStatus}</div>
      </Card>
      <Card>
        <h3 style={{ marginTop: 0 }}>Exports</h3>
        <div style={styles.grid4}>
          <Button secondary onClick={() => downloadCSV("tmdk_managers.csv", activeManagers)}>Export Managers</Button>
          <Button secondary onClick={() => downloadCSV("tmdk_customers.csv", activeCustomers)}>Export Customers</Button>
          <Button secondary onClick={() => downloadCSV("tmdk_visits.csv", visitRows)}>Export Visits</Button>
          <Button secondary onClick={() => downloadCSV("tmdk_audit_log.csv", auditRows)}>Export Audit Log</Button>
          <Button secondary onClick={() => downloadCSV("tmdk_auth_users.csv", activeAuthUsers)}>Export Auth Users</Button>
        </div>
      </Card>
    </div>
  );
}

function ChangePasswordDialog({ open, onClose, onSubmit }) {
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [status, setStatus] = useState("");
  useEffect(() => {
    if (!open) {
      setCurrentPassword("");
      setNewPassword("");
      setStatus("");
    }
  }, [open]);
  return (
    <Modal open={open} onClose={onClose} title="Change Password">
      <div style={{ display: "grid", gap: 12 }}>
        <Field label="Current Password"><input style={styles.input} type="password" value={currentPassword} onChange={(e) => setCurrentPassword(e.target.value)} /></Field>
        <Field label="New Password"><input style={styles.input} type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} /></Field>
        {status ? <div style={{ border: "1px solid #ecfdf5", borderRadius: 12, padding: 12, background: "#f8fafc" }}>{status}</div> : null}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}><Button secondary onClick={onClose}>Cancel</Button><Button onClick={async () => { try { setStatus("Changing password..."); await onSubmit(currentPassword, newPassword); setStatus("Password changed successfully."); onClose(); } catch (error) { setStatus(error instanceof Error ? error.message : "Password change failed."); } }}>Update Password</Button></div>
      </div>
    </Modal>
  );
}

export default function TMDKSalesForceApp() {
  const [db, setDb] = useLocalCache();
  const [settings, setSettings] = useSettings();
  const [session, setSession] = useBackendSession();
  const [syncStatus, setSyncStatus] = useState("Waiting for backend connection.");
  const [syncing, setSyncing] = useState(false);
  const [filters, setFilters] = useState({ from: "", to: "", region: "all", state: "all", managerId: "all", pipeline: "all" });
  const [showChangePassword, setShowChangePassword] = useState(false);
  const [tab, setTab] = useState("dashboard");

  const currentManager = useMemo(() => db.managers.find((m) => !m.deletedAt && m.id === session?.managerId) || null, [db.managers, session]);
  const visibleManagers = useMemo(() => getVisibleManagers(db.managers, currentManager), [db.managers, currentManager]);
  const isAdmin = session ? ADMIN_ROLES.includes(session.role) : false;

  const refreshFromBackend = async () => {
    if (!session) return;
    setSyncing(true);
    try {
      setSyncStatus("Pulling latest filtered data from backend...");
      const remoteDb = await backendRead(settings, session);
      setDb((prev) => mergeDatabases(prev, remoteDb));
      setSyncStatus(`Pull completed at ${new Date().toLocaleTimeString()}.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Pull failed.";
      setSyncStatus(message);
      if (/session/i.test(message)) setSession(null);
      throw error;
    } finally {
      setSyncing(false);
    }
  };

  const pushToBackend = async () => {
    if (!session) return;
    setSyncing(true);
    try {
      setSyncStatus("Pushing incremental merge to backend...");
      const workingDb = addAuditEntry(db, session, "sync_push_client", "sync", "google", "Client pushed incremental merge to backend");
      setDb(workingDb);
      await backendMerge(settings, session, workingDb);
      setSyncStatus(`Push completed at ${new Date().toLocaleTimeString()}.`);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Push failed.";
      setSyncStatus(message);
      if (/session/i.test(message)) setSession(null);
      throw error;
    } finally {
      setSyncing(false);
    }
  };

  useEffect(() => {
    if (!session || !settings.appsScriptUrl) return;
    backendGet(settings, "session", session)
      .then(() => setSyncStatus("Backend session active."))
      .catch((error) => {
        setSyncStatus(error instanceof Error ? error.message : "Session validation failed.");
        setSession(null);
      });
  }, [session?.token, settings.appsScriptUrl]);

  useEffect(() => {
    if (!session || !settings.appsScriptUrl) return;
    refreshFromBackend().catch(() => {});
  }, [session?.token, settings.appsScriptUrl]);

  useEffect(() => {
    if (!session || !settings.appsScriptUrl || !settings.autoPullEnabled) return;
    const intervalMs = Math.max(15, Number(settings.autoPullSeconds || 60)) * 1000;
    const id = window.setInterval(() => { refreshFromBackend().catch(() => {}); }, intervalMs);
    return () => window.clearInterval(id);
  }, [session?.token, settings.appsScriptUrl, settings.autoPullEnabled, settings.autoPullSeconds]);

  useEffect(() => {
    if (!session || !settings.appsScriptUrl || !settings.autoPushEnabled) return;
    const intervalMs = Math.max(30, Number(settings.autoPushSeconds || 120)) * 1000;
    const id = window.setInterval(() => { pushToBackend().catch(() => {}); }, intervalMs);
    return () => window.clearInterval(id);
  }, [session?.token, settings.appsScriptUrl, settings.autoPushEnabled, settings.autoPushSeconds, db]);

  if (!session || !currentManager) {
    return <LoginScreen settings={settings} setSettings={setSettings} onLogin={setSession} />;
  }

  const tabs = [
    { id: "dashboard", label: "Dashboard" },
    { id: "visits", label: "Visits" },
    { id: "customers", label: "Customers" },
    { id: "managers", label: "Managers" },
    ...(isAdmin ? [{ id: "users", label: "Users" }, { id: "audit", label: "Audit" }, { id: "exports", label: "Exports" }] : []),
  ];

  return (
    <div style={styles.page}>
      <div style={styles.container}>
        <AppHeader session={session} logoUrl={settings.logoUrl} onLogout={async () => { try { await backendLogout(settings, session); } catch (_) {} setSession(null); }} onChangePassword={() => setShowChangePassword(true)} />
        <div style={styles.grid4}>{[
          { label: "Managers", value: db.managers.filter((m) => !m.deletedAt).length, icon: "👥" },
          { label: "Customers", value: db.customers.filter((c) => !c.deletedAt).length, icon: "🏢" },
          { label: "Visits", value: db.visits.filter((v) => !v.deletedAt).length, icon: "📍" },
          { label: "Current Access", value: session.role, icon: "🛡️" },
          { label: "Audit Entries", value: db.auditLog.filter((a) => !a.deletedAt).length, icon: "🧾" },
        ].map((kpi) => <Card key={kpi.label}><div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}><div style={{ color: "#64748b", fontSize: 13 }}>{kpi.label}</div><div style={{ width: 34, height: 34, borderRadius: 12, background: "#f0fdf4", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>{kpi.icon}</div></div><div style={{ fontSize: 24, fontWeight: 700, marginTop: 10 }}>{kpi.value}</div></Card>)}</div>
        <Card><div>{syncStatus}</div></Card>
        <Filters filters={filters} setFilters={setFilters} managers={visibleManagers} customers={db.customers.filter((c) => !c.deletedAt)} />
        <div style={styles.tabs}>{tabs.map((t) => <button key={t.id} style={styles.tab(tab === t.id)} onClick={() => setTab(t.id)}>{t.label}</button>)}</div>
        {tab === "dashboard" ? <Dashboard visits={db.visits} customers={db.customers} managers={db.managers} currentManager={currentManager} filters={filters} /> : null}
        {tab === "visits" ? <VisitsTab visits={db.visits} customers={db.customers} managers={db.managers} setDb={setDb} currentManager={currentManager} session={session} /> : null}
        {tab === "customers" ? <CustomersTab customers={db.customers} setDb={setDb} session={session} /> : null}
        {tab === "managers" ? <ManagersTab managers={db.managers} authUsers={db.authUsers} setDb={setDb} currentManager={currentManager} session={session} /> : null}
        {tab === "users" && isAdmin ? <UsersTab authUsers={db.authUsers} managers={db.managers} setDb={setDb} session={session} /> : null}
        {tab === "audit" && isAdmin ? <AuditTab auditLog={db.auditLog} session={session} /> : null}
        {tab === "exports" && isAdmin ? <ExportTab managers={db.managers} customers={db.customers} visits={db.visits} auditLog={db.auditLog} authUsers={db.authUsers} settings={settings} setSettings={setSettings} onPull={refreshFromBackend} onPush={pushToBackend} syncStatus={syncStatus} syncing={syncing} session={session} /> : null}
        <ChangePasswordDialog open={showChangePassword} onClose={() => setShowChangePassword(false)} onSubmit={(currentPassword, newPassword) => backendChangePassword(settings, session, currentPassword, newPassword)} />
      </div>
    </div>
  );
}
