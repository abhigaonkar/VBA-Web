import React, { useState, useEffect } from "react";

// -- Input Styles ---

// General input style
export const inputSt: React.CSSProperties = {
  background: "#f1f5f9",
  border: "1.4px solid #e5e9f2",
  borderRadius: 8,
  padding: "7px 11px",
  fontSize: 15,
  width: "100%",
  outline: "none",
  boxSizing: "border-box",
};

// -- Navbtn
const navBtn = {
  marginTop: 0,
  background: "#2563eb",
  border: "none",
  color: "#fff",
  padding: "7px 16px",
  borderRadius: 7,
  fontWeight: 700,
  fontSize: 15,
  letterSpacing: ".02em",
  cursor: "pointer",
  outline: "none",
  marginRight: 5,
  marginBottom: 4
} as React.CSSProperties;

// Input with error example, use inline or in a helper
export const inputStError: React.CSSProperties = {
  ...inputSt,
  border: "1.6px solid #dc2626",
};

// Mini input (for batch/array or table row cells)
export const miniInputSt: React.CSSProperties = {
  ...inputSt,
  padding: "5px 7px",
  fontSize: 14,
  width: 80,
};

// Submit/save button 
export const submitBtn: React.CSSProperties = {
  marginTop: 17,
  background: "#2563eb",
  border: "none",
  color: "#fff",
  padding: "11px 22px",
  borderRadius: 8,
  fontWeight: 800,
  fontSize: 16,
  letterSpacing: ".02em",
  cursor: "pointer",
  outline: "none",
  transition: "background 0.2s",
};

// For secondary/cancel buttons
export const cancelBtn: React.CSSProperties = {
  ...submitBtn,
  background: "#64748b",
};

// Example badge style (you already have)
export function badge(color: string): React.CSSProperties {
  return {
    display: "inline-block",
    padding: '2px 11px',
    borderRadius: 14,
    background: color + '18',
    color: color,
    border: '1.1px solid ' + color,
    fontWeight: 700,
    fontSize: 13,
    marginLeft: 1
  }
}

// --- Roles & Users ---
const USERS = [
  { username: "admin", password: "admin", role: "Admin", display: "Ashley Admin" },
  { username: "manager", password: "manager", role: "Manager", display: "Morgan Manager" },
  { username: "op", password: "op", role: "Operator", display: "Olivia Operator" }
] as const;
type Role = "Admin" | "Manager" | "Operator";
type UserType = (typeof USERS)[number];

// --- Ovens/Channels ---
const OVEN_IDS = ["NV0A", "NV0B", "NV0C", "NV0D"];
type FurnaceStatus = "Empty" | "Loading" | "Baking" | "Cooling" | "Needs QA" | "Complete";
type FurnaceChannel = {
  id: string;
  status: FurnaceStatus;
  currentBatch?: string;
  operator?: string;
};

// --- Navigation ---
type LinkDef = {
  key: string;
  label: string;
  icon: React.ReactNode;
  roles: Role[];
};
const lucide = {
  dashboard: <svg width={22} height={22} stroke="currentColor" fill="none" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg>,
  assembly: <svg width={22} height={22} stroke="currentColor" fill="none" viewBox="0 0 24 24"><path d="M4 17v-7c0-1.104.896-2 2-2h4V4a2 2 0 1 1 4 0v4h4a2 2 0 0 1 2 2v7"/><rect x="4" y="17" width="16" height="3" rx="1"/></svg>,
  sand: <svg width={22} height={22} stroke="currentColor" fill="none" viewBox="0 0 24 24"><ellipse cx="12" cy="7" rx="7" ry="3"/><path d="M19 7v6.5c0 1.38-3.13 2.5-7 2.5s-7-1.12-7-2.5V7"/><path d="M19 13.5V17c0 1.38-3.13 2.5-7 2.5S5 18.38 5 17v-3.5"/></svg>,
  serial: <svg width={22} height={22} stroke="currentColor" fill="none" viewBox="0 0 24 24"><rect x="7" y="3" width="10" height="6" rx="2"/><rect x="3" y="15" width="7" height="6" rx="2"/><rect x="14" y="15" width="7" height="6" rx="2"/><path d="M12 9v6"/></svg>,
  hv: <svg width={22} height={22} stroke="currentColor" fill="none" viewBox="0 0 24 24"><polyline points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>,
  admin: <svg width={22} height={22} stroke="currentColor" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="7" r="4"/><path d="M5.5 21a1.65 1.65 0 0 1 1.35-2.67h10.3A1.65 1.65 0 0 1 18.5 21z"/></svg>
};
const ALL_LINKS: LinkDef[] = [
  { key: "dashboard", label: "Dashboard", icon: lucide.dashboard, roles: ["Admin", "Manager", "Operator"] },
  { key: "lamp", label: "Lamp Rod", icon: lucide.assembly, roles: ["Manager", "Operator"] },
  { key: "bake", label: "Baking Sand", icon: lucide.sand, roles: ["Manager", "Operator"] },
  { key: "serial", label: "Serializing", icon: lucide.serial, roles: ["Manager", "Operator"] },
  { key: "hv", label: "HV Pack", icon: lucide.hv, roles: ["Manager", "Operator"] },
  { key: "admin", label: "Admin", icon: lucide.admin, roles: ["Admin"] }
];

// --- Record Type ---
type LampRodEntryType = {
  id: string;
  furnaceId: string;
  view: "serialized" | "nonserial" | "sand";
  createdBy: string;
  createdAt: string;
  oven?: string;
  palletCard?: string;
  bakeHours?: string;
  processOrder?: string;
  bakedQty?: string;
  operator: string;
  comments?: string;
  unbakedItem?: string;
  bakedItem?: string;
  batches?: string[];
  hampers?: string[];
  strapping?: string;
  batch?: string;
  unbakedQty?: string;
  hamper?: string;
  scrapCode?: string;
  sandItem?: string;
  sandBatch?: string;
  sourceLoc?: string;
  destFurnace?: string;
  qty?: string;
};

// --- Persistence ---
const LAMPROD_LS_KEY = "lampRodRecordsV3";
const FURNACES_LS_KEY = "furnaceChannelsV3";
function saveRecordsToStorage(records: LampRodEntryType[]) {
  localStorage.setItem(LAMPROD_LS_KEY, JSON.stringify(records));
}
function getRecordsFromStorage(): LampRodEntryType[] {
  try { const d = localStorage.getItem(LAMPROD_LS_KEY); if (!d) return []; return JSON.parse(d); } catch { return []; }
}
function getInitialFurnaceChannels(): FurnaceChannel[] {
  try {
    const stored = localStorage.getItem(FURNACES_LS_KEY);
    if (stored) return JSON.parse(stored);
  } catch { }
  return OVEN_IDS.map(id => ({ id, status: "Empty" as FurnaceStatus }));
}

// --- App ---
export default function App() {
  // Login & Navigation
  const [user, setUser] = useState<UserType | null>(null);
  const [loginError, setLoginError] = useState("");
  const [sideTab, setSideTab] = useState("dashboard");
  const [lampEntryView, setLampEntryView] = useState<null | { ovenId: string; entryType: "serialized" | "nonserial" | "sand"; }>(null);

  // Oven states
  const [furnaceChannels, setFurnaceChannels] = useState<FurnaceChannel[]>(getInitialFurnaceChannels());
  useEffect(() => { localStorage.setItem(FURNACES_LS_KEY, JSON.stringify(furnaceChannels)); }, [furnaceChannels]);

  // Lamp rod entries
  const [lampRodRecords, setLampRodRecords] = useState<LampRodEntryType[]>(getRecordsFromStorage());
  useEffect(() => { saveRecordsToStorage(lampRodRecords); }, [lampRodRecords]);

  function handleLogout() { setUser(null); setSideTab("dashboard"); setLampEntryView(null); }

  // Sidebar navigation, per-role
  const SIDEBAR_LINKS = user ? ALL_LINKS.filter(link => link.roles.includes(user.role as Role)) : [];

  // Derived user display
  const { display, username } = user || { display: "", username: "" };
  const initials = user ? (display?.split(" ").map(s => s[0]).join("").slice(0, 2).toUpperCase() || username[0].toUpperCase()) : "";

  // ----------------- UI Render -------------------
  if (!user) return (
    <LoginForm
      onLogin={(u) => { setUser(u); setSideTab("dashboard"); setLampEntryView(null); setLoginError(""); }}
      error={loginError}
      setError={setLoginError}
    />
  );

  return (
    <div style={{ background: "#f8fafd", minHeight: "100vh", display: "flex", fontFamily: "Inter,sans-serif" }}>
      {/* Sidebar */}
      <aside style={{ width: 260, background: "#1c2230", color: "#e2e8f0", minHeight: "100vh", position: "sticky", top: 0, display: 'flex', flexDirection: 'column', zIndex: 11 }}>
        <div style={{ height: 64, borderBottom: "1px solid #232b40", display: "flex", alignItems: "center", gap: 13, padding: "0 24px" }}>
          <div style={{ width: 36, height: 36, borderRadius: 9, background: "#2563eb", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <svg width={20} height={20} stroke="#fff" fill="none" viewBox="0 0 24 24"><rect x="2.7" y="2.7" width="18.6" height="18.6" rx="4.2" strokeWidth={2.5}></rect>
              <circle cx="12" cy="12" r="4.2" strokeWidth={2}></circle></svg>
          </div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 16 }}>Modern Manufacturing</div>
            <div style={{ fontSize: 10.5, textTransform: "uppercase", letterSpacing: 2, opacity: .4, fontWeight: 700 }}>Dashboard</div>
          </div>
        </div>
        <nav aria-label="Sidebar" style={{ margin: "26px 0 0 0", flex: 1, overflowY: "auto" }}>
          <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
            {SIDEBAR_LINKS.map(link => (
              <li key={link.key} style={{ marginBottom: 2 }}>
                <button
                  onClick={() => { setSideTab(link.key); setLampEntryView(null); }}
                  style={{
                    width: "100%", border: "none",
                    background: sideTab === link.key ? "linear-gradient(90deg,rgba(37,99,235,0.15),rgba(37,99,235,0.23) 80%,#2563eb11 100%)" : "none",
                    color: sideTab === link.key ? "#2563eb" : "#e2e8f0", display: "flex", alignItems: "center", gap: 13,
                    fontWeight: sideTab === link.key ? 700 : 600, fontSize: 15, padding: "10px 18px", borderRadius: 10, cursor: "pointer", outline: 'none',
                    boxShadow: sideTab === link.key ? "0 1px 8px #2563eb14" : "none"
                  }}>
                  <span style={{
                    background: sideTab === link.key ? "#e0e8fb" : "#222c43",
                    color: sideTab === link.key ? "#2563eb" : "#c4d0e4",
                    borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", width: 30, height: 30,
                  }}>{link.icon}</span>
                  <span>{link.label}</span>
                </button>
              </li>
            ))}
          </ul>
        </nav>
        <div style={{
          borderTop: "1px solid #232b40", padding: "19px 24px", display: "flex", alignItems: "center", gap: 13
        }}>
          <div style={{
            width: 33, height: 33, borderRadius: "50%",
            background: "linear-gradient(130deg,#2563ebcc,#6366f186)", display: "flex", alignItems: "center", justifyContent: "center",
            color: "#fff", fontWeight: 700, fontSize: 15, boxShadow: "0 1px 6px #2563eb22"
          }}>{initials}</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>{display || username}</div>
            <div style={{ fontSize: 11, opacity: .68, fontWeight: 600 }}>{user.role}</div>
          </div>
          <button onClick={handleLogout} style={{
            marginLeft: "auto", background: "#fff", color: "#2563eb",
            border: "none", borderRadius: 8, fontWeight: 800, fontSize: 15, padding: "0 .85em", height: 34, cursor: "pointer"
          }}>Logout</button>
        </div>
      </aside>
      {/* Main Content */}
      <main style={{ flex: 1, minWidth: 0, background: "#f8fafd", minHeight: "100vh" }}>
        <header style={{
          height: 62, borderBottom: "1.5px solid #e5eaf0",
          display: "flex", alignItems: "center", justifyContent: "flex-end",
          padding: "0 26px", background: "#fff", position: "sticky", top: 0, zIndex: 12,
        }}>
          <span style={{ fontWeight: 500, fontSize: 15, color: "#475274", marginRight: 12 }}>
            {display || username} ({user.role})
          </span>
          <span style={{
            display: 'inline-flex',
            width: 37, height: 37, borderRadius: "50%", background: "linear-gradient(135deg,#2563eb 80%,#64748b 120%)",
            color: "#fff", fontSize: 15, fontWeight: 800, alignItems: "center", justifyContent: "center"
          }}>{initials}</span>
        </header>
        <div style={{ maxWidth: 1250, margin: "0 auto", padding: "36px 26px 48px 26px" }}>
          {sideTab === "dashboard" && <Dashboard />}
          {sideTab === "lamp" &&
            <>
              <LampRodFurnaceMonitor
                channels={furnaceChannels}
                onStatusUpdate={(ovenId, newStatus) => {
                  setFurnaceChannels(cs => cs.map(ch =>
                    ch.id === ovenId ? { ...ch, status: newStatus, ...(newStatus === "Empty" ? { currentBatch: undefined, operator: undefined } : {}) } : ch
                  ));
                }}
                onStartBatch={(ovenId) => {
                  // Only operators & managers can start
                  if (user.role === "Operator" || user.role === "Manager") setLampEntryView({ ovenId, entryType: "serialized" });
                }}
                user={user}
                showStartButtons={user.role === "Operator" || user.role === "Manager"}
              />
              {lampEntryView && (
                <LampRodDataEntry
                  entryType={lampEntryView.entryType}
                  operator={initials}
                  ovenId={lampEntryView.ovenId}
                  onSave={rec => {
                    setLampRodRecords(prev => [{ ...rec, id: Date.now() + "_" + Math.random(), createdBy: user.username, createdAt: new Date().toISOString(), furnaceId: lampEntryView.ovenId }, ...prev]);
                    setFurnaceChannels(cs => cs.map(ch => ch.id === lampEntryView.ovenId ? { ...ch, currentBatch: rec.batch || rec.palletCard, operator: initials, status: "Baking" } : ch));
                    setLampEntryView(null);
                  }}
                  onCancel={() => setLampEntryView(null)}
                />
              )}
              {!lampEntryView && (
                <LampRodRecords
                  records={lampRodRecords}
                  onNew={(type, ovenId) => setLampEntryView({ entryType: type, ovenId })}
                  user={user}
                  furnaceChannels={furnaceChannels}
                />
              )}
            </>
          }
          {/* ...other tabs as before... */}
          {sideTab === "bake" && (
            <Section
              icon={lucide.sand}
              name="Baking Sand Workflow"
              color="#eab308"
              steps={[
                { label: "Sand Batch Prep", caption: "Stage and test raw sand", done: true },
                { label: "FIFO Distribution", caption: "Deliver to furnaces, FIFO rule", done: true },
                { label: "Bagging", caption: "Pack, print lots, scan into system", done: false },
                { label: "Warehouse Transfer", caption: "Move to warehouse, SAP update", done: false }
              ]}
            />
          )}
          {sideTab === "serial" && (
            <Section
              icon={lucide.serial}
              name="Serializing"
              color="#a21caf"
              steps={[
                { label: "Serial Batch Generation", caption: "Create code series for new lots", done: true },
                { label: "Print & Apply", caption: "Affix, scan for QA", done: false },
                { label: "System Log", caption: "Track lot movement", done: false }
              ]}
            />
          )}
          {sideTab === "hv" && (
            <Section
              icon={lucide.hv}
              name="HV High Voltage Pack"
              color="#0ea5e9"
              steps={[
                { label: "Staging", caption: "Assemble HV components", done: true },
                { label: "Component QA", caption: "Verify matches & log", done: true },
                { label: "Smart Packing", caption: "Pack, print, and scan", done: false },
                { label: "Outbound", caption: "Move to dock, SAP stock", done: false }
              ]}
            />
          )}
          {sideTab === "admin" && user.role === "Admin" && (
            <div style={{
              background: "#fff",
              borderRadius: 14,
              boxShadow: '0 1.5px 8px #24304418',
              margin: '2.2em 0 28px 0',
              padding: '2em 2em 2em 2em',
              border: '1px solid #e2e8f0'
            }}>
              <div style={{ fontSize: 17, color: '#64748b', fontWeight: 700, marginBottom: 24 }}>User/Role Admin</div>
              <table style={{ width: '100%', background: '#f8fafc', borderCollapse: 'collapse' }}>
                <thead>
                  <tr style={{ background: '#243044', color: '#fff' }}>
                    <th style={{ padding: '9px 11px' }}>Username</th>
                    <th style={{ padding: '9px 11px' }}>Role</th>
                  </tr>
                </thead>
                <tbody>
                  {USERS.map(user => (
                    <tr key={user.username}>
                      <td style={{ padding: '9px 11px' }}>{user.display}</td>
                      <td style={{ padding: '9px 11px' }}><span style={badge(user.role === "Admin" ? "#2563eb" : user.role === "Manager" ? "#10b981" : "#64748b")}>{user.role}</span></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

// --- Supporting Components ---
// ...Placeholders for Dashboard, Section, Input styles, LoginForm, LampRodFurnaceMonitor, LampRodDataEntry, LampRodRecords, badge etc...

function Dashboard() {
  // These are just example KPIs, you can expand/adjust as needed
  const kpis = [
    {
      label: "Production Output",
      value: "12,120",
      delta: "+6.2%",
      color: "#059669",
      sub: "units completed"
    },
    {
      label: "Sand Used",
      value: "12.9t",
      delta: "-2.1%",
      color: "#e11d48",
      sub: "last 30 days"
    },
    {
      label: "Serializations",
      value: "1,157",
      delta: "+3.7%",
      color: "#2563eb",
      sub: "lots tracked"
    },
    {
      label: "High Voltage Packs",
      value: "783",
      delta: "+0.9%",
      color: "#0ea5e9",
      sub: "units shipped"
    }
  ];

  return (
    <div>
      <h1 style={{ fontWeight: 800, fontSize: 26, marginBottom: 8, letterSpacing: "-0.02em", color: "#161b22" }}>
        Welcome back!
      </h1>
      <div style={{ color: "#64748b", marginBottom: 32 }}>Your manufacturing KPIs at a glance.</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "24px" }}>
        {kpis.map(kpi => (
          <div key={kpi.label}
            style={{
              background: "#fff", borderRadius: 18, boxShadow: "0 4px 18px #2233610c",
              border: "1px solid #e2e8f0", padding: "28px 26px 20px 26px", display: "flex", flexDirection: "column",
              transition: "box-shadow .18s", position: "relative"
            }}>
            <div style={{ fontSize: 14, color: "#a1a4b3", fontWeight: 600 }}>{kpi.label}</div>
            <div style={{ fontWeight: 900, fontSize: 27, color: "#1c2230", margin: "7px 0" }}>{kpi.value}</div>
            <div style={{ color: "#64748b", fontSize: 14, fontWeight: 700, margin: "1px 0 0 0" }}>{kpi.sub}</div>
            <div style={{ display: "flex", alignItems: "center", marginTop: 10 }}>
              <span style={{
                fontSize: 13.5,
                color: kpi.color,
                fontWeight: 700,
                display: "inline-flex", alignItems: "center", marginRight: 7
              }}>
                {kpi.delta.includes("+") ? "▲" : "▼"} {kpi.delta}
              </span>
              <span style={{ fontSize: 13, color: "#a1a4b3", fontWeight: 500 }}>last month</span>
            </div>
          </div>
        ))}
      </div>
      <div style={{ marginTop: 36, fontSize: 18, color: "#64748b", fontWeight: 600 }}>
        Production highlights coming soon...
      </div>
    </div>
  );
}

function Section({
  icon,
  name,
  color,
  steps
}: {
  icon: React.ReactNode,
  name: string,
  color: string,
  steps: { label: string; caption: string; done: boolean }[]
}) {
  return (
    <section>
      <div style={{
        display: 'flex', alignItems: 'center', gap: 16, background: "#fff", borderRadius: 16,
        boxShadow: "0 2px 10px #2233610f", border: "1px solid #ecf1f7", padding: "22px 28px", marginBottom: 18
      }}>
        <span style={{
          background: color,
          color: "#fff",
          borderRadius: 13, width: 46, height: 46, fontSize: 23, display: "flex", alignItems: "center", justifyContent: "center",
          boxShadow: "0 2px 12px " + color + "22"
        }}>
          {icon}
        </span>
        <span style={{ fontSize: 22, fontWeight: 800, color: "#222943" }}>{name} — Process Flow</span>
      </div>
      <ol style={{ margin: 0, padding: "0 0 0 10px" }}>
        {steps.map((s, i) => (
          <li key={i} style={{
            listStyle: "none", marginBottom: 22,
            background: "#fff", border: "1.5px solid #eef0f6", borderRadius: 13,
            boxShadow: s.done ? "0 1.5px 4px #68d3911c" : "0 2px 8px #2563eb0a",
            padding: "18px 23px",
            display: "flex", alignItems: "center", gap: 19
          }}>
            <span style={{
              width: 26, height: 26, borderRadius: "50%",
              background: s.done ? "#bbf7d0" : "#c7d2fe",
              color: s.done ? "#059669" : "#2563eb",
              fontWeight: 700, fontSize: 17, display: "flex", alignItems: "center", justifyContent: "center"
            }}>
              {s.done ? "✓" : i + 1}
            </span>
            <span>
              <span style={{ fontWeight: 700, color: "#1a1f2c", fontSize: 17 }}>{s.label}</span>
              <span style={{ display: "block", color: "#64748b", fontSize: 14, fontWeight: 500 }}>{s.caption}</span>
            </span>
          </li>
        ))}
      </ol>
    </section>
  );
}

function LoginForm({
  onLogin,
  error,
  setError
}: {
  onLogin: (u: UserType) => void,
  error: string,
  setError: (e: string) => void
}) {
  const [user, setUser] = useState("");
  const [pass, setPass] = useState("");
  const [touched, setTouched] = useState(false);

  function submit(e: React.FormEvent) {
    e.preventDefault();
    setTouched(true);
    const found = USERS.find(u => u.username === user && u.password === pass);
    if (!found) {
      setError("Invalid username or password.");
    } else {
      setError("");
      onLogin(found);
    }
  }

  return (
    <div style={{
      minHeight: '100vh', background: "#ecf0fa",
      display: 'flex', justifyContent: "center", alignItems: "center"
    }}>
      <form onSubmit={submit} style={{
        background: "#fff", borderRadius: 18, boxShadow: "0 4px 24px #22336115", maxWidth: 370,
        border: "1.5px solid #e0e7ef", padding: "37px 27px"
      }}>
        <div style={{ textAlign: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 29, fontWeight: 800, color: "#243044" }}>
            Modern Manufacturing
          </div>
          <div style={{ color: "#64748b", marginTop: 4, fontWeight: 500 }}>
            Sign in to your account
          </div>
        </div>
        <div style={{ margin: "18px 0 10px 0" }}>
          <input
            value={user}
            onChange={e => setUser(e.target.value)}
            placeholder="Username"
            style={{
              ...inputSt,
              border: error && touched ? "1.6px solid #dc2626" : inputSt.border
            }}
            autoFocus
            autoComplete="username"
          />
        </div>
        <div style={{ margin: "12px 0 8px 0" }}>
          <input
            type="password"
            value={pass}
            onChange={e => setPass(e.target.value)}
            placeholder="Password"
            style={{
              ...inputSt,
              border: error && touched ? "1.6px solid #dc2626" : inputSt.border
            }}
            autoComplete="current-password"
          />
        </div>
        {error && touched &&
          <div style={{ color: "#dc2626", marginBottom: 15, fontWeight: 700 }}>
            {error}
          </div>
        }
        <button type="submit" style={{
          ...submitBtn, width: "100%", marginTop: 10, padding: "11px 0"
        }}>Sign in</button>
        <div style={{ fontSize: 12, color: "#888", marginTop: 14 }}>
          Try one of:<br />
          <b>admin</b>, <b>manager</b>, <b>op</b> (use same as password)
        </div>
      </form>
    </div>
  );
}

function LampRodFurnaceMonitor({
  channels,
  onStatusUpdate,
  onStartBatch,
  user,
  showStartButtons
}: {
  channels: FurnaceChannel[];
  onStatusUpdate: (ovenId: string, newStatus: FurnaceStatus) => void;
  onStartBatch: (ovenId: string) => void;
  user: UserType;
  showStartButtons: boolean;
}) {
  // Map status to color
  function badgeColor(status: FurnaceStatus) {
    switch (status) {
      case "Empty": return "#cbd5e1";
      case "Loading": return "#a3e635";
      case "Baking": return "#facc15";
      case "Cooling": return "#60a5fa";
      case "Needs QA": return "#f87171";
      case "Complete": return "#059669";
      default: return "#64748b";
    }
  }
  return (
    <div style={{ display: "flex", gap: 22, marginBottom: 12, flexWrap: "wrap" }}>
      {channels.map(ch => (
        <div key={ch.id}
          style={{
            background: "#fff", borderRadius: 11, border: "1.5px solid #e7eaf3", boxShadow: "0 1.5px 8px #2430440e", padding: "15px 16px", minWidth: 170, minHeight: 125,
            display: "flex", flexDirection: "column", alignItems: "flex-start", position: "relative"
          }}>
          <div style={{ fontWeight: 800, fontSize: 17 }}>{ch.id}</div>
          <div style={{ fontWeight: 600, color: badgeColor(ch.status), marginBottom: 3 }}>
            <span style={{
              display: "inline-block", padding: "0 10px", borderRadius: 12,
              background: badgeColor(ch.status) + "22"
            }}>{ch.status}</span>
          </div>
          {ch.currentBatch && <div style={{ fontSize: 13, color: "#64748b" }}>Batch: {ch.currentBatch}</div>}
          {ch.operator && <div style={{ fontSize: 12, color: "#475274" }}>Op: {ch.operator}</div>}
          {(ch.status === "Empty" && (user.role === "Operator" || user.role === "Manager")) &&
            <button style={{
              ...submitBtn, fontSize: 13, padding: "3px 10px", margin: "14px 0 0 0"
            }} onClick={() => onStartBatch(ch.id)}>Start Batch</button>
          }
          {(ch.status !== "Empty" && (user.role === "Operator" || user.role === "Manager")) &&
            <div style={{ marginTop: 13 }}>
              <StatusDropdown status={ch.status} onChange={s => onStatusUpdate(ch.id, s)} />
            </div>
          }
        </div>
      ))}
    </div>
  );
}

// Helper: Status Dropdown
function StatusDropdown({ status, onChange }:
  { status: FurnaceStatus, onChange: (s: FurnaceStatus) => void }) {
  const statuses: FurnaceStatus[] = ["Empty", "Loading", "Baking", "Cooling", "Needs QA", "Complete"];
  return (
    <select value={status} onChange={e => onChange(e.target.value as FurnaceStatus)} style={{ ...inputSt, fontSize: 13, width: "auto" }}>
      {statuses.map(s => <option key={s}>{s}</option>)}
    </select>
  );
}

function LampRodDataEntry({
  entryType,
  operator,
  ovenId,
  onSave,
  onCancel
}: {
  entryType: "serialized" | "nonserial" | "sand",
  operator: string,
  ovenId: string,
  onSave: (entry: any) => void,
  onCancel: () => void
}) {
  // Field states
  const [fields, setFields] = useState<any>(() => {
    if (entryType === "serialized") return {
      oven: ovenId, palletCard: "", bakeHours: "", processOrder: "",
      unbakedItem: "", bakedItem: "", batches: ["", "", "", ""],
      bakedQty: "", hampers: ["", ""], strapping: "", operator, comments: ""
    };
    if (entryType === "nonserial") return {
      oven: ovenId, palletCard: "", bakeHours: "", processOrder: "",
      batch: "", bakedQty: "", unbakedQty: "", hamper: "", strapping: "",
      scrapCode: "", operator, comments: ""
    };
    // sand
    return {
      sandItem: "", sandBatch: "", sourceLoc: "", destFurnace: "", qty: "", operator, comments: ""
    };
  });
  const [touched, setTouched] = useState<{ [k: string]: boolean }>({});
  const [errors, setErrors] = useState<{ [k: string]: string }>({});
  const [saving, setSaving] = useState(false);

  // Select options
  const ovenOptions = ["NV0A", "NV0B", "NV0C", "NV0D"];
  const hamperOptions = ['66" Hamper', '70" Hamper', '86" Hamper', '110" Hamper'];
  const strappingOptions = ["147664", "147665"];
  const sandLocs = ["0015 (Warehouse)", "0100 (Storage)", "0020 (Floor)"];
  const sandFurnace = ["NF21", "NF22", "NF23"];
  const defectCodes = ["D001 (Crack)", "D002 (Bubble)", "D003 (Flat Spot)"];

  // Validation
  function validateFields(vals = fields) {
    let e: { [k: string]: string } = {};
    if (entryType === "serialized") {
      if (!vals.oven) e.oven = "Oven is required.";
      if (!vals.palletCard) e.palletCard = "Pallet Card # is required.";
      if (!vals.bakeHours) e.bakeHours = "Bake hours required.";
      if (!vals.processOrder) e.processOrder = "Process order required.";
      if (!vals.unbakedItem) e.unbakedItem = "Unbaked Item # is required.";
      if (!vals.bakedItem) e.bakedItem = "Baked Item # is required.";
      if (!vals.batches?.filter((v: string) => v).length) e.batches = "At least 1 batch is required.";
      if (!vals.bakedQty) e.bakedQty = "Baked quantity required.";
      if (!(vals.hampers?.[0] || vals.hampers?.[1])) e.hampers = "At least 1 hamper required.";
      if (!vals.strapping) e.strapping = "Strapping is required.";
    }
    if (entryType === "nonserial") {
      if (!vals.oven) e.oven = "Oven is required.";
      if (!vals.palletCard) e.palletCard = "Pallet Card # is required.";
      if (!vals.bakeHours) e.bakeHours = "Bake hours required.";
      if (!vals.processOrder) e.processOrder = "Process order required.";
      if (!vals.batch) e.batch = "Batch # is required.";
      if (!vals.bakedQty) e.bakedQty = "Baked quantity required.";
      if (!vals.unbakedQty) e.unbakedQty = "Unbaked quantity required.";
      if (!vals.hamper) e.hamper = "Hamper is required.";
      if (!vals.strapping) e.strapping = "Strapping is required.";
    }
    if (entryType === "sand") {
      if (!vals.sandItem) e.sandItem = "Sand item is required.";
      if (!vals.sandBatch) e.sandBatch = "Batch is required.";
      if (!vals.sourceLoc) e.sourceLoc = "Source location is required.";
      if (!vals.destFurnace) e.destFurnace = "Destination furnace is required.";
      if (!vals.qty) e.qty = "Quantity is required.";
      if (vals.qty && (isNaN(Number(vals.qty)) || Number(vals.qty) <= 0)) e.qty = "Enter a valid positive quantity.";
    }
    return e;
  }

  function setField(k: string, v: any) {
    setFields((f: any) => ({ ...f, [k]: v }));
    setTouched(t => ({ ...t, [k]: true }));
  }

  function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setTouched(Object.keys(fields).reduce((acc, k) => ({ ...acc, [k]: true }), {}));
    const errs = validateFields();
    setErrors(errs);
    if (Object.keys(errs).length === 0) {
      setSaving(true);
      setTimeout(() => {
        setSaving(false);
        onSave({
          ...fields,
          view: entryType
        });
      }, 400);
    }
  }

  useEffect(() => {
    if (Object.keys(touched).length) setErrors(validateFields());
  }, [fields, touched]);

  function fieldBlock(label: string, k: string, children: React.ReactNode) {
    const hasErr = !!errors[k] && touched[k];
    return (
      <div style={{ marginBottom: 15 }}>
        <label style={{ color: "#374151", fontWeight: 600, fontSize: 13 }}>{label}</label>
        {children}
        {hasErr && <div style={{ color: "#dc2626", fontWeight: 500, fontSize: 12, marginTop: 2 }}>{errors[k]}</div>}
      </div>
    );
  }
  function inputStyle(key: string, multi?: boolean) {
    const err = !!errors[key] && touched[key];
    return {
      background: "#f1f5f9",
      border: err ? "1.6px solid #dc2626" : "1.4px solid #e5e9f2",
      borderRadius: 8,
      padding: multi ? "7px 11px" : "7px 11px",
      fontSize: 15,
      width: "100%",
      outline: "none"
    };
  }

  return (
    <form onSubmit={handleSubmit} style={{ background: "#fff", borderRadius: 13, padding: 26, maxWidth: 575, margin: "0 auto" }}>
      <div style={{ fontWeight: 700, fontSize: 21, marginBottom: 18 }}>
        {entryType === "serialized" && "Serialized Data Entry"}
        {entryType === "nonserial" && "Non-Serialized Data Entry"}
        {entryType === "sand" && "Sand Transfer Entry"}
      </div>

      {(entryType === "serialized" || entryType === "nonserial") && (
        <>
          {fieldBlock("Oven", "oven", <input value={fields.oven} readOnly style={inputStyle("oven")} />)}
          {fieldBlock("Pallet Card #", "palletCard", <input value={fields.palletCard} onChange={e => setField("palletCard", e.target.value)} style={inputStyle("palletCard")} onBlur={() => setTouched(t => ({ ...t, palletCard: true }))} />)}
          {fieldBlock("Bake Hours", "bakeHours", <input type="number" value={fields.bakeHours} onChange={e => setField("bakeHours", e.target.value)} style={inputStyle("bakeHours")} onBlur={() => setTouched(t => ({ ...t, bakeHours: true }))} />)}
          {fieldBlock("Process Order #", "processOrder", <input value={fields.processOrder} onChange={e => setField("processOrder", e.target.value)} style={inputStyle("processOrder")} onBlur={() => setTouched(t => ({ ...t, processOrder: true }))} />)}
        </>
      )}

      {entryType === "serialized" && (
        <>
          {fieldBlock("Unbaked Item #", "unbakedItem", <input value={fields.unbakedItem} onChange={e => setField("unbakedItem", e.target.value)} style={inputStyle("unbakedItem")} onBlur={() => setTouched(t => ({ ...t, unbakedItem: true }))} />)}
          {fieldBlock("Baked Item #", "bakedItem", <input value={fields.bakedItem} onChange={e => setField("bakedItem", e.target.value)} style={inputStyle("bakedItem")} onBlur={() => setTouched(t => ({ ...t, bakedItem: true }))} />)}
          <div style={{ marginBottom: 15 }}>
            <label style={{ color: "#374151", fontWeight: 600, fontSize: 13 }}>Batch Numbers (up to 4)</label>
            <div style={{ display: "flex", gap: 6 }}>
              {fields.batches.map((v: string, i: number) =>
                <input
                  key={i}
                  value={v}
                  placeholder={"Batch " + (i + 1)}
                  onChange={e => setField("batches", fields.batches.map((b: string, j: number) => i === j ? e.target.value : b))}
                  style={inputStyle("batches")}
                  onBlur={() => setTouched(t => ({ ...t, batches: true }))}
                />)}
            </div>
            {errors.batches && touched.batches && (
              <div style={{ color: "#dc2626", fontWeight: 500, fontSize: 12, marginTop: 2 }}>{errors.batches}</div>
            )}
          </div>
          {fieldBlock("Baked Qty", "bakedQty", <input type="number" value={fields.bakedQty} onChange={e => setField("bakedQty", e.target.value)} style={inputStyle("bakedQty")} onBlur={() => setTouched(t => ({ ...t, bakedQty: true }))} />)}
          <div style={{ display: "flex", gap: 10, marginBottom: 15 }}>
            <div style={{ flex: 1 }}>
              {fieldBlock("Hamper #1", "hamper1", <select value={fields.hampers[0]} onChange={e => setField("hampers", [e.target.value, fields.hampers[1]])} style={inputStyle("hampers")}>
                <option value="">Select Hamper</option>{hamperOptions.map(h => <option key={h}>{h}</option>)}
              </select>)}
            </div>
            <div style={{ flex: 1 }}>
              {fieldBlock("Hamper #2", "hamper2", <select value={fields.hampers[1]} onChange={e => setField("hampers", [fields.hampers[0], e.target.value])} style={inputStyle("hampers")}>
                <option value="">Select Hamper</option>{hamperOptions.map(h => <option key={h}>{h}</option>)}
              </select>)}
            </div>
          </div>
          {fieldBlock("Strapping", "strapping", <select value={fields.strapping} onChange={e => setField("strapping", e.target.value)} style={inputStyle("strapping")} onBlur={() => setTouched(t => ({ ...t, strapping: true }))}>
            <option value="">Select Strapping</option>{strappingOptions.map(s => <option key={s}>{s}</option>)}
          </select>)}
          {fieldBlock("Operator", "operator", <input value={fields.operator} readOnly style={inputStyle("operator")} />)}
          {fieldBlock("Comments", "comments", <textarea value={fields.comments} onChange={e => setField("comments", e.target.value)} style={inputStyle("comments", true)} />)}
        </>
      )}
      {entryType === "nonserial" && (
        <>
          {fieldBlock("Batch Number", "batch", <input value={fields.batch} onChange={e => setField("batch", e.target.value)} style={inputStyle("batch")} onBlur={() => setTouched(t => ({ ...t, batch: true }))} />)}
          {fieldBlock("Unbaked Qty", "unbakedQty", <input type="number" value={fields.unbakedQty} onChange={e => setField("unbakedQty", e.target.value)} style={inputStyle("unbakedQty")} onBlur={() => setTouched(t => ({ ...t, unbakedQty: true }))} />)}
          {fieldBlock("Baked Qty", "bakedQty", <input type="number" value={fields.bakedQty} onChange={e => setField("bakedQty", e.target.value)} style={inputStyle("bakedQty")} onBlur={() => setTouched(t => ({ ...t, bakedQty: true }))} />)}
          {fieldBlock("Hamper", "hamper", <select value={fields.hamper} onChange={e => setField("hamper", e.target.value)} style={inputStyle("hamper")} onBlur={() => setTouched(t => ({ ...t, hamper: true }))}>
            <option value="">Select</option>{hamperOptions.map(h => <option key={h}>{h}</option>)}
          </select>)}
          {fieldBlock("Strapping", "strapping", <select value={fields.strapping} onChange={e => setField("strapping", e.target.value)} style={inputStyle("strapping")} onBlur={() => setTouched(t => ({ ...t, strapping: true }))}>
            <option value="">Select</option>{strappingOptions.map(s => <option key={s}>{s}</option>)}
          </select>)}
          {fieldBlock("Scrap Code", "scrapCode", <select value={fields.scrapCode} onChange={e => setField("scrapCode", e.target.value)} style={inputStyle("scrapCode")}>
            <option value="">If Scrap</option>{defectCodes.map(c => <option key={c}>{c}</option>)}
          </select>)}
          {fieldBlock("Operator", "operator", <input value={fields.operator} readOnly style={inputStyle("operator")} />)}
          {fieldBlock("Comments", "comments", <textarea value={fields.comments} onChange={e => setField("comments", e.target.value)} style={inputStyle("comments", true)} />)}
        </>
      )}
      {entryType === "sand" && (
        <>
          {fieldBlock("Sand Item #", "sandItem", <input value={fields.sandItem} onChange={e => setField("sandItem", e.target.value)} style={inputStyle("sandItem")} onBlur={() => setTouched(t => ({ ...t, sandItem: true }))} />)}
          {fieldBlock("Batch", "sandBatch", <input value={fields.sandBatch} onChange={e => setField("sandBatch", e.target.value)} style={inputStyle("sandBatch")} onBlur={() => setTouched(t => ({ ...t, sandBatch: true }))} />)}
          {fieldBlock("Source Location", "sourceLoc", <select value={fields.sourceLoc} onChange={e => setField("sourceLoc", e.target.value)} style={inputStyle("sourceLoc")} onBlur={() => setTouched(t => ({ ...t, sourceLoc: true }))}>
            <option value="">Select</option>{sandLocs.map(l => <option key={l}>{l}</option>)}
          </select>)}
          {fieldBlock("Destination Furnace", "destFurnace", <select value={fields.destFurnace} onChange={e => setField("destFurnace", e.target.value)} style={inputStyle("destFurnace")} onBlur={() => setTouched(t => ({ ...t, destFurnace: true }))}>
            <option value="">Select</option>{sandFurnace.map(f => <option key={f}>{f}</option>)}
          </select>)}
          {fieldBlock("Transfer Quantity (kg)", "qty", <input type="number" value={fields.qty} onChange={e => setField("qty", e.target.value)} style={inputStyle("qty")} onBlur={() => setTouched(t => ({ ...t, qty: true }))} />)}
          {fieldBlock("Operator", "operator", <input value={fields.operator} readOnly style={inputStyle("operator")} />)}
          {fieldBlock("Comments", "comments", <textarea value={fields.comments} onChange={e => setField("comments", e.target.value)} style={inputStyle("comments", true)} />)}
        </>
      )}
      <div style={{ display: "flex", gap: 16, marginTop: 20 }}>
        <button style={{ ...submitBtn, opacity: saving ? 0.7 : 1 }} type="submit" disabled={saving}>
          {saving ? "Saving…" : "Save Entry"}
        </button>
        <button type="button" style={{ ...submitBtn, background: "#64748b" }} onClick={onCancel}>Cancel</button>
      </div>
    </form>
  );
}

function LampRodRecords({
  records,
  onNew,
  user,
  furnaceChannels
}: {
  records: LampRodEntryType[];
  onNew: (type: "serialized" | "nonserial" | "sand", ovenId: string) => void;
  user: UserType;
  furnaceChannels: FurnaceChannel[];
}) {
  // Provide list of ovens or let sand be entered ad hoc
  const ovenChannels = furnaceChannels.filter(fc => true); // could filter if wanted
  // For each oven, only allow entry if available (not in use)
  const canAdd = (ovenId: string) =>
    ["Operator", "Manager"].includes(user.role) &&
    furnaceChannels.find(f => f.id === ovenId)?.status === "Empty";

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 800, fontSize: 22 }}>
          Lamp Rod — All Records
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {ovenChannels.map(oc => (
            <React.Fragment key={oc.id}>
              <button
                style={{ ...navBtn, opacity: canAdd(oc.id) ? 1 : 0.6, pointerEvents: canAdd(oc.id) ? "auto" : "none" }}
                onClick={() => onNew("serialized", oc.id)}
                title={!canAdd(oc.id) ? "Oven busy" : undefined}
              >+ Serialized ({oc.id})</button>
              <button
                style={{ ...navBtn, background: "#10b981", opacity: canAdd(oc.id) ? 1 : 0.6, pointerEvents: canAdd(oc.id) ? "auto" : "none" }}
                onClick={() => onNew("nonserial", oc.id)}
                title={!canAdd(oc.id) ? "Oven busy" : undefined}
              >+ Non-Serial ({oc.id})</button>
            </React.Fragment>
          ))}
          <button style={{ ...navBtn, background: "#eab308" }} onClick={() => onNew("sand", "")}>+ Sand Transfer</button>
        </div>
      </div>
      {records.length === 0 && (
        <div style={{ padding: "13px", background: "#f3f8fe", color: "#64748b", fontWeight: 500, borderRadius: 10 }}>
          No records yet.
        </div>
      )}
      {records.length > 0 && (
        <div style={{
          overflowX: "auto", background: "#fff", borderRadius: 14, boxShadow: "0 2px 9px #2233610c",
          border: "1px solid #e2e8f0", marginTop: 6
        }}>
          <table style={{ width: "100%", fontSize: 15, borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ color: "#8492ac", textAlign: 'left', fontSize: 13 }}>
                <th>Type</th>
                <th>Oven</th>
                <th>Pallet/Batch/Sand</th>
                <th>Baked Qty / Qty</th>
                <th>Operator</th>
                <th>Created At</th>
                <th>By</th>
              </tr>
            </thead>
            <tbody>
              {records.map(row => (
                <tr key={row.id} style={{ borderTop: "1px solid #ecf1f7", background: row.view === "sand" ? "#fefce8" : undefined }}>
                  <td>
                    <span style={badge(
                      row.view === "serialized" ? "#2563eb" :
                        row.view === "nonserial" ? "#10b981" : "#eab308"
                    )}>{row.view}</span>
                  </td>
                  <td>{row.furnaceId}</td>
                  <td>
                    {row.view === "sand"
                      ? row.sandItem
                      : (row.batch || row.palletCard || "-")}
                  </td>
                  <td>{row.bakedQty || row.qty || "-"}</td>
                  <td>{row.operator}</td>
                  <td>{row.createdAt && (new Date(row.createdAt).toLocaleString())}</td>
                  <td>{row.createdBy}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}