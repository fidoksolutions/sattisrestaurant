import { useState, useEffect, useMemo } from "react";

import * as XLSX from "xlsx";

 

// ─────────────────────────────────────────────────────────────────────────────

// THEME

// ─────────────────────────────────────────────────────────────────────────────

const C = {

  bg:        "#080B10",

  surface:   "#0D1117",

  card:      "#111720",

  card2:     "#161E2C",

  border:    "#1E2D40",

  border2:   "#253550",

  flame:     "#FF4D1C",

  flameLt:   "#FF6B3D",

  gold:      "#F5A623",

  goldLt:    "#FFD166",

  green:     "#00C896",

  greenDk:   "#00A87E",

  red:       "#FF3B5C",

  blue:      "#2B7FFF",

  blueLt:    "#5CA3FF",

  purple:    "#9B5DFF",

  text:      "#EEF2FF",

  sub:       "#8899B4",

  muted:     "#4A5A72",

  white:     "#FFFFFF",

};

 

const FONT_HEAD = "'Clash Display', 'Syne', sans-serif";

const FONT_BODY = "'Cabinet Grotesk', 'DM Sans', sans-serif";

 

const GS = `

@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');

*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

html{scroll-behavior:smooth}

body{background:${C.bg};color:${C.text};font-family:${FONT_BODY};min-height:100vh;overflow-x:hidden}

input,select,textarea{outline:none;font-family:inherit}

button{cursor:pointer;border:none;outline:none;font-family:inherit}

::-webkit-scrollbar{width:4px}

::-webkit-scrollbar-track{background:${C.surface}}

::-webkit-scrollbar-thumb{background:${C.border2};border-radius:2px}

@keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}

@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}

@keyframes spin{to{transform:rotate(360deg)}}

@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}

.fade-up{animation:fadeUp .4s ease both}

.fade-up-1{animation:fadeUp .4s .05s ease both;opacity:0}

.fade-up-2{animation:fadeUp .4s .1s ease both;opacity:0}

.fade-up-3{animation:fadeUp .4s .15s ease both;opacity:0}

.fade-up-4{animation:fadeUp .4s .2s ease both;opacity:0}

`;

 

// ─────────────────────────────────────────────────────────────────────────────

// CONSTANTS & HELPERS

// ─────────────────────────────────────────────────────────────────────────────

const ROLES = { STAFF: "staff", ADMIN: "admin" };

const STATUS = { PENDING: "pending", APPROVED: "approved", REJECTED: "rejected" };

const STAFF_USERS = [

  { id: "samiya",  name: "Samiya",  pin: "1111", role: ROLES.STAFF },

  { id: "staff2",  name: "Staff",   pin: "2222", role: ROLES.STAFF },

];

const ADMIN_USERS = [

  { id: "admin",   name: "Admin",   pin: "0000", role: ROLES.ADMIN },

];

const ALL_USERS = [...STAFF_USERS, ...ADMIN_USERS];

 

const todayStr = () => new Date().toISOString().slice(0, 10);

const fmt = (v) => `£${(+v || 0).toFixed(2)}`;

const fmtShort = (v) => {

  const n = +v || 0;

  return n >= 1000 ? `£${(n / 1000).toFixed(1)}k` : fmt(n);

};

const dayName = (d) => d ? new Date(d + "T12:00:00").toLocaleDateString("en-GB", { weekday: "long" }) : "";

const shortDay = (d) => d ? new Date(d + "T12:00:00").toLocaleDateString("en-GB", { weekday: "short" }) : "";

const n = (v) => parseFloat(v) || 0;

 

const DENOMS = [

  { key: "d_50n", label: "£50 Notes", val: 50,   type: "note" },

  { key: "d_20n", label: "£20 Notes", val: 20,   type: "note" },

  { key: "d_10n", label: "£10 Notes", val: 10,   type: "note" },

  { key: "d_5n",  label: "£5 Notes",  val: 5,    type: "note" },

  { key: "d_2c",  label: "£2 Coins",  val: 2,    type: "coin" },

  { key: "d_1c",  label: "£1 Coins",  val: 1,    type: "coin" },

  { key: "d_50p", label: "50p",       val: 0.50, type: "coin" },

  { key: "d_20p", label: "20p",       val: 0.20, type: "coin" },

  { key: "d_10p", label: "10p",       val: 0.10, type: "coin" },

  { key: "d_5p",  label: "5p",        val: 0.05, type: "coin" },

  { key: "d_2p",  label: "2p",        val: 0.02, type: "coin" },

  { key: "d_1p",  label: "1p",        val: 0.01, type: "coin" },

];

 

const EMPTY_FORM = (openingFloat = 25) => ({

  date: todayStr(),

  opening_float: openingFloat,

  till_cash: "", till_sumup: "",

  je_paid: "", je_notpaid: "",

  deliveroo: "", uber: "", foodhub: "",

  exp_driver: "", exp_tips: "",

  exp_other1: "", exp_other1_lbl: "Other Expense 1",

  exp_other2: "", exp_other2_lbl: "Other Expense 2",

  exp_other3: "", exp_other3_lbl: "Other Expense 3",

  ...Object.fromEntries(DENOMS.map(d => [d.key, ""])),

  staff_name: "", manager: "", notes: "",

});

 

function calcTotals(f) {

  const tillCash   = n(f.till_cash);

  const tillSumup  = n(f.till_sumup);

  const tillTotal  = tillCash + tillSumup;

  const jePaid     = n(f.je_paid);

  const jeNotPaid  = n(f.je_notpaid);

  const jeTotal    = jePaid + jeNotPaid;

  const deliveroo  = n(f.deliveroo);

  const uber       = n(f.uber);

  const foodhub    = n(f.foodhub);

  const onlineTotal = deliveroo + uber + foodhub;

  const grandTotal  = tillTotal + jeTotal + onlineTotal;

  const totalExp    = n(f.exp_driver) + n(f.exp_tips) + n(f.exp_other1) + n(f.exp_other2) + n(f.exp_other3);

  const openFloat   = n(f.opening_float);

  const cashIn      = openFloat + tillCash + jeNotPaid;

  const expectedCash = cashIn - totalExp;

  const notesCash   = DENOMS.filter(d => d.type === "note").reduce((s, d) => s + n(f[d.key]) * d.val, 0);

  const coinsCash   = DENOMS.filter(d => d.type === "coin").reduce((s, d) => s + n(f[d.key]) * d.val, 0);

  const physCash    = notesCash + coinsCash;

  const variance    = physCash - expectedCash;

  return { tillCash, tillSumup, tillTotal, jePaid, jeNotPaid, jeTotal, deliveroo, uber, foodhub, onlineTotal, grandTotal, totalExp, openFloat, cashIn, expectedCash, notesCash, coinsCash, physCash, variance };

}

 

// ─────────────────────────────────────────────────────────────────────────────

// EXCEL GENERATOR — well formatted

// ─────────────────────────────────────────────────────────────────────────────

function generateExcel(submissions, filename) {

  const wb = XLSX.utils.book_new();

 

  submissions.forEach((sub) => {

    const f = sub.form;

    const t = calcTotals(f);

    const day = dayName(f.date);

 

    const data = [

      ["SATTIS PERI PERI – DAILY BALANCE SHEET", "", "", "", ""],

      ["", "", "", "", ""],

      ["Date:", f.date, "", "Day:", day],

      ["Staff:", f.staff_name, "", "Manager:", f.manager || "—"],

      ["Status:", (sub.status || "pending").toUpperCase(), "", "Submitted:", sub.timestamp || ""],

      ["", "", "", "", ""],

      ["━━━ SECTION 1: TILL SALES ━━━", "", "", "", ""],

      ["", "Cash (£)", "SUMUP (£)", "", "Total (£)"],

      ["Till Sales", t.tillCash, t.tillSumup, "", t.tillTotal],

      ["TILL TOTAL", "", "", "", t.tillTotal],

      ["", "", "", "", ""],

      ["━━━ SECTION 2: JUST EAT ━━━", "", "", "", ""],

      ["", "Amount (£)", "Payment", "", "Total (£)"],

      ["JustEat Paid (On App)", t.jePaid, "On App", "", t.jePaid],

      ["JustEat Not Paid (Cash)", t.jeNotPaid, "Cash", "", t.jeNotPaid],

      ["JUSTEAT TOTAL", "", "", "", t.jeTotal],

      ["", "", "", "", ""],

      ["━━━ SECTION 3: ONLINE PLATFORMS ━━━", "", "", "", ""],

      ["", "Amount (£)", "Payment", "", "Total (£)"],

      ["Deliveroo", t.deliveroo, "On App", "", t.deliveroo],

      ["Uber Eats", t.uber, "On App", "", t.uber],

      ["Foodhub", t.foodhub, "On App", "", t.foodhub],

      ["PLATFORMS TOTAL", "", "", "", t.onlineTotal],

      ["", "", "", "", ""],

      ["━━━ SECTION 4: GRAND TOTAL ━━━", "", "", "", ""],

      ["Sale Type", "Cash (£)", "On App (£)", "SUMUP (£)", "Total (£)"],

      ["Till Sales", t.tillCash, 0, t.tillSumup, t.tillTotal],

      ["JustEat (App)", 0, t.jePaid, 0, t.jePaid],

      ["JustEat (Cash)", t.jeNotPaid, 0, 0, t.jeNotPaid],

      ["Deliveroo", 0, t.deliveroo, 0, t.deliveroo],

      ["Uber Eats", 0, t.uber, 0, t.uber],

      ["Foodhub", 0, t.foodhub, 0, t.foodhub],

      ["GRAND TOTAL", t.tillCash + t.jeNotPaid, t.jePaid + t.deliveroo + t.uber + t.foodhub, t.tillSumup, t.grandTotal],

      ["", "", "", "", ""],

      ["━━━ SECTION 5: EXPENSES ━━━", "", "", "", ""],

      ["Expense", "Amount (£)", "", "", ""],

      ["Delivery Driver", n(f.exp_driver), "", "", ""],

      ["Driver Tips", n(f.exp_tips), "", "", ""],

      [f.exp_other1_lbl || "Other 1", n(f.exp_other1), "", "", ""],

      [f.exp_other2_lbl || "Other 2", n(f.exp_other2), "", "", ""],

      [f.exp_other3_lbl || "Other 3", n(f.exp_other3), "", "", ""],

      ["TOTAL EXPENSES", t.totalExp, "", "", ""],

      ["", "", "", "", ""],

      ["━━━ SECTION 6: CASH RECONCILIATION ━━━", "", "", "", ""],

      ["Opening Float", t.openFloat, "", "", ""],

      ["Cash Sales – Till", t.tillCash, "", "", ""],

      ["Cash Sales – JustEat", t.jeNotPaid, "", "", ""],

      ["TOTAL CASH IN", t.cashIn, "", "", ""],

      ["Less: Expenses", t.totalExp, "", "", ""],

      ["EXPECTED CASH IN HAND", t.expectedCash, "", "", ""],

      ["", "", "", "", ""],

      ["━━━ SECTION 7: PHYSICAL CASH COUNT ━━━", "", "", "", ""],

      ["Denomination", "Qty", "Value Each (£)", "", "Total (£)"],

      ...DENOMS.map(d => [d.label, n(f[d.key]), d.val, "", n(f[d.key]) * d.val]),

      ["Total Notes", "", "", "", t.notesCash],

      ["Total Coins", "", "", "", t.coinsCash],

      ["TOTAL PHYSICAL CASH", "", "", "", t.physCash],

      ["", "", "", "", ""],

      ["━━━ SECTION 8: VARIANCE ━━━", "", "", "", ""],

      ["Expected Cash", t.expectedCash, "", "", ""],

      ["Physical Cash", t.physCash, "", "", ""],

      ["VARIANCE", t.variance, "", "", t.variance === 0 ? "✅ BALANCED" : t.variance > 0 ? "⚠️ OVER" : "⚠️ SHORT"],

      ["", "", "", "", ""],

      ["━━━ NOTES ━━━", "", "", "", ""],

      [f.notes || "—", "", "", "", ""],

    ];

 

    const ws = XLSX.utils.aoa_to_sheet(data);

    ws["!cols"] = [{ wch: 30 }, { wch: 16 }, { wch: 16 }, { wch: 14 }, { wch: 16 }];

 

    // Style key rows

    const styleCell = (addr, bold, bg, color) => {

      if (!ws[addr]) return;

      ws[addr].s = { font: { bold, color: { rgb: (color || "EEF2FF").replace("#","") } }, fill: bg ? { fgColor: { rgb: bg.replace("#","") } } : undefined, alignment: { horizontal: "left" } };

    };

    ["A1"].forEach(a => styleCell(a, true, "#FF4D1C", "#FFFFFF"));

 

    const sheetName = (shortDay(f.date) + " " + (f.date || "").slice(5)).slice(0, 31);

    XLSX.utils.book_append_sheet(wb, ws, sheetName);

  });

 

  // Weekly Summary sheet

  if (submissions.length > 0) {

    const hdrs = ["Date", "Day", "Status", "Staff", "Till Cash", "Till SUMUP", "JustEat App", "JustEat Cash", "Deliveroo", "Uber Eats", "Foodhub", "Expenses", "Grand Total", "Expected Cash", "Physical Cash", "Variance"];

    const rows = submissions.map(sub => {

      const t = calcTotals(sub.form);

      return [sub.form.date, dayName(sub.form.date), (sub.status||"pending").toUpperCase(), sub.form.staff_name, t.tillCash, t.tillSumup, t.jePaid, t.jeNotPaid, t.deliveroo, t.uber, t.foodhub, t.totalExp, t.grandTotal, t.expectedCash, t.physCash, t.variance];

    });

    const totals = ["TOTAL", "", "", "", ...Array(12).fill(0).map((_, i) => rows.reduce((s, r) => s + (typeof r[i + 4] === "number" ? r[i + 4] : 0), 0))];

    const wsS = XLSX.utils.aoa_to_sheet([hdrs, ...rows, [], totals]);

    wsS["!cols"] = hdrs.map((_, i) => ({ wch: i < 4 ? 16 : 14 }));

    XLSX.utils.book_append_sheet(wb, wsS, "Weekly Summary");

  }

 

  XLSX.writeFile(wb, filename || "Sattis_Report.xlsx");

}

 

// ─────────────────────────────────────────────────────────────────────────────

// STORAGE

// ─────────────────────────────────────────────────────────────────────────────

const LS_KEY = "sattis_submissions";

const LS_FLOAT = "sattis_opening_float";

 

function loadSubmissions() {

  try { return JSON.parse(localStorage.getItem(LS_KEY) || "[]"); } catch { return []; }

}

function saveSubmissions(arr) {

  try { localStorage.setItem(LS_KEY, JSON.stringify(arr)); } catch {}

}

function loadFloat() {

  try { return parseFloat(localStorage.getItem(LS_FLOAT)) || 25; } catch { return 25; }

}

function saveFloat(v) {

  try { localStorage.setItem(LS_FLOAT, String(v)); } catch {}

}

 

// ─────────────────────────────────────────────────────────────────────────────

// UI PRIMITIVES

// ─────────────────────────────────────────────────────────────────────────────

const Lbl = ({ children, req }) => (

  <label style={{ display: "block", fontSize: 11, fontWeight: 600, color: C.muted, marginBottom: 5, letterSpacing: "0.08em", textTransform: "uppercase" }}>

    {children}{req && <span style={{ color: C.flame, marginLeft: 3 }}>*</span>}

  </label>

);

 

const Inp = ({ value, onChange, type = "text", placeholder, min, step, readOnly, accent }) => (

  <input type={type} value={value} onChange={onChange} placeholder={placeholder} min={min} step={step} readOnly={readOnly}

    style={{ width: "100%", padding: "10px 13px", background: readOnly ? C.surface : C.bg, border: `1.5px solid ${accent ? C.flameLt : readOnly ? C.border : C.border2}`, borderRadius: 9, color: readOnly ? C.muted : C.text, fontSize: 14, transition: "border-color .15s" }}

    onFocus={e => { if (!readOnly) e.target.style.borderColor = C.flame; }}

    onBlur={e => { e.target.style.borderColor = accent ? C.flameLt : readOnly ? C.border : C.border2; }}

  />

);

 

const Btn = ({ children, onClick, variant = "primary", size = "md", disabled, full }) => {

  const styles = {

    primary:  { bg: C.flame,   color: "#fff",   hover: C.flameLt },

    success:  { bg: C.green,   color: "#fff",   hover: C.greenDk },

    ghost:    { bg: "transparent", color: C.sub, hover: C.card2, border: `1px solid ${C.border}` },

    danger:   { bg: C.red,     color: "#fff",   hover: "#FF6070" },

    gold:     { bg: C.gold,    color: "#080B10", hover: C.goldLt },

  };

  const s = styles[variant] || styles.primary;

  const pad = size === "sm" ? "7px 16px" : size === "lg" ? "14px 32px" : "10px 22px";

  const fs  = size === "sm" ? 12 : size === "lg" ? 15 : 13;

  return (

    <button onClick={onClick} disabled={disabled}

      style={{ padding: pad, fontSize: fs, fontWeight: 700, borderRadius: 10, background: disabled ? C.muted : s.bg, color: disabled ? C.bg : s.color, border: s.border || "none", opacity: disabled ? 0.5 : 1, transition: "all .15s", width: full ? "100%" : undefined, letterSpacing: "0.03em" }}

      onMouseEnter={e => { if (!disabled) e.target.style.background = s.hover; }}

      onMouseLeave={e => { if (!disabled) e.target.style.background = disabled ? C.muted : s.bg; }}

    >{children}</button>

  );

};

 

const Card = ({ children, style }) => (

  <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, overflow: "hidden", ...style }}>{children}</div>

);

 

const CardHead = ({ icon, title, accent = C.flame, right }) => (

  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 18px", borderBottom: `1px solid ${C.border}`, background: C.surface }}>

    <div style={{ display: "flex", alignItems: "center", gap: 9 }}>

      <span style={{ fontSize: 16 }}>{icon}</span>

      <span style={{ fontFamily: FONT_HEAD, fontWeight: 700, fontSize: 13, color: accent, letterSpacing: "0.05em", textTransform: "uppercase" }}>{title}</span>

    </div>

    {right && <div>{right}</div>}

  </div>

);

 

const Row2 = ({ children }) => (

  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>{children}</div>

);

const Row3 = ({ children }) => (

  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>{children}</div>

);

 

const StatCard = ({ label, value, sub, color = C.flame, icon, delay = 0 }) => (

  <div className={`fade-up-${delay + 1}`} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 13, padding: "16px 18px", position: "relative", overflow: "hidden" }}>

    <div style={{ position: "absolute", top: 0, left: 0, width: 3, height: "100%", background: color, borderRadius: "2px 0 0 2px" }} />

    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>

      <div>

        <div style={{ fontSize: 11, color: C.muted, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>{label}</div>

        <div style={{ fontFamily: FONT_HEAD, fontSize: 22, fontWeight: 700, color, lineHeight: 1 }}>{value}</div>

        {sub && <div style={{ fontSize: 11, color: C.sub, marginTop: 5 }}>{sub}</div>}

      </div>

      {icon && <div style={{ fontSize: 24, opacity: 0.4 }}>{icon}</div>}

    </div>

  </div>

);

 

const Badge = ({ label, color, bg }) => (

  <span style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "3px 10px", borderRadius: 99, background: bg || `${color}22`, color, fontSize: 11, fontWeight: 700, letterSpacing: "0.04em", textTransform: "uppercase" }}>

    {label}

  </span>

);

 

const statusBadge = (status) => {

  if (status === STATUS.APPROVED) return <Badge label="Approved" color={C.green} />;

  if (status === STATUS.REJECTED) return <Badge label="Rejected" color={C.red} />;

  return <Badge label="Pending" color={C.gold} />;

};

 

// ─────────────────────────────────────────────────────────────────────────────

// LOGIN

// ─────────────────────────────────────────────────────────────────────────────

function Login({ onLogin }) {

  const [pin, setPin] = useState("");

  const [err, setErr] = useState("");

  const [shake, setShake] = useState(false);

 

  const handlePin = (digit) => {

    if (pin.length >= 4) return;

    const np = pin + digit;

    setPin(np);

    if (np.length === 4) {

      const user = ALL_USERS.find(u => u.pin === np);

      if (user) { onLogin(user); }

      else {

        setShake(true);

        setErr("Incorrect PIN. Try again.");

        setTimeout(() => { setPin(""); setShake(false); setErr(""); }, 900);

      }

    }

  };

 

  const pad = ["1","2","3","4","5","6","7","8","9","","0","⌫"];

 

  return (

    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", background: C.bg, padding: 24 }}>

      {/* Logo */}

      <div style={{ textAlign: "center", marginBottom: 36 }}>

        <div style={{ width: 64, height: 64, borderRadius: 18, background: `linear-gradient(135deg, ${C.flame}, ${C.gold})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 30, margin: "0 auto 16px" }}>🌶️</div>

        <div style={{ fontFamily: FONT_HEAD, fontSize: 22, fontWeight: 800, color: C.flame, letterSpacing: "0.04em" }}>SATTIS PERI PERI</div>

        <div style={{ fontSize: 12, color: C.muted, letterSpacing: "0.12em", textTransform: "uppercase", marginTop: 4 }}>Balance Sheet Portal</div>

      </div>

 

      {/* PIN Display */}

      <div style={{ animation: shake ? "shimmer .3s ease" : undefined, marginBottom: 24 }}>

        <div style={{ display: "flex", gap: 12, justifyContent: "center" }}>

          {[0,1,2,3].map(i => (

            <div key={i} style={{ width: 16, height: 16, borderRadius: "50%", background: pin.length > i ? C.flame : C.border, transition: "background .1s", boxShadow: pin.length > i ? `0 0 8px ${C.flame}88` : "none" }} />

          ))}

        </div>

        {err && <div style={{ color: C.red, fontSize: 12, textAlign: "center", marginTop: 10 }}>{err}</div>}

      </div>

 

      {/* Numpad */}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 72px)", gap: 10 }}>

        {pad.map((d, i) => (

          <button key={i} onClick={() => d === "⌫" ? setPin(p => p.slice(0, -1)) : d ? handlePin(d) : null}

            style={{ height: 64, borderRadius: 14, background: d === "" ? "transparent" : C.card, border: `1px solid ${d === "" ? "transparent" : C.border}`, color: d === "⌫" ? C.red : C.text, fontSize: d === "⌫" ? 18 : 22, fontFamily: FONT_HEAD, fontWeight: 700, cursor: d === "" ? "default" : "pointer", transition: "all .1s" }}

            onMouseEnter={e => { if (d && d !== "") { e.target.style.background = C.card2; e.target.style.borderColor = C.border2; }}}

            onMouseLeave={e => { if (d && d !== "") { e.target.style.background = C.card; e.target.style.borderColor = C.border; }}}

          >{d}</button>

        ))}

      </div>

 

      <div style={{ marginTop: 28, fontSize: 11, color: C.muted, textAlign: "center", lineHeight: 1.8 }}>

        Staff PIN: 1111 (Samiya) · 2222 (Other Staff)<br/>Admin PIN: 0000

      </div>

    </div>

  );

}

 

// ─────────────────────────────────────────────────────────────────────────────

// STAFF FORM

// ─────────────────────────────────────────────────────────────────────────────

function StaffForm({ user, openingFloat, submissions, onSubmit }) {

  const [form, setForm] = useState(() => EMPTY_FORM(openingFloat));

  const [step, setStep] = useState(0);

  const [done, setDone] = useState(false);

 

  // Pre-fill staff name

  useEffect(() => { setForm(f => ({ ...f, staff_name: user.name, opening_float: openingFloat })); }, [user, openingFloat]);

 

  const set = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  const t = useMemo(() => calcTotals(form), [form]);

 

  const STEPS = [

    { label: "Sales",     icon: "🏪" },

    { label: "Expenses",  icon: "💸" },

    { label: "Cash",      icon: "💰" },

    { label: "Sign Off",  icon: "✍️" },

  ];

 

  const handleSubmit = () => {

    const entry = {

      id: Date.now().toString(),

      form: { ...form },

      totals: calcTotals(form),

      timestamp: new Date().toLocaleString("en-GB"),

      submittedBy: user.id,

      status: STATUS.PENDING,

      rejectionNote: "",

      assignedTo: "",

    };

    onSubmit(entry);

    setDone(true);

  };

 

  if (done) return (

    <div className="fade-up" style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "65vh", gap: 18, textAlign: "center", padding: 24 }}>

      <div style={{ width: 80, height: 80, borderRadius: "50%", background: `linear-gradient(135deg, ${C.green}22, ${C.green}44)`, border: `2px solid ${C.green}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 36 }}>✅</div>

      <div style={{ fontFamily: FONT_HEAD, fontSize: 26, fontWeight: 800, color: C.green }}>Report Submitted!</div>

      <div style={{ color: C.sub, fontSize: 14 }}>{dayName(form.date)}, {form.date}</div>

      <div className="fade-up-1" style={{ background: C.card, borderRadius: 14, border: `1px solid ${C.border}`, padding: "20px 28px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px 28px", textAlign: "left", minWidth: 280 }}>

        {[["Grand Total", fmt(t.grandTotal), C.flame], ["Expected Cash", fmt(t.expectedCash), C.gold], ["Physical Cash", fmt(t.physCash), C.blue], ["Variance", fmt(t.variance), t.variance === 0 ? C.green : C.red]].map(([l, v, c]) => (

          <div key={l}><div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 3 }}>{l}</div><div style={{ fontFamily: FONT_HEAD, fontSize: 18, fontWeight: 700, color: c }}>{v}</div></div>

        ))}

      </div>

      <div className="fade-up-2" style={{ padding: "10px 20px", borderRadius: 10, background: t.variance === 0 ? `${C.green}18` : `${C.red}18`, border: `1px solid ${t.variance === 0 ? C.green : C.red}`, color: t.variance === 0 ? C.green : C.red, fontWeight: 700 }}>

        {t.variance === 0 ? "✅ Cash Balanced" : t.variance > 0 ? `⚠️ Over by ${fmt(Math.abs(t.variance))}` : `⚠️ Short by ${fmt(Math.abs(t.variance))}`}

      </div>

      <div style={{ fontSize: 12, color: C.muted }}>Awaiting admin review</div>

      <Btn onClick={() => { setForm(EMPTY_FORM(openingFloat)); setDone(false); setStep(0); }} variant="ghost">Start New Day</Btn>

    </div>

  );

 

  // Check if this staff has a rejected report to re-do

  const rejectedForMe = submissions.find(s => s.status === STATUS.REJECTED && s.assignedTo === user.id);

 

  return (

    <div style={{ maxWidth: 640, margin: "0 auto", padding: "0 0 40px" }}>

      {rejectedForMe && (

        <div className="fade-up" style={{ background: `${C.red}15`, border: `1px solid ${C.red}`, borderRadius: 12, padding: "14px 18px", marginBottom: 20, display: "flex", gap: 12, alignItems: "flex-start" }}>

          <span style={{ fontSize: 20 }}>⚠️</span>

          <div>

            <div style={{ fontWeight: 700, color: C.red, marginBottom: 3 }}>Report Rejected — Please Re-submit</div>

            <div style={{ fontSize: 13, color: C.sub }}>Date: {rejectedForMe.form.date} · Note: {rejectedForMe.rejectionNote || "No reason given"}</div>

          </div>

        </div>

      )}

 

      {/* Opening float info */}

      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center" }}>

        <div style={{ fontSize: 12, color: C.muted }}>Today's Opening Float</div>

        <div style={{ fontFamily: FONT_HEAD, fontSize: 18, fontWeight: 700, color: C.gold }}>{fmt(openingFloat)}</div>

      </div>

 

      {/* Step tabs */}

      <div style={{ display: "flex", gap: 6, marginBottom: 22 }}>

        {STEPS.map((s, i) => (

          <div key={i} onClick={() => setStep(i)} style={{ flex: 1, padding: "10px 6px", borderRadius: 11, background: step === i ? C.flame : C.card, border: `1px solid ${step === i ? C.flame : C.border}`, textAlign: "center", cursor: "pointer", transition: "all .2s" }}>

            <div style={{ fontSize: 15 }}>{s.icon}</div>

            <div style={{ fontSize: 10, fontWeight: step === i ? 700 : 400, color: step === i ? "#fff" : C.muted, marginTop: 2, letterSpacing: "0.04em" }}>{s.label}</div>

          </div>

        ))}

      </div>

 

      {/* Live totals bar */}

      <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "12px 16px", marginBottom: 18, display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>

        {[["Grand Total", fmt(t.grandTotal), C.flame], ["Cash Sales", fmt(t.tillCash + t.jeNotPaid), C.gold], ["Online Sales", fmt(t.jePaid + t.deliveroo + t.uber + t.foodhub), C.blue], ["Expenses", fmt(t.totalExp), C.red]].map(([l, v, c]) => (

          <div key={l} style={{ textAlign: "center" }}>

            <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase" }}>{l}</div>

            <div style={{ fontFamily: FONT_HEAD, fontSize: 15, fontWeight: 700, color: c }}>{v}</div>

          </div>

        ))}

      </div>

 

      {/* Step 0 – Sales */}

      {step === 0 && <div className="fade-up">

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="📅" title="Date" />

          <div style={{ padding: 18 }}>

            <Row2>

              <div><Lbl req>Date</Lbl><Inp type="date" value={form.date} onChange={set("date")} /></div>

              <div><Lbl>Day</Lbl><Inp value={dayName(form.date)} readOnly /></div>

            </Row2>

          </div>

        </Card>

 

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="🖥️" title="1. Till Sales" />

          <div style={{ padding: 18 }}>

            <Row2>

              <div><Lbl>Cash (£)</Lbl><Inp type="number" min="0" step="0.01" value={form.till_cash} onChange={set("till_cash")} placeholder="0.00" /></div>

              <div><Lbl>SUMUP (£)</Lbl><Inp type="number" min="0" step="0.01" value={form.till_sumup} onChange={set("till_sumup")} placeholder="0.00" /></div>

            </Row2>

            <div style={{ marginTop: 12, padding: "9px 13px", background: C.surface, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>

              <span style={{ color: C.sub, fontSize: 13 }}>Till Total</span>

              <span style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: C.flame }}>{fmt(t.tillTotal)}</span>

            </div>

          </div>

        </Card>

 

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="🥡" title="2. Just Eat" />

          <div style={{ padding: 18 }}>

            <Row2>

              <div>

                <Lbl>Paid Orders – On App (£)</Lbl>

                <Inp type="number" min="0" step="0.01" value={form.je_paid} onChange={set("je_paid")} placeholder="0.00" />

                <div style={{ fontSize: 11, color: C.blue, marginTop: 4 }}>💳 On App</div>

              </div>

              <div>

                <Lbl>Not Paid – Cash (£)</Lbl>

                <Inp type="number" min="0" step="0.01" value={form.je_notpaid} onChange={set("je_notpaid")} placeholder="0.00" />

                <div style={{ fontSize: 11, color: C.gold, marginTop: 4 }}>💵 Cash</div>

              </div>

            </Row2>

            <div style={{ marginTop: 12, padding: "9px 13px", background: C.surface, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>

              <span style={{ color: C.sub, fontSize: 13 }}>JustEat Total</span>

              <span style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: C.flame }}>{fmt(t.jeTotal)}</span>

            </div>

          </div>

        </Card>

 

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="📱" title="3. Online Platforms" />

          <div style={{ padding: 18 }}>

            <Row3>

              {[["deliveroo","Deliveroo","🟦"],["uber","Uber Eats","🟩"],["foodhub","Foodhub","🟥"]].map(([k,l,e]) => (

                <div key={k}>

                  <Lbl>{e} {l}</Lbl>

                  <Inp type="number" min="0" step="0.01" value={form[k]} onChange={set(k)} placeholder="0.00" />

                  <div style={{ fontSize: 10, color: C.blue, marginTop: 3 }}>💳 On App</div>

                </div>

              ))}

            </Row3>

            <div style={{ marginTop: 12, padding: "9px 13px", background: C.surface, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>

              <span style={{ color: C.sub, fontSize: 13 }}>Platforms Total</span>

              <span style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: C.flame }}>{fmt(t.onlineTotal)}</span>

            </div>

          </div>

        </Card>

      </div>}

 

      {/* Step 1 – Expenses */}

      {step === 1 && <div className="fade-up">

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="💸" title="5. Daily Expenses" accent={C.red} />

          <div style={{ padding: 18 }}>

            <Row2>

              <div><Lbl>Delivery Driver (£)</Lbl><Inp type="number" min="0" step="0.01" value={form.exp_driver} onChange={set("exp_driver")} placeholder="0.00" /></div>

              <div><Lbl>Driver Tips (£)</Lbl><Inp type="number" min="0" step="0.01" value={form.exp_tips} onChange={set("exp_tips")} placeholder="0.00" /></div>

            </Row2>

            <div style={{ marginTop: 14, borderTop: `1px solid ${C.border}`, paddingTop: 14 }}>

              {[["exp_other1","exp_other1_lbl","Other 1"],["exp_other2","exp_other2_lbl","Other 2"],["exp_other3","exp_other3_lbl","Other 3"]].map(([vk,lk,ph]) => (

                <div key={vk} style={{ display: "grid", gridTemplateColumns: "1fr 110px", gap: 10, marginBottom: 10 }}>

                  <div><Lbl>Description</Lbl><Inp value={form[lk]} onChange={set(lk)} placeholder={ph} /></div>

                  <div><Lbl>Amount (£)</Lbl><Inp type="number" min="0" step="0.01" value={form[vk]} onChange={set(vk)} placeholder="0.00" /></div>

                </div>

              ))}

            </div>

            <div style={{ padding: "9px 13px", background: C.surface, borderRadius: 8, display: "flex", justifyContent: "space-between" }}>

              <span style={{ color: C.sub, fontSize: 13 }}>Total Expenses</span>

              <span style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: C.red }}>{fmt(t.totalExp)}</span>

            </div>

          </div>

        </Card>

 

        <Card>

          <CardHead icon="🧾" title="6. Cash Reconciliation" accent={C.gold} />

          <div style={{ padding: "4px 18px 18px" }}>

            {[

              ["Opening Float", fmt(t.openFloat), C.sub],

              ["Cash Sales – Till", fmt(t.tillCash), C.text],

              ["Cash Sales – JustEat", fmt(t.jeNotPaid), C.text],

              ["Total Cash In", fmt(t.cashIn), C.gold],

              ["Less: Expenses", `(${fmt(t.totalExp)})`, C.red],

              ["Expected Cash in Hand", fmt(t.expectedCash), C.green],

            ].map(([l, v, c]) => (

              <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "10px 0", borderBottom: `1px solid ${C.border}` }}>

                <span style={{ color: C.sub, fontSize: 13 }}>{l}</span>

                <span style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: c, fontSize: 14 }}>{v}</span>

              </div>

            ))}

          </div>

        </Card>

      </div>}

 

      {/* Step 2 – Physical Cash */}

      {step === 2 && <div className="fade-up">

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="💰" title="7. Physical Cash Count" accent={C.green} />

          <div style={{ padding: 18 }}>

            <div style={{ marginBottom: 14, padding: "8px 12px", background: C.surface, borderRadius: 8, fontSize: 12, color: C.muted }}>

              Enter the <b style={{ color: C.text }}>quantity</b> of each note/coin you have

            </div>

            <div style={{ marginBottom: 10, fontSize: 11, color: C.muted, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em" }}>📄 Notes</div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginBottom: 16 }}>

              {DENOMS.filter(d => d.type === "note").map(d => {

                const qty = n(form[d.key]); const tot = qty * d.val;

                return (

                  <div key={d.key} style={{ background: C.surface, borderRadius: 9, padding: "10px 8px", border: `1px solid ${C.border}` }}>

                    <div style={{ fontSize: 10, color: C.muted, marginBottom: 5, textAlign: "center" }}>{d.label}</div>

                    <Inp type="number" min="0" step="1" value={form[d.key]} onChange={set(d.key)} placeholder="0" />

                    {qty > 0 && <div style={{ fontSize: 10, color: C.green, marginTop: 3, textAlign: "right" }}>{fmt(tot)}</div>}

                  </div>

                );

              })}

            </div>

            <div style={{ marginBottom: 10, fontSize: 11, color: C.muted, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em" }}>🪙 Coins</div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>

              {DENOMS.filter(d => d.type === "coin").map(d => {

                const qty = n(form[d.key]); const tot = qty * d.val;

                return (

                  <div key={d.key} style={{ background: C.surface, borderRadius: 9, padding: "10px 8px", border: `1px solid ${C.border}` }}>

                    <div style={{ fontSize: 10, color: C.muted, marginBottom: 5, textAlign: "center" }}>{d.label}</div>

                    <Inp type="number" min="0" step="1" value={form[d.key]} onChange={set(d.key)} placeholder="0" />

                    {qty > 0 && <div style={{ fontSize: 10, color: C.green, marginTop: 3, textAlign: "right" }}>{fmt(tot)}</div>}

                  </div>

                );

              })}

            </div>

            <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>

              {[["Notes", fmt(t.notesCash), C.blue], ["Coins", fmt(t.coinsCash), C.purple], ["Total Physical", fmt(t.physCash), C.green]].map(([l,v,c]) => (

                <div key={l} style={{ background: C.surface, borderRadius: 9, padding: "10px 12px", textAlign: "center", border: `1px solid ${C.border}` }}>

                  <div style={{ fontSize: 10, color: C.muted, marginBottom: 3 }}>{l}</div>

                  <div style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: c, fontSize: 15 }}>{v}</div>

                </div>

              ))}

            </div>

          </div>

        </Card>

 

        <Card>

          <CardHead icon="⚖️" title="8. Cash Variance" accent={t.variance === 0 ? C.green : C.red} />

          <div style={{ padding: 18 }}>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>

              {[["Expected", fmt(t.expectedCash), C.gold], ["Physical", fmt(t.physCash), C.blue], ["Variance", fmt(t.variance), t.variance === 0 ? C.green : C.red]].map(([l,v,c]) => (

                <div key={l} style={{ background: C.surface, borderRadius: 9, padding: "10px 12px", textAlign: "center", border: `1px solid ${C.border}` }}>

                  <div style={{ fontSize: 10, color: C.muted, marginBottom: 3 }}>{l}</div>

                  <div style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: c, fontSize: 15 }}>{v}</div>

                </div>

              ))}

            </div>

            <div style={{ padding: "12px 16px", borderRadius: 10, background: t.variance === 0 ? `${C.green}15` : `${C.red}15`, border: `1px solid ${t.variance === 0 ? C.green : C.red}`, textAlign: "center", fontFamily: FONT_HEAD, fontWeight: 700, color: t.variance === 0 ? C.green : C.red }}>

              {t.variance === 0 ? "✅ Cash Balanced" : t.variance > 0 ? `⚠️ Cash Over by ${fmt(Math.abs(t.variance))}` : `⚠️ Cash Short by ${fmt(Math.abs(t.variance))}`}

            </div>

          </div>

        </Card>

      </div>}

 

      {/* Step 3 – Sign Off */}

      {step === 3 && <div className="fade-up">

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="✍️" title="9. Sign Off" accent={C.blue} />

          <div style={{ padding: 18 }}>

            <Row2>

              <div><Lbl req>Staff Name</Lbl><Inp value={form.staff_name} onChange={set("staff_name")} placeholder="Your name" /></div>

              <div><Lbl>Manager Check</Lbl><Inp value={form.manager} onChange={set("manager")} placeholder="Manager name" /></div>

            </Row2>

            <div style={{ marginTop: 14 }}>

              <Lbl>Notes / Comments</Lbl>

              <textarea value={form.notes} onChange={set("notes")} placeholder="Any notes for today..."

                style={{ width: "100%", padding: "10px 13px", background: C.bg, border: `1.5px solid ${C.border2}`, borderRadius: 9, color: C.text, fontSize: 14, minHeight: 80, resize: "vertical", fontFamily: FONT_BODY }} />

            </div>

          </div>

        </Card>

 

        <Card style={{ marginBottom: 16 }}>

          <CardHead icon="📋" title="Summary" />

          <div style={{ padding: 18 }}>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>

              {[["Till Total", fmt(t.tillTotal)], ["JustEat Total", fmt(t.jeTotal)], ["Online Total", fmt(t.onlineTotal)], ["Expenses", fmt(t.totalExp)], ["Grand Total", fmt(t.grandTotal)], ["Expected Cash", fmt(t.expectedCash)]].map(([l,v]) => (

                <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "7px 10px", background: C.surface, borderRadius: 7 }}>

                  <span style={{ fontSize: 12, color: C.muted }}>{l}</span>

                  <span style={{ fontSize: 13, fontWeight: 700 }}>{v}</span>

                </div>

              ))}

            </div>

          </div>

        </Card>

      </div>}

 

      {/* Nav buttons */}

      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>

        <Btn variant="ghost" onClick={() => setStep(s => Math.max(0, s - 1))} disabled={step === 0}>← Back</Btn>

        {step < 3

          ? <Btn onClick={() => setStep(s => s + 1)}>Next →</Btn>

          : <Btn variant="success" onClick={handleSubmit} disabled={!form.staff_name}>✅ Submit Report</Btn>

        }

      </div>

    </div>

  );

}

 

// ─────────────────────────────────────────────────────────────────────────────

// ADMIN — DASHBOARD

// ─────────────────────────────────────────────────────────────────────────────

function AdminDashboard({ submissions }) {

  const approved = submissions.filter(s => s.status === STATUS.APPROVED);

  const pending  = submissions.filter(s => s.status === STATUS.PENDING);

  const rejected = submissions.filter(s => s.status === STATUS.REJECTED);

  const allT     = (key) => approved.reduce((s, sub) => s + (calcTotals(sub.form)[key] || 0), 0);

 

  // Last 7 days chart data

  const last7 = useMemo(() => {

    const days = [];

    for (let i = 6; i >= 0; i--) {

      const d = new Date(); d.setDate(d.getDate() - i);

      const ds = d.toISOString().slice(0, 10);

      const sub = approved.find(s => s.form.date === ds);

      days.push({ date: ds, day: shortDay(ds), total: sub ? calcTotals(sub.form).grandTotal : 0, hasData: !!sub });

    }

    return days;

  }, [submissions]);

 

  const maxTotal = Math.max(...last7.map(d => d.total), 1);

 

  return (

    <div>

      {/* KPI row */}

      <div className="fade-up" style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginBottom: 22 }}>

        <StatCard label="Week Grand Total" value={fmtShort(allT("grandTotal"))} color={C.flame} icon="💰" delay={0} />

        <StatCard label="Pending Reviews" value={pending.length} color={C.gold} icon="⏳" delay={1} />

        <StatCard label="Total Expenses" value={fmtShort(allT("totalExp"))} color={C.red} icon="💸" delay={2} />

        <StatCard label="Cash Variances" value={approved.filter(s => calcTotals(s.form).variance !== 0).length} color={approved.filter(s => calcTotals(s.form).variance !== 0).length > 0 ? C.red : C.green} icon="⚖️" delay={3} />

      </div>

 

      {/* 7-day bar chart */}

      <Card style={{ marginBottom: 18 }}>

        <CardHead icon="📈" title="Last 7 Days — Grand Total Sales" />

        <div style={{ padding: 20 }}>

          <div style={{ display: "flex", alignItems: "flex-end", gap: 8, height: 120 }}>

            {last7.map((d, i) => (

              <div key={i} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>

                <div style={{ fontSize: 10, color: C.flame, fontWeight: 700 }}>{d.total > 0 ? fmtShort(d.total) : ""}</div>

                <div style={{ width: "100%", background: d.hasData ? `linear-gradient(180deg, ${C.flame}, ${C.gold})` : C.border, borderRadius: "5px 5px 0 0", height: `${Math.max((d.total / maxTotal) * 80, d.hasData ? 4 : 0)}px`, transition: "height .4s ease", minHeight: d.hasData ? 4 : 0 }} />

                <div style={{ fontSize: 10, color: C.sub }}>{d.day}</div>

              </div>

            ))}

          </div>

        </div>

      </Card>

 

      {/* Platform breakdown */}

      <Card style={{ marginBottom: 18 }}>

        <CardHead icon="📊" title="Platform Breakdown (Approved Days)" />

        <div style={{ padding: 18 }}>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>

            {[["Till", allT("tillTotal"), C.flame], ["JustEat", allT("jeTotal"), C.gold], ["Deliveroo", allT("deliveroo"), C.blue], ["Uber Eats", allT("uber"), C.green], ["Foodhub", allT("foodhub"), C.red], ["Total Online", allT("onlineTotal"), C.purple]].map(([l,v,c]) => (

              <div key={l} style={{ background: C.surface, borderRadius: 10, padding: "12px 14px", border: `1px solid ${C.border}` }}>

                <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", marginBottom: 5 }}>{l}</div>

                <div style={{ fontFamily: FONT_HEAD, fontSize: 16, fontWeight: 700, color: c }}>{fmt(v)}</div>

              </div>

            ))}

          </div>

        </div>

      </Card>

 

      {/* Status summary */}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>

        {[["✅ Approved", approved.length, C.green], ["⏳ Pending", pending.length, C.gold], ["❌ Rejected", rejected.length, C.red]].map(([l,v,c]) => (

          <div key={l} style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 12, padding: "14px 16px", textAlign: "center" }}>

            <div style={{ fontSize: 11, color: C.muted, marginBottom: 5 }}>{l}</div>

            <div style={{ fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 800, color: c }}>{v}</div>

          </div>

        ))}

      </div>

    </div>

  );

}

 

// ─────────────────────────────────────────────────────────────────────────────

// ADMIN — REVIEW REPORTS

// ─────────────────────────────────────────────────────────────────────────────

function AdminReview({ submissions, onUpdate }) {

  const [selected, setSelected] = useState(null);

  const [rejectNote, setRejectNote] = useState("");

  const [filter, setFilter] = useState("all");

 

  const filtered = submissions.filter(s => filter === "all" ? true : s.status === filter).sort((a, b) => (b.id || 0) - (a.id || 0));

 

  const handleApprove = (sub) => {

    onUpdate(sub.id, { status: STATUS.APPROVED, rejectionNote: "" });

    setSelected(null);

  };

 

  const handleReject = (sub) => {

    if (!rejectNote.trim()) return;

    onUpdate(sub.id, { status: STATUS.REJECTED, rejectionNote: rejectNote, assignedTo: "samiya" });

    setRejectNote("");

    setSelected(null);

  };

 

  return (

    <div>

      {/* Filter tabs */}

      <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>

        {[["all","All"], ["pending","Pending"], ["approved","Approved"], ["rejected","Rejected"]].map(([v,l]) => (

          <button key={v} onClick={() => setFilter(v)}

            style={{ padding: "7px 16px", borderRadius: 10, background: filter === v ? C.flame : C.card, border: `1px solid ${filter === v ? C.flame : C.border}`, color: filter === v ? "#fff" : C.sub, fontWeight: 600, fontSize: 12, transition: "all .15s" }}>

            {l} {v !== "all" && <span style={{ marginLeft: 4, background: filter === v ? "#fff3" : C.border, padding: "1px 6px", borderRadius: 99, fontSize: 10 }}>{submissions.filter(s => s.status === v).length}</span>}

          </button>

        ))}

      </div>

 

      {filtered.length === 0 && (

        <div style={{ textAlign: "center", padding: 50, color: C.muted }}>

          <div style={{ fontSize: 36, marginBottom: 12 }}>📭</div>

          <div>No reports found</div>

        </div>

      )}

 

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>

        {filtered.map((sub) => {

          const t = calcTotals(sub.form);

          const isOpen = selected === sub.id;

          return (

            <div key={sub.id} style={{ background: C.card, border: `1px solid ${isOpen ? C.border2 : C.border}`, borderRadius: 13, overflow: "hidden", transition: "border-color .2s" }}>

              {/* Row */}

              <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 16px", cursor: "pointer" }} onClick={() => setSelected(isOpen ? null : sub.id)}>

                <div style={{ width: 38, height: 38, borderRadius: 10, background: sub.status === STATUS.APPROVED ? `${C.green}20` : sub.status === STATUS.REJECTED ? `${C.red}20` : `${C.gold}20`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>

                  {sub.status === STATUS.APPROVED ? "✅" : sub.status === STATUS.REJECTED ? "❌" : "⏳"}

                </div>

                <div style={{ flex: 1, minWidth: 0 }}>

                  <div style={{ fontWeight: 600, fontSize: 14, display: "flex", alignItems: "center", gap: 8 }}>

                    {dayName(sub.form.date)} <span style={{ color: C.muted, fontWeight: 400 }}>{sub.form.date}</span>

                    {statusBadge(sub.status)}

                  </div>

                  <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>By {sub.form.staff_name} · {sub.timestamp}</div>

                  {sub.status === STATUS.REJECTED && sub.rejectionNote && (

                    <div style={{ fontSize: 11, color: C.red, marginTop: 2 }}>↩ Reassigned to Samiya · "{sub.rejectionNote}"</div>

                  )}

                </div>

                <div style={{ textAlign: "right", flexShrink: 0 }}>

                  <div style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: C.flame, fontSize: 16 }}>{fmt(t.grandTotal)}</div>

                  <div style={{ fontSize: 11, color: t.variance === 0 ? C.green : C.red }}>{t.variance === 0 ? "Balanced" : `Var: ${fmt(t.variance)}`}</div>

                </div>

                <div style={{ color: C.muted, fontSize: 16, transform: isOpen ? "rotate(180deg)" : "none", transition: "transform .2s", flexShrink: 0 }}>▼</div>

              </div>

 

              {/* Expanded detail */}

              {isOpen && (

                <div style={{ borderTop: `1px solid ${C.border}`, padding: 18 }}>

                  {/* Totals grid */}

                  <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 16 }}>

                    {[["Till", fmt(t.tillTotal)], ["JustEat", fmt(t.jeTotal)], ["Platforms", fmt(t.onlineTotal)], ["Expenses", fmt(t.totalExp)], ["Expected Cash", fmt(t.expectedCash)], ["Physical Cash", fmt(t.physCash)]].map(([l,v]) => (

                      <div key={l} style={{ background: C.surface, borderRadius: 8, padding: "8px 12px" }}>

                        <div style={{ fontSize: 10, color: C.muted }}>{l}</div>

                        <div style={{ fontSize: 14, fontWeight: 700, marginTop: 2 }}>{v}</div>

                      </div>

                    ))}

                  </div>

 

                  {/* Variance status */}

                  <div style={{ padding: "9px 14px", borderRadius: 9, background: t.variance === 0 ? `${C.green}15` : `${C.red}15`, border: `1px solid ${t.variance === 0 ? C.green : C.red}`, color: t.variance === 0 ? C.green : C.red, fontWeight: 700, fontSize: 13, marginBottom: 16, textAlign: "center" }}>

                    {t.variance === 0 ? "✅ Cash Balanced" : t.variance > 0 ? `⚠️ Cash Over by ${fmt(Math.abs(t.variance))}` : `⚠️ Cash Short by ${fmt(Math.abs(t.variance))}`}

                  </div>

 

                  {sub.form.notes && (

                    <div style={{ padding: "9px 13px", background: C.surface, borderRadius: 8, fontSize: 13, color: C.sub, marginBottom: 16 }}>

                      📝 {sub.form.notes}

                    </div>

                  )}

 

                  {/* Action buttons */}

                  {sub.status === STATUS.PENDING && (

                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>

                      <div style={{ display: "flex", gap: 10 }}>

                        <Btn variant="success" onClick={() => handleApprove(sub)}>✅ Approve Report</Btn>

                        <Btn variant="ghost" onClick={() => generateExcel([sub], `Sattis_${sub.form.date}.xlsx`)}>⬇️ Download</Btn>

                      </div>

                      <div style={{ display: "flex", gap: 10, alignItems: "flex-end" }}>

                        <div style={{ flex: 1 }}>

                          <Lbl req>Rejection Reason (required to reject)</Lbl>

                          <Inp value={rejectNote} onChange={e => setRejectNote(e.target.value)} placeholder="Enter reason — will be shown to Samiya" />

                        </div>

                        <Btn variant="danger" onClick={() => handleReject(sub)} disabled={!rejectNote.trim()}>❌ Reject & Reassign</Btn>

                      </div>

                      <div style={{ fontSize: 11, color: C.muted }}>↩ Rejected reports are automatically reassigned to Samiya</div>

                    </div>

                  )}

 

                  {sub.status !== STATUS.PENDING && (

                    <div style={{ display: "flex", gap: 10 }}>

                      <Btn variant="ghost" onClick={() => generateExcel([sub], `Sattis_${sub.form.date}.xlsx`)}>⬇️ Download Excel</Btn>

                      <Btn variant="ghost" onClick={() => onUpdate(sub.id, { status: STATUS.PENDING, rejectionNote: "" })}>↩ Reset to Pending</Btn>

                    </div>

                  )}

                </div>

              )}

            </div>

          );

        })}

      </div>

    </div>

  );

}

 

// ─────────────────────────────────────────────────────────────────────────────

// ADMIN — SETTINGS

// ─────────────────────────────────────────────────────────────────────────────

function AdminSettings({ submissions, openingFloat, onFloatChange }) {

  const [floatInput, setFloatInput] = useState(openingFloat.toString());

  const [saved, setSaved] = useState(false);

 

  // Auto-pull yesterday's expected cash

  const yesterday = (() => {

    const d = new Date(); d.setDate(d.getDate() - 1);

    return d.toISOString().slice(0, 10);

  })();

  const yesterdaySub = submissions.filter(s => s.status === STATUS.APPROVED && s.form.date === yesterday)[0];

  const yesterdayExpected = yesterdaySub ? calcTotals(yesterdaySub.form).expectedCash : null;

 

  const handleSave = () => {

    const val = parseFloat(floatInput);

    if (isNaN(val) || val < 0) return;

    onFloatChange(val);

    setSaved(true);

    setTimeout(() => setSaved(false), 2000);

  };

 

  const handleAutoPull = () => {

    if (yesterdayExpected !== null) {

      setFloatInput(yesterdayExpected.toFixed(2));

    }

  };

 

  // Weekly download

  const handleWeeklyDownload = () => {

    const approved = submissions.filter(s => s.status === STATUS.APPROVED);

    if (approved.length === 0) { alert("No approved submissions to download."); return; }

    generateExcel(approved, `Sattis_WeeklyReport_${todayStr()}.xlsx`);

  };

 

  const handleAllDownload = () => {

    if (submissions.length === 0) { alert("No submissions yet."); return; }

    generateExcel(submissions, `Sattis_AllReports_${todayStr()}.xlsx`);

  };

 

  return (

    <div>

      {/* Opening Float */}

      <Card style={{ marginBottom: 18 }}>

        <CardHead icon="💷" title="Opening Float" accent={C.gold} />

        <div style={{ padding: 20 }}>

          <div style={{ marginBottom: 16, padding: "12px 16px", background: C.surface, borderRadius: 10, border: `1px solid ${C.border}` }}>

            <div style={{ fontSize: 12, color: C.muted, marginBottom: 3 }}>Current Opening Float</div>

            <div style={{ fontFamily: FONT_HEAD, fontSize: 26, fontWeight: 800, color: C.gold }}>{fmt(openingFloat)}</div>

          </div>

 

          {/* Auto-pull yesterday */}

          <div style={{ marginBottom: 16, padding: "12px 16px", background: `${C.blue}12`, border: `1px solid ${C.blue}40`, borderRadius: 10 }}>

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>

              <div>

                <div style={{ fontSize: 12, color: C.sub, marginBottom: 2 }}>Yesterday's Expected Cash in Hand</div>

                {yesterdayExpected !== null

                  ? <div style={{ fontFamily: FONT_HEAD, fontSize: 18, fontWeight: 700, color: C.blue }}>{fmt(yesterdayExpected)}</div>

                  : <div style={{ fontSize: 13, color: C.muted }}>No approved report for {yesterday}</div>

                }

              </div>

              <Btn variant="ghost" onClick={handleAutoPull} disabled={yesterdayExpected === null} size="sm">

                ⟳ Use as Today's Float

              </Btn>

            </div>

            <div style={{ fontSize: 11, color: C.muted, marginTop: 6 }}>Auto-pulls from yesterday's approved report Expected Cash in Hand</div>

          </div>

 

          <div style={{ display: "flex", gap: 10, alignItems: "flex-end" }}>

            <div style={{ flex: 1 }}>

              <Lbl>Set Opening Float (£)</Lbl>

              <Inp type="number" min="0" step="0.01" value={floatInput} onChange={e => setFloatInput(e.target.value)} placeholder="25.00" />

            </div>

            <Btn variant="gold" onClick={handleSave}>{saved ? "✅ Saved!" : "Save Float"}</Btn>

          </div>

          <div style={{ fontSize: 11, color: C.muted, marginTop: 8 }}>This value will be used as the opening float for all new staff entries</div>

        </div>

      </Card>

 

      {/* Downloads */}

      <Card style={{ marginBottom: 18 }}>

        <CardHead icon="⬇️" title="Download Reports" />

        <div style={{ padding: 20, display: "flex", flexDirection: "column", gap: 12 }}>

          <div style={{ display: "flex", gap: 12 }}>

            <div style={{ flex: 1 }}>

              <div style={{ fontWeight: 600, marginBottom: 3 }}>Weekly Report (Approved Only)</div>

              <div style={{ fontSize: 12, color: C.muted, marginBottom: 10 }}>Downloads all approved submissions with full weekly summary sheet</div>

              <Btn onClick={handleWeeklyDownload} full>⬇️ Download Weekly Excel</Btn>

            </div>

          </div>

          <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12 }}>

            <div style={{ fontWeight: 600, marginBottom: 3 }}>All Reports (Including Pending & Rejected)</div>

            <div style={{ fontSize: 12, color: C.muted, marginBottom: 10 }}>Includes every submission regardless of status</div>

            <Btn variant="ghost" onClick={handleAllDownload} full>⬇️ Download All Reports</Btn>

          </div>

        </div>

      </Card>

 

      {/* Submissions stats */}

      <Card>

        <CardHead icon="📊" title="All-Time Stats" />

        <div style={{ padding: 20 }}>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>

            {[

              ["Total Submissions", submissions.length, C.text],

              ["Approved", submissions.filter(s=>s.status===STATUS.APPROVED).length, C.green],

              ["Pending", submissions.filter(s=>s.status===STATUS.PENDING).length, C.gold],

              ["Rejected", submissions.filter(s=>s.status===STATUS.REJECTED).length, C.red],

              ["Avg Grand Total", fmt(submissions.filter(s=>s.status===STATUS.APPROVED).reduce((s,sub)=>s+calcTotals(sub.form).grandTotal,0) / (submissions.filter(s=>s.status===STATUS.APPROVED).length || 1)), C.flame],

              ["Total Revenue", fmt(submissions.filter(s=>s.status===STATUS.APPROVED).reduce((s,sub)=>s+calcTotals(sub.form).grandTotal,0)), C.flameLt],

            ].map(([l,v,c]) => (

              <div key={l} style={{ background: C.surface, borderRadius: 10, padding: "12px 14px", border: `1px solid ${C.border}` }}>

                <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>{l}</div>

                <div style={{ fontFamily: FONT_HEAD, fontWeight: 700, color: c, fontSize: 17 }}>{v}</div>

              </div>

            ))}

          </div>

        </div>

      </Card>

    </div>

  );

}

 

// ─────────────────────────────────────────────────────────────────────────────

// APP ROOT

// ─────────────────────────────────────────────────────────────────────────────

export default function App() {

  const [user, setUser]           = useState(null);

  const [adminTab, setAdminTab]   = useState("dashboard");

  const [submissions, setSubs]    = useState(loadSubmissions);

  const [openingFloat, setFloat]  = useState(loadFloat);

 

  useEffect(() => { saveSubmissions(submissions); }, [submissions]);

  useEffect(() => { saveFloat(openingFloat); }, [openingFloat]);

 

  const handleSubmit = (entry) => {

    setSubs(prev => [...prev, entry]);

  };

 

  const handleUpdate = (id, updates) => {

    setSubs(prev => prev.map(s => s.id === id ? { ...s, ...updates } : s));

  };

 

  const handleFloat = (val) => {

    setFloat(val);

  };

 

  const pendingCount = submissions.filter(s => s.status === STATUS.PENDING).length;

 

  // ADMIN TABS

  const adminTabs = [

    { key: "dashboard", label: "Dashboard",  icon: "📊" },

    { key: "review",    label: "Review",     icon: "📋", badge: pendingCount },

    { key: "settings",  label: "Settings",   icon: "⚙️" },

  ];

 

  if (!user) return (

    <>

      <style>{GS}</style>

      <Login onLogin={setUser} />

    </>

  );

 

  return (

    <>

      <style>{GS}</style>

 

      {/* Top Nav */}

      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, position: "sticky", top: 0, zIndex: 100, backdropFilter: "blur(12px)" }}>

        <div style={{ maxWidth: 820, margin: "0 auto", padding: "0 20px", display: "flex", alignItems: "center", height: 56, gap: 10 }}>

          {/* Logo */}

          <div style={{ display: "flex", alignItems: "center", gap: 9, marginRight: "auto" }}>

            <div style={{ width: 32, height: 32, borderRadius: 9, background: `linear-gradient(135deg, ${C.flame}, ${C.gold})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>🌶️</div>

            <div>

              <div style={{ fontFamily: FONT_HEAD, fontWeight: 800, fontSize: 14, color: C.flame, letterSpacing: "0.04em", lineHeight: 1 }}>SATTIS PERI PERI</div>

              <div style={{ fontSize: 9, color: C.muted, letterSpacing: "0.1em", textTransform: "uppercase" }}>Balance Sheet</div>

            </div>

          </div>

 

          {/* Admin tabs */}

          {user.role === ROLES.ADMIN && adminTabs.map(t => (

            <button key={t.key} onClick={() => setAdminTab(t.key)}

              style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 14px", borderRadius: 9, background: adminTab === t.key ? C.flame : "transparent", color: adminTab === t.key ? "#fff" : C.muted, fontWeight: 600, fontSize: 12, border: "none", transition: "all .15s", position: "relative" }}>

              {t.icon} {t.label}

              {t.badge > 0 && <span style={{ background: adminTab === t.key ? "#fff4" : C.red, color: "#fff", borderRadius: 99, fontSize: 9, fontWeight: 800, padding: "1px 5px" }}>{t.badge}</span>}

            </button>

          ))}

 

          {/* User + logout */}

          <div style={{ display: "flex", alignItems: "center", gap: 8, marginLeft: 8, paddingLeft: 12, borderLeft: `1px solid ${C.border}` }}>

            <div style={{ textAlign: "right" }}>

              <div style={{ fontSize: 12, fontWeight: 600 }}>{user.name}</div>

              <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", letterSpacing: "0.06em" }}>{user.role}</div>

            </div>

            <button onClick={() => setUser(null)} style={{ width: 30, height: 30, borderRadius: 8, background: C.card, border: `1px solid ${C.border}`, color: C.muted, fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center" }} title="Logout">⏻</button>

          </div>

        </div>

      </div>

 

      {/* Page */}

      <div style={{ maxWidth: 820, margin: "0 auto", padding: "24px 20px 60px" }}>

        {user.role === ROLES.STAFF && (

          <StaffForm user={user} openingFloat={openingFloat} submissions={submissions} onSubmit={handleSubmit} />

        )}

        {user.role === ROLES.ADMIN && adminTab === "dashboard" && <AdminDashboard submissions={submissions} />}

        {user.role === ROLES.ADMIN && adminTab === "review"    && <AdminReview submissions={submissions} onUpdate={handleUpdate} />}

        {user.role === ROLES.ADMIN && adminTab === "settings"  && <AdminSettings submissions={submissions} openingFloat={openingFloat} onFloatChange={handleFloat} />}

      </div>

    </>

  );

}
