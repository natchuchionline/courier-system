import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import {
  Package, Truck, Building2, ScanLine, Plus, LogOut, CheckCircle2, XCircle,
  RotateCcw, Search, ChevronDown, ChevronRight, Banknote, MapPin, Phone,
  Clock, ArrowRight, AlertTriangle, Users, Upload, FileText, X,
  Filter, RefreshCw, LayoutDashboard, ClipboardList, ShieldCheck, Lock,
  Printer, FileSpreadsheet, User, LocateFixed, MapPinned, Radar,
} from "lucide-react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";

/* ============================================================================
   KURSI — საკურიერო კომპანიის მართვის სისტემა
   ტოკენები:
   ფერი:  --bg #F4F5F1  --surface #FFFFFF  --border #DBDFDA  --ink #1C2321
          --ink-soft #62685F  --accent #DB5F1F (მარშრუტის ნარინჯისფერი)
          --track #2F5D8A (მარშრუტის ლურჯი)
   ტიპი:  "Noto Sans Georgian" ტექსტისთვის, "JetBrains Mono" — მხოლოდ
          სამარშრუტო კოდებისა და შტრიხკოდისთვის
   განლაგება: მარცხენა ნავიგაცია + ზედა სტატუს-ზოლი, კონტენტი — ზედნადების
          (waybill) სტილის ბარათები, დასაბეჭდი იარლიყი — ცალკე ბეჭდვის ფენა
============================================================================ */

const STATUS = {
  created:    { label: "შექმნილია",              color: "#62685F", bg: "#ECEDE9" },
  registered: { label: "მიღებულია საწყობში",       color: "#2F5D8A", bg: "#E9F0F7" },
  assigned:   { label: "გადაცემულია კურიერზე",      color: "#7452B0", bg: "#F0EAF8" },
  transit:    { label: "გზაშია",                  color: "#B8560E", bg: "#FBEAD9" },
  delivered:  { label: "მიწოდებულია",             color: "#1E8A5F", bg: "#E4F4EC" },
  failed:     { label: "ვერ ჩაბარდა",              color: "#B23A24", bg: "#FBE7E2" },
  returned:   { label: "დაბრუნებულია გამგზავნს",    color: "#8B5A3C", bg: "#F4EBE2" },
  cancelled:  { label: "გაუქმებულია",             color: "#9A9D97", bg: "#EFF0ED" },
};

const PAYMENT_METHODS = ["ქეშით კურიერთან (COD)", "წინასწარ გადახდილი", "ბანკო გადარიცხვა"];
const DELIVERY_TYPES = ["სტანდარტული", "ექსპრესი"];
const FAIL_REASONS = ["მიმღები არ იყო კონტაქტზე", "მისამართი არასწორია", "მიმღებმა თქვა უარი", "სხვა მიზეზი"];
const DEFAULT_DELIVERY_FEE = 5;
const DEFAULT_COMMISSION_TYPE = "percent"; // 'percent' | 'fixed'
const DEFAULT_COMMISSION_VALUE = 60; // 60% of the delivery fee by default

/** რამდენი ერგება კურიერს ამ ამანათის საკურიერო საფასურიდან */
function courierCommissionFor(parcel, courier) {
  if (!courier) return 0;
  const fee = Number(parcel.deliveryFee) || 0;
  const type = courier.commissionType || DEFAULT_COMMISSION_TYPE;
  const value = courier.commissionValue == null ? DEFAULT_COMMISSION_VALUE : Number(courier.commissionValue);
  if (type === "percent") return Math.max(0, fee * (value / 100));
  return Math.max(0, value);
}
/** რა ეკუთვნის გამგზავნ კომპანიას ამ ამანათზე (დადებითი = ჩაერიცხება, უარყოფითი = კომპანიის დავალიანება) */
function companyNetFor(parcel) {
  const fee = Number(parcel.deliveryFee) || 0;
  if (parcel.paymentMethod === PAYMENT_METHODS[0]) return (Number(parcel.codCollectedAmount || parcel.codAmount) || 0) - fee;
  return -fee;
}
function downloadCsv(filename, rows) {
  const csv = Papa.unparse(rows);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
function downloadXlsx(filename, rows, sheetName = "Sheet1") {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function pad(n) { return n < 10 ? "0" + n : "" + n; }
function genTrackingId() {
  const d = new Date();
  const rand = Math.floor(1000 + Math.random() * 9000);
  return `KR${(d.getFullYear() % 100)}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${rand}`;
}
function genId(prefix) {
  return prefix + "_" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function nowIso() { return new Date().toISOString(); }
function fmtDate(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toFixed(2) + " ₾";
}

/* ------------------------------ Code 39 barcode ---------------------------- */
/* გადამოწმებული "N"(ვიწრო)/"W"(განიერი) 9-ელემენტიანი ცხრილი Code 39
   სტანდარტისთვის — თითოეული სიმბოლო = 5 ზოლი + 4 დაშორება. */
const CODE39_WIDTHS = {
  "0": "NNNWWNWNN", "1": "WNNWNNNNW", "2": "NNWWNNNNW", "3": "WNWWNNNNN", "4": "NNNWWNNNW",
  "5": "WNNWWNNNN", "6": "NNWWWNNNN", "7": "NNNWNNWNW", "8": "WNNWNNWNN", "9": "NNWWNNWNN",
  "A": "WNNNNWNNW", "B": "NNWNNWNNW", "C": "WNWNNWNNN", "D": "NNNNWWNNW", "E": "WNNNWWNNN",
  "F": "NNWNWWNNN", "G": "NNNNNWWNW", "H": "WNNNNWWNN", "I": "NNWNNWWNN", "J": "NNNNWWWNN",
  "K": "WNNNNNNWW", "L": "NNWNNNNWW", "M": "WNWNNNNWN", "N": "NNNNWNNWW", "O": "WNNNWNNWN",
  "P": "NNWNWNNWN", "Q": "NNNNNNWWW", "R": "WNNNNNWWN", "S": "NNWNNNWWN", "T": "NNNNWNWWN",
  "U": "WWNNNNNNW", "V": "NWWNNNNNW", "W": "WWWNNNNNN", "X": "NWNNWNNNW", "Y": "WWNNWNNNN",
  "Z": "NWWNWNNNN", "-": "NWNNNNWNW", ".": "WWNNNNWNN", " ": "NWWNNNWNN", "$": "NWNWNWNNN",
  "/": "NWNWNNNWN", "+": "NWNNNWNWN", "%": "NNNWNWNWN", "*": "NWNNWNWNN",
};
function code39Elements(text) {
  const chars = ("*" + text.toUpperCase() + "*").split("");
  const elements = [];
  chars.forEach((ch, idx) => {
    const pattern = CODE39_WIDTHS[ch];
    if (!pattern) return;
    for (let i = 0; i < 9; i++) elements.push({ bar: i % 2 === 0, w: pattern[i] === "W" ? 3 : 1 });
    if (idx < chars.length - 1) elements.push({ bar: false, w: 1 });
  });
  return elements;
}
function Code39Barcode({ value, height = 54, unit = 2 }) {
  const elements = useMemo(() => code39Elements(value), [value]);
  const totalWidth = elements.reduce((s, e) => s + e.w, 0) * unit;
  let x = 0;
  const bars = [];
  elements.forEach((e, i) => {
    if (e.bar) bars.push(<rect key={i} x={x} y={0} width={e.w * unit} height={height} fill="#15181A" />);
    x += e.w * unit;
  });
  return (
    <svg viewBox={`0 0 ${totalWidth} ${height}`} width="100%" style={{ maxWidth: 320, height }} preserveAspectRatio="xMidYMid meet">
      <rect x={0} y={0} width={totalWidth} height={height} fill="#fff" />
      {bars}
    </svg>
  );
}

/* ---------------------------- storage helpers ---------------------------- */

/** ერთდროულად მაქსიმუმ `limit` ამოცანას უშვებს — თავიდან იცილებს იმას, რომ
    ათასობით ჩანაწერის ერთდროული წამოღება გადატვირთოს საცავს ან დაისაქმოს
    ბრაუზერის მეხსიერება; დიდი მოცულობის დროს სწრაფი და სტაბილურია. */
async function mapLimit(items, limit, fn) {
  const results = new Array(items.length);
  let idx = 0;
  async function worker() {
    while (idx < items.length) {
      const current = idx++;
      results[current] = await fn(items[current], current);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}
async function listPrefixed(prefix) {
  try {
    const { data, error } = await supabase.from("kv_store").select("value").like("key", `${prefix}%`);
    if (error) throw error;
    return (data || []).map((row) => row.value).filter(Boolean);
  } catch (e) { console.error("listPrefixed failed", prefix, e); return []; }
}
async function putShared(key, value) {
  try {
    const { error } = await supabase.from("kv_store").upsert({ key, value });
    if (error) throw error;
    return true;
  } catch (e) { console.error("putShared failed", key, e); return false; }
}
async function deleteShared(key) {
  try {
    const { error } = await supabase.from("kv_store").delete().eq("key", key);
    if (error) throw error;
    return true;
  } catch (e) { console.error("deleteShared failed", key, e); return false; }
}
async function getPersonal(key) {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : null; }
  catch (e) { return null; }
}
async function setPersonal(key, value) {
  try {
    if (value == null) localStorage.removeItem(key);
    else localStorage.setItem(key, JSON.stringify(value));
  } catch (e) { console.error("setPersonal failed", key, e); }
}

/* --------------------------------- style ---------------------------------- */

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Georgian:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');
      .kursi { --bg:#F4F5F1; --surface:#FFFFFF; --border:#DBDFDA; --ink:#1C2321; --ink-soft:#62685F;
               --accent:#DB5F1F; --accent-ink:#FFFFFF; --track:#2F5D8A;
               font-family:'Noto Sans Georgian','Noto Sans',sans-serif; color:var(--ink); background:var(--bg); }
      .kursi .mono { font-family:'JetBrains Mono',monospace; letter-spacing:.02em; }
      .kursi .panel { background:var(--surface); border:1px solid var(--border); border-radius:8px; }
      .kursi .btn { display:inline-flex; align-items:center; gap:6px; padding:9px 14px; border-radius:6px;
                    font-size:14px; font-weight:500; border:1px solid transparent; cursor:pointer; transition:filter .12s ease, background .12s ease; white-space:nowrap; }
      .kursi .btn:disabled { opacity:.45; cursor:not-allowed; }
      .kursi .btn-primary { background:var(--accent); color:#fff; }
      .kursi .btn-primary:hover:not(:disabled) { filter:brightness(0.93); }
      .kursi .btn-outline { background:transparent; border-color:var(--border); color:var(--ink); }
      .kursi .btn-outline:hover:not(:disabled) { background:#EFEFEB; }
      .kursi .btn-ghost { background:transparent; color:var(--ink-soft); }
      .kursi .btn-ghost:hover:not(:disabled) { background:#EFEFEB; color:var(--ink); }
      .kursi .btn-track { background:var(--track); color:#fff; }
      .kursi .btn-track:hover:not(:disabled) { filter:brightness(0.93); }
      .kursi .btn-danger { background:#B23A24; color:#fff; }
      .kursi .btn-danger:hover:not(:disabled) { filter:brightness(0.93); }
      .kursi .btn-sm { padding:6px 10px; font-size:13px; }
      .kursi .field-label { font-size:12.5px; color:var(--ink-soft); margin-bottom:4px; display:block; }
      .kursi .input, .kursi .select, .kursi .textarea {
        width:100%; border:1px solid var(--border); border-radius:6px; padding:8px 10px; font-size:14px;
        background:#fff; color:var(--ink); font-family:inherit; }
      .kursi .input:focus, .kursi .select:focus, .kursi .textarea:focus {
        outline:2px solid var(--track); outline-offset:1px; border-color:var(--track); }
      .kursi .textarea { resize:vertical; min-height:64px; }
      .kursi .pill { display:inline-flex; align-items:center; gap:5px; padding:3px 10px; border-radius:999px; font-size:12.5px; font-weight:600; }
      .kursi .navlink { display:flex; align-items:center; gap:10px; padding:10px 12px; border-radius:6px; font-size:14px; color:var(--ink-soft); cursor:pointer; }
      .kursi .navlink:hover { background:#ECEDE8; color:var(--ink); }
      .kursi .navlink.active { background:var(--ink); color:#fff; }
      .kursi .waybill { border:1px solid var(--border); border-radius:8px; background:#fff; padding:14px 16px; position:relative; }
      .kursi .waybill:before { content:''; position:absolute; left:14px; right:14px; top:64px; border-top:1px dashed var(--border); }
      .kursi table.data th { text-align:left; font-size:12px; color:var(--ink-soft); font-weight:600; padding:8px 10px; border-bottom:1px solid var(--border); }
      .kursi table.data td { padding:9px 10px; border-bottom:1px solid var(--border); font-size:13.5px; vertical-align:top; }
      .kursi table.data tr:last-child td { border-bottom:none; }
      .kursi .stat { padding:12px 14px; border:1px solid var(--border); border-radius:8px; background:#fff; }
      .kursi .scanbox { border:2px dashed var(--track); border-radius:10px; background:#F4F8FB; }
      .kursi .toast { position:fixed; top:16px; right:16px; z-index:60; padding:12px 16px; border-radius:8px; font-size:14px;
                      box-shadow:0 4px 16px rgba(0,0,0,.14); color:#fff; }
      .kursi .blink-red { animation: kursiBlink 1.1s ease-in-out infinite; }
      @keyframes kursiBlink { 0%,100% { opacity:1; } 50% { opacity:.35; } }
      @media (prefers-reduced-motion: reduce) { .kursi .blink-red { animation:none; } }
      .kursi ::-webkit-scrollbar { height:8px; width:8px; }
      .kursi ::-webkit-scrollbar-thumb { background:#CBCFC7; border-radius:4px; }
      @media (max-width:767px) {
        .kursi .wb-grid { grid-template-columns:1fr !important; }
        .kursi .wb-grid svg { transform:rotate(90deg); margin:2px 0; }
      }
      @media print {
        body * { visibility:hidden !important; }
        #print-label-root, #print-label-root * { visibility:visible !important; }
        #print-label-root { position:fixed !important; inset:0; background:#fff; padding:18px; }
        #print-label-root .no-print { display:none !important; }
      }
    `}</style>
  );
}

/* ------------------------------- small bits ------------------------------- */

function StatusPill({ status }) {
  const s = STATUS[status] || STATUS.created;
  return <span className="pill" style={{ color: s.color, background: s.bg }}>● {s.label}</span>;
}
function DispatchWarning({ parcel }) {
  if (parcel.status !== "created") return null;
  if (parcel.dispatchedAt) return <span className="pill" style={{ color: "#2F5D8A", background: "#E9F0F7" }}>◐ გზაშია საწყობისკენ</span>;
  return <span className="pill blink-red" style={{ color: "#fff", background: "#C6432B" }}><AlertTriangle size={11} /> საწყობში გაგზავნილი არ არის</span>;
}
function Toast({ toast }) {
  if (!toast) return null;
  const bg = toast.type === "error" ? "#B23A24" : toast.type === "warn" ? "#B8560E" : "#1E8A5F";
  return <div className="toast" style={{ background: bg }}>{toast.msg}</div>;
}
function EmptyState({ icon: Icon, title, hint, action }) {
  return (
    <div className="panel" style={{ padding: "40px 24px", textAlign: "center" }}>
      <Icon size={30} style={{ margin: "0 auto 10px", color: "var(--ink-soft)" }} />
      <div style={{ fontWeight: 600, marginBottom: 4 }}>{title}</div>
      {hint && <div style={{ fontSize: 13.5, color: "var(--ink-soft)", marginBottom: action ? 14 : 0 }}>{hint}</div>}
      {action}
    </div>
  );
}
function Field({ label, children }) {
  return <div><span className="field-label">{label}</span>{children}</div>;
}

/* ------------------------------ waybill card ------------------------------ */

function Waybill({ parcel, children }) {
  return (
    <div className="waybill">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8, flexWrap: "wrap" }}>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" }}>
          <StatusPill status={parcel.status} />
          <DispatchWarning parcel={parcel} />
        </div>
        <div className="mono" style={{ fontSize: 14, fontWeight: 600, textAlign: "right" }}>{parcel.id}</div>
      </div>
      <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: 4 }}>{parcel.companyName}</div>
      <div style={{ height: 14 }} />
      <div style={{ display: "grid", gridTemplateColumns: "1fr auto 1fr", gap: 10, alignItems: "center" }} className="wb-grid">
        <div>
          <div style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>გამგზავნი</div>
          <div style={{ fontWeight: 600, fontSize: 14 }}>{parcel.sender.name}</div>
          <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>{parcel.sender.address}{parcel.sender.city ? `, ${parcel.sender.city}` : ""}</div>
        </div>
        <ArrowRight size={16} color="var(--ink-soft)" />
        <div>
          <div style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>მიმღები</div>
          <div style={{ fontWeight: 600, fontSize: 14 }}>{parcel.receiver.name}</div>
          <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>{parcel.receiver.address}{parcel.receiver.city ? `, ${parcel.receiver.city}` : ""}</div>
        </div>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 14, marginTop: 12, fontSize: 12.5, color: "var(--ink-soft)" }}>
        <span>კურიერი: <b style={{ color: "var(--ink)" }}>{parcel.courierName || "არ არის მიმაგრებული"}</b></span>
        <span>წონა: <b style={{ color: "var(--ink)" }}>{parcel.weight || "—"} კგ</b></span>
        <span>საკურიერო საფასური: <b style={{ color: "var(--ink)" }}>{fmtMoney(parcel.deliveryFee)}</b></span>
        {Number(parcel.codAmount) > 0 && (
          <span style={{ color: "var(--accent)", fontWeight: 700 }}>
            <Banknote size={13} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />
            COD {fmtMoney(parcel.codAmount)}{parcel.codCollected ? " ✓ მიღებული" : ""}
          </span>
        )}
        <span>შექმნა: {fmtDate(parcel.createdAt)}</span>
      </div>
      {children && <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid var(--border)" }}>{children}</div>}
    </div>
  );
}

function HistoryTimeline({ history }) {
  if (!history || history.length === 0) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
      {history.slice().reverse().map((h, i) => (
        <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, color: "var(--ink-soft)" }}>
          <Clock size={13} style={{ marginTop: 2, flexShrink: 0 }} />
          <div><b style={{ color: "var(--ink)" }}>{STATUS[h.status]?.label || h.status}</b> — {fmtDate(h.at)}{h.note ? ` · ${h.note}` : ""}</div>
        </div>
      ))}
    </div>
  );
}

/* ------------------------------ GPS location ------------------------------- */
/* რუკის ჩართული ბიბლიოთეკის გარეშეც საიმედოდ მუშაობს: Google Maps-ს ბმულით
   ახალ ტაბში იხსნება ზუსტი კოორდინატებით. */
function mapsLink(lat, lng) { return `https://www.google.com/maps?q=${lat},${lng}`; }

function LocationStatus({ courier }) {
  if (!courier) return null;
  if (courier.lastLat != null && courier.lastLng != null) {
    return (
      <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>
        <MapPinned size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />
        ბოლო GPS მდებარეობა: {fmtDate(courier.lastLocationAt)} —{" "}
        <a href={mapsLink(courier.lastLat, courier.lastLng)} target="_blank" rel="noreferrer" style={{ color: "var(--track)" }}>რუკაზე ნახვა</a>
      </span>
    );
  }
  if (courier.lastLocationText) {
    return (
      <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>
        <MapPinned size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />
        ბოლო ცნობილი მდებარეობა (ხელით): {courier.lastLocationText} · {fmtDate(courier.lastLocationAt)}
      </span>
    );
  }
  return <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>მდებარეობა ჯერ არ არის განახლებული</span>;
}

function LocationUpdater({ courierId, onUpdateCourier }) {
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [manual, setManual] = useState(false);
  const [manualText, setManualText] = useState("");
  const [tracking, setTracking] = useState(false);
  const watchIdRef = useRef(null);
  const lastWriteRef = useRef(0);

  function writePosition(pos, { throttle }) {
    const now = Date.now();
    if (throttle && now - lastWriteRef.current < 20000) return; // მაქს. 1 ჩაწერა 20 წამში, საცავს რომ არ დავტვირთოთ
    lastWriteRef.current = now;
    onUpdateCourier(courierId, { lastLat: pos.coords.latitude, lastLng: pos.coords.longitude, lastLocationAt: nowIso(), lastLocationText: null });
  }
  function handleWatchError(geoErr) {
    setTracking(false);
    if (watchIdRef.current != null) { navigator.geolocation.clearWatch(watchIdRef.current); watchIdRef.current = null; }
    setErr(geoErr.code === 1 ? "GPS-ზე წვდომა უარყოფილია — შეამოწმეთ ბრაუზერის ნებართვები, ან შეიყვანეთ ხელით" : "მდებარეობის თვალთვალი შეწყდა — სცადეთ ხელახლა ან შეიყვანეთ ხელით");
    setManual(true);
  }

  function requestOnce() {
    if (!navigator.geolocation) { setErr("ამ ბრაუზერში GPS მხარდაჭერილი არ არის"); setManual(true); return; }
    setBusy(true); setErr("");
    navigator.geolocation.getCurrentPosition(
      (pos) => { setBusy(false); writePosition(pos, { throttle: false }); },
      (geoErr) => { setBusy(false); setErr(geoErr.code === 1 ? "GPS-ზე წვდომა უარყოფილია — შეამოწმეთ ბრაუზერის ნებართვები, ან შეიყვანეთ ხელით" : "მდებარეობის დადგენა ვერ მოხერხდა — სცადეთ ხელახლა ან შეიყვანეთ ხელით"); setManual(true); },
      { enableHighAccuracy: true, timeout: 8000 }
    );
  }
  function toggleTracking() {
    if (tracking) {
      if (watchIdRef.current != null) navigator.geolocation.clearWatch(watchIdRef.current);
      watchIdRef.current = null; setTracking(false);
      return;
    }
    if (!navigator.geolocation) { setErr("ამ ბრაუზერში GPS მხარდაჭერილი არ არის"); setManual(true); return; }
    setErr("");
    watchIdRef.current = navigator.geolocation.watchPosition(
      (pos) => writePosition(pos, { throttle: true }),
      handleWatchError,
      { enableHighAccuracy: true, maximumAge: 15000, timeout: 20000 }
    );
    setTracking(true);
  }
  useEffect(() => () => { if (watchIdRef.current != null) navigator.geolocation.clearWatch(watchIdRef.current); }, []);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        <button className="btn btn-outline btn-sm" disabled={busy} onClick={requestOnce}>
          <LocateFixed size={13} /> {busy ? "მიმდინარეობს..." : "ერთჯერადი განახლება"}
        </button>
        <button className={tracking ? "btn btn-danger btn-sm" : "btn btn-track btn-sm"} onClick={toggleTracking}>
          <Radar size={13} /> {tracking ? "თვალთვალის შეჩერება" : "უწყვეტი თვალთვალის დაწყება"}
        </button>
        {!manual && <button className="btn btn-ghost btn-sm" onClick={() => setManual(true)}>ხელით შეყვანა</button>}
      </div>
      {tracking && <div style={{ fontSize: 12, color: "var(--track)" }}>● თვალთვალი აქტიურია — მდებარეობა თავისთავად განახლდება, სანამ ეს გვერდი ღიაა და აქტიურია</div>}
      {err && <div style={{ color: "#B23A24", fontSize: 12.5 }}>{err}</div>}
      {manual && (
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <input className="input" style={{ maxWidth: 280 }} placeholder="მაგ. ვაჟა-ფშაველას გამზირი, თბილისი" value={manualText} onChange={(e) => setManualText(e.target.value)} />
          <button className="btn btn-primary btn-sm" disabled={!manualText.trim()} onClick={() => { onUpdateCourier(courierId, { lastLocationText: manualText.trim(), lastLocationAt: nowIso(), lastLat: null, lastLng: null }); setManualText(""); setManual(false); setErr(""); }}>შენახვა</button>
        </div>
      )}
    </div>
  );
}

/* ------------------------------ print label -------------------------------- */

function PrintLabelModal({ parcel, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(28,35,33,0.5)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 70 }} onClick={onClose}>
      <div id="print-label-root" style={{ background: "#fff", borderRadius: 10, width: "100%", maxWidth: 440, maxHeight: "92vh", overflowY: "auto", boxShadow: "0 12px 40px rgba(0,0,0,.25)" }} onClick={(e) => e.stopPropagation()}>
        <div className="no-print" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 16px", borderBottom: "1px solid var(--border)" }}>
          <div style={{ fontWeight: 600, fontSize: 14 }}>ამანათის იარლიყი</div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="btn btn-primary btn-sm" onClick={() => window.print()}><Printer size={13} /> ბეჭდვა</button>
            <button className="btn btn-ghost btn-sm" onClick={onClose}><X size={15} /></button>
          </div>
        </div>
        <div style={{ padding: 20, border: "2px solid #15181A", borderRadius: 6, margin: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 20, height: 20, borderRadius: 5, background: "#15181A", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Truck size={12} color="#fff" />
              </div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>კურსი</div>
            </div>
            <div style={{ fontSize: 11, color: "#62685F" }}>{parcel.deliveryType}</div>
          </div>
          <div style={{ height: 12 }} />
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
            <Code39Barcode value={parcel.id} />
            <div className="mono" style={{ fontSize: 16, fontWeight: 700, marginTop: 4, letterSpacing: "0.06em" }}>{parcel.id}</div>
          </div>
          <div style={{ height: 1, background: "#15181A", opacity: 0.15, margin: "14px 0" }} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div>
              <div style={{ fontSize: 10.5, color: "#62685F", fontWeight: 600 }}>გამგზავნი</div>
              <div style={{ fontWeight: 700, fontSize: 13.5 }}>{parcel.sender.name}</div>
              <div style={{ fontSize: 12 }}>{parcel.sender.phone}</div>
              <div style={{ fontSize: 12 }}>{parcel.sender.address}{parcel.sender.city ? `, ${parcel.sender.city}` : ""}</div>
            </div>
            <div>
              <div style={{ fontSize: 10.5, color: "#62685F", fontWeight: 600 }}>მიმღები</div>
              <div style={{ fontWeight: 700, fontSize: 13.5 }}>{parcel.receiver.name}</div>
              <div style={{ fontSize: 12 }}>{parcel.receiver.phone}</div>
              <div style={{ fontSize: 12 }}>{parcel.receiver.address}{parcel.receiver.city ? `, ${parcel.receiver.city}` : ""}</div>
            </div>
          </div>
          <div style={{ height: 1, background: "#15181A", opacity: 0.15, margin: "14px 0" }} />
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12 }}>
            <span>ცალები: <b>{parcel.pieces || 1}</b></span>
            <span>წონა: <b>{parcel.weight || "—"} კგ</b></span>
            <span>შემცველობა: <b>{parcel.contents || "—"}</b></span>
          </div>
          {Number(parcel.codAmount) > 0 && (
            <div style={{ marginTop: 12, padding: "10px 12px", background: "#FBEAD9", borderRadius: 6, textAlign: "center" }}>
              <div style={{ fontSize: 11, color: "#B8560E", fontWeight: 600 }}>ასაკრები თანხა — COD</div>
              <div style={{ fontSize: 22, fontWeight: 700, color: "#B8560E" }}>{fmtMoney(parcel.codAmount)}</div>
            </div>
          )}
          {parcel.notes && <div style={{ marginTop: 10, fontSize: 11.5, color: "#62685F" }}>შენიშვნა: {parcel.notes}</div>}
        </div>
      </div>
    </div>
  );
}

/* =============================== APP ROOT ================================ */

export default function App() {
  const [identity, setIdentity] = useState(null); // {role, id, name}
  const [booting, setBooting] = useState(true);
  const [companies, setCompanies] = useState([]);
  const [couriers, setCouriers] = useState([]);
  const [registrators, setRegistrators] = useState([]);
  const [admins, setAdmins] = useState([]);
  const [parcels, setParcels] = useState([]);
  const [loadingData, setLoadingData] = useState(true);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((msg, type = "ok") => {
    setToast({ msg, type });
    window.clearTimeout(showToast._t);
    showToast._t = window.setTimeout(() => setToast(null), 3200);
  }, []);

  const refreshAll = useCallback(async () => {
    setLoadingData(true);
    const [co, cr, rg, ad, pc] = await Promise.all([
      listPrefixed("company:"), listPrefixed("courier:"), listPrefixed("registrator:"), listPrefixed("admin:"), listPrefixed("parcel:"),
    ]);
    const byName = (a, b) => (a.name || "").localeCompare(b.name || "", "ka");
    co.sort(byName); cr.sort(byName); rg.sort(byName); ad.sort(byName);
    pc.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    setCompanies(co); setCouriers(cr); setRegistrators(rg); setAdmins(ad); setParcels(pc);
    setLoadingData(false);
  }, []);

  useEffect(() => {
    (async () => {
      const last = await getPersonal("last-identity");
      if (last) setIdentity(last);
      await refreshAll();
      setBooting(false);
    })();
  }, [refreshAll]);

  function chooseIdentity(id) {
    setIdentity(id);
    setPersonal("last-identity", id);
  }
  function logout() {
    setIdentity(null);
    setPersonal("last-identity", null);
  }
  function accountsFor(role) {
    return { company: companies, courier: couriers, registrator: registrators, admin: admins }[role] || [];
  }
  function tryLogin(role, username, pin) {
    const acct = accountsFor(role).find((a) => (a.username || "").trim().toLowerCase() === username.trim().toLowerCase());
    if (!acct) return { ok: false, reason: "notfound" };
    if ((acct.pin || "") !== pin) return { ok: false, reason: "badpin" };
    chooseIdentity({ role, id: acct.id, name: acct.name });
    return { ok: true };
  }

  /* --------------------------- account mutations --------------------------- */

  async function addCompany({ name, contactPerson, phone, address, taxId, username, pin }) {
    const c = { id: genId("company"), name, contactPerson, phone, address, taxId, username, pin, createdAt: nowIso() };
    const ok = await putShared("company:" + c.id, c);
    if (ok) { setCompanies((p) => [...p, c].sort((a, b) => a.name.localeCompare(b.name, "ka"))); showToast("კომპანია დარეგისტრირდა"); }
    else showToast("ვერ დაემატა კომპანია — სცადეთ ხელახლა", "error");
    return ok ? c : null;
  }
  async function addCourier({ name, phone, vehicle, zone, commissionType, commissionValue, username, pin }) {
    const c = {
      id: genId("courier"), name, phone, vehicle, zone, username, pin,
      commissionType: commissionType || "percent",
      commissionValue: commissionValue === "" || commissionValue == null ? 60 : Number(commissionValue),
      createdAt: nowIso(),
    };
    const ok = await putShared("courier:" + c.id, c);
    if (ok) { setCouriers((p) => [...p, c].sort((a, b) => a.name.localeCompare(b.name, "ka"))); showToast("კურიერი დარეგისტრირდა"); }
    else showToast("ვერ დაემატა კურიერი — სცადეთ ხელახლა", "error");
    return ok ? c : null;
  }
  async function addRegistrator({ name, username, pin }) {
    const r = { id: genId("registrator"), name, username, pin, createdAt: nowIso() };
    const ok = await putShared("registrator:" + r.id, r);
    if (ok) { setRegistrators((p) => [...p, r].sort((a, b) => a.name.localeCompare(b.name, "ka"))); showToast("რეგისტრატორი დარეგისტრირდა"); }
    else showToast("ვერ დაემატა — სცადეთ ხელახლა", "error");
    return ok ? r : null;
  }
  async function addAdmin({ name, username, pin }) {
    const a = { id: genId("admin"), name, username, pin, createdAt: nowIso() };
    const ok = await putShared("admin:" + a.id, a);
    if (ok) { setAdmins((p) => [...p, a].sort((x, y) => x.name.localeCompare(y.name, "ka"))); showToast("ადმინისტრატორი დარეგისტრირდა"); }
    else showToast("ვერ დაემატა — სცადეთ ხელახლა", "error");
    return ok ? a : null;
  }
  async function updateCourier(courierId, patch) {
    const courier = couriers.find((c) => c.id === courierId);
    if (!courier) return false;
    const updated = { ...courier, ...patch };
    const ok = await putShared("courier:" + updated.id, updated);
    if (ok) { setCouriers((p) => p.map((c) => (c.id === courierId ? updated : c))); showToast("კურიერის მონაცემები განახლდა"); }
    else showToast("ვერ განახლდა — სცადეთ ხელახლა", "error");
    return ok;
  }
  async function clearAllCompanies() {
    await mapLimit(companies, 25, (c) => deleteShared("company:" + c.id));
    setCompanies([]);
    showToast("ყველა კომპანია წაიშალა");
  }
  async function clearAllCouriers() {
    await mapLimit(couriers, 25, (c) => deleteShared("courier:" + c.id));
    setCouriers([]);
    showToast("ყველა კურიერი წაიშალა");
  }
  async function clearAllParcels() {
    await mapLimit(parcels, 25, (p) => deleteShared("parcel:" + p.id));
    setParcels([]);
    showToast("ყველა ამანათი წაიშალა");
  }
  async function clearAllRegistrators() {
    await mapLimit(registrators, 25, (r) => deleteShared("registrator:" + r.id));
    setRegistrators([]);
    showToast("ყველა რეგისტრატორი წაიშალა");
  }
  async function clearAllAdmins() {
    // მიმდინარე ადმინის საკუთარ ანგარიშს არ ვშლით, რომ თავად არ ჩაიკეტოთ სისტემიდან
    const toDelete = admins.filter((a) => a.id !== identity.id);
    await mapLimit(toDelete, 25, (a) => deleteShared("admin:" + a.id));
    setAdmins((prev) => prev.filter((a) => a.id === identity.id));
    showToast(toDelete.length === admins.length ? "ყველა ადმინისტრატორი წაიშალა" : "წაიშალა — გარდა თქვენი მიმდინარე ანგარიშისა");
  }

  /* --------------------------- parcel mutations ----------------------------- */

  async function createParcel(data) {
    const id = genTrackingId();
    const parcel = {
      id,
      companyId: identity.id, companyName: identity.name,
      sender: data.sender, receiver: data.receiver,
      pieces: data.pieces || 1, weight: data.weight, dimensions: data.dimensions,
      contents: data.contents, declaredValue: data.declaredValue,
      deliveryType: data.deliveryType, paymentMethod: data.paymentMethod,
      codAmount: data.paymentMethod === PAYMENT_METHODS[0] ? Number(data.codAmount) || 0 : 0,
      deliveryFee: data.deliveryFee === "" || data.deliveryFee == null ? DEFAULT_DELIVERY_FEE : Number(data.deliveryFee),
      codCollected: false, codCollectedAmount: 0, codCollectedAt: null, deliveredAt: null, dispatchedAt: null,
      preferredDate: data.preferredDate || null, notes: data.notes || "",
      status: "created", courierId: null, courierName: null,
      createdAt: nowIso(),
      history: [{ status: "created", at: nowIso(), note: "შეკვეთა შექმნილია კომპანიის კაბინეტიდან" }],
    };
    const ok = await putShared("parcel:" + id, parcel);
    if (ok) { setParcels((p) => [parcel, ...p]); }
    return ok ? parcel : null;
  }

  async function saveParcel(updated) {
    const ok = await putShared("parcel:" + updated.id, updated);
    if (ok) setParcels((p) => p.map((x) => (x.id === updated.id ? updated : x)));
    else showToast("ცვლილება ვერ შეინახა — სცადეთ ხელახლა", "error");
    return ok;
  }

  async function transition(parcel, status, note, extra = {}) {
    const updated = { ...parcel, ...extra, status, history: [...parcel.history, { status, at: nowIso(), note: note || "" }] };
    const ok = await saveParcel(updated);
    if (ok) showToast(`სტატუსი განახლდა: ${STATUS[status].label}`);
    return ok ? updated : null;
  }

  async function markDispatched(parcel) {
    const updated = { ...parcel, dispatchedAt: nowIso(), history: [...parcel.history, { status: parcel.status, at: nowIso(), note: "კომპანიამ მონიშნა, როგორც გაგზავნილი საწყობისკენ" }] };
    const ok = await saveParcel(updated);
    if (ok) showToast("მონიშნულია, როგორც გაგზავნილი საწყობისკენ");
  }

  async function registerParcel(trackingId, courierId) {
    const parcel = parcels.find((p) => p.id === trackingId.trim().toUpperCase());
    if (!parcel) return { ok: false, reason: "notfound" };
    if (parcel.status !== "created") return { ok: false, reason: "already", parcel };
    const courier = couriers.find((c) => c.id === courierId);
    const extra = courierId ? { courierId, courierName: courier?.name || null } : {};
    const status = courierId ? "assigned" : "registered";
    const note = courierId ? `დარეგისტრირდა და მიენიჭა კურიერს: ${courier?.name}` : "დარეგისტრირდა საწყობში";
    const updated = await transition(parcel, status, note, extra);
    return { ok: !!updated, parcel: updated || parcel };
  }

  async function assignCourier(parcel, courierId) {
    const courier = couriers.find((c) => c.id === courierId);
    if (!courier) return;
    const status = parcel.status === "created" || parcel.status === "registered" ? "assigned" : parcel.status;
    await transition(parcel, status, `მიმაგრდა კურიერი: ${courier.name}`, { courierId, courierName: courier.name });
  }

  async function cancelParcel(parcel) {
    await transition(parcel, "cancelled", "შეკვეთა გააუქმა კომპანიამ");
  }

  async function bulkImport(rows) {
    let count = 0;
    for (const r of rows) {
      const parcel = {
        id: genTrackingId(),
        companyId: identity.id, companyName: identity.name,
        sender: { name: identity.name, phone: "", address: "", city: "" },
        receiver: { name: r.receiverName, phone: r.receiverPhone, address: r.receiverAddress, city: r.city },
        pieces: 1, weight: r.weight, dimensions: "", contents: r.contents, declaredValue: r.declaredValue,
        deliveryType: "სტანდარტული", paymentMethod: r.paymentMethod || PAYMENT_METHODS[1],
        codAmount: Number(r.codAmount) || 0,
        deliveryFee: r.deliveryFee === "" || r.deliveryFee == null ? DEFAULT_DELIVERY_FEE : Number(r.deliveryFee),
        codCollected: false, codCollectedAmount: 0, codCollectedAt: null, deliveredAt: null, dispatchedAt: null,
        preferredDate: null, notes: r.notes || "",
        status: "created", courierId: null, courierName: null, createdAt: nowIso(),
        history: [{ status: "created", at: nowIso(), note: "შეკვეთა შექმნილია მასობრივი ატვირთვით" }],
      };
      // eslint-disable-next-line no-await-in-loop
      const ok = await putShared("parcel:" + parcel.id, parcel);
      if (ok) { count++; setParcels((p) => [parcel, ...p]); }
    }
    return count;
  }

  /* -------------------------------- render -------------------------------- */

  if (booting) {
    return (
      <div className="kursi" style={{ minHeight: 480, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <GlobalStyle />
        <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>იტვირთება...</div>
      </div>
    );
  }

  if (!identity) {
    return (
      <div className="kursi">
        <GlobalStyle />
        <LoginScreen
          companies={companies} couriers={couriers} registrators={registrators} admins={admins}
          onLogin={tryLogin} onChoose={chooseIdentity}
          onAddCompany={addCompany} onAddCourier={addCourier} onAddRegistrator={addRegistrator} onAddAdmin={addAdmin}
        />
      </div>
    );
  }

  return (
    <div className="kursi" style={{ minHeight: 560 }}>
      <GlobalStyle />
      <Toast toast={toast} />
      <Shell identity={identity} onLogout={logout}>
        {identity.role === "company" && (
          <CompanyView
            identity={identity} parcels={parcels.filter((p) => p.companyId === identity.id)}
            onCreate={createParcel} onCancel={cancelParcel} onBulkImport={bulkImport}
            onMarkDispatched={markDispatched} loading={loadingData}
          />
        )}
        {identity.role === "courier" && (
          <CourierView
            identity={identity} parcels={parcels.filter((p) => p.courierId === identity.id)} couriers={couriers}
            onTransition={transition} onUpdateCourier={updateCourier} loading={loadingData}
          />
        )}
        {identity.role === "registrator" && (
          <RegistratorView
            couriers={couriers} onRegister={registerParcel} recentParcels={parcels.slice(0, 12)}
            onAssignCourier={assignCourier} loading={loadingData}
          />
        )}
        {identity.role === "admin" && (
          <AdminView
            companies={companies} couriers={couriers} registrators={registrators} admins={admins} parcels={parcels}
            onAddCompany={addCompany} onAddCourier={addCourier} onUpdateCourier={updateCourier}
            onAddRegistrator={addRegistrator} onAddAdmin={addAdmin}
            onAssignCourier={assignCourier} onTransition={transition} loading={loadingData}
            onRefresh={refreshAll}
            onClearCompanies={clearAllCompanies} onClearCouriers={clearAllCouriers} onClearParcels={clearAllParcels}
            onClearRegistrators={clearAllRegistrators} onClearAdmins={clearAllAdmins}
          />
        )}
      </Shell>
    </div>
  );
}

/* ============================== LOGIN SCREEN ============================== */

function allUsernames(list) { return list.map((a) => (a.username || "").trim().toLowerCase()).filter(Boolean); }

function LoginForm({ role, onLogin }) {
  const [username, setUsername] = useState("");
  const [pin, setPin] = useState("");
  const [err, setErr] = useState("");
  function submit() {
    if (!username.trim() || !pin) { setErr("შეავსეთ მომხმარებლის სახელი და კოდი"); return; }
    const res = onLogin(role, username, pin);
    if (!res.ok) setErr(res.reason === "badpin" ? "კოდი არასწორია" : "ასეთი მომხმარებელი ვერ მოიძებნა");
    else setErr("");
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <Field label="მომხმარებლის სახელი">
        <input className="input" value={username} onChange={(e) => setUsername(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
      </Field>
      <Field label="კოდი (PIN / პაროლი)">
        <input className="input" type="password" value={pin} onChange={(e) => setPin(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} />
      </Field>
      {err && <div style={{ color: "#B23A24", fontSize: 12.5 }}>{err}</div>}
      <button className="btn btn-primary" style={{ justifyContent: "center" }} onClick={submit}><Lock size={14} /> შესვლა</button>
    </div>
  );
}

function LoginScreen({ companies, couriers, registrators, admins, onLogin, onChoose, onAddCompany, onAddCourier, onAddRegistrator, onAddAdmin }) {
  const [tab, setTab] = useState("company");
  const [showReg, setShowReg] = useState(false);

  const roleTabs = [
    { key: "company", label: "კომპანია", icon: Building2 },
    { key: "courier", label: "კურიერი", icon: Truck },
    { key: "registrator", label: "რეგისტრატორი", icon: ScanLine },
    { key: "admin", label: "ადმინისტრატორი", icon: ShieldCheck },
  ];

  function switchTab(k) { setTab(k); setShowReg(false); }

  return (
    <div style={{ minHeight: 560, display: "flex", alignItems: "center", justifyContent: "center", padding: "32px 16px" }}>
      <div style={{ width: "100%", maxWidth: 440 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 22, justifyContent: "center" }}>
          <div style={{ width: 34, height: 34, borderRadius: 8, background: "var(--ink)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Truck size={18} color="#fff" />
          </div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 18, lineHeight: 1 }}>კურსი</div>
            <div style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>საკურიერო მართვის სისტემა</div>
          </div>
        </div>

        <div className="panel" style={{ padding: 6, display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 4, marginBottom: 16 }}>
          {roleTabs.map((t) => (
            <button key={t.key} className="btn" onClick={() => switchTab(t.key)}
              style={{ flexDirection: "column", gap: 4, padding: "10px 4px", background: tab === t.key ? "var(--ink)" : "transparent", color: tab === t.key ? "#fff" : "var(--ink-soft)" }}>
              <t.icon size={16} /><span style={{ fontSize: 11.5 }}>{t.label}</span>
            </button>
          ))}
        </div>

        <div className="panel" style={{ padding: 18 }}>
          {tab === "company" && (
            <div>
              <div style={{ fontWeight: 600, marginBottom: 10 }}>კომპანიის ავტორიზაცია</div>
              <LoginForm role="company" onLogin={onLogin} />
              <div style={{ height: 1, background: "var(--border)", margin: "16px 0" }} />
              {!showReg ? (
                <button className="btn btn-ghost" style={{ width: "100%", justifyContent: "center" }} onClick={() => setShowReg(true)}><Plus size={15} /> ახალი კომპანიის რეგისტრაცია</button>
              ) : (
                <NewCompanyForm existingUsernames={allUsernames(companies)} onCancel={() => setShowReg(false)}
                  onSubmit={async (data) => { const c = await onAddCompany(data); if (c) onChoose({ role: "company", id: c.id, name: c.name }); }} />
              )}
            </div>
          )}
          {tab === "courier" && (
            <div>
              <div style={{ fontWeight: 600, marginBottom: 10 }}>კურიერის ავტორიზაცია</div>
              <LoginForm role="courier" onLogin={onLogin} />
              <div style={{ height: 1, background: "var(--border)", margin: "16px 0" }} />
              {!showReg ? (
                <button className="btn btn-ghost" style={{ width: "100%", justifyContent: "center" }} onClick={() => setShowReg(true)}><Plus size={15} /> ახალი კურიერის რეგისტრაცია</button>
              ) : (
                <NewCourierForm existingUsernames={allUsernames(couriers)} onCancel={() => setShowReg(false)}
                  onSubmit={async (data) => { const c = await onAddCourier(data); if (c) onChoose({ role: "courier", id: c.id, name: c.name }); }} />
              )}
            </div>
          )}
          {tab === "registrator" && (
            <div>
              <div style={{ fontWeight: 600, marginBottom: 10 }}>რეგისტრატორის ავტორიზაცია</div>
              <LoginForm role="registrator" onLogin={onLogin} />
              <div style={{ height: 1, background: "var(--border)", margin: "16px 0" }} />
              {!showReg ? (
                <button className="btn btn-ghost" style={{ width: "100%", justifyContent: "center" }} onClick={() => setShowReg(true)}><Plus size={15} /> ახალი რეგისტრატორის რეგისტრაცია</button>
              ) : (
                <NewStaffForm existingUsernames={allUsernames(registrators)} onCancel={() => setShowReg(false)}
                  onSubmit={async (data) => { const a = await onAddRegistrator(data); if (a) onChoose({ role: "registrator", id: a.id, name: a.name }); }} />
              )}
            </div>
          )}
          {tab === "admin" && (
            <div>
              <div style={{ fontWeight: 600, marginBottom: 10 }}>ადმინისტრატორის ავტორიზაცია</div>
              <LoginForm role="admin" onLogin={onLogin} />
              <div style={{ height: 1, background: "var(--border)", margin: "16px 0" }} />
              {!showReg ? (
                <button className="btn btn-ghost" style={{ width: "100%", justifyContent: "center" }} onClick={() => setShowReg(true)}><Plus size={15} /> ახალი ადმინისტრატორის რეგისტრაცია</button>
              ) : (
                <NewStaffForm existingUsernames={allUsernames(admins)} onCancel={() => setShowReg(false)}
                  onSubmit={async (data) => { const a = await onAddAdmin(data); if (a) onChoose({ role: "admin", id: a.id, name: a.name }); }} />
              )}
            </div>
          )}
        </div>
        <div style={{ fontSize: 12, color: "var(--ink-soft)", textAlign: "center", marginTop: 14, lineHeight: 1.5 }}>
          <User size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />
          თითოეულ ანგარიშს აქვს საკუთარი მომხმარებლის სახელი და კოდი — ეს იცავს შემთხვევით სხვისი ანგარიშის გახსნისგან, თუმცა არ არის კრიპტოგრაფიულად დაცული (რეალურ წარმოებაში საჭირო იქნება ნამდვილი სერვერული ავტორიზაცია).
        </div>
      </div>
    </div>
  );
}

function NewCompanyForm({ onSubmit, onCancel, existingUsernames = [] }) {
  const [f, setF] = useState({ name: "", contactPerson: "", phone: "", address: "", taxId: "", username: "", pin: "" });
  const [err, setErr] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  function submit() {
    if (!f.name.trim()) { setErr("შეიყვანეთ კომპანიის დასახელება"); return; }
    if (!f.username.trim() || !f.pin) { setErr("შეიყვანეთ მომხმარებლის სახელი და კოდი"); return; }
    if (existingUsernames.includes(f.username.trim().toLowerCase())) { setErr("ეს მომხმარებლის სახელი უკვე დაკავებულია"); return; }
    onSubmit(f);
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <Field label="კომპანიის დასახელება *"><input className="input" value={f.name} onChange={set("name")} placeholder="შპს მაგალითი" /></Field>
      <Field label="საკონტაქტო პირი"><input className="input" value={f.contactPerson} onChange={set("contactPerson")} /></Field>
      <Field label="ტელეფონი"><input className="input" value={f.phone} onChange={set("phone")} placeholder="+995 5__ __ __ __" /></Field>
      <Field label="მისამართი"><input className="input" value={f.address} onChange={set("address")} /></Field>
      <Field label="საიდენტიფიკაციო კოდი"><input className="input" value={f.taxId} onChange={set("taxId")} /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <Field label="მომხმარებლის სახელი *"><input className="input" value={f.username} onChange={set("username")} /></Field>
        <Field label="კოდი / PIN *"><input className="input" type="password" value={f.pin} onChange={set("pin")} /></Field>
      </div>
      {err && <div style={{ color: "#B23A24", fontSize: 12.5 }}>{err}</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
        <button className="btn btn-outline" style={{ flex: 1, justifyContent: "center" }} onClick={onCancel}>გაუქმება</button>
        <button className="btn btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={submit}>რეგისტრაცია</button>
      </div>
    </div>
  );
}
function NewCourierForm({ onSubmit, onCancel, existingUsernames = [] }) {
  const [f, setF] = useState({ name: "", phone: "", vehicle: "", zone: "", commissionType: "percent", commissionValue: 60, username: "", pin: "" });
  const [err, setErr] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  function submit() {
    if (!f.name.trim()) { setErr("შეიყვანეთ სახელი"); return; }
    if (!f.username.trim() || !f.pin) { setErr("შეიყვანეთ მომხმარებლის სახელი და კოდი"); return; }
    if (existingUsernames.includes(f.username.trim().toLowerCase())) { setErr("ეს მომხმარებლის სახელი უკვე დაკავებულია"); return; }
    onSubmit(f);
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <Field label="სახელი და გვარი *"><input className="input" value={f.name} onChange={set("name")} /></Field>
      <Field label="ტელეფონი"><input className="input" value={f.phone} onChange={set("phone")} placeholder="+995 5__ __ __ __" /></Field>
      <Field label="ტრანსპორტი"><input className="input" value={f.vehicle} onChange={set("vehicle")} placeholder="მოტოციკლი / ავტომობილი" /></Field>
      <Field label="ზონა / რაიონი"><input className="input" value={f.zone} onChange={set("zone")} /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <Field label="საკომისიო ტიპი">
          <select className="select" value={f.commissionType} onChange={set("commissionType")}>
            <option value="percent">% საკურიერო საფასურიდან</option>
            <option value="fixed">ფიქსირებული, ₾ / ამანათი</option>
          </select>
        </Field>
        <Field label={f.commissionType === "percent" ? "პროცენტი (%)" : "თანხა (₾)"}>
          <input className="input" value={f.commissionValue} onChange={set("commissionValue")} />
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <Field label="მომხმარებლის სახელი *"><input className="input" value={f.username} onChange={set("username")} /></Field>
        <Field label="კოდი / PIN *"><input className="input" type="password" value={f.pin} onChange={set("pin")} /></Field>
      </div>
      {err && <div style={{ color: "#B23A24", fontSize: 12.5 }}>{err}</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
        <button className="btn btn-outline" style={{ flex: 1, justifyContent: "center" }} onClick={onCancel}>გაუქმება</button>
        <button className="btn btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={submit}>რეგისტრაცია</button>
      </div>
    </div>
  );
}
function NewStaffForm({ onSubmit, onCancel, existingUsernames = [] }) {
  const [f, setF] = useState({ name: "", username: "", pin: "" });
  const [err, setErr] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  function submit() {
    if (!f.name.trim()) { setErr("შეიყვანეთ სახელი"); return; }
    if (!f.username.trim() || !f.pin) { setErr("შეიყვანეთ მომხმარებლის სახელი და კოდი"); return; }
    if (existingUsernames.includes(f.username.trim().toLowerCase())) { setErr("ეს მომხმარებლის სახელი უკვე დაკავებულია"); return; }
    onSubmit(f);
  }
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      <Field label="სახელი და გვარი *"><input className="input" value={f.name} onChange={set("name")} /></Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <Field label="მომხმარებლის სახელი *"><input className="input" value={f.username} onChange={set("username")} /></Field>
        <Field label="კოდი / PIN *"><input className="input" type="password" value={f.pin} onChange={set("pin")} /></Field>
      </div>
      {err && <div style={{ color: "#B23A24", fontSize: 12.5 }}>{err}</div>}
      <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
        <button className="btn btn-outline" style={{ flex: 1, justifyContent: "center" }} onClick={onCancel}>გაუქმება</button>
        <button className="btn btn-primary" style={{ flex: 1, justifyContent: "center" }} onClick={submit}>რეგისტრაცია</button>
      </div>
    </div>
  );
}

/* ================================== SHELL ================================= */

function Shell({ identity, onLogout, children }) {
  const roleMeta = {
    company: { icon: Building2, label: "კომპანიის კაბინეტი" },
    courier: { icon: Truck, label: "კურიერის კაბინეტი" },
    registrator: { icon: ScanLine, label: "რეგისტრატორის სამუშაო მაგიდა" },
    admin: { icon: ShieldCheck, label: "ადმინისტრატორის პანელი" },
  }[identity.role];

  return (
    <div>
      <div style={{ borderBottom: "1px solid var(--border)", background: "#fff" }}>
        <div style={{ maxWidth: 1180, margin: "0 auto", padding: "12px 18px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 30, height: 30, borderRadius: 7, background: "var(--ink)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <Truck size={16} color="#fff" />
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 15, lineHeight: 1.1 }}>კურსი</div>
              <div style={{ fontSize: 11.5, color: "var(--ink-soft)", display: "flex", alignItems: "center", gap: 4 }}>
                <roleMeta.icon size={12} />{roleMeta.label}
              </div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 13.5, fontWeight: 500 }}>{identity.name}</span>
            <button className="btn btn-ghost btn-sm" onClick={onLogout}><LogOut size={14} /> გამოსვლა</button>
          </div>
        </div>
      </div>
      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "20px 18px 40px" }}>{children}</div>
    </div>
  );
}

/* ================================ TABS UI ================================= */

function TabBar({ tabs, active, onChange }) {
  return (
    <div style={{ display: "flex", gap: 4, marginBottom: 18, borderBottom: "1px solid var(--border)", overflowX: "auto" }}>
      {tabs.map((t) => (
        <button key={t.key} onClick={() => onChange(t.key)}
          style={{
            padding: "9px 14px", fontSize: 13.5, fontWeight: 500, background: "transparent", border: "none",
            borderBottom: active === t.key ? "2px solid var(--accent)" : "2px solid transparent",
            color: active === t.key ? "var(--ink)" : "var(--ink-soft)", cursor: "pointer", whiteSpace: "nowrap",
            display: "flex", alignItems: "center", gap: 6,
          }}>
          <t.icon size={14} />{t.label}
        </button>
      ))}
    </div>
  );
}

function StatRow({ items }) {
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 18 }}>
      {items.map((it, i) => (
        <div key={i} className="stat" style={{ flex: "1 1 140px", minWidth: 130 }}>
          <div style={{ fontSize: 20, fontWeight: 700 }}>{it.value}</div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>{it.label}</div>
        </div>
      ))}
    </div>
  );
}

const PAGE_SIZE = 20;
/** დიდი სიების გვერდებად დაყოფა — რენდერს მხოლოდ მიმდინარე გვერდის
    ჩანაწერებს უზრუნველყოფს, რაც სისწრაფეს ინარჩუნებს ბევრი ჩანაწერის დროსაც. */
function Pagination({ page, setPage, total, pageSize = PAGE_SIZE }) {
  const pageCount = Math.max(1, Math.ceil(total / pageSize));
  if (pageCount <= 1) return null;
  const from = (page - 1) * pageSize + 1;
  const to = Math.min(total, page * pageSize);
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
      <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>{from}–{to} / {total}</span>
      <div style={{ display: "flex", gap: 6 }}>
        <button className="btn btn-outline btn-sm" disabled={page <= 1} onClick={() => setPage((p) => Math.max(1, p - 1))}>← წინა</button>
        <span style={{ fontSize: 12.5, color: "var(--ink-soft)", alignSelf: "center" }}>{page} / {pageCount}</span>
        <button className="btn btn-outline btn-sm" disabled={page >= pageCount} onClick={() => setPage((p) => Math.min(pageCount, p + 1))}>შემდეგი →</button>
      </div>
    </div>
  );
}

/** ორსაფეხურიანი დადასტურება დესტრუქციული, საზიარო მოქმედებისთვის —
    შემთხვევით დაჭერა ვერავის წაშლის მონაცემებს. */
function DangerZone({ label, onConfirm, disabled }) {
  const [confirming, setConfirming] = useState(false);
  const [busy, setBusy] = useState(false);
  if (!confirming) {
    return (
      <button className="btn btn-outline btn-sm" disabled={disabled} style={{ borderColor: "#B23A24", color: "#B23A24" }} onClick={() => setConfirming(true)}>
        <AlertTriangle size={13} /> {label}
      </button>
    );
  }
  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
      <span style={{ fontSize: 12.5, color: "#B23A24" }}>დარწმუნებული ხართ? ეს შეუქცევადია და ეხება ყველას, ვინც ამ სისტემას იყენებს.</span>
      <button className="btn btn-danger btn-sm" disabled={busy} onClick={async () => { setBusy(true); await onConfirm(); setBusy(false); setConfirming(false); }}>{busy ? "იშლება..." : "დიახ, წავშალო"}</button>
      <button className="btn btn-outline btn-sm" disabled={busy} onClick={() => setConfirming(false)}>გაუქმება</button>
    </div>
  );
}

/* =============================== COMPANY VIEW ============================== */

function CompanyView({ identity, parcels, onCreate, onCancel, onBulkImport, onMarkDispatched, loading }) {
  const [tab, setTab] = useState("orders");
  const [statusFilter, setStatusFilter] = useState("all");
  const [q, setQ] = useState("");
  const [printing, setPrinting] = useState(null);
  const [page, setPage] = useState(1);

  const filtered = useMemo(() => parcels.filter((p) => {
    if (statusFilter !== "all" && p.status !== statusFilter) return false;
    if (q && !(p.id.toLowerCase().includes(q.toLowerCase()) || p.receiver.name.toLowerCase().includes(q.toLowerCase()))) return false;
    return true;
  }), [parcels, statusFilter, q]);
  useEffect(() => { setPage(1); }, [statusFilter, q]);
  const pageItems = useMemo(() => filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE), [filtered, page]);

  const counts = useMemo(() => {
    const c = { total: parcels.length, active: 0, delivered: 0, cod: 0, unsent: 0 };
    parcels.forEach((p) => {
      if (["created", "registered", "assigned", "transit"].includes(p.status)) c.active++;
      if (p.status === "delivered") c.delivered++;
      if (Number(p.codAmount) > 0 && !p.codCollected) c.cod += Number(p.codAmount);
      if (p.status === "created" && !p.dispatchedAt) c.unsent++;
    });
    return c;
  }, [parcels]);

  return (
    <div>
      <StatRow items={[
        { label: "სულ შეკვეთა", value: counts.total },
        { label: "მიმდინარე", value: counts.active },
        { label: "მიწოდებული", value: counts.delivered },
        { label: "მოსალოდნელი COD", value: fmtMoney(counts.cod) },
        { label: "საწყობში გაუგზავნელი", value: counts.unsent },
      ]} />
      <TabBar active={tab} onChange={setTab} tabs={[
        { key: "orders", label: "ჩემი შეკვეთები", icon: ClipboardList },
        { key: "new", label: "ახალი შეკვეთა", icon: Plus },
        { key: "bulk", label: "მასობრივი ატვირთვა", icon: Upload },
      ]} />

      {tab === "orders" && (
        <div>
          <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
            <div style={{ position: "relative", flex: "1 1 220px" }}>
              <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: "var(--ink-soft)" }} />
              <input className="input" style={{ paddingLeft: 30 }} placeholder="ძებნა კოდით ან მიმღებით" value={q} onChange={(e) => setQ(e.target.value)} />
            </div>
            <select className="select" style={{ maxWidth: 220 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
              <option value="all">ყველა სტატუსი</option>
              {Object.entries(STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
          </div>
          {loading ? <div style={{ color: "var(--ink-soft)", fontSize: 13.5 }}>იტვირთება...</div> :
            filtered.length === 0 ? (
              <EmptyState icon={Package} title="შეკვეთები ვერ მოიძებნა" hint="ჯერ არ გაქვთ ამ ფილტრით შეკვეთა — შექმენით პირველი."
                action={<button className="btn btn-primary" onClick={() => setTab("new")}><Plus size={14} /> ახალი შეკვეთა</button>} />
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {pageItems.map((p) => (
                  <Waybill key={p.id} parcel={p}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
                      <HistoryTimeline history={p.history} />
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                        <button className="btn btn-outline btn-sm" onClick={() => setPrinting(p)}><Printer size={13} /> ბეჭდვა</button>
                        {p.status === "created" && !p.dispatchedAt && (
                          <button className="btn btn-track btn-sm" onClick={() => onMarkDispatched(p)}><CheckCircle2 size={13} /> გავგზავნე საწყობში</button>
                        )}
                        {p.status === "created" && (
                          <button className="btn btn-outline btn-sm" onClick={() => onCancel(p)}><X size={13} /> გაუქმება</button>
                        )}
                      </div>
                    </div>
                  </Waybill>
                ))}
              </div>
            )}
          <Pagination page={page} setPage={setPage} total={filtered.length} />
        </div>
      )}

      {tab === "new" && <NewOrderForm identity={identity} onCreate={onCreate} onDone={() => setTab("orders")} />}
      {tab === "bulk" && <BulkImportPanel onImport={onBulkImport} />}

      {printing && <PrintLabelModal parcel={printing} onClose={() => setPrinting(null)} />}
    </div>
  );
}

function NewOrderForm({ onCreate, onDone }) {
  const blank = {
    sender: { name: "", phone: "", address: "", city: "" },
    receiver: { name: "", phone: "", address: "", city: "" },
    pieces: 1, weight: "", dimensions: "", contents: "", declaredValue: "",
    deliveryType: DELIVERY_TYPES[0], paymentMethod: PAYMENT_METHODS[0], codAmount: "",
    deliveryFee: DEFAULT_DELIVERY_FEE, preferredDate: "", notes: "",
  };
  const [f, setF] = useState(blank);
  const [err, setErr] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const setSender = (k) => (e) => setF({ ...f, sender: { ...f.sender, [k]: e.target.value } });
  const setReceiver = (k) => (e) => setF({ ...f, receiver: { ...f.receiver, [k]: e.target.value } });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  async function submit() {
    if (!f.sender.name || !f.sender.phone || !f.sender.address) { setErr("შეავსეთ გამგზავნის სახელი, ტელეფონი და მისამართი"); return; }
    if (!f.receiver.name || !f.receiver.phone || !f.receiver.address) { setErr("შეავსეთ მიმღების სახელი, ტელეფონი და მისამართი"); return; }
    if (f.paymentMethod === PAYMENT_METHODS[0] && !f.codAmount) { setErr("მიუთითეთ ასაკრები თანხა (COD)"); return; }
    setErr(""); setSubmitting(true);
    const p = await onCreate(f);
    setSubmitting(false);
    if (p) { setF(blank); onDone(); }
    else setErr("შეკვეთის შენახვა ვერ მოხერხდა — სცადეთ ხელახლა");
  }

  return (
    <div className="panel" style={{ padding: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }} className="order-grid">
        <div>
          <div style={{ fontWeight: 600, marginBottom: 10 }}>გამგზავნი</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <Field label="სახელი *"><input className="input" value={f.sender.name} onChange={setSender("name")} /></Field>
            <Field label="ტელეფონი *"><input className="input" value={f.sender.phone} onChange={setSender("phone")} /></Field>
            <Field label="მისამართი *"><input className="input" value={f.sender.address} onChange={setSender("address")} /></Field>
            <Field label="ქალაქი"><input className="input" value={f.sender.city} onChange={setSender("city")} /></Field>
          </div>
        </div>
        <div>
          <div style={{ fontWeight: 600, marginBottom: 10 }}>მიმღები</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <Field label="სახელი *"><input className="input" value={f.receiver.name} onChange={setReceiver("name")} /></Field>
            <Field label="ტელეფონი *"><input className="input" value={f.receiver.phone} onChange={setReceiver("phone")} /></Field>
            <Field label="მისამართი *"><input className="input" value={f.receiver.address} onChange={setReceiver("address")} /></Field>
            <Field label="ქალაქი"><input className="input" value={f.receiver.city} onChange={setReceiver("city")} /></Field>
          </div>
        </div>
      </div>

      <div style={{ height: 1, background: "var(--border)", margin: "18px 0" }} />

      <div style={{ fontWeight: 600, marginBottom: 10 }}>ამანათის დეტალები</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10 }} className="details-grid">
        <Field label="ცალები"><input className="input" type="number" min="1" value={f.pieces} onChange={set("pieces")} /></Field>
        <Field label="წონა (კგ)"><input className="input" value={f.weight} onChange={set("weight")} placeholder="1.5" /></Field>
        <Field label="გაბარიტები"><input className="input" value={f.dimensions} onChange={set("dimensions")} placeholder="30x20x15" /></Field>
        <Field label="ღირებულება (₾)"><input className="input" value={f.declaredValue} onChange={set("declaredValue")} /></Field>
      </div>
      <div style={{ marginTop: 8 }}><Field label="შემცველობის აღწერა"><input className="input" value={f.contents} onChange={set("contents")} placeholder="მაგ. ტანსაცმელი, დოკუმენტები..." /></Field></div>

      <div style={{ height: 1, background: "var(--border)", margin: "18px 0" }} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }} className="details-grid">
        <Field label="მიწოდების ტიპი">
          <select className="select" value={f.deliveryType} onChange={set("deliveryType")}>
            {DELIVERY_TYPES.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="გადახდის მეთოდი">
          <select className="select" value={f.paymentMethod} onChange={set("paymentMethod")}>
            {PAYMENT_METHODS.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="საკურიერო მომსახურების საფასური (₾)"><input className="input" value={f.deliveryFee} onChange={set("deliveryFee")} /></Field>
        {f.paymentMethod === PAYMENT_METHODS[0] && (
          <Field label="ასაკრები თანხა — COD (₾) *"><input className="input" value={f.codAmount} onChange={set("codAmount")} /></Field>
        )}
      </div>
      {f.paymentMethod === PAYMENT_METHODS[0] && Number(f.codAmount) > 0 && (
        <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: 6 }}>
          კომპანიას ჩაერიცხება: <b style={{ color: "var(--ink)" }}>{fmtMoney((Number(f.codAmount) || 0) - (Number(f.deliveryFee) || 0))}</b> (COD − საკურიერო საფასური)
        </div>
      )}
      <div style={{ marginTop: 8 }}><Field label="შენიშვნა"><textarea className="textarea" value={f.notes} onChange={set("notes")} /></Field></div>

      {err && <div style={{ color: "#B23A24", fontSize: 13, marginTop: 12 }}>{err}</div>}
      <div style={{ marginTop: 16 }}>
        <button className="btn btn-primary" disabled={submitting} onClick={submit}>
          {submitting ? "იგზავნება..." : <><Plus size={15} /> შეკვეთის შექმნა</>}
        </button>
      </div>
      <style>{`@media (max-width:700px){ .order-grid{grid-template-columns:1fr !important;} .details-grid{grid-template-columns:1fr 1fr !important;} }`}</style>
    </div>
  );
}

const IMPORT_TEMPLATE_HEADERS = ["მიმღების_სახელი", "ტელეფონი", "მისამართი", "ქალაქი", "შემცველობა", "წონა", "ღირებულება", "გადახდის_მეთოდი", "cod_თანხა", "საკურიერო_საფასური", "შენიშვნა"];

function BulkImportPanel({ onImport }) {
  const [rows, setRows] = useState([]);
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(null);
  const fileRef = useRef(null);

  function rowsFromArrays(arr) {
    const t = (v) => (v == null ? "" : String(v)).trim();
    return arr.map((cols) => ({
      receiverName: t(cols[0]), receiverPhone: t(cols[1]), receiverAddress: t(cols[2]),
      city: t(cols[3]), contents: t(cols[4]), weight: t(cols[5]), declaredValue: t(cols[6]),
      paymentMethod: t(cols[7]) || PAYMENT_METHODS[1], codAmount: t(cols[8]), deliveryFee: t(cols[9]), notes: t(cols[10]),
    })).filter((r) => r.receiverName);
  }
  function parsePastedText(text) {
    const res = Papa.parse(text.trim(), { skipEmptyLines: true });
    setRows(rowsFromArrays(res.data)); setDone(null);
  }
  function onFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    const reader = new FileReader();
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      reader.onload = (ev) => {
        try {
          const wb = XLSX.read(ev.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const arr = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
          setRows(rowsFromArrays(arr.slice(1))); setDone(null); // slice(1): ჰედერის ხაზი გამოტოვებულია
        } catch (err) { console.error(err); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = () => {
        const res = Papa.parse(String(reader.result).trim(), { skipEmptyLines: true });
        setRows(rowsFromArrays(res.data.slice(1))); setDone(null); // CSV ფაილიც ჰედერით მოდის
      };
      reader.readAsText(file);
    }
    if (fileRef.current) fileRef.current.value = "";
  }
  async function runImport() {
    setBusy(true);
    const count = await onImport(rows);
    setBusy(false); setDone(count); setRows([]);
  }
  function downloadTemplate() {
    const example = { [IMPORT_TEMPLATE_HEADERS[0]]: "ნინო ბერიძე", [IMPORT_TEMPLATE_HEADERS[1]]: "599112233", [IMPORT_TEMPLATE_HEADERS[2]]: "ვაჟა-ფშაველას 12", [IMPORT_TEMPLATE_HEADERS[3]]: "თბილისი", [IMPORT_TEMPLATE_HEADERS[4]]: "დოკუმენტები", [IMPORT_TEMPLATE_HEADERS[5]]: "0.5", [IMPORT_TEMPLATE_HEADERS[6]]: "50", [IMPORT_TEMPLATE_HEADERS[7]]: PAYMENT_METHODS[0], [IMPORT_TEMPLATE_HEADERS[8]]: "55", [IMPORT_TEMPLATE_HEADERS[9]]: "5", [IMPORT_TEMPLATE_HEADERS[10]]: "სასწრაფოდ" };
    downloadXlsx("kursi-shabloni.xlsx", [example], "შეკვეთები");
  }

  return (
    <div className="panel" style={{ padding: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 6 }}>
        <div style={{ fontWeight: 600 }}>Excel / CSV ფაილით მასობრივი ატვირთვა</div>
        <button className="btn btn-outline btn-sm" onClick={downloadTemplate}><FileSpreadsheet size={13} /> Excel შაბლონის გადმოწერა</button>
      </div>
      <div style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12, lineHeight: 1.6 }}>
        ატვირთეთ .xlsx ან .csv ფაილი (პირველი ხაზი — სათაურები, გადმოწერეთ შაბლონი ზემოთ). სვეტების თანმიმდევრობა: {IMPORT_TEMPLATE_HEADERS.join(", ")}
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,text/csv" onChange={onFile} style={{ fontSize: 13 }} />
        <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>ან ჩასვით ტექსტი ქვემოთ (სათაურის გარეშე, თითო შეკვეთა ახალ ხაზზე)</span>
      </div>
      <textarea className="textarea" style={{ marginTop: 10, minHeight: 100 }}
        placeholder="ნინო ბერიძე, 599112233, ვაჟა-ფშაველას 12, თბილისი, დოკუმენტები, 0.5, 50, ქეშით კურიერთან (COD), 55, 5, სასწრაფოდ"
        onChange={(e) => parsePastedText(e.target.value)} />
      {rows.length > 0 && (
        <div style={{ marginTop: 14 }}>
          <div style={{ fontSize: 13, marginBottom: 8 }}>მოსატანი: <b>{rows.length}</b> შეკვეთა</div>
          <div style={{ overflowX: "auto", maxHeight: 220, border: "1px solid var(--border)", borderRadius: 6 }}>
            <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th>მიმღები</th><th>მისამართი</th><th>ტელეფონი</th><th>COD</th><th>საკურიერო</th></tr></thead>
              <tbody>{rows.map((r, i) => <tr key={i}><td>{r.receiverName}</td><td>{r.receiverAddress}{r.city ? `, ${r.city}` : ""}</td><td>{r.receiverPhone}</td><td>{r.codAmount || "—"}</td><td>{r.deliveryFee || DEFAULT_DELIVERY_FEE}</td></tr>)}</tbody>
            </table>
          </div>
          <button className="btn btn-primary" style={{ marginTop: 12 }} disabled={busy} onClick={runImport}>
            {busy ? "იტვირთება..." : <><Upload size={15} /> {rows.length} შეკვეთის დამატება</>}
          </button>
        </div>
      )}
      {done !== null && <div style={{ marginTop: 10, fontSize: 13.5, color: "#1E8A5F" }}><CheckCircle2 size={14} style={{ display: "inline", marginRight: 4 }} />{done} შეკვეთა წარმატებით აიტვირთა</div>}
    </div>
  );
}

/* =============================== COURIER VIEW =============================== */

function CourierView({ identity, parcels, couriers, onTransition, onUpdateCourier, loading }) {
  const [statusFilter, setStatusFilter] = useState("active");
  const [codModal, setCodModal] = useState(null);
  const [failModal, setFailModal] = useState(null);
  const [printing, setPrinting] = useState(null);
  const [page, setPage] = useState(1);
  const me = couriers.find((c) => c.id === identity.id);

  const active = parcels.filter((p) => ["assigned", "transit", "failed"].includes(p.status));
  const finished = parcels.filter((p) => ["delivered", "returned", "cancelled"].includes(p.status));
  const list = statusFilter === "active" ? active : statusFilter === "finished" ? finished : parcels;
  useEffect(() => { setPage(1); }, [statusFilter]);
  const pageItems = useMemo(() => list.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE), [list, page]);

  const counts = { assigned: parcels.filter((p) => p.status === "assigned").length, transit: parcels.filter((p) => p.status === "transit").length, delivered: parcels.filter((p) => p.status === "delivered").length };

  async function handleDeliver(p) {
    if (Number(p.codAmount) > 0) setCodModal(p);
    else onTransition(p, "delivered", "მიწოდებულია კურიერის მიერ", { deliveredAt: nowIso() });
  }

  return (
    <div>
      <StatRow items={[
        { label: "მისაღები", value: counts.assigned },
        { label: "გზაშია", value: counts.transit },
        { label: "დღეს მიწოდებული", value: counts.delivered },
      ]} />
      <div className="panel" style={{ padding: "12px 14px", marginBottom: 14, display: "flex", flexDirection: "column", gap: 8 }}>
        <LocationStatus courier={me} />
        <LocationUpdater courierId={identity.id} onUpdateCourier={onUpdateCourier} />
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
        {[["active", "აქტიური"], ["finished", "დასრულებული"], ["all", "ყველა"]].map(([k, l]) => (
          <button key={k} className="btn btn-sm" style={{ background: statusFilter === k ? "var(--ink)" : "transparent", color: statusFilter === k ? "#fff" : "var(--ink-soft)", border: "1px solid var(--border)" }} onClick={() => setStatusFilter(k)}>{l}</button>
        ))}
      </div>

      {loading ? <div style={{ color: "var(--ink-soft)", fontSize: 13.5 }}>იტვირთება...</div> :
        list.length === 0 ? (
          <EmptyState icon={Truck} title="ამანათები არ არის" hint="ამ ფილტრში ამჟამად ამანათი არ გიმაგრდებათ." />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {pageItems.map((p) => (
              <Waybill key={p.id} parcel={p}>
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  <div style={{ fontSize: 12.5, color: "var(--ink-soft)", display: "flex", gap: 14, flexWrap: "wrap" }}>
                    <span><Phone size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />{p.receiver.phone}</span>
                    <span><MapPin size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />{p.receiver.address}{p.receiver.city ? `, ${p.receiver.city}` : ""}</span>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button className="btn btn-outline btn-sm" onClick={() => setPrinting(p)}><Printer size={13} /> ბეჭდვა</button>
                    {p.status === "assigned" && (
                      <button className="btn btn-track btn-sm" onClick={() => onTransition(p, "transit", "წამოღებულია, გზაშია")}><Truck size={13} /> წამოღება — გზაზე</button>
                    )}
                    {p.status === "transit" && (
                      <>
                        <button className="btn btn-primary btn-sm" onClick={() => handleDeliver(p)}><CheckCircle2 size={13} /> მიწოდებულია</button>
                        <button className="btn btn-danger btn-sm" onClick={() => setFailModal(p)}><XCircle size={13} /> ვერ ჩაბარდა</button>
                      </>
                    )}
                    {p.status === "failed" && (
                      <>
                        <button className="btn btn-track btn-sm" onClick={() => onTransition(p, "transit", "განმეორებითი მცდელობა")}><RotateCcw size={13} /> ხელახლა მცდელობა</button>
                        <button className="btn btn-outline btn-sm" onClick={() => onTransition(p, "returned", "დაბრუნდა გამგზავნს")}>დაბრუნება გამგზავნს</button>
                      </>
                    )}
                  </div>
                  <HistoryTimeline history={p.history} />
                </div>
              </Waybill>
            ))}
          </div>
        )}
      <Pagination page={page} setPage={setPage} total={list.length} />

      {codModal && <CodModal parcel={codModal} onClose={() => setCodModal(null)}
        onConfirm={(amount) => { onTransition(codModal, "delivered", "მიწოდებულია, COD მიღებულია", { codCollected: true, codCollectedAmount: amount, codCollectedAt: nowIso(), deliveredAt: nowIso() }); setCodModal(null); }} />}
      {failModal && <FailModal parcel={failModal} onClose={() => setFailModal(null)}
        onConfirm={(reason) => { onTransition(failModal, "failed", reason); setFailModal(null); }} />}
      {printing && <PrintLabelModal parcel={printing} onClose={() => setPrinting(null)} />}
    </div>
  );
}

function ModalShell({ title, onClose, children }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(28,35,33,0.45)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 50 }} onClick={onClose}>
      <div className="panel" style={{ padding: 18, width: "100%", maxWidth: 380, background: "#fff" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontWeight: 600 }}>{title}</div>
          <button className="btn btn-ghost btn-sm" onClick={onClose}><X size={15} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}
function CodModal({ parcel, onClose, onConfirm }) {
  const [amount, setAmount] = useState(parcel.codAmount);
  return (
    <ModalShell title="ნაღდი ანგარიშსწორების დადასტურება" onClose={onClose}>
      <div style={{ fontSize: 13.5, color: "var(--ink-soft)", marginBottom: 10 }}>დაადასტურეთ მიმღებისგან მიღებული თანხა ამანათზე {parcel.id}</div>
      <Field label="მიღებული თანხა (₾)"><input className="input" value={amount} onChange={(e) => setAmount(e.target.value)} /></Field>
      <button className="btn btn-primary" style={{ width: "100%", justifyContent: "center", marginTop: 14 }} onClick={() => onConfirm(Number(amount) || 0)}>
        <Banknote size={14} /> დადასტურება და მიწოდების დახურვა
      </button>
    </ModalShell>
  );
}
function FailModal({ parcel, onClose, onConfirm }) {
  const [reason, setReason] = useState(FAIL_REASONS[0]);
  const [note, setNote] = useState("");
  return (
    <ModalShell title="მიწოდების მცდელობა ვერ განხორციელდა" onClose={onClose}>
      <Field label="მიზეზი">
        <select className="select" value={reason} onChange={(e) => setReason(e.target.value)}>
          {FAIL_REASONS.map((r) => <option key={r}>{r}</option>)}
        </select>
      </Field>
      <div style={{ marginTop: 8 }}><Field label="დამატებითი შენიშვნა"><input className="input" value={note} onChange={(e) => setNote(e.target.value)} /></Field></div>
      <button className="btn btn-danger" style={{ width: "100%", justifyContent: "center", marginTop: 14 }} onClick={() => onConfirm(note ? `${reason} — ${note}` : reason)}>
        დადასტურება
      </button>
    </ModalShell>
  );
}

/* ============================= REGISTRATOR VIEW ============================= */

function RegistratorView({ couriers, onRegister, recentParcels, onAssignCourier, loading }) {
  const [code, setCode] = useState("");
  const [log, setLog] = useState([]);
  const [pendingAssign, setPendingAssign] = useState(null);
  const inputRef = useRef(null);

  useEffect(() => { inputRef.current && inputRef.current.focus(); }, []);

  async function handleScan(courierId) {
    if (!code.trim()) return;
    const result = await onRegister(code, courierId || null);
    if (result.ok) {
      setLog((l) => [{ ok: true, id: result.parcel.id, msg: courierId ? "დარეგისტრირდა და მიენიჭა კურიერს" : "დარეგისტრირდა საწყობში", at: nowIso() }, ...l].slice(0, 20));
      setPendingAssign(courierId ? null : result.parcel);
    } else if (result.reason === "notfound") {
      setLog((l) => [{ ok: false, id: code.trim().toUpperCase(), msg: "ასეთი კოდით ამანათი ვერ მოიძებნა", at: nowIso() }, ...l].slice(0, 20));
    } else if (result.reason === "already") {
      setLog((l) => [{ ok: false, id: result.parcel.id, msg: `უკვე დარეგისტრირებულია (სტატუსი: ${STATUS[result.parcel.status].label})`, at: nowIso() }, ...l].slice(0, 20));
    }
    setCode("");
    inputRef.current && inputRef.current.focus();
  }

  return (
    <div>
      <div className="scanbox" style={{ padding: 22, marginBottom: 18, textAlign: "center" }}>
        <ScanLine size={26} color="var(--track)" style={{ marginBottom: 8 }} />
        <div style={{ fontWeight: 600, marginBottom: 4 }}>დაასკანერეთ ან შეიყვანეთ ამანათის კოდი</div>
        <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 12 }}>მუშაობს ბეჭდურ იარლიყზე დატანილ შტრიხკოდთან (Code 39) ნებისმიერი USB/Bluetooth სკანერით — სკანერი კოდს კლავიატურასავით აკრეფს და Enter-ს აგზავნის</div>
        <div style={{ display: "flex", gap: 8, maxWidth: 380, margin: "0 auto" }}>
          <input ref={inputRef} className="input mono" style={{ textAlign: "center", fontSize: 16 }} value={code}
            onChange={(e) => setCode(e.target.value)} placeholder="KR250905-1234"
            onKeyDown={(e) => { if (e.key === "Enter") handleScan(null); }} autoFocus />
          <button className="btn btn-primary" onClick={() => handleScan(null)}>რეგისტრაცია</button>
        </div>
      </div>

      {pendingAssign && (
        <div className="panel" style={{ padding: 16, marginBottom: 18, borderColor: "var(--track)" }}>
          <div style={{ fontWeight: 600, marginBottom: 2 }}>ამანათი {pendingAssign.id} დარეგისტრირდა — მიანიჭეთ კურიერი (სურვილისამებრ)</div>
          <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 10 }}>
            <MapPin size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />
            {pendingAssign.receiver.address}{pendingAssign.receiver.city ? `, ${pendingAssign.receiver.city}` : ""}
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {couriers.length === 0 && <span style={{ fontSize: 13, color: "var(--ink-soft)" }}>კურიერი ჯერ არ არის დამატებული — ეს შესაძლებელია ადმინის პანელიდან.</span>}
            {couriers.map((c) => (
              <button key={c.id} className="btn btn-outline btn-sm" onClick={async () => { await onAssignCourier(pendingAssign, c.id); setPendingAssign(null); }}>{c.name}{c.zone ? ` · ${c.zone}` : ""}</button>
            ))}
            <button className="btn btn-ghost btn-sm" onClick={() => setPendingAssign(null)}>მოგვიანებით</button>
          </div>
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1.1fr .9fr", gap: 18 }} className="reg-grid">
        <div>
          <div style={{ fontWeight: 600, marginBottom: 8, fontSize: 14 }}>სკანირების ისტორია</div>
          {log.length === 0 ? (
            <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>ჯერ არაფერია დასკანერებული ამ სესიაში.</div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {log.map((l, i) => (
                <div key={i} className="panel" style={{ padding: "8px 12px", display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    {l.ok ? <CheckCircle2 size={15} color="#1E8A5F" /> : <AlertTriangle size={15} color="#B23A24" />}
                    <span className="mono" style={{ fontSize: 13 }}>{l.id}</span>
                  </div>
                  <span style={{ fontSize: 12, color: "var(--ink-soft)" }}>{l.msg}</span>
                </div>
              ))}
            </div>
          )}
        </div>
        <div>
          <div style={{ fontWeight: 600, marginBottom: 8, fontSize: 14 }}>ბოლოს დამატებული შეკვეთები</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 320, overflowY: "auto" }}>
            {recentParcels.map((p) => (
              <div key={p.id} className="panel" style={{ padding: "8px 12px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                <div style={{ minWidth: 0 }}>
                  <div className="mono" style={{ fontSize: 12.5 }}>{p.id}</div>
                  <div style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>{p.companyName}</div>
                  <div style={{ fontSize: 11.5, color: "var(--ink-soft)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    <MapPin size={11} style={{ display: "inline", marginRight: 2, verticalAlign: "-1px" }} />
                    {p.receiver.address}{p.receiver.city ? `, ${p.receiver.city}` : ""}
                  </div>
                </div>
                <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 4 }}>
                  <StatusPill status={p.status} />
                  <DispatchWarning parcel={p} />
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
      <style>{`@media (max-width:760px){ .reg-grid{grid-template-columns:1fr !important;} }`}</style>
    </div>
  );
}

/* ================================= ADMIN VIEW ================================ */

function AdminView({ companies, couriers, registrators, admins, parcels, onAddCompany, onAddCourier, onUpdateCourier, onAddRegistrator, onAddAdmin, onAssignCourier, onTransition, loading, onRefresh, onClearCompanies, onClearCouriers, onClearParcels, onClearRegistrators, onClearAdmins }) {
  const [tab, setTab] = useState("overview");

  const byStatus = useMemo(() => {
    const m = {}; Object.keys(STATUS).forEach((k) => (m[k] = 0));
    parcels.forEach((p) => m[p.status]++);
    return m;
  }, [parcels]);
  const codOutstanding = parcels.filter((p) => Number(p.codAmount) > 0 && !p.codCollected).reduce((s, p) => s + Number(p.codAmount), 0);
  const codCollected = parcels.filter((p) => p.codCollected).reduce((s, p) => s + Number(p.codCollectedAmount), 0);
  const unsentCount = parcels.filter((p) => p.status === "created" && !p.dispatchedAt).length;

  return (
    <div>
      <StatRow items={[
        { label: "სულ ამანათი", value: parcels.length },
        { label: "კომპანია", value: companies.length },
        { label: "კურიერი", value: couriers.length },
        { label: "საწყობში გაუგზავნელი", value: unsentCount },
        { label: "მოსალოდნელი COD", value: fmtMoney(codOutstanding) },
        { label: "მიღებული COD", value: fmtMoney(codCollected) },
      ]} />
      <TabBar active={tab} onChange={setTab} tabs={[
        { key: "overview", label: "ყველა ამანათი", icon: LayoutDashboard },
        { key: "reports", label: "რეპორტინგი", icon: FileText },
        { key: "companies", label: "კომპანიები", icon: Building2 },
        { key: "couriers", label: "კურიერები", icon: Users },
        { key: "staff", label: "პერსონალი", icon: User },
      ]} />

      {tab === "overview" && (
        <AdminParcelTable parcels={parcels} couriers={couriers} onAssignCourier={onAssignCourier} onTransition={onTransition} byStatus={byStatus} loading={loading} onRefresh={onRefresh} onClearAll={onClearParcels} />
      )}
      {tab === "reports" && <ReportsView parcels={parcels} couriers={couriers} />}
      {tab === "companies" && <AdminCompanies companies={companies} onAdd={onAddCompany} parcels={parcels} onClearAll={onClearCompanies} />}
      {tab === "couriers" && <AdminCouriers couriers={couriers} onAdd={onAddCourier} onUpdate={onUpdateCourier} parcels={parcels} onClearAll={onClearCouriers} />}
      {tab === "staff" && <AdminStaff registrators={registrators} admins={admins} onAddRegistrator={onAddRegistrator} onAddAdmin={onAddAdmin} onClearRegistrators={onClearRegistrators} onClearAdmins={onClearAdmins} />}
    </div>
  );
}

function AdminParcelTable({ parcels, couriers, onAssignCourier, onTransition, loading, onRefresh, onClearAll }) {
  const [statusFilter, setStatusFilter] = useState("all");
  const [q, setQ] = useState("");
  const [expanded, setExpanded] = useState(null);
  const [printing, setPrinting] = useState(null);
  const [page, setPage] = useState(1);

  const filtered = parcels.filter((p) => {
    if (statusFilter !== "all" && p.status !== statusFilter) return false;
    if (q) {
      const s = q.toLowerCase();
      if (!(p.id.toLowerCase().includes(s) || p.companyName.toLowerCase().includes(s) || p.receiver.name.toLowerCase().includes(s) || (p.courierName || "").toLowerCase().includes(s) || p.receiver.address.toLowerCase().includes(s))) return false;
    }
    return true;
  });
  useEffect(() => { setPage(1); }, [statusFilter, q]);
  const pageItems = filtered.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  function exportExcel() {
    // ექსპორტი ყოველთვის სრულ გაფილტრულ სიას მოიცავს, არა მხოლოდ მიმდინარე გვერდს
    downloadXlsx("kursi-amanatebi.xlsx", filtered.map((p) => ({
      კოდი: p.id, კომპანია: p.companyName, მიმღები: p.receiver.name, ტელეფონი: p.receiver.phone,
      მისამართი: `${p.receiver.address}${p.receiver.city ? ", " + p.receiver.city : ""}`,
      კურიერი: p.courierName || "", სტატუსი: STATUS[p.status]?.label || p.status,
      cod: p.codAmount, საკურიერო_საფასური: p.deliveryFee, შექმნის_თარიღი: fmtDate(p.createdAt),
    })), "ამანათები");
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
        <div style={{ position: "relative", flex: "1 1 240px" }}>
          <Search size={14} style={{ position: "absolute", left: 10, top: 10, color: "var(--ink-soft)" }} />
          <input className="input" style={{ paddingLeft: 30 }} placeholder="ძებნა კოდით, კომპანიით, მიმღებით, მისამართით, კურიერით" value={q} onChange={(e) => setQ(e.target.value)} />
        </div>
        <select className="select" style={{ maxWidth: 220 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
          <option value="all">ყველა სტატუსი</option>
          {Object.entries(STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
        <button className="btn btn-outline btn-sm" onClick={onRefresh}><RefreshCw size={13} /> განახლება</button>
        <button className="btn btn-outline btn-sm" onClick={exportExcel} disabled={filtered.length === 0}><FileSpreadsheet size={13} /> Excel ექსპორტი</button>
        <DangerZone label="ყველა ამანათის წაშლა" disabled={parcels.length === 0} onConfirm={onClearAll} />
      </div>

      {loading ? <div style={{ color: "var(--ink-soft)", fontSize: 13.5 }}>იტვირთება...</div> :
        filtered.length === 0 ? <EmptyState icon={Filter} title="ამანათი ვერ მოიძებნა" hint="შეცვალეთ ფილტრი ან ძიების სიტყვა." /> : (
          <div className="panel" style={{ overflowX: "auto" }}>
            <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th>კოდი</th><th>კომპანია</th><th>მიმღები</th><th>მისამართი</th><th>კურიერი</th><th>სტატუსი</th><th>COD</th><th></th></tr></thead>
              <tbody>
                {pageItems.map((p) => (
                  <React.Fragment key={p.id}>
                    <tr style={{ cursor: "pointer" }} onClick={() => setExpanded(expanded === p.id ? null : p.id)}>
                      <td className="mono">{p.id}</td>
                      <td>{p.companyName}</td>
                      <td>{p.receiver.name}<div style={{ fontSize: 11, color: "var(--ink-soft)" }}>{p.receiver.phone}</div></td>
                      <td style={{ maxWidth: 200 }}>{p.receiver.address}{p.receiver.city ? `, ${p.receiver.city}` : ""}</td>
                      <td>{p.courierName || <span style={{ color: "var(--ink-soft)" }}>—</span>}</td>
                      <td>
                        <div style={{ display: "flex", flexDirection: "column", gap: 4, alignItems: "flex-start" }}>
                          <StatusPill status={p.status} /><DispatchWarning parcel={p} />
                        </div>
                      </td>
                      <td>{Number(p.codAmount) > 0 ? fmtMoney(p.codAmount) + (p.codCollected ? " ✓" : "") : "—"}</td>
                      <td>{expanded === p.id ? <ChevronDown size={15} /> : <ChevronRight size={15} />}</td>
                    </tr>
                    {expanded === p.id && (
                      <tr><td colSpan={8} style={{ background: "#FAFAF8" }}>
                        <div style={{ padding: "10px 4px", display: "flex", flexDirection: "column", gap: 10 }}>
                          <div style={{ display: "flex", gap: 20, flexWrap: "wrap", fontSize: 12.5, color: "var(--ink-soft)" }}>
                            <span><Phone size={12} style={{ display: "inline", marginRight: 3, verticalAlign: "-2px" }} />გამგზავნი: {p.sender.name}, {p.sender.phone}</span>
                            <span>წონა: {p.weight || "—"} კგ</span>
                            <span>გადახდა: {p.paymentMethod}</span>
                            <span>საკურიერო საფასური: {fmtMoney(p.deliveryFee)}</span>
                            {p.status === "delivered" && <span>ჩასარიცხი კომპანიას: <b style={{ color: "var(--ink)" }}>{fmtMoney(companyNetFor(p))}</b></span>}
                            <span>შენიშვნა: {p.notes || "—"}</span>
                          </div>
                          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                            <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>კურიერის მინიჭება:</span>
                            <select className="select" style={{ maxWidth: 200 }} value={p.courierId || ""} onChange={(e) => onAssignCourier(p, e.target.value)}>
                              <option value="">— აირჩიეთ —</option>
                              {couriers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                            </select>
                            <button className="btn btn-outline btn-sm" onClick={() => setPrinting(p)}><Printer size={13} /> ბეჭდვა</button>
                            {!["delivered", "cancelled", "returned"].includes(p.status) && (
                              <button className="btn btn-outline btn-sm" onClick={() => onTransition(p, "cancelled", "გააუქმა ადმინისტრატორმა")}>შეკვეთის გაუქმება</button>
                            )}
                          </div>
                          {p.courierId && <LocationStatus courier={couriers.find((c) => c.id === p.courierId)} />}
                          <HistoryTimeline history={p.history} />
                        </div>
                      </td></tr>
                    )}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
          </div>
        )}
      <Pagination page={page} setPage={setPage} total={filtered.length} />
      {printing && <PrintLabelModal parcel={printing} onClose={() => setPrinting(null)} />}
    </div>
  );
}

/* ================================ REPORTS VIEW =============================== */

function ReportsView({ parcels, couriers }) {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");

  const delivered = useMemo(() => parcels.filter((p) => {
    if (p.status !== "delivered") return false;
    const at = p.deliveredAt || p.history.find((h) => h.status === "delivered")?.at;
    if (!at) return true;
    const d = new Date(at);
    if (from && d < new Date(from)) return false;
    if (to && d > new Date(to + "T23:59:59")) return false;
    return true;
  }), [parcels, from, to]);

  const companyRows = useMemo(() => {
    const map = {};
    delivered.forEach((p) => {
      if (!map[p.companyId]) map[p.companyId] = { companyId: p.companyId, companyName: p.companyName, count: 0, codTotal: 0, feeTotal: 0, net: 0 };
      const r = map[p.companyId];
      r.count++; r.codTotal += Number(p.codAmount) || 0; r.feeTotal += Number(p.deliveryFee) || 0; r.net += companyNetFor(p);
    });
    return Object.values(map).sort((a, b) => b.net - a.net);
  }, [delivered]);

  const courierRows = useMemo(() => {
    const map = {};
    delivered.forEach((p) => {
      if (!p.courierId) return;
      const courier = couriers.find((c) => c.id === p.courierId);
      if (!map[p.courierId]) map[p.courierId] = { courierId: p.courierId, courierName: p.courierName, count: 0, feeTotal: 0, commission: 0 };
      const r = map[p.courierId];
      r.count++; r.feeTotal += Number(p.deliveryFee) || 0; r.commission += courierCommissionFor(p, courier);
    });
    return Object.values(map).sort((a, b) => b.commission - a.commission);
  }, [delivered, couriers]);

  const totalFee = delivered.reduce((s, p) => s + (Number(p.deliveryFee) || 0), 0);
  const totalPayroll = courierRows.reduce((s, r) => s + r.commission, 0);
  const totalToCompanies = companyRows.reduce((s, r) => s + r.net, 0);
  const totalCod = companyRows.reduce((s, r) => s + r.codTotal, 0);
  const netRevenue = totalFee - totalPayroll;

  function exportCompanies() {
    downloadXlsx("kompaniebtan-angarishsts.xlsx", companyRows.map((r) => ({
      კომპანია: r.companyName, ამანათები: r.count, cod_სულ: r.codTotal.toFixed(2),
      საკურიერო_საფასური: r.feeTotal.toFixed(2), ჩასარიცხი: r.net.toFixed(2),
    })), "ანგარიშსწორება");
  }
  function exportPayroll() {
    downloadXlsx("kurierebis-khelfasebi.xlsx", courierRows.map((r) => ({
      კურიერი: r.courierName, ამანათები: r.count, საკურიერო_შემოსავალი: r.feeTotal.toFixed(2), ხელფასი: r.commission.toFixed(2),
    })), "ხელფასები");
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", marginBottom: 16 }}>
        <Field label="პერიოდი — დან"><input type="date" className="input" style={{ maxWidth: 160 }} value={from} onChange={(e) => setFrom(e.target.value)} /></Field>
        <Field label="მდე"><input type="date" className="input" style={{ maxWidth: 160 }} value={to} onChange={(e) => setTo(e.target.value)} /></Field>
        {(from || to) && <button className="btn btn-ghost btn-sm" onClick={() => { setFrom(""); setTo(""); }}>გასუფთავება</button>}
        <span style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>რეპორტი ავტომატურად აითვლება მხოლოდ მიწოდებულ ამანათებზე</span>
      </div>

      <StatRow items={[
        { label: "მიწოდებული ამანათი", value: delivered.length },
        { label: "სულ COD შეგროვილი", value: fmtMoney(totalCod) },
        { label: "საკურიერო შემოსავალი", value: fmtMoney(totalFee) },
        { label: "გაცემული ხელფასები", value: fmtMoney(totalPayroll) },
        { label: "კომპანიის სუფთა შემოსავალი", value: fmtMoney(netRevenue) },
      ]} />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div style={{ fontWeight: 600, fontSize: 14 }}>კომპანიებთან ანგარიშსწორება — ავტომატური „ჩასარიცხი" რეპორტი</div>
        <button className="btn btn-outline btn-sm" onClick={exportCompanies} disabled={companyRows.length === 0}><FileSpreadsheet size={13} /> Excel</button>
      </div>
      <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 10 }}>ჩასარიცხი = COD-ით შეგროვილი თანხა მინუს საკურიერო საფასური. წინასწარ გადახდილ/ბანკო შეკვეთებზე თანხა უარყოფითია — ეს არის კომპანიის დავალიანება საკურიერო მომსახურებაზე.</div>
      {companyRows.length === 0 ? <EmptyState icon={Building2} title="ამ პერიოდში მიწოდებული ამანათი არ არის" /> : (
        <div className="panel" style={{ overflowX: "auto", marginBottom: 26 }}>
          <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>კომპანია</th><th>ამანათი</th><th>COD სულ</th><th>საკურიერო საფასური</th><th>ჩასარიცხი კომპანიას</th></tr></thead>
            <tbody>
              {companyRows.map((r) => (
                <tr key={r.companyId}>
                  <td>{r.companyName}</td><td>{r.count}</td><td>{fmtMoney(r.codTotal)}</td><td>{fmtMoney(r.feeTotal)}</td>
                  <td style={{ fontWeight: 700, color: r.net >= 0 ? "#1E8A5F" : "#B23A24" }}>{fmtMoney(r.net)}</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>ჯამი</td><td style={{ fontWeight: 700 }}>{delivered.length}</td>
                <td style={{ fontWeight: 700 }}>{fmtMoney(totalCod)}</td><td style={{ fontWeight: 700 }}>{fmtMoney(totalFee)}</td>
                <td style={{ fontWeight: 700 }}>{fmtMoney(totalToCompanies)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div style={{ fontWeight: 600, fontSize: 14 }}>კურიერების სახელფასო რეპორტი</div>
        <button className="btn btn-outline btn-sm" onClick={exportPayroll} disabled={courierRows.length === 0}><FileSpreadsheet size={13} /> Excel</button>
      </div>
      <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 10 }}>ხელფასი ითვლება ავტომატურად თითოეული კურიერის საკომისიო განაკვეთის მიხედვით (იხ. „კურიერები" ტაბი) — საკურიერო საფასურის % ან ფიქსირებული თანხა ამანათზე.</div>
      {courierRows.length === 0 ? <EmptyState icon={Truck} title="ამ პერიოდში კურიერზე მიწოდებული ამანათი არ არის" /> : (
        <div className="panel" style={{ overflowX: "auto" }}>
          <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>კურიერი</th><th>მიწოდებული ამანათი</th><th>საკურიერო შემოსავალი</th><th>ხელფასი (საკომისიო)</th></tr></thead>
            <tbody>
              {courierRows.map((r) => (
                <tr key={r.courierId}>
                  <td>{r.courierName}</td><td>{r.count}</td><td>{fmtMoney(r.feeTotal)}</td>
                  <td style={{ fontWeight: 700, color: "var(--accent)" }}>{fmtMoney(r.commission)}</td>
                </tr>
              ))}
              <tr>
                <td style={{ fontWeight: 700 }}>ჯამი</td><td style={{ fontWeight: 700 }}>{courierRows.reduce((s, r) => s + r.count, 0)}</td>
                <td style={{ fontWeight: 700 }}>{fmtMoney(totalFee)}</td><td style={{ fontWeight: 700 }}>{fmtMoney(totalPayroll)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function AdminCompanies({ companies, onAdd, parcels, onClearAll }) {
  const [showForm, setShowForm] = useState(false);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 13.5, color: "var(--ink-soft)" }}>სულ {companies.length} კომპანია</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <DangerZone label="ყველა კომპანიის წაშლა" disabled={companies.length === 0} onConfirm={onClearAll} />
          <button className="btn btn-primary btn-sm" onClick={() => setShowForm((s) => !s)}><Plus size={14} /> ახალი კომპანია</button>
        </div>
      </div>
      {showForm && (
        <div className="panel" style={{ padding: 16, marginBottom: 14, maxWidth: 420 }}>
          <NewCompanyForm existingUsernames={companies.map((c) => (c.username || "").toLowerCase())} onCancel={() => setShowForm(false)} onSubmit={async (d) => { await onAdd(d); setShowForm(false); }} />
        </div>
      )}
      {companies.length === 0 ? <EmptyState icon={Building2} title="კომპანია არ არის დამატებული" /> : (
        <div className="panel" style={{ overflowX: "auto" }}>
          <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>დასახელება</th><th>საკონტაქტო</th><th>ტელეფონი</th><th>მომხმარებელი</th><th>შეკვეთები</th></tr></thead>
            <tbody>{companies.map((c) => (
              <tr key={c.id}><td>{c.name}</td><td>{c.contactPerson || "—"}</td><td>{c.phone || "—"}</td><td className="mono">{c.username || "—"}</td>
                <td>{parcels.filter((p) => p.companyId === c.id).length}</td></tr>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function AdminCouriers({ couriers, onAdd, onUpdate, parcels, onClearAll }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 13.5, color: "var(--ink-soft)" }}>სულ {couriers.length} კურიერი</div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <DangerZone label="ყველა კურიერის წაშლა" disabled={couriers.length === 0} onConfirm={onClearAll} />
          <button className="btn btn-primary btn-sm" onClick={() => setShowForm((s) => !s)}><Plus size={14} /> ახალი კურიერი</button>
        </div>
      </div>
      {showForm && (
        <div className="panel" style={{ padding: 16, marginBottom: 14, maxWidth: 420 }}>
          <NewCourierForm existingUsernames={couriers.map((c) => (c.username || "").toLowerCase())} onCancel={() => setShowForm(false)} onSubmit={async (d) => { await onAdd(d); setShowForm(false); }} />
        </div>
      )}
      {couriers.length === 0 ? <EmptyState icon={Truck} title="კურიერი არ არის დამატებული" /> : (
        <div className="panel" style={{ overflowX: "auto" }}>
          <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr><th>სახელი</th><th>ტელეფონი</th><th>ტრანსპორტი</th><th>ზონა</th><th>მომხმარებელი</th><th>საკომისიო</th><th>აქტიური ამანათი</th><th>მდებარეობა</th><th></th></tr></thead>
            <tbody>{couriers.map((c) => (
              <React.Fragment key={c.id}>
                <tr>
                  <td>{c.name}</td><td>{c.phone || "—"}</td><td>{c.vehicle || "—"}</td><td>{c.zone || "—"}</td><td className="mono">{c.username || "—"}</td>
                  <td>{(c.commissionType || "percent") === "percent" ? `${c.commissionValue ?? 60}%` : fmtMoney(c.commissionValue ?? 0) + " / ამანათი"}</td>
                  <td>{parcels.filter((p) => p.courierId === c.id && ["assigned", "transit", "failed"].includes(p.status)).length}</td>
                  <td style={{ maxWidth: 170 }}><LocationStatus courier={c} /></td>
                  <td><button className="btn btn-ghost btn-sm" onClick={() => setEditing(editing === c.id ? null : c.id)}>ჩასწორება</button></td>
                </tr>
                {editing === c.id && (
                  <tr><td colSpan={9} style={{ background: "#FAFAF8" }}>
                    <CourierCommissionEdit courier={c} onSave={(patch) => { onUpdate(c.id, patch); setEditing(null); }} onCancel={() => setEditing(null)} />
                  </td></tr>
                )}
              </React.Fragment>
            ))}</tbody>
          </table>
        </div>
      )}
    </div>
  );
}
function CourierCommissionEdit({ courier, onSave, onCancel }) {
  const [type, setType] = useState(courier.commissionType || "percent");
  const [value, setValue] = useState(courier.commissionValue ?? 60);
  return (
    <div style={{ padding: "10px 4px", display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
      <Field label="საკომისიო ტიპი">
        <select className="select" value={type} onChange={(e) => setType(e.target.value)}>
          <option value="percent">% საკურიერო საფასურიდან</option>
          <option value="fixed">ფიქსირებული, ₾ / ამანათი</option>
        </select>
      </Field>
      <Field label={type === "percent" ? "პროცენტი (%)" : "თანხა (₾)"}>
        <input className="input" style={{ maxWidth: 120 }} value={value} onChange={(e) => setValue(e.target.value)} />
      </Field>
      <button className="btn btn-primary btn-sm" onClick={() => onSave({ commissionType: type, commissionValue: Number(value) || 0 })}>შენახვა</button>
      <button className="btn btn-outline btn-sm" onClick={onCancel}>გაუქმება</button>
    </div>
  );
}

function AdminStaff({ registrators, admins, onAddRegistrator, onAddAdmin, onClearRegistrators, onClearAdmins }) {
  const [showReg, setShowReg] = useState(false);
  const [showAdmin, setShowAdmin] = useState(false);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 26 }}>
      <div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 600 }}>რეგისტრატორები ({registrators.length})</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <DangerZone label="ყველა რეგისტრატორის წაშლა" disabled={registrators.length === 0} onConfirm={onClearRegistrators} />
            <button className="btn btn-primary btn-sm" onClick={() => setShowReg((s) => !s)}><Plus size={14} /> ახალი რეგისტრატორი</button>
          </div>
        </div>
        {showReg && (
          <div className="panel" style={{ padding: 16, marginBottom: 14, maxWidth: 420 }}>
            <NewStaffForm existingUsernames={registrators.map((r) => (r.username || "").toLowerCase())} onCancel={() => setShowReg(false)} onSubmit={async (d) => { await onAddRegistrator(d); setShowReg(false); }} />
          </div>
        )}
        {registrators.length === 0 ? <EmptyState icon={ScanLine} title="რეგისტრატორი არ არის დამატებული" /> : (
          <div className="panel" style={{ overflowX: "auto" }}>
            <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th>სახელი</th><th>მომხმარებელი</th><th>დამატებულია</th></tr></thead>
              <tbody>{registrators.map((r) => <tr key={r.id}><td>{r.name}</td><td className="mono">{r.username}</td><td>{fmtDate(r.createdAt)}</td></tr>)}</tbody>
            </table>
          </div>
        )}
      </div>
      <div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
          <div style={{ fontWeight: 600 }}>ადმინისტრატორები ({admins.length})</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <DangerZone label="ყველა ადმინისტრატორის წაშლა" disabled={admins.length <= 1} onConfirm={onClearAdmins} />
            <button className="btn btn-primary btn-sm" onClick={() => setShowAdmin((s) => !s)}><Plus size={14} /> ახალი ადმინისტრატორი</button>
          </div>
        </div>
        <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: -6, marginBottom: 10 }}>„ყველას წაშლა" ტოვებს თქვენს მიმდინარე ანგარიშს ხელუხლებელს, რომ არ ჩაიკეტოთ სისტემიდან.</div>
        {showAdmin && (
          <div className="panel" style={{ padding: 16, marginBottom: 14, maxWidth: 420 }}>
            <NewStaffForm existingUsernames={admins.map((a) => (a.username || "").toLowerCase())} onCancel={() => setShowAdmin(false)} onSubmit={async (d) => { await onAddAdmin(d); setShowAdmin(false); }} />
          </div>
        )}
        {admins.length === 0 ? <EmptyState icon={ShieldCheck} title="ადმინისტრატორი არ არის დამატებული" /> : (
          <div className="panel" style={{ overflowX: "auto" }}>
            <table className="data" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead><tr><th>სახელი</th><th>მომხმარებელი</th><th>დამატებულია</th></tr></thead>
              <tbody>{admins.map((a) => <tr key={a.id}><td>{a.name}</td><td className="mono">{a.username}</td><td>{fmtDate(a.createdAt)}</td></tr>)}</tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
