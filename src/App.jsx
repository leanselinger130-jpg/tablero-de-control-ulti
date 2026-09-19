import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { storage } from "./storageClient";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend,
  BarChart, Bar, Cell, ScatterChart, Scatter, LabelList,
} from "recharts";
import {
  UploadCloud, Radio, Trash2, AlertTriangle, Loader2, Table as TableIcon,
  LayoutGrid, X, CheckCircle2, RefreshCw, Users, Pencil, Target, SlidersHorizontal, Search,
  FileSpreadsheet, Printer, Trophy, Compass, MessageCircleQuestion, BookPlus, Info, BarChart3, LayoutDashboard,
  Factory, Wallet, TrendingUp, ClipboardList,
} from "lucide-react";

/* ---------------------------------------------------------------
   TOKENS
--------------------------------------------------------------- */
const T = {
  bg: "#F4F5FA", panel: "#FFFFFF", panelAlt: "#F8F9FC",
  border: "#E3E6EE", borderSoft: "#EDF0F6",
  text: "#12141C", textDim: "#5B6474", textFaint: "#98A2B3",
  amber: "#4F46E5", amberDim: "#EEF0FF",
  cyan: "#0D9488", cyanDim: "#E6F5F3",
  green: "#16A34A", red: "#DC2626", violet: "#7C3AED",
  overlay: "rgba(15,23,42,0.4)",
};
const CARD_SHADOW = "0 1px 2px rgba(15,23,42,0.05), 0 4px 16px rgba(15,23,42,0.04)";
const SERIES_COLORS = [T.amber, T.cyan, T.violet, T.green, T.red];
const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');`;
const ULTI_LOGO = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCADPAowDASIAAhEBAxEB/8QAHgABAAEEAwEBAAAAAAAAAAAAAAcFBggJAwQKAQL/xABYEAABAwMCAgQGCwkOBQIHAAABAAIDBAURBgcSIQgxQWETFCJRcbQJFRYZMjh1doGmszY3Qld0kaGy1BcjM0NiY2VocoKxwdHkGFJWZpVTczQ1ZIOSwsP/xAAcAQEAAQUBAQAAAAAAAAAAAAAABAECAwUGBwj/xABAEQACAQIDAwgHBwQCAQUAAAAAAQIDEQQFMRIhQQYyUWFxgaGxBxMUInKRwTNSYrLR4fAVIyRzJTSSFkKCwvH/2gAMAwEAAhEDEQA/ANqaIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIuN78ID9FwC+eECt7WOttM6DsVRqTVl3gt1vpx5cspPN3PDWtGS9x7GgEnzLG+79P3SNNcHwWTQt0rqRkhaKiaqZTl7R+E1mHdfmJBxjqPIb3K+TeaZzFywNFzS47kuy7aV+rUiYjHYfDO1WVmZYh4K/QIKiPaHpHbf7wvdQWWaegu7GOkdbqwNbKWA83MIJa8dRODkZ5hSsyTK1+Oy7E5bWeHxUHCa4P8Am9dZmpVqdeO3Td0c6L4DlfVCMoREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBEXHPPDTQyVNRKyKKJpfI97sNa0DJJJ6gAiV9yByLiiqqad744aiKR8Zw9rXglp7wOpauukZ0tNabxakrbRpi81to0ZTSuio6SmkdC+tYDjwtQWkF3F1hh8loxyJyTEen7jdLLWxXGz3GqoauE5jnppnRSMPnDmkEL1rLvRRi8ThVWxVdU5tX2dnat1N3W/psnY5Gpyto+ucKNPaiuN7X7FZm6ZFjF0P+kdeNyo59Aa8qxU32hg8Yo60tDXVkDcBzZMcjI3IOfwgSTzaScnV5znOT4nIsZLBYpe9HitGno11P9tTpsLiaeLpKrT0YREWrJAREQBERAEREAREQBERAfHHAXVmfwgnK7EnUujVuIafQslNXZbIwF6bGu7nqDdM6MMr2W7TdPEGwgnhfPNG2V0hHaeF7GjzYPnKx6DcqVelFl2/OqzjrlpvVYlGUcfaQvrLkzh6eFybCwpqy2IvvaTb722zz7Gyc8TUlLpfgzvafu1003eKO/2WskpK+gmbPTzMOCx7TkekecdRGQVtT0RqMar0pZtSiMRC7W+nrfBg5DPCRtfw57s4+hapmM7ls02OcRtPo9vmslF9i1efelahTnh8PXt7ybV+pq9vA2/J+ctucOFkSRGcrkXDGVzLwt6nWBERUAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAVh79VE1LsduDUU8hjlj0vdHMc3rafFZOY71fisDpA89itwx/wBr3T1WRTssV8dRT+/HzRHxf/Xqdj8jTtRM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";

/* ---------------------------------------------------------------
   GENERIC HELPERS
--------------------------------------------------------------- */
function isBlank(v) { return v === undefined || v === null || String(v).trim() === ""; }
function toNumberOrNull(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (isBlank(v)) return null;
  const s = String(v).trim();
  const cleaned = s.replace(/[$%\s]/g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return Number.isFinite(n) && /^-?[\d.]+$/.test(cleaned) ? n : null;
}
function fmtNumber(n) {
  if (!Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  if (abs >= 1_000_000) return (n / 1_000_000).toFixed(2) + "M";
  if (abs >= 1_000) return (n / 1_000).toFixed(1) + "K";
  if (Number.isInteger(n)) return n.toString();
  return n.toFixed(2);
}
function fmtCell(v) {
  if (v === null || v === undefined || v === "") return "—";
  if (typeof v === "number") return v.toLocaleString("es-AR", { maximumFractionDigits: 2 });
  return String(v);
}
function fmtPct(n, digits = 1) {
  if (!Number.isFinite(n)) return "—";
  return `${n.toFixed(digits)}%`;
}
// Nicely formats numbers inside chart tooltips (thousands separators, max 1 decimal)
// instead of dumping raw floats like "1511430.826481615".
function tooltipNumberFormatter(value, name) {
  if (typeof value !== "number" || !Number.isFinite(value)) return [value, name];
  const formatted = value.toLocaleString("es-AR", { maximumFractionDigits: 1 });
  return [formatted, name];
}

/* ---------------------------------------------------------------
   CESIM "Results" PARSER
--------------------------------------------------------------- */
function isCesimResultsSheet(aoa) {
  const a0 = aoa?.[0]?.[0];
  return typeof a0 === "string" && /^Resultados,\s*Ronda/i.test(a0.trim());
}
function detectRoundInfo(titleRaw) {
  const m = titleRaw.match(/(\d+)\s*$/);
  const roundNumber = m ? parseInt(m[1], 10) : null;
  const label = titleRaw.replace(/^Resultados,\s*/i, "").trim();
  return { roundNumber, label };
}
function categoryOf(title) {
  const t = title.toLowerCase();
  if (/(cuenta de resultados|hoja de balance|flujo de efectivo|ratios e indicadores|indicadores financieros|tasas de interés|valuaci[oó]n|creaci[oó]n de valor)/.test(t)) return "Financiero";
  if (/informe de mercado|cuota/.test(t)) return "Mercado";
  if (/fabricaci[oó]n|log[ií]stica|informe de costos|proveedor de componentes|demanda estimada|precio de venta/.test(t)) return "Producción";
  if (/rrhh|recursos humanos/.test(t)) return "RRHH";
  if (/esg|ambiental|social|gobernanza/.test(t)) return "ESG";
  return "Otros";
}
function rowIsEmpty(row) { return !row || row.every((c) => isBlank(c)); }

function parseCesimResults(aoa) {
  const titleRaw = String(aoa[0][0]).trim();
  const { roundNumber, label } = detectRoundInfo(titleRaw);

  let teams = null;
  for (let r = 1; r < aoa.length && r < 40; r++) {
    const row = aoa[r] || [];
    if (!isBlank(row[0])) continue;
    const rest = row.slice(1);
    const nonEmpty = rest.filter((c) => !isBlank(c));
    const textish = nonEmpty.filter((c) => toNumberOrNull(c) === null);
    if (nonEmpty.length >= 3 && textish.length === nonEmpty.length) {
      teams = rest.map((c) => (isBlank(c) ? null : String(c).trim()));
      break;
    }
  }
  if (!teams) return null;
  const teamColCount = teams.length;

  const sameTeams = (row) => {
    const rest = (row || []).slice(1, 1 + teamColCount);
    if (rest.length < teamColCount) return false;
    for (let i = 0; i < teamColCount; i++) {
      const cell = isBlank(rest[i]) ? null : String(rest[i]).trim();
      if (cell !== teams[i]) return false;
    }
    return true;
  };

  const headerRows = [];
  for (let r = 1; r < aoa.length; r++) {
    if (isBlank(aoa[r][0]) && sameTeams(aoa[r])) headerRows.push(r);
  }
  if (headerRows.length === 0) return null;

  const blocks = [];
  headerRows.forEach((h, i) => {
    let title = null;
    for (let r = h - 1; r >= Math.max(0, h - 8); r--) {
      const row = aoa[r];
      if (rowIsEmpty(row)) continue;
      if (!isBlank(row[0]) && row.slice(1).every((c) => isBlank(c))) title = String(row[0]).trim();
      break;
    }
    if (!title) title = `Sección fila ${h}`;

    const nextH = headerRows[i + 1] ?? aoa.length;
    const rawRows = [];
    for (let r = h + 1; r < nextH; r++) {
      const row = aoa[r] || [];
      if (rowIsEmpty(row)) continue;
      if (isBlank(row[0])) continue;
      const vals = teams.map((_, ci) => { const cell = row[1 + ci]; return isBlank(cell) ? null : cell; });
      const hasData = vals.some((v) => v !== null);
      if (hasData) rawRows.push({ kind: "metric", label: String(row[0]).trim(), values: vals });
      else rawRows.push({ kind: "group", label: String(row[0]).trim() });
    }
    const rows = rawRows.filter((row, idx) => {
      if (row.kind !== "group") return true;
      const next = rawRows[idx + 1];
      return !(next && next.kind === "metric" && next.label === row.label);
    });

    blocks.push({ title, category: categoryOf(title), rows });
  });

  return { titleRaw, roundNumber, label, teams, blocks };
}

function parseGenericSheet(aoa) {
  const cleaned = aoa.filter((r) => !rowIsEmpty(r));
  if (cleaned.length < 2) return null;
  const headers = cleaned[0].map((h) => (isBlank(h) ? "" : String(h).trim()));
  const rows = cleaned.slice(1);
  return { headers, rows };
}

async function parseFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv")) {
    const text = await file.text();
    const parsed = Papa.parse(text.trim(), { skipEmptyLines: true });
    const generic = parseGenericSheet(parsed.data);
    return { kind: "generic", sheets: generic ? { Datos: generic } : {} };
  }
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    if (isCesimResultsSheet(aoa)) {
      const parsed = parseCesimResults(aoa);
      if (parsed) return { kind: "cesim", data: parsed };
    }
  }
  const sheets = {};
  wb.SheetNames.forEach((sn) => {
    const ws = wb.Sheets[sn];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    const generic = parseGenericSheet(aoa);
    if (generic) sheets[sn] = generic;
  });
  return { kind: "generic", sheets };
}

/* ---------------------------------------------------------------
   STORAGE LAYER
--------------------------------------------------------------- */
async function loadRoundsMeta() {
  try { const res = await storage.get("rounds-meta"); return res ? JSON.parse(res.value) : []; } catch { return []; }
}
async function saveRoundsMeta(meta) { await storage.set("rounds-meta", JSON.stringify(meta)); }
async function loadRoundData(id) {
  try { const res = await storage.get(`round-data:${id}`); return res ? JSON.parse(res.value) : null; } catch { return null; }
}
async function saveRoundData(id, data) { await storage.set(`round-data:${id}`, JSON.stringify(data)); }
async function deleteRoundData(id) { try { await storage.delete(`round-data:${id}`); } catch {} }
async function loadOurTeam() {
  try { const res = await storage.get("our-team"); return res ? res.value : null; } catch { return null; }
}
async function saveOurTeam(team) { await storage.set("our-team", team); }
async function loadNote(roundId) {
  try { const res = await storage.get(`note:${roundId}`); return res ? res.value : ""; } catch { return ""; }
}
async function saveNote(roundId, text) { await storage.set(`note:${roundId}`, text); }
async function loadStrategyPlan() {
  try { const res = await storage.get("strategy-plan"); return res ? JSON.parse(res.value) : null; } catch { return null; }
}
async function saveStrategyPlan(plan) { await storage.set("strategy-plan", JSON.stringify(plan)); }
async function loadStrategySnapshot(roundId) {
  try { const res = await storage.get(`strategy-snapshot:${roundId}`); return res ? JSON.parse(res.value) : null; } catch { return null; }
}
async function saveStrategySnapshot(roundId, plan) { await storage.set(`strategy-snapshot:${roundId}`, JSON.stringify(plan)); }
async function loadStrategyComment(roundId) {
  try { const res = await storage.get(`strategy-comment:${roundId}`); return res ? res.value : ""; } catch { return ""; }
}
async function saveStrategyComment(roundId, text) { await storage.set(`strategy-comment:${roundId}`, text); }
async function loadPremises() {
  try { const res = await storage.get("premises-list"); return res ? JSON.parse(res.value) : []; } catch { return []; }
}
async function savePremises(list) { await storage.set("premises-list", JSON.stringify(list)); }
async function loadDashboardStory() {
  try { const res = await storage.get("dashboard-story"); return res ? JSON.parse(res.value) : null; } catch { return null; }
}
async function saveDashboardStory(story) { await storage.set("dashboard-story", JSON.stringify(story)); }
async function loadVmsData() {
  try { const res = await storage.get("vms-data"); return res ? JSON.parse(res.value) : null; } catch { return null; }
}
async function saveVmsData(data) { await storage.set("vms-data", JSON.stringify(data)); }

/* ---------------------------------------------------------------
   UI PRIMITIVES
--------------------------------------------------------------- */
function Panel({ children, style, ...rest }) {
  return <div style={{ background: T.panel, border: `1px solid ${T.border}`, borderRadius: 14, boxShadow: CARD_SHADOW, ...style }} {...rest}>{children}</div>;
}
function Eyebrow({ children, info }) {
  return (
    <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, letterSpacing: "0.14em", textTransform: "uppercase", color: T.textFaint }}>
        {children}
      </div>
      {info && <InfoBubble text={info} />}
    </div>
  );
}

/* Click-to-open info bubble — replaces permanently-visible helper subtitles */
function InfoBubble({ text }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const onClick = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", onClick);
    return () => document.removeEventListener("mousedown", onClick);
  }, [open]);
  return (
    <span ref={ref} style={{ position: "relative", display: "inline-flex", marginLeft: 7 }}>
      <button
        onClick={(e) => { e.stopPropagation(); setOpen((o) => !o); }}
        title="¿Qué muestra esto?"
        style={{
          width: 15, height: 15, borderRadius: "50%", border: `1px solid ${open ? T.amber : T.textFaint}`,
          background: open ? T.amberDim : "none", color: open ? T.amber : T.textFaint, fontSize: 10, fontWeight: 700,
          cursor: "pointer", display: "inline-flex", alignItems: "center", justifyContent: "center", padding: 0, lineHeight: 1,
        }}
      >i</button>
      {open && (
        <div style={{
          position: "absolute", top: "150%", left: 0, zIndex: 30, width: 250, background: T.panelAlt,
          border: `1px solid ${T.border}`, borderRadius: 8, padding: "10px 12px", fontSize: 12.5, color: T.textDim,
          lineHeight: 1.5, boxShadow: "0 6px 20px rgba(0,0,0,0.5)", fontFamily: "'Inter', sans-serif", fontWeight: 400,
          textTransform: "none", letterSpacing: "normal",
        }}>
          {text}
        </div>
      )}
    </span>
  );
}

const inputStyle = {
  background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text,
  padding: "8px 10px", fontSize: 13, fontFamily: "'IBM Plex Mono', monospace", outline: "none",
};
const primaryBtn = {
  display: "flex", alignItems: "center", gap: 6, background: T.amber, color: "#FFFFFF",
  border: "none", borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 600, cursor: "pointer",
};
const ghostBtn = {
  display: "flex", alignItems: "center", gap: 6, background: "transparent", color: T.textDim,
  border: `1px solid ${T.border}`, borderRadius: 7, padding: "9px 16px", fontSize: 13, fontWeight: 500, cursor: "pointer",
};

/* ---------------------------------------------------------------
   TEAM PICKER
--------------------------------------------------------------- */
function TeamPicker({ teams, current, onPick }) {
  const [open, setOpen] = useState(!current);
  useEffect(() => { if (!current) setOpen(true); }, [current]);
  if (!open) {
    return (
      <button onClick={() => setOpen(true)} style={{
        display: "flex", alignItems: "center", gap: 6, background: T.amberDim, border: `1px solid ${T.amber}`,
        borderRadius: 20, padding: "5px 12px", fontSize: 12, color: T.amber, cursor: "pointer", fontWeight: 600,
      }}><Users size={12} /> Nuestro equipo: {current} <Pencil size={11} style={{ opacity: 0.7 }} /></button>
    );
  }
  return (
    <Panel style={{ padding: 14, display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
      <span style={{ fontSize: 12.5, color: T.textDim, marginRight: 4 }}>¿Cuál es su equipo?</span>
      {teams.map((tm) => (
        <button key={tm} onClick={() => { onPick(tm); setOpen(false); }} style={{
          padding: "6px 12px", borderRadius: 6, fontSize: 12.5, cursor: "pointer", fontWeight: 600,
          border: `1px solid ${tm === current ? T.amber : T.border}`,
          background: tm === current ? T.amberDim : "transparent", color: tm === current ? T.amber : T.text,
        }}>{tm}</button>
      ))}
      {current && <button onClick={() => setOpen(false)} style={{ ...ghostBtn, padding: "5px 10px" }}><X size={12} /></button>}
    </Panel>
  );
}

/* ---------------------------------------------------------------
   UPLOAD FLOW
--------------------------------------------------------------- */
function UploadPanel({ rounds, onSaved }) {
  const [pending, setPending] = useState(null);
  const [roundNumber, setRoundNumber] = useState("");
  const [label, setLabel] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    if (!file) return;
    setErr(""); setBusy(true);
    try {
      const parsed = await parseFile(file);
      const isEmpty = parsed.kind === "cesim" ? !parsed.data : Object.keys(parsed.sheets).length === 0;
      if (isEmpty) { setErr("No se pudo detectar datos utilizables en el archivo."); setBusy(false); return; }
      setPending({ fileName: file.name, parsed });
      if (parsed.kind === "cesim") {
        setRoundNumber(parsed.data.roundNumber !== null ? String(parsed.data.roundNumber) : "");
        setLabel(parsed.data.label || "");
      } else {
        const nextN = rounds.length ? Math.max(...rounds.map((r) => r.roundNumber)) + 1 : 0;
        setRoundNumber(String(nextN)); setLabel("");
      }
    } catch (e) { setErr("No se pudo leer el archivo. Verificá que sea un .xls, .xlsx o .csv válido."); }
    setBusy(false);
  };

  const confirmSave = async () => {
    if (roundNumber === "" || Number.isNaN(Number(roundNumber))) { setErr("Ingresá un número de ronda válido."); return; }
    const dup = rounds.find((r) => r.roundNumber === Number(roundNumber));
    setBusy(true); setErr("");
    try {
      const id = dup ? dup.id : `r${Date.now()}`;
      const parsed = pending.parsed;
      const isCesim = parsed.kind === "cesim";
      await saveRoundData(id, isCesim ? { kind: "cesim", teams: parsed.data.teams, blocks: parsed.data.blocks } : { kind: "generic", sheets: parsed.sheets });
      const meta = await loadRoundsMeta();
      const others = meta.filter((m) => m.id !== id);
      const newEntry = {
        id, roundNumber: Number(roundNumber), label: label.trim() || `Ronda ${roundNumber}`,
        uploadedAt: new Date().toISOString(), kind: isCesim ? "cesim" : "generic",
        sheetNames: isCesim ? [] : Object.keys(parsed.sheets), teams: isCesim ? parsed.data.teams : [], fileName: pending.fileName,
      };
      const newMeta = [...others, newEntry].sort((a, b) => a.roundNumber - b.roundNumber);
      await saveRoundsMeta(newMeta);
      setPending(null); setLabel("");
      onSaved(newMeta);
    } catch (e) { setErr("Ocurrió un error al guardar. Probá de nuevo."); }
    setBusy(false);
  };

  if (pending) {
    const isCesim = pending.parsed.kind === "cesim";
    const dup = rounds.find((r) => r.roundNumber === Number(roundNumber));
    return (
      <Panel style={{ padding: 20 }}>
        <Eyebrow>Confirmar carga</Eyebrow>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6, flexWrap: "wrap" }}>
          <CheckCircle2 size={16} color={T.green} />
          <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: T.text }}>{pending.fileName}</span>
          {isCesim ? (
            <span style={{ color: T.textFaint, fontSize: 12 }}>· reporte Cesim detectado · {pending.parsed.data.teams.length} equipos · {pending.parsed.data.blocks.length} tablas</span>
          ) : (
            <span style={{ color: T.textFaint, fontSize: 12 }}>· {Object.keys(pending.parsed.sheets).length} hoja(s): {Object.keys(pending.parsed.sheets).join(", ")}</span>
          )}
        </div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
          <div>
            <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Número de ronda</div>
            <input value={roundNumber} onChange={(e) => setRoundNumber(e.target.value)} style={inputStyle} placeholder="0" />
          </div>
          <div style={{ flex: 1, minWidth: 180 }}>
            <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Etiqueta</div>
            <input value={label} onChange={(e) => setLabel(e.target.value)} style={{ ...inputStyle, width: "100%" }} placeholder={`Ronda ${roundNumber || ""}`} />
          </div>
        </div>
        {dup && <div style={{ display: "flex", gap: 6, alignItems: "center", color: T.amber, fontSize: 12, marginBottom: 10 }}><AlertTriangle size={13} /> Ya existe una ronda {roundNumber} ({dup.label}) — al guardar se va a reemplazar.</div>}
        {err && <div style={{ color: T.red, fontSize: 12.5, marginBottom: 10 }}>{err}</div>}
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={confirmSave} disabled={busy} style={primaryBtn}>{busy ? <Loader2 size={14} className="spin" /> : <CheckCircle2 size={14} />} Guardar ronda</button>
          <button onClick={() => { setPending(null); setErr(""); }} style={ghostBtn}><X size={14} /> Cancelar</button>
        </div>
      </Panel>
    );
  }

  return (
    <Panel
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files?.[0]); }}
      style={{ padding: 24, borderStyle: "dashed", borderColor: dragOver ? T.amber : T.border, background: dragOver ? T.amberDim : T.panel, transition: "all 0.15s ease", cursor: "pointer", textAlign: "center" }}
      onClick={() => inputRef.current?.click()}
    >
      <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv,.tsv" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files?.[0])} />
      {busy ? <Loader2 size={22} color={T.amber} className="spin" /> : <UploadCloud size={22} color={T.amber} />}
      <div style={{ marginTop: 10, color: T.text, fontSize: 14, fontWeight: 500 }}>Arrastrá el export de Cesim (.xls) acá, o hacé click para elegirlo</div>
      <div style={{ marginTop: 4, color: T.textFaint, fontSize: 12 }}>Reconoce el reporte "Results" completo — una ronda por archivo</div>
      {err && <div style={{ color: T.red, fontSize: 12.5, marginTop: 10 }}>{err}</div>}
    </Panel>
  );
}

/* ---------------------------------------------------------------
   ROUND TICKER
--------------------------------------------------------------- */
function RoundTicker({ rounds, onDelete }) {
  if (rounds.length === 0) return null;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 0, overflowX: "auto", padding: "4px 0" }}>
      {rounds.map((r, i) => (
        <React.Fragment key={r.id}>
          <div className="round-chip" style={{ display: "flex", alignItems: "center", gap: 8, background: T.panelAlt, border: `1px solid ${T.borderSoft}`, borderRadius: 8, padding: "6px 10px", flexShrink: 0 }}>
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: T.amber, fontWeight: 600 }}>R{r.roundNumber}</span>
            <span style={{ fontSize: 12, color: T.textDim, whiteSpace: "nowrap" }}>{r.label}</span>
            <button onClick={() => onDelete(r)} className="chip-del" style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", display: "flex", padding: 2 }} title="Eliminar ronda"><Trash2 size={12} /></button>
          </div>
          {i < rounds.length - 1 && <div style={{ width: 20, height: 1, background: T.borderSoft, flexShrink: 0 }} />}
        </React.Fragment>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------------
   KPIs
--------------------------------------------------------------- */
const KPI_DEFS = [
  { key: "revenue", label: "Ingresos por ventas", blockTitle: "Cuenta de resultados, miles USD, Global" },
  { key: "profit", label: "Beneficio de la ronda", blockTitle: "Cuenta de resultados, miles USD, Global" },
  { key: "cash", label: "Efectivo y equivalentes de efectivo", blockTitle: "Hoja de Balance, miles USD, Global" },
  { key: "share", label: "Total", blockTitle: "Informe de mercado, global" },
  { key: "marketcap", label: "Capitalización de mercado, miles USD", blockTitle: "Valuación - Global" },
];
function findMetricRow(block, label) {
  if (!block) return null;
  return block.rows.find((r) => r.kind === "metric" && r.label === label) || null;
}
function extractKpi(roundData, teamIdx, def) {
  if (!roundData || roundData.kind !== "cesim") return null;
  const block = roundData.blocks.find((b) => b.title === def.blockTitle);
  const row = findMetricRow(block, def.label);
  if (!row) return null;
  const v = row.values[teamIdx];
  return typeof v === "number" ? v : toNumberOrNull(v);
}
const PROMOTION_KPI_DEF = { key: "promotion", label: "Promoción", blockTitle: "Cuenta de resultados, miles USD, Global" };
// Best-effort: Cesim's regional market blocks don't always carry a
// promotion line — this returns null (not a fabricated 0) when absent,
// so the widget can say plainly that regional data isn't available.
function extractRegionalPromotion(roundData, region, teamIdx) {
  const block = roundData.blocks.find((b) => b.title === `Informe de mercado, ${region}`);
  if (!block) return null;
  const row = block.rows.find((r) => r.kind === "metric" && /promoci[oó]n/i.test(r.label));
  if (!row) return null;
  return toNumberOrNull(row.values[teamIdx]);
}
function computePromotionVariation(cur, prev) {
  if (!cur || cur.kind !== "cesim") return [];
  return cur.teams.map((tm, i) => {
    const now = extractKpi(cur, i, PROMOTION_KPI_DEF);
    const before = prev ? extractKpi(prev, i, PROMOTION_KPI_DEF) : null;
    const pctChange = now !== null && before !== null && before !== 0 ? ((now - before) / Math.abs(before)) * 100 : null;
    return { team: tm, now, before, pctChange };
  }).filter((r) => r.now !== null);
}

/* ---------------------------------------------------------------
   STRATEGIC PROXIMITY
--------------------------------------------------------------- */
const STRATEGY_VARS = [
  { key: "price", label: "Precio, % del promedio (Global)" },
  { key: "features", label: "Características, prom." },
  { key: "rd", label: "I+D" },
  { key: "promotion", label: "Promoción" },
  { key: "share", label: "Cuota de mercado global" },
];
function getTecMetric(block, label, teamIdx) {
  if (!block) return null;
  const rows = block.rows.filter((r) => r.kind === "metric" && r.label === label);
  for (const r of rows) { const n = toNumberOrNull(r.values[teamIdx]); if (n !== null) return n; }
  return null;
}
function getTecMetricPriceAnyCurrency(block, teamIdx) {
  if (!block) return null;
  const rows = block.rows.filter((r) => r.kind === "metric" && r.label.startsWith("Precio de venta,"));
  for (const r of rows) { const n = toNumberOrNull(r.values[teamIdx]); if (n !== null) return n; }
  return null;
}
const REGION_CURRENCY = { "EE.UU.": "USD", "Asia": "RMB", "Europa": "EUR" };

function computeRegionPriceIndex(roundData, region, teamIdx) {
  const block = roundData.blocks.find((b) => b.title === `Informe de mercado, ${region}`);
  if (!block) return null;
  const prices = roundData.teams.map((_, i) => getTecMetricPriceAnyCurrency(block, i)).filter((p) => p !== null);
  if (!prices.length) return null;
  const avg = prices.reduce((a, b) => a + b, 0) / prices.length;
  const own = getTecMetricPriceAnyCurrency(block, teamIdx);
  if (own === null || avg === 0) return null;
  return (own / avg) * 100; // % of that region's average price — currency-safe
}
// A team's price relative to the market average, averaged across the 3
// regions — this is what "Global" price should mean, since raw prices in
// USD/RMB/EUR can't be averaged directly.
function computeGlobalPriceIndex(roundData, teamIdx) {
  const idxs = REGIONS.map((r) => computeRegionPriceIndex(roundData, r, teamIdx)).filter((v) => v !== null);
  if (!idxs.length) return null;
  return idxs.reduce((a, b) => a + b, 0) / idxs.length;
}

function computeStrategicProfile(roundData, teamIdx) {
  if (!roundData || roundData.kind !== "cesim") return null;
  const marketBlocks = ["Informe de mercado, EE.UU.", "Informe de mercado, Asia", "Informe de mercado, Europa"]
    .map((t) => roundData.blocks.find((b) => b.title === t)).filter(Boolean);
  const avgAcrossMarkets = (label) => {
    const vals = marketBlocks.map((b) => getTecMetric(b, label, teamIdx)).filter((v) => v !== null);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };
  const plBlock = roundData.blocks.find((b) => b.title === "Cuenta de resultados, miles USD, Global");
  const mktBlock = roundData.blocks.find((b) => b.title === "Informe de mercado, global");
  const rdRow = findMetricRow(plBlock, "I+D");
  const promoRow = findMetricRow(plBlock, "Promoción");
  const shareRow = findMetricRow(mktBlock, "Total");
  return {
    price: computeGlobalPriceIndex(roundData, teamIdx),
    features: avgAcrossMarkets("Cantidad de características ofrecidas"),
    rd: rdRow ? toNumberOrNull(rdRow.values[teamIdx]) : null,
    promotion: promoRow ? toNumberOrNull(promoRow.values[teamIdx]) : null,
    share: shareRow ? toNumberOrNull(shareRow.values[teamIdx]) : null,
  };
}
function computeDistances(roundData, ourIdx) {
  if (!roundData || roundData.kind !== "cesim" || ourIdx < 0) return [];
  const teams = roundData.teams;
  const profiles = teams.map((_, i) => computeStrategicProfile(roundData, i));
  const keys = STRATEGY_VARS.map((v) => v.key);
  const stats = {};
  keys.forEach((k) => {
    const vals = profiles.map((p) => p?.[k]).filter((v) => v !== null && v !== undefined);
    const mean = vals.reduce((a, b) => a + b, 0) / (vals.length || 1);
    const variance = vals.reduce((a, b) => a + (b - mean) ** 2, 0) / (vals.length || 1);
    stats[k] = { mean, std: Math.sqrt(variance) || 1 };
  });
  const ourProfile = profiles[ourIdx];
  return teams.map((tm, i) => {
    if (i === ourIdx) return null;
    const p = profiles[i];
    let sumSq = 0, n = 0;
    keys.forEach((k) => {
      const a = ourProfile?.[k], b = p?.[k];
      if (a === null || a === undefined || b === null || b === undefined) return;
      const za = (a - stats[k].mean) / stats[k].std, zb = (b - stats[k].mean) / stats[k].std;
      sumSq += (za - zb) ** 2; n++;
    });
    return { team: tm, distance: n > 0 ? Math.sqrt(sumSq / n) : null, profile: p, coverage: n };
  }).filter(Boolean);
}

/* ---------------------------------------------------------------
   METRIC CATALOG
--------------------------------------------------------------- */
function buildMetricCatalog(rounds, dataById) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  if (!cesimRounds.length) return [];
  const ref = dataById[cesimRounds[cesimRounds.length - 1].id];
  const catalog = [];
  ref.blocks.forEach((block, bi) => {
    let lastGroup = null;
    block.rows.forEach((row, ri) => {
      if (row.kind === "group") { lastGroup = row.label; return; }
      catalog.push({ blockIndex: bi, rowIndex: ri, blockTitle: block.title, category: block.category, group: lastGroup, label: row.label, key: `${bi}::${ri}` });
    });
  });
  return catalog;
}

/* ---------------------------------------------------------------
   RISK RATIOS + ALERTS
--------------------------------------------------------------- */
const RATIOS_BLOCK_TITLE = "Ratios e indicadores financieros clave";
const CREDIT_RATING_ORDER = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"];

function getRatio(roundData, label, teamIdx) {
  if (!roundData || roundData.kind !== "cesim") return null;
  const block = roundData.blocks.find((b) => b.title === RATIOS_BLOCK_TITLE);
  const row = findMetricRow(block, label);
  if (!row) return null;
  const v = row.values[teamIdx];
  return typeof v === "number" || toNumberOrNull(v) !== null ? toNumberOrNull(v) : v;
}

function computeAlerts(cur, prev, teamIdx) {
  if (!cur || cur.kind !== "cesim") return [];
  const alerts = [];
  const profit = extractKpi(cur, teamIdx, KPI_DEFS.find((d) => d.key === "profit"));
  const cash = extractKpi(cur, teamIdx, KPI_DEFS.find((d) => d.key === "cash"));
  const share = extractKpi(cur, teamIdx, KPI_DEFS.find((d) => d.key === "share"));
  const leverage = getRatio(cur, "Endeudamiento neto/patrimonio (apalancamiento)", teamIdx);
  const roe = getRatio(cur, "Rendimiento de los Fondos Propios (ROE)", teamIdx);
  const rating = getRatio(cur, "Calificación crediticia", teamIdx);

  if (profit !== null && profit < 0) alerts.push({ level: "danger", text: `Beneficio negativo esta ronda (${fmtNumber(profit)} miles USD).` });
  if (typeof roe === "number" && roe < 0) alerts.push({ level: "danger", text: `ROE negativo (${roe.toFixed(1)}%).` });
  if (typeof leverage === "number" && leverage > 40) alerts.push({ level: "warning", text: `Apalancamiento alto — deuda neta/patrimonio en ${leverage.toFixed(1)}%.` });
  if (prev) {
    const cashPrev = extractKpi(prev, teamIdx, KPI_DEFS.find((d) => d.key === "cash"));
    if (cash !== null && cashPrev !== null && cashPrev !== 0) {
      const pct = (cash - cashPrev) / Math.abs(cashPrev);
      if (pct < -0.25) alerts.push({ level: "warning", text: `La caja cayó ${(pct * 100).toFixed(0)}% respecto a la ronda anterior.` });
    }
    const sharePrev = extractKpi(prev, teamIdx, KPI_DEFS.find((d) => d.key === "share"));
    if (share !== null && sharePrev !== null && share - sharePrev < -2) alerts.push({ level: "warning", text: `La cuota de mercado global cayó ${(sharePrev - share).toFixed(1)} puntos.` });
    const ratingPrev = getRatio(prev, "Calificación crediticia", teamIdx);
    if (typeof rating === "string" && typeof ratingPrev === "string") {
      const curRank = CREDIT_RATING_ORDER.indexOf(rating), prevRank = CREDIT_RATING_ORDER.indexOf(ratingPrev);
      if (curRank > -1 && prevRank > -1 && curRank > prevRank) alerts.push({ level: "warning", text: `La calificación crediticia bajó de ${ratingPrev} a ${rating}.` });
    }
  }
  return alerts;
}

function computeRank(roundData, def, ourIdx) {
  if (!roundData || roundData.kind !== "cesim" || ourIdx < 0) return null;
  const values = roundData.teams.map((_, i) => extractKpi(roundData, i, def));
  const withValues = values.map((v, i) => ({ i, v })).filter((x) => x.v !== null);
  if (!withValues.length) return null;
  withValues.sort((a, b) => b.v - a.v);
  const pos = withValues.findIndex((x) => x.i === ourIdx);
  return pos === -1 ? null : { rank: pos + 1, of: withValues.length };
}

function computeMargin(roundData, teamIdx) {
  const revenue = extractKpi(roundData, teamIdx, KPI_DEFS.find((d) => d.key === "revenue"));
  const profit = extractKpi(roundData, teamIdx, KPI_DEFS.find((d) => d.key === "profit"));
  if (revenue === null || profit === null || revenue === 0) return null;
  return (profit / revenue) * 100;
}
function computeMarginRank(roundData, ourIdx) {
  if (!roundData || roundData.kind !== "cesim" || ourIdx < 0) return null;
  const values = roundData.teams.map((_, i) => computeMargin(roundData, i));
  const withValues = values.map((v, i) => ({ i, v })).filter((x) => x.v !== null);
  if (!withValues.length) return null;
  withValues.sort((a, b) => b.v - a.v);
  const pos = withValues.findIndex((x) => x.i === ourIdx);
  return pos === -1 ? null : { rank: pos + 1, of: withValues.length };
}

/* ---------------------------------------------------------------
   TECH MIX
--------------------------------------------------------------- */
function getMetricUnderGroup(block, groupLabel, metricLabel, teamIdx) {
  if (!block) return null;
  const rows = block.rows;
  const groupIdx = rows.findIndex((r) => r.kind === "group" && r.label === groupLabel);
  if (groupIdx === -1) return null;
  for (let i = groupIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.kind === "group") break;
    if (r.kind === "metric" && r.label === metricLabel) return toNumberOrNull(r.values[teamIdx]);
  }
  return null;
}
// Same lookup, but for TEXT fields (e.g. "Enfoque de la estrategia de
// marketing") — getMetricUnderGroup forces a number conversion and would
// always return null for a string like "Marca".
function getRawMetricUnderGroup(block, groupLabel, metricLabel, teamIdx) {
  if (!block) return null;
  const rows = block.rows;
  const groupIdx = rows.findIndex((r) => r.kind === "group" && r.label === groupLabel);
  if (groupIdx === -1) return null;
  for (let i = groupIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.kind === "group") break;
    if (r.kind === "metric" && r.label === metricLabel) {
      const v = r.values[teamIdx];
      return v !== null && v !== undefined && v !== "" ? v : null;
    }
  }
  return null;
}
const TECH_LABELS = ["Tec 1", "Tec 2", "Tec 3", "Tec 4"];
const REGIONS = ["EE.UU.", "Asia", "Europa"];
const REGION_SCOPES = ["Global", ...REGIONS];

/* ---------------------------------------------------------------
   REGION FILTER — visible on every tab; scopes market/price/tech-mix
   views to a market. Company-wide financials (P&L, Balance, RRHH, ESG)
   have no per-region breakdown in Cesim and stay Global regardless.
--------------------------------------------------------------- */
function RegionFilter({ value, onChange }) {
  return (
    <div style={{ display: "flex", gap: 4, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 10, padding: 4 }}>
      {REGION_SCOPES.map((r) => (
        <button key={r} onClick={() => onChange(r)} style={{
          padding: "6px 13px", borderRadius: 7, fontSize: 12.5, fontWeight: 600, cursor: "pointer", border: "none",
          background: value === r ? T.amber : "transparent", color: value === r ? "#FFFFFF" : T.textDim,
          transition: "all 0.12s ease",
        }}>{r}</button>
      ))}
    </div>
  );
}
function regionMarketBlockTitle(region) {
  return region === "Global" ? "Informe de mercado, global" : `Informe de mercado, ${region}`;
}
function computeTechMix(roundData, teamIdx) {
  if (!roundData || roundData.kind !== "cesim") return null;
  const totals = [0, 0, 0, 0];
  REGIONS.forEach((region) => {
    const block = roundData.blocks.find((b) => b.title === `Informe de mercado, ${region}`);
    if (!block) return;
    TECH_LABELS.forEach((tech, i) => {
      const v = getMetricUnderGroup(block, tech, "Ventas, miles unidades", teamIdx);
      if (v !== null) totals[i] += v;
    });
  });
  const sum = totals.reduce((a, b) => a + b, 0);
  if (sum <= 0) return null;
  return TECH_LABELS.map((t, i) => ({ tech: t, units: totals[i], pct: (totals[i] / sum) * 100 })).filter((t) => t.units > 0);
}

/* ---------------------------------------------------------------
   STRATEGY PLAN
--------------------------------------------------------------- */
const MARKETING_FOCUS_OPTIONS = ["Precio bajo", "Características", "Marca", "Equilibrado"];

function getBlockMetricRaw(block, label, teamIdx) {
  if (!block) return null;
  const rows = block.rows.filter((r) => r.kind === "metric" && r.label === label);
  for (const r of rows) { const v = r.values[teamIdx]; if (v !== null && v !== undefined && v !== "") return v; }
  return null;
}

function computeStrategyAlignment(plan, roundData, teamIdx) {
  if (!plan || !roundData || roundData.kind !== "cesim" || teamIdx < 0) return [];
  const profile = computeStrategicProfile(roundData, teamIdx);
  const leverage = getRatio(roundData, "Endeudamiento neto/patrimonio (apalancamiento)", teamIdx);
  const checks = [];

  if (plan.rdMin != null || plan.rdMax != null) {
    const v = profile.rd;
    const ok = v !== null && (plan.rdMin == null || v >= plan.rdMin) && (plan.rdMax == null || v <= plan.rdMax);
    checks.push({ key: "rd", label: "I+D (compañía)", target: `${plan.rdMin ?? "–"} a ${plan.rdMax ?? "–"} miles USD`, actual: v !== null ? fmtNumber(v) : "—", ok, region: null });
  }
  if (plan.leverageMax != null) {
    const v = typeof leverage === "number" ? leverage : null;
    const ok = v !== null && v <= plan.leverageMax;
    checks.push({ key: "leverage", label: "Apalancamiento (compañía)", target: `≤ ${plan.leverageMax}%`, actual: v !== null ? `${v.toFixed(1)}%` : "—", ok, region: null });
  }

  REGIONS.forEach((region) => {
    const rp = plan.regions?.[region];
    if (!rp) return;
    const block = roundData.blocks.find((b) => b.title === `Informe de mercado, ${region}`);
    const currency = REGION_CURRENCY[region] || "";

    if (rp.shareMin != null) {
      const shareRow = findMetricRow(block, "Total");
      const v = shareRow ? toNumberOrNull(shareRow.values[teamIdx]) : null;
      const ok = v !== null && v >= rp.shareMin;
      checks.push({ key: `share-${region}`, label: `Cuota de mercado — ${region}`, target: `≥ ${rp.shareMin}%`, actual: v !== null ? `${v.toFixed(1)}%` : "—", ok, region });
    }

    TECH_LABELS.forEach((tech) => {
      const tp = rp.techs?.[tech];
      if (!tp) return;
      if (tp.priceMin != null || tp.priceMax != null) {
        const v = getMetricUnderGroupPriceAnyCurrency(block, tech, teamIdx);
        const ok = v !== null && (tp.priceMin == null || v >= tp.priceMin) && (tp.priceMax == null || v <= tp.priceMax);
        checks.push({ key: `price-${region}-${tech}`, label: `Precio — ${region} · ${tech} (${currency})`, target: `${tp.priceMin ?? "–"} a ${tp.priceMax ?? "–"}`, actual: v !== null ? v.toFixed(0) : "—", ok, region });
      }
      if (tp.marketingFocus) {
        const focus = getRawMetricUnderGroup(block, tech, "Enfoque de la estrategia de marketing", teamIdx);
        const ok = focus !== null && focus === tp.marketingFocus;
        checks.push({ key: `focus-${region}-${tech}`, label: `Enfoque — ${region} · ${tech}`, target: tp.marketingFocus, actual: focus ?? "—", ok, region });
      }
    });
  });

  return checks;
}

/* =================================================================
   GENERAL SECTION
================================================================= */
function GeneralSection({ rounds, dataById, ourTeam, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const teamIdx = ourTeam && cesimRounds.length ? cesimRounds[0].teams.indexOf(ourTeam) : -1;
  const latestRound = cesimRounds[cesimRounds.length - 1];
  const latestData = latestRound ? dataById[latestRound.id] : null;
  const prevRound = cesimRounds[cesimRounds.length - 2];
  const kpiDefs = useMemo(() => KPI_DEFS.map((d) => (d.key === "share" ? { ...d, blockTitle: regionMarketBlockTitle(regionFilter) } : d)), [regionFilter]);
  const prevData = prevRound ? dataById[prevRound.id] : null;

  const kpiSeries = useMemo(() => {
    if (teamIdx < 0) return {};
    const series = {};
    kpiDefs.forEach((def) => {
      const points = [];
      cesimRounds.forEach((r) => { const v = extractKpi(dataById[r.id], teamIdx, def); if (v !== null) points.push({ roundNumber: r.roundNumber, value: v }); });
      if (points.length) series[def.key] = { header: def.label, points };
    });
    return series;
  }, [cesimRounds, dataById, teamIdx, kpiDefs]);

  const alerts = useMemo(() => (teamIdx < 0 ? [] : computeAlerts(latestData, prevData, teamIdx)), [latestData, prevData, teamIdx]);

  const [plan, setPlan] = useState(null);
  useEffect(() => { loadStrategyPlan().then(setPlan); }, []);
  const alignmentChecks = useMemo(() => (plan && teamIdx >= 0 ? computeStrategyAlignment(plan, latestData, teamIdx) : []), [plan, latestData, teamIdx]);
  const alignmentAlerts = alignmentChecks.filter((c) => !c.ok).map((c) => ({ level: "warning", text: `Fuera de estrategia — ${c.label}: objetivo ${c.target}, actual ${c.actual}.`, strategic: true }));
  const allAlerts = [...alerts, ...alignmentAlerts];

  const marginNow = teamIdx >= 0 && latestData ? computeMargin(latestData, teamIdx) : null;
  const marginPrev = teamIdx >= 0 && prevData ? computeMargin(prevData, teamIdx) : null;
  const marginRank = teamIdx >= 0 && latestData ? computeMarginRank(latestData, teamIdx) : null;

  if (rounds.length === 0) {
    return (
      <div style={{ textAlign: "center", padding: "60px 20px", color: T.textFaint }}>
        <Radio size={28} style={{ marginBottom: 10, opacity: 0.5 }} />
        <div style={{ fontSize: 14 }}>Todavía no hay rondas cargadas.</div>
        <div style={{ fontSize: 12.5, marginTop: 4 }}>Subí el primer export de Cesim para empezar el panel.</div>
      </div>
    );
  }
  if (teamIdx < 0) return <div style={{ textAlign: "center", padding: "40px 20px", color: T.textFaint }}>Elegí cuál es su equipo (arriba) para ver los indicadores.</div>;

  const kpiKeys = Object.keys(kpiSeries);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      {allAlerts.length > 0 && <AlertBanner alerts={allAlerts} />}

      {regionFilter !== "Global" && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: T.textDim, background: T.cyanDim, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px" }}>
          <Info size={13} color={T.cyan} style={{ flexShrink: 0 }} />
          Mostrando cuota de mercado de <strong style={{ color: T.text }}>{regionFilter}</strong>. Ingresos, beneficio, caja y capitalización son siempre a nivel compañía — Cesim no reporta esos estados por región.
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12 }}>
        <StatCard label="Rondas cargadas" value={rounds.length} />
        <StatCard label="Última ronda" value={`R${rounds[rounds.length - 1].roundNumber}`} />
        {["revenue", "profit"].map((k) => {
          const s = kpiSeries[k]; if (!s) return null;
          const last = s.points[s.points.length - 1], prev = s.points[s.points.length - 2];
          const delta = prev ? last.value - prev.value : null;
          const def = kpiDefs.find((d) => d.key === k);
          const rank = latestData ? computeRank(latestData, def, teamIdx) : null;
          return <StatCard key={k} label={s.header} value={fmtNumber(last.value)} delta={delta} rank={rank} />;
        })}
        {marginNow !== null && <StatCard label="Margen sobre ventas" value={`${marginNow.toFixed(1)}%`} delta={marginPrev !== null ? marginNow - marginPrev : null} rank={marginRank} />}
        {["cash", "share", "marketcap"].map((k) => {
          const s = kpiSeries[k]; if (!s) return null;
          const last = s.points[s.points.length - 1], prev = s.points[s.points.length - 2];
          const delta = prev ? last.value - prev.value : null;
          const def = kpiDefs.find((d) => d.key === k);
          const rank = latestData ? computeRank(latestData, def, teamIdx) : null;
          return <StatCard key={k} label={s.header} value={fmtNumber(last.value)} delta={delta} rank={rank} />;
        })}
      </div>

      {kpiKeys.length > 0 ? (
        <Panel style={{ padding: 18 }}>
          <Eyebrow info="Evolución de los indicadores clave de su equipo, ronda a ronda.">Indicadores de {ourTeam} a través de las rondas</Eyebrow>
          <div style={{ height: 280 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart>
                <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
                <XAxis dataKey="roundNumber" type="number" domain={["dataMin", "dataMax"]} allowDecimals={false} stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} label={{ value: "Ronda", position: "insideBottom", offset: -4, fill: T.textFaint, fontSize: 11 }} />
                <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
                <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} labelFormatter={(v) => `Ronda ${v}`} formatter={tooltipNumberFormatter} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {kpiKeys.map((k, i) => <Line key={k} data={kpiSeries[k].points} dataKey="value" name={kpiSeries[k].header} stroke={SERIES_COLORS[i % SERIES_COLORS.length]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        </Panel>
      ) : (
        <Panel style={{ padding: 18, color: T.textFaint, fontSize: 13 }}>No se encontraron los indicadores estándar en las rondas cargadas todavía.</Panel>
      )}

      <RoundNotes rounds={rounds} />

      <Panel style={{ padding: 18 }}>
        <Eyebrow>Historial de cargas</Eyebrow>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {[...rounds].reverse().map((r) => (
            <div key={r.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: T.textDim, borderBottom: `1px solid ${T.borderSoft}`, paddingBottom: 6, gap: 10, flexWrap: "wrap" }}>
              <span style={{ color: T.text }}>R{r.roundNumber} · {r.label}</span>
              <span>{r.kind === "cesim" ? `${r.teams?.length || 0} equipos` : r.sheetNames.join(", ")}</span>
              <span style={{ color: T.textFaint }}>{new Date(r.uploadedAt).toLocaleString("es-AR")}</span>
            </div>
          ))}
        </div>
      </Panel>
    </div>
  );
}

function DashboardManualAlerts() {
  const [story, setStory] = useState(null);
  useEffect(() => { loadDashboardStory().then((s) => setStory(s || DEFAULT_DASHBOARD_STORY)); }, []);
  const saveAlerts = async (alerts) => {
    const next = { ...(story || DEFAULT_DASHBOARD_STORY), alerts };
    await saveDashboardStory(next);
    setStory(next);
  };
  if (!story) return null;
  return (
    <Panel style={{ padding: 16, borderColor: T.red, background: "#FFF6F6" }}>
      <Eyebrow info="Puntos de atención que el equipo carga a mano, además de las alertas automáticas de arriba.">Alertas — última ronda</Eyebrow>
      <StoryList items={story.alerts} onSave={saveAlerts} tone="danger" emptyLabel="Sin alertas manuales cargadas." />
    </Panel>
  );
}
function AlertBanner({ alerts }) {
  return (
    <Panel style={{ padding: 16, borderColor: alerts.some((a) => a.level === "danger") ? T.red : T.amber }}>
      <Eyebrow info="Riesgos financieros automáticos (caja, apalancamiento, calificación) y desvíos respecto al plan estratégico definido.">Alertas — última ronda</Eyebrow>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {alerts.map((a, i) => (
          <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 8, fontSize: 13, color: a.level === "danger" ? T.red : T.amber }}>
            {a.strategic ? <Target size={14} style={{ marginTop: 1, flexShrink: 0 }} /> : <AlertTriangle size={14} style={{ marginTop: 1, flexShrink: 0 }} />}
            <span>{a.text}</span>
          </div>
        ))}
      </div>
    </Panel>
  );
}

function RoundNotes({ rounds }) {
  const [selectedId, setSelectedId] = useState(rounds.length ? rounds[rounds.length - 1].id : null);
  const [text, setText] = useState(""); const [saved, setSaved] = useState(true); const [loading, setLoading] = useState(true);
  useEffect(() => { if (!selectedId) return; setLoading(true); loadNote(selectedId).then((t) => { setText(t); setSaved(true); setLoading(false); }); }, [selectedId]);
  useEffect(() => { if (rounds.length && !rounds.find((r) => r.id === selectedId)) setSelectedId(rounds[rounds.length - 1].id); }, [rounds.map((r) => r.id).join(",")]);
  const doSave = async () => { await saveNote(selectedId, text); setSaved(true); };
  return (
    <Panel style={{ padding: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
        <Eyebrow info="Bitácora compartida: registren qué decidieron cada ronda y por qué, para el informe final.">Notas de decisiones</Eyebrow>
        <select value={selectedId || ""} onChange={(e) => setSelectedId(e.target.value)} style={{ ...inputStyle, fontFamily: "'Inter', sans-serif" }}>
          {[...rounds].reverse().map((r) => <option key={r.id} value={r.id}>R{r.roundNumber} · {r.label}</option>)}
        </select>
      </div>
      <textarea value={text} onChange={(e) => { setText(e.target.value); setSaved(false); }} disabled={loading}
        placeholder="¿Qué decisiones tomaron esta ronda y por qué?"
        style={{ width: "100%", minHeight: 90, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, padding: 10, fontSize: 13, fontFamily: "'Inter', sans-serif", resize: "vertical", outline: "none" }} />
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
        <button onClick={doSave} disabled={saved} style={{ ...primaryBtn, opacity: saved ? 0.5 : 1, cursor: saved ? "default" : "pointer" }}><CheckCircle2 size={13} /> {saved ? "Guardado" : "Guardar nota"}</button>
      </div>
    </Panel>
  );
}

function buildExportWorkbook(rounds, dataById, ourTeam, teamIdx, latestRound) {
  const wb = XLSX.utils.book_new();
  const latestData = dataById[latestRound.id];
  const kpiRows = KPI_DEFS.map((def) => {
    const row = { Indicador: def.label };
    latestData.teams.forEach((tm, i) => { row[tm] = extractKpi(latestData, i, def); });
    return row;
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kpiRows), `KPIs R${latestRound.roundNumber}`);
  const distances = computeDistances(latestData, teamIdx).sort((a, b) => (a.distance ?? Infinity) - (b.distance ?? Infinity));
  const rivalRows = distances.map((d) => ({ Equipo: d.team, Distancia: d.distance !== null ? Number(d.distance.toFixed(3)) : null, ...Object.fromEntries(STRATEGY_VARS.map((v) => [v.label, d.profile?.[v.key] ?? null])) }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rivalRows), "Rivales cercanos");
  const historyRows = rounds.map((r) => ({ Ronda: r.roundNumber, Etiqueta: r.label, Cargado: r.uploadedAt }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(historyRows), "Historial de rondas");
  return wb;
}
function ExportButtons({ rounds, dataById, ourTeam, teamIdx, latestRound }) {
  const [busy, setBusy] = useState(false);
  if (!latestRound) return null;
  const exportExcel = () => {
    setBusy(true);
    try { XLSX.writeFile(buildExportWorkbook(rounds, dataById, ourTeam, teamIdx, latestRound), `cesim_${ourTeam}_R${latestRound.roundNumber}.xlsx`); }
    finally { setBusy(false); }
  };
  return (
    <>
      <button onClick={exportExcel} disabled={busy} style={ghostBtn}>{busy ? <Loader2 size={13} className="spin" /> : <FileSpreadsheet size={13} />} Exportar Excel</button>
      <button onClick={() => window.print()} style={ghostBtn}><Printer size={13} /> Exportar PDF</button>
    </>
  );
}
function StatCard({ label, value, delta, rank }) {
  return (
    <Panel style={{ padding: "14px 16px" }}>
      <div style={{ fontSize: 11, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 6 }}>{label}</div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: T.text }}>{value}</div>
        {Number.isFinite(delta) && <div style={{ fontSize: 12, color: delta >= 0 ? T.green : T.red, fontWeight: 600 }}>{delta >= 0 ? "▲" : "▼"} {fmtNumber(Math.abs(delta))}</div>}
        {rank && <div style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 11, color: rank.rank <= 2 ? T.green : rank.rank >= 6 ? T.red : T.textFaint, fontWeight: 600 }}><Trophy size={11} /> {rank.rank}°/{rank.of}</div>}
      </div>
    </Panel>
  );
}

const DEFAULT_VMS_DATA = {
  labelPrev: "Año 1", labelCur: "Año 2",
  totalPrev: 16.17, totalCur: 15.32,
  tec1Prev: 12.42, tec1Cur: 9.55,
  tec2Prev: 3.75, tec2Cur: 5.77,
  causes: [
    "Responsable puntual: PRELA hizo una guerra de precios en Asia Tec1 (-36% de precio), lo que le costó a ULTI 3,4 p.p. de cuota en ese mercado (el mayor golpe individual del análisis).",
    "NODO y NOSRE ganaron el boom de Tec2: capitalizaron mejor el crecimiento de +496% en Tec2, limitando cuánto pudo compensar ULTI el bache de Tec1.",
    "Efecto estructural de mix: Tec2 creció ~6x más rápido que Tec1 a nivel industria. Con 62% de sus ingresos aún en Tec1, ULTI queda en desventaja relativa.",
  ],
  breakdownInfo: "Tec2 hizo lo que se esperaba de la estrategia: su negocio más que se duplicó (US$350M → US$710M) y sumó +2 p.p. Pero Tec1 sigue siendo el 62% de nuestros ingresos, y su caída de casi -3 p.p. pesó más que esa ganancia.",
  news: [
    {
      title: "Asia Tec1: la guerra de precios de PRELA",
      impact: "negativo",
      stat: { value: "-3,4 p.p.", label: "de cuota de valor perdida por ULTI en ese único mercado" },
      body: "**PRELA bajó precio 36%** y casi duplicó volumen. El mayor golpe individual de todo el análisis: **+6,2 p.p.** de cuota de valor ganada por PRELA en Asia Tec1 — más del doble que cualquier otro movimiento.",
    },
    {
      title: "EE.UU. Tec1: un competidor que se retiró solo",
      impact: "positivo",
      stat: { value: "-70%", label: "de caída de volumen para MECHA en EE.UU. Tec1" },
      body: "**MECHA subió precio 43%** sin respaldo de marca y perdió volumen. ULTI ganó **+2,92 p.p.** repartiendo el espacio libre con NOSRE, NODO y LUCKY — no fue mérito de una pelea directa.",
    },
    {
      title: "Tec2: ¿el 80% del año 1 era un espejismo?",
      impact: "neutro",
      stat: { value: "+496%", label: "creció el mercado de Tec2 — aparecieron rivales reales" },
      body: "En el año 1 el liderazgo de ULTI en Tec2 era por ausencia de competencia, no por mérito. En el año 2 el mercado maduró: **el negocio Tec2 de ULTI casi se duplicó en dólares** (US$350M → US$710M) pese a la caída de cuota relativa.",
    },
  ],
};

function NewsForm({ initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || { title: "", impact: "neutro", statValue: "", statLabel: "", body: "" });
  const [saving, setSaving] = useState(false);
  const save = async () => { if (!form.title.trim()) return; setSaving(true); await onSave({ title: form.title, impact: form.impact, body: form.body, stat: (form.statValue || form.statLabel) ? { value: form.statValue, label: form.statLabel } : null }); setSaving(false); };
  return (
    <Panel style={{ padding: 16 }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <input value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} placeholder="Título de la novedad" style={{ ...inputStyle, width: "100%", fontFamily: "'Inter', sans-serif" }} />
        <div style={{ display: "flex", gap: 6 }}>
          {[["positivo", "Nos benefició", T.green], ["negativo", "Nos perjudicó", T.red], ["neutro", "Neutro", T.textFaint]].map(([key, lbl, color]) => (
            <button key={key} onClick={() => setForm({ ...form, impact: key })} style={{ padding: "6px 12px", borderRadius: 6, fontSize: 12, cursor: "pointer", fontWeight: 600, border: `1px solid ${form.impact === key ? color : T.border}`, background: form.impact === key ? color + "18" : "transparent", color: form.impact === key ? color : T.text }}>{lbl}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input value={form.statValue} onChange={(e) => setForm({ ...form, statValue: e.target.value })} placeholder="Número destacado (ej: -3,4 p.p.)" style={{ ...inputStyle, width: 160, fontFamily: "'Inter', sans-serif" }} />
          <input value={form.statLabel} onChange={(e) => setForm({ ...form, statLabel: e.target.value })} placeholder="Descripción del número" style={{ ...inputStyle, flex: 1, fontFamily: "'Inter', sans-serif" }} />
        </div>
        <textarea value={form.body} onChange={(e) => setForm({ ...form, body: e.target.value })} placeholder="Qué pasó — usá **texto** para resaltar en negrita"
          style={{ width: "100%", minHeight: 70, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, padding: 10, fontSize: 13, fontFamily: "'Inter', sans-serif", resize: "vertical", outline: "none" }} />
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={save} disabled={saving || !form.title.trim()} style={primaryBtn}>{saving ? <Loader2 size={13} className="spin" /> : <CheckCircle2 size={13} />} Guardar</button>
          <button onClick={onCancel} style={ghostBtn}><X size={13} /> Cancelar</button>
        </div>
      </div>
    </Panel>
  );
}
const NEWS_IMPACT_STYLE = {
  positivo: { color: T.green, bg: "#EAFBF0", border: "#CDEFDA", label: "Nos benefició" },
  negativo: { color: T.red, bg: "#FFF1F1", border: "#FBD5D5", label: "Nos perjudicó" },
  neutro: { color: T.textFaint, bg: T.panelAlt, border: T.border, label: "Neutro" },
};
function NewsSection({ news, onSave }) {
  const [adding, setAdding] = useState(false);
  const [editingIdx, setEditingIdx] = useState(null);
  const [confirmDeleteIdx, setConfirmDeleteIdx] = useState(null);
  const addItem = async (item) => { await onSave([...news, item]); setAdding(false); };
  const updateItem = async (idx, item) => { await onSave(news.map((n, i) => (i === idx ? item : n))); setEditingIdx(null); };
  const deleteItem = async (idx) => { await onSave(news.filter((_, i) => i !== idx)); setConfirmDeleteIdx(null); };

  return (
    <Panel style={{ padding: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <Eyebrow info="Qué pasó en el entorno competitivo esta ronda — movimientos de otros equipos que terminaron beneficiando o perjudicando a ULTI.">Novedades del entorno</Eyebrow>
        {!adding && <button onClick={() => setAdding(true)} style={{ ...ghostBtn, padding: "5px 11px", fontSize: 11.5 }}><BookPlus size={11} /> Agregar novedad</button>}
      </div>
      {adding && <div style={{ marginBottom: 12 }}><NewsForm onSave={addItem} onCancel={() => setAdding(false)} /></div>}
      {news.length === 0 && !adding && <div style={{ color: T.textFaint, fontSize: 12.5 }}>Sin novedades cargadas.</div>}
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {news.map((n, i) => {
          if (editingIdx === i) return <NewsForm key={i} initial={{ title: n.title, impact: n.impact, statValue: n.stat?.value || "", statLabel: n.stat?.label || "", body: n.body }} onSave={(item) => updateItem(i, item)} onCancel={() => setEditingIdx(null)} />;
          const style = NEWS_IMPACT_STYLE[n.impact] || NEWS_IMPACT_STYLE.neutro;
          return (
            <div key={i} style={{ border: `1px solid ${style.border}`, background: style.bg, borderRadius: 10, padding: "14px 16px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{n.title}</span>
                  <span style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: style.color, background: T.panel, border: `1px solid ${style.border}`, borderRadius: 20, padding: "2px 9px" }}>{style.label}</span>
                </div>
                <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                  <button onClick={() => setEditingIdx(i)} style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", padding: 3 }}><Pencil size={12} /></button>
                  <button onClick={() => setConfirmDeleteIdx(i)} style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", padding: 3 }}><Trash2 size={12} /></button>
                </div>
              </div>
              {n.stat && (
                <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 6 }}>
                  <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: style.color }}>{n.stat.value}</span>
                  <span style={{ fontSize: 12, color: T.textDim }}>{n.stat.label}</span>
                </div>
              )}
              <div style={{ fontSize: 13, color: T.text, lineHeight: 1.55 }}>{renderBoldText(n.body)}</div>
            </div>
          );
        })}
      </div>
      {confirmDeleteIdx !== null && (
        <div style={{ position: "fixed", inset: 0, background: T.overlay, display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }}>
          <Panel style={{ padding: 22, maxWidth: 320 }}>
            <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 8 }}>Eliminar novedad</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => deleteItem(confirmDeleteIdx)} style={{ ...primaryBtn, background: T.red, color: "#fff" }}><Trash2 size={13} /> Eliminar</button>
              <button onClick={() => setConfirmDeleteIdx(null)} style={ghostBtn}>Cancelar</button>
            </div>
          </Panel>
        </div>
      )}
    </Panel>
  );
}

function ValueMarketShareSection({ rounds }) {
  const [data, setData] = useState(null);
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(null);
  const [saving, setSaving] = useState(false);

  useEffect(() => { loadVmsData().then((d) => setData(d || DEFAULT_VMS_DATA)); }, []);

  const startEdit = () => { setDraft(data); setEditing(true); };
  const save = async () => { setSaving(true); await saveVmsData(draft); setData(draft); setSaving(false); setEditing(false); };

  if (!data) return <div style={{ padding: 20, color: T.textFaint, display: "flex", alignItems: "center", gap: 8 }}><Loader2 size={14} className="spin" /> Cargando…</div>;

  const totalVar = data.totalCur - data.totalPrev;
  const tec1Var = data.tec1Cur - data.tec1Prev;
  const tec2Var = data.tec2Cur - data.tec2Prev;
  const trendData = [{ label: data.labelPrev, value: data.totalPrev }, { label: data.labelCur, value: data.totalCur }];
  const breakdownChartData = [
    { tech: "Tec1", variacion: Number(tec1Var.toFixed(2)) },
    { tech: "Tec2", variacion: Number(tec2Var.toFixed(2)) },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <Panel style={{ padding: 18 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
          <Eyebrow info="Cuota de mercado en VALOR (no en unidades) — mide qué % de los ingresos totales del mercado captura ULTI.">Cuota de valor de mercado (VMS)</Eyebrow>
          {!editing && <button onClick={startEdit} style={{ ...ghostBtn, padding: "5px 11px", fontSize: 11.5 }}><Pencil size={11} /> Editar datos</button>}
        </div>
        {!editing ? (
          <>
            <div style={{ display: "flex", alignItems: "baseline", gap: 14, marginBottom: 14 }}>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 32, fontWeight: 700, color: totalVar < 0 ? T.red : T.green }}>{data.totalCur.toFixed(2)}%</div>
              <div style={{ fontSize: 13, color: T.textDim }}>vs. {data.totalPrev.toFixed(2)}% en {data.labelPrev} <span style={{ fontWeight: 700, color: totalVar < 0 ? T.red : T.green }}>({totalVar >= 0 ? "+" : ""}{totalVar.toFixed(2)} p.p.)</span></div>
            </div>
            <div style={{ height: 160 }}>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={trendData}>
                  <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
                  <XAxis dataKey="label" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
                  <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} unit="%" />
                  <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v) => [`${v}%`, "VMS"]} />
                  <Line dataKey="value" stroke={T.red} strokeWidth={2.5} dot={{ r: 4 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </>
        ) : (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
            <div><div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Etiqueta período anterior</div><input value={draft.labelPrev} onChange={(e) => setDraft({ ...draft, labelPrev: e.target.value })} style={{ ...inputStyle, width: 110, fontFamily: "'Inter', sans-serif" }} /></div>
            <div><div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Etiqueta período actual</div><input value={draft.labelCur} onChange={(e) => setDraft({ ...draft, labelCur: e.target.value })} style={{ ...inputStyle, width: 110, fontFamily: "'Inter', sans-serif" }} /></div>
            <NumField label="VMS total anterior, %" value={draft.totalPrev} onChange={(v) => setDraft({ ...draft, totalPrev: v })} />
            <NumField label="VMS total actual, %" value={draft.totalCur} onChange={(v) => setDraft({ ...draft, totalCur: v })} />
            <NumField label="Tec1 anterior, p.p." value={draft.tec1Prev} onChange={(v) => setDraft({ ...draft, tec1Prev: v })} />
            <NumField label="Tec1 actual, p.p." value={draft.tec1Cur} onChange={(v) => setDraft({ ...draft, tec1Cur: v })} />
            <NumField label="Tec2 anterior, p.p." value={draft.tec2Prev} onChange={(v) => setDraft({ ...draft, tec2Prev: v })} />
            <NumField label="Tec2 actual, p.p." value={draft.tec2Cur} onChange={(v) => setDraft({ ...draft, tec2Cur: v })} />
          </div>
        )}
      </Panel>

      <Panel style={{ padding: 18 }}>
        <Eyebrow info="Interpretación cargada por el equipo — no se auto-calcula, ya que requiere leer el porqué detrás de los números.">¿Causantes de la caída?</Eyebrow>
        <StoryList
          items={(editing ? draft.causes : data.causes).map((c) => ({ text: c, info: null }))}
          onSave={async (items) => { const causes = items.map((i) => i.text); const next = { ...data, causes }; await saveVmsData(next); setData(next); }}
          emptyLabel="Sin causas cargadas."
        />
      </Panel>

      <Panel style={{ padding: 18 }}>
        <div style={{ display: "flex", alignItems: "center", marginBottom: 10 }}>
          <Eyebrow>Desglose de la caída del valor por tecnología</Eyebrow>
          <InfoBubble text={data.breakdownInfo} />
        </div>
        <div style={{ overflowX: "auto", marginBottom: 16 }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>
            <thead><tr>
              <th style={{ textAlign: "left", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}` }} />
              <th style={{ textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>{data.labelPrev}</th>
              <th style={{ textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>{data.labelCur}</th>
              <th style={{ textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>Var.</th>
            </tr></thead>
            <tbody>
              {[["Tec1", data.tec1Prev, data.tec1Cur, tec1Var], ["Tec2", data.tec2Prev, data.tec2Cur, tec2Var], ["Total", data.totalPrev, data.totalCur, totalVar]].map(([label, prev, cur, v], i) => (
                <tr key={label} style={{ background: i % 2 ? T.panel : T.panelAlt, fontWeight: label === "Total" ? 700 : 400 }}>
                  <td style={{ padding: "7px 12px", color: T.text }}>{label}</td>
                  <td style={{ padding: "7px 12px", textAlign: "right", color: T.text }}>{prev.toFixed(2)} p.p.</td>
                  <td style={{ padding: "7px 12px", textAlign: "right", color: T.text }}>{cur.toFixed(2)} p.p.</td>
                  <td style={{ padding: "7px 12px", textAlign: "right", color: v >= 0 ? T.green : T.red, fontWeight: 600 }}>{v >= 0 ? "+" : ""}{v.toFixed(2)} p.p.</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ height: 160 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={breakdownChartData}>
              <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
              <XAxis dataKey="tech" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
              <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} unit=" p.p." />
              <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v) => [`${v} p.p.`, "Variación"]} />
              <Bar dataKey="variacion" radius={[4, 4, 0, 0]}>{breakdownChartData.map((d, i) => <Cell key={i} fill={d.variacion >= 0 ? T.green : T.red} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Panel>

      {editing && (
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={save} disabled={saving} style={primaryBtn}>{saving ? <Loader2 size={13} className="spin" /> : <CheckCircle2 size={13} />} Guardar</button>
          <button onClick={() => setEditing(false)} style={ghostBtn}>Cancelar</button>
        </div>
      )}

      <NewsSection news={data.news || []} onSave={async (news) => { const next = { ...data, news }; await saveVmsData(next); setData(next); }} />

      <RoundNotes rounds={rounds} />
    </div>
  );
}

/* =================================================================
   COMPETENCIA
================================================================= */
function computeTechMixForRegion(roundData, teamIdx, region) {
  if (!roundData || roundData.kind !== "cesim") return null;
  const regionsToUse = !region || region === "Global" ? REGIONS : [region];
  const totals = [0, 0, 0, 0];
  regionsToUse.forEach((r) => {
    const block = roundData.blocks.find((b) => b.title === `Informe de mercado, ${r}`);
    if (!block) return;
    TECH_LABELS.forEach((tech, i) => {
      const v = getMetricUnderGroup(block, tech, "Ventas, miles unidades", teamIdx);
      if (v !== null) totals[i] += v;
    });
  });
  const sum = totals.reduce((a, b) => a + b, 0);
  if (sum <= 0) return null;
  return TECH_LABELS.map((t, i) => ({ tech: t, units: totals[i], pct: (totals[i] / sum) * 100 })).filter((t) => t.units > 0);
}

function PositioningMap({ roundData, ourTeam, region = "Global" }) {
  const isGlobal = !region || region === "Global";
  const currency = isGlobal ? null : (REGION_CURRENCY[region] || "USD");
  const priceBlock = isGlobal ? null : roundData.blocks.find((b) => b.title === `Informe de mercado, ${region}`);
  const data = roundData.teams.map((tm, i) => {
    const price = isGlobal ? computeGlobalPriceIndex(roundData, i) : getTecMetricPriceAnyCurrency(priceBlock, i);
    const features = isGlobal
      ? REGIONS.map((r) => getTecMetric(roundData.blocks.find((b) => b.title === `Informe de mercado, ${r}`), "Cantidad de características ofrecidas", i)).filter((v) => v !== null).reduce((acc, v, _, arr) => acc + v / arr.length, 0) || null
      : getTecMetric(priceBlock, "Cantidad de características ofrecidas", i);
    return { team: tm, price, features };
  }).filter((d) => d.price !== null && d.features !== null);
  if (data.length === 0) return null;
  const xLabel = isGlobal ? "Precio, % del promedio regional (Global)" : `Precio, ${region} (${currency})`;
  const xUnit = isGlobal ? "%" : ` ${currency}`;
  return (
    <Panel style={{ padding: 18, marginBottom: 14 }}>
      <Eyebrow info={isGlobal
        ? "Un punto por equipo: precio propio como % del precio promedio de mercado, promediado entre EE.UU./Asia/Europa (así se puede comparar sin mezclar monedas) vs. características promedio."
        : `Un punto por equipo: precio en ${region} (${currency}) vs. características en ese mismo mercado.`}>Mapa de posicionamiento</Eyebrow>
      <div style={{ height: 280 }}>
        <ResponsiveContainer width="100%" height="100%">
          <ScatterChart margin={{ top: 20, right: 24, bottom: 20, left: 10 }}>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
            <XAxis type="number" dataKey="price" name="Precio" unit={xUnit} stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} label={{ value: xLabel, position: "insideBottom", offset: -6, fill: T.textFaint, fontSize: 11 }} />
            <YAxis type="number" dataKey="features" name="Características" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} label={{ value: "Características", angle: -90, position: "insideLeft", fill: T.textFaint, fontSize: 11 }} />
            <Tooltip cursor={{ strokeDasharray: "3 3" }} contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={tooltipNumberFormatter} labelFormatter={() => ""} />
            <Scatter data={data}>
              <LabelList dataKey="team" position="top" style={{ fontSize: 11, fill: T.textDim, fontFamily: "'IBM Plex Mono', monospace" }} />
              {data.map((d, i) => <Cell key={i} fill={d.team === ourTeam ? T.amber : T.cyan} />)}
            </Scatter>
          </ScatterChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );
}
function TechMixChart({ roundData, ourTeam, region = "Global" }) {
  const rows = roundData.teams.map((tm, i) => { const mix = computeTechMixForRegion(roundData, i, region); const point = { team: tm }; (mix || []).forEach((m) => { point[m.tech] = Number(m.pct.toFixed(1)); }); return point; });
  const techsPresent = TECH_LABELS.filter((t) => rows.some((r) => r[t] !== undefined));
  if (techsPresent.length === 0) return null;
  return (
    <Panel style={{ padding: 18, marginBottom: 14 }}>
      <Eyebrow info={`Qué % de las unidades vendidas de cada equipo${region !== "Global" ? ` en ${region}` : ""} corresponden a cada tecnología — para anticipar quién migra primero.`}>Mix de tecnología vendida{region !== "Global" ? ` — ${region}` : ""}</Eyebrow>
      <div style={{ height: 240 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={rows} layout="vertical" margin={{ left: 10 }}>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" horizontal={false} />
            <XAxis type="number" unit="%" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <YAxis type="category" dataKey="team" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 12 }} width={70} />
            <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v, n) => [`${Number(v).toFixed(1)}%`, n]} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            {techsPresent.map((t, i) => <Bar key={t} dataKey={t} stackId="mix" fill={SERIES_COLORS[i % SERIES_COLORS.length]} />)}
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );

}
function MarketOverviewTable({ roundData, ourTeam }) {
  const rows = roundData.teams.map((tm, i) => {
    const row = { team: tm };
    KPI_DEFS.forEach((def) => { row[def.key] = extractKpi(roundData, i, def); });
    row.leverage = getRatio(roundData, "Endeudamiento neto/patrimonio (apalancamiento)", i);
    row.roe = getRatio(roundData, "Rendimiento de los Fondos Propios (ROE)", i);
    return row;
  }).sort((a, b) => (b.revenue ?? -Infinity) - (a.revenue ?? -Infinity));
  const cols = [...KPI_DEFS.map((d) => ({ key: d.key, label: d.label })), { key: "leverage", label: "Apalancamiento, %" }, { key: "roe", label: "ROE, %" }];
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Los 7 equipos y sus indicadores clave en una sola tabla, ordenados por ingresos.">Todo el mercado — R{roundData.roundNumber ?? ""}</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <thead><tr>
            <th style={{ textAlign: "left", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>Equipo</th>
            {cols.map((c) => <th key={c.key} style={{ textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" }}>{c.label}</th>)}
          </tr></thead>
          <tbody>
            {rows.map((r, ri) => (
              <tr key={r.team} style={{ background: r.team === ourTeam ? T.amberDim : (ri % 2 ? T.panel : T.panelAlt) }}>
                <td style={{ padding: "6px 12px", color: r.team === ourTeam ? T.amber : T.text, fontWeight: r.team === ourTeam ? 700 : 400 }}>{r.team}</td>
                {cols.map((c) => <td key={c.key} style={{ padding: "6px 12px", textAlign: "right", color: r.team === ourTeam ? T.amber : T.text }}>{typeof r[c.key] === "number" ? fmtNumber(r[c.key]) : fmtCell(r[c.key])}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}
function getMetricUnderGroupPriceAnyCurrency(block, techLabel, teamIdx) {
  if (!block) return null;
  const rows = block.rows;
  const groupIdx = rows.findIndex((r) => r.kind === "group" && r.label === techLabel);
  if (groupIdx === -1) return null;
  for (let i = groupIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.kind === "group") break;
    if (r.kind === "metric" && r.label.startsWith("Precio de venta,")) return toNumberOrNull(r.values[teamIdx]);
  }
  return null;
}

function CompetitorFocusTable({ roundData, ourTeam }) {
  const ourIdx = roundData.teams.indexOf(ourTeam);
  if (ourIdx < 0) return null;
  const rows = [];
  REGIONS.forEach((region) => {
    const block = roundData.blocks.find((b) => b.title === `Informe de mercado, ${region}`);
    if (!block) return;
    TECH_LABELS.forEach((tech) => {
      const ourPrice = getMetricUnderGroupPriceAnyCurrency(block, tech, ourIdx);
      const ourFocus = getRawMetricUnderGroup(block, tech, "Enfoque de la estrategia de marketing", ourIdx);
      const brandTeams = roundData.teams
        .map((tm, i) => {
          if (i === ourIdx) return null;
          const focus = getRawMetricUnderGroup(block, tech, "Enfoque de la estrategia de marketing", i);
          if (focus !== "Marca") return null;
          return { team: tm, price: getMetricUnderGroupPriceAnyCurrency(block, tech, i) };
        })
        .filter(Boolean);
      if (brandTeams.length > 0) rows.push({ region, tech, currency: REGION_CURRENCY[region], ourPrice, ourFocus, brandTeams });
    });
  });

  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info='Para cada región y tecnología, qué competidores eligieron enfoque "Marca" y cómo se compara su precio contra el de ULTI en esa misma región/tecnología.'>Enfoque de marca — comparación de precios</Eyebrow>
      {rows.length === 0 ? (
        <div style={{ color: T.textFaint, fontSize: 13 }}>Ningún competidor tiene enfoque "Marca" en esta ronda.</div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {rows.map((r, i) => (
            <div key={i} style={{ border: `1px solid ${T.border}`, borderRadius: 10, padding: "12px 14px", background: T.panelAlt }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: T.text, marginBottom: 8 }}>{r.region} — {r.tech}</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 12px", borderRadius: 8, background: T.amberDim, border: `1px solid ${T.amber}` }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: T.amber }}>{ourTeam}</span>
                  <span style={{ fontSize: 11.5, color: T.textDim }}>{ourFocusLabel(r.ourFocus)} · {r.ourPrice !== null ? `${r.ourPrice.toFixed(0)} ${r.currency}` : "—"}</span>
                </div>
                {r.brandTeams.map((bt) => (
                  <div key={bt.team} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 12px", borderRadius: 8, background: T.violet + "15", border: `1px solid ${T.violet}55` }}>
                    <span style={{ fontSize: 12, fontWeight: 700, color: T.violet }}>{bt.team}</span>
                    <span style={{ fontSize: 11.5, color: T.textDim }}>Marca · {bt.price !== null ? `${bt.price.toFixed(0)} ${r.currency}` : "—"}</span>
                    {bt.price !== null && r.ourPrice !== null && (
                      <span style={{ fontSize: 11, fontWeight: 600, color: bt.price > r.ourPrice ? T.green : T.red }}>
                        (ULTI {bt.price > r.ourPrice ? "más barato" : "más caro"} por {Math.abs(((r.ourPrice - bt.price) / bt.price) * 100).toFixed(0)}%)
                      </span>
                    )}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </Panel>
  );
}
function ourFocusLabel(focus) {
  return focus || "sin dato";
}

function CompetitionSection({ rounds, dataById, ourTeam, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const [selectedId, setSelectedId] = useState(cesimRounds.length ? cesimRounds[cesimRounds.length - 1].id : null);
  useEffect(() => { if (cesimRounds.length && !cesimRounds.find((r) => r.id === selectedId)) setSelectedId(cesimRounds[cesimRounds.length - 1].id); }, [cesimRounds.map((r) => r.id).join(",")]);
  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Sin datos todavía.</div>;
  if (!ourTeam) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Elegí cuál es su equipo (arriba) para ver a los rivales más cercanos.</div>;

  const roundData = dataById[selectedId];
  const ourIdx = roundData.teams.indexOf(ourTeam);
  const distances = computeDistances(roundData, ourIdx).sort((a, b) => (a.distance ?? Infinity) - (b.distance ?? Infinity));
  const closest = distances.slice(0, 3);
  const ourProfile = computeStrategicProfile(roundData, ourIdx);
  const chartData = distances.filter((d) => d.distance !== null).map((d) => ({ team: d.team, distance: Number(d.distance.toFixed(2)) }));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        <span style={{ fontSize: 12, color: T.textDim }}>Ronda:</span>
        {cesimRounds.map((r) => (
          <button key={r.id} onClick={() => setSelectedId(r.id)} style={{ padding: "5px 10px", borderRadius: 6, fontSize: 12, cursor: "pointer", border: `1px solid ${r.id === selectedId ? T.amber : T.border}`, background: r.id === selectedId ? T.amberDim : "transparent", color: r.id === selectedId ? T.amber : T.textDim }}>R{r.roundNumber}</button>
        ))}
      </div>

      <PositioningMap roundData={roundData} ourTeam={ourTeam} region={regionFilter} />

      <Panel style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 16px 6px" }}><Eyebrow><Target size={11} style={{ verticalAlign: -2, marginRight: 4 }} />Rivales más directos vs. {ourTeam}</Eyebrow></div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
            <thead><tr>
              <th style={{ textAlign: "left", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}` }} />
              {STRATEGY_VARS.map((v) => <th key={v.key} style={{ textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" }}>{v.label}</th>)}
            </tr></thead>
            <tbody>
              <tr style={{ background: T.amberDim }}>
                <td style={{ padding: "6px 12px", color: T.amber, fontWeight: 700 }}>{ourTeam} (nosotros)</td>
                {STRATEGY_VARS.map((v) => <td key={v.key} style={{ padding: "6px 12px", textAlign: "right", color: T.amber, fontWeight: 600 }}>{fmtCell(ourProfile?.[v.key])}</td>)}
              </tr>
              {closest.map((d, i) => (
                <tr key={d.team} style={{ background: i % 2 ? T.panel : T.panelAlt }}>
                  <td style={{ padding: "6px 12px", color: T.text }}>{d.team} <span style={{ color: T.textFaint, fontSize: 11 }}>· dist. {d.distance !== null ? d.distance.toFixed(2) : "—"}</span></td>
                  {STRATEGY_VARS.map((v) => <td key={v.key} style={{ padding: "6px 12px", textAlign: "right", color: T.text }}>{fmtCell(d.profile?.[v.key])}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Panel>

      {regionFilter !== "Global" && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: T.textDim, background: T.cyanDim, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px" }}>
          <Info size={13} color={T.cyan} style={{ flexShrink: 0 }} />
          La distancia estratégica y "todo el mercado" siguen a nivel compañía (I+D, promoción y cuota global no tienen desglose por región en Cesim).
        </div>
      )}

      <Panel style={{ padding: 18 }}>
        <Eyebrow info="Distancia calculada con precio, características, I+D, promoción y cuota de mercado, normalizados entre los 7 equipos. Menor = estrategia más parecida.">Distancia estratégica a {ourTeam}</Eyebrow>
        <div style={{ height: 200 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={chartData} layout="vertical" margin={{ left: 10 }}>
              <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" horizontal={false} />
              <XAxis type="number" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
              <YAxis type="category" dataKey="team" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 12 }} width={70} />
              <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} />
              <Bar dataKey="distance" radius={[0, 4, 4, 0]}>{chartData.map((d, i) => <Cell key={i} fill={i < 2 ? T.amber : T.cyan} opacity={i < 2 ? 1 : 0.55} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Panel>

      <TechMixChart roundData={roundData} ourTeam={ourTeam} region={regionFilter} />

      <CompetitorFocusTable roundData={roundData} ourTeam={ourTeam} />

      <MarketOverviewTable roundData={{ ...roundData, roundNumber: cesimRounds.find((r) => r.id === selectedId)?.roundNumber }} ourTeam={ourTeam} />
    </div>
  );
}

/* =================================================================
   PANEL DINÁMICO
================================================================= */
function DynamicDashboard({ rounds, dataById, ourTeam, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const catalog = useMemo(() => buildMetricCatalog(rounds, dataById), [rounds, dataById]);
  const [query, setQuery] = useState(""); const [selected, setSelected] = useState(null); const [selectedTeams, setSelectedTeams] = useState([]);
  const teams = cesimRounds.length ? dataById[cesimRounds[0].id].teams : [];

  useEffect(() => {
    if (!ourTeam || !cesimRounds.length || selectedTeams.length) return;
    const latest = dataById[cesimRounds[cesimRounds.length - 1].id];
    const ourIdx = latest.teams.indexOf(ourTeam);
    const dists = computeDistances(latest, ourIdx).sort((a, b) => (a.distance ?? Infinity) - (b.distance ?? Infinity));
    setSelectedTeams([ourTeam, ...dists.slice(0, 2).map((d) => d.team)]);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ourTeam, cesimRounds.length]);

  const regionScoped = useMemo(() => {
    if (regionFilter === "Global") return catalog;
    return catalog.filter((c) => {
      const t = c.blockTitle.toLowerCase();
      const mentionsOtherRegion = REGIONS.filter((r) => r !== regionFilter).some((r) => t.includes(r.toLowerCase()));
      return !mentionsOtherRegion;
    });
  }, [catalog, regionFilter]);

  const filtered = useMemo(() => {
    if (!query.trim()) return regionScoped.slice(0, 30);
    const q = query.toLowerCase();
    return regionScoped.filter((c) => (c.label + " " + c.blockTitle + " " + (c.group || "")).toLowerCase().includes(q)).slice(0, 30);
  }, [regionScoped, query]);

  const chartData = useMemo(() => {
    if (!selected) return [];
    return cesimRounds.map((r) => {
      const rd = dataById[r.id]; const block = rd.blocks[selected.blockIndex]; const row = block?.rows?.[selected.rowIndex];
      const point = { roundNumber: r.roundNumber };
      if (row && row.kind === "metric" && row.label === selected.label) teams.forEach((tm, i) => { if (selectedTeams.includes(tm)) point[tm] = toNumberOrNull(row.values[i]); });
      return point;
    });
  }, [selected, cesimRounds, dataById, teams, selectedTeams]);

  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Sin datos todavía.</div>;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <Panel style={{ padding: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, padding: "8px 12px" }}>
          <Search size={14} color={T.textFaint} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar un indicador… (ej: precio, I+D, inventario, ratios)" style={{ flex: 1, background: "none", border: "none", outline: "none", color: T.text, fontSize: 13, fontFamily: "'Inter', sans-serif" }} />
        </div>
        <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 2, maxHeight: 220, overflowY: "auto" }}>
          {filtered.map((c) => (
            <button key={c.key} onClick={() => setSelected(c)} style={{ textAlign: "left", background: selected?.key === c.key ? T.amberDim : "transparent", border: "none", borderRadius: 6, padding: "7px 10px", cursor: "pointer", color: selected?.key === c.key ? T.amber : T.text }}>
              <div style={{ fontSize: 13, fontWeight: 500 }}>{c.label}</div>
              <div style={{ fontSize: 11, color: T.textFaint }}>{c.category} · {c.blockTitle}{c.group ? ` · ${c.group}` : ""}</div>
            </button>
          ))}
          {filtered.length === 0 && <div style={{ fontSize: 12.5, color: T.textFaint, padding: 8 }}>Sin resultados para "{query}".</div>}
        </div>
      </Panel>
      {selected && (
        <>
          <Panel style={{ padding: 14 }}>
            <Eyebrow>Equipos a comparar</Eyebrow>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {teams.map((tm) => {
                const on = selectedTeams.includes(tm);
                return <button key={tm} onClick={() => setSelectedTeams((cur) => on ? cur.filter((t) => t !== tm) : [...cur, tm])} style={{ padding: "5px 11px", borderRadius: 6, fontSize: 12, cursor: "pointer", fontWeight: 600, border: `1px solid ${on ? (tm === ourTeam ? T.amber : T.cyan) : T.border}`, background: on ? (tm === ourTeam ? T.amberDim : T.cyanDim) : "transparent", color: on ? (tm === ourTeam ? T.amber : T.cyan) : T.textFaint }}>{tm}</button>;
              })}
            </div>
          </Panel>
          <Panel style={{ padding: 18 }}>
            <Eyebrow>{selected.label} {selected.group ? `— ${selected.group}` : ""}</Eyebrow>
            <div style={{ height: 260 }}>
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={chartData}>
                  <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
                  <XAxis dataKey="roundNumber" allowDecimals={false} stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} label={{ value: "Ronda", position: "insideBottom", offset: -4, fill: T.textFaint, fontSize: 11 }} />
                  <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
                  <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} labelFormatter={(v) => `Ronda ${v}`} formatter={tooltipNumberFormatter} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  {selectedTeams.map((tm, i) => <Line key={tm} dataKey={tm} stroke={tm === ourTeam ? T.amber : SERIES_COLORS[(i + 1) % SERIES_COLORS.length]} strokeWidth={tm === ourTeam ? 3 : 2} dot={{ r: 3 }} connectNulls />)}
                </LineChart>
              </ResponsiveContainer>
            </div>
          </Panel>
        </>
      )}
    </div>
  );
}

/* =================================================================
   ESTRATEGIA
================================================================= */
function NumField({ label, value, onChange, width }) {
  return (
    <div>
      <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>{label}</div>
      <input type="number" value={value ?? ""} onChange={(e) => onChange(e.target.value === "" ? null : Number(e.target.value))} style={{ ...inputStyle, width: width || 130 }} />
    </div>
  );
}
function TechPlanEditor({ tech, currency, value, onChange }) {
  return (
    <div style={{ padding: 10, background: T.panel, borderRadius: 8, border: `1px solid ${T.borderSoft}` }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 8 }}>{tech}</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
        <NumField label={`Precio mín, ${currency}`} value={value.priceMin} onChange={(v) => onChange({ priceMin: v })} width={110} />
        <NumField label={`Precio máx, ${currency}`} value={value.priceMax} onChange={(v) => onChange({ priceMax: v })} width={110} />
      </div>
      <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
        {MARKETING_FOCUS_OPTIONS.map((opt) => (
          <button key={opt} onClick={() => onChange({ marketingFocus: value.marketingFocus === opt ? null : opt })} style={{ padding: "4px 9px", borderRadius: 6, fontSize: 11.5, cursor: "pointer", fontWeight: 600, border: `1px solid ${value.marketingFocus === opt ? T.amber : T.border}`, background: value.marketingFocus === opt ? T.amberDim : "transparent", color: value.marketingFocus === opt ? T.amber : T.text }}>{opt}</button>
        ))}
      </div>
    </div>
  );
}
function RegionPlanEditor({ region, value, onChange }) {
  const currency = REGION_CURRENCY[region] || "USD";
  const techs = value.techs || {};
  const setTech = (tech, patch) => onChange({ techs: { ...techs, [tech]: { ...(techs[tech] || {}), ...patch } } });
  return (
    <Panel style={{ padding: 14, background: T.panelAlt }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: T.text, marginBottom: 10 }}>{region} <span style={{ color: T.textFaint, fontWeight: 400 }}>· precios en {currency}</span></div>
      <div style={{ marginBottom: 12 }}>
        <NumField label="Cuota mín, % (región, todas las Tec)" value={value.shareMin} onChange={(v) => onChange({ shareMin: v })} width={220} />
      </div>
      <div style={{ fontSize: 11, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>Por tecnología</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 8 }}>
        {TECH_LABELS.map((tech) => (
          <TechPlanEditor key={tech} tech={tech} currency={currency} value={techs[tech] || {}} onChange={(patch) => setTech(tech, patch)} />
        ))}
      </div>
    </Panel>
  );
}
function PlanSummaryView({ plan }) {
  if (!plan) return <div style={{ color: T.textFaint, fontSize: 12.5 }}>No se guardó una estimación para esta ronda.</div>;
  const rows = [];
  REGIONS.forEach((region) => {
    const rp = plan.regions?.[region];
    if (!rp) return;
    TECH_LABELS.forEach((tech) => {
      const tp = rp.techs?.[tech];
      if (!tp || (tp.priceMin == null && tp.priceMax == null && !tp.marketingFocus)) return;
      rows.push({ region, tech, currency: REGION_CURRENCY[region], ...tp });
    });
  });
  return (
    <div>
      {plan.positioning && <div style={{ fontSize: 14, color: T.text, marginBottom: 8, fontWeight: 600 }}>{plan.positioning}</div>}
      {(plan.rdMin != null || plan.rdMax != null || plan.leverageMax != null) && (
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 12.5, color: T.textDim, marginBottom: 10 }}>
          {(plan.rdMin != null || plan.rdMax != null) && <span>I+D estimado: {plan.rdMin ?? "–"} a {plan.rdMax ?? "–"} miles USD</span>}
          {plan.leverageMax != null && <span>Apalancamiento máx: {plan.leverageMax}%</span>}
        </div>
      )}
      {rows.length > 0 ? (
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
            <thead><tr>
              <th style={{ textAlign: "left", padding: "6px 10px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>Región</th>
              <th style={{ textAlign: "left", padding: "6px 10px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>Tec</th>
              <th style={{ textAlign: "right", padding: "6px 10px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>Precio estimado</th>
              <th style={{ textAlign: "left", padding: "6px 10px", color: T.textDim, borderBottom: `1px solid ${T.border}` }}>Enfoque</th>
            </tr></thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={i} style={{ background: i % 2 ? T.panel : T.panelAlt }}>
                  <td style={{ padding: "6px 10px", color: T.text }}>{r.region}</td>
                  <td style={{ padding: "6px 10px", color: T.text }}>{r.tech}</td>
                  <td style={{ padding: "6px 10px", textAlign: "right", color: T.text }}>{r.priceMin != null || r.priceMax != null ? `${r.priceMin ?? "–"} a ${r.priceMax ?? "–"} ${r.currency}` : "—"}</td>
                  <td style={{ padding: "6px 10px", color: T.text }}>{r.marketingFocus || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : <div style={{ color: T.textFaint, fontSize: 12.5 }}>Sin precios estimados por región/tecnología.</div>}
    </div>
  );
}
function RoundComment({ roundId }) {
  const [text, setText] = useState(""); const [saved, setSaved] = useState(true); const [loading, setLoading] = useState(true);
  useEffect(() => { setLoading(true); loadStrategyComment(roundId).then((t) => { setText(t); setSaved(true); setLoading(false); }); }, [roundId]);
  const doSave = async () => { await saveStrategyComment(roundId, text); setSaved(true); };
  return (
    <Panel style={{ padding: 16 }}>
      <Eyebrow info='Por qué tomaron las decisiones de esta ronda — ej: "pusimos tanto en publicidad porque..."'>Comentario de la ronda</Eyebrow>
      <textarea value={text} onChange={(e) => { setText(e.target.value); setSaved(false); }} disabled={loading}
        placeholder="¿Por qué decidieron esto en esta ronda?"
        style={{ width: "100%", minHeight: 70, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, padding: 10, fontSize: 13, fontFamily: "'Inter', sans-serif", resize: "vertical", outline: "none" }} />
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
        <button onClick={doSave} disabled={saved} style={{ ...primaryBtn, opacity: saved ? 0.5 : 1, cursor: saved ? "default" : "pointer" }}><CheckCircle2 size={13} /> {saved ? "Guardado" : "Guardar comentario"}</button>
      </div>
    </Panel>
  );
}
function StrategySection({ rounds, dataById, ourTeam }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const [selectedId, setSelectedId] = useState(null);
  const [plan, setPlan] = useState(null); const [form, setForm] = useState({});
  const [snapshot, setSnapshot] = useState(null);
  const [loading, setLoading] = useState(true); const [editing, setEditing] = useState(false); const [saving, setSaving] = useState(false);

  const latestRound = cesimRounds[cesimRounds.length - 1];
  useEffect(() => { if (latestRound && !selectedId) setSelectedId(latestRound.id); }, [latestRound?.id]);
  const isLatest = selectedId === latestRound?.id;

  useEffect(() => { loadStrategyPlan().then((p) => { setPlan(p); setForm(p || {}); setLoading(false); setEditing(!p); }); }, []);
  useEffect(() => {
    if (!selectedId) return;
    if (selectedId === latestRound?.id) { setSnapshot(null); return; } // latest round shows the live editable plan instead
    loadStrategySnapshot(selectedId).then(setSnapshot);
  }, [selectedId, latestRound?.id]);

  const save = async () => {
    setSaving(true);
    await saveStrategyPlan(form);
    if (latestRound) await saveStrategySnapshot(latestRound.id, form); // keep a dated copy for the historial
    setPlan(form); setEditing(false); setSaving(false);
  };

  if (loading) return <div style={{ padding: 20, color: T.textFaint, display: "flex", alignItems: "center", gap: 8 }}><Loader2 size={14} className="spin" /> Cargando…</div>;
  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Cargá una ronda primero.</div>;
  if (!ourTeam) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Elegí cuál es su equipo (arriba).</div>;

  const selectedRoundData = selectedId ? dataById[selectedId] : null;
  const ourIdx = selectedRoundData ? selectedRoundData.teams.indexOf(ourTeam) : -1;
  const planForChecks = isLatest ? plan : snapshot;
  const checks = planForChecks && selectedRoundData ? computeStrategyAlignment(planForChecks, selectedRoundData, ourIdx) : [];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
        <span style={{ fontSize: 12, color: T.textDim }}>Ronda:</span>
        {cesimRounds.map((r) => (
          <button key={r.id} onClick={() => setSelectedId(r.id)} style={{ padding: "5px 10px", borderRadius: 6, fontSize: 12, cursor: "pointer", border: `1px solid ${r.id === selectedId ? T.amber : T.border}`, background: r.id === selectedId ? T.amberDim : "transparent", color: r.id === selectedId ? T.amber : T.textDim }}>R{r.roundNumber}</button>
        ))}
      </div>

      {!isLatest && (
        <Panel style={{ padding: 18 }}>
          <Eyebrow info="Lo que habían estimado para esta ronda ya jugada, comparado contra lo que terminó pasando.">Lo que estimamos para esta ronda</Eyebrow>
          <PlanSummaryView plan={snapshot} />
          {checks.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 14 }}>
              {checks.map((c) => (
                <div key={c.key} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6, fontSize: 13, padding: "9px 12px", background: T.panelAlt, borderRadius: 7, border: `1px solid ${c.ok ? T.borderSoft : T.red}` }}>
                  <span style={{ color: T.text, fontWeight: 500 }}>{c.label}</span>
                  <span style={{ color: T.textFaint, fontSize: 12 }}>estimado: {c.target}</span>
                  <span style={{ color: c.ok ? T.green : T.red, fontWeight: 700 }}>{c.ok ? "✓" : "✗"} {c.actual}</span>
                </div>
              ))}
            </div>
          )}
        </Panel>
      )}

      {isLatest && !editing && plan && (
        <Panel style={{ padding: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
            <Eyebrow info="Objetivos definidos por el equipo para la ronda actual. Se chequean automáticamente contra la última ronda cargada.">Nuestro plan estratégico — ronda actual</Eyebrow>
            <button onClick={() => { setForm(plan); setEditing(true); }} style={ghostBtn}><Pencil size={12} /> Editar</button>
          </div>
          {plan.notes && <div style={{ fontSize: 13, color: T.textDim, marginBottom: 14, whiteSpace: "pre-wrap" }}>{plan.notes}</div>}
          <PlanSummaryView plan={plan} />
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 14 }}>
            {checks.map((c) => (
              <div key={c.key} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 6, fontSize: 13, padding: "9px 12px", background: T.panelAlt, borderRadius: 7, border: `1px solid ${c.ok ? T.borderSoft : T.red}` }}>
                <span style={{ color: T.text, fontWeight: 500 }}>{c.label}</span>
                <span style={{ color: T.textFaint, fontSize: 12 }}>objetivo: {c.target}</span>
                <span style={{ color: c.ok ? T.green : T.red, fontWeight: 700 }}>{c.ok ? "✓" : "✗"} {c.actual}</span>
              </div>
            ))}
            {checks.length === 0 && <div style={{ color: T.textFaint, fontSize: 12.5 }}>Definí al menos un objetivo numérico para poder chequear alineación.</div>}
          </div>
        </Panel>
      )}
      {isLatest && editing && (
        <Panel style={{ padding: 18 }}>
          <Eyebrow>{plan ? "Editar plan estratégico" : "Definir plan estratégico"}</Eyebrow>
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            <div>
              <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Posicionamiento (libre)</div>
              <input value={form.positioning || ""} onChange={(e) => setForm({ ...form, positioning: e.target.value })} style={{ ...inputStyle, width: "100%", fontFamily: "'Inter', sans-serif" }} placeholder="Ej: Diferenciación por tecnología" />
            </div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <NumField label="I+D mín, miles USD (compañía)" value={form.rdMin} onChange={(v) => setForm({ ...form, rdMin: v })} width={170} />
              <NumField label="I+D máx, miles USD (compañía)" value={form.rdMax} onChange={(v) => setForm({ ...form, rdMax: v })} width={170} />
              <NumField label="Apalancamiento máx, % (compañía)" value={form.leverageMax} onChange={(v) => setForm({ ...form, leverageMax: v })} width={190} />
            </div>
            <div style={{ fontSize: 11, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.08em", marginTop: 4 }}>Objetivos por región (dejá en blanco lo que no quieran fijar) — cada región abre sus 4 tecnologías</div>
            {REGIONS.map((region) => (
              <RegionPlanEditor key={region} region={region} value={form.regions?.[region] || {}} onChange={(patch) => setForm({ ...form, regions: { ...form.regions, [region]: { ...(form.regions?.[region] || {}), ...patch } } })} />
            ))}
            <div>
              <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Notas / racional</div>
              <textarea value={form.notes || ""} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="¿Por qué eligieron esta estrategia?"
                style={{ width: "100%", minHeight: 70, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, padding: 10, fontSize: 13, fontFamily: "'Inter', sans-serif", resize: "vertical", outline: "none" }} />
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={save} disabled={saving} style={primaryBtn}>{saving ? <Loader2 size={13} className="spin" /> : <CheckCircle2 size={13} />} Guardar plan</button>
              {plan && <button onClick={() => setEditing(false)} style={ghostBtn}>Cancelar</button>}
            </div>
          </div>
        </Panel>
      )}

      {selectedId && <RoundComment roundId={selectedId} />}
    </div>
  );
}

/* =================================================================
   DASHBOARD COMPLETO — one consolidated decision-making view
================================================================= */
function buildExecutiveSummary(latestData, prevData, teamIdx, ourTeam, latestRound) {
  if (!latestData || teamIdx < 0) return [];
  const lines = [];
  const revenue = extractKpi(latestData, teamIdx, KPI_DEFS[0]);
  const profit = extractKpi(latestData, teamIdx, KPI_DEFS[1]);
  const margin = computeMargin(latestData, teamIdx);
  const share = extractKpi(latestData, teamIdx, KPI_DEFS[3]);
  const revenueRank = computeRank(latestData, KPI_DEFS[0], teamIdx);
  const profitRank = computeRank(latestData, KPI_DEFS[1], teamIdx);
  const marginRank = computeMarginRank(latestData, teamIdx);

  if (revenue !== null) {
    let s = `${ourTeam} cerró la Ronda ${latestRound.roundNumber} con ingresos de ${fmtNumber(revenue)} miles USD`;
    if (revenueRank) s += `, ${revenueRank.rank}° de ${revenueRank.of} equipos`;
    if (prevData) {
      const prevRevenue = extractKpi(prevData, teamIdx, KPI_DEFS[0]);
      if (prevRevenue !== null && prevRevenue !== 0) {
        const chg = ((revenue - prevRevenue) / Math.abs(prevRevenue)) * 100;
        s += ` (${chg >= 0 ? "+" : ""}${chg.toFixed(1)}% vs. ronda anterior)`;
      }
    }
    lines.push({ icon: "trend", text: s + "." });
  }
  if (profit !== null) {
    const good = profit >= 0;
    let s = good
      ? `El beneficio de la ronda fue positivo, ${fmtNumber(profit)} miles USD`
      : `Atención: el beneficio de la ronda fue negativo, ${fmtNumber(profit)} miles USD`;
    if (profitRank) s += ` (${profitRank.rank}° de ${profitRank.of})`;
    lines.push({ icon: good ? "ok" : "warn", text: s + "." });
  }
  if (margin !== null) {
    lines.push({ icon: marginRank && marginRank.rank <= 2 ? "ok" : "trend", text: `Margen sobre ventas: ${margin.toFixed(1)}%${marginRank ? ` (${marginRank.rank}° de ${marginRank.of})` : ""}.` });
  }
  if (share !== null) {
    lines.push({ icon: "trend", text: `Cuota de mercado global: ${share.toFixed(1)}%.` });
  }
  return lines;
}
// Renders **bold** markers inside a plain string as <strong> spans.
function renderBoldText(text) {
  const parts = String(text || "").split("**");
  return parts.map((part, i) => (i % 2 === 1 ? <strong key={i}>{part}</strong> : <React.Fragment key={i}>{part}</React.Fragment>));
}

const DEFAULT_DASHBOARD_STORY = {
  highlights: [
    { text: "ULTI cerró la ronda con el **mayor crecimiento porcentual de rentabilidad** entre las 7 competidoras (+129% vs. el año anterior).", info: "El beneficio de ULTI pasó de 237.8M a 546.2M y el crecimiento promedio de beneficio de la industria respecto al año 1 fue de +84.2%." },
    { text: "La apuesta por rentabilidad se reflejó en el retorno al accionista: **ULTI pasó del 5to al 2do puesto** entre las 7 competidoras, con un retorno de 30.51.", info: null },
    { text: "Comparte podio con NOSRE (31.88) y LUCKY (24.66).", info: null },
  ],
  alerts: [
    { text: "**ULTI cedió el liderazgo** en Tecnología 2: **la cuota global cayó de 83% a 30%**. NODO capturó 48% del mercado bajando precio y ampliando producción, y NOSRE entró como nuevo competidor con 22%.", info: null },
    { text: "La **cuota de valor global NO** aumentó: cayó 0,85 p.p. arrastrada por Tec1, pese a la mejora en Tec2.", info: null },
    { text: "67% **de demanda insatisfecha en Europa Tec2: fue el mayor problema de cobertura**.", info: null },
  ],
};

function StoryList({ items, onSave, tone = "text", emptyLabel }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(items);
  const [saving, setSaving] = useState(false);
  useEffect(() => { setDraft(items); }, [items]);

  const color = tone === "danger" ? T.red : T.text;

  if (!editing) {
    return (
      <div>
        {items.length === 0 && <div style={{ color: T.textFaint, fontSize: 12.5, marginBottom: 8 }}>{emptyLabel || "Sin puntos cargados."}</div>}
        <div style={{ display: "flex", flexDirection: "column", gap: 9, marginBottom: 8 }}>
          {items.map((item, i) => (
            <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 7, fontSize: 13.5, color, lineHeight: 1.5 }}>
              {tone === "danger" ? <AlertTriangle size={14} style={{ marginTop: 2, flexShrink: 0 }} /> : <span style={{ marginTop: 6, width: 5, height: 5, borderRadius: "50%", background: T.amber, flexShrink: 0 }} />}
              <span>{renderBoldText(item.text)} {item.info && <InfoBubble text={item.info} />}</span>
            </div>
          ))}
        </div>
        <button onClick={() => setEditing(true)} style={{ ...ghostBtn, padding: "5px 11px", fontSize: 11.5 }}><Pencil size={11} /> Editar puntos</button>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {draft.map((item, i) => (
        <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 4 }}>
            <textarea
              value={item.text}
              onChange={(e) => setDraft((d) => d.map((x, xi) => (xi === i ? { ...x, text: e.target.value } : x)))}
              placeholder="Usá **texto** para resaltar en negrita"
              style={{ width: "100%", minHeight: 50, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, padding: 8, fontSize: 12.5, fontFamily: "'Inter', sans-serif", resize: "vertical", outline: "none" }}
            />
            <input
              value={item.info || ""}
              onChange={(e) => setDraft((d) => d.map((x, xi) => (xi === i ? { ...x, info: e.target.value } : x)))}
              placeholder="Texto del ícono de info (opcional)"
              style={{ ...inputStyle, fontFamily: "'Inter', sans-serif", fontSize: 12 }}
            />
          </div>
          <button onClick={() => setDraft((d) => d.filter((_, xi) => xi !== i))} style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", padding: 4, marginTop: 4 }}><Trash2 size={13} /></button>
        </div>
      ))}
      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={() => setDraft((d) => [...d, { text: "", info: null }])} style={{ ...ghostBtn, padding: "6px 11px", fontSize: 12 }}>+ Agregar punto</button>
        <button onClick={async () => { setSaving(true); await onSave(draft.filter((x) => x.text.trim())); setSaving(false); setEditing(false); }} disabled={saving} style={{ ...primaryBtn, padding: "6px 13px", fontSize: 12 }}>{saving ? <Loader2 size={12} className="spin" /> : <CheckCircle2 size={12} />} Guardar</button>
        <button onClick={() => { setDraft(items); setEditing(false); }} style={{ ...ghostBtn, padding: "6px 11px", fontSize: 12 }}>Cancelar</button>
      </div>
    </div>
  );
}

function ExecutiveSummaryCard({ lines, ourTeam }) {
  const [story, setStory] = useState(null);
  useEffect(() => { loadDashboardStory().then((s) => setStory(s || DEFAULT_DASHBOARD_STORY)); }, []);
  const saveHighlights = async (highlights) => {
    const next = { ...(story || DEFAULT_DASHBOARD_STORY), highlights };
    await saveDashboardStory(next);
    setStory(next);
  };
  if (!lines.length && !story?.highlights?.length) return null;
  const iconFor = (kind) => kind === "ok" ? <CheckCircle2 size={15} color={T.green} /> : kind === "warn" ? <AlertTriangle size={15} color={T.red} /> : <TrendingUpIcon />;
  return (
    <Panel style={{ padding: 20, background: `linear-gradient(135deg, ${T.panel} 0%, ${T.amberDim} 130%)` }}>
      <Eyebrow info="La parte de arriba se genera sola con los datos cargados; los puntos destacados de abajo los edita el equipo — un punto de partida para la presentación a accionistas.">Resumen ejecutivo — {ourTeam}</Eyebrow>
      <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 4, marginBottom: story?.highlights?.length ? 14 : 0 }}>
        {lines.map((l, i) => (
          <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 9, fontSize: 13.5, color: T.text, lineHeight: 1.5 }}>
            <span style={{ marginTop: 1, flexShrink: 0 }}>{iconFor(l.icon)}</span>
            <span>{l.text}</span>
          </div>
        ))}
      </div>
      {story && <StoryList items={story.highlights} onSave={saveHighlights} emptyLabel="Sin puntos destacados cargados." />}
    </Panel>
  );
}
function TrendingUpIcon() {
  return <Compass size={15} color={T.cyan} />;
}

/* ---------------------------------------------------------------
   RETORNO AL ACCIONISTA (TSR) — real row if Cesim reports one,
   otherwise a clearly-labeled proxy from the change in market cap.
--------------------------------------------------------------- */
function findShareholderReturnRow(roundData) {
  if (!roundData || roundData.kind !== "cesim") return null;
  // Search every block — Cesim reports this inside "Ratios e indicadores
  // financieros clave", not under "Creación de valor" / "Valuación".
  for (const block of roundData.blocks) {
    const row = block.rows.find((r) => r.kind === "metric" && /retorno.*accion|TSR|rendimiento.*accion|shareholder/i.test(r.label));
    if (row) return row;
  }
  return null;
}
function computeShareholderReturn(roundData, prevData, teamIdx) {
  const row = findShareholderReturnRow(roundData);
  if (row) {
    const v = toNumberOrNull(row.values[teamIdx]);
    if (v !== null) return { value: v, label: row.label, isProxy: false };
  }
  // Proxy: % change in market capitalization vs. the previous round.
  if (!prevData) return { value: null, label: null, isProxy: true };
  const capDef = KPI_DEFS.find((d) => d.key === "marketcap");
  const capNow = extractKpi(roundData, teamIdx, capDef);
  const capPrev = extractKpi(prevData, teamIdx, capDef);
  if (capNow === null || capPrev === null || capPrev === 0) return { value: null, label: null, isProxy: true };
  return { value: ((capNow - capPrev) / Math.abs(capPrev)) * 100, label: "Variación de capitalización de mercado", isProxy: true };
}
function ShareholderReturnCard({ latestData, prevData, ourTeam, teamIdx }) {
  const own = computeShareholderReturn(latestData, prevData, teamIdx);
  const allTeams = latestData.teams.map((tm, i) => {
    const r = computeShareholderReturn(latestData, prevData, i);
    return { team: tm, value: r.value };
  }).filter((t) => t.value !== null).sort((a, b) => b.value - a.value);
  const ourRankPos = allTeams.findIndex((t) => t.team === ourTeam);

  if (own.value === null) return null;

  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info={own.isProxy
        ? "Cesim no expone una línea de TSR en este reporte, así que aproximamos el retorno al accionista con la variación de la capitalización de mercado ronda a ronda."
        : `Tomado directamente de "${own.label}" en el reporte de Cesim.`}>
        Retorno al accionista {own.isProxy ? "(aproximado)" : "(TSR)"}
      </Eyebrow>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 24, alignItems: "flex-end", marginBottom: 16 }}>
        <div>
          <div style={{ fontSize: 11, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 4 }}>{ourTeam} — esta ronda</div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 34, fontWeight: 700, color: own.value >= 0 ? T.green : T.red }}>
            {own.value >= 0 ? "+" : ""}{own.value.toFixed(1)}%
          </div>
        </div>
        {ourRankPos > -1 && (
          <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: T.textDim, paddingBottom: 6 }}>
            <Trophy size={15} color={ourRankPos < 2 ? T.green : ourRankPos >= allTeams.length - 2 ? T.red : T.textFaint} />
            {ourRankPos + 1}° de {allTeams.length} equipos
          </div>
        )}
      </div>
      <div style={{ height: 190 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={allTeams} layout="vertical" margin={{ left: 10 }}>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" horizontal={false} />
            <XAxis type="number" unit="%" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <YAxis type="category" dataKey="team" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 12 }} width={70} />
            <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v, n) => [`${Number(v).toFixed(1)}%`, n]} />
            <Bar dataKey="value" radius={[0, 4, 4, 0]}>
              {allTeams.map((t, i) => <Cell key={i} fill={t.team === ourTeam ? T.amber : (t.value >= 0 ? T.cyan : T.red)} opacity={t.team === ourTeam ? 1 : 0.7} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );
}

function PromotionTrendCard({ rounds, dataById, ourTeam, teamIdx }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const points = cesimRounds.map((r) => ({ roundNumber: r.roundNumber, value: extractKpi(dataById[r.id], teamIdx, PROMOTION_KPI_DEF) })).filter((p) => p.value !== null);
  if (points.length < 2) return null;
  const first = points[0], last = points[points.length - 1];
  const pctChange = first.value !== 0 ? ((last.value - first.value) / Math.abs(first.value)) * 100 : null;
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Gasto de promoción de la compañía, tomado de la cuenta de resultados, ronda a ronda.">Evolución del gasto de promoción — {ourTeam}</Eyebrow>
      {pctChange !== null && (
        <div style={{ display: "flex", alignItems: "baseline", gap: 10, marginBottom: 10 }}>
          <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 26, fontWeight: 700, color: pctChange >= 0 ? T.green : T.red }}>{pctChange >= 0 ? "+" : ""}{pctChange.toFixed(0)}%</span>
          <span style={{ fontSize: 12.5, color: T.textDim }}>entre R{first.roundNumber} y R{last.roundNumber} ({fmtNumber(first.value)} → {fmtNumber(last.value)} miles USD)</span>
        </div>
      )}
      <div style={{ height: 160 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={points}>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
            <XAxis dataKey="roundNumber" tickFormatter={(v) => `R${v}`} stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} labelFormatter={(v) => `Ronda ${v}`} formatter={tooltipNumberFormatter} />
            <Bar dataKey="value" radius={[4, 4, 0, 0]} fill={T.amber} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );
}
function PromotionRivalsCard({ latestData, prevData, ourTeam }) {
  const rows = computePromotionVariation(latestData, prevData).filter((r) => r.pctChange !== null).sort((a, b) => b.pctChange - a.pctChange);
  if (rows.length === 0) return null;
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Variación % del gasto de promoción entre la ronda anterior y esta, para los equipos que reportan cambios significativos.">Variación del gasto de promoción — vs. rivales</Eyebrow>
      <div style={{ height: Math.max(160, rows.length * 34) }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={rows} layout="vertical" margin={{ left: 10 }}>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" horizontal={false} />
            <XAxis type="number" unit="%" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <YAxis type="category" dataKey="team" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 12 }} width={70} />
            <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v, n) => [`${Number(v).toFixed(0)}%`, n]} />
            <Bar dataKey="pctChange" radius={[0, 4, 4, 0]}>{rows.map((r, i) => <Cell key={i} fill={r.team === ourTeam ? T.amber : T.cyan} />)}</Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );
}
function PromotionByRegionCard({ latestData, prevData, ourTeam }) {
  const rows = REGIONS.map((region) => {
    const cur = extractRegionalPromotion(latestData, region, latestData.teams.indexOf(ourTeam));
    const prev = prevData ? extractRegionalPromotion(prevData, region, prevData.teams.indexOf(ourTeam)) : null;
    return { region, cur, prev };
  });
  const available = rows.filter((r) => r.cur !== null);
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Reparto del presupuesto de promoción entre regiones — solo disponible si Cesim reporta esa línea por región en este curso.">Presupuesto de promoción por región</Eyebrow>
      {available.length === 0 ? (
        <div style={{ color: T.textFaint, fontSize: 12.5 }}>Este reporte de Cesim no incluye el gasto de promoción desglosado por región (solo a nivel compañía). Podés cargar los montos a mano en "Novedades del entorno" si los tienen del juego.</div>
      ) : (
        <div style={{ height: 200 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={rows.filter((r) => r.cur !== null || r.prev !== null)}>
              <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
              <XAxis dataKey="region" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
              <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
              <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={tooltipNumberFormatter} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Bar dataKey="prev" name="Ronda anterior" fill={T.cyan} radius={[4, 4, 0, 0]} />
              <Bar dataKey="cur" name="Esta ronda" fill={T.amber} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
    </Panel>
  );
}

function ClosestRivalsTable({ latestData, ourTeam }) {
  const ourIdx = latestData.teams.indexOf(ourTeam);
  const distances = computeDistances(latestData, ourIdx).sort((a, b) => (a.distance ?? Infinity) - (b.distance ?? Infinity));
  const closest = distances.slice(0, 3);
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 6px" }}><Eyebrow info="Los 3 equipos con la estrategia más parecida a la de ustedes.">Rivales más cercanos</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <tbody>
            {closest.map((d, i) => (
              <tr key={d.team} style={{ background: i % 2 ? T.panel : T.panelAlt }}>
                <td style={{ padding: "8px 16px", color: T.text, fontWeight: 600 }}>{d.team}</td>
                <td style={{ padding: "8px 16px", color: T.textFaint, textAlign: "right" }}>distancia {d.distance !== null ? d.distance.toFixed(2) : "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}
function KpiTrendChart({ rounds, dataById, teamIdx, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const kpiDefs = KPI_DEFS.map((d) => (d.key === "share" ? { ...d, blockTitle: regionMarketBlockTitle(regionFilter) } : d));
  const kpiSeries = {};
  kpiDefs.forEach((def) => {
    const points = [];
    cesimRounds.forEach((r) => { const v = extractKpi(dataById[r.id], teamIdx, def); if (v !== null) points.push({ roundNumber: r.roundNumber, value: v }); });
    if (points.length) kpiSeries[def.key] = { header: def.label, points };
  });
  const kpiKeys = Object.keys(kpiSeries);
  if (kpiKeys.length === 0) return null;
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Evolución de los indicadores clave, ronda a ronda.">Tendencia</Eyebrow>
      <div style={{ height: 240 }}>
        <ResponsiveContainer width="100%" height="100%">
          <LineChart>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
            <XAxis dataKey="roundNumber" type="number" domain={["dataMin", "dataMax"]} allowDecimals={false} stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} labelFormatter={(v) => `Ronda ${v}`} formatter={tooltipNumberFormatter} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {kpiKeys.map((k, i) => <Line key={k} data={kpiSeries[k].points} dataKey="value" name={kpiSeries[k].header} stroke={SERIES_COLORS[i % SERIES_COLORS.length]} strokeWidth={2} dot={{ r: 3 }} connectNulls />)}
          </LineChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );
}

function DashboardCompleto({ rounds, dataById, ourTeam, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const teamIdx = ourTeam && cesimRounds.length ? cesimRounds[0].teams.indexOf(ourTeam) : -1;
  const latestRound = cesimRounds[cesimRounds.length - 1];
  const latestData = latestRound ? dataById[latestRound.id] : null;
  const prevRound = cesimRounds[cesimRounds.length - 2];
  const prevData = prevRound ? dataById[prevRound.id] : null;

  const kpiDefs = useMemo(() => KPI_DEFS.map((d) => (d.key === "share" ? { ...d, blockTitle: regionMarketBlockTitle(regionFilter) } : d)), [regionFilter]);

  const [plan, setPlan] = useState(null);
  useEffect(() => { loadStrategyPlan().then(setPlan); }, []);

  if (cesimRounds.length === 0) return <div style={{ textAlign: "center", padding: "60px 20px", color: T.textFaint }}><Radio size={28} style={{ marginBottom: 10, opacity: 0.5 }} /><div style={{ fontSize: 14 }}>Todavía no hay rondas cargadas.</div></div>;
  if (teamIdx < 0) return <div style={{ textAlign: "center", padding: "40px 20px", color: T.textFaint }}>Elegí cuál es su equipo (arriba) para ver el dashboard.</div>;

  const summary = buildExecutiveSummary(latestData, prevData, teamIdx, ourTeam, latestRound);
  const alerts = computeAlerts(latestData, prevData, teamIdx);
  const alignmentChecks = plan ? computeStrategyAlignment(plan, latestData, teamIdx) : [];
  const alignmentAlerts = alignmentChecks.filter((c) => !c.ok).map((c) => ({ level: "warning", text: `Fuera de estrategia — ${c.label}: objetivo ${c.target}, actual ${c.actual}.`, strategic: true }));
  const allAlerts = [...alerts, ...alignmentAlerts];

  const ourIdx = latestData.teams.indexOf(ourTeam);
  const distances = computeDistances(latestData, ourIdx).sort((a, b) => (a.distance ?? Infinity) - (b.distance ?? Infinity));
  const closest = distances.slice(0, 3);

  const kpiSeries = {};
  kpiDefs.forEach((def) => {
    const points = [];
    cesimRounds.forEach((r) => { const v = extractKpi(dataById[r.id], teamIdx, def); if (v !== null) points.push({ roundNumber: r.roundNumber, value: v }); });
    if (points.length) kpiSeries[def.key] = { header: def.label, points };
  });
  const kpiKeys = Object.keys(kpiSeries);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <ExecutiveSummaryCard lines={summary} ourTeam={ourTeam} />

      <ShareholderReturnCard latestData={latestData} prevData={prevData} ourTeam={ourTeam} teamIdx={teamIdx} />

      {allAlerts.length > 0 && <AlertBanner alerts={allAlerts} />}
      <DashboardManualAlerts />

      <ClosestRivalsTable latestData={latestData} ourTeam={ourTeam} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12 }}>
        {kpiDefs.map((def) => {
          const s = kpiSeries[def.key]; if (!s) return null;
          const last = s.points[s.points.length - 1], prev = s.points[s.points.length - 2];
          const delta = prev ? last.value - prev.value : null;
          const rank = computeRank(latestData, def, teamIdx);
          return <StatCard key={def.key} label={s.header} value={fmtNumber(last.value)} delta={delta} rank={rank} />;
        })}
      </div>

      <KpiTrendChart rounds={rounds} dataById={dataById} teamIdx={teamIdx} regionFilter={regionFilter} />

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14 }}>
        <PositioningMap roundData={latestData} ourTeam={ourTeam} region={regionFilter} />
        <TechMixChart roundData={latestData} ourTeam={ourTeam} region={regionFilter} />
      </div>

      {plan && alignmentChecks.length > 0 && (
        <Panel style={{ padding: 18 }}>
          <Eyebrow info="Chequeo automático del plan estratégico contra la última ronda.">Alineación estratégica</Eyebrow>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {alignmentChecks.map((c) => (
              <div key={c.key} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 12px", borderRadius: 20, fontSize: 12, fontWeight: 600, background: c.ok ? "#EAFBF0" : "#FEEEEE", color: c.ok ? T.green : T.red, border: `1px solid ${c.ok ? "#CDEFDA" : "#FBD5D5"}` }}>
                {c.ok ? <CheckCircle2 size={12} /> : <AlertTriangle size={12} />} {c.label}
              </div>
            ))}
          </div>
        </Panel>
      )}

      <MarketOverviewTable roundData={{ ...latestData, roundNumber: latestRound.roundNumber }} ourTeam={ourTeam} />
    </div>
  );
}

/* =================================================================
   ASISTENTE IA
================================================================= */
function buildAiContext(rounds, dataById, ourTeam, plan, premises) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  // Build the round list newest-first and keep adding older rounds only
  // while they fit a safe character budget. This guarantees the MOST
  // RECENT round is never the one that gets cut — unlike a blind
  // character-slice truncation (which chops off whatever ends up last
  // in the JSON, i.e. exactly the newest round).
  const CONTEXT_BUDGET = 450000;
  const newestFirst = [...cesimRounds].reverse();
  const selected = [];
  let size = 0;
  for (const r of newestFirst) {
    const entry = { roundNumber: r.roundNumber, label: r.label, teams: dataById[r.id].teams, blocks: dataById[r.id].blocks };
    const entrySize = JSON.stringify(entry).length;
    if (selected.length > 0 && size + entrySize > CONTEXT_BUDGET) break; // always keep at least the latest round intact
    selected.push(entry);
    size += entrySize;
  }
  selected.reverse();
  return {
    ourTeam, strategyPlan: plan || null, premises: premises || [],
    rounds: selected,
  };
}
function AiAssistant({ rounds, dataById, ourTeam }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const [premises, setPremises] = useState([]); const [plan, setPlan] = useState(null);
  const [messages, setMessages] = useState([]); const [input, setInput] = useState(""); const [busy, setBusy] = useState(false); const [err, setErr] = useState("");
  useEffect(() => { loadStrategyPlan().then(setPlan); }, []);
  useEffect(() => { loadPremises().then(setPremises); }, []);

  const ask = async () => {
    const q = input.trim(); if (!q || busy) return;
    setMessages((m) => [...m, { role: "user", text: q }]); setInput(""); setBusy(true); setErr("");
    try {
      const context = buildAiContext(rounds, dataById, ourTeam, plan, premises);
      const res = await fetch("/.netlify/functions/ask-ai", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ question: q, context }) });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || `Error ${res.status}`);
      setMessages((m) => [...m, { role: "ai", text: data.answer }]);
    } catch (e) { setErr(e.message || "Ocurrió un error al consultar la IA."); }
    setBusy(false);
  };

  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Cargá al menos una ronda para poder preguntarle a la IA.</div>;
  if (!ourTeam) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Elegí cuál es su equipo (arriba) primero.</div>;

  return (
    <Panel style={{ padding: 16 }}>
      <Eyebrow info="Prioriza siempre la ronda más reciente y suma rondas anteriores mientras entren en el paquete de datos, además del plan de estrategia y las premisas cargadas. No inventa números que no estén en los datos.">Preguntale a la IA sobre estos datos</Eyebrow>
      <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: 380, overflowY: "auto", marginBottom: 12, paddingRight: 4 }}>
        {messages.length === 0 && <div style={{ color: T.textFaint, fontSize: 12.5, fontStyle: "italic" }}>Ej: "¿Cómo venimos de rentabilidad comparado con MECHA?" · "¿Estamos alineados con nuestra estrategia esta ronda?"</div>}
        {messages.map((m, i) => (
          <div key={i} style={{ alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "88%", background: m.role === "user" ? T.amberDim : T.panelAlt, color: m.role === "user" ? T.amber : T.text, border: `1px solid ${T.border}`, borderRadius: 10, padding: "9px 13px", fontSize: 13, whiteSpace: "pre-wrap", lineHeight: 1.5 }}>{m.text}</div>
        ))}
        {busy && <div style={{ color: T.textFaint, fontSize: 12.5, display: "flex", alignItems: "center", gap: 6 }}><Loader2 size={13} className="spin" /> Pensando…</div>}
      </div>
      {err && <div style={{ color: T.red, fontSize: 12.5, marginBottom: 8 }}>{err}</div>}
      <div style={{ display: "flex", gap: 8 }}>
        <input value={input} onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => e.key === "Enter" && ask()} placeholder="Escribí tu pregunta…" style={{ flex: 1, ...inputStyle, fontFamily: "'Inter', sans-serif" }} />
        <button onClick={ask} disabled={busy || !input.trim()} style={primaryBtn}>Preguntar</button>
      </div>
    </Panel>
  );
}

/* =================================================================
   PREMISAS
================================================================= */
const PREMISE_CATEGORIES = ["General", "Financiero", "Mercado", "Producción", "RRHH", "ESG", "Otros"];
function PremiseForm({ initial, onSave, onCancel }) {
  const [form, setForm] = useState(initial || { title: "", category: "General", content: "", fromRound: null });
  const [saving, setSaving] = useState(false);
  const save = async () => { if (!form.title.trim() || !form.content.trim()) return; setSaving(true); await onSave(form); setSaving(false); };
  return (
    <Panel style={{ padding: 16 }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 200 }}>
            <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Título</div>
            <input value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} style={{ ...inputStyle, width: "100%", fontFamily: "'Inter', sans-serif" }} placeholder="Ej: Tasa de interés fija para la ronda 3" />
          </div>
          <div>
            <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Categoría</div>
            <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })} style={{ ...inputStyle, fontFamily: "'Inter', sans-serif" }}>
              {PREMISE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Aplica desde ronda</div>
            <input type="number" value={form.fromRound ?? ""} onChange={(e) => setForm({ ...form, fromRound: e.target.value === "" ? null : Number(e.target.value) })} style={{ ...inputStyle, width: 100 }} placeholder="opcional" />
          </div>
        </div>
        <div>
          <div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Contenido</div>
          <textarea value={form.content} onChange={(e) => setForm({ ...form, content: e.target.value })} placeholder="Pegá o escribí la premisa tal como la dio el profesor o Cesim…"
            style={{ width: "100%", minHeight: 90, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, padding: 10, fontSize: 13, fontFamily: "'Inter', sans-serif", resize: "vertical", outline: "none" }} />
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={save} disabled={saving || !form.title.trim() || !form.content.trim()} style={primaryBtn}>{saving ? <Loader2 size={13} className="spin" /> : <CheckCircle2 size={13} />} Guardar</button>
          <button onClick={onCancel} style={ghostBtn}><X size={13} /> Cancelar</button>
        </div>
      </div>
    </Panel>
  );
}
function PremisesSection() {
  const [premises, setPremises] = useState(null); const [query, setQuery] = useState(""); const [categoryFilter, setCategoryFilter] = useState("Todas");
  const [adding, setAdding] = useState(false); const [editingId, setEditingId] = useState(null); const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  useEffect(() => { loadPremises().then(setPremises); }, []);
  const addPremise = async (form) => { const next = [...(premises || []), { ...form, id: `p${Date.now()}`, createdAt: new Date().toISOString() }]; await savePremises(next); setPremises(next); setAdding(false); };
  const updatePremise = async (id, form) => { const next = (premises || []).map((p) => (p.id === id ? { ...p, ...form } : p)); await savePremises(next); setPremises(next); setEditingId(null); };
  const deletePremise = async (id) => { const next = (premises || []).filter((p) => p.id !== id); await savePremises(next); setPremises(next); setConfirmDeleteId(null); };
  if (premises === null) return <div style={{ padding: 20, color: T.textFaint, display: "flex", alignItems: "center", gap: 8 }}><Loader2 size={14} className="spin" /> Cargando…</div>;
  const filtered = premises.filter((p) => {
    const matchesQuery = !query.trim() || (p.title + " " + p.content).toLowerCase().includes(query.toLowerCase());
    const matchesCat = categoryFilter === "Todas" || p.category === categoryFilter;
    return matchesQuery && matchesCat;
  });
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", flex: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 7, padding: "7px 11px", minWidth: 200 }}>
            <Search size={13} color={T.textFaint} />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar premisa…" style={{ flex: 1, background: "none", border: "none", outline: "none", color: T.text, fontSize: 13, fontFamily: "'Inter', sans-serif" }} />
          </div>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ ...inputStyle, fontFamily: "'Inter', sans-serif" }}>
            <option value="Todas">Todas las categorías</option>
            {PREMISE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        {!adding && <button onClick={() => setAdding(true)} style={primaryBtn}><BookPlus size={14} /> Agregar premisa</button>}
      </div>
      {adding && <PremiseForm onSave={addPremise} onCancel={() => setAdding(false)} />}
      {filtered.length === 0 && !adding && (
        <div style={{ textAlign: "center", padding: "40px 20px", color: T.textFaint }}>
          {premises.length === 0 ? "Todavía no cargaron ninguna premisa. Sumen las reglas, supuestos o restricciones que les dio el profesor o el propio Cesim." : "Ninguna premisa coincide con la búsqueda."}
        </div>
      )}
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filtered.map((p) => editingId === p.id ? (
          <PremiseForm key={p.id} initial={p} onSave={(form) => updatePremise(p.id, form)} onCancel={() => setEditingId(null)} />
        ) : (
          <Panel key={p.id} style={{ padding: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 6 }}>
              <div>
                <div style={{ fontSize: 14, fontWeight: 600, color: T.text }}>{p.title}</div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 3 }}>
                  <span style={{ fontSize: 11, color: T.amber, background: T.amberDim, borderRadius: 4, padding: "1px 7px", fontWeight: 600 }}>{p.category}</span>
                  {p.fromRound != null && <span style={{ fontSize: 11, color: T.textFaint }}>desde ronda {p.fromRound}</span>}
                </div>
              </div>
              <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                <button onClick={() => setEditingId(p.id)} style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", padding: 4 }} title="Editar"><Pencil size={13} /></button>
                <button onClick={() => setConfirmDeleteId(p.id)} style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", padding: 4 }} title="Eliminar"><Trash2 size={13} /></button>
              </div>
            </div>
            <div style={{ fontSize: 13, color: T.textDim, whiteSpace: "pre-wrap", lineHeight: 1.5 }}>{p.content}</div>
          </Panel>
        ))}
      </div>
      {confirmDeleteId && (
        <div style={{ position: "fixed", inset: 0, background: T.overlay, display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }}>
          <Panel style={{ padding: 22, maxWidth: 340 }}>
            <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 8 }}>Eliminar premisa</div>
            <div style={{ fontSize: 13, color: T.textDim, marginBottom: 16 }}>Esta acción no se puede deshacer.</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => deletePremise(confirmDeleteId)} style={{ ...primaryBtn, background: T.red, color: "#fff" }}><Trash2 size={13} /> Eliminar</button>
              <button onClick={() => setConfirmDeleteId(null)} style={ghostBtn}>Cancelar</button>
            </div>
          </Panel>
        </div>
      )}
    </div>
  );
}

/* =================================================================
   ANÁLISIS FINANCIERO — vertical y horizontal
================================================================= */
const FIN_STATEMENTS = {
  pl: { label: "Estado de resultados", blockTitle: "Cuenta de resultados, miles USD, Global", baseLabel: "Ingresos por ventas" },
  bs: { label: "Balance", blockTitle: "Hoja de Balance, miles USD, Global", baseLabelPattern: /total.*activ/i },
};

function findBaseRow(block, stmtDef) {
  if (!block) return null;
  if (stmtDef.baseLabel) return findMetricRow(block, stmtDef.baseLabel);
  if (stmtDef.baseLabelPattern) return block.rows.find((r) => r.kind === "metric" && stmtDef.baseLabelPattern.test(r.label)) || null;
  return null;
}

function AnalisisFinancieroSection({ rounds, dataById, ourTeam, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const [stmtKey, setStmtKey] = useState("pl");
  const [viewMode, setViewMode] = useState("valores"); // valores | vertical | horizontal

  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Cargá al menos una ronda para ver el análisis.</div>;
  if (!ourTeam) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Elegí cuál es su equipo (arriba) primero.</div>;

  const statements = {
    ...FIN_STATEMENTS,
    mkt: { label: `Informe de mercado${regionFilter !== "Global" ? ` — ${regionFilter}` : " — Global"}`, blockTitle: regionMarketBlockTitle(regionFilter) },
  };
  const effectiveStmtKey = statements[stmtKey] ? stmtKey : "pl";
  const stmtDef = statements[effectiveStmtKey];
  const teamIdx = cesimRounds[0].teams.indexOf(ourTeam);

  const refRound = cesimRounds[cesimRounds.length - 1];
  const refBlock = dataById[refRound.id].blocks.find((b) => b.title === stmtDef.blockTitle);

  const rowsData = useMemo(() => {
    if (!refBlock) return [];
    const metricRows = refBlock.rows.filter((r) => r.kind === "metric");
    return metricRows.map((refRow) => {
      const perRound = cesimRounds.map((r) => {
        const block = dataById[r.id].blocks.find((b) => b.title === stmtDef.blockTitle);
        const row = block ? findMetricRow(block, refRow.label) : null;
        const base = findBaseRow(block, stmtDef);
        const value = row ? toNumberOrNull(row.values[teamIdx]) : null;
        const baseValue = base ? toNumberOrNull(base.values[teamIdx]) : null;
        return { roundNumber: r.roundNumber, value, baseValue };
      });
      return { label: refRow.label, perRound };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [refBlock, cesimRounds, dataById, teamIdx, effectiveStmtKey, regionFilter]);

  const hasBase = !!(stmtDef.baseLabel || stmtDef.baseLabelPattern);

  if (!refBlock) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>No se encontró "{stmtDef.blockTitle}" en la última ronda cargada.</div>;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "center" }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {Object.entries(statements).map(([key, def]) => (
            <button key={key} onClick={() => setStmtKey(key)} style={{ padding: "7px 14px", borderRadius: 7, fontSize: 12.5, cursor: "pointer", fontWeight: 600, border: `1px solid ${effectiveStmtKey === key ? T.amber : T.border}`, background: effectiveStmtKey === key ? T.amberDim : "transparent", color: effectiveStmtKey === key ? T.amber : T.text }}>{def.label}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {[["valores", "Valores"], ["vertical", "Análisis vertical"], ["horizontal", "Análisis horizontal"]].map(([key, lbl]) => (
            <button key={key} onClick={() => setViewMode(key)} style={{ padding: "7px 14px", borderRadius: 7, fontSize: 12.5, cursor: "pointer", fontWeight: 600, border: `1px solid ${viewMode === key ? T.cyan : T.border}`, background: viewMode === key ? T.cyanDim : "transparent", color: viewMode === key ? T.cyan : T.text }}>{lbl}</button>
          ))}
          <InfoBubble text={
            viewMode === "vertical"
              ? (hasBase ? `Cada línea como % de "${stmtDef.baseLabel || "Total Activos"}" en esa misma ronda — muestra la estructura relativa del estado.` : "El informe de mercado no tiene una única línea base común (mezcla precio, unidades, %) — usá Valores u Horizontal para este informe.")
              : viewMode === "horizontal"
              ? "Variación % de cada línea respecto a la ronda anterior — muestra tendencia y crecimiento período a período."
              : "Valores tal como los reporta Cesim para esa ronda."
          } />
        </div>
      </div>
      {!hasBase && viewMode === "vertical" ? (
        <Panel style={{ padding: 18, color: T.textDim, fontSize: 13 }}>El informe de mercado combina precio, unidades y porcentajes en distintas unidades — no tiene una base única para un % vertical. Probá "Valores" o "Análisis horizontal".</Panel>
      ) : (
      <Panel style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "14px 16px 6px" }}>
          <Eyebrow>{stmtDef.label} — {ourTeam}</Eyebrow>
        </div>
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
            <thead>
              <tr>
                <th style={{ position: "sticky", left: 0, background: T.panelAlt, textAlign: "left", padding: "7px 12px", borderBottom: `1px solid ${T.border}`, color: T.textDim, whiteSpace: "nowrap" }}>Concepto</th>
                {cesimRounds.map((r) => <th key={r.id} style={{ textAlign: "right", padding: "7px 12px", borderBottom: `1px solid ${T.border}`, color: T.textDim, whiteSpace: "nowrap" }}>R{r.roundNumber}</th>)}
              </tr>
            </thead>
            <tbody>
              {rowsData.map((row, ri) => (
                <tr key={ri} style={{ background: ri % 2 ? T.panel : T.panelAlt }}>
                  <td style={{ position: "sticky", left: 0, padding: "6px 12px", color: T.text, whiteSpace: "nowrap", borderBottom: `1px solid ${T.borderSoft}`, background: "inherit" }}>{row.label}</td>
                  {row.perRound.map((cell, ci) => {
                    let display = "—";
                    let color = T.text;
                    if (viewMode === "valores") {
                      display = fmtCell(cell.value);
                    } else if (viewMode === "vertical") {
                      if (cell.value !== null && cell.baseValue) { const pct = (cell.value / cell.baseValue) * 100; display = fmtPct(pct); }
                    } else if (viewMode === "horizontal") {
                      const prev = ci > 0 ? row.perRound[ci - 1] : null;
                      if (prev && prev.value !== null && cell.value !== null && prev.value !== 0) {
                        const pct = ((cell.value - prev.value) / Math.abs(prev.value)) * 100;
                        display = fmtPct(pct);
                        color = pct > 0 ? T.green : pct < 0 ? T.red : T.text;
                      }
                    }
                    return <td key={ci} style={{ padding: "6px 12px", textAlign: "right", color, borderBottom: `1px solid ${T.borderSoft}`, whiteSpace: "nowrap" }}>{display}</td>;
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Panel>
      )}
    </div>
  );
}

/* =================================================================
   RESUMEN PARA ACCIONISTAS — pick real widgets from other tabs into
   one board; double-click a block to jump to where it lives.
================================================================= */
function VmsTrendWidget() {
  const [data, setData] = useState(null);
  useEffect(() => { loadVmsData().then((d) => setData(d || DEFAULT_VMS_DATA)); }, []);
  if (!data) return null;
  const totalVar = data.totalCur - data.totalPrev;
  const trendData = [{ label: data.labelPrev, value: data.totalPrev }, { label: data.labelCur, value: data.totalCur }];
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Cuota de mercado en VALOR — qué % de los ingresos totales del mercado captura ULTI.">Cuota de valor de mercado (VMS)</Eyebrow>
      <div style={{ display: "flex", alignItems: "baseline", gap: 14, marginBottom: 12 }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 28, fontWeight: 700, color: totalVar < 0 ? T.red : T.green }}>{data.totalCur.toFixed(2)}%</div>
        <div style={{ fontSize: 12.5, color: T.textDim }}>vs. {data.totalPrev.toFixed(2)}% en {data.labelPrev} <span style={{ fontWeight: 700, color: totalVar < 0 ? T.red : T.green }}>({totalVar >= 0 ? "+" : ""}{totalVar.toFixed(2)} p.p.)</span></div>
      </div>
      <div style={{ height: 140 }}>
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={trendData}>
            <CartesianGrid stroke={T.borderSoft} strokeDasharray="3 3" />
            <XAxis dataKey="label" stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} />
            <YAxis stroke={T.textFaint} tick={{ fill: T.textDim, fontSize: 11 }} unit="%" />
            <Tooltip contentStyle={{ background: T.panelAlt, border: `1px solid ${T.border}`, borderRadius: 8, fontSize: 12 }} formatter={(v) => [`${v}%`, "VMS"]} />
            <Line dataKey="value" stroke={T.red} strokeWidth={2.5} dot={{ r: 4 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>
    </Panel>
  );
}
function NewsWidget() {
  const [data, setData] = useState(null);
  useEffect(() => { loadVmsData().then((d) => setData(d || DEFAULT_VMS_DATA)); }, []);
  if (!data || !data.news?.length) return null;
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Movimientos de otros equipos que terminaron beneficiando o perjudicando a ULTI.">Novedades del entorno</Eyebrow>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {data.news.map((n, i) => {
          const style = NEWS_IMPACT_STYLE[n.impact] || NEWS_IMPACT_STYLE.neutro;
          return (
            <div key={i} style={{ border: `1px solid ${style.border}`, background: style.bg, borderRadius: 10, padding: "12px 14px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 6 }}>
                <span style={{ fontSize: 13.5, fontWeight: 700, color: T.text }}>{n.title}</span>
                <span style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: style.color, background: T.panel, border: `1px solid ${style.border}`, borderRadius: 20, padding: "2px 8px" }}>{style.label}</span>
              </div>
              {n.stat && <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 5 }}><span style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 18, fontWeight: 700, color: style.color }}>{n.stat.value}</span><span style={{ fontSize: 11.5, color: T.textDim }}>{n.stat.label}</span></div>}
              <div style={{ fontSize: 12.5, color: T.text, lineHeight: 1.5 }}>{renderBoldText(n.body)}</div>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}
function PlanWidget() {
  const [plan, setPlan] = useState(null); const [loaded, setLoaded] = useState(false);
  useEffect(() => { loadStrategyPlan().then((p) => { setPlan(p); setLoaded(true); }); }, []);
  if (!loaded || !plan) return null;
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="El plan estratégico vigente, con precios estimados por región y tecnología.">Plan estratégico — ronda actual</Eyebrow>
      <PlanSummaryView plan={plan} />
    </Panel>
  );
}

const SUMMARY_WIDGETS = [
  { id: "resumen-ejecutivo", label: "Resumen ejecutivo", sourceTab: "dashboard", Component: (ctx) => <ExecutiveSummaryCard lines={buildExecutiveSummary(ctx.latestData, ctx.prevData, ctx.teamIdx, ctx.ourTeam, ctx.latestRound)} ourTeam={ctx.ourTeam} /> },
  { id: "retorno-accionista", label: "Retorno al accionista", sourceTab: "dashboard", Component: (ctx) => <ShareholderReturnCard latestData={ctx.latestData} prevData={ctx.prevData} ourTeam={ctx.ourTeam} teamIdx={ctx.teamIdx} /> },
  { id: "tendencia-kpis", label: "Tendencia de indicadores", sourceTab: "dashboard", Component: (ctx) => <KpiTrendChart rounds={ctx.rounds} dataById={ctx.dataById} teamIdx={ctx.teamIdx} regionFilter={ctx.regionFilter} /> },
  { id: "rivales-cercanos", label: "Rivales más cercanos", sourceTab: "dashboard", Component: (ctx) => <ClosestRivalsTable latestData={ctx.latestData} ourTeam={ctx.ourTeam} /> },
  { id: "mapa-posicionamiento", label: "Mapa de posicionamiento", sourceTab: "competencia", Component: (ctx) => <PositioningMap roundData={ctx.latestData} ourTeam={ctx.ourTeam} region={ctx.regionFilter} /> },
  { id: "mix-tecnologia", label: "Mix de tecnología", sourceTab: "competencia", Component: (ctx) => <TechMixChart roundData={ctx.latestData} ourTeam={ctx.ourTeam} region={ctx.regionFilter} /> },
  { id: "enfoque-marca", label: "Enfoque de marca — comparación de precios", sourceTab: "competencia", Component: (ctx) => <CompetitorFocusTable roundData={ctx.latestData} ourTeam={ctx.ourTeam} /> },
  { id: "todo-el-mercado", label: "Todo el mercado", sourceTab: "competencia", Component: (ctx) => <MarketOverviewTable roundData={{ ...ctx.latestData, roundNumber: ctx.latestRound.roundNumber }} ourTeam={ctx.ourTeam} /> },
  { id: "vms", label: "Cuota de valor de mercado (VMS)", sourceTab: "general", Component: () => <VmsTrendWidget /> },
  { id: "novedades", label: "Novedades del entorno", sourceTab: "general", Component: () => <NewsWidget /> },
  { id: "plan-estrategico", label: "Plan estratégico — precios estimados", sourceTab: "estrategia", Component: () => <PlanWidget /> },
  { id: "promocion-evolucion", label: "Evolución del gasto de promoción", sourceTab: "dashboard", Component: (ctx) => <PromotionTrendCard rounds={ctx.rounds} dataById={ctx.dataById} ourTeam={ctx.ourTeam} teamIdx={ctx.teamIdx} /> },
  { id: "promocion-rivales", label: "Variación de promoción vs. rivales", sourceTab: "dashboard", Component: (ctx) => <PromotionRivalsCard latestData={ctx.latestData} prevData={ctx.prevData} ourTeam={ctx.ourTeam} /> },
  { id: "promocion-region", label: "Promoción por región", sourceTab: "dashboard", Component: (ctx) => <PromotionByRegionCard latestData={ctx.latestData} prevData={ctx.prevData} ourTeam={ctx.ourTeam} /> },
];

async function loadSummaryBoard() {
  try { const res = await storage.get("summary-board"); return res ? JSON.parse(res.value) : ["resumen-ejecutivo", "retorno-accionista", "vms", "novedades"]; } catch { return ["resumen-ejecutivo", "retorno-accionista", "vms", "novedades"]; }
}
async function saveSummaryBoard(ids) { await storage.set("summary-board", JSON.stringify(ids)); }

function AccionistasSummarySection({ rounds, dataById, ourTeam, regionFilter, onNavigate }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const teamIdx = ourTeam && cesimRounds.length ? cesimRounds[0].teams.indexOf(ourTeam) : -1;
  const latestRound = cesimRounds[cesimRounds.length - 1];
  const latestData = latestRound ? dataById[latestRound.id] : null;
  const prevRound = cesimRounds[cesimRounds.length - 2];
  const prevData = prevRound ? dataById[prevRound.id] : null;

  const [board, setBoard] = useState(null);
  const [picking, setPicking] = useState(false);
  useEffect(() => { loadSummaryBoard().then(setBoard); }, []);

  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Cargá al menos una ronda para armar el resumen.</div>;
  if (teamIdx < 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Elegí cuál es su equipo (arriba) primero.</div>;
  if (!board) return <div style={{ padding: 20, color: T.textFaint, display: "flex", alignItems: "center", gap: 8 }}><Loader2 size={14} className="spin" /> Cargando…</div>;

  const ctx = { rounds, dataById, ourTeam, regionFilter, teamIdx, latestData, prevData, latestRound };
  const available = SUMMARY_WIDGETS.filter((w) => !board.includes(w.id));

  const addWidget = async (id) => { const next = [...board, id]; await saveSummaryBoard(next); setBoard(next); setPicking(false); };
  const removeWidget = async (id) => { const next = board.filter((x) => x !== id); await saveSummaryBoard(next); setBoard(next); };
  const move = async (idx, dir) => {
    const next = [...board];
    const j = idx + dir;
    if (j < 0 || j >= next.length) return;
    [next[idx], next[j]] = [next[j], next[idx]];
    await saveSummaryBoard(next); setBoard(next);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 12.5, color: T.textDim, display: "flex", alignItems: "center", gap: 6 }}>
          <Info size={13} color={T.cyan} /> Doble click en cualquier bloque para ir a la pestaña de donde sale ese dato.
        </div>
        <button onClick={() => setPicking((p) => !p)} style={primaryBtn}><BookPlus size={13} /> Agregar bloque</button>
      </div>

      {picking && (
        <Panel style={{ padding: 14 }}>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {available.length === 0 && <div style={{ color: T.textFaint, fontSize: 12.5 }}>Ya agregaste todos los bloques disponibles.</div>}
            {available.map((w) => (
              <button key={w.id} onClick={() => addWidget(w.id)} style={{ padding: "8px 13px", borderRadius: 8, fontSize: 12.5, cursor: "pointer", fontWeight: 600, border: `1px solid ${T.border}`, background: T.panelAlt, color: T.text }}>+ {w.label}</button>
            ))}
          </div>
        </Panel>
      )}

      {board.length === 0 && <div style={{ textAlign: "center", padding: "50px 20px", color: T.textFaint }}>Todavía no armaste el resumen. Usá "Agregar bloque" para traer datos y gráficos de las otras pestañas.</div>}

      {board.map((id, idx) => {
        const w = SUMMARY_WIDGETS.find((x) => x.id === id);
        if (!w) return null;
        return (
          <div key={id} onDoubleClick={() => onNavigate(w.sourceTab)} style={{ position: "relative", cursor: "pointer" }} title={`Doble click para ir a "${TAB_LABELS[w.sourceTab] || w.sourceTab}"`}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, padding: "0 2px" }}>
              <span style={{ fontSize: 10.5, color: T.textFaint, fontFamily: "'IBM Plex Mono', monospace", textTransform: "uppercase", letterSpacing: "0.06em" }}>
                fuente: {TAB_LABELS[w.sourceTab] || w.sourceTab}
              </span>
              <div className="no-print" style={{ display: "flex", gap: 2 }}>
                <button onClick={(e) => { e.stopPropagation(); move(idx, -1); }} disabled={idx === 0} style={{ background: "none", border: "none", color: T.textFaint, cursor: idx === 0 ? "default" : "pointer", opacity: idx === 0 ? 0.3 : 1, padding: 3, fontSize: 12 }}>▲</button>
                <button onClick={(e) => { e.stopPropagation(); move(idx, 1); }} disabled={idx === board.length - 1} style={{ background: "none", border: "none", color: T.textFaint, cursor: idx === board.length - 1 ? "default" : "pointer", opacity: idx === board.length - 1 ? 0.3 : 1, padding: 3, fontSize: 12 }}>▼</button>
                <button onClick={(e) => { e.stopPropagation(); removeWidget(id); }} style={{ background: "none", border: "none", color: T.textFaint, cursor: "pointer", padding: 3 }}><Trash2 size={12} /></button>
              </div>
            </div>
            {w.Component(ctx)}
          </div>
        );
      })}
    </div>
  );
}

/* =================================================================
   GENERIC SHEET SECTION (fallback for non-Cesim files)
================================================================= */
function GenericSheetSection({ sheetName, rounds, dataById }) {
  const roundsWithSheet = rounds.filter((r) => dataById[r.id]?.kind === "generic" && dataById[r.id]?.sheets?.[sheetName]);
  const [selectedRoundId, setSelectedRoundId] = useState(roundsWithSheet.length ? roundsWithSheet[roundsWithSheet.length - 1].id : null);
  useEffect(() => { if (roundsWithSheet.length && !roundsWithSheet.find((r) => r.id === selectedRoundId)) setSelectedRoundId(roundsWithSheet[roundsWithSheet.length - 1].id); }, [roundsWithSheet.map((r) => r.id).join(",")]);
  if (roundsWithSheet.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Sin datos para esta hoja.</div>;
  const activeSheet = dataById[selectedRoundId].sheets[sheetName];
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ overflowX: "auto", maxHeight: 420 }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <thead><tr>{activeSheet.headers.map((h, i) => <th key={i} style={{ position: "sticky", top: 0, background: T.panelAlt, color: T.textDim, textAlign: "left", padding: "8px 12px", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" }}>{h || `Col ${i + 1}`}</th>)}</tr></thead>
          <tbody>
            {activeSheet.rows.map((row, ri) => (
              <tr key={ri} style={{ background: ri % 2 ? T.panel : T.panelAlt }}>
                {activeSheet.headers.map((_, ci) => <td key={ci} style={{ padding: "7px 12px", color: T.text, whiteSpace: "nowrap", borderBottom: `1px solid ${T.borderSoft}` }}>{row[ci] === undefined || row[ci] === "" ? "—" : String(row[ci])}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}

/* =================================================================
   ROOT APP
================================================================= */
const CATEGORY_ORDER = ["Financiero", "Mercado", "Producción", "RRHH", "ESG", "Otros"];

function BlockTable({ block, teams, ourTeam }) {
  return (
    <Panel style={{ padding: 0, overflow: "hidden", marginBottom: 14 }}>
      <div style={{ padding: "12px 16px 8px" }}><Eyebrow>{block.title}</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <thead><tr>
            <th style={{ position: "sticky", left: 0, background: T.panelAlt, textAlign: "left", padding: "7px 12px", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" }} />
            {teams.map((tm) => <th key={tm} style={{ background: tm === ourTeam ? T.amberDim : T.panelAlt, color: tm === ourTeam ? T.amber : T.textDim, textAlign: "right", padding: "7px 12px", borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" }}>{tm}</th>)}
          </tr></thead>
          <tbody>
            {block.rows.map((row, ri) => row.kind === "group" ? (
              <tr key={ri}><td colSpan={teams.length + 1} style={{ padding: "8px 12px 4px", color: T.textDim, fontWeight: 600, fontFamily: "'Inter', sans-serif", fontSize: 11.5, borderBottom: `1px solid ${T.borderSoft}` }}>{row.label}</td></tr>
            ) : (
              <tr key={ri} style={{ background: ri % 2 ? T.panel : T.panelAlt }}>
                <td style={{ position: "sticky", left: 0, padding: "6px 12px", color: T.text, whiteSpace: "nowrap", borderBottom: `1px solid ${T.borderSoft}`, background: "inherit" }}>{row.label}</td>
                {row.values.map((v, ci) => <td key={ci} style={{ padding: "6px 12px", textAlign: "right", whiteSpace: "nowrap", borderBottom: `1px solid ${T.borderSoft}`, color: teams[ci] === ourTeam ? T.amber : T.text, fontWeight: teams[ci] === ourTeam ? 600 : 400 }}>{fmtCell(v)}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}
function CategorySection({ category, rounds, dataById, ourTeam, regionFilter = "Global" }) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const [selectedId, setSelectedId] = useState(cesimRounds.length ? cesimRounds[cesimRounds.length - 1].id : null);
  useEffect(() => { if (cesimRounds.length && !cesimRounds.find((r) => r.id === selectedId)) setSelectedId(cesimRounds[cesimRounds.length - 1].id); }, [cesimRounds.map((r) => r.id).join(",")]);
  if (cesimRounds.length === 0) return <div style={{ color: T.textFaint, fontSize: 13, padding: 20 }}>Sin datos todavía.</div>;
  const roundData = dataById[selectedId];
  const teams = roundData.teams;
  const blocksAll = roundData.blocks.filter((b) => b.category === category);
  const blocks = regionFilter === "Global"
    ? blocksAll
    : blocksAll.filter((b) => {
        const t = b.title.toLowerCase();
        const mentionsOtherRegion = REGIONS.filter((r) => r !== regionFilter).some((r) => t.includes(r.toLowerCase()));
        return !mentionsOtherRegion;
      });
  return (
    <div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 16 }}>
        <span style={{ fontSize: 12, color: T.textDim }}>Ronda:</span>
        {cesimRounds.map((r) => <button key={r.id} onClick={() => setSelectedId(r.id)} style={{ padding: "5px 10px", borderRadius: 6, fontSize: 12, cursor: "pointer", border: `1px solid ${r.id === selectedId ? T.amber : T.border}`, background: r.id === selectedId ? T.amberDim : "transparent", color: r.id === selectedId ? T.amber : T.textDim }}>R{r.roundNumber}</button>)}
      </div>
      {blocks.length === 0 ? <div style={{ color: T.textFaint, fontSize: 13 }}>Esta ronda no tiene tablas en esta categoría{regionFilter !== "Global" ? ` para ${regionFilter}` : ""}.</div> : blocks.map((b, i) => <BlockTable key={i} block={b} teams={teams} ourTeam={ourTeam} />)}
    </div>
  );
}

function TabButton({ children, icon, active, onClick }) {
  return (
    <button onClick={onClick} className="nav-tab" style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", padding: "9px 4px", marginRight: 20, borderBottom: `2px solid ${active ? T.amber : "transparent"}`, color: active ? T.text : T.textFaint, fontSize: 13, fontWeight: active ? 600 : 500, whiteSpace: "nowrap" }}>
      {icon} {children}
    </button>
  );
}

const TAB_LABELS = {
  resumen: "Resumen Ejecutivo", mercado: "Mercado y Demanda", produccion: "Producción y Operaciones",
  finanzas: "Finanzas", ratios: "Ratios y Valoración", decisiones: "Decisiones y Registro", asistente: "Asistente IA",
  dashboard: "Dashboard", general: "Value Market Share", premisas: "Premisas", competencia: "Competencia",
  dinamico: "Panel dinámico", estrategia: "Estrategia", financiero: "Análisis Financiero",
  accionistas: "Resumen para Accionistas",
};

/* Solapas principales (áreas de decisión del simulador). Los ids viejos se conservan como alias
   para que los widgets que navegan con "dashboard", "general", etc. sigan funcionando. */
const MAIN_TABS = ["resumen", "mercado", "produccion", "finanzas", "ratios", "decisiones"];
// Categorías crudas del simulador que ahora se muestran dentro de una solapa principal.
const CATEGORIES_IN_MAIN_TABS = ["Financiero", "Mercado", "Producción"];
const LEGACY_TAB_ALIAS = {
  dashboard: "resumen", accionistas: "resumen",
  general: "mercado", competencia: "mercado",
  financiero: "finanzas", dinamico: "ratios",
  estrategia: "decisiones", premisas: "decisiones",
};

/* ---------------------------------------------------------------
   HELPERS DE DATOS PARA LAS SOLAPAS NUEVAS
   El .xls cambia de forma entre rondas (y el parser a veces titula mal un bloque,
   ej. "Tec 4"), así que se busca por título estable, por grupo o por etiqueta —
   nunca por número de fila.
--------------------------------------------------------------- */
const MARKET_REGIONS = [
  { key: "EE.UU.", currency: "USD" }, { key: "Asia", currency: "RMB" }, { key: "Europa", currency: "EUR" },
];
const TL_COLORS = { green: T.green, yellow: "#D97706", red: T.red, none: T.textFaint };

function fmtInt(n) { return Number.isFinite(n) ? n.toLocaleString("es-AR", { maximumFractionDigits: 0 }) : "—"; }
function fmtSigned(n, digits = 1, suffix = "") { return Number.isFinite(n) ? `${n >= 0 ? "+" : ""}${n.toFixed(digits)}${suffix}` : "—"; }
function avg(list) { const xs = list.filter((v) => Number.isFinite(v)); return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null; }

function getRoundContext(rounds, dataById, ourTeam) {
  const cesimRounds = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]);
  const latestRound = cesimRounds[cesimRounds.length - 1] || null;
  const prevRound = cesimRounds[cesimRounds.length - 2] || null;
  const latestData = latestRound ? dataById[latestRound.id] : null;
  const prevData = prevRound ? dataById[prevRound.id] : null;
  const teamIdx = latestData && ourTeam ? latestData.teams.indexOf(ourTeam) : -1;
  return { cesimRounds, latestRound, latestData, prevData, teamIdx };
}
function findBlockWithGroup(roundData, re) {
  return roundData?.blocks?.find((b) => b.rows.some((r) => r.kind === "group" && re.test(r.label))) || null;
}
function findMetricAnywhere(roundData, re) {
  for (const b of roundData?.blocks || []) { const row = b.rows.find((r) => r.kind === "metric" && re.test(r.label)); if (row) return row; }
  return null;
}
function getRatioBy(roundData, re, teamIdx) {
  const block = roundData?.blocks?.find((b) => b.title === RATIOS_BLOCK_TITLE);
  const row = block?.rows.find((r) => r.kind === "metric" && re.test(r.label));
  return row ? toNumberOrNull(row.values[teamIdx]) : null;
}
function getKpiByKey(roundData, key, teamIdx) { return extractKpi(roundData, teamIdx, KPI_DEFS.find((d) => d.key === key)); }
// "Capacidad Disponible" (capacidad ociosa, %) cuelga de la región, pero Tec 3/Tec 4 vacíos
// aparecen como grupos intermedios, por eso no sirve getMetricUnderGroup.
function getIdleCapacity(roundData, region, teamIdx) {
  const block = findBlockWithGroup(roundData, /^Capacidad empleada/i);
  if (!block) return null;
  const rows = block.rows;
  const start = rows.findIndex((r) => r.kind === "group" && r.label === region);
  if (start < 0) return null;
  for (let i = start + 1; i < rows.length; i++) {
    const r = rows[i];
    if (r.kind === "group" && !/^Tec \d/.test(r.label)) break;
    if (r.kind === "metric" && /^Capacidad Disponible/i.test(r.label)) return toNumberOrNull(r.values[teamIdx]);
  }
  return null;
}

// Tecnologías presentes en una ronda ("Tec 1", "Tec 2"… las que existan en el archivo), en orden numérico.
// Así, si una ronda futura agrega Tec 5 o más, aparece sola sin tocar el código.
const _techCache = new WeakMap();
function techsInRound(roundData) {
  if (!roundData?.blocks) return TECH_LABELS;
  if (_techCache.has(roundData)) return _techCache.get(roundData);
  const nums = new Set();
  roundData.blocks.forEach((b) => b.rows.forEach((r) => { const m = /^Tec (\d+)(?!\d)/.exec(r.label); if (m) nums.add(parseInt(m[1], 10)); }));
  const list = nums.size ? [...nums].sort((a, b) => a - b).map((n) => `Tec ${n}`) : TECH_LABELS;
  _techCache.set(roundData, list);
  return list;
}

// Una fila por mercado y tecnología que el equipo efectivamente opera/estimó.
function computeMarketRows(roundData, teamIdx) {
  if (!roundData || roundData.kind !== "cesim" || teamIdx < 0) return [];
  const estBlock = roundData.blocks.find((b) => /^Demanda estimada, miles unid/i.test(b.title));
  const rows = [];
  MARKET_REGIONS.forEach(({ key, currency }) => {
    const mBlock = roundData.blocks.find((b) => b.title === regionMarketBlockTitle(key));
    const marginBlock = roundData.blocks.find((b) => b.title === `Desglose de margen por tec, miles USD, ${key}`);
    techsInRound(roundData).forEach((tec) => {
      const sales = getMetricUnderGroup(mBlock, tec, "Ventas, miles unidades", teamIdx);
      const demand = getMetricUnderGroup(mBlock, tec, "Demanda, miles unidades", teamIdx);
      const estDemand = getMetricUnderGroup(estBlock, key, tec, teamIdx);
      if (!((sales || 0) > 0 || (demand || 0) > 0 || (estDemand || 0) > 0)) return;
      rows.push({
        region: key, tec, currency, sales, demand, estDemand,
        price: getMetricUnderGroup(mBlock, tec, `Precio de venta, ${currency}`, teamIdx),
        revenue: getMetricUnderGroup(marginBlock, tec, "Ingresos por ventas", teamIdx),
      });
    });
  });
  return rows;
}

function rankAmongTeams(roundData, getter, ourIdx) {
  if (!roundData || ourIdx < 0) return null;
  const vals = roundData.teams.map((_, i) => ({ i, v: getter(roundData, i) })).filter((x) => Number.isFinite(x.v));
  if (!vals.length) return null;
  vals.sort((a, b) => b.v - a.v);
  const pos = vals.findIndex((x) => x.i === ourIdx);
  return pos < 0 ? null : { rank: pos + 1, of: vals.length };
}

/* ---------------------------------------------------------------
   SEMÁFOROS DE ALERTA
--------------------------------------------------------------- */
function computeTrafficLights(cur, prev, teamIdx, regionFilter = "Global") {
  const lights = [];

  // 1) Fill rate = ventas / demanda (unidades)
  const rows = computeMarketRows(cur, teamIdx).filter((r) => regionFilter === "Global" || r.region === regionFilter);
  const sales = rows.reduce((a, r) => a + (r.sales || 0), 0), demand = rows.reduce((a, r) => a + (r.demand || 0), 0);
  const fill = demand > 0 ? (sales / demand) * 100 : null;
  lights.push({
    key: "fill", label: `Fill rate ${regionFilter === "Global" ? "global" : regionFilter}`, hint: "Verde ≥95% · Amarillo 80–95% · Rojo <80%",
    value: fill === null ? null : fmtPct(fill), status: fill === null ? "none" : fill >= 95 ? "green" : fill >= 80 ? "yellow" : "red",
    detail: fill === null ? "" : `${fmtInt(sales)} vendidas de ${fmtInt(demand)} demandadas (miles u.)`,
  });

  // 2) Capacidad ociosa EE.UU.
  const idle = getIdleCapacity(cur, "EE.UU.", teamIdx);
  lights.push({
    key: "idle", label: "Capacidad ociosa EE.UU.", hint: "Verde <20% · Amarillo 20–50% · Rojo >50%",
    value: idle === null ? null : fmtPct(idle), status: idle === null ? "none" : idle < 20 ? "green" : idle <= 50 ? "yellow" : "red", detail: "",
  });

  // 3) Variación de cuota global vs ronda anterior (p.p.)
  const share = getKpiByKey(cur, "share", teamIdx), sharePrev = prev ? getKpiByKey(prev, "share", teamIdx) : null;
  const dShare = share !== null && sharePrev !== null ? share - sharePrev : null;
  lights.push({
    key: "share", label: "Cuota global vs ronda anterior", hint: "Verde ≥0 · Amarillo 0 a -2 p.p. · Rojo < -2 p.p.",
    value: dShare === null ? null : fmtSigned(dShare, 2, " p.p."), status: dShare === null ? "none" : dShare >= 0 ? "green" : dShare >= -2 ? "yellow" : "red",
    detail: dShare === null ? "Necesita una ronda anterior cargada" : `${sharePrev.toFixed(1)}% → ${share.toFixed(1)}%`,
  });

  // 4) Desvío de ingresos vs estimado
  const dev = computeRevenueDeviation(cur, teamIdx);
  const devPct = dev ? dev.pct : null;
  lights.push({
    key: "dev", label: "Desvío de ingresos vs estimado", hint: "Verde ±5% · Amarillo ±5–15% · Rojo >±15%",
    value: devPct === null ? null : fmtSigned(devPct, 1, "%"), status: devPct === null ? "none" : Math.abs(devPct) <= 5 ? "green" : Math.abs(devPct) <= 15 ? "yellow" : "red",
    detail: dev ? `Real ${fmtInt(dev.real)} vs estimado ${fmtInt(dev.est)} (miles USD)` : "",
  });

  // 5) Margen EBITDA vs promedio de la industria
  const ebitdaRe = /EBITDA/i;
  const ebitda = getRatioBy(cur, ebitdaRe, teamIdx);
  const industry = avg(cur.teams.map((_, i) => getRatioBy(cur, ebitdaRe, i)));
  const dE = ebitda !== null && industry !== null ? ebitda - industry : null;
  lights.push({
    key: "ebitda", label: "Margen EBITDA vs industria", hint: "Verde > promedio · Amarillo hasta -5 p.p. · Rojo < -5 p.p.",
    value: ebitda === null ? null : fmtPct(ebitda), status: dE === null ? "none" : dE > 0 ? "green" : dE >= -5 ? "yellow" : "red",
    detail: dE === null ? "" : `Promedio industria ${fmtPct(industry)} (${fmtSigned(dE, 1, " p.p.")})`,
  });

  // 6) Calificación crediticia
  const rating = getRatio(cur, "Calificación crediticia", teamIdx);
  const rIdx = typeof rating === "string" ? CREDIT_RATING_ORDER.indexOf(rating.trim()) : -1;
  lights.push({
    key: "rating", label: "Calificación crediticia", hint: "Verde A+ o mejor · Amarillo BBB a A · Rojo peor que BBB",
    value: rIdx < 0 ? null : rating, status: rIdx < 0 ? "none" : rIdx <= CREDIT_RATING_ORDER.indexOf("A+") ? "green" : rIdx <= CREDIT_RATING_ORDER.indexOf("BBB-") ? "yellow" : "red", detail: "",
  });
  return lights;
}

function TrafficLightsCard({ latestData, prevData, teamIdx, regionFilter }) {
  if (!latestData || teamIdx < 0) return null;
  const lights = computeTrafficLights(latestData, prevData, teamIdx, regionFilter);
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info="Semáforos sobre la última ronda cargada. Pasá el mouse por cada uno para ver los umbrales.">Semáforos de alerta</Eyebrow>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12 }}>
        {lights.map((l) => (
          <div key={l.key} title={l.hint} style={{ display: "flex", gap: 12, padding: "12px 14px", borderRadius: 12, border: `1px solid ${T.border}`, background: T.panelAlt }}>
            <div style={{ width: 14, height: 14, borderRadius: "50%", marginTop: 3, flexShrink: 0, background: TL_COLORS[l.status], boxShadow: l.status === "none" ? "none" : `0 0 8px ${TL_COLORS[l.status]}66` }} />
            <div style={{ minWidth: 0 }}>
              <div style={{ fontSize: 11, color: T.textFaint, textTransform: "uppercase", letterSpacing: "0.06em" }}>{l.label}</div>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, fontWeight: 700, color: l.value === null ? T.textFaint : T.text }}>{l.value ?? "Sin datos"}</div>
              {l.detail && <div style={{ fontSize: 11, color: T.textDim, marginTop: 2 }}>{l.detail}</div>}
            </div>
          </div>
        ))}
      </div>
    </Panel>
  );
}

/* ---------------------------------------------------------------
   KPIs PRINCIPALES (valor, variación vs ronda anterior, ranking)
--------------------------------------------------------------- */
const RESUMEN_KPIS = [
  { key: "revenue", label: "Ingresos", get: (rd, i) => getKpiByKey(rd, "revenue", i), fmt: fmtNumber },
  { key: "profit", label: "Beneficio neto", get: (rd, i) => getKpiByKey(rd, "profit", i), fmt: fmtNumber },
  { key: "ebitda", label: "Margen EBITDA", get: (rd, i) => getRatioBy(rd, /EBITDA/i, i), fmt: (v) => fmtPct(v) },
  { key: "roe", label: "ROE", get: (rd, i) => getRatioBy(rd, /\(ROE\)/i, i), fmt: (v) => fmtPct(v) },
  { key: "eps", label: "EPS (USD)", get: (rd, i) => getRatioBy(rd, /\(EPS\)/i, i), fmt: (v) => v.toFixed(2) },
  { key: "share", label: "Cuota global", get: (rd, i) => getKpiByKey(rd, "share", i), fmt: (v) => fmtPct(v) },
];
function KpisPrincipalesCards({ latestData, prevData, teamIdx }) {
  if (!latestData || teamIdx < 0) return null;
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12 }}>
      {RESUMEN_KPIS.map((k) => {
        const v = k.get(latestData, teamIdx);
        if (v === null) return null;
        const before = prevData ? k.get(prevData, teamIdx) : null;
        return <StatCard key={k.key} label={k.label} value={k.fmt(v)} delta={before !== null ? v - before : null} rank={rankAmongTeams(latestData, k.get, teamIdx)} />;
      })}
    </div>
  );
}

/* ---------------------------------------------------------------
   DESVÍO PRESUPUESTARIO (estimado vs real)
--------------------------------------------------------------- */
// Estimado: "Demanda estimada, miles USD" · Real: ingresos por ventas del P&L global.
function computeRevenueDeviation(roundData, teamIdx) {
  const estRow = findMetricAnywhere(roundData, /^Demanda estimada, miles USD/i);
  const est = estRow ? toNumberOrNull(estRow.values[teamIdx]) : null;
  const real = getKpiByKey(roundData, "revenue", teamIdx);
  if (est === null || real === null || est === 0) return null;
  return { est, real, diff: real - est, pct: ((real - est) / est) * 100 };
}
function DevCell({ pct }) {
  if (!Number.isFinite(pct)) return <span style={{ color: T.textFaint }}>—</span>;
  const a = Math.abs(pct);
  return <span style={{ color: TL_COLORS[a <= 5 ? "green" : a <= 15 ? "yellow" : "red"], fontWeight: 600 }}>{fmtSigned(pct, 1, "%")}</span>;
}
function DesvioPresupuestarioCard({ latestData, teamIdx, regionFilter }) {
  if (!latestData || teamIdx < 0) return null;
  const rows = computeMarketRows(latestData, teamIdx).filter((r) => regionFilter === "Global" || r.region === regionFilter).map((r) => {
    // El .xls solo trae ingresos estimados en total; por mercado se valoriza la demanda estimada
    // con el ingreso unitario real (ingresos reales / unidades vendidas), en USD.
    const unitRev = r.revenue !== null && r.sales > 0 ? r.revenue / r.sales : null;
    const estRevenue = unitRev !== null && r.estDemand !== null ? r.estDemand * unitRev : null;
    return { ...r, estRevenue, demandDev: r.estDemand > 0 && r.demand !== null ? ((r.demand - r.estDemand) / r.estDemand) * 100 : null, revDiff: estRevenue !== null && r.revenue !== null ? r.revenue - estRevenue : null };
  });
  const total = computeRevenueDeviation(latestData, teamIdx);
  const estProfitRow = findMetricAnywhere(latestData, /^Ingresos netos proyectados/i);
  const estProfit = estProfitRow ? toNumberOrNull(estProfitRow.values[teamIdx]) : null;
  const realProfit = getKpiByKey(latestData, "profit", teamIdx);
  const th = { textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" };
  const td = { padding: "6px 12px", textAlign: "right", color: T.text };
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Compara lo que el simulador estimó al decidir contra lo que realmente pasó. Unidades en miles; ingresos en miles USD. Por mercado, el ingreso estimado es la demanda estimada valorizada al ingreso unitario real.">Desvío presupuestario — estimado vs real</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <thead><tr>
            <th style={{ ...th, textAlign: "left" }}>Mercado</th><th style={{ ...th, textAlign: "left" }}>Tec</th>
            <th style={th}>Demanda estimada</th><th style={th}>Demanda real</th><th style={th}>Desvío %</th>
            <th style={th}>Ingresos estimados</th><th style={th}>Ingresos reales</th><th style={th}>Desvío USD</th>
          </tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.region + r.tec} style={{ background: i % 2 ? T.panel : T.panelAlt }}>
                <td style={{ ...td, textAlign: "left" }}>{r.region}</td><td style={{ ...td, textAlign: "left" }}>{r.tec}</td>
                <td style={td}>{fmtInt(r.estDemand)}</td><td style={td}>{fmtInt(r.demand)}</td><td style={td}><DevCell pct={r.demandDev} /></td>
                <td style={td}>{fmtInt(r.estRevenue)}</td><td style={td}>{fmtInt(r.revenue)}</td>
                <td style={{ ...td, color: r.revDiff === null ? T.textFaint : r.revDiff >= 0 ? T.green : T.red, fontWeight: 600 }}>{r.revDiff === null ? "—" : `${r.revDiff >= 0 ? "+" : ""}${fmtInt(r.revDiff)}`}</td>
              </tr>
            ))}
            {regionFilter === "Global" && total && (
              <tr style={{ background: T.amberDim, fontWeight: 700 }}>
                <td style={{ ...td, textAlign: "left", color: T.amber }} colSpan={2}>Total compañía</td>
                <td style={td}>—</td><td style={td}>—</td><td style={td}><DevCell pct={total.pct} /></td>
                <td style={td}>{fmtInt(total.est)}</td><td style={td}>{fmtInt(total.real)}</td>
                <td style={{ ...td, color: total.diff >= 0 ? T.green : T.red }}>{`${total.diff >= 0 ? "+" : ""}${fmtInt(total.diff)}`}</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {estProfit !== null && realProfit !== null && regionFilter === "Global" && (
        <div style={{ padding: "10px 16px 14px", fontSize: 12, color: T.textDim }}>
          Beneficio proyectado {fmtInt(estProfit)} vs real {fmtInt(realProfit)} (miles USD) — <DevCell pct={estProfit !== 0 ? ((realProfit - estProfit) / Math.abs(estProfit)) * 100 : null} />
        </div>
      )}
      {rows.length === 0 && <div style={{ padding: 16, fontSize: 13, color: T.textFaint }}>No se encontraron datos de demanda estimada en la última ronda.</div>}
    </Panel>
  );
}

/* ---------------------------------------------------------------
   MERCADO Y DEMANDA
--------------------------------------------------------------- */
const TH = { textAlign: "right", padding: "7px 12px", color: T.textDim, borderBottom: `1px solid ${T.border}`, whiteSpace: "nowrap" };
const TD = { padding: "6px 12px", textAlign: "right", color: T.text };
const TEAM_PALETTE = [T.cyan, T.violet, T.green, T.red, "#D97706", "#0EA5E9", "#64748B", "#DB2777"];
const MONO_TABLE = { borderCollapse: "collapse", width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 };

// Cuota (%) de un equipo en una región ("Global" o una de MARKET_REGIONS) y tecnología ("Tec 1"… o "Total").
function getShareValue(roundData, region, tec, teamIdx) {
  const block = roundData?.blocks?.find((b) => b.title === regionMarketBlockTitle(region));
  if (!block) return null;
  const gi = block.rows.findIndex((r) => r.kind === "group" && /cuotas de mercado/i.test(r.label));
  if (gi < 0) return null;
  for (let i = gi + 1; i < block.rows.length; i++) {
    const r = block.rows[i];
    if (r.kind === "group") break;
    if (r.kind === "metric" && r.label === tec) return toNumberOrNull(r.values[teamIdx]);
  }
  return null;
}
function getMarketPrice(roundData, region, currency, tec, teamIdx) {
  const block = roundData?.blocks?.find((b) => b.title === regionMarketBlockTitle(region));
  return getMetricUnderGroup(block, tec, `Precio de venta, ${currency}`, teamIdx);
}
function TeamHeaderCells({ teams, ourTeam }) {
  return teams.map((tm) => <th key={tm} style={{ ...TH, color: tm === ourTeam ? T.amber : T.textDim, background: tm === ourTeam ? T.amberDim : "transparent" }}>{tm}</th>);
}

function CuotasHeatmapCard({ latestData, ourTeam, regionFilter }) {
  if (!latestData) return null;
  const teams = latestData.teams;
  const regions = MARKET_REGIONS.filter((r) => regionFilter === "Global" || r.key === regionFilter);
  const rows = [];
  regions.forEach(({ key }) => {
    techsInRound(latestData).forEach((tec) => {
      const vals = teams.map((_, i) => getShareValue(latestData, key, tec, i));
      if (vals.some((v) => (v || 0) > 0)) rows.push({ label: `${key} ${tec}`, vals, total: false });
    });
    const totals = teams.map((_, i) => getShareValue(latestData, key, "Total", i));
    if (totals.some((v) => v !== null)) rows.push({ label: `Total ${key}`, vals: totals, total: true });
  });
  const max = Math.max(1, ...rows.filter((r) => !r.total).flatMap((r) => r.vals.map((v) => v || 0)));
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Cuota de mercado en unidades (%) por región y tecnología. Cuanto más intenso el color, mayor la cuota. El líder de cada fila va en negrita.">Cuotas de mercado por región y tecnología</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={MONO_TABLE}>
          <thead><tr><th style={{ ...TH, textAlign: "left" }}>Mercado</th><TeamHeaderCells teams={teams} ourTeam={ourTeam} /></tr></thead>
          <tbody>
            {rows.map((r) => {
              const top = Math.max(...r.vals.map((v) => v || 0));
              return (
                <tr key={r.label} style={{ borderTop: r.total ? `1px solid ${T.border}` : "none" }}>
                  <td style={{ ...TD, textAlign: "left", fontWeight: r.total ? 700 : 400 }}>{r.label}</td>
                  {r.vals.map((v, i) => {
                    const isUs = teams[i] === ourTeam;
                    const alpha = r.total || !v ? 0 : 0.08 + 0.42 * (v / max);
                    return <td key={teams[i]} style={{ ...TD, background: alpha ? `rgba(79,70,229,${alpha.toFixed(2)})` : (isUs ? T.amberDim : "transparent"), fontWeight: isUs || (v && v === top) ? 700 : 400, color: isUs ? T.amber : T.text, outline: isUs ? `1px solid ${T.amber}55` : "none", outlineOffset: -1 }}>{v === null || v === 0 ? "—" : `${v.toFixed(1)}%`}</td>;
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {rows.length === 0 && <div style={{ padding: 16, fontSize: 13, color: T.textFaint }}>No se encontraron cuotas por región en la última ronda.</div>}
    </Panel>
  );
}

function PreciosCompetenciaCard({ latestData, ourTeam, regionFilter }) {
  if (!latestData) return null;
  const teams = latestData.teams, ourIdx = teams.indexOf(ourTeam);
  const rows = [];
  MARKET_REGIONS.filter((r) => regionFilter === "Global" || r.key === regionFilter).forEach(({ key, currency }) => {
    techsInRound(latestData).forEach((tec) => {
      const vals = teams.map((_, i) => getMarketPrice(latestData, key, currency, tec, i));
      if (vals.some((v) => (v || 0) > 0)) rows.push({ label: `${key} ${tec}`, currency, vals, mean: avg(vals.filter((v) => v > 0)) });
    });
  });
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Precio de venta por mercado y tecnología, en la moneda local de cada mercado. El promedio considera solo equipos que venden esa tecnología.">Precios vs competencia</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={MONO_TABLE}>
          <thead><tr><th style={{ ...TH, textAlign: "left" }}>Mercado</th><TeamHeaderCells teams={teams} ourTeam={ourTeam} /><th style={TH}>Promedio</th><th style={TH}>{ourTeam || "Nosotros"} vs prom.</th></tr></thead>
          <tbody>
            {rows.map((r) => {
              const mine = ourIdx >= 0 ? r.vals[ourIdx] : null;
              const diff = mine > 0 && r.mean ? ((mine - r.mean) / r.mean) * 100 : null;
              return (
                <tr key={r.label}>
                  <td style={{ ...TD, textAlign: "left" }}>{r.label} <span style={{ color: T.textFaint }}>({r.currency})</span></td>
                  {r.vals.map((v, i) => <td key={teams[i]} style={{ ...TD, background: teams[i] === ourTeam ? T.amberDim : "transparent", color: teams[i] === ourTeam ? T.amber : T.text, fontWeight: teams[i] === ourTeam ? 700 : 400 }}>{v > 0 ? fmtInt(v) : "—"}</td>)}
                  <td style={TD}>{fmtInt(r.mean)}</td>
                  <td style={{ ...TD, fontWeight: 700, color: diff === null ? T.textFaint : diff >= 0 ? T.violet : T.cyan }}>{diff === null ? "—" : `${diff >= 0 ? "▲" : "▼"} ${fmtSigned(diff, 1, "%")}`}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {rows.length === 0 && <div style={{ padding: 16, fontSize: 13, color: T.textFaint }}>No se encontraron precios en la última ronda.</div>}
    </Panel>
  );
}

function FillRateCard({ latestData, teamIdx, regionFilter }) {
  if (!latestData || teamIdx < 0) return null;
  const rows = computeMarketRows(latestData, teamIdx).filter((r) => regionFilter === "Global" || r.region === regionFilter)
    .map((r) => ({ ...r, fill: r.demand > 0 ? ((r.sales || 0) / r.demand) * 100 : null, lost: Math.max(0, (r.demand || 0) - (r.sales || 0)) }));
  const status = (f) => (f === null ? "none" : f >= 95 ? "green" : f >= 80 ? "yellow" : "red");
  const totDemand = rows.reduce((a, r) => a + (r.demand || 0), 0), totSales = rows.reduce((a, r) => a + (r.sales || 0), 0);
  const totFill = totDemand > 0 ? (totSales / totDemand) * 100 : null;
  const Dot = ({ s }) => <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: "50%", background: TL_COLORS[s], verticalAlign: "middle" }} />;
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Fill rate = ventas / demanda, en miles de unidades. Verde ≥95%, amarillo 80–95%, rojo <80%. Las unidades perdidas son demanda que no pudimos atender.">Fill rate por mercado y tecnología</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={MONO_TABLE}>
          <thead><tr><th style={{ ...TH, textAlign: "left" }}>Mercado</th><th style={{ ...TH, textAlign: "left" }}>Tec</th><th style={TH}>Demanda</th><th style={TH}>Ventas</th><th style={TH}>Fill rate</th><th style={TH}>Unidades perdidas</th></tr></thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={r.region + r.tec} style={{ background: i % 2 ? T.panel : T.panelAlt }}>
                <td style={{ ...TD, textAlign: "left" }}>{r.region}</td><td style={{ ...TD, textAlign: "left" }}>{r.tec}</td>
                <td style={TD}>{fmtInt(r.demand)}</td><td style={TD}>{fmtInt(r.sales)}</td>
                <td style={{ ...TD, fontWeight: 700, color: TL_COLORS[status(r.fill)] }}><Dot s={status(r.fill)} /> {r.fill === null ? "—" : fmtPct(r.fill)}</td>
                <td style={{ ...TD, color: r.lost > 0 ? T.red : T.textFaint }}>{r.lost > 0 ? fmtInt(r.lost) : "—"}</td>
              </tr>
            ))}
            {rows.length > 0 && (
              <tr style={{ background: T.amberDim, fontWeight: 700 }}>
                <td style={{ ...TD, textAlign: "left", color: T.amber }} colSpan={2}>Total {regionFilter === "Global" ? "global" : regionFilter}</td>
                <td style={TD}>{fmtInt(totDemand)}</td><td style={TD}>{fmtInt(totSales)}</td>
                <td style={{ ...TD, color: TL_COLORS[status(totFill)] }}><Dot s={status(totFill)} /> {totFill === null ? "—" : fmtPct(totFill)}</td>
                <td style={{ ...TD, color: T.red }}>{fmtInt(totDemand - totSales)}</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {rows.length === 0 && <div style={{ padding: 16, fontSize: 13, color: T.textFaint }}>No se encontraron ventas y demanda en la última ronda.</div>}
    </Panel>
  );
}

function CuotaEvolucionChart({ rounds, dataById, ourTeam, regionFilter }) {
  const cesim = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]).sort((a, b) => a.roundNumber - b.roundNumber);
  if (!cesim.length) return null;
  const teams = dataById[cesim[cesim.length - 1].id].teams;
  const data = cesim.map((r) => {
    const rd = dataById[r.id], point = { name: `R${r.roundNumber}` };
    teams.forEach((tm) => { const i = rd.teams.indexOf(tm); point[tm] = i >= 0 ? getShareValue(rd, regionFilter, "Total", i) : null; });
    return point;
  });
  let paletteIdx = 0;
  return (
    <Panel style={{ padding: 18 }}>
      <Eyebrow info={`Cuota de mercado total (%) de cada equipo, ronda a ronda${regionFilter === "Global" ? "" : ` en ${regionFilter}`}.`}>Evolución de cuota {regionFilter === "Global" ? "global" : regionFilter} por ronda</Eyebrow>
      <div style={{ width: "100%", height: 280 }}>
        <ResponsiveContainer>
          <LineChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
            <CartesianGrid stroke={T.borderSoft} vertical={false} />
            <XAxis dataKey="name" tick={{ fontSize: 11, fill: T.textDim }} />
            <YAxis unit="%" tick={{ fontSize: 11, fill: T.textDim }} domain={["auto", "auto"]} />
            <Tooltip formatter={(v, n) => [Number.isFinite(v) ? `${v.toFixed(2)}%` : "—", n]} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {teams.map((tm) => {
              const us = tm === ourTeam;
              return <Line key={tm} dataKey={tm} stroke={us ? T.amber : TEAM_PALETTE[paletteIdx++ % TEAM_PALETTE.length]} strokeWidth={us ? 3.5 : 1.8} strokeOpacity={us ? 1 : 0.8} dot={{ r: us ? 4 : 3 }} connectNulls />;
            })}
          </LineChart>
        </ResponsiveContainer>
      </div>
      {cesim.length < 2 && <div style={{ fontSize: 12, color: T.textFaint, marginTop: 6 }}>Hay una sola ronda cargada: cargá las anteriores para ver la evolución.</div>}
    </Panel>
  );
}

/* ---------------------------------------------------------------
   PRODUCCIÓN Y OPERACIONES
   Las secciones del .xls (capacidad, inventario, costos, defectuosos…) siguen todas el patrón
   "sección > región > Tec N", con grupos "Tec N" vacíos intercalados y títulos de bloque poco fiables.
   readSection busca la sección por el texto de su grupo y devuelve {region: {tecs, extras}}.
--------------------------------------------------------------- */
function readSection(roundData, startRe) {
  for (const block of roundData?.blocks || []) {
    const rows = block.rows;
    const start = rows.findIndex((r) => r.kind === "group" && startRe.test(r.label));
    if (start < 0) continue;
    const out = {};
    let region = null;
    for (let i = start + 1; i < rows.length; i++) {
      const r = rows[i];
      if (r.kind === "group") {
        if (MARKET_REGIONS.some((m) => m.key === r.label)) { region = r.label; out[region] = { tecs: {}, extras: {} }; continue; }
        if (/^Tec \d+/.test(r.label)) continue;
        break; // empieza otra sección
      }
      if (!region) continue;
      const m = /^(Tec \d+)/.exec(r.label);
      if (m) out[region].tecs[m[1]] = r.values; else out[region].extras[r.label] = r.values;
    }
    return out;
  }
  return null;
}
function readFactories(roundData) {
  for (const block of roundData?.blocks || []) {
    const start = block.rows.findIndex((r) => r.kind === "group" && /^N.mero de f.bricas/i.test(r.label));
    if (start < 0) continue;
    const out = [];
    let period = null;
    for (let i = start + 1; i < block.rows.length; i++) {
      const r = block.rows[i];
      if (r.kind === "group") { if (/^(Ronda actual|Pr.xima ronda|Despu.s de la pr.xima)/i.test(r.label)) { period = r.label; continue; } break; }
      if (period) out.push({ period, region: r.label, values: r.values });
    }
    return out;
  }
  return [];
}
function fmtDec(n, d = 1) { return Number.isFinite(n) ? n.toLocaleString("es-AR", { maximumFractionDigits: d }) : "—"; }
const numOrNull = (v) => (v === null || v === undefined || v === "" ? null : toNumberOrNull(v));

// Filas de matriz (una por región/tec) para una sección; los renglones "extras" (Capacidad Disponible,
// Costos de gestión, Promedio ponderado…) van en negrita al final de cada región.
function sectionMatrixRows(roundData, startRe, regions, { prefix = "", vs = null, extras = true } = {}) {
  const sec = readSection(roundData, startRe);
  if (!sec) return [];
  const rows = [];
  regions.forEach((region) => {
    const s = sec[region];
    if (!s) return;
    const body = [];
    Object.keys(s.tecs).sort((a, b) => parseInt(a.slice(4), 10) - parseInt(b.slice(4), 10)).forEach((tec) => {
      const values = s.tecs[tec].map(numOrNull);
      if (values.some((v) => v)) body.push({ label: tec, values, vs });
    });
    if (extras) Object.entries(s.extras).forEach(([label, vals]) => {
      const values = vals.map(numOrNull);
      if (values.some((v) => v !== null)) body.push({ label, values, bold: true, vs: /promedio/i.test(label) ? vs : null });
    });
    if (body.length) rows.push({ header: `${prefix}${region}` }, ...body);
  });
  return rows;
}

function MatrixCard({ title, info, teams, ourTeam, rows, fmt = (v) => fmtDec(v, 1), emptyText = "No se encontraron datos en la última ronda." }) {
  const ourIdx = teams.indexOf(ourTeam);
  const hasVs = rows.some((r) => r.vs);
  const Dot = ({ s }) => <span style={{ display: "inline-block", width: 9, height: 9, borderRadius: "50%", background: TL_COLORS[s], marginRight: 5, verticalAlign: "middle" }} />;
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info={info}>{title}</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={MONO_TABLE}>
          <thead><tr><th style={{ ...TH, textAlign: "left" }}></th><TeamHeaderCells teams={teams} ourTeam={ourTeam} />{hasVs && <th style={TH}>{ourTeam || "Nosotros"} vs rivales</th>}</tr></thead>
          <tbody>
            {rows.map((r, ri) => {
              if (r.header) return <tr key={`h${ri}`}><td colSpan={teams.length + (hasVs ? 2 : 1)} style={{ padding: "8px 12px", background: T.panelAlt, color: T.textDim, fontWeight: 700, borderTop: `1px solid ${T.border}` }}>{r.header}</td></tr>;
              let diff = null;
              if (r.vs && ourIdx >= 0 && r.values[ourIdx] > 0) {
                const others = avg(r.values.filter((v, i) => i !== ourIdx && v > 0));
                if (others) diff = ((r.values[ourIdx] - others) / others) * 100;
              }
              const good = r.vs === "low" ? diff <= 0 : diff >= 0;
              return (
                <tr key={ri}>
                  <td style={{ ...TD, textAlign: "left", fontWeight: r.bold ? 700 : 400 }}>{r.label}</td>
                  {r.values.map((v, i) => {
                    const isUs = teams[i] === ourTeam;
                    return <td key={teams[i]} style={{ ...TD, background: isUs ? T.amberDim : "transparent", color: isUs ? T.amber : T.text, fontWeight: isUs || r.bold ? 700 : 400 }}>{v === null ? "—" : <>{r.status && <Dot s={r.status(v)} />}{(r.fmt || fmt)(v)}</>}</td>;
                  })}
                  {hasVs && <td style={{ ...TD, fontWeight: 700, color: diff === null ? T.textFaint : good ? T.green : T.red }}>{diff === null ? "—" : `${diff >= 0 ? "▲" : "▼"} ${fmtSigned(diff, 1, "%")}`}</td>}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {rows.length === 0 && <div style={{ padding: 16, fontSize: 13, color: T.textFaint }}>{emptyText}</div>}
    </Panel>
  );
}

function ProduccionCards({ latestData, ourTeam, regionFilter }) {
  if (!latestData) return null;
  const teams = latestData.teams;
  const scope = (list) => list.filter((r) => regionFilter === "Global" || r === regionFilter);
  const allRegions = MARKET_REGIONS.map((m) => m.key);
  const factoryRegions = scope(["EE.UU.", "Asia"]);
  const pct = (v) => `${fmtDec(v, 1)}%`;

  // Capacidad: por tec = % de la capacidad total que ocupa; total empleada = 100 - capacidad disponible.
  const capRows = sectionMatrixRows(latestData, /^Capacidad empleada/i, factoryRegions, { prefix: "Fábricas ", extras: true });
  const capFinal = [];
  capRows.forEach((r) => {
    if (r.label && /^Capacidad Disponible/i.test(r.label)) {
      capFinal.push({ label: "Capacidad ociosa", values: r.values, bold: true });
      capFinal.push({ label: "Capacidad empleada total", values: r.values.map((v) => (v === null ? null : 100 - v)), bold: true, status: (v) => (v > 70 ? "green" : v >= 40 ? "yellow" : "red") });
    } else capFinal.push(r);
  });

  const factories = readFactories(latestData).filter((f) => factoryRegions.includes(f.region));
  const factoryRows = [];
  factoryRegions.forEach((region) => {
    const list = factories.filter((f) => f.region === region);
    if (list.length) factoryRows.push({ header: region }, ...list.map((f) => ({ label: f.period, values: f.values.map(numOrNull) })));
  });

  const prodRows = [
    ...sectionMatrixRows(latestData, /^Producci.n interna, miles/i, scope(["EE.UU.", "Asia"]), { prefix: "Interna · ", extras: false }),
    ...sectionMatrixRows(latestData, /^Producci.n contratada, miles/i, scope(["EE.UU.", "Asia"]), { prefix: "Contratada · ", extras: false }),
  ];
  const costRows = [
    ...sectionMatrixRows(latestData, /^Costo de producci.n interna por unidad/i, scope(["EE.UU.", "Asia"]), { prefix: "Producción interna · ", vs: "low", extras: false }),
    ...sectionMatrixRows(latestData, /^Costos? de fabricaci.n contratada por unidad/i, scope(["EE.UU.", "Asia"]), { prefix: "Fabricación contratada · ", vs: "low", extras: false }),
    ...sectionMatrixRows(latestData, /^Costo unitario promedio/i, scope(allRegions), { prefix: "Costo unitario promedio · ", vs: "low" }),
  ];
  const invRows = sectionMatrixRows(latestData, /^Inventario, miles/i, scope(["EE.UU.", "Asia"]), { prefix: "Inventario · " });
  const defRows = sectionMatrixRows(latestData, /^Producto Defectuoso/i, scope(["EE.UU.", "Asia"]), { prefix: "Defectuosos · ", vs: "low" });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      {regionFilter === "Europa" && <Panel style={{ padding: 14, fontSize: 13, color: T.textDim }}>Europa no tiene fábricas propias: se abastece desde EE.UU. y Asia. Solo se muestran los costos unitarios de ese mercado.</Panel>}
      <MatrixCard title="Capacidad empleada" teams={teams} ourTeam={ourTeam} rows={capFinal} fmt={pct}
        info="Cada tecnología muestra qué % de la capacidad total de la fábrica ocupa; 'empleada total' es 100% menos la capacidad ociosa. Semáforo inverso: verde >70%, amarillo 40–70%, rojo <40%." />
      <MatrixCard title="Fábricas" teams={teams} ourTeam={ourTeam} rows={factoryRows} fmt={(v) => fmtDec(v, 0)}
        info="Cantidad de fábricas en la ronda actual, la próxima y la siguiente (según las decisiones ya tomadas)." />
      <MatrixCard title="Producción interna vs contratada (miles de unidades)" teams={teams} ourTeam={ourTeam} rows={prodRows} fmt={(v) => fmtDec(v, 1)}
        info="Unidades producidas en fábricas propias y por fabricantes contratados, por región y tecnología." />
      <MatrixCard title="Costos unitarios (USD por unidad)" teams={teams} ourTeam={ourTeam} rows={costRows} fmt={(v) => fmtDec(v, 1)}
        info="Costo por unidad de producción interna, de fabricación contratada y costo unitario promedio por producto vendido. La última columna compara a ULTI con el promedio de los rivales: verde = más barato." />
      <MatrixCard title="Inventario (miles de unidades)" teams={teams} ourTeam={ourTeam} rows={invRows} fmt={(v) => fmtDec(v, 1)}
        info="Unidades en stock al cierre por región y tecnología, con el costo de gestión del inventario en miles de USD." />
      <MatrixCard title="Producto defectuoso (%)" teams={teams} ourTeam={ourTeam} rows={defRows} fmt={pct}
        info="Porcentaje de producto defectuoso por región y tecnología. Verde en la última columna = menos defectos que el promedio de rivales." />
    </div>
  );
}

/* ---------------------------------------------------------------
   FINANZAS
--------------------------------------------------------------- */
const BALANCE_GLOBAL = "Hoja de Balance, miles USD, Global";
function balanceValues(roundData, label) {
  const block = roundData?.blocks?.find((b) => b.title === BALANCE_GLOBAL);
  const row = block?.rows.find((r) => r.kind === "metric" && r.label === label);
  return row ? row.values.map(numOrNull) : null;
}
// Métricas de un grupo (hasta el próximo grupo), buscado por el texto del grupo.
function flatGroupRows(roundData, groupRe) {
  for (const block of roundData?.blocks || []) {
    const start = block.rows.findIndex((r) => r.kind === "group" && groupRe.test(r.label));
    if (start < 0) continue;
    const out = [];
    for (let i = start + 1; i < block.rows.length && block.rows[i].kind === "metric"; i++) out.push({ label: block.rows[i].label, values: block.rows[i].values.map(numOrNull) });
    return out;
  }
  return [];
}
const RATING_TIER = (rating) => {
  const i = CREDIT_RATING_ORDER.indexOf(String(rating || "").trim());
  return i < 0 ? "none" : i <= CREDIT_RATING_ORDER.indexOf("A+") ? "green" : i <= CREDIT_RATING_ORDER.indexOf("BBB-") ? "yellow" : "red";
};

function CapitalStructureCard({ latestData, ourTeam }) {
  const lp = balanceValues(latestData, "Deudas a largo plazo"), cp = balanceValues(latestData, "Deudas a corto plazo (no planificadas)");
  const eq = balanceValues(latestData, "Total patrimonio neto"), cash = balanceValues(latestData, "Efectivo y equivalentes de efectivo");
  if (!lp && !cp && !eq && !cash) return null;
  const teams = latestData.teams, at = (arr, i) => (arr ? arr[i] : null);
  const data = teams.map((tm, i) => ({ name: tm === ourTeam ? `${tm} ★` : tm, "Patrimonio neto": at(eq, i) || 0, "Deuda largo plazo": at(lp, i) || 0, "Deuda corto plazo": at(cp, i) || 0 }));
  const netDebt = teams.map((_, i) => (at(lp, i) ?? 0) + (at(cp, i) ?? 0) - (at(cash, i) ?? 0));
  const rows = [
    { label: "Patrimonio neto", values: eq || teams.map(() => null) }, { label: "Deuda largo plazo", values: lp || teams.map(() => null) },
    { label: "Deuda corto plazo", values: cp || teams.map(() => null) }, { label: "Efectivo", values: cash || teams.map(() => null) },
    { label: "Deuda neta (deuda - efectivo)", values: netDebt, bold: true },
  ];
  return (
    <>
      <Panel style={{ padding: 18 }}>
        <Eyebrow info="Cómo se financia cada empresa: patrimonio neto y deuda (largo y corto plazo), en miles de USD. Balance global de la compañía (no cambia con el selector de región).">Estructura de capital</Eyebrow>
        <div style={{ width: "100%", height: 280 }}>
          <ResponsiveContainer>
            <BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
              <CartesianGrid stroke={T.borderSoft} vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11, fill: T.textDim }} />
              <YAxis tick={{ fontSize: 11, fill: T.textDim }} tickFormatter={(v) => fmtNumber(v)} />
              <Tooltip formatter={tooltipNumberFormatter} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Patrimonio neto" stackId="cap" fill={T.cyan} />
              <Bar dataKey="Deuda largo plazo" stackId="cap" fill={T.amber} />
              <Bar dataKey="Deuda corto plazo" stackId="cap" fill={T.red} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Panel>
      <MatrixCard title="Estructura de capital — detalle (miles USD)" teams={teams} ourTeam={ourTeam} rows={rows} fmt={fmtInt} info="Deuda neta negativa = la empresa tiene más efectivo que deuda." />
    </>
  );
}

function CreditRatingEvolutionCard({ rounds, dataById, ourTeam }) {
  const cesim = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]).sort((a, b) => a.roundNumber - b.roundNumber);
  if (!cesim.length) return null;
  const teams = dataById[cesim[cesim.length - 1].id].teams;
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Calificación crediticia de cada equipo, ronda a ronda. Verde A+ o mejor, amarillo BBB a A, rojo peor que BBB.">Calificación crediticia por ronda</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={MONO_TABLE}>
          <thead><tr><th style={{ ...TH, textAlign: "left" }}>Equipo</th>{cesim.map((r) => <th key={r.id} style={TH}>R{r.roundNumber}</th>)}</tr></thead>
          <tbody>
            {teams.map((tm) => (
              <tr key={tm} style={{ background: tm === ourTeam ? T.amberDim : "transparent" }}>
                <td style={{ ...TD, textAlign: "left", color: tm === ourTeam ? T.amber : T.text, fontWeight: tm === ourTeam ? 700 : 400 }}>{tm}</td>
                {cesim.map((r) => {
                  const rd = dataById[r.id], i = rd.teams.indexOf(tm);
                  const rating = i >= 0 ? getRatio(rd, "Calificación crediticia", i) : null;
                  return <td key={r.id} style={{ ...TD, fontWeight: 700, color: TL_COLORS[RATING_TIER(rating)] }}>{typeof rating === "string" ? rating : "—"}</td>;
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}

function InterestRatesCard({ latestData, ourTeam }) {
  const rows = flatGroupRows(latestData, /^Tasas de inter.s/i);
  return <MatrixCard title="Tasas de interés por mercado y plazo (%)" teams={latestData.teams} ourTeam={ourTeam} rows={rows.map((r) => ({ ...r, vs: "low" }))} fmt={(v) => `${fmtDec(v, 2)}%`}
    info="Tasa de la deuda específica de cada empresa según su calificación. La última columna compara a ULTI con el promedio de rivales: verde = paga menos." emptyText="No se encontraron tasas de interés en la última ronda." />;
}

function FinancialExpensesCard({ latestData, ourTeam, regionFilter }) {
  const all = flatGroupRows(latestData, /^Acreedores/i);
  const rows = all.filter((r) => regionFilter === "Global" ? true : new RegExp(regionFilter.replace(".", "\\."), "i").test(r.label) && !/^Total/i.test(r.label))
    .map((r) => ({ ...r, bold: /^Total/i.test(r.label) }));
  return <MatrixCard title="Gastos financieros netos por región (miles USD)" teams={latestData.teams} ourTeam={ourTeam} rows={rows} fmt={fmtInt}
    info="Intereses pagados menos intereses ganados. Un valor negativo significa que la región generó más ingresos financieros que gastos." emptyText="No se encontró el desglose de gastos financieros en la última ronda." />;
}

function FinanzasSaludCards({ rounds, dataById, ourTeam, regionFilter }) {
  const { latestData } = getRoundContext(rounds, dataById, ourTeam);
  if (!latestData) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <CapitalStructureCard latestData={latestData} ourTeam={ourTeam} />
      <CreditRatingEvolutionCard rounds={rounds} dataById={dataById} ourTeam={ourTeam} />
      <InterestRatesCard latestData={latestData} ourTeam={ourTeam} />
      <FinancialExpensesCard latestData={latestData} ourTeam={ourTeam} regionFilter={regionFilter} />
    </div>
  );
}

/* ---------------------------------------------------------------
   RATIOS Y VALORACIÓN
--------------------------------------------------------------- */
const RATIO_PANEL = [
  { key: "roe", label: "ROE", re: /\(ROE\)/i, unit: "%" },
  { key: "roce", label: "ROCE", re: /\(ROCE\)/i, unit: "%" },
  { key: "ros", label: "ROS (rentabilidad de ventas)", re: /\(ROS\)/i, unit: "%" },
  { key: "gm", label: "Margen bruto", re: /^Margen bruto/i, unit: "%" },
  { key: "ebitda", label: "Margen EBITDA", re: /EBITDA/i, unit: "%" },
  { key: "eps", label: "EPS (USD)", re: /\(EPS\)/i, unit: "", dec: 2 },
  { key: "pe", label: "P/E", re: /P\/E/i, unit: "x" },
  { key: "lev", label: "Apalancamiento (deuda neta / patrimonio)", re: /apalancamiento/i, unit: "%", lowerBetter: true },
  { key: "eq", label: "Ratio patrimonio neto", re: /^Ratio Patrimonio neto/i, unit: "%" },
];

function RatiosClaveCard({ rounds, dataById, ourTeam }) {
  const cesim = rounds.filter((r) => r.kind === "cesim" && dataById[r.id]).sort((a, b) => a.roundNumber - b.roundNumber);
  if (!cesim.length) return null;
  const latest = dataById[cesim[cesim.length - 1].id], ourIdx = latest.teams.indexOf(ourTeam);
  if (ourIdx < 0) return null;
  const fmtV = (v, def) => (v === null ? "—" : `${fmtDec(v, def.dec ?? 1)}${def.unit}`);
  return (
    <Panel style={{ padding: 0, overflow: "hidden" }}>
      <div style={{ padding: "14px 16px 8px" }}><Eyebrow info="Ratios clave de ULTI ronda a ronda, con el ranking entre los equipos en la última ronda (1° = mejor; en apalancamiento, menor es mejor) y quién lidera.">Panel de ratios clave</Eyebrow></div>
      <div style={{ overflowX: "auto" }}>
        <table style={MONO_TABLE}>
          <thead><tr>
            <th style={{ ...TH, textAlign: "left" }}>Ratio</th>{cesim.map((r) => <th key={r.id} style={TH}>R{r.roundNumber}</th>)}
            <th style={TH}>Var.</th><th style={TH}>Ranking</th><th style={{ ...TH, textAlign: "left" }}>Líder</th>
          </tr></thead>
          <tbody>
            {RATIO_PANEL.map((def, ri) => {
              const series = cesim.map((r) => { const rd = dataById[r.id], i = rd.teams.indexOf(ourTeam); return i >= 0 ? getRatioBy(rd, def.re, i) : null; });
              const cur = series[series.length - 1], prev = series.length > 1 ? series[series.length - 2] : null;
              const sign = def.lowerBetter ? -1 : 1;
              const rank = rankAmongTeams(latest, (rd, i) => { const v = getRatioBy(rd, def.re, i); return v === null ? null : sign * v; }, ourIdx);
              const all = latest.teams.map((tm, i) => ({ tm, v: getRatioBy(latest, def.re, i) })).filter((x) => x.v !== null);
              const leader = all.length ? all.reduce((a, b) => (sign * b.v > sign * a.v ? b : a)) : null;
              const delta = cur !== null && prev !== null ? cur - prev : null;
              const goodDelta = delta === null ? null : def.lowerBetter ? delta <= 0 : delta >= 0;
              return (
                <tr key={def.key} style={{ background: ri % 2 ? T.panel : T.panelAlt }}>
                  <td style={{ ...TD, textAlign: "left" }}>{def.label}</td>
                  {series.map((v, i) => <td key={cesim[i].id} style={{ ...TD, fontWeight: i === series.length - 1 ? 700 : 400, color: i === series.length - 1 ? T.amber : T.text }}>{fmtV(v, def)}</td>)}
                  <td style={{ ...TD, fontWeight: 700, color: delta === null ? T.textFaint : goodDelta ? T.green : T.red }}>{delta === null ? "—" : fmtSigned(delta, def.dec ?? 1)}</td>
                  <td style={{ ...TD, fontWeight: 700, color: !rank ? T.textFaint : rank.rank <= 2 ? T.green : rank.rank >= rank.of - 1 ? T.red : T.text }}>{rank ? `${rank.rank}°/${rank.of}` : "—"}</td>
                  <td style={{ ...TD, textAlign: "left", color: T.textDim }}>{leader ? `${leader.tm} (${fmtV(leader.v, def)})` : "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {cesim.length < 2 && <div style={{ padding: "8px 16px 12px", fontSize: 12, color: T.textFaint }}>Hay una sola ronda cargada: la variación aparece al cargar la anterior.</div>}
    </Panel>
  );
}

function valuationRows(roundData, region) {
  const block = roundData.blocks.find((b) => b.title === `Valuación - ${region === "Global" ? "Global" : region}`);
  if (!block) return [];
  const rows = block.rows.filter((r) => r.kind === "metric");
  const nTeams = roundData.teams.length;
  const find = (re, from = 0) => { for (let i = from; i < rows.length; i++) if (re.test(rows[i].label)) return i; return -1; };
  const vals = (i) => (i < 0 ? null : rows[i].values.map(numOrNull));
  const out = [];
  const push = (label, values, extra = {}) => { if (values && values.some((v) => v !== null)) out.push({ label, values, ...extra }); };

  const evMarketIdx = find(/^Valor de la empresa/);
  const evMarket = vals(evMarketIdx);
  out.push({ header: "Capitalización y valor de la empresa" });
  push("Capitalización de mercado", vals(find(/^Capitalización de mercado, miles/)), { bold: true });
  if (region === "Global") push("Precio de la acción, USD", roundData.teams.map((_, i) => getRatioBy(roundData, /^Precio de la acci.n al final/i, i)), { fmt: (v) => fmtDec(v, 1) });
  push("Valor de la empresa (mercado)", evMarket);
  push("Exceso de efectivo (+)", vals(find(/^Exceso de efectivo/)));
  push("Valor de la deuda (-)", vals(find(/^Valor de la deuda/)));
  push("WACC", vals(find(/^WACC/)), { fmt: (v) => `${fmtDec(v, 2)}%` });

  out.push({ header: "Valor de la empresa según cada método" });
  const dcfIdx = find(/^Valor de la empresa/, find(/^WACC/) + 1);
  push("Flujos de fondos descontados (DCF)", vals(dcfIdx));
  const methods = [
    ["EV/Ventas", /^EV\/Ventas/i, /^Ventas ajustadas/i], ["EV/EBITDA", /^EV\/EBITDA/i, /^EBITDA ajustado/i],
    ["EV/EBIT", /^EV\/EBIT(?!DA)/i, /^EBIT ajustado/i], ["EV/NOPLAT", /^EV\/NOPLAT/i, /^NOPLAT ajustado/i],
  ];
  const implied = [];
  methods.forEach(([name, anchorRe, adjRe]) => {
    const a = find(anchorRe);
    if (a < 0) return;
    const mult = numOrNull(rows[a].values[0]);
    const adj = vals(find(adjRe, a));
    push(`${name} (${fmtDec(mult, 2)}x)`, vals(find(/^Valor de la empresa/, a)));
    if (adj && evMarket) implied.push({ label: name, values: evMarket.map((ev, i) => (ev !== null && adj[i] ? ev / adj[i] : null)) });
  });
  if (implied.length) {
    out.push({ header: "Múltiplos implícitos en el valor de mercado (valor de la empresa / métrica ajustada)" });
    implied.forEach((r) => push(`${r.label} implícito`, r.values, { fmt: (v) => `${fmtDec(v, 2)}x` }));
  }
  return nTeams ? out : [];
}
function ValoracionCard({ latestData, ourTeam, regionFilter }) {
  const rows = valuationRows(latestData, regionFilter);
  return <MatrixCard title={`Valoración de la empresa${regionFilter === "Global" ? "" : ` — ${regionFilter}`}`} teams={latestData.teams} ourTeam={ourTeam} rows={rows} fmt={fmtInt}
    info="Valor de la empresa (EV) por capitalización de mercado, flujos descontados y múltiplos del simulador. Los múltiplos implícitos muestran cuántas veces las ventas/EBITDA/EBIT/NOPLAT ajustados vale el mercado a cada equipo, para comparar con el múltiplo de referencia." emptyText="No se encontró el bloque de valuación en la última ronda." />;
}

function readValueCreation(roundData) {
  const block = roundData.blocks.find((b) => /^Creaci.n de valor/i.test(b.title));
  if (!block) return null;
  const parts = []; let cur = null, grand = null;
  block.rows.forEach((r) => {
    if (r.kind === "group") { cur = { label: r.label, items: [], total: null }; parts.push(cur); return; }
    const values = r.values.map(numOrNull);
    if (/^Valor total creado/i.test(r.label)) { grand = values; return; }
    if (!cur) return;
    if (r.label === "Total") cur.total = values; else cur.items.push({ label: r.label, values });
  });
  return { parts: parts.filter((p) => p.total), grand };
}
function CreacionValorCard({ latestData, ourTeam }) {
  const vc = readValueCreation(latestData);
  if (!vc || !vc.parts.length) return null;
  const teams = latestData.teams;
  const colors = [T.amber, T.cyan, T.violet, T.green, "#D97706", "#64748B"];
  const data = teams.map((tm, i) => { const p = { name: tm === ourTeam ? `${tm} ★` : tm }; vc.parts.forEach((part) => { p[part.label] = part.total[i] || 0; }); return p; });
  const rows = [];
  vc.parts.forEach((part) => {
    rows.push({ header: part.label }, ...part.items.map((it) => ({ ...it })), { label: `Total ${part.label}`, values: part.total, bold: true });
  });
  if (vc.grand) rows.push({ header: "Valor total creado" }, { label: "Valor total creado", values: vc.grand, bold: true });
  return (
    <>
      <Panel style={{ padding: 18 }}>
        <Eyebrow info="Cómo se reparte el valor que genera cada empresa entre accionistas, acreedores, personal, gobierno y proveedores (miles USD). Un tramo negativo significa que ese grupo aportó valor a la empresa (por ejemplo, ingresos financieros netos).">Creación de valor por grupo de interés</Eyebrow>
        <div style={{ width: "100%", height: 300 }}>
          <ResponsiveContainer>
            <BarChart data={data} margin={{ top: 8, right: 16, left: 0, bottom: 0 }}>
              <CartesianGrid stroke={T.borderSoft} vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11, fill: T.textDim }} />
              <YAxis tick={{ fontSize: 11, fill: T.textDim }} tickFormatter={(v) => fmtNumber(v)} />
              <Tooltip formatter={tooltipNumberFormatter} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              {vc.parts.map((part, i) => <Bar key={part.label} dataKey={part.label} stackId="vc" fill={colors[i % colors.length]} />)}
            </BarChart>
          </ResponsiveContainer>
        </div>
      </Panel>
      <MatrixCard title="Creación de valor — detalle (miles USD)" teams={teams} ourTeam={ourTeam} rows={rows} fmt={fmtInt} />
    </>
  );
}

function RatiosValoracionCards({ rounds, dataById, ourTeam, regionFilter }) {
  const { latestData } = getRoundContext(rounds, dataById, ourTeam);
  if (!latestData) return null;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <RatiosClaveCard rounds={rounds} dataById={dataById} ourTeam={ourTeam} />
      <ValoracionCard latestData={latestData} ourTeam={ourTeam} regionFilter={regionFilter} />
      <CreacionValorCard latestData={latestData} ourTeam={ourTeam} />
    </div>
  );
}

function TabSection({ title, children }) {
  return (
    <section style={{ marginBottom: 34 }}>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 15, fontWeight: 600, color: T.text, margin: "0 0 12px", paddingBottom: 8, borderBottom: `1px solid ${T.borderSoft}` }}>{title}</div>
      {children}
    </section>
  );
}

function ResumenEjecutivoTab({ rounds, dataById, ourTeam, regionFilter, onNavigate }) {
  const { latestData, prevData, teamIdx } = getRoundContext(rounds, dataById, ourTeam);
  return (
    <>
      {latestData && teamIdx >= 0 && (
        <TabSection title="Estado de la ronda">
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            <TrafficLightsCard latestData={latestData} prevData={prevData} teamIdx={teamIdx} regionFilter={regionFilter} />
            <KpisPrincipalesCards latestData={latestData} prevData={prevData} teamIdx={teamIdx} />
            <DesvioPresupuestarioCard latestData={latestData} teamIdx={teamIdx} regionFilter={regionFilter} />
          </div>
        </TabSection>
      )}
      <TabSection title="Dashboard"><DashboardCompleto rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
      <TabSection title="Resumen para accionistas"><AccionistasSummarySection rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} onNavigate={onNavigate} /></TabSection>
    </>
  );
}
function MercadoTab({ rounds, dataById, ourTeam, regionFilter, categories }) {
  const { latestData, teamIdx } = getRoundContext(rounds, dataById, ourTeam);
  return (
    <>
      {latestData && (
        <TabSection title="Posición en el mercado">
          <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
            <CuotasHeatmapCard latestData={latestData} ourTeam={ourTeam} regionFilter={regionFilter} />
            <PreciosCompetenciaCard latestData={latestData} ourTeam={ourTeam} regionFilter={regionFilter} />
            <FillRateCard latestData={latestData} teamIdx={teamIdx} regionFilter={regionFilter} />
            <CuotaEvolucionChart rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} />
          </div>
        </TabSection>
      )}
      <TabSection title="Cuota de valor de mercado (VMS) y novedades del entorno"><ValueMarketShareSection rounds={rounds} /></TabSection>
      <TabSection title="Competencia"><CompetitionSection rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
      {categories.includes("Mercado") && <TabSection title="Informes de mercado (detalle del simulador)"><CategorySection category="Mercado" rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>}
    </>
  );
}
function ProduccionTab({ rounds, dataById, ourTeam, regionFilter, hasProductionData }) {
  const { latestData } = getRoundContext(rounds, dataById, ourTeam);
  return (
    <>
      {latestData && <TabSection title="Producción y operaciones"><ProduccionCards latestData={latestData} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>}
      {hasProductionData
        ? <TabSection title="Detalle del simulador (datos crudos)"><CategorySection category="Producción" rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
        : !latestData && <Panel style={{ padding: 20, color: T.textFaint, fontSize: 13 }}>Cargá una ronda para ver los datos de producción.</Panel>}
    </>
  );
}
function FinanzasTab({ rounds, dataById, ourTeam, regionFilter, categories }) {
  return (
    <>
      <TabSection title="Salud financiera"><FinanzasSaludCards rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
      <TabSection title="Análisis vertical y horizontal"><AnalisisFinancieroSection rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
      {categories.includes("Financiero") && <TabSection title="Estados financieros (detalle del simulador)"><CategorySection category="Financiero" rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>}
    </>
  );
}
function RatiosTab({ rounds, dataById, ourTeam, regionFilter }) {
  return (
    <>
      <TabSection title="Ratios clave, valoración y creación de valor"><RatiosValoracionCards rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
      <TabSection title="Buscador de ratios e indicadores"><DynamicDashboard rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} /></TabSection>
    </>
  );
}
function DecisionesTab({ rounds, dataById, ourTeam }) {
  return (
    <>
      <TabSection title="Estrategia"><StrategySection rounds={rounds} dataById={dataById} ourTeam={ourTeam} /></TabSection>
      <TabSection title="Premisas"><PremisesSection /></TabSection>
    </>
  );
}
function TabExportBar({ tabName }) {
  const [busy, setBusy] = useState(false);
  const label = TAB_LABELS[tabName] || tabName;

  const exportExcel = () => {
    setBusy(true);
    try {
      const container = document.getElementById("tab-print-area");
      const tables = container ? Array.from(container.querySelectorAll("table")) : [];
      if (!tables.length) { window.alert(`No hay tablas para exportar en "${label}" — probá Exportar PDF para capturar gráficos y texto.`); return; }
      const wb = XLSX.utils.book_new();
      tables.forEach((table, i) => {
        try {
          const ws = XLSX.utils.table_to_sheet(table);
          const sheetName = `${label} ${i + 1}`.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || `Tabla ${i + 1}`;
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        } catch (e) { /* skip a table that fails to convert */ }
      });
      XLSX.writeFile(wb, `cesim_${tabName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } finally { setBusy(false); }
  };

  return (
    <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 14 }}>
      <button onClick={exportExcel} disabled={busy} style={ghostBtn}>{busy ? <Loader2 size={13} className="spin" /> : <FileSpreadsheet size={13} />} Exportar Excel</button>
      <button onClick={() => window.print()} style={ghostBtn}><Printer size={13} /> Exportar PDF</button>
    </div>
  );
}

export default function App() {
  const [rounds, setRounds] = useState([]);
  const [dataById, setDataById] = useState({});
  const [loading, setLoading] = useState(true);
  const [loadErr, setLoadErr] = useState("");
  const [active, setActiveRaw] = useState("resumen");
  const setActive = useCallback((id) => setActiveRaw(LEGACY_TAB_ALIAS[id] || id), []);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [ourTeam, setOurTeam] = useState(null);
  const [regionFilter, setRegionFilter] = useState("Global");

  const refreshAll = useCallback(async (metaOverride) => {
    setLoading(true); setLoadErr("");
    try {
      const meta = metaOverride || (await loadRoundsMeta());
      setRounds(meta);
      const entries = await Promise.all(meta.map(async (r) => [r.id, await loadRoundData(r.id)]));
      const byId = {}; entries.forEach(([id, d]) => { if (d) byId[id] = d; });
      setDataById(byId);
      setOurTeam(await loadOurTeam());
    } catch (e) { setLoadErr("No se pudieron cargar los datos guardados. Probá recargar."); }
    setLoading(false);
  }, []);
  useEffect(() => { refreshAll(); }, [refreshAll]);

  const allTeams = useMemo(() => { const set = new Set(); rounds.forEach((r) => (r.teams || []).forEach((t) => set.add(t))); return Array.from(set); }, [rounds]);
  const availableCategories = useMemo(() => {
    const set = new Set();
    rounds.forEach((r) => { const d = dataById[r.id]; if (d?.kind === "cesim") d.blocks.forEach((b) => set.add(b.category)); });
    return CATEGORY_ORDER.filter((c) => set.has(c));
  }, [rounds, dataById]);
  const genericSheetNames = useMemo(() => { const set = new Set(); rounds.forEach((r) => { if (r.kind === "generic") r.sheetNames.forEach((sn) => set.add(sn)); }); return Array.from(set); }, [rounds]);

  const extraCategories = useMemo(() => availableCategories.filter((c) => !CATEGORIES_IN_MAIN_TABS.includes(c)), [availableCategories]);
  const tabs = useMemo(() => [...MAIN_TABS, "asistente", ...extraCategories, ...genericSheetNames], [extraCategories, genericSheetNames]);
  useEffect(() => { if (!tabs.includes(active)) setActive("resumen"); }, [tabs.join(","), active]);

  const doDelete = async (r) => {
    await deleteRoundData(r.id);
    const meta = await loadRoundsMeta();
    const next = meta.filter((m) => m.id !== r.id);
    await saveRoundsMeta(next);
    setConfirmDelete(null);
    refreshAll(next);
  };
  const handleTeamPick = async (tm) => { setOurTeam(tm); await saveOurTeam(tm); };

  return (
    <div style={{ minHeight: "100vh", background: T.bg, color: T.text, fontFamily: "'Inter', -apple-system, sans-serif", padding: "0 0 40px" }}>
      <style>{`
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        input:focus, select:focus { border-color: ${T.amber} !important; }
        .spin { animation: spin 0.9s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .chip-del { opacity: 0.5; transition: opacity 0.15s; }
        .round-chip:hover .chip-del { opacity: 1; }
        .nav-tab { transition: all 0.12s ease; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: ${T.border}; border-radius: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        @media print {
          body * { visibility: hidden; }
          #tab-print-area, #tab-print-area * { visibility: visible; }
          #tab-print-area { position: absolute; left: 0; top: 0; width: 100%; }
          .no-print { display: none !important; }
        }
      `}</style>

      <div style={{ maxWidth: 1040, margin: "0 auto", padding: "28px 20px 0" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 4 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 8, height: 8, borderRadius: "50%", background: T.green, boxShadow: `0 0 8px ${T.green}` }} />
            <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: T.textFaint, letterSpacing: "0.1em", textTransform: "uppercase" }}>Datos compartidos con todo el equipo</span>
          </div>
          {allTeams.length > 0 && <TeamPicker teams={allTeams} current={ourTeam} onPick={handleTeamPick} />}
        </div>

        <div style={{ marginBottom: 16 }}>
          <RegionFilter value={regionFilter} onChange={setRegionFilter} />
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 16, margin: "10px 0 22px" }}>
          <img src={ULTI_LOGO} alt="ULTI" style={{ height: 46, width: "auto" }} />
          <div style={{ width: 1, height: 34, background: T.border }} />
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 21, fontWeight: 700, color: T.text, letterSpacing: "-0.01em", lineHeight: 1.15 }}>Tablero de Control de Gestión</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: T.textFaint, letterSpacing: "0.08em", textTransform: "uppercase", marginTop: 2 }}>Cesim Global Challenge</div>
          </div>
        </div>

        {rounds.length > 0 && <div style={{ marginBottom: 22 }}><RoundTicker rounds={rounds} onDelete={setConfirmDelete} /></div>}
        <div style={{ marginBottom: 26 }}><UploadPanel rounds={rounds} onSaved={(meta) => refreshAll(meta)} /></div>

        {loading ? (
          <div style={{ display: "flex", alignItems: "center", gap: 10, color: T.textFaint, padding: 30 }}><Loader2 size={16} className="spin" /> Cargando datos del equipo…</div>
        ) : loadErr ? (
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: T.red, padding: 20 }}><AlertTriangle size={16} /> {loadErr}<button onClick={() => refreshAll()} style={{ ...ghostBtn, padding: "4px 10px" }}><RefreshCw size={12} /> Reintentar</button></div>
        ) : (
          <>
            <div style={{ display: "flex", gap: 6, borderBottom: `1px solid ${T.borderSoft}`, marginBottom: 20, overflowX: "auto" }}>
              <TabButton icon={<LayoutDashboard size={13} />} active={active === "resumen"} onClick={() => setActive("resumen")}>Resumen Ejecutivo</TabButton>
              <TabButton icon={<Target size={13} />} active={active === "mercado"} onClick={() => setActive("mercado")}>Mercado y Demanda</TabButton>
              <TabButton icon={<Factory size={13} />} active={active === "produccion"} onClick={() => setActive("produccion")}>Producción y Operaciones</TabButton>
              <TabButton icon={<Wallet size={13} />} active={active === "finanzas"} onClick={() => setActive("finanzas")}>Finanzas</TabButton>
              <TabButton icon={<TrendingUp size={13} />} active={active === "ratios"} onClick={() => setActive("ratios")}>Ratios y Valoración</TabButton>
              <TabButton icon={<ClipboardList size={13} />} active={active === "decisiones"} onClick={() => setActive("decisiones")}>Decisiones y Registro</TabButton>
              <TabButton icon={<MessageCircleQuestion size={13} />} active={active === "asistente"} onClick={() => setActive("asistente")}>Asistente IA</TabButton>
              {extraCategories.map((c) => <TabButton key={c} icon={<TableIcon size={13} />} active={active === c} onClick={() => setActive(c)}>{c}</TabButton>)}
              {genericSheetNames.map((sn) => <TabButton key={sn} icon={<TableIcon size={13} />} active={active === sn} onClick={() => setActive(sn)}>{sn}</TabButton>)}
            </div>

            <TabExportBar tabName={active} />
            <div id="tab-print-area">
              {active === "resumen" ? <ResumenEjecutivoTab rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} onNavigate={setActive} />
                : active === "mercado" ? <MercadoTab rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} categories={availableCategories} />
                : active === "produccion" ? <ProduccionTab rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} hasProductionData={availableCategories.includes("Producción")} />
                : active === "finanzas" ? <FinanzasTab rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} categories={availableCategories} />
                : active === "ratios" ? <RatiosTab rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} />
                : active === "decisiones" ? <DecisionesTab rounds={rounds} dataById={dataById} ourTeam={ourTeam} />
                : active === "asistente" ? <AiAssistant rounds={rounds} dataById={dataById} ourTeam={ourTeam} />
                : CATEGORY_ORDER.includes(active) ? <CategorySection category={active} rounds={rounds} dataById={dataById} ourTeam={ourTeam} regionFilter={regionFilter} />
                : <GenericSheetSection sheetName={active} rounds={rounds} dataById={dataById} />}
            </div>
          </>
        )}
      </div>

      {confirmDelete && (
        <div style={{ position: "fixed", inset: 0, background: T.overlay, display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50, padding: 20 }}>
          <Panel style={{ padding: 22, maxWidth: 360 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><AlertTriangle size={16} color={T.red} /><span style={{ fontWeight: 600, fontSize: 14 }}>Eliminar ronda</span></div>
            <div style={{ fontSize: 13, color: T.textDim, marginBottom: 16 }}>Se va a borrar R{confirmDelete.roundNumber} ({confirmDelete.label}) para todo el equipo. Esta acción no se puede deshacer.</div>
            <div style={{ display: "flex", gap: 10 }}>
              <button onClick={() => doDelete(confirmDelete)} style={{ ...primaryBtn, background: T.red, color: "#fff" }}><Trash2 size={13} /> Eliminar</button>
              <button onClick={() => setConfirmDelete(null)} style={ghostBtn}>Cancelar</button>
            </div>
          </Panel>
        </div>
      )}
    </div>
  );
}
