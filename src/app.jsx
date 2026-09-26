// Diario dell'assistenza
// Copyright (C) 2026 na103
//
// Questo programma è software libero: puoi ridistribuirlo e/o modificarlo secondo i termini
// della GNU General Public License pubblicata dalla Free Software Foundation, versione 3
// o (a tua scelta) qualunque versione successiva. È distribuito senza alcuna garanzia.
// Il testo completo è nel file LICENSE.

import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Moon, Sun, Activity, List, BarChart3, Settings, Plus, Minus,
  Trash2, Download, Upload, ChevronLeft, ChevronRight, X, Pill, Pencil, Droplet, Coffee,
} from "lucide-react";
import { version as APP_VERSION } from "../package.json";

/* ------------------------------------------------------------------ */
/*  Costanti                                                           */
/* ------------------------------------------------------------------ */

const KEY = "diario-assistenza:v1";

const AGIT = [
  { v: 0, label: "Tranquilla", desc: "Nessun disturbo" },
  { v: 1, label: "Confusa ma gestibile", desc: "Qualche richiamo, si calma con poco" },
  { v: 2, label: "Agitata, delirio", desc: "Difficile da calmare" },
  { v: 3, label: "Crisi", desc: "Urla, scende dal letto, allucinazioni" },
];

// Stessi valori della lista a tendina del foglio Excel
const EP_TYPES = [
  "Delirio", "Agitazione", "Allucinazioni", "Rifiuto cure", "Vagabondaggio",
  "Caduta", "Aggressivita", "Apatia marcata", "Altro",
];
const EP_LABEL = { Aggressivita: "Aggressività" };

const INTENS = [
  { v: 1, label: "Lieve" },
  { v: 2, label: "Moderata" },
  { v: 3, label: "Forte" },
];

const ASSIST_CHIPS = [0, 5, 10, 15, 20, 30, 45, 60];

const GLU_TAGS = ["A digiuno", "Prima del pasto", "Dopo il pasto", "Sera", "Altro"];
// Colori solo indicativi, per leggere il grafico a colpo d'occhio: i valori di riferimento li dà il medico.
const gluColor = (v) => (v == null ? "var(--muted)" : v < 70 ? "var(--a3)" : v <= 180 ? "var(--a0)" : v <= 250 ? "var(--a2)" : "var(--a3)");

// Pressione misurata a casa (ESH 2023): 135/85 è il limite della norma, da 160/100 è decisamente alta.
// Verso il basso conta solo la massima: sotto 100 è bassa, sotto 90 è ipotensione.
const bpLevel = (s, d) => {
  if (s == null || d == null) return null;
  if (s >= 160 || d >= 100 || s < 90) return 3;
  if (s >= 135 || d >= 85 || s < 100) return 1;
  return 0;
};
const bpColor = (s, d) => { const l = bpLevel(s, d); return l == null ? "var(--muted)" : `var(--a${l})`; };
const BP_LEGEND = [{ l: 0, label: "nella norma" }, { l: 1, label: "da tenere d'occhio" }, { l: 3, label: "alta o bassa" }];

const GIORNI = ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"];
const GIORNI_XLS = ["Domenica", "Lunedi", "Martedi", "Mercoledi", "Giovedi", "Venerdi", "Sabato"];
const MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
  "agosto", "settembre", "ottobre", "novembre", "dicembre"];

/* ------------------------------------------------------------------ */
/*  Utilità                                                            */
/* ------------------------------------------------------------------ */

const pad = (n) => String(n).padStart(2, "0");
const toISO = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromISO = (s) => { const [y, m, d] = s.split("-").map(Number); return new Date(y, m - 1, d); };
const addDays = (s, n) => { const d = fromISO(s); d.setDate(d.getDate() + n); return toISO(d); };
const fmtDay = (s) => { const d = fromISO(s); return `${GIORNI[d.getDay()]} ${d.getDate()}`; };
const fmtDayShort = (s) => { const d = fromISO(s); return `${GIORNI[d.getDay()].slice(0, 3)} ${d.getDate()}`; };
const monthLabel = (ym) => { const [y, m] = ym.split("-").map(Number); const s = `${MESI[m - 1]} ${y}`; return s[0].toUpperCase() + s.slice(1); };
const fmtNum = (x, dec = 1) =>
  x == null || isNaN(x) ? "–" : x.toLocaleString("it-IT", { maximumFractionDigits: dec });
const toMin = (t) => { if (!t) return null; const [h, m] = t.split(":").map(Number); return isNaN(h) ? null : h * 60 + m; };
const duration = (a, b) => { const x = toMin(a), y = toMin(b); if (x == null || y == null) return null; return ((y - x) % 1440 + 1440) % 1440; };
const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36);
const nowHM = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const plural = (n, s, p) => `${n} ${n === 1 ? s : p}`;
const parseDec = (s) => { if (s == null || String(s).trim() === "") return null; const x = parseFloat(String(s).replace(",", ".")); return isNaN(x) ? null : x; };
const splitHM = (x) => { if (x == null || isNaN(x)) return { h: "", m: "" }; const t = Math.round(x * 60); return { h: String(Math.floor(t / 60)), m: String(t % 60) }; };
const fromHM = (h, m) => { if (String(h).trim() === "" && String(m).trim() === "") return null; const t = (parseInt(h, 10) || 0) * 60 + (parseInt(m, 10) || 0); return Math.round((t / 60) * 10000) / 10000; };
const fmtHM = (x) => { if (x == null || isNaN(x)) return "–"; const { h, m } = splitHM(x); return `${h} h ${pad(m)} min`; };

// Prima delle 14 si registra la notte appena passata (iniziata ieri sera)
const defaultNightDate = () => { const now = new Date(); const d = toISO(now); return now.getHours() < 14 ? addDays(d, -1) : d; };

const avg = (arr) => { const v = arr.filter((x) => x != null && !isNaN(x)); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
const roundOrNull = (x) => (x == null ? null : Math.round(x));

// Media delle due misurazioni di pressione; conta solo quelle con massima e minima
function bpAvg(b) {
  const rs = [[b.s1, b.d1, b.p1], [b.s2, b.d2, b.p2]].filter(([s, d]) => s != null && d != null);
  if (!rs.length) return null;
  return { sys: roundOrNull(avg(rs.map((r) => r[0]))), dia: roundOrNull(avg(rs.map((r) => r[1]))), pulse: roundOrNull(avg(rs.map((r) => r[2]))), n: rs.length };
}

// Una misurazione per giorno, la prima della giornata: { "2026-09-01": media, ... }
function bpByDay(bp, ym) {
  const out = {};
  [...bp].filter((b) => b.date.startsWith(ym)).sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
    .forEach((b) => { const a = bpAvg(b); if (a && !out[b.date]) out[b.date] = a; });
  return out;
}

const bpMean = (list) => (list.length ? { sys: roundOrNull(avg(list.map((a) => a.sys))), dia: roundOrNull(avg(list.map((a) => a.dia))) } : null);

function monthStats(data, ym) {
  const ns = data.nights.filter((n) => n.date.startsWith(ym));
  const es = data.episodes.filter((e) => e.date.startsWith(ym));
  return {
    ns, es,
    nights: ns.length,
    wakes: avg(ns.map((n) => n.nWakes)),
    assist: avg(ns.map((n) => n.assistMin)),
    mySleep: avg(ns.map((n) => n.mySleep)),
    agitated: ns.filter((n) => n.agit != null && n.agit >= 2).length,
    episodes: es.length,
    glu: (data.glucose || []).filter((g) => g.date.startsWith(ym) && g.value != null),
    bp: Object.values(bpByDay(data.bp || [], ym)),
    epDur: avg(es.map((e) => duration(e.start, e.end))),
  };
}

// Quanti giorni di fila, a partire dall'ultimo registrato, risulta che non abbia fatto
function noPoopStreak(nights) {
  const seq = [...nights].filter((n) => n.poop != null).sort((a, b) => b.date.localeCompare(a.date));
  let k = 0;
  for (const n of seq) { if (n.poop === false) k++; else break; }
  return k;
}

const monthsWithData = (data) =>
  [...new Set([...data.nights, ...data.episodes].map((x) => x.date.slice(0, 7)))].sort();

function recentDistinct(items, field, max = 3) {
  const out = [];
  [...items].sort((a, b) => b.date.localeCompare(a.date)).forEach((it) => {
    const v = (it[field] || "").trim();
    if (v && !out.includes(v) && out.length < max) out.push(v);
  });
  return out;
}

// Frasi che si ripetono nelle note, proposte come suggerimenti
function noteFragments(nights) {
  const c = {};
  nights.forEach((n) =>
    (n.notes || "").split(/[,.;]/).map((s) => s.trim().toLowerCase())
      .filter((s) => s.length > 3 && s.length <= 45)
      .forEach((s) => { c[s] = (c[s] || 0) + 1; })
  );
  return Object.entries(c).filter(([, k]) => k >= 2).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([s]) => s);
}

const appendText = (base, frag) => (base && base.trim() ? `${base.trim().replace(/[,.]$/, "")}, ${frag}` : frag);

/* ------------------------------------------------------------------ */
/*  Salvataggio                                                        */
/* ------------------------------------------------------------------ */

const errText = (e) => { try { return String((e && (e.message || e.name)) || e); } catch (x) { return "sconosciuto"; } };

function storageApi() {
  try {
    const ls = window.localStorage;
    const probe = "diario-prova";
    ls.setItem(probe, "1");
    ls.removeItem(probe);
    return ls;
  } catch (e) {
    return null;
  }
}

// I dati stanno nella memoria del browser, su questo telefono. Nessun server, nessun account.
async function loadData() {
  const ls = storageApi();
  if (!ls) return { status: "nostorage", detail: "il browser non permette di salvare in questa pagina" };
  let raw = null;
  try { raw = ls.getItem(KEY); }
  catch (e) { return { status: "nostorage", detail: errText(e) }; }
  if (raw == null || raw === "") return { status: "empty" };
  try { return { status: "ok", data: JSON.parse(raw) }; }
  catch (e) { return { status: "unreadable", detail: `dati salvati non validi: ${errText(e)}` }; }
}

async function saveData(data) {
  const ls = storageApi();
  if (!ls) return { ok: false, detail: "salvataggio non disponibile in questa pagina" };
  const payload = JSON.stringify(data);
  try {
    ls.setItem(KEY, payload);
    if (ls.getItem(KEY) !== payload) return { ok: false, detail: "salvato, ma la verifica non restituisce gli stessi dati" };
    return { ok: true };
  } catch (e) {
    return { ok: false, detail: errText(e) };
  }
}

// Unisce i dati trovati salvati con quelli inseriti nel frattempo (vincono i nuovi)
function mergeData(stored, cur) {
  const nm = new Map((stored.nights || []).map((n) => [n.date, n]));
  cur.nights.forEach((n) => nm.set(n.date, n));
  const ek = (e) => `${e.date}|${e.start}`;
  const em = new Map((stored.episodes || []).map((e) => [ek(e), e]));
  cur.episodes.forEach((e) => em.set(ek(e), e));
  const tm = new Map((stored.therapy || []).map((t) => [t.time, t]));
  (cur.therapy || []).forEach((t) => tm.set(t.time, t));
  const gk = (g) => `${g.date}|${g.time}`;
  const gm = new Map((stored.glucose || []).map((g) => [gk(g), g]));
  (cur.glucose || []).forEach((g) => gm.set(gk(g), g));
  const bm = new Map((stored.bp || []).map((b) => [gk(b), b]));
  (cur.bp || []).forEach((b) => bm.set(gk(b), b));
  return { ...cur, therapy: [...tm.values()].map((t) => ({ ...t, id: t.id || uid() })), glucose: [...gm.values()].map((g) => ({ ...g, id: g.id || uid() })), bp: [...bm.values()].map((b) => ({ ...b, id: b.id || uid() })), nights: [...nm.values()].map((n) => ({ ...n, id: n.id || uid() })), episodes: [...em.values()].map((e) => ({ ...e, id: e.id || uid() })) };
}

/* ------------------------------------------------------------------ */
/*  Excel: esportazione e importazione                                 */
/* ------------------------------------------------------------------ */

const serial = (iso) => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
const tfrac = (t) => { const m = toMin(t); return m == null ? null : m / 1440; };
const round1 = (x) => (x == null ? null : Math.round(x * 10) / 10);

function fmtCells(ws, rows, col, fmt) {
  for (let r = 1; r <= rows; r++) {
    const ref = XLSX.utils.encode_cell({ r, c: col });
    if (ws[ref]) ws[ref].z = fmt;
  }
}

function buildWorkbook(data) {
  const wb = XLSX.utils.book_new();
  const byDate = (a, b) => a.date.localeCompare(b.date) || (a.start || "").localeCompare(b.start || "");

  // Riepilogo
  const months = monthsWithData(data);
  const rh = ["Mese", "Notti registrate", "Media risvegli", "Media min. assistenza", "Mie ore di sonno (media)", "Notti agitate (>=2)", "Episodi diurni", "Durata media episodi (min)", "Pressione media"];
  const rrows = months.map((ym) => {
    const s = monthStats(data, ym);
    const bm = bpMean(s.bp);
    return [monthLabel(ym), s.nights, round1(s.wakes), round1(s.assist), round1(s.mySleep), s.agitated, s.episodes, round1(s.epDur), bm ? `${bm.sys}/${bm.dia}` : null];
  });
  const wsR = XLSX.utils.aoa_to_sheet([rh, ...rrows]);
  wsR["!cols"] = [16, 15, 14, 20, 22, 18, 14, 24, 16].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsR, "Riepilogo");

  // Diario notti
  const nh = ["Data", "Giorno", "Ora a letto", "1o risveglio", "2o risveglio", "N. risvegli", "Min. assistenza", "Agitazione (0-3)", "Mie ore di sonno", "Farmaci / variazioni", "Note", "Cacca"];
  const nights = [...data.nights].sort(byDate);
  const nrows = nights.map((n) => {
    const extra = (n.wakes || []).slice(2).filter(Boolean);
    const notes = [n.notes, extra.length ? `altri risvegli: ${extra.join(", ")}` : ""].filter(Boolean).join(" | ");
    return [serial(n.date), GIORNI_XLS[fromISO(n.date).getDay()], tfrac(n.bed), tfrac(n.wakes?.[0]), tfrac(n.wakes?.[1]),
      n.nWakes, n.assistMin, n.agit, n.mySleep, n.meds || null, notes || null,
      n.poop == null ? null : (n.poop ? "Si" : "No")];
  });
  const wsN = XLSX.utils.aoa_to_sheet([nh, ...nrows]);
  fmtCells(wsN, nrows.length, 0, "dd/mm/yyyy");
  [2, 3, 4].forEach((c) => fmtCells(wsN, nrows.length, c, "hh:mm"));
  wsN["!cols"] = [12, 11, 11, 12, 12, 11, 15, 16, 16, 28, 60, 9].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsN, "Diario notti");

  // Episodi diurni
  const eh = ["Data", "Ora inizio", "Ora fine", "Durata (min)", "Tipo", "Intensita' (1-3)", "Possibile causa", "Cosa ha aiutato", "Note"];
  const eps = [...data.episodes].sort(byDate);
  const erows = eps.map((e) => [serial(e.date), tfrac(e.start), tfrac(e.end), duration(e.start, e.end), e.type || null,
    e.intensity, e.cause || null, e.helped || null, e.notes || null]);
  const wsE = XLSX.utils.aoa_to_sheet([eh, ...erows]);
  fmtCells(wsE, erows.length, 0, "dd/mm/yyyy");
  [1, 2].forEach((c) => fmtCells(wsE, erows.length, c, "hh:mm"));
  wsE["!cols"] = [12, 11, 11, 12, 16, 15, 28, 32, 40].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsE, "Episodi diurni");

  // Glicemia
  const gh = ["Data", "Ora", "Valore (mg/dL)", "Contesto", "Note"];
  const glu = [...(data.glucose || [])].sort(byDate);
  if (glu.length) {
    const grows = glu.map((g) => [serial(g.date), tfrac(g.time), g.value, g.tag || null, g.note || null]);
    const wsG = XLSX.utils.aoa_to_sheet([gh, ...grows]);
    fmtCells(wsG, grows.length, 0, "dd/mm/yyyy");
    fmtCells(wsG, grows.length, 1, "hh:mm");
    wsG["!cols"] = [12, 10, 15, 18, 40].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsG, "Glicemia");
  }

  // Pressione
  const bh = ["Data", "Ora", "Massima 1", "Minima 1", "Battiti 1", "Massima 2", "Minima 2", "Battiti 2", "Massima (media)", "Minima (media)", "Battiti (media)", "Note"];
  const bps = [...(data.bp || [])].sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
  if (bps.length) {
    const brows = bps.map((b) => {
      const a = bpAvg(b) || {};
      return [serial(b.date), tfrac(b.time), b.s1, b.d1, b.p1, b.s2, b.d2, b.p2, a.sys ?? null, a.dia ?? null, a.pulse ?? null, b.note || null];
    });
    const wsB = XLSX.utils.aoa_to_sheet([bh, ...brows]);
    fmtCells(wsB, brows.length, 0, "dd/mm/yyyy");
    fmtCells(wsB, brows.length, 1, "hh:mm");
    wsB["!cols"] = [12, 8, 10, 9, 9, 10, 9, 9, 15, 14, 14, 40].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsB, "Pressione");
  }

  // Terapia
  const th = ["Ora", "Farmaci", "Note"];
  const ther = [...(data.therapy || [])].sort((a, b) => (a.time || "").localeCompare(b.time || ""));
  if (ther.length) {
    const trows = ther.map((t) => [tfrac(t.time), t.meds || null, t.note || null]);
    const wsT = XLSX.utils.aoa_to_sheet([th, ...trows]);
    fmtCells(wsT, trows.length, 0, "hh:mm");
    wsT["!cols"] = [10, 50, 30].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsT, "Terapia");
  }

  return wb;
}

const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function xlsxFile(data) {
  const out = XLSX.write(buildWorkbook(data), { bookType: "xlsx", type: "array" });
  return new File([out], `diario_assistenza_${toISO(new Date())}.xlsx`, { type: XLSX_TYPE });
}

function backupFile(data) {
  const text = JSON.stringify({ nights: data.nights, episodes: data.episodes, therapy: data.therapy || [], glucose: data.glucose || [], bp: data.bp || [], savedAt: new Date().toISOString() }, null, 1);
  return new File([text], `diario_backup_${toISO(new Date())}.json`, { type: "application/json" });
}

const serToISO = (v) => {
  if (typeof v === "number") {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  if (typeof v === "string") {
    const iso = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return iso[0];
    const m = v.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
    if (m) { let y = +m[3]; if (y < 100) y += 2000; return `${y}-${pad(+m[2])}-${pad(+m[1])}`; }
  }
  return null;
};
const fracToHM = (v) => {
  if (v == null || v === "") return "";
  if (typeof v === "number") { const m = Math.round((v % 1) * 1440) % 1440; return `${pad(Math.floor(m / 60))}:${pad(m % 60)}`; }
  const m = String(v).match(/(\d{1,2})[:.](\d{2})/);
  return m ? `${pad(+m[1])}:${m[2]}` : "";
};
const cellNum = (v) => (typeof v === "number" ? v : parseDec(v));
const cellBool = (v) => {
  if (v == null || v === "") return null;
  if (typeof v === "boolean") return v;
  const t = String(v).trim().toLowerCase();
  if (/^(si|sì|s|1|x|vero|true)$/.test(t)) return true;
  if (/^(no|n|0|falso|false)$/.test(t)) return false;
  return null;
};
const cellTxt = (v) => (v == null ? "" : String(v).trim());

async function parseXlsx(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const rowsOf = (name, head = "Data") => {
    const ws = wb.Sheets[name];
    if (!ws) return [];
    const a = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const h = a.findIndex((r) => r && r[0] === head);
    if (h < 0) return [];
    // Nel modello originale la riga subito sotto le intestazioni è l'esempio: il Riepilogo la ignora, qui pure
    const start = h > 0 && head === "Data" ? h + 2 : h + 1;
    return a.slice(start).filter((r) => r && (head === "Data" ? serToISO(r[0]) : r[0] != null && r[0] !== ""));
  };
  const nights = rowsOf("Diario notti").map((r) => {
    const w = [fracToHM(r[3]), fracToHM(r[4])];
    const nW = cellNum(r[5]) ?? w.filter(Boolean).length;
    return {
      date: serToISO(r[0]), bed: fracToHM(r[2]), nWakes: nW,
      wakes: Array.from({ length: nW }, (_, i) => w[i] || ""),
      assistMin: cellNum(r[6]), agit: cellNum(r[7]), mySleep: cellNum(r[8]),
      meds: cellTxt(r[9]), notes: cellTxt(r[10]), poop: cellBool(r[11]),
    };
  });
  const episodes = rowsOf("Episodi diurni").map((r) => ({
    date: serToISO(r[0]), start: fracToHM(r[1]), end: fracToHM(r[2]),
    type: cellTxt(r[4]), intensity: cellNum(r[5]),
    cause: cellTxt(r[6]), helped: cellTxt(r[7]), notes: cellTxt(r[8]),
  }));
  const glucose = rowsOf("Glicemia").map((r) => ({
    date: serToISO(r[0]), time: fracToHM(r[1]), value: cellNum(r[2]), tag: cellTxt(r[3]), note: cellTxt(r[4]),
  })).filter((g) => g.value != null);
  const bp = rowsOf("Pressione").map((r) => ({
    date: serToISO(r[0]), time: fracToHM(r[1]),
    s1: cellNum(r[2]), d1: cellNum(r[3]), p1: cellNum(r[4]), s2: cellNum(r[5]), d2: cellNum(r[6]), p2: cellNum(r[7]),
    note: cellTxt(r[11]),
  })).filter((b) => bpAvg(b));
  const therapy = rowsOf("Terapia", "Ora").map((r) => ({
    time: fracToHM(r[0]), meds: cellTxt(r[1]), note: cellTxt(r[2]),
  })).filter((t) => t.time);
  return { nights, episodes, therapy, glucose, bp };
}

function downloadFile(file) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(file);
  a.download = file.name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

/* ------------------------------------------------------------------ */
/*  Stile                                                              */
/* ------------------------------------------------------------------ */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Atkinson+Hyperlegible:wght@400;700&display=swap');
.app{--bg:#1B2030;--surface:#242A3B;--raised:#2D3447;--line:#3B4259;--text:#EDE6D8;--muted:#A8A296;
  --accent:#F0B95A;--accent-ink:#2A1F0C;--a0:#8DB89E;--a1:#D9C27A;--a2:#E39A62;--a3:#DE6E70;--danger:#E58383;
  color-scheme:dark;}
.app[data-theme="day"]{--bg:#ECEEF3;--surface:#FFFFFF;--raised:#F4F5F8;--line:#D2D6DF;--text:#1B2030;--muted:#5A6172;
  --accent:#A86A12;--accent-ink:#FFFFFF;--a0:#3F7F59;--a1:#A5861F;--a2:#C0652A;--a3:#B83C3F;--danger:#B23B3B;
  color-scheme:light;}
.app{min-height:100vh;background:var(--bg);color:var(--text);
  font-family:'Atkinson Hyperlegible',system-ui,-apple-system,'Segoe UI',sans-serif;font-size:17px;line-height:1.45;
  -webkit-tap-highlight-color:transparent;}
.app *{box-sizing:border-box;}
.app button{font:inherit;color:inherit;cursor:pointer;}
.app :focus-visible{outline:3px solid var(--accent);outline-offset:2px;}
.topbar{position:sticky;top:0;z-index:5;background:var(--bg);display:flex;align-items:center;justify-content:space-between;
  padding:10px 16px;max-width:560px;margin:0 auto;}
.brand{font-weight:700;font-size:15px;color:var(--muted);letter-spacing:.01em;}
.main{max-width:560px;margin:0 auto;padding:4px 16px 170px;}
.iconbtn{width:46px;height:46px;border-radius:50%;border:1px solid var(--line);background:var(--surface);
  display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;}
.iconbtn:disabled{opacity:.35;cursor:default;}
.iconbtn.ghost{border-color:transparent;background:transparent;}
.nighthead{display:flex;align-items:flex-start;gap:10px;margin:6px 0 4px;}
.nighthead .mid{flex:1;min-width:0;}
.h1{font-size:28px;line-height:1.15;font-weight:700;margin:2px 0 8px;letter-spacing:-.01em;}
.meta{display:flex;align-items:center;gap:10px;flex-wrap:wrap;color:var(--muted);font-size:15px;}
.pill{display:inline-block;padding:3px 10px;border-radius:999px;border:1px solid var(--line);font-size:14px;}
.pill.on{border-color:var(--accent);color:var(--accent);}
.datein{background:transparent;border:none;color:var(--muted);font:inherit;font-size:15px;padding:0;text-decoration:underline;text-underline-offset:3px;}
.sec{padding:20px 0;border-bottom:1px solid var(--line);}
.sec:last-of-type{border-bottom:none;}
.sech{font-size:18px;font-weight:700;margin:0 0 4px;}
.hint{color:var(--muted);font-size:15px;margin:0 0 12px;}
.row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;}
.grow{flex:1;min-width:0;}
.input,.textarea{width:100%;background:var(--raised);border:1px solid var(--line);border-radius:12px;color:var(--text);
  font:inherit;font-size:18px;padding:0 14px;}
.input{height:52px;}
.textarea{padding:12px 14px;min-height:96px;resize:vertical;font-size:17px;}
.input::placeholder,.textarea::placeholder{color:var(--muted);opacity:.8;}
.timein{max-width:150px;}
.stepper{display:flex;align-items:center;gap:0;border:1px solid var(--line);border-radius:14px;overflow:hidden;background:var(--raised);width:max-content;}
.stepper button{width:56px;height:56px;border:none;background:transparent;display:flex;align-items:center;justify-content:center;}
.stepper button:disabled{opacity:.3;}
.stepper .val{min-width:84px;text-align:center;font-size:24px;font-weight:700;}
.stepper input.val{background:transparent;border:none;color:var(--text);font:inherit;font-size:24px;font-weight:700;width:96px;height:56px;}
.stepper input.val.hm{width:72px;}
.unit{color:var(--muted);font-size:16px;}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px;}
.chip{border:1px solid var(--line);background:var(--surface);border-radius:999px;padding:9px 14px;font-size:15px;line-height:1.2;text-align:left;}
.chip[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:var(--accent-ink);font-weight:700;}
.chip.suggest{border-style:dashed;}
.wakes{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px;margin-top:14px;}
.wakes .lab{font-size:14px;color:var(--muted);margin-bottom:4px;}
.tiles{display:flex;flex-direction:column;gap:8px;}
.tile{display:flex;align-items:center;gap:14px;text-align:left;padding:12px 14px;border-radius:14px;border:2px solid var(--line);background:var(--surface);}
.tile .num{font-size:30px;font-weight:700;width:34px;text-align:center;line-height:1;}
.tile .t{font-weight:700;}
.tile .d{font-size:14px;color:var(--muted);}
.tile[aria-pressed="true"]{border-color:var(--c);background:color-mix(in srgb,var(--c) 16%,var(--surface));}
.tiles.h{flex-direction:row;}
.tiles.h .tile{flex:1;flex-direction:column;gap:2px;padding:12px 6px;text-align:center;}
.savebar{position:fixed;left:0;right:0;bottom:64px;z-index:6;padding:10px 16px;
  background:linear-gradient(to bottom,transparent,var(--bg) 30%);}
.savebar .inner{max-width:560px;margin:0 auto;}
.btn{height:54px;border-radius:14px;border:1px solid var(--line);background:var(--surface);padding:0 18px;
  font-weight:700;display:inline-flex;align-items:center;justify-content:center;gap:8px;}
.btn.primary{width:100%;background:var(--accent);border-color:var(--accent);color:var(--accent-ink);font-size:18px;}
.btn.primary:disabled{opacity:.45;cursor:default;}
.btn.wide{width:100%;}
.btn.danger{color:var(--danger);border-color:color-mix(in srgb,var(--danger) 50%,var(--line));background:transparent;}
.btn.small{height:44px;font-size:15px;padding:0 14px;}
.filepick{position:relative;overflow:hidden;}
.filepick input{position:absolute;inset:0;width:100%;height:100%;opacity:0;cursor:pointer;font-size:0;}
.filepick:focus-within{outline:3px solid var(--accent);outline-offset:2px;}
.linkbtn{display:inline-block;background:none;border:none;color:var(--accent);font-weight:700;padding:6px 0;font-size:15px;}
.tabbar{position:fixed;left:0;right:0;bottom:0;height:64px;z-index:7;background:var(--surface);border-top:1px solid var(--line);}
.tabbar .inner{max-width:560px;margin:0 auto;display:grid;grid-template-columns:repeat(6,1fr);height:100%;}
.tab{border:none;background:transparent;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;
  font-size:10.5px;color:var(--muted);min-width:0;padding:0 1px;overflow:hidden;}
.tab .lbl{max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.tab svg{flex-shrink:0;}
@media (max-width:359px){.tab{font-size:9.5px;}}
.tab[aria-current="page"]{color:var(--accent);font-weight:700;}
.bprow{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;}
.bprow label{display:flex;flex-direction:column;gap:4px;}
.bprow .input{font-size:22px;font-weight:700;text-align:center;padding:0 4px;}
.seg{display:grid;grid-template-columns:1fr 1fr;border:1px solid var(--line);border-radius:12px;overflow:hidden;margin:8px 0 16px;}
.seg button{height:46px;border:none;background:transparent;}
.seg button[aria-pressed="true"]{background:var(--raised);font-weight:700;color:var(--accent);}
.mhead{font-size:15px;color:var(--muted);margin:18px 0 8px;font-weight:700;}
.entry{display:flex;gap:12px;width:100%;text-align:left;padding:12px;border-radius:12px;border:1px solid var(--line);background:var(--surface);margin-bottom:8px;}
.entry .bar{width:6px;border-radius:3px;background:var(--c);flex-shrink:0;}
.entry .t{font-weight:700;}
.entry .d{font-size:15px;color:var(--muted);display:flex;gap:14px;flex-wrap:wrap;}
.entry .dose{display:block;white-space:pre-line;margin-top:2px;}
.entry .n{font-size:15px;margin-top:4px;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;}
.stats{width:100%;border-collapse:collapse;}
.stats td{padding:12px 0;border-bottom:1px solid var(--line);}
.stats td:last-child{text-align:right;font-weight:700;font-size:20px;white-space:nowrap;}
.stats tr:last-child td{border-bottom:none;}
.chart{margin:8px 0 4px;}
.legend{display:flex;flex-wrap:wrap;gap:12px;font-size:14px;color:var(--muted);margin-top:6px;}
.legend span{display:inline-flex;align-items:center;gap:6px;}
.legend i{width:12px;height:12px;border-radius:3px;background:var(--c);display:inline-block;}
.legend i.dot{width:10px;height:10px;border-radius:50%;}
.legend i.hollow{background:transparent;border:1px solid var(--muted);opacity:.7;}
.callout{border-left:4px solid var(--accent);padding:10px 14px;background:var(--surface);border-radius:0 12px 12px 0;margin:16px 0;font-size:16px;}
.empty{padding:28px 0;color:var(--muted);}
.empty p{margin:0 0 14px;}
.toast{position:fixed;top:14px;left:50%;transform:translateX(-50%);z-index:20;background:var(--text);color:var(--bg);
  padding:12px 18px;border-radius:12px;font-weight:700;box-shadow:0 6px 24px rgba(0,0,0,.25);animation:tin .2s ease-out;max-width:90vw;font-size:15px;line-height:1.35;}
@keyframes tin{from{opacity:0;transform:translate(-50%,-8px);}to{opacity:1;transform:translate(-50%,0);}}
.banner{max-width:560px;margin:0 auto 8px;padding:12px 16px;background:color-mix(in srgb,var(--danger) 18%,var(--surface));border-radius:12px;font-size:15px;}
.banner.info{background:var(--surface);border:1px solid var(--line);}
.banner .detail{display:block;margin-top:6px;font-size:13px;color:var(--muted);word-break:break-word;}
.sheet{position:fixed;inset:0;z-index:15;background:var(--bg);overflow:auto;}
.sheet .inner{max-width:560px;margin:0 auto;padding:12px 16px 40px;}
.muted{color:var(--muted);}
.small{font-size:15px;}
@media (prefers-reduced-motion:reduce){.toast{animation:none;}}
`;

/* ------------------------------------------------------------------ */
/*  Componenti di base                                                 */
/* ------------------------------------------------------------------ */

function Stepper({ value, onChange, min = 0, max = 99, step = 1, label }) {
  return (
    <div className="stepper" role="group" aria-label={label}>
      <button type="button" aria-label="Meno" disabled={value <= min} onClick={() => onChange(Math.max(min, value - step))}><Minus size={22} /></button>
      <div className="val" aria-live="polite">{value}</div>
      <button type="button" aria-label="Più" disabled={value >= max} onClick={() => onChange(Math.min(max, value + step))}><Plus size={22} /></button>
    </div>
  );
}

function Chips({ items, onPick, suggest = false, isOn, render }) {
  if (!items.length) return null;
  return (
    <div className="chips">
      {items.map((it) => (
        <button key={String(it)} type="button" className={`chip${suggest ? " suggest" : ""}`}
          aria-pressed={isOn ? isOn(it) : undefined} onClick={() => onPick(it)}>
          {render ? render(it) : it}
        </button>
      ))}
    </div>
  );
}

// Il tocco arriva direttamente sul campo file (invisibile, steso sopra il pulsante):
// è il modo che funziona anche nelle app per telefono, dove aprire il selettore "da codice" viene bloccato.
// Nessun filtro sul tipo di file, perché su Android fa apparire i file .xlsx grigi e non selezionabili.
function FilePick({ className, onFile, children }) {
  return (
    <label className={`${className} filepick`}>
      {children}
      <input type="file" onChange={(e) => { const file = e.target.files?.[0]; e.target.value = ""; if (file) onFile(file); }} />
    </label>
  );
}

function Section({ title, hint, children }) {
  return (
    <section className="sec">
      <h2 className="sech">{title}</h2>
      {hint && <p className="hint">{hint}</p>}
      {!hint && <div style={{ height: 8 }} />}
      {children}
    </section>
  );
}

function ConfirmDelete({ label, onConfirm }) {
  const [ask, setAsk] = useState(false);
  return ask ? (
    <div className="row">
      <button type="button" className="btn danger small" onClick={onConfirm}><Trash2 size={18} />Sì, elimina</button>
      <button type="button" className="btn small" onClick={() => setAsk(false)}>Annulla</button>
    </div>
  ) : (
    <button type="button" className="btn danger small" onClick={() => setAsk(true)}><Trash2 size={18} />{label}</button>
  );
}

/* ------------------------------------------------------------------ */
/*  Notte                                                              */
/* ------------------------------------------------------------------ */

function blankNight(date, data) {
  const last = [...data.nights].sort((a, b) => b.date.localeCompare(a.date))[0];
  return { id: null, date, bed: last?.bed || "20:30", nWakes: 0, wakes: [], assistMin: null, agit: null, mySleep: null, poop: null, meds: "", notes: "" };
}

function NightForm({ data, date, setDate, onSave, onDelete }) {
  const existing = data.nights.find((n) => n.date === date);
  const [f, setF] = useState(() => (existing ? { ...existing, wakes: [...(existing.wakes || [])] } : blankNight(date, data)));
  const [sleepH, setSleepH] = useState(() => splitHM(existing?.mySleep).h);
  const [sleepM, setSleepM] = useState(() => splitHM(existing?.mySleep).m);
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const setWakes = (n) => setF((p) => { const w = [...p.wakes]; while (w.length < n) w.push(""); return { ...p, nWakes: n, wakes: w.slice(0, n) }; });
  const setHM = (h, m) => { setSleepH(h); setSleepM(m); set("mySleep", fromHM(h, m)); };
  // pulsanti +/-: ore di un'ora in un'ora, minuti di 5 in 5 (con passaggio all'ora)
  const bump = (dMin) => {
    const cur = f.mySleep == null ? 7 * 60 : Math.round(f.mySleep * 60);
    const t = Math.max(0, Math.min(24 * 60, cur + dMin));
    setHM(String(Math.floor(t / 60)), String(t % 60));
  };
  const digits = (v, max) => { const d = v.replace(/\D/g, "").slice(0, 2); return d === "" ? "" : String(Math.min(max, Number(d))); };
  const today = toISO(new Date());
  const medsChips = recentDistinct(data.nights, "meds");
  const noteChips = noteFragments(data.nights);

  return (
    <div>
      <div className="nighthead">
        <button type="button" className="iconbtn" aria-label="Notte precedente" onClick={() => setDate(addDays(date, -1))}><ChevronLeft /></button>
        <div className="mid">
          <h1 className="h1">Notte tra {fmtDay(date)} e {fmtDay(addDays(date, 1))}</h1>
          <div className="meta">
            <span className={`pill${existing ? " on" : ""}`}>{existing ? "Già registrata, la stai modificando" : "Nuova"}</span>
            <input type="date" className="datein" value={date} max={today} aria-label="Scegli la data della sera"
              onChange={(e) => e.target.value && setDate(e.target.value)} />
          </div>
        </div>
        <button type="button" className="iconbtn" aria-label="Notte successiva" disabled={date >= today} onClick={() => setDate(addDays(date, 1))}><ChevronRight /></button>
      </div>

      <Section title="Ora a letto">
        <input type="time" className="input timein" value={f.bed} onChange={(e) => set("bed", e.target.value)} aria-label="Ora a letto" />
      </Section>

      <Section title="Risvegli" hint="Quante volte si è svegliata. L'orario è facoltativo.">
        <Stepper label="Numero di risvegli" value={f.nWakes} max={10} onChange={setWakes} />
        {f.nWakes > 0 && (
          <div className="wakes">
            {f.wakes.map((w, i) => (
              <div key={i}>
                <div className="lab">{i + 1}° risveglio</div>
                <input type="time" className="input" value={w} aria-label={`Ora del ${i + 1}° risveglio`}
                  onChange={(e) => setF((p) => { const ws = [...p.wakes]; ws[i] = e.target.value; return { ...p, wakes: ws }; })} />
              </div>
            ))}
          </div>
        )}
      </Section>

      <Section title="Minuti di assistenza" hint="Il tempo in cui sei dovuto intervenire, in totale.">
        <div className="row">
          <input className="input" style={{ maxWidth: 120 }} inputMode="numeric" aria-label="Minuti di assistenza"
            value={f.assistMin ?? ""} placeholder="0"
            onChange={(e) => { const v = e.target.value.replace(/\D/g, ""); set("assistMin", v === "" ? null : Number(v)); }} />
          <span className="unit">minuti</span>
        </div>
        <Chips items={ASSIST_CHIPS} isOn={(m) => f.assistMin === m} onPick={(m) => set("assistMin", m)} render={(m) => `${m}`} />
      </Section>

      <Section title="Agitazione">
        <div className="tiles">
          {AGIT.map((a) => (
            <button key={a.v} type="button" className="tile" style={{ "--c": `var(--a${a.v})` }} aria-pressed={f.agit === a.v}
              onClick={() => set("agit", f.agit === a.v ? null : a.v)}>
              <span className="num" style={{ color: `var(--a${a.v})` }}>{a.v}</span>
              <span><span className="t">{a.label}</span><br /><span className="d">{a.desc}</span></span>
            </button>
          ))}
        </div>
      </Section>

      <Section title="Le tue ore di sonno" hint="Quelle dormite davvero, non quelle passate a letto. È un dato clinico su di te.">
        <div className="row" style={{ gap: 14 }}>
          <div className="stepper" role="group" aria-label="Ore">
            <button type="button" aria-label="Meno un'ora" onClick={() => bump(-60)}><Minus size={22} /></button>
            <input className="val hm" inputMode="numeric" value={sleepH} placeholder="–" aria-label="Ore di sonno"
              onChange={(e) => setHM(digits(e.target.value, 24), sleepM)} />
            <button type="button" aria-label="Più un'ora" onClick={() => bump(60)}><Plus size={22} /></button>
          </div>
          <span className="unit">ore</span>
        </div>
        <div className="row" style={{ gap: 14, marginTop: 10 }}>
          <div className="stepper" role="group" aria-label="Minuti">
            <button type="button" aria-label="Meno cinque minuti" onClick={() => bump(-5)}><Minus size={22} /></button>
            <input className="val hm" inputMode="numeric" value={sleepM} placeholder="–" aria-label="Minuti di sonno"
              onChange={(e) => setHM(sleepH, digits(e.target.value, 59))} />
            <button type="button" aria-label="Più cinque minuti" onClick={() => bump(5)}><Plus size={22} /></button>
          </div>
          <span className="unit">minuti</span>
        </div>
      </Section>

      <Section title="Ha fatto la cacca?" hint="Se saltano più giorni di fila può diventare lei stessa una causa di agitazione e di delirio.">
        <div className="tiles h">
          <button type="button" className="tile" style={{ "--c": "var(--a0)" }} aria-pressed={f.poop === true}
            onClick={() => set("poop", f.poop === true ? null : true)}>
            <span className="t">Sì</span>
          </button>
          <button type="button" className="tile" style={{ "--c": "var(--a2)" }} aria-pressed={f.poop === false}
            onClick={() => set("poop", f.poop === false ? null : false)}>
            <span className="t">No</span>
          </button>
        </div>
      </Section>

      <Section title="Farmaci o variazioni" hint="Ogni cambio di dose o di orario. Serve a collegare un peggioramento a una modifica.">
        <input className="input" value={f.meds} onChange={(e) => set("meds", e.target.value)} placeholder="Farmaco, dose e orario" />
        <Chips items={medsChips.filter((m) => m !== f.meds)} suggest onPick={(m) => set("meds", m)} />
      </Section>

      <Section title="Note">
        <textarea className="textarea" value={f.notes} onChange={(e) => set("notes", e.target.value)} placeholder="Cosa è successo" />
        <Chips items={noteChips} suggest onPick={(frag) => set("notes", appendText(f.notes, frag))} render={(s) => `+ ${s}`} />
      </Section>

      {existing && (
        <div style={{ padding: "8px 0 20px" }}>
          <ConfirmDelete label="Elimina questa notte" onConfirm={() => onDelete(existing.id)} />
        </div>
      )}

      <div className="savebar"><div className="inner">
        <button type="button" className="btn primary" onClick={() => onSave({ ...f, date })}>
          {existing ? "Salva le modifiche" : "Salva la notte"}
        </button>
      </div></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Episodio diurno                                                    */
/* ------------------------------------------------------------------ */

function EpisodeForm({ data, episode, onSave, onDelete, onCancel }) {
  const [f, setF] = useState(() => episode ? { ...episode } : {
    id: null, date: toISO(new Date()), start: nowHM(), end: "", type: "", intensity: null, cause: "", helped: "", notes: "",
  });
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const dur = duration(f.start, f.end);
  const today = toISO(new Date());

  return (
    <div>
      <div className="nighthead">
        <div className="mid">
          <h1 className="h1">{episode ? "Modifica episodio" : "Episodio di giorno"}</h1>
          <p className="hint" style={{ margin: 0 }}>Solo quando succede qualcosa fuori dalla routine. Se la giornata va liscia, non serve scrivere niente.</p>
        </div>
        {episode && <button type="button" className="iconbtn ghost" aria-label="Chiudi senza salvare" onClick={onCancel}><X /></button>}
      </div>

      <Section title="Quando">
        <div className="row" style={{ marginBottom: 12 }}>
          <input type="date" className="input" style={{ maxWidth: 200 }} value={f.date} max={today} onChange={(e) => e.target.value && set("date", e.target.value)} aria-label="Data" />
          {f.date !== today && <button type="button" className="linkbtn" onClick={() => set("date", today)}>Oggi</button>}
        </div>
        <div className="wakes" style={{ marginTop: 0 }}>
          <div>
            <div className="lab">Inizio</div>
            <div className="row" style={{ flexWrap: "nowrap" }}>
              <input type="time" className="input" value={f.start} onChange={(e) => set("start", e.target.value)} aria-label="Ora di inizio" />
            </div>
            <button type="button" className="linkbtn" onClick={() => set("start", nowHM())}>Adesso</button>
          </div>
          <div>
            <div className="lab">Fine</div>
            <input type="time" className="input" value={f.end} onChange={(e) => set("end", e.target.value)} aria-label="Ora di fine" />
            <button type="button" className="linkbtn" onClick={() => set("end", nowHM())}>Adesso</button>
          </div>
        </div>
        <p className="muted small" style={{ margin: "6px 0 0" }}>Durata: {dur == null ? "segna inizio e fine" : `${dur} minuti`}</p>
      </Section>

      <Section title="Tipo">
        <Chips items={EP_TYPES} isOn={(t) => f.type === t} onPick={(t) => set("type", f.type === t ? "" : t)} render={(t) => EP_LABEL[t] || t} />
      </Section>

      <Section title="Intensità">
        <div className="tiles h">
          {INTENS.map((a) => (
            <button key={a.v} type="button" className="tile" style={{ "--c": `var(--a${a.v})` }} aria-pressed={f.intensity === a.v}
              onClick={() => set("intensity", f.intensity === a.v ? null : a.v)}>
              <span className="num" style={{ color: `var(--a${a.v})`, width: "auto" }}>{a.v}</span>
              <span className="d">{a.label}</span>
            </button>
          ))}
        </div>
      </Section>

      <Section title="Possibile causa">
        <input className="input" value={f.cause} onChange={(e) => set("cause", e.target.value)} placeholder="Es. fame, caldo, stanchezza" />
        <Chips items={recentDistinct(data.episodes, "cause").filter((x) => x !== f.cause)} suggest onPick={(x) => set("cause", x)} />
      </Section>

      <Section title="Cosa ha aiutato" hint="La cosa che tra sei settimane non ricorderai.">
        <input className="input" value={f.helped} onChange={(e) => set("helped", e.target.value)} placeholder="Es. mangiare qualcosa, una voce calma" />
        <Chips items={recentDistinct(data.episodes, "helped").filter((x) => x !== f.helped)} suggest onPick={(x) => set("helped", x)} />
      </Section>

      <Section title="Note">
        <textarea className="textarea" value={f.notes} onChange={(e) => set("notes", e.target.value)} placeholder="Altri dettagli" />
      </Section>

      {episode && (
        <div style={{ padding: "8px 0 20px" }}>
          <ConfirmDelete label="Elimina questo episodio" onConfirm={() => onDelete(episode.id)} />
        </div>
      )}

      <div className="savebar"><div className="inner">
        <button type="button" className="btn primary" onClick={() => onSave(f)}>{episode ? "Salva le modifiche" : "Salva l'episodio"}</button>
      </div></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Terapia                                                            */
/* ------------------------------------------------------------------ */

const BLANK_DOSE = { id: null, time: "08:00", meds: "", note: "" };

function Terapia({ items, onSave, onDelete }) {
  const [editing, setEditing] = useState(null); // null = elenco, altrimenti la voce in modifica
  const sorted = [...items].sort((a, b) => (a.time || "").localeCompare(b.time || ""));
  const now = nowHM();
  const next = sorted.find((x) => (x.time || "") >= now) || sorted[0];

  if (editing) {
    const set = (k, v) => setEditing((p) => ({ ...p, [k]: v }));
    return (
      <div>
        <div className="nighthead">
          <h1 className="h1 mid" style={{ margin: 0 }}>{editing.id ? "Modifica orario" : "Nuovo orario"}</h1>
          <button type="button" className="iconbtn ghost" aria-label="Chiudi senza salvare" onClick={() => setEditing(null)}><X /></button>
        </div>

        <Section title="Ora">
          <input type="time" className="input timein" value={editing.time} aria-label="Ora"
            onChange={(e) => set("time", e.target.value)} />
        </Section>

        <Section title="Medicine" hint="Una per riga, con la dose.">
          <textarea className="textarea" value={editing.meds} onChange={(e) => set("meds", e.target.value)}
            placeholder={"Nome e dose\nUn'altra medicina"} style={{ minHeight: 130 }} />
        </Section>

        <Section title="Nota">
          <input className="input" value={editing.note} onChange={(e) => set("note", e.target.value)} placeholder="Es. a stomaco pieno" />
        </Section>

        {editing.id && (
          <div style={{ padding: "8px 0 20px" }}>
            <ConfirmDelete label="Elimina questo orario" onConfirm={() => { onDelete(editing.id); setEditing(null); }} />
          </div>
        )}

        <div className="savebar"><div className="inner">
          <button type="button" className="btn primary" onClick={() => { onSave(editing); setEditing(null); }}>
            {editing.id ? "Salva le modifiche" : "Salva l'orario"}
          </button>
        </div></div>
      </div>
    );
  }

  return (
    <div>
      <h1 className="h1" style={{ marginTop: 6 }}>Terapia</h1>
      <p className="hint">Il promemoria delle medicine da dare, ora per ora. Non registra niente: serve solo a ricordare.</p>

      {!sorted.length ? (
        <div className="empty"><p>Nessun orario impostato.</p></div>
      ) : sorted.map((it) => (
        <div key={it.id} className="entry" style={{ "--c": next && it.id === next.id ? "var(--accent)" : "var(--line)" }}>
          <span className="bar" />
          <span className="grow">
            <span className="t">{it.time}{next && it.id === next.id && <span className="muted small"> prossima</span>}</span>
            <span className="dose">{it.meds}</span>
            {it.note && <span className="d">{it.note}</span>}
          </span>
          <button type="button" className="iconbtn" aria-label={`Modifica l'orario delle ${it.time}`} onClick={() => setEditing({ ...it })}><Pencil size={18} /></button>
        </div>
      ))}

      <div className="savebar"><div className="inner">
        <button type="button" className="btn primary" onClick={() => setEditing({ ...BLANK_DOSE })}><Plus size={20} />Aggiungi un orario</button>
      </div></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Glicemia                                                           */
/* ------------------------------------------------------------------ */

const blankGlu = () => ({ id: null, date: toISO(new Date()), time: nowHM(), value: null, tag: "", note: "" });

function Glicemia({ items, onSave, onDelete }) {
  const [f, setF] = useState(blankGlu);
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const today = toISO(new Date());
  const recenti = [...items].sort((a, b) => (b.date + b.time).localeCompare(a.date + a.time)).slice(0, 10);

  return (
    <div>
      <div className="nighthead">
        <div className="mid">
          <h1 className="h1" style={{ marginBottom: 4 }}>{f.id ? "Modifica misurazione" : "Glicemia"}</h1>
          <p className="hint" style={{ margin: 0 }}>Il valore letto sul glucometro, con l'ora e il momento della giornata.</p>
        </div>
        {f.id && <button type="button" className="iconbtn ghost" aria-label="Annulla la modifica" onClick={() => setF(blankGlu())}><X /></button>}
      </div>

      <Section title="Valore">
        <div className="row">
          <input className="input" style={{ maxWidth: 140, fontSize: 26, fontWeight: 700, color: gluColor(f.value) }}
            inputMode="numeric" aria-label="Valore della glicemia" placeholder="---"
            value={f.value ?? ""}
            onChange={(e) => { const v = e.target.value.replace(/\D/g, "").slice(0, 3); set("value", v === "" ? null : Number(v)); }} />
          <span className="unit">mg/dL</span>
        </div>
      </Section>

      <Section title="Quando">
        <div className="row">
          <input type="date" className="input" style={{ maxWidth: 190 }} value={f.date} max={today}
            aria-label="Data" onChange={(e) => e.target.value && set("date", e.target.value)} />
          <input type="time" className="input timein" value={f.time} aria-label="Ora" onChange={(e) => set("time", e.target.value)} />
        </div>
        <button type="button" className="linkbtn" onClick={() => setF((p) => ({ ...p, date: today, time: nowHM() }))}>Adesso</button>
      </Section>

      <Section title="Momento">
        <Chips items={GLU_TAGS} isOn={(t) => f.tag === t} onPick={(t) => set("tag", f.tag === t ? "" : t)} />
      </Section>

      <Section title="Nota">
        <input className="input" value={f.note} onChange={(e) => set("note", e.target.value)} placeholder="Es. dopo pranzo abbondante" />
      </Section>

      {f.id && (
        <div style={{ padding: "8px 0 4px" }}>
          <ConfirmDelete label="Elimina questa misurazione" onConfirm={() => { onDelete(f.id); setF(blankGlu()); }} />
        </div>
      )}

      {!!recenti.length && (
        <section className="sec">
          <h2 className="sech">Ultime misurazioni</h2>
          <div style={{ height: 8 }} />
          {recenti.map((g) => (
            <button key={g.id} type="button" className="entry" style={{ "--c": gluColor(g.value) }} onClick={() => setF({ ...g })}>
              <span className="bar" />
              <span className="grow">
                <span className="t"><span style={{ color: gluColor(g.value) }}>{g.value}</span> <span className="muted small">mg/dL</span></span>
                <span className="d">
                  <span>{fmtDayShort(g.date)}, {g.time}</span>
                  {g.tag && <span>{g.tag.toLowerCase()}</span>}
                </span>
                {g.note && <span className="n">{g.note}</span>}
              </span>
            </button>
          ))}
        </section>
      )}

      <div className="savebar"><div className="inner">
        <button type="button" className="btn primary" disabled={f.value == null}
          onClick={() => { onSave(f); setF(blankGlu()); }}>
          {f.id ? "Salva le modifiche" : "Salva la misurazione"}
        </button>
      </div></div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Pressione                                                          */
/* ------------------------------------------------------------------ */

const blankBp = () => ({ id: null, date: toISO(new Date()), time: nowHM(), s1: null, d1: null, p1: null, s2: null, d2: null, p2: null, note: "" });

function BpReading({ f, n, set }) {
  const field = (k, label) => (
    <label>
      <span className="muted small">{label}</span>
      <input className="input" inputMode="numeric" placeholder="---"
        aria-label={`${label}, ${n === 1 ? "prima" : "seconda"} misurazione`}
        value={f[k + n] ?? ""}
        onChange={(e) => { const v = e.target.value.replace(/\D/g, "").slice(0, 3); set(k + n, v === "" ? null : Number(v)); }} />
    </label>
  );
  return <div className="bprow">{field("s", "Massima")}{field("d", "Minima")}{field("p", "Battiti")}</div>;
}

function Pressione({ items, onSave, onDelete }) {
  const [f, setF] = useState(blankBp);
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const today = toISO(new Date());
  const a = bpAvg(f);
  const recenti = [...items].sort((x, y) => (y.date + y.time).localeCompare(x.date + x.time)).slice(0, 10);

  return (
    <div>
      <div className="nighthead">
        <div className="mid">
          <h1 className="h1" style={{ marginBottom: 4 }}>{f.id ? "Modifica misurazione" : "Pressione"}</h1>
        </div>
        {f.id && <button type="button" className="iconbtn ghost" aria-label="Annulla la modifica" onClick={() => setF(blankBp())}><X /></button>}
      </div>

      <Section title="Prima misurazione"><BpReading f={f} n={1} set={set} /></Section>
      <Section title="Seconda misurazione"><BpReading f={f} n={2} set={set} /></Section>

      <Section title="Media">
        {a ? (
          <div className="row" style={{ alignItems: "baseline" }}>
            <span style={{ fontSize: 30, fontWeight: 700, color: bpColor(a.sys, a.dia) }}>{a.sys}/{a.dia}</span>
            <span className="unit">mmHg</span>
            {a.pulse != null && <span className="unit">{a.pulse} battiti</span>}
          </div>
        ) : (
          <p className="hint" style={{ margin: 0 }}>Compare quando inserisci massima e minima.</p>
        )}
      </Section>

      <Section title="Quando">
        <div className="row">
          <input type="date" className="input" style={{ maxWidth: 190 }} value={f.date} max={today}
            aria-label="Data" onChange={(e) => e.target.value && set("date", e.target.value)} />
          <input type="time" className="input timein" value={f.time} aria-label="Ora" onChange={(e) => set("time", e.target.value)} />
        </div>
        <button type="button" className="linkbtn" onClick={() => setF((p) => ({ ...p, date: today, time: nowHM() }))}>Adesso</button>
      </Section>

      <Section title="Nota">
        <input className="input" value={f.note} onChange={(e) => set("note", e.target.value)} />
      </Section>

      {f.id && (
        <div style={{ padding: "8px 0 4px" }}>
          <ConfirmDelete label="Elimina questa misurazione" onConfirm={() => { onDelete(f.id); setF(blankBp()); }} />
        </div>
      )}

      {!!recenti.length && (
        <section className="sec">
          <h2 className="sech">Ultime misurazioni</h2>
          <div style={{ height: 8 }} />
          {recenti.map((b) => {
            const m = bpAvg(b);
            const col = m ? bpColor(m.sys, m.dia) : "var(--muted)";
            return (
              <button key={b.id} type="button" className="entry" style={{ "--c": col }} onClick={() => setF({ ...b })}>
                <span className="bar" />
                <span className="grow">
                  <span className="t"><span style={{ color: col }}>{m ? `${m.sys}/${m.dia}` : "–"}</span> <span className="muted small">mmHg</span></span>
                  <span className="d">
                    <span>{fmtDayShort(b.date)}, {b.time}</span>
                    {m && m.pulse != null && <span>{m.pulse} battiti</span>}
                    {m && m.n === 2 && <span>media di 2</span>}
                  </span>
                  {b.note && <span className="n">{b.note}</span>}
                </span>
              </button>
            );
          })}
        </section>
      )}

      <div className="savebar"><div className="inner">
        <button type="button" className="btn primary" disabled={!a}
          onClick={() => { onSave(f); setF(blankBp()); }}>
          {f.id ? "Salva le modifiche" : "Salva la misurazione"}
        </button>
      </div></div>
    </div>
  );
}

// Glicemia e pressione condividono la scheda: la barra in basso non ha posto per un'altra icona
function Misure({ kind, setKind, glucose, bp, onSaveGlu, onDeleteGlu, onSaveBp, onDeleteBp }) {
  return (
    <div>
      <div className="seg" style={{ marginTop: 6 }}>
        <button type="button" aria-pressed={kind === "glicemia"} onClick={() => setKind("glicemia")}>Glicemia</button>
        <button type="button" aria-pressed={kind === "pressione"} onClick={() => setKind("pressione")}>Pressione</button>
      </div>
      {kind === "pressione"
        ? <Pressione items={bp} onSave={onSaveBp} onDelete={onDeleteBp} />
        : <Glicemia items={glucose} onSave={onSaveGlu} onDelete={onDeleteGlu} />}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Diario (elenco)                                                    */
/* ------------------------------------------------------------------ */

function EmptyState({ onSettings }) {
  return (
    <div className="empty">
      <p>Qui compariranno le notti e gli episodi che registri. Comincia dalla scheda Notte, oppure porta dentro quello che hai già nel tuo Excel.</p>
      <p><button type="button" className="btn wide" onClick={onSettings}><Upload size={18} />Importa il tuo file Excel</button></p>
    </div>
  );
}

function Diario({ data, onEditNight, onEditEp, onSettings }) {
  const [mode, setMode] = useState("notti");
  const isEmpty = !data.nights.length && !data.episodes.length;
  const items = mode === "notti" ? data.nights : data.episodes;
  const sorted = [...items].sort((a, b) => b.date.localeCompare(a.date) || (b.start || "").localeCompare(a.start || ""));
  const groups = [];
  sorted.forEach((it) => {
    const ym = it.date.slice(0, 7);
    if (!groups.length || groups[groups.length - 1].ym !== ym) groups.push({ ym, items: [] });
    groups[groups.length - 1].items.push(it);
  });

  return (
    <div>
      <h1 className="h1" style={{ marginTop: 6 }}>Diario</h1>
      {isEmpty ? <EmptyState onSettings={onSettings} /> : (
        <>
          <div className="seg">
            <button type="button" aria-pressed={mode === "notti"} onClick={() => setMode("notti")}>Notti ({data.nights.length})</button>
            <button type="button" aria-pressed={mode === "episodi"} onClick={() => setMode("episodi")}>Episodi ({data.episodes.length})</button>
          </div>
          {!sorted.length && <p className="empty">{mode === "notti" ? "Nessuna notte registrata." : "Nessun episodio registrato."}</p>}
          {groups.map((g) => (
            <div key={g.ym}>
              <div className="mhead">{monthLabel(g.ym)}</div>
              {g.items.map((it) => mode === "notti" ? (
                <button key={it.id} type="button" className="entry" style={{ "--c": it.agit == null ? "var(--line)" : `var(--a${it.agit})` }} onClick={() => onEditNight(it.date)}>
                  <span className="bar" />
                  <span className="grow">
                    <span className="t">{fmtDayShort(it.date)} → {fmtDayShort(addDays(it.date, 1))}</span>
                    <span className="d">
                      <span>{plural(it.nWakes || 0, "risveglio", "risvegli")}</span>
                      {it.assistMin != null && <span>{it.assistMin} min</span>}
                      {it.agit != null && <span>agitazione {it.agit}</span>}
                      {it.mySleep != null && <span>tu {fmtHM(it.mySleep)}</span>}
                      {it.poop === false && <span>niente cacca</span>}
                    </span>
                    {it.meds && <span className="n" style={{ color: "var(--accent)" }}>{it.meds}</span>}
                    {it.notes && <span className="n">{it.notes}</span>}
                  </span>
                </button>
              ) : (
                <button key={it.id} type="button" className="entry" style={{ "--c": it.intensity ? `var(--a${it.intensity})` : "var(--line)" }} onClick={() => onEditEp(it.id)}>
                  <span className="bar" />
                  <span className="grow">
                    <span className="t">{fmtDayShort(it.date)}, {it.start || "?"}{it.end ? `–${it.end}` : ""}</span>
                    <span className="d">
                      <span>{EP_LABEL[it.type] || it.type || "Tipo non indicato"}</span>
                      {duration(it.start, it.end) != null && <span>{duration(it.start, it.end)} min</span>}
                      {it.intensity && <span>intensità {it.intensity}</span>}
                    </span>
                    {it.helped && <span className="n">Ha aiutato: {it.helped}</span>}
                  </span>
                </button>
              ))}
            </div>
          ))}
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Riepilogo                                                          */
/* ------------------------------------------------------------------ */

function MonthChart({ data, ym }) {
  const [y, m] = ym.split("-").map(Number);
  const days = new Date(y, m, 0).getDate();
  const W = days * 12, H = 126, base = 96, scale = 8; // 10 ore = 80 px
  const dotY = base + 7;
  const byDate = Object.fromEntries(data.nights.map((n) => [n.date, n]));
  const y7 = base - 7 * scale;
  return (
    <svg className="chart" viewBox={`0 0 ${W} ${H}`} width="100%" role="img"
      aria-label="Ogni barra è una notte: l'altezza sono le tue ore di sonno, il colore l'agitazione. Sotto, un pallino pieno nei giorni in cui ha fatto la cacca">
      <line x1="0" x2={W} y1={y7} y2={y7} style={{ stroke: "var(--line)" }} strokeDasharray="3 3" />
      <text x={W - 2} y={y7 - 3} fontSize="8" textAnchor="end" style={{ fill: "var(--muted)" }}>7 h</text>
      <line x1="0" x2={W} y1={base} y2={base} style={{ stroke: "var(--line)" }} />
      {Array.from({ length: days }, (_, i) => {
        const iso = `${ym}-${pad(i + 1)}`;
        const n = byDate[iso];
        const x = i * 12 + 2.5;
        const col = n && n.agit != null ? `var(--a${n.agit})` : "var(--muted)";
        let bar = null;
        if (n) {
          const h = n.mySleep != null ? Math.max(3, Math.min(n.mySleep, 11) * scale) : 3;
          bar = <rect x={x} y={base - h} width="7" height={h} rx="2" style={{ fill: col }} opacity={n.mySleep != null ? 1 : 0.6} />;
        }
        const lab = i === 0 || (i + 1) % 5 === 0;
        return (
          <g key={iso}>
            {bar}
            {!n && <circle cx={x + 3.5} cy={base - 2} r="1" style={{ fill: "var(--line)" }} />}
            {n && n.poop === true && <circle cx={x + 3.5} cy={dotY} r="2.6" style={{ fill: "var(--a0)" }} />}
            {n && n.poop === false && <circle cx={x + 3.5} cy={dotY} r="2.2" fill="none" strokeWidth="1" style={{ stroke: "var(--muted)" }} opacity="0.7" />}
            {lab && <text x={x + 3.5} y={base + 21} fontSize="8.5" textAnchor="middle" style={{ fill: "var(--muted)" }}>{i + 1}</text>}
          </g>
        );
      })}
    </svg>
  );
}

// Stesso passo orizzontale del grafico delle notti, così i due si leggono in colonna
function GlucoseChart({ data, ym }) {
  const [y, m] = ym.split("-").map(Number);
  const days = new Date(y, m, 0).getDate();
  const W = days * 12, top = 8, base = 92, lo = 40, hi = 320;
  const yOf = (v) => base - ((Math.max(lo, Math.min(hi, v)) - lo) / (hi - lo)) * (base - top);
  const glu = (data.glucose || []).filter((g) => g.date.startsWith(ym) && g.value != null);
  const eps = (data.episodes || []).filter((e) => e.date.startsWith(ym));
  const xOf = (iso) => (Number(iso.slice(8, 10)) - 1) * 12 + 5.5;

  return (
    <svg className="chart" viewBox={`0 0 ${W} ${base + 30}`} width="100%" role="img"
      aria-label="Glicemia del mese, con sotto i giorni in cui ci sono stati episodi">
      {[100, 180, 250].map((v) => (
        <g key={v}>
          <line x1="0" x2={W} y1={yOf(v)} y2={yOf(v)} style={{ stroke: "var(--line)" }} strokeDasharray="3 3" />
          <text x={W - 2} y={yOf(v) - 2} fontSize="8" textAnchor="end" style={{ fill: "var(--muted)" }}>{v}</text>
        </g>
      ))}
      <line x1="0" x2={W} y1={base} y2={base} style={{ stroke: "var(--line)" }} />
      {glu.map((g) => (
        <circle key={g.id} cx={xOf(g.date)} cy={yOf(g.value)} r="3" style={{ fill: gluColor(g.value) }} />
      ))}
      {eps.map((e) => (
        <rect key={e.id} x={xOf(e.date) - 3} y={base + 4} width="6" height="6" rx="1.5"
          style={{ fill: e.intensity ? `var(--a${e.intensity})` : "var(--muted)" }} />
      ))}
      {Array.from({ length: days }, (_, i) => (
        (i === 0 || (i + 1) % 5 === 0) &&
        <text key={i} x={i * 12 + 5.5} y={base + 24} fontSize="8.5" textAnchor="middle" style={{ fill: "var(--muted)" }}>{i + 1}</text>
      ))}
    </svg>
  );
}

// Una barra per giorno, dalla minima alla massima, sulla stessa scala dei giorni degli altri grafici
function BpChart({ data, ym }) {
  const [y, m] = ym.split("-").map(Number);
  const days = new Date(y, m, 0).getDate();
  const W = days * 12, top = 8, base = 100, lo = 40, hi = 200;
  const yOf = (v) => base - ((Math.max(lo, Math.min(hi, v)) - lo) / (hi - lo)) * (base - top);
  const xOf = (iso) => (Number(iso.slice(8, 10)) - 1) * 12 + 5.5;
  const byDay = bpByDay(data.bp || [], ym);

  return (
    <svg className="chart" viewBox={`0 0 ${W} ${base + 22}`} width="100%" role="img"
      aria-label="Pressione del mese: una barra per giorno, dalla minima alla massima, colorata secondo la media">
      {[85, 135].map((v) => (
        <g key={v}>
          <line x1="0" x2={W} y1={yOf(v)} y2={yOf(v)} style={{ stroke: "var(--line)" }} strokeDasharray="3 3" />
          <text x={W - 2} y={yOf(v) - 2} fontSize="8" textAnchor="end" style={{ fill: "var(--muted)" }}>{v}</text>
        </g>
      ))}
      <line x1="0" x2={W} y1={base} y2={base} style={{ stroke: "var(--line)" }} />
      {Object.entries(byDay).map(([iso, a]) => {
        const x = xOf(iso), col = bpColor(a.sys, a.dia);
        return (
          <g key={iso}>
            <line x1={x} x2={x} y1={yOf(a.sys)} y2={yOf(a.dia)} strokeWidth="2.5" strokeLinecap="round" style={{ stroke: col }} />
            <circle cx={x} cy={yOf(a.sys)} r="3" style={{ fill: col }} />
            <circle cx={x} cy={yOf(a.dia)} r="3" style={{ fill: col }} />
          </g>
        );
      })}
      {Array.from({ length: days }, (_, i) => (
        (i === 0 || (i + 1) % 5 === 0) &&
        <text key={i} x={i * 12 + 5.5} y={base + 16} fontSize="8.5" textAnchor="middle" style={{ fill: "var(--muted)" }}>{i + 1}</text>
      ))}
    </svg>
  );
}

function Riepilogo({ data, onExport, onSettings }) {
  const months = monthsWithData(data);
  const cur = toISO(new Date()).slice(0, 7);
  const [ym, setYm] = useState(months.includes(cur) || !months.length ? cur : months[months.length - 1]);
  const s = monthStats(data, ym);
  const [y, m] = ym.split("-").map(Number);
  const prevYm = `${m === 1 ? y - 1 : y}-${pad(m === 1 ? 12 : m - 1)}`;
  const nextYm = `${m === 12 ? y + 1 : y}-${pad(m === 12 ? 1 : m + 1)}`;
  const prev = monthStats(data, prevYm);
  const streak = noPoopStreak(data.nights);
  const sleepDrop = s.mySleep != null && prev.mySleep != null && prev.mySleep - s.mySleep >= 0.5;

  if (!months.length) {
    return (
      <div>
        <h1 className="h1" style={{ marginTop: 6 }}>Riepilogo</h1>
        <EmptyState onSettings={onSettings} />
      </div>
    );
  }

  return (
    <div>
      <div className="nighthead">
        <button type="button" className="iconbtn" aria-label="Mese precedente" onClick={() => setYm(prevYm)}><ChevronLeft /></button>
        <div className="mid" style={{ textAlign: "center" }}>
          <h1 className="h1" style={{ marginBottom: 0 }}>{monthLabel(ym)}</h1>
          <div className="muted small">{plural(s.nights, "notte registrata", "notti registrate")}</div>
        </div>
        <button type="button" className="iconbtn" aria-label="Mese successivo" disabled={ym >= cur} onClick={() => setYm(nextYm)}><ChevronRight /></button>
      </div>

      <section className="sec">
        <MonthChart data={data} ym={ym} />
        <div className="legend">
          {AGIT.map((a) => <span key={a.v} style={{ "--c": `var(--a${a.v})` }}><i />{a.v} {a.label.split(",")[0].toLowerCase()}</span>)}
        </div>
        <div className="legend">
          <span style={{ "--c": "var(--a0)" }}><i className="dot" />ha fatto la cacca</span>
          <span><i className="dot hollow" />niente</span>
        </div>
        <p className="hint" style={{ marginTop: 8 }}>Ogni barra è una notte: l'altezza sono le tue ore di sonno, il colore è l'agitazione. I pallini sotto la linea sono la cacca.</p>
      </section>

      {streak >= 3 && (
        <div className="callout">
          Risulta che non faccia la cacca da {plural(streak, "giorno", "giorni")} di fila. La stitichezza è una causa frequente di agitazione e delirio: se continua, sentine il medico.
        </div>
      )}

      {sleepDrop && (
        <div className="callout">
          Le tue ore di sonno sono scese da {fmtHM(prev.mySleep)} a {fmtHM(s.mySleep)} in media rispetto a {MESI[Number(prevYm.slice(5)) - 1]}. Dillo al geriatra: il carico di chi assiste è un dato che usano per decidere.
        </div>
      )}

      {(s.glu.length > 0 || s.es.length > 0) && (
        <section className="sec">
          <h2 className="sech">Glicemia ed episodi</h2>
          <div style={{ height: 4 }} />
          <GlucoseChart data={data} ym={ym} />
          <p className="hint" style={{ marginTop: 8 }}>
            Ogni pallino è una misurazione, i quadratini sotto la linea sono gli episodi di quel giorno. Le scale dei giorni coincidono con il grafico delle notti qui sopra, quindi le colonne si leggono in verticale.
            {s.glu.length > 0 && <> Media del mese: {fmtNum(avg(s.glu.map((g) => g.value)), 0)} mg/dL su {plural(s.glu.length, "misurazione", "misurazioni")}.</>}
          </p>
        </section>
      )}

      {s.bp.length > 0 && (() => {
        const bm = bpMean(s.bp);
        return (
          <section className="sec">
            <h2 className="sech">Pressione</h2>
            <div style={{ height: 4 }} />
            <BpChart data={data} ym={ym} />
            <div className="legend">
              {BP_LEGEND.map((l) => <span key={l.l} style={{ "--c": `var(--a${l.l})` }}><i />{l.label}</span>)}
            </div>
            <p className="hint" style={{ marginTop: 8 }}>
              Ogni barra va dalla minima alla massima del giorno, media delle due misurazioni.
              {" "}Media del mese: <b style={{ color: bpColor(bm.sys, bm.dia) }}>{bm.sys}/{bm.dia}</b> mmHg su {plural(s.bp.length, "giorno", "giorni")}.
            </p>
          </section>
        );
      })()}

      <section className="sec">
        <table className="stats"><tbody>
          <tr><td>Media risvegli</td><td>{fmtNum(s.wakes)}</td></tr>
          <tr><td>Media minuti di assistenza</td><td>{fmtNum(s.assist, 0)}</td></tr>
          <tr><td>Le tue ore di sonno (media)</td><td>{fmtHM(s.mySleep)}</td></tr>
          <tr><td>Notti agitate (2 o più)</td><td>{s.agitated}</td></tr>
          <tr><td>Episodi di giorno</td><td>{s.episodes}</td></tr>
          <tr><td>Durata media episodi</td><td>{s.epDur == null ? "–" : `${fmtNum(s.epDur, 0)} min`}</td></tr>
        </tbody></table>
      </section>

      {s.ns.some((n) => n.meds) && (
        <section className="sec">
          <h2 className="sech">Farmaci e variazioni del mese</h2>
          {s.ns.filter((n) => n.meds).sort((a, b) => a.date.localeCompare(b.date)).map((n) => (
            <p key={n.id} style={{ margin: "6px 0" }}><span className="muted">{fmtDayShort(n.date)}:</span> {n.meds}</p>
          ))}
        </section>
      )}

      <section className="sec">
        <button type="button" className="btn wide" onClick={onExport}><Download size={20} />Scarica Excel per la visita</button>
        <p className="hint" style={{ marginTop: 8 }}>Contiene il riepilogo di tutti i mesi, il diario delle notti, gli episodi, la glicemia, la pressione e la terapia. Una volta scaricato lo condividi dai download, con il tasto condividi del telefono.</p>
      </section>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Impostazioni                                                       */
/* ------------------------------------------------------------------ */

function SettingsSheet({ data, onClose, onImport, onRestore, onExport, onClear, toast }) {
  const handleXls = async (file) => {
    if (!/\.xlsx?$/i.test(file.name)) { toast("Scegli un file Excel (.xlsx)"); return; }
    try {
      const res = await parseXlsx(file);
      if (!res.nights.length && !res.episodes.length) { toast("Nel file non ho trovato i fogli «Diario notti» o «Episodi diurni» con dati"); return; }
      onImport(res);
    } catch (err) {
      toast("Non riesco a leggere questo file. Serve un file .xlsx");
    }
  };

  const handleJson = async (file) => {
    try {
      const d = JSON.parse(await file.text());
      if (!Array.isArray(d.nights) || !Array.isArray(d.episodes)) throw new Error("formato");
      onRestore(d);
    } catch (err) {
      toast("Questo file non è un backup del diario");
    }
  };

  const backup = () => {
    try {
      downloadFile(backupFile(data));
      toast("Backup scaricato");
    } catch (e) {
      toast("Download non riuscito");
    }
  };

  return (
    <div className="sheet" role="dialog" aria-modal="true" aria-label="Impostazioni">
      <div className="inner">
        <div className="nighthead" style={{ alignItems: "center" }}>
          <h1 className="h1 mid" style={{ margin: 0 }}>Dati e backup</h1>
          <button type="button" className="iconbtn ghost" aria-label="Chiudi" onClick={onClose}><X /></button>
        </div>

        <Section title="Cosa c'è salvato" hint="I dati restano su questo telefono, nella memoria del browser. Non passano da nessun server. Scarica un backup ogni tanto: se cambi telefono o cancelli i dati di navigazione, è l'unico modo per ritrovarli.">
          <p style={{ margin: 0 }}>{plural(data.nights.length, "notte", "notti")} e {plural(data.episodes.length, "episodio", "episodi")}.</p>
        </Section>

        <Section title="Importa dal tuo Excel" hint="Legge i fogli «Diario notti» ed «Episodi diurni», salta la riga d'esempio come fa il tuo Riepilogo e sostituisce le notti con la stessa data.">
          <FilePick className="btn" onFile={handleXls}><Upload size={18} />Scegli il file Excel</FilePick>
        </Section>

        <Section title="Esporta">
          <div className="row">
            <button type="button" className="btn" onClick={onExport}><Download size={18} />Scarica Excel</button>
            <button type="button" className="btn" onClick={backup}><Download size={18} />Scarica backup</button>
          </div>
          <p className="hint" style={{ marginTop: 10, marginBottom: 0 }}>I file finiscono nella cartella Download: da lì li condividi con il tasto del telefono.</p>
          <div style={{ marginTop: 10 }}>
            <FilePick className="linkbtn" onFile={handleJson}>Ripristina da un backup</FilePick>
          </div>
        </Section>

        <Section title="Cancella tutto" hint="Elimina tutte le notti e gli episodi da questa app. Non si può annullare.">
          <ConfirmDelete label="Cancella tutti i dati" onConfirm={onClear} />
        </Section>

        <Section title="Versione">
          <p style={{ margin: 0 }}>Diario dell'assistenza {APP_VERSION}</p>
        </Section>

        <section className="sec" style={{ borderBottom: "none" }}>
          <p style={{ margin: "0 0 12px" }}>Se ti piace, considera di supportarmi con un caffè.</p>
          <a className="btn" href="https://ko-fi.com/na103" target="_blank" rel="noopener noreferrer" style={{ textDecoration: "none", color: "var(--text)" }}>
            <Coffee size={18} />Offrimi un caffè su Ko-fi
          </a>
        </section>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  App                                                                */
/* ------------------------------------------------------------------ */

export default function App() {
  const [data, setData] = useState({ nights: [], episodes: [], therapy: [], glucose: [], bp: [], theme: "night" });
  const [status, setStatus] = useState("loading"); // loading | ready | memory | unreadable
  const [saveErr, setSaveErr] = useState(null);
  const [notice, setNotice] = useState(null); // { kind, detail }
  const [tab, setTab] = useState("notte");
  const [measureKind, setMeasureKind] = useState("glicemia");
  const [nightDate, setNightDate] = useState(defaultNightDate);
  const [nonce, setNonce] = useState(0);
  const [editEpId, setEditEpId] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [toastMsg, setToastMsg] = useState(null);
  const toastTimer = useRef();
  const skipSave = useRef(true);
  const verifyBeforeWrite = useRef(false);

  const toast = (msg) => {
    setToastMsg(msg);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastMsg(null), 2600);
  };

  const init = async () => {
    setStatus("loading");
    const r = await loadData();
    skipSave.current = true;
    if (r.status === "ok") {
      const d = r.data || {};
      setData((cur) => {
        const loaded = { nights: d.nights || [], episodes: d.episodes || [], therapy: d.therapy || [], glucose: d.glucose || [], bp: d.bp || [], theme: d.theme || cur.theme || "night" };
        // se nel frattempo era stato inserito qualcosa, lo tengo
        if (cur.nights.length || cur.episodes.length) { skipSave.current = false; return mergeData(loaded, cur); }
        return loaded;
      });
      verifyBeforeWrite.current = false;
      setNotice(null);
      setStatus("ready");
    } else if (r.status === "empty") {
      verifyBeforeWrite.current = false;
      setNotice(null);
      setStatus("ready");
    } else if (r.status === "unknown") {
      verifyBeforeWrite.current = true;
      setNotice({ kind: "unknown", detail: r.detail });
      setStatus("ready");
    } else if (r.status === "nostorage") {
      setNotice({ kind: "nostorage", detail: r.detail });
      setStatus("memory");
    } else {
      setNotice({ kind: "unreadable", detail: r.detail });
      setStatus("unreadable");
    }
    setNonce((x) => x + 1);
  };

  useEffect(() => { init(); }, []);

  useEffect(() => {
    if (status !== "ready") return;
    if (skipSave.current) { skipSave.current = false; return; }
    (async () => {
      if (verifyBeforeWrite.current) {
        // Prima di scrivere per la prima volta riprovo a leggere, per non cancellare niente
        const r = await loadData();
        if (r.status === "ok") {
          verifyBeforeWrite.current = false;
          setNotice(null);
          setData((cur) => mergeData(r.data || {}, cur)); // il salvataggio riparte con i dati uniti
          setNonce((x) => x + 1);
          return;
        }
        if (r.status === "unreadable") {
          setNotice({ kind: "unreadable", detail: r.detail });
          setStatus("unreadable");
          return;
        }
        verifyBeforeWrite.current = false;
      }
      const res = await saveData(data);
      if (!res.ok && /not available/i.test(res.detail || "")) {
        setNotice({ kind: "nostorage", detail: res.detail });
        setStatus("memory");
        return;
      }
      if (res.ok) { setSaveErr(null); setNotice((n) => (n && n.kind === "unknown" ? null : n)); }
      else setSaveErr(res.detail || "errore sconosciuto");
    })();
  }, [data, status]);

  const forceStart = () => {
    verifyBeforeWrite.current = false;
    skipSave.current = false;
    setNotice(null);
    setStatus("ready");
  };

  const changeNightDate = (d) => { setNightDate(d); setNonce((x) => x + 1); };

  const saveNight = (n) => {
    setData((d) => {
      const others = d.nights.filter((x) => x.date !== n.date && x.id !== n.id);
      const clean = { ...n, id: n.id || uid(), wakes: (n.wakes || []).slice(0, n.nWakes) };
      return { ...d, nights: [...others, clean] };
    });
    setNonce((x) => x + 1);
    toast("Notte salvata");
  };
  const deleteNight = (id) => {
    setData((d) => ({ ...d, nights: d.nights.filter((x) => x.id !== id) }));
    setNonce((x) => x + 1);
    toast("Notte eliminata");
  };
  const saveEpisode = (e) => {
    const isNew = !e.id;
    setData((d) => ({ ...d, episodes: [...d.episodes.filter((x) => x.id !== e.id), { ...e, id: e.id || uid() }] }));
    setEditEpId(null);
    setNonce((x) => x + 1);
    toast(isNew ? "Episodio salvato" : "Episodio aggiornato");
  };
  const deleteEpisode = (id) => {
    setData((d) => ({ ...d, episodes: d.episodes.filter((x) => x.id !== id) }));
    setEditEpId(null);
    setNonce((x) => x + 1);
    toast("Episodio eliminato");
  };

  const saveTherapy = (t) => {
    const isNew = !t.id;
    setData((d) => ({ ...d, therapy: [...(d.therapy || []).filter((x) => x.id !== t.id), { ...t, id: t.id || uid() }] }));
    toast(isNew ? "Orario aggiunto" : "Orario aggiornato");
  };
  const deleteTherapy = (id) => {
    setData((d) => ({ ...d, therapy: (d.therapy || []).filter((x) => x.id !== id) }));
    toast("Orario eliminato");
  };

  const saveGlucose = (g) => {
    const isNew = !g.id;
    setData((d) => ({ ...d, glucose: [...(d.glucose || []).filter((x) => x.id !== g.id), { ...g, id: g.id || uid() }] }));
    toast(isNew ? "Misurazione salvata" : "Misurazione aggiornata");
  };
  const deleteGlucose = (id) => {
    setData((d) => ({ ...d, glucose: (d.glucose || []).filter((x) => x.id !== id) }));
    toast("Misurazione eliminata");
  };

  const saveBp = (b) => {
    const isNew = !b.id;
    setData((d) => ({ ...d, bp: [...(d.bp || []).filter((x) => x.id !== b.id), { ...b, id: b.id || uid() }] }));
    toast(isNew ? "Misurazione salvata" : "Misurazione aggiornata");
  };
  const deleteBp = (id) => {
    setData((d) => ({ ...d, bp: (d.bp || []).filter((x) => x.id !== id) }));
    toast("Misurazione eliminata");
  };

  const mergeIn = (inc, msg) => {
    setData((d) => {
      const nm = new Map(d.nights.map((n) => [n.date, n]));
      inc.nights.forEach((n) => nm.set(n.date, { ...n, id: nm.get(n.date)?.id || uid() }));
      const ek = (e) => `${e.date}|${e.start}`;
      const em = new Map(d.episodes.map((e) => [ek(e), e]));
      inc.episodes.forEach((e) => em.set(ek(e), { ...e, id: em.get(ek(e))?.id || uid() }));
      const tm = new Map((d.therapy || []).map((t) => [t.time, t]));
      (inc.therapy || []).forEach((t) => tm.set(t.time, { ...t, id: tm.get(t.time)?.id || uid() }));
      const gk = (g) => `${g.date}|${g.time}`;
      const gm = new Map((d.glucose || []).map((g) => [gk(g), g]));
      (inc.glucose || []).forEach((g) => gm.set(gk(g), { ...g, id: gm.get(gk(g))?.id || uid() }));
      const bm = new Map((d.bp || []).map((b) => [gk(b), b]));
      (inc.bp || []).forEach((b) => bm.set(gk(b), { ...b, id: bm.get(gk(b))?.id || uid() }));
      return { ...d, nights: [...nm.values()], episodes: [...em.values()], therapy: [...tm.values()], glucose: [...gm.values()], bp: [...bm.values()] };
    });
    setNonce((x) => x + 1);
    toast(msg || `Importate ${plural(inc.nights.length, "notte", "notti")} e ${plural(inc.episodes.length, "episodio", "episodi")}`);
  };
  const restore = (d) => {
    setData((old) => ({ ...old, nights: d.nights.map((n) => ({ ...n, id: n.id || uid() })), episodes: d.episodes.map((e) => ({ ...e, id: e.id || uid() })), therapy: (d.therapy || []).map((t) => ({ ...t, id: t.id || uid() })), glucose: (d.glucose || []).map((g) => ({ ...g, id: g.id || uid() })), bp: (d.bp || []).map((b) => ({ ...b, id: b.id || uid() })) }));
    setNonce((x) => x + 1);
    setShowSettings(false);
    toast("Backup ripristinato");
  };
  const clearAll = () => {
    setData((d) => ({ ...d, nights: [], episodes: [], therapy: [], glucose: [], bp: [] }));
    setNonce((x) => x + 1);
    setShowSettings(false);
    toast("Dati cancellati");
  };
  const doExport = () => {
    if (!data.nights.length && !data.episodes.length) { toast("Non ci sono ancora dati da esportare"); return; }
    try { downloadFile(xlsxFile(data)); toast("Excel scaricato"); }
    catch (e) { toast("Download non riuscito"); }
  };
  const toggleTheme = () => setData((d) => ({ ...d, theme: d.theme === "day" ? "night" : "day" }));

  const editEpisode = data.episodes.find((e) => e.id === editEpId) || null;

  const TABS = [
    { id: "notte", label: "Notte", Icon: Moon },
    { id: "episodio", label: "Episodi", Icon: Activity },
    { id: "terapia", label: "Terapia", Icon: Pill },
    { id: "glicemia", label: "Misure", Icon: Droplet },
    { id: "diario", label: "Diario", Icon: List },
    { id: "riepilogo", label: "Riepilogo", Icon: BarChart3 },
  ];

  return (
    <div className="app" data-theme={data.theme}>
      <style>{CSS}</style>

      <header className="topbar">
        <span className="brand">Diario dell'assistenza</span>
        <span className="row" style={{ gap: 6 }}>
          <button type="button" className="iconbtn ghost" onClick={toggleTheme}
            aria-label={data.theme === "day" ? "Passa ai colori per la notte" : "Passa ai colori per il giorno"}>
            {data.theme === "day" ? <Moon size={22} /> : <Sun size={22} />}
          </button>
          <button type="button" className="iconbtn ghost" aria-label="Dati e backup" onClick={() => setShowSettings(true)}><Settings size={22} /></button>
        </span>
      </header>

      {notice && notice.kind === "unknown" && (
        <div className="banner info">
          Il salvataggio risponde lentamente. Puoi usare l'app normalmente: al primo salvataggio controllo di nuovo e, se trovo dati salvati prima, li unisco a quelli nuovi.
          <span className="detail">Dettaglio: {notice.detail}</span>
        </div>
      )}
      {notice && notice.kind === "nostorage" && (
        <div className="banner">
          Il browser non permette di salvare in questa pagina: succede per esempio in navigazione anonima. Apri l'app da una scheda normale, altrimenti quello che registri si perde quando la chiudi.
          <span className="detail">Dettaglio: {notice.detail}</span>
        </div>
      )}
      {status === "unreadable" && notice && (
        <div className="banner">
          Risultano dati salvati, ma non riesco a leggerli. Per non sovrascriverli non salvo niente finché non si risolve.
          <span className="detail">Dettaglio: {notice.detail}</span>
          <span className="row" style={{ marginTop: 10 }}>
            <button type="button" className="btn small" onClick={init}>Riprova</button>
            <ConfirmDelete label="Ricomincia da zero" onConfirm={forceStart} />
          </span>
        </div>
      )}
      {saveErr && status === "ready" && (
        <div className="banner">
          L'ultimo salvataggio non è riuscito, quindi quello che hai inserito potrebbe perdersi alla chiusura. Scarica un backup dalle impostazioni e scrivimi il dettaglio qui sotto.
          <span className="detail">Dettaglio: {saveErr}</span>
        </div>
      )}

      <main className="main">
        {status === "loading" ? (
          <p className="empty">Apro il diario…</p>
        ) : status === "unreadable" ? null : (
          <>
            {tab === "notte" && (
              <NightForm key={`${nightDate}:${nonce}`} data={data} date={nightDate} setDate={changeNightDate}
                onSave={saveNight} onDelete={deleteNight} />
            )}
            {tab === "episodio" && (
              <EpisodeForm key={`${editEpId || "new"}:${nonce}`} data={data} episode={editEpisode}
                onSave={saveEpisode} onDelete={deleteEpisode} onCancel={() => { setEditEpId(null); setTab("diario"); }} />
            )}
            {tab === "terapia" && (
              <Terapia items={data.therapy || []} onSave={saveTherapy} onDelete={deleteTherapy} />
            )}
            {tab === "glicemia" && (
              <Misure kind={measureKind} setKind={setMeasureKind} glucose={data.glucose || []} bp={data.bp || []}
                onSaveGlu={saveGlucose} onDeleteGlu={deleteGlucose} onSaveBp={saveBp} onDeleteBp={deleteBp} />
            )}
            {tab === "diario" && (
              <Diario data={data} onSettings={() => setShowSettings(true)}
                onEditNight={(d) => { changeNightDate(d); setTab("notte"); }}
                onEditEp={(id) => { setEditEpId(id); setTab("episodio"); }} />
            )}
            {tab === "riepilogo" && (
              <Riepilogo data={data} onExport={doExport} onSettings={() => setShowSettings(true)} />
            )}
          </>
        )}
      </main>

      <nav className="tabbar" aria-label="Sezioni">
        <div className="inner">
          {TABS.map(({ id, label, Icon }) => (
            <button key={id} type="button" className="tab" aria-current={tab === id ? "page" : undefined}
              onClick={() => {
                if (id === "episodio" && tab !== "episodio") setEditEpId(null);
                if (id === "notte" && tab !== "notte") changeNightDate(defaultNightDate());
                setTab(id);
              }}>
              <Icon size={21} />
              <span className="lbl">{label}</span>
            </button>
          ))}
        </div>
      </nav>

      {showSettings && (
        <SettingsSheet data={data} onClose={() => setShowSettings(false)} onImport={(r) => { mergeIn(r); setShowSettings(false); }}
          onRestore={restore} onExport={doExport} onClear={clearAll} toast={toast} />
      )}

      {toastMsg && <div className="toast" role="status">{toastMsg}</div>}
    </div>
  );
}
