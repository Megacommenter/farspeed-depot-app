import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { storageGet, storageSet } from "./storage";

const FARSPEED_LOGO_DATA_URI = "data:image/jpeg;base64,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";
const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');`;

const FONT_DISPLAY = "'Barlow Condensed','Noto Sans TC',sans-serif";
const FONT_BODY = "'Inter','Noto Sans TC',sans-serif";
const FONT_MONO = "'IBM Plex Mono','Noto Sans TC',monospace";

const LIGHT_COLORS = {
  bg: "#F3F1EA",
  surface: "#FFFFFF",
  surfaceDim: "#E7E3D8",
  line: "#D8D3C4",
  amber: "#D98A1E",
  amberSoft: "#F4E0BC",
  amberText: "#8A5A0E",
  green: "#2D6E5C",
  greenSoft: "#DCEAE5",
  red: "#B23A2E",
  redSoft: "#F4DDD8",
  ink: "#1A2332",
  inkFaint: "#6B7280",
  navy: "#161E2E",
  navySoft: "#212B40",
  onDark: "#F3F1EA",
};

const DARK_COLORS = {
  bg: "#12161F",
  surface: "#1B2130",
  surfaceDim: "#232B3D",
  line: "#333D52",
  amber: "#E0983A",
  amberSoft: "#3A2E15",
  amberText: "#F2B45C",
  green: "#4FAE93",
  greenSoft: "#16302A",
  red: "#E8796A",
  redSoft: "#3A1B18",
  ink: "#EDEBE3",
  inkFaint: "#9AA3B5",
  navy: "#0B0F17",
  navySoft: "#1B2130",
  onDark: "#F3F1EA",
};

const CLIENTS = ["OTIS", "Schindler", "Kone", "TK Elevator", "Mitsubishi", "Fujitec", "Chevalier", "Sigma", "Lecturn", "Hitachi", "Other"];
const DEFAULT_CBM_RATES = {
  Chevalier: 110,
  Fujitec: 165,
  Kone: 162,
  Mitsubishi: 147,
  Schindler: 145,
  Sigma: 160,
  "TK Elevator": 145,
  OTIS: 160,
};
let cbmRateOverrides = {};
function setCbmRateOverridesGlobal(rates) {
  cbmRateOverrides = rates && typeof rates === "object" ? rates : {};
}
function cbmRateFor(client) {
  if (cbmRateOverrides && cbmRateOverrides[client] != null && cbmRateOverrides[client] !== "") {
    const v = Number(cbmRateOverrides[client]);
    if (v >= 0) return v;
  }
  return DEFAULT_CBM_RATES[client] != null ? DEFAULT_CBM_RATES[client] : null;
}

// --- Storage billing: monthly $/CBM rate. First partial month is pro-rated by day;
// every calendar month after that is billed in full regardless of when mid-month
// the goods leave. Billing starts the day after the free-storage period ends. ---
function monthIndexOf(d) { return d.getFullYear() * 12 + d.getMonth(); }
function lastDayOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0); }
function daysInMonthOf(d) { return lastDayOfMonth(d).getDate(); }
function toDateOnly(s) { return new Date(`${s}T00:00:00`); }
function computeStorageCharge(arrivalDateStr, endDateStr, freeDays, ratePerCbmMonth, cbm) {
  if (!arrivalDateStr || !(ratePerCbmMonth > 0) || !(cbm > 0)) return null;
  const arrival = toDateOnly(arrivalDateStr);
  const billStart = new Date(arrival);
  billStart.setDate(billStart.getDate() + (freeDays || 0));
  const end = toDateOnly(endDateStr || todayStr());
  const breakdown = [];
  if (end < billStart) return { billStart: dateToLocalISO(billStart), total: 0, breakdown };
  const firstMonthEnd = lastDayOfMonth(billStart);
  const daysInFirstMonth = daysInMonthOf(billStart);
  let total = 0;
  if (end <= firstMonthEnd) {
    const daysUsed = Math.round((end - billStart) / 86400000) + 1;
    const amt = ratePerCbmMonth * cbm * (daysUsed / daysInFirstMonth);
    breakdown.push({ label: `${fmt(dateToLocalISO(billStart))} \u2013 ${fmt(dateToLocalISO(end))}`, detail: `pro-rata, ${daysUsed}/${daysInFirstMonth} days`, amount: amt, year: billStart.getFullYear(), month: billStart.getMonth() });
    total = amt;
  } else {
    const daysUsed = Math.round((firstMonthEnd - billStart) / 86400000) + 1;
    const amt1 = ratePerCbmMonth * cbm * (daysUsed / daysInFirstMonth);
    breakdown.push({ label: `${fmt(dateToLocalISO(billStart))} \u2013 ${fmt(dateToLocalISO(firstMonthEnd))}`, detail: `pro-rata, ${daysUsed}/${daysInFirstMonth} days`, amount: amt1, year: billStart.getFullYear(), month: billStart.getMonth() });
    total += amt1;
    const numFullMonths = monthIndexOf(end) - monthIndexOf(billStart);
    for (let i = 1; i <= numFullMonths; i++) {
      const mDate = new Date(billStart.getFullYear(), billStart.getMonth() + i, 1);
      const isLast = i === numFullMonths;
      const label = mDate.toLocaleString("en-US", { month: "short", year: "numeric" });
      breakdown.push({ label, detail: isLast && end < lastDayOfMonth(mDate) ? `full month \u2014 left ${fmt(dateToLocalISO(end))}` : "full month", amount: ratePerCbmMonth * cbm, year: mDate.getFullYear(), month: mDate.getMonth() });
      total += ratePerCbmMonth * cbm;
    }
  }
  return { billStart: dateToLocalISO(billStart), total: Math.round(total * 100) / 100, breakdown };
}
// Builds one billing row per arrival batch (or per whole item if no split arrivals were
// used), further split by delivery date where a batch's cases left storage on different
// dates. Rows are not grouped or separated by construction site.
function computeItemBillingRows(item) {
  const rate = cbmRateFor(item.client);
  if (!rate) return [];
  const freeDays = freeDaysFor(item);
  const rows = [];
  if (usesArrivalBatches(item)) {
    const deliveredAt = {};
    activeDeliveries(item).forEach((d) => (d.codes || []).forEach((c) => { deliveredAt[c] = d.date; }));
    activeArrivals(item).forEach((batch) => {
      if (!batch.date) return;
      const byEnd = new Map();
      (batch.codes || []).forEach((code) => {
        const pkg = (item.packages || []).find((p) => p.code === code);
        const cbm = pkg ? Number(pkg.cbm) || 0 : 0;
        const end = deliveredAt[code] || null;
        const key = end || "__ongoing__";
        if (!byEnd.has(key)) byEnd.set(key, { end, cbm: 0, codes: [] });
        const g = byEnd.get(key);
        g.cbm += cbm;
        g.codes.push(code);
      });
      for (const g of byEnd.values()) {
        if (!(g.cbm > 0)) continue;
        const calc = computeStorageCharge(batch.date, g.end, freeDays, rate, g.cbm);
        if (calc) rows.push({ item, rate, freeDays, batchDate: batch.date, batchType: batch.type, codes: g.codes, cbm: g.cbm, endDate: g.end, ongoing: !g.end, ...calc });
      }
    });
  } else {
    const arrivalDate = effectiveDepotArrivalDate(item);
    if (!arrivalDate) return [];
    const cbmTotal = (item.packages && item.packages.length) ? item.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0) : Number(item.volumeCbm) || 0;
    if (!(cbmTotal > 0)) return [];
    const status = deriveStatus(item);
    if (status === "delivered") {
      const end = lastDeliveryDate(item);
      const calc = computeStorageCharge(arrivalDate, end, freeDays, rate, cbmTotal);
      if (calc) rows.push({ item, rate, freeDays, batchDate: arrivalDate, cbm: cbmTotal, endDate: end, ongoing: false, ...calc });
    } else if (status === "partial") {
      const delivered = new Set(deliveredCodes(item));
      const pkgs = item.packages || [];
      const haveCbmData = pkgs.length > 0 && pkgs.every((p) => p.cbm !== "" && p.cbm != null && !isNaN(Number(p.cbm)));
      let deliveredCbm, remainingCbm, exact;
      if (haveCbmData) {
        deliveredCbm = pkgs.filter((p) => delivered.has(p.code)).reduce((s, p) => s + Number(p.cbm), 0);
        remainingCbm = cbmTotal - deliveredCbm;
        exact = true;
      } else {
        const deliveredShare = deliveredUnits(item) / (totalUnits(item) || 1);
        deliveredCbm = cbmTotal * deliveredShare;
        remainingCbm = cbmTotal - deliveredCbm;
        exact = false;
      }
      const lastDelEnd = lastDeliveryDate(item);
      if (deliveredCbm > 0) {
        const c = computeStorageCharge(arrivalDate, lastDelEnd, freeDays, rate, deliveredCbm);
        if (c) rows.push({ item, rate, freeDays, batchDate: arrivalDate, cbm: deliveredCbm, endDate: lastDelEnd, ongoing: false, estimated: !exact, ...c });
      }
      if (remainingCbm > 0) {
        const c = computeStorageCharge(arrivalDate, null, freeDays, rate, remainingCbm);
        if (c) rows.push({ item, rate, freeDays, batchDate: arrivalDate, cbm: remainingCbm, endDate: null, ongoing: true, estimated: !exact, ...c });
      }
    } else if (status === "at_depot") {
      const c = computeStorageCharge(arrivalDate, null, freeDays, rate, cbmTotal);
      if (c) rows.push({ item, rate, freeDays, batchDate: arrivalDate, cbm: cbmTotal, endDate: null, ongoing: true, ...c });
    }
  }
  return rows;
}
// Aggregates every row's breakdown lines into the given calendar month/year (month is
// 0-indexed, JS Date convention), grouped by client - this is what should match what
// MYOB invoices for that month, for cross-checking.
function computeMonthlyBillingSummary(items, year, month) {
  const byClient = new Map();
  let grandTotal = 0;
  for (const item of items) {
    for (const row of computeItemBillingRows(item)) {
      for (const line of row.breakdown) {
        if (line.year !== year || line.month !== month) continue;
        const amt = Math.round(line.amount * 100) / 100;
        if (!byClient.has(item.client)) byClient.set(item.client, { client: item.client, total: 0, lines: [] });
        const g = byClient.get(item.client);
        g.total += amt;
        g.lines.push({
          project: item.constructionSite || item.project || "",
          jobNumber: item.jobNumber || "",
          batchDate: row.batchDate,
          cbm: row.cbm,
          detail: line.detail,
          label: line.label,
          amount: amt,
          estimated: !!row.estimated,
        });
        grandTotal += amt;
      }
    }
  }
  const clients = [...byClient.values()].map((g) => ({ ...g, total: Math.round(g.total * 100) / 100 })).sort((a, b) => b.total - a.total);
  return { clients, grandTotal: Math.round(grandTotal * 100) / 100 };
}
const ITEM_TYPES = ["Container", "Separate Items"];
const DEPOTS = ["Farspeed Depot 1", "Farspeed Depot 3"];
const DEPOT_LABELS_ZH = {
  "Farspeed Depot 1": "快達一號倉",
  "Farspeed Depot 3": "快達三號倉",
};
function depotLabel(value, lang) {
  if (lang === "zh" && DEPOT_LABELS_ZH[value]) return DEPOT_LABELS_ZH[value];
  return value;
}
function depotDisplay(value, lang) {
  if (!value) return "—";
  if (lang === "zh" && DEPOT_LABELS_ZH[value]) return DEPOT_LABELS_ZH[value];
  return value.replace("Farspeed ", "");
}
const ARRIVING_TYPES = ["Devan", "CFS"];
const DEFAULT_ROLES = ["Account Officer", "CEO / Business Manager", "Warehouse Depot Head", "Admin"];
const DEFAULT_EMPLOYEES = [
  { id: "E1", name: "Irene Lee", role: "Account Officer" },
  { id: "E2", name: "Nana Chan", role: "Account Officer" },
  { id: "E3", name: "Polly Lee", role: "Account Officer" },
  { id: "E4", name: "Cheng Wai Kee", role: "CEO / Business Manager" },
  { id: "E5", name: "Bhatt Wai Lee", role: "Warehouse Depot Head" },
  { id: "E6", name: "Mega Chan", role: "Admin" },
  { id: "E7", name: "Simon Chan", role: "Staff" },
];
// Login accounts - names match the Employees directory so "Recorded By" and the logged-in
// identity stay consistent. Default password for every account is "Farspeed"; each person
// can change their own password after logging in.
const LOGIN_ACCOUNTS = ["Irene Lee", "Nana Chan", "Cheng Wai Kee", "Mega Chan", "Polly Lee", "Simon Chan", "Bhatt Wai Lee"];
const DEFAULT_PASSWORD = "Farspeed";
async function hashPassword(pw) {
  const enc = new TextEncoder().encode(`farspeed-depot-app:${pw}`);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
const DEFAULT_DIRECTORY = [
  {
    id: "SITE1",
    siteEn: "Kwu Tung North Area 19 Phase 1A & 1B",
    siteZh: "古洞北19區1A & 1B",
    client: "TK Elevator",
    jobRef: "KTN-002",
    orderedBy: "Alex Tam",
    accountOfficer: "Irene Lee",
  },
];
const FREE_DAYS = 14;
const CLIENT_FREE_DAYS = { Schindler: 21 };
let freeStorageRules = [];
function setFreeStorageRulesGlobal(rules) {
  freeStorageRules = Array.isArray(rules) ? rules : [];
}
function freeDaysFor(item) {
  const proj = String(item.project || "").trim().toLowerCase();
  const site = String(item.constructionSite || "").trim().toLowerCase();
  for (const r of freeStorageRules) {
    const key = String(r.project || "").trim().toLowerCase();
    if (!key) continue;
    if ((proj && proj.includes(key)) || (site && site.includes(key))) {
      const d = Number(r.days);
      if (d > 0) return d;
    }
  }
  return CLIENT_FREE_DAYS[item.client] || FREE_DAYS;
}

const FIELD_DEFS = [
  { key: "client", label: "Client", aliases: ["client", "customer"] },
  { key: "project", label: "Project", aliases: ["project", "project name"] },
  { key: "invoiceNumber", label: "Invoice No.", aliases: ["invoice no", "invoice number", "invoice"] },
  { key: "constructionSite", label: "Construction Site Name", aliases: ["construction site", "site", "site name"] },
  { key: "itemType", label: "Item Type", aliases: ["item type", "type"] },
  { key: "packageCount", label: "# of Packages", aliases: ["packages", "no. of packages", "package count", "# of packages"] },
  { key: "unitCode", label: "Escalator/Elevator # or Code", aliases: ["unit code", "escalator/elevator no", "unit no", "code", "unit"] },
  { key: "description", label: "Description", aliases: ["description", "desc"] },
  { key: "shkNumber", label: "Reference / SHK No.", aliases: ["shk no", "shk number", "reference", "reference/shk no"] },
  { key: "weightKg", label: "Weight (KG)", aliases: ["weight", "weight (kg)", "kg"] },
  { key: "volumeCbm", label: "Volume (CBM)", aliases: ["volume", "cbm", "volume (cbm)"] },
  { key: "containers20", label: "No. of 20' Containers", aliases: ["20' containers", "containers 20", "no. of 20' containers", "20ft"] },
  { key: "containers40", label: "No. of 40' Containers", aliases: ["40' containers", "containers 40", "no. of 40' containers", "40ft"] },
  { key: "arrivingType", label: "Arriving Type", aliases: ["arriving type", "arrival type"] },
  { key: "depot", label: "Depot", aliases: ["depot"] },
  { key: "depotLocation", label: "Depot Location / Bay", aliases: ["depot location", "bay", "location"] },
  { key: "terminalArrivalDate", label: "Terminal Arrival Date", aliases: ["terminal arrival date", "terminal arrival"] },
  { key: "terminalLFD", label: "Terminal Last Free Day", aliases: ["terminal lfd", "last free day", "lfd"] },
  { key: "confirmedCollectionDate", label: "Confirmed Collection Date", aliases: ["confirmed collection date", "collection date"] },
  { key: "depotArrivalDate", label: "Depot Arrival Date", aliases: ["depot arrival date", "depot arrival"] },
  { key: "plannedDeliveryDate", label: "Planned Delivery Date", aliases: ["planned delivery date", "delivery date", "estimated delivery"] },
  { key: "jobNumber", label: "Job No.", aliases: ["job no", "job no.", "job number"] },
  { key: "orderedBy", label: "Ordered By", aliases: ["ordered by"] },
  { key: "poNumber", label: "P.O. No.", aliases: ["p.o. no", "po no", "purchase order no"] },
  { key: "jobRef", label: "Job Ref. (Site Code)", aliases: ["job ref", "job ref.", "site code"] },
  { key: "notes", label: "Notes", aliases: ["notes", "remarks"] },
];

function normalizeHeader(h) {
  return String(h || "").toLowerCase().trim().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}
function matchField(header) {
  const norm = normalizeHeader(header);
  const hit = FIELD_DEFS.find((f) => f.aliases.includes(norm) || normalizeHeader(f.label) === norm);
  return hit ? hit.key : null;
}

// --- Client packing-list parser ---------------------------------------
// Different clients (TK Elevator, Schindler, OTIS, etc.) all send packing
// lists with completely different layouts. Rather than expecting a fixed
// header row, this scans for the row that looks most like a table header
// (by keyword match) and maps its columns by meaning, then groups the
// case/package rows underneath by whichever "lot" column is present
// (e.g. a unit code like L0MO-027837.006, or a lift name like L2/L3, or a
// lift no. like "L8 Batch 1") so each lot becomes its own manifest entry.
const PL_HEADER_ALIASES = {
  containerNo: ["container no.", "container no", "container"],
  caseNo: ["case no.", "case no", "case\nno", "case", "cases discript", "cases     discript", "pkg#", "pkg #", "pkg no.", "pkg no", "package no."],
  qty: ["qty", "quantity", "qyt = quantity", "qyt"],
  lot: ["project no.", "project no", "lift name", "lift no.", "lift no", "lift", "sap no.", "sap no", "sap"],
  orderNo: ["omc sales order no.", "omc sales order no", "sales order no.", "sales order no", "order no."],
  description: ["description", "material description"],
  grossWeight: ["g.weight", "gross weight", "gross", "actual   weight", "actual weight", "g.w./kg", "g.w.", "g.w"],
  netWeight: ["n.weight", "net weight", "net", "estimated  weight", "estimated weight", "n.w./kg", "n.w.", "n.w"],
  cbm: ["cbm", "volume(m3)", "volume (m3)", "volume"],
  dimension: ["dimension", "dimension (mm)", "dimensions", "size"],
  dimensionCm: ["dim(cm)", "dim (cm)", "dimension(cm)", "dimension (cm)", "dimensions(cm)", "dimensions (cm)"],
};
const PL_CLIENT_HINTS = [
  ["otis", "OTIS"], ["schindler", "Schindler"], ["kone", "Kone"], ["tk elevator", "TK Elevator"],
  ["mitsubishi", "Mitsubishi"], ["fujitec", "Fujitec"], ["chevalier", "Chevalier"], ["sigma", "Sigma"],
  ["lecturn", "Lecturn"], ["hitachi", "Hitachi"],
];

function plNorm(v) {
  return String(v || "").toLowerCase().replace(/\n/g, " ").replace(/\s+/g, " ").trim();
}
function plNum(v) {
  if (v === "" || v === null || v === undefined) return 0;
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  const cleaned = String(v).replace(/,/g, "").trim();
  const n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}
function plCbmFromDimension(v, unit) {
  // Parses strings like "2000*850*600" (mm) or "140*90*96" (cm) into m3.
  const parts = String(v || "").split(/[x*×]/i).map((s) => Number(s.replace(/,/g, "").trim()));
  if (parts.length !== 3 || parts.some((n) => !n || isNaN(n))) return 0;
  const divisor = unit === "cm" ? 1e6 : 1e9;
  return (parts[0] * parts[1] * parts[2]) / divisor;
}
function plScoreRow(row) {
  let score = 0;
  row.forEach((cell) => {
    const n = plNorm(cell);
    if (!n) return;
    for (const aliases of Object.values(PL_HEADER_ALIASES)) {
      if (aliases.some((a) => n === a || n.includes(a))) { score++; break; }
    }
  });
  return score;
}
function plDetectHeaderRow(rows) {
  let bestIdx = -1, bestScore = 0;
  rows.slice(0, 40).forEach((row, idx) => {
    const s = plScoreRow(row);
    if (s > bestScore) { bestScore = s; bestIdx = idx; }
  });
  return bestScore >= 3 ? bestIdx : -1;
}
function plMapColumns(headerRow) {
  const colMap = {};
  headerRow.forEach((cell, idx) => {
    const n = plNorm(cell);
    if (!n) return;
    for (const [field, aliases] of Object.entries(PL_HEADER_ALIASES)) {
      if (colMap[field] !== undefined) continue;
      if (aliases.some((a) => n === a || n.includes(a))) { colMap[field] = idx; break; }
    }
  });
  return colMap;
}
// Manufacturer packing lists (like TK Elevator's) often list a "Marks:" block where the
// first line is the project/building name (sometimes with unit codes and city tacked on,
// e.g. "1881 Heritage E3&4,HongKong") followed by lines pairing a part/SAP number with its
// unit code (e.g. "410217-501-003/E3"). This pulls out both: the clean project name, and a
// lookup from part number to unit code so case groups can be labeled "E3"/"E4" instead of
// showing the raw part number.
function plGuessMarksBlock(rows) {
  const legend = {};
  let rawProjectLine = "";
  for (let r = 0; r < Math.min(rows.length, 25); r++) {
    const row = rows[r];
    for (let i = 0; i < row.length; i++) {
      const n = plNorm(row[i]);
      if (n !== "marks" && n !== "marks:" && n !== "唛頭" && n !== "唛头" && n !== "shipping marks") continue;
      let valueCol = -1;
      for (let j = i + 1; j < row.length; j++) {
        if (row[j] !== "" && row[j] != null) { rawProjectLine = String(row[j]).trim(); valueCol = j; break; }
      }
      if (valueCol === -1 && rows[r + 1] && row[i] !== undefined) {
        if (rows[r + 1][i] !== "" && rows[r + 1][i] != null) { rawProjectLine = String(rows[r + 1][i]).trim(); valueCol = i; r = r + 1; }
      }
      if (valueCol === -1) continue;
      for (let rr = r + 1; rr < Math.min(r + 6, rows.length); rr++) {
        const cell = rows[rr] ? rows[rr][valueCol] : undefined;
        if (cell === "" || cell == null) break;
        const text = String(cell).trim();
        const m = text.match(/^([\w\-]+)\s*\/\s*([A-Za-z0-9]+)$/);
        if (!m) break;
        legend[m[1].trim()] = m[2].trim();
      }
      return { rawProjectLine, legend };
    }
  }
  return { rawProjectLine, legend };
}
function plCleanProjectName(rawLine) {
  if (!rawLine) return "";
  // Drop a trailing unit-code list and city (e.g. "1881 Heritage E3&4,HongKong" -> "1881 Heritage")
  const beforeComma = rawLine.split(",")[0];
  const stripped = beforeComma.replace(/\s*[A-Za-z]\d+(\s*&\s*[A-Za-z]?\d+)*\s*$/i, "").trim();
  return stripped || beforeComma.trim();
}
function plGuessClient(rows) {
  for (const row of rows.slice(0, 25)) {
    for (const cell of row) {
      const n = plNorm(cell);
      if (!n) continue;
      for (const [needle, label] of PL_CLIENT_HINTS) {
        if (n.includes(needle)) return label;
      }
    }
  }
  return null;
}
function plGuessProject(rows) {
  for (const row of rows.slice(0, 25)) {
    for (let i = 0; i < row.length; i++) {
      const n = plNorm(row[i]);
      if (n === "project" || n === "project:" || n === "project name" || n === "project name:") {
        for (let j = i + 1; j < row.length; j++) {
          if (row[j] !== "" && row[j] !== null && row[j] !== undefined) return String(row[j]).trim();
        }
      }
    }
  }
  // Fallback: manufacturer packing lists often have no literal "Project" label at all -
  // pull the cleaned building name out of the Marks block instead.
  const { rawProjectLine } = plGuessMarksBlock(rows);
  return plCleanProjectName(rawProjectLine);
}
function parsePackingListSheet(rows, legend) {
  const headerIdx = plDetectHeaderRow(rows);
  if (headerIdx === -1) return null;
  const colMap = plMapColumns(rows[headerIdx]);
  if (colMap.lot === undefined && colMap.caseNo === undefined) return null;

  // Translates a raw lot value (which may be a manufacturer part/SAP number, or several
  // comma-separated ones for cases shared across units) into friendly unit codes via the
  // legend pulled from the Marks block, e.g. "410217-501-003,502-004" -> "E3, E4". The
  // second part often omits the shared numeric prefix, so that's restored first.
  function expandAbbreviatedParts(raw) {
    const parts = raw.split(",").map((p) => p.trim()).filter(Boolean);
    if (parts.length <= 1) return parts;
    const firstSegs = parts[0].split("-");
    const prefix = firstSegs.length > 2 ? firstSegs.slice(0, firstSegs.length - 2).join("-") : null;
    return parts.map((p, idx) => {
      if (idx === 0) return p;
      const segs = p.split("-");
      if (prefix && segs.length < firstSegs.length) return `${prefix}-${p}`;
      return p;
    });
  }
  function translateLot(raw) {
    if (!legend || !Object.keys(legend).length) return raw;
    const parts = expandAbbreviatedParts(raw);
    const translated = parts.map((p) => legend[p] || p);
    const allTranslated = parts.every((p) => legend[p]);
    return allTranslated ? [...new Set(translated)].join(", ") : raw;
  }

  const groups = {};
  const order = [];
  let lastLot = "", lastContainer = "", lastCase = "", lastOrderNo = "";
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (row.every((c) => String(c || "").trim() === "")) continue;
    const rowText = row.map((c) => String(c || "").toLowerCase()).join(" ");
    if (rowText.includes("total")) break;

    const lot = colMap.lot !== undefined ? String(row[colMap.lot] || "").trim() : "";
    const container = colMap.containerNo !== undefined ? String(row[colMap.containerNo] || "").trim() : "";
    const caseNo = colMap.caseNo !== undefined ? String(row[colMap.caseNo] || "").trim() : "";
    const orderNo = colMap.orderNo !== undefined ? String(row[colMap.orderNo] || "").trim() : "";
    const description = colMap.description !== undefined ? String(row[colMap.description] || "").trim() : "";
    const grossVal = colMap.grossWeight !== undefined && row[colMap.grossWeight] !== "" && row[colMap.grossWeight] != null ? plNum(row[colMap.grossWeight]) : null;
    const netVal = colMap.netWeight !== undefined && row[colMap.netWeight] !== "" && row[colMap.netWeight] != null ? plNum(row[colMap.netWeight]) : null;
    // Use whichever weight is bigger - usually gross, but this doesn't assume it.
    const weight = grossVal != null && netVal != null ? Math.max(grossVal, netVal) : (grossVal != null ? grossVal : (netVal != null ? netVal : 0));
    let cbm = 0;
    if (colMap.dimensionCm !== undefined && row[colMap.dimensionCm]) cbm = plCbmFromDimension(row[colMap.dimensionCm], "cm");
    if (!cbm && colMap.dimension !== undefined && row[colMap.dimension]) cbm = plCbmFromDimension(row[colMap.dimension]);
    if (!cbm && colMap.cbm !== undefined && row[colMap.cbm] !== "" && row[colMap.cbm] != null) cbm = plNum(row[colMap.cbm]);

    // Many bilingual packing lists have a second header row directly below the first,
    // repeating the same column labels in Chinese (e.g. "项目号/行号/梯号" under "SAP NO.").
    // That row has no real weight or CBM data, so skip it rather than counting it as a case.
    if (i === headerIdx + 1 && !weight && !cbm) continue;

    if (lot) lastLot = lot;
    if (container) lastContainer = container;
    if (caseNo) lastCase = caseNo;
    if (orderNo) lastOrderNo = orderNo;
    // A description usually signals a real data row, but some packing lists (like ones
    // that only list PKG#/dimensions/weight) have no description column at all - a case
    // number paired with real weight or CBM is equally good evidence of a genuine row.
    if (!description && !(caseNo && (weight || cbm))) continue;

    const key = translateLot(lastLot) || "UNSPECIFIED";
    if (!groups[key]) { groups[key] = { lot: key, packages: [], containers: new Set(), totalWeight: 0, totalCbm: 0 }; order.push(key); }
    groups[key].packages.push({ code: lastCase || String(groups[key].packages.length + 1), orderNo: lastOrderNo, description, weightKg: weight ? String(weight) : "", cbm: cbm ? String(cbm) : "" });
    if (lastContainer) groups[key].containers.add(lastContainer);
    groups[key].totalWeight += weight;
    groups[key].totalCbm += cbm;
  }
  // A single lot can combine multiple separate order numbers, each restarting its own case
  // numbering (e.g. two different orders both having a case "1/1") - disambiguate any case
  // code that collides within the same lot so every package has a genuinely unique code.
  for (const key of order) {
    const pkgs = groups[key].packages;
    const counts = {};
    pkgs.forEach((p) => { counts[p.code] = (counts[p.code] || 0) + 1; });
    const seen = {};
    for (const p of pkgs) {
      if (counts[p.code] > 1 && p.orderNo) {
        seen[p.code] = (seen[p.code] || 0) + 1;
        p.code = `${p.orderNo}/${p.code}`;
      }
    }
  }
  return { groups: order.map((k) => ({ ...groups[k], containers: [...groups[k].containers] })), hasLotColumn: colMap.lot !== undefined };
}
function parsePackingListWorkbook(workbook) {
  let bestGroups = null;
  let bestHasLotColumn = false;
  let client = null;
  let project = "";
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    if (!client) client = plGuessClient(rows);
    if (!project) project = plGuessProject(rows);
    const { legend } = plGuessMarksBlock(rows);
    const result = parsePackingListSheet(rows, legend);
    if (!result || !result.groups || result.groups.length === 0) continue;
    const { groups, hasLotColumn } = result;
    // A sheet that actually identifies lift/lot numbers is preferred over one that had to
    // lump everything into a single UNSPECIFIED group, even if the latter has more raw
    // rows (that's often a material/component breakdown sheet, not the case-level one).
    const better = !bestGroups || (hasLotColumn && !bestHasLotColumn) || (hasLotColumn === bestHasLotColumn && groups.length > bestGroups.length);
    if (better) { bestGroups = groups; bestHasLotColumn = hasLotColumn; }
  }
  return { groups: bestGroups, client, project };
}

function sigPart(v) {
  return String(v || "").trim().toLowerCase();
}
function itemSignature(item) {
  const projectKey = item.directoryId ? `dir:${item.directoryId}` : sigPart(item.project);
  if (!projectKey) return null;
  return [sigPart(item.client), projectKey, sigPart(item.unitCode), sigPart(item.itemType), sigPart(item.depotArrivalDate), sigPart(item.invoiceNumber)].join("|");
}
function findDuplicateGroups(items) {
  const map = {};
  items.forEach((it) => {
    const sig = itemSignature(it);
    if (!sig) return;
    (map[sig] = map[sig] || []).push(it);
  });
  return Object.values(map).filter((g) => g.length > 1);
}

function todayStr() {
  return dateToLocalISO(new Date());
}
function daysBetween(a, b) {
  if (!a || !b) return null;
  const d1 = new Date(a + "T00:00:00");
  const d2 = new Date(b + "T00:00:00");
  return Math.round((d2 - d1) / 86400000);
}
function addDays(dateStr, n) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() + n);
  return dateToLocalISO(d);
}
function fmt(dateStr) {
  if (!dateStr) return "—";
  const d = new Date(dateStr + "T00:00:00");
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function totalUnits(item) {
  if (item.packages && item.packages.length > 0) return item.packages.length;
  const n = Number(item.packageCount);
  return n > 0 ? n : 1;
}
function activeDeliveries(item) {
  return (item.deliveries || []).filter((d) => !d.cancelled);
}
function deliveredUnits(item) {
  if (item.packages && item.packages.length > 0) {
    return activeDeliveries(item).reduce((s, d) => s + (d.codes ? d.codes.length : 0), 0);
  }
  return activeDeliveries(item).reduce((s, d) => s + (Number(d.packageCount) || 0), 0);
}
function remainingUnits(item) {
  return Math.max(0, totalUnits(item) - deliveredUnits(item));
}
function deliveredCodes(item) {
  return activeDeliveries(item).flatMap((d) => d.codes || []);
}
function remainingPackages(item) {
  if (!item.packages || item.packages.length === 0) return [];
  const done = new Set(deliveredCodes(item));
  return item.packages.filter((p) => !done.has(p.code));
}
function activeArrivals(item) {
  return item.arrivals || [];
}
function arrivedCodesSet(item) {
  return new Set(activeArrivals(item).flatMap((a) => a.codes || []));
}
function usesArrivalBatches(item) {
  return activeArrivals(item).length > 0 && (item.packages || []).length > 0;
}
function notYetArrivedPackages(item) {
  if (!usesArrivalBatches(item)) return [];
  const done = arrivedCodesSet(item);
  return item.packages.filter((p) => !done.has(p.code));
}
function earliestArrivalDate(item) {
  const ds = activeArrivals(item).map((a) => a.date).filter(Boolean).sort();
  return ds.length ? ds[0] : "";
}
function effectiveDepotArrivalDate(item) {
  return item.depotArrivalDate || earliestArrivalDate(item);
}
function lastDeliveryDate(item) {
  const ds = activeDeliveries(item).map((d) => d.date).filter(Boolean).sort();
  return ds.length ? ds[ds.length - 1] : null;
}
function remainingShare(item) {
  const total = totalUnits(item) || 1;
  return remainingUnits(item) / total;
}
function remainingWeightKg(item) {
  if (!item.weightKg) return 0;
  return Number(item.weightKg) * remainingShare(item);
}
function remainingVolumeCbm(item) {
  if (!item.volumeCbm) return 0;
  return Number(item.volumeCbm) * remainingShare(item);
}
function depotRemainingTotals(items, depotValue) {
  let kg = 0, cbm = 0, count = 0;
  items.forEach((it) => {
    if (it.depot !== depotValue) return;
    const status = deriveStatus(it);
    if (status !== "at_depot" && status !== "partial") return;
    kg += remainingWeightKg(it);
    cbm += remainingVolumeCbm(it);
    count += 1;
  });
  return { kg: Math.round(kg * 10) / 10, cbm: Math.round(cbm * 1000) / 1000, count };
}

function deriveStatus(item) {
  if (!effectiveDepotArrivalDate(item)) return "pending_collection";
  const hasDeliveries = activeDeliveries(item).length > 0;
  const remaining = remainingUnits(item);
  if (hasDeliveries && remaining <= 0) return "delivered";
  if (hasDeliveries && remaining > 0) return "partial";
  return "at_depot";
}

function lfdAlert(item) {
  if (deriveStatus(item) !== "pending_collection" || !item.terminalLFD) return null;
  const d = daysBetween(todayStr(), item.terminalLFD);
  if (d === null) return null;
  if (d < 0) return { level: "overdue", days: d };
  if (d <= 3) return { level: "soon", days: d };
  return null;
}

function storageInfo(item) {
  const status = deriveStatus(item);
  if (status === "pending_collection") return null;
  const freeDays = freeDaysFor(item);
  const today = todayStr();

  if (usesArrivalBatches(item)) {
    // Each batch's cases run their own free-storage clock from that batch's devan/CFS date.
    const deliveredAt = {};
    activeDeliveries(item).forEach((d) => (d.codes || []).forEach((c) => { deliveredAt[c] = d.date; }));
    let maxBillable = 0;
    let maxHeld = 0;
    let minLeft = Infinity;
    let anyHolding = false;
    activeArrivals(item).forEach((a) => {
      if (!a.date) return;
      (a.codes || []).forEach((code) => {
        const end = deliveredAt[code] || today;
        const held = daysBetween(a.date, end) || 0;
        const over = held - freeDays;
        if (over > 0) maxBillable = Math.max(maxBillable, over);
        if (held > maxHeld) maxHeld = held;
        if (!deliveredAt[code]) {
          anyHolding = true;
          const left = freeDays - held;
          if (left < minLeft) minLeft = left;
        }
      });
    });
    const freeUntil = addDays(earliestArrivalDate(item), freeDays);
    return {
      freeUntil,
      daysHeld: maxHeld,
      billableDays: maxBillable,
      billable: maxBillable > 0,
      freeDays,
      daysLeft: anyHolding && minLeft !== Infinity ? minLeft : null,
    };
  }

  const arrivalDate = effectiveDepotArrivalDate(item);
  const freeUntil = addDays(arrivalDate, freeDays);
  const endDate = status === "delivered" ? lastDeliveryDate(item) || today : today;
  const daysHeld = daysBetween(arrivalDate, endDate);
  const overFree = daysBetween(freeUntil, endDate);
  const billableDays = overFree && overFree > 0 ? overFree : 0;
  return { freeUntil, daysHeld, billableDays, billable: billableDays > 0, freeDays, daysLeft: freeDays - daysHeld };
}

function currentYyMm() {
  const now = new Date();
  return String(now.getFullYear()).slice(-2) + String(now.getMonth() + 1).padStart(2, "0");
}
function allJobNumbersUsed(items) {
  const nums = [];
  (items || []).forEach((it) => {
    if (it.jobNumber) nums.push(it.jobNumber);
    (it.deliveries || []).forEach((d) => {
      if (d.jobNumber) nums.push(d.jobNumber);
    });
  });
  return nums;
}
function nextJobNumber(items) {
  const prefix = currentYyMm();
  let maxSeq = 0;
  allJobNumbersUsed(items).forEach((jn) => {
    if (jn.startsWith(prefix) && jn.length === 7) {
      const seq = Number(jn.slice(4));
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  return `${prefix}${String(maxSeq + 1).padStart(3, "0")}`;
}

function emptyForm() {
  return {
    client: CLIENTS[0],
    project: "",
    invoiceNumber: "",
    itemType: ITEM_TYPES[0],
    packageCount: "",
    unitCode: "",
    constructionSite: "",
    description: "",
    weightKg: "",
    volumeCbm: "",
    shkNumber: "",
    ssDoNo: "",
    containers20: "",
    containers40: "",
    arrivingType: ARRIVING_TYPES[0],
    terminalArrivalDate: "",
    terminalLFD: "",
    confirmedCollectionDate: "",
    depot: DEPOTS[0],
    depotArrivalDate: "",
    depotLocation: "",
    plannedDeliveryDate: "",
    deliveries: [],
    arrivals: [],
    packages: [],
    jobNumber: "",
    orderedBy: "",
    poNumber: "",
    jobRef: "",
    directoryId: "",
    recordedBy: "",
    notes: "",
  };
}
const TEXT = {
  en: {
    appSubtitle: "Depot & Storage Manifest",
    navDashboard: "Dashboard",
    navInventory: "Inventory",
    navNewEntry: "New Entry",
    navDeliveries: "Deliveries",
    navDuplicates: "Duplicates",
    navDuplicatesCount: (n) => `Duplicates (${n})`,
    navImport: "Import",
    themeToggleLabel: "Toggle dark mode",
    langToggleLabel: "中文",
    loadingMsg: "Loading depot records…",
    saveErrorMsg: "Could not save — please try again.",

    statAtDepot: "At Depot",
    statPending: "Pending Collection",
    statBillable: "Billable Now",
    statLfd: "LFD Warnings",
    dupBanner: (n) => `${n} possible duplicate ${n === 1 ? "entry" : "entries"} found — same client, project, unit and depot arrival date.`,
    reviewDuplicatesBtn: "Review Duplicates",
    lfdSectionTitle: "Terminal Last Free Day — Action Needed",
    billableSectionTitle: "Currently Billable Storage",
    billableEmptyMsg: "Nothing past the 14-day free period right now.",
    sinceLabel: "since",

    searchLabel: "Search",
    searchPlaceholder: "Project, client, ID, SHK no., case no., job no.",
    clientLabel: "Client",
    statusLabel: "Status",
    statusAll: "All",
    statusPending: "Pending Collection",
    statusAtDepot: "At Depot",
    statusPartial: "Partially Delivered",
    statusDelivered: "Delivered",
    depotLabel: "Depot",
    exportBtn: (n) => `Export to Excel (${n})`,
    newEntryBtn: "+ New Entry",
    colId: "ID",
    colClient: "Client",
    colProjectSite: "Project / Site",
    colUnit: "Unit",
    deliverySearchPlaceholder: "Client, project, unit no., job no., or case no.",
    colDepot: "Depot",
    colDepotArrival: "Depot Arrival",
    colStatus: "Status",
    noRecordsMsg: "No records match. Add a new entry to get started.",
    editBtn: "Edit",
    deleteBtn: "Delete",
    deliverBtn: "Deliver",
    duplicateBadge: "DUPLICATE?",

    titleNew: "New Manifest Entry",
    titleEdit: "Edit Manifest Entry",
    fClient: "Client",
    fProject: "Project",
    fProjectEn: "Project / Site (English)",
    fProjectZh: "Project / Site (Chinese)",
    fProjectPlaceholder: "e.g. MTR Yau Tong Station MOD",
    fInvoiceNo: "Invoice No.",
    fInvoiceHint: "Fill in once invoiced",
    sectionSite: "Construction Site",
    fSiteName: "Construction Site Name",
    fSiteHint: "English or Traditional Chinese — one is fine, no need for both",
    fSitePlaceholder: "e.g. 油塘站 or MTR Yau Tong Station",
    sectionCargo: "Cargo Detail",
    fItemType: "Item Type",
    fPackages: "# of Packages",
    fUnitCode: "Escalator / Elevator # or Code",
    fUnitCodeHint: "e.g. L1, E02",
    fDescription: "Description",
    fDescriptionHint: "e.g. unit no. / equipment description",
    fReference: "Reference / SHK No.",
    fReferenceHint: "Schindler reference, if applicable",
    fWeight: "Weight (KG)",
    fVolume: "Volume (CBM)",
    f20: "No. of 20' Containers",
    f40: "No. of 40' Containers",
    sectionArrival: "Arrival & Depot",
    fArrivingType: "Arriving Type",
    fArrivingTypeHint: "Devan = we unpack a container ourselves. CFS = client delivers loose packages, no devanning needed.",
    fDepot: "Depot",
    fDepotHint: "Which depot this is going into",
    fDepotLocation: "Depot Location / Bay",
    fTerminalArrival: "Terminal Arrival Date",
    fTerminalArrivalHint: "ETA at container terminal",
    fTerminalLFD: "Terminal Last Free Day",
    fTerminalLFDHint: "Latest date to collect before demurrage",
    fConfirmedCollection: "Confirmed Collection Date",
    fConfirmedCollectionHint: "Date we go to terminal",
    fDepotArrival: "Depot Arrival Date",
    fDepotArrivalHint: "Starts the 14-day free storage clock",
    fPlannedDelivery: "Planned Delivery Date",
    fPlannedDeliveryHint: "Estimate only — record real deliveries via the Deliveries tab",
    deliveryProgress: (del, tot, count) => `${del} of ${tot} unit(s) delivered across ${count} delivery record(s), last on `,
    deliveryProgressManage: "Manage deliveries via the Deliveries tab.",
    fNotes: "Notes",
    saveBtn: "Save Entry",
    cancelBtn: "Cancel",

    deliveryTitlePrefix: "Record Delivery —",
    combinedDeliveryTitle: (n) => `Record Combined Delivery — ${n} entries`,
    combinedHistoryHiddenNote: "Past deliveries for these entries aren't shown here — open each entry individually (from Inventory) to view or cancel its delivery history.",
    recordCombinedBtn: (n) => n > 0 ? `Record Delivery for Selected (${n})` : "Record Delivery for Selected",
    combinedSelectionHint: "Tick every entry going out together on the same job \u2014 e.g. two different batches heading to the same site on the same day \u2014 then record them as one combined delivery with a single job sheet.",
    selectedTag: "Included \u2713",
    includeInDeliveryBtn: "Include in this delivery",
    addCombinedDeliveryBtn: "Record Combined Delivery",
    combinedPrintLabel: "Combined Delivery Job Sheet",
    addMoreBatchesBtn: "+ Add more batches to this delivery",
    addMoreBatchesSearchPlaceholder: "Search by unit no. or job no.",
    addMoreBatchesScopeNote: (site) => `Showing other entries at ${site} only.`,
    addMoreBatchesNoneMsg: "No other entries at this same construction site are available to add.",
    plannedWasText: (date) => ` · planned delivery was ${date}`,
    progressOf: "of",
    progressDeliveredSoFar: "unit(s) delivered so far",
    progressRemaining: "remaining at the depot",
    colDate: "Date",
    colQty: "Qty",
    colDeliveredTo: "Delivered To",
    colReceivedBy: "Received By",
    removeBtn: "Remove",
    fDeliveryDate: "Delivery Date",
    fDeliveryDateHint: "This is what closes the storage clock once fully delivered",
    fQty: "Quantity Delivered",
    fQtyHint: (r) => `${r} remaining`,
    fDeliveredTo: "Delivered To",
    fDeliveredToHint: "Defaults to the construction site",
    fReceivedBy: "Received By",
    fReceivedByHint: "If left blank, uses the Ordered By contact for this project",
    overshootMsg: (r) => `Only ${r} unit(s) remain at the depot for this entry — reduce the quantity.`,
    addDeliveryBtn: "Add Delivery Record",
    closeBtn: "Close",
    allDeliveredMsg: "All units for this entry have been delivered.",
    selectItemMsg: "Select an item to record a delivery. Items with more than one package (e.g. a lot of escalator crates) can be delivered in stages — each stage reduces what's left at the depot.",
    nothingAtDepotMsg: "Nothing is currently at the depot.",
    recordDeliveryBtn: "Record Delivery",

    noneFoundMsg: "No duplicates found. Entries are flagged when client, project, unit code, item type, depot arrival date, and invoice number all match another entry.",
    matchingEntries: (n) => `${n} matching entries`,
    deleteAllBtn: (n) => `Delete all ${n}`,
    colInvoiceNo: "Invoice No.",
    colAddedOn: "Added On",
    keepDeleteBtn: "Keep this, delete others",

    tabExcel: "Excel Upload",
    tabPdf: "PDF Scan",
    tabManualPackingList: "Manual Entry",
    manualPackingListDesc: "For older jobs with no packing list file, or one that's incomplete \u2014 type in the case list by hand. This creates one Incoming shipment the same way an uploaded file would, ready to check in via Devan/CFS.",
    legacyManualEntryNote: "Manually entered \u2014 no packing list file on file for this shipment.",
    excelTitle: "Import from Excel",
    excelDesc: 'Column headers are matched against the depot\'s field names automatically (e.g. "Invoice No.", "Depot Arrival Date"). Unrecognized columns are skipped and listed below.',
    chooseFileBtn: "Choose File (.xlsx, .xls, .csv)",
    downloadTemplateBtn: "Download blank template",
    selectedCount: (sel, tot) => `${sel} of ${tot} selected to import.`,
    selectAllBtn: "Select all",
    selectNonDupBtn: "Select non-duplicates only",
    unmatchedMsg: "Columns not recognized (skipped): ",
    prevColClient: "Client",
    prevColProject: "Project",
    prevColItemType: "Item Type",
    prevColDepot: "Depot",
    prevColDepotArrival: "Depot Arrival",
    prevColMatch: "Match",
    importBtn: (n) => `Import ${n} Record(s)`,
    discardBtn: "Discard",
    pdfTitle: "Scan a PDF Packing List",
    pdfDesc: "Upload a client's packing list, delivery memo, or shipping list PDF — any layout. It reads the lift/lot breakdown, cases, weights and CBM automatically, the same way the Excel packing list import does.",
    choosePdfBtn: "Choose PDF",
    scanningMsg: "Reading document…",
    pdfReadErrorMsg: "Couldn't automatically read this PDF. Please check the file, or enter the details manually below.",
    pdfTruncatedMsg: "This document is very dense (lots of cases/items) and the automatic read got cut off partway through. Try again \u2014 it sometimes succeeds on a second attempt \u2014 or enter the details manually below.",
    pdfKeyWarning: "This uses your own Anthropic API key, entered below and saved only in this browser. Since it's used directly from this page, anyone who opens this file's developer tools could see it — fine for internal testing among trusted staff, not for wider distribution.",
    pdfApiKeyLabel: "Anthropic API Key",
    pdfApiKeyHint: "From console.anthropic.com — saved only in this browser",
    pdfSaveKeyBtn: "Save Key",
    pdfKeySavedBadge: "Key saved",
    pdfNeedKeyMsg: "Add and save an API key above to enable PDF scanning.",
    reviewWarningMsg: "Review the extracted details below before saving — automatic reads can miss or misplace things.",
    dupWarningMsg: (match) => `This looks like it may match an existing entry: ${match.id} (${match.client} · ${match.project}).`,
    excelErrorMsg: "Couldn't read that file. Make sure it's a .xlsx, .xls, or .csv export.",
    excelNoRowsMsg: "No rows found in that file.",

    badgePendingCollection: "PENDING COLLECTION",
    badgeLfdOverdue: (d) => `LFD OVERDUE ${d} DAYS`,
    badgeLfdSoon: (d) => `LFD IN ${d} DAYS`,
    badgeBillable: (d) => `BILLABLE · ${d} DAYS`,
    badgeFree: (d) => `FREE · ${d} DAYS LEFT`,
    badgePartial: (del, tot, extra) => `PARTIAL ${del}/${tot}${extra}`,
    badgeDelivered: (extra) => `DELIVERED${extra}`,
    badgeBilledSuffix: (d) => ` · BILLED ${d} DAYS`,
    badgeBillableSuffix: (d) => ` · ${d} DAYS`,
    badgeDupOf: (id) => `DUP OF ${id}`,
    badgeNew: "NEW",

    sectionPackages: "Itemized Packages",
    packagesHint: "Optional \u2014 add individual case/package codes so specific ones can be delivered separately (e.g. case 2A, 13A). Leave empty to just track a total count instead.",
    bulkAddLabel: "Quick add (comma or line separated)",
    bulkAddPlaceholder: "e.g. 1A, 2A, 3A, 12A, 13A",
    bulkAddBtn: "Add Codes",
    packageCodeCol: "Code",
    packageDescCol: "Description",
    packageWeightCol: "Weight (kg)",
    packageCbmCol: "CBM",
    removePackageBtn: "Remove",
    packagesCountSummary: (n) => `${n} itemized package(s)`,
    noPackagesMsg: "No itemized packages yet \u2014 using the plain count above instead.",

    uploadModePackingList: "Packing List (Excel/PDF)",
    uploadModeLegacy: "Legacy Upload (CFS/Devan/Delivery)",
    tabPackingList: "Packing List Import",
    packingListTitle: "Import a Client Packing List",
    packingListDesc: "Upload a client's packing list as-is (TK Elevator, Schindler, OTIS, etc. all use different layouts \u2014 this reads the columns automatically). Each distinct lift/lot found becomes its own manifest entry with its cases itemized.",
    choosePackingListBtn: "Choose Packing List File",
    packingListNoStructure: "Couldn't recognize a packing-list table in that file. It may use a layout this hasn't seen before \u2014 try Excel Upload instead, or add entries manually.",
    packingListDetectedTitle: (n) => `Found ${n} lift/lot(s) in this file`,
    packingListCommonFieldsTitle: "Apply to all of these",
    packingListApplyClient: "Client",
    packingListApplyProject: "Project",
    packingListApplyDepot: "Depot",
    packingListApplyDepotArrival: "Depot Arrival Date (all items arrived together)",
    packingListApplyDepotLocation: "Depot Location / Bay",
    packingListImportBtn: (n) => `Import ${n} Manifest Entr${n === 1 ? "y" : "ies"}`,
    colLot: "Lift / Lot",
    colPackages: "Packages",
    colContainers: "Container(s)",
    colWeight: "Weight (kg)",
    selectByBatchLabel: "Select by batch (this delivery can pull from more than one arrival batch)",
    selectCodesLabel: "Select which package codes are going out",
    noCodesRemainingMsg: "All itemized packages for this entry have been delivered.",

    colCbm: "CBM",
    colKg: "KG",
    colJobNo: "Job No.",

    sectionJobSheet: "Job Sheet",
    fJobNumber: "Job No.",
    fJobNumberHint: "YYMM + running number, shared across Devan/CFS/Delivery",
    generateJobNoBtn: "Generate Job No.",
    fOrderedBy: "Ordered By",
    fPoNumber: "P.O. No.",
    fPoNumberHint: "Provided by the client",
    fJobRef: "Job Ref. (Site Code)",
    fJobRefHint: "e.g. KTN-002, GAGE-001",
    printJobSheetBtn: "Print Job Sheet / Save as PDF",
    printBtn: "Print / Save as PDF",
    closePreviewBtn: "Close",

    jsTitle: "JOB SHEET",
    jsTitleZh: "工　單",
    jsFrom: "FROM",
    jsFromZh: "由",
    jsTo: "TO",
    jsToZh: "送",
    jsAccount: "ACCOUNT",
    jsAccountZh: "客戶",
    jsJobNo: "JOB NO.",
    jsJobNoZh: "快達單號",
    jsDate: "DATE",
    jsDateZh: "日期",
    jsOrderedBy: "ORDERED BY",
    jsOrderedByZh: "落單人",
    jsPoNo: "P.O. NO.",
    jsPoNoZh: "採購編號",
    jsJobRef: "JOB REF.",
    jsJobRefZh: "地盤代號",
    jsDescription: "DESCRIPTION",
    jsDescriptionZh: "貨物資料／工作程序",
    jsIssuedBy: "ISSUED BY",
    jsIssuedByZh: "出單人",
    jsTotal: "TOTAL",
    jsTotalZh: "共",
    jsPkgs: "PKGS",
    jsKgs: "KGS",
    jsCbm: "CBM",
    jsDevanFrom: (dep) => `DEVAN AT ${dep}`,
    jsCfsFrom: "CFS \u2014 CLIENT DELIVERED, NO DEVANNING",
    jsDeliveryType: "DELIVERY",
    jsDevanType: "DEVAN",
    jsCfsType: "CFS",
    jsSignatureLine: "Customer signature confirming above work completed:",
    jsTemplateLabel: "Template",
    jsCfsFromPreset: "CFS location…",
    jsEditableHint: "Dashed boxes can be edited before printing.",
    jsOversizeLabel: "Oversize Cases 超長/超大件:",
    jsOversizePlaceholder: "#case codes @X.XXCBM per line, then (合共 Total: X CBM)  xN 倍",
    jsOversizeNote: "Oversize CBM is charged at the listed multiplier. Applies to Schindler & Chevalier accounts only.",
    deliveryHistoryLabel: "Delivery History",
    signedDocArrivalLabel: "Signed arrival form (Devan/CFS)",
    signedDocUploadBtn: "Upload signed copy",
    signedDocReplaceBtn: "Replace signed copy",
    signedDocViewBtn: "View signed copy",
    signedDocSavingMsg: "Saving…",
    signedDocFailMsg: "Couldn't save this file. Photos are compressed automatically, but PDFs must be under about 3 MB — try a photo of the signed sheet instead.",
    fSsDoNo: "SS/D.O. No. (提單資料)",
    fSsDoNoHint: "Vessel, voyage & container numbers — prefills the SS/D.O. row on Devan/CFS job sheets",
    jsEstimatedNote: "~ Weight/CBM estimated as a proportional share of the full entry — not individually weighed per package.",

    navDirectory: "Directory",
    tabSitesAccounts: "Sites & Accounts",
    tabEmployees: "Employees",

    dirTitle: "Site & Account Directory",
    dirDesc: "One entry per real construction site \u2014 keeps client, job ref, and contact consistent, and stops the same site being typed differently (English vs Chinese) on different entries.",
    dirAddBtn: "+ Add Site",
    fSiteEn: "Site Name (English)",
    fSiteZh: "Site Name (Chinese)",
    fDirClient: "Account (Client)",
    fDirJobRef: "Job Ref. (Site Code)",
    fDirOrderedBy: "Default Ordered By",
    fDirOfficer: "Responsible Account Officer",
    dirColSite: "Site",
    dirColClient: "Account",
    dirColJobRef: "Job Ref.",
    dirColOfficer: "Account Officer",
    dirColOrderedBy: "Ordered By",
    dirNoneMsg: "No sites added yet.",
    selectFromDirectory: "Fill from Directory",
    selectFromDirectoryPlaceholder: "\u2014 pick a site to auto-fill \u2014",
    saveNewSiteToDirectory: (name) => `Save "${name}" as a new site in the Directory, so future imports recognize it automatically`,

    siteTotalsTitle: "CBM & KG Remaining by Construction Site",
    siteTotalsColSite: "Construction Site",
    siteTotalsColClient: "Client",
    siteTotalsColItems: "Items",
    siteTotalsColCbm: "CBM Left",
    siteTotalsColKg: "KG Left",
    siteTotalsNoneMsg: "Nothing currently at the depot.",
    siteTotalsToggleHide: "Hide",
    siteTotalsToggleShow: "Show",

    empTitle: "Employees",
    empDesc: "Add anyone who does data entry, devans/CFS, or deliveries. Roles can be picked from the list or typed fresh.",
    empAddBtn: "+ Add Employee",
    fEmpName: "Name",
    fEmpRole: "Role",
    fEmpRolePlaceholder: "Pick or type a role",
    empColName: "Name",
    empColRole: "Role",
    empNoneMsg: "No employees added yet.",

    signedInAs: "Signed in as",
    mobileMenuLabel: "Menu",
    loginNameLabel: "Name",
    loginNamePlaceholder: "Select your name",
    loginPasswordLabel: "Password",
    loginErrorMissing: "Enter your name and password.",
    loginErrorWrong: "Wrong name or password.",
    loginBtn: "Log In",
    loginBusyMsg: "Checking\u2026",
    loginDefaultPwHint: "First time? The default password is \"Farspeed\" \u2014 you can change it after logging in.",
    changePasswordLink: "Change Password",
    changePasswordTitle: "Change Password",
    logoutBtn: "Log Out",
    pwCurrentLabel: "Current Password",
    pwNewLabel: "New Password",
    pwConfirmLabel: "Confirm New Password",
    pwTooShortMsg: "New password must be at least 4 characters.",
    pwMismatchMsg: "New password and confirmation don't match.",
    pwCurrentWrongMsg: "Current password is incorrect.",
    pwChangedMsg: "Password changed. Use your new password next time you log in.",
    pwSaveBtn: "Save New Password",
    signedInNone: "Not signed in",

    fRecordedBy: "Recorded By",
    fRecordedByHint: "Required \u2014 who is doing this Devan / CFS / Delivery",
    recordedByRequiredMsg: "Select who is recording this before saving.",

    resetBtn: "Reset All Deliveries (demo)",
    resetConfirmMsg: "This clears every delivery record on every item so everything shows as not yet delivered. It does not delete any manifest entries. Continue?",
    resetDoneMsg: "All delivery records cleared.",

    navJobLog: "Job Log",
    navBilling: "Billing",
    navUpload: "Upload",
    navIncoming: "Incoming",
    incomingTitle: "Incoming",
    incomingDesc: "Cases from uploaded packing lists that haven't been checked into the depot yet. Select which cases arrived via Devan or CFS and when \u2014 this creates or updates the real inventory entry.",
    incomingUploadToggle: "Upload Packing List",
    incomingUploadShow: "Show",
    incomingUploadHide: "Hide",
    incomingShowCompleted: "Show fully checked-in shipments",
    incomingNoneMsg: "Nothing incoming right now \u2014 upload a packing list to add cases here.",
    incomingCaseCount: (n) => `${n} case${n === 1 ? "" : "s"} on packing list`,
    incomingLinkedTo: (id) => `linked to ${id}`,
    incomingFullyCheckedIn: "Fully checked in",
    incomingRemainingBadge: (remaining, total) => `${remaining} of ${total} not yet checked in`,
    incomingSelectCasesLabel: "Select which cases are coming in",
    incomingCheckInBtn: (n) => n > 0 ? `Check In (${n})` : "Check In",
    incomingCheckedInNote: (incId) => `Checked in from Incoming ${incId}`,
    packingListAddToIncomingBtn: (n) => `Add ${n} Group${n === 1 ? "" : "s"} to Incoming`,
    packingListRemoveGroupBtn: "Remove this group",
    incomingDeleteBtn: "Delete this Incoming shipment",
    incomingDeleteConfirm: "Delete this Incoming shipment? This only removes it from Incoming \u2014 any inventory entry already checked in from it is not affected.",
    packingListIncomingHint: "This just adds the cases to Incoming \u2014 depot, Devan/CFS, job number, and date get decided later when you check them in from the Incoming tab.",
    billingTitle: "Storage Billing",
    billingDesc: "Search by client, project, job number, or case to see storage charges. Billed per arrival batch (not split by construction site) at the client's $/CBM monthly rate: free days first, the remainder of that month pro-rated by day, then every following month in full — even if the goods leave partway through it.",
    billingSearchPlaceholder: "Client, project, job no., or case no.",
    billingColClient: "Client",
    billingColProject: "Project / Site",
    billingColJobNo: "Job No.",
    billingColBatchDate: "Arrival",
    billingColCbm: "CBM",
    billingColRate: "Rate",
    billingColStatus: "Status",
    billingColTotal: "Total",
    billingNoneMsg: "No billable storage found for this search.",
    billingPerCbmMonth: "CBM/mo",
    billingOngoing: "ONGOING",
    billingClosed: "DELIVERED",
    billingShowBtn: "Show breakdown",
    billingHideBtn: "Hide",
    billingCasesLabel: "Cases",
    billingFreeDaysNote: (d) => `${d} free days applied before billing starts.`,
    billingEstimatedNote: "* CBM split estimated from package count share (no per-case CBM available for the delivered/remaining split).",
    billingGrandTotal: "Grand Total",
    billingFootnote: "Ongoing rows are calculated up to today and will keep growing until the goods are marked delivered. Rates are set in Directory → CBM Pricing.",
    billingDeleteItemBtn: "Delete this entry (admin password required)",
    adminConfirmTitle: "Confirm Admin Action",
    adminConfirmDesc: (name) => `Re-enter ${name || "your"}'s password to continue. This permanently deletes the underlying inventory entry, not just this billing row.`,
    adminConfirmBtn: "Delete Permanently",
    billingModeSearch: "Search",
    billingModeMonthly: "Monthly Summary",
    billingMonthLabel: "Month",
    billingYearLabel: "Year",
    billingMonthNoneMsg: "No storage charges fall in this month.",
    billingMonthFootnote: "Each client's total is what should match their MYOB invoice for this month — use this to double-check before billing. Click a client to see every line item behind their total.",
    jobLogTitle: "All Job Numbers Used",
    jobLogDesc: "Every Devan, CFS, and Delivery job number ever created, most recent first. Click any row to view and reprint that job sheet.",
    jobLogColJobNo: "Job No.",
    jobLogColType: "Type",
    jobLogColDate: "Date",
    jobLogColClient: "Client",
    jobLogColSite: "Project / Site",
    jobLogColRecordedBy: "Recorded By",
    jobLogNoneMsg: "No job numbers created yet.",
    viewReprintBtn: "View / Reprint",

    navCancelledJobs: "Cancelled Jobs",
    cancelJobBtn: "Cancel",
    cancelledJobsTitle: "Cancelled Jobs",
    cancelledJobsDesc: "Voided Devan, CFS, and Delivery job sheets. They're kept here (not the main Job Log) so the job number stays reserved. Since this is a demo, you can permanently delete them from here to clean up test data.",
    cancelledJobsNoneMsg: "Nothing cancelled.",
    restoreBtn: "Restore",
    purgeBtn: "Permanently Delete",
    permanentDeleteConfirmMsg: "This permanently deletes the record and frees up its job number. This cannot be undone. Continue?",

    inventoryRemainingLabel: "remaining in depot",
    inventoryNoRemainingPkgsMsg: "All itemized packages have left the depot.",

    settingsLabel: "Settings",
    navDuplicatesShort: "Duplicates",

    depotOverviewTitle: "Depot Overview",
    depotOverviewItemsLabel: "item(s)",

    newEntryManual: "Manual",
    newEntryImport: "Import",

    splitArrivalLabel: "Split Arrival — Cases Arriving on Different Days",
    splitArrivalHint: "If this packing list isn't all devanned / delivered to the depot on the same day, record each batch here with its own date and type. The main Depot Arrival Date will follow the earliest batch automatically.",
    splitArrivalCasesCol: "Cases",
    splitArrivalAddBtn: "Add Arrival Batch",
    splitArrivalSelectHint: "Tap the cases that arrived in this batch, then add it.",
    splitArrivalAllAssignedMsg: "All cases are assigned to an arrival batch.",
    badgePartialArrival: (a, b) => `AT DEPOT ${a}/${b} · MORE ARRIVING`,
    notYetArrivedTag: "not yet arrived",
    notYetArrivedHint: "This case has not yet arrived at the depot, so it can't be delivered out yet.",
    pendingArrivalNotice: (n) => `${n} case(s) have not yet arrived at the depot and can't be selected for delivery. Record their arrival batch on the Edit screen when they land.`,

    tabFreeStorage: "Free Storage",
    freeStorageTitle: "Free Storage Days",
    freeStorageDesc: "Standard free storage is 14 days from each devan / CFS arrival (Schindler: 21 days). For projects with a special arrangement — e.g. Otis, which varies by project — add a rule below. The rule applies when the project or construction site name contains the text you enter, and overrides the client default.",
    fFreeProject: "Project name contains",
    fFreeProjectHint: "e.g. Kwu Tung North Area 19",
    fFreeDays: "Free days",
    freeStorageAddBtn: "Add Rule",
    freeColProject: "Project match",
    freeColDays: "Free days",
    freeStorageNoneMsg: "No project rules yet — standard 14 days applies (Schindler: 21 days).",
    tabPricing: "CBM Pricing",
    pricingTitle: "CBM Rate per Client",
    pricingDesc: "Storage in the Warehouse and Open Yard is billed by CBM. These rates apply per CBM of storage; KG is only used when calculating Delivery and CFS charges. Leave a field blank to use the default rate shown as its placeholder.",
    pricingColClient: "Client",
    pricingColRate: "Rate",
    pricingPerCbm: "/ CBM",
    pricingResetBtn: "Reset to default",
    tabLegacy: "Legacy Uploads",
    legacyUploadTitle: "Legacy Upload",
    legacyUploadDesc: "Bulk-upload old Devan, CFS, Delivery, Shifting, Hoisting and other job sheet files. Best done in two rounds: upload all Devan/CFS files first (these create the storage arrivals), then upload Delivery files afterward \u2014 each one automatically finds and closes the arrival it refers to. Every file type is archived and kept searchable below regardless.",
    legacyChooseFilesBtn: "Choose Files\u2026",
    legacyDocType: "Doc Type",
    legacyProjectSite: "Project / Site",
    legacyUnitCode: "Unit / Lift No.",
    legacyPkgs: "Packages",
    legacyWeightKg: "Weight (kg)",
    legacyCbm: "CBM",
    legacyAlreadyDelivered: "Already delivered \u2014 close this job immediately instead of leaving it at the depot",
    legacyProcessBtn: (n) => `Process ${n} File${n === 1 ? "" : "s"}`,
    legacyProcessingMsg: "Processing\u2026",
    legacyResultsMsg: (archived, created, delivered, enriched, checkedIn) => `Archived ${archived} file${archived === 1 ? "" : "s"}${created > 0 ? ` \u2014 created ${created} inventory ${created === 1 ? "entry" : "entries"}` : ""}${checkedIn > 0 ? ` \u2014 checked in ${checkedIn} case group${checkedIn === 1 ? "" : "s"} against Incoming` : ""}${enriched > 0 ? ` \u2014 matched ${enriched} to existing ${enriched === 1 ? "entry" : "entries"} and added its details` : ""}${delivered > 0 ? ` \u2014 recorded ${delivered} deliver${delivered === 1 ? "y" : "ies"} against existing arrivals` : ""}.`,
    legacyImportedNote: (name) => `Imported from legacy file: ${name}`,
    legacyAutoClosedNote: "Auto-closed on legacy import (marked already delivered).",
    legacyBacklogTitle: "Backlog",
    legacyBacklogDesc: "Every file uploaded through Legacy Upload, whether it created an inventory entry or was archived only.",
    legacyBacklogNoneMsg: "No legacy files uploaded yet.",
    legacyColFile: "File",
    legacyColLinked: "Linked Entry",
    legacyArchivedOnly: "Archived only",
    legacyEditLinkedHint: "Only fixes this archive listing \u2014 edit the linked FS-#### entry directly in Inventory if its data needs correcting too.",
    legacyClientUnresolved: "\u2014 select client \u2014",
    legacyClientRequiredSummaryMsg: "One or more files don't have a recognized client \u2014 select the correct client for each file before processing.",
    legacyDeliveredFrom: (id) => `Delivered from ${id}`,
    legacyUnmatchedReferral: (job) => `No arrival found for job ${job}`,
    legacyUnmatchedHint: "This delivery refers to a job number that isn't in the system yet - upload its Devan/CFS file first, or link it manually later.",
    legacyProjectSiteEn: "Project / Site (English)",
    legacyProjectSiteZh: "Project / Site (Chinese)",
    legacySiteRequiredMsg: "Enter at least one of English or Chinese site name.",
    legacySiteRequiredSummaryMsg: "One or more files are missing a site name in both languages \u2014 fill in at least one (English or Chinese) for each file before processing.",
    legacyScanningMsg: "Reading files\u2026",
    legacyAutoDetectHint: "Excel files (.xlsx/.xls/.csv) are scanned automatically for client, site, job number, date, SS/D.O. info, and package totals \u2014 review and correct the pre-filled fields below before processing. PDFs and images still need manual entry.",
    legacyAutoDetectedTag: "Auto-detected from file \u2014 please check",
    legacyReferLine: (job, date) => `refers to job no. ${job} on ${date}`,
    legacyReferJobNoLabel: "Refers to Arrival Job No.",
    legacyEnrichedNote: (name) => `Enriched from legacy file: ${name}`,
    legacyMatchedIncoming: (id) => `Matched to Incoming ${id} \u2014 select which of its cases arrived in this file`,
    legacyMatchedIncomingCount: (n) => `Matched ${n} Incoming shipment${n === 1 ? "" : "s"} at this site \u2014 select which cases from each arrived in this file`,
    legacyReferJobNoHint: "Optional \u2014 narrows the match to one specific arrival job number. Leave blank to match by client + site instead (can find several).",
    legacyMatchedItemsCount: (n) => `Matched ${n} inventory ${n === 1 ? "entry" : "entries"} with cases still at the depot \u2014 select which left in this delivery`,
    legacyTypeSelectPlaceholder: "e.g. 1,3-5,7",
    legacyTypeSelectBtn: "Add",
    legacySelectedTotals: (count, kg, cbm) => `Selected: ${count} pkg${count === 1 ? "" : "s"} \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacySelectedTotalsGrand: (count, kg, cbm) => `Total selected across all entries: ${count} pkg${count === 1 ? "" : "s"} \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacyMatchedItem: (id) => `Delivering from ${id}`,
    legacyArrivalStaysOpenHint: "Stays open at the depot until a matching Delivery file is uploaded (or you record a delivery for it normally).",
    legacyNoReferralHint: "No \"Ref Job no.\" line detected \u2014 enter the arrival's job number manually, or this file will only be archived.",
    legacyNoArrivalFoundHint: (job) => `No inventory entry found with job number ${job} \u2014 check the number is correct, or upload that Devan/CFS file first. This file will still be archived, but no delivery will be recorded against it.`,
  },
  zh: {
    appSubtitle: "倉庫及貨物存倉表",
    navDashboard: "總覽",
    navInventory: "存倉記錄",
    navNewEntry: "新增記錄",
    navDeliveries: "送貨記錄",
    navDuplicates: "重複記錄",
    navDuplicatesCount: (n) => `重複記錄（${n}）`,
    navImport: "匯入",
    themeToggleLabel: "切換深色模式",
    langToggleLabel: "EN",
    loadingMsg: "正在載入倉存記錄…",
    saveErrorMsg: "儲存失敗，請重試。",

    statAtDepot: "在倉",
    statPending: "待提取",
    statBillable: "現正收費",
    statLfd: "截關期提示",
    dupBanner: (n) => `發現 ${n} 組可能重複的記錄 — 客戶、項目、單位編號及抵倉日期相同。`,
    reviewDuplicatesBtn: "查看重複記錄",
    lfdSectionTitle: "貨櫃碼頭截關期 — 需要處理",
    billableSectionTitle: "現正收費之存倉項目",
    billableEmptyMsg: "暫時沒有超過14天免費存倉期的項目。",
    sinceLabel: "由",

    searchLabel: "搜尋",
    searchPlaceholder: "項目、客戶、編號、SHK編號、件號、工單號",
    clientLabel: "客戶",
    statusLabel: "狀態",
    statusAll: "全部",
    statusPending: "待提取",
    statusAtDepot: "在倉",
    statusPartial: "部分已送貨",
    statusDelivered: "已送貨",
    depotLabel: "貨倉",
    exportBtn: (n) => `匯出Excel（${n}）`,
    newEntryBtn: "+ 新增記錄",
    colId: "編號",
    colClient: "客戶",
    colProjectSite: "項目/地盤",
    colUnit: "單位編號",
    deliverySearchPlaceholder: "客戶、項目、單位編號、工單號或件號",
    colDepot: "貨倉",
    colDepotArrival: "抵倉日期",
    colStatus: "狀態",
    noRecordsMsg: "沒有符合的記錄。請新增記錄開始使用。",
    editBtn: "編輯",
    deleteBtn: "刪除",
    deliverBtn: "送貨",
    duplicateBadge: "重複？",

    titleNew: "新增倉存記錄",
    titleEdit: "編輯倉存記錄",
    fClient: "客戶",
    fProject: "項目",
    fProjectEn: "項目／地盤（英文）",
    fProjectZh: "項目／地盤（中文）",
    fProjectPlaceholder: "例如：港鐵油塘站現代化工程",
    fInvoiceNo: "發票編號",
    fInvoiceHint: "開發票後填寫",
    sectionSite: "地盤資料",
    fSiteName: "地盤名稱",
    fSiteHint: "英文或繁體中文皆可，毋須兩者兼填",
    fSitePlaceholder: "例如：油塘站 or MTR Yau Tong Station",
    sectionCargo: "貨物詳情",
    fItemType: "貨物類型",
    fPackages: "件數",
    fUnitCode: "扶手電梯/升降機編號",
    fUnitCodeHint: "例如：L1、E02",
    fDescription: "描述",
    fDescriptionHint: "例如：單位編號／設備描述",
    fReference: "參考編號／SHK編號",
    fReferenceHint: "適用於迅達（Schindler）之參考編號",
    fWeight: "重量（公斤）",
    fVolume: "體積（立方米）",
    f20: "20呎貨櫃數量",
    f40: "40呎貨櫃數量",
    sectionArrival: "抵達及倉存資料",
    fArrivingType: "抵達方式",
    fArrivingTypeHint: "Devan（拆櫃）= 由我方自行拆櫃；CFS = 客戶直接送來散件，毋須拆櫃。",
    fDepot: "貨倉",
    fDepotHint: "貨物存放之貨倉",
    fDepotLocation: "倉位/存放位置",
    fTerminalArrival: "貨櫃碼頭抵港日期",
    fTerminalArrivalHint: "預計抵達貨櫃碼頭日期",
    fTerminalLFD: "貨櫃碼頭截關期",
    fTerminalLFDHint: "須於此日期前提取貨物，否則產生額外費用",
    fConfirmedCollection: "確認提貨日期",
    fConfirmedCollectionHint: "我方前往貨櫃碼頭提貨之日期",
    fDepotArrival: "抵倉日期",
    fDepotArrivalHint: "開始計算14天免費存倉期",
    fPlannedDelivery: "預計送貨日期",
    fPlannedDeliveryHint: "僅供預算之用 — 實際送貨請於「送貨記錄」分頁記錄",
    deliveryProgress: (del, tot, count) => `已送出 ${tot} 件中之 ${del} 件，共 ${count} 次送貨記錄，最近一次為 `,
    deliveryProgressManage: "請於「送貨記錄」分頁管理送貨情況。",
    fNotes: "備註",
    saveBtn: "儲存記錄",
    cancelBtn: "取消",

    deliveryTitlePrefix: "記錄送貨 —",
    combinedDeliveryTitle: (n) => `記錄合併送貨 — 共${n}項記錄`,
    combinedHistoryHiddenNote: "此處不會顯示這些記錄各自的送貨歷史 — 如需查看或取消個別記錄的送貨，請從存倉記錄個別打開。",
    recordCombinedBtn: (n) => n > 0 ? `為已選項目記錄送貨 (${n})` : "為已選項目記錄送貨",
    combinedSelectionHint: "勾選所有要一同送出的記錄（例如兩個不同批次但同日送往同一地盤），然後合併記錄為一張工單。",
    selectedTag: "已包含 ✓",
    includeInDeliveryBtn: "加入此次送貨",
    addCombinedDeliveryBtn: "記錄合併送貨",
    combinedPrintLabel: "合併送貨工單",
    addMoreBatchesBtn: "+ 加入更多批次到此送貨記錄",
    addMoreBatchesSearchPlaceholder: "按單位編號或工單號搜尋",
    addMoreBatchesScopeNote: (site) => `只顯示同屬 ${site} 的其他記錄。`,
    addMoreBatchesNoneMsg: "沒有其他屬於同一地盤的記錄可供加入。",
    plannedWasText: (date) => ` · 預計送貨日期為 ${date}`,
    progressOf: "／",
    progressDeliveredSoFar: "件已送出",
    progressRemaining: "件尚在倉",
    colDate: "日期",
    colQty: "數量",
    colDeliveredTo: "送達地點",
    colReceivedBy: "接收人",
    removeBtn: "移除",
    fDeliveryDate: "送貨日期",
    fDeliveryDateHint: "全部送出後將以此日期結算存倉期",
    fQty: "送貨數量",
    fQtyHint: (r) => `尚餘 ${r} 件`,
    fDeliveredTo: "送達地點",
    fDeliveredToHint: "預設為地盤名稱",
    fReceivedBy: "接收人",
    fReceivedByHint: "留空則使用此項目之落單人",
    overshootMsg: (r) => `此記錄現時只剩 ${r} 件在倉，請減少數量。`,
    addDeliveryBtn: "新增送貨記錄",
    closeBtn: "關閉",
    allDeliveredMsg: "此記錄之貨物已全部送出。",
    selectItemMsg: "請選擇項目以記錄送貨。多於一件之貨物（例如一批電梯零件）可分批送貨 — 每次記錄將扣減在倉數量。",
    nothingAtDepotMsg: "目前貨倉內沒有貨物。",
    recordDeliveryBtn: "記錄送貨",

    noneFoundMsg: "未發現重複記錄。當客戶、項目、單位編號、貨物類型、抵倉日期及發票編號均與另一記錄相同時，系統會標示為重複。",
    matchingEntries: (n) => `${n} 個相符記錄`,
    deleteAllBtn: (n) => `全部刪除（${n}項）`,
    colInvoiceNo: "發票編號",
    colAddedOn: "新增日期",
    keepDeleteBtn: "保留此項，刪除其他",

    tabExcel: "上載Excel",
    tabPdf: "掃描PDF",
    tabManualPackingList: "手動輸入",
    manualPackingListDesc: "適用於較舊、沒有裝箱單檔案或資料不完整的工作 \u2014 直接手動輸入件號清單。此操作會建立一項待到倉貨件，效果與上載檔案相同，可於拆櫃/CFS辦理到倉。",
    legacyManualEntryNote: "手動輸入 \u2014 此貨件沒有裝箱單檔案存檔。",
    excelTitle: "從Excel匯入",
    excelDesc: "系統會自動將欄位標題與倉存系統之欄位配對（例如「發票編號」、「抵倉日期」）。無法識別之欄位將被略過並於下方列出。",
    chooseFileBtn: "選擇檔案（.xlsx、.xls、.csv）",
    downloadTemplateBtn: "下載空白範本",
    selectedCount: (sel, tot) => `已選擇 ${sel} 項，共 ${tot} 項可匯入。`,
    selectAllBtn: "全選",
    selectNonDupBtn: "只選擇非重複項目",
    unmatchedMsg: "無法識別之欄位（已略過）：",
    prevColClient: "客戶",
    prevColProject: "項目",
    prevColItemType: "貨物類型",
    prevColDepot: "貨倉",
    prevColDepotArrival: "抵倉日期",
    prevColMatch: "配對結果",
    importBtn: (n) => `匯入 ${n} 項記錄`,
    discardBtn: "捨棄",
    pdfTitle: "掃描PDF裝箱單",
    pdfDesc: "上載客戶之裝箱單、送貨通知或Shipping List PDF — 任何格式皆可。系統會自動讀取升降機/批次分類、件號、重量及CBM，效果與Excel裝箱單匯入相同。",
    choosePdfBtn: "選擇PDF檔案",
    scanningMsg: "正在讀取文件…",
    pdfReadErrorMsg: "未能自動讀取此PDF。請檢查檔案，或於下方手動輸入資料。",
    pdfTruncatedMsg: "此文件內容非常密集（件數/項目眾多），自動讀取中途被截斷。請再試一次（有時第二次會成功），或於下方手動輸入資料。",
    pdfKeyWarning: "此功能使用您自己的Anthropic API金鑰，於下方輸入並只保存在此瀏覽器內。由於是直接從此頁面使用，任何開啟此檔案開發人員工具的人都有機會看到金鑰 — 適合內部信任員工測試，不適合對外派發。",
    pdfApiKeyLabel: "Anthropic API 金鑰",
    pdfApiKeyHint: "從 console.anthropic.com 取得 — 只保存在此瀏覽器",
    pdfSaveKeyBtn: "儲存金鑰",
    pdfKeySavedBadge: "金鑰已儲存",
    pdfNeedKeyMsg: "請於上方輸入並儲存API金鑰以啟用PDF掃描功能。",
    reviewWarningMsg: "請於儲存前核對以下自動讀取之資料 — 自動讀取或有遺漏或錯置。",
    dupWarningMsg: (match) => `此文件可能與現有記錄相符：${match.id}（${match.client} · ${match.project}）。`,
    excelErrorMsg: "未能讀取此檔案。請確保檔案為 .xlsx、.xls 或 .csv 格式。",
    excelNoRowsMsg: "此檔案內沒有資料列。",

    badgePendingCollection: "待提取",
    badgeLfdOverdue: (d) => `已過截關期 ${d} 日`,
    badgeLfdSoon: (d) => `截關期尚餘 ${d} 日`,
    badgeBillable: (d) => `計費中 · ${d}日`,
    badgeFree: (d) => `免費 · 尚餘${d}日`,
    badgePartial: (del, tot, extra) => `部分送貨 ${del}/${tot}${extra}`,
    badgeDelivered: (extra) => `已送貨${extra}`,
    badgeBilledSuffix: (d) => ` · 已收費${d}日`,
    badgeBillableSuffix: (d) => ` · ${d}日`,
    badgeDupOf: (id) => `重複於 ${id}`,
    badgeNew: "新項目",

    sectionPackages: "貨物件號清單",
    packagesHint: "選填 — 加入個別件號（例如 2A、13A），方便分開送貨。留空則只記錄總件數。",
    bulkAddLabel: "快速加入（以逗號或換行分隔）",
    bulkAddPlaceholder: "例如：1A, 2A, 3A, 12A, 13A",
    bulkAddBtn: "加入件號",
    packageCodeCol: "件號",
    packageDescCol: "描述",
    packageWeightCol: "重量（公斤）",
    packageCbmCol: "CBM",
    removePackageBtn: "移除",
    packagesCountSummary: (n) => `已列出 ${n} 件`,
    noPackagesMsg: "尚未列出個別件號 — 將使用上方之總件數。",

    uploadModePackingList: "裝箱單（Excel／PDF）",
    uploadModeLegacy: "舊資料上載（CFS／拆櫃／送貨）",
    tabPackingList: "匯入客戶裝箱單",
    packingListTitle: "匯入客戶裝箱單",
    packingListDesc: "直接上載客戶原本的裝箱單（TK Elevator、Schindler、OTIS 等各自格式不同 — 系統會自動辨識欄位）。文件內每部升降機/每個批次將自動分成獨立記錄，並列出其貨物件號。",
    choosePackingListBtn: "選擇裝箱單檔案",
    packingListNoStructure: "未能在此檔案中識別裝箱單表格結構，可能是未見過的格式 — 請改用「上載Excel」，或手動新增記錄。",
    packingListDetectedTitle: (n) => `在此檔案中找到 ${n} 個升降機/批次`,
    packingListCommonFieldsTitle: "套用至以下全部項目",
    packingListApplyClient: "客戶",
    packingListApplyProject: "項目",
    packingListApplyDepot: "貨倉",
    packingListApplyDepotArrival: "抵倉日期（全部貨物一同抵達）",
    packingListApplyDepotLocation: "倉位/存放位置",
    packingListImportBtn: (n) => `匯入 ${n} 項記錄`,
    colLot: "升降機/批次",
    colPackages: "件數",
    colContainers: "貨櫃",
    colWeight: "重量（公斤）",
    selectByBatchLabel: "按到倉批次選取（此送貨可同時包含多於一個到倉批次）",
    selectCodesLabel: "選擇要送出的件號",
    noCodesRemainingMsg: "此記錄之個別件號已全部送出。",

    colCbm: "CBM",
    colKg: "KG",
    colJobNo: "快達單號",

    sectionJobSheet: "工單",
    fJobNumber: "快達單號",
    fJobNumberHint: "格式為年月+流水號，Devan/CFS/送貨共用同一組編號",
    generateJobNoBtn: "產生單號",
    fOrderedBy: "落單人",
    fPoNumber: "採購編號",
    fPoNumberHint: "由客戶提供",
    fJobRef: "地盤代號",
    fJobRefHint: "例如：KTN-002、GAGE-001",
    printJobSheetBtn: "列印工單／匯出PDF",
    printBtn: "列印／匯出PDF",
    closePreviewBtn: "關閉",

    jsTitle: "JOB SHEET",
    jsTitleZh: "工　單",
    jsFrom: "FROM",
    jsFromZh: "由",
    jsTo: "TO",
    jsToZh: "送",
    jsAccount: "ACCOUNT",
    jsAccountZh: "客戶",
    jsJobNo: "JOB NO.",
    jsJobNoZh: "快達單號",
    jsDate: "DATE",
    jsDateZh: "日期",
    jsOrderedBy: "ORDERED BY",
    jsOrderedByZh: "落單人",
    jsPoNo: "P.O. NO.",
    jsPoNoZh: "採購編號",
    jsJobRef: "JOB REF.",
    jsJobRefZh: "地盤代號",
    jsDescription: "DESCRIPTION",
    jsDescriptionZh: "貨物資料／工作程序",
    jsIssuedBy: "ISSUED BY",
    jsIssuedByZh: "出單人",
    jsTotal: "TOTAL",
    jsTotalZh: "共",
    jsPkgs: "PKGS",
    jsKgs: "KGS",
    jsCbm: "CBM",
    jsDevanFrom: (dep) => `DEVAN AT ${dep}`,
    jsCfsFrom: "CFS \u2014 客戶直接送到，毋須拆櫃",
    jsDeliveryType: "送貨 DELIVERY",
    jsDevanType: "拆櫃 DEVAN",
    jsCfsType: "CFS",
    jsSignatureLine: "客戶簽署確認 (按以上工作完成):",
    jsTemplateLabel: "工單類型",
    jsCfsFromPreset: "CFS 起運地點…",
    jsEditableHint: "虛線框可於列印前編輯。",
    jsOversizeLabel: "超長/超大件 Oversize Cases:",
    jsOversizePlaceholder: "每行 #件號 @X.XXCBM，最後一行 (合共 Total: X CBM)  xN 倍",
    jsOversizeNote: "超大件體積按上列倍數計算運費，僅適用於 Schindler 及 Chevalier 客戶。",
    deliveryHistoryLabel: "送貨記錄",
    signedDocArrivalLabel: "已簽收到倉工單 (拆櫃/CFS)",
    signedDocUploadBtn: "上載已簽工單",
    signedDocReplaceBtn: "更換已簽工單",
    signedDocViewBtn: "查看已簽工單",
    signedDocSavingMsg: "儲存中…",
    signedDocFailMsg: "未能儲存此檔案。相片會自動壓縮，但PDF須小於約3MB — 可改為上載已簽工單的相片。",
    fSsDoNo: "提單資料 SS/D.O. No.",
    fSsDoNoHint: "船名、航次及櫃號 — 會自動填入拆櫃/CFS工單的提單資料欄",
    jsEstimatedNote: "~ 重量／CBM為按比例估算，並非逐件過磅。",

    navDirectory: "目錄",
    tabSitesAccounts: "地盤及客戶",
    tabEmployees: "員工",

    dirTitle: "地盤及客戶目錄",
    dirDesc: "每個真實地盤只需登記一次 — 統一客戶、地盤代號及聯絡人資料，避免同一地盤因英文／中文名稱不同而被當作不同記錄。",
    dirAddBtn: "+ 新增地盤",
    fSiteEn: "地盤名稱（英文）",
    fSiteZh: "地盤名稱（中文）",
    fDirClient: "客戶",
    fDirJobRef: "地盤代號",
    fDirOrderedBy: "預設落單人",
    fDirOfficer: "負責客戶主任",
    dirColSite: "地盤",
    dirColClient: "客戶",
    dirColJobRef: "地盤代號",
    dirColOfficer: "客戶主任",
    dirColOrderedBy: "落單人",
    dirNoneMsg: "尚未新增任何地盤。",
    selectFromDirectory: "從目錄自動填寫",
    selectFromDirectoryPlaceholder: "— 選擇地盤以自動填寫 —",
    saveNewSiteToDirectory: (name) => `將「${name}」儲存為目錄中的新地盤，日後匯入將自動識別`,

    siteTotalsTitle: "各地盤存倉之CBM及KG",
    siteTotalsColSite: "地盤",
    siteTotalsColClient: "客戶",
    siteTotalsColItems: "項目數",
    siteTotalsColCbm: "剩餘CBM",
    siteTotalsColKg: "剩餘KG",
    siteTotalsNoneMsg: "目前貨倉內沒有貨物。",
    siteTotalsToggleHide: "隱藏",
    siteTotalsToggleShow: "顯示",

    empTitle: "員工",
    empDesc: "加入所有負責資料輸入、拆櫃／CFS或送貨的同事。職位可從清單選擇，亦可自行輸入新職位。",
    empAddBtn: "+ 新增員工",
    fEmpName: "姓名",
    fEmpRole: "職位",
    fEmpRolePlaceholder: "選擇或輸入職位",
    empColName: "姓名",
    empColRole: "職位",
    empNoneMsg: "尚未新增任何員工。",

    signedInAs: "登入身份",
    mobileMenuLabel: "選單",
    loginNameLabel: "姓名",
    loginNamePlaceholder: "請選擇你的姓名",
    loginPasswordLabel: "密碼",
    loginErrorMissing: "請輸入姓名及密碼。",
    loginErrorWrong: "姓名或密碼錯誤。",
    loginBtn: "登入",
    loginBusyMsg: "驗證中…",
    loginDefaultPwHint: "第一次使用？預設密碼為「Farspeed」，登入後可自行更改。",
    changePasswordLink: "更改密碼",
    changePasswordTitle: "更改密碼",
    logoutBtn: "登出",
    pwCurrentLabel: "目前密碼",
    pwNewLabel: "新密碼",
    pwConfirmLabel: "確認新密碼",
    pwTooShortMsg: "新密碼須至少4個字元。",
    pwMismatchMsg: "新密碼與確認密碼不相符。",
    pwCurrentWrongMsg: "目前密碼不正確。",
    pwChangedMsg: "密碼已更改，下次登入請使用新密碼。",
    pwSaveBtn: "儲存新密碼",
    signedInNone: "未登入",

    fRecordedBy: "記錄人",
    fRecordedByHint: "必填 — 由誰負責此次拆櫃／CFS／送貨",
    recordedByRequiredMsg: "儲存前請選擇記錄人。",

    resetBtn: "重設所有送貨記錄（示範用）",
    resetConfirmMsg: "此操作將清除所有項目之送貨記錄，令所有貨物顯示為尚未送貨，但不會刪除任何倉存記錄。是否繼續？",
    resetDoneMsg: "所有送貨記錄已清除。",

    navJobLog: "單號記錄",
    navBilling: "帳單",
    navUpload: "上載",
    navIncoming: "待到倉",
    incomingTitle: "待到倉",
    incomingDesc: "已上載裝箱單但尚未辦理到倉手續的貨件。選擇哪些件號經拆櫃或CFS到倉及日期 \u2014 系統會建立或更新相應的存倉記錄。",
    incomingUploadToggle: "上載裝箱單",
    incomingUploadShow: "顯示",
    incomingUploadHide: "收起",
    incomingShowCompleted: "顯示已全部到倉的貨件",
    incomingNoneMsg: "目前沒有待到倉貨件 \u2014 上載裝箱單以新增。",
    incomingCaseCount: (n) => `裝箱單共 ${n} 件`,
    incomingLinkedTo: (id) => `已連結至 ${id}`,
    incomingFullyCheckedIn: "已全部到倉",
    incomingRemainingBadge: (remaining, total) => `尚有 ${remaining}／${total} 件未到倉`,
    incomingSelectCasesLabel: "選擇哪些件號到倉",
    incomingCheckInBtn: (n) => n > 0 ? `辦理到倉 (${n})` : "辦理到倉",
    incomingCheckedInNote: (incId) => `由待到倉記錄 ${incId} 辦理到倉`,
    packingListAddToIncomingBtn: (n) => `新增 ${n} 組至待到倉`,
    packingListRemoveGroupBtn: "移除此組",
    incomingDeleteBtn: "刪除此待到倉貨件",
    incomingDeleteConfirm: "確定刪除此待到倉貨件？此操作只會移除待到倉記錄 \u2014 已辦理到倉的存倉記錄不受影響。",
    packingListIncomingHint: "此步驟只會將件號加入待到倉名單 \u2014 倉庫、拆櫃/CFS、工單號及日期會在稍後於「待到倉」分頁辦理到倉時才決定。",
    billingTitle: "存倉收費",
    billingDesc: "可按客戶、項目、工單號或件號搜尋存倉收費。收費以到倉批次計算（不按地盤分開），按客戶的每CBM月費計：先扣免費日數，該月餘下日數按日比例計算，其後每個月則全額收取——即使貨物於月中送出亦然。",
    billingSearchPlaceholder: "客戶、項目、工單號或件號",
    billingColClient: "客戶",
    billingColProject: "項目／地盤",
    billingColJobNo: "工單號",
    billingColBatchDate: "到倉日期",
    billingColCbm: "CBM",
    billingColRate: "收費",
    billingColStatus: "狀態",
    billingColTotal: "總額",
    billingNoneMsg: "此搜尋未有可收費的存倉記錄。",
    billingPerCbmMonth: "CBM/月",
    billingOngoing: "計算中",
    billingClosed: "已送出",
    billingShowBtn: "顯示明細",
    billingHideBtn: "收起",
    billingCasesLabel: "件號",
    billingFreeDaysNote: (d) => `已扣除${d}日免費存倉。`,
    billingEstimatedNote: "＊按件數比例估算已送/餘下CBM分配（此記錄沒有逐件CBM資料）。",
    billingGrandTotal: "總計",
    billingFootnote: "計算中的項目按至今日計算，直至貨物標記為已送出才會停止累加。收費可於 目錄 → CBM 收費 設定。",
    billingDeleteItemBtn: "刪除此記錄（需要管理密碼）",
    adminConfirmTitle: "確認管理員操作",
    adminConfirmDesc: (name) => `請重新輸入 ${name || "您"} 的密碼以繼續。此操作會永久刪除相關的存倉記錄，不只是這一行帳單。`,
    adminConfirmBtn: "永久刪除",
    billingModeSearch: "搜尋",
    billingModeMonthly: "每月總覽",
    billingMonthLabel: "月份",
    billingYearLabel: "年份",
    billingMonthNoneMsg: "此月份沒有存倉收費。",
    billingMonthFootnote: "各客戶總額應與其MYOB該月發票金額相符，可用作出單前核對。點擊客戶可查看組成總額的每一項明細。",
    jobLogTitle: "所有已使用的工作單號",
    jobLogDesc: "所有曾建立的Devan、CFS及送貨單號，最新在前。點擊任何一行可查看及重印該工單。",
    jobLogColJobNo: "單號",
    jobLogColType: "類型",
    jobLogColDate: "日期",
    jobLogColClient: "客戶",
    jobLogColSite: "項目/地盤",
    jobLogColRecordedBy: "記錄人",
    jobLogNoneMsg: "尚未建立任何單號。",
    viewReprintBtn: "查看／重印",

    navCancelledJobs: "已取消工作",
    cancelJobBtn: "取消",
    cancelledJobsTitle: "已取消工作",
    cancelledJobsDesc: "已作廢之Devan、CFS及送貨工單。此類記錄不會出現在主要之「單號記錄」，並保留其單號以免被重用。由於此為示範用途，可於此處永久刪除以清理測試資料。",
    cancelledJobsNoneMsg: "沒有已取消之記錄。",
    restoreBtn: "還原",
    purgeBtn: "永久刪除",
    permanentDeleteConfirmMsg: "此操作將永久刪除該記錄並釋放其單號，無法還原。是否繼續？",

    inventoryRemainingLabel: "存倉中",
    inventoryNoRemainingPkgsMsg: "所有件號均已送出，倉內已無剩餘。",

    settingsLabel: "設定",
    navDuplicatesShort: "重複記錄",

    depotOverviewTitle: "各倉存倉概覽",
    depotOverviewItemsLabel: "項",

    newEntryManual: "手動輸入",
    newEntryImport: "匯入",

    splitArrivalLabel: "分批到倉（貨箱不同日子到貨）",
    splitArrivalHint: "如同一張裝箱單不是同一日全部拆櫃／送到倉，可在此分批記錄，每批有自己的日期及類型。抵倉日期會自動跟隨最早一批。",
    splitArrivalCasesCol: "貨箱",
    splitArrivalAddBtn: "新增到倉批次",
    splitArrivalSelectHint: "點選此批次到達的貨箱，然後新增。",
    splitArrivalAllAssignedMsg: "所有貨箱已分配到倉批次。",
    badgePartialArrival: (a, b) => `已到倉 ${a}/${b} · 其餘未到`,
    notYetArrivedTag: "未到倉",
    notYetArrivedHint: "此貨箱尚未到倉，未能安排送出。",
    pendingArrivalNotice: (n) => `${n} 箱尚未到倉，未能選擇送貨。貨到後請在編輯頁記錄其到倉批次。`,

    tabFreeStorage: "免費存倉",
    freeStorageTitle: "免費存倉日數",
    freeStorageDesc: "標準免費存倉為每批拆櫃／CFS到倉日起計14日（Schindler為21日）。如個別項目另有安排（例如Otis按項目而定），可在下方新增規則：當項目或地盤名稱包含所輸入的文字時，即採用指定日數，並優先於客戶預設。",
    fFreeProject: "項目名稱包含",
    fFreeProjectHint: "例如：Kwu Tung North Area 19",
    fFreeDays: "免費日數",
    freeStorageAddBtn: "新增規則",
    freeColProject: "項目條件",
    freeColDays: "免費日數",
    freeStorageNoneMsg: "未有項目規則 — 採用標準14日（Schindler 21日）。",
    tabPricing: "CBM 收費",
    pricingTitle: "各客戶 CBM 收費",
    pricingDesc: "貨倉及露天場地存倉均以CBM計算收費，此處設定每CBM的收費；KG只用於計算送貨及CFS費用。留空即採用預設收費（顯示於欄位提示中）。",
    pricingColClient: "客戶",
    pricingColRate: "收費",
    pricingPerCbm: "／ CBM",
    pricingResetBtn: "回復預設",
    tabLegacy: "舊資料上載",
    legacyUploadTitle: "舊資料上載",
    legacyUploadDesc: "批量上載舊有的拆櫃、CFS、送貨、調位、吊運等工單檔案。建議分兩輪上載：先上載所有拆櫃/CFS檔案（建立到倉記錄），再上載送貨檔案 — 系統會自動配對並完結對應的到倉記錄。不論類型，所有檔案均會存檔並可於下方搜尋。",
    legacyChooseFilesBtn: "選擇檔案…",
    legacyDocType: "工單類型",
    legacyProjectSite: "項目／地盤",
    legacyUnitCode: "單位／升降機編號",
    legacyPkgs: "件數",
    legacyWeightKg: "重量 (kg)",
    legacyCbm: "CBM",
    legacyAlreadyDelivered: "已送出 — 直接完結此工作，不留在倉內",
    legacyProcessBtn: (n) => `處理 ${n} 個檔案`,
    legacyProcessingMsg: "處理中…",
    legacyResultsMsg: (archived, created, delivered, enriched, checkedIn) => `已存檔 ${archived} 個檔案${created > 0 ? `，並建立 ${created} 項存倉記錄` : ""}${checkedIn > 0 ? `，並為 ${checkedIn} 組件號辦理待到倉核對` : ""}${enriched > 0 ? `，並配對 ${enriched} 項現有記錄並補充其資料` : ""}${delivered > 0 ? `，並為 ${delivered} 項已到倉記錄登記送貨` : ""}。`,
    legacyImportedNote: (name) => `由舊資料檔案匯入：${name}`,
    legacyAutoClosedNote: "舊資料匯入時自動完結（標記為已送出）。",
    legacyBacklogTitle: "待處理記錄",
    legacyBacklogDesc: "所有透過舊資料上載的檔案，不論是否建立了存倉記錄。",
    legacyBacklogNoneMsg: "尚未上載任何舊資料檔案。",
    legacyColFile: "檔案",
    legacyColLinked: "連結記錄",
    legacyArchivedOnly: "僅存檔",
    legacyEditLinkedHint: "只會修正此存檔記錄 \u2014 如需同時修正相應的 FS-#### 存倉記錄，請直接於存倉列表編輯。",
    legacyClientUnresolved: "— 請選擇客戶 —",
    legacyClientRequiredSummaryMsg: "部分檔案未能識別客戶 — 處理前請為每個檔案選擇正確客戶。",
    legacyDeliveredFrom: (id) => `送出自 ${id}`,
    legacyUnmatchedReferral: (job) => `找不到工單號 ${job} 的到倉記錄`,
    legacyUnmatchedHint: "此送貨記錄指向的工單號尚未存在系統內 — 請先上載其拆櫃/CFS檔案，或稍後手動連結。",
    legacyProjectSiteEn: "項目／地盤（英文）",
    legacyProjectSiteZh: "項目／地盤（中文）",
    legacySiteRequiredMsg: "請至少填寫英文或中文地盤名稱其中一項。",
    legacySiteRequiredSummaryMsg: "部分檔案的地盤名稱（中英文）均未填寫 — 處理前請至少填寫其中一種語言。",
    legacyScanningMsg: "讀取檔案中…",
    legacyAutoDetectHint: "Excel檔案 (.xlsx/.xls/.csv) 會自動掃描客戶、地盤、工單號、日期、提單資料及件數/重量/CBM等資料 — 請於處理前檢查並修正下方已預填的欄位。PDF及圖片檔仍需手動輸入。",
    legacyAutoDetectedTag: "已從檔案自動偵測 — 請核對",
    legacyReferLine: (job, date) => `指向工單號 ${job}，日期 ${date}`,
    legacyReferJobNoLabel: "指向到倉工單號",
    legacyEnrichedNote: (name) => `由舊資料檔案補充資料：${name}`,
    legacyMatchedIncoming: (id) => `已配對至待到倉 ${id} \u2014 請選擇此檔案中已到達的件號`,
    legacyMatchedIncomingCount: (n) => `此地盤配對到 ${n} 項待到倉貨件 \u2014 請分別選擇此檔案中已到達的件號`,
    legacyReferJobNoHint: "可選填 \u2014 填寫後只會配對該工單號的到倉記錄；留空則按客戶＋地盤配對（可能配對多項）。",
    legacyMatchedItemsCount: (n) => `配對到 ${n} 項倉內仍有貨件的存倉記錄 \u2014 請選擇此次送貨送出的件號`,
    legacyTypeSelectPlaceholder: "例如 1,3-5,7",
    legacyTypeSelectBtn: "加入",
    legacySelectedTotals: (count, kg, cbm) => `已選：${count} 件 \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacySelectedTotalsGrand: (count, kg, cbm) => `全部已選（合計）：${count} 件 \u00b7 ${kg} kg \u00b7 ${cbm} cbm`,
    legacyMatchedItem: (id) => `送出自 ${id}`,
    legacyArrivalStaysOpenHint: "此記錄會保持在倉狀態，直至上載對應的送貨檔案（或日後手動記錄送貨）為止。",
    legacyNoReferralHint: "未有偵測到「Ref Job no.」字句 — 請手動輸入到倉工單號，否則此檔案只會被存檔。",
    legacyNoArrivalFoundHint: (job) => `找不到工單號 ${job} 的存倉記錄 \u2014 請檢查號碼是否正確，或先上載該拆櫃/CFS檔案。此檔案仍會被存檔，但不會記錄任何送貨。`,
  },
};

function inputStyleFor(colors) {
  return {
    fontFamily: FONT_BODY,
    border: `1px solid ${colors.line}`,
    background: colors.surface,
    color: colors.ink,
  };
}
const inputClass = "px-2.5 py-1.5 rounded text-sm outline-none focus:ring-2";

function Badge({ children, tone = "grey", colors }) {
  const map = {
    grey: { bg: colors.surfaceDim, fg: colors.inkFaint },
    amber: { bg: colors.amberSoft, fg: colors.amberText },
    green: { bg: colors.greenSoft, fg: colors.green },
    red: { bg: colors.redSoft, fg: colors.red },
    navy: { bg: colors.navySoft, fg: colors.onDark },
  };
  const c = map[tone];
  return (
    <span
      className="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold tracking-wide"
      style={{ background: c.bg, color: c.fg, fontFamily: FONT_DISPLAY }}
    >
      {children}
    </span>
  );
}

function StatusBadge({ item, colors, t }) {
  const status = deriveStatus(item);
  if (status === "pending_collection") {
    const alert = lfdAlert(item);
    if (alert && alert.level === "overdue") return <Badge tone="red" colors={colors}>{t.badgeLfdOverdue(Math.abs(alert.days))}</Badge>;
    if (alert && alert.level === "soon") return <Badge tone="amber" colors={colors}>{t.badgeLfdSoon(alert.days)}</Badge>;
    return <Badge tone="grey" colors={colors}>{t.badgePendingCollection}</Badge>;
  }
  if (status === "at_depot") {
    const info = storageInfo(item);
    if (info.billable) return <Badge tone="red" colors={colors}>{t.badgeBillable(info.billableDays)}</Badge>;
    const pendingArr = notYetArrivedPackages(item).length;
    if (pendingArr > 0) {
      const arrived = (item.packages || []).length - pendingArr;
      return <Badge tone="amber" colors={colors}>{t.badgePartialArrival(arrived, (item.packages || []).length)}</Badge>;
    }
    const left = info.daysLeft != null ? info.daysLeft : info.freeDays - info.daysHeld;
    return <Badge tone="green" colors={colors}>{t.badgeFree(left)}</Badge>;
  }
  if (status === "partial") {
    const info = storageInfo(item);
    const del = deliveredUnits(item);
    const tot = totalUnits(item);
    return (
      <Badge tone={info.billable ? "red" : "amber"} colors={colors}>
        {t.badgePartial(del, tot, info.billable ? t.badgeBillableSuffix(info.billableDays) : "")}
      </Badge>
    );
  }
  const info = storageInfo(item);
  return <Badge tone="navy" colors={colors}>{t.badgeDelivered(info.billable ? t.badgeBilledSuffix(info.billableDays) : "")}</Badge>;
}

function Field({ label, children, hint, colors }) {
  return (
    <label className="flex flex-col gap-1">
      <span className="text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
        {label}
      </span>
      {children}
      {hint && <span className="text-[11px]" style={{ color: colors.inkFaint }}>{hint}</span>}
    </label>
  );
}

function PackagesEditor({ form, setForm, colors, t }) {
  const [bulkText, setBulkText] = useState("");
  const inputStyle = inputStyleFor(colors);
  const packages = form.packages || [];

  function addCodes() {
    const codes = bulkText.split(/[,\n]/).map((s) => s.trim()).filter(Boolean);
    if (codes.length === 0) return;
    const existing = new Set(packages.map((p) => p.code));
    const additions = codes.filter((c) => !existing.has(c)).map((c) => ({ code: c, description: "", weightKg: "", cbm: "" }));
    setForm((f) => ({ ...f, packages: [...(f.packages || []), ...additions] }));
    setBulkText("");
  }

  function updatePackage(idx, key, value) {
    setForm((f) => {
      const next = [...(f.packages || [])];
      next[idx] = { ...next[idx], [key]: value };
      return { ...f, packages: next };
    });
  }

  function removePackage(idx) {
    setForm((f) => ({ ...f, packages: (f.packages || []).filter((_, i) => i !== idx) }));
  }

  return (
    <div className="mt-2">
      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionPackages}
      </div>
      <p className="text-xs mb-3" style={{ color: colors.inkFaint }}>{t.packagesHint}</p>

      <div className="flex flex-col md:flex-row gap-2 mb-3">
        <textarea
          className={inputClass + " flex-1"}
          style={inputStyle}
          rows={2}
          placeholder={t.bulkAddPlaceholder}
          value={bulkText}
          onChange={(e) => setBulkText(e.target.value)}
        />
        <button
          className="px-4 py-2 rounded text-sm font-semibold h-fit"
          style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={addCodes}
        >
          {t.bulkAddBtn}
        </button>
      </div>

      {packages.length === 0 ? (
        <p className="text-xs" style={{ color: colors.inkFaint }}>{t.noPackagesMsg}</p>
      ) : (
        <div className="rounded overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
          <table className="w-full text-xs" style={{ background: colors.surface }}>
            <thead>
              <tr style={{ background: colors.surfaceDim }}>
                {[t.packageCodeCol, t.packageDescCol, t.packageWeightCol, t.packageCbmCol, ""].map((h) => (
                  <th key={h} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {packages.map((p, idx) => (
                <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-1 py-1">
                    <input className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.code} onChange={(e) => updatePackage(idx, "code", e.target.value)} />
                  </td>
                  <td className="px-1 py-1">
                    <input className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.description} onChange={(e) => updatePackage(idx, "description", e.target.value)} />
                  </td>
                  <td className="px-1 py-1">
                    <input type="number" className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.weightKg} onChange={(e) => updatePackage(idx, "weightKg", e.target.value)} />
                  </td>
                  <td className="px-1 py-1">
                    <input type="number" className="w-full px-1.5 py-1 rounded text-xs" style={inputStyle} value={p.cbm || ""} onChange={(e) => updatePackage(idx, "cbm", e.target.value)} />
                  </td>
                  <td className="px-1 py-1 text-right">
                    <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => removePackage(idx)}>{t.removePackageBtn}</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function ItemForm({ initial, onSave, onCancel, onPrintJobSheet, directory, employees, currentUser, items, colors, t, lang }) {
  const [form, setForm] = useState(initial || { ...emptyForm(), recordedBy: currentUser || "" });
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const inputStyle = inputStyleFor(colors);
  const siteSuggestions = useMemo(() => {
    const fromDirectory = (directory || []).map((s) => s.siteEn).filter(Boolean);
    const fromItems = (items || []).map((i) => i.project).filter(Boolean);
    return [...new Set([...fromDirectory, ...fromItems])];
  }, [directory, items]);

  function applyDirectoryEntry(id) {
    const site = (directory || []).find((s) => s.id === id);
    if (!site) return;
    setForm((f) => ({
      ...f,
      directoryId: site.id,
      client: CLIENTS.includes(site.client) ? site.client : f.client,
      project: site.siteEn || f.project,
      constructionSite: site.siteZh || site.siteEn || f.constructionSite,
      jobRef: site.jobRef || f.jobRef,
      orderedBy: site.orderedBy || f.orderedBy,
    }));
  }

  return (
    <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <h3 className="text-lg font-bold mb-4" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
        {initial ? t.titleEdit : t.titleNew}
      </h3>

      {(directory || []).length > 0 && (
        <div className="mb-4">
          <Field label={t.selectFromDirectory} colors={colors}>
            <select className={inputClass} style={inputStyle} value="" onChange={(e) => applyDirectoryEntry(e.target.value)}>
              <option value="">{t.selectFromDirectoryPlaceholder}</option>
              {directory.map((s) => <option key={s.id} value={s.id}>{s.siteEn} — {s.client}</option>)}
            </select>
          </Field>
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fClient} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.client} onChange={set("client")}>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.fProjectEn} colors={colors}>
          <input list="itemform-site-suggestions" className={inputClass} style={inputStyle} placeholder={t.fProjectPlaceholder} value={form.project} onChange={set("project")} />
          <datalist id="itemform-site-suggestions">
            {siteSuggestions.map((s) => <option key={s} value={s} />)}
          </datalist>
        </Field>
        <Field label={t.fProjectZh} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.constructionSite} onChange={set("constructionSite")} />
        </Field>
        <Field label={t.fInvoiceNo} hint={t.fInvoiceHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.invoiceNumber} onChange={set("invoiceNumber")} />
        </Field>
      </div>

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionCargo}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fItemType} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.itemType} onChange={set("itemType")}>
            {ITEM_TYPES.map((tt) => <option key={tt}>{tt}</option>)}
          </select>
        </Field>
        {form.itemType === "Separate Items" && (!form.packages || form.packages.length === 0) && (
          <Field label={t.fPackages} colors={colors}>
            <input type="number" min="0" className={inputClass} style={inputStyle} value={form.packageCount} onChange={set("packageCount")} />
          </Field>
        )}
        {form.itemType === "Separate Items" && form.packages && form.packages.length > 0 && (
          <Field label={t.fPackages} colors={colors}>
            <div className="px-2.5 py-1.5 rounded text-sm" style={{ ...inputStyle, background: colors.surfaceDim }}>
              {t.packagesCountSummary(form.packages.length)}
            </div>
          </Field>
        )}
        <Field label={t.fUnitCode} hint={t.fUnitCodeHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.unitCode} onChange={set("unitCode")} />
        </Field>

        <Field label={t.fDescription} hint={t.fDescriptionHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.description} onChange={set("description")} />
        </Field>
        <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.shkNumber} onChange={set("shkNumber")} />
        </Field>
        <Field label={t.fSsDoNo} hint={t.fSsDoNoHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.ssDoNo || ""} onChange={set("ssDoNo")} placeholder={'ex ss."SHIP" V.___; CONTAINERS NO. ___'} />
        </Field>

        <Field label={t.fWeight} colors={colors}>
          <input type="number" className={inputClass} style={inputStyle} value={form.weightKg} onChange={set("weightKg")} />
        </Field>
        <Field label={t.fVolume} colors={colors}>
          <input type="number" className={inputClass} style={inputStyle} value={form.volumeCbm} onChange={set("volumeCbm")} />
        </Field>
        <div />

        <Field label={t.f20} colors={colors}>
          <input type="number" min="0" className={inputClass} style={inputStyle} value={form.containers20} onChange={set("containers20")} />
        </Field>
        <Field label={t.f40} colors={colors}>
          <input type="number" min="0" className={inputClass} style={inputStyle} value={form.containers40} onChange={set("containers40")} />
        </Field>
        <div />
      </div>

      {form.itemType === "Separate Items" && (
        <PackagesEditor form={form} setForm={setForm} colors={colors} t={t} />
      )}

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionArrival}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fArrivingType} hint={t.fArrivingTypeHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.arrivingType} onChange={set("arrivingType")}>
            {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
          </select>
        </Field>
        <Field label={t.fDepot} hint={t.fDepotHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.depot} onChange={set("depot")}>
            {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
          </select>
        </Field>
        <Field label={t.fDepotLocation} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.depotLocation} onChange={set("depotLocation")} />
        </Field>

        <Field label={t.fTerminalArrival} hint={t.fTerminalArrivalHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.terminalArrivalDate} onChange={set("terminalArrivalDate")} />
        </Field>
        <Field label={t.fTerminalLFD} hint={t.fTerminalLFDHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.terminalLFD} onChange={set("terminalLFD")} />
        </Field>
        <Field label={t.fConfirmedCollection} hint={t.fConfirmedCollectionHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.confirmedCollectionDate} onChange={set("confirmedCollectionDate")} />
        </Field>

        <Field label={t.fDepotArrival} hint={t.fDepotArrivalHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.depotArrivalDate} onChange={set("depotArrivalDate")} />
        </Field>
        <Field label={t.fPlannedDelivery} hint={t.fPlannedDeliveryHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.plannedDeliveryDate} onChange={set("plannedDeliveryDate")} />
        </Field>
        <div />
        <div />

        {(form.packages || []).length > 0 && (
          <ArrivalBatchesEditor form={form} setForm={setForm} colors={colors} t={t} lang={lang} />
        )}

        {(form.deliveries || []).length > 0 && (
          <div className="col-span-2 md:col-span-3 px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>
            {t.deliveryProgress(deliveredUnits(form), totalUnits(form), form.deliveries.length)}
            <strong>{fmt(lastDeliveryDate(form))}</strong>. {t.deliveryProgressManage}
          </div>
        )}

        <div className="col-span-2 md:col-span-3">
          <Field label={t.fNotes} colors={colors}>
            <textarea className={inputClass} style={inputStyle} rows={2} value={form.notes} onChange={set("notes")} />
          </Field>
        </div>
      </div>

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionJobSheet}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
          <div className="flex gap-2">
            <input className={inputClass + " flex-1"} style={inputStyle} value={form.jobNumber} onChange={set("jobNumber")} />
            <button
              type="button"
              className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => setForm((f) => ({ ...f, jobNumber: nextJobNumber(items) }))}
            >
              {t.generateJobNoBtn}
            </button>
          </div>
        </Field>
        <Field label={t.fOrderedBy} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.orderedBy} onChange={set("orderedBy")} />
        </Field>
        <Field label={t.fPoNumber} hint={t.fPoNumberHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.poNumber} onChange={set("poNumber")} />
        </Field>
        <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.jobRef} onChange={set("jobRef")} />
        </Field>
        <Field label={t.fRecordedBy} hint={t.fRecordedByHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.recordedBy} onChange={set("recordedBy")}>
            <option value=""></option>
            {(employees || []).map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
          </select>
        </Field>
        {form.jobNumber && (
          <div className="col-span-2 md:col-span-3">
            <button
              type="button"
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ border: `1px solid ${colors.line}`, color: colors.amberText, fontFamily: FONT_DISPLAY }}
              onClick={() => onPrintJobSheet({ type: form.arrivingType, item: form })}
            >
              {t.printJobSheetBtn}
            </button>
          </div>
        )}
      </div>

      {form.depotArrivalDate && !form.recordedBy && (
        <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
          {t.recordedByRequiredMsg}
        </div>
      )}

      <div className="flex gap-2 mt-5">
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
          onClick={() => {
            if (!form.project || !form.description) return;
            if (form.depotArrivalDate && !form.recordedBy) return;
            onSave(form);
          }}
        >
          {t.saveBtn}
        </button>
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={onCancel}
        >
          {t.cancelBtn}
        </button>
      </div>
    </div>
  );
}

function ArrivalBatchesEditor({ form, setForm, colors, t, lang }) {
  const inputStyle = inputStyleFor(colors);
  const [draft, setDraft] = useState({ date: todayStr(), type: ARRIVING_TYPES[0], codes: [] });
  const arrivals = form.arrivals || [];
  const assigned = new Set(arrivals.flatMap((a) => a.codes || []));
  const unassigned = (form.packages || []).filter((p) => !assigned.has(p.code));

  function toggleDraftCode(code) {
    setDraft((d) => ({ ...d, codes: d.codes.includes(code) ? d.codes.filter((c) => c !== code) : [...d.codes, code] }));
  }
  function syncEarliestDate(list) {
    const ds = list.map((a) => a.date).filter(Boolean).sort();
    return ds.length ? ds[0] : "";
  }
  function addBatch() {
    if (!draft.date || draft.codes.length === 0) return;
    const next = [...arrivals, { id: `arr-${Date.now()}`, date: draft.date, type: draft.type, codes: draft.codes }];
    setForm((f) => ({ ...f, arrivals: next, depotArrivalDate: syncEarliestDate(next) }));
    setDraft({ date: todayStr(), type: ARRIVING_TYPES[0], codes: [] });
  }
  function removeBatch(id) {
    const next = arrivals.filter((a) => a.id !== id);
    setForm((f) => ({ ...f, arrivals: next, depotArrivalDate: next.length ? syncEarliestDate(next) : f.depotArrivalDate }));
  }

  return (
    <div className="col-span-2 md:col-span-3 rounded p-3" style={{ border: `1px dashed ${colors.line}` }}>
      <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
        {t.splitArrivalLabel}
      </div>
      <div className="text-[11px] mb-2" style={{ color: colors.inkFaint }}>{t.splitArrivalHint}</div>

      {arrivals.length > 0 && (
        <div className="mb-3 rounded overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
          <table className="w-full text-xs" style={{ background: colors.surface }}>
            <thead>
              <tr style={{ background: colors.surfaceDim }}>
                {[t.colDate, t.fArrivingType, t.splitArrivalCasesCol, ""].map((h, idx) => (
                  <th key={idx} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...arrivals].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((a) => (
                <tr key={a.id} style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{fmt(a.date)}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{a.type}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{(a.codes || []).join(", ")}</td>
                  <td className="px-2 py-1.5 text-right">
                    <button type="button" className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => removeBatch(a.id)}>{t.deleteBtn}</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {unassigned.length > 0 ? (
        <div className="flex flex-col gap-2">
          <div className="flex flex-wrap items-end gap-3">
            <Field label={t.colDate} colors={colors}>
              <input type="date" className={inputClass} style={inputStyle} value={draft.date} onChange={(e) => setDraft((d) => ({ ...d, date: e.target.value }))} />
            </Field>
            <Field label={t.fArrivingType} colors={colors}>
              <select className={inputClass} style={inputStyle} value={draft.type} onChange={(e) => setDraft((d) => ({ ...d, type: e.target.value }))}>
                {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
              </select>
            </Field>
            <button
              type="button"
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: draft.codes.length === 0 ? 0.5 : 1 }}
              disabled={draft.codes.length === 0}
              onClick={addBatch}
            >
              {t.splitArrivalAddBtn}
            </button>
          </div>
          <div className="text-[11px]" style={{ color: colors.inkFaint }}>{t.splitArrivalSelectHint}</div>
          <div className="flex flex-wrap gap-2">
            {unassigned.map((p) => (
              <button
                key={p.code}
                type="button"
                onClick={() => toggleDraftCode(p.code)}
                className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                style={{
                  border: `1px solid ${draft.codes.includes(p.code) ? colors.amber : colors.line}`,
                  background: draft.codes.includes(p.code) ? colors.amberSoft : colors.surface,
                  color: draft.codes.includes(p.code) ? colors.amberText : colors.ink,
                }}
                title={p.description}
              >
                {p.code}
              </button>
            ))}
          </div>
        </div>
      ) : arrivals.length > 0 ? (
        <div className="text-xs" style={{ color: colors.green }}>{t.splitArrivalAllAssignedMsg}</div>
      ) : null}
    </div>
  );
}

function DeliveryForm({ deliveryItems, onAddDelivery, onAddCombinedDelivery, onDeleteDelivery, onCancel, onPrintJobSheet, employees, currentUser, items, colors, t, lang }) {
  const firstItem = deliveryItems[0];
  const [extraItemIds, setExtraItemIds] = useState([]);
  const [addPickerOpen, setAddPickerOpen] = useState(false);
  const [addSearch, setAddSearch] = useState("");
  const baseIds = new Set(deliveryItems.map((i) => i.id));
  const allDeliveryItems = [...deliveryItems, ...items.filter((i) => extraItemIds.includes(i.id) && !baseIds.has(i.id))];
  const isCombined = allDeliveryItems.length > 1;
  const addCandidates = useMemo(() => {
    const excluded = new Set(allDeliveryItems.map((i) => i.id));
    const siteKey = (i) => (i.constructionSite || i.project || "").trim().toLowerCase();
    const targetSite = siteKey(firstItem);
    return items.filter((i) => !excluded.has(i.id) && ["at_depot", "partial"].includes(deriveStatus(i)) && (!targetSite || siteKey(i) === targetSite));
  }, [items, allDeliveryItems, firstItem]);
  const addFiltered = addSearch.trim()
    ? addCandidates.filter((i) => {
        const q = addSearch.toLowerCase();
        return i.unitCode?.toLowerCase().includes(q) || i.jobNumber?.toLowerCase().includes(q) || i.id?.toLowerCase().includes(q);
      })
    : addCandidates;

  const [form, setForm] = useState({
    date: todayStr(),
    deliveredTo: firstItem.constructionSite || firstItem.project || "",
    receivedBy: "",
    notes: "",
    jobNumber: "",
    recordedBy: currentUser || "",
  });
  const [perItem, setPerItem] = useState({});
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const inputStyle = inputStyleFor(colors);

  function getSel(itemId) {
    return perItem[itemId] || { codes: [], qty: "" };
  }
  function setCodes(itemId, updater) {
    setPerItem((prev) => {
      const cur = prev[itemId] || { codes: [], qty: "" };
      const nextCodes = typeof updater === "function" ? updater(cur.codes) : updater;
      return { ...prev, [itemId]: { ...cur, codes: nextCodes } };
    });
  }
  function setQty(itemId, qty) {
    setPerItem((prev) => ({ ...prev, [itemId]: { ...(prev[itemId] || { codes: [] }), qty } }));
  }
  function toggleCode(itemId, code) {
    setCodes(itemId, (prev) => (prev.includes(code) ? prev.filter((c) => c !== code) : [...prev, code]));
  }
  function toggleBatch(itemId, batch) {
    setCodes(itemId, (prev) => {
      const allSelected = batch.remainingCodes.every((c) => prev.includes(c));
      if (allSelected) return prev.filter((c) => !batch.remainingCodes.includes(c));
      return [...new Set([...prev, ...batch.remainingCodes])];
    });
  }
  function addBatchToDelivery(itemId) {
    setExtraItemIds((prev) => (prev.includes(itemId) ? prev : [...prev, itemId]));
    setAddSearch("");
    setAddPickerOpen(false);
  }
  function removeExtraBatch(itemId) {
    setExtraItemIds((prev) => prev.filter((id) => id !== itemId));
    setPerItem((prev) => { const next = { ...prev }; delete next[itemId]; return next; });
  }

  const anySelected = allDeliveryItems.some((it) => {
    const itemized = (it.packages || []).length > 0;
    const sel = getSel(it.id);
    return itemized ? sel.codes.length > 0 : Number(sel.qty) > 0;
  });
  const anyOvershoot = allDeliveryItems.some((it) => {
    const itemized = (it.packages || []).length > 0;
    if (itemized) return false;
    const sel = getSel(it.id);
    return Number(sel.qty) > remainingUnits(it);
  });

  function handleAddClick() {
    const jobNumber = form.jobNumber || nextJobNumber(items);
    const deliveredTo = form.deliveredTo || firstItem.constructionSite || firstItem.project || "";
    const receivedBy = form.receivedBy || firstItem.orderedBy || "";
    const entries = [];
    for (const it of allDeliveryItems) {
      const itemized = (it.packages || []).length > 0;
      const sel = getSel(it.id);
      if (itemized) {
        if (sel.codes.length === 0) continue;
        entries.push({ itemId: it.id, delivery: { date: form.date, deliveredTo, receivedBy, jobNumber, recordedBy: form.recordedBy, notes: form.notes, codes: sel.codes } });
      } else {
        const qty = Number(sel.qty) || 0;
        if (qty <= 0) continue;
        entries.push({ itemId: it.id, delivery: { date: form.date, deliveredTo, receivedBy, jobNumber, recordedBy: form.recordedBy, notes: form.notes, packageCount: qty } });
      }
    }
    if (entries.length === 0) return;
    setPerItem({});
    setExtraItemIds([]);
    setForm({ date: todayStr(), deliveredTo: firstItem.constructionSite || firstItem.project || "", receivedBy: "", notes: "", jobNumber: "", recordedBy: currentUser || "" });
    if (entries.length === 1 && !isCombined) {
      const record = onAddDelivery(entries[0].delivery, entries[0].itemId);
      onPrintJobSheet({ type: "Delivery", item: firstItem, delivery: record || entries[0].delivery });
    } else {
      const recorded = onAddCombinedDelivery(entries);
      const groups = recorded.map(({ itemId, record }) => ({ item: allDeliveryItems.find((i) => i.id === itemId), delivery: record }));
      onPrintJobSheet({ type: "Delivery", combined: true, groups, jobNumber, date: form.date, deliveredTo, receivedBy });
    }
  }

  return (
    <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
        {isCombined ? t.combinedDeliveryTitle(allDeliveryItems.length) : `${t.deliveryTitlePrefix} ${firstItem.id}`}
      </h3>
      <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>
        {isCombined
          ? allDeliveryItems.map((it) => `${it.unitCode || it.id}`).join(" · ")
          : <>{firstItem.client} · {firstItem.project}{firstItem.plannedDeliveryDate ? t.plannedWasText(fmt(firstItem.plannedDeliveryDate)) : ""}</>}
      </p>

      {!isCombined && activeDeliveries(firstItem).length > 0 && (
        <div className="mb-4 rounded overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
          <table className="w-full text-xs" style={{ background: colors.surface }}>
            <thead>
              <tr style={{ background: colors.surfaceDim }}>
                {[t.colDate, t.colQty, t.colDeliveredTo, t.colReceivedBy, t.colJobNo, ""].map((h) => (
                  <th key={h} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {[...activeDeliveries(firstItem)].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((d) => (
                <tr key={d.id} style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{fmt(d.date)}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.codes ? d.codes.join(", ") : d.packageCount}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.deliveredTo || "—"}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.receivedBy || "—"}</td>
                  <td className="px-2 py-1.5" style={{ fontFamily: FONT_MONO, color: colors.ink }}>{d.jobNumber || "—"}</td>
                  <td className="px-2 py-1.5 text-right whitespace-nowrap">
                    <button className="text-xs font-semibold mr-2" style={{ color: colors.amberText }} onClick={() => onPrintJobSheet({ type: "Delivery", item: firstItem, delivery: d })}>{t.printBtn}</button>
                    <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => onDeleteDelivery(d.id, firstItem.id)}>{t.cancelJobBtn}</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {isCombined && (
        <div className="mb-4 px-3 py-2 rounded text-xs" style={{ background: colors.surfaceDim, color: colors.inkFaint }}>
          {t.combinedHistoryHiddenNote}
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fDeliveryDate} hint={t.fDeliveryDateHint} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={form.date} onChange={set("date")} />
        </Field>
        <div />
        <div />
        <Field label={t.fDeliveredTo} hint={t.fDeliveredToHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.deliveredTo} onChange={set("deliveredTo")} />
        </Field>
        <Field label={t.fReceivedBy} hint={t.fReceivedByHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.receivedBy} onChange={set("receivedBy")} />
        </Field>
        <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
          <div className="flex gap-2">
            <input className={inputClass + " flex-1"} style={inputStyle} value={form.jobNumber} onChange={set("jobNumber")} />
            <button
              type="button"
              className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => setForm((f) => ({ ...f, jobNumber: nextJobNumber(items) }))}
            >
              {t.generateJobNoBtn}
            </button>
          </div>
        </Field>
        <Field label={t.fRecordedBy} hint={t.fRecordedByHint} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.recordedBy} onChange={set("recordedBy")}>
            <option value=""></option>
            {(employees || []).map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
          </select>
        </Field>
        <div className="col-span-2 md:col-span-3">
          <Field label={t.fNotes} colors={colors}>
            <textarea className={inputClass} style={inputStyle} rows={2} value={form.notes} onChange={set("notes")} />
          </Field>
        </div>
      </div>

      {allDeliveryItems.map((it) => {
        const remaining = remainingUnits(it);
        const multiUnit = totalUnits(it) > 1;
        const itemized = (it.packages || []).length > 0;
        const remainingPkgs = remainingPackages(it);
        const pendingArrival = new Set(notYetArrivedPackages(it).map((p) => p.code));
        const remainingCodesSet = new Set(remainingPkgs.map((p) => p.code));
        const selectableBatches = usesArrivalBatches(it)
          ? activeArrivals(it).map((b) => ({ ...b, remainingCodes: (b.codes || []).filter((c) => remainingCodesSet.has(c)) })).filter((b) => b.remainingCodes.length > 0)
          : [];
        const sel = getSel(it.id);

        if (remaining <= 0) {
          return (
            <div key={it.id} className="mt-4 px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>
              {(it.unitCode || it.id)}: {itemized ? t.noCodesRemainingMsg : t.allDeliveredMsg}
            </div>
          );
        }

        return (
          <div key={it.id} className="mt-5 pt-4" style={{ borderTop: `1px solid ${colors.line}` }}>
            {isCombined && (
              <div className="text-sm font-bold mb-2" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
                {it.unitCode || it.id} · {it.client} · {it.project}{it.jobNumber ? ` · ${t.colJobNo}: ${it.jobNumber}` : ""}
              </div>
            )}
            {itemized ? (
              <div>
                {selectableBatches.length > 1 && (
                  <div className="mb-3">
                    <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                      {t.selectByBatchLabel}
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {selectableBatches.map((b) => {
                        const allSelected = b.remainingCodes.every((c) => sel.codes.includes(c));
                        const someSelected = !allSelected && b.remainingCodes.some((c) => sel.codes.includes(c));
                        return (
                          <button
                            key={b.id}
                            type="button"
                            onClick={() => toggleBatch(it.id, b)}
                            className="px-3 py-1.5 rounded text-xs font-semibold text-left"
                            style={{
                              border: `1px solid ${allSelected || someSelected ? colors.amber : colors.line}`,
                              background: allSelected ? colors.amberSoft : someSelected ? colors.surfaceDim : colors.surface,
                              color: allSelected || someSelected ? colors.amberText : colors.ink,
                            }}
                          >
                            {fmt(b.date)}{b.type ? ` · ${b.type}` : ""} · {b.remainingCodes.length} {t.jsPkgs}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
                <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                  {t.selectCodesLabel}
                </div>
                <div className="flex flex-wrap gap-2">
                  {remainingPkgs.map((p) => {
                    const notHere = pendingArrival.has(p.code);
                    if (notHere) {
                      return (
                        <span
                          key={p.code}
                          className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                          style={{ border: `1px dashed ${colors.line}`, background: colors.surfaceDim, color: colors.inkFaint, cursor: "not-allowed" }}
                          title={t.notYetArrivedHint}
                        >
                          {p.code} · {t.notYetArrivedTag}
                        </span>
                      );
                    }
                    return (
                    <button
                      key={p.code}
                      type="button"
                      onClick={() => toggleCode(it.id, p.code)}
                      className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                      style={{
                        border: `1px solid ${sel.codes.includes(p.code) ? colors.amber : colors.line}`,
                        background: sel.codes.includes(p.code) ? colors.amberSoft : colors.surface,
                        color: sel.codes.includes(p.code) ? colors.amberText : colors.ink,
                      }}
                      title={p.description}
                    >
                      {p.code}{p.description ? ` — ${p.description}` : ""}
                    </button>
                    );
                  })}
                </div>
                {pendingArrival.size > 0 && (
                  <div className="mt-2 px-3 py-2 rounded text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
                    {t.pendingArrivalNotice(pendingArrival.size)}
                  </div>
                )}
              </div>
            ) : (
              <div className="max-w-xs">
                {multiUnit ? (
                  <Field label={t.fQty} hint={t.fQtyHint(remaining)} colors={colors}>
                    <input type="number" min="1" max={remaining} className={inputClass} style={inputStyle} value={sel.qty} onChange={(e) => setQty(it.id, e.target.value)} />
                  </Field>
                ) : (
                  <button
                    type="button"
                    className="px-2.5 py-1.5 rounded text-xs font-semibold"
                    style={{
                      border: `1px solid ${Number(sel.qty) > 0 ? colors.amber : colors.line}`,
                      background: Number(sel.qty) > 0 ? colors.amberSoft : colors.surface,
                      color: Number(sel.qty) > 0 ? colors.amberText : colors.ink,
                    }}
                    onClick={() => setQty(it.id, Number(sel.qty) > 0 ? "" : "1")}
                  >
                    {Number(sel.qty) > 0 ? t.selectedTag : t.includeInDeliveryBtn}
                  </button>
                )}
                {Number(sel.qty) > remaining && (
                  <div className="mt-2 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
                    {t.overshootMsg(remaining)}
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}

      <div className="mt-5 pt-4" style={{ borderTop: `1px solid ${colors.line}` }}>
        {extraItemIds.length > 0 && (
          <div className="flex flex-wrap gap-2 mb-3">
            {items.filter((i) => extraItemIds.includes(i.id)).map((i) => (
              <span key={i.id} className="inline-flex items-center gap-2 px-2.5 py-1.5 rounded text-xs font-semibold" style={{ background: colors.amberSoft, color: colors.amberText, border: `1px solid ${colors.amber}` }}>
                {i.unitCode || i.id} · {i.client}
                <button type="button" onClick={() => removeExtraBatch(i.id)} style={{ color: colors.amberText, fontWeight: "bold" }}>&times;</button>
              </span>
            ))}
          </div>
        )}
        {!addPickerOpen ? (
          <button
            type="button"
            className="px-3 py-2 rounded text-sm font-semibold"
            style={{ border: `1px dashed ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
            onClick={() => setAddPickerOpen(true)}
          >
            {t.addMoreBatchesBtn}
          </button>
        ) : (
          <div className="rounded p-3" style={{ border: `1px solid ${colors.line}`, background: colors.surfaceDim }}>
            <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>
              {t.addMoreBatchesScopeNote(firstItem.constructionSite || firstItem.project || "")}
            </div>
            <div className="flex items-center gap-2 mb-2">
              <input
                autoFocus
                className={inputClass}
                style={{ ...inputStyle, flex: 1 }}
                placeholder={t.addMoreBatchesSearchPlaceholder}
                value={addSearch}
                onChange={(e) => setAddSearch(e.target.value)}
              />
              <button type="button" className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => { setAddPickerOpen(false); setAddSearch(""); }}>{t.closeBtn}</button>
            </div>
            <div className="flex flex-col gap-1 max-h-56 overflow-y-auto">
              {addFiltered.length === 0 && (
                <div className="text-xs px-2 py-2" style={{ color: colors.inkFaint }}>{t.addMoreBatchesNoneMsg}</div>
              )}
              {addFiltered.slice(0, 30).map((i) => (
                <button
                  key={i.id}
                  type="button"
                  onClick={() => addBatchToDelivery(i.id)}
                  className="text-left px-2.5 py-1.5 rounded text-xs"
                  style={{ background: colors.surface, color: colors.ink, border: `1px solid ${colors.line}` }}
                >
                  <span className="font-semibold">{i.unitCode || i.id}</span> · {i.client} · {i.constructionSite || i.project}
                </button>
              ))}
            </div>
          </div>
        )}
      </div>

      {!form.recordedBy && (
        <div className="mt-4 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
          {t.recordedByRequiredMsg}
        </div>
      )}

      <div className="flex gap-2 mt-5">
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
          disabled={!anySelected || anyOvershoot || !form.date || !form.recordedBy}
          onClick={handleAddClick}
        >
          {isCombined ? t.addCombinedDeliveryBtn : t.addDeliveryBtn}
        </button>
        <button
          className="px-4 py-2 rounded text-sm font-semibold"
          style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
          onClick={onCancel}
        >
          {t.closeBtn}
        </button>
      </div>
    </div>
  );
}

const JOB_SHEET_TEMPLATES = ["Devan", "CFS", "Delivery", "Shifting", "Hoisting", "Day Work", "Dismantle", "Dis & Removal of Lifting Tools", "Job Cancel", "Pick-up", "Position", "Re-position", "Retain of Safety Ropes"];
const JOB_SHEET_ITEMIZED = ["Devan", "CFS", "Delivery"];
const CFS_FROM_PRESETS = [
  { key: "金田物流", text: "金田物流有限公司\n油麻地海輝道38號新油麻地貨物公眾貨物\n裝卸區7號 3-4口水位高鴻躉船\n(車入閘后轉右直行約350米)\nTEL: 27850666 (亞昌/ 珍珍)" },
  { key: "悅昇物流", text: "悅昇物流有限公司\n香港新界葵涌昂船洲479號\nTEL: 24974884" },
];
const JOB_SHEET_BODY_SEEDS = {
  "Day Work": "- 按客戶要求將扶梯配件(共: _______箱) 由暫存位置攝出供客人開箱檢查貨件、\n  完成後用回原有包裝包妥，並攝回存倉位置擺放。",
  "Dismantle": "DISMANTLE WORK FOR ESCALATOR NO. ______\n\n1) 於下夜將所有吊重工具送交工地並移入所屬梯井。\n2) 將吊重工具掛上吊碼上。\n3) 將扶梯配件拆走。\n4) 將梯架吊離梯井。\n5) 用等鋰子切割機將扶梯切割至可移走呎吋。\n6) 機房內設置/搭建一套通架吊台供吊走摩打台。\n7) 出口外設置攝貨平台及圍封。\n8) 收走梯架及機房台。",
  "Dis & Removal of Lifting Tools": "FOR ESCALATOR NO. ______\n\n1/ 將有關工作台通架料於下夜送交工地及搬入______梯圍板內。\n2/ 於梯面搭建工作台將所有吊梯工具拆下並搬出站外上貨位。\n3/ 收離工地。",
  "Job Cancel": "- 因______，按______要求停止進行______工作。",
  "Pick-up": "- 客人安排運輸到快達倉取走全部______。(共: ______ 箱)",
  "Position": "POSITION WORK FOR ESCALATOR NO. ______\n\n1/ 設置吊重工具，以供吊嵌扶梯之用。\n2/ 吊臂車將扶梯及全部配件運到地盤。\n3/ 將扶梯攝入所屬梯井，並進行接駁、吊崁及安放到指定梯井位內。",
  "Re-position": "於______提供______(共______名)吊運人員將以上扶梯重新吊回梯井座好。",
  "Retain of Safety Ropes": "1) 於梯井設置羊眼圈及安全繩(共______組)供我司拆配件及吊走扶梯之用。\n2) 我司吊運工作完成後按客人要求保留梯井內(共______組)安全繩連羊眼圈供安裝工作之用。",
  "Shifting": "- 提供人手及工具將______攝到______暫存位置擺放。\n(TOTAL: ______ PKG / ______ KGS / ______ CBM)",
  "Hoisting": "HOISTING WORK FOR ______\n\n1/ ",
};

async function compressFileToDataUri(file, maxDim = 1400, quality = 0.72) {
  const readAsDataUrl = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(new Error("read-failed")); r.readAsDataURL(f); });
  if (file.type === "application/pdf") {
    if (file.size > 3.2 * 1024 * 1024) throw new Error("too-large");
    return await readAsDataUrl(file);
  }
  const dataUrl = await readAsDataUrl(file);
  const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error("bad-image")); i.src = dataUrl; });
  const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.width * scale));
  canvas.height = Math.max(1, Math.round(img.height * scale));
  canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/jpeg", quality);
}

function SignedDocControl({ docKey, colors, t }) {
  const [status, setStatus] = useState("loading");
  const [busy, setBusy] = useState(false);
  useEffect(() => {
    let on = true;
    (async () => {
      try {
        const r = await storageGet(docKey);
        if (on) setStatus(r ? "saved" : "none");
      } catch (e) {
        if (on) setStatus("none");
      }
    })();
    return () => { on = false; };
  }, [docKey]);

  async function handleFile(e) {
    const f = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setBusy(true);
    try {
      const uri = await compressFileToDataUri(f);
      const ok = await storageSet(docKey, JSON.stringify({ uri, name: f.name, at: todayStr() }));
      if (!ok) throw new Error("save-failed");
      setStatus("saved");
    } catch (err) {
      alert(t.signedDocFailMsg);
    }
    setBusy(false);
  }

  async function view() {
    try {
      const r = await storageGet(docKey);
      if (!r) return;
      const { uri, name } = JSON.parse(r.value);
      const w = window.open("", "_blank");
      if (!w) return;
      if (uri.startsWith("data:application/pdf")) {
        w.document.write(`<title>${name || "Signed copy"}</title><embed src="${uri}" type="application/pdf" style="width:100%;height:100vh;">`);
      } else {
        w.document.write(`<title>${name || "Signed copy"}</title><body style="margin:0;background:#333;display:flex;justify-content:center;"><img src="${uri}" style="max-width:100%;height:auto;"></body>`);
      }
    } catch (e) { /* noop */ }
  }

  if (status === "loading") return <span className="text-xs" style={{ color: colors.inkFaint }}>…</span>;
  return (
    <span className="inline-flex items-center gap-1.5">
      {status === "saved" && (
        <button type="button" className="text-xs font-semibold underline" style={{ color: colors.green }} onClick={view}>
          ✓ {t.signedDocViewBtn}
        </button>
      )}
      <label className="text-xs font-semibold underline cursor-pointer" style={{ color: colors.amberText, opacity: busy ? 0.5 : 1 }}>
        {busy ? t.signedDocSavingMsg : status === "saved" ? t.signedDocReplaceBtn : t.signedDocUploadBtn}
        <input type="file" accept="image/*,application/pdf" style={{ display: "none" }} onChange={handleFile} disabled={busy} />
      </label>
    </span>
  );
}


const OVERSIZE_RULES = {
  Schindler: [
    { min: 3.25, max: 5.50, mult: 1.5 },
    { min: 5.501, max: 7.50, mult: 2.0 },
    { min: 7.501, max: Infinity, mult: 2.5 },
  ],
  Chevalier: [
    { min: 4.0, max: 6.50, mult: 2.0 },
    { min: 6.501, max: 8.00, mult: 2.5 },
    { min: 8.001, max: Infinity, mult: 3.0 },
  ],
};
function oversizeTierFor(client, cbm) {
  const rules = OVERSIZE_RULES[client];
  if (!rules || !(cbm > 0)) return null;
  for (const r of rules) {
    if (cbm >= r.min && cbm <= r.max) return r.mult;
  }
  return null;
}
function computeOversizeText(item) {
  const rules = OVERSIZE_RULES[item.client];
  if (!rules || !(item.packages || []).length) return "";
  const byMult = new Map();
  for (const p of item.packages) {
    const cbm = Number(p.cbm);
    const mult = oversizeTierFor(item.client, cbm);
    if (!mult) continue;
    if (!byMult.has(mult)) byMult.set(mult, []);
    byMult.get(mult).push({ code: p.code, cbm });
  }
  if (byMult.size === 0) return "";
  const lines = [];
  const sortedMults = [...byMult.keys()].sort((a, b) => a - b);
  for (const mult of sortedMults) {
    const entries = byMult.get(mult);
    const byCbm = new Map();
    for (const e of entries) {
      const key = String(e.cbm);
      if (!byCbm.has(key)) byCbm.set(key, []);
      byCbm.get(key).push(e.code);
    }
    for (const [cbmVal, codes] of byCbm) {
      lines.push(`#${codes.join(",")} @${cbmVal}CBM`);
    }
    const subtotal = Math.round(entries.reduce((s, e) => s + e.cbm, 0) * 100) / 100;
    lines.push(`(合共 Total: ${subtotal} CBM)  x${mult} 倍`);
    lines.push("");
  }
  return lines.join("\n").trim();
}

function deliveryTotalsFor(it, del) {
  const pkgs = del.codes ? del.codes.length : Number(del.packageCount) || 0;
  let kgs = "", cbm = "", estimated = false;
  if (del.codes && (it.packages || []).length > 0) {
    const delivered = it.packages.filter((p) => del.codes.includes(p.code));
    const haveAllWeights = delivered.length > 0 && delivered.every((p) => p.weightKg !== "" && p.weightKg != null);
    const haveAllCbm = delivered.length > 0 && delivered.every((p) => p.cbm !== "" && p.cbm != null);
    kgs = haveAllWeights ? String(Math.round(delivered.reduce((s, p) => s + Number(p.weightKg), 0) * 10) / 10) : "";
    cbm = haveAllCbm ? String(Math.round(delivered.reduce((s, p) => s + Number(p.cbm), 0) * 1000) / 1000) : "";
    if (!haveAllWeights || !haveAllCbm) estimated = true;
  } else {
    estimated = true;
  }
  if (estimated) {
    const totalU = totalUnits(it) || 1;
    const share = pkgs / totalU;
    if (!kgs && it.weightKg) kgs = `~${Math.round(Number(it.weightKg) * share * 10) / 10}`;
    if (!cbm && it.volumeCbm) cbm = `~${Math.round(Number(it.volumeCbm) * share * 1000) / 1000}`;
  }
  return { pkgs, kgs, cbm, estimated };
}
function refTextFor(it, del) {
  const jn = it.jobNumber || "";
  let dates = [];
  if (usesArrivalBatches(it) && del && del.codes) {
    const set = new Set(del.codes);
    dates = [...new Set(activeArrivals(it).filter((a) => (a.codes || []).some((c) => set.has(c))).map((a) => a.date).filter(Boolean))];
  }
  if (dates.length === 0) dates = [effectiveDepotArrivalDate(it)].filter(Boolean);
  if (dates.length === 0 || !jn) return "";
  return dates.map((d) => `Refer to job no. ${jn} on ${fmt(d)}`).join("\n");
}

function CombinedDeliveryPrint({ sheet, onClose, directory, colors, t, lang }) {
  const { groups, jobNumber, date, deliveredTo } = sheet;
  const firstItem = groups[0].item;
  const depotEn = firstItem.depot || "";
  const depotZh = DEPOT_LABELS_ZH[firstItem.depot] || "";
  const dirMatch = (directory || []).find((s) => s.siteEn && (s.siteEn === firstItem.constructionSite || s.siteEn === firstItem.project));
  const siteZh = (dirMatch && dirMatch.siteZh) || "";
  const siteContact = dirMatch ? [dirMatch.contactName, dirMatch.contactPhone].filter(Boolean).join(" ") : "";

  const site = firstItem.constructionSite || firstItem.project || deliveredTo || "";
  const toTop = [site ? (/^site\s+at/i.test(site) ? site : `SITE AT ${site}`) : "", siteZh].filter(Boolean).join("\n");
  const toBottom = siteContact;

  const rows = groups.map((g) => ({ ...g, totals: deliveryTotalsFor(g.item, g.delivery), ref: refTextFor(g.item, g.delivery) }));
  const anyEstimated = rows.some((r) => r.totals.estimated);
  const totalPkgs = rows.reduce((s, r) => s + r.totals.pkgs, 0);
  const numOrNull = (v) => { const n = Number(String(v || "").replace("~", "")); return isNaN(n) ? null : n; };
  const kgVals = rows.map((r) => numOrNull(r.totals.kgs));
  const cbmVals = rows.map((r) => numOrNull(r.totals.cbm));
  const totalKgs = kgVals.every((v) => v != null) ? Math.round(kgVals.reduce((s, v) => s + v, 0) * 10) / 10 : "";
  const totalCbm = cbmVals.every((v) => v != null) ? Math.round(cbmVals.reduce((s, v) => s + v, 0) * 1000) / 1000 : "";

  const lbl = { border: "1px solid #111", fontWeight: "bold", padding: "4px 6px", verticalAlign: "top", width: 78, fontSize: 13, wordBreak: "break-word", overflowWrap: "break-word" };
  const cel = { border: "1px solid #111", padding: 6, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };
  const cel0 = { border: "1px solid #111", padding: 0, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };

  return (
    <div id="job-sheet-print-root" className="fixed inset-0 z-50 flex flex-col" style={{ background: "rgba(0,0,0,0.5)" }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #job-sheet-print-root, #job-sheet-print-root * { visibility: visible; }
          #job-sheet-print-root { position: absolute; inset: auto; left: 0; top: 0; width: 100%; background: #fff !important; }
          #job-sheet-print-toolbar { display: none !important; }
          #job-sheet-print-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          #job-sheet-print-area { margin: 0 !important; max-width: 100% !important; }
        }
      `}</style>
      <div id="job-sheet-print-toolbar" className="flex flex-wrap items-center gap-2 p-3" style={{ background: colors.navy }}>
        <span className="text-sm font-semibold" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>{t.combinedPrintLabel}</span>
        <div className="ml-auto flex gap-2">
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => window.print()}>
            {t.printBtn}
          </button>
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.onDark}`, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>
            {t.closePreviewBtn}
          </button>
        </div>
      </div>
      <div id="job-sheet-print-scroll" className="flex-1 overflow-y-auto p-6" style={{ background: colors.bg }}>
        <div id="job-sheet-print-area" className="mx-auto" style={{ background: "#fff", color: "#111", maxWidth: 700, padding: 24, fontFamily: "Arial, sans-serif", fontSize: 15 }}>

          <div className="flex items-start justify-between" style={{ marginBottom: 8 }}>
            <div className="flex items-center gap-3">
              <img src={FARSPEED_LOGO_DATA_URI} alt="Farspeed" style={{ height: 78, width: "auto" }} />
              <div>
                <div style={{ fontSize: 24, fontWeight: 700, color: "#111", lineHeight: 1.1 }}>FARSPEED Contractors Limited</div>
                <div style={{ fontSize: 11, color: "#111" }}>P. O. Box No. 1985, Yuen Long Post Office, Yuen Long, N.T., Hong Kong</div>
                <div style={{ fontSize: 11, color: "#111" }}>Tel: +852 5337-9500&nbsp;&nbsp;Fax: +852 2402-4450&nbsp;&nbsp;http://www.farspeed.hk</div>
              </div>
            </div>
            <div className="flex items-start" style={{ fontSize: 16, fontWeight: 700, lineHeight: 1.5 }}>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3 }}>有限公司</div>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3, marginLeft: 4 }}>快達承判</div>
            </div>
          </div>
          <div style={{ borderTop: "2.5px solid #111", marginBottom: 10 }} />

          <div className="text-center font-bold mb-3" style={{ fontSize: 22, letterSpacing: 6 }}>
            {t.jsTitleZh}&nbsp;&nbsp;{t.jsTitle}
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18 }}>
            <tbody>
              <tr>
                <td style={lbl}>{t.jsFromZh}<br />{t.jsFrom}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: "1px solid #111" }}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20 }}>{[depotZh, depotEn].filter(Boolean).join("\n")}</div>
                </td>
                <td style={lbl} rowSpan={toBottom ? 1 : 1}>{t.jsToZh}<br />{t.jsTo}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: toBottom ? "none" : "1px solid #111" }} rowSpan={toBottom ? 1 : 1}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{toTop}</div>
                </td>
              </tr>
              {toBottom && (
                <tr>
                  <td style={lbl}></td>
                  <td style={cel0}></td>
                  <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderBottom: "1px solid #111", borderTop: "none" }}>
                    <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{toBottom}</div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>

          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18, marginTop: -1 }}>
            <tbody>
              <tr>
                <td style={lbl}>{t.jsAccountZh}<br />{t.jsAccount}</td>
                <td style={cel}>{firstItem.client}</td>
                <td style={lbl}>{t.jsJobNoZh}<br />{t.jsJobNo}</td>
                <td style={{ ...cel, fontWeight: "bold" }}>{jobNumber || "—"}</td>
                <td style={lbl}>{t.jsDateZh}<br />{t.jsDate}</td>
                <td style={cel}>{fmt(date)}</td>
              </tr>
              <tr>
                <td style={lbl}>{t.jsOrderedByZh}<br />{t.jsOrderedBy}</td>
                <td style={cel}>{firstItem.orderedBy || "—"}</td>
                <td style={lbl}>{t.jsPoNoZh}<br />{t.jsPoNo}</td>
                <td style={cel}>{firstItem.poNumber || "—"}</td>
                <td style={lbl}>{t.jsJobRefZh}<br />{t.jsJobRef}</td>
                <td style={cel}>{firstItem.jobRef || "—"}</td>
              </tr>
              <tr>
                <td style={lbl}>提單資料<br />SS/D.O. NO.</td>
                <td style={cel} colSpan={5}>—</td>
              </tr>
            </tbody>
          </table>

          <div style={{ border: "1px solid #111", borderTop: "none", padding: "4px 6px", display: "flex", alignItems: "center", fontSize: 16 }}>
            <span style={{ flex: 1, textAlign: "center", fontWeight: "bold" }}>
              {t.jsDescriptionZh}<br /><span style={{ letterSpacing: 2 }}>{t.jsDescription}</span>
            </span>
            <span style={{ fontSize: 12 }}>{t.jsIssuedByZh}: {groups[0].delivery.recordedBy || "—"}</span>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", border: "1px solid #111", borderTop: "none", fontSize: 16 }}>
            <tbody>
              <tr>
                <td style={{ padding: 8, verticalAlign: "top", height: 400 }}>
                  {rows.map((r, idx) => (
                    <div key={r.item.id} style={{ marginBottom: idx < rows.length - 1 ? 14 : 0 }}>
                      {r.ref && <div style={{ marginBottom: 4, textDecoration: "underline", whiteSpace: "pre-line" }}>{r.ref}</div>}
                      {r.item.shkNumber && <div style={{ fontWeight: "bold" }}>{r.item.shkNumber}</div>}
                      {r.item.unitCode && <div style={{ fontWeight: "bold" }}>{r.item.unitCode}</div>}
                      {r.item.description && <div>{r.item.description}</div>}
                      {r.delivery.codes && r.delivery.codes.length > 0 && (
                        <div style={{ fontSize: 12, color: "#111" }}>
                          C/S NO. {r.delivery.codes.join(", ")} &nbsp;&nbsp; {r.totals.pkgs} {t.jsPkgs}
                          {r.totals.kgs ? ` \u00b7 ${r.totals.kgs} ${t.jsKgs}` : ""}{r.totals.cbm ? ` \u00b7 ${r.totals.cbm} ${t.jsCbm}` : ""}
                        </div>
                      )}
                    </div>
                  ))}
                  <div style={{ borderTop: "1px solid #111", marginTop: 12, paddingTop: 6 }}>
                    共:&nbsp;&nbsp;&nbsp;{totalPkgs} {t.jsPkgs} &nbsp;&nbsp;&nbsp; {totalKgs !== "" ? totalKgs : "—"} {t.jsKgs} &nbsp;&nbsp;&nbsp; {totalCbm !== "" ? totalCbm : "—"} {t.jsCbm}
                  </div>
                  {anyEstimated && <div style={{ fontSize: 10, color: "#900", marginTop: 4 }}>{t.jsEstimatedNote}</div>}
                </td>
              </tr>
              <tr>
                <td style={{ padding: "0 8px", fontSize: 16 }}>
                  客戶簽署確認 / Customer&nbsp;signature:&nbsp;______________________________&nbsp;(工作妥當及完成)
                  <br /><br /><br /><br />
                </td>
              </tr>
            </tbody>
          </table>

          <div style={{ borderTop: "1px solid #ccc", paddingTop: 5, fontSize: 8, color: "#111", textAlign: "center", marginTop: 8 }}>
            Office and Depot: 21D Wang Toi Shan, Hung Mo Tam, Kam Tin. NT., HK
            <div style={{ marginTop: 2 }}>N.B. Farspeed Contractors Ltd. is a private company. All transaction(s) taken into account are subject to the STANDARD BUSINESS CONDITIONS of the company, details as behind.</div>
            <div style={{ marginTop: 2 }}>(a member of FARSPEED Group)</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function JobSheetPrint({ sheet, onClose, directory, colors, t, lang }) {
  if (sheet.combined) {
    return <CombinedDeliveryPrint sheet={sheet} onClose={onClose} directory={directory} colors={colors} t={t} lang={lang} />;
  }
  const { type, item, delivery } = sheet;
  const isDelivery = type === "Delivery";
  const depotEn = item.depot || "";
  const depotZh = DEPOT_LABELS_ZH[item.depot] || "";
  const dirMatch = (directory || []).find((s) => s.siteEn && (s.siteEn === item.constructionSite || s.siteEn === item.project));
  const siteZh = (dirMatch && dirMatch.siteZh) || "";
  const siteContact = dirMatch ? [dirMatch.contactName, dirMatch.contactPhone].filter(Boolean).join(" ") : "";

  const initialTemplate = isDelivery ? "Delivery" : (JOB_SHEET_TEMPLATES.includes(type) ? type : "Devan");
  const [template, setTemplate] = useState(initialTemplate);
  const [cfsPresetKey, setCfsPresetKey] = useState(CFS_FROM_PRESETS[0].key);

  function fromTopText(tpl) {
    if (tpl === "Devan") return `客人安排運輸送到${depotZh || depotEn}\n(共1櫃)`;
    if (tpl === "CFS") return (CFS_FROM_PRESETS.find((p) => p.key === cfsPresetKey) || CFS_FROM_PRESETS[0]).text;
    if (["Delivery", "Pick-up", "Day Work"].includes(tpl)) return [depotZh, depotEn].filter(Boolean).join("\n");
    return "";
  }
  function fromBottomText(tpl) {
    if (tpl === "Devan") return "*由快達拆櫃";
    if (tpl === "CFS") return "*由客戶自行CFS";
    return "";
  }
  function toTopText() {
    const site = item.constructionSite || item.project || "";
    const out = [];
    if (site) out.push(/^site\s+at/i.test(site) ? site : `SITE AT ${site}`);
    if (siteZh) out.push(siteZh);
    return out.join("\n");
  }
  function toBottomText(tpl) {
    if (tpl === "Devan" || tpl === "CFS") return `暫存${depotZh || depotEn}`;
    if (tpl === "Delivery") return siteContact || "";
    return "";
  }
  function ssText(tpl) {
    if (tpl === "Devan" || tpl === "CFS") return item.ssDoNo || "";
    return "";
  }
  function refText(tpl) {
    if (tpl !== "Delivery") return "";
    const jn = item.jobNumber || "";
    let dates = [];
    if (usesArrivalBatches(item) && delivery && delivery.codes) {
      const set = new Set(delivery.codes);
      dates = [...new Set(activeArrivals(item).filter((a) => (a.codes || []).some((c) => set.has(c))).map((a) => a.date).filter(Boolean))];
    }
    if (dates.length === 0) dates = [effectiveDepotArrivalDate(item)].filter(Boolean);
    if (dates.length === 0 || !jn) return "";
    return dates.map((d) => `Refer to job no. ${jn} on ${fmt(d)}`).join("\n");
  }
  function bodyText(tpl) {
    if (JOB_SHEET_ITEMIZED.includes(tpl)) return "";
    const site = (item.constructionSite || item.project || "").trim();
    const seed = JOB_SHEET_BODY_SEEDS[tpl] || "";
    return [site, seed].filter(Boolean).join("\n\n");
  }

  const itemized = JOB_SHEET_ITEMIZED.includes(template);
  const dateText = isDelivery ? delivery.date : item.depotArrivalDate || effectiveDepotArrivalDate(item);
  const jobNo = isDelivery ? delivery.jobNumber : item.jobNumber;
  const issuedBy = (isDelivery ? delivery.recordedBy : item.recordedBy) || "";
  const showOversize = ["Devan", "CFS"].includes(template) && !!OVERSIZE_RULES[item.client];
  const oversizeText = showOversize ? computeOversizeText(item) : "";

  const pkgs = isDelivery ? (delivery.codes ? delivery.codes.length : Number(delivery.packageCount) || 0) : totalUnits(item);
  let kgs = item.weightKg || "";
  let cbm = item.volumeCbm || "";
  let estimated = false;
  if (isDelivery) {
    if (delivery.codes && (item.packages || []).length > 0) {
      const delivered = item.packages.filter((p) => delivery.codes.includes(p.code));
      const haveAllWeights = delivered.length > 0 && delivered.every((p) => p.weightKg !== "" && p.weightKg != null);
      const haveAllCbm = delivered.length > 0 && delivered.every((p) => p.cbm !== "" && p.cbm != null);
      kgs = haveAllWeights ? String(Math.round(delivered.reduce((s, p) => s + Number(p.weightKg), 0) * 10) / 10) : "";
      cbm = haveAllCbm ? String(Math.round(delivered.reduce((s, p) => s + Number(p.cbm), 0) * 1000) / 1000) : "";
      if (!haveAllWeights || !haveAllCbm) estimated = true;
    } else {
      estimated = true;
    }
    if (estimated) {
      const totalU = totalUnits(item) || 1;
      const share = pkgs / totalU;
      if (!kgs && item.weightKg) kgs = `~${Math.round(Number(item.weightKg) * share * 10) / 10}`;
      if (!cbm && item.volumeCbm) cbm = `~${Math.round(Number(item.volumeCbm) * share * 1000) / 1000}`;
    }
  }
  const csLine = isDelivery
    ? (delivery.codes ? delivery.codes.join(", ") : "")
    : (item.packages || []).map((p) => p.code).join(", ");

  const lbl = { border: "1px solid #111", fontWeight: "bold", padding: "4px 6px", verticalAlign: "top", width: 78, fontSize: 13, wordBreak: "break-word", overflowWrap: "break-word" };
  const cel = { border: "1px solid #111", padding: 6, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };
  const cel0 = { border: "1px solid #111", padding: 0, verticalAlign: "top", wordBreak: "break-word", overflowWrap: "break-word" };
  const fTop = fromTopText(template), fBottom = fromBottomText(template);
  const tTop = toTopText(), tBottom = toBottomText(template);
  const fromHasBottom = !!fBottom, toHasBottom = !!tBottom;
  const bText = bodyText(template), rText = refText(template), sText = ssText(template);

  return (
    <div id="job-sheet-print-root" className="fixed inset-0 z-50 flex flex-col" style={{ background: "rgba(0,0,0,0.5)" }}>
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #job-sheet-print-root, #job-sheet-print-root * { visibility: visible; }
          #job-sheet-print-root { position: absolute; inset: auto; left: 0; top: 0; width: 100%; background: #fff !important; }
          #job-sheet-print-toolbar { display: none !important; }
          #job-sheet-print-scroll { overflow: visible !important; height: auto !important; padding: 0 !important; }
          #job-sheet-print-area { margin: 0 !important; max-width: 100% !important; }
        }
      `}</style>
      <div id="job-sheet-print-toolbar" className="flex flex-wrap items-center gap-2 p-3" style={{ background: colors.navy }}>
        <span className="text-sm font-semibold" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>{t.jsTemplateLabel}:</span>
        <select
          value={template}
          onChange={(e) => setTemplate(e.target.value)}
          className="px-2 py-1.5 rounded text-sm"
          style={{ background: colors.surface, color: colors.ink }}
        >
          {JOB_SHEET_TEMPLATES.map((tp) => <option key={tp} value={tp}>{tp}</option>)}
        </select>
        {template === "CFS" && (
          <select
            value={cfsPresetKey}
            onChange={(e) => setCfsPresetKey(e.target.value)}
            className="px-2 py-1.5 rounded text-sm"
            style={{ background: colors.surface, color: colors.ink }}
          >
            {CFS_FROM_PRESETS.map((p) => <option key={p.key} value={p.key}>{p.key}</option>)}
          </select>
        )}
        <div className="ml-auto flex gap-2">
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => window.print()}>
            {t.printBtn}
          </button>
          <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.onDark}`, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>
            {t.closePreviewBtn}
          </button>
        </div>
      </div>
      <div id="job-sheet-print-scroll" className="flex-1 overflow-y-auto p-6" style={{ background: colors.bg }}>
        <div id="job-sheet-print-area" className="mx-auto" style={{ background: "#fff", color: "#111", maxWidth: 700, padding: 24, fontFamily: "Arial, sans-serif", fontSize: 15 }}>

          {/* Letterhead */}
          <div className="flex items-start justify-between" style={{ marginBottom: 8 }}>
            <div className="flex items-center gap-3">
              <img src={FARSPEED_LOGO_DATA_URI} alt="Farspeed" style={{ height: 78, width: "auto" }} />
              <div>
                <div style={{ fontSize: 24, fontWeight: 700, color: "#111", lineHeight: 1.1 }}>FARSPEED Contractors Limited</div>
                <div style={{ fontSize: 11, color: "#111" }}>P. O. Box No. 1985, Yuen Long Post Office, Yuen Long, N.T., Hong Kong</div>
                <div style={{ fontSize: 11, color: "#111" }}>Tel: +852 5337-9500&nbsp;&nbsp;Fax: +852 2402-4450&nbsp;&nbsp;http://www.farspeed.hk</div>
              </div>
            </div>
            <div className="flex items-start" style={{ fontSize: 16, fontWeight: 700, lineHeight: 1.5 }}>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3 }}>有限公司</div>
              <div style={{ writingMode: "vertical-rl", color: "#111", letterSpacing: 3, marginLeft: 4 }}>快達承判</div>
            </div>
          </div>
          <div style={{ borderTop: "2.5px solid #111", marginBottom: 10 }} />

          <div className="text-center font-bold mb-3" style={{ fontSize: 22, letterSpacing: 6 }}>
            {t.jsTitleZh}&nbsp;&nbsp;{t.jsTitle}
          </div>

          {/* FROM / TO — one real 2-row table so the two sides line up on the same horizontal, no divider between rows */}
          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18 }}>
            <tbody>
              <tr>
                <td style={lbl} rowSpan={2}>{t.jsFromZh}<br />{t.jsFrom}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: fromHasBottom ? "none" : "1px solid #111" }} rowSpan={fromHasBottom ? 1 : 2}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20 }}>{fTop}</div>
                </td>
                <td style={lbl} rowSpan={2}>{t.jsToZh}<br />{t.jsTo}</td>
                <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderTop: "1px solid #111", borderBottom: toHasBottom ? "none" : "1px solid #111" }} rowSpan={toHasBottom ? 1 : 2}>
                  <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{tTop}</div>
                </td>
              </tr>
              <tr>
                {fromHasBottom && (
                  <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderBottom: "1px solid #111", borderTop: "none" }}>
                    <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{fBottom}</div>
                  </td>
                )}
                {toHasBottom && (
                  <td style={{ ...cel0, borderLeft: "1px solid #111", borderRight: "1px solid #111", borderBottom: "1px solid #111", borderTop: "none" }}>
                    <div style={{ padding: "6px 8px", whiteSpace: "pre-line", fontSize: 20, fontWeight: "bold" }}>{tBottom}</div>
                  </td>
                )}
              </tr>
            </tbody>
          </table>

          <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed", fontSize: 18, marginTop: -1 }}>
            <tbody>
              <tr>
                <td style={lbl}>{t.jsAccountZh}<br />{t.jsAccount}</td>
                <td style={cel}>{item.client}</td>
                <td style={lbl}>{t.jsJobNoZh}<br />{t.jsJobNo}</td>
                <td style={{ ...cel, fontWeight: "bold" }}>{jobNo || "—"}</td>
                <td style={lbl}>{t.jsDateZh}<br />{t.jsDate}</td>
                <td style={cel}>{fmt(dateText)}</td>
              </tr>
              <tr>
                <td style={lbl}>{t.jsOrderedByZh}<br />{t.jsOrderedBy}</td>
                <td style={cel}>{item.orderedBy || "—"}</td>
                <td style={lbl}>{t.jsPoNoZh}<br />{t.jsPoNo}</td>
                <td style={cel}>{item.poNumber || "—"}</td>
                <td style={lbl}>{t.jsJobRefZh}<br />{t.jsJobRef}</td>
                <td style={cel}>{item.jobRef || "—"}</td>
              </tr>
              <tr>
                <td style={lbl}>提單資料<br />SS/D.O. NO.</td>
                <td style={cel} colSpan={5}>{sText || "—"}</td>
              </tr>
            </tbody>
          </table>

          <div style={{ border: "1px solid #111", borderTop: "none", padding: "4px 6px", display: "flex", alignItems: "center", fontSize: 16 }}>
            <span style={{ flex: 1, textAlign: "center", fontWeight: "bold" }}>
              {t.jsDescriptionZh}<br /><span style={{ letterSpacing: 2 }}>{t.jsDescription}</span>
            </span>
            <span style={{ fontSize: 12 }}>{t.jsIssuedByZh}: {issuedBy || "—"}</span>
          </div>

          <table style={{ width: "100%", borderCollapse: "collapse", border: "1px solid #111", borderTop: "none", fontSize: 16 }}>
            <tbody>
              <tr>
                <td style={{ padding: 8, verticalAlign: "top", height: 400 }}>
                  {itemized ? (
                    <>
                      {template === "Delivery" && rText && <div style={{ marginBottom: 6, textDecoration: "underline", whiteSpace: "pre-line" }}>{rText}</div>}
                      {template === "Devan" && item.shkNumber && <div style={{ fontWeight: "bold" }}>{item.shkNumber}</div>}
                      {item.unitCode && <div style={{ fontWeight: "bold" }}>{item.unitCode}</div>}
                      {item.description && <div>{item.description}</div>}
                      {csLine && <div style={{ fontSize: 12, color: "#111" }}>C/S NO. {csLine}</div>}
                      <div style={{ borderTop: "1px solid #111", marginTop: 12, paddingTop: 6 }}>
                        共:&nbsp;&nbsp;&nbsp;{pkgs} {t.jsPkgs} &nbsp;&nbsp;&nbsp; {kgs || "—"} {t.jsKgs} &nbsp;&nbsp;&nbsp; {cbm || "—"} {t.jsCbm}
                      </div>
                      {estimated && <div style={{ fontSize: 10, color: "#900", marginTop: 4 }}>{t.jsEstimatedNote}</div>}
                      {showOversize && oversizeText && (
                        <div style={{ borderTop: "1px solid #111", marginTop: 12, paddingTop: 8 }}>
                          <div style={{ fontWeight: "bold", textDecoration: "underline" }}>{t.jsOversizeLabel}</div>
                          <div style={{ marginTop: 4, fontSize: 13, whiteSpace: "pre-line" }}>{oversizeText}</div>
                          <div style={{ fontSize: 10, color: "#111", marginTop: 4 }}>{t.jsOversizeNote}</div>
                        </div>
                      )}
                    </>
                  ) : (
                    <div style={{ whiteSpace: "pre-line" }}>{bText}</div>
                  )}
                </td>
              </tr>
              <tr>
                <td style={{ padding: "0 8px", fontSize: 16 }}>
                  {template === "Pick-up" ? (
                    <>
                      簽署確認收妥 / Signature:&nbsp;______________________________&nbsp;(車輛號碼 Vehicle&nbsp;No.:________ )
                      <br /><br /><br />
                    </>
                  ) : (
                    <>
                      客戶簽署確認 / Customer&nbsp;signature:&nbsp;______________________________&nbsp;(工作妥當及完成)
                      <br /><br /><br /><br />
                    </>
                  )}
                </td>
              </tr>
            </tbody>
          </table>

          <div style={{ borderTop: "1px solid #ccc", paddingTop: 5, fontSize: 8, color: "#111", textAlign: "center", marginTop: 8 }}>
            Office and Depot: 21D Wang Toi Shan, Hung Mo Tam, Kam Tin. NT., HK
            <div style={{ marginTop: 2 }}>N.B. Farspeed Contractors Ltd. is a private company. All transaction(s) taken into account are subject to the STANDARD BUSINESS CONDITIONS of the company, details as behind.</div>
            <div style={{ marginTop: 2 }}>(a member of FARSPEED Group)</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function BillingPanel({ items, onDeleteItem, authUser, colors, t, lang }) {
  const now = new Date();
  const [mode, setMode] = useState("search");
  const [search, setSearch] = useState("");
  const [filterClient, setFilterClient] = useState("All");
  const [filterProject, setFilterProject] = useState("All");
  const [filterJobNo, setFilterJobNo] = useState("All");
  const [expanded, setExpanded] = useState(null);
  const [summaryYear, setSummaryYear] = useState(now.getFullYear());
  const [summaryMonth, setSummaryMonth] = useState(now.getMonth());
  const [expandedClient, setExpandedClient] = useState(null);
  const [pendingDeleteItem, setPendingDeleteItem] = useState(null);

  const allRows = useMemo(() => {
    const rows = [];
    for (const item of items) {
      for (const r of computeItemBillingRows(item)) rows.push(r);
    }
    return rows;
  }, [items]);

  const clientOptions = useMemo(() => [...new Set(allRows.map((r) => r.item.client).filter(Boolean))].sort(), [allRows]);
  const projectOptions = useMemo(() => [...new Set(allRows.map((r) => r.item.constructionSite || r.item.project).filter(Boolean))].sort(), [allRows]);
  const jobNoOptions = useMemo(() => [...new Set(allRows.map((r) => r.item.jobNumber).filter(Boolean))].sort(), [allRows]);

  const monthlySummary = useMemo(() => computeMonthlyBillingSummary(items, summaryYear, summaryMonth), [items, summaryYear, summaryMonth]);
  const yearOptions = useMemo(() => {
    const ys = new Set([now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1]);
    for (const item of items) {
      const d = item.depotArrivalDate || effectiveDepotArrivalDate(item);
      if (d) ys.add(toDateOnly(d).getFullYear());
    }
    return [...ys].sort();
  }, [items]);
  const monthNames = lang === "zh"
    ? ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
    : ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const filtered = useMemo(() => {
    return allRows.filter((r) => {
      const i = r.item;
      if (filterClient !== "All" && i.client !== filterClient) return false;
      if (filterProject !== "All" && (i.constructionSite || i.project) !== filterProject) return false;
      if (filterJobNo !== "All" && i.jobNumber !== filterJobNo) return false;
      if (!search.trim()) return true;
      const q = search.toLowerCase();
      return (
        i.client?.toLowerCase().includes(q) || i.project?.toLowerCase().includes(q) ||
        i.constructionSite?.toLowerCase().includes(q) || i.jobNumber?.toLowerCase().includes(q) ||
        i.id?.toLowerCase().includes(q) || i.shkNumber?.toLowerCase().includes(q) ||
        (r.codes || []).some((c) => (c || "").toLowerCase().includes(q))
      );
    });
  }, [allRows, search, filterClient, filterProject, filterJobNo]);

  const grandTotal = Math.round(filtered.reduce((s, r) => s + r.total, 0) * 100) / 100;
  const money = (n) => `$${n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

  return (
    <div className="flex flex-col gap-4">
      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.billingTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.billingDesc}</p>
        <div className="flex gap-1 rounded-lg p-1 mb-3" style={{ background: colors.surfaceDim, width: "fit-content" }}>
          {[["search", t.billingModeSearch], ["monthly", t.billingModeMonthly]].map(([k, label]) => (
            <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
              {label}
            </button>
          ))}
        </div>
        {mode === "search" && (
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.searchLabel} colors={colors}>
            <input
              className={inputClass}
              style={{ ...inputStyleFor(colors), minWidth: 240 }}
              placeholder={t.billingSearchPlaceholder}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </Field>
          <Field label={t.clientLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={filterClient} onChange={(e) => setFilterClient(e.target.value)}>
              <option>All</option>
              {clientOptions.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label={t.billingColProject} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={filterProject} onChange={(e) => setFilterProject(e.target.value)}>
              <option>All</option>
              {projectOptions.map((p) => <option key={p}>{p}</option>)}
            </select>
          </Field>
          <Field label={t.billingColJobNo} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={filterJobNo} onChange={(e) => setFilterJobNo(e.target.value)}>
              <option>All</option>
              {jobNoOptions.map((j) => <option key={j}>{j}</option>)}
            </select>
          </Field>
        </div>
        )}
        {mode === "monthly" && (
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.billingMonthLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={summaryMonth} onChange={(e) => setSummaryMonth(Number(e.target.value))}>
              {monthNames.map((m, idx) => <option key={idx} value={idx}>{m}</option>)}
            </select>
          </Field>
          <Field label={t.billingYearLabel} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={summaryYear} onChange={(e) => setSummaryYear(Number(e.target.value))}>
              {yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
          </Field>
        </div>
        )}
      </div>

      {mode === "monthly" && (
        <>
          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.billingColClient, t.billingColTotal, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {monthlySummary.clients.length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.billingMonthNoneMsg}</td></tr>
                )}
                {monthlySummary.clients.map((g) => {
                  const isOpen = expandedClient === g.client;
                  return (
                    <React.Fragment key={g.client}>
                      <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }} onClick={() => setExpandedClient(isOpen ? null : g.client)}>
                        <td className="px-3 py-2 font-semibold">{g.client}</td>
                        <td className="px-3 py-2 font-semibold">{money(g.total)}</td>
                        <td className="px-3 py-2 text-right text-xs" style={{ color: colors.amberText }}>{isOpen ? t.billingHideBtn : t.billingShowBtn}</td>
                      </tr>
                      {isOpen && (
                        <tr style={{ background: colors.surfaceDim }}>
                          <td colSpan={3} className="px-4 py-3">
                            <table className="w-full text-xs" style={{ color: colors.ink }}>
                              <thead>
                                <tr>
                                  {[t.billingColProject, t.billingColJobNo, t.billingColCbm, t.billingColBatchDate, "", ""].map((h) => (
                                    <th key={h} className="text-left pr-4 pb-1 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {g.lines.map((l, i) => (
                                  <tr key={i} style={{ borderTop: `1px solid ${colors.line}` }}>
                                    <td className="pr-4 py-1">{l.project}</td>
                                    <td className="pr-4 py-1" style={{ fontFamily: FONT_MONO }}>{l.jobNumber || "—"}</td>
                                    <td className="pr-4 py-1">{l.cbm.toFixed(3)}{l.estimated ? " *" : ""}</td>
                                    <td className="pr-4 py-1">{fmt(l.batchDate)}</td>
                                    <td className="pr-4 py-1" style={{ color: colors.inkFaint }}>{l.detail}</td>
                                    <td className="py-1 text-right font-semibold">{money(l.amount)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
              {monthlySummary.clients.length > 0 && (
                <tfoot>
                  <tr style={{ borderTop: `2px solid ${colors.line}` }}>
                    <td className="px-3 py-2 text-right font-semibold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.billingGrandTotal}</td>
                    <td className="px-3 py-2 font-bold" style={{ color: colors.ink }}>{money(monthlySummary.grandTotal)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
          <div className="text-xs" style={{ color: colors.inkFaint }}>{t.billingMonthFootnote}</div>
        </>
      )}

      {mode === "search" && (
      <>
        <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
        <table className="w-full text-sm" style={{ background: colors.surface }}>
          <thead>
            <tr style={{ background: colors.surfaceDim }}>
              {[t.billingColClient, t.billingColProject, t.billingColJobNo, t.billingColBatchDate, t.billingColCbm, t.billingColRate, t.billingColStatus, t.billingColTotal, "", ""].map((h, idx) => (
                <th key={idx} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={9} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.billingNoneMsg}</td></tr>
            )}
            {filtered.map((r, idx) => {
              const key = `${r.item.id}-${idx}`;
              const isOpen = expanded === key;
              return (
                <React.Fragment key={key}>
                  <tr style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }} onClick={() => setExpanded(isOpen ? null : key)}>
                    <td className="px-3 py-2">{r.item.client}</td>
                    <td className="px-3 py-2">{r.item.constructionSite || r.item.project}</td>
                    <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.item.jobNumber || "—"}</td>
                    <td className="px-3 py-2">{fmt(r.batchDate)}</td>
                    <td className="px-3 py-2">{r.cbm.toFixed(3)}{r.estimated ? <span title={t.billingEstimatedNote} style={{ color: colors.amberText, cursor: "help" }}> *</span> : ""}</td>
                    <td className="px-3 py-2">${r.rate}/{t.billingPerCbmMonth}</td>
                    <td className="px-3 py-2">
                      {r.ongoing ? <Badge tone="amber" colors={colors}>{t.billingOngoing}</Badge> : <Badge tone="grey" colors={colors}>{t.billingClosed}</Badge>}
                    </td>
                    <td className="px-3 py-2 font-semibold">{money(r.total)}</td>
                    <td className="px-3 py-2 text-right text-xs" style={{ color: colors.amberText }}>{isOpen ? t.billingHideBtn : t.billingShowBtn}</td>
                    <td className="px-3 py-2 text-right">
                      <button
                        type="button"
                        title={t.billingDeleteItemBtn}
                        aria-label={t.billingDeleteItemBtn}
                        className="w-6 h-6 rounded-full inline-flex items-center justify-center font-bold"
                        style={{ background: colors.redSoft, color: colors.red, lineHeight: 1 }}
                        onClick={(e) => { e.stopPropagation(); setPendingDeleteItem(r.item); }}
                      >
                        &minus;
                      </button>
                    </td>
                  </tr>
                  {isOpen && (
                    <tr style={{ background: colors.surfaceDim }}>
                      <td colSpan={10} className="px-4 py-3">
                        <div className="text-xs mb-2" style={{ color: colors.inkFaint }}>
                          {t.billingFreeDaysNote(r.freeDays)}{r.codes && r.codes.length > 0 ? ` · ${t.billingCasesLabel}: ${r.codes.join(", ")}` : ""}
                        </div>
                        <table className="text-xs" style={{ color: colors.ink }}>
                          <tbody>
                            {r.breakdown.map((b, i) => (
                              <tr key={i}>
                                <td className="pr-4 py-0.5 font-semibold">{b.label}</td>
                                <td className="pr-4 py-0.5" style={{ color: colors.inkFaint }}>{b.detail}</td>
                                <td className="py-0.5 text-right">{money(b.amount)}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        {r.estimated && <div className="text-xs mt-2" style={{ color: colors.amberText }}>{t.billingEstimatedNote}</div>}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
          {filtered.length > 0 && (
            <tfoot>
              <tr style={{ borderTop: `2px solid ${colors.line}` }}>
                <td colSpan={7} className="px-3 py-2 text-right font-semibold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.billingGrandTotal}</td>
                <td className="px-3 py-2 font-bold" style={{ color: colors.ink }}>{money(grandTotal)}</td>
                <td></td>
                <td></td>
              </tr>
            </tfoot>
          )}
        </table>
        </div>
        <div className="text-xs" style={{ color: colors.inkFaint }}>{t.billingFootnote}</div>
      </>
      )}
      {pendingDeleteItem && (
        <AdminConfirmModal
          authUser={authUser}
          onConfirm={() => { onDeleteItem(pendingDeleteItem.id); setPendingDeleteItem(null); }}
          onClose={() => setPendingDeleteItem(null)}
          colors={colors}
          t={t}
        />
      )}
    </div>
  );
}

const LEGACY_DOC_TYPES = JOB_SHEET_TEMPLATES;
function guessDocTypeFromName(name) {
  const n = (name || "").toLowerCase();
  const map = [
    ["devan", "Devan"], ["cfs", "CFS"], ["delivery", "Delivery"],
    ["shifting", "Shifting"], ["hoisting", "Hoisting"], ["day_work", "Day Work"], ["day work", "Day Work"],
    ["dismantle", "Dismantle"], ["dis_n_removal", "Dis & Removal of Lifting Tools"], ["removal", "Dis & Removal of Lifting Tools"],
    ["job_cancel", "Job Cancel"], ["cancel", "Job Cancel"], ["pick-up", "Pick-up"], ["pickup", "Pick-up"],
    ["re-position", "Re-position"], ["reposition", "Re-position"], ["position", "Position"],
    ["retain", "Retain of Safety Ropes"],
  ];
  for (const [needle, type] of map) if (n.includes(needle)) return type;
  return "Devan";
}
function guessJobNumberFromName(name) {
  const m = (name || "").match(/\b(\d{6,8})\b/);
  return m ? m[1] : "";
}

// Scans a job-sheet-style Excel (Devan/CFS/Delivery layouts) for known Chinese/English
// labels and pulls whatever value sits next to each one - to the right first, then below.
// This is a best-effort pre-fill: the person still reviews and can correct every field
// before committing, same as the PDF/Excel packing-list importers already do.
const CHINESE_CLIENT_ALIASES = {
  "三菱": "Mitsubishi", "迅達": "Schindler", "蒂升": "TK Elevator", "通力": "Kone",
  "富士達": "Fujitec", "其士": "Chevalier", "希格馬": "Sigma", "日立": "Hitachi",
};
function resolveClientGuess(text) {
  if (!text) return "";
  const t = text.trim();
  for (const [zh, client] of Object.entries(CHINESE_CLIENT_ALIASES)) {
    if (t.includes(zh)) return client;
  }
  const lower = t.toLowerCase();
  const match = CLIENTS.find((c) => lower.includes(c.toLowerCase()));
  return match || "";
}
const JOBSHEET_LABEL_ALIASES = {
  account: ["客戶", "account"],
  jobNo: ["快達單號", "job no", "job no."],
  date: ["日期", "date"],
  orderedBy: ["落單人", "ordered by"],
  poNo: ["採購編號", "p.o. no", "po no"],
  jobRef: ["地盤代號", "job ref"],
  ssDoNo: ["提單資料", "ss/d.o. no", "ss/do no"],
  to: ["送", "to"],
  from: ["由", "from"],
};
const ALL_JOBSHEET_LABELS = Object.values(JOBSHEET_LABEL_ALIASES).flat();
function normCell(v) {
  return String(v == null ? "" : v).replace(/\s+/g, "").toLowerCase();
}
function normLabelCompare(v) {
  return normCell(v).replace(/[.:]+$/g, "");
}
function isKnownLabel(v) {
  const n = normLabelCompare(v);
  if (!n) return false;
  // Exact match only (after trimming trailing punctuation) - a substring check here
  // would flag real data as a label whenever a short alias like "to" happens to
  // appear inside a word (e.g. "Elevator").
  return ALL_JOBSHEET_LABELS.some((a) => n === normLabelCompare(a));
}
// Tries every cell matching one of the given labels (not just the first), and for each
// match looks to its right then below for a value - skipping any candidate that is
// itself a recognized label (adjacent field labels sit right next to each other in
// these sheets, e.g. "JOB NO." next to "DATE", so a naive "next cell" grab picks up
// the neighboring label's text instead of real data).
function findLabelValue(rows, aliases) {
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const cellNorm = normCell(rows[r][c]);
      if (!cellNorm) continue;
      if (!aliases.some((a) => cellNorm.includes(normCell(a)))) continue;
      for (let cc = c + 1; cc < Math.min(c + 6, rows[r].length); cc++) {
        const v = rows[r][cc];
        if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) return String(v).trim();
      }
      for (let rr = r + 1; rr < Math.min(r + 3, rows.length); rr++) {
        const v = rows[rr][c];
        if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) return String(v).trim();
      }
      // This occurrence's neighbors were empty or just more labels - keep scanning
      // the sheet in case the same label appears again (often once in Chinese, once
      // in English) with real data next to the other occurrence.
    }
  }
  return "";
}
// Collects the address block after a FROM/TO-style label - these often span several
// lines (English site name, then a separate Chinese site name below it), so this
// gathers consecutive non-empty lines and splits them by script rather than grabbing
// just the single adjacent cell.
function findAddressLines(rows, aliases, maxLines = 8) {
  const aliasesNorm = aliases.map((a) => normLabelCompare(a));
  // Prefer an exact label match over a substring match - a single Chinese character
  // alias like "送"/"由" can appear embedded inside an unrelated sentence elsewhere on
  // the sheet (e.g. "运输送到" contains "送"), which would otherwise be mistaken for
  // the real FROM/TO label cell.
  let hasExactMatch = false;
  for (const row of rows) {
    if (row.some((cell) => aliasesNorm.includes(normLabelCompare(cell)))) { hasExactMatch = true; break; }
  }
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const cellNorm = normCell(rows[r][c]);
      if (!cellNorm) continue;
      const isMatch = hasExactMatch ? aliasesNorm.includes(normLabelCompare(rows[r][c])) : aliases.some((a) => cellNorm.includes(normCell(a)));
      if (!isMatch) continue;
      let startR = -1, startC = -1;
      for (let cc = c + 1; cc < Math.min(c + 6, rows[r].length); cc++) {
        const v = rows[r][cc];
        if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) { startR = r; startC = cc; break; }
      }
      if (startR === -1) {
        for (let rr = r + 1; rr < Math.min(r + 3, rows.length); rr++) {
          const v = rows[rr][c];
          if (v != null && String(v).trim() !== "" && !isKnownLabel(v)) { startR = rr; startC = c; break; }
        }
      }
      if (startR === -1) continue;
      const enLines = [], zhLines = [];
      for (let rr = startR; rr < Math.min(startR + maxLines, rows.length); rr++) {
        const v = rows[rr][startC];
        if (v == null || String(v).trim() === "") break;
        const text = String(v).trim();
        if (isKnownLabel(text)) break;
        if (/^(暫存|temp\.?\s*storage)/i.test(text)) break;
        if (/[\u4e00-\u9fff]/.test(text)) zhLines.push(text);
        else enLines.push(text);
      }
      if (enLines.length || zhLines.length) return { en: enLines.join(" ").trim(), zh: zhLines.join("").trim() };
    }
  }
  return { en: "", zh: "" };
}
// Converts a Date object to a plain YYYY-MM-DD string using its LOCAL calendar date,
// not toISOString() (which converts through UTC and silently shifts the date backward
// by a day for anyone in a timezone ahead of UTC, e.g. Hong Kong).
function dateToLocalISO(d) {
  const yr = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${yr}-${mo}-${da}`;
}
// Reads the workbook's own Excel "Print Area" setting for a sheet, if one is defined.
// Farspeed's rule: only rows within the printable A4 area count as actually arriving or
// leaving - anything beyond that (e.g. extra case groups that overflowed onto a page that
// was never actually printed/sent) is not part of this job. Returns the last row number
// covered by the print area (1-indexed, matching Excel), or null if none is set.
function getPrintAreaLastRow(wb, sheetIndex) {
  try {
    const names = wb.Workbook && wb.Workbook.Names;
    if (!names) return null;
    for (const n of names) {
      if (!n || !n.Name || !/print_area/i.test(n.Name)) continue;
      if (n.Sheet !== undefined && n.Sheet !== null && n.Sheet !== sheetIndex) continue;
      const ref = String(n.Ref || "");
      const m = ref.match(/!\$?[A-Za-z]+\$?(\d+)(?::\$?[A-Za-z]+\$?(\d+))?/);
      if (m) return Math.max(Number(m[1]), Number(m[2] || m[1]));
    }
  } catch (e) { /* fall back to using the whole sheet */ }
  return null;
}
function guessFieldsFromWorkbook(wb) {
  const sheetIndex = 0;
  const sheet = wb.Sheets[wb.SheetNames[sheetIndex]];
  let rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  const printAreaLastRow = getPrintAreaLastRow(wb, sheetIndex);
  if (printAreaLastRow) rows = rows.slice(0, printAreaLastRow);
  const flatText = rows.map((r) => r.join(" ")).join("\n");

  const siteBlock = findAddressLines(rows, JOBSHEET_LABEL_ALIASES.to);
  const out = {
    client: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.account),
    jobNumber: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.jobNo),
    date: "",
    projectEn: siteBlock.en,
    projectZh: siteBlock.zh,
    orderedBy: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.orderedBy),
    poNumber: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.poNo),
    jobRef: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.jobRef),
    ssDoNo: findLabelValue(rows, JOBSHEET_LABEL_ALIASES.ssDoNo),
    unitCode: "",
    referJobNumber: "",
    referDate: "",
    packageCount: "",
    weightKg: "",
    volumeCbm: "",
  };

  const rawDate = findLabelValue(rows, JOBSHEET_LABEL_ALIASES.date);
  if (rawDate) {
    const d = new Date(rawDate);
    if (!isNaN(d)) out.date = dateToLocalISO(d);
    else {
      const m = rawDate.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
      if (m) {
        const yr = m[3].length === 2 ? `20${m[3]}` : m[3];
        out.date = `${yr}-${String(m[2]).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;
      }
    }
  }

  const referMatches = [...flatText.matchAll(/ref(?:er)?\.?\s*(?:to\s+)?job\s*no\.?\s*([A-Za-z0-9\-]+)\s*(?:on\s*([\d\/\.\- ]+\d))?/gi)];
  if (referMatches.length) {
    out.referJobNumber = [...new Set(referMatches.map((m) => m[1].trim()))].join(", ");
    const firstWithDate = referMatches.find((m) => m[2]);
    if (firstWithDate) {
      const d = new Date(firstWithDate[2].trim());
      if (!isNaN(d)) out.referDate = dateToLocalISO(d);
    }
  }

  const liftMatch = flatText.match(/lift\s*no\.?\s*(?:phase\s*\d+\s*)?(#[0-9,\s]+)/i)
    || flatText.match(/unit\s*(?:no\.?|code)\s*[:\-]?\s*([A-Za-z0-9\-\.\/#, ]{2,30})/i)
    || flatText.match(/escalator\s*no\.?\s*([A-Za-z0-9 &,]{2,20})/i);
  if (liftMatch) {
    out.unitCode = liftMatch[1].trim().replace(/\s+/g, " ");
  } else {
    // Fallback: some sheets list unit codes only as "REFNUMBER/E1", "REFNUMBER2/E2" lines
    // rather than a labeled heading - collect every such pair's unit code.
    const pairMatches = [...flatText.matchAll(/\b\d{6,12}\/([A-Za-z]\d+)\b/g)];
    if (pairMatches.length) out.unitCode = [...new Set(pairMatches.map((m) => m[1]))].join(", ");
  }

  // Prefer the explicit "共:" / "Total:" line for aggregate totals - a job sheet with
  // multiple case groups before the total would otherwise match the FIRST group's
  // numbers instead of the actual total.
  const totalLineMatch = flatText.match(/(?:^|\n)[^\n]*(?:共|total)[:\uff1a][^\n]*/i);
  const totalsText = totalLineMatch ? totalLineMatch[0] : flatText;
  const pkgsMatch = totalsText.match(/(\d+)\s*PKGS?/i) || flatText.match(/(\d+)\s*PKGS?/i);
  if (pkgsMatch) out.packageCount = pkgsMatch[1];
  const kgsMatch = totalsText.match(/([\d,]+(?:\.\d+)?)\s*KGS?/i) || flatText.match(/([\d,]+(?:\.\d+)?)\s*KGS?/i);
  if (kgsMatch) out.weightKg = kgsMatch[1].replace(/,/g, "");
  const cbmMatch = totalsText.match(/([\d,]+(?:\.\d+)?)\s*CBM/i) || flatText.match(/([\d,]+(?:\.\d+)?)\s*CBM/i);
  if (cbmMatch) out.volumeCbm = cbmMatch[1].replace(/,/g, "");

  const ssShipMatch = flatText.match(/ex\s*ss\.?\s*"[^"]+"[^\n]*/i);
  if (ssShipMatch && !out.ssDoNo) out.ssDoNo = ssShipMatch[0].trim();

  return out;
}

function siteKeyFor(en, zh) {
  return [(en || "").trim().toLowerCase(), (zh || "").trim()].filter(Boolean);
}
function normalizeSiteForMatch(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/^(site at|no\.?)\s*/i, "")
    .replace(/,?\s*(hong ?kong|hk|wan ?chai)\s*$/i, "")
    .replace(/[^\w\u4e00-\u9fff]+/g, " ")
    .trim();
}
function sitesLooselyMatch(enA, zhA, enB, zhB) {
  const aEn = normalizeSiteForMatch(enA), bEn = normalizeSiteForMatch(enB);
  const aZh = String(zhA || "").trim(), bZh = String(zhB || "").trim();
  if (aEn && bEn && (aEn.includes(bEn) || bEn.includes(aEn))) return true;
  if (aZh && bZh && (aZh.includes(bZh) || bZh.includes(aZh))) return true;
  return false;
}
// Parses typed case-number input like "1,3-5,7" into a Set of integers, supporting
// both single numbers and ranges.
function parseRangeInput(text) {
  const nums = new Set();
  for (const part of String(text || "").split(",").map((s) => s.trim()).filter(Boolean)) {
    const rangeMatch = part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const a = Number(rangeMatch[1]), b = Number(rangeMatch[2]);
      for (let n = Math.min(a, b); n <= Math.max(a, b); n++) nums.add(n);
    } else if (/^\d+$/.test(part)) {
      nums.add(Number(part));
    }
  }
  return nums;
}
// Extracts the leading case number from a code like "26/40" -> 26, so typed numbers can
// be matched against it regardless of the "/total" suffix.
function codeLeadingNumber(code) {
  const m = String(code).match(/^\d+/);
  return m ? Number(m[0]) : null;
}
function sumSelectedPackages(packages, codes) {
  const set = new Set(codes);
  const matched = (packages || []).filter((p) => set.has(p.code));
  return {
    count: matched.length,
    weight: matched.reduce((s, p) => s + (Number(p.weightKg) || 0), 0),
    cbm: matched.reduce((s, p) => s + (Number(p.cbm) || 0), 0),
  };
}
function groupPackagesByOrder(pkgs) {
  const hasAnyOrderNo = pkgs.some((p) => p.orderNo);
  if (!hasAnyOrderNo) return [{ orderNo: "", packages: pkgs }];
  const order = [];
  const groups = {};
  for (const p of pkgs) {
    const key = p.orderNo || "";
    if (!groups[key]) { groups[key] = []; order.push(key); }
    groups[key].push(p);
  }
  return order.map((k) => ({ orderNo: k, packages: groups[k] }));
}
function LegacyUploadRow({ row, onChange, onRemove, incoming, items, onProcessAll, processing, processDisabled, colors, t, lang }) {
  const inputStyle = inputStyleFor(colors);
  const set = (k) => (e) => onChange({ ...row, [k]: e.target.value });
  const itemized = JOB_SHEET_ITEMIZED.includes(row.docType);
  const canMatchIncoming = row.docType === "Devan" || row.docType === "CFS";
  const matchedIncomings = canMatchIncoming && row.client && (row.projectEn || row.projectZh)
    ? (incoming || []).filter((inc) => {
        if (inc.client !== row.client) return false;
        const done = new Set(inc.checkedInCodes || []);
        const remaining = (inc.packages || []).filter((p) => !done.has(p.code));
        if (remaining.length === 0) return false;
        return sitesLooselyMatch(row.projectEn, row.projectZh, inc.project, inc.constructionSite);
      })
    : [];
  const selectedByIncoming = row.selectedByIncoming || {};
  function toggleIncomingCode(incId, code) {
    const cur = selectedByIncoming[incId] || [];
    const next = cur.includes(code) ? cur.filter((c) => c !== code) : [...cur, code];
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: next } });
  }
  function selectAllIncoming(incId, remainingPkgs) {
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: remainingPkgs.map((p) => p.code) } });
  }

  // Delivery: match against real inventory items (itemized, with cases still remaining)
  // instead of asking for a flat package count - prefer a typed "Refers to Arrival Job
  // No." if given, otherwise match by client + site the same way as Devan/CFS.
  const matchedItems = row.docType === "Delivery" && row.client
    ? items.filter((it) => {
        if (it.client !== row.client) return false;
        if (!(it.packages || []).length) return false;
        if (remainingPackages(it).length === 0) return false;
        const refNos = String(row.referJobNumber || "").split(",").map((s) => s.trim()).filter(Boolean);
        if (refNos.length) return refNos.includes(String(it.jobNumber || "").trim());
        return (row.projectEn || row.projectZh) && sitesLooselyMatch(row.projectEn, row.projectZh, it.project, it.constructionSite);
      })
    : [];
  const selectedByItem = row.selectedByItem || {};
  function toggleItemCode(itemId, code) {
    const cur = selectedByItem[itemId] || [];
    const next = cur.includes(code) ? cur.filter((c) => c !== code) : [...cur, code];
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: next } });
  }
  function selectAllItem(itemId, remainingPkgs) {
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: remainingPkgs.map((p) => p.code) } });
  }
  const [typeSelectText, setTypeSelectText] = useState({});
  function applyTypeSelectItem(itemId, remainingPkgs) {
    const nums = parseRangeInput(typeSelectText[itemId]);
    if (nums.size === 0) return;
    const matchingCodes = remainingPkgs.filter((p) => nums.has(codeLeadingNumber(p.code))).map((p) => p.code);
    const cur = selectedByItem[itemId] || [];
    onChange({ ...row, selectedByItem: { ...selectedByItem, [itemId]: [...new Set([...cur, ...matchingCodes])] } });
    setTypeSelectText((prev) => ({ ...prev, [itemId]: "" }));
  }
  function applyTypeSelectIncoming(incId, remainingPkgs) {
    const nums = parseRangeInput(typeSelectText[incId]);
    if (nums.size === 0) return;
    const matchingCodes = remainingPkgs.filter((p) => nums.has(codeLeadingNumber(p.code))).map((p) => p.code);
    const cur = selectedByIncoming[incId] || [];
    onChange({ ...row, selectedByIncoming: { ...selectedByIncoming, [incId]: [...new Set([...cur, ...matchingCodes])] } });
    setTypeSelectText((prev) => ({ ...prev, [incId]: "" }));
  }
  return (
    <div className="rounded p-3 flex flex-col gap-3" style={{ border: `1px solid ${colors.line}`, background: colors.surface }}>
      <div className="flex items-start justify-between gap-2">
        <div className="text-sm font-semibold" style={{ color: colors.ink, wordBreak: "break-all" }}>{row.file.name}</div>
        <button type="button" className="text-xs font-semibold whitespace-nowrap" style={{ color: colors.red }} onClick={onRemove}>{t.deleteBtn}</button>
      </div>
      {row.autoDetected && (
        <div className="px-2 py-1 rounded text-xs w-fit" style={{ background: colors.amberSoft, color: colors.amberText }}>
          {t.legacyAutoDetectedTag}
        </div>
      )}
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
        <Field label={t.legacyDocType} colors={colors}>
          <select className={inputClass} style={inputStyle} value={row.docType} onChange={set("docType")}>
            {LEGACY_DOC_TYPES.map((tp) => <option key={tp} value={tp}>{tp}</option>)}
          </select>
        </Field>
        <Field label={t.clientLabel} colors={colors}>
          <select className={inputClass} style={{ ...inputStyle, borderColor: !row.client ? colors.red : inputStyle.borderColor }} value={row.client} onChange={set("client")}>
            <option value="">{t.legacyClientUnresolved}</option>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.legacyProjectSiteEn} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, borderColor: !row.projectEn && !row.projectZh ? colors.red : inputStyle.borderColor }} value={row.projectEn} onChange={set("projectEn")} />
        </Field>
        <Field label={t.legacyProjectSiteZh} colors={colors}>
          <input className={inputClass} style={{ ...inputStyle, borderColor: !row.projectEn && !row.projectZh ? colors.red : inputStyle.borderColor }} value={row.projectZh} onChange={set("projectZh")} />
        </Field>
        <Field label={t.fJobNumber} colors={colors}>
          <input className={inputClass} style={inputStyle} value={row.jobNumber} onChange={set("jobNumber")} />
        </Field>
        <Field label={t.colDate} colors={colors}>
          <input type="date" className={inputClass} style={inputStyle} value={row.date} onChange={set("date")} />
        </Field>
        {itemized && (
          <>
            {(matchedIncomings.length > 0 || matchedItems.length > 0) ? null : (
              <Field label={t.legacyUnitCode} colors={colors}>
                <input className={inputClass} style={inputStyle} value={row.unitCode} onChange={set("unitCode")} />
              </Field>
            )}
            {(matchedIncomings.length > 0 || matchedItems.length > 0) ? null : (
              <>
                <Field label={t.legacyPkgs} colors={colors}>
                  <input type="number" min="0" className={inputClass} style={inputStyle} value={row.packageCount} onChange={set("packageCount")} />
                </Field>
                <Field label={t.legacyWeightKg} colors={colors}>
                  <input type="number" min="0" className={inputClass} style={inputStyle} value={row.weightKg} onChange={set("weightKg")} />
                </Field>
                <Field label={t.legacyCbm} colors={colors}>
                  <input type="number" min="0" step="0.001" className={inputClass} style={inputStyle} value={row.volumeCbm} onChange={set("volumeCbm")} />
                </Field>
              </>
            )}
            {(row.docType === "Devan" || row.docType === "CFS") && (
              <>
                <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.shkNumber} onChange={set("shkNumber")} />
                </Field>
                <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.jobRef} onChange={set("jobRef")} />
                </Field>
                <div className="col-span-2 md:col-span-3">
                  <Field label={t.fSsDoNo} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={row.ssDoNo} onChange={set("ssDoNo")} placeholder={'ex ss."SHIP" V.___; CONTAINERS NO. ___'} />
                  </Field>
                </div>
              </>
            )}
            {row.docType === "Delivery" && (
              <>
                <Field label={t.legacyReferJobNoLabel} hint={t.legacyReferJobNoHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.referJobNumber} onChange={set("referJobNumber")} />
                </Field>
                <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={row.shkNumber} onChange={set("shkNumber")} />
                </Field>
              </>
            )}
          </>
        )}
      </div>
      {onProcessAll && (matchedIncomings.length > 0 || matchedItems.length > 0) && (
        <button
          type="button"
          className="px-4 py-2 rounded text-sm font-semibold w-fit"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: processDisabled ? 0.5 : 1 }}
          disabled={processDisabled}
          onClick={onProcessAll}
        >
          {processing ? t.legacyProcessingMsg : t.saveBtn}
        </button>
      )}
      {matchedIncomings.length > 0 && (
        <div className="rounded p-3" style={{ background: colors.greenSoft, border: `1px solid ${colors.green}` }}>
          <div className="text-sm font-semibold mb-2" style={{ color: colors.green }}>
            {t.legacyMatchedIncomingCount(matchedIncomings.length)}
          </div>
          <div className="mb-3 max-w-xs">
            <Field label={t.packingListApplyDepot} colors={colors}>
              <select className={inputClass} style={inputStyle} value={row.depot || DEPOTS[0]} onChange={set("depot")}>
                {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
              </select>
            </Field>
          </div>
          <div className="flex flex-col gap-4">
            {matchedIncomings.map((inc) => {
              const done = new Set(inc.checkedInCodes || []);
              const remainingPkgs = (inc.packages || []).filter((p) => !done.has(p.code));
              const selectedCodes = selectedByIncoming[inc.id] || [];
              const totals = sumSelectedPackages(inc.packages, selectedCodes);
              return (
                <div key={inc.id} style={{ borderTop: `1px solid ${colors.green}`, paddingTop: 10 }}>
                  <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
                    <div className="text-xs font-semibold" style={{ color: colors.green, fontFamily: FONT_DISPLAY }}>
                      {t.legacyMatchedIncoming(inc.id)}{inc.unitCode ? ` \u00b7 ${inc.unitCode}` : ""}
                    </div>
                    <div className="flex items-center gap-2">
                      <input
                        className={inputClass}
                        style={{ ...inputStyleFor(colors), width: 140, fontSize: 12, padding: "4px 8px" }}
                        placeholder={t.legacyTypeSelectPlaceholder}
                        value={typeSelectText[inc.id] || ""}
                        onChange={(e) => setTypeSelectText((prev) => ({ ...prev, [inc.id]: e.target.value }))}
                        onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); applyTypeSelectIncoming(inc.id, remainingPkgs); } }}
                      />
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => applyTypeSelectIncoming(inc.id, remainingPkgs)}>{t.legacyTypeSelectBtn}</button>
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => selectAllIncoming(inc.id, remainingPkgs)}>{t.selectAllBtn}</button>
                    </div>
                  </div>
                  {selectedCodes.length > 0 && (
                    <div className="text-xs mb-2 font-semibold" style={{ color: colors.ink }}>
                      {t.legacySelectedTotals(totals.count, Math.round(totals.weight * 10) / 10, Math.round(totals.cbm * 1000) / 1000)}
                    </div>
                  )}
                  <div className="flex flex-col gap-3">
                    {groupPackagesByOrder(remainingPkgs).map((grp) => (
                      <div key={grp.orderNo || "_"}>
                        {grp.orderNo && (
                          <div className="text-xs font-semibold mb-1.5" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                            {grp.orderNo}{inc.unitCode ? ` / ${inc.unitCode}` : ""}
                          </div>
                        )}
                        <div className="flex flex-wrap gap-2">
                          {grp.packages.map((p) => (
                            <button
                              key={p.code}
                              type="button"
                              onClick={() => toggleIncomingCode(inc.id, p.code)}
                              className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                              style={{
                                border: `1px solid ${selectedCodes.includes(p.code) ? colors.amber : colors.line}`,
                                background: selectedCodes.includes(p.code) ? colors.amberSoft : colors.surface,
                                color: selectedCodes.includes(p.code) ? colors.amberText : colors.ink,
                              }}
                              title={p.description}
                            >
                              {p.code}{p.description ? ` \u2014 ${p.description}` : ""}
                            </button>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
      {matchedItems.length > 0 && (
        <div className="rounded p-3" style={{ background: colors.greenSoft, border: `1px solid ${colors.green}` }}>
          <div className="text-sm font-semibold mb-2" style={{ color: colors.green }}>
            {t.legacyMatchedItemsCount(matchedItems.length)}
          </div>
          {(() => {
            const grand = matchedItems.reduce((acc, it) => {
              const s = sumSelectedPackages(it.packages, selectedByItem[it.id] || []);
              return { count: acc.count + s.count, weight: acc.weight + s.weight, cbm: acc.cbm + s.cbm };
            }, { count: 0, weight: 0, cbm: 0 });
            return grand.count > 0 ? (
              <div className="text-xs mb-3 font-semibold px-2 py-1.5 rounded" style={{ color: colors.ink, background: colors.surface, width: "fit-content" }}>
                {t.legacySelectedTotalsGrand(grand.count, Math.round(grand.weight * 10) / 10, Math.round(grand.cbm * 1000) / 1000)}
              </div>
            ) : null;
          })()}
          <div className="flex flex-col gap-4">
            {matchedItems.map((it) => {
              const remainingPkgs = remainingPackages(it);
              const selectedCodes = selectedByItem[it.id] || [];
              const totals = sumSelectedPackages(it.packages, selectedCodes);
              return (
                <div key={it.id} style={{ borderTop: `1px solid ${colors.green}`, paddingTop: 10 }}>
                  <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
                    <div className="text-xs font-semibold" style={{ color: colors.green, fontFamily: FONT_DISPLAY }}>
                      {t.legacyMatchedItem(it.id)}{it.unitCode ? ` \u00b7 ${it.unitCode}` : ""}{it.jobNumber ? ` \u00b7 ${t.fJobNumber}: ${it.jobNumber}` : ""}
                    </div>
                    <div className="flex items-center gap-2">
                      <input
                        className={inputClass}
                        style={{ ...inputStyleFor(colors), width: 140, fontSize: 12, padding: "4px 8px" }}
                        placeholder={t.legacyTypeSelectPlaceholder}
                        value={typeSelectText[it.id] || ""}
                        onChange={(e) => setTypeSelectText((prev) => ({ ...prev, [it.id]: e.target.value }))}
                        onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); applyTypeSelectItem(it.id, remainingPkgs); } }}
                      />
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => applyTypeSelectItem(it.id, remainingPkgs)}>{t.legacyTypeSelectBtn}</button>
                      <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => selectAllItem(it.id, remainingPkgs)}>{t.selectAllBtn}</button>
                    </div>
                  </div>
                  {selectedCodes.length > 0 && (
                    <div className="text-xs mb-2 font-semibold" style={{ color: colors.ink }}>
                      {t.legacySelectedTotals(totals.count, Math.round(totals.weight * 10) / 10, Math.round(totals.cbm * 1000) / 1000)}
                    </div>
                  )}
                  <div className="flex flex-wrap gap-2">
                    {remainingPkgs.map((p) => (
                      <button
                        key={p.code}
                        type="button"
                        onClick={() => toggleItemCode(it.id, p.code)}
                        className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                        style={{
                          border: `1px solid ${selectedCodes.includes(p.code) ? colors.amber : colors.line}`,
                          background: selectedCodes.includes(p.code) ? colors.amberSoft : colors.surface,
                          color: selectedCodes.includes(p.code) ? colors.amberText : colors.ink,
                        }}
                        title={p.description}
                      >
                        {p.code}{p.description ? ` \u2014 ${p.description}` : ""}
                      </button>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
      {!row.projectEn && !row.projectZh && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {t.legacySiteRequiredMsg}
        </div>
      )}
      {(row.docType === "Devan" || row.docType === "CFS") && (
        <div className="text-xs" style={{ color: colors.inkFaint }}>{t.legacyArrivalStaysOpenHint}</div>
      )}
      {row.docType === "Delivery" && matchedItems.length === 0 && (
        <div className="px-2 py-1.5 rounded text-xs" style={{ background: colors.redSoft, color: colors.red }}>
          {row.referJobNumber ? t.legacyNoArrivalFoundHint(row.referJobNumber) : t.legacyNoReferralHint}
        </div>
      )}
    </div>
  );
}

function IncomingPanel({ incoming, setIncoming, items, directory, setDirectory, onCheckIn, onAddIncoming, colors, t, lang }) {
  const [search, setSearch] = useState("");
  const [filterClient, setFilterClient] = useState("All");
  const [showCompleted, setShowCompleted] = useState(false);
  const [expandedId, setExpandedId] = useState(null);
  const [selectedByShipment, setSelectedByShipment] = useState({});
  const [formByShipment, setFormByShipment] = useState({});
  const inputStyle = inputStyleFor(colors);

  function remainingPkgs(inc) {
    const done = new Set(inc.checkedInCodes || []);
    return (inc.packages || []).filter((p) => !done.has(p.code));
  }
  function isComplete(inc) {
    return remainingPkgs(inc).length === 0;
  }

  const clientOptions = useMemo(() => [...new Set(incoming.map((i) => i.client).filter(Boolean))].sort(), [incoming]);
  const filtered = incoming.filter((inc) => {
    if (filterClient !== "All" && inc.client !== filterClient) return false;
    if (!showCompleted && isComplete(inc)) return false;
    if (!search.trim()) return true;
    const q = search.toLowerCase();
    return inc.client?.toLowerCase().includes(q) || inc.project?.toLowerCase().includes(q) ||
      inc.constructionSite?.toLowerCase().includes(q) || inc.unitCode?.toLowerCase().includes(q) ||
      (inc.packages || []).some((p) => (p.code || "").toLowerCase().includes(q));
  });

  function getSel(id) { return selectedByShipment[id] || []; }
  function toggleCode(id, code) {
    setSelectedByShipment((prev) => {
      const cur = prev[id] || [];
      return { ...prev, [id]: cur.includes(code) ? cur.filter((c) => c !== code) : [...cur, code] };
    });
  }
  function selectAll(inc) {
    setSelectedByShipment((prev) => ({ ...prev, [inc.id]: remainingPkgs(inc).map((p) => p.code) }));
  }
  function getForm(id) {
    return formByShipment[id] || { type: "Devan", depot: DEPOTS[0], jobNumber: "", date: todayStr(), ssDoNo: "" };
  }
  function setForm(id, patch) {
    setFormByShipment((prev) => ({ ...prev, [id]: { ...getForm(id), ...patch } }));
  }

  function submitCheckIn(inc) {
    const codes = getSel(inc.id);
    const form = getForm(inc.id);
    if (codes.length === 0 || !form.jobNumber || !form.date) return;
    onCheckIn({ incomingId: inc.id, codes, type: form.type, depot: form.depot, jobNumber: form.jobNumber, date: form.date, ssDoNo: form.ssDoNo });
    setSelectedByShipment((prev) => ({ ...prev, [inc.id]: [] }));
    setFormByShipment((prev) => ({ ...prev, [inc.id]: { ...getForm(inc.id), jobNumber: "" } }));
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.incomingTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.incomingDesc}</p>
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.searchLabel} colors={colors}>
            <input className={inputClass} style={{ ...inputStyle, minWidth: 220 }} value={search} onChange={(e) => setSearch(e.target.value)} />
          </Field>
          <Field label={t.clientLabel} colors={colors}>
            <select className={inputClass} style={inputStyle} value={filterClient} onChange={(e) => setFilterClient(e.target.value)}>
              <option>All</option>
              {clientOptions.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <label className="flex items-center gap-2 text-sm pb-2" style={{ color: colors.inkFaint }}>
            <input type="checkbox" checked={showCompleted} onChange={(e) => setShowCompleted(e.target.checked)} />
            {t.incomingShowCompleted}
          </label>
        </div>
      </div>

      <div className="flex flex-col gap-3">
        {filtered.length === 0 && (
          <div className="px-3 py-6 text-center text-sm rounded-lg" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.inkFaint }}>
            {t.incomingNoneMsg}
          </div>
        )}
        {filtered.map((inc) => {
          const remaining = remainingPkgs(inc);
          const complete = isComplete(inc);
          const isOpen = expandedId === inc.id;
          const sel = getSel(inc.id);
          const form = getForm(inc.id);
          return (
            <div key={inc.id} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}`, background: colors.surface }}>
              <div
                className="px-4 py-3 flex flex-wrap items-center justify-between gap-2 cursor-pointer"
                style={{ background: colors.surfaceDim }}
                onClick={() => setExpandedId(isOpen ? null : inc.id)}
              >
                <div>
                  <div className="text-sm font-bold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
                    {inc.client} · {inc.project || inc.constructionSite}{inc.unitCode ? ` \u00b7 ${inc.unitCode}` : ""}
                  </div>
                  <div className="text-xs" style={{ color: colors.inkFaint }}>
                    {t.incomingCaseCount((inc.packages || []).length)}{inc.linkedItemId ? ` \u00b7 ${t.incomingLinkedTo(inc.linkedItemId)}` : ""}
                  </div>
                </div>
                {complete ? (
                  <Badge tone="green" colors={colors}>{t.incomingFullyCheckedIn}</Badge>
                ) : (
                  <Badge tone="amber" colors={colors}>{t.incomingRemainingBadge(remaining.length, (inc.packages || []).length)}</Badge>
                )}
                <button
                  type="button"
                  title={t.incomingDeleteBtn}
                  aria-label={t.incomingDeleteBtn}
                  className="w-6 h-6 rounded-full inline-flex items-center justify-center font-bold"
                  style={{ background: colors.redSoft, color: colors.red, lineHeight: 1 }}
                  onClick={(e) => { e.stopPropagation(); if (window.confirm(t.incomingDeleteConfirm)) setIncoming((prev) => prev.filter((i) => i.id !== inc.id)); }}
                >
                  &minus;
                </button>
              </div>
              {isOpen && (
                <div className="p-4 flex flex-col gap-4">
                  <div className="max-w-xs">
                    <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                      <input
                        className={inputClass}
                        style={inputStyleFor(colors)}
                        value={inc.shkNumber || ""}
                        onChange={(e) => setIncoming((prev) => prev.map((i) => (i.id === inc.id ? { ...i, shkNumber: e.target.value } : i)))}
                      />
                    </Field>
                  </div>
                  {remaining.length > 0 ? (
                    <>
                      <div>
                        <div className="flex items-center justify-between mb-2">
                          <div className="text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{t.incomingSelectCasesLabel}</div>
                          <button type="button" className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => selectAll(inc)}>{t.selectAllBtn}</button>
                        </div>
                        <div className="flex flex-col gap-3">
                          {groupPackagesByOrder(remaining).map((grp) => (
                            <div key={grp.orderNo || "_"}>
                              {grp.orderNo && (
                                <div className="text-xs font-semibold mb-1.5" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                                  {grp.orderNo}{inc.unitCode ? ` / ${inc.unitCode}` : ""}
                                </div>
                              )}
                              <div className="flex flex-wrap gap-2">
                                {grp.packages.map((p) => (
                                  <button
                                    key={p.code}
                                    type="button"
                                    onClick={() => toggleCode(inc.id, p.code)}
                                    className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                                    style={{
                                      border: `1px solid ${sel.includes(p.code) ? colors.amber : colors.line}`,
                                      background: sel.includes(p.code) ? colors.amberSoft : colors.surface,
                                      color: sel.includes(p.code) ? colors.amberText : colors.ink,
                                    }}
                                    title={p.description}
                                  >
                                    {p.code}{p.description ? ` \u2014 ${p.description}` : ""}
                                  </button>
                                ))}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
                        <Field label={t.fArrivingType} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={form.type} onChange={(e) => setForm(inc.id, { type: e.target.value })}>
                            {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
                          </select>
                        </Field>
                        <Field label={t.packingListApplyDepot} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={form.depot} onChange={(e) => setForm(inc.id, { depot: e.target.value })}>
                            {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                          </select>
                        </Field>
                        <Field label={t.colDate} colors={colors}>
                          <input type="date" className={inputClass} style={inputStyle} value={form.date} onChange={(e) => setForm(inc.id, { date: e.target.value })} />
                        </Field>
                        <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
                          <div className="flex gap-2">
                            <input className={inputClass + " flex-1"} style={inputStyle} value={form.jobNumber} onChange={(e) => setForm(inc.id, { jobNumber: e.target.value })} />
                            <button
                              type="button"
                              className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
                              style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
                              onClick={() => setForm(inc.id, { jobNumber: nextJobNumber(items) })}
                            >
                              {t.generateJobNoBtn}
                            </button>
                          </div>
                        </Field>
                        {form.type === "Devan" && (
                          <div className="col-span-1 sm:col-span-2 md:col-span-4">
                            <Field label={t.fSsDoNo} hint={t.fSsDoNoHint} colors={colors}>
                              <input className={inputClass} style={inputStyle} value={form.ssDoNo} onChange={(e) => setForm(inc.id, { ssDoNo: e.target.value })} placeholder={'ex ss."SHIP" V.___; CONTAINERS NO. ___'} />
                            </Field>
                          </div>
                        )}
                      </div>
                      <button
                        className="px-4 py-2 rounded text-sm font-semibold w-fit"
                        style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: (sel.length === 0 || !form.jobNumber || !form.date) ? 0.5 : 1 }}
                        disabled={sel.length === 0 || !form.jobNumber || !form.date}
                        onClick={() => submitCheckIn(inc)}
                      >
                        {t.incomingCheckInBtn(sel.length)}
                      </button>
                    </>
                  ) : (
                    <div className="text-sm" style={{ color: colors.green }}>{t.incomingFullyCheckedIn}</div>
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

function LegacyUploadsPanel({ legacyArchive, setLegacyArchive, items, incoming, onLegacyCheckIn, onLegacyCheckInBatch, directory, onLegacyImport, onLegacyDeliver, onLegacyEnrich, colors, t, lang }) {
  const [rows, setRows] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState(null);
  const [backlogSearch, setBacklogSearch] = useState("");
  const [backlogTypeFilter, setBacklogTypeFilter] = useState("All");
  const [editingBacklogId, setEditingBacklogId] = useState(null);
  const [backlogEditDraft, setBacklogEditDraft] = useState(null);
  const fileInputRef = React.useRef(null);

  const [scanning, setScanning] = useState(false);

  async function handleFilesSelected(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    setScanning(true);
    const newRows = [];
    for (const file of files) {
      const base = {
        file,
        docType: guessDocTypeFromName(file.name),
        client: "",
        projectEn: "",
        projectZh: "",
        jobNumber: guessJobNumberFromName(file.name),
        date: "",
        unitCode: "",
        packageCount: "",
        weightKg: "",
        volumeCbm: "",
        ssDoNo: "",
        shkNumber: "",
        jobRef: "",
        referJobNumber: "",
        referDate: "",
        autoDetected: false,
      };
      const isExcel = /\.(xlsx|xls|csv)$/i.test(file.name);
      if (isExcel) {
        try {
          const buf = await file.arrayBuffer();
          const wb = XLSX.read(buf, { type: "array", cellDates: true });
          const guessed = guessFieldsFromWorkbook(wb);
          const clientMatch = resolveClientGuess(guessed.client);
          Object.assign(base, {
            client: clientMatch || base.client,
            projectEn: guessed.projectEn || base.projectEn,
            projectZh: guessed.projectZh || base.projectZh,
            jobNumber: guessed.jobNumber || base.jobNumber,
            date: guessed.date || base.date,
            unitCode: guessed.unitCode || base.unitCode,
            packageCount: guessed.packageCount || base.packageCount,
            weightKg: guessed.weightKg || base.weightKg,
            volumeCbm: guessed.volumeCbm || base.volumeCbm,
            ssDoNo: guessed.ssDoNo || base.ssDoNo,
            jobRef: guessed.jobRef || base.jobRef,
            referJobNumber: guessed.referJobNumber || base.referJobNumber,
            referDate: guessed.referDate || base.referDate,
            autoDetected: true,
          });
          if (base.docType === "Delivery" && guessed.referJobNumber) base.jobNumber = base.jobNumber || guessed.referJobNumber;
        } catch (err) { /* fall back to filename-only guesses */ }
      }
      newRows.push(base);
    }
    setRows((prev) => [...prev, ...newRows]);
    setResults(null);
    setScanning(false);
  }

  function updateRow(idx, next) {
    setRows((prev) => prev.map((r, i) => (i === idx ? next : r)));
  }
  function removeRow(idx) {
    setRows((prev) => prev.filter((_, i) => i !== idx));
  }

  async function processAll() {
    setProcessing(true);
    const archiveEntries = [];
    const importRows = [];
    const jobNoToImportIdx = new Map(); // arrival job number -> index into importRows, for same-batch linking
    const pendingEnrichments = [];
    const checkInOps = []; // {op, archiveEntry} pairs, resolved in one batch call after the loop
    const fileUriById = {};

    // Pass 1: archive every file, and create the arrival (Devan/CFS = add to storage).
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const id = `LEG${Date.now()}${Math.floor(Math.random() * 1000)}-${i}`;
      let fileUri = null;
      try {
        fileUri = await compressFileToDataUri(row.file);
      } catch (e) { /* archive without file if compression fails */ }
      if (fileUri) {
        try { await storageSet(`legacyDoc:${id}`, JSON.stringify({ uri: fileUri, name: row.file.name, at: todayStr() })); } catch (e) {}
      }
      fileUriById[i] = fileUri;

      const hasSite = !!(row.projectEn || row.projectZh);
      let linkedItemIndex = null;
      let enrichTargetId = null;
      const matchedIncomings = (row.docType === "Devan" || row.docType === "CFS") && row.client && hasSite
        ? (incoming || []).filter((inc) => {
            if (inc.client !== row.client) return false;
            const done = new Set(inc.checkedInCodes || []);
            if ((inc.packages || []).filter((p) => !done.has(p.code)).length === 0) return false;
            return sitesLooselyMatch(row.projectEn, row.projectZh, inc.project, inc.constructionSite);
          })
        : [];
      const selectedByIncoming = row.selectedByIncoming || {};
      const hasAnyIncomingSelection = matchedIncomings.some((inc) => (selectedByIncoming[inc.id] || []).length > 0);
      const archiveEntry = {
        id, rowIndex: i, fileName: row.file.name, docType: row.docType, client: row.client,
        project: [row.projectEn, row.projectZh].filter(Boolean).join(" / "),
        jobNumber: row.jobNumber, date: row.date, uploadedAt: todayStr(), hasFile: !!fileUri,
        linkedItemId: null, linkedItemIds: [], __linkedItemIndex: null,
      };
      archiveEntries.push(archiveEntry);

      if (row.docType !== "Delivery" && row.client && hasSite) {
        if (hasAnyIncomingSelection && onLegacyCheckInBatch) {
          // Matched to one or more real Incoming shipments with cases selected - queue a
          // check-in for each one (creates or extends its linked inventory item, same as
          // the Incoming tab), resolved together with every other row's check-ins after
          // this loop.
          for (const inc of matchedIncomings) {
            const codes = selectedByIncoming[inc.id] || [];
            if (codes.length === 0) continue;
            checkInOps.push({
              op: {
                incomingId: inc.id,
                codes,
                type: row.docType,
                depot: row.depot || DEPOTS[0],
                jobNumber: row.jobNumber,
                date: row.date,
                ssDoNo: row.ssDoNo,
                shkNumber: row.shkNumber,
                jobRef: row.jobRef,
              },
              archiveEntry,
            });
          }
        } else {
        const jobNoTrim = String(row.jobNumber || "").trim();
        const existingByJobNo = jobNoTrim ? items.find((it) => String(it.jobNumber || "").trim() === jobNoTrim) : null;
        if (existingByJobNo) {
          // A real item already exists for this job number - most likely created via the
          // proper Packing List import, which has real per-case CBM/KG/case numbers.
          // Enrich it with this file's metadata rather than creating a duplicate that
          // would only carry flat totals and shadow the good per-case data.
          const patch = {};
          if (!existingByJobNo.ssDoNo && row.ssDoNo) patch.ssDoNo = row.ssDoNo;
          if (!existingByJobNo.shkNumber && row.shkNumber) patch.shkNumber = row.shkNumber;
          if (!existingByJobNo.jobRef && row.jobRef) patch.jobRef = row.jobRef;
          if (!existingByJobNo.project && row.projectEn) patch.project = row.projectEn;
          if (!existingByJobNo.constructionSite && row.projectZh) patch.constructionSite = row.projectZh;
          if (!existingByJobNo.depotArrivalDate && row.date) patch.depotArrivalDate = row.date;
          patch.notes = [existingByJobNo.notes, t.legacyEnrichedNote(row.file.name)].filter(Boolean).join(" \u00b7 ");
          enrichTargetId = existingByJobNo.id;
          pendingEnrichments.push({ itemId: existingByJobNo.id, patch });
        } else {
          const dirMatch = (directory || []).find((s) =>
            (row.projectEn && s.siteEn && s.siteEn.toLowerCase() === row.projectEn.toLowerCase()) ||
            (row.projectZh && s.siteZh && s.siteZh === row.projectZh)
          );
          const newItem = {
            ...emptyForm(),
            client: row.client,
            project: row.projectEn || (dirMatch ? dirMatch.siteEn : ""),
            constructionSite: row.projectZh || (dirMatch ? dirMatch.siteZh : ""),
            jobNumber: row.jobNumber,
            jobRef: row.jobRef || "",
            depotArrivalDate: row.date,
            unitCode: row.unitCode,
            packageCount: row.packageCount || "",
            weightKg: row.weightKg || "",
            volumeCbm: row.volumeCbm || "",
            arrivingType: row.docType === "CFS" ? "CFS" : "Devan",
            ssDoNo: row.ssDoNo || "",
            shkNumber: row.shkNumber || "",
            notes: t.legacyImportedNote(row.file.name),
            deliveries: [],
          };
          linkedItemIndex = importRows.length;
          if (row.jobNumber) jobNoToImportIdx.set(String(row.jobNumber).trim(), linkedItemIndex);
          importRows.push(newItem);
        }
        }
      }
      archiveEntry.linkedItemId = enrichTargetId;
      archiveEntry.__linkedItemIndex = linkedItemIndex;
    }

    if (checkInOps.length > 0 && onLegacyCheckInBatch) {
      const results = onLegacyCheckInBatch(checkInOps.map((c) => c.op));
      results.forEach((itemId, idx) => {
        if (!itemId) return;
        const entry = checkInOps[idx].archiveEntry;
        if (!entry.linkedItemIds.includes(itemId)) entry.linkedItemIds.push(itemId);
        entry.linkedItemId = entry.linkedItemIds.join(", ");
      });
    }

    // Pass 2: Delivery = subtract from storage. Rows matched to real inventory items (via
    // the case-selection UI) deliver those exact codes; rows with no match fall back to
    // the old referJobNumber + flat package count behaviour.
    const existingDeliveryEntries = [];
    let sameBatchDeliveryCount = 0;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row.docType !== "Delivery") continue;
      const archiveEntry = archiveEntries[i];
      const selectedByItem = row.selectedByItem || {};
      const hasItemSelections = Object.values(selectedByItem).some((codes) => (codes || []).length > 0);

      if (hasItemSelections) {
        for (const [itemId, codes] of Object.entries(selectedByItem)) {
          if (!codes || codes.length === 0) continue;
          existingDeliveryEntries.push({
            itemId,
            delivery: {
              date: row.date || todayStr(), deliveredTo: row.projectEn || row.projectZh, receivedBy: "",
              jobNumber: row.jobNumber, recordedBy: "", notes: t.legacyImportedNote(row.file.name),
              shkNumber: row.shkNumber || "",
              codes,
            },
            archiveEntry,
          });
        }
        continue;
      }

      const refNos = String(row.referJobNumber || "").split(",").map((s) => s.trim()).filter(Boolean);
      if (!refNos.length) continue;
      const deliveryRecord = {
        date: row.date || todayStr(), deliveredTo: row.projectEn || row.projectZh, receivedBy: "",
        jobNumber: row.jobNumber, recordedBy: "", notes: t.legacyImportedNote(row.file.name),
        shkNumber: row.shkNumber || "",
        packageCount: row.packageCount || 1,
      };
      let anyRefMatched = false;
      for (const refNo of refNos) {
        const sameBatchIdx = jobNoToImportIdx.get(refNo);
        if (sameBatchIdx != null) {
          importRows[sameBatchIdx].deliveries = [...(importRows[sameBatchIdx].deliveries || []), { ...deliveryRecord, id: `D${Date.now()}${i}` }];
          archiveEntry.__linkedItemIndex = sameBatchIdx;
          sameBatchDeliveryCount++;
          anyRefMatched = true;
          continue;
        }
        const existing = items.find((it) => String(it.jobNumber || "").trim() === refNo);
        if (existing) {
          existingDeliveryEntries.push({ itemId: existing.id, delivery: deliveryRecord, archiveEntry });
          anyRefMatched = true;
        } else {
          archiveEntry.unmatchedReferral = refNo;
        }
      }
    }

    const createdItems = importRows.length > 0 ? onLegacyImport(importRows) : [];
    for (const entry of archiveEntries) {
      if (entry.__linkedItemIndex != null) entry.linkedItemId = createdItems[entry.__linkedItemIndex] ? createdItems[entry.__linkedItemIndex].id : null;
      delete entry.__linkedItemIndex;
      delete entry.rowIndex;
    }
    if (pendingEnrichments.length > 0 && onLegacyEnrich) onLegacyEnrich(pendingEnrichments);
    if (existingDeliveryEntries.length > 0 && onLegacyDeliver) {
      const results = onLegacyDeliver(existingDeliveryEntries.map((e) => ({ itemId: e.itemId, delivery: e.delivery })));
      results.forEach((r, idx) => {
        const entry = existingDeliveryEntries[idx].archiveEntry;
        if (!entry.linkedItemIds.includes(r.itemId)) entry.linkedItemIds.push(r.itemId);
        entry.linkedItemId = entry.linkedItemIds.join(", ");
      });
    }

    setLegacyArchive((prev) => [...archiveEntries, ...prev]);
    setResults({ archived: archiveEntries.length, created: createdItems.length, delivered: sameBatchDeliveryCount + existingDeliveryEntries.length, enriched: pendingEnrichments.length, checkedIn: checkInOps.length });
    setRows([]);
    setProcessing(false);
  }

  const backlogFiltered = (legacyArchive || []).filter((r) => {
    if (backlogTypeFilter !== "All" && r.docType !== backlogTypeFilter) return false;
    if (!backlogSearch.trim()) return true;
    const q = backlogSearch.toLowerCase();
    return r.fileName?.toLowerCase().includes(q) || r.client?.toLowerCase().includes(q) ||
      r.project?.toLowerCase().includes(q) || r.jobNumber?.toLowerCase().includes(q) ||
      r.linkedItemId?.toLowerCase().includes(q);
  });

  async function viewArchivedFile(id) {
    try {
      const res = await storageGet(`legacyDoc:${id}`);
      if (!res) return;
      const { uri, name } = JSON.parse(res.value);
      const w = window.open("", "_blank");
      if (!w) return;
      if (uri.startsWith("data:application/pdf")) {
        w.document.write(`<title>${name}</title><embed src="${uri}" type="application/pdf" style="width:100%;height:100vh;">`);
      } else {
        w.document.write(`<title>${name}</title><body style="margin:0;background:#333;display:flex;justify-content:center;"><img src="${uri}" style="max-width:100%;height:auto;"></body>`);
      }
    } catch (e) { /* noop */ }
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.legacyUploadTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.legacyUploadDesc}</p>
        <input ref={fileInputRef} type="file" multiple accept=".xlsx,.xls,.pdf,image/*" className="hidden" onChange={handleFilesSelected} disabled={scanning} />
        <button
          className="px-3 py-1.5 rounded text-sm font-semibold"
          style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: scanning ? 0.6 : 1 }}
          disabled={scanning}
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
        >
          {scanning ? t.legacyScanningMsg : t.legacyChooseFilesBtn}
        </button>
        {!scanning && <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>{t.legacyAutoDetectHint}</div>}
      </div>

      {rows.length > 0 && (
        <div className="flex flex-col gap-3">
          {rows.map((row, idx) => (
            <LegacyUploadRow key={idx} row={row} onChange={(next) => updateRow(idx, next)} onRemove={() => removeRow(idx)} incoming={incoming} items={items} onProcessAll={processAll} processing={processing} processDisabled={processing || rows.some((r) => !r.projectEn && !r.projectZh) || rows.some((r) => !r.client)} colors={colors} t={t} lang={lang} />
          ))}
          {rows.some((r) => !r.client) && (
            <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
              {t.legacyClientRequiredSummaryMsg}
            </div>
          )}
          {rows.some((r) => !r.projectEn && !r.projectZh) && (
            <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
              {t.legacySiteRequiredSummaryMsg}
            </div>
          )}
          <div className="flex gap-2">
            <button
              className="px-4 py-2 rounded text-sm font-semibold"
              style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: (processing || rows.some((r) => !r.projectEn && !r.projectZh) || rows.some((r) => !r.client)) ? 0.6 : 1 }}
              disabled={processing || rows.some((r) => !r.projectEn && !r.projectZh) || rows.some((r) => !r.client)}
              onClick={processAll}
            >
              {processing ? t.legacyProcessingMsg : t.legacyProcessBtn(rows.length)}
            </button>
          </div>
        </div>
      )}
      {results && (
        <div className="px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>
          {t.legacyResultsMsg(results.archived, results.created, results.delivered, results.enriched, results.checkedIn)}
        </div>
      )}

      <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.legacyBacklogTitle}</h3>
        <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.legacyBacklogDesc}</p>
        <div className="flex flex-wrap gap-3 items-end">
          <Field label={t.searchLabel} colors={colors}>
            <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }} value={backlogSearch} onChange={(e) => setBacklogSearch(e.target.value)} />
          </Field>
          <Field label={t.legacyDocType} colors={colors}>
            <select className={inputClass} style={inputStyleFor(colors)} value={backlogTypeFilter} onChange={(e) => setBacklogTypeFilter(e.target.value)}>
              <option>All</option>
              {LEGACY_DOC_TYPES.map((tp) => <option key={tp}>{tp}</option>)}
            </select>
          </Field>
        </div>
      </div>

      <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
        <table className="w-full text-sm" style={{ background: colors.surface }}>
          <thead>
            <tr style={{ background: colors.surfaceDim }}>
              {[t.legacyColFile, t.legacyDocType, t.clientLabel, t.legacyProjectSite, t.fJobNumber, t.colDate, t.legacyColLinked, ""].map((h) => (
                <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {backlogFiltered.length === 0 && (
              <tr><td colSpan={8} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.legacyBacklogNoneMsg}</td></tr>
            )}
            {backlogFiltered.map((r) => {
              const isEditing = editingBacklogId === r.id;
              const inputStyle = inputStyleFor(colors);
              if (isEditing) {
                const d = backlogEditDraft;
                return (
                  <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, background: colors.amberSoft }}>
                    <td colSpan={8} className="px-3 py-3">
                      <div className="text-xs font-semibold mb-2 truncate" style={{ color: colors.inkFaint }}>{r.fileName}</div>
                      <div className="flex flex-wrap gap-3">
                        <Field label={t.legacyDocType} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={d.docType} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, docType: e.target.value }))}>
                            {LEGACY_DOC_TYPES.map((tp) => <option key={tp}>{tp}</option>)}
                          </select>
                        </Field>
                        <Field label={t.clientLabel} colors={colors}>
                          <select className={inputClass} style={inputStyle} value={d.client} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, client: e.target.value }))}>
                            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                          </select>
                        </Field>
                        <Field label={t.legacyProjectSite} colors={colors}>
                          <input className={inputClass} style={inputStyle} value={d.project} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, project: e.target.value }))} />
                        </Field>
                        <Field label={t.fJobNumber} colors={colors}>
                          <input className={inputClass} style={inputStyle} value={d.jobNumber} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, jobNumber: e.target.value }))} />
                        </Field>
                        <Field label={t.colDate} colors={colors}>
                          <input type="date" className={inputClass} style={inputStyle} value={d.date} onChange={(e) => setBacklogEditDraft((p) => ({ ...p, date: e.target.value }))} />
                        </Field>
                      </div>
                      <div className="text-xs mt-2" style={{ color: colors.inkFaint }}>{t.legacyEditLinkedHint}</div>
                      <div className="mt-3">
                        <button
                          className="text-xs font-semibold mr-3"
                          style={{ color: colors.green }}
                          onClick={() => {
                            setLegacyArchive((prev) => prev.map((row) => (row.id === r.id ? { ...row, ...backlogEditDraft } : row)));
                            setEditingBacklogId(null);
                            setBacklogEditDraft(null);
                          }}
                        >
                          {t.saveBtn}
                        </button>
                        <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => { setEditingBacklogId(null); setBacklogEditDraft(null); }}>
                          {t.cancelBtn}
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              }
              return (
              <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                <td className="px-3 py-2 max-w-[200px] truncate">{r.fileName}</td>
                <td className="px-3 py-2">{r.docType}</td>
                <td className="px-3 py-2">{r.client}</td>
                <td className="px-3 py-2 max-w-[180px] truncate">{r.project}</td>
                <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{r.jobNumber || "—"}</td>
                <td className="px-3 py-2">{r.date ? fmt(r.date) : "—"}</td>
                <td className="px-3 py-2">
                  {r.linkedItemId ? (
                    <span style={{ color: colors.green, fontWeight: 600 }}>{r.docType === "Delivery" ? t.legacyDeliveredFrom(r.linkedItemId) : r.linkedItemId}</span>
                  ) : r.unmatchedReferral ? (
                    <span style={{ color: colors.red }} title={t.legacyUnmatchedHint}>{t.legacyUnmatchedReferral(r.unmatchedReferral)}</span>
                  ) : (
                    <span style={{ color: colors.inkFaint }}>{t.legacyArchivedOnly}</span>
                  )}
                </td>
                <td className="px-3 py-2 text-right whitespace-nowrap">
                  {r.hasFile && <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => viewArchivedFile(r.id)}>{t.signedDocViewBtn}</button>}
                  <button
                    className="text-xs font-semibold"
                    style={{ color: colors.inkFaint }}
                    onClick={() => {
                      setEditingBacklogId(r.id);
                      setBacklogEditDraft({ docType: r.docType, client: r.client, project: r.project, jobNumber: r.jobNumber || "", date: r.date || "" });
                    }}
                  >
                    {t.editBtn}
                  </button>
                </td>
              </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DirectoryPanel({ directory, setDirectory, employees, setEmployees, freeRules, setFreeRules, cbmRates, setCbmRates, legacyArchive, setLegacyArchive, items, incoming, onLegacyImport, onLegacyDeliver, onLegacyEnrich, onLegacyCheckIn, onLegacyCheckInBatch, colors, t, lang }) {
  const [mode, setMode] = useState("sites");
  const [editingSite, setEditingSite] = useState(null);
  const [siteForm, setSiteForm] = useState(null);
  const [editingEmp, setEditingEmp] = useState(null);
  const [empForm, setEmpForm] = useState(null);
  const [ruleDraft, setRuleDraft] = useState({ project: "", days: "" });
  const inputStyle = inputStyleFor(colors);
  const allRoles = [...new Set([...DEFAULT_ROLES, ...employees.map((e) => e.role)])];

  function addFreeRule() {
    if (!ruleDraft.project.trim() || !(Number(ruleDraft.days) > 0)) return;
    setFreeRules((rs) => [...(rs || []), { id: `FR${Date.now()}`, project: ruleDraft.project.trim(), days: Number(ruleDraft.days) }]);
    setRuleDraft({ project: "", days: "" });
  }
  function deleteFreeRule(id) {
    setFreeRules((rs) => (rs || []).filter((r) => r.id !== id));
  }

  function setRate(client, value) {
    setCbmRates((r) => ({ ...(r || {}), [client]: value }));
  }
  function resetRate(client) {
    setCbmRates((r) => {
      const next = { ...(r || {}) };
      delete next[client];
      return next;
    });
  }

  function newSiteForm() {
    return { siteEn: "", siteZh: "", client: CLIENTS[0], jobRef: "", orderedBy: "", accountOfficer: "" };
  }
  function saveSite() {
    if (!siteForm.siteEn.trim()) return;
    if (editingSite) {
      setDirectory((d) => d.map((s) => (s.id === editingSite.id ? { ...s, ...siteForm } : s)));
    } else {
      setDirectory((d) => [...d, { ...siteForm, id: `SITE${Date.now()}` }]);
    }
    setEditingSite(null);
    setSiteForm(null);
  }
  function deleteSite(id) {
    setDirectory((d) => d.filter((s) => s.id !== id));
  }

  function newEmpForm() {
    return { name: "", role: allRoles[0] || "" };
  }
  function saveEmp() {
    if (!empForm.name.trim()) return;
    if (editingEmp) {
      setEmployees((es) => es.map((e) => (e.id === editingEmp.id ? { ...e, ...empForm } : e)));
    } else {
      setEmployees((es) => [...es, { ...empForm, id: `EMP${Date.now()}` }]);
    }
    setEditingEmp(null);
    setEmpForm(null);
  }
  function deleteEmp(id) {
    setEmployees((es) => es.filter((e) => e.id !== id));
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["sites", t.tabSitesAccounts], ["employees", t.tabEmployees], ["freedays", t.tabFreeStorage], ["pricing", t.tabPricing]].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>

      {mode === "freedays" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.freeStorageTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.freeStorageDesc}</p>
            <div className="flex flex-wrap items-end gap-3">
              <Field label={t.fFreeProject} hint={t.fFreeProjectHint} colors={colors}>
                <input className={inputClass} style={{ ...inputStyle, minWidth: "260px" }} value={ruleDraft.project} onChange={(e) => setRuleDraft((d) => ({ ...d, project: e.target.value }))} />
              </Field>
              <Field label={t.fFreeDays} colors={colors}>
                <input type="number" min="1" className={inputClass} style={{ ...inputStyle, width: "90px" }} value={ruleDraft.days} onChange={(e) => setRuleDraft((d) => ({ ...d, days: e.target.value }))} />
              </Field>
              <button
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY, opacity: !ruleDraft.project.trim() || !(Number(ruleDraft.days) > 0) ? 0.5 : 1 }}
                onClick={addFreeRule}
              >
                {t.freeStorageAddBtn}
              </button>
            </div>
          </div>

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.freeColProject, t.freeColDays, ""].map((h, idx) => (
                    <th key={idx} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {(freeRules || []).length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.freeStorageNoneMsg}</td></tr>
                )}
                {(freeRules || []).map((r) => (
                  <tr key={r.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                    <td className="px-3 py-2">{r.project}</td>
                    <td className="px-3 py-2">{r.days}</td>
                    <td className="px-3 py-2 text-right">
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => deleteFreeRule(r.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {mode === "pricing" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.pricingTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.pricingDesc}</p>
          </div>

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.pricingColClient, t.pricingColRate, ""].map((h, idx) => (
                    <th key={idx} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.keys(DEFAULT_CBM_RATES).map((client) => {
                  const override = (cbmRates || {})[client];
                  const hasOverride = override != null && override !== "";
                  return (
                    <tr key={client} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                      <td className="px-3 py-2 font-semibold">{client}</td>
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-2">
                          <span style={{ color: colors.inkFaint }}>$</span>
                          <input
                            type="number" min="0" step="0.01"
                            className={inputClass}
                            style={{ ...inputStyle, width: "110px" }}
                            placeholder={String(DEFAULT_CBM_RATES[client])}
                            value={hasOverride ? override : ""}
                            onChange={(e) => setRate(client, e.target.value)}
                          />
                          <span style={{ color: colors.inkFaint }}>{t.pricingPerCbm}</span>
                        </div>
                      </td>
                      <td className="px-3 py-2 text-right">
                        {hasOverride && (
                          <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => resetRate(client)}>{t.pricingResetBtn}</button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {mode === "sites" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.dirTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.dirDesc}</p>
            <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => { setEditingSite(null); setSiteForm(newSiteForm()); }}>
              {t.dirAddBtn}
            </button>
          </div>

          {siteForm && (
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                <Field label={t.fSiteEn} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.siteEn} onChange={(e) => setSiteForm((f) => ({ ...f, siteEn: e.target.value }))} />
                </Field>
                <Field label={t.fSiteZh} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.siteZh} onChange={(e) => setSiteForm((f) => ({ ...f, siteZh: e.target.value }))} />
                </Field>
                <Field label={t.fDirClient} colors={colors}>
                  <select className={inputClass} style={inputStyle} value={siteForm.client} onChange={(e) => setSiteForm((f) => ({ ...f, client: e.target.value }))}>
                    {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                  </select>
                </Field>
                <Field label={t.fDirJobRef} hint={t.fJobRefHint} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.jobRef} onChange={(e) => setSiteForm((f) => ({ ...f, jobRef: e.target.value }))} />
                </Field>
                <Field label={t.fDirOrderedBy} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={siteForm.orderedBy} onChange={(e) => setSiteForm((f) => ({ ...f, orderedBy: e.target.value }))} />
                </Field>
                <Field label={t.fDirOfficer} colors={colors}>
                  <select className={inputClass} style={inputStyle} value={siteForm.accountOfficer} onChange={(e) => setSiteForm((f) => ({ ...f, accountOfficer: e.target.value }))}>
                    <option value=""></option>
                    {employees.map((e) => <option key={e.id} value={e.name}>{e.name}</option>)}
                  </select>
                </Field>
              </div>
              <div className="flex gap-2 mt-4">
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={saveSite}>
                  {t.saveBtn}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setEditingSite(null); setSiteForm(null); }}>
                  {t.cancelBtn}
                </button>
              </div>
            </div>
          )}

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.dirColSite, t.dirColClient, t.dirColJobRef, t.dirColOfficer, t.dirColOrderedBy, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {directory.length === 0 && (
                  <tr><td colSpan={6} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.dirNoneMsg}</td></tr>
                )}
                {directory.map((s) => (
                  <tr key={s.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                    <td className="px-3 py-2">
                      <div>{s.siteEn}</div>
                      {s.siteZh && <div className="text-xs" style={{ color: colors.inkFaint }}>{s.siteZh}</div>}
                    </td>
                    <td className="px-3 py-2">{s.client}</td>
                    <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.jobRef || "—"}</td>
                    <td className="px-3 py-2">{s.accountOfficer || "—"}</td>
                    <td className="px-3 py-2">{s.orderedBy || "—"}</td>
                    <td className="px-3 py-2 text-right whitespace-nowrap">
                      <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => { setEditingSite(s); setSiteForm({ ...s }); }}>{t.editBtn}</button>
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => deleteSite(s.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {mode === "employees" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.empTitle}</h3>
            <p className="text-sm mb-3" style={{ color: colors.inkFaint }}>{t.empDesc}</p>
            <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={() => { setEditingEmp(null); setEmpForm(newEmpForm()); }}>
              {t.empAddBtn}
            </button>
          </div>

          {empForm && (
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                <Field label={t.fEmpName} colors={colors}>
                  <input className={inputClass} style={inputStyle} value={empForm.name} onChange={(e) => setEmpForm((f) => ({ ...f, name: e.target.value }))} />
                </Field>
                <Field label={t.fEmpRole} colors={colors}>
                  <input
                    className={inputClass}
                    style={inputStyle}
                    list="role-options"
                    placeholder={t.fEmpRolePlaceholder}
                    value={empForm.role}
                    onChange={(e) => setEmpForm((f) => ({ ...f, role: e.target.value }))}
                  />
                  <datalist id="role-options">
                    {allRoles.map((r) => <option key={r} value={r} />)}
                  </datalist>
                </Field>
              </div>
              <div className="flex gap-2 mt-4">
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={saveEmp}>
                  {t.saveBtn}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setEditingEmp(null); setEmpForm(null); }}>
                  {t.cancelBtn}
                </button>
              </div>
            </div>
          )}

          <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
            <table className="w-full text-sm" style={{ background: colors.surface }}>
              <thead>
                <tr style={{ background: colors.surfaceDim }}>
                  {[t.empColName, t.empColRole, ""].map((h) => (
                    <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {employees.length === 0 && (
                  <tr><td colSpan={3} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.empNoneMsg}</td></tr>
                )}
                {employees.map((e) => (
                  <tr key={e.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                    <td className="px-3 py-2">{e.name}</td>
                    <td className="px-3 py-2">{e.role}</td>
                    <td className="px-3 py-2 text-right whitespace-nowrap">
                      <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => { setEditingEmp(e); setEmpForm({ ...e }); }}>{t.editBtn}</button>
                      <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => deleteEmp(e.id)}>{t.deleteBtn}</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, tone, colors }) {
  const toneColor = { grey: colors.ink, amber: colors.amber, red: colors.red, green: colors.green }[tone || "grey"];
  return (
    <div className="rounded-lg p-4 flex flex-col gap-1" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <span className="text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{label}</span>
      <span className="text-3xl font-bold" style={{ color: toneColor, fontFamily: FONT_MONO }}>{value}</span>
    </div>
  );
}

function downloadTemplate() {
  const headerRow = FIELD_DEFS.map((f) => f.label);
  const ws = XLSX.utils.aoa_to_sheet([headerRow]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Import Template");
  XLSX.writeFile(wb, "farspeed-import-template.xlsx");
}

function exportToExcel(items) {
  const inventoryRows = items.map((it) => {
    const row = {};
    FIELD_DEFS.forEach((f) => { row[f.label] = it[f.key] ?? ""; });
    row["Status"] = deriveStatus(it).replace("_", " ");
    row["Total Units"] = totalUnits(it);
    row["Delivered Units"] = deliveredUnits(it);
    row["Remaining Units"] = remainingUnits(it);
    const info = storageInfo(it);
    row["Billable Storage Days"] = info ? info.billableDays : 0;
    return row;
  });

  const deliveryRows = [];
  items.forEach((it) => {
    (it.deliveries || []).forEach((d) => {
      deliveryRows.push({
        "Item ID": it.id, Client: it.client, Project: it.project, "Unit Code": it.unitCode,
        "Delivery Date": d.date, "Quantity Delivered": d.packageCount, "Delivered To": d.deliveredTo,
        "Received By": d.receivedBy, Notes: d.notes,
      });
    });
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(inventoryRows), "Inventory");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deliveryRows.length ? deliveryRows : [{ "Item ID": "", "Delivery Date": "", "Quantity Delivered": "" }]), "Deliveries");
  XLSX.writeFile(wb, `farspeed-depot-export-${todayStr()}.xlsx`);
}

function UploadPanel({ onImportRows, onAddIncoming, existingItems, directory, setDirectory, legacyArchive, setLegacyArchive, items, incoming, onLegacyImport, onLegacyCheckIn, onLegacyCheckInBatch, onLegacyDeliver, onLegacyEnrich, colors, t, lang }) {
  const [mode, setMode] = useState("packinglist");
  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["packinglist", t.uploadModePackingList], ["legacy", t.uploadModeLegacy]].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>
      {mode === "packinglist" && (
        <ImportPanel onImportRows={onImportRows} onAddIncoming={onAddIncoming} existingItems={existingItems} directory={directory} setDirectory={setDirectory} colors={colors} t={t} lang={lang} hideExcelMode />
      )}
      {mode === "legacy" && (
        <LegacyUploadsPanel legacyArchive={legacyArchive} setLegacyArchive={setLegacyArchive} items={items} incoming={incoming} onLegacyCheckIn={onLegacyCheckIn} onLegacyCheckInBatch={onLegacyCheckInBatch} directory={directory} onLegacyImport={onLegacyImport} onLegacyDeliver={onLegacyDeliver} onLegacyEnrich={onLegacyEnrich} colors={colors} t={t} lang={lang} />
      )}
    </div>
  );
}

function ImportPanel({ onImportRows, onAddIncoming, existingItems, directory, setDirectory, colors, t, lang, hideExcelMode }) {
  const [mode, setMode] = useState("packinglist");
  const [excelPreview, setExcelPreview] = useState(null);
  const [included, setIncluded] = useState([]);
  const [excelError, setExcelError] = useState("");
  const [plPreview, setPlPreview] = useState(null);
  const [plError, setPlError] = useState("");
  const [plCommon, setPlCommon] = useState(null);
  const [pdfStatus, setPdfStatus] = useState("idle"); // idle | scanning
  const [pdfError, setPdfError] = useState("");
  const [manualForm, setManualForm] = useState({
    client: CLIENTS[0], project: "", constructionSite: "", orderedBy: "", jobRef: "", shkNumber: "", unitCode: "", directoryId: "", saveToDirectory: false, packages: [],
  });
  const inputStyle = inputStyleFor(colors);
  const siteSuggestions = useMemo(() => {
    const fromDirectory = (directory || []).map((s) => s.siteEn).filter(Boolean);
    const fromItems = (existingItems || []).map((i) => i.project).filter(Boolean);
    return [...new Set([...fromDirectory, ...fromItems])];
  }, [directory, existingItems]);

  function applyParsedResult({ groups, client, project }) {
    if (!groups || groups.length === 0) { return false; }
    setPlPreview(groups);
    const guess = String(project || "").toLowerCase();
    const matchedSite = guess ? (directory || []).find((s) =>
      (s.siteEn && guess.includes(s.siteEn.toLowerCase())) ||
      (s.siteZh && guess.includes(s.siteZh.toLowerCase())) ||
      (s.siteEn && s.siteEn.toLowerCase().includes(guess))
    ) : null;
    const resolvedClient = resolveClientGuess(client);
    setPlCommon({
      client: matchedSite ? matchedSite.client : (resolvedClient || CLIENTS[CLIENTS.length - 1]),
      project: matchedSite ? matchedSite.siteEn : (project || ""),
      directoryId: matchedSite ? matchedSite.id : "",
      jobRef: matchedSite ? matchedSite.jobRef : "",
      orderedBy: matchedSite ? matchedSite.orderedBy : "",
      shkNumber: "",
      constructionSite: matchedSite ? (matchedSite.siteZh || matchedSite.siteEn) : (project || ""),
      saveToDirectory: !matchedSite,
    });
    return true;
  }

  async function handlePdfScanFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setPdfError("");
    setPlPreview(null);
    setPdfStatus("scanning");
    try {
      const base64 = await new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result.split(",")[1]);
        r.onerror = () => reject(new Error("read failed"));
        r.readAsDataURL(file);
      });
      const prompt = `This is a packing list, delivery memo, shipping list, or similar logistics document for elevator/escalator materials, possibly in English, Traditional or Simplified Chinese, or mixed.

Follow these extraction rules exactly - they keep the output compact even for long, dense documents:

1. Each row/block of the table is ONE case/package. For each case extract exactly these 5 things:
   - "code": the case/box number as printed (e.g. "C01", "C21", "17A10A01").
   - "lot": the lift/unit number. This is very often printed right next to the case number in parentheses, like "(#.01)" or "(#.23)" - extract just the number (e.g. "01", "23"). Different cases sharing the same case-number prefix (e.g. multiple "C21" cases) but different lift numbers are DIFFERENT packages in DIFFERENT lots. If there's no such lift marker anywhere, use the shop order number or another batch identifier as the lot instead.
   - "description": ONLY the general category/heading text for that case (e.g. "Guide Rail", "Rail Bracket", "Traction Machine", "Installation Material"). Do NOT list the individual part numbers or sub-components underneath it even if the document itemizes many - this is the single most important rule for keeping output size manageable on long documents.
   - "weightKg": use the GROSS weight (毛重 / GROSS column), not net weight - gross is what matters here.
   - "cbm": if a CBM/volume figure is already given directly for that case, use it as-is. Otherwise read that case's dimensions (often shown as "L*W*H", e.g. "500*20*20") and check the column header/label to see whether dimensions are in cm or mm. If in cm, compute cbm = (L*W*H) / 1,000,000. If in mm, compute cbm = (L*W*H) / 1,000,000,000.
2. Group all cases by their lot/lift number into the "groups" array - one entry per distinct lot.
2b. Also look for shipping/bill-of-lading details anywhere on the document: vessel/ship name (often after "ex ss." or 船名), voyage number, and container numbers. Combine them into one line for "ssDoNo" in roughly this style: ex ss."SHIP NAME" V.VOYAGE; CONTAINERS NO. XXXX/40GP. If none present, use ''.
3. Keep everything as compact as possible: short descriptions, no commentary, no repeated sub-item lists.

Respond with ONLY a raw JSON object in EXACTLY this shape and nothing else (no markdown fences, no commentary, no explanation before or after):
{"client": "best-guess client name or ''", "project": "site/building/project name found in the document, or ''", "ssDoNo": "vessel + voyage + container line or ''", "groups": [{"lot": "lift/lot/shop-order number identifying this batch", "containers": ["container numbers if any, else empty array"], "packages": [{"code": "case/package number", "description": "short category name, a few words only", "weightKg": number_or_empty_string, "cbm": number_or_empty_string}]}]}
If the document only has one overall lot/shipment with no explicit lift/case breakdown, put everything under a single group with a sensible lot name.`;
      const response = await fetch("/api/scan-pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 16000,
          messages: [{ role: "user", content: [
            { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
            { type: "text", text: prompt },
          ] }],
        }),
      });
      const data = await response.json();
      if (data.error) throw new Error(data.error.message || data.error || "API error");
      const text = (data.content || []).map((b) => b.text || "").join("");
      const clean = text.replace(/```json|```/g, "").trim();
      let parsed;
      try {
        parsed = JSON.parse(clean);
      } catch (parseErr) {
        throw new Error("truncated-or-invalid-json");
      }
      const normalizedGroups = (parsed.groups || []).map((g) => ({
        lot: g.lot || "UNSPECIFIED",
        containers: g.containers || [],
        totalWeight: (g.packages || []).reduce((s, p) => s + (Number(p.weightKg) || 0), 0),
        totalCbm: (g.packages || []).reduce((s, p) => s + (Number(p.cbm) || 0), 0),
        packages: (g.packages || []).map((p) => ({
          code: p.code || "",
          description: p.description || "",
          weightKg: p.weightKg !== "" && p.weightKg != null ? String(p.weightKg) : "",
          cbm: p.cbm !== "" && p.cbm != null ? String(p.cbm) : "",
        })),
      }));
      const ok = applyParsedResult({ groups: normalizedGroups, client: parsed.client, project: parsed.project, ssDoNo: parsed.ssDoNo || "" });
      if (!ok) setPdfError(t.packingListNoStructure);
      setPdfStatus("idle");
    } catch (err) {
      if (err.message === "truncated-or-invalid-json") {
        setPdfError(t.pdfTruncatedMsg);
      } else {
        const friendly = t.pdfReadErrorMsg || "Couldn't read this PDF. Please check the file, or enter the details manually below.";
        const detail = err && err.message ? err.message : "";
        setPdfError(detail ? `${friendly} (${detail})` : friendly);
      }
      setPdfStatus("idle");
    }
    e.target.value = "";
  }

  async function handleExcelFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setExcelError("");
    setExcelPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (rows.length === 0) { setExcelError(t.excelNoRowsMsg); return; }
      const unmatched = new Set();
      const mapped = rows.map((row) => {
        const item = emptyForm();
        Object.entries(row).forEach(([header, value]) => {
          const key = matchField(header);
          if (!key) { if (String(value).trim() !== "") unmatched.add(header); return; }
          item[key] = value instanceof Date ? dateToLocalISO(value) : String(value).trim();
        });
        return item;
      });

      const existingSigs = new Map();
      (existingItems || []).forEach((it) => { const sig = itemSignature(it); if (sig) existingSigs.set(sig, it.id); });
      const seenInBatch = new Map();
      const withDupInfo = mapped.map((item, idx) => {
        const sig = itemSignature(item);
        let dupOf = null;
        if (sig) {
          if (existingSigs.has(sig)) dupOf = existingSigs.get(sig);
          else if (seenInBatch.has(sig)) dupOf = `row ${seenInBatch.get(sig) + 1}`;
          else seenInBatch.set(sig, idx);
        }
        return { item, dupOf };
      });

      setExcelPreview({ rows: withDupInfo, unmatched: [...unmatched] });
      setIncluded(withDupInfo.map((r) => !r.dupOf));
    } catch (err) {
      setExcelError(t.excelErrorMsg);
    }
    e.target.value = "";
  }

  async function handlePackingListFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setPlError("");
    setPlPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const { groups, client, project } = parsePackingListWorkbook(wb);
      const ok = applyParsedResult({ groups, client, project });
      if (!ok) setPlError(t.packingListNoStructure);
    } catch (err) {
      setPlError(t.packingListNoStructure);
    }
    e.target.value = "";
  }

  function addToIncoming() {
    let effectiveDirectoryId = plCommon.directoryId || "";

    if (plCommon.saveToDirectory && !plCommon.directoryId && plCommon.project) {
      const newSite = {
        id: `SITE${Date.now()}`,
        siteEn: plCommon.project,
        siteZh: plCommon.constructionSite && plCommon.constructionSite !== plCommon.project ? plCommon.constructionSite : "",
        client: plCommon.client,
        jobRef: plCommon.jobRef || "",
        orderedBy: plCommon.orderedBy || "",
        accountOfficer: "",
      };
      setDirectory((d) => [...d, newSite]);
      effectiveDirectoryId = newSite.id;
    }

    const newIncoming = plPreview.map((g) => ({
      client: plCommon.client,
      project: plCommon.project,
      constructionSite: plCommon.constructionSite || "",
      jobRef: plCommon.jobRef || "",
      orderedBy: plCommon.orderedBy || "",
      shkNumber: plCommon.shkNumber || "",
      directoryId: effectiveDirectoryId,
      unitCode: g.lot,
      packages: g.packages,
      notes: g.containers.length ? `Container(s): ${g.containers.join(", ")}` : "",
    }));
    onAddIncoming(newIncoming);
    setPlPreview(null);
    setPlCommon(null);
  }

  function addManualToIncoming() {
    let effectiveDirectoryId = manualForm.directoryId || "";
    if (manualForm.saveToDirectory && !manualForm.directoryId && manualForm.project) {
      const newSite = {
        id: `SITE${Date.now()}`,
        siteEn: manualForm.project,
        siteZh: manualForm.constructionSite && manualForm.constructionSite !== manualForm.project ? manualForm.constructionSite : "",
        client: manualForm.client,
        jobRef: manualForm.jobRef || "",
        orderedBy: manualForm.orderedBy || "",
        accountOfficer: "",
      };
      setDirectory((d) => [...d, newSite]);
      effectiveDirectoryId = newSite.id;
    }
    onAddIncoming([{
      client: manualForm.client,
      project: manualForm.project,
      constructionSite: manualForm.constructionSite || "",
      jobRef: manualForm.jobRef || "",
      orderedBy: manualForm.orderedBy || "",
      shkNumber: manualForm.shkNumber || "",
      directoryId: effectiveDirectoryId,
      unitCode: manualForm.unitCode || "",
      packages: manualForm.packages || [],
      notes: t.legacyManualEntryNote,
    }]);
    setManualForm({ client: CLIENTS[0], project: "", constructionSite: "", orderedBy: "", jobRef: "", shkNumber: "", unitCode: "", directoryId: "", saveToDirectory: false, packages: [] });
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["packinglist", t.tabPackingList], ["pdf", t.tabPdf], ["manual", t.tabManualPackingList], ...(hideExcelMode ? [] : [["excel", t.tabExcel]])].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>

      {mode === "packinglist" && (
        <div className="rounded-lg p-5 flex flex-col gap-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
          <div>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.packingListTitle}</h3>
            <p className="text-sm" style={{ color: colors.inkFaint }}>{t.packingListDesc}</p>
          </div>
          <label className="px-3 py-2 rounded text-sm font-semibold cursor-pointer w-fit" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}>
            {t.choosePackingListBtn}
            <input type="file" accept=".xlsx,.xls,.xlsm,.csv" className="hidden" onChange={handlePackingListFile} />
          </label>
          {plError && <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{plError}</div>}
        </div>
      )}

      {mode === "pdf" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5 flex flex-col gap-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <div>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.pdfTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.pdfDesc}</p>
            </div>

            <label className="px-3 py-2 rounded text-sm font-semibold cursor-pointer w-fit" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}>
              {t.choosePdfBtn}
              <input type="file" accept="application/pdf" className="hidden" onChange={handlePdfScanFile} />
            </label>
            {pdfStatus === "scanning" && <div className="text-sm" style={{ color: colors.inkFaint }}>{t.scanningMsg}</div>}
            {pdfError && <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{pdfError}</div>}
          </div>
        </div>
      )}

      {plPreview && plCommon && (
            <div className="flex flex-col gap-4">
              <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
                <h4 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
                  {t.packingListCommonFieldsTitle}
                </h4>
                {(directory || []).length > 0 && (
                  <div className="mb-4">
                    <Field label={t.selectFromDirectory} colors={colors}>
                      <select
                        className={inputClass}
                        style={inputStyle}
                        value={plCommon.directoryId || ""}
                        onChange={(e) => {
                          const site = directory.find((s) => s.id === e.target.value);
                          if (!site) return;
                          setPlCommon((c) => ({
                            ...c,
                            directoryId: site.id,
                            client: CLIENTS.includes(site.client) ? site.client : c.client,
                            project: site.siteEn || c.project,
                            constructionSite: site.siteZh || site.siteEn || c.constructionSite,
                            jobRef: site.jobRef || c.jobRef,
                            orderedBy: site.orderedBy || c.orderedBy,
                            saveToDirectory: false,
                          }));
                        }}
                      >
                        <option value="">{t.selectFromDirectoryPlaceholder}</option>
                        {directory.map((s) => <option key={s.id} value={s.id}>{s.siteEn} — {s.client}</option>)}
                      </select>
                    </Field>
                  </div>
                )}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                  <Field label={t.packingListApplyClient} colors={colors}>
                    <select className={inputClass} style={inputStyle} value={plCommon.client} onChange={(e) => setPlCommon((c) => ({ ...c, client: e.target.value }))}>
                      {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label={t.packingListApplyProject} colors={colors}>
                    <input list="site-name-suggestions" className={inputClass} style={inputStyle} value={plCommon.project} onChange={(e) => setPlCommon((c) => ({ ...c, project: e.target.value, directoryId: "" }))} />
                    <datalist id="site-name-suggestions">
                      {siteSuggestions.map((s) => <option key={s} value={s} />)}
                    </datalist>
                  </Field>
                  <Field label={t.fOrderedBy} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.orderedBy || ""} onChange={(e) => setPlCommon((c) => ({ ...c, orderedBy: e.target.value }))} />
                  </Field>
                  <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.jobRef || ""} onChange={(e) => setPlCommon((c) => ({ ...c, jobRef: e.target.value }))} />
                  </Field>
                  <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.shkNumber || ""} onChange={(e) => setPlCommon((c) => ({ ...c, shkNumber: e.target.value }))} />
                  </Field>
                </div>
                <div className="mt-3 px-3 py-2 rounded text-xs" style={{ background: colors.surfaceDim, color: colors.inkFaint }}>
                  {t.packingListIncomingHint}
                </div>
                {!plCommon.directoryId && plCommon.project && (
                  <label className="flex items-center gap-2 mt-4 text-sm" style={{ color: colors.ink }}>
                    <input
                      type="checkbox"
                      checked={!!plCommon.saveToDirectory}
                      onChange={(e) => setPlCommon((c) => ({ ...c, saveToDirectory: e.target.checked }))}
                    />
                    {t.saveNewSiteToDirectory(plCommon.project)}
                  </label>
                )}
              </div>

              <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-2 text-sm font-semibold" style={{ background: colors.amberSoft, color: colors.amberText, fontFamily: FONT_DISPLAY }}>
                  {t.packingListDetectedTitle(plPreview.length)}
                </div>
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {[t.colLot, t.colPackages, t.colContainers, t.colWeight, t.colCbm, ""].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {plPreview.map((g, idx) => (
                      <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <td className="px-3 py-2">
                          <input
                            className={inputClass}
                            style={{ ...inputStyleFor(colors), minWidth: 100, fontWeight: 600 }}
                            value={g.lot}
                            onChange={(e) => setPlPreview((prev) => prev.map((grp, i) => (i === idx ? { ...grp, lot: e.target.value } : grp)))}
                          />
                        </td>
                        <td className="px-3 py-2">{g.packages.length}</td>
                        <td className="px-3 py-2 text-xs" style={{ color: colors.inkFaint }}>{g.containers.join(", ") || "—"}</td>
                        <td className="px-3 py-2">{Math.round(g.totalWeight)}</td>
                        <td className="px-3 py-2">{g.totalCbm ? Math.round(g.totalCbm * 1000) / 1000 : "—"}</td>
                        <td className="px-3 py-2 text-right">
                          <button
                            type="button"
                            title={t.packingListRemoveGroupBtn}
                            aria-label={t.packingListRemoveGroupBtn}
                            className="w-6 h-6 rounded-full inline-flex items-center justify-center font-bold"
                            style={{ background: colors.redSoft, color: colors.red, lineHeight: 1 }}
                            onClick={() => setPlPreview((prev) => prev.filter((_, i) => i !== idx))}
                          >
                            &minus;
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex gap-2">
                <button
                  className="px-4 py-2 rounded text-sm font-semibold w-fit"
                  style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: plPreview.length === 0 ? 0.5 : 1 }}
                  disabled={plPreview.length === 0}
                  onClick={addToIncoming}
                >
                  {t.packingListAddToIncomingBtn(plPreview.length)}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setPlPreview(null); setPlCommon(null); }}>
                  {t.discardBtn}
                </button>
              </div>
            </div>
      )}

      {mode === "manual" && (
        <div className="flex flex-col gap-4">
          <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.tabManualPackingList}</h3>
            <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>{t.manualPackingListDesc}</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
              <Field label={t.packingListApplyClient} colors={colors}>
                <select className={inputClass} style={inputStyle} value={manualForm.client} onChange={(e) => setManualForm((f) => ({ ...f, client: e.target.value }))}>
                  {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              <Field label={t.legacyProjectSiteEn} colors={colors}>
                <input list="manual-site-suggestions" className={inputClass} style={inputStyle} value={manualForm.project} onChange={(e) => setManualForm((f) => ({ ...f, project: e.target.value, directoryId: "" }))} />
                <datalist id="manual-site-suggestions">
                  {siteSuggestions.map((s) => <option key={s} value={s} />)}
                </datalist>
              </Field>
              <Field label={t.legacyProjectSiteZh} colors={colors}>
                <input className={inputClass} style={inputStyle} value={manualForm.constructionSite} onChange={(e) => setManualForm((f) => ({ ...f, constructionSite: e.target.value }))} />
              </Field>
              <Field label={t.fOrderedBy} colors={colors}>
                <input className={inputClass} style={inputStyle} value={manualForm.orderedBy} onChange={(e) => setManualForm((f) => ({ ...f, orderedBy: e.target.value }))} />
              </Field>
              <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
                <input className={inputClass} style={inputStyle} value={manualForm.jobRef} onChange={(e) => setManualForm((f) => ({ ...f, jobRef: e.target.value }))} />
              </Field>
              <Field label={t.fReference} hint={t.fReferenceHint} colors={colors}>
                <input className={inputClass} style={inputStyle} value={manualForm.shkNumber} onChange={(e) => setManualForm((f) => ({ ...f, shkNumber: e.target.value }))} />
              </Field>
              <Field label={t.legacyUnitCode} colors={colors}>
                <input className={inputClass} style={inputStyle} value={manualForm.unitCode} onChange={(e) => setManualForm((f) => ({ ...f, unitCode: e.target.value }))} />
              </Field>
            </div>
            {!manualForm.directoryId && manualForm.project && (
              <label className="flex items-center gap-2 mt-4 text-sm" style={{ color: colors.inkFaint }}>
                <input type="checkbox" checked={manualForm.saveToDirectory} onChange={(e) => setManualForm((f) => ({ ...f, saveToDirectory: e.target.checked }))} />
                {t.saveNewSiteToDirectory(manualForm.project)}
              </label>
            )}
            <PackagesEditor form={manualForm} setForm={setManualForm} colors={colors} t={t} />
          </div>
          <div className="flex gap-2">
            <button
              className="px-4 py-2 rounded text-sm font-semibold w-fit"
              style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: (!manualForm.client || !(manualForm.project || manualForm.constructionSite) || (manualForm.packages || []).length === 0) ? 0.5 : 1 }}
              disabled={!manualForm.client || !(manualForm.project || manualForm.constructionSite) || (manualForm.packages || []).length === 0}
              onClick={addManualToIncoming}
            >
              {t.packingListAddToIncomingBtn(1)}
            </button>
          </div>
        </div>
      )}

      {mode === "excel" && (
        <div className="rounded-lg p-5 flex flex-col gap-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
          <div>
            <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.excelTitle}</h3>
            <p className="text-sm" style={{ color: colors.inkFaint }}>{t.excelDesc}</p>
          </div>
          <div className="flex gap-3 items-center flex-wrap">
            <label className="px-3 py-2 rounded text-sm font-semibold cursor-pointer" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}>
              {t.chooseFileBtn}
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelFile} />
            </label>
            <button className="text-sm font-semibold underline" style={{ color: colors.amberText }} onClick={downloadTemplate}>
              {t.downloadTemplateBtn}
            </button>
          </div>
          {excelError && <div className="px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{excelError}</div>}
          {excelPreview && (
            <div className="flex flex-col gap-3">
              <div className="text-sm flex items-center gap-3 flex-wrap" style={{ color: colors.inkFaint }}>
                <span>{t.selectedCount(included.filter(Boolean).length, excelPreview.rows.length)}</span>
                <button className="underline text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => setIncluded(excelPreview.rows.map(() => true))}>{t.selectAllBtn}</button>
                <button className="underline text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => setIncluded(excelPreview.rows.map((r) => !r.dupOf))}>{t.selectNonDupBtn}</button>
              </div>
              {excelPreview.unmatched.length > 0 && (
                <div className="px-3 py-2 rounded text-sm" style={{ background: colors.amberSoft, color: colors.amberText }}>
                  {t.unmatchedMsg}{excelPreview.unmatched.join(", ")}
                </div>
              )}
              <div className="rounded overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
                <table className="w-full text-xs" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {["", t.prevColClient, t.prevColProject, t.prevColItemType, t.prevColDepot, t.prevColDepotArrival, t.prevColMatch].map((h) => (
                        <th key={h} className="text-left px-2 py-1.5 font-semibold" style={{ color: colors.inkFaint }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {excelPreview.rows.map((r, idx) => (
                      <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, opacity: included[idx] ? 1 : 0.5 }}>
                        <td className="px-2 py-1.5">
                          <input type="checkbox" checked={!!included[idx]} onChange={(e) => setIncluded((prev) => prev.map((v, i) => (i === idx ? e.target.checked : v)))} />
                        </td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.client}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.project}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.itemType}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.depot}</td>
                        <td className="px-2 py-1.5" style={{ color: colors.ink }}>{r.item.depotArrivalDate}</td>
                        <td className="px-2 py-1.5">{r.dupOf ? <Badge tone="amber" colors={colors}>{t.badgeDupOf(r.dupOf)}</Badge> : <Badge tone="green" colors={colors}>{t.badgeNew}</Badge>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex gap-2">
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
                  onClick={() => {
                    const rowsToImport = excelPreview.rows.filter((_, idx) => included[idx]).map((r) => r.item);
                    onImportRows(rowsToImport);
                    setExcelPreview(null);
                    setIncluded([]);
                  }}>
                  {t.importBtn(included.filter(Boolean).length)}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setExcelPreview(null); setIncluded([]); }}>
                  {t.discardBtn}
                </button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

async function ensureUserAccountsSeeded() {
  try {
    const res = await storageGet("userAccounts");
    if (res) return JSON.parse(res.value);
  } catch (e) { /* not seeded yet */ }
  const defaultHash = await hashPassword(DEFAULT_PASSWORD);
  const seeded = LOGIN_ACCOUNTS.map((name) => ({ name, passwordHash: defaultHash }));
  try { await storageSet("userAccounts", JSON.stringify(seeded)); } catch (e) {}
  return seeded;
}

function LoginScreen({ onLoggedIn, colors, t, lang }) {
  const [accounts, setAccounts] = useState(null);
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let on = true;
    ensureUserAccountsSeeded().then((accts) => { if (on) setAccounts(accts); });
    return () => { on = false; };
  }, []);

  async function handleLogin(e) {
    e.preventDefault();
    if (!name || !password) { setError(t.loginErrorMissing); return; }
    setBusy(true);
    setError("");
    try {
      const accts = accounts || (await ensureUserAccountsSeeded());
      const acct = accts.find((a) => a.name === name);
      const hash = await hashPassword(password);
      if (!acct || acct.passwordHash !== hash) {
        setError(t.loginErrorWrong);
        setBusy(false);
        return;
      }
      try {
        window.localStorage.setItem("farspeed_session", JSON.stringify({ name, at: Date.now() }));
      } catch (e2) {}
      onLoggedIn(name);
    } catch (e3) {
      setError(t.loginErrorWrong);
    }
    setBusy(false);
  }

  const inputStyle = inputStyleFor(colors);
  return (
    <div className="fixed inset-0 flex items-center justify-center p-6" style={{ background: colors.bg }}>
      <form onSubmit={handleLogin} className="w-full max-w-sm rounded-lg p-6" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <div className="text-center mb-5">
          <div className="text-xl font-bold" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>FARSPEED CONTRACTORS LTD</div>
          <div className="text-xs mt-1" style={{ color: colors.amberText, fontFamily: FONT_DISPLAY }}>{t.appSubtitle}</div>
        </div>
        <Field label={t.loginNameLabel} colors={colors}>
          <select className={inputClass} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} required>
            <option value="">{t.loginNamePlaceholder}</option>
            {LOGIN_ACCOUNTS.map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
        </Field>
        <div className="mt-3">
          <Field label={t.loginPasswordLabel} colors={colors}>
            <input type="password" className={inputClass} style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" required />
          </Field>
        </div>
        {error && (
          <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>
        )}
        <button
          type="submit"
          className="w-full mt-4 px-4 py-2.5 rounded text-sm font-semibold"
          style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: busy ? 0.6 : 1 }}
          disabled={busy}
        >
          {busy ? t.loginBusyMsg : t.loginBtn}
        </button>
        <div className="text-xs mt-4 text-center" style={{ color: colors.inkFaint }}>{t.loginDefaultPwHint}</div>
      </form>
    </div>
  );
}

// Simple "admin privilege" gate: re-enter the current user's own password before a
// destructive action is allowed to proceed. Not a separate admin role for now - just a
// friction check, per current requirements (may become a real permission later).
function AdminConfirmModal({ authUser, onConfirm, onClose, colors, t }) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const inputStyle = inputStyleFor(colors);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      const accts = await ensureUserAccountsSeeded();
      const acct = accts.find((a) => a.name === authUser);
      const hash = await hashPassword(password);
      if (!acct || acct.passwordHash !== hash) {
        setError(t.loginErrorWrong);
        setBusy(false);
        return;
      }
      onConfirm();
    } catch (e2) {
      setError(t.loginErrorWrong);
      setBusy(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-6" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="w-full max-w-sm rounded-lg p-6" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.adminConfirmTitle}</h3>
        <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>{t.adminConfirmDesc(authUser)}</p>
        <form onSubmit={handleSubmit}>
          <Field label={t.loginPasswordLabel} colors={colors}>
            <input type="password" autoFocus className={inputClass} style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} required />
          </Field>
          {error && <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>}
          <div className="flex gap-2 mt-4">
            <button type="submit" className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.red, color: "#fff", fontFamily: FONT_DISPLAY, opacity: busy ? 0.6 : 1 }} disabled={busy}>
              {busy ? t.loginBusyMsg : t.adminConfirmBtn}
            </button>
            <button type="button" className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={onClose}>
              {t.closeBtn}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

function ChangePasswordModal({ name, onClose, colors, t }) {
  const [current, setCurrent] = useState("");
  const [next, setNext] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [success, setSuccess] = useState(false);
  const [busy, setBusy] = useState(false);
  const inputStyle = inputStyleFor(colors);

  async function handleSubmit(e) {
    e.preventDefault();
    setError("");
    if (next.length < 4) { setError(t.pwTooShortMsg); return; }
    if (next !== confirm) { setError(t.pwMismatchMsg); return; }
    setBusy(true);
    try {
      const accts = await ensureUserAccountsSeeded();
      const idx = accts.findIndex((a) => a.name === name);
      const currentHash = await hashPassword(current);
      if (idx === -1 || accts[idx].passwordHash !== currentHash) {
        setError(t.pwCurrentWrongMsg);
        setBusy(false);
        return;
      }
      const newHash = await hashPassword(next);
      const updated = accts.map((a, i) => (i === idx ? { ...a, passwordHash: newHash } : a));
      await storageSet("userAccounts", JSON.stringify(updated));
      setSuccess(true);
    } catch (e2) {
      setError(t.pwCurrentWrongMsg);
    }
    setBusy(false);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-6" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="w-full max-w-sm rounded-lg p-6" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
        <h3 className="text-lg font-bold mb-4" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.changePasswordTitle}</h3>
        {success ? (
          <>
            <div className="px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>{t.pwChangedMsg}</div>
            <button className="w-full mt-4 px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>{t.closeBtn}</button>
          </>
        ) : (
          <form onSubmit={handleSubmit}>
            <Field label={t.pwCurrentLabel} colors={colors}>
              <input type="password" className={inputClass} style={inputStyle} value={current} onChange={(e) => setCurrent(e.target.value)} required />
            </Field>
            <div className="mt-3">
              <Field label={t.pwNewLabel} colors={colors}>
                <input type="password" className={inputClass} style={inputStyle} value={next} onChange={(e) => setNext(e.target.value)} required />
              </Field>
            </div>
            <div className="mt-3">
              <Field label={t.pwConfirmLabel} colors={colors}>
                <input type="password" className={inputClass} style={inputStyle} value={confirm} onChange={(e) => setConfirm(e.target.value)} required />
              </Field>
            </div>
            {error && <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>}
            <div className="flex gap-2 mt-4">
              <button type="submit" className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY, opacity: busy ? 0.6 : 1 }} disabled={busy}>
                {busy ? t.loginBusyMsg : t.pwSaveBtn}
              </button>
              <button type="button" className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={onClose}>
                {t.closeBtn}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}

export default function FarspeedInventory() {
  const [authUser, setAuthUser] = useState(undefined); // undefined = checking session, null = logged out, string = logged in
  const [changePasswordOpen, setChangePasswordOpen] = useState(false);
  const [items, setItems] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("dashboard");
  const [editing, setEditing] = useState(null);
  const [exitingItems, setExitingItems] = useState([]);
  const [deliveryPickerSelection, setDeliveryPickerSelection] = useState([]);
  const [deliverySearch, setDeliverySearch] = useState("");
  const [deliveryFilterClient, setDeliveryFilterClient] = useState("All");
  const [deliveryFilterDepot, setDeliveryFilterDepot] = useState("All");
  const [deliveryFilterStatus, setDeliveryFilterStatus] = useState("All");
  const [printJobSheet, setPrintJobSheet] = useState(null);
  const [expandedRowId, setExpandedRowId] = useState(null);
  const [siteTotalsOpen, setSiteTotalsOpen] = useState(true);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [newEntryMenuOpen, setNewEntryMenuOpen] = useState(false);
  const [directory, setDirectoryState] = useState([]);
  const [employees, setEmployeesState] = useState([]);
  const [freeRules, setFreeRulesState] = useState([]);
  const [cbmRates, setCbmRatesState] = useState({});
  const [legacyArchive, setLegacyArchiveState] = useState([]);
  const [incoming, setIncomingState] = useState([]);
  const [currentUser, setCurrentUserState] = useState("");
  const [filterClient, setFilterClient] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [filterDepot, setFilterDepot] = useState("All");
  const [search, setSearch] = useState("");
  const [error, setError] = useState("");
  const [lang, setLang] = useState("en");
  const [theme, setTheme] = useState("light");

  const t = TEXT[lang];
  const colors = theme === "dark" ? DARK_COLORS : LIGHT_COLORS;

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem("farspeed_session");
      if (raw) {
        const session = JSON.parse(raw);
        if (session && session.name && LOGIN_ACCOUNTS.includes(session.name)) {
          setAuthUser(session.name);
          return;
        }
      }
    } catch (e) {}
    setAuthUser(null);
  }, []);

  function handleLoggedIn(name) {
    setAuthUser(name);
    setCurrentUser(name);
  }
  function handleLogout() {
    try { window.localStorage.removeItem("farspeed_session"); } catch (e) {}
    setAuthUser(null);
  }

  useEffect(() => {
    (async () => {
      try {
        const res = await storageGet("items");
        setItems(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setItems([]);
      }
      try {
        const res = await storageGet("directory");
        setDirectoryState(res ? JSON.parse(res.value) : DEFAULT_DIRECTORY);
        if (!res) storageSet("directory", JSON.stringify(DEFAULT_DIRECTORY));
      } catch (e) {
        setDirectoryState(DEFAULT_DIRECTORY);
      }
      try {
        const res = await storageGet("employees");
        setEmployeesState(res ? JSON.parse(res.value) : DEFAULT_EMPLOYEES);
        if (!res) storageSet("employees", JSON.stringify(DEFAULT_EMPLOYEES));
      } catch (e) {
        setEmployeesState(DEFAULT_EMPLOYEES);
      }
      try {
        const res = await storageGet("freeStorageRules");
        const rules = res ? JSON.parse(res.value) : [];
        setFreeRulesState(rules);
        setFreeStorageRulesGlobal(rules);
      } catch (e) {
        setFreeRulesState([]);
        setFreeStorageRulesGlobal([]);
      }
      try {
        const res = await storageGet("cbmRates", true);
        const rates = res ? JSON.parse(res.value) : {};
        setCbmRatesState(rates);
        setCbmRateOverridesGlobal(rates);
      } catch (e) {
        setCbmRatesState({});
        setCbmRateOverridesGlobal({});
      }
      try {
        const res = await storageGet("legacyArchive");
        setLegacyArchiveState(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setLegacyArchiveState([]);
      }
      try {
        const res = await storageGet("incoming");
        setIncomingState(res ? JSON.parse(res.value) : []);
      } catch (e) {
        setIncomingState([]);
      }
      try {
        setCurrentUserState(window.localStorage.getItem("farspeed_current_user") || "");
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  function setDirectory(updater) {
    setDirectoryState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("directory", JSON.stringify(next));
      return next;
    });
  }
  function setEmployees(updater) {
    setEmployeesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("employees", JSON.stringify(next));
      return next;
    });
  }
  function setFreeRules(updater) {
    setFreeRulesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("freeStorageRules", JSON.stringify(next));
      setFreeStorageRulesGlobal(next);
      return next;
    });
  }
  function setCbmRates(updater) {
    setCbmRatesState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("cbmRates", JSON.stringify(next), true);
      setCbmRateOverridesGlobal(next);
      return next;
    });
  }
  function setLegacyArchive(updater) {
    setLegacyArchiveState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("legacyArchive", JSON.stringify(next));
      return next;
    });
  }
  function setIncoming(updater) {
    setIncomingState((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      storageSet("incoming", JSON.stringify(next));
      return next;
    });
  }
  // Adds shipments from a Packing List upload to the Incoming pool - informational only,
  // not yet checked into the depot via Devan or CFS.
  function handleAddIncoming(rows) {
    let counter = incoming.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const newRows = rows.map((r) => {
      counter += 1;
      return { ...r, numericId: counter, id: `INC-${String(counter).padStart(4, "0")}`, checkedInCodes: [], linkedItemId: null, createdAt: todayStr() };
    });
    setIncoming((prev) => [...prev, ...newRows]);
    return newRows;
  }
  // Checks selected cases from an Incoming shipment into the depot via Devan or CFS.
  // First check-in for a shipment creates the real inventory item (carrying the full case
  // pool so later partial check-ins can keep adding arrival batches to the same item,
  // reusing the existing Split Arrival mechanism); later check-ins just add another batch.
  // Batch-safe version of handleCheckIn for when several check-ins happen in one go (e.g.
  // Legacy Upload processing several Devan/CFS rows at once). Calling handleCheckIn
  // synchronously in a loop would have each call read the same stale `items`/`incoming`
  // snapshot, risking duplicate FS-#### ids or one check-in silently overwriting another's
  // arrival batch. This computes everything against one working copy instead.
  function handleCheckInBatch(operations) {
    let workingItems = [...items];
    let counter = workingItems.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const incomingPatches = new Map();
    const results = [];
    for (const op of operations) {
      const inc = incoming.find((i) => i.id === op.incomingId);
      if (!inc || op.codes.length === 0) { results.push(null); continue; }
      const batch = { id: `ARR${Date.now()}${Math.floor(Math.random() * 1000)}`, date: op.date, type: op.type, codes: op.codes };
      const prior = incomingPatches.get(op.incomingId);
      const effectiveLinkedId = prior ? prior.linkedItemId : inc.linkedItemId;
      let resultItemId;
      if (!effectiveLinkedId) {
        counter += 1;
        const totalWeight = inc.packages.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
        const totalCbm = inc.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0);
        const newItem = {
          ...emptyForm(),
          client: inc.client, project: inc.project, constructionSite: inc.constructionSite || "",
          jobRef: op.jobRef || inc.jobRef || "", orderedBy: inc.orderedBy || "", directoryId: inc.directoryId || "",
          itemType: "Separate Items", unitCode: inc.unitCode || "",
          depot: op.depot, depotArrivalDate: op.date, arrivingType: op.type, jobNumber: op.jobNumber,
          ssDoNo: op.type === "Devan" ? (op.ssDoNo || "") : "",
          shkNumber: op.shkNumber || inc.shkNumber || "",
          weightKg: totalWeight ? String(Math.round(totalWeight * 10) / 10) : "",
          volumeCbm: totalCbm ? String(Math.round(totalCbm * 1000) / 1000) : "",
          packages: inc.packages, arrivals: [batch], deliveries: [],
          notes: t.incomingCheckedInNote(inc.id), numericId: counter,
          id: `FS-${String(counter).padStart(4, "0")}`, createdAt: todayStr(),
        };
        workingItems = [...workingItems, newItem];
        resultItemId = newItem.id;
      } else {
        workingItems = workingItems.map((it) => (it.id === effectiveLinkedId ? { ...it, arrivals: [...(it.arrivals || []), batch] } : it));
        resultItemId = effectiveLinkedId;
      }
      incomingPatches.set(op.incomingId, { linkedItemId: resultItemId, addedCodes: [...(prior ? prior.addedCodes : []), ...op.codes] });
      results.push(resultItemId);
    }
    if (operations.length > 0) {
      persist(workingItems);
      setIncoming((prev) => prev.map((inc) => {
        const patch = incomingPatches.get(inc.id);
        if (!patch) return inc;
        return { ...inc, checkedInCodes: [...new Set([...(inc.checkedInCodes || []), ...patch.addedCodes])], linkedItemId: patch.linkedItemId };
      }));
    }
    return results;
  }
  function handleCheckIn({ incomingId, codes, type, depot, jobNumber, date, ssDoNo }) {
    const inc = incoming.find((i) => i.id === incomingId);
    if (!inc || codes.length === 0) return null;
    const batch = { id: `ARR${Date.now()}${Math.floor(Math.random() * 1000)}`, date, type, codes };
    let resultItemId = inc.linkedItemId;

    if (!inc.linkedItemId) {
      let counter = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0) + 1;
      const totalWeight = inc.packages.reduce((s, p) => s + (Number(p.weightKg) || 0), 0);
      const totalCbm = inc.packages.reduce((s, p) => s + (Number(p.cbm) || 0), 0);
      const newItem = {
        ...emptyForm(),
        client: inc.client,
        project: inc.project,
        constructionSite: inc.constructionSite || "",
        jobRef: inc.jobRef || "",
        orderedBy: inc.orderedBy || "",
        directoryId: inc.directoryId || "",
        itemType: "Separate Items",
        unitCode: inc.unitCode || "",
        depot,
        depotArrivalDate: date,
        arrivingType: type,
        jobNumber,
        ssDoNo: type === "Devan" ? (ssDoNo || "") : "",
        shkNumber: inc.shkNumber || "",
        weightKg: totalWeight ? String(Math.round(totalWeight * 10) / 10) : "",
        volumeCbm: totalCbm ? String(Math.round(totalCbm * 1000) / 1000) : "",
        packages: inc.packages,
        arrivals: [batch],
        deliveries: [],
        notes: t.incomingCheckedInNote(inc.id),
        numericId: counter,
        id: `FS-${String(counter).padStart(4, "0")}`,
        createdAt: todayStr(),
      };
      persist([...items, newItem]);
      resultItemId = newItem.id;
    } else {
      persist(items.map((it) => (it.id === inc.linkedItemId ? { ...it, arrivals: [...(it.arrivals || []), batch] } : it)));
    }

    setIncoming((prev) => prev.map((i) => (i.id === incomingId ? { ...i, checkedInCodes: [...new Set([...(i.checkedInCodes || []), ...codes])], linkedItemId: resultItemId } : i)));
    return resultItemId;
  }
  function setCurrentUser(name) {
    setCurrentUserState(name);
    try {
      window.localStorage.setItem("farspeed_current_user", name);
    } catch (e) {}
  }

  function handleResetDeliveries() {
    if (!window.confirm(t.resetConfirmMsg)) return;
    persist(items.map((i) => ({ ...i, deliveries: [] })));
    window.alert(t.resetDoneMsg);
  }

  async function persist(next) {
    setItems(next);
    try {
      const res = await storageSet("items", JSON.stringify(next));
      if (!res) setError(t.saveErrorMsg);
      else setError("");
    } catch (e) {
      setError(t.saveErrorMsg);
    }
  }

  function nextId() {
    const max = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    return { numericId: max + 1, id: `FS-${String(max + 1).padStart(4, "0")}` };
  }

  function handleSave(form) {
    if (editing) {
      persist(items.map((i) => (i.id === editing.id ? { ...editing, ...form } : i)));
    } else {
      const idFields = nextId();
      persist([...items, { ...idFields, ...form, createdAt: todayStr() }]);
    }
    setEditing(null);
    setView("inventory");
  }

  function handleDelete(id) {
    persist(items.filter((i) => i.id !== id));
  }

  function handleAddDelivery(delivery, itemId) {
    const record = { ...delivery, id: `D${Date.now()}${Math.floor(Math.random() * 1000)}` };
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: [...(i.deliveries || []), record] } : i)));
    setExitingItems((prev) => prev.map((i) => (i.id === itemId ? { ...i, deliveries: [...(i.deliveries || []), record] } : i)));
    return record;
  }

  // Records one combined delivery across multiple items at once (same job number, date,
  // destination and receiver) - each item gets its own delivery entry with its own codes,
  // linked only by sharing that job number/date, and prints as a single combined job sheet.
  // Merges Legacy Upload metadata (SS/D.O., notes, missing site names) onto an item
  // that already exists - used when the job number matches an item created earlier via
  // the proper Packing List import, so its real per-case package data is never touched
  // or duplicated, only enriched.
  function handleLegacyEnrich(entries) {
    const byItemId = new Map(entries.map((e) => [e.itemId, e.patch]));
    persist(items.map((i) => (byItemId.has(i.id) ? { ...i, ...byItemId.get(i.id) } : i)));
    return entries.map((e) => ({ itemId: e.itemId }));
  }

  function handleAddCombinedDelivery(entries) {
    const records = entries.map(({ itemId, delivery }) => ({ itemId, record: { ...delivery, id: `D${Date.now()}${Math.floor(Math.random() * 10000)}-${itemId}` } }));
    const byItemId = new Map();
    for (const r of records) {
      if (!byItemId.has(r.itemId)) byItemId.set(r.itemId, []);
      byItemId.get(r.itemId).push(r.record);
    }
    persist(items.map((i) => (byItemId.has(i.id) ? { ...i, deliveries: [...(i.deliveries || []), ...byItemId.get(i.id)] } : i)));
    setExitingItems((prev) => prev.map((i) => (byItemId.has(i.id) ? { ...i, deliveries: [...(i.deliveries || []), ...byItemId.get(i.id)] } : i)));
    return records;
  }

  function handleDeleteDelivery(deliveryId, itemId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: true } : d)) } : i)));
    setExitingItems((prev) => prev.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: true } : d)) } : i)));
  }

  function handleCancelItem(itemId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, cancelled: true } : i)));
  }

  function handleRestoreItem(itemId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, cancelled: false } : i)));
  }

  function handleRestoreDelivery(itemId, deliveryId) {
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: false } : d)) } : i)));
  }

  function handlePermanentlyDeleteDelivery(itemId, deliveryId) {
    if (!window.confirm(t.permanentDeleteConfirmMsg)) return;
    persist(items.map((i) => (i.id === itemId ? { ...i, deliveries: (i.deliveries || []).filter((d) => d.id !== deliveryId) } : i)));
  }

  function handlePermanentlyDeleteItem(itemId) {
    if (!window.confirm(t.permanentDeleteConfirmMsg)) return;
    persist(items.filter((i) => i.id !== itemId));
  }

  function handleImportRows(rows) {
    let counter = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const newItems = rows.map((r) => {
      counter += 1;
      return { ...r, numericId: counter, id: `FS-${String(counter).padStart(4, "0")}`, createdAt: todayStr() };
    });
    persist([...items, ...newItems]);
    setView("inventory");
  }

  // Same as handleImportRows but stays on the current screen and returns the created
  // items (with their assigned FS ids) - used by the Legacy Uploads bulk importer.
  function handleLegacyImport(rows) {
    let counter = items.reduce((m, i) => Math.max(m, i.numericId || 0), 0);
    const newItems = rows.map((r) => {
      counter += 1;
      return { ...r, numericId: counter, id: `FS-${String(counter).padStart(4, "0")}`, createdAt: todayStr() };
    });
    persist([...items, ...newItems]);
    return newItems;
  }

  function handleKeepOne(groupIds, keepId) {
    persist(items.filter((i) => !(groupIds.includes(i.id) && i.id !== keepId)));
  }

  function handleDeleteGroup(groupIds) {
    persist(items.filter((i) => !groupIds.includes(i.id)));
  }

  const activeItemsList = useMemo(() => items.filter((i) => !i.cancelled), [items]);

  const filtered = useMemo(() => {
    return activeItemsList
      .filter((i) => filterClient === "All" || i.client === filterClient)
      .filter((i) => filterStatus === "All" || deriveStatus(i) === filterStatus)
      .filter((i) => filterDepot === "All" || i.depot === filterDepot)
      .filter((i) => {
        if (!search.trim()) return true;
        const q = search.toLowerCase();
        return (
          i.project?.toLowerCase().includes(q) || i.description?.toLowerCase().includes(q) ||
          i.client?.toLowerCase().includes(q) || i.id?.toLowerCase().includes(q) ||
          i.shkNumber?.toLowerCase().includes(q) || i.unitCode?.toLowerCase().includes(q) ||
          i.constructionSite?.toLowerCase().includes(q) || i.invoiceNumber?.toLowerCase().includes(q) ||
          i.jobNumber?.toLowerCase().includes(q) || i.ssDoNo?.toLowerCase().includes(q) ||
          (i.packages || []).some((p) => (p.code || "").toLowerCase().includes(q)) ||
          (i.deliveries || []).some((d) => (d.jobNumber || "").toLowerCase().includes(q))
        );
      })
      .sort((a, b) => b.numericId - a.numericId);
  }, [activeItemsList, filterClient, filterStatus, filterDepot, search]);

  const atDepot = activeItemsList.filter((i) => deriveStatus(i) === "at_depot");
  const partial = activeItemsList.filter((i) => deriveStatus(i) === "partial");
  const openForDelivery = [...atDepot, ...partial];
  const filteredForDelivery = openForDelivery.filter((i) => {
    if (deliveryFilterClient !== "All" && i.client !== deliveryFilterClient) return false;
    if (deliveryFilterDepot !== "All" && i.depot !== deliveryFilterDepot) return false;
    if (deliveryFilterStatus !== "All" && deriveStatus(i) !== deliveryFilterStatus) return false;
    if (!deliverySearch.trim()) return true;
    const q = deliverySearch.toLowerCase();
    return (
      i.client?.toLowerCase().includes(q) || i.project?.toLowerCase().includes(q) ||
      i.constructionSite?.toLowerCase().includes(q) || i.unitCode?.toLowerCase().includes(q) ||
      i.jobNumber?.toLowerCase().includes(q) || i.id?.toLowerCase().includes(q) ||
      i.shkNumber?.toLowerCase().includes(q) ||
      (i.packages || []).some((p) => (p.code || "").toLowerCase().includes(q))
    );
  });
  const pending = activeItemsList.filter((i) => deriveStatus(i) === "pending_collection");
  const billable = openForDelivery.filter((i) => storageInfo(i)?.billable);
  const lfdWarnings = activeItemsList.filter((i) => { const a = lfdAlert(i); return a && (a.level === "soon" || a.level === "overdue"); });
  const duplicateGroups = useMemo(() => findDuplicateGroups(activeItemsList), [activeItemsList]);
  const duplicateIds = useMemo(() => new Set(duplicateGroups.flat().map((i) => i.id)), [duplicateGroups]);
  const siteTotals = useMemo(() => {
    const map = {};
    openForDelivery.forEach((it) => {
      const key = it.directoryId || sigPart(it.project) || "unspecified";
      if (!map[key]) {
        const dirEntry = it.directoryId ? (directory || []).find((d) => d.id === it.directoryId) : null;
        map[key] = {
          key,
          label: dirEntry ? dirEntry.siteEn : (it.project || "—"),
          labelZh: dirEntry ? dirEntry.siteZh : (it.constructionSite || ""),
          client: it.client,
          cbm: 0,
          kg: 0,
          count: 0,
        };
      }
      map[key].cbm += remainingVolumeCbm(it);
      map[key].kg += remainingWeightKg(it);
      map[key].count += 1;
    });
    return Object.values(map)
      .map((s) => ({ ...s, cbm: Math.round(s.cbm * 1000) / 1000, kg: Math.round(s.kg * 10) / 10 }))
      .sort((a, b) => b.cbm - a.cbm);
  }, [openForDelivery, directory]);
  const jobLog = useMemo(() => {
    const rows = [];
    items.forEach((it) => {
      if (it.jobNumber && !it.cancelled) {
        rows.push({
          jobNumber: it.jobNumber,
          type: it.arrivingType,
          date: it.depotArrivalDate,
          client: it.client,
          site: it.project,
          recordedBy: it.recordedBy,
          sheet: { type: it.arrivingType, item: it },
        });
      }
      (it.deliveries || []).forEach((d) => {
        if (d.jobNumber && !d.cancelled) {
          rows.push({
            jobNumber: d.jobNumber,
            type: "Delivery",
            date: d.date,
            client: it.client,
            site: it.project,
            recordedBy: d.recordedBy,
            sheet: { type: "Delivery", item: it, delivery: d },
          });
        }
      });
    });
    return rows.sort((a, b) => b.jobNumber.localeCompare(a.jobNumber));
  }, [items]);

  const cancelledJobs = useMemo(() => {
    const rows = [];
    items.forEach((it) => {
      if (it.jobNumber && it.cancelled) {
        rows.push({
          jobNumber: it.jobNumber,
          type: it.arrivingType,
          date: it.depotArrivalDate,
          client: it.client,
          site: it.project,
          recordedBy: it.recordedBy,
          sheet: { type: it.arrivingType, item: it },
          onRestore: () => handleRestoreItem(it.id),
          onPurge: () => handlePermanentlyDeleteItem(it.id),
        });
      }
      (it.deliveries || []).forEach((d) => {
        if (d.jobNumber && d.cancelled) {
          rows.push({
            jobNumber: d.jobNumber,
            type: "Delivery",
            date: d.date,
            client: it.client,
            site: it.project,
            recordedBy: d.recordedBy,
            sheet: { type: "Delivery", item: it, delivery: d },
            onRestore: () => handleRestoreDelivery(it.id, d.id),
            onPurge: () => handlePermanentlyDeleteDelivery(it.id, d.id),
          });
        }
      });
    });
    return rows.sort((a, b) => b.jobNumber.localeCompare(a.jobNumber));
  }, [items]);

  if (authUser === undefined) {
    return <div className="p-8 text-sm" style={{ color: colors.inkFaint }}>{t.loadingMsg}</div>;
  }
  if (authUser === null) {
    return <LoginScreen onLoggedIn={handleLoggedIn} colors={colors} t={t} lang={lang} />;
  }

  if (!loaded) {
    return <div className="p-8 text-sm" style={{ color: colors.inkFaint }}>{t.loadingMsg}</div>;
  }

  return (
    <div style={{ background: colors.bg, minHeight: "100%", fontFamily: FONT_BODY }} className="w-full">
      <style>{FONT_IMPORT}</style>

      <div style={{ background: colors.navy }} className="px-4 md:px-6 py-3 md:py-4 flex items-center justify-between gap-3">
        <div className="min-w-0">
          <div className="text-lg md:text-2xl font-bold tracking-tight truncate" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>FARSPEED CONTRACTORS LTD</div>
          <div className="text-[10px] md:text-xs tracking-widest uppercase" style={{ color: colors.amber }}>{t.appSubtitle}</div>
        </div>
        <button
          className="md:hidden flex items-center justify-center rounded"
          style={{ width: 40, height: 40, background: colors.navySoft, color: colors.onDark, fontSize: 20, flexShrink: 0 }}
          onClick={() => setMobileMenuOpen((o) => !o)}
          aria-label={t.mobileMenuLabel}
        >
          {mobileMenuOpen ? "\u2715" : "\u2630"}
        </button>
        <div className="hidden md:flex items-center gap-3 flex-wrap">
          <div className="flex gap-1 rounded-lg p-1" style={{ background: colors.navySoft }}>
            {[
              ["dashboard", t.navDashboard],
              ["inventory", t.navInventory],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setNewEntryMenuOpen(false); setSettingsOpen(false); setView(k); }}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark }}>
                {label}
              </button>
            ))}

            <button
              onClick={() => { setEditing(null); setSettingsOpen(false); setView("add"); }}
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, background: view === "add" ? colors.amber : "transparent", color: view === "add" ? colors.ink : colors.onDark }}
            >
              {t.navNewEntry}
            </button>

            {[
              ["upload", t.navUpload],
              ["incoming", t.navIncoming],
              ["billing", t.navBilling],
              ["directory", t.navDirectory],
              ["joblog", t.navJobLog],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setNewEntryMenuOpen(false); setSettingsOpen(false); setView(k); }}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark }}>
                {label}
              </button>
            ))}

            <div className="relative">
              <button
                onClick={() => { setSettingsOpen((o) => !o); setNewEntryMenuOpen(false); }}
                title={t.settingsLabel}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: ["duplicates", "cancelledjobs"].includes(view) ? colors.amber : "transparent", color: ["duplicates", "cancelledjobs"].includes(view) ? colors.ink : colors.onDark }}
              >
                ⚙
              </button>
              {settingsOpen && (
                <div className="absolute right-0 mt-1 rounded-lg overflow-hidden z-20" style={{ background: colors.surface, border: `1px solid ${colors.line}`, minWidth: 180 }}>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}
                    onClick={() => { setView("duplicates"); setSettingsOpen(false); }}
                  >
                    {duplicateGroups.length > 0 ? t.navDuplicatesCount(duplicateGroups.length) : t.navDuplicatesShort}
                  </button>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY, borderTop: `1px solid ${colors.surfaceDim}` }}
                    onClick={() => { setView("cancelledjobs"); setSettingsOpen(false); }}
                  >
                    {cancelledJobs.length > 0 ? `${t.navCancelledJobs} (${cancelledJobs.length})` : t.navCancelledJobs}
                  </button>
                </div>
              )}
            </div>
          </div>
          <div className="flex items-center gap-2 rounded-lg px-2 py-1" style={{ background: colors.navySoft }}>
            <span className="text-xs" style={{ color: colors.onDark, opacity: 0.7, fontFamily: FONT_DISPLAY }}>{t.signedInAs}:</span>
            <span className="text-sm font-semibold" style={{ color: colors.amber, fontFamily: FONT_DISPLAY }}>{authUser}</span>
            <button className="text-xs font-semibold underline" style={{ color: colors.onDark, opacity: 0.8 }} onClick={() => setChangePasswordOpen(true)}>{t.changePasswordLink}</button>
            <button className="text-xs font-semibold underline" style={{ color: colors.onDark, opacity: 0.8 }} onClick={handleLogout}>{t.logoutBtn}</button>
          </div>
          <div className="flex gap-1 rounded-lg p-1" style={{ background: colors.navySoft }}>
            <button
              title={t.langToggleLabel}
              onClick={() => setLang((l) => (l === "en" ? "zh" : "en"))}
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark }}
            >
              {lang === "en" ? "中文" : "EN"}
            </button>
            <button
              title={t.themeToggleLabel}
              onClick={() => setTheme((th) => (th === "light" ? "dark" : "light"))}
              className="px-3 py-1.5 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark }}
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </div>
      </div>

      {mobileMenuOpen && (
        <div className="md:hidden flex flex-col" style={{ background: colors.navySoft, borderTop: `1px solid ${colors.navy}` }}>
          {[
            ["dashboard", t.navDashboard],
            ["inventory", t.navInventory],
            ["add", t.navNewEntry],
            ["upload", t.navUpload],
            ["incoming", t.navIncoming],
            ["billing", t.navBilling],
            ["directory", t.navDirectory],
            ["joblog", t.navJobLog],
            ["duplicates", duplicateGroups.length > 0 ? t.navDuplicatesCount(duplicateGroups.length) : t.navDuplicatesShort],
            ["cancelledjobs", cancelledJobs.length > 0 ? `${t.navCancelledJobs} (${cancelledJobs.length})` : t.navCancelledJobs],
          ].map(([k, label]) => (
            <button
              key={k}
              onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setNewEntryMenuOpen(false); setSettingsOpen(false); setMobileMenuOpen(false); setView(k); }}
              className="text-left px-5 py-3 text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark, borderBottom: `1px solid ${colors.navy}` }}
            >
              {label}
            </button>
          ))}
          <div className="flex items-center justify-between px-5 py-3" style={{ borderBottom: `1px solid ${colors.navy}` }}>
            <span className="text-xs" style={{ color: colors.onDark, opacity: 0.7, fontFamily: FONT_DISPLAY }}>{t.signedInAs}: <span className="font-semibold" style={{ color: colors.amber }}>{authUser}</span></span>
          </div>
          <button
            className="text-left px-5 py-3 text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, borderBottom: `1px solid ${colors.navy}` }}
            onClick={() => { setChangePasswordOpen(true); setMobileMenuOpen(false); }}
          >
            {t.changePasswordLink}
          </button>
          <button
            className="text-left px-5 py-3 text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, borderBottom: `1px solid ${colors.navy}` }}
            onClick={handleLogout}
          >
            {t.logoutBtn}
          </button>
          <div className="flex gap-2 px-5 py-3">
            <button
              className="flex-1 px-3 py-2 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, background: colors.navy }}
              onClick={() => setLang((l) => (l === "en" ? "zh" : "en"))}
            >
              {lang === "en" ? "中文" : "EN"}
            </button>
            <button
              className="flex-1 px-3 py-2 rounded text-sm font-semibold"
              style={{ fontFamily: FONT_DISPLAY, color: colors.onDark, background: colors.navy }}
              onClick={() => setTheme((th) => (th === "light" ? "dark" : "light"))}
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </div>
      )}

      <div className="p-3 md:p-6 max-w-6xl mx-auto">
        {error && <div className="mb-4 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>{error}</div>}

        {view === "dashboard" && (
          <div className="flex flex-col gap-6">
            <div className="flex justify-end">
              <button
                className="text-xs font-semibold underline"
                style={{ color: colors.inkFaint }}
                onClick={handleResetDeliveries}
              >
                {t.resetBtn}
              </button>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <StatCard label={t.statAtDepot} value={openForDelivery.length} colors={colors} />
              <StatCard label={t.statPending} value={pending.length} tone="grey" colors={colors} />
              <StatCard label={t.statBillable} value={billable.length} tone="red" colors={colors} />
              <StatCard label={t.statLfd} value={lfdWarnings.length} tone="amber" colors={colors} />
            </div>

            {duplicateGroups.length > 0 && (
              <div className="rounded-lg p-4 flex items-center justify-between flex-wrap gap-2" style={{ background: colors.amberSoft }}>
                <span className="text-sm" style={{ color: colors.amberText }}>{t.dupBanner(duplicateGroups.length)}</span>
                <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => setView("duplicates")}>
                  {t.reviewDuplicatesBtn}
                </button>
              </div>
            )}

            {lfdWarnings.length > 0 && (
              <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
                <h3 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.red }}>{t.lfdSectionTitle}</h3>
                <div className="flex flex-col gap-2">
                  {lfdWarnings.map((i) => (
                    <div key={i.id} className="flex items-center justify-between text-sm py-1.5" style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                      <span style={{ fontFamily: FONT_MONO }}>{i.id}</span>
                      <span className="flex-1 mx-3 truncate">{i.client} · {i.project}</span>
                      <span className="mr-3" style={{ color: colors.inkFaint }}>LFD {fmt(i.terminalLFD)}</span>
                      <StatusBadge item={i} colors={colors} t={t} />
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.depotOverviewTitle}</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {DEPOTS.map((d) => {
                  const totals = depotRemainingTotals(activeItemsList, d);
                  return (
                    <div key={d} className="rounded-lg p-3" style={{ background: colors.surfaceDim }}>
                      <div className="text-sm font-semibold mb-2" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>
                        {depotLabel(d, lang)} <span style={{ color: colors.inkFaint, fontWeight: 400 }}>({totals.count} {t.depotOverviewItemsLabel})</span>
                      </div>
                      <div className="flex gap-6">
                        <div>
                          <div className="text-2xl font-bold" style={{ color: colors.ink, fontFamily: FONT_MONO }}>{totals.cbm}</div>
                          <div className="text-xs" style={{ color: colors.inkFaint }}>{t.jsCbm}</div>
                        </div>
                        <div>
                          <div className="text-2xl font-bold" style={{ color: colors.ink, fontFamily: FONT_MONO }}>{totals.kg}</div>
                          <div className="text-xs" style={{ color: colors.inkFaint }}>{t.jsKgs}</div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="rounded-lg p-4" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-sm font-bold uppercase tracking-wider mb-3" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.billableSectionTitle}</h3>
              {billable.length === 0 ? (
                <p className="text-sm" style={{ color: colors.inkFaint }}>{t.billableEmptyMsg}</p>
              ) : (
                <div className="flex flex-col gap-2">
                  {billable.map((i) => {
                    const info = storageInfo(i);
                    return (
                      <div key={i.id} className="flex items-center justify-between text-sm py-1.5" style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <span style={{ fontFamily: FONT_MONO }}>{i.id}</span>
                        <span className="flex-1 mx-3 truncate">{i.client} · {i.project}</span>
                        <span className="mr-3" style={{ color: colors.inkFaint }}>{t.sinceLabel} {fmt(info.freeUntil)}</span>
                        <StatusBadge item={i} colors={colors} t={t} />
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        )}

        {view === "inventory" && (
          <div className="flex flex-col gap-4">
            <div className="flex justify-end">
              <button
                onClick={() => { setEditing(null); setExitingItems([]); setDeliveryPickerSelection([]); setView("exit"); }}
                className="px-4 py-2 rounded text-sm font-semibold"
                style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
              >
                {t.navDeliveries} &rarr;
              </button>
            </div>
            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <div
                className="px-4 py-2 flex items-center justify-between cursor-pointer"
                style={{ background: colors.surfaceDim }}
                onClick={() => setSiteTotalsOpen((o) => !o)}
              >
                <span className="text-sm font-bold" style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}>{t.siteTotalsTitle}</span>
                <span className="text-xs font-semibold" style={{ color: colors.amberText }}>
                  {siteTotalsOpen ? t.siteTotalsToggleHide : t.siteTotalsToggleShow}
                </span>
              </div>
              {siteTotalsOpen && (
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {[t.siteTotalsColSite, t.siteTotalsColClient, t.siteTotalsColItems, t.siteTotalsColCbm, t.siteTotalsColKg].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {siteTotals.length === 0 && (
                      <tr><td colSpan={5} className="px-3 py-4 text-center text-sm" style={{ color: colors.inkFaint }}>{t.siteTotalsNoneMsg}</td></tr>
                    )}
                    {siteTotals.map((s) => (
                      <tr key={s.key} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <td className="px-3 py-2">
                          <div>{s.label}</div>
                          {s.labelZh && <div className="text-xs" style={{ color: colors.inkFaint }}>{s.labelZh}</div>}
                        </td>
                        <td className="px-3 py-2">{s.client}</td>
                        <td className="px-3 py-2">{s.count}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.cbm}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{s.kg}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>

            <div className="flex flex-wrap gap-3 items-end">
              <Field label={t.searchLabel} colors={colors}>
                <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }} placeholder={t.searchPlaceholder} value={search} onChange={(e) => setSearch(e.target.value)} />
              </Field>
              <Field label={t.clientLabel} colors={colors}>
                <select className={inputClass} style={inputStyleFor(colors)} value={filterClient} onChange={(e) => setFilterClient(e.target.value)}>
                  <option>All</option>
                  {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                </select>
              </Field>
              <Field label={t.statusLabel} colors={colors}>
                <select className={inputClass} style={inputStyleFor(colors)} value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)}>
                  <option value="All">{t.statusAll}</option>
                  <option value="pending_collection">{t.statusPending}</option>
                  <option value="at_depot">{t.statusAtDepot}</option>
                  <option value="partial">{t.statusPartial}</option>
                  <option value="delivered">{t.statusDelivered}</option>
                </select>
              </Field>
              <Field label={t.depotLabel} colors={colors}>
                <select className={inputClass} style={inputStyleFor(colors)} value={filterDepot} onChange={(e) => setFilterDepot(e.target.value)}>
                  <option>All</option>
                  {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                </select>
              </Field>
              <button className="px-3 py-1.5 rounded text-sm font-semibold ml-auto" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => exportToExcel(filtered)}>
                {t.exportBtn(filtered.length)}
              </button>
              <button className="px-3 py-1.5 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => { setEditing(null); setView("add"); }}>
                {t.newEntryBtn}
              </button>
            </div>

            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <div className="overflow-x-auto">
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {[t.colId, t.colJobNo, t.colClient, t.colProjectSite, t.colUnit, t.colDepot, t.colDepotArrival, t.colStatus, t.colPackages, t.colCbm, t.colKg, ""].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.length === 0 && (
                      <tr><td colSpan={12} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.noRecordsMsg}</td></tr>
                    )}
                    {filtered.map((i) => (
                      <React.Fragment key={i.id}>
                      <tr
                        style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, cursor: "pointer" }}
                        onClick={() => setExpandedRowId((prev) => (prev === i.id ? null : i.id))}
                      >
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{i.id}</td>
                        <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{i.jobNumber || "—"}</td>
                        <td className="px-3 py-2">{i.client}</td>
                        <td className="px-3 py-2 max-w-[220px]">
                          <div className="truncate">{i.project}</div>
                          {i.constructionSite && <div className="truncate text-xs" style={{ color: colors.inkFaint }}>{i.constructionSite}</div>}
                        </td>
                        <td className="px-3 py-2">{i.unitCode || "—"}</td>
                        <td className="px-3 py-2 whitespace-nowrap">{depotDisplay(i.depot, lang)}</td>
                        <td className="px-3 py-2">{fmt(i.depotArrivalDate)}</td>
                        <td className="px-3 py-2">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <StatusBadge item={i} colors={colors} t={t} />
                            {duplicateIds.has(i.id) && <Badge tone="amber" colors={colors}>{t.duplicateBadge}</Badge>}
                          </div>
                        </td>
                        <td className="px-3 py-2">{totalUnits(i)}</td>
                        <td className="px-3 py-2">{i.volumeCbm || "—"}</td>
                        <td className="px-3 py-2">{i.weightKg || "—"}</td>
                        <td className="px-3 py-2" onClick={(e) => e.stopPropagation()}>
                          <div className="flex flex-wrap justify-end items-center gap-x-3 gap-y-1">
                            {["at_depot", "partial"].includes(deriveStatus(i)) && (
                              <button className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => { setExitingItems([i]); setView("exit"); }}>{t.deliverBtn}</button>
                            )}
                            <button className="text-xs font-semibold" style={{ color: colors.amberText }} onClick={() => { setEditing(i); setView("add"); }}>{t.editBtn}</button>
                            {i.jobNumber && (
                              <button className="text-xs font-semibold" style={{ color: colors.inkFaint }} onClick={() => handleCancelItem(i.id)}>{t.cancelJobBtn}</button>
                            )}
                            <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => handleDelete(i.id)}>{t.deleteBtn}</button>
                          </div>
                        </td>
                      </tr>
                      {expandedRowId === i.id && (
                        <tr style={{ background: colors.surfaceDim }}>
                          <td colSpan={12} className="px-4 py-3">
                            {(() => {
                              const remaining = remainingPackages(i);
                              const remKg = Math.round(remainingWeightKg(i) * 10) / 10;
                              const remCbm = Math.round(remainingVolumeCbm(i) * 1000) / 1000;
                              return (
                                <div className="flex flex-col gap-2">
                                  <div className="flex flex-wrap gap-4 text-sm" style={{ color: colors.ink }}>
                                    <span><strong>{remainingUnits(i)}</strong> / {totalUnits(i)} {t.jsPkgs} {t.inventoryRemainingLabel}</span>
                                    <span><strong>{remCbm || 0}</strong> {t.jsCbm} {t.inventoryRemainingLabel}</span>
                                    <span><strong>{remKg || 0}</strong> {t.jsKgs} {t.inventoryRemainingLabel}</span>
                                  </div>
                                  {i.packages && i.packages.length > 0 && (
                                    remaining.length > 0 ? (
                                      <div className="flex flex-wrap gap-1.5">
                                        {(() => {
                                          const awaiting = new Set(notYetArrivedPackages(i).map((p) => p.code));
                                          return remaining.map((p) => (
                                            awaiting.has(p.code) ? (
                                              <span key={p.code} className="px-2 py-1 rounded text-xs" style={{ background: colors.surfaceDim, border: `1px dashed ${colors.line}`, color: colors.inkFaint }} title={t.notYetArrivedHint}>
                                                {p.code} · {t.notYetArrivedTag}
                                              </span>
                                            ) : (
                                              <span key={p.code} className="px-2 py-1 rounded text-xs" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.ink }}>
                                                {p.code}{p.description ? ` — ${p.description}` : ""}
                                              </span>
                                            )
                                          ));
                                        })()}
                                      </div>
                                    ) : (
                                      <span className="text-xs" style={{ color: colors.inkFaint }}>{t.inventoryNoRemainingPkgsMsg}</span>
                                    )
                                  )}
                                  {activeDeliveries(i).length > 0 && (
                                    <div className="mt-1">
                                      <div className="text-xs font-semibold uppercase tracking-wider mb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{t.deliveryHistoryLabel}</div>
                                      <div className="flex flex-col gap-1.5">
                                        {[...activeDeliveries(i)].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((d) => (
                                          <div key={d.id} className="flex flex-wrap items-center gap-2 text-xs" style={{ color: colors.ink }}>
                                            <span className="font-semibold" style={{ fontFamily: FONT_MONO }}>{d.jobNumber || "—"}</span>
                                            <span>{fmt(d.date)}</span>
                                            <span style={{ color: colors.inkFaint }}>
                                              {d.codes && d.codes.length > 0 ? d.codes.join(", ") : `${d.packageCount || "?"} ${t.jsPkgs}`}
                                            </span>
                                            <SignedDocControl docKey={`signedDoc:${i.id}:${d.id}`} colors={colors} t={t} />
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  )}
                                  <div className="flex items-center gap-2 text-xs mt-1" style={{ color: colors.inkFaint }}>
                                    <span>{t.signedDocArrivalLabel}:</span>
                                    <SignedDocControl docKey={`signedDoc:${i.id}:arrival`} colors={colors} t={t} />
                                  </div>
                                </div>
                              );
                            })()}
                          </td>
                        </tr>
                      )}
                      </React.Fragment>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {view === "exit" && (
          <div className="flex flex-col gap-4">
            {exitingItems.length > 0 ? (
              <DeliveryForm deliveryItems={exitingItems} onAddDelivery={handleAddDelivery} onAddCombinedDelivery={handleAddCombinedDelivery} onDeleteDelivery={handleDeleteDelivery}
                onCancel={() => { setExitingItems([]); setDeliveryPickerSelection([]); setView("inventory"); }} onPrintJobSheet={setPrintJobSheet}
                employees={employees} currentUser={currentUser} items={items} colors={colors} t={t} lang={lang} />
            ) : (
              <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-3 flex flex-wrap gap-3 items-end" style={{ background: colors.surfaceDim }}>
                  <Field label={t.searchLabel} colors={colors}>
                    <input className={inputClass} style={{ ...inputStyleFor(colors), minWidth: 220 }} placeholder={t.deliverySearchPlaceholder} value={deliverySearch} onChange={(e) => setDeliverySearch(e.target.value)} />
                  </Field>
                  <Field label={t.clientLabel} colors={colors}>
                    <select className={inputClass} style={inputStyleFor(colors)} value={deliveryFilterClient} onChange={(e) => setDeliveryFilterClient(e.target.value)}>
                      <option>All</option>
                      {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label={t.depotLabel} colors={colors}>
                    <select className={inputClass} style={inputStyleFor(colors)} value={deliveryFilterDepot} onChange={(e) => setDeliveryFilterDepot(e.target.value)}>
                      <option>All</option>
                      {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                    </select>
                  </Field>
                  <Field label={t.statusLabel} colors={colors}>
                    <select className={inputClass} style={inputStyleFor(colors)} value={deliveryFilterStatus} onChange={(e) => setDeliveryFilterStatus(e.target.value)}>
                      <option value="All">{t.statusAll}</option>
                      <option value="at_depot">{t.statusAtDepot}</option>
                      <option value="partial">{t.statusPartial}</option>
                    </select>
                  </Field>
                  <div className="ml-auto">
                    <button
                      className="text-sm font-semibold px-3 py-2 rounded"
                      style={{ background: deliveryPickerSelection.length > 0 ? colors.amber : colors.surfaceDim, color: deliveryPickerSelection.length > 0 ? colors.ink : colors.inkFaint, fontFamily: FONT_DISPLAY, cursor: deliveryPickerSelection.length > 0 ? "pointer" : "default" }}
                      disabled={deliveryPickerSelection.length === 0}
                      onClick={() => setExitingItems(filteredForDelivery.filter((i) => deliveryPickerSelection.includes(i.id)))}
                    >
                      {t.recordCombinedBtn(deliveryPickerSelection.length)}
                    </button>
                  </div>
                </div>
                {deliveryPickerSelection.length > 0 && (
                  <div className="px-4 py-2 text-xs" style={{ background: colors.amberSoft, color: colors.amberText }}>
                    {t.combinedSelectionHint}
                  </div>
                )}
                <div className="overflow-x-auto">
                  <table className="w-full text-sm" style={{ background: colors.surface }}>
                    <thead>
                      <tr style={{ background: colors.surfaceDim }}>
                        <th className="px-3 py-2" style={{ width: 32 }}></th>
                        {[t.colClient, t.colProjectSite, t.colUnit, t.colDepot, t.colDepotArrival, t.colStatus, ""].map((h) => (
                          <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredForDelivery.length === 0 && (
                        <tr><td colSpan={8} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.nothingAtDepotMsg}</td></tr>
                      )}
                      {filteredForDelivery.map((i) => {
                        const checked = deliveryPickerSelection.includes(i.id);
                        return (
                        <tr key={i.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, background: checked ? colors.amberSoft : "transparent" }}>
                          <td className="px-3 py-2">
                            <input type="checkbox" checked={checked} onChange={() => setDeliveryPickerSelection((prev) => checked ? prev.filter((x) => x !== i.id) : [...prev, i.id])} />
                          </td>
                          <td className="px-3 py-2">{i.client}</td>
                          <td className="px-3 py-2 max-w-[220px] truncate">{i.project}</td>
                          <td className="px-3 py-2">{i.unitCode || "—"}</td>
                          <td className="px-3 py-2 whitespace-nowrap">{depotDisplay(i.depot, lang)}</td>
                          <td className="px-3 py-2">{fmt(i.depotArrivalDate)}</td>
                          <td className="px-3 py-2"><StatusBadge item={i} colors={colors} t={t} /></td>
                          <td className="px-3 py-2 text-right">
                            <button className="text-xs font-semibold px-2 py-1 rounded" style={{ background: colors.amber, color: colors.ink }} onClick={() => setExitingItems([i])}>
                              {t.recordDeliveryBtn}
                            </button>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {view === "duplicates" && (
          <div className="flex flex-col gap-4">
            {duplicateGroups.length === 0 ? (
              <div className="rounded-lg p-6 text-sm text-center" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.inkFaint }}>
                {t.noneFoundMsg}
              </div>
            ) : (
              duplicateGroups.map((group, gi) => {
                const groupIds = group.map((i) => i.id);
                const sorted = [...group].sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));
                return (
                  <div key={gi} className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.amber}` }}>
                    <div className="px-4 py-2 flex items-center justify-between flex-wrap gap-2" style={{ background: colors.amberSoft }}>
                      <span className="text-sm font-semibold" style={{ color: colors.amberText }}>
                        {group[0].client} · {group[0].project}{group[0].unitCode ? ` · ${group[0].unitCode}` : ""} — {t.matchingEntries(group.length)}
                      </span>
                      <button className="text-xs font-semibold underline" style={{ color: colors.red }} onClick={() => handleDeleteGroup(groupIds)}>
                        {t.deleteAllBtn(group.length)}
                      </button>
                    </div>
                    <table className="w-full text-sm" style={{ background: colors.surface }}>
                      <thead>
                        <tr style={{ background: colors.surfaceDim }}>
                          {[t.colId, t.colDepot, t.colDepotArrival, t.colInvoiceNo, t.colAddedOn, t.colStatus, ""].map((h) => (
                            <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sorted.map((i) => (
                          <tr key={i.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                            <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{i.id}</td>
                            <td className="px-3 py-2">{depotDisplay(i.depot, lang)}</td>
                            <td className="px-3 py-2">{fmt(i.depotArrivalDate)}</td>
                            <td className="px-3 py-2">{i.invoiceNumber || "—"}</td>
                            <td className="px-3 py-2">{fmt(i.createdAt)}</td>
                            <td className="px-3 py-2"><StatusBadge item={i} colors={colors} t={t} /></td>
                            <td className="px-3 py-2 text-right whitespace-nowrap">
                              <button className="text-xs font-semibold mr-3" style={{ color: colors.green }} onClick={() => handleKeepOne(groupIds, i.id)}>{t.keepDeleteBtn}</button>
                              <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => handleDelete(i.id)}>{t.deleteBtn}</button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                );
              })
            )}
          </div>
        )}

        {view === "upload" && (
          <UploadPanel
            onImportRows={handleImportRows} onAddIncoming={handleAddIncoming} existingItems={items}
            directory={directory} setDirectory={setDirectory}
            legacyArchive={legacyArchive} setLegacyArchive={setLegacyArchive}
            items={items} incoming={incoming}
            onLegacyImport={handleLegacyImport}
            onLegacyCheckIn={handleCheckIn} onLegacyCheckInBatch={handleCheckInBatch}
            onLegacyDeliver={handleAddCombinedDelivery} onLegacyEnrich={handleLegacyEnrich}
            colors={colors} t={t} lang={lang}
          />
        )}

        {view === "incoming" && (
          <IncomingPanel incoming={incoming} setIncoming={setIncoming} items={items} directory={directory} setDirectory={setDirectory} onCheckIn={handleCheckIn} onAddIncoming={handleAddIncoming} colors={colors} t={t} lang={lang} />
        )}

        {view === "billing" && (
          <BillingPanel items={items} onDeleteItem={handleDelete} authUser={authUser} colors={colors} t={t} lang={lang} />
        )}

        {view === "directory" && (
          <DirectoryPanel directory={directory} setDirectory={setDirectory} employees={employees} setEmployees={setEmployees} freeRules={freeRules} setFreeRules={setFreeRules} cbmRates={cbmRates} setCbmRates={setCbmRates} legacyArchive={legacyArchive} setLegacyArchive={setLegacyArchive} items={items} incoming={incoming} onLegacyImport={handleLegacyImport} onLegacyDeliver={handleAddCombinedDelivery} onLegacyEnrich={handleLegacyEnrich} onLegacyCheckIn={handleCheckIn} onLegacyCheckInBatch={handleCheckInBatch} colors={colors} t={t} lang={lang} />
        )}

        {view === "joblog" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.jobLogTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.jobLogDesc}</p>
            </div>
            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <table className="w-full text-sm" style={{ background: colors.surface }}>
                <thead>
                  <tr style={{ background: colors.surfaceDim }}>
                    {[t.jobLogColJobNo, t.jobLogColType, t.jobLogColDate, t.jobLogColClient, t.jobLogColSite, t.jobLogColRecordedBy, ""].map((h) => (
                      <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {jobLog.length === 0 && (
                    <tr><td colSpan={7} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.jobLogNoneMsg}</td></tr>
                  )}
                  {jobLog.map((row, idx) => (
                    <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                      <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{row.jobNumber}</td>
                      <td className="px-3 py-2">
                        <Badge tone={row.type === "Delivery" ? "navy" : row.type === "Devan" ? "amber" : "green"} colors={colors}>
                          {row.type === "Devan" ? t.jsDevanType : row.type === "CFS" ? t.jsCfsType : t.jsDeliveryType}
                        </Badge>
                      </td>
                      <td className="px-3 py-2">{fmt(row.date)}</td>
                      <td className="px-3 py-2">{row.client}</td>
                      <td className="px-3 py-2 max-w-[220px] truncate">{row.site}</td>
                      <td className="px-3 py-2">{row.recordedBy || "—"}</td>
                      <td className="px-3 py-2 text-right">
                        <button className="text-xs font-semibold px-2 py-1 rounded" style={{ background: colors.amber, color: colors.ink }} onClick={() => setPrintJobSheet(row.sheet)}>
                          {t.viewReprintBtn}
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === "cancelledjobs" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.cancelledJobsTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.cancelledJobsDesc}</p>
            </div>
            <div className="rounded-lg overflow-x-auto" style={{ border: `1px solid ${colors.line}` }}>
              <table className="w-full text-sm" style={{ background: colors.surface }}>
                <thead>
                  <tr style={{ background: colors.surfaceDim }}>
                    {[t.jobLogColJobNo, t.jobLogColType, t.jobLogColDate, t.jobLogColClient, t.jobLogColSite, t.jobLogColRecordedBy, ""].map((h) => (
                      <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {cancelledJobs.length === 0 && (
                    <tr><td colSpan={7} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.cancelledJobsNoneMsg}</td></tr>
                  )}
                  {cancelledJobs.map((row, idx) => (
                    <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink, opacity: 0.75 }}>
                      <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{row.jobNumber}</td>
                      <td className="px-3 py-2">
                        <Badge tone="grey" colors={colors}>
                          {row.type === "Devan" ? t.jsDevanType : row.type === "CFS" ? t.jsCfsType : t.jsDeliveryType}
                        </Badge>
                      </td>
                      <td className="px-3 py-2">{fmt(row.date)}</td>
                      <td className="px-3 py-2">{row.client}</td>
                      <td className="px-3 py-2 max-w-[220px] truncate">{row.site}</td>
                      <td className="px-3 py-2">{row.recordedBy || "—"}</td>
                      <td className="px-3 py-2 text-right whitespace-nowrap">
                        <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => setPrintJobSheet(row.sheet)}>{t.viewReprintBtn}</button>
                        <button className="text-xs font-semibold mr-3" style={{ color: colors.green }} onClick={row.onRestore}>{t.restoreBtn}</button>
                        <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={row.onPurge}>{t.purgeBtn}</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === "add" && (
          <ItemForm
            initial={editing}
            onSave={handleSave}
            onCancel={() => { setEditing(null); setView("inventory"); }}
            onPrintJobSheet={setPrintJobSheet}
            directory={directory}
            employees={employees}
            currentUser={currentUser}
            items={items}
            colors={colors}
            t={t}
            lang={lang}
          />
        )}
      </div>

      {printJobSheet && (
        <JobSheetPrint sheet={printJobSheet} onClose={() => setPrintJobSheet(null)} directory={directory} colors={colors} t={t} lang={lang} />
      )}
      {changePasswordOpen && (
        <ChangePasswordModal name={authUser} onClose={() => setChangePasswordOpen(false)} colors={colors} t={t} />
      )}
    </div>
  );
}

