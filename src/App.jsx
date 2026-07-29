import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { storageGet, storageSet } from "./storage";

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
];
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
  caseNo: ["case no.", "case no", "case\nno", "case", "cases discript", "cases     discript"],
  qty: ["qty", "quantity", "qyt = quantity", "qyt"],
  lot: ["project no.", "project no", "lift name", "lift no.", "lift no", "lift"],
  description: ["description", "material description"],
  grossWeight: ["g.weight", "gross weight", "gross", "actual   weight", "actual weight"],
  netWeight: ["n.weight", "net weight", "net", "estimated  weight", "estimated weight"],
  cbm: ["cbm", "volume(m3)", "volume (m3)", "volume"],
  dimension: ["dimension", "dimension (mm)", "dimensions", "size"],
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
function plCbmFromDimension(v) {
  // Parses strings like "2000*850*600" or "2000 x 850 x 600" (mm) into m3.
  const parts = String(v || "").split(/[x*×]/i).map((s) => Number(s.replace(/,/g, "").trim()));
  if (parts.length !== 3 || parts.some((n) => !n || isNaN(n))) return 0;
  return (parts[0] * parts[1] * parts[2]) / 1e9;
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
  return "";
}
function parsePackingListSheet(rows) {
  const headerIdx = plDetectHeaderRow(rows);
  if (headerIdx === -1) return null;
  const colMap = plMapColumns(rows[headerIdx]);
  if (colMap.lot === undefined && colMap.caseNo === undefined) return null;

  const groups = {};
  const order = [];
  let lastLot = "", lastContainer = "", lastCase = "";
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (row.every((c) => String(c || "").trim() === "")) continue;
    const rowText = row.map((c) => String(c || "").toLowerCase()).join(" ");
    if (rowText.includes("total")) break;

    const lot = colMap.lot !== undefined ? String(row[colMap.lot] || "").trim() : "";
    const container = colMap.containerNo !== undefined ? String(row[colMap.containerNo] || "").trim() : "";
    const caseNo = colMap.caseNo !== undefined ? String(row[colMap.caseNo] || "").trim() : "";
    const description = colMap.description !== undefined ? String(row[colMap.description] || "").trim() : "";
    let weight = 0;
    if (colMap.grossWeight !== undefined && row[colMap.grossWeight] !== "" && row[colMap.grossWeight] != null) weight = plNum(row[colMap.grossWeight]);
    else if (colMap.netWeight !== undefined && row[colMap.netWeight] !== "" && row[colMap.netWeight] != null) weight = plNum(row[colMap.netWeight]);
    let cbm = 0;
    if (colMap.dimension !== undefined && row[colMap.dimension]) cbm = plCbmFromDimension(row[colMap.dimension]);
    if (!cbm && colMap.cbm !== undefined && row[colMap.cbm] !== "" && row[colMap.cbm] != null) cbm = plNum(row[colMap.cbm]);

    if (lot) lastLot = lot;
    if (container) lastContainer = container;
    if (caseNo) lastCase = caseNo;
    if (!description) continue;

    const key = lastLot || "UNSPECIFIED";
    if (!groups[key]) { groups[key] = { lot: key, packages: [], containers: new Set(), totalWeight: 0, totalCbm: 0 }; order.push(key); }
    groups[key].packages.push({ code: lastCase || String(groups[key].packages.length + 1), description, weightKg: weight ? String(weight) : "", cbm: cbm ? String(cbm) : "" });
    if (lastContainer) groups[key].containers.add(lastContainer);
    groups[key].totalWeight += weight;
    groups[key].totalCbm += cbm;
  }
  return order.map((k) => ({ ...groups[k], containers: [...groups[k].containers] }));
}
function parsePackingListWorkbook(workbook) {
  let bestGroups = null;
  let client = null;
  let project = "";
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
    if (!client) client = plGuessClient(rows);
    if (!project) project = plGuessProject(rows);
    const groups = parsePackingListSheet(rows);
    if (groups && groups.length > 0) {
      if (!bestGroups || groups.length > bestGroups.length) bestGroups = groups;
    }
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
  return new Date().toISOString().slice(0, 10);
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
  return d.toISOString().slice(0, 10);
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
  if (!item.depotArrivalDate) return "pending_collection";
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
  const freeUntil = addDays(item.depotArrivalDate, FREE_DAYS);
  const endDate = status === "delivered" ? lastDeliveryDate(item) || todayStr() : todayStr();
  const daysHeld = daysBetween(item.depotArrivalDate, endDate);
  const overFree = daysBetween(freeUntil, endDate);
  const billableDays = overFree && overFree > 0 ? overFree : 0;
  return { freeUntil, daysHeld, billableDays, billable: billableDays > 0 };
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
    searchPlaceholder: "Project, client, ID, SHK no.",
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
    badgeLfdOverdue: (d) => `LFD OVERDUE ${d}D`,
    badgeLfdSoon: (d) => `LFD IN ${d}D`,
    badgeBillable: (d) => `BILLABLE · ${d}D`,
    badgeFree: (d) => `FREE · ${d}D LEFT`,
    badgePartial: (del, tot, extra) => `PARTIAL ${del}/${tot}${extra}`,
    badgeDelivered: (extra) => `DELIVERED${extra}`,
    badgeBilledSuffix: (d) => ` · BILLED ${d}D`,
    badgeBillableSuffix: (d) => ` · ${d}D`,
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
    signedInNone: "Not signed in",

    fRecordedBy: "Recorded By",
    fRecordedByHint: "Required \u2014 who is doing this Devan / CFS / Delivery",
    recordedByRequiredMsg: "Select who is recording this before saving.",

    resetBtn: "Reset All Deliveries (demo)",
    resetConfirmMsg: "This clears every delivery record on every item so everything shows as not yet delivered. It does not delete any manifest entries. Continue?",
    resetDoneMsg: "All delivery records cleared.",

    navJobLog: "Job Log",
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
    searchPlaceholder: "項目、客戶、編號、SHK編號",
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
    signedInNone: "未登入",

    fRecordedBy: "記錄人",
    fRecordedByHint: "必填 — 由誰負責此次拆櫃／CFS／送貨",
    recordedByRequiredMsg: "儲存前請選擇記錄人。",

    resetBtn: "重設所有送貨記錄（示範用）",
    resetConfirmMsg: "此操作將清除所有項目之送貨記錄，令所有貨物顯示為尚未送貨，但不會刪除任何倉存記錄。是否繼續？",
    resetDoneMsg: "所有送貨記錄已清除。",

    navJobLog: "單號記錄",
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
    const left = FREE_DAYS - info.daysHeld;
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

      <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
        <Field label={t.fClient} colors={colors}>
          <select className={inputClass} style={inputStyle} value={form.client} onChange={set("client")}>
            {CLIENTS.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label={t.fProject} colors={colors}>
          <input className={inputClass} style={inputStyle} placeholder={t.fProjectPlaceholder} value={form.project} onChange={set("project")} />
        </Field>
        <Field label={t.fInvoiceNo} hint={t.fInvoiceHint} colors={colors}>
          <input className={inputClass} style={inputStyle} value={form.invoiceNumber} onChange={set("invoiceNumber")} />
        </Field>
      </div>

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionSite}
      </div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
        <div className="col-span-2 md:col-span-3">
          <Field label={t.fSiteName} hint={t.fSiteHint} colors={colors}>
            <input className={inputClass} style={inputStyle} placeholder={t.fSitePlaceholder} value={form.constructionSite} onChange={set("constructionSite")} />
          </Field>
        </div>
      </div>

      <div className="text-xs font-semibold uppercase tracking-wider mt-5 mb-2 pb-1" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY, borderBottom: `1px solid ${colors.surfaceDim}` }}>
        {t.sectionCargo}
      </div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
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
        <div />

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
      <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
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
      <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
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

function DeliveryForm({ item, onAddDelivery, onDeleteDelivery, onCancel, onPrintJobSheet, employees, currentUser, items, colors, t, lang }) {
  const remaining = remainingUnits(item);
  const multiUnit = totalUnits(item) > 1;
  const itemized = (item.packages || []).length > 0;
  const [form, setForm] = useState({ date: todayStr(), packageCount: multiUnit ? "" : "1", deliveredTo: item.constructionSite || item.project || "", receivedBy: "", notes: "", jobNumber: "", recordedBy: currentUser || "" });
  const [selectedCodes, setSelectedCodes] = useState([]);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const inputStyle = inputStyleFor(colors);
  const qty = Number(form.packageCount) || 0;
  const overshoot = qty > remaining;
  const remainingPkgs = remainingPackages(item);

  function toggleCode(code) {
    setSelectedCodes((prev) => (prev.includes(code) ? prev.filter((c) => c !== code) : [...prev, code]));
  }

  function handleAddClick() {
    const jobNumber = form.jobNumber || nextJobNumber(items);
    const deliveredTo = form.deliveredTo || item.constructionSite || item.project || "";
    const receivedBy = form.receivedBy || item.orderedBy || "";
    const record = itemized
      ? { ...form, jobNumber, deliveredTo, receivedBy, codes: selectedCodes }
      : { ...form, jobNumber, deliveredTo, receivedBy, packageCount: multiUnit ? qty : 1 };
    onAddDelivery(record);
    setSelectedCodes([]);
    setForm({ date: todayStr(), packageCount: multiUnit ? "" : "1", deliveredTo: item.constructionSite || item.project || "", receivedBy: "", notes: "", jobNumber: "", recordedBy: currentUser || "" });
    onPrintJobSheet({ type: "Delivery", item, delivery: record });
  }

  return (
    <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
      <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>
        {t.deliveryTitlePrefix} {item.id}
      </h3>
      <p className="text-sm mb-4" style={{ color: colors.inkFaint }}>
        {item.client} · {item.project}{item.plannedDeliveryDate ? t.plannedWasText(fmt(item.plannedDeliveryDate)) : ""}
      </p>

      <div className="px-3 py-2 rounded text-sm mb-4" style={{ background: colors.surfaceDim, color: colors.ink }}>
        <strong>{deliveredUnits(item)}</strong> {t.progressOf} <strong>{totalUnits(item)}</strong> {t.progressDeliveredSoFar} ·{" "}
        <strong>{remaining}</strong> {t.progressRemaining}
      </div>

      {activeDeliveries(item).length > 0 && (
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
              {[...activeDeliveries(item)].sort((a, b) => (a.date || "").localeCompare(b.date || "")).map((d) => (
                <tr key={d.id} style={{ borderTop: `1px solid ${colors.surfaceDim}` }}>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{fmt(d.date)}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.codes ? d.codes.join(", ") : d.packageCount}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.deliveredTo || "—"}</td>
                  <td className="px-2 py-1.5" style={{ color: colors.ink }}>{d.receivedBy || "—"}</td>
                  <td className="px-2 py-1.5" style={{ fontFamily: FONT_MONO, color: colors.ink }}>{d.jobNumber || "—"}</td>
                  <td className="px-2 py-1.5 text-right whitespace-nowrap">
                    <button className="text-xs font-semibold mr-2" style={{ color: colors.amberText }} onClick={() => onPrintJobSheet({ type: "Delivery", item, delivery: d })}>{t.printBtn}</button>
                    <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => onDeleteDelivery(d.id)}>{t.cancelJobBtn}</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {remaining > 0 ? (
        <>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
            <Field label={t.fDeliveryDate} hint={t.fDeliveryDateHint} colors={colors}>
              <input type="date" className={inputClass} style={inputStyle} value={form.date} onChange={set("date")} />
            </Field>
            {itemized ? <div /> : multiUnit ? (
              <Field label={t.fQty} hint={t.fQtyHint(remaining)} colors={colors}>
                <input type="number" min="1" max={remaining} className={inputClass} style={inputStyle} value={form.packageCount} onChange={set("packageCount")} />
              </Field>
            ) : <div />}
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

          {itemized && (
            <div className="mt-4">
              <div className="text-xs font-semibold uppercase tracking-wider mb-2" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>
                {t.selectCodesLabel}
              </div>
              <div className="flex flex-wrap gap-2">
                {remainingPkgs.map((p) => (
                  <button
                    key={p.code}
                    type="button"
                    onClick={() => toggleCode(p.code)}
                    className="px-2.5 py-1.5 rounded text-xs font-semibold text-left"
                    style={{
                      border: `1px solid ${selectedCodes.includes(p.code) ? colors.amber : colors.line}`,
                      background: selectedCodes.includes(p.code) ? colors.amberSoft : colors.surface,
                      color: selectedCodes.includes(p.code) ? colors.amberText : colors.ink,
                    }}
                    title={p.description}
                  >
                    {p.code}{p.description ? ` — ${p.description}` : ""}
                  </button>
                ))}
              </div>
            </div>
          )}

          {!itemized && overshoot && (
            <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
              {t.overshootMsg(remaining)}
            </div>
          )}

          {!form.recordedBy && (
            <div className="mt-3 px-3 py-2 rounded text-sm" style={{ background: colors.redSoft, color: colors.red }}>
              {t.recordedByRequiredMsg}
            </div>
          )}

          <div className="flex gap-2 mt-5">
            <button
              className="px-4 py-2 rounded text-sm font-semibold"
              style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }}
              disabled={(itemized ? (selectedCodes.length === 0 || !form.date) : (qty <= 0 || overshoot || !form.date)) || !form.recordedBy}
              onClick={handleAddClick}
            >
              {t.addDeliveryBtn}
            </button>
            <button
              className="px-4 py-2 rounded text-sm font-semibold"
              style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
              onClick={onCancel}
            >
              {t.closeBtn}
            </button>
          </div>
        </>
      ) : (
        <div className="flex flex-col gap-3">
          <div className="px-3 py-2 rounded text-sm" style={{ background: colors.greenSoft, color: colors.green }}>{itemized ? t.noCodesRemainingMsg : t.allDeliveredMsg}</div>
          <button
            className="px-4 py-2 rounded text-sm font-semibold w-fit"
            style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
            onClick={onCancel}
          >
            {t.closeBtn}
          </button>
        </div>
      )}
    </div>
  );
}

function JobSheetPrint({ sheet, onClose, colors, t, lang }) {
  const { type, item, delivery } = sheet;
  const isDelivery = type === "Delivery";
  const [issuedBy, setIssuedBy] = useState((isDelivery ? delivery.recordedBy : item.recordedBy) || "");
  const typeLabel = type === "Devan" ? t.jsDevanType : type === "CFS" ? t.jsCfsType : t.jsDeliveryType;

  const fromText = isDelivery
    ? `${depotDisplay(item.depot, lang)}${item.depotLocation ? " — " + item.depotLocation : ""}`
    : type === "Devan"
      ? t.jsDevanFrom(depotDisplay(item.depot, lang))
      : t.jsCfsFrom;
  const toText = item.constructionSite || item.project;
  const dateText = isDelivery ? delivery.date : item.depotArrivalDate;

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
  const jobNo = isDelivery ? delivery.jobNumber : item.jobNumber;

  const csLine = isDelivery
    ? (delivery.codes ? delivery.codes.join(", ") : "")
    : (item.packages || []).map((p) => p.code).join(", ");
  const cell = { border: "1px solid #999", padding: 6 };
  const label = { ...cell, fontWeight: "bold", background: "#F5F5F5", width: "16%" };

  return (
    <div className="fixed inset-0 z-50 flex flex-col" style={{ background: "rgba(0,0,0,0.5)" }}>
      <div className="no-print flex justify-end gap-2 p-3" style={{ background: colors.navy }}>
        <button className="px-4 py-2 rounded text-sm font-semibold" style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }} onClick={() => window.print()}>
          {t.printBtn}
        </button>
        <button className="px-4 py-2 rounded text-sm font-semibold" style={{ border: `1px solid ${colors.onDark}`, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={onClose}>
          {t.closePreviewBtn}
        </button>
      </div>
      <div className="flex-1 overflow-y-auto p-6" style={{ background: colors.bg }}>
        <div className="print-area mx-auto" style={{ background: "#fff", color: "#111", maxWidth: 800, padding: 28, fontFamily: "Arial, sans-serif", fontSize: 13 }}>

          {/* Letterhead */}
          <div className="flex items-start justify-between" style={{ marginBottom: 10 }}>
            <div className="flex items-center gap-3">
              <svg width="56" height="44" viewBox="0 0 56 44">
                <polygon points="4,40 4,4 40,40" fill="#C0392B" />
                <polygon points="10,40 22,14 34,40" fill="#2D6E5C" />
                <polygon points="18,40 32,10 46,40" fill="#2A6FB0" />
              </svg>
              <div>
                <div style={{ fontSize: 22, fontWeight: 800, color: "#2D6E5C", letterSpacing: 0.5 }}>FARSPEED<span style={{ color: "#111", fontWeight: 500 }}> Contractors Limited</span></div>
                <div style={{ fontSize: 10, color: "#333" }}>P. O. Box No. 1985, Yuen Long Post Office, Yuen Long, N.T., Hong Kong</div>
                <div style={{ fontSize: 10, color: "#333" }}>Tel: +852 5337-9500&nbsp;&nbsp;Fax: +852 2402-4450&nbsp;&nbsp;http://www.farspeed.hk</div>
              </div>
            </div>
            <div className="flex items-stretch" style={{ fontSize: 13, fontWeight: 700, lineHeight: 1.15 }}>
              <div style={{ writingMode: "vertical-rl", color: "#C0392B" }}>快達承判</div>
              <div style={{ writingMode: "vertical-rl", color: "#2D6E5C", marginLeft: 2 }}>有限公司</div>
            </div>
          </div>
          <div style={{ borderTop: "2px solid #2D6E5C", marginBottom: 10 }} />

          <div className="text-center font-bold mb-3" style={{ fontSize: 20, letterSpacing: 3 }}>
            {t.jsTitleZh}&nbsp;&nbsp;{t.jsTitle}
          </div>
          <div className="mb-2 text-xs font-bold" style={{ color: "#900" }}>{typeLabel} &nbsp;—&nbsp; JOB NO. {jobNo}</div>

          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 0 }}>
            <tbody>
              <tr>
                <td style={label}>{t.jsFromZh}<br />{t.jsFrom}</td>
                <td style={{ ...cell, width: "34%" }}>{fromText}</td>
                <td style={label}>{t.jsToZh}<br />{t.jsTo}</td>
                <td style={{ ...cell, width: "34%" }}>{toText}</td>
              </tr>
              <tr>
                <td style={label}>{t.jsAccountZh}<br />{t.jsAccount}</td>
                <td style={cell}>{item.client}</td>
                <td style={label}>{t.jsJobNoZh}<br />{t.jsJobNo}</td>
                <td style={{ ...cell, fontWeight: "bold" }}>{jobNo}</td>
              </tr>
              <tr>
                <td style={label}>{t.jsOrderedByZh}<br />{t.jsOrderedBy}</td>
                <td style={cell}>{item.orderedBy || "—"}</td>
                <td style={label}>{t.jsDateZh}<br />{t.jsDate}</td>
                <td style={cell}>{fmt(dateText)}</td>
              </tr>
              <tr>
                <td style={label}>{t.jsPoNoZh}<br />{t.jsPoNo}</td>
                <td style={cell}>{item.poNumber || "—"}</td>
                <td style={label}>{t.jsJobRefZh}<br />{t.jsJobRef}</td>
                <td style={cell}>{item.jobRef || "—"}</td>
              </tr>
            </tbody>
          </table>

          <div className="flex justify-between items-end" style={{ borderLeft: "1px solid #999", borderRight: "1px solid #999", borderTop: "1px solid #999", padding: "4px 6px" }}>
            <div className="font-bold" style={{ fontSize: 12 }}>{t.jsDescriptionZh}／{t.jsDescription}</div>
            <div style={{ fontSize: 11 }}>
              <span className="no-print">
                {t.jsIssuedByZh}: <input style={{ border: "1px solid #999", padding: "1px 4px", fontSize: 11 }} value={issuedBy} onChange={(e) => setIssuedBy(e.target.value)} />
              </span>
              <span className="print-only-inline">{t.jsIssuedByZh}: {issuedBy || "—"}</span>
            </div>
          </div>
          <div style={{ border: "1px solid #999", borderTop: "none", padding: 8, minHeight: 130, marginBottom: 16 }}>
            {item.unitCode && <div style={{ fontWeight: "bold" }}>{item.unitCode}</div>}
            {item.description && <div>{item.description}</div>}
            {csLine && <div style={{ fontSize: 11, color: "#333" }}>C/S NO. {csLine}</div>}
            <div style={{ marginTop: 8 }}>
              {pkgs} {t.jsPkgs} &nbsp;&nbsp; {kgs || "—"} {t.jsKgs} &nbsp;&nbsp; {cbm || "—"} {t.jsCbm}
            </div>
            {estimated && <div style={{ fontSize: 10, color: "#900", marginTop: 4 }}>{t.jsEstimatedNote}</div>}
          </div>

          <div style={{ fontSize: 12, marginBottom: 24 }}>
            {t.jsSignatureLine}&nbsp;_________________________________
          </div>

          <div style={{ borderTop: "1px solid #ccc", paddingTop: 6, fontSize: 9, color: "#666", textAlign: "center" }}>
            N.B. Farspeed Contractors Ltd. is a private company. All transaction(s) taken into account are subject to the STANDARD BUSINESS CONDITIONS of the company, details as behind.
            <div style={{ marginTop: 2 }}>(a member of FARSPEED Group)</div>
          </div>
        </div>
      </div>
    </div>
  );
}

function DirectoryPanel({ directory, setDirectory, employees, setEmployees, colors, t }) {
  const [mode, setMode] = useState("sites");
  const [editingSite, setEditingSite] = useState(null);
  const [siteForm, setSiteForm] = useState(null);
  const [editingEmp, setEditingEmp] = useState(null);
  const [empForm, setEmpForm] = useState(null);
  const inputStyle = inputStyleFor(colors);
  const allRoles = [...new Set([...DEFAULT_ROLES, ...employees.map((e) => e.role)])];

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
        {[["sites", t.tabSitesAccounts], ["employees", t.tabEmployees]].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className="px-3 py-1.5 rounded text-sm font-semibold"
            style={{ fontFamily: FONT_DISPLAY, background: mode === k ? colors.surface : "transparent", color: colors.ink }}>
            {label}
          </button>
        ))}
      </div>

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
              <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
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

          <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
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
              <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
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

          <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
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

function ImportPanel({ onImportRows, existingItems, directory, colors, t, lang }) {
  const [mode, setMode] = useState("packinglist");
  const [excelPreview, setExcelPreview] = useState(null);
  const [included, setIncluded] = useState([]);
  const [excelError, setExcelError] = useState("");
  const [plPreview, setPlPreview] = useState(null);
  const [plError, setPlError] = useState("");
  const [plCommon, setPlCommon] = useState(null);
  const [pdfStatus, setPdfStatus] = useState("idle"); // idle | scanning
  const [pdfError, setPdfError] = useState("");
  const inputStyle = inputStyleFor(colors);

  function applyParsedResult({ groups, client, project }) {
    if (!groups || groups.length === 0) { return false; }
    setPlPreview(groups);
    const guess = String(project || "").toLowerCase();
    const matchedSite = guess ? (directory || []).find((s) =>
      (s.siteEn && guess.includes(s.siteEn.toLowerCase())) ||
      (s.siteZh && guess.includes(s.siteZh.toLowerCase())) ||
      (s.siteEn && s.siteEn.toLowerCase().includes(guess))
    ) : null;
    setPlCommon({
      client: matchedSite ? matchedSite.client : (CLIENTS.includes(client) ? client : CLIENTS[0]),
      project: matchedSite ? matchedSite.siteEn : (project || ""),
      depot: DEPOTS[0],
      depotArrivalDate: todayStr(),
      depotLocation: "",
      arrivingType: ARRIVING_TYPES[0],
      jobNumber: "",
      directoryId: matchedSite ? matchedSite.id : "",
      jobRef: matchedSite ? matchedSite.jobRef : "",
      orderedBy: matchedSite ? matchedSite.orderedBy : "",
      constructionSite: matchedSite ? (matchedSite.siteZh || matchedSite.siteEn) : (project || ""),
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
      const prompt = `This is a packing list, delivery memo, shipping list, or similar logistics document for elevator/escalator materials, possibly in English, Traditional or Simplified Chinese, or mixed. It may cover one or more lifts/lots/field modules, each containing individual cases/packages with descriptions, quantities, weights, and CBM or dimensions. Extract a JSON object with EXACTLY this shape and nothing else (no markdown fences, no commentary):
{"client": "best-guess client name or ''", "project": "site/building/project name found in the document, or ''", "groups": [{"lot": "unit/lift/lot/shop-order code that identifies this batch", "containers": ["container numbers if any, else empty array"], "packages": [{"code": "case/package number", "description": "item description", "weightKg": number_or_empty_string, "cbm": number_or_empty_string}]}]}
If CBM isn't given directly but dimensions in mm are (e.g. LxWxH), compute cbm as L*W*H/1000000000. If the document only has one overall lot/shipment with no explicit lift/case breakdown, put everything under a single group with a sensible lot name.`;
      const response = await fetch("/api/scan-pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-5",
          max_tokens: 4000,
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
      const parsed = JSON.parse(clean);
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
      const ok = applyParsedResult({ groups: normalizedGroups, client: parsed.client, project: parsed.project });
      if (!ok) setPdfError(t.packingListNoStructure);
      setPdfStatus("idle");
    } catch (err) {
      setPdfError(t.pdfReadErrorMsg || "Couldn't read this PDF. Check your API key and try again.");
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
          item[key] = value instanceof Date ? value.toISOString().slice(0, 10) : String(value).trim();
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

  function importPackingList() {
    const sharedJobNumber = plCommon.jobNumber || nextJobNumber(existingItems);
    const newItems = plPreview.map((g) => {
      const base = emptyForm();
      return {
        ...base,
        client: plCommon.client,
        project: plCommon.project,
        constructionSite: plCommon.constructionSite || "",
        depot: plCommon.depot,
        depotArrivalDate: plCommon.depotArrivalDate,
        depotLocation: plCommon.depotLocation,
        arrivingType: plCommon.arrivingType,
        jobNumber: sharedJobNumber,
        jobRef: plCommon.jobRef || "",
        orderedBy: plCommon.orderedBy || "",
        directoryId: plCommon.directoryId || "",
        itemType: "Separate Items",
        unitCode: g.lot,
        weightKg: g.totalWeight ? String(Math.round(g.totalWeight * 10) / 10) : "",
        volumeCbm: g.totalCbm ? String(Math.round(g.totalCbm * 1000) / 1000) : "",
        packages: g.packages,
        notes: g.containers.length ? `Container(s): ${g.containers.join(", ")}` : "",
      };
    });
    onImportRows(newItems);
    setPlPreview(null);
    setPlCommon(null);
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="flex gap-1 rounded-lg p-1 w-fit" style={{ background: colors.surfaceDim }}>
        {[["packinglist", t.tabPackingList], ["pdf", t.tabPdf], ["excel", t.tabExcel]].map(([k, label]) => (
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
                          }));
                        }}
                      >
                        <option value="">{t.selectFromDirectoryPlaceholder}</option>
                        {directory.map((s) => <option key={s.id} value={s.id}>{s.siteEn} — {s.client}</option>)}
                      </select>
                    </Field>
                  </div>
                )}
                <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
                  <Field label={t.packingListApplyClient} colors={colors}>
                    <select className={inputClass} style={inputStyle} value={plCommon.client} onChange={(e) => setPlCommon((c) => ({ ...c, client: e.target.value }))}>
                      {CLIENTS.map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                  <Field label={t.packingListApplyProject} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.project} onChange={(e) => setPlCommon((c) => ({ ...c, project: e.target.value }))} />
                  </Field>
                  <Field label={t.fOrderedBy} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.orderedBy || ""} onChange={(e) => setPlCommon((c) => ({ ...c, orderedBy: e.target.value }))} />
                  </Field>
                  <Field label={t.fJobRef} hint={t.fJobRefHint} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.jobRef || ""} onChange={(e) => setPlCommon((c) => ({ ...c, jobRef: e.target.value }))} />
                  </Field>
                  <Field label={t.packingListApplyDepot} colors={colors}>
                    <select className={inputClass} style={inputStyle} value={plCommon.depot} onChange={(e) => setPlCommon((c) => ({ ...c, depot: e.target.value }))}>
                      {DEPOTS.map((d) => <option key={d} value={d}>{depotLabel(d, lang)}</option>)}
                    </select>
                  </Field>
                  <Field label={t.packingListApplyDepotArrival} colors={colors}>
                    <input type="date" className={inputClass} style={inputStyle} value={plCommon.depotArrivalDate} onChange={(e) => setPlCommon((c) => ({ ...c, depotArrivalDate: e.target.value }))} />
                  </Field>
                  <Field label={t.packingListApplyDepotLocation} colors={colors}>
                    <input className={inputClass} style={inputStyle} value={plCommon.depotLocation} onChange={(e) => setPlCommon((c) => ({ ...c, depotLocation: e.target.value }))} />
                  </Field>
                  <Field label={t.fArrivingType} hint={t.fArrivingTypeHint} colors={colors}>
                    <select className={inputClass} style={inputStyle} value={plCommon.arrivingType} onChange={(e) => setPlCommon((c) => ({ ...c, arrivingType: e.target.value }))}>
                      {ARRIVING_TYPES.map((a) => <option key={a}>{a}</option>)}
                    </select>
                  </Field>
                  <Field label={t.fJobNumber} hint={t.fJobNumberHint} colors={colors}>
                    <div className="flex gap-2">
                      <input className={inputClass + " flex-1"} style={inputStyle} value={plCommon.jobNumber} onChange={(e) => setPlCommon((c) => ({ ...c, jobNumber: e.target.value }))} />
                      <button
                        type="button"
                        className="px-2.5 py-1.5 rounded text-xs font-semibold whitespace-nowrap"
                        style={{ background: colors.amber, color: colors.ink, fontFamily: FONT_DISPLAY }}
                        onClick={() => setPlCommon((c) => ({ ...c, jobNumber: nextJobNumber(existingItems) }))}
                      >
                        {t.generateJobNoBtn}
                      </button>
                    </div>
                  </Field>
                </div>
              </div>

              <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-2 text-sm font-semibold" style={{ background: colors.amberSoft, color: colors.amberText, fontFamily: FONT_DISPLAY }}>
                  {t.packingListDetectedTitle(plPreview.length)}
                </div>
                <table className="w-full text-sm" style={{ background: colors.surface }}>
                  <thead>
                    <tr style={{ background: colors.surfaceDim }}>
                      {[t.colLot, t.colPackages, t.colContainers, t.colWeight, t.colCbm].map((h) => (
                        <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {plPreview.map((g, idx) => (
                      <tr key={idx} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                        <td className="px-3 py-2 font-semibold">{g.lot}</td>
                        <td className="px-3 py-2">{g.packages.length}</td>
                        <td className="px-3 py-2 text-xs" style={{ color: colors.inkFaint }}>{g.containers.join(", ") || "—"}</td>
                        <td className="px-3 py-2">{Math.round(g.totalWeight)}</td>
                        <td className="px-3 py-2">{g.totalCbm ? Math.round(g.totalCbm * 1000) / 1000 : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex gap-2">
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ background: colors.navy, color: colors.onDark, fontFamily: FONT_DISPLAY }} onClick={importPackingList}>
                  {t.packingListImportBtn(plPreview.length)}
                </button>
                <button className="px-4 py-2 rounded text-sm font-semibold w-fit" style={{ border: `1px solid ${colors.line}`, color: colors.ink, fontFamily: FONT_DISPLAY }}
                  onClick={() => { setPlPreview(null); setPlCommon(null); }}>
                  {t.discardBtn}
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

export default function FarspeedInventory() {
  const [items, setItems] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState("dashboard");
  const [editing, setEditing] = useState(null);
  const [exitingItem, setExitingItem] = useState(null);
  const [printJobSheet, setPrintJobSheet] = useState(null);
  const [expandedRowId, setExpandedRowId] = useState(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [newEntryMenuOpen, setNewEntryMenuOpen] = useState(false);
  const [directory, setDirectoryState] = useState([]);
  const [employees, setEmployeesState] = useState([]);
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

  function handleAddDelivery(delivery) {
    const record = { ...delivery, id: `D${Date.now()}${Math.floor(Math.random() * 1000)}` };
    persist(items.map((i) => (i.id === exitingItem.id ? { ...i, deliveries: [...(i.deliveries || []), record] } : i)));
    setExitingItem((prev) => (prev ? { ...prev, deliveries: [...(prev.deliveries || []), record] } : prev));
  }

  function handleDeleteDelivery(deliveryId) {
    persist(items.map((i) => (i.id === exitingItem.id ? { ...i, deliveries: (i.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: true } : d)) } : i)));
    setExitingItem((prev) => (prev ? { ...prev, deliveries: (prev.deliveries || []).map((d) => (d.id === deliveryId ? { ...d, cancelled: true } : d)) } : prev));
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
          i.constructionSite?.toLowerCase().includes(q) || i.invoiceNumber?.toLowerCase().includes(q)
        );
      })
      .sort((a, b) => b.numericId - a.numericId);
  }, [activeItemsList, filterClient, filterStatus, filterDepot, search]);

  const atDepot = activeItemsList.filter((i) => deriveStatus(i) === "at_depot");
  const partial = activeItemsList.filter((i) => deriveStatus(i) === "partial");
  const openForDelivery = [...atDepot, ...partial];
  const pending = activeItemsList.filter((i) => deriveStatus(i) === "pending_collection");
  const billable = openForDelivery.filter((i) => storageInfo(i)?.billable);
  const lfdWarnings = activeItemsList.filter((i) => { const a = lfdAlert(i); return a && (a.level === "soon" || a.level === "overdue"); });
  const duplicateGroups = useMemo(() => findDuplicateGroups(activeItemsList), [activeItemsList]);
  const duplicateIds = useMemo(() => new Set(duplicateGroups.flat().map((i) => i.id)), [duplicateGroups]);
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

  if (!loaded) {
    return <div className="p-8 text-sm" style={{ color: colors.inkFaint }}>{t.loadingMsg}</div>;
  }

  return (
    <div style={{ background: colors.bg, minHeight: "100%", fontFamily: FONT_BODY }} className="w-full">
      <style>{FONT_IMPORT}</style>

      <div style={{ background: colors.navy }} className="px-6 py-4 flex items-center justify-between flex-wrap gap-3">
        <div>
          <div className="text-2xl font-bold tracking-tight" style={{ color: colors.onDark, fontFamily: FONT_DISPLAY }}>FARSPEED CONTRACTORS LTD</div>
          <div className="text-xs tracking-widest uppercase" style={{ color: colors.amber }}>{t.appSubtitle}</div>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          <div className="flex gap-1 rounded-lg p-1" style={{ background: colors.navySoft }}>
            {[
              ["dashboard", t.navDashboard],
              ["inventory", t.navInventory],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setEditing(null); setExitingItem(null); setNewEntryMenuOpen(false); setSettingsOpen(false); setView(k); }}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: view === k ? colors.amber : "transparent", color: view === k ? colors.ink : colors.onDark }}>
                {label}
              </button>
            ))}

            <div className="relative">
              <button
                onClick={() => { setNewEntryMenuOpen((o) => !o); setSettingsOpen(false); }}
                className="px-3 py-1.5 rounded text-sm font-semibold"
                style={{ fontFamily: FONT_DISPLAY, background: ["add", "import"].includes(view) ? colors.amber : "transparent", color: ["add", "import"].includes(view) ? colors.ink : colors.onDark }}
              >
                {t.navNewEntry} ▾
              </button>
              {newEntryMenuOpen && (
                <div className="absolute left-0 mt-1 rounded-lg overflow-hidden z-20" style={{ background: colors.surface, border: `1px solid ${colors.line}`, minWidth: 140 }}>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY }}
                    onClick={() => { setEditing(null); setView("add"); setNewEntryMenuOpen(false); }}
                  >
                    {t.newEntryManual}
                  </button>
                  <button
                    className="block w-full text-left px-3 py-2 text-sm font-semibold"
                    style={{ color: colors.ink, fontFamily: FONT_DISPLAY, borderTop: `1px solid ${colors.surfaceDim}` }}
                    onClick={() => { setView("import"); setNewEntryMenuOpen(false); }}
                  >
                    {t.newEntryImport}
                  </button>
                </div>
              )}
            </div>

            {[
              ["exit", t.navDeliveries],
              ["directory", t.navDirectory],
              ["joblog", t.navJobLog],
            ].map(([k, label]) => (
              <button key={k} onClick={() => { setEditing(null); setExitingItem(null); setNewEntryMenuOpen(false); setSettingsOpen(false); setView(k); }}
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
          <div className="flex items-center gap-1 rounded-lg px-2 py-1" style={{ background: colors.navySoft }}>
            <span className="text-xs" style={{ color: colors.onDark, opacity: 0.7, fontFamily: FONT_DISPLAY }}>{t.signedInAs}:</span>
            <select
              className="text-sm font-semibold bg-transparent"
              style={{ color: colors.amber, fontFamily: FONT_DISPLAY }}
              value={currentUser}
              onChange={(e) => setCurrentUser(e.target.value)}
            >
              <option value="" style={{ color: "#000" }}>{t.signedInNone}</option>
              {employees.map((e) => <option key={e.id} value={e.name} style={{ color: "#000" }}>{e.name}</option>)}
            </select>
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

      <div className="p-6 max-w-6xl mx-auto">
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

            <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
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
                        <td className="px-3 py-2 text-right whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                          {["at_depot", "partial"].includes(deriveStatus(i)) && (
                            <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => { setExitingItem(i); setView("exit"); }}>{t.deliverBtn}</button>
                          )}
                          <button className="text-xs font-semibold mr-3" style={{ color: colors.amberText }} onClick={() => { setEditing(i); setView("add"); }}>{t.editBtn}</button>
                          {i.jobNumber && (
                            <button className="text-xs font-semibold mr-3" style={{ color: colors.inkFaint }} onClick={() => handleCancelItem(i.id)}>{t.cancelJobBtn}</button>
                          )}
                          <button className="text-xs font-semibold" style={{ color: colors.red }} onClick={() => handleDelete(i.id)}>{t.deleteBtn}</button>
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
                                        {remaining.map((p) => (
                                          <span key={p.code} className="px-2 py-1 rounded text-xs" style={{ background: colors.surface, border: `1px solid ${colors.line}`, color: colors.ink }}>
                                            {p.code}{p.description ? ` — ${p.description}` : ""}
                                          </span>
                                        ))}
                                      </div>
                                    ) : (
                                      <span className="text-xs" style={{ color: colors.inkFaint }}>{t.inventoryNoRemainingPkgsMsg}</span>
                                    )
                                  )}
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
            {exitingItem ? (
              <DeliveryForm item={exitingItem} onAddDelivery={handleAddDelivery} onDeleteDelivery={handleDeleteDelivery}
                onCancel={() => { setExitingItem(null); setView("inventory"); }} onPrintJobSheet={setPrintJobSheet}
                employees={employees} currentUser={currentUser} items={items} colors={colors} t={t} lang={lang} />
            ) : (
              <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
                <div className="px-4 py-3 text-sm" style={{ background: colors.surfaceDim, color: colors.inkFaint }}>{t.selectItemMsg}</div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm" style={{ background: colors.surface }}>
                    <thead>
                      <tr style={{ background: colors.surfaceDim }}>
                        {[t.colId, t.colClient, t.colProjectSite, t.colDepot, t.colDepotArrival, t.colStatus, ""].map((h) => (
                          <th key={h} className="text-left px-3 py-2 text-xs font-semibold uppercase tracking-wider" style={{ color: colors.inkFaint, fontFamily: FONT_DISPLAY }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {openForDelivery.length === 0 && (
                        <tr><td colSpan={7} className="px-3 py-6 text-center text-sm" style={{ color: colors.inkFaint }}>{t.nothingAtDepotMsg}</td></tr>
                      )}
                      {openForDelivery.map((i) => (
                        <tr key={i.id} style={{ borderTop: `1px solid ${colors.surfaceDim}`, color: colors.ink }}>
                          <td className="px-3 py-2" style={{ fontFamily: FONT_MONO }}>{i.id}</td>
                          <td className="px-3 py-2">{i.client}</td>
                          <td className="px-3 py-2 max-w-[220px] truncate">{i.project}</td>
                          <td className="px-3 py-2 whitespace-nowrap">{depotDisplay(i.depot, lang)}</td>
                          <td className="px-3 py-2">{fmt(i.depotArrivalDate)}</td>
                          <td className="px-3 py-2"><StatusBadge item={i} colors={colors} t={t} /></td>
                          <td className="px-3 py-2 text-right">
                            <button className="text-xs font-semibold px-2 py-1 rounded" style={{ background: colors.amber, color: colors.ink }} onClick={() => setExitingItem(i)}>
                              {t.recordDeliveryBtn}
                            </button>
                          </td>
                        </tr>
                      ))}
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

        {view === "import" && (
          <ImportPanel onImportRows={handleImportRows} existingItems={items} directory={directory} colors={colors} t={t} lang={lang} />
        )}

        {view === "directory" && (
          <DirectoryPanel directory={directory} setDirectory={setDirectory} employees={employees} setEmployees={setEmployees} colors={colors} t={t} />
        )}

        {view === "joblog" && (
          <div className="flex flex-col gap-4">
            <div className="rounded-lg p-5" style={{ background: colors.surface, border: `1px solid ${colors.line}` }}>
              <h3 className="text-lg font-bold mb-1" style={{ fontFamily: FONT_DISPLAY, color: colors.ink }}>{t.jobLogTitle}</h3>
              <p className="text-sm" style={{ color: colors.inkFaint }}>{t.jobLogDesc}</p>
            </div>
            <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
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
            <div className="rounded-lg overflow-hidden" style={{ border: `1px solid ${colors.line}` }}>
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
        <JobSheetPrint sheet={printJobSheet} onClose={() => setPrintJobSheet(null)} colors={colors} t={t} lang={lang} />
      )}
    </div>
  );
}

