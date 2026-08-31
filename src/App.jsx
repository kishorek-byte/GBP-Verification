import React, { useState, useEffect, useRef, useMemo } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, UploadCloud, Database, AlertTriangle, Users, TrendingUp,
  ScrollText, Settings as SettingsIcon, HelpCircle, Search, Download, RefreshCw,
  CheckCircle2, Clock, XCircle, CircleHelp, ChevronRight, X, Play, FileText,
  Sun, ArrowLeft, Info, Sparkles, ChevronDown, PlusCircle, Menu, Copy, Check
} from "lucide-react";

/* ============================================================================
   GBP CHANGE VERIFICATION & MONITORING
   ----------------------------------------------------------------------------
   A verification/monitoring layer on top of the existing SI platform.
   This tool does NOT sync data. It compares an "expected change" against the
   live GBP snapshot and classifies it GREEN / YELLOW / RED / GRAY.
   ============================================================================ */

/* ---------------------------- CONSTANTS ---------------------------------- */

const CHANGE_TYPES = [
  "Phone", "Business Name", "Address", "Business Hours", "Map Pin (Lat/Long)",
  "Category", "Website", "WhatsApp Number", "Services", "Products", "Photos",
  "Service Area", "Description",
];

const FIELD_KEY = {
  "Phone": "phone",
  "Business Name": "businessName",
  "Address": "address",
  "Business Hours": "hours",
  "Map Pin (Lat/Long)": "latlng", // virtual — computed from lat/lng, not a plain field lookup
  "Category": "category",
  "Website": "website",
  "WhatsApp Number": "whatsapp",
  "Services": "services",
  "Products": "products",
  "Photos": "photos",
  "Service Area": "serviceArea",
  "Description": "description",
};

const STATUS_META = {
  GREEN:  { label: "VERIFIED",           dot: "#16A34A", text: "#15803D", bg: "#ECFDF3", border: "#B7EFC5", icon: CheckCircle2 },
  YELLOW: { label: "PENDING",            dot: "#D97706", text: "#B45309", bg: "#FFFBEB", border: "#FDE9B8", icon: Clock },
  RED:    { label: "EXCEPTION",          dot: "#DC2626", text: "#B91C1C", bg: "#FEF2F2", border: "#F9C7C4", icon: XCircle },
  GRAY:   { label: "MATCHING REQUIRED",  dot: "#6B7280", text: "#4B5563", bg: "#F3F4F6", border: "#E2E4E8", icon: CircleHelp },
};

// Real-world default recheck schedule (used for "Next Check" display / Settings)
const DEFAULT_INTERVALS_MIN = [30, 120, 360, 60 * 18]; // 30m, 2h, 6h, ~next business day
const INTERVAL_LABELS = ["30 minutes", "2 hours", "6 hours", "Next business day"];

// Compressed timing so the recheck pipeline can be *watched* in a browser demo.
// This does not run once the tab is closed — see Help / Settings for the honest caveat.
const DEMO_SECONDS = [10, 18, 26, 34];

const PRIORITIES = ["Low", "Normal", "High", "Urgent"];

/* ---------------------------- UTILITIES ----------------------------------- */

let __id = 1;
const nextId = (p) => `${p}_${__id++}_${Math.random().toString(36).slice(2, 7)}`;

function fmtDateTime(d) {
  if (!d) return "-";
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt.getTime())) return "-";
  return dt.toLocaleString("en-IN", { day: "2-digit", month: "short", hour: "numeric", minute: "2-digit", hour12: true });
}
function fmtDateTimeFull(d) {
  if (!d) return "-";
  const dt = d instanceof Date ? d : new Date(d);
  if (isNaN(dt.getTime())) return "-";
  return dt.toLocaleString("en-IN", { day: "2-digit", month: "long", year: "numeric", hour: "numeric", minute: "2-digit", hour12: true });
}
function addMinutes(d, m) { return new Date(d.getTime() + m * 60000); }
function minutesBetween(a, b) { return Math.round((b.getTime() - a.getTime()) / 60000); }
function fmtDuration(mins) {
  if (mins == null || isNaN(mins)) return "-";
  if (mins < 60) return `${mins}m`;
  const h = Math.floor(mins / 60), m = mins % 60;
  return m ? `${h}h ${m}m` : `${h}h`;
}

/* --------------------------- NORMALIZATION -------------------------------- */

function normalizePhone(v) {
  if (!v) return "";
  let digits = String(v).replace(/[^\d]/g, "");
  if (digits.length === 12 && digits.startsWith("91")) digits = digits.slice(2);
  if (digits.length === 11 && digits.startsWith("0")) digits = digits.slice(1);
  return digits;
}
function normalizeName(v) {
  if (!v) return "";
  return String(v).trim().replace(/\s+/g, " ").toLowerCase();
}
function normalizeAddress(v) {
  if (!v) return "";
  return String(v)
    .trim().toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[.,]/g, "")
    .replace(/\bstreet\b/g, "st")
    .replace(/\broad\b/g, "rd")
    .replace(/\bfloor\b/g, "fl");
}
function normalizeHours(v) {
  if (!v) return "";
  return String(v)
    .trim().toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[–—]/g, "-")
    .replace(/\./g, "")
    .replace(/\bam\b/g, "am").replace(/\bpm\b/g, "pm");
}
// "lat,lng" strings — trims whitespace only; equality isn't meaningful for
// coordinates (see haversineMeters / valuesMatch), this is just for display.
function normalizeLatLng(v) {
  if (!v) return "";
  return String(v).trim().replace(/\s+/g, "");
}
function normalizeUrl(v) {
  if (!v) return "";
  return String(v).trim().toLowerCase()
    .replace(/^https?:\/\//, "")
    .replace(/^www\./, "")
    .replace(/\/+$/, "");
}
// Services / Products / Service Area are usually multi-value — compare as an
// unordered set so re-ordering the same list doesn't read as a mismatch.
function normalizeListValue(v) {
  if (!v) return "";
  return String(v).split(/[,;|/]/).map(s => s.trim().toLowerCase()).filter(Boolean).sort().join("|");
}
function normalizeLongText(v) {
  if (!v) return "";
  return String(v).trim().toLowerCase().replace(/\s+/g, " ");
}
const NORMALIZER = {
  "Phone": normalizePhone,
  "Business Name": normalizeName,
  "Address": normalizeAddress,
  "Business Hours": normalizeHours,
  "Map Pin (Lat/Long)": normalizeLatLng,
  "Category": normalizeName,
  "Website": normalizeUrl,
  "WhatsApp Number": normalizePhone,
  "Services": normalizeListValue,
  "Products": normalizeListValue,
  "Photos": normalizeName,
  "Service Area": normalizeListValue,
  "Description": normalizeLongText,
};

function manualCheckHint(changeType) {
  const map = {
    "Map Pin (Lat/Long)": "Open the listing on Google Maps / Street View and check whether the pin now sits at the right storefront.",
    "Photos": "Open the listing's Photos tab on Google and check the photo(s) directly — images can't be compared from a data file.",
    "Services": "Open the listing's Services section on Google and compare it against what was requested.",
    "Products": "Open the listing's Products section on Google and compare it against what was requested.",
    "Service Area": "Open the listing's Service Area section on Google and compare it against what was requested.",
    "Description": "Open the listing's About/Description section on Google and compare it against what was requested.",
  };
  return map[changeType] || `No "${changeType}" data was included in your GBP Data upload for this location — open the listing on Google directly and check this field.`;
}

function parseLatLng(v) {
  if (!v) return null;
  const parts = String(v).split(",").map(s => parseFloat(s.trim()));
  if (parts.length !== 2 || parts.some(n => Number.isNaN(n))) return null;
  return { lat: parts[0], lng: parts[1] };
}

// Great-circle distance in meters between two lat/lng points.
function haversineMeters(a, b) {
  if (!a || !b) return Infinity;
  const R = 6371000;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(b.lat - a.lat), dLng = toRad(b.lng - a.lng);
  const s = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(s));
}

// Whether the current GBP value satisfies the expected value for a given
// change type. Coordinates need a distance tolerance, not string equality —
// GPS/geocoding precision means two "correct" pins are rarely byte-identical.
function valuesMatch(changeType, currentVal, expectedVal, settings) {
  if (changeType === "Map Pin (Lat/Long)") {
    const a = parseLatLng(currentVal), b = parseLatLng(expectedVal);
    if (!a || !b) return false;
    const toleranceM = (settings && settings.coordToleranceMeters) || 50;
    return haversineMeters(a, b) <= toleranceM;
  }
  const normalizer = NORMALIZER[changeType];
  const nc = normalizer(currentVal), ne = normalizer(expectedVal);
  return nc !== "" && nc === ne;
}

// Reads the "current" value for a change type off a GBP location record.
// Lat/Long isn't a plain field — it's assembled from separate lat/lng columns.
function getCurrentValueForLocation(changeType, location) {
  if (!location) return null;
  if (changeType === "Map Pin (Lat/Long)") {
    const lat = location.lat, lng = location.lng;
    if (lat === undefined || lat === "" || lng === undefined || lng === "") return null;
    return `${lat},${lng}`;
  }
  return location[FIELD_KEY[changeType]];
}

// simple bigram Dice-coefficient similarity, used only for fuzzy name matching
function similarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const bigrams = (s) => {
    const out = [];
    for (let i = 0; i < s.length - 1; i++) out.push(s.slice(i, i + 2));
    return out;
  };
  const A = bigrams(a), B = bigrams(b);
  if (!A.length || !B.length) return 0;
  const map = new Map();
  A.forEach((g) => map.set(g, (map.get(g) || 0) + 1));
  let hits = 0;
  B.forEach((g) => {
    const c = map.get(g) || 0;
    if (c > 0) { hits++; map.set(g, c - 1); }
  });
  return (2 * hits) / (A.length + B.length);
}

/* --------------------------- MATCHING ENGINE ------------------------------- */

function matchLocation(change, gbpLocations) {
  // If the change specifies a client and any uploaded GBP rows are tagged with
  // a client, narrow the candidate pool to that client first. This prevents a
  // "Mumbai" branch belonging to Client A from ever being matched against a
  // change request meant for Client B's "Mumbai" branch.
  const taggedPool = gbpLocations.some(g => g.clientName);
  const pool = (change.clientName && taggedPool)
    ? gbpLocations.filter(g => !g.clientName || normalizeName(g.clientName) === normalizeName(change.clientName))
    : gbpLocations;
  const searchIn = pool.length ? pool : gbpLocations;

  if (change.gbpLocationId) {
    const m = searchIn.find(g => g.gbpLocationId && g.gbpLocationId.trim().toLowerCase() === change.gbpLocationId.trim().toLowerCase());
    if (m) return { location: m, method: "GBP Location ID", confidence: "high" };
  }
  if (change.storeCode) {
    const m = searchIn.find(g => g.storeCode && g.storeCode.trim().toLowerCase() === change.storeCode.trim().toLowerCase());
    if (m) return { location: m, method: "Store Code", confidence: "high" };
  }
  if (change.locationId) {
    const m = searchIn.find(g => g.internalLocationId && g.internalLocationId.trim().toLowerCase() === change.locationId.trim().toLowerCase());
    if (m) return { location: m, method: "Internal Location ID", confidence: "high" };
  }
  if (change.changeType === "Phone" && change.oldValue) {
    const norm = normalizePhone(change.oldValue);
    if (norm) {
      const m = searchIn.find(g => normalizePhone(g.phone) === norm);
      if (m) return { location: m, method: "Phone", confidence: "high" };
    }
  }
  if (change.branchName) {
    let best = null, bestScore = 0;
    searchIn.forEach(g => {
      const nameScore = similarity(normalizeName(change.branchName), normalizeName(g.businessName));
      const addrScore = change.address ? similarity(normalizeAddress(change.address), normalizeAddress(g.address)) : 0;
      const combined = change.address ? nameScore * 0.6 + addrScore * 0.4 : nameScore;
      if (combined > bestScore) { bestScore = combined; best = g; }
    });
    if (best) {
      if (bestScore >= 0.9) return { location: best, method: "Business Name + Address", confidence: "high" };
      if (bestScore >= 0.6) return { location: best, method: "Business Name + Address", confidence: "low" };
    }
  }
  return { location: null, method: "No Match Found", confidence: "none" };
}

/* --------------------------- GBP DATA PROVIDER -----------------------------
   Abstraction layer. MVP implementation reads from an in-memory array that was
   populated by file upload. A future GoogleBusinessProfileAPIProvider could
   implement the same three methods against the live GBP API without any other
   part of the application changing.
------------------------------------------------------------------------------ */

// Every location gets a stable internal _uid so lookups never collide, even
// when GBP Location ID is blank for multiple rows (common in real exports —
// this was the root cause of a real bug where two ID-less listings' checks
// got mixed up). Matching logic still uses GBP Location ID / Store Code as
// before; _uid is purely an internal bookkeeping key, never shown to the user.
function keyOfLocation(l) {
  if (l.gbpLocationId && String(l.gbpLocationId).trim()) return String(l.gbpLocationId).trim().toLowerCase();
  if (l.storeCode && String(l.storeCode).trim()) {
    // Store codes are sometimes reused/duplicated in real exports across
    // genuinely different listings. Folding the business name in here means
    // two different businesses sharing a store code become two distinct
    // internal records instead of silently merging into one garbled entry —
    // the same store code + same name still correctly re-merges on re-upload.
    const namePart = l.businessName ? `|${normalizeName(l.businessName)}` : "";
    return `sc:${String(l.storeCode).trim().toLowerCase()}${namePart}`;
  }
  return null;
}
function ensureUid(loc) {
  if (loc._uid) return loc;
  return { ...loc, _uid: keyOfLocation(loc) || nextId("loc") };
}
// Merges freshly uploaded/pasted rows into the existing location pool,
// matching on GBP Location ID or Store Code where available so re-uploads
// update the same record instead of duplicating it; anything with neither
// gets its own fresh internal identity rather than colliding with others.
function mergeGbpLocations(prev, incoming) {
  const byKey = new Map();
  prev.forEach(l => { const loc = ensureUid(l); byKey.set(keyOfLocation(loc) || loc._uid, loc); });
  incoming.forEach(m => {
    const k = keyOfLocation(m);
    if (k && byKey.has(k)) {
      const existing = byKey.get(k);
      byKey.set(k, ensureUid({ ...existing, ...m, _uid: existing._uid }));
    } else {
      const fresh = ensureUid(m);
      byKey.set(k || fresh._uid, fresh);
    }
  });
  return Array.from(byKey.values());
}

class FileUploadGBPProvider {
  constructor(locations) { this.locations = locations || []; }
  getLocations() { return this.locations; }
  getLocation(uid) { return this.locations.find(l => l._uid === uid) || null; }
  getLocationField(uid, field) {
    const loc = this.getLocation(uid);
    return loc ? loc[field] : undefined;
  }
  get sourceLabel() { return "GBP Data Provider (File Upload)"; }
}

/* ------------------------------ DEMO DATA ---------------------------------- */

const CITIES = [
  "Kochi","Pune","Mumbai","Chennai","Bangalore","Hyderabad","Delhi","Kolkata",
  "Ahmedabad","Jaipur","Lucknow","Indore","Nagpur","Surat","Coimbatore",
  "Vadodara","Bhopal","Patna","Chandigarh","Nashik",
];
const CITY_CODE = {
  Kochi:"KOC",Pune:"PUN",Mumbai:"MUM",Chennai:"CHE",Bangalore:"BLR",Hyderabad:"HYD",
  Delhi:"DEL",Kolkata:"KOL",Ahmedabad:"AMD",Jaipur:"JAI",Lucknow:"LKO",Indore:"IND",
  Nagpur:"NAG",Surat:"SUR",Coimbatore:"CBE",Vadodara:"BAR",Bhopal:"BHO",Patna:"PAT",
  Chandigarh:"CHD",Nashik:"NAS",
};

function buildDemoGbpLocations() {
  return CITIES.map((city, i) => {
    const code = CITY_CODE[city];
    return {
      clientName: "ICL Fincorp",
      gbpLocationId: `GBP-ICL-${code}`,
      storeCode: `ICL-KL-${code}-00${(i % 9) + 1}`,
      internalLocationId: `ICL-${code}-001`,
      businessName: `ICL Fincorp - ${city}`,
      phone: `98765${(43200 + i).toString()}`,
      address: `${100 + i}, MG Road, ${city}`,
      hours: "Mon-Sat 9:30 AM-6:30 PM, Sun Closed",
      lat: 10 + i * 0.4,
      lng: 76 + i * 0.3,
    };
  });
}

// hand-authored seed records so the demo narrates the exact scenarios in the brief
function buildDemoRecords(gbpLocations) {
  const g = (code) => gbpLocations.find(l => l.gbpLocationId === `GBP-ICL-${code}`);
  const today = new Date();
  today.setHours(10, 5, 0, 0);
  const mk = (overrides) => {
    const base = {
      id: nextId("rec"),
      client: "ICL Fincorp",
      priority: "Normal",
      requestedBy: "R. Menon",
      unexpectedChange: null,
      isDemo: true,
      source: "GBP Data Provider (File Upload)",
    };
    return { ...base, ...overrides };
  };

  const records = [];

  // 1. Kochi phone — GREEN, matches the worked example in the brief exactly
  {
    const loc = g("KOC");
    const created = new Date(today); created.setHours(10, 5, 0, 0);
    const c1 = new Date(today); c1.setHours(10, 35, 0, 0);
    const c2 = new Date(today); c2.setHours(12, 5, 0, 0);
    const c3 = new Date(today); c3.setHours(15, 5, 0, 0);
    loc.phone = "9876543210";
    records.push(mk({
      client: "ICL Fincorp", branch: "Kochi", locationId: "ICL-KL-KOC-001", storeCode: "ICL-KL-KOC-001",
      changeType: "Phone", oldValue: "9876543200", expectedValue: "9876543210",
      matchedLocationId: loc.gbpLocationId, matchMethod: "Store Code", matchConfidence: "high",
      currentGbpValue: loc.phone, attempts: 3, status: "GREEN",
      createdAt: created, verifiedAt: c3, nextCheckDemoAt: null, requestDate: "2026-07-31",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: "9876543200", result: "PENDING" },
        { n: 2, time: c2, gbpValue: "9876543200", result: "PENDING" },
        { n: 3, time: c3, gbpValue: "9876543210", result: "VERIFIED" },
      ],
    }));
  }

  // 2. Pune phone — RED, matches the worked example exactly
  {
    const loc = g("PUN");
    loc.phone = "9876543200"; // still old — this is why it's an exception
    const created = new Date(today); created.setHours(10, 30, 0, 0);
    const c1 = new Date(today); c1.setHours(11, 0, 0, 0);
    const c2 = new Date(today); c2.setHours(13, 0, 0, 0);
    const c3 = new Date(today); c3.setHours(19, 0, 0, 0);
    const c4 = new Date(today); c4.setHours(15, 30, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Pune", locationId: "ICL-KL-PUN-004", storeCode: "ICL-KL-PUN-004",
      changeType: "Phone", oldValue: "9876543200", expectedValue: "9876543222",
      matchedLocationId: loc.gbpLocationId, matchMethod: "Store Code", matchConfidence: "high",
      currentGbpValue: loc.phone, attempts: 4, status: "RED",
      createdAt: created, requestDate: "2026-07-31",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: "9876543200", result: "PENDING" },
        { n: 2, time: c2, gbpValue: "9876543200", result: "PENDING" },
        { n: 3, time: c3, gbpValue: "9876543200", result: "PENDING" },
        { n: 4, time: c4, gbpValue: "9876543200", result: "EXCEPTION" },
      ],
      reviewed: false, notes: [],
    }));
  }

  // 3. Mumbai business hours — YELLOW, live-pending, will tick in the demo clock
  {
    const loc = g("MUM");
    loc.hours = "Mon-Sat 9:30 AM-6:30 PM, Sun Closed"; // unchanged so far
    const created = new Date(); created.setMinutes(created.getMinutes() - 12);
    const c1 = new Date(); c1.setMinutes(c1.getMinutes() - 2);
    records.push(mk({
      client: "ICL Fincorp", branch: "Mumbai", locationId: "ICL-KL-MUM-002", storeCode: "ICL-KL-MUM-002",
      changeType: "Business Hours", oldValue: "Mon-Sat 9:30 AM-6:30 PM, Sun Closed",
      expectedValue: "Mon-Sat 9:00 AM-7:00 PM, Sun 10:00 AM-2:00 PM",
      matchedLocationId: loc.gbpLocationId, matchMethod: "GBP Location ID", matchConfidence: "high",
      currentGbpValue: loc.hours, attempts: 1, status: "YELLOW",
      createdAt: created, requestDate: "2026-08-03",
      nextCheckDemoAt: Date.now() + DEMO_SECONDS[1] * 1000,
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: loc.hours, result: "PENDING" },
      ],
      reviewed: false, notes: [],
    }));
  }

  // 4. Chennai address — GRAY, low-confidence match (no store code / gbp id given)
  {
    const created = new Date(); created.setHours(9, 50, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Chennai Annex", locationId: "", storeCode: "",
      changeType: "Address", oldValue: "12, Anna Salai, Chennai",
      expectedValue: "14, Anna Salai, Chennai",
      matchedLocationId: null, matchMethod: "No Match Found", matchConfidence: "none",
      currentGbpValue: null, attempts: 0, status: "GRAY",
      createdAt: created, requestDate: "2026-08-01",
      checks: [{ n: 0, time: created, event: "Change request created — no confident location match" }],
      reviewed: false, notes: [],
    }));
  }

  // 5. Bangalore phone — GREEN but with an unexpected business-name change flagged
  {
    const loc = g("BLR");
    loc.phone = "9876543214";
    loc.businessName = "ICL Fincorp Pvt Ltd - Bangalore Whitefield"; // drifted from SI's "Bangalore"
    const created = new Date(today); created.setHours(9, 0, 0, 0);
    const c1 = new Date(today); c1.setHours(9, 40, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Bangalore", locationId: "ICL-KL-BLR-005", storeCode: "ICL-KL-BLR-005",
      changeType: "Phone", oldValue: "9876543204", expectedValue: "9876543214",
      matchedLocationId: loc.gbpLocationId, matchMethod: "Store Code", matchConfidence: "high",
      currentGbpValue: loc.phone, attempts: 1, status: "GREEN",
      createdAt: created, verifiedAt: c1, requestDate: "2026-07-30",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: loc.phone, result: "VERIFIED" },
      ],
      unexpectedChange: { field: "Business Name", from: "ICL Fincorp - Bangalore", to: loc.businessName },
      reviewed: false, notes: [],
    }));
  }

  // 6. Hyderabad business name — GREEN
  {
    const loc = g("HYD");
    loc.businessName = "ICL Fincorp Limited - Hyderabad";
    const created = new Date(today); created.setHours(8, 30, 0, 0);
    const c1 = new Date(today); c1.setHours(9, 5, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Hyderabad", locationId: "ICL-KL-HYD-006", storeCode: "ICL-KL-HYD-006",
      changeType: "Business Name", oldValue: "ICL Fincorp - Hyderabad",
      expectedValue: "ICL Fincorp Limited - Hyderabad",
      matchedLocationId: loc.gbpLocationId, matchMethod: "GBP Location ID", matchConfidence: "high",
      currentGbpValue: loc.businessName, attempts: 1, status: "GREEN",
      createdAt: created, verifiedAt: c1, requestDate: "2026-07-29",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: loc.businessName, result: "VERIFIED" },
      ],
      reviewed: false, notes: [],
    }));
  }

  // 7. Kolkata address — RED
  {
    const loc = g("KOL");
    const created = new Date(today); created.setHours(9, 15, 0, 0);
    const c1 = new Date(today); c1.setHours(9, 45, 0, 0);
    const c2 = new Date(today); c2.setHours(11, 45, 0, 0);
    const c3 = new Date(today); c3.setHours(17, 45, 0, 0);
    const c4 = new Date(today); c4.setHours(14, 0, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Kolkata", locationId: "ICL-KL-KOL-007", storeCode: "ICL-KL-KOL-007",
      changeType: "Address", oldValue: "107, MG Road, Kolkata", expectedValue: "45, Park Street, Kolkata",
      matchedLocationId: loc.gbpLocationId, matchMethod: "Store Code", matchConfidence: "high",
      currentGbpValue: loc.address, attempts: 4, status: "RED",
      createdAt: created, requestDate: "2026-07-28",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: loc.address, result: "PENDING" },
        { n: 2, time: c2, gbpValue: loc.address, result: "PENDING" },
        { n: 3, time: c3, gbpValue: loc.address, result: "PENDING" },
        { n: 4, time: c4, gbpValue: loc.address, result: "EXCEPTION" },
      ],
      reviewed: false, notes: [],
    }));
  }

  // 8. Ahmedabad phone — YELLOW, fewer attempts so far, live-ticking
  {
    const loc = g("AMD");
    const created = new Date(); created.setMinutes(created.getMinutes() - 4);
    records.push(mk({
      client: "ICL Fincorp", branch: "Ahmedabad", locationId: "ICL-KL-AMD-008", storeCode: "ICL-KL-AMD-008",
      changeType: "Phone", oldValue: "9876543208", expectedValue: "9876543299",
      matchedLocationId: loc.gbpLocationId, matchMethod: "Store Code", matchConfidence: "high",
      currentGbpValue: loc.phone, attempts: 0, status: "YELLOW",
      createdAt: created, requestDate: "2026-08-03",
      nextCheckDemoAt: Date.now() + DEMO_SECONDS[0] * 1000,
      checks: [{ n: 0, time: created, event: "Change request created" }],
      reviewed: false, notes: [],
    }));
  }

  // 9. Delhi address — GREEN
  {
    const loc = g("DEL");
    loc.address = "22, Connaught Place, Delhi";
    const created = new Date(today); created.setHours(8, 0, 0, 0);
    const c1 = new Date(today); c1.setHours(8, 35, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Delhi", locationId: "ICL-KL-DEL-009", storeCode: "ICL-KL-DEL-009",
      changeType: "Address", oldValue: "19, Connaught Place, Delhi", expectedValue: "22, Connaught Place, Delhi",
      matchedLocationId: loc.gbpLocationId, matchMethod: "GBP Location ID", matchConfidence: "high",
      currentGbpValue: loc.address, attempts: 1, status: "GREEN",
      createdAt: created, verifiedAt: c1, requestDate: "2026-07-27",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: loc.address, result: "VERIFIED" },
      ],
      reviewed: false, notes: [],
    }));
  }

  // 10. Jaipur hours — RED, already reviewed (for exception-management demo)
  {
    const loc = g("JAI");
    const created = new Date(today); created.setHours(7, 0, 0, 0);
    const c1 = new Date(today); c1.setHours(7, 30, 0, 0);
    const c2 = new Date(today); c2.setHours(9, 30, 0, 0);
    const c3 = new Date(today); c3.setHours(13, 30, 0, 0);
    const c4 = new Date(today); c4.setHours(9, 0, 0, 0);
    records.push(mk({
      client: "ICL Fincorp", branch: "Jaipur", locationId: "ICL-KL-JAI-010", storeCode: "ICL-KL-JAI-010",
      changeType: "Business Hours", oldValue: "Mon-Sat 9:30 AM-6:30 PM, Sun Closed",
      expectedValue: "Mon-Sun 9:00 AM-9:00 PM",
      matchedLocationId: loc.gbpLocationId, matchMethod: "Store Code", matchConfidence: "high",
      currentGbpValue: loc.hours, attempts: 4, status: "RED",
      createdAt: created, requestDate: "2026-07-26",
      checks: [
        { n: 0, time: created, event: "Change request created" },
        { n: 1, time: c1, gbpValue: loc.hours, result: "PENDING" },
        { n: 2, time: c2, gbpValue: loc.hours, result: "PENDING" },
        { n: 3, time: c3, gbpValue: loc.hours, result: "PENDING" },
        { n: 4, time: c4, gbpValue: loc.hours, result: "EXCEPTION" },
      ],
      reviewed: true, notes: [{ time: c4, text: "Confirmed with CSM — GBP change was never submitted on Google's side. Re-requesting.", author: "P. Sharma" }],
    }));
  }

  return records;
}

// A handful of *unprocessed* expected-change rows so "RUN VERIFICATION" has
// something real to do live in the demo.
function buildDemoUnprocessedChanges(gbpLocations) {
  const g = (code) => gbpLocations.find(l => l.gbpLocationId === `GBP-ICL-${code}`);
  const rows = [
    { city: "Lucknow", code: "LKO", changeType: "Phone", newSuffix: "543280" },
    { city: "Indore", code: "IND", changeType: "Business Name", newVal: "ICL Fincorp Pvt Ltd - Indore" },
    { city: "Nagpur", code: "NAG", changeType: "Address", newVal: "9, Civil Lines, Nagpur" },
    { city: "Surat", code: "SUR", changeType: "Business Hours", newVal: "Mon-Sat 10:00 AM-7:00 PM, Sun Closed" },
    { city: "Coimbatore", code: "CBE", changeType: "Phone", newSuffix: "543294" },
    { city: "Vadodara", code: "BAR", changeType: "Phone", newSuffix: "543295" },
  ];
  return rows.map((r, i) => {
    const loc = g(r.code);
    const oldValue = loc[FIELD_KEY[r.changeType]];
    const expectedNewValue = r.newVal || `98765${r.newSuffix}`;
    return {
      id: nextId("chg"),
      clientName: "ICL Fincorp",
      locationId: loc.internalLocationId,
      storeCode: loc.storeCode,
      branchName: r.city,
      gbpLocationId: loc.gbpLocationId,
      changeType: r.changeType,
      oldValue,
      expectedNewValue,
      requestDate: "2026-08-03",
      requestedBy: "R. Menon",
      priority: i % 3 === 0 ? "High" : "Normal",
    };
  });
}

/* ------------------------- FILE PARSING HELPERS ---------------------------- */

// Keyword sets used for fuzzy header detection — a header only needs to
// CONTAIN one of these phrases (not match exactly) to be recognized. This is
// deliberately broad so real-world exports (different naming conventions,
// abbreviations, extra words) still import without the sheet being edited.
const EXPECTED_KEYWORDS = {
  clientName: ["client name", "client"],
  locationId: ["internal location id", "internal id", "location id"],
  storeCode: ["store code", "storecode", "outlet code", "branch code"],
  branchName: ["branch name", "branch", "outlet name", "outlet", "listing name", "location name"],
  gbpLocationId: ["gbp location id", "gbp id", "google location id", "google business profile id"],
  changeType: ["change type", "field changed", "attribute", "what changed", "type of change"],
  oldValue: ["old value", "previous value", "current value", "existing value", "old"],
  expectedNewValue: ["expected new value", "expected value", "new value", "updated value", "expected"],
  requestDate: ["request date", "date requested", "date of change", "date"],
  requestedBy: ["requested by", "updated by", "changed by", "done by", "raised by"],
  priority: ["priority"],
};
const GBP_KEYWORDS = {
  clientName: ["client name", "client"],
  gbpLocationId: ["gbp location id", "gbp id", "google location id"],
  storeCode: ["store code", "storecode", "outlet code", "branch code"],
  internalLocationId: ["internal location id", "internal id", "location id"],
  businessName: ["business name", "listing name", "outlet name", "location name", "name"],
  phone: ["phone number", "phone", "mobile number", "mobile", "contact number", "contact no"],
  address: ["address", "full address", "street address"],
  hours: ["business hours", "working hours", "hours", "timing"],
  lat: ["latitude", "lat"],
  lng: ["longitude", "long", "lng"],
  category: ["category", "business category", "primary category"],
  website: ["website", "web address", "url"],
  whatsapp: ["whatsapp number", "whatsapp"],
  services: ["services"],
  products: ["products"],
  photos: ["photos", "photo count", "images"],
  serviceArea: ["service area", "service areas", "coverage area"],
  description: ["description", "business description", "about"],
};

const EXPECTED_FIELDS_META = [
  { key: "clientName", label: "Client Name", required: false },
  { key: "branchName", label: "Branch Name", required: false, identifier: true },
  { key: "storeCode", label: "Store Code", required: false, identifier: true },
  { key: "gbpLocationId", label: "GBP Location ID", required: false, identifier: true },
  { key: "locationId", label: "Internal Location ID", required: false, identifier: true },
  { key: "changeType", label: "Change Type", required: true },
  { key: "oldValue", label: "Old Value", required: true },
  { key: "expectedNewValue", label: "Expected New Value", required: true },
  { key: "requestDate", label: "Request Date", required: false },
  { key: "requestedBy", label: "Requested / Updated By", required: false },
  { key: "priority", label: "Priority", required: false },
];
const GBP_FIELDS_META = [
  { key: "clientName", label: "Client Name", required: false },
  { key: "businessName", label: "Business Name", required: false, identifier: true },
  { key: "storeCode", label: "Store Code", required: false, identifier: true },
  { key: "gbpLocationId", label: "GBP Location ID", required: false, identifier: true },
  { key: "internalLocationId", label: "Internal Location ID", required: false },
  { key: "phone", label: "Phone", required: false },
  { key: "address", label: "Address", required: false },
  { key: "hours", label: "Business Hours", required: false },
  { key: "lat", label: "Latitude", required: false },
  { key: "lng", label: "Longitude", required: false },
  { key: "category", label: "Category", required: false },
  { key: "website", label: "Website", required: false },
  { key: "whatsapp", label: "WhatsApp Number", required: false },
  { key: "services", label: "Services", required: false },
  { key: "products", label: "Products", required: false },
  { key: "photos", label: "Photos", required: false },
  { key: "serviceArea", label: "Service Area", required: false },
  { key: "description", label: "Description", required: false },
];

function normalizeHeader(h) {
  return String(h || "").trim().toLowerCase().replace(/[_.\-]+/g, " ").replace(/\s+/g, " ");
}

function scoreHeaderForField(header, keywords) {
  const h = normalizeHeader(header);
  if (!h) return 0;
  let best = 0;
  keywords.forEach(k => {
    if (h === k) best = Math.max(best, 100);
    else if (h.includes(k)) best = Math.max(best, 60 + Math.min(k.length, 30));
    else if (k.includes(h) && h.length > 2) best = Math.max(best, 45);
  });
  return best;
}

// Greedily assigns each field to its best-scoring, not-yet-used header.
function autoMapHeaders(headers, keywordMap) {
  const mapping = {};
  const used = new Set();
  const order = Object.entries(keywordMap).sort((a, b) => b[1].join("").length - a[1].join("").length);
  order.forEach(([field, keywords]) => {
    let best = null, bestScore = 0;
    headers.forEach(h => {
      if (used.has(h)) return;
      const s = scoreHeaderForField(h, keywords);
      if (s > bestScore) { bestScore = s; best = h; }
    });
    mapping[field] = bestScore >= 45 ? best : null;
    if (mapping[field]) used.add(mapping[field]);
  });
  return mapping;
}

function applyMapping(rows, mapping) {
  return rows.map(row => {
    const out = {};
    Object.entries(mapping).forEach(([field, header]) => {
      out[field] = header != null ? String(row[header] ?? "").trim() : "";
    });
    return out;
  });
}

// Loosely maps whatever text is in the Change Type cell (e.g. "Mobile No.",
// "Business Name Update", "Timing change") onto one of the four supported
// change types, so the source sheet doesn't need to use our exact wording.
function normalizeChangeTypeValue(v) {
  const s = normalizeHeader(v);
  if (!s) return "";
  const exact = CHANGE_TYPES.find(ct => ct.toLowerCase() === s);
  if (exact) return exact;
  if (/whatsapp/.test(s)) return "WhatsApp Number";
  if (/phone|mobile|contact/.test(s)) return "Phone";
  if (/categor/.test(s)) return "Category";
  if (/website|web site|\burl\b/.test(s)) return "Website";
  if (/lat|long|coordinate|gps|geo|pin/.test(s)) return "Map Pin (Lat/Long)";
  if (/service area|coverage area/.test(s)) return "Service Area";
  if (/service/.test(s)) return "Services";
  if (/product/.test(s)) return "Products";
  if (/photo|image/.test(s)) return "Photos";
  if (/description|about/.test(s)) return "Description";
  if (/name/.test(s)) return "Business Name";
  if (/address|street|location/.test(s)) return "Address";
  if (/hour|timing|time/.test(s)) return "Business Hours";
  return v;
}

function isExpectedMappingUsable(mapping) {
  const identifierPresent = ["gbpLocationId", "storeCode", "locationId", "branchName"].some(k => mapping[k]);
  return !!(mapping.changeType && mapping.oldValue && mapping.expectedNewValue && identifierPresent);
}
function isGbpMappingUsable(mapping) {
  return !!(mapping.gbpLocationId || mapping.storeCode || mapping.businessName);
}

function parseUploadedFile(file, onData, onError) {
  const isCSV = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  if (isCSV) {
    reader.onload = (e) => {
      const result = Papa.parse(e.target.result, { header: true, skipEmptyLines: true });
      onData(result.data);
    };
    reader.onerror = () => onError("Could not read the file.");
    reader.readAsText(file);
  } else {
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        onData(json);
      } catch (err) { onError("Could not parse this Excel file. Try exporting as CSV."); }
    };
    reader.onerror = () => onError("Could not read the file.");
    reader.readAsBinaryString(file);
  }
}

function downloadCSV(filename, rows) {
  const safeRows = rows && rows.length ? rows : [{ "No data": "" }];
  const headers = Object.keys(safeRows[0]);
  const esc = (v) => {
    let s = v == null ? "" : String(v);
    if (/[",\n]/.test(s)) s = `"${s.replace(/"/g, '""')}"`;
    return s;
  };
  const csv = [headers.join(","), ...safeRows.map(r => headers.map(h => esc(r[h])).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

/* --------------------------- VERIFICATION ENGINE ---------------------------- */

function processExpectedChange(change, provider, settings) {
  const gbpLocations = provider.getLocations();
  const match = matchLocation(change, gbpLocations);
  const now = new Date();
  const base = {
    id: nextId("rec"),
    changeId: change.id,
    client: change.clientName || "Unknown Client",
    branch: change.branchName || "-",
    locationId: change.locationId || "-",
    storeCode: change.storeCode || "-",
    changeType: change.changeType,
    oldValue: change.oldValue,
    expectedValue: change.expectedNewValue,
    priority: change.priority || "Normal",
    requestDate: change.requestDate || now.toISOString().slice(0, 10),
    requestedBy: change.requestedBy || "-",
    matchMethod: match.method,
    matchConfidence: match.confidence,
    matchedLocationId: match.location ? match.location._uid : null,
    createdAt: now,
    attempts: 0,
    currentGbpValue: null,
    unexpectedChange: null,
    nextCheckDemoAt: null,
    source: provider.sourceLabel,
    reviewed: false,
    notes: [],
    isDemo: false,
    checks: [{ n: 0, time: now, event: "Change request created" }],
  };

  if (!match.location || match.confidence === "low") {
    return { ...base, status: "GRAY" };
  }

  const currentVal = getCurrentValueForLocation(change.changeType, match.location);

  // If your GBP Data upload doesn't include this field (common for Photos,
  // Services, Products, Description, Service Area, Map Pin — these rarely
  // have a simple bulk export), don't force a comparison against nothing.
  // Hand it to a manual visual/human check instead of guessing or silently
  // escalating to a false exception over time.
  if (currentVal == null) {
    return {
      ...base, status: "YELLOW", requiresManualCheck: true, currentGbpValue: null,
      checks: [...base.checks, { n: 0, time: now, event: `Waiting on manual check — no "${change.changeType}" data supplied in the GBP Data upload for this location` }],
    };
  }

  const isVerified = valuesMatch(change.changeType, currentVal, change.expectedNewValue, settings);

  let unexpectedChange = null;
  if (change.changeType !== "Business Name" && change.branchName) {
    const sim = similarity(normalizeName(change.branchName), normalizeName(match.location.businessName));
    if (sim < 0.55) {
      unexpectedChange = { field: "Business Name", from: change.branchName, to: match.location.businessName };
    }
  }

  const checks = [...base.checks, {
    n: 1, time: now, gbpValue: currentVal, result: isVerified ? "VERIFIED" : "PENDING",
  }];

  if (isVerified) {
    return { ...base, currentGbpValue: currentVal, attempts: 1, checks, status: "GREEN", verifiedAt: now, unexpectedChange };
  }
  return {
    ...base, currentGbpValue: currentVal, attempts: 1, checks, status: "YELLOW", unexpectedChange,
    nextCheckDemoAt: Date.now() + DEMO_SECONDS[0] * 1000,
  };
}

function advanceCheck(record, provider, settings) {
  const loc = provider.getLocation(record.matchedLocationId);
  const currentVal = loc ? getCurrentValueForLocation(record.changeType, loc) : record.currentGbpValue;
  const attempts = record.attempts + 1;
  const isVerified = currentVal != null && valuesMatch(record.changeType, currentVal, record.expectedValue, settings);
  const now = new Date();
  const result = isVerified ? "VERIFIED" : (attempts >= settings.maxAttempts ? "EXCEPTION" : "PENDING");
  const checks = [...record.checks, { n: attempts, time: now, gbpValue: currentVal, result }];

  let status = record.status, nextCheckDemoAt = null, verifiedAt = record.verifiedAt;
  if (isVerified) { status = "GREEN"; verifiedAt = now; }
  else if (attempts >= settings.maxAttempts) { status = "RED"; }
  else {
    status = "YELLOW";
    const idx = Math.min(attempts, DEMO_SECONDS.length - 1);
    nextCheckDemoAt = Date.now() + DEMO_SECONDS[idx] * 1000;
  }
  return { ...record, attempts, checks, currentGbpValue: currentVal, status, nextCheckDemoAt, verifiedAt };
}

/* -------------------------------- APP -------------------------------------- */

const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "quickcheck", label: "Quick Check", icon: Search },
  { key: "today", label: "Today's Verification", icon: Sun },
  { key: "expected", label: "Expected Changes", icon: FileText },
  { key: "gbpdata", label: "GBP Data (Master)", icon: Database },
  { key: "exceptions", label: "Exceptions", icon: AlertTriangle },
  { key: "clients", label: "Clients", icon: Users },
  { key: "insights", label: "Insights", icon: TrendingUp },
  { key: "audit", label: "Audit Log", icon: ScrollText },
  { key: "settings", label: "Settings", icon: SettingsIcon },
  { key: "help", label: "Help", icon: HelpCircle },
];

function AppInner() {
  const [gbpLocations, setGbpLocations] = useState(() => buildDemoGbpLocations().map(ensureUid));
  const [expectedChanges, setExpectedChanges] = useState(() => buildDemoUnprocessedChanges(buildDemoGbpLocations()));
  const [records, setRecords] = useState(() => {
    const locs = gbpLocations;
    return buildDemoRecords(locs);
  });
  const [auditLog, setAuditLog] = useState(() => seedAuditFromRecords(records));
  // Supports a one-click entry point: a bookmarklet or link like
  // ?qc=1&q=Business+Name&changeType=Phone opens the app straight into Quick
  // Check with the search box (and optionally the field) pre-filled — see
  // the Quick Check page for the bookmarklet snippet itself.
  const urlParams = useMemo(() => new URLSearchParams(window.location.search), []);
  const initialQuickCheck = urlParams.get("qc") ? {
    query: urlParams.get("q") || "",
    changeType: urlParams.get("changeType") || "",
  } : null;

  const [settings, setSettings] = useState({
    intervalsMinutes: DEFAULT_INTERVALS_MIN, maxAttempts: 4, demoAutoCheck: true, currentUser: "R. Menon",
    coordToleranceMeters: 50,
  });
  const [page, setPage] = useState(initialQuickCheck ? "quickcheck" : "dashboard");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [selectedRecordId, setSelectedRecordId] = useState(null);
  const [filters, setFilters] = useState({ client: "All", changeType: "All", status: "All", priority: "All", q: "" });
  const [toast, setToast] = useState(null);
  const [selectedClient, setSelectedClient] = useState("ICL Fincorp");
  const [pendingImport, setPendingImport] = useState(null); // { kind: 'expected'|'gbp', headers, rows, mapping }

  const gbpRef = useRef(gbpLocations);
  const settingsRef = useRef(settings);
  useEffect(() => { gbpRef.current = gbpLocations; }, [gbpLocations]);
  useEffect(() => { settingsRef.current = settings; }, [settings]);
  const provider = useMemo(() => new FileUploadGBPProvider(gbpLocations), [gbpLocations]);

  function pushAudit(entries) {
    setAuditLog(prev => [...entries, ...prev]);
  }

  function showToast(msg, tone = "default") {
    setToast({ msg, tone, id: nextId("t") });
    setTimeout(() => setToast(t => (t && t.msg === msg ? null : t)), 3200);
  }

  // Live demo recheck clock — only runs while this tab is open.
  useEffect(() => {
    const timer = setInterval(() => {
      if (!settingsRef.current.demoAutoCheck) return;
      const now = Date.now();
      const liveProvider = new FileUploadGBPProvider(gbpRef.current);
      let audits = [];
      setRecords(prev => prev.map(r => {
        if (r.status !== "YELLOW" || !r.nextCheckDemoAt || now < r.nextCheckDemoAt) return r;
        const updated = advanceCheck(r, liveProvider, settingsRef.current);
        audits.push(auditEntryFor(updated, settingsRef.current, "Automatic recheck (simulated)"));
        return updated;
      }));
      if (audits.length) pushAudit(audits);
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  function auditEntryFor(record, settingsSnapshot, actionTaken) {
    const lastCheck = record.checks[record.checks.length - 1];
    return {
      id: nextId("aud"), timestamp: new Date(), user: settingsSnapshot.currentUser,
      client: record.client, location: record.branch, changeType: record.changeType,
      expectedValue: record.expectedValue, gbpValue: record.currentGbpValue,
      result: record.status, attemptNumber: record.attempts, source: record.source,
      actionTaken, notes: lastCheck && lastCheck.event ? lastCheck.event : "",
    };
  }

  function runVerification() {
    const unprocessed = expectedChanges.filter(c => !records.some(r => r.changeId === c.id));
    if (unprocessed.length === 0) {
      showToast("No new expected changes to process — upload a file or add rows first.");
      return;
    }
    const newRecords = unprocessed.map(c => processExpectedChange(c, provider, settings));
    setRecords(prev => [...newRecords, ...prev]);
    pushAudit(newRecords.map(r => auditEntryFor(r, settings, "Verification run")));
    showToast(`Run complete — processed ${newRecords.length} expected change${newRecords.length > 1 ? "s" : ""}.`, "success");
  }

  // Quick Check: check one listing right now, no spreadsheet involved. Reuses
  // the exact same matching + comparison engine as the bulk upload path —
  // just skips straight to it with a location the user already picked.
  function submitQuickCheck(location, changeType, oldValue, expectedNewValue) {
    const change = {
      id: nextId("chg"),
      clientName: location.clientName || "Unknown Client",
      locationId: location.internalLocationId || "",
      storeCode: location.storeCode || "",
      branchName: location.businessName || location.storeCode || "",
      gbpLocationId: location.gbpLocationId || "",
      changeType, oldValue, expectedNewValue,
      requestDate: new Date().toISOString().slice(0, 10),
      requestedBy: settings.currentUser,
      priority: "Normal",
    };
    setExpectedChanges(prev => [...prev, change]);
    const record = processExpectedChange(change, provider, settings);
    setRecords(prev => [record, ...prev]);
    pushAudit([auditEntryFor(record, settings, "Quick Check")]);
    return record;
  }

  function retryVerification(recordId) {
    setRecords(prev => prev.map(r => {
      if (r.id !== recordId) return r;
      if (r.requiresManualCheck) return r; // handled via manuallyConfirmRecord, not the automatic pipeline
      const updated = advanceCheck({ ...r, status: "YELLOW" }, provider, settings);
      pushAudit([auditEntryFor(updated, settings, "Manual retry")]);
      return updated;
    }));
    showToast("Retry check complete.");
  }

  // Available on ANY pending/exception record, not just ones lacking data —
  // ops sometimes eyeballs the live listing directly (faster than waiting on
  // a data refresh). This records the outcome AND patches the underlying
  // master GBP data for this one field/listing, so the record doesn't drift
  // back to Pending on the next check just because the bulk data is stale.
  function manuallyConfirmRecord(recordId, isLive) {
    const now = new Date();
    let target = null;
    setRecords(prev => prev.map(r => {
      if (r.id !== recordId) return r;
      const attempts = r.attempts + 1;
      const checks = [...r.checks, {
        n: attempts, time: now,
        gbpValue: isLive ? r.expectedValue : (r.currentGbpValue ?? "Confirmed not changed (manual check)"),
        result: isLive ? "VERIFIED" : "EXCEPTION",
      }];
      const updated = {
        ...r, attempts, checks, status: isLive ? "GREEN" : "RED",
        verifiedAt: isLive ? now : r.verifiedAt,
        currentGbpValue: isLive ? r.expectedValue : r.currentGbpValue,
        nextCheckDemoAt: null,
      };
      pushAudit([auditEntryFor(updated, settings, isLive ? "Manually confirmed live (checked directly on Google)" : "Manually confirmed still incorrect (checked directly on Google)")]);
      target = updated;
      return updated;
    }));
    if (target && isLive && target.matchedLocationId && target.changeType !== "Map Pin (Lat/Long)") {
      setGbpLocations(gl => gl.map(l => l._uid === target.matchedLocationId
        ? { ...l, [FIELD_KEY[target.changeType]]: target.expectedValue }
        : l));
    }
    showToast(isLive ? "Marked verified — your master data for this listing was updated too, so it won't drift back to Pending." : "Marked as exception from your manual check.");
  }

  function markReviewed(recordId) {
    setRecords(prev => prev.map(r => r.id === recordId ? { ...r, reviewed: true } : r));
    const r = records.find(x => x.id === recordId);
    if (r) pushAudit([auditEntryFor(r, settings, "Marked as reviewed")]);
    showToast("Marked as reviewed.");
  }

  function addNote(recordId, text) {
    if (!text.trim()) return;
    setRecords(prev => prev.map(r => r.id === recordId
      ? { ...r, notes: [...(r.notes || []), { time: new Date(), text, author: settings.currentUser }] }
      : r));
    const r = records.find(x => x.id === recordId);
    if (r) pushAudit([{ ...auditEntryFor(r, settings, "Note added"), notes: text }]);
  }

  function assignRecord(recordId, assignee) {
    setRecords(prev => prev.map(r => r.id === recordId ? { ...r, assignedTo: assignee } : r));
    showToast(`Assigned to ${assignee}.`);
  }

  // ---- uploads ----
  function importExpectedRows(mappedRaw) {
    const normalized = mappedRaw.map(r => ({ ...r, changeType: normalizeChangeTypeValue(r.changeType) }));
    const valid = normalized.filter(r => CHANGE_TYPES.includes(r.changeType) && r.oldValue !== "" && r.expectedNewValue !== "");
    if (!valid.length) {
      showToast(`Columns were mapped, but ${normalized.length} row(s) didn't have a recognizable Change Type, Old Value, and Expected Value together. Nothing was imported.`, "error");
      return;
    }
    const withIds = valid.map(r => ({ ...r, id: nextId("chg") }));
    setExpectedChanges(prev => [...prev, ...withIds]);
    const skipped = normalized.length - valid.length;
    showToast(`Imported ${withIds.length} expected change row(s)${skipped ? ` (${skipped} skipped — missing values)` : ""}. Click "Run Verification" to process.`, "success");
  }
  function importGbpRows(mapped) {
    const valid = mapped.filter(r => r.businessName || r.gbpLocationId || r.storeCode);
    if (!valid.length) {
      showToast("Columns were mapped, but no row had a Business Name, GBP Location ID, or Store Code. Nothing was imported.", "error");
      return;
    }
    // Flag store codes that map to more than one distinct business name in
    // this batch — that's a real data-quality issue upstream (duplicate or
    // reused store codes), and it's the other realistic way two different
    // listings could get mixed up during matching. We still import the data
    // (matching falls back to GBP Location ID / name+address for these), but
    // ops should know.
    const nameByCode = new Map();
    valid.forEach(r => {
      if (!r.storeCode) return;
      const names = nameByCode.get(r.storeCode) || new Set();
      if (r.businessName) names.add(normalizeName(r.businessName));
      nameByCode.set(r.storeCode, names);
    });
    const dupeCodes = [...nameByCode.entries()].filter(([, names]) => names.size > 1).map(([code]) => code);

    setGbpLocations(prev => mergeGbpLocations(prev, valid));
    showToast(
      dupeCodes.length
        ? `Imported ${valid.length} row(s) — heads up: ${dupeCodes.length} store code(s) are shared by more than one business name in this file (e.g. ${dupeCodes.slice(0, 3).join(", ")}), which can cause a Store Code match to pick the wrong listing.`
        : `Imported ${valid.length} GBP location row(s).`,
      dupeCodes.length ? "error" : "success"
    );
  }

  function handleIncomingRows(rows, kind) {
    try {
      if (!rows || !rows.length) { showToast("That looks empty — nothing to import.", "error"); return; }
      const headers = Object.keys(rows[0]).filter(h => h && h.trim());
      if (!headers.length) { showToast("Couldn't find any column headers in that data.", "error"); return; }
      const keywordMap = kind === "expected" ? EXPECTED_KEYWORDS : GBP_KEYWORDS;
      const mapping = autoMapHeaders(headers, keywordMap);
      const usable = kind === "expected" ? isExpectedMappingUsable(mapping) : isGbpMappingUsable(mapping);
      if (usable) {
        const mapped = applyMapping(rows, mapping);
        if (kind === "expected") importExpectedRows(mapped); else importGbpRows(mapped);
      } else {
        setPendingImport({ kind, headers, rows, mapping });
      }
    } catch (e) {
      showToast("Something about that data couldn't be read. Double-check it has a header row and try again.", "error");
    }
  }

  function handleExpectedUpload(file) {
    parseUploadedFile(file, (rows) => handleIncomingRows(rows, "expected"), (err) => showToast(err, "error"));
  }
  function handleGbpUpload(file) {
    parseUploadedFile(file, (rows) => handleIncomingRows(rows, "gbp"), (err) => showToast(err, "error"));
  }
  function handlePasteRows(text, kind) {
    try {
      const result = Papa.parse(text.trim(), { header: true, skipEmptyLines: true });
      if (result.errors && result.errors.length && (!result.data || !result.data.length)) {
        showToast("Couldn't read that as rows — make sure the first line is column headers.", "error");
        return;
      }
      handleIncomingRows(result.data, kind);
    } catch (e) {
      showToast("Couldn't read that pasted data.", "error");
    }
  }

  function updateImportMapping(field, header) {
    setPendingImport(p => p ? { ...p, mapping: { ...p.mapping, [field]: header || null } } : p);
  }
  function confirmImport() {
    if (!pendingImport) return;
    const mapped = applyMapping(pendingImport.rows, pendingImport.mapping);
    if (pendingImport.kind === "expected") importExpectedRows(mapped);
    else importGbpRows(mapped);
    setPendingImport(null);
  }
  function cancelImport() { setPendingImport(null); }

  function simulateGoogleUpdate(recordId) {
    const r = records.find(x => x.id === recordId);
    if (!r || !r.matchedLocationId) return;
    setGbpLocations(gl => gl.map(l => {
      if (l._uid !== r.matchedLocationId) return l;
      if (r.changeType === "Map Pin (Lat/Long)") {
        const parsed = parseLatLng(r.expectedValue);
        return parsed ? { ...l, lat: parsed.lat, lng: parsed.lng } : l;
      }
      return { ...l, [FIELD_KEY[r.changeType]]: r.expectedValue };
    }));
    showToast("Demo only: pushed the expected value into the GBP snapshot so you can watch the next check verify it.");
  }

  const clients = useMemo(() => Array.from(new Set(records.map(r => r.client))), [records]);

  const filteredRecords = useMemo(() => {
    return records.filter(r => {
      if (filters.client !== "All" && r.client !== filters.client) return false;
      if (filters.changeType !== "All" && r.changeType !== filters.changeType) return false;
      if (filters.status !== "All" && r.status !== filters.status) return false;
      if (filters.priority !== "All" && r.priority !== filters.priority) return false;
      if (filters.q) {
        const q = filters.q.toLowerCase();
        const hay = `${r.client} ${r.branch} ${r.locationId} ${r.changeType} ${r.oldValue} ${r.expectedValue}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [records, filters]);

  const counts = useMemo(() => ({
    total: filteredRecords.length,
    green: filteredRecords.filter(r => r.status === "GREEN").length,
    yellow: filteredRecords.filter(r => r.status === "YELLOW").length,
    red: filteredRecords.filter(r => r.status === "RED").length,
    gray: filteredRecords.filter(r => r.status === "GRAY").length,
  }), [filteredRecords]);

  const selectedRecord = records.find(r => r.id === selectedRecordId) || null;

  function goDetail(id) { setSelectedRecordId(id); setPage("detail"); }

  return (
    <div style={{ fontFamily: "'Inter', ui-sans-serif, system-ui, -apple-system, sans-serif" }} className="min-h-screen w-full bg-slate-50 text-slate-900 flex">
      {/* -------------------------------- SIDEBAR -------------------------------- */}
      {mobileNavOpen && (
        <div className="fixed inset-0 bg-black/30 z-40 lg:hidden" onClick={() => setMobileNavOpen(false)} />
      )}
      <aside className={`flex flex-col w-60 shrink-0 bg-white border-r border-slate-200 py-5 px-3 overflow-y-auto
        fixed inset-y-0 left-0 z-50 h-screen transition-transform duration-200
        ${mobileNavOpen ? "translate-x-0" : "-translate-x-full"} lg:translate-x-0 lg:sticky lg:top-0`}>
        <div className="px-2 mb-6 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center text-white font-bold text-sm">GV</div>
            <div>
              <div className="font-semibold text-[13.5px] leading-tight text-slate-900">GBP Change</div>
              <div className="text-[11px] text-slate-500 leading-tight">Verification &amp; Monitoring</div>
            </div>
          </div>
          <button onClick={() => setMobileNavOpen(false)} className="lg:hidden p-1 text-slate-400 hover:text-slate-700">
            <X size={18} />
          </button>
        </div>
        <nav className="flex-1 space-y-0.5">
          {NAV.map(item => {
            const Icon = item.icon;
            const active = page === item.key || (page === "detail" && item.key === "dashboard");
            return (
              <button key={item.key} onClick={() => { setPage(item.key); setMobileNavOpen(false); }}
                className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-md text-[13px] font-medium transition-colors ${active ? "bg-indigo-50 text-indigo-700" : "text-slate-600 hover:bg-slate-100"}`}>
                <Icon size={16} strokeWidth={2} />
                {item.label}
              </button>
            );
          })}
        </nav>
        <div className="mt-4 mx-1 p-3 rounded-lg bg-slate-50 border border-slate-200 text-[11px] text-slate-500 leading-snug">
          Complements the SI <b>Out of Sync</b> feature. This tool verifies requested GBP changes and monitors for exceptions — it does not sync data.
        </div>
      </aside>

      {/* -------------------------------- MAIN -------------------------------- */}
      <div className="flex-1 min-w-0 flex flex-col">
        <TopBar
          onRun={runVerification}
          onMenu={() => setMobileNavOpen(true)}
          currentPageLabel={(NAV.find(n => n.key === page) || {}).label || "GBP Change Verification"}
          pendingCount={expectedChanges.filter(c => !records.some(r => r.changeId === c.id)).length}
          onExport={() => downloadCSV("full_verification_report.csv", records.map(exportRow))}
          filters={filters} setFilters={setFilters} clients={clients}
        />
        {toast && (
          <div className={`mx-4 mt-3 md:mx-6 px-4 py-2.5 rounded-lg text-[13px] border flex items-center gap-2 ${toast.tone === "success" ? "bg-emerald-50 border-emerald-200 text-emerald-800" : toast.tone === "error" ? "bg-red-50 border-red-200 text-red-800" : "bg-indigo-50 border-indigo-200 text-indigo-800"}`}>
            <Info size={14} /> {toast.msg}
          </div>
        )}
        {pendingImport && (
          <div className="mx-4 mt-3 md:mx-6">
            <ColumnMappingPanel
              pendingImport={pendingImport}
              fieldsMeta={pendingImport.kind === "expected" ? EXPECTED_FIELDS_META : GBP_FIELDS_META}
              onChange={updateImportMapping} onConfirm={confirmImport} onCancel={cancelImport}
            />
          </div>
        )}

        <main className="flex-1 p-4 md:p-6 max-w-[1400px] w-full mx-auto">
          {page === "dashboard" && (
            <Dashboard counts={counts} records={filteredRecords} onView={goDetail} onExport={downloadCSV} onQuickCheck={() => setPage("quickcheck")} />
          )}
          {page === "quickcheck" && (
            <QuickCheckPage gbpLocations={gbpLocations} onSubmit={submitQuickCheck} onView={goDetail} onGoToGbpData={() => setPage("gbpdata")} initial={initialQuickCheck} />
          )}
          {page === "today" && (
            <TodayView records={records} onView={goDetail} clients={clients} />
          )}
          {page === "expected" && (
            <ExpectedChangesPage
              expectedChanges={expectedChanges} records={records}
              onUpload={handleExpectedUpload} onRun={runVerification}
              onPasteRows={(text) => handlePasteRows(text, "expected")}
            />
          )}
          {page === "gbpdata" && (
            <GbpDataPage gbpLocations={gbpLocations} onUpload={handleGbpUpload}
              onPasteRows={(text) => handlePasteRows(text, "gbp")} />
          )}
          {page === "exceptions" && (
            <ExceptionsPage records={records} onView={goDetail} onReview={markReviewed} onRetry={retryVerification} onExport={downloadCSV} />
          )}
          {page === "clients" && (
            <ClientsPage records={records} clients={clients} selectedClient={selectedClient} setSelectedClient={setSelectedClient} />
          )}
          {page === "insights" && (
            <InsightsPage records={records} />
          )}
          {page === "audit" && (
            <AuditLogPage auditLog={auditLog} onExport={downloadCSV} />
          )}
          {page === "settings" && (
            <SettingsPage settings={settings} setSettings={setSettings} />
          )}
          {page === "help" && <HelpPage />}
          {page === "detail" && selectedRecord && (
            <DetailPage
              record={selectedRecord} onBack={() => setPage("dashboard")}
              onRetry={retryVerification} onReview={markReviewed} onNote={addNote}
              onAssign={assignRecord} onSimulateGoogleUpdate={simulateGoogleUpdate}
              onManualResolve={manuallyConfirmRecord}
              settings={settings}
            />
          )}
        </main>
      </div>
    </div>
  );
}

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { hasError: false }; }
  static getDerivedStateFromError() { return { hasError: true }; }
  componentDidCatch(error) { console.error("GBP Change Verification app error:", error); }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ fontFamily: "ui-sans-serif, system-ui, sans-serif" }} className="min-h-screen w-full bg-slate-50 flex items-center justify-center p-6">
          <div className="max-w-sm text-center bg-white rounded-xl border border-slate-200 p-6">
            <div className="text-[15px] font-semibold text-slate-800 mb-1.5">Something went wrong loading this view</div>
            <div className="text-[13px] text-slate-500 mb-4">Your uploaded data wasn't lost — this usually clears up with a refresh.</div>
            <button onClick={() => this.setState({ hasError: false })} className="px-4 py-1.5 rounded-md bg-indigo-600 text-white text-[13px] font-semibold hover:bg-indigo-700">
              Try again
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppInner />
    </ErrorBoundary>
  );
}

function seedAuditFromRecords(records) {
  const entries = [];
  records.forEach(r => {
    r.checks.forEach(c => {
      entries.push({
        id: nextId("aud"), timestamp: c.time, user: "System", client: r.client, location: r.branch,
        changeType: r.changeType, expectedValue: r.expectedValue, gbpValue: c.gbpValue ?? "-",
        result: c.result || "CREATED", attemptNumber: c.n, source: r.source,
        actionTaken: c.n === 0 ? "Verification record created" : "Automatic check",
        notes: c.event || "",
      });
    });
  });
  return entries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

function exportRow(r) {
  return {
    Client: r.client, Branch: r.branch, "Store Code": r.storeCode || "-", "Location ID": r.locationId, "Change Type": r.changeType,
    "Old Value": r.oldValue, "Expected Value": r.expectedValue, "Current GBP Value": r.currentGbpValue ?? "-",
    Status: `${r.status} — ${STATUS_META[r.status].label}`, "Last Checked": fmtDateTime(r.checks[r.checks.length - 1]?.time),
    "Attempt Count": r.attempts, "Match Method": r.matchMethod, Priority: r.priority, "Requested By": r.requestedBy,
  };
}

/* ------------------------------- COMPONENTS -------------------------------- */

function ColumnMappingPanel({ pendingImport, fieldsMeta, onChange, onConfirm, onCancel }) {
  const { kind, headers, rows, mapping } = pendingImport;
  const usable = kind === "expected" ? isExpectedMappingUsable(mapping) : isGbpMappingUsable(mapping);
  const previewRows = rows.slice(0, 3);
  return (
    <div className="bg-white rounded-xl border-2 border-indigo-300 p-4 mb-5">
      <div className="flex items-start gap-2 mb-3">
        <Info size={16} className="text-indigo-500 shrink-0 mt-0.5" />
        <div>
          <div className="text-[13.5px] font-semibold text-slate-800">We couldn't confidently match every column — point us to the right one</div>
          <div className="text-[12px] text-slate-500 mt-0.5">Nothing about your file needs to change. Pick which of your columns holds each piece of data below, then import.</div>
        </div>
      </div>

      <div className="grid sm:grid-cols-2 gap-x-6 gap-y-2 mb-4">
        {fieldsMeta.map(f => (
          <div key={f.key} className="flex items-center gap-2">
            <label className="text-[12.5px] text-slate-600 w-44 shrink-0">{f.label}{f.required && <span className="text-red-500"> *</span>}</label>
            <select value={mapping[f.key] || ""} onChange={e => onChange(f.key, e.target.value)}
              className="flex-1 text-[12.5px] px-2 py-1.5 rounded-md border border-slate-200 bg-white min-w-0">
              <option value="">— Not in file —</option>
              {headers.map(h => <option key={h} value={h}>{h}</option>)}
            </select>
          </div>
        ))}
      </div>

      {!usable && (
        <div className="text-[12px] text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-1.5 mb-3">
          {kind === "expected"
            ? "Map Change Type, Old Value, Expected New Value, plus at least one of Store Code / GBP Location ID / Branch Name to continue."
            : "Map at least Business Name, Store Code, or GBP Location ID to continue."}
        </div>
      )}

      {previewRows.length > 0 && (
        <div className="mb-4 overflow-x-auto border border-slate-100 rounded-md">
          <table className="text-[11.5px] w-full">
            <thead><tr className="bg-slate-50">{headers.map(h => <th key={h} className="px-2 py-1 text-left font-medium text-slate-500 whitespace-nowrap">{h}</th>)}</tr></thead>
            <tbody>
              {previewRows.map((r, i) => (
                <tr key={i} className="border-t border-slate-100">
                  {headers.map(h => <td key={h} className="px-2 py-1 whitespace-nowrap text-slate-600 max-w-[140px] truncate">{String(r[h] ?? "")}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="flex gap-2">
        <button onClick={onConfirm} disabled={!usable}
          className={`px-3.5 py-1.5 rounded-md text-[12.5px] font-semibold ${usable ? "bg-indigo-600 text-white hover:bg-indigo-700" : "bg-slate-100 text-slate-400 cursor-not-allowed"}`}>
          Import {rows.length} row{rows.length !== 1 ? "s" : ""}
        </button>
        <button onClick={onCancel} className="px-3.5 py-1.5 rounded-md border border-slate-200 text-[12.5px] font-medium text-slate-600 hover:bg-slate-50">Cancel</button>
      </div>
    </div>
  );
}

function CopyButton({ value, label = "Copy" }) {
  const [copied, setCopied] = useState(false);
  if (!value || value === "-") return null;
  const doCopy = async () => {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(value);
      } else {
        const ta = document.createElement("textarea");
        ta.value = value; ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta); ta.focus(); ta.select();
        document.execCommand("copy"); document.body.removeChild(ta);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch (e) { /* clipboard unavailable — silently no-op */ }
  };
  return (
    <button onClick={doCopy} title={`${label}: ${value}`}
      className="inline-flex items-center gap-1 text-[11px] font-medium text-indigo-600 hover:text-indigo-800 px-1.5 py-0.5 rounded hover:bg-indigo-50">
      {copied ? <><Check size={11} /> Copied</> : <><Copy size={11} /> Copy</>}
    </button>
  );
}

function StatusBadge({ status, size = "md" }) {
  const m = STATUS_META[status];
  const Icon = m.icon;
  const padding = size === "sm" ? "px-2 py-0.5 text-[11px]" : "px-2.5 py-1 text-[12px]";
  return (
    <span className={`inline-flex items-center gap-1.5 rounded-full font-semibold border ${padding}`}
      style={{ background: m.bg, color: m.text, borderColor: m.border }}>
      <Icon size={size === "sm" ? 11 : 13} />
      {status} — {m.label}
    </span>
  );
}

function PriorityChip({ p }) {
  const map = { Urgent: "bg-red-100 text-red-700", High: "bg-orange-100 text-orange-700", Normal: "bg-slate-100 text-slate-600", Low: "bg-slate-50 text-slate-400" };
  return <span className={`text-[11px] font-medium px-1.5 py-0.5 rounded ${map[p] || map.Normal}`}>{p}</span>;
}

function TopBar({ onRun, onMenu, currentPageLabel, pendingCount, onExport, filters, setFilters, clients }) {
  return (
    <div className="sticky top-0 z-20 bg-white/90 backdrop-blur border-b border-slate-200">
      <div className="lg:hidden flex items-center gap-2 px-4 pt-3">
        <button onClick={onMenu} className="p-1.5 -ml-1.5 rounded-md text-slate-500 hover:bg-slate-100">
          <Menu size={18} />
        </button>
        <span className="text-[13px] font-semibold text-slate-800">{currentPageLabel}</span>
      </div>
      <div className="px-4 md:px-6 py-3 flex flex-wrap items-center gap-2 max-w-[1400px] mx-auto w-full">
        <div className="relative flex-1 min-w-[180px] max-w-xs">
          <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={filters.q} onChange={e => setFilters(f => ({ ...f, q: e.target.value }))}
            placeholder="Search client, branch, value…"
            className="w-full pl-8 pr-3 py-1.5 text-[13px] rounded-md border border-slate-200 bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 focus:ring-indigo-200 focus:border-indigo-400" />
        </div>
        <FilterSelect label="Client" value={filters.client} onChange={v => setFilters(f => ({ ...f, client: v }))} options={["All", ...clients]} />
        <FilterSelect label="Change type" value={filters.changeType} onChange={v => setFilters(f => ({ ...f, changeType: v }))} options={["All", ...CHANGE_TYPES]} />
        <FilterSelect label="Status" value={filters.status} onChange={v => setFilters(f => ({ ...f, status: v }))} options={["All", "GREEN", "YELLOW", "RED", "GRAY"]} />
        <FilterSelect label="Priority" value={filters.priority} onChange={v => setFilters(f => ({ ...f, priority: v }))} options={["All", ...PRIORITIES]} />
        <div className="flex-1" />
        <button onClick={onExport} className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-[13px] font-medium border border-slate-200 text-slate-600 hover:bg-slate-50">
          <Download size={14} /> Export
        </button>
        <button onClick={onRun} className="relative inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-md text-[13px] font-semibold bg-indigo-600 text-white hover:bg-indigo-700 shadow-sm">
          <Play size={14} /> Run Verification
          {pendingCount > 0 && <span className="absolute -top-1.5 -right-1.5 bg-amber-500 text-white text-[10px] font-bold w-5 h-5 min-w-[18px] px-1 rounded-full flex items-center justify-center">{pendingCount}</span>}
        </button>
      </div>
    </div>
  );
}

function FilterSelect({ label, value, onChange, options }) {
  return (
    <select value={value} onChange={e => onChange(e.target.value)}
      title={label}
      className="text-[12.5px] px-2 py-1.5 rounded-md border border-slate-200 bg-white text-slate-600 focus:outline-none focus:ring-2 focus:ring-indigo-200 max-w-[140px]">
      {options.map(o => <option key={o} value={o}>{o === "All" ? `${label}: All` : o}</option>)}
    </select>
  );
}

function SummaryCards({ counts }) {
  const cards = [
    { label: "Total Changes", value: counts.total, tone: "slate" },
    { label: "Verified", value: counts.green, tone: "GREEN" },
    { label: "Pending", value: counts.yellow, tone: "YELLOW" },
    { label: "Exceptions", value: counts.red, tone: "RED" },
    { label: "Matching Required", value: counts.gray, tone: "GRAY" },
  ];
  return (
    <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-5">
      {cards.map(c => {
        const meta = STATUS_META[c.tone];
        return (
          <div key={c.label} className="bg-white rounded-xl border border-slate-200 p-4">
            <div className="text-[12px] text-slate-500 font-medium mb-1.5">{c.label}</div>
            <div className="text-2xl font-bold" style={{ color: meta ? meta.text : "#0F172A" }}>{c.value}</div>
          </div>
        );
      })}
    </div>
  );
}

function RecordsTable({ records, onView, compact }) {
  if (!records.length) {
    return (
      <div className="bg-white rounded-xl border border-slate-200 p-10 text-center text-slate-400 text-[13px]">
        No records match the current filters.
      </div>
    );
  }
  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full text-[12.5px]">
          <thead>
            <tr className="bg-slate-50 text-slate-500 text-left border-b border-slate-200">
              {["Client", "Branch", "Store Code", "Location ID", "Change Type", "Old Value", "Expected", "Current GBP", "Status", "Last Checked", "Next Check", "Attempts", "Match Method", ""].map(h => (
                <th key={h} className="px-3 py-2 font-semibold whitespace-nowrap">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {records.map(r => {
              const last = r.checks[r.checks.length - 1];
              return (
                <tr key={r.id} className="border-b border-slate-100 hover:bg-slate-50/70 transition-colors">
                  <td className="px-3 py-2 whitespace-nowrap font-medium text-slate-700">{r.client}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{r.branch}</td>
                  <td className="px-3 py-2 whitespace-nowrap font-mono text-[11.5px] text-slate-700 font-semibold">
                    <span className="inline-flex items-center gap-1">{r.storeCode || "-"}<CopyButton value={r.storeCode} label="Store Code" /></span>
                  </td>
                  <td className="px-3 py-2 whitespace-nowrap font-mono text-[11.5px] text-slate-500">{r.locationId}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{r.changeType}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500 max-w-[120px] truncate">{r.oldValue}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[120px] truncate">{r.expectedValue}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[120px] truncate">{r.currentGbpValue ?? (r.requiresManualCheck ? "Needs manual check" : "Waiting for GBP data")}</td>
                  <td className="px-3 py-2 whitespace-nowrap"><StatusBadge status={r.status} size="sm" /></td>
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500">{fmtDateTime(last?.time)}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-slate-400">{r.status === "YELLOW" ? (r.requiresManualCheck ? "Awaiting your check" : "Scheduled (simulated)") : "-"}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-center">{r.attempts}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500">{r.matchMethod}</td>
                  <td className="px-3 py-2 whitespace-nowrap">
                    <button onClick={() => onView(r.id)} className="text-indigo-600 hover:text-indigo-800 font-medium text-[12px]">
                      {r.status === "RED" ? "Investigate" : "View"} <ChevronRight size={12} className="inline" />
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

/* --------------------------------- PAGES ------------------------------------ */

function Dashboard({ counts, records, onView, onExport, onQuickCheck }) {
  return (
    <div>
      <PageHeader title="Dashboard" subtitle="Every expected GBP change, matched against the live GBP snapshot and classified automatically." />
      <button onClick={onQuickCheck} className="w-full text-left bg-indigo-600 hover:bg-indigo-700 transition-colors rounded-xl p-4 mb-5 flex items-center justify-between gap-3 text-white">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-white/15 flex items-center justify-center shrink-0"><Search size={18} /></div>
          <div>
            <div className="text-[14px] font-semibold">Check one listing right now</div>
            <div className="text-[12px] text-indigo-100">Pick a listing, tell it what changed, get an instant answer — no spreadsheet needed.</div>
          </div>
        </div>
        <ChevronRight size={18} />
      </button>
      <SummaryCards counts={counts} />
      <div className="flex items-center justify-between mb-3">
        <div className="text-[13px] font-semibold text-slate-700">All verification records</div>
        <button onClick={() => onExport("full_verification_report.csv", records.map(exportRow))} className="text-[12px] text-indigo-600 font-medium flex items-center gap-1">
          <Download size={13} /> Export view
        </button>
      </div>
      <RecordsTable records={records} onView={onView} />
    </div>
  );
}

function BookmarkletHelper() {
  const [open, setOpen] = useState(false);
  const appUrl = typeof window !== "undefined" ? window.location.origin + window.location.pathname : "";
  const bookmarklet = `javascript:(function(){var t=document.title||'';t=t.replace(/\\s*-\\s*Google (Maps|Search).*$/i,'').trim();window.open('${appUrl}?qc=1&q='+encodeURIComponent(t),'_blank');})();`;

  return (
    <div className="mt-4">
      <button onClick={() => setOpen(o => !o)} className="text-[12px] text-indigo-600 font-medium hover:text-indigo-800">
        {open ? "Hide one-click setup ▲" : "Set up one-click checking from Google Maps ▾"}
      </button>
      {open && (
        <div className="mt-2 bg-white rounded-lg border border-slate-200 p-3.5">
          <div className="text-[12px] text-slate-500 mb-3 leading-relaxed">
            Drag this button to your bookmarks bar. Next time you're looking at a listing on Google Maps or Search, click it — this app opens straight into Quick Check with that listing already found (and auto-selected if there's only one match).
          </div>
          <div className="flex items-center gap-2 flex-wrap">
            <a href={bookmarklet} draggable
              onClick={(e) => e.preventDefault()}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-md bg-indigo-600 text-white text-[12.5px] font-semibold cursor-grab active:cursor-grabbing">
              <Search size={13} /> Check on GBP Tool
            </a>
            <CopyButton value={bookmarklet} label="Bookmarklet code" />
          </div>
          <div className="text-[11px] text-slate-400 mt-2">
            If dragging doesn't work in your browser: add any new bookmark, then paste the copied code as its URL.
          </div>
        </div>
      )}
    </div>
  );
}

function QuickCheckPage({ gbpLocations, onSubmit, onView, onGoToGbpData, initial }) {
  const [query, setQuery] = useState(initial?.query || "");
  const [selected, setSelected] = useState(null);
  const [changeType, setChangeType] = useState(initial?.changeType && CHANGE_TYPES.includes(initial.changeType) ? initial.changeType : "Phone");
  const [oldValue, setOldValue] = useState("");
  const [newValue, setNewValue] = useState("");
  const [result, setResult] = useState(null);
  const [autoPrefilled] = useState(!!initial?.query);

  const matches = useMemo(() => {
    if (!query.trim() || selected) return [];
    const q = query.trim().toLowerCase();
    return gbpLocations.filter(l =>
      (l.businessName || "").toLowerCase().includes(q) ||
      (l.storeCode || "").toLowerCase().includes(q) ||
      (l.gbpLocationId || "").toLowerCase().includes(q)
    ).slice(0, 8);
  }, [query, selected, gbpLocations]);

  // Came in from the bookmarklet/link with a specific listing name — if that
  // uniquely identifies one listing, jump straight to it with no extra click.
  useEffect(() => {
    if (autoPrefilled && matches.length === 1 && !selected) {
      pickLocation(matches[0]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [autoPrefilled, matches.length]);

  function pickLocation(loc) {
    setSelected(loc);
    setQuery("");
    setResult(null);
    const current = getCurrentValueForLocation(changeType, loc);
    setOldValue(current || "");
    setNewValue("");
  }
  function changeChangeType(ct) {
    setChangeType(ct);
    if (selected) setOldValue(getCurrentValueForLocation(ct, selected) || "");
    setNewValue("");
    setResult(null);
  }
  function reset() {
    setSelected(null); setQuery(""); setOldValue(""); setNewValue(""); setResult(null);
  }
  function submit() {
    if (!selected || !newValue.trim()) return;
    const record = onSubmit(selected, changeType, oldValue, newValue.trim());
    setResult(record);
  }

  if (!gbpLocations.length) {
    return (
      <div>
        <PageHeader title="Quick Check" subtitle="Check one listing right now — no spreadsheet needed." />
        <div className="bg-white rounded-xl border border-slate-200 p-10 text-center">
          <Database size={26} className="mx-auto text-slate-300 mb-3" />
          <div className="text-[14px] font-semibold text-slate-700 mb-1">Upload your GBP master data first</div>
          <div className="text-[12.5px] text-slate-500 mb-4 max-w-sm mx-auto">Quick Check searches the listings already loaded on the GBP Data page. Upload that once, then this stays available every day.</div>
          <button onClick={onGoToGbpData} className="px-4 py-1.5 rounded-md bg-indigo-600 text-white text-[13px] font-semibold hover:bg-indigo-700">Go to GBP Data</button>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-2xl">
      <PageHeader title="Quick Check" subtitle="Pick a listing from your master data, say what changed, get an instant answer." />

      {!selected ? (
        <div>
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input value={query} onChange={e => setQuery(e.target.value)} autoFocus
              placeholder="Search by business name, store code, or GBP Location ID…"
              className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-200 text-[13.5px] focus:outline-none focus:ring-2 focus:ring-indigo-200 focus:border-indigo-400" />
          </div>
          {matches.length > 0 && (
            <div className="mt-2 bg-white rounded-lg border border-slate-200 divide-y divide-slate-100 overflow-hidden">
              {matches.map(l => (
                <button key={l._uid} onClick={() => pickLocation(l)} className="w-full text-left px-3.5 py-2.5 hover:bg-slate-50">
                  <div className="text-[13px] font-medium text-slate-800">{l.businessName}</div>
                  <div className="text-[11.5px] text-slate-400">{l.clientName ? `${l.clientName} · ` : ""}Store Code: {l.storeCode || "-"}</div>
                </button>
              ))}
            </div>
          )}
          {query.trim() && matches.length === 0 && (
            <div className="mt-2 text-[12.5px] text-slate-400 px-1">No listing found matching "{query}".</div>
          )}
          <BookmarkletHelper />
        </div>
      ) : (
        <div>
          <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4 flex items-start justify-between gap-3">
            <div>
              <div className="text-[14px] font-semibold text-slate-800">{selected.businessName}</div>
              <div className="text-[12px] text-slate-400 mt-0.5">{selected.clientName ? `${selected.clientName} · ` : ""}Store Code: {selected.storeCode || "-"}</div>
              <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2 text-[12px] text-slate-500">
                {selected.phone && <span>📞 {selected.phone}</span>}
                {selected.address && <span>📍 {selected.address}</span>}
                {selected.category && <span>🏷 {selected.category}</span>}
                {selected.website && <span>🌐 {selected.website}</span>}
              </div>
            </div>
            <button onClick={reset} className="text-[12px] text-indigo-600 font-medium shrink-0">Change listing</button>
          </div>

          {!result ? (
            <div className="bg-white rounded-xl border border-slate-200 p-4">
              <label className="block text-[12px] font-medium text-slate-500 mb-1.5">What changed?</label>
              <select value={changeType} onChange={e => changeChangeType(e.target.value)}
                className="w-full px-3 py-2 rounded-md border border-slate-200 text-[13.5px] mb-3">
                {CHANGE_TYPES.map(ct => <option key={ct} value={ct}>{ct}</option>)}
              </select>

              <label className="block text-[12px] font-medium text-slate-500 mb-1.5">Old value {oldValue && <span className="text-slate-400 font-normal">(auto-filled from master data — edit if needed)</span>}</label>
              <input value={oldValue} onChange={e => setOldValue(e.target.value)}
                placeholder="What it was before"
                className="w-full px-3 py-2 rounded-md border border-slate-200 text-[13.5px] mb-3" />

              <label className="block text-[12px] font-medium text-slate-500 mb-1.5">New value you set it to</label>
              <input value={newValue} onChange={e => setNewValue(e.target.value)}
                placeholder="What you changed it to in GBP"
                className="w-full px-3 py-2 rounded-md border border-slate-200 text-[13.5px] mb-4" />

              <button onClick={submit} disabled={!newValue.trim()}
                className={`w-full py-2 rounded-md text-[13.5px] font-semibold ${newValue.trim() ? "bg-indigo-600 text-white hover:bg-indigo-700" : "bg-slate-100 text-slate-400 cursor-not-allowed"}`}>
                Check Now
              </button>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-slate-200 p-4">
              <div className="flex items-center justify-between mb-3">
                <div className="text-[13.5px] font-semibold text-slate-800">Result</div>
                <StatusBadge status={result.status} />
              </div>
              <div className="grid grid-cols-2 gap-3 text-[12.5px] mb-4">
                <ValueCard label="Expected" value={result.expectedValue} />
                <ValueCard label="Current in Master Data" value={result.currentGbpValue ?? (result.requiresManualCheck ? "Needs manual check" : "No data")} highlight />
              </div>
              {result.status === "GREEN" && <div className="text-[12.5px] text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-md px-3 py-2 mb-3">Matches your master data — verified.</div>}
              {result.status === "YELLOW" && !result.requiresManualCheck && <div className="text-[12.5px] text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-2 mb-3">Doesn't match yet — re-check after your master data refreshes with a newer snapshot.</div>}
              {result.requiresManualCheck && <div className="text-[12.5px] text-indigo-700 bg-indigo-50 border border-indigo-200 rounded-md px-3 py-2 mb-3">No "{result.changeType}" data in your master data for this listing — open the record to confirm manually.</div>}
              {result.status === "GRAY" && <div className="text-[12.5px] text-slate-600 bg-slate-50 border border-slate-200 rounded-md px-3 py-2 mb-3">Couldn't confidently match this listing — check the location data.</div>}
              <div className="flex gap-2">
                <button onClick={() => onView(result.id)} className="flex-1 py-2 rounded-md bg-indigo-600 text-white text-[13px] font-semibold hover:bg-indigo-700">View full details</button>
                <button onClick={reset} className="flex-1 py-2 rounded-md border border-slate-200 text-slate-600 text-[13px] font-medium hover:bg-slate-50">Check another listing</button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function TodayView({ records, onView, clients }) {
  const todayStr = new Date().toDateString();
  const isToday = (d) => d && new Date(d).toDateString() === todayStr;
  const todays = records.filter(r => isToday(r.createdAt));
  const yestPending = records.filter(r => !isToday(r.createdAt) && r.status === "YELLOW");
  const prevExceptions = records.filter(r => !isToday(r.createdAt) && r.status === "RED" && !r.reviewed);
  const awaitingReview = records.filter(r => r.status === "GRAY");

  return (
    <div>
      <PageHeader title="Today's Verification" subtitle="A rolling view of everything created or checked today, plus anything carried over that still needs attention." />
      {clients.map(client => {
        const rows = todays.filter(r => r.client === client);
        if (!rows.length) return null;
        const v = rows.filter(r => r.status === "GREEN").length;
        const p = rows.filter(r => r.status === "YELLOW").length;
        const e = rows.filter(r => r.status === "RED").length;
        return (
          <div key={client} className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
            <div className="flex items-center justify-between mb-3">
              <div>
                <div className="text-[11px] text-slate-400 font-medium uppercase tracking-wide">Client</div>
                <div className="text-[15px] font-semibold">{client}</div>
              </div>
              <div className="flex gap-4 text-[12.5px]">
                <span className="text-slate-500">Today's changes: <b className="text-slate-800">{rows.length}</b></span>
                <span className="text-emerald-600 font-medium">{v} Verified</span>
                <span className="text-amber-600 font-medium">{p} Pending</span>
                <span className="text-red-600 font-medium">{e} Exceptions</span>
              </div>
            </div>
            <div className="divide-y divide-slate-100">
              {rows.map(r => (
                <button key={r.id} onClick={() => onView(r.id)} className="w-full flex items-center justify-between py-2 text-left hover:bg-slate-50 px-1.5 -mx-1.5 rounded">
                  <div className="text-[13px]">
                    <span className="font-medium text-slate-700">{r.branch}</span>
                    <span className="text-slate-400"> · {r.changeType} · </span>
                    <span className="text-slate-500">{fmtDateTime(r.createdAt)}</span>
                  </div>
                  <StatusBadge status={r.status} size="sm" />
                </button>
              ))}
            </div>
          </div>
        );
      })}
      {!records.some(r => isToday(r.createdAt)) && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 text-slate-400 text-[13px] mb-4">No verification activity recorded yet today.</div>
      )}

      <div className="grid md:grid-cols-3 gap-4 mt-2">
        <CarryoverCard title="Yesterday's Pending" icon={Clock} tone="YELLOW" rows={yestPending} onView={onView} />
        <CarryoverCard title="Previous Exceptions" icon={AlertTriangle} tone="RED" rows={prevExceptions} onView={onView} />
        <CarryoverCard title="Awaiting Manual Match" icon={CircleHelp} tone="GRAY" rows={awaitingReview} onView={onView} />
      </div>
    </div>
  );
}

function CarryoverCard({ title, icon: Icon, tone, rows, onView }) {
  const meta = STATUS_META[tone];
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4">
      <div className="flex items-center gap-2 mb-3">
        <Icon size={14} style={{ color: meta.text }} />
        <div className="text-[13px] font-semibold text-slate-700">{title}</div>
        <span className="ml-auto text-[11px] font-semibold px-1.5 py-0.5 rounded" style={{ background: meta.bg, color: meta.text }}>{rows.length}</span>
      </div>
      {rows.length === 0 ? (
        <div className="text-[12px] text-slate-400">Nothing here.</div>
      ) : (
        <div className="space-y-1.5">
          {rows.slice(0, 5).map(r => (
            <button key={r.id} onClick={() => onView(r.id)} className="w-full text-left text-[12px] text-slate-600 hover:text-indigo-600 truncate block">
              {r.branch} — {r.changeType}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function PageHeader({ title, subtitle, right }) {
  return (
    <div className="flex items-start justify-between mb-5 gap-4 flex-wrap">
      <div>
        <h1 className="text-[19px] font-bold text-slate-900">{title}</h1>
        {subtitle && <p className="text-[13px] text-slate-500 mt-0.5 max-w-2xl">{subtitle}</p>}
      </div>
      {right}
    </div>
  );
}

function PasteRowsBox({ onRows }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const doImport = () => {
    if (!text.trim()) return;
    onRows(text);
    setText("");
    setOpen(false);
  };
  return (
    <div className="mt-3">
      <button onClick={() => setOpen(o => !o)} className="text-[12px] text-indigo-600 font-medium hover:text-indigo-800">
        {open ? "Hide paste option ▲" : "Or paste rows instead of uploading a file ▾"}
      </button>
      {open && (
        <div className="mt-2 bg-white rounded-lg border border-slate-200 p-3">
          <div className="text-[11.5px] text-slate-500 mb-2">
            Select your rows in Google Sheets/Excel — including the header row — copy, and paste below. No file needed.
          </div>
          <textarea value={text} onChange={e => setText(e.target.value)} rows={5}
            placeholder="Paste copied rows here (first row should be your column headers)…"
            className="w-full text-[12px] font-mono px-2.5 py-2 rounded-md border border-slate-200 focus:outline-none focus:ring-2 focus:ring-indigo-200" />
          <button onClick={doImport} disabled={!text.trim()}
            className={`mt-2 px-3.5 py-1.5 rounded-md text-[12.5px] font-semibold ${text.trim() ? "bg-indigo-600 text-white hover:bg-indigo-700" : "bg-slate-100 text-slate-400 cursor-not-allowed"}`}>
            Import pasted rows
          </button>
        </div>
      )}
    </div>
  );
}

function UploadBox({ label, hint, onFile, templateRows, templateName, onPasteRows }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  return (
    <div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); const f = e.dataTransfer.files?.[0]; if (f) onFile(f); }}
        className={`rounded-xl border-2 border-dashed p-6 text-center transition-colors ${dragOver ? "border-indigo-400 bg-indigo-50/50" : "border-slate-200 bg-white"}`}>
        <UploadCloud size={26} className="mx-auto text-indigo-500 mb-2" />
        <div className="text-[13.5px] font-semibold text-slate-700">{label}</div>
        <div className="text-[12px] text-slate-500 mt-1 mb-3 max-w-sm mx-auto">{hint}</div>
        <div className="flex items-center justify-center gap-2">
          <button onClick={() => inputRef.current?.click()} className="px-3.5 py-1.5 rounded-md bg-indigo-600 text-white text-[12.5px] font-semibold hover:bg-indigo-700">
            Choose file
          </button>
          {templateRows && (
            <button onClick={() => downloadCSV(templateName, templateRows)} className="px-3.5 py-1.5 rounded-md border border-slate-200 text-slate-600 text-[12.5px] font-medium hover:bg-slate-50">
              Download template
            </button>
          )}
        </div>
        <input ref={inputRef} type="file" accept=".csv,.xlsx,.xls" className="hidden"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }} />
        <div className="text-[11px] text-slate-400 mt-2">Accepts .csv, .xlsx, .xls — or drag a file here</div>
      </div>
      {onPasteRows && <PasteRowsBox onRows={onPasteRows} />}
    </div>
  );
}

function ExpectedChangesPage({ expectedChanges, records, onUpload, onRun, onPasteRows }) {
  const unprocessed = expectedChanges.filter(c => !records.some(r => r.changeId === c.id));
  const template = [{
    "Client Name": "ICL Fincorp", "Location ID": "ICL-KL-KOC-001", "Store Code": "ICL-KL-KOC-001",
    "Branch Name": "Kochi", "GBP Location ID": "GBP-ICL-KOC", "Change Type": "Phone",
    "Old Value": "9876543200", "Expected New Value": "9876543210", "Request Date": "2026-07-31",
    "Requested By": "R. Menon", "Priority": "Normal",
  }];
  return (
    <div>
      <PageHeader title="Expected Changes" subtitle="Upload the changes requested through SI / V1 / V3. Each row becomes a verification record once you run verification." />
      <UploadBox label="Upload expected changes" hint="Excel or CSV export from SI, V1, V3, or a manually built list of requested GBP changes." onFile={onUpload} templateRows={template} templateName="expected_changes_template.csv" onPasteRows={onPasteRows} />

      <div className="flex items-center justify-between mt-6 mb-3">
        <div className="text-[13px] font-semibold text-slate-700">Uploaded rows ({expectedChanges.length})</div>
        <button onClick={onRun} disabled={!unprocessed.length}
          className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-md text-[13px] font-semibold ${unprocessed.length ? "bg-indigo-600 text-white hover:bg-indigo-700" : "bg-slate-100 text-slate-400 cursor-not-allowed"}`}>
          <Play size={14} /> Run Verification {unprocessed.length ? `(${unprocessed.length} pending)` : ""}
        </button>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-[12.5px]">
            <thead>
              <tr className="bg-slate-50 text-slate-500 text-left border-b border-slate-200">
                {["Client", "Branch", "Store Code", "GBP Location ID", "Change Type", "Old Value", "Expected", "Priority", "Requested", "Processed?"].map(h => <th key={h} className="px-3 py-2 font-semibold whitespace-nowrap">{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {expectedChanges.map(c => {
                const processed = records.some(r => r.changeId === c.id);
                return (
                  <tr key={c.id} className="border-b border-slate-100">
                    <td className="px-3 py-2 whitespace-nowrap font-medium">{c.clientName}</td>
                    <td className="px-3 py-2 whitespace-nowrap">{c.branchName}</td>
                    <td className="px-3 py-2 whitespace-nowrap font-mono text-[11.5px] text-slate-500">{c.storeCode}</td>
                    <td className="px-3 py-2 whitespace-nowrap font-mono text-[11.5px] text-slate-500">{c.gbpLocationId || "-"}</td>
                    <td className="px-3 py-2 whitespace-nowrap">{c.changeType}</td>
                    <td className="px-3 py-2 whitespace-nowrap text-slate-500">{c.oldValue}</td>
                    <td className="px-3 py-2 whitespace-nowrap">{c.expectedNewValue}</td>
                    <td className="px-3 py-2 whitespace-nowrap"><PriorityChip p={c.priority || "Normal"} /></td>
                    <td className="px-3 py-2 whitespace-nowrap text-slate-500">{c.requestDate}</td>
                    <td className="px-3 py-2 whitespace-nowrap">
                      {processed ? <span className="text-emerald-600 text-[11.5px] font-medium">Processed</span> : <span className="text-amber-600 text-[11.5px] font-medium">Awaiting run</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function GbpDataPage({ gbpLocations, onUpload, onPasteRows }) {
  const [clientFilter, setClientFilter] = useState("All");
  const [q, setQ] = useState("");
  const clients = useMemo(() => Array.from(new Set(gbpLocations.map(l => l.clientName).filter(Boolean))), [gbpLocations]);
  const filtered = useMemo(() => gbpLocations.filter(l => {
    if (clientFilter !== "All" && l.clientName !== clientFilter) return false;
    if (q) {
      const hay = `${l.businessName} ${l.storeCode} ${l.gbpLocationId} ${l.phone} ${l.address}`.toLowerCase();
      if (!hay.includes(q.toLowerCase())) return false;
    }
    return true;
  }), [gbpLocations, clientFilter, q]);
  const RENDER_CAP = 500;
  const shown = filtered.slice(0, RENDER_CAP);

  const template = [{
    "Client Name": "ICL Fincorp", "Business Name": "ICL Fincorp - Kochi", "Store Code": "ICL-KL-KOC-001",
    "GBP Location ID": "GBP-ICL-KOC", "Internal Location ID": "ICL-KOC-001", "Phone": "9876543210",
    "Address": "101, MG Road, Kochi", "Business Hours": "Mon-Sat 9:30 AM-6:30 PM, Sun Closed",
    "Latitude": "9.93", "Longitude": "76.26", "Category": "Financial Consultant",
    "Website": "https://iclfincorp.com/kochi", "WhatsApp Number": "9876543210",
    "Services": "Personal Loans, Home Loans, Insurance", "Products": "", "Photos": "",
    "Service Area": "Kochi, Ernakulam", "Description": "ICL Fincorp Kochi branch offering loans and insurance services.",
  }];
  const perClientCounts = clients.map(c => ({ c, n: gbpLocations.filter(l => l.clientName === c).length }));

  return (
    <div>
      <PageHeader title="GBP Data" subtitle="This is the current-state snapshot the comparison engine checks against. In the MVP it comes from a file export; the same interface can later be backed by the Google Business Profile API." />
      <div className="mb-5 p-3.5 rounded-lg bg-indigo-50 border border-indigo-100 text-[12.5px] text-indigo-800 flex gap-2">
        <Info size={15} className="shrink-0 mt-0.5" />
        <div>
          <b>Data provider:</b> File Upload Provider (active). A future <code className="bg-white/60 px-1 rounded">GoogleBusinessProfileAPIProvider</code> can replace this without changing matching, comparison, or the dashboard — see Settings.
        </div>
      </div>
      <UploadBox label="Upload GBP data" hint="A Google Business Profile export, or an equivalent snapshot of current listing values, per location. Only Business Name/Store Code/GBP Location ID are needed to match a location — every other column (Phone, Category, Website, Services, Description, etc.) is optional. Whatever you don't include just falls back to a manual check instead of failing." onFile={onUpload} templateRows={template} templateName="gbp_data_template.csv" onPasteRows={onPasteRows} />

      {clients.length > 0 && (
        <div className="flex flex-wrap gap-2 mt-4">
          {perClientCounts.map(({ c, n }) => (
            <div key={c} className="px-3 py-1.5 rounded-full text-[12px] font-medium bg-white border border-slate-200 text-slate-600">
              {c} <span className="text-slate-400">· {n} listings</span>
            </div>
          ))}
        </div>
      )}

      <div className="flex items-center justify-between mt-6 mb-3 gap-2 flex-wrap">
        <div className="text-[13px] font-semibold text-slate-700">Current GBP snapshot ({filtered.length} of {gbpLocations.length} locations)</div>
        <div className="flex gap-2">
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search name, store code, phone…"
            className="text-[12.5px] px-2.5 py-1.5 rounded-md border border-slate-200 bg-white w-56" />
          {clients.length > 0 && (
            <select value={clientFilter} onChange={e => setClientFilter(e.target.value)}
              className="text-[12.5px] px-2 py-1.5 rounded-md border border-slate-200 bg-white text-slate-600">
              <option value="All">Client: All</option>
              {clients.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          )}
        </div>
      </div>
      {filtered.length > RENDER_CAP && (
        <div className="mb-2 text-[12px] text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-1.5">
          Showing the first {RENDER_CAP} of {filtered.length} matching locations. Use the client filter or search to narrow the list — this doesn't affect matching or verification, only what's rendered here.
        </div>
      )}
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto max-h-[520px] overflow-y-auto">
          <table className="w-full text-[12.5px]">
            <thead className="sticky top-0 bg-slate-50">
              <tr className="text-slate-500 text-left border-b border-slate-200">
                {["Client", "Business Name", "Store Code", "GBP Location ID", "Phone", "Address", "Business Hours"].map(h => <th key={h} className="px-3 py-2 font-semibold whitespace-nowrap bg-slate-50">{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {shown.map(l => (
                <tr key={l._uid} className="border-b border-slate-100">
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500">{l.clientName || "-"}</td>
                  <td className="px-3 py-2 whitespace-nowrap font-medium">{l.businessName}</td>
                  <td className="px-3 py-2 whitespace-nowrap font-mono text-[11.5px] text-slate-500">{l.storeCode}</td>
                  <td className="px-3 py-2 whitespace-nowrap font-mono text-[11.5px] text-slate-500">{l.gbpLocationId}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{l.phone}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[220px] truncate">{l.address}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[220px] truncate">{l.hours}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ExceptionsPage({ records, onView, onReview, onRetry, onExport }) {
  const exceptions = records.filter(r => r.status === "RED");
  const outstanding = exceptions.filter(r => !r.reviewed);
  const reviewed = exceptions.filter(r => r.reviewed);
  return (
    <div>
      <PageHeader title="Exceptions" subtitle="Requested changes that did not appear on Google after the maximum number of verification attempts. This is the only list Ops needs to work from."
        right={<button onClick={() => onExport("exceptions.csv", exceptions.map(exportRow))} className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-[12.5px] font-medium border border-slate-200 text-slate-600 hover:bg-slate-50"><Download size={13} /> Export exceptions</button>} />

      {exceptions.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-10 text-center text-slate-400 text-[13px]">No exceptions right now.</div>
      ) : (
        <>
          <div className="text-[12.5px] font-semibold text-slate-500 mb-2">Needs attention ({outstanding.length})</div>
          <div className="space-y-3 mb-6">
            {outstanding.map(r => <ExceptionCard key={r.id} r={r} onView={onView} onReview={onReview} onRetry={onRetry} />)}
          </div>
          {reviewed.length > 0 && (
            <>
              <div className="text-[12.5px] font-semibold text-slate-500 mb-2">Reviewed ({reviewed.length})</div>
              <div className="space-y-3 opacity-70">
                {reviewed.map(r => <ExceptionCard key={r.id} r={r} onView={onView} onReview={onReview} onRetry={onRetry} />)}
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}

function ExceptionCard({ r, onView, onReview, onRetry }) {
  const first = r.checks.find(c => c.n === 1);
  const last = r.checks[r.checks.length - 1];
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4">
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div>
          <div className="flex items-center gap-2 mb-1.5">
            <StatusBadge status="RED" size="sm" />
            {r.reviewed && <span className="text-[11px] text-slate-400 font-medium">Reviewed</span>}
          </div>
          <div className="text-[14.5px] font-semibold text-slate-800">{r.branch} Branch <span className="text-slate-400 font-normal">— {r.client}</span></div>
          <div className="text-[12.5px] text-slate-500 mt-1 flex items-center gap-1 flex-wrap">
            Change: <b className="text-slate-700">{r.changeType}</b> · Store Code: <span className="font-mono font-semibold text-slate-700">{r.storeCode && r.storeCode !== "-" ? r.storeCode : "Not provided"}</span>
            <CopyButton value={r.storeCode} label="Store Code" />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-x-8 gap-y-1 text-[12.5px] text-right">
          <div className="text-slate-400">Expected</div><div className="font-medium text-slate-700">{r.expectedValue}</div>
          <div className="text-slate-400">Current GBP</div><div className="font-medium text-slate-700">{r.currentGbpValue ?? "-"}</div>
          <div className="text-slate-400">Checks</div><div className="font-medium text-slate-700">{r.attempts}</div>
          <div className="text-slate-400">First check</div><div className="text-slate-600">{fmtDateTime(first?.time)}</div>
          <div className="text-slate-400">Last check</div><div className="text-slate-600">{fmtDateTime(last?.time)}</div>
        </div>
      </div>
      <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between flex-wrap gap-2">
        <div className="text-[12px] text-slate-500 max-w-md"><b className="text-slate-600">Recommended action:</b> Review GBP status and confirm whether the requested change was successfully submitted through SI.</div>
        <div className="flex gap-2 flex-wrap">
          {!r.reviewed && <button onClick={() => onReview(r.id)} className="px-2.5 py-1 rounded-md border border-slate-200 text-[12px] font-medium text-slate-600 hover:bg-slate-50">Mark as Reviewed</button>}
          <button onClick={() => onRetry(r.id)} className="px-2.5 py-1 rounded-md border border-slate-200 text-[12px] font-medium text-slate-600 hover:bg-slate-50">Retry Verification</button>
          <button onClick={() => onView(r.id)} className="px-2.5 py-1 rounded-md bg-indigo-600 text-white text-[12px] font-semibold hover:bg-indigo-700">Open</button>
        </div>
      </div>
    </div>
  );
}

function ClientsPage({ records, clients, selectedClient, setSelectedClient }) {
  const rows = records.filter(r => r.client === selectedClient);
  const total = rows.length;
  const verified = rows.filter(r => r.status === "GREEN").length;
  const pending = rows.filter(r => r.status === "YELLOW").length;
  const exceptions = rows.filter(r => r.status === "RED").length;
  const health = total ? Math.round((verified / total) * 100) : 0;

  const verifiedDurations = rows.filter(r => r.status === "GREEN" && r.verifiedAt).map(r => minutesBetween(new Date(r.createdAt), new Date(r.verifiedAt)));
  const avgMin = verifiedDurations.length ? Math.round(verifiedDurations.reduce((a, b) => a + b, 0) / verifiedDurations.length) : null;

  const breakdown = CHANGE_TYPES.map(ct => {
    const sub = rows.filter(r => r.changeType === ct);
    const v = sub.filter(r => r.status === "GREEN").length;
    return { ct, pct: sub.length ? Math.round((v / sub.length) * 100) : null, count: sub.length };
  });

  return (
    <div>
      <PageHeader title="Clients" subtitle="Per-client verification health, drawn from actual recorded verification data." />
      <div className="flex gap-2 mb-5 flex-wrap">
        {clients.map(c => (
          <button key={c} onClick={() => setSelectedClient(c)}
            className={`px-3 py-1.5 rounded-full text-[12.5px] font-medium border ${selectedClient === c ? "bg-indigo-600 text-white border-indigo-600" : "bg-white border-slate-200 text-slate-600 hover:bg-slate-50"}`}>
            {c}
          </button>
        ))}
      </div>

      {total === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 p-10 text-center text-slate-400 text-[13px]">No verification records for this client yet.</div>
      ) : (
        <>
          <div className="grid md:grid-cols-2 gap-4 mb-5">
            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <div className="text-[12px] text-slate-400 font-medium uppercase tracking-wide mb-1">Verification Health</div>
              <div className="text-3xl font-bold text-slate-900">{health}%</div>
              <div className="w-full h-2 rounded-full bg-slate-100 mt-3 overflow-hidden">
                <div className="h-full bg-emerald-500" style={{ width: `${health}%` }} />
              </div>
              <div className="grid grid-cols-4 gap-2 mt-4 text-center">
                <div><div className="text-[11px] text-slate-400">Monitored</div><div className="font-semibold text-slate-700">{total}</div></div>
                <div><div className="text-[11px] text-slate-400">Verified</div><div className="font-semibold text-emerald-600">{verified}</div></div>
                <div><div className="text-[11px] text-slate-400">Pending</div><div className="font-semibold text-amber-600">{pending}</div></div>
                <div><div className="text-[11px] text-slate-400">Exceptions</div><div className="font-semibold text-red-600">{exceptions}</div></div>
              </div>
              <div className="mt-4 pt-3 border-t border-slate-100 text-[12.5px] text-slate-500">
                Average verification time: <b className="text-slate-700">{avgMin != null ? fmtDuration(avgMin) : "Not enough verified records yet"}</b>
              </div>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <div className="text-[12px] text-slate-400 font-medium uppercase tracking-wide mb-3">Change-type breakdown</div>
              <div className="space-y-3">
                {breakdown.map(b => (
                  <div key={b.ct}>
                    <div className="flex justify-between text-[12.5px] mb-1">
                      <span className="text-slate-600">{b.ct}</span>
                      <span className="text-slate-500">{b.pct != null ? `${b.pct}% verified` : "No data"} <span className="text-slate-300">({b.count})</span></span>
                    </div>
                    <div className="w-full h-1.5 rounded-full bg-slate-100 overflow-hidden">
                      <div className="h-full bg-indigo-500" style={{ width: `${b.pct || 0}%` }} />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function InsightsPage({ records }) {
  if (records.length === 0) {
    return (
      <div>
        <PageHeader title="Insights" subtitle="Calculated only from recorded verification data — nothing here is estimated." />
        <div className="bg-white rounded-xl border border-slate-200 p-10 text-center text-slate-400 text-[13px]">No data yet — run verification on some expected changes first.</div>
      </div>
    );
  }

  const byType = CHANGE_TYPES.map(ct => {
    const sub = records.filter(r => r.changeType === ct);
    const resolved = sub.filter(r => r.status === "GREEN" || r.status === "RED");
    const exceptionRate = resolved.length ? Math.round((sub.filter(r => r.status === "RED").length / resolved.length) * 100) : null;
    const verifiedDurations = sub.filter(r => r.status === "GREEN" && r.verifiedAt).map(r => minutesBetween(new Date(r.createdAt), new Date(r.verifiedAt)));
    const avg = verifiedDurations.length ? Math.round(verifiedDurations.reduce((a, b) => a + b, 0) / verifiedDurations.length) : null;
    return { ct, exceptionRate, avg, count: sub.length };
  }).filter(b => b.count > 0);

  const worstException = [...byType].filter(b => b.exceptionRate != null).sort((a, b) => b.exceptionRate - a.exceptionRate)[0];
  const bestException = [...byType].filter(b => b.exceptionRate != null).sort((a, b) => a.exceptionRate - b.exceptionRate)[0];

  // repeat-exception locations
  const branchExceptions = {};
  records.forEach(r => { if (r.status === "RED") branchExceptions[`${r.client} — ${r.branch}`] = (branchExceptions[`${r.client} — ${r.branch}`] || 0) + 1; });
  const repeatLocations = Object.entries(branchExceptions).filter(([, n]) => n > 1).sort((a, b) => b[1] - a[1]);

  // client exception rates
  const clients = Array.from(new Set(records.map(r => r.client)));
  const clientRates = clients.map(c => {
    const sub = records.filter(r => r.client === c);
    const resolved = sub.filter(r => r.status === "GREEN" || r.status === "RED");
    const rate = resolved.length ? Math.round((sub.filter(r => r.status === "RED").length / resolved.length) * 100) : null;
    return { c, rate, count: sub.length };
  }).filter(c => c.rate != null);
  const avgClientRate = clientRates.length ? clientRates.reduce((a, b) => a + b.rate, 0) / clientRates.length : 0;
  const highClients = clientRates.filter(c => c.rate > avgClientRate + 10);

  const pendingCount = records.filter(r => r.status === "YELLOW").length;
  const exceptionCount = records.filter(r => r.status === "RED").length;

  return (
    <div>
      <PageHeader title="Insights" subtitle="Calculated only from recorded verification data — nothing here is estimated." />

      <div className="grid md:grid-cols-3 gap-3 mb-5">
        <MiniStat label="Pending changes" value={pendingCount} tone="YELLOW" />
        <MiniStat label="Open exceptions" value={exceptionCount} tone="RED" />
        <MiniStat label="Change types with data" value={byType.length} tone="slate" />
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5 mb-4">
        <div className="text-[13px] font-semibold text-slate-700 mb-3">Exception rate by change type</div>
        <div className="space-y-3">
          {byType.map(b => (
            <div key={b.ct}>
              <div className="flex justify-between text-[12.5px] mb-1">
                <span className="text-slate-600">{b.ct}</span>
                <span className="text-slate-500">{b.exceptionRate != null ? `${b.exceptionRate}% exception rate` : "Not enough resolved records"} · avg verify time {b.avg != null ? fmtDuration(b.avg) : "-"}</span>
              </div>
              <div className="w-full h-1.5 rounded-full bg-slate-100 overflow-hidden">
                <div className="h-full bg-red-400" style={{ width: `${b.exceptionRate || 0}%` }} />
              </div>
            </div>
          ))}
        </div>
        {worstException && bestException && worstException.ct !== bestException.ct && (
          <div className="mt-4 pt-3 border-t border-slate-100 text-[12.5px] text-slate-600">
            <Sparkles size={13} className="inline text-indigo-500 mr-1" />
            {worstException.ct} changes have a higher exception rate ({worstException.exceptionRate}%) than {bestException.ct} changes ({bestException.exceptionRate}%) in the current data set.
          </div>
        )}
      </div>

      <div className="grid md:grid-cols-2 gap-4">
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="text-[13px] font-semibold text-slate-700 mb-3">Locations with repeat exceptions</div>
          {repeatLocations.length === 0 ? (
            <div className="text-[12.5px] text-slate-400">No location has more than one exception yet.</div>
          ) : (
            <div className="space-y-2">
              {repeatLocations.map(([loc, n]) => (
                <div key={loc} className="flex justify-between text-[12.5px]"><span className="text-slate-600">{loc}</span><span className="font-semibold text-red-600">{n} exceptions</span></div>
              ))}
            </div>
          )}
        </div>
        <div className="bg-white rounded-xl border border-slate-200 p-5">
          <div className="text-[13px] font-semibold text-slate-700 mb-3">Clients with above-average exception rates</div>
          {highClients.length === 0 ? (
            <div className="text-[12.5px] text-slate-400">No client is meaningfully above average right now.</div>
          ) : (
            <div className="space-y-2">
              {highClients.map(c => (
                <div key={c.c} className="flex justify-between text-[12.5px]"><span className="text-slate-600">{c.c}</span><span className="font-semibold text-red-600">{c.rate}%</span></div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function MiniStat({ label, value, tone }) {
  const meta = STATUS_META[tone];
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4">
      <div className="text-[12px] text-slate-500 font-medium mb-1">{label}</div>
      <div className="text-2xl font-bold" style={{ color: meta ? meta.text : "#0F172A" }}>{value}</div>
    </div>
  );
}

function AuditLogPage({ auditLog, onExport }) {
  return (
    <div>
      <PageHeader title="Audit Log" subtitle="Every verification action, recorded for accountability."
        right={<button onClick={() => onExport("audit_log.csv", auditLog.map(a => ({
          Timestamp: fmtDateTimeFull(a.timestamp), User: a.user, Client: a.client, Location: a.location,
          "Change Type": a.changeType, "Expected Value": a.expectedValue, "GBP Value": a.gbpValue,
          Result: a.result, "Attempt Number": a.attemptNumber, Source: a.source, "Action Taken": a.actionTaken, Notes: a.notes,
        })))} className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-[12.5px] font-medium border border-slate-200 text-slate-600 hover:bg-slate-50"><Download size={13} /> Export log</button>} />
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
          <table className="w-full text-[12px]">
            <thead className="sticky top-0 bg-slate-50">
              <tr className="text-slate-500 text-left border-b border-slate-200">
                {["Timestamp", "User", "Client", "Location", "Change Type", "Expected", "GBP Value", "Result", "Attempt", "Source", "Action Taken", "Notes"].map(h => <th key={h} className="px-3 py-2 font-semibold whitespace-nowrap bg-slate-50">{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {auditLog.map(a => (
                <tr key={a.id} className="border-b border-slate-100">
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500">{fmtDateTime(a.timestamp)}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{a.user}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{a.client}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{a.location}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{a.changeType}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[120px] truncate">{a.expectedValue}</td>
                  <td className="px-3 py-2 whitespace-nowrap max-w-[120px] truncate">{a.gbpValue ?? "-"}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{a.result}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-center">{a.attemptNumber}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500">{a.source}</td>
                  <td className="px-3 py-2 whitespace-nowrap">{a.actionTaken}</td>
                  <td className="px-3 py-2 whitespace-nowrap text-slate-500 max-w-[160px] truncate">{a.notes}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function SettingsPage({ settings, setSettings }) {
  return (
    <div>
      <PageHeader title="Settings" subtitle="Configure the recheck schedule and review the data-provider architecture." />

      <div className="bg-white rounded-xl border border-slate-200 p-5 mb-4 max-w-2xl">
        <div className="text-[13px] font-semibold text-slate-700 mb-3">Automatic recheck schedule</div>
        <div className="text-[12px] text-slate-500 mb-4">These are the real-world intervals a production backend scheduler would use. Attempt 1 fires this long after a change request is created; each following attempt fires this long after the previous one.</div>
        <div className="space-y-2">
          {settings.intervalsMinutes.map((m, i) => (
            <div key={i} className="flex items-center gap-3 text-[12.5px]">
              <span className="w-20 text-slate-500">Check {i + 1}</span>
              <input type="number" min="1" value={m === 60 * 18 ? 1080 : m}
                onChange={e => {
                  const v = Number(e.target.value) || 1;
                  setSettings(s => ({ ...s, intervalsMinutes: s.intervalsMinutes.map((x, xi) => xi === i ? v : x) }));
                }}
                className="w-24 px-2 py-1 rounded-md border border-slate-200" />
              <span className="text-slate-400">minutes ({fmtDuration(m)})</span>
            </div>
          ))}
        </div>
        <div className="flex items-center gap-3 mt-4 pt-4 border-t border-slate-100 text-[12.5px]">
          <span className="w-40 text-slate-500">Max attempts before Exception</span>
          <input type="number" min="1" value={settings.maxAttempts}
            onChange={e => setSettings(s => ({ ...s, maxAttempts: Number(e.target.value) || 1 }))}
            className="w-20 px-2 py-1 rounded-md border border-slate-200" />
        </div>
        <div className="flex items-center gap-3 mt-3 pt-3 border-t border-slate-100 text-[12.5px]">
          <span className="w-40 text-slate-500">Map pin match tolerance</span>
          <input type="number" min="1" value={settings.coordToleranceMeters}
            onChange={e => setSettings(s => ({ ...s, coordToleranceMeters: Number(e.target.value) || 1 }))}
            className="w-20 px-2 py-1 rounded-md border border-slate-200" />
          <span className="text-slate-400">meters — GPS/geocoding precision means pins are rarely byte-identical</span>
        </div>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5 mb-4 max-w-2xl">
        <div className="text-[13px] font-semibold text-slate-700 mb-2">Live demo auto-check</div>
        <div className="text-[12px] text-slate-500 mb-3 leading-relaxed">
          This browser prototype has no backend job scheduler, so it cannot check Google in the background once the tab is closed. With this toggle on, pending records recheck themselves on a <b>compressed</b> timer (~10–35 seconds) while this tab stays open, so you can watch the pipeline work. Turn it off to freeze pending records exactly where they are.
        </div>
        <label className="flex items-center gap-2 text-[13px] font-medium text-slate-700">
          <input type="checkbox" checked={settings.demoAutoCheck} onChange={e => setSettings(s => ({ ...s, demoAutoCheck: e.target.checked }))} />
          Enable live demo auto-check (this tab only)
        </label>
      </div>

      <div className="bg-white rounded-xl border border-slate-200 p-5 max-w-2xl">
        <div className="text-[13px] font-semibold text-slate-700 mb-2">GBP data provider</div>
        <div className="text-[12.5px] text-slate-600 space-y-2 leading-relaxed">
          <p>The application reads GBP location data through a single interface, <code className="bg-slate-100 px-1 rounded">GBPDataProvider</code>, with three methods: <code className="bg-slate-100 px-1 rounded">getLocations()</code>, <code className="bg-slate-100 px-1 rounded">getLocation(id)</code>, and <code className="bg-slate-100 px-1 rounded">getLocationField(id, field)</code>.</p>
          <p><b>Active today:</b> <code className="bg-slate-100 px-1 rounded">FileUploadGBPProvider</code> — reads from the file you upload on the GBP Data page.</p>
          <p><b>Not connected:</b> a future <code className="bg-slate-100 px-1 rounded">GoogleBusinessProfileAPIProvider</code> would implement the same three methods against the live GBP API, so matching, comparison, the dashboard, and every other module would work unchanged. This requires GBP API credentials and access that this prototype does not have — nothing in this tool pretends that connection exists.</p>
        </div>
      </div>
    </div>
  );
}

function HelpPage() {
  const items = [
    { t: "How to upload expected changes", b: "Go to Expected Changes, then upload a CSV or Excel export from SI / V1 / V3 (or download the template first). Each row needs at minimum a Change Type, Old Value, Expected New Value, and enough identifying information to find the right location — ideally a GBP Location ID or Store Code." },
    { t: "How to upload GBP data", b: "Go to GBP Data and upload a Google Business Profile export or equivalent snapshot. This is the current-state data the comparison engine checks against. Re-upload periodically to refresh the snapshot." },
    { t: "Which fields are supported", b: "Business Name, Address, Phone, Business Hours, Map Pin (Lat/Long), Category, Website, WhatsApp Number, Services, Products, Photos, Service Area, and Description. If your GBP Data upload includes a column for a field, it auto-verifies like any other change. If it doesn't — common for Photos, Services, Products, Description, Service Area, and Map Pin, since GBP doesn't offer a simple bulk export of these — the record waits as Pending with a manual \"Confirmed live\" / \"Still not correct\" check instead of guessing or forcing you to hunt down data that isn't easily exportable." },
    { t: "How matching works", b: "Each expected change is matched to a GBP location using a hierarchy: GBP Location ID → Store Code → Internal Location ID → Phone → Business Name + Address. The method used is always shown. If no method reaches a confident match, the record is marked GRAY — Matching Required rather than risk comparing the wrong branch." },
    { t: "What Verified (GREEN) means", b: "The current GBP value, once normalized, matches the expected value. This is only ever shown when the actual uploaded/checked GBP data contains the expected value — never assumed." },
    { t: "What Pending (YELLOW) means", b: "Either the expected value hasn't yet appeared on Google and it's still within the allowed recheck attempts, or the field needs a manual visual check (see \"Which fields are supported\") and is waiting on you." },
    { t: "What Exception (RED) means", b: "The expected value still did not appear after the maximum number of recheck attempts, or you manually confirmed it's still wrong. This is the list Ops should work from — it does not mean SI is wrong, only that the change is not yet reflected on Google and needs a look." },
    { t: "How automatic rechecking works", b: "In production, a backend job scheduler would recheck each pending record at the intervals configured in Settings (default: 30 min, 2h, 6h, next business day) until it verifies or exhausts the max attempts. This browser prototype has no backend, so it simulates the same pipeline on a compressed timer while the tab is open (Settings → Live demo auto-check). You can also trigger a check manually with Retry Verification. Records waiting on a manual check never auto-tick — they wait for you." },
    { t: "How to export the report", b: "Use the Export button in the top bar for the full report, or the Export buttons on the Exceptions and Audit Log pages for a scoped export. All exports download as CSV." },
  ];
  return (
    <div>
      <PageHeader title="Help" subtitle="What this tool does, and how it fits alongside SI." />
      <div className="bg-indigo-50 border border-indigo-100 rounded-xl p-4 mb-5 text-[13px] text-indigo-900 leading-relaxed">
        This solution complements the existing SI Out of Sync functionality. It is designed specifically for operational verification of requested GBP changes and exception monitoring — it does not sync, replace, or second-guess SI's data.
      </div>
      <div className="space-y-3">
        {items.map((it, i) => (
          <details key={i} className="bg-white rounded-xl border border-slate-200 p-4 group" open={i === 0}>
            <summary className="text-[13.5px] font-semibold text-slate-800 cursor-pointer list-none flex items-center justify-between">
              {it.t} <ChevronDown size={15} className="text-slate-400 group-open:rotate-180 transition-transform" />
            </summary>
            <p className="text-[12.5px] text-slate-600 mt-2 leading-relaxed">{it.b}</p>
          </details>
        ))}
      </div>
    </div>
  );
}

function DetailPage({ record: r, onBack, onRetry, onReview, onNote, onAssign, onSimulateGoogleUpdate, onManualResolve, settings }) {
  const [noteText, setNoteText] = useState("");
  const [assignee, setAssignee] = useState("");
  const isCoord = r.changeType === "Map Pin (Lat/Long)";
  const normalizer = NORMALIZER[r.changeType];
  const rawMismatch = r.currentGbpValue != null && !valuesMatch(r.changeType, r.currentGbpValue, r.expectedValue, settings);
  const coordDistance = isCoord && r.currentGbpValue != null ? haversineMeters(parseLatLng(r.currentGbpValue), parseLatLng(r.expectedValue)) : null;
  const currentValueDisplay = r.currentGbpValue ?? (r.requiresManualCheck ? "Requires manual visual check (Street View)" : (r.status === "GRAY" ? "Waiting for GBP data" : (r.isDemo ? "Demo Data" : "Waiting for GBP data")));

  return (
    <div className="max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1.5 text-[12.5px] text-slate-500 hover:text-slate-800 mb-4">
        <ArrowLeft size={14} /> Back
      </button>

      <div className="flex items-start justify-between flex-wrap gap-3 mb-1">
        <div>
          <div className="text-[12px] text-slate-400 font-medium">{r.client} — {r.branch}</div>
          <h1 className="text-[20px] font-bold text-slate-900">{r.changeType} Verification</h1>
          <div className="mt-1.5 flex items-center gap-1.5 text-[12px] text-slate-500">
            <span>Store Code:</span>
            <span className="font-mono font-semibold text-slate-700 bg-slate-100 px-1.5 py-0.5 rounded">{r.storeCode && r.storeCode !== "-" ? r.storeCode : "Not provided"}</span>
            <CopyButton value={r.storeCode} label="Store Code" />
            {r.storeCode && r.storeCode !== "-" && <span className="text-slate-400">— use this to find the listing directly in GBP for a manual cross-check</span>}
          </div>
        </div>
        <StatusBadge status={r.status} />
      </div>
      {r.unexpectedChange && (
        <div className="mt-3 p-3 rounded-lg bg-amber-50 border border-amber-200 text-[12.5px] text-amber-800 flex gap-2">
          <AlertTriangle size={15} className="shrink-0 mt-0.5" />
          <div>
            <b>Unexpected change detected.</b> Expected: {r.changeType} update {r.status === "GREEN" ? "✓" : ""}. Detected: <b>{r.unexpectedChange.field}</b> changed from "{r.unexpectedChange.from}" to "{r.unexpectedChange.to}". This is shown for review — it is not automatically classified as an error.
          </div>
        </div>
      )}

      <div className="grid sm:grid-cols-3 gap-3 my-5">
        <ValueCard label="Old Value" value={r.oldValue} />
        <ValueCard label="Expected Value" value={r.expectedValue} />
        <ValueCard label="Current GBP Value" value={currentValueDisplay} highlight />
      </div>

      {rawMismatch && isCoord && (
        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-5">
          <div className="text-[12px] font-semibold text-slate-500 uppercase tracking-wide mb-2">Pin distance</div>
          <div className="text-[13px] text-slate-700">
            The current GBP pin is <b>{Number.isFinite(coordDistance) ? `${Math.round(coordDistance)} meters` : "an unreadable distance"}</b> from the expected pin — outside the {settings.coordToleranceMeters}m match tolerance set in Settings.
          </div>
          <div className="grid grid-cols-2 gap-4 text-[12.5px] mt-3">
            <div className="text-slate-700">Expected: <span className="font-mono">{r.expectedValue}</span></div>
            <div className="text-slate-700">GBP: <span className="font-mono">{r.currentGbpValue}</span></div>
          </div>
        </div>
      )}
      {rawMismatch && !isCoord && (
        <div className="bg-white rounded-xl border border-slate-200 p-4 mb-5">
          <div className="text-[12px] font-semibold text-slate-500 uppercase tracking-wide mb-2">Raw vs. normalized comparison</div>
          <div className="grid grid-cols-2 gap-4 text-[12.5px]">
            <div>
              <div className="text-slate-400 mb-1">Raw</div>
              <div className="text-slate-700">Expected: <span className="font-mono">{r.expectedValue}</span></div>
              <div className="text-slate-700">GBP: <span className="font-mono">{r.currentGbpValue}</span></div>
            </div>
            <div>
              <div className="text-slate-400 mb-1">Normalized</div>
              <div className="text-slate-700">Expected: <span className="font-mono">{normalizer(r.expectedValue)}</span></div>
              <div className="text-slate-700">GBP: <span className="font-mono">{normalizer(r.currentGbpValue)}</span></div>
            </div>
          </div>
        </div>
      )}

      <div className="grid sm:grid-cols-2 gap-4 mb-5">
        <div className="bg-white rounded-xl border border-slate-200 p-4">
          <div className="text-[12px] font-semibold text-slate-500 uppercase tracking-wide mb-2">Match</div>
          <div className="text-[13px] text-slate-700">Matched using: <b>{r.matchMethod}</b></div>
          <div className="text-[12px] text-slate-400 mt-1">Confidence: {r.matchConfidence}</div>
        </div>
        <div className="bg-white rounded-xl border border-slate-200 p-4">
          <div className="text-[12px] font-semibold text-slate-500 uppercase tracking-wide mb-2">Verification source</div>
          <div className="text-[13px] text-slate-700">{r.source} {r.isDemo && <span className="text-amber-600 font-medium">(Demo Data)</span>}</div>
          {r.status === "GREEN" && r.verifiedAt && <div className="text-[12px] text-slate-400 mt-1">Verified at {fmtDateTimeFull(r.verifiedAt)}</div>}
        </div>
      </div>

      {/* TIMELINE */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-5">
        <div className="text-[12px] font-semibold text-slate-500 uppercase tracking-wide mb-3">Verification timeline</div>
        <div className="space-y-0">
          {r.checks.map((c, i) => (
            <div key={i} className="flex gap-3 relative">
              <div className="flex flex-col items-center">
                <div className={`w-2.5 h-2.5 rounded-full mt-1 ${c.result === "VERIFIED" ? "bg-emerald-500" : c.result === "EXCEPTION" ? "bg-red-500" : c.result === "PENDING" ? "bg-amber-500" : "bg-slate-300"}`} />
                {i < r.checks.length - 1 && <div className="w-px flex-1 bg-slate-200" />}
              </div>
              <div className="pb-4">
                <div className="text-[12px] text-slate-400">{fmtDateTime(c.time)}</div>
                <div className="text-[13px] text-slate-700 font-medium">{c.n === 0 ? c.event : `Check #${c.n}`}</div>
                {c.n > 0 && (
                  <div className="text-[12.5px] text-slate-500 mt-0.5">
                    GBP {c.result === "VERIFIED" ? "shows expected value" : "still shows previous value"}: <span className="font-mono">{c.gbpValue}</span> — <span className={c.result === "VERIFIED" ? "text-emerald-600 font-semibold" : c.result === "EXCEPTION" ? "text-red-600 font-semibold" : "text-amber-600 font-semibold"}>{c.result}</span>
                  </div>
                )}
              </div>
            </div>
          ))}
          {r.status === "YELLOW" && !r.requiresManualCheck && (
            <div className="text-[12px] text-slate-400 pl-5 ml-0.5">
              {settings.demoAutoCheck ? "Next check scheduled — simulated timing (see Settings)." : "Auto-check paused (Settings → Live demo auto-check is off)."}
            </div>
          )}
        </div>
      </div>

      {(r.status === "YELLOW" || r.status === "RED") && (
        <div className="bg-indigo-50 border border-indigo-200 rounded-xl p-4 mb-5">
          <div className="flex items-start gap-2">
            <Info size={16} className="text-indigo-500 shrink-0 mt-0.5" />
            <div className="text-[12.5px] text-indigo-900">
              {r.requiresManualCheck
                ? <><b>This one needs a manual check.</b> {manualCheckHint(r.changeType)} Then record what you found below.</>
                : <><b>Already checked this yourself on Google?</b> The automated check only compares against your uploaded master data — if that's stale, confirm what you actually saw directly and this record (and your master data) will update immediately instead of waiting on a fresh file upload.</>}
            </div>
          </div>
          <div className="flex gap-2 mt-3">
            <button onClick={() => onManualResolve(r.id, true)} className="px-3.5 py-1.5 rounded-md bg-emerald-600 text-white text-[12.5px] font-semibold hover:bg-emerald-700">
              ✓ Confirmed live on Google
            </button>
            <button onClick={() => onManualResolve(r.id, false)} className="px-3.5 py-1.5 rounded-md bg-red-600 text-white text-[12.5px] font-semibold hover:bg-red-700">
              ✕ Still not correct
            </button>
          </div>
        </div>
      )}

      {/* ACTIONS */}
      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-5">
        <div className="text-[12px] font-semibold text-slate-500 uppercase tracking-wide mb-3">Actions</div>
        <div className="flex flex-wrap gap-2 mb-3">
          {r.status === "RED" && !r.reviewed && <button onClick={() => onReview(r.id)} className="px-3 py-1.5 rounded-md border border-slate-200 text-[12.5px] font-medium text-slate-600 hover:bg-slate-50">Mark as Reviewed</button>}
          {(r.status === "RED" || r.status === "YELLOW") && !r.requiresManualCheck && <button onClick={() => onRetry(r.id)} className="px-3 py-1.5 rounded-md border border-slate-200 text-[12.5px] font-medium text-slate-600 hover:bg-slate-50 inline-flex items-center gap-1"><RefreshCw size={13} /> Retry Verification</button>}
          {r.status === "YELLOW" && r.matchedLocationId && !r.requiresManualCheck && (
            <button onClick={() => onSimulateGoogleUpdate(r.id)} className="px-3 py-1.5 rounded-md border border-dashed border-indigo-300 text-[12.5px] font-medium text-indigo-600 hover:bg-indigo-50">
              Demo: Simulate Google update
            </button>
          )}
          <button onClick={() => downloadCSV(`record_${r.id}.csv`, [exportRow(r)])} className="px-3 py-1.5 rounded-md border border-slate-200 text-[12.5px] font-medium text-slate-600 hover:bg-slate-50 inline-flex items-center gap-1"><Download size={13} /> Export</button>
        </div>

        <div className="flex gap-2 mb-3">
          <input value={assignee} onChange={e => setAssignee(e.target.value)} placeholder="Assign to team member…"
            className="flex-1 px-3 py-1.5 rounded-md border border-slate-200 text-[12.5px]" />
          <button onClick={() => { if (assignee.trim()) { onAssign(r.id, assignee); setAssignee(""); } }}
            className="px-3 py-1.5 rounded-md border border-slate-200 text-[12.5px] font-medium text-slate-600 hover:bg-slate-50">Assign</button>
        </div>
        {r.assignedTo && <div className="text-[12px] text-slate-500 mb-3">Currently assigned to <b>{r.assignedTo}</b></div>}

        <div className="flex gap-2">
          <input value={noteText} onChange={e => setNoteText(e.target.value)} placeholder="Add a note…"
            className="flex-1 px-3 py-1.5 rounded-md border border-slate-200 text-[12.5px]" />
          <button onClick={() => { onNote(r.id, noteText); setNoteText(""); }}
            className="px-3 py-1.5 rounded-md bg-indigo-600 text-white text-[12.5px] font-semibold hover:bg-indigo-700 inline-flex items-center gap-1"><PlusCircle size={13} /> Add Note</button>
        </div>
        {(r.notes || []).length > 0 && (
          <div className="mt-3 space-y-2">
            {r.notes.map((n, i) => (
              <div key={i} className="text-[12.5px] bg-slate-50 rounded-md p-2.5 border border-slate-100">
                <div className="text-slate-400 text-[11px] mb-0.5">{fmtDateTime(n.time)} · {n.author}</div>
                <div className="text-slate-700">{n.text}</div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function ValueCard({ label, value, highlight }) {
  return (
    <div className={`rounded-xl border p-4 ${highlight ? "border-indigo-200 bg-indigo-50/40" : "border-slate-200 bg-white"}`}>
      <div className="text-[11px] text-slate-400 font-semibold uppercase tracking-wide mb-1">{label}</div>
      <div className="text-[14px] font-semibold text-slate-800 break-words">{value || "-"}</div>
    </div>
  );
}
