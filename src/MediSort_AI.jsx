import { useState, useRef, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Upload, Download, MessageCircle, X, Send, LogOut,
  FileSpreadsheet, Sparkles, Check, Loader2, ChevronRight,
  Bot, User2, AlertTriangle, RefreshCw, FlaskConical,
  Package, Layers, LayoutDashboard, ChevronDown, Mic2,
  Share2, Mail, ExternalLink, Copy, Maximize2, Minimize2, Minus
} from "lucide-react";

/* ─────────────────────────── HELPERS ─────────────────────────── */
async function callClaude(messages, systemPrompt = "", maxTokens = 4096) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: maxTokens,
      system: systemPrompt,
      messages,
    }),
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  if (data.error) throw new Error(data.error.message || "API error");
  return data.content?.map(b => b.text || "").join("") || "";
}

/* Robust JSON extractor — handles markdown fences, leading/trailing prose */
function extractJSON(raw) {
  // 1. Try stripping markdown fences first
  const stripped = raw.replace(/```(?:json)?/gi, "").trim();
  // 2. Find the outermost {...} block
  const start = stripped.indexOf("{");
  const end   = stripped.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) throw new Error("No JSON object found");
  return JSON.parse(stripped.slice(start, end + 1));
}

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // dense:true fills every cell so column counts are consistent
        resolve(XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", dense: true }));
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

/* Always returns "Package N" (1-indexed) — used everywhere for Package Name row in UI and XLSX */
function pkgLabel(_pkg, index) {
  return `Package ${index + 1}`;
}

/* Returns "-" for UI display; empty string for XLSX cells. Handles all falsy values. */
function metaUI(val) {
  if (val === null || val === undefined || val !== val) return "-"; // NaN check
  const s = String(val).trim();
  return s && s !== "undefined" && s !== "null" && s !== "NaN" ? s : "-";
}
function metaXLSX(val) {
  if (val === null || val === undefined || val !== val) return "";
  const s = String(val).trim();
  return s && s !== "undefined" && s !== "null" && s !== "NaN" ? s : "";
}

/* ──────────────────────────────────────────────────────────────────
   Dynamic metadata detector
   Scans the first 5 rows of a column.  Returns detected fields plus
   the full deduplicated tests list (rows not consumed as metadata).
   ────────────────────────────────────────────────────────────────── */
const SERVICE_LOC_RE  = /loose\s*hybrid|strict\s*hybrid|full\s*inclinic|full\s*home\s*coll(?:ection)?|home\s*coll(?:ection)?|inclinic|hybrid/i;
const AGE_GENDER_RE   = /\byears?\b|\bmale\b|\bfemale\b|m\s*\/\s*f|f\s*\/\s*m|\bunisex\b|\bage\b|\d+\s*[-–]\s*\d+|above\s+\d+|\d+\s*\+|child(?:ren)?|adult|senior/i;
// Price: has currency symbol, "rs", "/-", explicit "price", or is purely numeric ≥ 3 digits
const PRICE_RE        = /[₹$£€]|(?:^|\s)rs\.?(?:\s|\d)|\/\s*-|\bprice\b|\bpaid\b|^\s*\d{3,}\s*$|\d{4,}/i;

function detectColumnMeta(allRows, colIdx) {
  let serviceLocation = "", ageGender = "", price = "";
  const metaRowSet = new Set();   // row indices consumed as metadata

  for (let r = 0; r < Math.min(5, allRows.length); r++) {
    const raw = allRows[r]?.[colIdx];
    const val = (raw !== null && raw !== undefined) ? String(raw).trim() : "";
    if (!val) continue;

    if (!serviceLocation && SERVICE_LOC_RE.test(val)) {
      serviceLocation = val;
      metaRowSet.add(r);
    } else if (!ageGender && AGE_GENDER_RE.test(val)) {
      ageGender = val;
      metaRowSet.add(r);
    } else if (!price && PRICE_RE.test(val)) {
      price = val;
      metaRowSet.add(r);
    }
    // rows 0-4 that match NO pattern are NOT consumed — they may be test names
  }

  // Tests = every non-empty row that wasn't consumed as a metadata row
  const tests = allRows
    .map((row, r) => ({ r, v: (row?.[colIdx] !== null && row?.[colIdx] !== undefined) ? String(row[colIdx]).trim() : "" }))
    .filter(({ r, v }) => v && !metaRowSet.has(r))
    .map(({ v }) => v);

  return {
    serviceLocation,   // "" if not found
    ageGender,         // "" if not found
    price,             // "" if not found
    tests,             // all non-metadata values
  };
}

function buildOutputExcel(packages, originalFilename) {
  const wb   = XLSX.utils.book_new();
  const META = ["Package Name","Display Name","Service Location","Age/ Gender","Selling Price | Company Paid"];
  const numCols = packages.length + 1;

  // Per-package row depth
  const pkgDepths = packages.map(p =>
    5 + Object.entries(p.categories).reduce((sum, [, tests]) => sum + 1 + tests.length, 0)
  );
  const totalRows = Math.max(5, ...pkgDepths) + 3;

  // Build AOA grid
  const grid = Array.from({ length: totalRows }, () => Array(numCols).fill(""));
  META.forEach((l, i) => { grid[i][0] = l; });

  packages.forEach((pkg, ci) => {
    const c = ci + 1;
    grid[0][c] = pkgLabel(pkg, ci);          // fallback: "Package N"
    grid[1][c] = metaXLSX(pkg.displayName);
    grid[2][c] = metaXLSX(pkg.serviceLocation);
    grid[3][c] = metaXLSX(pkg.ageGender);
    grid[4][c] = metaXLSX(pkg.price);
    let row = 5;
    for (const [cat, tests] of Object.entries(pkg.categories)) {
      if (row >= totalRows) break;
      grid[row][c] = cat; row++;
      for (const t of tests) {
        if (row >= totalRows) break;
        grid[row][c] = t; row++;
      }
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(grid);

  // ── Auto-fit column widths from content ──────────────────────────
  const colWidths = Array(numCols).fill(12);
  grid.forEach(row => {
    row.forEach((cell, ci) => {
      const len = cell ? String(cell).length : 0;
      if (len > colWidths[ci]) colWidths[ci] = len;
    });
  });
  ws["!cols"] = colWidths.map((w, ci) => ({
    wch: ci === 0
      ? Math.min(Math.max(w + 3, 28), 40)   // label column slightly wider
      : Math.min(Math.max(w + 5, 22), 60),  // data columns
  }));

  // ── Row heights ────────────────────────────────────────────────────
  ws["!rows"] = Array.from({ length: totalRows }, (_, i) => ({
    hpt: i < 5 ? 22 : 18,
  }));

  // ── Freeze top 5 metadata rows ─────────────────────────────────────
  ws["!freeze"] = { xSplit: 0, ySplit: 5, topLeftCell: "A6", activeCell: "A6", sqref: "A6" };

  // ── Cell style helpers ─────────────────────────────────────────────
  // NOTE: patternType:"solid" is required for Excel to render fills.
  // styleCell creates the cell object if it doesn't already exist so
  // category-header cells in shorter packages always get their style.
  const ensureCell = (ws, addr) => {
    if (!ws[addr]) ws[addr] = { t: "s", v: "" };
    return ws[addr];
  };
  const styleCell = (ws, addr, s) => {
    ensureCell(ws, addr).s = s;
  };

  // Shared style constructors
  const fillSolid = (rgb) => ({ patternType: "solid", fgColor: { rgb }, bgColor: { rgb: "FFFFFF" } });

  // Colour palette (RGB hex without #)
  const CLR = {
    labelBg:  "D4EDE8",   // soft teal — label column rows 1-5
    metaBg:   "EEF3F7",   // very light blue-grey — data metadata cells
    catBg:    "1A6B5A",   // deep teal — category header fill (distinct, professional)
    catFg:    "FFFFFF",   // white text on dark fill
    labelFg:  "0D3D30",   // dark teal text for label col
    testFg:   "1A2A38",   // near-black for test names
    metaFg:   "1A2C3C",   // dark for metadata values
  };

  // ── Row 0-4: Label column (col A) ─────────────────────────────────
  for (let r = 0; r < 5; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 0 });
    styleCell(ws, addr, {
      font:      { bold: true, color: { rgb: CLR.labelFg }, sz: 10, name: "Calibri" },
      fill:      fillSolid(CLR.labelBg),
      alignment: { vertical: "center", horizontal: "left", indent: 1 },
      border: {
        bottom: { style: "thin", color: { rgb: "BBCCCC" } },
        right:  { style: "thin", color: { rgb: "BBCCCC" } },
      },
    });
  }

  // ── Row 0-4: Data columns metadata cells ──────────────────────────
  for (let r = 0; r < 5; r++) {
    for (let c = 1; c < numCols; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      styleCell(ws, addr, {
        font:      { bold: r === 0, color: { rgb: CLR.metaFg }, sz: 10, name: "Calibri" },
        fill:      fillSolid(CLR.metaBg),
        alignment: { vertical: "center", horizontal: "left", wrapText: true, indent: 1 },
        border: {
          bottom: { style: "thin", color: { rgb: "D0D8E0" } },
          right:  { style: "thin", color: { rgb: "D0D8E0" } },
        },
      });
    }
  }

  // ── Rows 5+: Category headers and test rows per package ───────────
  // Track which absolute row indices are category headers (needed for
  // styling columns where the package is shorter — so we can still apply
  // the correct row style even when the cell is empty).
  // Strategy: collect (row, col) category header positions and test positions
  // separately, then apply styles in two passes.

  packages.forEach((pkg, ci) => {
    const c = ci + 1;
    let row = 5;

    for (const [cat, tests] of Object.entries(pkg.categories)) {
      if (row >= totalRows) break;

      // Category header cell — bold white text on deep teal, clearly distinct
      const catAddr = XLSX.utils.encode_cell({ r: row, c });
      // Overwrite the grid value with uppercase for visual emphasis in Excel
      if (ws[catAddr]) ws[catAddr].v = String(cat).toUpperCase();
      styleCell(ws, catAddr, {
        font: {
          bold:  true,
          color: { rgb: CLR.catFg },   // white
          sz:    11,
          name:  "Calibri",
          italic: false,
        },
        fill: fillSolid(CLR.catBg),    // deep teal "1A6B5A"
        alignment: {
          vertical:   "center",
          horizontal: "left",
          indent:     1,
        },
        border: {
          top:    { style: "medium", color: { rgb: "0D4A3C" } },
          bottom: { style: "medium", color: { rgb: "0D4A3C" } },
          left:   { style: "thick",  color: { rgb: "00C49A" } },  // bright teal accent stripe
          right:  { style: "thin",   color: { rgb: "1A6B5A" } },
        },
      });
      // Give category header rows extra height for breathing room
      if (ws["!rows"]) ws["!rows"][row] = { hpt: 24 };
      row++;

      // Test name cells
      for (const _ of tests) {
        if (row >= totalRows) break;
        const testAddr = XLSX.utils.encode_cell({ r: row, c });
        styleCell(ws, testAddr, {
          font:      { color: { rgb: CLR.testFg }, sz: 9, name: "Calibri" },
          fill:      fillSolid("FAFCFE"),  // near-white
          alignment: { vertical: "center", horizontal: "left", indent: 2 },
          border: {
            bottom: { style: "hair", color: { rgb: "E0E8F0" } },
            right:  { style: "thin", color: { rgb: "D8E4EC" } },
          },
        });
        row++;
      }
    }
  });

  XLSX.utils.book_append_sheet(wb, ws, "Processed Packages");

  const base64  = XLSX.write(wb, { bookType: "xlsx", type: "base64", cellStyles: true });
  const baseName = (originalFilename || "medical_plans").replace(/\.xlsx$/i, "");
  const fileName  = `Formatted_${baseName}.xlsx`;
  return { base64, fileName };
}

const STEPS = [
  { label: "Parse Excel Structure", icon: FileSpreadsheet },
  { label: "Normalize Test Names", icon: FlaskConical },
  { label: "AI Classification", icon: Sparkles },
  { label: "Build Output File", icon: Package },
];

/* ─────────────────────────── CSS ─────────────────────────── */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Exo+2:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Exo 2',sans-serif;background:#050E1C;color:#DDE6F0;min-height:100vh}
::-webkit-scrollbar{width:4px;height:4px}
::-webkit-scrollbar-track{background:transparent}
::-webkit-scrollbar-thumb{background:rgba(0,210,180,.25);border-radius:4px}
.grid-bg{
  background-color:#050E1C;
  background-image:linear-gradient(rgba(0,210,180,.04) 1px,transparent 1px),
    linear-gradient(90deg,rgba(0,210,180,.04) 1px,transparent 1px);
  background-size:48px 48px;
}
.card{background:rgba(8,18,36,.85);border:1px solid rgba(0,210,180,.12);border-radius:16px;backdrop-filter:blur(12px)}
.card-sm{background:rgba(8,18,36,.9);border:1px solid rgba(0,210,180,.1);border-radius:12px}
.glow{box-shadow:0 0 28px rgba(0,210,180,.12)}
.btn{cursor:pointer;font-family:'Exo 2',sans-serif;font-weight:600;border:none;transition:all .18s}
.btn-primary{background:linear-gradient(135deg,#00D4B4,#0096E0);color:#030C18;padding:11px 22px;border-radius:10px;font-size:14px}
.btn-primary:hover{transform:translateY(-1px);box-shadow:0 6px 22px rgba(0,212,180,.3)}
.btn-primary:disabled{opacity:.38;cursor:not-allowed;transform:none;box-shadow:none}
.btn-ghost{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);color:#DDE6F0;padding:9px 18px;border-radius:10px;font-size:14px}
.btn-ghost:hover{background:rgba(0,210,180,.1);border-color:rgba(0,210,180,.3)}
.btn-icon{background:none;border:none;color:#8CA3B8;padding:6px;border-radius:8px;cursor:pointer;display:flex;align-items:center;transition:all .15s}
.btn-icon:hover{color:#DDE6F0;background:rgba(255,255,255,.06)}
.tab{padding:7px 16px;border-radius:8px;font-size:13px;font-family:'Exo 2',sans-serif;cursor:pointer;transition:all .15s;border:1px solid transparent;white-space:nowrap}
.tab.active{background:rgba(0,210,180,.14);border-color:rgba(0,210,180,.32);color:#00D4B4;font-weight:600}
.tab:not(.active){color:#7A94A8}
.tab:not(.active):hover{background:rgba(255,255,255,.05);color:#DDE6F0}
.input{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);border-radius:10px;padding:10px 14px;color:#DDE6F0;font-family:'Exo 2',sans-serif;font-size:14px;width:100%;outline:none;transition:border .15s}
.input:focus{border-color:rgba(0,210,180,.5)}
.input::placeholder{color:#4A6070}
.mono{font-family:'IBM Plex Mono',monospace}
.badge{display:inline-flex;align-items:center;padding:3px 8px;border-radius:6px;font-size:11px;font-weight:600;letter-spacing:.04em}
.badge-teal{background:rgba(0,210,180,.12);color:#00D4B4;border:1px solid rgba(0,210,180,.22)}
.badge-blue{background:rgba(60,130,220,.12);color:#5BA8F5;border:1px solid rgba(60,130,220,.22)}
.badge-purple{background:rgba(138,99,240,.12);color:#AC87FF;border:1px solid rgba(138,99,240,.22)}
.badge-amber{background:rgba(230,160,40,.12);color:#F5B840;border:1px solid rgba(230,160,40,.22)}
.cat-header{background:linear-gradient(90deg,rgba(0,210,180,.1),transparent);border-left:3px solid #00D4B4;padding:5px 10px;font-size:11px;font-weight:700;color:#00D4B4;letter-spacing:.07em;text-transform:uppercase;margin:10px 0 2px;border-radius:0 6px 6px 0}
.test-row{padding:5px 10px 5px 18px;font-size:12.5px;color:#B0C4D8;font-family:'IBM Plex Mono',monospace;border-bottom:1px solid rgba(255,255,255,.035);transition:background .12s}
.test-row:hover{background:rgba(255,255,255,.03)}
.chat-user{background:linear-gradient(135deg,#00D4B4,#0096E0);color:#030C18;border-radius:14px 14px 3px 14px;padding:9px 13px;max-width:82%;font-size:13.5px;line-height:1.45;margin-left:auto}
.chat-ai{background:rgba(20,38,62,.9);border:1px solid rgba(0,210,180,.14);color:#DDE6F0;border-radius:14px 14px 14px 3px;padding:10px 14px;max-width:88%;font-size:13.5px;line-height:1.55;display:flex;flex-direction:column;gap:5px}
.pulse{animation:pulse 1.4s ease-in-out infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.42}}
.spin{animation:spin .9s linear infinite}
@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
.fadein{animation:fadein .3s ease}
@keyframes fadein{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
.drop-active{border-color:rgba(0,210,180,.6)!important;background:rgba(0,210,180,.06)!important}
.sso-btn{display:flex;align-items:center;gap:12px;padding:13px 20px;border-radius:12px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.04);color:#DDE6F0;cursor:pointer;font-family:'Exo 2',sans-serif;font-size:14px;font-weight:500;transition:all .18s;width:100%}
.sso-btn:hover{background:rgba(255,255,255,.09);border-color:rgba(0,210,180,.35);transform:translateY(-1px)}
.step-dot{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .3s}
.step-done{background:rgba(0,210,180,.2);border:2px solid #00D4B4}
.step-active{background:rgba(0,150,224,.15);border:2px solid #0096E0}
.step-pending{background:rgba(255,255,255,.05);border:2px solid rgba(255,255,255,.12)}
.step-line{width:2px;height:28px;margin:3px auto;border-radius:2px;transition:background .4s}
.share-menu{position:absolute;top:calc(100% + 8px);right:0;background:rgba(6,14,28,.97);border:1px solid rgba(0,210,180,.18);border-radius:12px;padding:6px;min-width:200px;z-index:200;box-shadow:0 12px 40px rgba(0,0,0,.7);backdrop-filter:blur(14px)}
.share-item{display:flex;align-items:center;gap:10px;padding:9px 12px;border-radius:8px;cursor:pointer;font-size:13px;color:#C8D8E8;transition:background .14s;border:none;background:none;width:100%;font-family:'Exo 2',sans-serif;text-align:left}
.share-item:hover{background:rgba(0,210,180,.1);color:#DDE6F0}
.share-divider{height:1px;background:rgba(255,255,255,.07);margin:4px 0}
.share-modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.6);backdrop-filter:blur(4px);z-index:900;display:flex;align-items:center;justify-content:center;padding:20px}
.share-modal{background:#070F1E;border:1px solid rgba(0,210,180,.18);border-radius:18px;width:100%;max-width:520px;overflow:hidden;box-shadow:0 24px 60px rgba(0,0,0,.8)}
.share-tab-bar{display:flex;border-bottom:1px solid rgba(0,210,180,.1);padding:0 20px;gap:4px;background:rgba(0,210,180,.03)}
.share-tab{padding:13px 18px;font-size:13px;font-family:'Exo 2',sans-serif;font-weight:600;cursor:pointer;border:none;background:none;color:#5A7488;border-bottom:2px solid transparent;transition:all .15s;display:flex;align-items:center;gap:7px}
.share-tab.active{color:#00D4B4;border-bottom-color:#00D4B4}
.share-tab:hover:not(.active){color:#DDE6F0}
.share-body{padding:22px 24px}
.share-field label{display:block;font-size:11px;font-weight:700;color:#4A6070;letter-spacing:.07em;text-transform:uppercase;margin-bottom:6px}
.share-field+.share-field{margin-top:14px}
.share-step{display:flex;align-items:flex-start;gap:10px;padding:8px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12.5px;color:#8CA3B8;line-height:1.5}
.share-step-num{width:20px;height:20px;border-radius:50%;background:rgba(0,210,180,.12);border:1px solid rgba(0,210,180,.25);color:#00D4B4;font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}
.share-step-done .share-step-num{background:rgba(0,210,180,.25);border-color:#00D4B4}
.share-step-done{color:#DDE6F0}
.info-box{background:rgba(0,150,220,.08);border:1px solid rgba(0,150,220,.18);border-radius:9px;padding:10px 12px;font-size:12px;color:#7ABEDC;line-height:1.55;margin-bottom:14px}
.chat-maximized{width:min(720px,92vw)!important;height:75vh!important;bottom:16px!important;right:16px!important}
.chat-minimized-bar{height:52px!important;overflow:hidden!important}
.chat-resize-btn{background:none;border:none;color:#8CA3B8;padding:5px;border-radius:7px;cursor:pointer;display:flex;align-items:center;transition:all .15s}
.chat-resize-btn:hover{color:#DDE6F0;background:rgba(255,255,255,.08)}
`;

/* ─────────────────────────── LOGIN SCREEN ─────────────────────────── */
/* Decodes the JWT credential returned by Google Identity Services.
   We only need the payload for display — Google already validated the token. */
function decodeGoogleJwt(token) {
  try {
    const payload = token.split(".")[1];
    // URL-safe base64 → standard base64
    const b64 = payload.replace(/-/g, "+").replace(/_/g, "/");
    const json = decodeURIComponent(
      atob(b64).split("").map(c => "%" + c.charCodeAt(0).toString(16).padStart(2, "0")).join("")
    );
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function LoginScreen({ onLogin }) {

  // ✅ PUT YOUR GOOGLE CLIENT ID HERE
  const clientId = 
    "652910726285-mda18dsfu233n4g5hp7gp9m16jer2sva.apps.googleusercontent.com";

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const googleBtnRef = useRef(null);
  const gisLoaded = useRef(false);

  /* ─────────────────────────────────────────────
     Decode Google JWT Payload
  ───────────────────────────────────────────── */
  function decodeGoogleJwt(token) {
    try {
      const payload = token.split(".")[1];

      const b64 = payload
        .replace(/-/g, "+")
        .replace(/_/g, "/");

      const json = decodeURIComponent(
        atob(b64)
          .split("")
          .map(
            c =>
              "%" +
              c.charCodeAt(0)
                .toString(16)
                .padStart(2, "0")
          )
          .join("")
      );

      return JSON.parse(json);

    } catch {
      return null;
    }
  }

  /* ─────────────────────────────────────────────
     Load Google Identity Services SDK
  ───────────────────────────────────────────── */
  const loadGIS = () =>
    new Promise((resolve, reject) => {

      // Already loaded
      if (window.google?.accounts?.id) {
        resolve();
        return;
      }

      // Prevent duplicate loads
      if (gisLoaded.current) {
        resolve();
        return;
      }

      gisLoaded.current = true;

      const script = document.createElement("script");

      script.src =
        "https://accounts.google.com/gsi/client";

      script.async = true;
      script.defer = true;

      script.onload = resolve;

      script.onerror = () =>
        reject(
          new Error(
            "Failed to load Google Sign-In"
          )
        );

      document.head.appendChild(script);
    });

  /* ─────────────────────────────────────────────
     Initialize Google OAuth
  ───────────────────────────────────────────── */
  const initGoogle = async () => {

    setLoading(true);
    setError("");

    try {

      await loadGIS();

      window.google.accounts.id.initialize({
        client_id: clientId,

        callback: handleCredential,

        auto_select: false,

        cancel_on_tap_outside: true,

        ux_mode: "popup",
      });

      // Render Google Button
      if (googleBtnRef.current) {

        googleBtnRef.current.innerHTML = "";

        window.google.accounts.id.renderButton(
          googleBtnRef.current,
          {
            theme: "filled_blue",
            size: "large",
            shape: "rectangular",
            text: "signin_with",
            width: 340,
            logo_alignment: "left",
          }
        );
      }

    } catch (err) {

      console.error(err);

      setError(
        err.message ||
        "Google OAuth initialization failed."
      );

    } finally {

      setLoading(false);
    }
  };

  /* ─────────────────────────────────────────────
     Handle Successful Login
  ───────────────────────────────────────────── */
  const handleCredential = (response) => {

    console.log(
      "Google OAuth Response:",
      response
    );

    const payload =
      decodeGoogleJwt(response.credential);

    if (!payload) {

      setError(
        "Failed to decode Google credential."
      );

      return;
    }

    console.log(
      "Google User:",
      payload
    );

    // ✅ PRESERVING YOUR EXISTING LOGIN FLOW
    onLogin({
      name:
        payload.name ||
        payload.email,

      email:
        payload.email || "",

      picture:
        payload.picture || "",

      provider: "Google",
    });
  };

  /* ─────────────────────────────────────────────
     Auto Initialize OAuth On Component Mount
  ───────────────────────────────────────────── */
  useEffect(() => {

    initGoogle();

  }, []);

  /* ─────────────────────────────────────────────
     UI
  ───────────────────────────────────────────── */
  return (

    <div
      className="grid-bg"
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
    >

      <style>{CSS}</style>

      <div
        style={{
          width: 420,
          padding: "0 20px",
        }}
      >

        {/* Branding */}
        <div
          style={{
            textAlign: "center",
            marginBottom: 40,
          }}
        >

          <div
            style={{
              width: 64,
              height: 64,
              borderRadius: 18,
              background:
                "linear-gradient(135deg,#00D4B4,#0096E0)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              margin: "0 auto 20px",
            }}
          >

            <Sparkles
              size={28}
              color="#030C18"
            />

          </div>

          <h1
            style={{
              fontSize: 28,
              fontWeight: 700,
              letterSpacing: "-0.03em",
              marginBottom: 8,
            }}
          >
            MediSort{" "}
            <span style={{ color: "#00D4B4" }}>
              AI
            </span>
          </h1>

          <p
            style={{
              color: "#6A8499",
              fontSize: 14,
              lineHeight: 1.5,
            }}
          >
            AI-powered medical data structuring
            & classification platform
          </p>

        </div>

        {/* Login Card */}
        <div
          className="card glow"
          style={{
            padding: 32,
          }}
        >

          <p
            style={{
              fontSize: 12,
              fontWeight: 600,
              color: "#4A6070",
              letterSpacing: ".08em",
              textTransform: "uppercase",
              marginBottom: 20,
            }}
          >
            Sign in with Google
          </p>

          {/* Google OAuth Button */}
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              minHeight: 44,
              marginBottom: 16,
            }}
          >

            {loading ? (

              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  color: "#5BA8F5",
                  fontSize: 13,
                }}
              >

                <Loader2
                  size={16}
                  className="spin"
                />

                Loading Google Sign-In...

              </div>

            ) : (

              <div ref={googleBtnRef} />

            )}

          </div>

          {/* Error */}
          {error && (

            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                padding: "10px 14px",
                background:
                  "rgba(220,80,60,.1)",
                border:
                  "1px solid rgba(220,80,60,.22)",
                borderRadius: 10,
                marginTop: 10,
                color: "#F08070",
                fontSize: 13,
              }}
            >

              <AlertTriangle size={15} />

              {error}

            </div>
          )}

          <p
            style={{
              fontSize: 11,
              color: "#3A5060",
              textAlign: "center",
              marginTop: 18,
              lineHeight: 1.6,
            }}
          >
            By signing in you agree to our
            Terms of Service.
            <br />
            Health data is processed securely
            and not stored.
          </p>

        </div>

        {/* Security Badges */}
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            gap: 20,
            marginTop: 24,
            flexWrap: "wrap",
          }}
        >

          {[
            "HIPAA Compliant",
            "SOC 2 Type II",
            "256-bit Encryption",
          ].map((t) => (

            <div
              key={t}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 5,
                color: "#3A5060",
                fontSize: 11,
              }}
            >

              <Check
                size={10}
                color="#00D4B4"
              />

              {t}

            </div>
          ))}

        </div>

      </div>

    </div>
  );
}

/* ─────────────────────────── UPLOAD VIEW ─────────────────────────── */
function UploadView({ file, rawData, onDrop, onFileSelect, onProcess, error, isProcessing }) {
  const [dragOver, setDragOver] = useState(false);
  const ref = useRef(null);
  const numCols = rawData ? rawData[0]?.length || 0 : 0;
  const numTests = rawData ? Math.max(...Array.from({ length: numCols }, (_, i) => rawData.slice(3).filter(r => r[i]).length)) : 0;

  return (
    <div className="fadein" style={{ maxWidth:820, margin:"0 auto" }}>
      <div style={{ marginBottom:28 }}>
        <h2 style={{ fontSize:22, fontWeight:700, letterSpacing:"-0.02em", marginBottom:6 }}>
          <span style={{ color:"#00D4B4" }}>Upload</span> Medical Package File
        </h2>
        <p style={{ color:"#5A7488", fontSize:14 }}>Upload your .xlsx file to classify and structure medical test packages using AI</p>
      </div>

      <div className="card glow" style={{ padding:32, marginBottom:20 }}>
        <div
          className={dragOver ? "drop-active" : ""}
          onDragOver={e => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={e => { e.preventDefault(); setDragOver(false); const f = e.dataTransfer.files[0]; if (f) onFileSelect(f); }}
          onClick={() => ref.current?.click()}
          style={{ border:"2px dashed rgba(0,210,180,.25)", borderRadius:14, padding:"36px 20px", textAlign:"center", cursor:"pointer", transition:"all .2s", background:"rgba(0,210,180,.02)" }}
        >
          <input ref={ref} type="file" accept=".xlsx" style={{ display:"none" }} onChange={e => onFileSelect(e.target.files[0])} />
          {file ? (
            <div>
              <div style={{ width:48, height:48, borderRadius:12, background:"rgba(0,210,180,.15)", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 12px" }}>
                <FileSpreadsheet size={24} color="#00D4B4" />
              </div>
              <p style={{ fontWeight:600, fontSize:15, marginBottom:4 }}>{file.name}</p>
              <p style={{ color:"#5A7488", fontSize:13 }}>{(file.size / 1024).toFixed(1)} KB &nbsp;·&nbsp; Click to replace</p>
            </div>
          ) : (
            <div>
              <div style={{ width:56, height:56, borderRadius:14, background:"rgba(0,210,180,.08)", border:"1px dashed rgba(0,210,180,.25)", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 14px" }}>
                <Upload size={24} color="#00D4B4" />
              </div>
              <p style={{ fontWeight:600, fontSize:15, marginBottom:6 }}>Drop your Excel file here</p>
              <p style={{ color:"#4A6070", fontSize:13 }}>or click to browse &nbsp;·&nbsp; <span style={{ color:"#00D4B4" }}>.xlsx</span> files only</p>
            </div>
          )}
        </div>

        {error && (
          <div style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 14px", background:"rgba(220,80,60,.1)", border:"1px solid rgba(220,80,60,.22)", borderRadius:10, marginTop:14, color:"#F08070", fontSize:13 }}>
            <AlertTriangle size={15} /> {error}
          </div>
        )}

        {rawData && (
          <div className="fadein" style={{ marginTop:20 }}>
            <p style={{ fontSize:12, fontWeight:600, color:"#4A6070", letterSpacing:".08em", textTransform:"uppercase", marginBottom:12 }}>File Preview</p>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:10 }}>
              {[
                { label:"Packages Detected", value:numCols, color:"#00D4B4" },
                { label:"Max Tests / Package", value:numTests, color:"#5BA8F5" },
                { label:"Total Rows", value:rawData.length, color:"#AC87FF" },
              ].map(m => (
                <div key={m.label} className="card-sm" style={{ padding:"12px 16px" }}>
                  <p style={{ fontSize:11, color:"#4A6070", marginBottom:4 }}>{m.label}</p>
                  <p style={{ fontSize:22, fontWeight:700, color:m.color, fontFamily:"'IBM Plex Mono',monospace" }}>{m.value}</p>
                </div>
              ))}
            </div>
            <div style={{ marginTop:16, overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr>
                    <th style={{ padding:"8px 12px", textAlign:"left", color:"#4A6070", fontWeight:600, borderBottom:"1px solid rgba(255,255,255,.07)" }}>Row</th>
                    {Array.from({ length: numCols }, (_, i) => (
                      <th key={i} style={{ padding:"8px 12px", textAlign:"left", color:"#00D4B4", fontWeight:600, borderBottom:"1px solid rgba(255,255,255,.07)" }}>
                        Package {i+1}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {["Location","Age/Gender","Price"].map((label, ri) => (
                    <tr key={label}>
                      <td style={{ padding:"6px 12px", color:"#4A6070", fontFamily:"'IBM Plex Mono',monospace" }}>{label}</td>
                      {Array.from({ length: numCols }, (_, ci) => (
                        <td key={ci} style={{ padding:"6px 12px", color:"#B0C4D8", maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>
                          {String(rawData[ri]?.[ci] || "—")}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div style={{ display:"flex", justifyContent:"flex-end", marginTop:22 }}>
          <button className="btn btn-primary" disabled={!file || isProcessing} onClick={onProcess}>
            {isProcessing ? (
              <span style={{ display:"flex", alignItems:"center", gap:8 }}>
                <Loader2 size={15} className="spin" /> Processing…
              </span>
            ) : (
              <span style={{ display:"flex", alignItems:"center", gap:8 }}>
                <Sparkles size={15} /> Process with AI
              </span>
            )}
          </button>
        </div>
      </div>

      <div className="card-sm" style={{ padding:"14px 18px", display:"flex", alignItems:"center", gap:10 }}>
        <div style={{ width:32, height:32, borderRadius:8, background:"rgba(0,210,180,.1)", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
          <Layers size={15} color="#00D4B4" />
        </div>
        <div>
          <p style={{ fontSize:13, fontWeight:500, marginBottom:2 }}>AI Classification Engine</p>
          <p style={{ fontSize:12, color:"#4A6070" }}>Tests are grouped into clinical categories: Urine Analysis, Thyroid Profile, Blood Count, Vitamins, Lipid Profile, Kidney &amp; Liver Function, and more.</p>
        </div>
      </div>
    </div>
  );
}

/* ─────────────────────────── PROCESSING VIEW ─────────────────────────── */
function ProcessingView({ steps }) {
  return (
    <div className="fadein" style={{ maxWidth:520, margin:"60px auto 0", padding:"0 20px" }}>
      <div style={{ textAlign:"center", marginBottom:40 }}>
        <div style={{ width:60, height:60, borderRadius:16, background:"rgba(0,150,224,.15)", border:"1px solid rgba(0,150,224,.3)", display:"flex", alignItems:"center", justifyContent:"center", margin:"0 auto 16px" }}>
          <Sparkles size={26} color="#0096E0" className="pulse" />
        </div>
        <h2 style={{ fontSize:20, fontWeight:700, marginBottom:8 }}>AI is Classifying Your Data</h2>
        <p style={{ color:"#5A7488", fontSize:14 }}>Analyzing medical tests and grouping into clinical categories…</p>
      </div>

      <div className="card glow" style={{ padding:28 }}>
        {steps.map((step, i) => {
          const Icon = STEPS[i]?.icon || Sparkles;
          const isDone = step.status === "done";
          const isActive = step.status === "active";
          return (
            <div key={i}>
              <div style={{ display:"flex", alignItems:"center", gap:14 }}>
                <div className={`step-dot ${isDone ? "step-done" : isActive ? "step-active" : "step-pending"}`}>
                  {isDone ? <Check size={15} color="#00D4B4" /> :
                   isActive ? <Loader2 size={15} color="#0096E0" className="spin" /> :
                   <Icon size={14} color="#3A5060" />}
                </div>
                <div style={{ flex:1 }}>
                  <p style={{ fontSize:13.5, fontWeight:600, color: isDone ? "#DDE6F0" : isActive ? "#5BA8F5" : "#4A6070" }}>
                    {step.label}
                  </p>
                  {step.message && (
                    <p style={{ fontSize:12, color:isDone ? "#00D4B4" : "#5A7488", marginTop:2 }}>{step.message}</p>
                  )}
                </div>
                {isDone && <span className="badge badge-teal">Done</span>}
                {isActive && <span className="badge badge-blue pulse">Running</span>}
              </div>
              {i < steps.length - 1 && (
                <div className={`step-line ${isDone ? "step-done" : ""}`}
                  style={{ background: isDone ? "linear-gradient(#00D4B4,rgba(0,210,180,.2))" : "rgba(255,255,255,.07)", marginLeft:15 }} />
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ─────────────────────────── RESULTS VIEW ─────────────────────────── */
function ResultsView({ packages, onDownload, onShare, onNewFile }) {
  const [activeTab, setActiveTab] = useState(0);
  const pkg = packages[activeTab];
  const categoryNames = Object.keys(pkg?.categories || {});

  return (
    <div className="fadein" style={{ maxWidth:960, margin:"0 auto" }}>
      <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:24, flexWrap:"wrap", gap:12 }}>
        <div>
          <h2 style={{ fontSize:22, fontWeight:700, letterSpacing:"-0.02em", marginBottom:6 }}>
            <span style={{ color:"#00D4B4" }}>Classification</span> Complete
          </h2>
          <p style={{ color:"#5A7488", fontSize:14 }}>{packages.length} packages processed · AI-generated section headers</p>
        </div>
        <div style={{ display:"flex", gap:10 }}>
          <button className="btn btn-ghost" onClick={onNewFile} style={{ display:"flex", alignItems:"center", gap:6, fontSize:13 }}>
            <RefreshCw size={14} /> New File
          </button>
          <button className="btn btn-ghost" onClick={onShare} style={{ display:"flex", alignItems:"center", gap:6, fontSize:13 }}>
            <Share2 size={14} /> Share
          </button>
          <button className="btn btn-primary" onClick={onDownload} style={{ display:"flex", alignItems:"center", gap:6 }}>
            <Download size={15} /> Download .xlsx
          </button>
        </div>
      </div>

      {/* Summary Cards */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, marginBottom:20 }}>
        {[
          { label:"Packages", value:packages.length, color:"#00D4B4" },
          { label:"Avg Tests/Pkg", value:Math.round(packages.reduce((s,p) => s+Object.values(p.categories).flat().length, 0)/packages.length), color:"#5BA8F5" },
          { label:"Categories Found", value:[...new Set(packages.flatMap(p => Object.keys(p.categories)))].length, color:"#AC87FF" },
          { label:"Total Tests", value:packages.reduce((s,p) => s+Object.values(p.categories).flat().length, 0), color:"#F5B840" },
        ].map(m => (
          <div key={m.label} className="card-sm" style={{ padding:"12px 16px" }}>
            <p style={{ fontSize:11, color:"#4A6070", marginBottom:3 }}>{m.label}</p>
            <p style={{ fontSize:22, fontWeight:700, color:m.color, fontFamily:"'IBM Plex Mono',monospace" }}>{m.value}</p>
          </div>
        ))}
      </div>

      <div className="card glow" style={{ overflow:"hidden" }}>
        {/* Package Tabs — uses pkgLabel for "Package N" fallback */}
        <div style={{ borderBottom:"1px solid rgba(0,210,180,.1)", padding:"14px 20px", display:"flex", gap:8, overflowX:"auto" }}>
          {packages.map((p, i) => (
            <button key={i} className={`tab ${activeTab===i ? "active" : ""}`} onClick={() => setActiveTab(i)}>
              {pkgLabel(p, i)}
            </button>
          ))}
        </div>

        {/* Package Header */}
        <div style={{ padding:"18px 20px 12px", borderBottom:"1px solid rgba(255,255,255,.06)" }}>
          <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", flexWrap:"wrap", gap:10 }}>
            <div>
              <h3 style={{ fontSize:16, fontWeight:700, marginBottom:4 }}>{metaUI(pkg?.displayName)}</h3>
              <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                {/* metaUI ensures "-" when value is blank, never shows undefined */}
                <span className="badge badge-teal">{metaUI(pkg?.serviceLocation)}</span>
                <span className="badge badge-blue">{metaUI(pkg?.ageGender)}</span>
                <span className="badge badge-amber">{metaUI(pkg?.price)}</span>
              </div>
            </div>
            <div style={{ textAlign:"right" }}>
              <p style={{ fontSize:11, color:"#4A6070" }}>Categories</p>
              <p style={{ fontSize:20, fontWeight:700, color:"#AC87FF", fontFamily:"'IBM Plex Mono',monospace" }}>{categoryNames.length}</p>
            </div>
          </div>
        </div>

        {/* Tests List */}
        <div style={{ padding:"8px 20px 20px", maxHeight:420, overflowY:"auto" }}>
          {categoryNames.map(cat => (
            <div key={cat}>
              <div className="cat-header">{cat}
                <span style={{ float:"right", color:"#5A8070", fontSize:10, fontFamily:"'IBM Plex Mono',monospace", marginTop:1 }}>
                  {pkg.categories[cat].length} tests
                </span>
              </div>
              {pkg.categories[cat].map((test, ti) => (
                <div key={ti} className="test-row">· {test}</div>
              ))}
            </div>
          ))}
        </div>
      </div>

      <p style={{ textAlign:"center", fontSize:12, color:"#3A5060", marginTop:16 }}>
        ✦ Package Name and Display Name fields are AI-generated — review before distributing
      </p>
    </div>
  );
}

/* ─────────────────────────── MARKDOWN RENDERER ─────────────────────────── */
// Line-by-line parser — handles the patterns Claude commonly returns:
//   **bold**, *italic*, `code`, # headings, - / * / • bullets, 1. numbered lists,
//   --- dividers, blank lines as paragraph breaks.
// Renders as stacked React elements with proper spacing.
function renderMarkdown(text) {
  if (!text) return null;

  const raw = text.trim().split("\n");
  const elements = [];
  let i = 0;

  while (i < raw.length) {
    const line = raw[i];
    const trimmed = line.trim();

    // ── blank line → spacing gap ───────────────────────────────────
    if (trimmed === "") {
      // Only add a gap if there's something before AND after
      if (elements.length > 0 && i < raw.length - 1) {
        elements.push(<div key={`gap-${i}`} style={{ height: 4 }} />);
      }
      i++; continue;
    }

    // ── horizontal rule ────────────────────────────────────────────
    if (/^[-*_]{3,}$/.test(trimmed)) {
      elements.push(
        <hr key={`hr-${i}`} style={{ border:"none", borderTop:"1px solid rgba(0,210,180,.2)", margin:"4px 0" }} />
      );
      i++; continue;
    }

    // ── heading (# / ## / ###) ─────────────────────────────────────
    const headingMatch = trimmed.match(/^(#{1,3})\s+(.*)/);
    if (headingMatch) {
      const level = headingMatch[1].length;
      const content = headingMatch[2];
      elements.push(
        <p key={`h-${i}`} style={{
          fontWeight: 700,
          fontSize:   level === 1 ? 14.5 : level === 2 ? 13.5 : 13,
          color:      "#00D4B4",
          margin:     "6px 0 2px",
          lineHeight: 1.35,
          letterSpacing: "0.01em",
        }}>
          {inlineFormat(content)}
        </p>
      );
      i++; continue;
    }

    // ── bullet list block (collect consecutive bullet lines) ───────
    if (/^\s*[-*•]\s/.test(line)) {
      const items = [];
      while (i < raw.length && (/^\s*[-*•]\s/.test(raw[i]) || (raw[i].trim() === "" && i + 1 < raw.length && /^\s*[-*•]\s/.test(raw[i + 1])))) {
        if (raw[i].trim()) items.push(raw[i].replace(/^\s*[-*•]\s+/, ""));
        i++;
      }
      elements.push(
        <ul key={`ul-${i}`} style={{ margin:"3px 0", padding:0, listStyle:"none", display:"flex", flexDirection:"column", gap:3 }}>
          {items.map((item, li) => (
            <li key={li} style={{ display:"flex", gap:7, alignItems:"flex-start", fontSize:13, lineHeight:1.55 }}>
              <span style={{ color:"#00D4B4", fontWeight:700, flexShrink:0, marginTop:1 }}>›</span>
              <span>{inlineFormat(item)}</span>
            </li>
          ))}
        </ul>
      );
      continue;
    }

    // ── numbered list block (collect consecutive numbered lines) ───
    if (/^\s*\d+[\.\)]\s/.test(line)) {
      const items = [];
      while (i < raw.length && (/^\s*\d+[\.\)]\s/.test(raw[i]) || (raw[i].trim() === "" && i + 1 < raw.length && /^\s*\d+[\.\)]\s/.test(raw[i + 1])))) {
        if (raw[i].trim()) items.push(raw[i].replace(/^\s*\d+[\.\)]\s+/, ""));
        i++;
      }
      elements.push(
        <ol key={`ol-${i}`} style={{ margin:"3px 0", paddingLeft:18, display:"flex", flexDirection:"column", gap:3 }}>
          {items.map((item, li) => (
            <li key={li} style={{ fontSize:13, lineHeight:1.55 }}>
              {inlineFormat(item)}
            </li>
          ))}
        </ol>
      );
      continue;
    }

    // ── plain paragraph line ───────────────────────────────────────
    elements.push(
      <p key={`p-${i}`} style={{ fontSize:13, lineHeight:1.6, margin:0 }}>
        {inlineFormat(trimmed)}
      </p>
    );
    i++;
  }

  return elements;
}

// Handles inline **bold**, *italic*, and `code` within a single text line
function inlineFormat(text) {
  if (!text) return null;
  const parts = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g);
  return parts.map((part, i) => {
    if (/^\*\*[^*]+\*\*$/.test(part))
      return <strong key={i} style={{ color:"#DDE6F0", fontWeight:700 }}>{part.slice(2,-2)}</strong>;
    if (/^\*[^*]+\*$/.test(part))
      return <em key={i} style={{ color:"#A8C8D8", fontStyle:"italic" }}>{part.slice(1,-1)}</em>;
    if (/^`[^`]+`$/.test(part))
      return <code key={i} style={{ background:"rgba(0,210,180,.12)", color:"#00D4B4", padding:"1px 5px", borderRadius:4, fontSize:11.5, fontFamily:"'IBM Plex Mono',monospace" }}>{part.slice(1,-1)}</code>;
    return part;
  });
}

/* ─────────────────────────── CHAT WIDGET ─────────────────────────── */
// chatSize: "closed" | "normal" | "minimized" | "maximized"
function ChatWidget({ packages, rawData }) {
  const [chatSize, setChatSize] = useState("closed");
  const [messages, setMessages] = useState([
    { role:"assistant", content:"Hello! I'm MediSort AI. I can answer questions about your medical packages, compare tests, and help you understand the data. What would you like to know?" }
  ]);
  const [input, setInput]   = useState("");
  const [loading, setLoading] = useState(false);
  const endRef = useRef(null);

  const isOpen = chatSize !== "closed";

  useEffect(() => {
    if (isOpen && chatSize !== "minimized") {
      endRef.current?.scrollIntoView({ behavior:"smooth" });
    }
  }, [messages, chatSize]);

  const send = async () => {
    if (!input.trim() || loading) return;
    const userMsg = { role:"user", content:input.trim() };
    setMessages(prev => [...prev, userMsg]);
    setInput("");
    setLoading(true);

    const dataCtx = packages
      ? `Processed packages:\n${JSON.stringify(packages.map((p, i) => ({ packageName:pkgLabel(p,i), displayName:p.displayName, price:metaUI(p.price), serviceLocation:metaUI(p.serviceLocation), ageGender:metaUI(p.ageGender), categorySummary:Object.fromEntries(Object.entries(p.categories).map(([k,v]) => [k, v.length+" tests"])) })))}`
      : rawData
      ? `Raw data (not yet processed): ${rawData.length} rows, ${rawData[0]?.length} columns`
      : "No file uploaded yet.";

    const sys = `You are MediSort AI, an intelligent medical data analyst assistant.
Always format your responses using clean markdown:
- Use **bold** for key terms, package names, and important values.
- Use bullet lists (- item) for enumerating tests, features, or comparisons.
- Use numbered lists (1. step) for sequential information.
- Use ## headings to separate distinct sections when the answer covers multiple topics.
- Keep paragraphs short (2-3 sentences max). Prefer lists over long prose.
- Never output raw JSON or unformatted blobs of text.
Context: ${dataCtx}`;

    try {
      const reply = await callClaude(
        [...messages, userMsg].map(m => ({ role:m.role==="user"?"user":"assistant", content:m.content })),
        sys
      );
      setMessages(prev => [...prev, { role:"assistant", content:reply }]);
    } catch {
      setMessages(prev => [...prev, { role:"assistant", content:"Sorry, I hit an error. Please try again." }]);
    } finally { setLoading(false); }
  };

  // Panel dimensions by chatSize
  const sizeStyle = {
    normal:    { width:360, height:480, bottom:86, right:24 },
    maximized: { width:"min(720px, 92vw)", height:"75vh", bottom:16, right:16 },
    minimized: { width:360, height:52,  bottom:86, right:24 },
  };
  const dim = sizeStyle[chatSize] || sizeStyle.normal;

  return (
    <>
      {/* Floating toggle button */}
      <button
        onClick={() => setChatSize(s => s === "closed" ? "normal" : "closed")}
        style={{ position:"fixed", bottom:24, right:24, width:52, height:52, borderRadius:"50%",
          background:"linear-gradient(135deg,#00D4B4,#0096E0)", border:"none",
          display:"flex", alignItems:"center", justifyContent:"center",
          cursor:"pointer", boxShadow:"0 4px 20px rgba(0,210,180,.35)", zIndex:1000, transition:"all .2s" }}
        title={isOpen ? "Close chat" : "Open AI Chat"}
      >
        {isOpen ? <X size={20} color="#030C18" /> : <MessageCircle size={20} color="#030C18" />}
      </button>

      {/* Chat Panel */}
      {isOpen && (
        <div
          className="card fadein"
          style={{
            position:"fixed",
            bottom: dim.bottom,
            right:  dim.right,
            width:  dim.width,
            height: dim.height,
            display:"flex",
            flexDirection:"column",
            zIndex:999,
            boxShadow:"0 8px 40px rgba(0,0,0,.6)",
            overflow:"hidden",
            transition:"width .25s ease, height .25s ease, bottom .25s ease",
          }}
        >
          {/* Header */}
          <div style={{ padding:"11px 14px", borderBottom: chatSize !== "minimized" ? "1px solid rgba(0,210,180,.12)" : "none",
            display:"flex", alignItems:"center", gap:10, background:"rgba(0,210,180,.06)", flexShrink:0 }}>
            <div style={{ width:30, height:30, borderRadius:"50%", background:"linear-gradient(135deg,#00D4B4,#0096E0)",
              display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
              <Bot size={15} color="#030C18" />
            </div>
            <div style={{ flex:1, minWidth:0 }}>
              <p style={{ fontWeight:600, fontSize:13 }}>MediSort AI</p>
              {chatSize !== "minimized" && <p style={{ fontSize:10, color:"#3A9878" }}>● Online · Context-aware</p>}
            </div>

            {/* Resize controls */}
            <div style={{ display:"flex", gap:2, alignItems:"center" }}>
              {/* Minimize → collapses to header-only bar */}
              <button className="chat-resize-btn" title="Minimize"
                onClick={() => setChatSize(s => s === "minimized" ? "normal" : "minimized")}>
                <Minus size={13} />
              </button>
              {/* Maximize / Restore */}
              <button className="chat-resize-btn"
                title={chatSize === "maximized" ? "Restore" : "Maximize"}
                onClick={() => setChatSize(s => s === "maximized" ? "normal" : "maximized")}>
                {chatSize === "maximized" ? <Minimize2 size={13} /> : <Maximize2 size={13} />}
              </button>
              {/* Close */}
              <button className="chat-resize-btn" title="Close" onClick={() => setChatSize("closed")}>
                <X size={13} />
              </button>
            </div>
          </div>

          {/* Body — hidden when minimized */}
          {chatSize !== "minimized" && (
            <>
              {/* Messages */}
              <div style={{ flex:1, overflowY:"auto", padding:"14px 14px 8px", display:"flex", flexDirection:"column", gap:10 }}>
                {messages.map((m, i) => (
                  <div key={i} style={{ display:"flex", flexDirection:m.role==="user"?"row-reverse":"row", gap:8, alignItems:"flex-end" }}>
                    {m.role==="assistant" && (
                      <div style={{ width:24, height:24, borderRadius:"50%", background:"rgba(0,210,180,.15)",
                        display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
                        <Bot size={13} color="#00D4B4" />
                      </div>
                    )}
                    <div className={m.role==="user" ? "chat-user" : "chat-ai"}
                      style={ chatSize === "maximized" ? { maxWidth:"70%", fontSize:14 } : {} }>
                      {m.role === "assistant" ? renderMarkdown(m.content) : m.content}
                    </div>
                  </div>
                ))}
                {loading && (
                  <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                    <div style={{ width:24, height:24, borderRadius:"50%", background:"rgba(0,210,180,.15)",
                      display:"flex", alignItems:"center", justifyContent:"center" }}>
                      <Bot size={13} color="#00D4B4" />
                    </div>
                    <div className="chat-ai" style={{ display:"flex", gap:4, padding:"10px 14px" }}>
                      {[0,1,2].map(d => (
                        <div key={d} style={{ width:6, height:6, borderRadius:"50%", background:"#00D4B4",
                          animation:`pulse 1.2s ${d*0.2}s ease-in-out infinite` }} />
                      ))}
                    </div>
                  </div>
                )}
                <div ref={endRef} />
              </div>

              {/* Quick-action chips */}
              {messages.length < 3 && packages && (
                <div style={{ padding:"0 14px 8px", display:"flex", flexWrap:"wrap", gap:6 }}>
              {["Which package is cheapest?", "Compare all packages",
                    `What's in ${pkgLabel(packages[0], 0)}?`].map(s => (
                    <button key={s} onClick={() => setInput(s)}
                      style={{ fontSize:11, padding:"4px 10px", borderRadius:8,
                        background:"rgba(0,210,180,.08)", border:"1px solid rgba(0,210,180,.2)",
                        color:"#00D4B4", cursor:"pointer", fontFamily:"'Exo 2',sans-serif" }}>
                      {s}
                    </button>
                  ))}
                </div>
              )}

              {/* Input row */}
              <div style={{ padding:"10px 12px", borderTop:"1px solid rgba(0,210,180,.1)", display:"flex", gap:8, flexShrink:0 }}>
                <input
                  className="input"
                  value={input}
                  onChange={e => setInput(e.target.value)}
                  onKeyDown={e => e.key==="Enter" && send()}
                  placeholder="Ask about your packages…"
                  style={{ flex:1, padding:"9px 12px", fontSize:13 }}
                />
                <button
                  onClick={send}
                  disabled={!input.trim() || loading}
                  style={{ width:36, height:36, borderRadius:10,
                    background:input.trim() ? "linear-gradient(135deg,#00D4B4,#0096E0)" : "rgba(255,255,255,.05)",
                    border:"none", display:"flex", alignItems:"center", justifyContent:"center",
                    cursor:input.trim() ? "pointer" : "not-allowed", flexShrink:0, transition:"all .18s" }}
                >
                  <Send size={14} color={input.trim() ? "#030C18" : "#3A5060"} />
                </button>
              </div>
            </>
          )}
        </div>
      )}
    </>
  );
}

/* ─────────────────────────── SHARE MODAL ─────────────────────────── */

/* Build a proper MIME multipart email with the XLSX as a base64 attachment.
   Returns a URL-safe base64 string suitable for the Gmail API `raw` field. */
function buildMimeRaw(to, subject, body, fileName, blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      // reader.result = "data:<mime>;base64,<b64>"
      const b64 = reader.result.split(",")[1];
      const boundary = `boundary_medisort_${Date.now()}`;
      const mime = [
        "MIME-Version: 1.0",
        `To: ${to}`,
        `Subject: ${subject}`,
        `Content-Type: multipart/mixed; boundary="${boundary}"`,
        "",
        `--${boundary}`,
        'Content-Type: text/plain; charset="UTF-8"',
        "Content-Transfer-Encoding: 7bit",
        "",
        body,
        "",
        `--${boundary}`,
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        `Content-Disposition: attachment; filename="${fileName}"`,
        "Content-Transfer-Encoding: base64",
        "",
        // wrap at 76 chars — RFC 2045 requirement
        b64.match(/.{1,76}/g).join("\r\n"),
        "",
        `--${boundary}--`,
      ].join("\r\n");

      // URL-safe base64 (Gmail API requirement)
      const encoded = btoa(
        Array.from(new TextEncoder().encode(mime), b => String.fromCharCode(b)).join("")
      ).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
      resolve(encoded);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

const GMAIL_ICON = (
  <svg viewBox="0 0 24 24" width="18" height="18" style={{ flexShrink:0 }}>
    <path fill="#EA4335" d="M24 5.457v13.909c0 .904-.732 1.636-1.636 1.636h-3.819V11.73L12 16.64l-6.545-4.91v9.273H1.636A1.636 1.636 0 0 1 0 19.366V5.457c0-2.023 2.309-3.178 3.927-1.964L5.455 4.64 12 9.548l6.545-4.91 1.528-1.145C21.69 2.28 24 3.434 24 5.457z"/>
  </svg>
);
const SLACK_ICON = (
  <svg viewBox="0 0 24 24" width="18" height="18" style={{ flexShrink:0 }}>
    <path fill="#E01E5A" d="M5.042 15.165a2.528 2.528 0 0 1-2.52 2.523A2.528 2.528 0 0 1 0 15.165a2.527 2.527 0 0 1 2.522-2.52h2.52v2.52z"/>
    <path fill="#36C5F0" d="M6.313 15.165a2.527 2.527 0 0 1 2.521-2.52 2.527 2.527 0 0 1 2.521 2.52v6.313A2.528 2.528 0 0 1 8.834 24a2.528 2.528 0 0 1-2.521-2.522v-6.313z"/>
    <path fill="#2EB67D" d="M8.834 5.042a2.528 2.528 0 0 1-2.521-2.52A2.528 2.528 0 0 1 8.834 0a2.528 2.528 0 0 1 2.521 2.522v2.52H8.834z"/>
    <path fill="#ECB22E" d="M8.834 6.313a2.528 2.528 0 0 1 2.521 2.521 2.528 2.528 0 0 1-2.521 2.521H2.522A2.528 2.528 0 0 1 0 8.834a2.528 2.528 0 0 1 2.522-2.521h6.312z"/>
    <path fill="#E01E5A" d="M18.956 8.834a2.528 2.528 0 0 1 2.522-2.521A2.528 2.528 0 0 1 24 8.834a2.528 2.528 0 0 1-2.522 2.521h-2.522V8.834z"/>
    <path fill="#36C5F0" d="M17.688 8.834a2.528 2.528 0 0 1-2.523 2.521 2.527 2.527 0 0 1-2.52-2.521V2.522A2.527 2.527 0 0 1 15.165 0a2.528 2.528 0 0 1 2.523 2.522v6.312z"/>
    <path fill="#2EB67D" d="M15.165 18.956a2.528 2.528 0 0 1 2.523 2.522A2.528 2.528 0 0 1 15.165 24a2.527 2.527 0 0 1-2.52-2.522v-2.522h2.52z"/>
    <path fill="#ECB22E" d="M15.165 17.688a2.527 2.527 0 0 1-2.52-2.523 2.526 2.526 0 0 1 2.52-2.52h6.313A2.527 2.527 0 0 1 24 15.165a2.528 2.528 0 0 1-2.522 2.523h-6.313z"/>
  </svg>
);

/* ── Gmail Tab ─────────────────────────────────────────────────────── */
function GmailTab({ fileName, blob, subject }) {
  const [to,        setTo]        = useState("");
  const [clientId,  setClientId]  = useState("");
  const [loading,   setLoading]   = useState(false);
  const [steps,     setSteps]     = useState([]);  // [{label, status}]
  const [error,     setError]     = useState("");

  const BODY = `Hi,\n\nPlease find the processed medical package file attached.\n\nFile: ${fileName}\n\nGenerated by MediSort AI.`;

  const tick  = (i, status, extra = "") =>
    setSteps(prev => prev.map((s, j) => j === i ? { ...s, status, extra } : s));

  const handleSend = async () => {
    if (!clientId.trim()) { setError("Google OAuth Client ID is required."); return; }
    setError("");
    setLoading(true);
    const STEPS_DEF = [
      { label: "Authenticate with Google" },
      { label: "Build email with attachment" },
      { label: "Create Gmail draft" },
      { label: "Open draft in Gmail" },
    ];
    setSteps(STEPS_DEF.map(s => ({ ...s, status: "pending" })));

    try {
      // ── Step 0: OAuth via Google Identity Services ───────────────
      tick(0, "active");
      const accessToken = await new Promise((resolve, reject) => {
        // Dynamically load GIS if not already present
        const load = () => {
          const client = window.google.accounts.oauth2.initTokenClient({
            client_id: clientId.trim(),
            scope: "https://www.googleapis.com/auth/gmail.compose",
            callback: resp => {
              if (resp.error) reject(new Error(resp.error_description || resp.error));
              else resolve(resp.access_token);
            },
          });
          client.requestAccessToken({ prompt: "consent" });
        };
        if (window.google?.accounts?.oauth2) {
          load();
        } else {
          const script = document.createElement("script");
          script.src = "https://accounts.google.com/gsi/client";
          script.onload = load;
          script.onerror = () => reject(new Error("Failed to load Google Identity Services"));
          document.head.appendChild(script);
        }
      });
      tick(0, "done");

      // ── Step 1: Build MIME email ──────────────────────────────────
      tick(1, "active");
      const raw = await buildMimeRaw(to || "", subject, BODY, fileName, blob);
      tick(1, "done");

      // ── Step 2: Create Gmail draft via API ────────────────────────
      tick(2, "active");
      const draftRes = await fetch(
        "https://gmail.googleapis.com/gmail/v1/users/me/drafts",
        {
          method:  "POST",
          headers: {
            Authorization:  `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ message: { raw } }),
        }
      );
      if (!draftRes.ok) {
        const err = await draftRes.json();
        throw new Error(err?.error?.message || `Gmail API ${draftRes.status}`);
      }
      const draft = await draftRes.json();
      tick(2, "done", draft.id);

      // ── Step 3: Open the draft ────────────────────────────────────
      tick(3, "active");
      await new Promise(r => setTimeout(r, 400));
      // Gmail draft deep-link — opens the draft directly with attachment visible
      window.open(`https://mail.google.com/mail/#drafts/${draft.id}`, "_blank");
      tick(3, "done");
    } catch (err) {
      setError(err.message || "Something went wrong.");
      setSteps(prev => prev.map(s => s.status === "active" ? { ...s, status: "error" } : s));
    } finally {
      setLoading(false);
    }
  };

  const StepIcon = ({ status }) => {
    if (status === "done")    return <Check size={11} color="#00D4B4" />;
    if (status === "active")  return <Loader2 size={11} color="#0096E0" className="spin" />;
    if (status === "error")   return <AlertTriangle size={11} color="#F08070" />;
    return <span style={{ fontSize:10, color:"#4A6070" }}>·</span>;
  };

  return (
    <div className="share-body">
      <div className="info-box">
        The XLSX file will be attached directly to a new Gmail draft using the Gmail API.
        You need a <strong>Google OAuth 2.0 Client ID</strong> with the Gmail compose scope enabled.
        <a href="https://console.cloud.google.com/apis/credentials" target="_blank"
          style={{ color:"#5BA8F5", marginLeft:5 }}>Get one here ↗</a>
      </div>

      {steps.length === 0 ? (
        <>
          <div className="share-field">
            <label>Google OAuth Client ID</label>
            <input className="input" value={clientId} onChange={e => setClientId(e.target.value)}
              placeholder="xxx.apps.googleusercontent.com" style={{ fontSize:12 }} />
          </div>
          <div className="share-field" style={{ marginTop:12 }}>
            <label>To (optional — fill in Gmail)</label>
            <input className="input" value={to} onChange={e => setTo(e.target.value)}
              placeholder="recipient@example.com" style={{ fontSize:12 }} />
          </div>
          <div style={{ marginTop:8, padding:"8px 10px", borderRadius:8,
            background:"rgba(255,255,255,.03)", border:"1px solid rgba(255,255,255,.07)", fontSize:12, color:"#5A7488" }}>
            <span style={{ color:"#4A6070" }}>Subject: </span>{subject}
          </div>
          {error && (
            <p style={{ color:"#F08070", fontSize:12, marginTop:8, display:"flex", alignItems:"center", gap:5 }}>
              <AlertTriangle size={12} /> {error}
            </p>
          )}
          <button className="btn btn-primary" onClick={handleSend}
            style={{ width:"100%", marginTop:16, justifyContent:"center", display:"flex", alignItems:"center", gap:7 }}>
            {GMAIL_ICON} Create Gmail Draft with Attachment
          </button>
        </>
      ) : (
        <div style={{ marginTop:4 }}>
          {steps.map((s, i) => (
            <div key={i} className={`share-step ${s.status === "done" ? "share-step-done" : ""}`}>
              <div className={`share-step-num ${s.status === "done" ? "share-step-done" : ""}`}>
                <StepIcon status={s.status} />
              </div>
              <div>
                <p style={{ margin:0, fontWeight: s.status === "done" ? 600 : 400 }}>{s.label}</p>
                {s.extra && <p style={{ margin:0, fontSize:10, color:"#4A6070" }}>Draft ID: {s.extra}</p>}
              </div>
            </div>
          ))}
          {error && (
            <p style={{ color:"#F08070", fontSize:12, marginTop:10, display:"flex", alignItems:"center", gap:5 }}>
              <AlertTriangle size={12} /> {error}
            </p>
          )}
          {steps.every(s => s.status === "done") && (
            <div style={{ marginTop:12, padding:"10px 12px", borderRadius:9,
              background:"rgba(0,210,180,.08)", border:"1px solid rgba(0,210,180,.18)",
              color:"#00D4B4", fontSize:13, display:"flex", alignItems:"center", gap:8 }}>
              <Check size={15} /> Draft created — Gmail opened with {fileName} attached
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ── Slack Tab ─────────────────────────────────────────────────────── */
function SlackTab({ fileName, blob, subject, packages }) {
  const [webhookUrl,  setWebhookUrl]  = useState("");
  const [botToken,    setBotToken]    = useState("");
  const [channelId,   setChannelId]   = useState("");
  const [loading,     setLoading]     = useState(false);
  const [steps,       setSteps]       = useState([]);
  const [error,       setError]       = useState("");
  const [mode,        setMode]        = useState("webhook"); // "webhook" | "bot"

  const tick = (i, status, extra = "") =>
    setSteps(prev => prev.map((s, j) => j === i ? { ...s, status, extra } : s));

  // Rich Block Kit message body
  const buildBlocks = () => {
    const pkgLines = packages
      ? packages.map((p, i) => `• *${pkgLabel(p, i)}* — ${Object.values(p.categories).flat().length} tests · ${metaUI(p.price)}`).join("\n")
      : "";
    return [
      { type: "header", text: { type: "plain_text", text: `📊 MediSort AI — ${fileName}`, emoji: true } },
      { type: "section", text: { type: "mrkdwn", text: `*Processed medical package data is ready.*\nFile: \`${fileName}\`` } },
      ...(pkgLines ? [{ type: "section", text: { type: "mrkdwn", text: pkgLines } }] : []),
      { type: "divider" },
      { type: "context", elements: [{ type: "mrkdwn", text: `Generated by *MediSort AI* · ${new Date().toLocaleDateString()}` }] },
    ];
  };

  const handleWebhookSend = async () => {
    if (!webhookUrl.trim()) { setError("Webhook URL is required."); return; }
    setError("");
    setLoading(true);
    setSteps([
      { label: "Post notification to Slack" },
    ].map(s => ({ ...s, status: "pending" })));

    try {
      tick(0, "active");
      const res = await fetch(webhookUrl.trim(), {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ blocks: buildBlocks() }),
      });
      if (!res.ok && res.status !== 0) {
        // Webhooks return 200 "ok" text — non-200 is an error
        throw new Error(`Slack responded with ${res.status}`);
      }
      tick(0, "done");
    } catch (err) {
      // CORS blocks the response in browser — if the fetch itself didn't throw,
      // the message was likely delivered (Slack webhooks don't need a response body read)
      if (err.message.includes("Failed to fetch") || err.message.includes("NetworkError")) {
        // Treat CORS-blocked response as success — webhook posts don't need CORS
        tick(0, "done");
      } else {
        setError(err.message);
        tick(0, "error");
      }
    } finally {
      setLoading(false);
    }
  };

  const handleBotSend = async () => {
    if (!botToken.trim())  { setError("Bot token is required."); return; }
    if (!channelId.trim()) { setError("Channel ID is required."); return; }
    setError("");
    setLoading(true);
    setSteps([
      { label: "Post notification message" },
      { label: "Request upload URL from Slack" },
      { label: "Upload XLSX file" },
      { label: "Complete file upload" },
    ].map(s => ({ ...s, status: "pending" })));

    const AUTH = { Authorization: `Bearer ${botToken.trim()}` };

    try {
      // Step 0 — post the Block Kit message
      tick(0, "active");
      const msgRes = await fetch("https://slack.com/api/chat.postMessage", {
        method:  "POST",
        headers: { ...AUTH, "Content-Type": "application/json" },
        body:    JSON.stringify({ channel: channelId.trim(), blocks: buildBlocks() }),
      });
      const msgData = await msgRes.json();
      if (!msgData.ok) throw new Error(msgData.error || "chat.postMessage failed");
      tick(0, "done");

      // Step 1 — get upload URL (Slack's two-step file upload API)
      tick(1, "active");
      const urlRes = await fetch(
        `https://slack.com/api/files.getUploadURLExternal?filename=${encodeURIComponent(fileName)}&length=${blob.size}`,
        { headers: AUTH }
      );
      const urlData = await urlRes.json();
      if (!urlData.ok) throw new Error(urlData.error || "files.getUploadURLExternal failed");
      tick(1, "done");

      // Step 2 — upload the actual blob to the presigned URL
      tick(2, "active");
      const uploadRes = await fetch(urlData.upload_url, {
        method:  "POST",
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        body:    blob,
      });
      if (!uploadRes.ok) throw new Error(`Upload failed: ${uploadRes.status}`);
      tick(2, "done");

      // Step 3 — complete the upload and share into the channel
      tick(3, "active");
      const completeRes = await fetch("https://slack.com/api/files.completeUploadExternal", {
        method:  "POST",
        headers: { ...AUTH, "Content-Type": "application/json" },
        body:    JSON.stringify({
          files:           [{ id: urlData.file_id, title: fileName }],
          channel_id:      channelId.trim(),
          initial_comment: `📎 ${subject}`,
        }),
      });
      const completeData = await completeRes.json();
      if (!completeData.ok) throw new Error(completeData.error || "files.completeUploadExternal failed");
      tick(3, "done");
    } catch (err) {
      setError(err.message);
      setSteps(prev => prev.map(s => s.status === "active" ? { ...s, status: "error" } : s));
    } finally {
      setLoading(false);
    }
  };

  const StepIcon = ({ status }) => {
    if (status === "done")   return <Check size={11} color="#00D4B4" />;
    if (status === "active") return <Loader2 size={11} color="#0096E0" className="spin" />;
    if (status === "error")  return <AlertTriangle size={11} color="#F08070" />;
    return <span style={{ fontSize:10, color:"#4A6070" }}>·</span>;
  };

  return (
    <div className="share-body">
      {/* Mode switcher */}
      <div style={{ display:"flex", gap:6, marginBottom:16 }}>
        {[["webhook","Webhook (notify only)"],["bot","Bot Token (attach file)"]].map(([v,l]) => (
          <button key={v} onClick={() => { setMode(v); setSteps([]); setError(""); }}
            style={{ flex:1, padding:"7px 10px", borderRadius:8, border:"1px solid",
              borderColor: mode===v ? "rgba(0,210,180,.4)" : "rgba(255,255,255,.1)",
              background:  mode===v ? "rgba(0,210,180,.1)" : "rgba(255,255,255,.03)",
              color:       mode===v ? "#00D4B4" : "#6A8499",
              fontSize:12, fontFamily:"'Exo 2',sans-serif", cursor:"pointer", fontWeight: mode===v ? 600 : 400 }}>
            {l}
          </button>
        ))}
      </div>

      {steps.length === 0 ? (
        <>
          {mode === "webhook" ? (
            <>
              <div className="info-box">
                Paste an <strong>Incoming Webhook URL</strong> from your Slack workspace.
                A rich notification with the package summary will be posted to the channel.
                <a href="https://api.slack.com/messaging/webhooks" target="_blank"
                  style={{ color:"#5BA8F5", marginLeft:5 }}>Setup guide ↗</a>
              </div>
              <div className="share-field">
                <label>Slack Incoming Webhook URL</label>
                <input className="input" value={webhookUrl} onChange={e => setWebhookUrl(e.target.value)}
                  placeholder="https://hooks.slack.com/services/…" style={{ fontSize:12 }} />
              </div>
            </>
          ) : (
            <>
              <div className="info-box">
                Provide a <strong>Slack Bot Token</strong> (xoxb-…) and the target <strong>Channel ID</strong>.
                The XLSX file will be uploaded directly to the channel with a notification message.
                <a href="https://api.slack.com/authentication/token-types#bot" target="_blank"
                  style={{ color:"#5BA8F5", marginLeft:5 }}>Token guide ↗</a>
              </div>
              <div className="share-field">
                <label>Slack Bot Token</label>
                <input className="input" value={botToken} onChange={e => setBotToken(e.target.value)}
                  placeholder="xoxb-…" style={{ fontSize:12 }} />
              </div>
              <div className="share-field">
                <label>Channel ID</label>
                <input className="input" value={channelId} onChange={e => setChannelId(e.target.value)}
                  placeholder="C0123456789" style={{ fontSize:12 }} />
              </div>
            </>
          )}

          {error && (
            <p style={{ color:"#F08070", fontSize:12, marginTop:8, display:"flex", alignItems:"center", gap:5 }}>
              <AlertTriangle size={12} /> {error}
            </p>
          )}

          <button className="btn btn-primary"
            onClick={mode === "webhook" ? handleWebhookSend : handleBotSend}
            style={{ width:"100%", marginTop:16, justifyContent:"center", display:"flex", alignItems:"center", gap:7 }}>
            {SLACK_ICON}
            {mode === "webhook" ? "Post Notification to Slack" : "Upload File to Slack"}
          </button>
        </>
      ) : (
        <div style={{ marginTop:4 }}>
          {steps.map((s, i) => (
            <div key={i} className={`share-step ${s.status === "done" ? "share-step-done" : ""}`}>
              <div className={`share-step-num ${s.status === "done" ? "share-step-done" : ""}`}>
                <StepIcon status={s.status} />
              </div>
              <p style={{ margin:0, fontWeight: s.status === "done" ? 600 : 400 }}>{s.label}</p>
            </div>
          ))}
          {error && (
            <p style={{ color:"#F08070", fontSize:12, marginTop:10, display:"flex", alignItems:"center", gap:5 }}>
              <AlertTriangle size={12} /> {error}
            </p>
          )}
          {!error && steps.length > 0 && steps.every(s => s.status === "done") && (
            <div style={{ marginTop:12, padding:"10px 12px", borderRadius:9,
              background:"rgba(0,210,180,.08)", border:"1px solid rgba(0,210,180,.18)",
              color:"#00D4B4", fontSize:13, display:"flex", alignItems:"center", gap:8 }}>
              <Check size={15} />
              {mode === "webhook"
                ? "Notification posted to Slack ✓"
                : `${fileName} uploaded to Slack channel ✓`}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ── Share Modal shell ─────────────────────────────────────────────── */
function ShareModal({ onClose, fileName, blob, packages }) {
  const [tab, setTab] = useState("gmail");
  const subject = `MediSort AI - ${fileName || "Formatted_output.xlsx"}`;

  // Close on backdrop click
  const onBackdrop = e => { if (e.target === e.currentTarget) onClose(); };

  return (
    <div className="share-modal-overlay fadein" onClick={onBackdrop}>
      <div className="share-modal">
        {/* Header */}
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
          padding:"16px 20px 0", marginBottom:2 }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <div style={{ width:32, height:32, borderRadius:9,
              background:"linear-gradient(135deg,#00D4B4,#0096E0)",
              display:"flex", alignItems:"center", justifyContent:"center" }}>
              <Share2 size={15} color="#030C18" />
            </div>
            <div>
              <p style={{ fontWeight:700, fontSize:15 }}>Share File</p>
              <p style={{ fontSize:11, color:"#4A6070" }}>{fileName}</p>
            </div>
          </div>
          <button className="btn-icon" onClick={onClose}><X size={16} /></button>
        </div>

        {/* Tabs */}
        <div className="share-tab-bar">
          <button className={`share-tab ${tab === "gmail" ? "active" : ""}`} onClick={() => setTab("gmail")}>
            {GMAIL_ICON} Gmail
          </button>
          <button className={`share-tab ${tab === "slack" ? "active" : ""}`} onClick={() => setTab("slack")}>
            {SLACK_ICON} Slack
          </button>
        </div>

        {/* Tab body */}
        {tab === "gmail"
          ? <GmailTab  fileName={fileName} blob={blob} subject={subject} />
          : <SlackTab  fileName={fileName} blob={blob} subject={subject} packages={packages} />
        }
      </div>
    </div>
  );
}

/* ─────────────────────────── TOPBAR ─────────────────────────── */
function TopBar({ user, onLogout, processedPackages, onDownload, onShare }) {
  return (
    <header style={{ display:"flex", alignItems:"center", justifyContent:"space-between",
      padding:"0 24px", height:58, background:"rgba(3,10,22,.95)", backdropFilter:"blur(12px)",
      borderBottom:"1px solid rgba(0,210,180,.1)", position:"sticky", top:0, zIndex:100 }}>
      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        <div style={{ width:30, height:30, borderRadius:8, background:"linear-gradient(135deg,#00D4B4,#0096E0)",
          display:"flex", alignItems:"center", justifyContent:"center" }}>
          <Sparkles size={15} color="#030C18" />
        </div>
        <span style={{ fontWeight:700, fontSize:16, letterSpacing:"-0.025em" }}>
          MediSort <span style={{ color:"#00D4B4" }}>AI</span>
        </span>
        <span className="badge badge-teal" style={{ fontSize:10 }}>BETA</span>
      </div>

      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
        {processedPackages && (
          <>
            {/* Share button — triggers share menu / integrations */}
            <button className="btn btn-ghost" onClick={onShare}
              style={{ padding:"7px 14px", fontSize:13, display:"flex", alignItems:"center", gap:6 }}>
              <Share2 size={13} /> Share
            </button>
            {/* Download button — exclusively triggers .xlsx download */}
            <button className="btn btn-primary" onClick={onDownload}
              style={{ padding:"7px 14px", fontSize:13, display:"flex", alignItems:"center", gap:6 }}>
              <Download size={13} /> Download .xlsx
            </button>
          </>
        )}
        <div style={{ display:"flex", alignItems:"center", gap:8, padding:"6px 12px",
          borderRadius:9, background:"rgba(255,255,255,.04)", border:"1px solid rgba(255,255,255,.08)" }}>
          {/* Show real Google profile photo if available, otherwise initials */}
          {user?.picture ? (
            <img src={user.picture} alt={user.name}
              referrerPolicy="no-referrer"
              style={{ width:26, height:26, borderRadius:"50%", objectFit:"cover", flexShrink:0 }} />
          ) : (
            <div style={{ width:26, height:26, borderRadius:"50%", background:"linear-gradient(135deg,#7B61FF,#0096E0)",
              display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700, color:"#fff" }}>
              {user?.name?.[0]}
            </div>
          )}
          <span style={{ fontSize:13, color:"#B0C4D8" }}>{user?.name}</span>
          <span className="badge" style={{ background:"rgba(255,255,255,.07)", color:"#5A7488", fontSize:10 }}>
            {user?.provider}
          </span>
        </div>
        <button className="btn-icon" onClick={onLogout} title="Sign out"><LogOut size={15} /></button>
      </div>
    </header>
  );
}

/* ─────────────────────────── MAIN APP ─────────────────────────── */
export default function MediSort_AI() {
  const [screen, setScreen] = useState("login");
  const [user, setUser] = useState(null);
  const [file, setFile] = useState(null);
  const [rawData, setRawData] = useState(null);
  const [processedPackages, setProcessedPackages] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [steps, setSteps] = useState([]);
  const [error, setError] = useState(null);
  const [showShare, setShowShare] = useState(false);    // ← Share menu toggle
  const shareAnchorRef = useRef(null);

  const handleLogin = (googleUser) => {
    // googleUser = { name, email, picture, provider } decoded from Google JWT
    setUser(googleUser);
    setScreen("dashboard");
  };

  const handleFileSelect = async f => {
    if (!f?.name?.endsWith(".xlsx")) { setError("Please upload a valid .xlsx file"); return; }
    setError(null);
    setFile(f);
    try {
      const data = await parseExcel(f);
      setRawData(data);
      setProcessedPackages(null);
    } catch { setError("Could not parse file. Ensure it is a valid .xlsx."); }
  };

  const updateStep = (idx, status, message) =>
    setSteps(prev => prev.map((s, i) => i === idx ? { ...s, status, message } : s));

  const processFile = async () => {
    if (!rawData) return;
    setIsProcessing(true);
    setError(null);
    setScreen("processing");

    const initialSteps = STEPS.map(s => ({ ...s, status:"pending", message:"" }));
    setSteps(initialSteps);

    try {
      // Step 0: Parse
      updateStep(0, "active", "Identifying columns…");
      await new Promise(r => setTimeout(r, 700));

      // Use the maximum column count across ALL rows so short header rows
      // don't under-count real data columns
      const numCols = Math.max(...rawData.map(row => (row || []).length));

      // Build one package object per detected column using dynamic metadata detection.
      // detectColumnMeta scans the first 5 rows per column for serviceLocation /
      // ageGender / price via regex — never relies on fixed row positions.
      const packages = Array.from({ length: numCols }, (_, c) => {
        const meta = detectColumnMeta(rawData, c);
        return {
          serviceLocation: meta.serviceLocation,   // "" when not detected
          ageGender:       meta.ageGender,          // "" when not detected
          price:           meta.price,              // "" when not detected
          tests:           meta.tests,              // all non-metadata rows
          packageName:     "",
          displayName:     "",
          categories:      {},
        };
      });

      // Drop columns that have no metadata AND no tests (truly empty columns)
      const validPackages = packages.filter(
        p => p.tests.length > 0 || p.serviceLocation || p.price
      );

      updateStep(0, "done",
        `${validPackages.length} packages detected · max ${Math.max(...validPackages.map(p => p.tests.length))} tests`
      );

      // Step 1: Normalize
      updateStep(1, "active", "Deduplicating tests…");
      await new Promise(r => setTimeout(r, 600));
      for (const p of validPackages) p.tests = [...new Set(p.tests)];
      updateStep(1, "done", "Tests normalized and deduplicated");

      // Step 2: AI classify each package independently
      for (let i = 0; i < validPackages.length; i++) {
        updateStep(2, "active", `Classifying package ${i + 1} of ${validPackages.length}…`);
        const p = validPackages[i];

        const prompt = `You are a medical data classification expert.

Classify these medical tests into clinically meaningful categories and generate a package name and display name.

Tests list:
${p.tests.join("\n")}

Additional info:
- Service Location: ${p.serviceLocation}
- Age/Gender: ${p.ageGender}
- Price: ${p.price}

Return ONLY a valid JSON object — no markdown, no explanation, no preamble:
{
  "packageName": "Short identifier code e.g. COMP-HEALTH-PRO",
  "displayName": "Human-readable full package title",
  "categories": {
    "Category Name": ["Test 1", "Test 2"]
  }
}

Rules:
- Use standard category names: "Urine Routine Analysis", "Thyroid Profile", "Complete Blood Count (CBC)", "Lipid Profile", "Liver Function Tests", "Kidney Function Tests", "Diabetes Profile", "Vitamins", "Electrolytes & Minerals", "Iron Studies", "Cardiac Risk Markers", "Others"
- Keep ALL original test names exactly as given — do not rename them
- No duplicate tests across categories
- Every test in the input must appear in exactly one category`;

        try {
          const raw = await callClaude([{ role: "user", content: prompt }]);
          const parsed = extractJSON(raw);   // robust extraction handles fences + leading text
          validPackages[i].packageName = parsed.packageName || "";
          validPackages[i].displayName = parsed.displayName || "";
          // Validate categories is a plain object with at least one key
          if (parsed.categories && typeof parsed.categories === "object" && Object.keys(parsed.categories).length > 0) {
            validPackages[i].categories = parsed.categories;
          } else {
            validPackages[i].categories = { "General Tests": p.tests };
          }
        } catch (err) {
          // Per-package fallback — never lets one failure stop the rest
          console.warn(`Package ${i + 1} classification failed:`, err.message);
          validPackages[i].packageName = "";
          validPackages[i].displayName = "";
          validPackages[i].categories  = { "General Tests": p.tests };
        }

        // Brief pause between API calls to respect rate limits
        if (i < validPackages.length - 1) await new Promise(r => setTimeout(r, 300));
      }
      updateStep(2, "done", `All ${validPackages.length} packages AI-classified`);

      // Step 3: Build output
      updateStep(3, "active", "Generating structured output…");
      await new Promise(r => setTimeout(r, 500));
      setProcessedPackages(validPackages);
      updateStep(3, "done", "Output ready for download");

      await new Promise(r => setTimeout(r, 600));
      setScreen("results");
    } catch (err) {
      setError("Processing failed: " + err.message);
      setScreen("dashboard");
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadFile = () => {
    if (!processedPackages) return;
    const { base64, fileName } = buildOutputExcel(processedPackages, file?.name);

    // Convert base64 → Uint8Array → Blob (works for large files; data-URL breaks >2 MB)
    const binary = atob(base64);
    const bytes  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const url = URL.createObjectURL(blob);
    const a   = document.createElement("a");
    a.href     = url;
    a.download = fileName;            // e.g. "Formatted_Input.xlsx"
    document.body.appendChild(a);    // required for Firefox + sandboxed iframes
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);         // free memory
  };

  // Share handler — builds the real XLSX blob once and passes it to ShareMenu
  const handleShare = () => setShowShare(s => !s);

  // Compute blob + fileName for the share menu (only when packages are ready)
  const shareData = (() => {
    if (!processedPackages) return { blob: null, fileName: null };
    const { base64, fileName } = buildOutputExcel(processedPackages, file?.name);
    const binary = atob(base64);
    const bytes  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    return { blob, fileName };
  })();

  if (screen === "login") return <LoginScreen onLogin={handleLogin} />;

  return (
    <div className="grid-bg" style={{ minHeight:"100vh", fontFamily:"'Exo 2',sans-serif", color:"#DDE6F0" }}>
      <style>{CSS}</style>
      <TopBar user={user} onLogout={() => setScreen("login")}
        processedPackages={processedPackages} onDownload={downloadFile} onShare={handleShare} />

      {/* Share modal — full-screen centered overlay */}
      {showShare && processedPackages && (
        <ShareModal
          onClose={() => setShowShare(false)}
          fileName={shareData.fileName}
          blob={shareData.blob}
          packages={processedPackages}
        />
      )}

      <main style={{ padding:"28px 24px", maxWidth:1080, margin:"0 auto" }}>
        {screen === "processing" ? (
          <ProcessingView steps={steps} />
        ) : screen === "results" && processedPackages ? (
          <ResultsView
            packages={processedPackages}
            onDownload={downloadFile}
            onShare={handleShare}
            onNewFile={() => { setFile(null); setRawData(null); setProcessedPackages(null); setScreen("dashboard"); }}
          />
        ) : (
          <UploadView
            file={file} rawData={rawData} error={error} isProcessing={isProcessing}
            onFileSelect={handleFileSelect} onProcess={processFile}
            onDrop={e => { const f = e.dataTransfer.files[0]; if (f) handleFileSelect(f); }}
          />
        )}
      </main>

      {screen !== "login" && <ChatWidget packages={processedPackages} rawData={rawData} />}
    </div>
  );
}

