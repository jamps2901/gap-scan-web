/* global XLSX, Tabulator */

const MB51_REQUIRED = [
  "Plant", "Material", "Material Description", "Storage Location",
  "Movement Type", "Movement Type Text", "Posting Date", "Time of Entry",
  "Qty in unit of entry",
];

const SALES_MVTS = ["251", "601"];
const RECEIPT_MVTS = ["101"];
const COUNT_MVTS = ["701", "702"];

let mb51File = null;
let mb5bFile = null;
let resultRows = [];
let table = null;

const el = (id) => document.getElementById(id);

function setStatus(msg){ el("status").textContent = msg; }

function normIntish(x){
  if (x === null || x === undefined) return "";
  let s = String(x).trim();
  if (/^\d+\.0$/.test(s)) s = s.split(".")[0];
  return s.trim();
}

function zfill(s, n){
  s = String(s);
  return s.length >= n ? s : "0".repeat(n - s.length) + s;
}

function normMaterial(x, padTo=0){
  const s = normIntish(x);
  if (padTo && /^\d+$/.test(s)) return zfill(s, padTo);
  return s;
}

function normSloc(x, padTo=4){
  const s = normIntish(x);
  if (padTo && /^\d+$/.test(s)) return zfill(s, padTo);
  return s;
}

function makeKey(plant, material, sloc){
  return `${plant}|${material}|${sloc}`;
}

function safeNumber(x){
  const n = Number(String(x).replace(/,/g,"").trim());
  return Number.isFinite(n) ? n : 0.0;
}

/** Excel date (serial) -> JS Date */
function excelSerialToDate(n){
  // Excel 1900 date system; SheetJS commonly uses this conversion
  const utcDays = Math.floor(n - 25569);
  const utcValue = utcDays * 86400; // seconds
  const dateInfo = new Date(utcValue * 1000);
  const fractional = n - Math.floor(n);
  const seconds = Math.round(fractional * 86400);
  return new Date(dateInfo.getTime() + seconds * 1000);
}

function parseDateCell(v){
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number") {
    const d = excelSerialToDate(v);
    return isNaN(d) ? null : d;
  }
  const s = String(v).trim();
  // Try Date(...) parsing
  const d = new Date(s);
  if (!isNaN(d)) return d;
  return null;
}

function parseTimeToMs(v){
  if (v === null || v === undefined || v === "") return 0;
  // Excel time might be fraction of day (0..1)
  if (typeof v === "number") {
    return Math.round(v * 86400 * 1000);
  }
  const s = String(v).trim();
  // Accept HH:MM:SS or HH:MM
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m){
    const hh = Number(m[1] || 0);
    const mm = Number(m[2] || 0);
    const ss = Number(m[3] || 0);
    return ((hh*3600 + mm*60 + ss) * 1000);
  }
  // Fallback: try Date parse for times
  const d = new Date(`1970-01-01T${s}Z`);
  if (!isNaN(d)) return d.getUTCHours()*3600000 + d.getUTCMinutes()*60000 + d.getUTCSeconds()*1000;
  return 0;
}

function daysSince(dt){
  if (!dt) return NaN;
  const ms = Date.now() - dt.getTime();
  return ms / 86400000.0;
}

/** Read file (xlsx/xls/csv) to rows of objects (for tabular sheets) */
async function readWorkbookAsObjects(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  // Convert to JSON objects with header row
  return XLSX.utils.sheet_to_json(ws, { defval:"" });
}

/** Read file to rows as arrays (for MB5B block parsing) */
async function readWorkbookAsArrays(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header:1, defval:"" }); // array of arrays
}

/** MB51 cleaning */
function cleanMb51(rows, matPad, slocPad){
  // Validate headers
  const cols = new Set(Object.keys(rows[0] || {}).map(c => String(c).trim()));
  const missing = MB51_REQUIRED.filter(c => !cols.has(c));
  if (missing.length) throw new Error(`MB51 missing columns: ${missing.join(", ")}`);

  const out = rows.map(r => {
    const plant = normIntish(r["Plant"]);
    const material = normMaterial(r["Material"], matPad);
    const sloc = normSloc(r["Storage Location"], slocPad);

    const postDate = parseDateCell(r["Posting Date"]);
    const timeMs = parseTimeToMs(r["Time of Entry"]);
    const postDT = postDate ? new Date(postDate.getTime() + timeMs) : null;

    const mvt = String(r["Movement Type"]).trim();
    const mvtTxt = String(r["Movement Type Text"]).trim();
    const qty = safeNumber(r["Qty in unit of entry"]);

    const key = makeKey(plant, material, sloc);

    return {
      Plant: plant,
      Material: material,
      "Material Description": String(r["Material Description"] ?? "").trim(),
      "Storage Location": sloc,
      "Movement Type": mvt,
      "Movement Type Text": mvtTxt,
      "Posting Date": postDate,
      "Time of Entry": r["Time of Entry"],
      "Post DateTime": postDT,
      Qty: qty,
      Key: key,
    };
  });

  out.sort((a,b) => {
    if (a.Key < b.Key) return -1;
    if (a.Key > b.Key) return 1;
    const at = a["Post DateTime"] ? a["Post DateTime"].getTime() : 0;
    const bt = b["Post DateTime"] ? b["Post DateTime"].getTime() : 0;
    return at - bt;
  });

  return out;
}

/** MB5B parsing (block/detail export) */
function parseQtyLine(prefix, s){
  s = (s === null || s === undefined) ? "" : String(s);
  if (!s.trim().startsWith(prefix)) return null;

  // Parse only after prefix (avoid parsing date inside prefix)
  const tail = s.trim().slice(prefix.length).trim();
  const m = tail.match(/([0-9]+(?:\.[0-9]+)?)\s*(-)?\s*(EA|PC|ST|KG|L|)\b/i);
  if (!m) return null;
  let num = Number(m[1]);
  if (m[2] === "-") num = -num;
  return Number.isFinite(num) ? num : null;
}

function parseMb5bBlocks(rows2d, matPad, slocPad){
  const col0 = rows2d.map(r => String((r && r[0] !== undefined) ? r[0] : "").trim());
  const plantRows = [];
  for (let i=0; i<col0.length; i++){
    if (col0[i].startsWith("Plant")) plantRows.push(i);
  }
  if (!plantRows.length) throw new Error("MB5B format not recognized (no 'Plant' blocks found). Export MB5B in detail/block format.");

  const blocks = [];

  for (let bi=0; bi<plantRows.length; bi++){
    const start = plantRows[bi];
    const end = (bi+1 < plantRows.length) ? plantRows[bi+1] : rows2d.length;
    const block = rows2d.slice(start, end);

    const lines0 = block
      .map(r => r && r[0] !== undefined ? String(r[0]) : "")
      .filter(x => String(x).trim().length)
      .map(x => String(x));

    if (!lines0.length) continue;

    const plantLine = lines0[0];
    const pm = plantLine.match(/Plant\s+(\d+)/);
    const plant = pm ? pm[1] : "";

    let material = "";
    let desc = "";
    let closeQty = null;

    for (let i=0; i<Math.min(lines0.length, 140); i++){
      const line = lines0[i];

      if (!material && line.includes("Material")){
        const mm = line.match(/Material\s+(\d+)/);
        if (mm) material = mm[1];
      }
      if (!desc && line.trim().startsWith("Description")){
        const dm = line.match(/Description\s+(.*)$/);
        if (dm) desc = (dm[1] || "").trim();
      }
      if (closeQty === null){
        const q = parseQtyLine("Stock on 31.12.9999", line);
        if (q !== null) closeQty = q;
      }
    }

    // infer SLoc from detail table if possible: find row where col1="Loca" col2="MvT"
    const slocs = new Set();
    for (let r=0; r<block.length; r++){
      const c1 = block[r]?.[1] ?? "";
      const c2 = block[r]?.[2] ?? "";
      if (String(c1).trim() === "Loca" && String(c2).trim() === "MvT"){
        for (let rr=r+2; rr<block.length; rr++){
          const v0 = block[rr]?.[0] ?? "";
          if (typeof v0 === "string" && String(v0).trim().startsWith("Plant")) break;
          const loca = block[rr]?.[1];
          if (loca !== null && loca !== undefined && String(loca).trim() !== ""){
            slocs.add(normSloc(loca, slocPad));
          }
        }
        break;
      }
    }

    let sloc = "";
    if (slocs.size === 1) sloc = [...slocs][0];
    else if (slocs.size > 1) sloc = "MULTI";

    const plantN = normIntish(plant);
    const matN = normMaterial(material, matPad);
    const slocN = (sloc === "MULTI") ? "MULTI" : normSloc(sloc, slocPad);

    blocks.push({
      Plant: plantN,
      Material: matN,
      "Storage Location": slocN,
      "Material Description": desc,
      SAP_SOH_MB5B: (closeQty !== null) ? Number(closeQty) : 0.0,
      Key: makeKey(plantN, matN, slocN),
    });
  }

  // drop duplicates by Key
  const seen = new Set();
  const out = [];
  for (const b of blocks){
    if (seen.has(b.Key)) continue;
    seen.add(b.Key);
    out.push(b);
  }
  return out;
}

/** Group sum Qty per Key */
function computeExpectedSohMb51(mb51){
  const map = new Map();
  for (const r of mb51){
    map.set(r.Key, (map.get(r.Key) || 0) + (r.Qty || 0));
  }
  return map;
}

function lastEvent(mb51, mvts, label){
  const last = new Map(); // Key -> row
  const mvSet = new Set(mvts);
  for (const r of mb51){
    if (!mvSet.has(r["Movement Type"])) continue;
    if (!r["Post DateTime"]) continue;
    const cur = last.get(r.Key);
    if (!cur || cur["Post DateTime"].getTime() < r["Post DateTime"].getTime()){
      last.set(r.Key, r);
    }
  }
  // convert to map Key -> fields
  const out = new Map();
  for (const [key, r] of last.entries()){
    out.set(key, {
      [`${label}_DT`]: r["Post DateTime"],
      [`${label}_Qty`]: r.Qty,
      [`${label}_MvT`]: r["Movement Type"],
      [`${label}_Txt`]: r["Movement Type Text"],
    });
  }
  return out;
}

function expectationAndReason(row, tol){
  const sap = Number(row.SAP_SOH_MB5B || 0);
  const exp = Number(row.Expected_SOH_MB51 || 0);
  const delta = Number(row.Delta_SAP_minus_Expected || 0);

  const dCount = row.Days_Since_LastCount;
  const dRec = row.Days_Since_LastReceipt;
  const dSale = row.Days_Since_LastSale;

  const loss702 = Number(row.Loss702_Sum || 0);

  const reasons = [];

  if (sap <= 0){
    return { Expectation:"N/A", Summary:"SAP shows 0 on-hand (no stock expected)." };
  }

  let base = "MEDIUM";
  if (Math.abs(delta) <= tol){
    reasons.push("SAP SOH matches movement replay (consistent).");
    base = "HIGH";
  } else {
    reasons.push(`SAP SOH differs from movement replay by ${delta.toFixed(2)} (mismatch).`);
    base = (Math.abs(delta) <= 5) ? "MEDIUM" : "LOW";
  }

  if (Number.isFinite(dRec) && dRec <= 14){
    reasons.push("Recent receipt → stock likely exists somewhere (backroom possible).");
    if (base === "MEDIUM") base = "HIGH";
  } else if (!Number.isFinite(dRec) || dRec > 90){
    reasons.push("No recent receipts → less likely to be in backroom.");
    if (base === "HIGH") base = "MEDIUM";
  }

  if (Number.isFinite(dSale) && dSale <= 14){
    reasons.push("Recent sales → item is active (stock movement ongoing).");
  }

  if (loss702 < 0){
    reasons.push(`702 loss history (${loss702.toFixed(2)}) → higher chance of shrink / missing stock.`);
    if (base === "HIGH") base = "MEDIUM";
    else if (base === "MEDIUM") base = "LOW";
  }

  if (!Number.isFinite(dCount)){
    reasons.push("No 701/702 count event found → confidence weaker.");
    if (base === "HIGH") base = "MEDIUM";
  } else if (dCount > 180){
    reasons.push("Last count is old → more uncertainty.");
    if (base === "HIGH") base = "MEDIUM";
  }

  return { Expectation: base, Summary: reasons.slice(0,3).join(" ") };
}

function formatDate(d){
  if (!d) return "";
  try{
    return d.toISOString().replace("T"," ").slice(0,19);
  }catch{
    return String(d);
  }
}

function buildDetailsText(r){
  const lines = [];
  lines.push(`KEY: ${r.Key}`);
  lines.push(`Material: ${r.Material}  |  SLoc: ${r["Storage Location"]}  |  Plant: ${r.Plant}`);
  lines.push(`Description: ${r["Material Description"] || ""}`);
  lines.push("");
  lines.push("WHAT SAP SAYS NOW:");
  lines.push(`  SAP SOH (from MB5B): ${Number(r.SAP_SOH_MB5B || 0).toFixed(2)}`);
  lines.push("");
  lines.push("WHAT MOVEMENT HISTORY IMPLIES:");
  lines.push(`  Expected SOH (from MB51 replay): ${Number(r.Expected_SOH_MB51 || 0).toFixed(2)}`);
  lines.push(`  Difference (SAP - Expected): ${Number(r.Delta_SAP_minus_Expected || 0).toFixed(2)}`);
  lines.push("");
  lines.push("SHOULD IT BE THERE (WITHOUT COUNTING)?");
  lines.push(`  Expectation: ${r.Expectation}`);
  lines.push(`  Why: ${r.Summary || ""}`);
  lines.push("");
  lines.push("RECENT CHECKPOINTS (helpful for gap-scan judgement):");
  lines.push(`  Last count adjustment (701/702): ${formatDate(r.LastCount_DT)}  | Qty: ${r.LastCount_Qty ?? ""}`);
  lines.push(`  Last sale movement (251/601): ${formatDate(r.LastSale_DT)}  | Qty: ${r.LastSale_Qty ?? ""}`);
  lines.push(`  Last receipt (101): ${formatDate(r.LastReceipt_DT)}  | Qty: ${r.LastReceipt_Qty ?? ""}`);
  lines.push(`  702 loss total: ${Number(r.Loss702_Sum || 0).toFixed(2)}`);
  lines.push("");
  lines.push("HOW TO USE THIS IN YOUR PROCESS:");
  lines.push("  - If shelf is empty AND Expectation is HIGH → likely exists somewhere (backroom check first).");
  lines.push("  - If shelf is empty AND Expectation is LOW → SAP may be overstated; consider COUNT before write-off.");
  lines.push("  - If SAP SOH is small (1–2) → low exposure; policy/value can drive whether you count or write off.");
  return lines.join("\n");
}

function initTabs(){
  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tabpane").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      const which = btn.dataset.tab;
      el(`tab-${which}`).classList.add("active");
    });
  });
}

function initTable(){
  table = new Tabulator("#table", {
    height: "620px",
    layout: "fitColumns",
    placeholder: "Run analysis to see results.",
    data: [],
    rowFormatter: function(row){
      const d = row.getData();
      row.getElement().classList.remove("row-high","row-medium","row-low","row-na");
      const exp = String(d.Expectation || "").toUpperCase();
      if (exp === "HIGH") row.getElement().classList.add("row-high");
      else if (exp === "MEDIUM") row.getElement().classList.add("row-medium");
      else if (exp === "LOW") row.getElement().classList.add("row-low");
      else row.getElement().classList.add("row-na");
    },
    columns: [
      { title:"Plant", field:"Plant", width:110 },
      { title:"Material", field:"Material", width:160 },
      { title:"Storage Location", field:"Storage Location", width:160 },
      { title:"Material Description", field:"Material Description", width:340 },
      { title:"SAP_SOH_MB5B", field:"SAP_SOH_MB5B", hozAlign:"right", width:140, formatter:(c)=>Number(c.getValue()||0).toFixed(2) },
      { title:"Expected_SOH_MB51", field:"Expected_SOH_MB51", hozAlign:"right", width:170, formatter:(c)=>Number(c.getValue()||0).toFixed(2) },
      { title:"Delta_SAP_minus_Expected", field:"Delta_SAP_minus_Expected", hozAlign:"right", width:210, formatter:(c)=>Number(c.getValue()||0).toFixed(2) },
      { title:"Expectation", field:"Expectation", width:120, hozAlign:"center" },
      { title:"Summary", field:"Summary", widthGrow:3, tooltip:true },
    ],
    rowClick: function(e, row){
      const d = row.getData();
      el("keySelect").value = d.Key;
      el("detailText").textContent = buildDetailsText(d);
      // jump to Details tab
      document.querySelector('.tab[data-tab="detail"]').click();
    }
  });
}

function populateKeySelect(rows){
  const sel = el("keySelect");
  sel.innerHTML = "";
  for (const r of rows){
    const opt = document.createElement("option");
    opt.value = r.Key;
    opt.textContent = r.Key;
    sel.appendChild(opt);
  }
  sel.addEventListener("change", () => {
    const k = sel.value;
    const r = resultRows.find(x => x.Key === k);
    el("detailText").textContent = r ? buildDetailsText(r) : "";
  });
}

async function runAnalysis(){
  if (!mb51File || !mb5bFile){
    setStatus("Please load both MB51 and MB5B first.");
    return;
  }

  try{
    el("exportBtn").disabled = true;
    setStatus("Reading files...");

    const matPad = Number(el("matPad").value);
    const slocPad = Number(el("slocPad").value);
    const tol = Number(el("tol").value);

    // MB51
    const mb51Raw = await readWorkbookAsObjects(mb51File);
    const mb51 = cleanMb51(mb51Raw, matPad, slocPad);

    // MB5B
    const mb5b2d = await readWorkbookAsArrays(mb5bFile);
    const mb5b = parseMb5bBlocks(mb5b2d, matPad, slocPad);

    setStatus("Computing metrics...");

    const expectedMap = computeExpectedSohMb51(mb51);

    const lastCnt = lastEvent(mb51, COUNT_MVTS, "LastCount");
    const lastSale = lastEvent(mb51, SALES_MVTS, "LastSale");
    const lastRec = lastEvent(mb51, RECEIPT_MVTS, "LastReceipt");

    // Loss 702 sum
    const loss702Map = new Map();
    for (const r of mb51){
      if (r["Movement Type"] !== "702") continue;
      loss702Map.set(r.Key, (loss702Map.get(r.Key) || 0) + (r.Qty || 0));
    }

    const rows = mb5b.map(b => {
      const expected = expectedMap.get(b.Key) || 0.0;
      const delta = (b.SAP_SOH_MB5B || 0) - expected;

      const row = {
        ...b,
        Expected_SOH_MB51: expected,
        Delta_SAP_minus_Expected: delta,
      };

      Object.assign(row, lastCnt.get(b.Key) || {});
      Object.assign(row, lastSale.get(b.Key) || {});
      Object.assign(row, lastRec.get(b.Key) || {});

      row.Days_Since_LastCount = row.LastCount_DT ? daysSince(row.LastCount_DT) : NaN;
      row.Days_Since_LastSale = row.LastSale_DT ? daysSince(row.LastSale_DT) : NaN;
      row.Days_Since_LastReceipt = row.LastReceipt_DT ? daysSince(row.LastReceipt_DT) : NaN;
      row.Loss702_Sum = loss702Map.get(b.Key) || 0.0;

      const expRes = expectationAndReason(row, tol);
      row.Expectation = expRes.Expectation;
      row.Summary = expRes.Summary;

      return row;
    });

    // Sort: SAP stock > 0 first, then LOW first, then SAP desc
    const order = { "LOW":0, "MEDIUM":1, "HIGH":2, "N/A":3 };
    rows.sort((a,b) => {
      const ap = (a.SAP_SOH_MB5B > 0) ? 1 : 0;
      const bp = (b.SAP_SOH_MB5B > 0) ? 1 : 0;
      if (ap !== bp) return bp - ap;

      const ar = order[a.Expectation] ?? 9;
      const br = order[b.Expectation] ?? 9;
      if (ar !== br) return ar - br;

      return (b.SAP_SOH_MB5B || 0) - (a.SAP_SOH_MB5B || 0);
    });

    resultRows = rows;

    table.setData(rows);

    populateKeySelect(rows);
    if (rows.length){
      el("keySelect").value = rows[0].Key;
      el("detailText").textContent = buildDetailsText(rows[0]);
    }

    const mb51Keys = new Set(mb51.map(r => r.Key));
    const mb5bKeys = new Set(mb5b.map(r => r.Key));
    let overlap = 0;
    for (const k of mb51Keys) if (mb5bKeys.has(k)) overlap++;

    el("exportBtn").disabled = rows.length === 0;
    setStatus(`Done. Keys matched between MB51 and MB5B: ${overlap}`);
  }catch(err){
    console.error(err);
    setStatus(`Error: ${err.message || err}`);
    alert(`Error:\n\n${err.message || err}`);
  }
}

function exportExcel(){
  if (!resultRows.length) return;

  const exportCols = [
    "Plant", "Material", "Storage Location", "Material Description",
    "SAP_SOH_MB5B", "Expected_SOH_MB51", "Delta_SAP_minus_Expected",
    "Expectation", "Summary",
    "LastCount_DT", "LastCount_Qty",
    "LastSale_DT", "LastSale_Qty",
    "LastReceipt_DT", "LastReceipt_Qty",
    "Loss702_Sum"
  ];

  const out = resultRows.map(r => {
    const obj = {};
    for (const c of exportCols){
      if (c.endsWith("_DT")) obj[c] = r[c] ? formatDate(r[c]) : "";
      else obj[c] = (r[c] ?? "");
    }
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "GapScan");
  XLSX.writeFile(wb, "gap_scan_results.xlsx");
}

function init(){
  initTabs();
  initTable();

  el("mb51File").addEventListener("change", (e) => {
    mb51File = e.target.files?.[0] || null;
    el("mb51Hint").textContent = mb51File ? `Loaded: ${mb51File.name}` : "No file loaded.";
    setStatus("MB51 selected. Load MB5B and run analysis.");
  });

  el("mb5bFile").addEventListener("change", (e) => {
    mb5bFile = e.target.files?.[0] || null;
    el("mb5bHint").textContent = mb5bFile ? `Loaded: ${mb5bFile.name}` : "No file loaded.";
    setStatus("MB5B selected. Load MB51 and run analysis.");
  });

  el("runBtn").addEventListener("click", runAnalysis);
  el("exportBtn").addEventListener("click", exportExcel);
}

init();
