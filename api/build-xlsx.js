// api/build-xlsx.js
const ExcelJS = require("exceljs");
const Papa = require("papaparse");
const fetch = require("node-fetch");
const { PassThrough } = require("stream");

/* -------------------- Node-safe: workbook -> Buffer -------------------- */
function toBuffer(workbook) {
  return new Promise((resolve, reject) => {
    const stream = new PassThrough();
    const chunks = [];
    stream.on("data", c => chunks.push(c));
    stream.on("end", () => resolve(Buffer.concat(chunks)));
    stream.on("error", reject);
    workbook.xlsx.write(stream).then(() => stream.end()).catch(reject);
  });
}

/* -------------------- HTTP/inline CSV reader -------------------- */
async function readCsvText({ url, csv, headers }) {
  if (csv && typeof csv === "string") return csv;
  if (!url) return "";
  const r = await fetch(url, { headers });
  if (!r.ok) throw new Error(`Failed to fetch CSV: ${url} (status ${r.status})`);
  return await r.text();
}

/* -------------------- Config & styles -------------------- */
const SHEET_DET = "Detaljeret oversigt";
const SHEET_SAP = "SAP";

const TITLE_ROW_INDEX    = 1; // Periode… (merged)
const SUBTITLE_ROW_INDEX = 2; // optional (merged)
const BLANK_ROW_INDEX    = 3; // reserved
const HEADER_ROW_INDEX   = 4;

const BORDER = { style: "thin", color: { argb: "FF000000" } };

const SAP_HEADERS = [
  "Kontrakt","Position","Artskonto","Besrkivelse","Profitcenter",
  "Pris inkl. moms","Momskode","Pris ex. moms","Lokation/rute","Kunde"
];
const SAP_WIDTHS = [9.0, 8.43, 9.71, 34.43, 11.86, 14.86, 10.86, 13.71, 19.29, 33.57];

const DET_BASE_HEADERS = [
  "Øko-ID","Vintercentral","Område","Distrikt","Rute/lokation","Lokationsnavn","Evt. id",
  "Lokationsadresse","Lokationspostnr","I alt, kr","Saltning, kr","Salt, antal","Salt, gns. pris",
  "Kombi, kr","Kombi, antal","Kombi, gns. pris","Snerydning, kr","Sne, antal","Sne, gns. pris",
  "Andet, kr","Andet, antal","Andet, gns. pris","Salt, kg"
];
const DET_BASE_WIDTHS = [
  9.0, 12.71, 15.86, 7.71, 19.29, 33.57, 6.57, 19.29, 15.71,
  9.14, 11.0, 9.86, 13.14, 9.43, 12.14, 15.43, 14.0, 9.86, 13.14,
  9.0, 11.71, 15.0, 7.43
];
const DAY_COL_WIDTH = 3.0;

/* -------------------- Numeric coercion -------------------- */
const DET_NUMERIC_HEADERS = new Set([
  "I alt, kr","Saltning, kr","Salt, antal","Salt, gns. pris",
  "Kombi, kr","Kombi, antal","Kombi, gns. pris",
  "Snerydning, kr","Sne, antal","Sne, gns. pris",
  "Andet, kr","Andet, antal","Andet, gns. pris","Salt, kg"
]);
const SAP_NUMERIC_HEADERS = new Set([
  "Pris inkl. moms","Pris ex. moms"
]);

/* Danish/intl-friendly number parsing */
function toNumberMaybe(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  let s = String(v).trim();
  if (!s) return null;

  s = s.replace(/\s+/g, "");
  // 1.234,56 -> 1234.56
  if (/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
  // 1234,56 -> 1234.56
  else if (/^\d+,\d+$/.test(s)) s = s.replace(",", ".");
  // 1,234,567.89 -> 1234567.89
  if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, "");

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/* Keep day columns (1–31) as text; coerce only known numeric fields */
function coerceByHeader(header, value) {
  const h = (header || "").trim();
  if (/^\d{1,2}$/.test(h)) return value === "" ? null : String(value); // day columns -> text
  if (DET_NUMERIC_HEADERS.has(h) || SAP_NUMERIC_HEADERS.has(h)) return toNumberMaybe(value);
  return value === "" ? null : value;
}

/* -------------------- Helpers -------------------- */
const norm = s => (s ?? "").toString().trim().toLowerCase();
const stripComma = s => norm(s).replace(/,/g, "");

function todayDk() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

function setTitle(ws, text, totalCols) {
  if (!text) return;
  const c = ws.getCell(TITLE_ROW_INDEX, 1);
  c.value = text;
  c.font = { bold: true, size: 12 };
  ws.mergeCells(TITLE_ROW_INDEX, 1, TITLE_ROW_INDEX, totalCols);
}

function setSubtitle(ws, text, totalCols) {
  if (!text) return;
  const c = ws.getCell(SUBTITLE_ROW_INDEX, 1);
  c.value = text;
  c.font = { italic: true, size: 10, color: { argb: "FF333333" } };
  ws.mergeCells(SUBTITLE_ROW_INDEX, 1, SUBTITLE_ROW_INDEX, totalCols);
}

function writeHeader(ws, labels) {
  const r = ws.getRow(HEADER_ROW_INDEX);
  labels.forEach((label, i) => {
    const cell = r.getCell(i + 1);
    cell.value = label;
    cell.font = { bold: true };
    cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    cell.border = { top: BORDER, left: BORDER, right: BORDER, bottom: BORDER };
  });
  r.height = 18;
  r.commit();
}

function addBorders(ws, startRow, endRow, startCol, endCol) {
  for (let r = startRow; r <= endRow; r++) {
    const row = ws.getRow(r);
    for (let c = startCol; c <= endCol; c++) {
      row.getCell(c).border = { top: BORDER, left: BORDER, right: BORDER, bottom: BORDER };
    }
    row.commit();
  }
}

function writeFooter(ws, startCol, endCol, afterRow, extractedAt) {
  const row = (afterRow || HEADER_ROW_INDEX) + 2;
  const cell = ws.getCell(row, startCol);
  cell.value = `Udtrukket ${extractedAt}`;
  cell.font = { italic: true, size: 9, color: { argb: "FF444444" } };
  cell.alignment = { horizontal: "left", vertical: "middle" };
  ws.mergeCells(row, startCol, row, endCol);
}

function numberToColumn(n) {
  let s = "";
  while (n > 0) { const t = (n - 1) % 26; s = String.fromCharCode(65 + t) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function writeSumFormulas(ws, totalRow, colIndexes, startRow, endRow) {
  colIndexes.forEach(ci => {
    const colLetter = numberToColumn(ci);
    const cell = ws.getCell(totalRow, ci);
    cell.value = { formula: `SUM(${colLetter}${startRow}:${colLetter}${endRow})` };
    cell.font = { bold: true };
  });
  const maxCol = Math.max(...colIndexes);
  for (let c = 1; c <= maxCol; c++) {
    ws.getCell(totalRow, c).border = { top: BORDER, left: BORDER, right: BORDER, bottom: BORDER };
  }
}

/* Loose header lookup (case/commas) */
function getValueForHeader(obj, hdr) {
  if (obj[hdr] !== undefined) return obj[hdr];
  const target = stripComma(hdr);
  const k = Object.keys(obj).find(key => norm(key) === norm(hdr) || stripComma(key) === target);
  return k ? obj[k] : null;
}

/* -------------------- CSV parsing without row dropping -------------------- */
function findHeaderIndex(rows, expectCandidates) {
  let bestIdx = -1, bestScore = -1;
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const r = rows[i].map(x => (x ?? "").toString().trim());
    const score = expectCandidates.reduce(
      (acc, h) => acc + (r.some(v => norm(v) === norm(h) || stripComma(v) === stripComma(h)) ? 1 : 0), 0
    );
    if (score > bestScore) { bestScore = score; bestIdx = i; }
  }
  return { idx: bestIdx, score: bestScore };
}

function parseWithDelimiter(raw, delimiter, expectedHeaders) {
  const parsed = Papa.parse(raw, { header: false, delimiter, skipEmptyLines: false });
  const rows = parsed.data.filter(Array.isArray);
  const { idx, score } = findHeaderIndex(rows.slice(0, 12), expectedHeaders);
  return { rows, headerIdx: idx, score };
}

function isNoiseRow(arr) {
  const first = (arr[0] ?? "").toString().trim().toLowerCase();
  const allEmpty = arr.every(v => (v ?? "").toString().trim() === "");
  if (allEmpty) return true;
  if (first.startsWith("udtrukket")) return true;
  return false;
}

function parseCsvFlexible(text, expectedHeaders) {
  const clean = (text || "").replace(/^\uFEFF/, "");

  const candidates = [";", ",", "\t"];
  let best = { rows: [], headerIdx: -1, score: -1, delimiter: ";" };
  for (const d of candidates) {
    const attempt = parseWithDelimiter(clean, d, expectedHeaders);
    if (attempt.score > best.score) best = { ...attempt, delimiter: d };
  }
  const rows = best.rows;

  // Title in first row if "Periode…"
  let title = "";
  const t0 = (rows[0]?.[0] || "").toString().replace(/^"+|"+$/g, "");
  if (norm(t0).startsWith("periode")) title = t0;

  const headerIdx = best.headerIdx >= 0 ? best.headerIdx : 0;

  // Subtitle: first non-empty line between title and header
  let subtitle = "";
  for (let i = 1; i < headerIdx; i++) {
    const line = rows[i] || [];
    const pieces = line.map(x => (x ?? "").toString().trim()).filter(Boolean);
    if (pieces.length) { subtitle = pieces.join("  "); break; }
  }

  // Headers & data (NO heuristic dropping)
  const headers = (rows[headerIdx] || []).map(x => (x ?? "").toString().trim());
  const data = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const arr = rows[i];
    if (isNoiseRow(arr)) continue;

    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      const key = headers[c] || `col_${c + 1}`;
      obj[key] = arr[c] ?? "";
    }
    data.push(obj);
  }

  return { title, subtitle, headers, rows: data };
}

/* ------------------------------ Handler ------------------------------ */
module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Use POST" });
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body || "{}") : (req.body || {});
    const {
      sheet1_url, sheet1_csv, sheet1_headers,   // SAP
      sheet2_url, sheet2_csv, sheet2_headers,   // Detaljeret
      extracted_at,
      file_name = "winter_fg.xlsx"
    } = body;

    const detDaysAll = Array.from({ length: 31 }, (_, i) => String(i + 1));

    // Parse CSVs
    const sapText = await readCsvText({ url: sheet1_url, csv: sheet1_csv, headers: sheet1_headers });
    const detText = await readCsvText({ url: sheet2_url, csv: sheet2_csv, headers: sheet2_headers });

    const sapParsed = parseCsvFlexible(sapText, SAP_HEADERS);
    const detParsed = parseCsvFlexible(detText, [...DET_BASE_HEADERS, ...detDaysAll]);

    // Dynamic day columns present in Detaljeret
    const detDays = detParsed.headers.map(h => (h || "").trim()).filter(h => /^\d{1,2}$/.test(h));
    const DET_HEADERS = [...DET_BASE_HEADERS, ...detDays];
    const DET_WIDTHS  = [...DET_BASE_WIDTHS, ...new Array(detDays.length).fill(DAY_COL_WIDTH)];
    const footerDate = extracted_at || todayDk();

    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true;

    /* -------- Detaljeret oversigt (first tab) -------- */
    const wsDET = wb.addWorksheet(SHEET_DET);
    wsDET.columns = DET_WIDTHS.map(w => ({ width: w }));

    setTitle(wsDET, detParsed.title || sapParsed.title || "", DET_HEADERS.length);
    const subtitleText = detParsed.subtitle || sapParsed.subtitle || "";
    setSubtitle(wsDET, subtitleText, DET_HEADERS.length);

    writeHeader(wsDET, DET_HEADERS);

    let r = HEADER_ROW_INDEX + 1;
    detParsed.rows.forEach(obj => {
      const row = wsDET.getRow(r);
      DET_HEADERS.forEach((hdr, i) => {
        const raw = getValueForHeader(obj, hdr);
        const val = coerceByHeader(hdr, raw);
        const cell = row.getCell(i + 1);

        if (/^\d{1,2}$/.test(hdr)) {
          // Day columns: keep as text (K/S markers, etc.)
          cell.value = val === null ? null : String(val);
          cell.numFmt = "General";
        } else if (typeof val === "number") {
          // Numeric columns: keep numeric; format by type
          cell.value = Number(val);

          // ——— Formatting rules (J→P are covered by these headers) ———
          // Any header ending with ', kr' or containing 'gns. pris' => 3 decimals
          // Counts/weights ('antal', 'kg') => 0 decimals
          if (/, kr$/i.test(hdr) || /gns\. pris/i.test(hdr)) {
            cell.numFmt = "#,##0.000";
          } else if (/antal|kg/i.test(hdr)) {
            cell.numFmt = "0";
          } else {
            cell.numFmt = "#,##0.000"; // fallback: 3 decimals to match your example
          }
        } else {
          // Text columns
          cell.value = val === null ? null : String(val);
        }
      });
      row.commit();
      r++;
    });

    const detFirstDataRow = HEADER_ROW_INDEX + 1;
    const detLast = r - 1;
    if (detLast >= detFirstDataRow) {
      addBorders(wsDET, detFirstDataRow, detLast, 1, DET_HEADERS.length);

      // Totals row: keep your previous summed columns J..M
      const j = DET_HEADERS.indexOf("I alt, kr") + 1;
      const k = DET_HEADERS.indexOf("Saltning, kr") + 1;
      const l = DET_HEADERS.indexOf("Salt, antal") + 1;
      const m = DET_HEADERS.indexOf("Salt, gns. pris") + 1;
      const detTotalsRow = detLast + 1;
      const colsToSum = [j, k, l, m].filter(Boolean);
      if (colsToSum.length) writeSumFormulas(wsDET, detTotalsRow, colsToSum, detFirstDataRow, detLast);

      // Make sure visible totals use correct formats too
      if (j) wsDET.getCell(detTotalsRow, j).numFmt = "#,##0.000";
      if (k) wsDET.getCell(detTotalsRow, k).numFmt = "#,##0.000";
      if (l) wsDET.getCell(detTotalsRow, l).numFmt = "0";
      if (m) wsDET.getCell(detTotalsRow, m).numFmt = "#,##0.000";

      writeFooter(wsDET, 1, 3, detTotalsRow, footerDate);
    } else {
      writeFooter(wsDET, 1, 3, HEADER_ROW_INDEX, footerDate);
    }

    /* -------- SAP (second tab) -------- */
    const wsSAP = wb.addWorksheet(SHEET_SAP);
    wsSAP.columns = SAP_WIDTHS.map(w => ({ width: w }));
    setTitle(wsSAP, sapParsed.title || detParsed.title || "", SAP_HEADERS.length);
    writeHeader(wsSAP, SAP_HEADERS);

    r = HEADER_ROW_INDEX + 1;
    sapParsed.rows.forEach(obj => {
      const row = wsSAP.getRow(r);
      SAP_HEADERS.forEach((hdr, i) => {
        const raw = getValueForHeader(obj, hdr);
        const val = coerceByHeader(hdr, raw);
        const cell = row.getCell(i + 1);

        if (typeof val === "number") {
          cell.value = Number(val);
          // Prices keep two decimals
          if (/moms/i.test(hdr)) cell.numFmt = "#,##0.00";
          else cell.numFmt = "0";
        } else {
          cell.value = val === null ? null : String(val);
        }
      });
      row.commit();
      r++;
    });

    const sapFirstDataRow = HEADER_ROW_INDEX + 1;
    const sapLast = r - 1;
    if (sapLast >= sapFirstDataRow) {
      addBorders(wsSAP, sapFirstDataRow, sapLast, 1, SAP_HEADERS.length);
      const h = SAP_HEADERS.indexOf("Pris ex. moms") + 1;
      const sapTotalsRow = sapLast + 1;
      if (h) writeSumFormulas(wsSAP, sapTotalsRow, [h], sapFirstDataRow, sapLast);
      writeFooter(wsSAP, 1, 3, sapTotalsRow, footerDate);
    } else {
      writeFooter(wsSAP, 1, 3, HEADER_ROW_INDEX, footerDate);
    }

    // Return file
    const buf = await toBuffer(wb);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${file_name}"`);
    return res.status(200).send(buf);
  } catch (e) {
    return res.status(400).json({ error: e.message, stack: e.stack });
  }
};
