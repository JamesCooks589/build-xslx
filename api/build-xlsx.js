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

const TITLE_ROW_INDEX    = 1; // Periode…
const SUBTITLE_ROW_INDEX = 2; // long supplier list
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

/* -------------------- Column typing helpers -------------------- */
const MONEY_HEADERS = new Set([
  "I alt, kr","Saltning, kr","Kombi, kr","Snerydning, kr","Andet, kr",
  "Pris inkl. moms","Pris ex. moms"
]);

const AVG_PRICE_HEADERS = new Set([
  "Salt, gns. pris","Kombi, gns. pris","Sne, gns. pris","Andet, gns. pris"
]);

const COUNT_HEADERS = new Set([
  "Salt, antal","Kombi, antal","Sne, antal","Andet, antal","Salt, kg"
]);

/* -------------------- Number parsing with ; grouping --------------------
   - Treat ';' as thousands/grouping separator (never decimal).
   - Support '.' or ',' as decimals where applicable.
   - Return { num, fmt, decimals, groupingUsed } so we can format the Excel cell.
----------------------------------------------------------------------------- */
function parseNumberAndFormat(rawText) {
  if (rawText === null || rawText === undefined) return null;
  let s = String(rawText).trim();
  if (!s) return null;

  // normalize thin spaces and NBSP
  s = s.replace(/\u00A0|\u2009/g, "");

  // Detect semicolon grouping (your data uses ";" for thousands)
  const hadSemicolonGrouping = s.includes(";");
  if (hadSemicolonGrouping) {
    s = s.replace(/;/g, ""); // remove grouping to parse numerically
  }

  const hasDot = s.includes(".");
  const hasComma = s.includes(",");
  let decimals = 0;
  let groupingUsed = hadSemicolonGrouping; // true if we saw any ';'
  let num = null;
  let fmt = null;

  const decFmt = d => (d > 0 ? "0." + "0".repeat(d) : "0");
  const decFmtGrouped = d => (d > 0 ? "#,##0." + "0".repeat(d) : "#,##0");

  // Case: both '.' and ',' present -> decide decimal
  if (hasDot && hasComma) {
    const lastDot = s.lastIndexOf(".");
    const lastComma = s.lastIndexOf(",");
    // If comma is the last separator, treat comma as decimal (European "1.234,56").
    if (lastComma > lastDot) {
      const noDots = s.replace(/\./g, "");
      const canonical = noDots.replace(",", ".");
      num = Number(canonical);
      if (!Number.isFinite(num)) return null;
      const tail = s.slice(lastComma + 1);
      decimals = /^\d+$/.test(tail) ? tail.length : 0;
      groupingUsed = true; // dots were grouping
      fmt = decFmtGrouped(decimals);
      return { num, fmt, decimals, groupingUsed };
    }
    // Else dot is decimal (rare mix), treat comma as grouping
    const noCommas = s.replace(/,/g, "");
    num = Number(noCommas);
    if (!Number.isFinite(num)) return null;
    const tail = s.slice(lastDot + 1);
    decimals = /^\d+$/.test(tail) ? tail.length : 0;
    groupingUsed = true;
    fmt = decFmtGrouped(decimals);
    return { num, fmt, decimals, groupingUsed };
  }

  // Only comma present
  if (hasComma && !hasDot) {
    const parts = s.split(",");
    const tail = parts[1] || "";
    if (parts.length === 2 && /^\d{1,3}$/.test(tail)) {
      // decimal comma
      decimals = tail.length;
      num = Number(parts[0] + "." + tail);
      if (!Number.isFinite(num)) return null;
      fmt = groupingUsed ? decFmtGrouped(decimals) : decFmt(decimals);
      return { num, fmt, decimals, groupingUsed };
    }
    // otherwise treat comma as grouping
    num = Number(s.replace(/,/g, ""));
    if (!Number.isFinite(num)) return null;
    decimals = 0;
    groupingUsed = true || groupingUsed;
    fmt = decFmtGrouped(0);
    return { num, fmt, decimals, groupingUsed };
  }

  // Only dot present
  if (!hasComma && hasDot) {
    const parts = s.split(".");
    const tail = parts[1] || "";
    if (parts.length === 2 && /^\d{1,3}$/.test(tail)) {
      // decimal dot
      decimals = tail.length;
      num = Number(s);
      if (!Number.isFinite(num)) return null;
      fmt = groupingUsed ? decFmtGrouped(decimals) : decFmt(decimals);
      return { num, fmt, decimals, groupingUsed };
    }
    // multiple dots or unclear -> strip grouping dots
    num = Number(s.replace(/\./g, ""));
    if (!Number.isFinite(num)) return null;
    decimals = 0;
    groupingUsed = true || groupingUsed;
    fmt = decFmtGrouped(0);
    return { num, fmt, decimals, groupingUsed };
  }

  // No dot or comma -> integer
  num = Number(s);
  if (!Number.isFinite(num)) return null;
  decimals = 0;
  fmt = groupingUsed ? decFmtGrouped(0) : "0";
  return { num, fmt, decimals, groupingUsed };
}

/* -------------------- Utilities -------------------- */
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

function getValueForHeader(obj, hdr) {
  if (obj[hdr] !== undefined) return obj[hdr];
  const target = stripComma(hdr);
  const k = Object.keys(obj).find(key => norm(key) === norm(hdr) || stripComma(key) === target);
  return k ? obj[k] : null;
}

/* -------------------- Robust CSV parsing -------------------- */
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

/* Returns { title, subtitle, headers, rows } */
function parseCsvFlexible(text, expectedHeaders, opts = {}) {
  const { identityKeys = [], valueKeys = [], dayKeys = [] } = opts;
  const clean = (text || "").replace(/^\uFEFF/, "");

  // Try multiple delimiters and pick best
  const candidates = [";", ",", "\t"];
  let best = { rows: [], headerIdx: -1, score: -1, delimiter: ";" };
  for (const d of candidates) {
    const attempt = parseWithDelimiter(clean, d, expectedHeaders);
    if (attempt.score > best.score) best = { ...attempt, delimiter: d };
  }
  const rows = best.rows;

  // Title on first row if "Periode…"
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

  // Headers & data
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

    const nonEmptyCount = arr.filter(v => (v ?? "").toString().trim() !== "").length;
    const firstNonEmpty  = (arr[0] ?? "").toString().trim() !== "";

    const hasIdentity = identityKeys.some(k => ((obj[k] ?? "").toString().trim() !== ""));
    const hasValues   = valueKeys.some(k   => ((obj[k] ?? "").toString().trim() !== ""));
    const hasDays     = dayKeys.some(k     => ((obj[k] ?? "").toString().trim() !== ""));

    if (hasIdentity || hasValues || hasDays || (firstNonEmpty && nonEmptyCount === 1)) {
      data.push(obj);
    }
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

    const sapParsed = parseCsvFlexible(sapText, SAP_HEADERS, {
      identityKeys: ["Kontrakt", "Lokation/rute", "Kunde"],
      valueKeys: ["Pris inkl. moms", "Pris ex. moms"],
    });

    const detParsed = parseCsvFlexible(detText, [...DET_BASE_HEADERS, ...detDaysAll], {
      identityKeys: ["Øko-ID", "Rute/lokation", "Lokationsnavn"],
      valueKeys: [
        "I alt, kr","Saltning, kr","Salt, antal","Salt, gns. pris",
        "Kombi, kr","Kombi, antal","Kombi, gns. pris",
        "Snerydning, kr","Sne, antal","Sne, gns. pris",
        "Andet, kr","Andet, antal","Salt, kg",
      ],
      dayKeys: detDaysAll,
    });

    // Dynamic day columns
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
    setSubtitle(wsDET, detParsed.subtitle || sapParsed.subtitle || "", DET_HEADERS.length);
    writeHeader(wsDET, DET_HEADERS);

    // Track per-column decimal precision & whether grouping was used at least once
    const detDecimalsByHeader = Object.create(null);  // max decimals seen
    const detGroupingByHeader = Object.create(null);  // any grouping used

    let r = HEADER_ROW_INDEX + 1;
    detParsed.rows.forEach(obj => {
      const row = wsDET.getRow(r);
      DET_HEADERS.forEach((hdr, i) => {
        const raw = getValueForHeader(obj, hdr);
        const cell = row.getCell(i + 1);

        // Day columns stay as text (K/S)
        if (/^\d{1,2}$/.test(hdr)) {
          cell.value = (raw === "" || raw === null) ? null : raw;
          cell.numFmt = "General";
          return;
        }

        // Counts: numeric integers unless decimals present
        if (COUNT_HEADERS.has(hdr)) {
          const parsed = parseNumberAndFormat(raw);
          if (parsed) {
            cell.value = parsed.num;
            const zeros = parsed.decimals > 0 ? "0." + "0".repeat(parsed.decimals) : "0";
            // If groups present on counts (rare), allow grouping
            cell.numFmt = parsed.groupingUsed ? (parsed.decimals ? "#,##0." + "0".repeat(parsed.decimals) : "#,##0") : zeros;
            detDecimalsByHeader[hdr] = Math.max(detDecimalsByHeader[hdr] || 0, parsed.decimals);
            if (parsed.groupingUsed) detGroupingByHeader[hdr] = true;
          } else {
            cell.value = (raw === "" || raw === null) ? null : raw;
          }
          return;
        }

        // Money & avg price columns
        if (MONEY_HEADERS.has(hdr) || AVG_PRICE_HEADERS.has(hdr)) {
          const parsed = parseNumberAndFormat(raw);
          if (parsed) {
            cell.value = parsed.num;
            cell.numFmt = parsed.fmt; // grouped vs non-grouped chosen by parser
            detDecimalsByHeader[hdr] = Math.max(detDecimalsByHeader[hdr] || 0, parsed.decimals);
            if (parsed.groupingUsed) detGroupingByHeader[hdr] = true;
          } else {
            cell.value = (raw === "" || raw === null) ? null : raw;
          }
          return;
        }

        // Other text
        cell.value = (raw === "" || raw === null) ? null : raw;
      });
      row.commit();
      r++;
    });
    const detLast = r - 1;

    if (detLast >= HEADER_ROW_INDEX + 1) {
      addBorders(wsDET, HEADER_ROW_INDEX + 1, detLast, 1, DET_HEADERS.length);

      // Totals row for a few key columns
      const colIndex = name => DET_HEADERS.indexOf(name) + 1;
      const j = colIndex("I alt, kr");
      const k = colIndex("Saltning, kr");
      const l = colIndex("Salt, antal");
      const m = colIndex("Salt, gns. pris");
      const detTotalsRow = detLast + 1;
      const colsToSum = [j, k, l, m].filter(Boolean);
      if (colsToSum.length) writeSumFormulas(wsDET, detTotalsRow, colsToSum, HEADER_ROW_INDEX + 1, detLast);

      // Format totals based on column stats (max decimals & whether grouping was used)
      const applyTotalFmt = (hdr, colIdx) => {
        if (!colIdx) return;
        const d = detDecimalsByHeader[hdr] || 0;
        const grouping = !!detGroupingByHeader[hdr];
        let fmt;
        if (grouping || MONEY_HEADERS.has(hdr)) {
          fmt = (d > 0) ? "#,##0." + "0".repeat(d) : "#,##0";
        } else {
          fmt = (d > 0) ? "0." + "0".repeat(d) : "0";
        }
        wsDET.getCell(detTotalsRow, colIdx).numFmt = fmt;
      };

      applyTotalFmt("I alt, kr", j);
      applyTotalFmt("Saltning, kr", k);
      applyTotalFmt("Salt, antal", l);
      applyTotalFmt("Salt, gns. pris", m);

      writeFooter(wsDET, 1, 3, detTotalsRow, footerDate);
    } else {
      writeFooter(wsDET, 1, 3, HEADER_ROW_INDEX, footerDate);
    }

    /* -------- SAP (second tab) -------- */
    const wsSAP = wb.addWorksheet(SHEET_SAP);
    wsSAP.columns = SAP_WIDTHS.map(w => ({ width: w }));
    setTitle(wsSAP, sapParsed.title || detParsed.title || "", SAP_HEADERS.length);
    writeHeader(wsSAP, SAP_HEADERS);

    const sapDecimalsByHeader = Object.create(null);
    const sapGroupingByHeader = Object.create(null);

    r = HEADER_ROW_INDEX + 1;
    sapParsed.rows.forEach(obj => {
      const row = wsSAP.getRow(r);
      SAP_HEADERS.forEach((hdr, i) => {
        const raw = getValueForHeader(obj, hdr);
        const cell = row.getCell(i + 1);

        if (MONEY_HEADERS.has(hdr) || AVG_PRICE_HEADERS.has(hdr)) {
          const parsed = parseNumberAndFormat(raw);
          if (parsed) {
            cell.value = parsed.num;
            cell.numFmt = parsed.fmt;
            sapDecimalsByHeader[hdr] = Math.max(sapDecimalsByHeader[hdr] || 0, parsed.decimals);
            if (parsed.groupingUsed) sapGroupingByHeader[hdr] = true;
          } else {
            cell.value = (raw === "" || raw === null) ? null : raw;
          }
          return;
        }

        // Other fields: try numeric-inference; else text
        const parsed = parseNumberAndFormat(raw);
        if (parsed) {
          cell.value = parsed.num;
          cell.numFmt = parsed.fmt;
          sapDecimalsByHeader[hdr] = Math.max(sapDecimalsByHeader[hdr] || 0, parsed.decimals);
          if (parsed.groupingUsed) sapGroupingByHeader[hdr] = true;
        } else {
          cell.value = (raw === "" || raw === null) ? null : raw;
        }
      });
      row.commit();
      r++;
    });
    const sapLast = r - 1;

    if (sapLast >= HEADER_ROW_INDEX + 1) {
      addBorders(wsSAP, HEADER_ROW_INDEX + 1, sapLast, 1, SAP_HEADERS.length);

      // Example total for Pris ex. moms
      const hIdx = SAP_HEADERS.indexOf("Pris ex. moms") + 1;
      const sapTotalsRow = sapLast + 1;
      if (hIdx) {
        writeSumFormulas(wsSAP, sapTotalsRow, [hIdx], HEADER_ROW_INDEX + 1, sapLast);
        const d = sapDecimalsByHeader["Pris ex. moms"] || 2;
        const grouping = !!sapGroupingByHeader["Pris ex. moms"];
        wsSAP.getCell(sapTotalsRow, hIdx).numFmt = grouping
          ? "#,##0." + "0".repeat(d)
          : "0." + "0".repeat(d);
      }

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
