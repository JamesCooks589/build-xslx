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

const TITLE_ROW_INDEX    = 1;
const SUBTITLE_ROW_INDEX = 2;
const BLANK_ROW_INDEX    = 3;
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

const DET_NUMERIC_HEADERS = new Set([
  "I alt, kr","Saltning, kr","Salt, antal","Salt, gns. pris",
  "Kombi, kr","Kombi, antal","Kombi, gns. pris",
  "Snerydning, kr","Sne, antal","Sne, gns. pris",
  "Andet, kr","Andet, antal","Andet, gns. pris","Salt, kg"
]);
const SAP_NUMERIC_HEADERS = new Set([
  "Pris inkl. moms","Pris ex. moms"
]);

/* -------------------- Number parsing -------------------- */
function toNumberMaybe(v) {
  if (v == null) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/\s+/g, "").replace(/;/g, "");

  if (/,/.test(s) && /^\d{1,3}(\.\d{3})*(,\d+)?$/.test(s))
    s = s.replace(/\./g, "").replace(",", ".");
  else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s))
    s = s.replace(/,/g, "");
  else if (/^\d+,\d+$/.test(s))
    s = s.replace(",", ".");
  else if (/^\d{1,3}(\.\d{3})+$/.test(s))
    s = s.replace(/\./g, "");

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function coerceByHeader(header, value) {
  const h = (header || "").trim();
  if (/^\d{1,2}$/.test(h)) return value === "" ? null : String(value);
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

function getValueForHeader(obj, hdr) {
  if (obj[hdr] !== undefined) return obj[hdr];
  const target = stripComma(hdr);
  const k = Object.keys(obj).find(key => norm(key) === norm(hdr) || stripComma(key) === target);
  return k ? obj[k] : null;
}

/* -------------------- Parsing -------------------- */
function parseCsvFlexible(text, expectedHeaders) {
  const clean = (text || "").replace(/^\uFEFF/, "");
  const parsed = Papa.parse(clean, { header: true, skipEmptyLines: true });
  const headers = parsed.meta.fields || expectedHeaders;
  return { title: "", subtitle: "", headers, rows: parsed.data };
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
      sheet1_url, sheet1_csv, sheet2_url, sheet2_csv,
      extracted_at, file_name = "winter_fg.xlsx"
    } = body;

    const detDaysAll = Array.from({ length: 31 }, (_, i) => String(i + 1));
    const sapText = await readCsvText({ url: sheet1_url, csv: sheet1_csv });
    const detText = await readCsvText({ url: sheet2_url, csv: sheet2_csv });

    const sapParsed = parseCsvFlexible(sapText, SAP_HEADERS);
    const detParsed = parseCsvFlexible(detText, [...DET_BASE_HEADERS, ...detDaysAll]);

    const detDays = detParsed.headers.filter(h => /^\d{1,2}$/.test(h));
    const DET_HEADERS = [...DET_BASE_HEADERS, ...detDays];
    const DET_WIDTHS  = [...DET_BASE_WIDTHS, ...new Array(detDays.length).fill(DAY_COL_WIDTH)];
    const footerDate = extracted_at || todayDk();

    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true;

    /* -------- Detaljeret oversigt -------- */
    const wsDET = wb.addWorksheet(SHEET_DET);
    wsDET.columns = DET_WIDTHS.map(w => ({ width: w }));
    writeHeader(wsDET, DET_HEADERS);

    const decimalsForHeader = hdr => (/antal|kg/i.test(hdr) ? 0 : 2);

    let r = HEADER_ROW_INDEX + 1;
    detParsed.rows.forEach(obj => {
      const row = wsDET.getRow(r);
      DET_HEADERS.forEach((hdr, i) => {
        const raw = getValueForHeader(obj, hdr);
        const val = coerceByHeader(hdr, raw);
        const cell = row.getCell(i + 1);

        if (/^\d{1,2}$/.test(hdr)) {
          cell.value = val === null ? null : String(val);
        } else if (typeof val === "number") {
          const dec = decimalsForHeader(hdr);
          cell.value = val;
          cell.numFmt = dec === 0 ? "0" : "#,##0.00";
        } else {
          cell.value = val === null ? null : String(val);
        }
      });
      row.commit();
      r++;
    });

    const first = HEADER_ROW_INDEX + 1;
    const last = r - 1;
    if (last >= first) {
      addBorders(wsDET, first, last, 1, DET_HEADERS.length);
      const j = DET_HEADERS.indexOf("I alt, kr") + 1;
      const k = DET_HEADERS.indexOf("Saltning, kr") + 1;
      const l = DET_HEADERS.indexOf("Salt, antal") + 1;
      const m = DET_HEADERS.indexOf("Salt, gns. pris") + 1;
      const totalRow = last + 1;
      const colsToSum = [j, k, l, m].filter(Boolean);
      if (colsToSum.length) writeSumFormulas(wsDET, totalRow, colsToSum, first, last);
      writeFooter(wsDET, 1, 3, totalRow, footerDate);
    }

    /* -------- SAP -------- */
    const wsSAP = wb.addWorksheet(SHEET_SAP);
    wsSAP.columns = SAP_WIDTHS.map(w => ({ width: w }));
    writeHeader(wsSAP, SAP_HEADERS);

    r = HEADER_ROW_INDEX + 1;
    sapParsed.rows.forEach(obj => {
      const row = wsSAP.getRow(r);
      SAP_HEADERS.forEach((hdr, i) => {
        const raw = getValueForHeader(obj, hdr);
        const val = coerceByHeader(hdr, raw);
        const cell = row.getCell(i + 1);

        if (typeof val === "number") {
          cell.value = val;
          if (/moms/i.test(hdr)) cell.numFmt = "#,##0.00";
          else cell.numFmt = "0";
        } else {
          cell.value = val === null ? null : String(val);
        }
      });
      row.commit();
      r++;
    });

    const sapFirst = HEADER_ROW_INDEX + 1;
    const sapLast = r - 1;
    if (sapLast >= sapFirst) {
      addBorders(wsSAP, sapFirst, sapLast, 1, SAP_HEADERS.length);
      const h = SAP_HEADERS.indexOf("Pris ex. moms") + 1;
      const sapTotalsRow = sapLast + 1;
      if (h) writeSumFormulas(wsSAP, sapTotalsRow, [h], sapFirst, sapLast);
      writeFooter(wsSAP, 1, 3, sapTotalsRow, footerDate);
    }

    const buf = await toBuffer(wb);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${file_name}"`);
    return res.status(200).send(buf);
  } catch (e) {
    return res.status(400).json({ error: e.message, stack: e.stack });
  }
};
