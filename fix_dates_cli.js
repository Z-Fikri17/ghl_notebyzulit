#!/usr/bin/env node
// fix_dates_cli.js — command-line twin of the "Fix Dates in an Existing Excel File"
// feature in csv_to_excel_appender_settled_txn_3.html.
//
// Usage:
//   node fix_dates_cli.js <file.csv|xlsx> [out.xlsx]
//
// - For a raw settled-txn CSV: parses with its own RFC4180 parser and writes an
//   .xlsx where Settlement Date / Transaction Time are strict dd/mm/yyyy text.
// - For an already-produced .xlsx (dates mis-read by Excel, e.g. 03/07/2026 → 07
//   March 2026): recovers the true dates using the range stated in the filename
//   (settled-txn-<id>-<start>-to-<end>.<ext>), swapping day/month back when one
//   interpretation fits that range.
const fs = require('fs');
const path = require('path');

const XLSX = require('xlsx');

const FNAME_RE = /^settled-txn-(\d+)-(\d{4}-\d{2}-\d{2})-to-(\d{4}-\d{2}-\d{2})\.(csv|xlsx)$/i;
const NUMERIC_COLS = ['Txn Amount(RM)', 'MDR(RM)', 'Net Amount(RM)'];
const DATE_COLS = ['Settlement Date', 'Transaction Time'];

function pad2(n) { return String(n).padStart(2, '0'); }

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function toUnamb(d, mo, y, hh, mm, ss) {
  return `${pad2(d)}-${MONTHS[mo - 1]}-${y} ${pad2(hh)}:${pad2(mm)}:${pad2(ss)}`;
}

function parseDateTimeParts(s) {
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?$/);
  if (m) return { d: +m[1], mo: +m[2], y: +m[3], hh: m[4] || '00', mm: m[5] || '00', ss: m[6] || '00' };
  m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?$/);
  if (m) {
    const mi = MONTHS.indexOf(m[2][0].toUpperCase() + m[2].slice(1).toLowerCase());
    if (mi !== -1) return { d: +m[1], mo: mi + 1, y: +m[3], hh: m[4] || '00', mm: m[5] || '00', ss: m[6] || '00', monthNamed: true };
  }
  return null;
}

function parseCSV(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = [];
  let row = [], field = '', inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\r') { /* skip */ }
      else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else field += c;
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
  return rows;
}

function expectedWindow(name) {
  const m = String(name || '').match(/^settled-txn-(\d+)-(\d{4}-\d{2}-\d{2})-to-(\d{4}-\d{2}-\d{2})\.(?:csv|xlsx)$/i);
  if (!m) return null;
  return {
    start: new Date(m[2] + 'T00:00:00Z'),
    end: new Date(m[3] + 'T00:00:00Z'),
    year: parseInt(m[2].slice(0, 4), 10),
    month: parseInt(m[2].slice(5, 7), 10) - 1
  };
}

function daysInMonth(year, mo0) {
  return new Date(Date.UTC(year, mo0 + 1, 0)).getUTCDate();
}

function pickDate(year, d1, mo1, d2, mo2, win) {
  const valid = (d, mo) => mo >= 1 && mo <= 12 && d >= 1 && d <= daysInMonth(year, mo - 1);
  if (d1 === d2 && mo1 === mo2) return { d: d1, mo: mo1, note: 'ok' };
  const cands = [
    { d: d1, mo: mo1, i: 0 },
    { d: d2, mo: mo2, i: 1 }
  ];
  const v = cands.filter(c => valid(c.d, c.mo));
  if (v.length === 0) return { d: d1, mo: mo1, note: 'invalid' };
  if (v.length === 1) {
    const c = v[0];
    return { d: c.d, mo: c.mo, note: c.i === 1 ? 'swapped' : 'ok' };
  }
  const score = c => {
    let s = 0;
    if (c.mo - 1 === win.month && year === win.year) s += 1000;
    const t = Date.UTC(year, c.mo - 1, c.d);
    if (win && t >= win.start.getTime() && t <= win.end.getTime()) s += 500;
    else if (win) {
      const dist = Math.min(Math.abs(t - win.start.getTime()), Math.abs(t - win.end.getTime()));
      s -= Math.min(dist / 86400000, 1000);
    }
    return s;
  };
  const [a, b] = v;
  const sa = score(a), sb = score(b);
  if (sa !== sb) {
    const c = sa > sb ? a : b;
    return { d: c.d, mo: c.mo, note: c.i === 1 ? 'swapped' : 'ok' };
  }
  return { d: d1, mo: mo1, note: 'ambiguous' };
}

function normalizeDateTimeSmart(raw, ctx, win, warnings) {
  if (raw === undefined || raw === null || raw === '') return { text: '', note: 'ok' };
  const s = String(raw).trim();
  const parts = parseDateTimeParts(s);
  if (!parts) {
    warnings.push(`${ctx}: unrecognised date format "${s}" — left as-is`);
    return { text: s, note: 'unrecognised' };
  }
  const { d, mo, y, hh, mm, ss } = parts;
  if (parts.monthNamed) {
    return { text: toUnamb(d, mo, y, hh, mm, ss), note: 'ok' };
  }
  const pick = win ? pickDate(y, d, mo, mo, d, win) : { d, mo, note: 'ok' };
  if (pick.note === 'invalid') {
    warnings.push(`${ctx}: invalid date "${s}" (day=${d}, month=${mo}) — left as-is, please check source`);
    return { text: s, note: 'invalid' };
  }
  if (pick.note === 'ambiguous') {
    warnings.push(`${ctx}: "${pad2(d)}/${pad2(mo)}" is genuinely ambiguous (${pad2(d)}/${pad2(mo)} vs ${pad2(mo)}/${pad2(d)} both fit) — left as-is, please check`);
    return { text: toUnamb(d, mo, y, hh, mm, ss), note: 'ambiguous' };
  }
  if (pick.note === 'swapped') {
    warnings.push(`${ctx}: "${pad2(d)}/${pad2(mo)}" doesn't fit the file's stated range — swapped to "${pad2(pick.d)}/${pad2(pick.mo)}"`);
    return { text: toUnamb(pick.d, pick.mo, y, hh, mm, ss), note: 'swapped' };
  }
  return { text: toUnamb(pick.d, pick.mo, y, hh, mm, ss), note: 'ok' };
}

function stripForcedQuote(v) {
  if (typeof v !== 'string') return v;
  if (v.length >= 2 && v.startsWith("'") && v.endsWith("'")) return v.slice(1, -1);
  return v;
}

function excelDateToText(dt) {
  return `${pad2(dt.getUTCDate())}/${pad2(dt.getUTCMonth() + 1)}/${dt.getUTCFullYear()} ${pad2(dt.getUTCHours())}:${pad2(dt.getUTCMinutes())}:${pad2(dt.getUTCSeconds())}`;
}

function roundToSecond(dt) {
  return new Date(Math.round(dt.getTime() / 1000) * 1000);
}

function buildWorkbookFromCsv(text, name) {
  const rows = parseCSV(text);
  if (rows.length < 2) throw new Error('Empty or unreadable CSV.');
  const headers = rows[0];
  const dataRows = rows.slice(1).filter(r => r.length > 1 || (r[0] !== undefined && r[0] !== ''));
  const numericIdx = headers.map(h => NUMERIC_COLS.includes(h));
  const dateIdx = headers.map(h => DATE_COLS.includes(h));
  const win = expectedWindow(name);

  const ws = {};
  headers.forEach((h, c) => { ws[XLSX.utils.encode_cell({ r: 0, c })] = { t: 's', v: h ?? '' }; });

  let sheetFixed = 0, sheetOk = 0, sheetAmbiguous = 0, sheetUnrecognized = 0;
  const warnings = [];
  dataRows.forEach((row, rIdx) => {
    for (let c = 0; c < headers.length; c++) {
      const ref = XLSX.utils.encode_cell({ r: rIdx + 1, c });
      let val = row[c] !== undefined ? row[c] : '';
      if (dateIdx[c]) {
        const res = normalizeDateTimeSmart(val, `row ${rIdx + 2} (${headers[c]})`, win, warnings);
        if (res.note === 'unrecognised') sheetUnrecognized++;
        else if (res.note === 'invalid' || res.note === 'ambiguous') sheetAmbiguous++;
        else if (res.note === 'swapped') sheetFixed++;
        else sheetOk++;
        ws[ref] = { t: 's', v: res.text };
      } else if (numericIdx[c]) {
        const trimmed = String(val).trim(); const num = parseFloat(trimmed);
        ws[ref] = (trimmed === '' || isNaN(num)) ? { t: 's', v: val } : { t: 'n', v: num };
      } else {
        ws[ref] = { t: 's', v: stripForcedQuote(String(val)) };
      }
    }
  });
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: dataRows.length, c: headers.length - 1 } });
  ws['!cols'] = dateIdx.map(isDate => isDate ? { wch: 20 } : null);

  const m = name.match(FNAME_RE);
  const sheetName = m ? m[1] : 'Sheet1';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return { wb, stats: { sheetFixed, sheetOk, sheetAmbiguous, sheetUnrecognized, warnings } };
}

function fixWorkbookFromXlsx(buf, name) {
  const win = expectedWindow(name);
  const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
  let totalFixed = 0, totalOk = 0, totalAmbiguous = 0, totalUnrecognized = 0;
  const warnings = [];

  wb.SheetNames.forEach(sheetName => {
    const ws = wb.Sheets[sheetName];
    if (!ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);

    const dateCols = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const headerCell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
      if (headerCell && DATE_COLS.includes(headerCell.v)) dateCols.push(c);
    }
    if (dateCols.length === 0) return;

    let sheetFixed = 0, sheetOk = 0, sheetAmbiguous = 0, sheetUnrecognized = 0;
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      dateCols.forEach(c => {
        const ref = XLSX.utils.encode_cell({ r, c });
        const cell = ws[ref];
        if (!cell || cell.v === undefined || cell.v === '') return;

        if (cell.t === 'd' || cell.v instanceof Date) {
          const dt = roundToSecond(cell.v);
          const y = dt.getUTCFullYear();
          const excelMo = dt.getUTCMonth() + 1;
          const excelD = dt.getUTCDate();
          const hh = dt.getUTCHours(), mm = dt.getUTCMinutes(), ss = dt.getUTCSeconds();
          const pick = win ? pickDate(y, excelD, excelMo, excelMo, excelD, win) : { d: excelD, mo: excelMo, note: 'ok' };
          const text = toUnamb(pick.d, pick.mo, y, hh, mm, ss);
          ws[ref] = { t: 's', v: text };
          if (pick.note === 'swapped') {
            sheetFixed++;
            warnings.push(`  ✓ ${sheetName}!${ref}: "${excelDateToText(dt)}" → "${text}" (swapped to match ${name})`);
          } else if (pick.note === 'ambiguous' || pick.note === 'invalid') {
            sheetAmbiguous++;
            warnings.push(`  ⚠ ${sheetName}!${ref}: "${excelDateToText(dt)}" ${pick.note === 'invalid' ? 'is invalid' : 'is genuinely ambiguous'} — left as-is, please check`);
          } else {
            sheetOk++;
          }
        } else if (cell.t === 's') {
          const before = cell.v;
          const res = normalizeDateTimeSmart(before, `${sheetName}!${ref}`, win, warnings);
          if (res.note === 'unrecognised') sheetUnrecognized++;
          else if (res.note === 'invalid' || res.note === 'ambiguous') sheetAmbiguous++;
          else if (res.note === 'swapped' || res.text !== before) {
            sheetFixed++;
            warnings.push(`  ✓ ${sheetName}!${ref}: "${before}" → "${res.text}"`);
          } else sheetOk++;
          ws[ref] = { t: 's', v: res.text };
        } else {
          sheetUnrecognized++;
          warnings.push(`  ⚠ ${sheetName}!${ref}: unexpected value "${cell.v}" in a date column — left as-is`);
        }
      });
    }
    const line = `── ${sheetName}: ${sheetFixed} fixed, ${sheetOk} already correct, ${sheetAmbiguous} ambiguous, ${sheetUnrecognized} unrecognised ──`;
    console.log(line);
    totalFixed += sheetFixed; totalOk += sheetOk; totalAmbiguous += sheetAmbiguous; totalUnrecognized += sheetUnrecognized;
  });

  return { wb, stats: { sheetFixed: totalFixed, sheetOk: totalOk, sheetAmbiguous: totalAmbiguous, sheetUnrecognized: totalUnrecognized, warnings } };
}

function main() {
  const input = process.argv[2];
  const output = process.argv[3];
  if (!input) {
    console.error('Usage: node fix_dates_cli.js <file.csv|xlsx> [out.xlsx]');
    process.exit(1);
  }
  if (!fs.existsSync(input)) {
    console.error(`File not found: ${input}`);
    process.exit(1);
  }
  const name = path.basename(input);
  const ext = name.toLowerCase().split('.').pop();
  const isCsv = ext === 'csv';

  let wb, stats;
  if (isCsv) {
    const text = fs.readFileSync(input, 'utf8');
    const r = buildWorkbookFromCsv(text, name);
    wb = r.wb; stats = r.stats;
    console.log(`── ${wb.SheetNames[0]}: ${stats.sheetFixed} fixed, ${stats.sheetOk} already correct, ${stats.sheetAmbiguous} ambiguous, ${stats.sheetUnrecognized} unrecognised ──`);
  } else {
    const buf = fs.readFileSync(input);
    const r = fixWorkbookFromXlsx(buf, name);
    wb = r.wb; stats = r.stats;
    console.log(`── Done: ${stats.sheetFixed} fixed, ${stats.sheetOk} already correct, ${stats.sheetAmbiguous} ambiguous, ${stats.sheetUnrecognized} unrecognised ──`);
  }

  if (stats.warnings.length) {
    console.log('Warnings:');
    stats.warnings.slice(0, 30).forEach(w => console.log(w));
    if (stats.warnings.length > 30) console.log(`…and ${stats.warnings.length - 30} more`);
  }

  const outName = output || path.join(path.dirname(input), name.replace(/\.[^.]+$/, '') + '_dates_fixed.xlsx');
  XLSX.writeFile(wb, outName);
  console.log(`Wrote ${outName}`);
}

if (require.main === module) main();

module.exports = {
  FNAME_RE, NUMERIC_COLS, DATE_COLS,
  parseCSV, expectedWindow, pickDate, normalizeDateTimeSmart,
  stripForcedQuote, excelDateToText, roundToSecond, toUnamb, parseDateTimeParts,
  buildWorkbookFromCsv, fixWorkbookFromXlsx
};
