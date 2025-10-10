/***** SETTINGS *****/
const SRC_SHEET_NAME = 'Data source';   // source
const DATE_COL = 5;                    // column E (1-based)
const COPY_COL_START = 2;              // B
const COPY_COL_END = 6;                // F
const HEADER_ROW = 1;                  // header row
const BLANK_GAP = 2;                   // blank rows between blocks
const TOTAL_BG = '#b7e1cd';            // fill color for the total row (as requested)

/** Map: month number (0-11) -> sheet name in Hebrew */
const HE_MONTHS = [
  '01/2025','02/2025','03/2025','04/2025','05/2025','06/2025',
  '07/2025','08/2025','09/2025','10/2025','11/2025','12/2025'
];

function menuTransferByDate() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Введите дату', 'Формат: dd/MM/yyyy (например 08/09/2025)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const dateStr = res.getResponseText().trim();
  transferByDate(dateStr);
}

/***** CORE LOGIC *****/

/**
 * Transfers all rows with the specified date as a single block:
 * - B..F are copied consecutively
 * - A = 1..N
 * - block rows are grouped (the total row is NOT in the group)
 * - right under the block: E = date, F = SUM(F), the row is filled with TOTAL_BG across the entire width
 * - exactly BLANK_GAP blank rows between blocks
 */
function transferByDate(dateStr) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SRC_SHEET_NAME);
  if (!src) throw new Error(`Лист-источник "${SRC_SHEET_NAME}" не найден`);

  const { d } = parseDDMMYYYY(dateStr);
  const targetName = HE_MONTHS[d.getMonth()];
  const target = ensureTargetSheet(ss, targetName);

  const lastRow = src.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const width = COPY_COL_END - COPY_COL_START + 1;
  const allValues = src.getRange(HEADER_ROW + 1, COPY_COL_START, lastRow - HEADER_ROW, width).getValues();
  const dateValues = src.getRange(HEADER_ROW + 1, DATE_COL, lastRow - HEADER_ROW, 1).getValues();

  const rowsToCopy = [];
  for (let i = 0; i < dateValues.length; i++) {
    const cDate = coerceDate(dateValues[i][0]);
    if (sameDay(cDate, d)) {
      const row = allValues[i].slice();
      const eIdx = DATE_COL - COPY_COL_START;
      if (row[eIdx]) row[eIdx] = formatDateOnly(cDate); // date only in E
      rowsToCopy.push(row);
    }
  }

  if (rowsToCopy.length === 0) {
    SpreadsheetApp.getUi().alert(`Нет строк с датой ${dateStr}`);
    return;
  }

  // Block start (taking the offset into account)
  const startRow = nextBlockStartRow(target, COPY_COL_START, COPY_COL_END, BLANK_GAP);

  // 1) Insert B..F consecutively
  target.getRange(startRow, COPY_COL_START, rowsToCopy.length, width).setValues(rowsToCopy);

  // 2) Column A: 1..N
  const counter = Array.from({ length: rowsToCopy.length }, (_, i) => [i + 1]);
  target.getRange(startRow, 1, rowsToCopy.length, 1).setValues(counter);

  // 3) Group BLOCK rows (without the total)
  target
    .getRange(startRow, 1, rowsToCopy.length, target.getLastColumn())
    .shiftRowGroupDepth(1); // create row group

  // 4) Total row right under the block
  const totalsRowIndex = startRow + rowsToCopy.length;
  const dateText = formatDateOnly(d);
  const sumF = sumColumnF(rowsToCopy);

  // clear the row and set E and F
  target.getRange(totalsRowIndex, 1, 1, target.getLastColumn()).clearContent();
  target.getRange(totalsRowIndex, DATE_COL).setValue(dateText).setNumberFormat('dd/MM/yyyy'); // E
  target.getRange(totalsRowIndex, COPY_COL_END).setValue(sumF);                                // F

  // fill the total row ACROSS THE WHOLE ROW
  target.getRange(totalsRowIndex, 1, 1, target.getLastColumn()).setBackground(TOTAL_BG);
  
  SpreadsheetApp.getUi().alert(
    `Перенесено ${rowsToCopy.length} строк в лист "${targetName}". Итог: ${sumF}.`
  );
}

/**
 * Batch transfer for all dates (each block + grouping + total row).
 */
function transferAllDates() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SRC_SHEET_NAME);
  if (!src) throw new Error(`Лист-источник "${SRC_SHEET_NAME}" не найден`);

  const lastRow = src.getLastRow();
  if (lastRow <= HEADER_ROW) return;

  const width = COPY_COL_END - COPY_COL_START + 1;
  const allValues = src.getRange(HEADER_ROW + 1, COPY_COL_START, lastRow - HEADER_ROW, width).getValues();
  const dateValues = src.getRange(HEADER_ROW + 1, DATE_COL, lastRow - HEADER_ROW, 1).getValues();

  // group rows by “clean” date
  const grouped = {};
  for (let i = 0; i < dateValues.length; i++) {
    const cDate = coerceDate(dateValues[i][0]);
    if (!cDate) continue;
    const key = formatDateOnly(cDate);
    if (!grouped[key]) grouped[key] = [];
    const row = allValues[i].slice();
    const eIdx = DATE_COL - COPY_COL_START;
    if (row[eIdx]) row[eIdx] = key;
    grouped[key].push(row);
  }

  let totalRows = 0;
  for (const key of Object.keys(grouped).sort()) {
    const d = parseDDMMYYYY(key).d;
    const targetName = HE_MONTHS[d.getMonth()];
    const target = ensureTargetSheet(ss, targetName);

    const rows = grouped[key];
    const startRow = nextBlockStartRow(target, COPY_COL_START, COPY_COL_END, BLANK_GAP);

    // B..F
    target.getRange(startRow, COPY_COL_START, rows.length, width).setValues(rows);

    // A: 1..N
    const counter = Array.from({ length: rows.length }, (_, i) => [i + 1]);
    target.getRange(startRow, 1, rows.length, 1).setValues(counter);

    // Group block rows
    target
      .getRange(startRow, 1, rows.length, target.getLastColumn())
      .shiftRowGroupDepth(1);

    // Total row
    const totalsRowIndex = startRow + rows.length;
    const sumF = sumColumnF(rows);
    target.getRange(totalsRowIndex, 1, 1, target.getLastColumn()).clearContent();
    target.getRange(totalsRowIndex, DATE_COL).setValue(key).setNumberFormat('dd/MM/yyyy');
    target.getRange(totalsRowIndex, COPY_COL_END).setValue(sumF);
    target.getRange(totalsRowIndex, 1, 1, target.getLastColumn()).setBackground(TOTAL_BG);

    totalRows += rows.length;
  }

  SpreadsheetApp.getUi().alert(
    `Готово. Перенесено всего строк: ${totalRows}. Блоки сгруппированы; итоги добавлены; отступ = ${BLANK_GAP} пустых строк.`
  );
}

/***** UTILITIES *****/
function ensureTargetSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, 6).setValues([['A','B','C','D','E','F']]);
  return sh;
}

function nextBlockStartRow(sheet, colStart, colEnd, blankGap) {
  const last = sheet.getLastRow();
  if (last <= HEADER_ROW) return HEADER_ROW + 1 + blankGap;
  return last + 1 + blankGap;
}

function parseDDMMYYYY(str) {
  const m = String(str).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return { d, mIndex: d.getMonth() };
  }
  const d = new Date(str);
  return { d, mIndex: d.getMonth() };
}

function coerceDate(val) {
  if (val instanceof Date) return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  if (typeof val === 'number') {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  if (typeof val === 'string') {
    const s = val.split(' ')[0];
    const parts = s.split(/[\/.\-]/);
    if (parts.length === 3) {
      if (parts[2].length === 4) return new Date(parts[2], parts[1] - 1, parts[0]); // dd/mm/yyyy
      if (parts[0].length === 4) return new Date(parts[0], parts[1] - 1, parts[2]); // yyyy/mm/dd
    }
    const d = new Date(s);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

function sameDay(d1, d2) {
  return d1 && d2 &&
    d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate();
}

function formatDateOnly(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/** Sum of column F in an array of B..F rows */
function sumColumnF(rowsBF) {
  const fIdx = COPY_COL_END - COPY_COL_START; // index of F within the B..F slice
  let s = 0;
  for (const r of rowsBF) {
    const v = r[fIdx];
    const num = parseNumber(v);
    if (!isNaN(num)) s += num;
  }
  return s;
}

/** Parse a number from string/value, supports "1,5" and "1.5" */
function parseNumber(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const trimmed = v.trim().replace(/\s+/g, '');
    const normalized = trimmed.replace(',', '.');
    const n = Number(normalized);
    return isNaN(n) ? NaN : n;
  }
  return NaN;
}
