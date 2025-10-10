/***** SETTINGS *****/
const VALID_SHEETS = ["COLORS", "STORAGE"];
const DATA_START_ROW = 3; // the row where items start

/***** UTILITIES *****/
function columnToLetter(column) {
  let letter = "";
  while (column > 0) {
    let temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
function _norm(v){ return String(v || "").toLowerCase().trim(); }

/** Find the (1-based) index of the "TOTAL NOW" column using row 1 */
function getTotalNowCol(sheet) {
  const header1 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = header1.findIndex(h => String(h).trim() === "TOTAL NOW");
  if (idx === -1) throw new Error(`Колонка "TOTAL NOW" не найдена`);
  return idx + 1;
}

/** Blocks map: supports 2-, 3-, and 4-column blocks; indices are 0-based */
function getBlocksMap(headerRow2Values) {
  const h = headerRow2Values.map(_norm);
  const res = [];
  for (let i = 0; i < h.length; i++) {
    const a = h[i], b = h[i+1], c = h[i+2], d = h[i+3];

    // 2-column: (in… | out… | out to floor…) + Total
    if (((a && a.startsWith("in")) || (a && a.startsWith("out")) || (a && a.startsWith("out to floor"))) && b === "total") {
      res.push({ startCol: i, size: 2, labels: [headerRow2Values[i], headerRow2Values[i+1]] });
      i += 1;
      continue;
    }

    // 3-column: in… | out to floor… | Total
    if ((a && a.startsWith("in")) &&
        (b && b.startsWith("out to floor")) &&
        c === "total") {
      res.push({ startCol: i, size: 3, labels: [headerRow2Values[i], headerRow2Values[i+1], headerRow2Values[i+2]] });
      i += 2;
      continue;
    }

    // 4-column: in… | out… | out to floor… | Total
    if ((a && a.startsWith("in")) &&
        (b && b.startsWith("out")) &&
        (c && c.startsWith("out to floor")) &&
        d === "total") {
      res.push({ startCol: i, size: 4, labels: [
        headerRow2Values[i], headerRow2Values[i+1],
        headerRow2Values[i+2], headerRow2Values[i+3]
      ]});
      i += 3;
      continue;
    }
  }
  return res;
}

/** First column of the very first block (1-based). If there are no blocks — fallback F=6; otherwise the column to the left of the first Total. */
function getFirstBlockStartCol(sheet) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return 6;
  const header2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const blocks = getBlocksMap(header2);
  if (blocks.length) return blocks[0].startCol + 1; // 0-based -> 1-based

  const h = header2.map(_norm);
  const firstTotalIdx = h.findIndex(v => v === "total");
  if (firstTotalIdx > 0) return firstTotalIdx; // column to the left of the first Total (1-based)
  return 6;
}

/** TOTAL NOW formula: take the value from the last column where row 2 equals "Total" */
function buildTotalNowFormulaA1(row, firstBlockStartCol, lastCol) {
  const firstColLetter = columnToLetter(firstBlockStartCol);
  const lastColLetter  = columnToLetter(lastCol);
  return `=IFERROR(INDEX(${row}:${row}, MAX(FILTER(COLUMN(${firstColLetter}$2:${lastColLetter}$2), ${firstColLetter}$2:${lastColLetter}$2="Total"))), "")`;
}

/** Numeric format for a column (so the status bar always shows a sum) */
function enforceNumberFormatForColumn(sheet, startRow, col, numRows) {
  if (numRows <= 0) return;
  sheet.getRange(startRow, col, numRows, 1).setNumberFormat("0.############");
}

/** Apply numeric validation (>=0 by default) to a range.
 *  allowInvalid=true — show a warning but do NOT block (so formulas pass).
 */
function applyNumberValidation(range, minZero = true, allowInvalid = true) {
  const b = SpreadsheetApp.newDataValidation();
  const builder = minZero
    ? b.requireNumberGreaterThanOrEqualTo(0)
    : b.requireNumberBetween(-1e12, 1e12);
  range.setDataValidation(builder.setAllowInvalid(allowInvalid).build());
}

/** Remove protections if present */
function removeProtections(sheet) {
  try { (sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || []).forEach(p => p.remove()); } catch (e) {}
  try { (sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET) || []).forEach(p => p.remove()); } catch (e) {}
}

/** Conditional formatting for TOTAL NOW: <20 orange, <10 red */
function applyTotalNowConditionalFormatting(sheet, totalNowCol) {
  const rules = sheet.getConditionalFormatRules() || [];
  const startRow = DATA_START_ROW;
  const rowsCount = Math.max(sheet.getLastRow() - startRow + 1, 1);
  const totalNowRange = sheet.getRange(startRow, totalNowCol, rowsCount, 1);

  const orangeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(20).setFontColor("#FFA500").setRanges([totalNowRange]).build();

  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(10).setFontColor("#FF0000").setRanges([totalNowRange]).build();

  const targetKey = totalNowRange.getA1Notation() + "@" + sheet.getName();
  const filtered = rules.filter(r => {
    const rgs = r.getRanges().map(rg => rg.getA1Notation() + "@" + rg.getSheet().getName());
    return !rgs.includes(targetKey);
  });

  sheet.setConditionalFormatRules([...filtered, orangeRule, redRule]);
}

/***** ADDING A NEW DAILY BLOCK *****/
/**
 * mode: "IN" | "OUT" | "out to floor"
 * Inserts TWO columns: <mode> | Total
 * New Total is calculated from the previous Total: +IN, −OUT, −out to floor
 * The input column gets numeric format and validation (>=0); its contents are cleared.
 */
function addNewDayBlock(mode, sheetName) {
  if (!VALID_SHEETS.includes(sheetName)) return;
  mode = String(mode || "").toLowerCase().trim();
  if (!["in", "out", "out to floor"].includes(mode)) {
    throw new Error(`Некорректный режим: "${mode}". Должно быть "IN", "OUT" или "out to floor".`);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Лист "${sheetName}" не найден`);
  removeProtections(sheet);

  const totalRows = sheet.getLastRow();
  const lastColBefore = sheet.getLastColumn();
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

  // Add 2 new columns: <mode> | Total
  sheet.insertColumnsAfter(lastColBefore, 2);
  const startCol = lastColBefore + 1;
  const valueCol = startCol;
  const totalCol = startCol + 1;

  // Row 1: date spanning 2 columns
  sheet.getRange(1, startCol, 1, 2)
    .merge()
    .setValue(date)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Row 2: labels
  const label = (mode === "in") ? "IN" : (mode === "out") ? "OUT" : "out to floor";
  sheet.getRange(2, valueCol).setValue(label);
  sheet.getRange(2, totalCol).setValue("Total");

  // Block header formatting
  sheet.getRange(2, valueCol, 1, 2)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  try { sheet.autoResizeRows(2, 1); } catch (e) {}

  // Safety: normalize the entire row 2 alignment
  try { normalizeHeaderRow2Alignment(sheetName); } catch (e) {}

  // Formulas & formatting if there are data rows
  if (totalRows >= DATA_START_ROW) {
    const rowsCount = totalRows - DATA_START_ROW + 1;

    // 1) Prepare the input column: clear + format + validation (>=0, but allowInvalid=true)
    const inputRange = sheet.getRange(DATA_START_ROW, valueCol, rowsCount, 1);
    inputRange.clearContent();
    enforceNumberFormatForColumn(sheet, DATA_START_ROW, valueCol, rowsCount);
    applyNumberValidation(inputRange, true, true); // warn but don’t block (to allow formulas)

    // 2) Formula Total = prevTotal ± value
    const prevTotalCol = lastColBefore;
    const formulas = [];
    for (let r = DATA_START_ROW; r <= totalRows; r++) {
      const prevA1  = sheet.getRange(r, prevTotalCol).getA1Notation();
      const valueA1 = sheet.getRange(r, valueCol).getA1Notation();
      const sign = (mode === "in") ? "+" : "-";
      formulas.push([`=IFERROR(VALUE(TRIM(${prevA1})),0)${sign}IFERROR(VALUE(TRIM(${valueA1})),0)`]);
    }
    sheet.getRange(DATA_START_ROW, totalCol, rowsCount, 1).setFormulas(formulas);
    enforceNumberFormatForColumn(sheet, DATA_START_ROW, totalCol, rowsCount);

    // 3) TOTAL NOW from the rightmost "Total"
    const totalNowCol = getTotalNowCol(sheet);
    const firstBlockStartCol = getFirstBlockStartCol(sheet);
    const totalNowFormulas = [];
    const lastCol = sheet.getLastColumn();
    for (let r = DATA_START_ROW; r <= totalRows; r++) {
      totalNowFormulas.push([buildTotalNowFormulaA1(r, firstBlockStartCol, lastCol)]);
    }
    sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).setFormulas(totalNowFormulas);

    // 4) Conditional formatting for TOTAL NOW
    applyTotalNowConditionalFormatting(sheet, totalNowCol);
  }

  SpreadsheetApp.flush();
}

/***** RESTORING FORMULAS IN ALL "Total"
 * Rules:
 *  - 4-col.: Total = N(in) + N(out) - N(out to floor) - N(prevTotal)
 *  - 3-col.: Total = N(in) + N(out to floor) - N(prevTotal)
 *  - 2-col.: Total = N(prevTotal) ± N(value), sign from the label on the left (in => +, out => -)
 * N(x) := IFERROR(VALUE(TRIM(x)),0)
*****/
function ensureFormulasInAllTotals(sheet) {
  const totalRows = sheet.getLastRow();
  const totalCols = sheet.getLastColumn();
  if (totalRows < DATA_START_ROW || totalCols === 0) return;

  const headerRow2 = sheet.getRange(2, 1, 1, totalCols).getValues()[0];
  const h2 = headerRow2.map(_norm);
  const rowsCount = totalRows - DATA_START_ROW + 1;

  for (let j = 1; j <= totalCols; j++) {
    if (_norm(headerRow2[j - 1]) !== "total") continue;

    const is4col =
      j >= 4 &&
      (h2[j - 4] && h2[j - 4].startsWith("in")) &&
      (h2[j - 3] && h2[j - 3].startsWith("out")) &&
      (h2[j - 2] && h2[j - 2].startsWith("out to floor"));

    const is3col =
      !is4col &&
      j >= 3 &&
      (h2[j - 3] && h2[j - 3].startsWith("in")) &&
      (h2[j - 2] && h2[j - 2].startsWith("out to floor"));

    const range = sheet.getRange(DATA_START_ROW, j, rowsCount, 1);

    if (is4col) {
      const r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-4])),0) + IFERROR(VALUE(TRIM(R[0]C[-3])),0) - IFERROR(VALUE(TRIM(R[0]C[-2])),0) - IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
      range.setFormulasR1C1(Array.from({ length: rowsCount }, () => [r1c1]));
      continue;
    }

    if (is3col) {
      const r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-3])),0) + IFERROR(VALUE(TRIM(R[0]C[-2])),0) - IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
      range.setFormulasR1C1(Array.from({ length: rowsCount }, () => [r1c1]));
      continue;
    }

    const mode = _norm(headerRow2[j - 2] || "");
    const sign = mode.startsWith("in") ? "+" : (mode.startsWith("out") ? "-" : null);
    if (!sign) continue;
    const r1c1 = `=IFERROR(VALUE(TRIM(R[0]C[-2])),0)${sign}IFERROR(VALUE(TRIM(R[0]C[-1])),0)`;
    range.setFormulasR1C1(Array.from({ length: rowsCount }, () => [r1c1]));
  }
}

/***** UPDATE THE CURRENT DAILY BLOCK (2/3/4-column) *****/
function updateDayBlock(sheetName) {
  if (!VALID_SHEETS.includes(sheetName)) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Лист "${sheetName}" не найден`);
  removeProtections(sheet);

  const startRow = DATA_START_ROW;
  const totalRows = sheet.getLastRow();
  const totalCols = sheet.getLastColumn();
  if (totalRows < startRow || totalCols === 0) return;

  const headerRow2Raw = sheet.getRange(2, 1, 1, totalCols).getValues()[0];
  const h2 = headerRow2Raw.map(_norm);

  // all "Total" columns (1-based)
  const totalColIdxs = [];
  for (let i = 0; i < h2.length; i++) if (h2[i] === "total") totalColIdxs.push(i + 1);
  if (!totalColIdxs.length) return;

  const newestTotalCol = totalColIdxs[totalColIdxs.length - 1];
  const nRows = totalRows - startRow + 1;

  // determine block type by headers to the left
  const is4col =
    newestTotalCol >= 4 &&
    (h2[newestTotalCol - 4] && h2[newestTotalCol - 4].startsWith("in")) &&
    (h2[newestTotalCol - 3] && h2[newestTotalCol - 3].startsWith("out")) &&
    (h2[newestTotalCol - 2] && h2[newestTotalCol - 2].startsWith("out to floor"));

  const is3col =
    !is4col &&
    newestTotalCol >= 3 &&
    (h2[newestTotalCol - 3] && h2[newestTotalCol - 3].startsWith("in")) &&
    (h2[newestTotalCol - 2] && h2[newestTotalCol - 2].startsWith("out to floor"));

  if (is4col) {
    const formulasR1C1 = Array.from({ length: nRows }, () => [
      `=IFERROR(VALUE(TRIM(R[0]C[-4])),0) + IFERROR(VALUE(TRIM(R[0]C[-3])),0) - IFERROR(VALUE(TRIM(R[0]C[-2])),0) - IFERROR(VALUE(TRIM(R[0]C[-1])),0)`
    ]);
    sheet.getRange(startRow, newestTotalCol, nRows, 1).setFormulasR1C1(formulasR1C1);
  } else if (is3col) {
    const formulasR1C1 = Array.from({ length: nRows }, () => [
      `=IFERROR(VALUE(TRIM(R[0]C[-3])),0) + IFERROR(VALUE(TRIM(R[0]C[-2])),0) - IFERROR(VALUE(TRIM(R[0]C[-1])),0)`
    ]);
    sheet.getRange(startRow, newestTotalCol, nRows, 1).setFormulasR1C1(formulasR1C1);
  } else {
    const modeRaw = headerRow2Raw[newestTotalCol - 2] || "";
    const mode = _norm(modeRaw);
    const sign = mode.startsWith("in") ? "+" : (mode.startsWith("out") ? "-" : null);
    if (!sign) return;
    const formulasR1C1 = Array.from({ length: nRows }, () => [
      `=IFERROR(VALUE(TRIM(R[0]C[-2])),0)${sign}IFERROR(VALUE(TRIM(R[0]C[-1])),0)`
    ]);
    sheet.getRange(startRow, newestTotalCol, nRows, 1).setFormulasR1C1(formulasR1C1);
  }

  // As a safeguard: restore formulas in all other "Total"
  ensureFormulasInAllTotals(sheet);

  // Rebuild TOTAL NOW (take the value from the rightmost "Total")
  const totalNowCol = getTotalNowCol(sheet);
  const firstBlockStartCol = getFirstBlockStartCol(sheet);
  const lastColLetter  = columnToLetter(totalCols);
  const firstColLetter = columnToLetter(firstBlockStartCol);

  const totalNowFormulas = [];
  for (let r = startRow; r <= totalRows; r++) {
    totalNowFormulas.push([`=IFERROR(INDEX(${r}:${r}, MAX(FILTER(COLUMN(${firstColLetter}$2:${lastColLetter}$2), ${firstColLetter}$2:${lastColLetter}$2="Total"))), "")`]);
  }
  sheet.getRange(startRow, totalNowCol, nRows, 1).setFormulas(totalNowFormulas);

  // Styling and numeric format for value/Total
  applyTotalNowConditionalFormatting(sheet, totalNowCol);
  enforceNumberFormatForColumn(sheet, startRow, newestTotalCol - 1, nRows); // value
  enforceNumberFormatForColumn(sheet, startRow, newestTotalCol,     nRows); // Total

  SpreadsheetApp.flush();
}

/***** HEADER ALIGNMENT (ROW 2) *****/
function normalizeHeaderRow2Alignment(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Лист "${sheetName}" не найден`);

  const lastCol = sh.getLastColumn();
  if (!lastCol) return;

  const hdr = sh.getRange(2, 1, 1, lastCol);
  hdr.setHorizontalAlignment("center")
     .setVerticalAlignment("middle")
     .setWrap(true);

  try { sh.autoResizeRows(2, 1); } catch (e) {}
}

/***** SIMPLIFY/RELAX VALIDATION FOR EXISTING BLOCKS *****/
function relaxValidationForInputColumns(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < DATA_START_ROW || !lastCol) return;

  const rowsCount = lastRow - DATA_START_ROW + 1;
  const h2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(_norm);

  for (let c = 1; c <= lastCol; c++) {
    const head = h2[c - 1];
    if (!head) continue;
    if (head.startsWith("in") || head.startsWith("out") || head.startsWith("out to floor")) {
      const rng = sheet.getRange(DATA_START_ROW, c, rowsCount, 1);
      applyNumberValidation(rng, true, true); // warn, don’t block
    }
  }
}

/***** UI AND INITIALIZATION *****/
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  VALID_SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    // Always rewrite the action list in A1
    ensureActionDropdown(sh);

    // Align header (row 2)
    try { normalizeHeaderRow2Alignment(name); } catch (e) {}

    // Relax validation in input columns so formulas aren’t blocked
    try { relaxValidationForInputColumns(sh); } catch (e) {}
  });
}

/** Dropdown in A1: same for all sheets */
function ensureActionDropdown(sheet) {
  const actions = [
    "➕ Add new day block (IN)",
    "➕ Add new day block (OUT)",
    "➕ Add new day block (out to floor)",
    "🔁 Update new day block",
  ];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(actions, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange("A1").setDataValidation(rule).setValue("");
}

/***** EDIT HANDLER WITH DEBOUNCE AGAINST DUPLICATES *****/
function _debounceOnce_(key, windowMs) {
  const props = PropertiesService.getDocumentProperties();
  const now = Date.now();
  const raw = props.getProperty('debounce:' + key);
  if (raw) {
    const ts = parseInt(raw, 10) || 0;
    if (now - ts < windowMs) return false; // too soon — duplicate
  }
  props.setProperty('debounce:' + key, String(now));
  return true;
}

/** Coerce input value to a number (supports , and .). Returns {ok,value} */
function _coerceToNumber_(raw) {
  const s = String(raw || "").trim();
  if (s === "") return { ok: true, value: "" }; // allow empty
  const normalized = s.replace(/\s+/g, "").replace(",", ".");
  const num = Number(normalized);
  if (Number.isFinite(num)) return { ok: true, value: num };
  return { ok: false, value: "" };
}

function onEdit(e) {
  if (!e || !e.range || !e.range.getSheet) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!VALID_SHEETS.includes(sheetName)) return;

  const a1 = e.range.getA1Notation();

  // 1) Handle dropdown in A1
  if (a1 === "A1") {
    const val = String(e.value || "");
    if (!val) return;

    // Anti-duplicate #1
    const liveBefore = String(sheet.getRange("A1").getValue() || "");
    if (liveBefore !== val) return;

    // Anti-duplicate #2
    const ok = _debounceOnce_(sheet.getSheetId() + '|' + val, 2500);
    if (!ok) return;

    // Serialize (take lock)
    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);

    try {
      // Anti-duplicate #3
      const liveNow = String(sheet.getRange("A1").getValue() || "");
      if (liveNow !== val) return;

      if (val === "➕ Add new day block (IN)") {
        addNewDayBlock("IN", sheetName);
      } else if (val === "➕ Add new day block (OUT)") {
        addNewDayBlock("OUT", sheetName);
      } else if (val === "➕ Add new day block (out to floor)") {
        addNewDayBlock("out to floor", sheetName);
      } else if (val === "🔁 Update new day block") {
        updateDayBlock(sheetName);
      }

      // clear the dropdown
      sheet.getRange("A1").setValue("");

    } finally {
      lock.releaseLock();
    }
    return; // handled A1
  }

  // 2) Strict coercion to number only for editable input cells (in/out/out to floor)
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < DATA_START_ROW) return; // only rows in the data area are editable

  // the label in row 2 above the edited column
  const header2val = String(sheet.getRange(2, col).getValue() || "");
  const head = _norm(header2val);
  const isInputCol = head.startsWith("in") || head.startsWith("out") || head.startsWith("out to floor");
  if (!isInputCol) return;

  // if the user entered a formula — allow it
  const hasFormula = !!e.range.getFormula();
  if (hasFormula) {
    e.range.setNumberFormat("0.############"); // format result as a number
    return;
  }

  // numeric coercion for plain input
  const r = _coerceToNumber_(e.value);
  if (!r.ok) {
    e.range.setValue(""); // reset invalid input
    try { SpreadsheetApp.getActive().toast("Введите число (разделитель: . или ,).", "Неверный ввод", 3); } catch (err) {}
    return;
  }

  if (r.value === "") {
    // empty — leave empty, but enforce numeric format
    e.range.setNumberFormat("0.############");
    return;
  }

  // valid number: write as number and format
  e.range.setValue(r.value);
  e.range.setNumberFormat("0.############");
}

/***** SERVICE UTILITIES *****/
/** One-time fix of formulas in all Total and TOTAL NOW */
function fixAllTotalsAndTotalNow(sheetName) {
  if (!VALID_SHEETS.includes(sheetName)) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Лист "${sheetName}" не найден`);

  const totalRows = sheet.getLastRow();
  if (totalRows < DATA_START_ROW) return;

  ensureFormulasInAllTotals(sheet);

  const totalNowCol = getTotalNowCol(sheet);
  const firstBlockStartCol = getFirstBlockStartCol(sheet);
  const lastCol = sheet.getLastColumn();
  const rowsCount = totalRows - DATA_START_ROW + 1;

  const totalNowFormulas = [];
  for (let r = DATA_START_ROW; r <= totalRows; r++) {
    totalNowFormulas.push([buildTotalNowFormulaA1(r, firstBlockStartCol, lastCol)]);
  }
  sheet.getRange(DATA_START_ROW, totalNowCol, rowsCount, 1).setFormulas(totalNowFormulas);
  applyTotalNowConditionalFormatting(sheet, totalNowCol);
}

/** Remove installable onEdit triggers if any remain */
function removeInstallableOnEditTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getEventType && t.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** One-off utility: relax validations everywhere (useful after updating the script) */
function onceAllowFormulasEverywhere() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  VALID_SHEETS.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh) relaxValidationForInputColumns(sh);
  });
}
