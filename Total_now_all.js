/**
 * Result: sheet "Total now all"
 * A: ITEM / DESCRIPTION (title case, merged case-insensitively)
 * B: SKU CC# as in cataloque (YU/BY/HH; if missing — ITEM from BISQUE)
 * C: BISQUE_IL (pcs)
 * D: YU_Storage (pcs)
 * E: BY_Storage (pcs)
 * F: HH_Storage (pcs)
 * G: TOTAL NOW ALL (pcs) = C+D+E+F (centered, bold; font color by condition)
 */
function buildTotalNowAll() {
  const TARGET_SHEET_NAME = 'Total now all';

  const SRC = {
    BISQUE: { id: 'https://docs.google.com/spreadsheets/d/1Lv02hdv9RF0sbIFhhIiCFdHXavsWk3UjrUXUygS_gJM/edit?gid=1644122405#gid=1644122405', sheetName: 'חדש' },
    YU:     { id: 'https://docs.google.com/spreadsheets/d/1U2yzVzaSGf2w3TN4Xgb783Rrrz_hb3IWYd9Q_IRy65g/edit?gid=1764326434#gid=1764326434', sheetName: 'STORAGE' },
    BY:     { id: 'https://docs.google.com/spreadsheets/d/1RUnWm1ruPwQULgfgbg4kkfOXbk0NRP5_uTBawfV57EA/edit?gid=1764326434#gid=1764326434', sheetName: 'STORAGE' },
    HH:     { id: 'https://docs.google.com/spreadsheets/d/1bI2ISWPc6_UCblCfRhHSA0LdjdEiZOMIa57YJeBTsgM/edit?gid=1764326434#gid=1764326434', sheetName: 'STORAGE' },
  };

  const HEAD = {
    descStrict: ['description','תיאור'],
    itemAny:    ['item','item name','name','sku','description','שם','פריט','תיאור'],
    totalAny:   ['total now (pcs)','total now','total','total_now','totalnow','total now all','סהכ','סה״כ','סה"כ'],
    skuCcAny:   [
      'sku cc# as in cataloque','sku cc as in cataloque',
      'sku cc# as in catalog','sku cc as in catalog',
      'sku cc# as in catalogue','sku cc as in catalogue'
    ],
    itemOnly:   ['item']
  };
  const MAX_SCAN_ROWS = 10;

  const normalizeId = (s) => {
    const str = String(s || '').trim();
    const m = str.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return m ? m[1] : str;
  };
  const norm = (s) => String(s || '').replace(/\s+/g,' ').replace(/[’‘']/g,'"').trim().toLowerCase();
  const nkey = (s) => String(s||'').normalize('NFKC').trim().replace(/\s+/g,' ').toLowerCase();

  function titleCase(s){
    return String(s||'').toLowerCase().trim().replace(/\s+/g,' ')
      .split(' ')
      .map(w => w.split('-').map(p => p? p[0].toUpperCase()+p.slice(1):p).join('-'))
      .join(' ');
  }

  function findHeaderPositions_(sh, namesArr) {
    const lastCol = sh.getLastColumn();
    const rows = Math.min(MAX_SCAN_ROWS, sh.getLastRow());
    if (!rows || !lastCol) return {};
    const vals = sh.getRange(1, 1, rows, lastCol).getDisplayValues();

    const want = Object.entries(namesArr).map(([key, arr]) => [key, new Set(arr)]);
    const pos = {};
    for (let r = 0; r < vals.length; r++) {
      for (let c = 0; c < lastCol; c++) {
        const cell = norm(vals[r][c]);
        for (const [key, set] of want) {
          if (!pos[key] && set.has(cell)) pos[key] = { col: c+1, row: r+1 };
        }
      }
    }
    return pos;
  }

  /**
   * Fast single-pass sheet reader:
   *  - sum: Map(key → TOTAL sum)
   *  - label: Map(key → original first label)
   *  - sku: Map(key → SKU CC#)
   *  - descItem: Map(key → fallback ITEM (BISQUE only))
   */
  function readSheetData(fileIdOrUrl, sheetName, needSku, needDescItem) {
    const ss = SpreadsheetApp.openById(normalizeId(fileIdOrUrl));
    const sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error(`Лист "${sheetName}" не найден.`);

    const pos = findHeaderPositions_(sh, {
      descStrict: HEAD.descStrict,
      itemAny:    HEAD.itemAny,
      totalAny:   HEAD.totalAny,
      skuCcAny:   HEAD.skuCcAny,
      itemOnly:   HEAD.itemOnly
    });

    const keyPos  = pos.descStrict || pos.itemAny;
    const totPos  = pos.totalAny;
    const skuPos  = needSku ? pos.skuCcAny : null;
    const itemPos = needDescItem ? pos.itemOnly : null;

    if (!keyPos || !totPos) {
      return { sum: new Map(), label: new Map(), sku: new Map(), descItem: new Map() };
    }

    const headerRow = Math.max(keyPos.row, totPos.row, skuPos?.row || 1, itemPos?.row || 1);
    const lastRow = sh.getLastRow();
    if (lastRow <= headerRow) {
      return { sum: new Map(), label: new Map(), sku: new Map(), descItem: new Map() };
    }

    const nRows = lastRow - headerRow;
    const cMin = Math.min(keyPos.col, totPos.col, skuPos?.col || keyPos.col, itemPos?.col || keyPos.col);
    const cMax = Math.max(keyPos.col, totPos.col, skuPos?.col || keyPos.col, itemPos?.col || keyPos.col);
    const width = cMax - cMin + 1;

    const block = sh.getRange(headerRow + 1, cMin, nRows, width).getValues();

    const idxKey  = keyPos.col  - cMin;
    const idxTot  = totPos.col  - cMin;
    const idxSku  = skuPos ? (skuPos.col - cMin) : null;
    const idxItem = itemPos ? (itemPos.col - cMin) : null;

    const sum = new Map();
    const label = new Map();
    const sku = new Map();
    const descItem = new Map();

    for (let i = 0; i < block.length; i++) {
      const row = block[i];
      const rawKey = String(row[idxKey] ?? '').trim();
      if (!rawKey) continue;
      const k = nkey(rawKey);

      const tVal = row[idxTot];
      const num = (typeof tVal === 'number') ? tVal : Number(String(tVal || '').replace(/\s+/g,'').replace(/,/g,'.'));
      const val = isNaN(num) ? 0 : num;
      sum.set(k, (sum.get(k) || 0) + val);

      if (!label.has(k)) label.set(k, rawKey);

      if (idxSku != null) {
        const s = String(row[idxSku] ?? '').trim();
        if (s && !sku.has(k)) sku.set(k, s);
      }

      if (idxItem != null) {
        const it = String(row[idxItem] ?? '').trim();
        if (it && !descItem.has(k)) descItem.set(k, it);
      }
    }

    return { sum, label, sku, descItem };
  }

  // Reading sources (each — one call)
  const bisque = readSheetData(SRC.BISQUE.id, SRC.BISQUE.sheetName, false, true);
  const yu     = readSheetData(SRC.YU.id,     SRC.YU.sheetName,     true,  false);
  const by     = readSheetData(SRC.BY.id,     SRC.BY.sheetName,     true,  false);
  const hh     = readSheetData(SRC.HH.id,     SRC.HH.sheetName,     true,  false);

  // Full set of keys
  const keySet = new Set([
    ...bisque.sum.keys(), ...yu.sum.keys(), ...by.sum.keys(), ...hh.sum.keys(),
    ...yu.sku.keys(), ...by.sku.keys(), ...hh.sku.keys(), ...bisque.descItem.keys()
  ]);

  // Sort by a “nice” label from BISQUE (if present)
  const keys = Array.from(keySet).sort((a,b) => {
    const la = bisque.label.get(a) || a;
    const lb = bisque.label.get(b) || b;
    return la.localeCompare(lb, undefined, {sensitivity:'base'});
  });

  const header = [
    'ITEM / DESCRIPTION',
    'SKU CC# as in cataloque',
    'BISQUE_IL (pcs)',
    'YU_Storage (pcs)',
    'BY_Storage (pcs)',
    'HH_Storage (pcs)',
    'TOTAL NOW ALL (pcs)'
  ];

  const rowsAF = keys.map(k => {
    const shown = bisque.label.get(k) || k;
    const displayName = titleCase(shown);
    const sku =
      (yu.sku.get(k) || '').trim() ||
      (by.sku.get(k) || '').trim() ||
      (hh.sku.get(k) || '').trim() ||
      (bisque.descItem.get(k) || '').trim();
    return [
      displayName,
      sku,
      bisque.sum.get(k) || 0,
      yu.sum.get(k) || 0,
      by.sum.get(k) || 0,
      hh.sum.get(k) || 0
    ];
  });

  // Render to the target sheet
  const ssTarget = SpreadsheetApp.getActiveSpreadsheet();
  const shTarget = ssTarget.getSheetByName(TARGET_SHEET_NAME) || ssTarget.insertSheet(TARGET_SHEET_NAME);

  shTarget.clearFormats();
  shTarget.clear();

  // Header
  shTarget.getRange(1, 1, 1, header.length)
    .setValues([header])
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  if (rowsAF.length) {
    // Data A–F
    shTarget.getRange(2, 1, rowsAF.length, 6).setValues(rowsAF);

    // Column G – formula and format
    shTarget.getRange(2, 7, rowsAF.length, 1)
      .setFormulasR1C1(Array(rowsAF.length).fill(['=SUM(RC[-4]:RC[-1])']))
      .setNumberFormat('0')
      .setHorizontalAlignment('center')
      .setFontWeight('bold');

    // Number format in C:F, alignment B:G
    shTarget.getRange(2, 3, rowsAF.length, 4).setNumberFormat('0');
    shTarget.getRange(2, 2, rowsAF.length, 6).setHorizontalAlignment('center');

    // Conditional coloring for G
    const gRange = shTarget.getRange(2, 7, rowsAF.length, 1);
    const gEndRow = 1 + rowsAF.length;
    let rules = shTarget.getConditionalFormatRules() || [];
    rules = rules.filter(rule => {
      const ranges = rule.getRanges ? rule.getRanges() : [];
      return !ranges.some(r => {
        if (r.getSheet().getName() !== TARGET_SHEET_NAME) return false;
        const c1 = r.getColumn();
        const c2 = c1 + r.getNumColumns() - 1;
        const r1 = r.getRow();
        const r2 = r1 + r.getNumRows() - 1;
        const colsOverlap = (c1 <= 7 && 7 <= c2);
        const rowsOverlap = !(r2 < 2 || r1 > gEndRow);
        return colsOverlap && rowsOverlap;
      });
    });

    const orangeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($G2<20,$G2>=10)')
      .setFontColor('#FFA500')
      .setRanges([gRange])
      .build();

    const redRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2<10')
      .setFontColor('#FF0000')
      .setRanges([gRange])
      .build();

    rules.push(orangeRule, redRule);
    shTarget.setConditionalFormatRules(rules);
  }

  // Freeze header row
  shTarget.setFrozenRows(1);

  // Manual banded fill (compatible everywhere)
  shTarget.getBandings().forEach(b => b.remove());
  const totalCols = 7;
  const HEADER_BG = '#AFC4E2';
  const ROW1_BG   = '#EEF4FB';
  const ROW2_BG   = '#FFFFFF';

  shTarget.getRange(1, 1, 1, totalCols).setBackground(HEADER_BG);
  if (rowsAF.length > 0) {
    const bodyRange = shTarget.getRange(2, 1, rowsAF.length, totalCols);
    const bg = [];
    for (let i = 0; i < rowsAF.length; i++) {
      bg.push(Array(totalCols).fill((i % 2 === 0) ? ROW1_BG : ROW2_BG));
    }
    bodyRange.setBackgrounds(bg);
  }

  // Auto width
  shTarget.autoResizeColumns(1, header.length);
}

/**
 * Sort by column G (TOTAL NOW ALL)
 */
function sortTotalNowAll_(asc) {
  const TARGET_SHEET_NAME = 'Total now all';
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sh) throw new Error(`Лист "${TARGET_SHEET_NAME}" не найден`);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  sh.getRange(2, 1, lastRow - 1, 7).sort([{ column: 7, ascending: !!asc }]);
  sh.setFrozenRows(1);
}
function sortTotalNowAllDesc() { sortTotalNowAll_(false); }
function sortTotalNowAllAsc()  { sortTotalNowAll_(true);  }

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TotalNow')
    .addItem('Collect Data', 'buildTotalNowAll')
    .addSeparator()
    .addItem('Sort by TOTAL (desc)', 'sortTotalNowAllDesc')
    .addItem('Sort by TOTAL (asc)',  'sortTotalNowAllAsc')
    .addToUi();
}
