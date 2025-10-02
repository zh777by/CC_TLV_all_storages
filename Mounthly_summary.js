/***** Цвета как на фото *****/
const HEADER_BG   = '#63D297';
const ODD_BG      = '#E8F8F2';
const EVEN_BG     = '#F4FBF8';
const HEADER_FONT = '#000000';

/***** Имена месячных листов теперь в формате MM/YYYY (пример: '09/2025') *****/
const SUMMARY_SHEET_NAME = 'PIVOT';
const OLD_SUMMARY_SHEET_NAMES = ['סיכום חודשי'];

/**
 * Сводная для активного месячного листа.
 */
function buildMonthlySummaryTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();

  const sheetMonth = _monthNumFromSheetName(sheet.getName());
  if (!sheetMonth) { ui.alert('Откройте месячный лист с именем в формате MM/YYYY (например 09/2025).'); return; }

  const cols = _findCols(sheet);
  if (!cols) { ui.alert('Не нашёл заголовки "נפתחה" и "מספר פריטים".'); return; }
  const { headerRow, colDate, colItems } = cols;

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) { ui.alert('После заголовка нет данных.'); return; }

  const width = Math.max(6, colDate, colItems);
  const rng   = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, width);
  const vals  = rng.getValues();
  const disp  = rng.getDisplayValues();

  const detected = _detectYearMonth(vals, disp, colDate, sheetMonth);
  if (!detected) { ui.alert('Не удалось определить месяц/год по данным листа.'); return; }
  const { year: TARGET_YEAR, month: TARGET_MONTH } = detected;
  const monthYearTag = `${_pad2(TARGET_MONTH)}/${TARGET_YEAR}`;

  const byDate = new Map();
  for (let i = 0; i < vals.length; i++) {
    const rowV = vals[i], rowD = disp[i];
    const aVal      = rowV[0];
    const itemsCell = rowV[colItems - 1];
    const dateVal   = rowV[colDate - 1];
    const dateDisp  = (rowD[colDate - 1] || '').trim();

    const ymd = _ymdFromValueOrDisplay(dateVal, dateDisp);
    if (!ymd || ymd.y !== TARGET_YEAR || ymd.m !== TARGET_MONTH) continue;

    const key = `${_pad2(ymd.d)}/${_pad2(ymd.m)}/${ymd.y}`;
    const dateObj = new Date(ymd.y, ymd.m - 1, ymd.d);
    const isSummaryRow = /^\d{2}\/\d{2}\/\d{4}$/.test(dateDisp);

    if (!byDate.has(key)) byDate.set(key, { countA: 0, maxA: 0, sumItems: 0, summaryItems: null, dateObj });
    const agg = byDate.get(key);

    const nA = _toNumber(aVal);
    if (Number.isFinite(nA)) { agg.countA++; if (nA > agg.maxA) agg.maxA = nA; }

    const nF = _toNumber(itemsCell);
    if (Number.isFinite(nF)) {
      if (isSummaryRow) agg.summaryItems = nF; else agg.sumItems += nF;
    }
  }
  if (byDate.size === 0) { SpreadsheetApp.getUi().alert(`Нет строк за ${monthYearTag}.`); return; }

  const rows = Array.from(byDate.entries())
    .sort((a, b) => a[1].dateObj - b[1].dateObj)
    .map(([key, v], idx) => {
      const orders = v.maxA > 0 ? v.maxA : v.countA;
      const items  = (v.summaryItems !== null ? v.summaryItems : v.sumItems);
      return [orders, items, key, idx + 1];
    });

  const lastDateStr = rows[rows.length - 1][2];
  const anchorRow = _findLastOccurrenceOfDateString(sheet, lastDateStr) || sheet.getLastRow();
  const startRow = anchorRow + 3, startCol = 1;

  sheet.getRange(startRow, startCol, 1, 4)
       .setValues([['ORDERS','ITEMS','DATE','DAYS']])
       .setFontWeight('bold');
  sheet.getRange(startRow + 1, startCol, rows.length, 4).setValues(rows);
  sheet.getRange(startRow + 1, startCol + 2, rows.length, 1).setNumberFormat('dd/mm/yyyy');

  const totalOrders = rows.reduce((s, r) => s + (Number(r[0]) || 0), 0);
  const totalItems  = rows.reduce((s, r) => s + (Number(r[1]) || 0), 0);
  const days = rows.length;
  const avgOrders = +(totalOrders / days).toFixed(2);
  const avgItems  = +(totalItems  / days).toFixed(2);

  const totalsStartRow = startRow + 1 + rows.length;
  sheet.getRange(totalsStartRow, startCol, 2, 4)
       .setValues([
         [totalOrders, totalItems, `${monthYearTag} (total)`,       'Total days'],
         [avgOrders,   avgItems,   `${monthYearTag} (average/day)`, days]
       ])
       .setFontWeight('bold');

  const bandRows = 1 + rows.length;
  const totalH   = bandRows + 2;

  sheet.getRange(startRow, startCol, 1, 4)
       .setBackground(HEADER_BG).setFontColor(HEADER_FONT)
       .setHorizontalAlignment('center');

  const dataColors = [];
  for (let r = 0; r < rows.length; r++) {
    const color = (r % 2 === 0) ? ODD_BG : EVEN_BG;
    dataColors.push([color, color, color, color]);
  }
  if (rows.length > 0) sheet.getRange(startRow + 1, startCol, rows.length, 4).setBackgrounds(dataColors);

  sheet.getRange(totalsStartRow, startCol, 2, 4)
       .setBackground(HEADER_BG).setFontColor(HEADER_FONT)
       .setHorizontalAlignment('center');

  sheet.getRange(startRow, startCol, totalH, 4).setHorizontalAlignment('center');
  sheet.autoResizeColumns(startCol, 4);

  _removeExistingChartsByTitlePrefix(sheet, 'ORDERS and ITEMS');
  const headerPlusData = bandRows;
  const dateRange   = sheet.getRange(startRow, startCol + 2, headerPlusData, 1);
  const itemsRange  = sheet.getRange(startRow, startCol + 1, headerPlusData, 1);
  const ordersRange = sheet.getRange(startRow, startCol + 0, headerPlusData, 1);

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dateRange)
    .addRange(itemsRange)
    .addRange(ordersRange)
    .setOption('title', `ORDERS and ITEMS — ${monthYearTag}`)
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { title: 'DATE', slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 10 } })
    .setOption('vAxis', { viewWindow: { min: 0 } })
    .setOption('bar', { groupWidth: '70%' })
    .setOption('series', {
      0: { color: '#D9534F', dataLabel: 'value', labelInLegend: 'items' },
      1: { color: '#007BFF', dataLabel: 'value', labelInLegend: 'orders' }
    })
    .setOption('trendlines', { 0: { type: 'linear', color: '#F5A5A5', lineWidth: 1, opacity: 0.6, visibleInLegend: false } })
    .setPosition(startRow, startCol + 6, 0, 0)
    .build();
  sheet.insertChart(chart);
}

/**
 * СВОДНАЯ ПО ВСЕМ МЕСЯЧНЫМ ЛИСТАМ → лист PIVOT.
 */
function buildAllMonthsSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const rows = [];
  for (const sh of sheets) {
    const mnum = _monthNumFromSheetName(sh.getName()); // теперь ждём 'MM/YYYY'
    if (!mnum) continue;

    const region = _findMonthlyTableRegion(sh);
    if (!region) continue;
    const { headerRow, firstCol, dataLen } = region;

    const ordVals = sh.getRange(headerRow + 1, firstCol + 0, dataLen, 1).getValues().flat();
    const itmVals = sh.getRange(headerRow + 1, firstCol + 1, dataLen, 1).getValues().flat();
    const dateDisp = sh.getRange(headerRow + 1, firstCol + 2, dataLen, 1).getDisplayValues().flat();

    const m = (dateDisp[0] || '').match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) continue;
    const mm = +m[2], yy = +m[3];

    const ordersTotal = ordVals.reduce((s,v)=>s+(Number(v)||0),0);
    const itemsTotal  = itmVals.reduce((s,v)=>s+(Number(v)||0),0);
    const days = dataLen;
    const avgO = +(ordersTotal/days).toFixed(2);
    const avgI = +(itemsTotal /days).toFixed(2);

    rows.push({ y: yy, m: mm, tag: `${_pad2(mm)}/${yy}`, days, ordersTotal, itemsTotal, avgO, avgI });
  }

  if (rows.length === 0) {
    SpreadsheetApp.getUi().alert('Не нашёл ни одной готовой месячной таблицы.');
    return;
  }

  rows.sort((a,b)=> a.y!==b.y ? a.y-b.y : a.m-b.m);

  // создаём/переименовываем лист сводной в PIVOT
  const ssObj = SpreadsheetApp.getActiveSpreadsheet();
  let sum = ssObj.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sum) {
    for (const old of OLD_SUMMARY_SHEET_NAMES) {
      const oldSheet = ssObj.getSheetByName(old);
      if (oldSheet) { oldSheet.setName(SUMMARY_SHEET_NAME); sum = oldSheet; break; }
    }
  }
  if (!sum) sum = ssObj.insertSheet(SUMMARY_SHEET_NAME); else sum.clear();

  const header = ['MONTH','DAYS','ORDERS (total)','ITEMS (total)','ORDERS (avg/day)','ITEMS (avg/day)'];
  const values = rows.map(r => [r.tag, r.days, r.ordersTotal, r.itemsTotal, r.avgO, r.avgI]);

  sum.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
  sum.getRange(2,1,values.length,header.length).setValues(values);

  sum.getRange(1,1,1,header.length)
     .setBackground(HEADER_BG).setFontColor(HEADER_FONT)
     .setHorizontalAlignment('center');

  const dataColors = [];
  for (let i=0;i<values.length;i++){
    const c = (i%2===0)?ODD_BG:EVEN_BG;
    dataColors.push([c,c,c,c,c,c]);
  }
  sum.getRange(2,1,values.length,header.length).setBackgrounds(dataColors);
  sum.getRange(1,1,values.length+1,header.length).setHorizontalAlignment('center');

  sum.getRange(2,2,values.length,1).setNumberFormat('0');
  sum.getRange(2,3,values.length,2).setNumberFormat('#,##0');
  sum.getRange(2,5,values.length,2).setNumberFormat('0.00');
  sum.autoResizeColumns(1, header.length);

  _removeExistingChartsByTitlePrefix(sum, 'ORDERS and ITEMS — all months');
  const monthRange = sum.getRange(1,1,values.length+1,1);
  const itemsRange = sum.getRange(1,4,values.length+1,1);
  const ordersRange= sum.getRange(1,3,values.length+1,1);

  const chart = sum.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(monthRange)
    .addRange(itemsRange)
    .addRange(ordersRange)
    .setOption('title', 'ORDERS and ITEMS — all months')
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { slantedText: true, slantedTextAngle: 45 })
    .setOption('vAxis', { viewWindow: { min: 0 } })
    .setOption('series', {
      0: { color: '#D9534F', dataLabel: 'value', labelInLegend: 'items' },
      1: { color: '#007BFF', dataLabel: 'value', labelInLegend: 'orders' }
    })
    .setPosition(1, header.length + 3, 0, 0)
    .build();
  sum.insertChart(chart);
}

/***** Совместимость со старым именем кнопки *****/
function buildJuneSummaryTable(){ return buildMonthlySummaryTable(); }

/***** Хелперы *****/

/** Имя листа должно быть строго в формате 'MM/YYYY'. Возвращает номер месяца 1..12 или null. */
function _monthNumFromSheetName(name){
  const m = String(name || '').trim().match(/^(\d{2})\/(\d{4})$/);
  if (!m) return null;
  const mm = Number(m[1]);
  return (mm >= 1 && mm <= 12) ? mm : null;
}

function _findCols(sheet){
  const lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
  const scanRows = Math.min(100, lastRow || 1);
  const grid = sheet.getRange(1,1,scanRows,lastCol||1).getDisplayValues();
  let colDate=null, colItems=null, headerRow=null;
  for (let r=0;r<grid.length;r++){
    for (let c=0;c<grid[r].length;c++){
      const t=(grid[r][c]||'').toString().trim();
      if (t==='נפתחה'){ colDate=c+1; headerRow=r+1; }
      if (t==='מספר פריטים'){ colItems=c+1; headerRow=headerRow||(r+1); }
      if (colDate && colItems) return { headerRow, colDate, colItems };
    }
  }
  return null;
}

function _ymdFromValueOrDisplay(val, disp){
  if (val instanceof Date) return { y: val.getFullYear(), m: val.getMonth()+1, d: val.getDate() };
  const m = (disp||'').match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? { d:+m[1], m:+m[2], y:+m[3] } : null;
}
function _toNumber(v){ if (typeof v==='number') return v; const s=(v==null?'':String(v)).replace(/\s/g,'').replace(/,/g,''); const n=Number(s); return Number.isFinite(n)?n:NaN; }
function _pad2(n){ return ('0'+n).slice(-2); }

function _detectYearMonth(vals, disp, colDate, preferMonth){
  const counts = new Map(), yearCounts = new Map();
  for (let i=0;i<vals.length;i++){
    const ymd = _ymdFromValueOrDisplay(vals[i][colDate-1], (disp[i][colDate-1]||'').trim());
    if (!ymd) continue;
    if (preferMonth && ymd.m !== preferMonth) continue;
    const key = `${ymd.y}-${_pad2(ymd.m)}`;
    counts.set(key,(counts.get(key)||0)+1);
    if (preferMonth) yearCounts.set(String(ymd.y),(yearCounts.get(String(ymd.y))||0)+1);
  }
  if (preferMonth){
    let bestYear=null,best=-1;
    for (const [y,c] of yearCounts.entries()){ if (c>best){ best=c; bestYear=Number(y);} }
    if (bestYear!=null) return { year: bestYear, month: preferMonth };
  }
  let bestKey=null,best=-1;
  for (const [k,c] of counts.entries()){ if (c>best){ best=c; bestKey=k; } }
  if (!bestKey) return null;
  const [yy,mm]=bestKey.split('-'); return { year:Number(yy), month:Number(mm) };
}

function _findLastOccurrenceOfDateString(sheet, dateStr){
  const lastRow=sheet.getLastRow(), lastCol=sheet.getLastColumn();
  if (lastRow===0 || lastCol===0) return null;
  const chunk=200;
  for (let start=lastRow; start>0; start-=chunk){
    const h=Math.min(chunk,start), top=start-h+1;
    const vals=sheet.getRange(top,1,h,lastCol).getDisplayValues();
    for (let i=vals.length-1;i>=0;i--){
      const row=vals[i];
      for (let j=0;j<row.length;j++){
        if ((row[j]||'').toString().trim().startsWith(dateStr)) return top+i;
      }
    }
  }
  return null;
}
function _removeExistingChartsByTitlePrefix(sheet, prefix){
  sheet.getCharts().forEach(ch=>{
    try{
      const t=String((ch.getOptions()&&ch.getOptions().title)||'');
      if (t.startsWith(prefix)) sheet.removeChart(ch);
    }catch(e){}
  });
}
function _findMonthlyTableRegion(sheet){
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) return null;
  const scanH = Math.min(600, lastRow);
  const start = lastRow - scanH + 1;
  const grid  = sheet.getRange(start, 1, scanH, lastCol).getDisplayValues();

  for (let r = grid.length - 1; r >= 0; r--) {
    for (let c = 0; c < grid[r].length - 3; c++) {
      const a = (grid[r][c]||'').toString().trim();
      const b = (grid[r][c+1]||'').toString().trim();
      const d = (grid[r][c+2]||'').toString().trim();
      const e = (grid[r][c+3]||'').toString().trim();
      if (a==='ORDERS' && b==='ITEMS' && d==='DATE' && e==='DAYS') {
        const headerRow = start + r, firstCol = c + 1;
        let len = 0;
        for (let rr = headerRow + 1; rr <= lastRow; rr++) {
          const s = (sheet.getRange(rr, firstCol + 2).getDisplayValue() || '').trim();
          if (!/^\d{2}\/\d{2}\/\d{4}$/.test(s)) break;
          len++;
        }
        return len > 0 ? { headerRow, firstCol, dataLen: len } : null;
      }
    }
  }
  return null;
}
