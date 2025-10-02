/* ===== CONFIG ===== */
const TEMPLATE_SHEET = 'August 2025 Delivery Tracker';   // your template tab
const DISPATCH_LIVE  = 'Dispatch_Live';                  // tab with =IMPORTRANGE(...,"Export!A1:Z")

// How far down to apply dropdowns automatically
const VALIDATION_ROWS_ON_SYNC = 5000;
const VALIDATION_ROWS_ON_FIX  = 10000;

// Column header names in your month sheets
const COLS = {
  order:        'Order #',
  broker:       'Broker',
  demo:         'Demographic',
  driver:       'Driver',
  truck:        'Truck',
  driverTruck:  'Driver Truck',

  puDate:       'Pick Up Date',
  puTime1:      'Appt Time ',       // pickup (first)
  puAddr:       'Pick Up Address',

  delDate:      'Delivery Date',
  delTime2a:    'Appt Time .1',     // delivery (preferred)
  delTime2b:    'Appt Time ',       // delivery (second)
  delAddr:      'Delivery Address',

  deadhead:     'Dead Head Miles',
  loaded:       'Loaded Miles ',
  loadedAlt:    'Loaded Miles',
  paid:         'Paid Miles',

  // manual/admin to carry over
  ppg:          'Price Per Gallon',
  mpg:          'Miles Per Gallon',
  bol:          'BOL RCVD',
  ow:           'OW',
  detention:    'Detention',
  brokerRateA:  'Broker Rate',
  brokerRateB:  'Boker Rate',       // tolerated typo
  invDate:      'Invoice Date',
  estPayout:    'Est Payout Date',
  invNum:       'Invoice #',

  // calculated columns to fill from row 2
  calc_estFuel: 'Est Fuel Cost',
  calc_drvPay:  'Driver Pay',
  calc_dispPay: 'Dispatcher Pay',
  calc_runCost: 'Run Total Cost',
  calc_company: 'Company Earnings'
};

// ===== INVOICE CONFIG =====
// Put the Google Sheets File ID of the 03-Invoice_Master_Workbook here:
const INVOICE_BOOK_ID = '1vNjSvLf7KpJ2lWTiJfBBpRrAduevIugi5_0NlNo046g';
const INVOICE_TEMPLATE_SHEET = 'Invoice_Template'; // a tab inside the invoice workbook to copy

// Cell mappings inside the invoice template:
const CELL_CLIENT    = 'B9';    // Broker/Client name
const CELL_DATE      = 'A11';   // Delivery Date
const CELL_ORDER     = 'B11';   // Container/PO# = Order #
const CELL_RATE      = 'D11';   // <-- Broker Rate column (was L11)
const CELL_ROW_TOTAL = 'M11';   // <-- Row total formula =SUM(D11:L11) 
const CELL_COL_TOTAL = 'M16';   // <-- Column total formula =SUM(M11:M15)


/* ===== MENU (clean labels) ===== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('TB&S Month')
    .addItem('New Month (clone template)…', 'newMonthClone')
    .addItem('Sync From Dispatch',          'syncFromDispatch')
    .addSeparator()
    .addItem('Carry-Over Invoice Fields',   'carryOverInvoiceFields')
    .addItem('Diagnose Carry-Over',         'diagnoseCarryOver')
    .addSeparator()
    .addItem('Apply Dropdowns Wide',        'applyDropdownsWideOnActive')
    .addSeparator()
    .addItem('Create Invoice from Order…',  'createInvoiceFromOrder')
    .addToUi();
}


function createInvoiceFromOrder(){
  const ui = SpreadsheetApp.getUi();
  if (!INVOICE_BOOK_ID || INVOICE_BOOK_ID === '1vNjSvLf7KpJ2lWTiJfBBpRrAduevIugi5_0NlNo046gPUT_INVOICE_FILE_ID_HERE'){
    ui.alert('Invoice workbook ID is not set. Please set INVOICE_BOOK_ID at the top of the script.');
    return;
  }
  if (!INVOICE_BOOK_ID || INVOICE_BOOK_ID.length < 20) {
  ui.alert('Invoice workbook ID is not set. Please set INVOICE_BOOK_ID at the top of the script.');
  return;
}
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  // ask for Order #
  const resp = ui.prompt('Create Invoice', 'Enter Order # to invoice:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const order = resp.getResponseText().trim();
  if (!order){ ui.alert('Order # is required.'); return; }

  // find order in current tab
  const hdr = headers_(sh), m = hmap_(hdr);
  const cOrder = m[COLS.order];
  const cBroker = m[COLS.broker];
  const cDelDate = m[COLS.delDate];
  const cBrokerRate = (m[COLS.brokerRateA] || m[COLS.brokerRateB] || 0);

  if (!cOrder || !cBroker || !cDelDate){
    ui.alert('Missing one of the required columns (Order #, Broker, Delivery Date) in this tab.');
    return;
  }

  const startRow = 3;
  const max = Math.max(sh.getMaxRows() - (startRow-1), 1);
  const orders = sh.getRange(startRow, cOrder, max, 1).getDisplayValues();
  let rowIndex = -1;
  for (let i=0;i<orders.length;i++){
    const v = String(orders[i][0]||'').trim();
    if (!v) break;
    if (v === order){ rowIndex = startRow + i; break; }
  }
  if (rowIndex === -1){ ui.alert('Order # not found on this tab.'); return; }

  const broker     = sh.getRange(rowIndex, cBroker).getDisplayValue();
  const delDate    = sh.getRange(rowIndex, cDelDate).getValue();
  const brokerRate = cBrokerRate ? sh.getRange(rowIndex, cBrokerRate).getValue() : '';

  // open invoice workbook
  const invSS = SpreadsheetApp.openById(INVOICE_BOOK_ID);

  const tmpl  = invSS.getSheetByName(INVOICE_TEMPLATE_SHEET);

  if (!tmpl){ ui.alert('Invoice template tab not found in the invoice workbook.'); return; }

  // create new invoice tab
  const safeBroker = broker.replace(/[\\\/\?\*\[\]:]/g,' ').trim();
  const newName = `${safeBroker || 'Broker'}_${order}`;
  let inv = invSS.getSheetByName(newName);
  if (inv) invSS.deleteSheet(inv);
  inv = tmpl.copyTo(invSS).setName(newName);

  // write cells
  inv.getRange(CELL_CLIENT).setValue(broker);
  inv.getRange(CELL_DATE).setValue(delDate);
  inv.getRange(CELL_ORDER).setValue(order);
  if (CELL_RATE) inv.getRange(CELL_RATE).setValue(brokerRate);
  
  // make currency cells look right
inv.getRange('D11:K11').setNumberFormat('$#,##0.00');
inv.getRange(CELL_ROW_TOTAL).setNumberFormat('$#,##0.00');
inv.getRange(CELL_COL_TOTAL).setNumberFormat('$#,##0.00');

// if RATE might come with a "$", coerce to number
const rateNum = Number(String(brokerRate).replace(/[^0-9.\-]/g,''));
inv.getRange(CELL_RATE).setValue(rateNum || brokerRate);


  // set formulas
  if (CELL_ROW_TOTAL) inv.getRange(CELL_ROW_TOTAL).setFormula('=SUM(D11:L11)');
  if (CELL_COL_TOTAL) inv.getRange(CELL_COL_TOTAL).setFormula('=SUM(M11:M15)');

  // optional: number formats
  if (CELL_DATE) inv.getRange(CELL_DATE).setNumberFormat('mm/dd/yyyy');
  

  ui.alert(`Invoice sheet created: "${newName}" in 03-Invoice_Master_Workbook.`);
}

// Admin-only tools are available in code but NOT shown in menu:
    // .addItem('Repair Validations on Active Sheet', 'repairValidationsOnActive')
    // .addItem('Fix Row 2 (format + protect)',       'fixRow2FormatAndProtection')}


/* ===== HELPERS ===== */
function _rowSignatureFromSource_(r){
  // Build a short signature from the Export row (same mapped fields we write)
  // r indices: 0 Order, 1 Broker, 2 Demo, 3 Driver/Truck, 4 PU Date, 5 PU Time, 6 PU Addr, 7 Del Date, 8 Del Time, 9 Del Addr, 10 Dead, 11 Loaded, 12 Paid
  const dt = (v)=>String(v||'').trim().toUpperCase();
  return [
    dt(r[1]), dt(r[2]), dt(r[3]),
    dt(r[4]), dt(r[5]), dt(r[6]),
    dt(r[7]), dt(r[8]), dt(r[9]),
    dt(r[10]), dt(r[11]), dt(r[12])
  ].join('|');
}

function _rowSignatureFromSheet_(vals){
  // vals = array of strings from the **target sheet** for the mapped columns
  const dt = (v)=>String(v||'').trim().toUpperCase();
  return vals.map(dt).join('|');
}

function _ensureSyncLog_(){
  const ss = SpreadsheetApp.getActive();
  let log = ss.getSheetByName('Sync_Log');
  if (!log){
    log = ss.insertSheet('Sync_Log');
    log.appendRow(['Timestamp','Month Tab','Added','Removed','Changed','Added Orders','Removed Orders','Changed Orders']);
  }
  return log;
}

function headers_(sh){ return sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]; }
function hmap_(hdr){ const m={}; hdr.forEach((h,i)=>{ if(h && !(h in m)) m[h]=i+1; }); return m; }
function idxOf_(hdr, name, nth){ let c=0; for (let i=0;i<hdr.length;i++){ if(hdr[i]===name){ c++; if(c===nth) return i+1; } } return 0; }
function firstOf_(map, names){ for (const n of names){ if(map[n]) return map[n]; } return 0; }
function normTxt_(v){ return String(v==null?'':v).replace(/\u00A0/g,' ').replace(/\s+/g,' ').trim().toUpperCase(); }

function clearBelowRow2_(sh){
  const rows = sh.getMaxRows(); const cols = sh.getLastColumn();
  if (rows>2) sh.getRange(3,1,rows-2,cols).clearContent();
}
function setCol_(sh, startRow, col, values){
  if (!col || !values || !values.length) return;
  sh.getRange(startRow, col, values.length, 1).setValues(values);
}

// copy Row 2 formulas down n rows (wrapped in IFERROR) — preserves row 2 formatting/content
function fillCalcFormulas_(sh, nRows){
  if (nRows<=0) return;
  const hdr = headers_(sh), m = hmap_(hdr);
  const calcCols = [
    m[COLS.calc_estFuel],
    m[COLS.calc_drvPay],
    m[COLS.calc_dispPay],
    m[COLS.calc_runCost],
    m[COLS.calc_company]
  ].filter(Boolean);

  // also include ANY col whose row 2 has a formula (safety net)
  const lastCol = sh.getLastColumn();
  for (let c=1;c<=lastCol;c++){
    const f2 = sh.getRange(2,c).getFormulaR1C1();
    if (f2 && calcCols.indexOf(c)===-1) calcCols.push(c);
  }

  const startRow = 3;
  calcCols.forEach(c=>{
    const f2 = sh.getRange(2, c).getFormulaR1C1();
    if (!f2) return;
    const wrapped = '=IFERROR(' + f2.substring(1) + ',"")';
    sh.getRange(startRow, c, nRows, 1).setFormulaR1C1(wrapped);
    const fmt = sh.getRange(2, c).getNumberFormat();
    if (fmt) sh.getRange(startRow, c, nRows, 1).setNumberFormat(fmt);
  });
}

// Build validation from template row 3 and paint it down many rows
// IMPORTANT: we set allowInvalid(true) so existing text won’t show red “invalid” flags.
function applyDropdownsFromTemplate_(targetSh, templateName, colIndexes, startRow, rows) {
  const ss   = targetSh.getParent();
  const tmpl = ss.getSheetByName(templateName);
  if (!tmpl) return;

  colIndexes.filter(Boolean).forEach(c=>{
    const baseRule = tmpl.getRange(3, c).getDataValidation();
    if (!baseRule) return;
    const rule = baseRule.copy().setAllowInvalid(true).build(); // warning-only, still shows dropdown options
    const grid = Array.from({length: rows}, ()=>[rule]);
    targetSh.getRange(startRow, c, rows, 1).setDataValidations(grid);
  });
}

// Clear validations on specific columns
function clearValidationsForCols_(sh, cols, startRow, numRows) {
  cols.filter(Boolean).forEach(c => {
    sh.getRange(startRow, c, Math.max(numRows, 1), 1).clearDataValidations();
  });
}

// Audit: highlight written rows + append to Sync_Log
function logSyncAudit_(sh, name, rows, startRow) {
  const ss = SpreadsheetApp.getActive();
  if (rows.length > 0) {
    sh.getRange(startRow, 1, rows.length, sh.getLastColumn())
      .setBackground('#fff2cc'); // light yellow
  }
  let log = ss.getSheetByName('Sync_Log');
  if (!log) { log = ss.insertSheet('Sync_Log'); log.appendRow(['Timestamp','Month Tab','Order #s']); }
  const orderList = rows.map(r => r[0]).join(', ');
  log.appendRow([new Date(), name, orderList]);
}

// Row 2 protection (no formatting change)
function protectRow2_(sh, strict){
  const rng = sh.getRange(2,1,1,sh.getLastColumn());
  (sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || []).forEach(p=>{
    const r = p.getRange(); if (r.getRow()===2 && r.getNumRows()===1) p.remove();
  });
  const prot = rng.protect();
  prot.setDescription('Row 2 formula row — do not edit');
  if (strict){
    prot.setWarningOnly(false);
    const me = Session.getEffectiveUser();
    prot.removeEditors(prot.getEditors());
    prot.addEditor(me);
  } else {
    prot.setWarningOnly(true);
  }
}

/* ===== NEW MONTH (keeps row 2 formatting) ===== */
function newMonthClone(){
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const tmpl = ss.getSheetByName(TEMPLATE_SHEET);
  if (!tmpl){ ui.alert('Template tab not found: '+TEMPLATE_SHEET); return; }

  const resp = ui.prompt('New Month', 'Enter new month tab name (e.g., September_2025):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const newName = resp.getResponseText().trim();
  if (!newName){ ui.alert('Month name is required.'); return; }

  const copy = tmpl.copyTo(ss);        // copies ALL formatting & protections
  copy.setName(newName);
  ss.setActiveSheet(copy);

  // Keep header + row2 formulas; clear rows 3+ (content only)
  clearBelowRow2_(copy);

  // Ensure row 2 is protected (doesn’t affect formatting)
  protectRow2_(copy, /*strict=*/false);

  // Ensure dropdowns are painted far down initially
  const hdr = headers_(copy), m = hmap_(hdr);
  applyDropdownsFromTemplate_(copy, TEMPLATE_SHEET,
    [m['Broker'], m['Status'], m['BOL RCVD']].filter(Boolean),
    3,
    VALIDATION_ROWS_ON_SYNC
  );

  ui.alert(`Created "${newName}". Now run "Sync From Dispatch".`);
}

/* ===== SYNC FROM DISPATCH ===== */
function syncFromDispatch(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const name = sh.getName();

  const dl = ss.getSheetByName(DISPATCH_LIVE);
  if (!dl) throw new Error(`Missing sheet "${DISPATCH_LIVE}". Create it with IMPORTRANGE to Export!A1:Z.`);
  const dVals = dl.getDataRange().getValues();
  if (dVals.length < 2){ SpreadsheetApp.getActive().toast('Dispatch has no rows.', 'TB&S', 4); return; }

  const rows = dVals.slice(1).filter(r => String(r.join('')).trim()!=='');
  const n = rows.length;
  if (n === 0){ SpreadsheetApp.getActive().toast('No non-empty rows from Dispatch.', 'TB&S', 4); return; }

  const hdr = headers_(sh);
  const m   = hmap_(hdr);

  // Resolve columns
  const cOrder   = m[COLS.order];
  const cBroker  = m[COLS.broker];
  const cDemo    = m[COLS.demo];
  const cDrv     = m[COLS.driver];
  const cTrk     = m[COLS.truck];
  const cDrvTrk  = m[COLS.driverTruck];

  const cPUDate  = m[COLS.puDate];
  const cPUTime  = idxOf_(hdr, COLS.puTime1, 1);
  const cPUAddr  = m[COLS.puAddr];

  const cDelDate = m[COLS.delDate];
  const cDelTime = idxOf_(hdr, COLS.delTime2a, 1) || idxOf_(hdr, COLS.delTime2b, 2);
  const cDelAddr = m[COLS.delAddr];

  const cDead    = m[COLS.deadhead];
  const cLoaded  = firstOf_(m, [COLS.loaded, COLS.loadedAlt]);
  const cPaid    = m[COLS.paid];

  const startRow = 3;
  const usedRows = Math.max(sh.getLastRow() - 2, 0);

  // ===== Build "old" snapshot for audit (by Order #) =====
  const oldMap = {};      // order -> signature
  const oldSet = new Set();
  if (usedRows > 0 && cOrder){
    const takeCols = [cBroker,cDemo,(cDrv||cDrvTrk),(cTrk||0),cPUDate,cPUTime,cPUAddr,cDelDate,cDelTime,cDelAddr,cDead,cLoaded,cPaid].filter(Boolean);
    const oldOrders = sh.getRange(startRow, cOrder, usedRows, 1).getDisplayValues().map(r=>String(r[0]||'').trim());
    const oldCols   = takeCols.map(c => sh.getRange(startRow, c, usedRows, 1).getDisplayValues());
    for (let i=0;i<usedRows;i++){
      const ord = oldOrders[i];
      if (!ord) continue;
      oldSet.add(ord);
      const vals = oldCols.map(col => col[i][0]);
      oldMap[ord] = _rowSignatureFromSheet_(vals);
    }
  }

  // Columns we will write
  const writeCols = [
    cOrder,cBroker,cDemo,(cDrv||cDrvTrk),(cTrk||null),
    cPUDate,cPUTime,cPUAddr,cDelDate,cDelTime,cDelAddr,
    cDead,cLoaded,cPaid
  ].filter(Boolean);

  // Temporarily remove validations on written columns (avoid rejects)
  clearValidationsForCols_(sh, writeCols, startRow, Math.max(usedRows, n));

  // Clear old content (only used range)
  writeCols.forEach(c => { if (usedRows > 0) sh.getRange(startRow, c, usedRows, 1).clearContent(); });

  // Parse driver/truck
  const drv = [], trk = [], drvTrk = [];
  rows.forEach(r=>{
    const s = String(r[3]||'').trim(); // "Driver/Truck[/...]"
    if (!s){ drv.push(['']); trk.push(['']); drvTrk.push(['']); return; }
    const parts = s.split('/');
    const d = (parts[0]||'').trim();
    const t = parts.length>1 ? parts.slice(1).join('/').trim() : '';
    drv.push([d]); trk.push([t]); drvTrk.push([d + (t?'/'+t:'')]);
  });

  // Write columns
  function setCol(col, arr){ if (col) sh.getRange(startRow, col, n, 1).setValues(arr); }

  setCol(cOrder,  rows.map(r=>[r[0]]));
  setCol(cBroker, rows.map(r=>[r[1]]));
  setCol(cDemo,   rows.map(r=>[r[2]]));

  if (cDrvTrk) setCol(cDrvTrk, drvTrk); else { setCol(cDrv, drv); setCol(cTrk, trk); }

  setCol(cPUDate, rows.map(r=>[r[4]]));
  setCol(cPUTime, rows.map(r=>[r[5]]));
  setCol(cPUAddr, rows.map(r=>[r[6]]));

  setCol(cDelDate, rows.map(r=>[r[7]]));
  setCol(cDelTime, rows.map(r=>[r[8]]));
  setCol(cDelAddr, rows.map(r=>[r[9]]));

  setCol(cDead,   rows.map(r=>[r[10]]));
  setCol(cLoaded, rows.map(r=>[r[11]]));
  setCol(cPaid,   rows.map(r=>[r[12]]));

  // Date/time formats
  if (cPUDate) sh.getRange(startRow, cPUDate, n, 1).setNumberFormat('mm/dd/yyyy');
  if (cDelDate) sh.getRange(startRow, cDelDate, n, 1).setNumberFormat('mm/dd/yyyy');
  if (cPUTime)  sh.getRange(startRow, cPUTime,  n, 1).setNumberFormat('HH:mm');
  if (cDelTime) sh.getRange(startRow, cDelTime, n, 1).setNumberFormat('HH:mm');

  // Fill calculated columns from Row 2
  fillCalcFormulas_(sh, n);

  // Re-apply intended dropdowns for the rows we wrote
  applyDropdownsFromTemplate_(sh, TEMPLATE_SHEET,
    [m['Broker'], m['Status'], m['BOL RCVD']].filter(Boolean),
    startRow,
    n
  );

  // ===== Build "new" snapshot (from source rows we just wrote) =====
  const newSet = new Set();
  const newMap = {};   // order -> signature (from source rows)
  rows.forEach(r=>{
    const ord = String(r[0]||'').trim();
    if (!ord) return;
    newSet.add(ord);
    newMap[ord] = _rowSignatureFromSource_(r);
  });

  // Diff
  const added = [], removed = [], changed = [];
  newSet.forEach(o => { if (!oldSet.has(o)) added.push(o); });
  oldSet.forEach(o => { if (!newSet.has(o)) removed.push(o); });
  newSet.forEach(o => {
    if (oldSet.has(o) && oldMap[o] !== newMap[o]) changed.push(o);
  });

  // Audit highlight (simple highlight of all synced rows)
  if (n > 0) sh.getRange(startRow, 1, n, sh.getLastColumn()).setBackground('#fff2cc');

  // Log
  const log = _ensureSyncLog_();
  log.appendRow([
    new Date(),
    name,
    added.length, removed.length, changed.length,
    added.join(', '),
    removed.join(', '),
    changed.join(', ')
  ]);

  SpreadsheetApp.getActive().toast(`Synced ${n} rows • Added ${added.length} • Removed ${removed.length} • Changed ${changed.length}`, 'TB&S', 6);
}

/* ===== Apply Dropdowns Wide (manual fix for existing tabs) ===== */
function applyDropdownsWideOnActive(){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const hdr = headers_(sh), m = hmap_(hdr);

  applyDropdownsFromTemplate_(sh, TEMPLATE_SHEET,
    [m['Broker'], m['Status'], m['BOL RCVD']].filter(Boolean),
    3,
    VALIDATION_ROWS_ON_FIX
  );
  SpreadsheetApp.getActive().toast('Dropdowns applied wide on this sheet.', 'TB&S', 4);
}

/* ===== Carry-Over (robust, leading zeros ignored for numeric Order #) ===== */
function _normHeader_(s) {
  return String(s||'').toLowerCase()
    .replace(/\u00a0/g,' ')
    .replace(/\s+/g,' ')
    .trim()
    .replace(/[^a-z0-9#]/g,'');
}
function _headerAliasMap_(headers) {
  const m = {};
  headers.forEach((h,i)=>{ const k=_normHeader_(h||''); if (k && !(k in m)) m[k]=i+1; });
  return m;
}
function _colByAliases_(aliasMap, names) {
  for (const n of names) { const c = aliasMap[_normHeader_(n)]; if (c) return c; }
  return 0;
}
// numeric order numbers match with leading zeros ignored
function _normOrderKey_(v){
  let s = String(v==null?'':v).trim();
  if (s === '') return '';
  if (typeof v === 'number') return String(Math.floor(v));
  if (/^\d+$/.test(s)) return String(parseInt(s,10)); // "0206925" -> "206925"
  return s.toUpperCase();
}

function carryOverInvoiceFields(){
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  const resp = ui.prompt('Previous Month', 'Enter previous month tab name (e.g., August 2025 Delivery Tracker):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const prevName = resp.getResponseText().trim();
  const prev = ss.getSheetByName(prevName);
  if (!prev){ ui.alert('Previous month tab not found: '+prevName); return; }

  // Current & previous header maps
  const hdrT = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const mapT = _headerAliasMap_(hdrT);
  const hdrP = prev.getRange(1,1,1,prev.getLastColumn()).getDisplayValues()[0];
  const mapP = _headerAliasMap_(hdrP);

  // Resolve Order # in both
  const cOrder = _colByAliases_(mapT, ['Order #','Order#','Order']);
  const pOrder = _colByAliases_(mapP, ['Order #','Order#','Order']);
  if (!cOrder || !pOrder) { ui.alert('Missing "Order #" in one of the tabs.'); return; }

  // Determine N by scanning Order # down
  const startRow = 3;
  const max = Math.max(sh.getMaxRows() - (startRow-1), 1);
  const curOrders = sh.getRange(startRow, cOrder, max, 1).getValues();
  let n = 0; for (let i=0;i<curOrders.length;i++) if (String(curOrders[i][0]).trim()!=='') n=i+1;
  if (n === 0) { ui.alert('No Order # rows found. Run "Sync From Dispatch" first.'); return; }

  // Build prev-month lookup by normalized key (leading zeros ignored)
  const pVals = prev.getDataRange().getValues();
  const lookup = {};
  for (let i=1;i<pVals.length;i++) {
    const k = _normOrderKey_(pVals[i][pOrder-1]);
    if (k) lookup[k] = pVals[i];
  }

  // Which columns to carry (with header aliasing)
  const spec = [
    // miles & fuel
    { label:'Dead Head Miles', tgt:_colByAliases_(mapT,['Dead Head Miles','Deadhead Miles']),  src:_colByAliases_(mapP,['Dead Head Miles','Deadhead Miles']) },
    { label:'Loaded Miles',    tgt:_colByAliases_(mapT,['Loaded Miles','Loaded Miles ']),       src:_colByAliases_(mapP,['Loaded Miles','Loaded Miles ']) },
    { label:'Paid Miles',      tgt:_colByAliases_(mapT,['Paid Miles']),                         src:_colByAliases_(mapP,['Paid Miles']) },
    { label:'Price Per Gallon',tgt:_colByAliases_(mapT,['Price Per Gallon','PPG']),             src:_colByAliases_(mapP,['Price Per Gallon','PPG']) },
    { label:'Miles Per Gallon',tgt:_colByAliases_(mapT,['Miles Per Gallon','MPG']),             src:_colByAliases_(mapP,['Miles Per Gallon','MPG']) },

    // dropdown / admin
    { label:'BOL RCVD',        tgt:_colByAliases_(mapT,['BOL RCVD','BOL Received']),            src:_colByAliases_(mapP,['BOL RCVD','BOL Received']) },
    { label:'Status',          tgt:_colByAliases_(mapT,['Status']),                             src:_colByAliases_(mapP,['Status']) },

    // money & billing
    { label:'Broker Rate',     tgt:_colByAliases_(mapT,['Broker Rate','Boker Rate']),           src:_colByAliases_(mapP,['Broker Rate','Boker Rate']) },
    { label:'Invoice Date',    tgt:_colByAliases_(mapT,['Invoice Date']),                       src:_colByAliases_(mapP,['Invoice Date']) },
    { label:'Est Payout Date', tgt:_colByAliases_(mapT,['Est Payout Date','Estimated Payout Date']), src:_colByAliases_(mapP,['Est Payout Date','Estimated Payout Date']) },
    { label:'Invoice #',       tgt:_colByAliases_(mapT,['Invoice #','Invoice Number']),         src:_colByAliases_(mapP,['Invoice #','Invoice Number']) },
  ].filter(s => s.tgt && s.src);

  if (!spec.length) { ui.alert('No matching invoice/admin columns found.'); return; }

  // Temporarily relax validation on targets we’ll write
  const tgtCols = spec.map(s=>s.tgt);
  tgtCols.forEach(c => sh.getRange(startRow, c, n, 1).clearDataValidations());

  // Set values using normalized order key
  spec.forEach(s => {
    const out = new Array(n);
    for (let i=0;i<n;i++){
      const key = _normOrderKey_(curOrders[i][0]);
      const row = key && lookup[key] ? lookup[key] : null;
      out[i] = [ row ? row[s.src-1] : '' ];
    }
    sh.getRange(startRow, s.tgt, n, 1).setValues(out);
  });

  // Re-apply dropdowns wide so copied values still show options (allowInvalid=true)
  const hdr = headers_(sh), m = hmap_(hdr);
  applyDropdownsFromTemplate_(sh, TEMPLATE_SHEET,
    [m['Status'], m['BOL RCVD']].filter(Boolean),
    3,
    VALIDATION_ROWS_ON_SYNC
  );

  SpreadsheetApp.getActive().toast('Carry-over (invoice/admin) complete.', 'TB&S', 4);
}

/* ===== Diagnose Carry-Over ===== */
function diagnoseCarryOver(){
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  const resp = ui.prompt('Previous Month (diagnose)', 'Enter previous month tab name:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const prevName = resp.getResponseText().trim();
  const prev = ss.getSheetByName(prevName);
  if (!prev){ ui.alert('Previous month tab not found: ' + prevName); return; }

  const hdrT = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const hdrP = prev.getRange(1,1,1,prev.getLastColumn()).getDisplayValues()[0];
  const mapT = _headerAliasMap_(hdrT);
  const mapP = _headerAliasMap_(hdrP);

  const cOrder = _colByAliases_(mapT,['Order #','Order#','Order']);
  const pOrder = _colByAliases_(mapP,['Order #','Order#','Order']);
  if (!cOrder || !pOrder) { ui.alert('Missing "Order #" column.'); return; }

  const startRow = 3;
  const max = Math.max(sh.getMaxRows() - (startRow-1), 1);
  const curOrders = sh.getRange(startRow, cOrder, max, 1).getValues();
  let n = 0; for (let i=0;i<curOrders.length;i++) if (String(curOrders[i][0]).trim()!=='') n=i+1;

  const pVals = prev.getDataRange().getValues();
  const lookup = {};
  for (let i=1;i<pVals.length;i++) { const nk = _normOrderKey_(pVals[i][pOrder-1]); if (nk) lookup[nk] = true; }

  let matches = 0;
  for (let i=0;i<n;i++) { const k = _normOrderKey_(curOrders[i][0]); if (k && lookup[k]) matches++; }

  ui.alert(`Diagnose:\nRows with Order # in current: ${n}\nRows that will match in previous (normalized): ${matches}`);
}

/* ===== Admin-only helpers (not on menu) ===== */
function repairValidationsOnActive() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const m = hmap_(headers);
  const startRow = 3;

  // Clear ALL validations below Row 2
  sh.getRange(startRow, 1, Math.max(sh.getMaxRows()-2,1), sh.getLastColumn()).clearDataValidations();

  // Re-apply from template down to a wide range (allowInvalid=true inside helper)
  applyDropdownsFromTemplate_(sh, TEMPLATE_SHEET,
    [m['Broker'], m['Status'], m['BOL RCVD']].filter(Boolean),
    3,
    VALIDATION_ROWS_ON_FIX
  );

  SpreadsheetApp.getActive().toast('Validations repaired & applied wide.', 'TB&S', 4);
}
function fixRow2FormatAndProtection(){
  const sh = SpreadsheetApp.getActiveSheet();
  protectRow2_(sh, /*strict=*/false);
  SpreadsheetApp.getActive().toast('Row 2 protected (format unchanged).', 'TB&S', 4);
}
