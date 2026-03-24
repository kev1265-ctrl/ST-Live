// ============================================================
// FILE 2 OF 4 — Setup.gs
// Supertrend Live Tracker v1.1
// Run firstTimeSetup() ONCE to build all sheets.
// Run setupTriggers() ONCE to schedule the scripts.
//
// v1.1 CHANGES:
//   - Config sheet: "Volume confirmation" + "SMA trend" sections
//     replaced by "Entry filters (bracket system)" section
//   - BRACKET_FILTER checkbox at row 19 (replaces VOL_FILTER)
//   - Row 20 is now an informational note (thresholds in script)
//   - VOL_LOOKBACK (row 21) and SMA_PERIOD (row 26) kept
// ============================================================


// ════════════════════════════════════════════════════════════
// PUBLIC ENTRY POINTS
// ════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📈 Live Tracker')
    .addItem('🔧 First-Time Setup (run once)',    'firstTimeSetup')
    .addItem('⏱  Setup Triggers (run once)',       'setupTriggers')
    .addItem('🛑 Remove All Triggers',             'removeTriggers')
    .addSeparator()
    .addItem('🌙 Run Nightly Batch Now',           'runNightlyBatch')
    .addItem('🔄 Run Daytime Sweep Now',           'runDaytimeSweep')
    .addSeparator()
    .addItem('📋 Import Tickers from Clipboard',  'showTickerImportInfo')
    .addToUi();
}

function firstTimeSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _createConfigSheet(ss);
  _createLiveTrackerSheet(ss);
  _createWatchlistSheet(ss);
  _createTickerListSheet(ss);
  _createJournalSheet(ss);
  _createBTPlaceholder(ss);
  _orderSheets(ss);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete!\n\n'
    + 'Next steps:\n'
    + '1. Paste your ticker list into 📋 Ticker List (col A, format: EXCHANGE:TICKER)\n'
    + '2. Edit ⚙️ Config — set your email address\n'
    + '3. Run "⏱ Setup Triggers" from the menu\n'
    + '4. Run "🌙 Run Nightly Batch Now" for the first population\n\n'
    + 'Optional: Copy your Supertrend_BT tab from the old file into 🔍 Supertrend_BT.'
  );
}

function setupTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('runNightlyBatch').timeBased().everyDays(1).atHour(21).create();
  ScriptApp.newTrigger('runDaytimeSweep').timeBased().everyMinutes(30).create();
  SpreadsheetApp.getUi().alert(
    '⏱ Triggers Active!\n\n'
    + 'Nightly batch : daily at 9 PM (set project timezone to America/New_York)\n'
    + 'Daytime sweep : every 30 minutes (skips weekends & outside market hours automatically)\n\n'
    + 'IMPORTANT: In Apps Script → Project Settings → Script timezone → set to America/New_York'
  );
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['runNightlyBatch', 'runDaytimeSweep'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function showTickerImportInfo() {
  SpreadsheetApp.getUi().alert(
    '📋 Ticker Import Instructions\n\n'
    + 'Go to the 📋 Ticker List sheet.\n'
    + 'Paste your tickers in column A starting at row 2.\n\n'
    + 'Format: EXCHANGE:TICKER\n'
    + 'Examples:\n'
    + '  NASDAQ:AAPL\n'
    + '  NYSE:WM\n'
    + '  TSX:SHOP   (or TSE:SHOP)\n'
    + '  CSE:AABC   (or CNSX:AABC)\n\n'
    + 'Column B (Name) is optional — leave blank or fill manually.\n'
    + 'The nightly batch reads column A only.'
  );
}


// ════════════════════════════════════════════════════════════
// SHEET BUILDERS
// ════════════════════════════════════════════════════════════

function _createConfigSheet(ss) {
  let sh = ss.getSheetByName(SHEET_CONFIG);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_CONFIG);
  sh.setTabColor('#4285F4');
  sh.clearContents();

  // Row 1 — title
  sh.getRange(1, 1, 1, 5).merge()
    .setValue('⚙️  Supertrend Live Tracker — Configuration')
    .setBackground('#1a73e8').setFontColor('#fff')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 36);

  // Row 2-3 status
  [[2,'Last nightly run:'],[3,'Last daytime sweep:']].forEach(([row, lbl]) => {
    sh.getRange(row, 1).setValue(lbl);
    sh.getRange(row, 2).setValue('—').setFontStyle('italic');
    sh.getRange(row, 1, 1, 4).setBackground('#E8F0FE');
  });

  // Row 4 note
  sh.getRange(4, 1, 1, 5).merge()
    .setValue('ℹ️  Parameters are read at runtime. Change any value — next run picks it up automatically. Never edit grey "Status" rows.')
    .setBackground('#FFF8E1').setFontStyle('italic').setFontSize(9).setFontColor('#6D4C41').setWrap(true);
  sh.setRowHeight(4, 36);

  // Helpers
  function sh2(startRow, title) {
    sh.getRange(startRow, 1, 1, 5).merge().setValue('  ' + title)
      .setBackground('#D2E3FC').setFontColor('#1a73e8').setFontWeight('bold').setFontSize(10);
    sh.getRange(startRow + 1, 1, 1, 4)
      .setValues([['Parameter', 'Value', 'Unit', 'Description']])
      .setBackground('#4285F4').setFontColor('#fff').setFontWeight('bold');
    return startRow + 2;
  }
  function pr(row, label, val, unit, desc) {
    sh.getRange(row, 1, 1, 4).setValues([[label, val, unit, desc]])
      .setBackground(row % 2 === 0 ? '#F8F9FA' : '#FFFFFF');
    return row + 1;
  }

  // ── Core Supertrend (rows 6-10) ───────────────────────────
  let row = sh2(6, '━━  Core Supertrend parameters  ━━');  // data starts row 8 → ATR_PERIOD
  row = pr(row, 'ATR Period',         10,     'days',   'Smoothing window — match your Supertrend_BT sheet');
  row = pr(row, 'ATR Multiplier',     4.0,    '×',      'Band width — higher = fewer signals');
  row = pr(row, 'History to fetch',   60,     'days',   '60 days is enough for ATR(10) + Supertrend live');

  // ── Stop loss (rows 12-15) ────────────────────────────────
  row++; // blank → 11
  row = sh2(row, '━━  Stop loss (applies to live exits)  ━━');  // data row 14 → STOP_TYPE
  row = pr(row, 'Stop type',    'Fixed',  'Fixed / Trailing / None', 'Change anytime — read on every sweep');
  row = pr(row, 'Stop loss %',  0.15,     'ratio',                   '0.15 = 15% below entry (fixed) or peak (trailing)');

  // ── Entry filters — bracket system (rows 17-21) ──────── ◀ CHANGED SECTION
  row++; // blank → 16
  row = sh2(row, '━━  Entry filters (price-bracket system)  ━━');  // data row 19 → BRACKET_FILTER
  row = pr(row, 'Enable bracket filter',  true,  'ON/OFF',
    'When ON, 10 price brackets auto-gate ATR%, vol ratio, and SMA ratio at entry');
  row = pr(row, '(Bracket thresholds)',   'See SharedUtils.gs → PRICE_BRACKETS',  '',
    'Ranges: $0.50–$1, $1–$10, $10–$25 … $300+. Stocks < $0.50 always rejected.');
  row = pr(row, 'Volume lookback',  20,  'days',
    'Lookback for vol-ratio computation (used by bracket filter)');

  // ── SMA computation (rows 23-26) ──────────────────────── ◀ CHANGED SECTION
  row++; // blank → 22
  row = sh2(row, '━━  SMA computation (used by bracket filter)  ━━');  // data row 25
  row = pr(row, '(SMA ratio)',  'close / SMA — checked per bracket',  '',
    'SMA ratio must fall within bracket min/max. SMA must be converged or stock is rejected.');
  row = pr(row, 'SMA period',  50,  'days',  'Period for SMA computation');

  // ── Position sizing (rows 28-30) ──────────────────────────
  row++; // blank → 27
  row = sh2(row, '━━  Position sizing  ━━');  // data row 30 → BASE_POS
  row = pr(row, 'Base position size', 25, '$', 'Default size per trade — overrideable in Live Tracker');

  // ── Watch list threshold (rows 32-34) ─────────────────────
  row++; // blank → 31
  row = sh2(row, '━━  Watch list threshold  ━━');  // data row 34 → WATCH_THRESH
  row = pr(row, 'Watch threshold', 70, '%', 'Flag stocks within this proximity to BUY signal');

  // ── Notifications (rows 36-39) ────────────────────────────
  row++; // blank → 35
  row = sh2(row, '━━  Notifications  ━━');  // data row 38 → EMAIL
  row = pr(row, 'Email address',    'Kev1265@gmail.com', '',        'Receives BUY alerts, SELL alerts, morning digest');
  row = pr(row, 'Error threshold',  10,                  'errors',  'Email warning if nightly errors exceed this');

  // ── Script status (rows 41-49) ────────────────────────────
  row++; // blank → 40
  row = sh2(row, '━━  Script status (written by scripts — do not edit)  ━━'); // data row 44
  ['Tickers scanned','Active trades','Watch stocks','Errors last run','Last nightly run','Last daytime sweep']
    .forEach(lbl => { pr(row, lbl, '—', '', ''); row++; });

  // ── Formats ───────────────────────────────────────────────
  sh.getRange(CFG_ROW.MULTIPLIER,    2).setNumberFormat('0.00');
  sh.getRange(CFG_ROW.STOP_PCT,      2).setNumberFormat('0.00%');
  sh.getRange(CFG_ROW.BASE_POS,      2).setNumberFormat('$#,##0.00');
  sh.getRange(CFG_ROW.WATCH_THRESH,  2).setNumberFormat('0');

  // Checkbox for bracket filter                               ◀ CHANGED
  const cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(CFG_ROW.BRACKET_FILTER, 2).setDataValidation(cbRule);

  // Stop type dropdown
  sh.getRange(CFG_ROW.STOP_TYPE, 2).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Fixed','Trailing','None'], true).build()
  );

  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 400);
  sh.setFrozenRows(4);
}

function _createLiveTrackerSheet(ss) {
  let sh = ss.getSheetByName(SHEET_TRACKER);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_TRACKER);
  sh.setTabColor('#34A853');

  // Row 1 — title
  sh.getRange(1, 1, 1, TRK.NCOLS_ACT).merge()
    .setValue('📺  Supertrend Live Tracker')
    .setBackground('#0D47A1').setFontColor('#fff')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1, 32);

  // Row 2 — summary metrics
  const metricLabels = ['Active trades', '', 'Unrealized P&L', '', 'Watch stocks', '', 'Universe', ''];
  sh.getRange(2, 1, 1, 8).setValues([metricLabels]).setBackground('#1565C0').setFontColor('#fff').setFontWeight('bold').setFontSize(10);
  sh.getRange(2, 2).setValue('—'); sh.getRange(2, 4).setValue('—');
  sh.getRange(2, 6).setValue('—'); sh.getRange(2, 8).setValue('—');

  // Row 3 — meta
  sh.getRange(3, 1, 1, TRK.NCOLS_ACT).merge()
    .setValue('Last refresh: — | Next run: — | Nightly batch: — | Stop mode: —')
    .setBackground('#E8F0FE').setFontStyle('italic').setFontSize(9).setFontColor('#1a73e8');
  sh.setRowHeight(3, 22);

  sh.setRowHeight(4, 8);

  // Row 5 — Active trades section header
  sh.getRange(5, 1, 1, TRK.NCOLS_ACT).merge()
    .setValue('  🟢  ACTIVE TRADES')
    .setBackground('#137333').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(5, 26);

  // Row 6 — column headers (active trades)
  const actHdrs = ['Ticker','Exchange','Entry date','Entry $','Current $',
                   'Actual size $','Unreal. P&L $','P&L %','Stop $','ST direction','Notes','Peak $ ▸'];
  sh.getRange(6, 1, 1, actHdrs.length).setValues([actHdrs])
    .setBackground('#34A853').setFontColor('#fff').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(6, 32);

  // Rows 7-56 — active trades data zone (pre-format)
  const actZone = sh.getRange(TRK_ACTIVE_START, 1, TRK_ACTIVE_END - TRK_ACTIVE_START + 1, TRK.NCOLS_ACT);
  actZone.setBackground('#F1F8E9');
  sh.getRange(TRK_ACTIVE_START, TRK.ENTRY_DATE, TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(TRK_ACTIVE_START, TRK.ENTRY_PX,   TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(TRK_ACTIVE_START, TRK.CURR_PX,    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(TRK_ACTIVE_START, TRK.SIZE,        TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('$#,##0.00');
  sh.getRange(TRK_ACTIVE_START, TRK.PNL_DOL,    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('$#,##0.00');
  sh.getRange(TRK_ACTIVE_START, TRK.PNL_PCT,    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('0.00%');
  sh.getRange(TRK_ACTIVE_START, TRK.STOP_PX,    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).setNumberFormat('$#,##0.000');
  [TRK.ENTRY_PX, TRK.SIZE, TRK.STOP_PX, TRK.NOTES].forEach(col => {
    sh.getRange(TRK_ACTIVE_START, col, TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1)
      .setBorder(null, null, true, null, null, null, '#1a73e8', SpreadsheetApp.BorderStyle.DASHED);
  });
  sh.hideColumns(TRK.PEAK_PX);

  // Row 57 — spacer
  sh.getRange(57, 1, 1, TRK.NCOLS_ACT).merge().setBackground('#E8F0FE');
  sh.setRowHeight(57, 8);

  // Row 58 — Watch list section header
  sh.getRange(TRK_WATCH_HDR_ROW, 1, 1, TRK.NCOLS_WCH).merge()
    .setValue('  🟡  WATCH LIST — stocks near a BUY signal (sorted by proximity)')
    .setBackground('#E65100').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(TRK_WATCH_HDR_ROW, 26);

  // Row 59 — watch column headers
  const watchHdrs = ['Ticker','Exchange','Proximity %','Current $','ST band $',
                     'Distance $','ATR','Vol ratio','SMA ok?','ST direction','Notes'];
  sh.getRange(TRK_WATCH_COL_HDR, 1, 1, watchHdrs.length).setValues([watchHdrs])
    .setBackground('#EF6C00').setFontColor('#fff').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(TRK_WATCH_COL_HDR, 32);

  // Rows 60-209 — watch data zone
  const watchZone = sh.getRange(TRK_WATCH_START, 1, TRK_WATCH_END - TRK_WATCH_START + 1, TRK.NCOLS_WCH);
  watchZone.setBackground('#FFF8E1');
  sh.getRange(TRK_WATCH_START, TRK.W_CURR,   TRK_WATCH_END - TRK_WATCH_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(TRK_WATCH_START, TRK.W_BAND,   TRK_WATCH_END - TRK_WATCH_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(TRK_WATCH_START, TRK.W_DIST,   TRK_WATCH_END - TRK_WATCH_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(TRK_WATCH_START, TRK.W_ATR,    TRK_WATCH_END - TRK_WATCH_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(TRK_WATCH_START, TRK.W_VOLRAT, TRK_WATCH_END - TRK_WATCH_START + 1, 1).setNumberFormat('0.0"×"');
  sh.getRange(TRK_WATCH_START, TRK.W_PROX,   TRK_WATCH_END - TRK_WATCH_START + 1, 1).setNumberFormat('0"%"');
  sh.getRange(TRK_WATCH_START, TRK.W_NOTES,  TRK_WATCH_END - TRK_WATCH_START + 1, 1)
    .setBorder(null, null, true, null, null, null, '#E65100', SpreadsheetApp.BorderStyle.DASHED);

  // Column widths
  const widths = [100, 70, 90, 90, 90, 90, 70, 80, 70, 90, 160, 70];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(6);
}

function _createWatchlistSheet(ss) {
  let sh = ss.getSheetByName(SHEET_WATCHLIST);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_WATCHLIST);
  sh.setTabColor('#FF6D00');

  const hdrs = ['Ticker (EXCHANGE:TICKER)', 'Yahoo Symbol', 'Exchange',
    'Last Close', 'ATR', 'ST Band', 'ST Direction', 'Proximity %',
    'Vol Ratio', 'ATR %', 'SMA Ratio',
    'BK ATR% Min', 'BK ATR% Max', 'BK Vol Min', 'BK Vol Max',
    'BK SMA Min', 'BK SMA Max',
    'Status', 'Last Updated'];
  sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
    .setBackground('#FF6D00').setFontColor('#fff').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(1, 36);
  sh.setFrozenRows(1);

  // Bracket threshold header cells — different background to distinguish
  sh.getRange(1, WL.BK_ATR_MIN, 1, 6).setBackground('#BF360C');

  sh.getRange(1, 1, 1, 1).merge();
  sh.setColumnWidths(1, hdrs.length, 90);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 100);
  sh.setColumnWidth(WL.UPDATED, 130);

  sh.getRange(2, 1, 1, hdrs.length).merge()
    .setValue('⚙️  Written exclusively by the nightly batch. Do not edit manually. '
      + 'Daytime sweep reads ACTIVE + WATCH rows only. You may hide this sheet.')
    .setBackground('#FFF3E0').setFontStyle('italic').setFontSize(9).setFontColor('#E65100').setWrap(true);
  sh.setRowHeight(2, 28);

  // Pre-format data columns
  sh.getRange(3, WL.LAST_CLOSE, 1000, 1).setNumberFormat('$#,##0.000');
  sh.getRange(3, WL.ATR,        1000, 1).setNumberFormat('$#,##0.000');
  sh.getRange(3, WL.ST_BAND,    1000, 1).setNumberFormat('$#,##0.000');
  sh.getRange(3, WL.PROXIMITY,  1000, 1).setNumberFormat('0"%"');
  sh.getRange(3, WL.VOL_RATIO,  1000, 1).setNumberFormat('0.0"×"');
  sh.getRange(3, WL.ATR_PCT,    1000, 1).setNumberFormat('0.00%');
  sh.getRange(3, WL.SMA_RATIO,  1000, 1).setNumberFormat('0.000');
  // Bracket threshold columns
  sh.getRange(3, WL.BK_ATR_MIN, 1000, 1).setNumberFormat('0.00%');
  sh.getRange(3, WL.BK_ATR_MAX, 1000, 1).setNumberFormat('0.00%');
  sh.getRange(3, WL.BK_VOL_MIN, 1000, 1).setNumberFormat('0.00"×"');
  sh.getRange(3, WL.BK_VOL_MAX, 1000, 1).setNumberFormat('0.00"×"');
  sh.getRange(3, WL.BK_SMA_MIN, 1000, 1).setNumberFormat('0.00');
  sh.getRange(3, WL.BK_SMA_MAX, 1000, 1).setNumberFormat('0.00');
  sh.getRange(3, WL.UPDATED,    1000, 1).setNumberFormat('yyyy-mm-dd HH:mm');
}

function _createTickerListSheet(ss) {
  let sh = ss.getSheetByName(SHEET_TICKERS);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_TICKERS);
  sh.setTabColor('#9C27B0');

  sh.getRange(1, 1, 1, 3).setValues([['Ticker (EXCHANGE:TICKER)', 'Name (optional)', 'Notes (optional)']])
    .setBackground('#9C27B0').setFontColor('#fff').setFontWeight('bold');
  sh.getRange(2, 1, 1, 3).merge()
    .setValue('Paste your tickers in column A starting row 2. Format: EXCHANGE:TICKER '
      + '(e.g. NASDAQ:AAPL, NYSE:WM, TSX:SHOP, CSE:AABC). Columns B and C are optional.')
    .setBackground('#F3E5F5').setFontStyle('italic').setFontSize(9).setFontColor('#6A1B9A').setWrap(true);
  sh.setRowHeight(2, 28);
  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 250);
  sh.setFrozenRows(2);

  const samples = [['NASDAQ:AAPL','Apple Inc.',''],['NYSE:WM','Waste Management',''],
    ['TSX:SHOP','Shopify Inc.',''],['CSE:AABC','','']];
  sh.getRange(3, 1, samples.length, 3).setValues(samples)
    .setFontStyle('italic').setFontColor('#999');
}

function _createJournalSheet(ss) {
  let sh = ss.getSheetByName(SHEET_JOURNAL);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SHEET_JOURNAL);
  sh.setTabColor('#00897B');

  sh.getRange(1, 1, 1, JCLOSED.NCOLS).merge()
    .setValue('📓  Trade Journal — Permanent Record')
    .setBackground('#00695C').setFontColor('#fff').setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setRowHeight(1, 30);

  sh.getRange(2, 1, 1, JCLOSED.NCOLS).merge()
    .setValue('ℹ️  Entry $ and Exit $ are editable — correct them to match your actual fills. '
      + 'Open trades move to Closed automatically when an exit is detected.')
    .setBackground('#E0F2F1').setFontStyle('italic').setFontSize(9).setFontColor('#00695C').setWrap(true);
  sh.setRowHeight(2, 28);

  sh.setRowHeight(3, 8);

  const statsStart = JRN_CLOSED_START;
  const statsData = [
    ['Total closed trades',   '=IFERROR(COUNTA(B' + statsStart + ':B10000),0)',                              '0'],
    ['Win rate',              '=IFERROR(COUNTIF(J' + statsStart + ':J10000,">"&0)/MAX(COUNTA(J' + statsStart + ':J10000),1),"—")', '0.00%'],
    ['Avg win %',             '=IFERROR(AVERAGEIF(J' + statsStart + ':J10000,">"&0,I' + statsStart + ':I10000),"—")', '0.00%'],
    ['Avg loss %',            '=IFERROR(AVERAGEIF(J' + statsStart + ':J10000,"<="&0,I' + statsStart + ':I10000),"—")', '0.00%'],
    ['Total P&L $',           '=IFERROR(SUM(J' + statsStart + ':J10000),0)',                                 '$#,##0.00'],
    ['Total P&L %',           '=IFERROR(SUM(J' + statsStart + ':J10000)/SUMIF(E' + statsStart + ':E10000,">"&0,E' + statsStart + ':E10000),"—")', '0.00%'],
    ['Best trade $',          '=IFERROR(MAX(J' + statsStart + ':J10000),"—")',                               '$#,##0.00'],
    ['Worst trade $',         '=IFERROR(MIN(J' + statsStart + ':J10000),"—")',                               '$#,##0.00'],
    ['Avg duration (days)',   '=IFERROR(AVERAGE(K' + statsStart + ':K10000),"—")',                           '0.0'],
  ];
  statsData.forEach(([lbl, formula, fmt], i) => {
    const row = 4 + i;
    const bg  = row % 2 === 0 ? '#F8F9FA' : '#FFFFFF';
    sh.getRange(row, 1).setValue(lbl).setFontColor('#555').setBackground(bg).setFontSize(12);
    sh.getRange(row, 2).setFormula(formula).setFontWeight('bold').setBackground(bg).setFontSize(13);
    if (fmt && fmt !== '0') sh.getRange(row, 2).setNumberFormat(fmt);
    sh.setRowHeight(row, 26);
  });
  sh.getRange(4, 5, 9, JCLOSED.NCOLS - 4).setBackground('#F8F9FA');

  sh.getRange(14, 1, 1, JOPEN.NCOLS).merge()
    .setValue('  🟢  Currently open trades (script-managed — do not delete rows)')
    .setBackground('#137333').setFontColor('#fff').setFontWeight('bold');
  sh.setRowHeight(14, 26);

  const openHdrs = ['Trade #','Ticker','Entry date','Entry $','Actual size $','Current $','Unreal. P&L $','Notes'];
  sh.getRange(15, 1, 1, openHdrs.length).setValues([openHdrs])
    .setBackground('#34A853').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange(JRN_OPEN_START, JOPEN.ENTRY_DATE, JRN_OPEN_END - JRN_OPEN_START + 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(JRN_OPEN_START, JOPEN.ENTRY_PX,   JRN_OPEN_END - JRN_OPEN_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(JRN_OPEN_START, JOPEN.SIZE,        JRN_OPEN_END - JRN_OPEN_START + 1, 1).setNumberFormat('$#,##0.00');
  sh.getRange(JRN_OPEN_START, JOPEN.CURR_PX,    JRN_OPEN_END - JRN_OPEN_START + 1, 1).setNumberFormat('$#,##0.000');
  sh.getRange(JRN_OPEN_START, JOPEN.PNL_DOL,    JRN_OPEN_END - JRN_OPEN_START + 1, 1).setNumberFormat('$#,##0.00');
  [JOPEN.ENTRY_PX, JOPEN.SIZE, JOPEN.NOTES].forEach(col => {
    sh.getRange(JRN_OPEN_START, col, JRN_OPEN_END - JRN_OPEN_START + 1, 1)
      .setBorder(null, null, true, null, null, null, '#137333', SpreadsheetApp.BorderStyle.DASHED);
  });

  sh.setRowHeight(67, 8); sh.setRowHeight(68, 8);

  sh.getRange(JRN_CLOSED_HDR, 1, 1, JCLOSED.NCOLS).merge()
    .setValue('  🔒  Closed trades — permanent history (never deleted by script)')
    .setBackground('#1565C0').setFontColor('#fff').setFontWeight('bold');
  sh.setRowHeight(JRN_CLOSED_HDR, 26);

  const closedHdrs = ['Trade #','Ticker','Entry date','Entry $','Actual size $',
    'Exit date','Exit $','Exit reason','P&L %','P&L $','Days','Notes'];
  sh.getRange(JRN_CLOSED_COL_H, 1, 1, closedHdrs.length).setValues([closedHdrs])
    .setBackground('#1976D2').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(JRN_CLOSED_COL_H, 28);

  [50,100,100,90,90,100,90,140,80,90,60,200].forEach((w, i) => sh.setColumnWidth(i+1, w));
  sh.setFrozenRows(3);
}

function _createBTPlaceholder(ss) {
  let sh = ss.getSheetByName(SHEET_BT);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BT);
    sh.setTabColor('#9E9E9E');
    sh.getRange(1, 1).setValue(
      '🔍 Supertrend_BT — Manual Inspection Sheet\n\n'
      + 'Copy your Supertrend_BT tab from the old backtester file into this tab,\n'
      + 'or use this sheet as a standalone single-ticker inspector.\n'
      + 'This sheet is NOT used by the automated scripts.'
    ).setWrap(true).setFontStyle('italic').setFontColor('#777');
    sh.setRowHeight(1, 120);
    sh.setColumnWidth(1, 500);
  }
}

function _orderSheets(ss) {
  const order = [SHEET_CONFIG, SHEET_TRACKER, SHEET_WATCHLIST, SHEET_TICKERS, SHEET_JOURNAL, SHEET_BT];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh) && ss.moveActiveSheet(i + 1);
  });
}
