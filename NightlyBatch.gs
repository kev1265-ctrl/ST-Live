// ============================================================
// FILE 3 OF 4 — NightlyBatch.gs
// Supertrend Live Tracker v1.2
//
// Self-chaining chunked execution to avoid the 6-minute
// Apps Script timeout.  Each chunk processes up to 500 tickers,
// stores accumulated results in PropertiesService, then sets a
// time-based trigger for the next chunk.  The final chunk writes
// the Watchlist, updates the Live Tracker, and sends emails.
//
// Pipeline state key:  LT_NIGHTLY_STATE  (JSON in ScriptProperties)
//
// Public entry points:
//   runNightlyBatch()     — called by the scheduled trigger or menu
//   _nightlyChunkResume() — called by self-chaining trigger
//   resetNightlyPipeline()— manual abort / cleanup
// ============================================================

const NIGHTLY_CHUNK_SIZE = 500;   // tickers per chunk
const NIGHTLY_CHUNK_DELAY = 15;    // seconds between chunks
const NIGHTLY_STATE_KEY  = 'LT_NIGHTLY_STATE';


// ════════════════════════════════════════════════════════════
// PUBLIC: Start nightly batch (trigger or menu)
// ════════════════════════════════════════════════════════════

function runNightlyBatch() {
  if (isWeekend_()) {
    Logger.log('Weekend — nightly batch skipped.');
    return;
  }

  // Guard: if a previous pipeline is still running, warn and bail
  const existing = _getNightlyState_();
  if (existing && !existing.done) {
    Logger.log('⚠️ Previous nightly pipeline still in progress (chunk '
      + existing.chunkIndex + '). Use resetNightlyPipeline() to abort.');
    return;
  }

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig_();

  Logger.log('═══════════════════════════════════════');
  Logger.log('  NIGHTLY BATCH START — ' + tsNow_());
  Logger.log('  Bracket filter: ' + (cfg.bracketFilter ? 'ON' : 'OFF'));
  Logger.log('═══════════════════════════════════════');

  // ── 1. Read ticker list ───────────────────────────────────
  const tickerSh = ss.getSheetByName(SHEET_TICKERS);
  if (!tickerSh) { Logger.log('❌ Ticker List sheet not found.'); return; }
  const rawTickers = tickerSh.getRange(3, 1, Math.max(tickerSh.getLastRow() - 2, 1), 1)
    .getValues().map(r => String(r[0]).trim()).filter(t => t && t.includes(':'));

  if (rawTickers.length === 0) {
    Logger.log('No tickers found in Ticker List.');
    return;
  }

  // ── 2. Read active trades (needed for status classification) ──
  const activeTrades  = readActiveTrades_(ss);
  const activeTickers = activeTrades.map(t => t.ticker);

  const totalChunks = Math.ceil(rawTickers.length / NIGHTLY_CHUNK_SIZE);

  Logger.log('Tickers: ' + rawTickers.length + ' → ' + totalChunks + ' chunks of ' + NIGHTLY_CHUNK_SIZE);
  setConfigStatus_('TICKERS_SCANNED', rawTickers.length);

  // ── 3. Initialise pipeline state ──────────────────────────
  const state = {
    startedAt     : tsNow_(),
    rawTickers    : rawTickers,
    activeTickers : activeTickers,
    totalChunks   : totalChunks,
    chunkIndex    : 0,
    // Accumulated results (serialisable snapshot summaries)
    snapshots     : [],
    errorCount    : 0,
    bracketRejects: 0,
    done          : false,
  };

  _setNightlyState_(state);

  // ── 4. Process first chunk immediately ────────────────────
  _processNightlyChunk_();
}


// ════════════════════════════════════════════════════════════
// SELF-CHAIN RESUME (called by trigger)
// ════════════════════════════════════════════════════════════

function _nightlyChunkResume() {
  // Clean up the one-shot trigger that called us
  _deleteResumeTriggers_();
  _processNightlyChunk_();
}


// ════════════════════════════════════════════════════════════
// CORE CHUNK PROCESSOR
// ════════════════════════════════════════════════════════════

function _processNightlyChunk_() {
  const state = _getNightlyState_();
  if (!state || state.done) {
    Logger.log('No active nightly pipeline. Nothing to do.');
    return;
  }

  const cfg           = getConfig_();
  const chunkIdx      = state.chunkIndex;
  const startIdx      = chunkIdx * NIGHTLY_CHUNK_SIZE;
  const chunkTickers  = state.rawTickers.slice(startIdx, startIdx + NIGHTLY_CHUNK_SIZE);
  const activeTkSet   = new Set(state.activeTickers);

  Logger.log('── Chunk ' + (chunkIdx + 1) + '/' + state.totalChunks
    + ' | Tickers ' + (startIdx + 1) + '–' + (startIdx + chunkTickers.length)
    + ' of ' + state.rawTickers.length + ' ──');

  // ── 1. Convert & fetch ────────────────────────────────────
  const yahooMap = {};
  chunkTickers.forEach(t => { yahooMap[toYahoo_(t)] = t; });
  const yahooSymbols = Object.keys(yahooMap);

  Logger.log('Fetching OHLCV for ' + yahooSymbols.length + ' symbols...');
  const ohlcvData = fetchYahooBatch_(yahooSymbols, cfg.historyDays);
  Logger.log('Fetch complete.');

  // ── 2. Analyse each ticker ────────────────────────────────
  let chunkErrors   = 0;
  let chunkRejects  = 0;
  const chunkSnaps  = [];

  yahooSymbols.forEach(sym => {
    const ticker   = yahooMap[sym];
    const exchange = exchangeOnly_(ticker);
    const ohlcv    = ohlcvData[sym];
    const snap     = ohlcv ? analyzeOHLCV_(ohlcv, cfg) : null;

    if (!snap) {
      chunkErrors++;
      chunkSnaps.push({
        ticker, sym, exchange,
        lastClose: 0, lastHL2: 0, atr: 0, stBand: 0,
        stDirection: '', proximity: 0, volRatio: 0, smaOK: false,
        lastSMA: 0,
        status: 'ERROR',
      });
      return;
    }

    let status;
    if (activeTkSet.has(ticker)) {
      status = 'ACTIVE';
    } else if (snap.stDirection === 'SELL' && snap.proximity >= cfg.watchThresh) {
      if (cfg.bracketFilter) {
        const bf = passesBracketFilter_(snap);
        if (bf.pass) {
          status = 'WATCH';
        } else {
          status = 'NEUTRAL';
          chunkRejects++;
        }
      } else {
        status = 'WATCH';
      }
    } else {
      status = 'NEUTRAL';
    }

    chunkSnaps.push({ ticker, sym, exchange, status, ...snap });
  });

  // ── 3. Accumulate into state ──────────────────────────────
  // Trim to essential fields to stay within PropertiesService limits
  const trimmed = chunkSnaps.map(s => ({
    ticker      : s.ticker,
    sym         : s.sym,
    exchange    : s.exchange,
    lastClose   : s.lastClose   || 0,
    lastHL2     : s.lastHL2     || 0,
    atr         : s.atr         || 0,
    stBand      : s.stBand      || 0,
    stDirection : s.stDirection || '',
    proximity   : s.proximity   || 0,
    volRatio    : s.volRatio    || 0,
    smaOK       : !!s.smaOK,
    lastSMA     : s.lastSMA     || 0,
    status      : s.status,
  }));

  state.snapshots      = state.snapshots.concat(trimmed);
  state.errorCount    += chunkErrors;
  state.bracketRejects+= chunkRejects;
  state.chunkIndex     = chunkIdx + 1;

  Logger.log('Chunk ' + (chunkIdx + 1) + ' done: +' + chunkSnaps.length + ' snapshots'
    + ' | errors: ' + chunkErrors
    + ' | bracket rejects: ' + chunkRejects
    + ' | accumulated: ' + state.snapshots.length);

  // ── 4. Send progress email ────────────────────────────────
  if (cfg.email) {
    const pct   = Math.round(state.chunkIndex / state.totalChunks * 100);
    const subj  = '🔄 Nightly batch — chunk ' + state.chunkIndex + '/' + state.totalChunks
                + ' (' + pct + '%)';
    const html  = _wrap(
      _hdr('#1565C0', 'Nightly Batch Progress', tsNow_() + ' EST') +
      _body('<table style="width:100%;border-collapse:collapse;font-size:13px">'
        + _row('Chunk',         state.chunkIndex + ' of ' + state.totalChunks)
        + _row('Tickers so far', state.snapshots.length + ' / ' + state.rawTickers.length)
        + _row('Errors (total)', String(state.errorCount))
        + _row('Bracket rejects', String(state.bracketRejects))
        + _row('Progress',       pct + '%')
        + '</table>')
    );
    sendAlert_(subj, html, cfg);
  }

  // ── 5. More chunks? → chain.  Last chunk? → finalise. ────
  if (state.chunkIndex < state.totalChunks) {
    _setNightlyState_(state);
    Utilities.sleep(NIGHTLY_CHUNK_DELAY * 1000);
    _scheduleResumeTrigger_();
    Logger.log('Next chunk scheduled. Exiting this execution.');
  } else {
    state.done = true;
    _setNightlyState_(state);
    Logger.log('All chunks fetched. Running finalisation...');
    _finaliseNightlyBatch_(state);
  }
}


// ════════════════════════════════════════════════════════════
// FINALISATION — write sheets, send digest, cleanup
// ════════════════════════════════════════════════════════════

function _finaliseNightlyBatch_(state) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig_();

  const snapshots = state.snapshots;

  Logger.log('═══════════════════════════════════════');
  Logger.log('  NIGHTLY FINALISATION — ' + tsNow_());
  Logger.log('  Total snapshots: ' + snapshots.length
    + ' | Errors: ' + state.errorCount
    + ' | Bracket rejects: ' + state.bracketRejects);
  Logger.log('═══════════════════════════════════════');

  setConfigStatus_('ERRORS_LAST', state.errorCount);

  // ── 1. Sort: ACTIVE first, then WATCH by proximity desc, NEUTRAL, ERROR
  snapshots.sort((a, b) => {
    const order = { ACTIVE: 0, WATCH: 1, NEUTRAL: 2, ERROR: 3 };
    const od = (order[a.status] || 2) - (order[b.status] || 2);
    if (od !== 0) return od;
    return b.proximity - a.proximity;
  });

  // ── 2. Write to Watchlist sheet ───────────────────────────
  const wlSh  = ss.getSheetByName(SHEET_WATCHLIST);
  const now   = new Date();
  const wlRows = snapshots.map(s => {
    const atrPct   = (s.lastHL2 && s.lastHL2 > 0 && s.atr) ? s.atr / s.lastHL2 : 0;
    const smaRatio = (s.lastSMA && s.lastSMA > 0 && s.lastClose) ? s.lastClose / s.lastSMA : 0;
    const bracket  = getBracket_(s.lastClose);
    return [
      s.ticker,
      s.sym,
      s.exchange,
      s.lastClose   || 0,
      s.atr         || 0,
      s.stBand      || 0,
      s.stDirection || '',
      s.proximity   || 0,
      s.volRatio    || 0,
      atrPct,
      smaRatio,
      bracket ? bracket.atrPctMin : '',
      bracket ? bracket.atrPctMax : '',
      bracket ? bracket.volMin    : '',
      bracket ? bracket.volMax    : '',
      bracket ? bracket.smaMin    : '',
      bracket ? bracket.smaMax    : '',
      s.status,
      now,
    ];
  });

  if (wlRows.length > 0) {
    const clearEnd = Math.max(wlSh.getLastRow(), 3);
    if (clearEnd > 2) wlSh.getRange(3, 1, clearEnd - 2, WL.NCOLS).clearContent();
    wlSh.getRange(3, 1, wlRows.length, WL.NCOLS).setValues(wlRows);
    const statusColors = { ACTIVE: '#E6F4EA', WATCH: '#FFF8E1', NEUTRAL: '#FFFFFF', ERROR: '#FCE8E6' };
    wlRows.forEach((row, i) => {
      const bg = statusColors[row[WL.STATUS - 1]] || '#FFFFFF';
      wlSh.getRange(3 + i, 1, 1, WL.NCOLS).setBackground(bg);
    });
    Logger.log('Watchlist written: ' + wlRows.length + ' rows.');
  }

  // ── 3. Update Live Tracker watch section ──────────────────
  const activeTrades = readActiveTrades_(ss);
  const watchStocks  = snapshots.filter(s => s.status === 'WATCH');
  _updateLiveTrackerWatchSection_(ss, watchStocks, cfg);
  setConfigStatus_('WATCH_STOCKS',  watchStocks.length);
  setConfigStatus_('ACTIVE_TRADES', activeTrades.length);

  // ── 4. Update Config status & timestamps ─────────────────
  setConfigStatus_('LAST_NIGHTLY', tsNow_());
  _updateTrackerMetaBar_(ss, cfg, activeTrades.length, watchStocks.length);

  // ── 5. Send morning digest email ─────────────────────────
  if (cfg.email) {
    const html = buildMorningDigest_(watchStocks, activeTrades, state.errorCount, cfg);
    sendAlert_('Supertrend Morning Digest — ' + fmtDate_(now), html, cfg);
    Logger.log('Morning digest email sent.');
  }

  // ── 6. Error threshold notification ─────────────────────
  if (state.errorCount >= cfg.errorThresh && cfg.email) {
    sendAlert_(
      '[Supertrend] ⚠️ Nightly batch — ' + state.errorCount + ' errors',
      _wrap(_hdr('#a50e0e', 'Nightly Batch Error Alert', tsNow_())
        + _body('<p style="font-size:14px">The nightly batch returned <strong>'
          + state.errorCount + '</strong> ticker errors (threshold: ' + cfg.errorThresh + ').</p>'
          + '<p style="font-size:13px;color:#555">Check the 🌙 Watchlist sheet and filter Status column'
          + ' = ERROR to see which tickers failed.</p>')),
      cfg
    );
  }

  // ── 7. Completion email ───────────────────────────────────
  if (cfg.email) {
    const watchCnt = watchStocks.length;
    const subj = '✅ Nightly batch complete — ' + snapshots.length + ' tickers'
               + ' | ' + watchCnt + ' WATCH | ' + state.errorCount + ' errors';
    const html = _wrap(
      _hdr('#137333', 'Nightly Batch Complete', tsNow_() + ' EST') +
      _body('<table style="width:100%;border-collapse:collapse;font-size:13px">'
        + _row('Started',          state.startedAt)
        + _row('Finished',         tsNow_())
        + _row('Total tickers',    String(snapshots.length))
        + _row('Chunks',           state.totalChunks + ' × ' + NIGHTLY_CHUNK_SIZE)
        + _row('Active trades',    String(activeTrades.length))
        + _row('Watch stocks',     String(watchCnt))
        + _row('Errors',           String(state.errorCount))
        + _row('Bracket rejects',  String(state.bracketRejects))
        + _row('Bracket filter',   cfg.bracketFilter ? 'ON' : 'OFF')
        + '</table>')
    );
    sendAlert_(subj, html, cfg);
  }

  // ── 8. Cleanup state ──────────────────────────────────────
  _clearNightlyState_();
  _deleteResumeTriggers_();

  SpreadsheetApp.flush();
  Logger.log('✅ Nightly batch complete — ' + tsNow_());
}


// ════════════════════════════════════════════════════════════
// PIPELINE STATE MANAGEMENT (PropertiesService)
// ════════════════════════════════════════════════════════════

function _getNightlyState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(NIGHTLY_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); }
  catch(e) { Logger.log('State parse error: ' + e.message); return null; }
}

function _setNightlyState_(state) {
  const json = JSON.stringify(state);
  const sizeKB = Math.round(json.length / 1024);
  Logger.log('State size: ' + sizeKB + ' KB');
  // PropertiesService limit is 500 KB per property
  if (json.length > 490000) {
    Logger.log('⚠️ State approaching size limit (' + sizeKB + ' KB). '
      + 'Trimming NEUTRAL snapshots to save space...');
    state.snapshots = state.snapshots.filter(s => s.status !== 'NEUTRAL');
    const trimmed = JSON.stringify(state);
    Logger.log('Trimmed state size: ' + Math.round(trimmed.length / 1024) + ' KB');
    PropertiesService.getScriptProperties().setProperty(NIGHTLY_STATE_KEY, trimmed);
    return;
  }
  PropertiesService.getScriptProperties().setProperty(NIGHTLY_STATE_KEY, json);
}

function _clearNightlyState_() {
  PropertiesService.getScriptProperties().deleteProperty(NIGHTLY_STATE_KEY);
}


// ════════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ════════════════════════════════════════════════════════════

function _scheduleResumeTrigger_() {
  ScriptApp.newTrigger('_nightlyChunkResume')
    .timeBased()
    .after(1)   // 1 ms — fires as soon as possible (usually ~30-60s)
    .create();
}

function _deleteResumeTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '_nightlyChunkResume')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// Manual abort — clears state and removes chain triggers
function resetNightlyPipeline() {
  _clearNightlyState_();
  _deleteResumeTriggers_();
  Logger.log('🛑 Nightly pipeline state cleared and resume triggers removed.');
  try {
    SpreadsheetApp.getUi().alert('🛑 Nightly pipeline has been reset.\n\n'
      + 'All accumulated progress has been cleared.\n'
      + 'You can start a fresh run from the menu.');
  } catch(e) { /* no UI in trigger context */ }
}


// ════════════════════════════════════════════════════════════
// LIVE TRACKER WATCH SECTION UPDATER
// ════════════════════════════════════════════════════════════

function _updateLiveTrackerWatchSection_(ss, watchStocks, cfg) {
  const sh = ss.getSheetByName(SHEET_TRACKER);
  if (!sh) return;

  const watchRows  = TRK_WATCH_END - TRK_WATCH_START + 1;

  const existingNotes = {};
  const existingVals  = sh.getRange(TRK_WATCH_START, 1, watchRows, TRK.NCOLS_WCH).getValues();
  existingVals.forEach(r => {
    const tk = String(r[TRK.W_TICKER - 1]).trim();
    if (tk) existingNotes[tk] = String(r[TRK.W_NOTES - 1]).trim();
  });

  sh.getRange(TRK_WATCH_START, 1, watchRows, TRK.NCOLS_WCH).clearContent().setBackground('#FFF8E1');

  if (watchStocks.length === 0) return;

  const rows = watchStocks.slice(0, watchRows).map(s => [
    s.ticker,
    s.exchange,
    s.proximity,
    '',
    s.stBand  || 0,
    Math.max(0, (s.stBand || 0) - (s.lastClose || 0)),
    s.atr     || 0,
    s.volRatio || 0,
    s.smaOK ? 'YES' : 'NO',
    s.stDirection || '',
    existingNotes[s.ticker] || '',
  ]);

  sh.getRange(TRK_WATCH_START, 1, rows.length, TRK.NCOLS_WCH).setValues(rows);

  for (let i = 0; i < rows.length; i++) {
    const r   = TRK_WATCH_START + i;
    const tkr = rows[i][0];
    sh.getRange(r, TRK.W_CURR).setFormula('=IFERROR(GOOGLEFINANCE("' + tkr + '","price"),"—")');
    const prox = rows[i][2];
    const bg   = prox >= 90 ? '#FFE0B2' : prox >= 75 ? '#FFF3E0' : '#FFF8E1';
    sh.getRange(r, 1, 1, TRK.NCOLS_WCH).setBackground(bg);
  }

  Logger.log('Watch section updated: ' + rows.length + ' stocks.');
}


// ════════════════════════════════════════════════════════════
// TRACKER META BAR UPDATER (rows 2-3 summary panel)
// ════════════════════════════════════════════════════════════

function _updateTrackerMetaBar_(ss, cfg, activeCount, watchCount) {
  const sh = ss.getSheetByName(SHEET_TRACKER);
  if (!sh) return;
  sh.getRange(2, 2).setValue(activeCount);
  sh.getRange(2, 6).setValue(watchCount);
  const ts = tsNow_();
  sh.getRange(3, 1, 1, TRK.NCOLS_ACT).merge()
    .setValue('Last refresh: ' + ts + '  |  Nightly batch: ' + ts
      + '  |  Stop mode: ' + cfg.stopType + ' ' + (cfg.stopPct * 100).toFixed(0) + '%'
      + (cfg.bracketFilter ? '  |  Bracket filter: ON' : ''));
  SpreadsheetApp.flush();
}
