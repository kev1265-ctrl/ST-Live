// ============================================================
// FILE 1 OF 4 — SharedUtils.gs
// Supertrend Live Tracker v1.1
// Shared constants, config reader, calculation engine,
// ticker converter, Yahoo Finance fetcher, email builders.
//
// v1.1 CHANGES:
//   - Added PRICE_BRACKETS constant (10 brackets)
//   - Added getBracket_() and passesBracketFilter_() functions
//   - analyzeOHLCV_() now returns lastHL2
//   - CFG_ROW: VOL_FILTER → BRACKET_FILTER; removed VOL_THRESHOLD, SMA_FILTER
//   - getConfig_(): bracketFilter replaces volFilter/volThreshold/smaFilter
// ============================================================


// ════════════════════════════════════════════════════════════
// SECTION 1 — SHEET & LAYOUT CONSTANTS
// ════════════════════════════════════════════════════════════

const SHEET_CONFIG    = '⚙️ Config';
const SHEET_TRACKER   = '📺 Live Tracker';
const SHEET_WATCHLIST = '🌙 Watchlist';
const SHEET_TICKERS   = '📋 Ticker List';
const SHEET_JOURNAL   = '📓 Journal';
const SHEET_BT        = '🔍 Supertrend_BT';

// Config sheet — row number for each parameter (value in col B)
const CFG_ROW = {
  ATR_PERIOD      : 8,
  MULTIPLIER      : 9,
  HISTORY_DAYS    : 10,
  STOP_TYPE       : 14,
  STOP_PCT        : 15,
  BRACKET_FILTER  : 19,   // ◀ CHANGED — was VOL_FILTER
  // VOL_THRESHOLD removed — bracket table replaces it
  VOL_LOOKBACK    : 21,   // kept — still used for vol-ratio computation
  // SMA_FILTER removed — bracket table replaces it
  SMA_PERIOD      : 26,   // kept — still used for SMA computation
  BASE_POS        : 30,
  WATCH_THRESH    : 34,
  EMAIL           : 38,
  ERROR_THRESH    : 39,
  // Status rows — written by scripts, never edited by hand
  TICKERS_SCANNED : 44,
  ACTIVE_TRADES   : 45,
  WATCH_STOCKS    : 46,
  ERRORS_LAST     : 47,
  LAST_NIGHTLY    : 48,
  LAST_DAYTIME    : 49,
};

// Live Tracker row layout
const TRK_ACTIVE_START  = 7;    // first active-trade data row
const TRK_ACTIVE_END    = 56;   // last  active-trade data row  (50 rows)
const TRK_WATCH_HDR_ROW = 58;   // "WATCH LIST" section label
const TRK_WATCH_COL_HDR = 59;   // watch column headers
const TRK_WATCH_START   = 60;   // first watch data row
const TRK_WATCH_END     = 209;  // last  watch data row (150 rows)

// Live Tracker columns (1-based)
const TRK = {
  TICKER    : 1,   // A  EXCHANGE:TICKER
  EXCHANGE  : 2,   // B
  ENTRY_DATE: 3,   // C
  ENTRY_PX  : 4,   // D  editable
  CURR_PX   : 5,   // E  =GOOGLEFINANCE formula
  SIZE      : 6,   // F  editable
  PNL_DOL   : 7,   // G  formula
  PNL_PCT   : 8,   // H  formula
  STOP_PX   : 9,   // I  editable display value
  ST_DIR    : 10,  // J
  NOTES     : 11,  // K  editable
  PEAK_PX   : 12,  // L  hidden — trailing-stop tracking
  NCOLS_ACT : 12,  // total active-trades columns
  // Watch list columns (reuse same A-K slots, different meaning)
  W_TICKER  : 1,   // A
  W_EXCHANGE: 2,   // B
  W_PROX    : 3,   // C  proximity %
  W_CURR    : 4,   // D  =GOOGLEFINANCE formula
  W_BAND    : 5,   // E  ST band $
  W_DIST    : 6,   // F  distance $
  W_ATR     : 7,   // G
  W_VOLRAT  : 8,   // H  vol ratio
  W_SMAOK   : 9,   // I  YES / NO
  W_STDIR   : 10,  // J
  W_NOTES   : 11,  // K  editable
  NCOLS_WCH : 11,
};

// Journal row layout
const JRN_OPEN_START   = 17;   // first open-trade row
const JRN_OPEN_END     = 66;   // last  open-trade row (50 rows)
const JRN_CLOSED_HDR   = 69;   // "CLOSED TRADES" section label
const JRN_CLOSED_COL_H = 70;   // closed-trades column headers
const JRN_CLOSED_START = 71;   // first closed-trade row

// Journal open-trades columns (1-based)
const JOPEN = {
  TRADE_NUM : 1,   // A
  TICKER    : 2,   // B
  ENTRY_DATE: 3,   // C
  ENTRY_PX  : 4,   // D  editable
  SIZE      : 5,   // E  editable
  CURR_PX   : 6,   // F  =GOOGLEFINANCE
  PNL_DOL   : 7,   // G  formula
  NOTES     : 8,   // H  editable
  NCOLS     : 8,
};

// Journal closed-trades columns (1-based)
const JCLOSED = {
  TRADE_NUM : 1,   // A
  TICKER    : 2,   // B
  ENTRY_DATE: 3,   // C
  ENTRY_PX  : 4,   // D  editable
  SIZE      : 5,   // E
  EXIT_DATE : 6,   // F
  EXIT_PX   : 7,   // G  editable
  EXIT_RSN  : 8,   // H
  PNL_PCT   : 9,   // I
  PNL_DOL   : 10,  // J
  DURATION  : 11,  // K
  NOTES     : 12,  // L  editable
  NCOLS     : 12,
};

// Watchlist columns (1-based)
const WL = {
  TICKER    : 1,   // A
  YAHOO_SYM : 2,   // B
  EXCHANGE  : 3,   // C
  LAST_CLOSE: 4,   // D
  ATR       : 5,   // E
  ST_BAND   : 6,   // F
  ST_DIR    : 7,   // G
  PROXIMITY : 8,   // H
  VOL_RATIO : 9,   // I  actual computed value
  ATR_PCT   : 10,  // J  actual ATR / HL2
  SMA_RATIO : 11,  // K  actual close / SMA
  BK_ATR_MIN: 12,  // L  bracket threshold
  BK_ATR_MAX: 13,  // M  bracket threshold
  BK_VOL_MIN: 14,  // N  bracket threshold
  BK_VOL_MAX: 15,  // O  bracket threshold
  BK_SMA_MIN: 16,  // P  bracket threshold
  BK_SMA_MAX: 17,  // Q  bracket threshold
  STATUS    : 18,  // R  ACTIVE / WATCH / NEUTRAL / ERROR
  UPDATED   : 19,  // S
  NCOLS     : 19,
};


// ════════════════════════════════════════════════════════════
// SECTION 1B — PRICE BRACKET FILTER TABLE           ◀ NEW
// ════════════════════════════════════════════════════════════

// Each bracket defines the acceptable range for ATR% (ATR/HL2),
// volume ratio, and SMA ratio (close/SMA) at entry time.
// Stocks below $0.50 are auto-rejected (no matching bracket).
const PRICE_BRACKETS = [
  { minPx:   0.50, maxPx:     1, atrPctMin: 0.06, atrPctMax: 0.14, volMin: 1.22, volMax: 17.30, smaMin: 0.88, smaMax: 1.16 },
  { minPx:   1,    maxPx:    10, atrPctMin: 0.03, atrPctMax: 0.10, volMin: 0.73, volMax:  6.89, smaMin: 0.96, smaMax: 1.16 },
  { minPx:  10,    maxPx:    25, atrPctMin: 0.02, atrPctMax: 0.06, volMin: 0.68, volMax:  3.68, smaMin: 0.99, smaMax: 1.10 },
  { minPx:  25,    maxPx:    50, atrPctMin: 0.02, atrPctMax: 0.05, volMin: 0.72, volMax:  3.30, smaMin: 1.00, smaMax: 1.09 },
  { minPx:  50,    maxPx:    75, atrPctMin: 0.02, atrPctMax: 0.05, volMin: 0.71, volMax:  3.04, smaMin: 1.00, smaMax: 1.09 },
  { minPx:  75,    maxPx:   100, atrPctMin: 0.02, atrPctMax: 0.05, volMin: 0.74, volMax:  2.96, smaMin: 1.00, smaMax: 1.09 },
  { minPx: 100,    maxPx:   150, atrPctMin: 0.02, atrPctMax: 0.05, volMin: 0.74, volMax:  2.76, smaMin: 1.00, smaMax: 1.09 },
  { minPx: 150,    maxPx:   200, atrPctMin: 0.02, atrPctMax: 0.05, volMin: 0.79, volMax:  2.60, smaMin: 0.99, smaMax: 1.09 },
  { minPx: 200,    maxPx:   300, atrPctMin: 0.02, atrPctMax: 0.04, volMin: 0.69, volMax:  2.65, smaMin: 1.00, smaMax: 1.09 },
  { minPx: 300,    maxPx: Infinity, atrPctMin: 0.02, atrPctMax: 0.05, volMin: 0.72, volMax: 2.54, smaMin: 1.00, smaMax: 1.09 },
];

// Returns the matching bracket object, or null if price < $0.50.
function getBracket_(price) {
  if (!price || price < 0.50) return null;
  return PRICE_BRACKETS.find(b => price >= b.minPx && price < b.maxPx) || null;
}

// Checks whether a snapshot passes all bracket filter criteria.
// Returns { pass: boolean, reasons: string[] }
// snap must include: lastClose, lastHL2, atr, volRatio, lastSMA
function passesBracketFilter_(snap) {
  if (!snap || isNaN(snap.lastClose) || snap.lastClose <= 0) {
    return { pass: false, reasons: ['No data'] };
  }

  const bracket = getBracket_(snap.lastClose);
  if (!bracket) {
    return { pass: false, reasons: ['Price $' + snap.lastClose.toFixed(3) + ' < $0.50'] };
  }

  const reasons = [];

  // 1) ATR% = ATR / HL2
  if (!snap.lastHL2 || snap.lastHL2 <= 0 || isNaN(snap.atr)) {
    reasons.push('ATR or HL2 unavailable');
  } else {
    const atrPct = snap.atr / snap.lastHL2;
    if (atrPct < bracket.atrPctMin || atrPct > bracket.atrPctMax) {
      reasons.push('ATR% ' + (atrPct * 100).toFixed(2) + '% outside ['
        + (bracket.atrPctMin * 100).toFixed(0) + '–' + (bracket.atrPctMax * 100).toFixed(0) + '%]');
    }
  }

  // 2) Volume ratio
  if (isNaN(snap.volRatio) || snap.volRatio <= 0) {
    reasons.push('Vol ratio unavailable');
  } else if (snap.volRatio < bracket.volMin || snap.volRatio > bracket.volMax) {
    reasons.push('Vol ' + snap.volRatio.toFixed(1) + '× outside ['
      + bracket.volMin.toFixed(2) + '–' + bracket.volMax.toFixed(2) + ']');
  }

  // 3) SMA ratio = close / SMA  (reject if SMA hasn't converged)
  if (isNaN(snap.lastSMA) || snap.lastSMA <= 0) {
    reasons.push('SMA not converged');
  } else {
    const smaRatio = snap.lastClose / snap.lastSMA;
    if (smaRatio < bracket.smaMin || smaRatio > bracket.smaMax) {
      reasons.push('SMA ratio ' + smaRatio.toFixed(3) + ' outside ['
        + bracket.smaMin.toFixed(2) + '–' + bracket.smaMax.toFixed(2) + ']');
    }
  }

  return { pass: reasons.length === 0, reasons };
}


// ════════════════════════════════════════════════════════════
// SECTION 2 — CONFIG READER
// ════════════════════════════════════════════════════════════

function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('Config sheet not found. Run firstTimeSetup() first.');
  const g = r => sh.getRange(r, 2).getValue();
  const b = r => { const v = g(r); return v === true || String(v).toLowerCase() === 'true'; };
  return {
    atrPeriod     : Math.max(1, Math.round(+g(CFG_ROW.ATR_PERIOD)    || 10)),
    multiplier    : +g(CFG_ROW.MULTIPLIER)    || 4.0,
    historyDays   : Math.max(30, Math.round(+g(CFG_ROW.HISTORY_DAYS) || 60)),
    stopType      : String(g(CFG_ROW.STOP_TYPE) || 'Fixed'),
    stopPct       : +g(CFG_ROW.STOP_PCT)      || 0.15,
    bracketFilter : b(CFG_ROW.BRACKET_FILTER),   // ◀ CHANGED — replaces volFilter + smaFilter
    volLookback   : Math.round(+g(CFG_ROW.VOL_LOOKBACK) || 20),  // kept for computation
    smaPeriod     : Math.round(+g(CFG_ROW.SMA_PERIOD) || 50),    // kept for computation
    basePos       : +g(CFG_ROW.BASE_POS)      || 25,
    watchThresh   : +g(CFG_ROW.WATCH_THRESH)  || 70,
    email         : String(g(CFG_ROW.EMAIL)   || '').trim(),
    errorThresh   : Math.round(+g(CFG_ROW.ERROR_THRESH) || 10),
  };
}

function setConfigStatus_(key, value) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
    if (sh && CFG_ROW[key]) sh.getRange(CFG_ROW[key], 2).setValue(value);
  } catch(e) { Logger.log('setConfigStatus_ error (' + key + '): ' + e.message); }
}


// ════════════════════════════════════════════════════════════
// SECTION 3 — TICKER FORMAT CONVERTER
// ════════════════════════════════════════════════════════════

// EXCHANGE:TICKER  →  Yahoo Finance symbol
function toYahoo_(raw) {
  const s = String(raw || '').trim();
  if (!s.includes(':')) return s.toUpperCase();
  const [ex, tk] = s.split(':');
  const E = ex.trim().toUpperCase();
  const T = tk.trim().toUpperCase();
  if (E === 'TSX' || E === 'TSE') return T + '.TO';
  if (E === 'CSE' || E === 'CNSX') return T + '.CN';
  return T;   // NASDAQ, NYSE, etc.
}

// Returns just the exchange label (normalised)
function exchangeOnly_(raw) {
  const s = String(raw || '').trim();
  if (!s.includes(':')) return '';
  const ex = s.split(':')[0].trim().toUpperCase();
  if (ex === 'TSE') return 'TSX';
  if (ex === 'CNSX') return 'CSE';
  return ex;
}


// ════════════════════════════════════════════════════════════
// SECTION 4 — YAHOO FINANCE BATCH FETCHER
// ════════════════════════════════════════════════════════════

function fetchYahooBatch_(yahooSymbols, days) {
  const endTs   = Math.floor(Date.now() / 1000);
  const startTs = endTs - (days + 14) * 86400;
  const results = {};
  const BATCH   = 100;

  for (let b = 0; b < yahooSymbols.length; b += BATCH) {
    const slice    = yahooSymbols.slice(b, b + BATCH);
    const requests = slice.map(sym => ({
      url              : 'https://query1.finance.yahoo.com/v8/finance/chart/'
                         + encodeURIComponent(sym)
                         + '?period1=' + startTs + '&period2=' + endTs
                         + '&interval=1d&includePrePost=false',
      muteHttpExceptions: true,
      headers          : { 'User-Agent': 'Mozilla/5.0' },
    }));

    let responses;
    try { responses = UrlFetchApp.fetchAll(requests); }
    catch(e) {
      Logger.log('fetchAll error batch ' + b + ': ' + e.message);
      slice.forEach(s => { results[s] = null; });
      continue;
    }

    for (let i = 0; i < slice.length; i++) {
      const sym  = slice[i];
      const resp = responses[i];
      try {
        if (resp.getResponseCode() !== 200) { results[sym] = null; continue; }
        const json  = JSON.parse(resp.getContentText());
        const chart = json?.chart?.result?.[0];
        if (!chart) { results[sym] = null; continue; }
        const ts  = chart.timestamp || [];
        const q   = chart.indicators?.quote?.[0] || {};
        const rows = [];
        for (let j = 0; j < ts.length; j++) {
          if (!q.close?.[j]) continue;
          rows.push({
            date  : new Date(ts[j] * 1000),
            open  : q.open?.[j]   || q.close[j],
            high  : q.high?.[j]   || q.close[j],
            low   : q.low?.[j]    || q.close[j],
            close : q.close[j],
            volume: q.volume?.[j] || 0,
          });
        }
        rows.sort((a, b) => a.date - b.date);
        results[sym] = rows.slice(-days);
      } catch(e) {
        Logger.log('Parse error ' + sym + ': ' + e.message);
        results[sym] = null;
      }
    }
    if (b + BATCH < yahooSymbols.length) Utilities.sleep(300);
  }
  return results;
}

function fetchYahoo_(yahooSymbol, days) {
  const r = fetchYahooBatch_([yahooSymbol], days);
  return r[yahooSymbol];
}


// ════════════════════════════════════════════════════════════
// SECTION 5 — SUPERTREND CALCULATION ENGINE
// ════════════════════════════════════════════════════════════

function computeATR_(ohlcv, period) {
  const n   = ohlcv.length;
  const atr = new Array(n).fill(NaN);
  if (n < period + 1) return atr;
  const tr = ohlcv.map((r, i) => {
    if (i === 0) return r.high - r.low;
    return Math.max(r.high - r.low,
                    Math.abs(r.high - ohlcv[i-1].close),
                    Math.abs(r.low  - ohlcv[i-1].close));
  });
  let sum = 0;
  for (let i = 0; i < period; i++) sum += tr[i];
  atr[period - 1] = sum / period;
  for (let i = period; i < n; i++)
    atr[i] = (atr[i-1] * (period - 1) + tr[i]) / period;
  return atr;
}

function computeSMA_(closes, period) {
  return closes.map((_, i) => {
    if (i < period - 1) return NaN;
    let s = 0;
    for (let j = i - period + 1; j <= i; j++) s += closes[j];
    return s / period;
  });
}

function computeSupertrend_(ohlcv, period, multiplier) {
  const n   = ohlcv.length;
  const atr = computeATR_(ohlcv, period);
  const out = Array.from({length: n}, () => ({ band: NaN, direction: '', signal: '' }));
  let prevUpper = NaN, prevLower = NaN, dir = 'SELL';
 
  for (let i = period - 1; i < n; i++) {
    const a = atr[i];
    if (isNaN(a)) continue;
    const hl2 = (ohlcv[i].high + ohlcv[i].low) / 2;
    const rawU = hl2 + multiplier * a;
    const rawL = hl2 - multiplier * a;
 
    // FIXED: Proper band smoothing
    // Upper band: keep previous if it's lower (bands should converge, not oscillate)
    // Lower band: keep previous if it's higher
    const upper = (!isNaN(prevUpper) && rawU > prevUpper) ? prevUpper : rawU;
    const lower = (!isNaN(prevLower) && rawL < prevLower) ? prevLower : rawL;
 
    let newDir = dir;
    if (dir === 'SELL' && ohlcv[i].close > upper) newDir = 'BUY';
    if (dir === 'BUY'  && ohlcv[i].close < lower) newDir = 'SELL';
 
    out[i].band      = newDir === 'BUY' ? lower : upper;
    out[i].direction = newDir;
    out[i].signal    = newDir !== dir ? newDir : '';
    dir = newDir;
    prevUpper = upper;
    prevLower = lower;
  }
  return out;
}
 

// Full analysis for one ticker. Returns snapshot object or null.
function analyzeOHLCV_(ohlcv, cfg) {
  if (!ohlcv || ohlcv.length < cfg.atrPeriod + 5) return null;
  const closes  = ohlcv.map(r => r.close);
  const volumes = ohlcv.map(r => r.volume);
  const n       = ohlcv.length;
  const st      = computeSupertrend_(ohlcv, cfg.atrPeriod, cfg.multiplier);
  const smaArr  = computeSMA_(closes, cfg.smaPeriod);
  const atrArr  = computeATR_(ohlcv, cfg.atrPeriod);

  // Latest valid bar
  let last = n - 1;
  while (last >= 0 && !st[last]?.direction) last--;
  if (last < 0) return null;

  const lastST    = st[last];
  const lastClose = closes[last];
  const lastATR   = atrArr[last];
  const lastSMA   = smaArr[last];
  const lastHL2   = (ohlcv[last].high + ohlcv[last].low) / 2;   // ◀ NEW

  // Volume ratio
  const lb      = Math.min(cfg.volLookback, last);
  const volSlice = volumes.slice(Math.max(0, last - lb), last);
  const volAvg  = volSlice.length > 0 ? volSlice.reduce((s, v) => s + v, 0) / volSlice.length : 0;
  const volRatio = volAvg > 0 ? volumes[last] / volAvg : 0;
  const smaOK   = !isNaN(lastSMA) && lastClose >= lastSMA;

  // Proximity (% of the way from 1 ATR away up to the ST band)
  let proximity = 0;
  if (lastST.direction === 'SELL' && !isNaN(lastATR) && lastATR > 0) {
    const dist = lastST.band - lastClose;
    proximity  = Math.max(0, Math.min(100, Math.round((1 - dist / lastATR) * 100)));
  }

  return {
    lastClose,
    lastDate    : ohlcv[last].date,
    lastHL2,                                    // ◀ NEW — used by bracket filter
    atr         : lastATR,
    stBand      : lastST.band,
    stDirection : lastST.direction,
    stSignal    : lastST.signal,   // 'BUY' | 'SELL' | ''
    proximity,
    volRatio    : Math.round(volRatio * 10) / 10,
    smaOK,
    lastSMA,
  };
}

// Checks stop condition for an open trade. Returns { hit, reason } or { hit:false }
function checkStop_(entryPx, peakPx, currentClose, cfg) {
  if (!cfg.stopType || cfg.stopType === 'None') return { hit: false };
  const level = cfg.stopType === 'Trailing'
    ? peakPx * (1 - cfg.stopPct)
    : entryPx * (1 - cfg.stopPct);
  if (currentClose <= level) {
    return { hit: true, reason: cfg.stopType + ' stop ' + (cfg.stopPct * 100).toFixed(0) + '%' };
  }
  return { hit: false };
}


// ════════════════════════════════════════════════════════════
// SECTION 6 — LIVE TRACKER SHEET HELPERS
// ════════════════════════════════════════════════════════════

function readActiveTrades_(ss) {
  const sh   = ss.getSheetByName(SHEET_TRACKER);
  if (!sh) return [];
  const rows = sh.getRange(TRK_ACTIVE_START, 1,
    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, TRK.NCOLS_ACT).getValues();
  return rows
    .map((r, i) => ({ sheetRow: TRK_ACTIVE_START + i, data: r }))
    .filter(o => o.data[TRK.TICKER - 1] && String(o.data[TRK.TICKER - 1]).trim() !== '')
    .map(o => ({
      sheetRow  : o.sheetRow,
      ticker    : String(o.data[TRK.TICKER     - 1]).trim(),
      exchange  : String(o.data[TRK.EXCHANGE   - 1]).trim(),
      entryDate : o.data[TRK.ENTRY_DATE - 1],
      entryPx   : +o.data[TRK.ENTRY_PX  - 1] || 0,
      size      : +o.data[TRK.SIZE      - 1] || 0,
      stopPx    : +o.data[TRK.STOP_PX   - 1] || 0,
      stDir     : String(o.data[TRK.ST_DIR    - 1]).trim(),
      notes     : String(o.data[TRK.NOTES     - 1]).trim(),
      peakPx    : +o.data[TRK.PEAK_PX   - 1] || 0,
    }));
}

function findActiveRow_(trackerSh, ticker) {
  const vals = trackerSh.getRange(TRK_ACTIVE_START, TRK.TICKER,
    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === ticker) return TRK_ACTIVE_START + i;
  }
  return -1;
}

function firstEmptyActiveRow_(trackerSh) {
  const vals = trackerSh.getRange(TRK_ACTIVE_START, TRK.TICKER,
    TRK_ACTIVE_END - TRK_ACTIVE_START + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][0] || String(vals[i][0]).trim() === '') return TRK_ACTIVE_START + i;
  }
  return 0;
}

function firstEmptyOpenRow_(journalSh) {
  const vals = journalSh.getRange(JRN_OPEN_START, JOPEN.TICKER,
    JRN_OPEN_END - JRN_OPEN_START + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][0] || String(vals[i][0]).trim() === '') return JRN_OPEN_START + i;
  }
  return 0;
}

function findOpenJournalRow_(journalSh, ticker) {
  const vals = journalSh.getRange(JRN_OPEN_START, JOPEN.TICKER,
    JRN_OPEN_END - JRN_OPEN_START + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === ticker) return JRN_OPEN_START + i;
  }
  return -1;
}

function appendClosedTrade_(journalSh, row) {
  const last    = journalSh.getLastRow();
  const nextRow = Math.max(last + 1, JRN_CLOSED_START);
  journalSh.getRange(nextRow, 1, 1, row.length).setValues([row]);
  journalSh.getRange(nextRow, JCLOSED.ENTRY_DATE, 1, 1).setNumberFormat('yyyy-mm-dd');
  journalSh.getRange(nextRow, JCLOSED.EXIT_DATE,  1, 1).setNumberFormat('yyyy-mm-dd');
  journalSh.getRange(nextRow, JCLOSED.ENTRY_PX,   1, 1).setNumberFormat('$#,##0.000');
  journalSh.getRange(nextRow, JCLOSED.EXIT_PX,    1, 1).setNumberFormat('$#,##0.000');
  journalSh.getRange(nextRow, JCLOSED.PNL_PCT,    1, 1).setNumberFormat('0.00%');
  journalSh.getRange(nextRow, JCLOSED.PNL_DOL,    1, 1).setNumberFormat('$#,##0.00');
  const pnl = row[JCLOSED.PNL_DOL - 1];
  const bg  = pnl > 0 ? '#E6F4EA' : '#FCE8E6';
  const fc  = pnl > 0 ? '#137333' : '#C5221F';
  journalSh.getRange(nextRow, 1, 1, JCLOSED.NCOLS).setBackground(bg).setFontColor(fc);
}

function nextTradeNum_() {
  const props = PropertiesService.getScriptProperties();
  const n     = parseInt(props.getProperty('LT_TRADE_COUNTER') || '0', 10) + 1;
  props.setProperty('LT_TRADE_COUNTER', String(n));
  return n;
}


// ════════════════════════════════════════════════════════════
// SECTION 7 — WATCHLIST SHEET READER
// ════════════════════════════════════════════════════════════

function readWatchlist_(ss) {
  const sh      = ss.getSheetByName(SHEET_WATCHLIST);
  if (!sh || sh.getLastRow() < 2) return [];
  const numRows = sh.getLastRow() - 1;
  return sh.getRange(2, 1, numRows, WL.NCOLS).getValues()
    .filter(r => r[0] && String(r[0]).trim())
    .map((r, i) => ({
      sheetRow  : 2 + i,
      ticker    : String(r[WL.TICKER     - 1]).trim(),
      yahooSym  : String(r[WL.YAHOO_SYM  - 1]).trim(),
      exchange  : String(r[WL.EXCHANGE   - 1]).trim(),
      lastClose : +r[WL.LAST_CLOSE - 1] || 0,
      atr       : +r[WL.ATR       - 1] || 0,
      stBand    : +r[WL.ST_BAND   - 1] || 0,
      stDir     : String(r[WL.ST_DIR     - 1]).trim(),
      proximity : +r[WL.PROXIMITY  - 1] || 0,
      volRatio  : +r[WL.VOL_RATIO  - 1] || 0,
      atrPct    : +r[WL.ATR_PCT    - 1] || 0,
      smaRatio  : +r[WL.SMA_RATIO  - 1] || 0,
      smaOK     : (+r[WL.SMA_RATIO - 1] || 0) >= 1.0,  // derived for backward compat
      status    : String(r[WL.STATUS    - 1]).trim(),
      updated   : r[WL.UPDATED    - 1],
    }));
}


// ════════════════════════════════════════════════════════════
// SECTION 8 — EMAIL BUILDERS
// ════════════════════════════════════════════════════════════

function sendAlert_(subject, html, cfg) {
  const em = (cfg || getConfig_()).email;
  if (!em) return;
  try { MailApp.sendEmail({ to: em, subject, htmlBody: html }); }
  catch(e) { Logger.log('Email error: ' + e.message); }
}

function _wrap(content) {
  return '<div style="font-family:Arial,sans-serif;max-width:640px;margin:auto">' + content + '</div>';
}
function _hdr(color, title, sub) {
  return '<div style="background:' + color + ';color:#fff;padding:18px 20px;border-radius:8px 8px 0 0">'
    + '<h2 style="margin:0;font-size:19px">' + title + '</h2>'
    + '<p style="margin:5px 0 0;opacity:.8;font-size:12px">' + sub + '</p></div>';
}
function _body(inner) {
  return '<div style="background:#f8f9fa;border:1px solid #dee2e6;border-top:none;'
    + 'padding:20px;border-radius:0 0 8px 8px">' + inner
    + '<p style="margin:18px 0 0;font-size:11px;color:#999">Supertrend Live Tracker — automated alert</p></div>';
}
function _row(label, val) {
  return '<tr><td style="padding:5px 0;color:#555;font-size:13px">' + label
    + '</td><td style="padding:5px 0;font-weight:bold;font-size:13px">' + val + '</td></tr>';
}

function buildBuyEmail_(ticker, exchange, entryPx, snap, cfg) {
  const ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const stop = cfg.stopType === 'None' ? 'None'
    : '$' + (entryPx * (1 - cfg.stopPct)).toFixed(3) + ' (' + cfg.stopType + ' ' + (cfg.stopPct*100).toFixed(0) + '%)';

  // ◀ NEW — bracket context in BUY email
  const bracket   = getBracket_(entryPx);
  const atrPct    = snap.lastHL2 > 0 ? snap.atr / snap.lastHL2 : 0;
  const smaRatio  = snap.lastSMA > 0 ? snap.lastClose / snap.lastSMA : 0;
  const bracketLbl = bracket
    ? '$' + bracket.minPx + '–' + (bracket.maxPx === Infinity ? '∞' : '$' + bracket.maxPx)
    : 'N/A';

  return _wrap(
    _hdr('#137333', 'BUY Signal — ' + ticker, ts + ' EST | ' + exchange) +
    _body('<table style="width:100%;border-collapse:collapse">'
      + _row('Entry price',  '$' + entryPx.toFixed(3))
      + _row('ST band',      '$' + snap.stBand.toFixed(3))
      + _row('ATR',          '$' + snap.atr.toFixed(3))
      + _row('ATR %',        (atrPct * 100).toFixed(2) + '% (HL2)')       // ◀ NEW
      + _row('Vol ratio',    snap.volRatio.toFixed(1) + '×')
      + _row('SMA ratio',    smaRatio > 0 ? smaRatio.toFixed(3) : '—')    // ◀ NEW
      + _row('Price bracket', bracketLbl)                                   // ◀ NEW
      + _row('Stop level',   stop)
      + '</table>')
  );
}

function buildSellEmail_(ticker, exchange, entryPx, exitPx, exitReason, pnlPct, pnlDol, days) {
  const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const pos = pnlDol >= 0;
  const col = pos ? '#137333' : '#a50e0e';
  const pnlStr = (pos ? '+' : '') + '$' + pnlDol.toFixed(2) + '  ('
               + (pos ? '+' : '') + (pnlPct * 100).toFixed(2) + '%)';
  return _wrap(
    _hdr(col, 'EXIT Signal — ' + ticker, ts + ' EST | ' + exchange) +
    _body('<table style="width:100%;border-collapse:collapse">'
      + _row('Exit reason',  exitReason)
      + _row('Entry price',  '$' + entryPx.toFixed(3))
      + _row('Exit price',   '$' + exitPx.toFixed(3))
      + _row('P&L',          '<span style="color:' + col + ';font-weight:bold">' + pnlStr + '</span>')
      + _row('Duration',     days + ' days')
      + '</table>')
  );
}

function buildMorningDigest_(watchRows, activeTickers, errorCount, cfg) {
  const ts  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const top = watchRows.slice(0, 15);
  const errNote = errorCount >= cfg.errorThresh
    ? '<p style="color:#a50e0e;font-size:13px;margin:0 0 12px">⚠️ '
      + errorCount + ' tickers returned errors — check Watchlist sheet (ERROR rows).</p>'
    : (errorCount > 0 ? '<p style="color:#888;font-size:12px;margin:0 0 12px">' + errorCount + ' minor errors.</p>' : '');
  const rows = top.map(r =>
    '<tr style="background:' + (top.indexOf(r) % 2 === 0 ? '#fff' : '#f8f9fa') + '">'
    + '<td style="padding:5px 8px;font-weight:bold">' + r.ticker + '</td>'
    + '<td style="padding:5px 8px">' + r.exchange + '</td>'
    + '<td style="padding:5px 8px">' + r.proximity + '%</td>'
    + '<td style="padding:5px 8px">$' + r.lastClose.toFixed(3) + '</td>'
    + '<td style="padding:5px 8px">$' + r.stBand.toFixed(3) + '</td>'
    + '<td style="padding:5px 8px">' + r.volRatio.toFixed(1) + '×</td>'
    + '<td style="padding:5px 8px">' + (r.smaOK ? '✅' : '—') + '</td>'
    + '</tr>'
  ).join('');
  return _wrap(
    _hdr('#1a1a2e',
      'Morning Watchlist Digest',
      ts + ' EST  |  Active: ' + activeTickers.length + '  |  Watch: ' + watchRows.length
      + (cfg.bracketFilter ? '  |  Bracket filter: ON' : '')) +                     // ◀ NEW
    _body(errNote
      + '<h3 style="color:#1a73e8;font-size:14px;margin:0 0 10px">Top ' + top.length
      + ' watch stocks (sorted by proximity)</h3>'
      + '<table style="width:100%;border-collapse:collapse;font-size:12px">'
      + '<thead><tr style="background:#4285F4;color:#fff">'
      + '<th style="padding:6px 8px;text-align:left">Ticker</th>'
      + '<th style="padding:6px 8px;text-align:left">Exch.</th>'
      + '<th style="padding:6px 8px;text-align:left">Proximity</th>'
      + '<th style="padding:6px 8px;text-align:left">Last close</th>'
      + '<th style="padding:6px 8px;text-align:left">ST band</th>'
      + '<th style="padding:6px 8px;text-align:left">Vol ratio</th>'
      + '<th style="padding:6px 8px;text-align:left">SMA ok?</th>'
      + '</tr></thead><tbody>' + rows + '</tbody></table>')
  );
}


// ════════════════════════════════════════════════════════════
// SECTION 9 — DATE & SCHEDULE UTILITIES
// ════════════════════════════════════════════════════════════

function fmtDate_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function daysBetween_(d1, d2) {
  const a = d1 instanceof Date ? d1 : new Date(d1);
  const b = d2 instanceof Date ? d2 : new Date(d2);
  return Math.max(0, Math.round(Math.abs(b - a) / 86400000));
}

function isWeekend_() {
  return [0, 6].includes(new Date().getDay());
}

function isMarketHours_() {
  const now  = new Date();
  const tz   = 'America/New_York';
  const h    = parseInt(Utilities.formatDate(now, tz, 'H'), 10);
  const m    = parseInt(Utilities.formatDate(now, tz, 'm'), 10);
  const mins = h * 60 + m;
  const day  = now.getDay();
  return day >= 1 && day <= 5 && mins >= 570 && mins <= 990; // 9:30–16:30
}

function tsNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}
