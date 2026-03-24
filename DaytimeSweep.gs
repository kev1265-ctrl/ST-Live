// ============================================================
// FILE 4 OF 4 — DaytimeSweep.gs
// Supertrend Live Tracker v1.1
//
// v1.1 CHANGES:
//   - BUY entry filter logic replaced: old volFilter + smaFilter
//     toggles replaced by single bracketFilter check via
//     passesBracketFilter_() from SharedUtils.gs
//   - Exit logic unchanged (entry-only filter)
//   - Enhanced logging with bracket reject reasons
// ============================================================

function runDaytimeSweep() {
  if (isWeekend_()) { Logger.log('Weekend — sweep skipped.'); return; }
  if (!isMarketHours_()) { Logger.log('Outside market hours — sweep skipped.'); return; }

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig_();
  Logger.log('── Daytime sweep ' + tsNow_()
    + ' | Bracket filter: ' + (cfg.bracketFilter ? 'ON' : 'OFF') + ' ──');

  // ── 1. Read Watchlist — ACTIVE + WATCH only ───────────────
  const allWL   = readWatchlist_(ss);
  const toScan  = allWL.filter(r => r.status === 'ACTIVE' || r.status === 'WATCH');
  if (toScan.length === 0) {
    Logger.log('No ACTIVE or WATCH stocks. Sweep done.');
    _updateSweepTimestamp_(ss, cfg);
    return;
  }
  Logger.log('Scanning ' + toScan.length + ' stocks (ACTIVE + WATCH).');

  // ── 2. Fetch fresh data (parallel) ───────────────────────
  const yahooSyms  = toScan.map(r => r.yahooSym);
  const freshData  = fetchYahooBatch_(yahooSyms, cfg.historyDays);

  // ── 3. Load active trades from Live Tracker ───────────────
  const activeTrades = readActiveTrades_(ss);
  const activeMap    = {};
  activeTrades.forEach(t => { activeMap[t.ticker] = t; });

  const trackerSh = ss.getSheetByName(SHEET_TRACKER);
  const journalSh = ss.getSheetByName(SHEET_JOURNAL);
  const wlSh      = ss.getSheetByName(SHEET_WATCHLIST);

  const updatedWL = [];

  // ── 4. Process each scanned stock ────────────────────────
  for (const wlRow of toScan) {
    const ohlcv = freshData[wlRow.yahooSym];
    const snap  = ohlcv ? analyzeOHLCV_(ohlcv, cfg) : null;

    if (!snap) {
      Logger.log('  SKIP (no data): ' + wlRow.ticker);
      continue;
    }

    updatedWL.push({ wlRow, snap });

    // ── ACTIVE trade — check for exit (unchanged) ─────────
    if (wlRow.status === 'ACTIVE') {
      const trade = activeMap[wlRow.ticker];
      if (!trade) {
        Logger.log('  ⚠️ ACTIVE in WL but not found in Live Tracker: ' + wlRow.ticker);
        continue;
      }

      const currentClose = snap.lastClose;
      const newPeak      = Math.max(trade.peakPx || trade.entryPx, currentClose);
      if (newPeak > (trade.peakPx || 0)) {
        trackerSh.getRange(trade.sheetRow, TRK.PEAK_PX).setValue(newPeak);
      }

      let exitReason = '';
      let exitPrice  = currentClose;

      if (snap.stSignal === 'SELL' || snap.stDirection === 'SELL') {
        exitReason = 'Supertrend SELL';
        exitPrice  = snap.stBand;
      } else {
        const stopCheck = checkStop_(trade.entryPx, newPeak, currentClose, cfg);
        if (stopCheck.hit) {
          exitReason = stopCheck.reason;
          exitPrice  = cfg.stopType === 'Trailing'
            ? newPeak * (1 - cfg.stopPct)
            : trade.entryPx * (1 - cfg.stopPct);
        }
      }

      if (exitReason) {
        Logger.log('  EXIT: ' + wlRow.ticker + ' — ' + exitReason);
        _closeTrade_(ss, trackerSh, journalSh, trade, exitPrice, exitReason, snap, cfg, wlRow.sheetRow, wlSh);
      } else {
        trackerSh.getRange(trade.sheetRow, TRK.ST_DIR).setValue(snap.stDirection);
        Logger.log('  HOLD: ' + wlRow.ticker + ' | ST: ' + snap.stDirection + ' | Close: $' + currentClose.toFixed(3));
      }
    }

    // ── WATCH stock — check for BUY entry ─────────────────
    else if (wlRow.status === 'WATCH') {
      const hasBuySignal = snap.stSignal === 'BUY';

      // ◀ CHANGED — bracket filter replaces old vol/sma checks
      let filtered = false;
      let filterReasons = [];

      if (hasBuySignal && cfg.bracketFilter) {
        const bf = passesBracketFilter_(snap);
        filtered = !bf.pass;
        filterReasons = bf.reasons;
      }

      if (hasBuySignal && !filtered) {
        Logger.log('  BUY: ' + wlRow.ticker + ' | Entry: $' + snap.lastClose.toFixed(3));
        _openTrade_(ss, trackerSh, journalSh, wlRow.ticker, wlRow.exchange,
                    snap, cfg, wlRow.sheetRow, wlSh);
      } else if (hasBuySignal && filtered) {
        Logger.log('  BUY signal FILTERED: ' + wlRow.ticker + ' — ' + filterReasons.join('; '));
      } else {
        Logger.log('  WATCH: ' + wlRow.ticker + ' | Proximity: ' + snap.proximity + '%');
      }
    }
  }

  // ── 5. Batch-update proximity scores in Live Tracker watch zone ──
  _refreshWatchZoneProximity_(trackerSh, updatedWL.filter(u => u.wlRow.status === 'WATCH'), cfg);

  // ── 6. Update Watchlist proximity values for WATCH rows ──
  updatedWL.filter(u => u.wlRow.status === 'WATCH').forEach(({ wlRow, snap }) => {
    try {
      wlSh.getRange(wlRow.sheetRow, WL.PROXIMITY).setValue(snap.proximity);
      wlSh.getRange(wlRow.sheetRow, WL.LAST_CLOSE).setValue(snap.lastClose);
    } catch(e) { Logger.log('WL update error: ' + e.message); }
  });

  // ── 7. Update summary panel ───────────────────────────────
  _updateSweepTimestamp_(ss, cfg);
  Logger.log('── Sweep complete ' + tsNow_() + ' ──');
}


// ════════════════════════════════════════════════════════════
// OPEN TRADE
// ════════════════════════════════════════════════════════════

function _openTrade_(ss, trackerSh, journalSh, ticker, exchange, snap, cfg, wlSheetRow, wlSh) {
  const entryPx   = snap.lastClose;
  const entryDate = snap.lastDate instanceof Date ? snap.lastDate : new Date();
  const tradeNum  = nextTradeNum_();
  const size      = cfg.basePos;
  const stopPx    = cfg.stopType === 'None' ? 0 : entryPx * (1 - cfg.stopPct);

  // ── Live Tracker — active zone ─────────────────────────
  const trkRow = firstEmptyActiveRow_(trackerSh);
  if (trkRow === 0) {
    Logger.log('⚠️ Active trades zone full — cannot add ' + ticker);
  } else {
    const pnlPctFormula = '=(E' + trkRow + '-D' + trkRow + ')/D' + trkRow;
    const pnlDolFormula = '=' + pnlPctFormula + '*F' + trkRow;
    trackerSh.getRange(trkRow, TRK.TICKER).setValue(ticker);
    trackerSh.getRange(trkRow, TRK.EXCHANGE).setValue(exchange);
    trackerSh.getRange(trkRow, TRK.ENTRY_DATE).setValue(entryDate).setNumberFormat('yyyy-mm-dd');
    trackerSh.getRange(trkRow, TRK.ENTRY_PX).setValue(entryPx).setNumberFormat('$#,##0.000');
    trackerSh.getRange(trkRow, TRK.CURR_PX)
      .setFormula('=IFERROR(GOOGLEFINANCE("' + ticker + '","price"),"—")')
      .setNumberFormat('$#,##0.000');
    trackerSh.getRange(trkRow, TRK.SIZE).setValue(size).setNumberFormat('$#,##0.00');
    trackerSh.getRange(trkRow, TRK.PNL_DOL).setFormula(pnlDolFormula).setNumberFormat('$#,##0.00');
    trackerSh.getRange(trkRow, TRK.PNL_PCT).setFormula(pnlPctFormula).setNumberFormat('0.00%');
    trackerSh.getRange(trkRow, TRK.STOP_PX).setValue(stopPx > 0 ? stopPx : '—').setNumberFormat('$#,##0.000');
    trackerSh.getRange(trkRow, TRK.ST_DIR).setValue('BUY');
    trackerSh.getRange(trkRow, TRK.PEAK_PX).setValue(entryPx);
    trackerSh.getRange(trkRow, 1, 1, TRK.NCOLS_ACT).setBackground('#DCEDC8').setFontColor('#1B5E20');
  }

  // ── Journal — open trades zone ─────────────────────────
  const jRow = firstEmptyOpenRow_(journalSh);
  if (jRow > 0) {
    const jPnlFormula = '=(F' + jRow + '-D' + jRow + ')/D' + jRow + '*E' + jRow;
    journalSh.getRange(jRow, JOPEN.TRADE_NUM).setValue(tradeNum);
    journalSh.getRange(jRow, JOPEN.TICKER).setValue(ticker);
    journalSh.getRange(jRow, JOPEN.ENTRY_DATE).setValue(entryDate).setNumberFormat('yyyy-mm-dd');
    journalSh.getRange(jRow, JOPEN.ENTRY_PX).setValue(entryPx).setNumberFormat('$#,##0.000');
    journalSh.getRange(jRow, JOPEN.SIZE).setValue(size).setNumberFormat('$#,##0.00');
    journalSh.getRange(jRow, JOPEN.CURR_PX)
      .setFormula('=IFERROR(GOOGLEFINANCE("' + ticker + '","price"),"—")')
      .setNumberFormat('$#,##0.000');
    journalSh.getRange(jRow, JOPEN.PNL_DOL).setFormula(jPnlFormula).setNumberFormat('$#,##0.00');
    journalSh.getRange(jRow, 1, 1, JOPEN.NCOLS).setBackground('#E8F5E9').setFontColor('#1B5E20');
  }

  // ── Watchlist — mark as ACTIVE ─────────────────────────
  if (wlSheetRow) {
    wlSh.getRange(wlSheetRow, WL.STATUS).setValue('ACTIVE');
    wlSh.getRange(wlSheetRow, WL.PROXIMITY).setValue(0);
    wlSh.getRange(wlSheetRow, WL.UPDATED).setValue(new Date());
    wlSh.getRange(wlSheetRow, 1, 1, WL.NCOLS).setBackground('#E6F4EA');
  }

  // ── Remove from Live Tracker watch zone ────────────────
  _removeFromWatchZone_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TRACKER), ticker);

  // ── Send BUY email ─────────────────────────────────────
  const subject = '🟢 BUY Signal — ' + ticker + ' @ $' + entryPx.toFixed(3);
  const html    = buildBuyEmail_(ticker, exchange, entryPx, snap, cfg);
  sendAlert_(subject, html, cfg);

  Logger.log('  ✅ Trade opened: ' + ticker + ' | #' + tradeNum + ' | Entry: $' + entryPx.toFixed(3));
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════
// CLOSE TRADE
// ════════════════════════════════════════════════════════════

function _closeTrade_(ss, trackerSh, journalSh, trade, exitPx, exitReason, snap, cfg, wlSheetRow, wlSh) {
  const exitDate = snap.lastDate instanceof Date ? snap.lastDate : new Date();
  const pnlPct   = trade.entryPx > 0 ? (exitPx - trade.entryPx) / trade.entryPx : 0;
  const pnlDol   = pnlPct * (trade.size || cfg.basePos);
  const days     = daysBetween_(trade.entryDate, exitDate);

  const jOpenRow = findOpenJournalRow_(journalSh, trade.ticker);
  let   tradeNum = '—', size = trade.size || cfg.basePos;
  if (jOpenRow > 0) {
    const jData = journalSh.getRange(jOpenRow, 1, 1, JOPEN.NCOLS).getValues()[0];
    tradeNum = jData[JOPEN.TRADE_NUM - 1] || tradeNum;
    size     = +jData[JOPEN.SIZE - 1] || size;
    journalSh.getRange(jOpenRow, 1, 1, JOPEN.NCOLS).clearContent().setBackground('#FFFFFF').setFontColor('#000');
  }

  const pnlDolActual = pnlPct * size;

  const closedRow = [
    tradeNum,
    trade.ticker,
    trade.entryDate,
    trade.entryPx,
    size,
    exitDate,
    exitPx,
    exitReason,
    pnlPct,
    pnlDolActual,
    days,
    trade.notes || '',
  ];
  appendClosedTrade_(journalSh, closedRow);

  trackerSh.getRange(trade.sheetRow, 1, 1, TRK.NCOLS_ACT)
    .clearContent().setBackground('#FFFFFF').setFontColor('#000');

  if (wlSheetRow) {
    // ◀ CHANGED — also apply bracket filter when deciding post-exit status
    let newStatus = 'NEUTRAL';
    if (snap.stDirection === 'SELL' && snap.proximity >= cfg.watchThresh) {
      if (cfg.bracketFilter) {
        const bf = passesBracketFilter_(snap);
        newStatus = bf.pass ? 'WATCH' : 'NEUTRAL';
      } else {
        newStatus = 'WATCH';
      }
    }
    wlSh.getRange(wlSheetRow, WL.STATUS).setValue(newStatus);
    wlSh.getRange(wlSheetRow, WL.UPDATED).setValue(new Date());
    wlSh.getRange(wlSheetRow, 1, 1, WL.NCOLS).setBackground(newStatus === 'WATCH' ? '#FFF8E1' : '#FFFFFF');
  }

  const subject = (pnlDolActual >= 0 ? '🔴 EXIT (+)' : '🔴 EXIT (-)') + ' — ' + trade.ticker
    + ' | ' + (pnlPct >= 0 ? '+' : '') + (pnlPct * 100).toFixed(2) + '%';
  const html = buildSellEmail_(
    trade.ticker, trade.exchange,
    trade.entryPx, exitPx, exitReason,
    pnlPct, pnlDolActual, days
  );
  sendAlert_(subject, html, cfg);

  Logger.log('  ✅ Trade closed: ' + trade.ticker + ' | ' + exitReason
    + ' | P&L: ' + (pnlPct >= 0 ? '+' : '') + (pnlPct * 100).toFixed(2) + '%'
    + ' ($' + pnlDolActual.toFixed(2) + ')');
  SpreadsheetApp.flush();
}


// ════════════════════════════════════════════════════════════
// WATCH ZONE HELPERS
// ════════════════════════════════════════════════════════════

function _removeFromWatchZone_(trackerSh, ticker) {
  const watchRows = TRK_WATCH_END - TRK_WATCH_START + 1;
  const vals      = trackerSh.getRange(TRK_WATCH_START, TRK.W_TICKER, watchRows, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === ticker) {
      trackerSh.getRange(TRK_WATCH_START + i, 1, 1, TRK.NCOLS_WCH)
        .clearContent().setBackground('#FFF8E1');
      return;
    }
  }
}

function _refreshWatchZoneProximity_(trackerSh, updatedWatchItems, cfg) {
  if (!updatedWatchItems.length) return;
  const snapMap = {};
  updatedWatchItems.forEach(({ wlRow, snap }) => { snapMap[wlRow.ticker] = snap; });

  const watchRows = TRK_WATCH_END - TRK_WATCH_START + 1;
  const vals      = trackerSh.getRange(TRK_WATCH_START, 1, watchRows, TRK.NCOLS_WCH).getValues();

  vals.forEach((row, i) => {
    const ticker = String(row[TRK.W_TICKER - 1]).trim();
    if (!ticker || !snapMap[ticker]) return;
    const snap   = snapMap[ticker];
    const shRow  = TRK_WATCH_START + i;
    trackerSh.getRange(shRow, TRK.W_PROX).setValue(snap.proximity);
    trackerSh.getRange(shRow, TRK.W_BAND).setValue(snap.stBand);
    trackerSh.getRange(shRow, TRK.W_DIST).setValue(Math.max(0, snap.stBand - snap.lastClose));
    trackerSh.getRange(shRow, TRK.W_ATR).setValue(snap.atr);
    trackerSh.getRange(shRow, TRK.W_VOLRAT).setValue(snap.volRatio);
    trackerSh.getRange(shRow, TRK.W_SMAOK).setValue(snap.smaOK ? 'YES' : 'NO');
    const prox = snap.proximity;
    const bg   = prox >= 90 ? '#FFE0B2' : prox >= 75 ? '#FFF3E0' : '#FFF8E1';
    trackerSh.getRange(shRow, 1, 1, TRK.NCOLS_WCH).setBackground(bg);
  });
}


// ════════════════════════════════════════════════════════════
// SUMMARY PANEL UPDATER
// ════════════════════════════════════════════════════════════

function _updateSweepTimestamp_(ss, cfg) {
  const trackerSh   = ss.getSheetByName(SHEET_TRACKER);
  if (!trackerSh) return;

  const activeTrades = readActiveTrades_(ss);
  const activeCount  = activeTrades.length;

  let unrealizedPnl = 0;
  activeTrades.forEach(t => {
    const val = trackerSh.getRange(t.sheetRow, TRK.PNL_DOL).getValue();
    if (typeof val === 'number') unrealizedPnl += val;
  });

  const wlAll    = readWatchlist_(ss);
  const wCount   = wlAll.filter(r => r.status === 'WATCH').length;
  const universe = wlAll.length;

  trackerSh.getRange(2, 2).setValue(activeCount);
  const pnlStr = (unrealizedPnl >= 0 ? '+$' : '-$') + Math.abs(unrealizedPnl).toFixed(2);
  trackerSh.getRange(2, 4).setValue(pnlStr);
  trackerSh.getRange(2, 6).setValue(wCount);
  trackerSh.getRange(2, 8).setValue(universe);

  const ts = tsNow_();
  trackerSh.getRange(3, 1, 1, TRK.NCOLS_ACT).merge()
    .setValue('Last refresh: ' + ts + '  |  Stop mode: '
      + cfg.stopType + ' ' + (cfg.stopPct * 100).toFixed(0) + '%'
      + '  |  Universe: ' + universe + ' tickers'
      + (cfg.bracketFilter ? '  |  Bracket filter: ON' : ''));   // ◀ NEW

  setConfigStatus_('ACTIVE_TRADES', activeCount);
  setConfigStatus_('WATCH_STOCKS',  wCount);
  setConfigStatus_('LAST_DAYTIME',  ts);

  SpreadsheetApp.flush();
}
