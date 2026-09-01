/************************************************************
 * STOCK UPDATER — SINGLE SCS FLOW
 *
 * ENTRY POINT:
 * - updateStockPricesSequentially()
 *
 * SUMMARY:
 * - Single execution flow
 * - Updates indices first
 * - Fetches each symbol once from SCS Trade
 * - Writes price + change together
 * - No separate price/change entry points
 * - No price batching
 * - No recovery worker
 * - No DPS dependency
 *
 * SHEET:
 * - Status:        Column T
 * - Symbol:        Column U
 * - Price:         Column V
 * - Change:        Column W
 *
 * INDICES:
 * - KSE100 %:      E3
 * - KMI30 %:       E4
 * - KSE100 value:  I3
 * - KMI30 value:   I4
 ************************************************************/

/********************
 * CONFIG
 ********************/
const CONFIG = {
  timezone: "Asia/Karachi",

  marketHours: {
    monThu: {
      open: { hour: 9, minute: 30 },
      close: { hour: 15, minute: 45 }
    },
    fri: {
      open: { hour: 9, minute: 0 },
      close: { hour: 16, minute: 45 },
      breakStart: { hour: 12, minute: 15 },
      breakEnd: { hour: 14, minute: 30 }
    }
  },

  afterClose: {
    firstRunAfterCloseMins: 30,
    toleranceMins: 5
  },

  urls: {
    scsIndices: "https://www.scstrade.com/Default.aspx",
    scsTrade: "https://www.scstrade.com/stockscreening/SS_CompanySnapShot.aspx?symbol="
  },

  requestOptions: {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    },
    timeout: 30000
  },

  sheetName: "MAIN",

  logCells: {
    status: "B1",

    // Index % change cells
    kse100: "E3",
    kmi30: "E4",

    // Index current value cells
    kse100Current: "I3",
    kmi30Current: "I4",

    afterCloseState: "AA1"
  },

  columns: {
    statusIndex: 20, // T
    symbolIndex: 21, // U
    priceIndex: 22,  // V
    changeIndex: 23  // W
  },

  rows: {
    alhisf: { start: 2, end: 20 }
  },

  afterRowDelayMs: 500,

  maxRetries: 5,
  maxRuntimeMs: 5 * 60 * 1000,
  runtimeSafetyMs: 5000,

  throttle: {
    minDelayMs: 500,
    maxDelayMs: 12000,
    upFactor: 2.2,
    downFactor: 0.9,
    successStreakToCool: 5,
    basePenaltyMs: 600,
    jitterMs: 400
  },

  locks: {
    waitMs: 1000
  }
};

/********************
 * RUNTIME STATE
 ********************/
const THROTTLE_STATE = {
  scs: { delayMs: 0, okStreak: 0 },
  indices: { delayMs: 500, okStreak: 0 }
};

/********************
 * SINGLE ENTRY POINT
 ********************/
function updateStockPricesSequentially() {
  withScriptLock_("STOCK UPDATE", function () {
    const ctx = createRunContext_();

    logRunStart_(ctx, "STOCK UPDATE");

    const gate = isRunAllowedNow_Stateful_(ctx.sheet, ctx.now);

    if (!gate.isAllowed) {
      Logger.log("[%s] ⛔ STOCK UPDATE blocked: %s", ctx.ts, gate.reason);
      return;
    }

    Logger.log("[%s] ✅ STOCK UPDATE allowed: %s", ctx.ts, gate.reason);

    Logger.log("[%s] 🌐 Updating indices first", ctx.ts);
    updateIndicesWithRetry_(ctx.sheet, CONFIG.maxRetries);

    updateRowsLineByLine_(ctx);

    ctx.sheet.getRange(CONFIG.logCells.status).setValue(fmt_(new Date(), "yyyy-MM-dd HH:mm:ss"));

    logRunFinish_(ctx, "STOCK UPDATE");
  });
}

/********************
 * CORE CONTEXT / LOGGING
 ********************/
function createRunContext_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
  const now = new Date();

  return {
    sheet,
    now,
    ts: fmt_(now, "yyyy-MM-dd HH:mm:ss"),
    startMs: Date.now()
  };
}

function logRunStart_(ctx, label) {
  Logger.log("[%s] ▶ START %s", ctx.ts, label);
}

function logRunFinish_(ctx, label) {
  Logger.log("[%s] ✅ FINISH %s", fmt_(new Date(), "yyyy-MM-dd HH:mm:ss"), label);
}

function isRuntimeNearlyExceeded_(ctx) {
  return Date.now() - ctx.startMs > CONFIG.maxRuntimeMs - CONFIG.runtimeSafetyMs;
}

/********************
 * LOCKING
 ********************/
function withScriptLock_(label, fn) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(CONFIG.locks.waitMs)) {
    Logger.log("⛔ %s skipped: another execution is already running.", label);
    return;
  }

  try {
    fn();
  } catch (e) {
    Logger.log("❌ %s failed: %s", label, e.message || e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/********************
 * ROW UPDATE
 ********************/
function updateRowsLineByLine_(ctx) {
  const startRow = CONFIG.rows.alhisf.start;
  const endRow = getLastSymbolRow_(ctx.sheet);

  Logger.log("[%s] 📋 stock scan | rows=%s-%s", ctx.ts, startRow, endRow);

  for (let row = startRow; row <= endRow; row++) {
    if (isRuntimeNearlyExceeded_(ctx)) {
      Logger.log("[%s] ⏱ runtime guard hit in stock scan", ctx.ts);
      return;
    }

    processSingleRow_(ctx.sheet, row);
  }

  Logger.log("[%s] ✅ stock scan completed", ctx.ts);
}

function processSingleRow_(sheet, row) {
  const symbol = safeTrim_(sheet.getRange(row, CONFIG.columns.symbolIndex).getValue());

  if (!symbol) {
    Logger.log("   ↷ row=%s skipped | blank symbol", row);
    return { row, symbol: "", outcome: "blank" };
  }

  const statusCell = sheet.getRange(row, CONFIG.columns.statusIndex);
  const priceCell = sheet.getRange(row, CONFIG.columns.priceIndex);
  const changeCell = sheet.getRange(row, CONFIG.columns.changeIndex);

  const prevStatusText = safeTrim_(statusCell.getValue());
  const lastKnownTs = getLastKnownUpdateTimestampText_(prevStatusText);
  const now = new Date();

  const skipDecision = shouldSkipRowByStatus_(prevStatusText, now);

  if (skipDecision.skip) {
    if (skipDecision.writeSkip) {
      statusCell.setValue("Skip @ " + fmt_(now, "HH:mm:ss"));
      SpreadsheetApp.flush();
    }

    Logger.log("   ↷ row=%s | symbol=%s | skipped | reason=%s", row, symbol, skipDecision.reason);
    return { row, symbol, outcome: "skipped", reason: skipDecision.reason };
  }

  statusCell.setValue("Updating @ " + fmt_(new Date(), "HH:mm:ss"));
  SpreadsheetApp.flush();

  Logger.log("   → row=%s | symbol=%s | fetching SCS snapshot", row, symbol);

  try {
    const snapshot = fetchSCSTradeSnapshotWithRetry_(symbol, CONFIG.maxRetries);

    const priceOk = snapshot && isFiniteNumber_(snapshot.price);
    const changeOk = snapshot && isFiniteNumber_(snapshot.changePct);

    if (priceOk) {
      priceCell.setValue(snapshot.price);
    }

    if (changeOk) {
      changeCell.setValue(snapshot.changePct / 100);
    }

    if (priceOk && changeOk) {
      statusCell.setValue("Done @ " + fmt_(new Date(), "HH:mm:ss"));
      SpreadsheetApp.flush();

      Logger.log(
        "   ✓ row=%s | symbol=%s | price=%s | change=%s%%",
        row,
        symbol,
        snapshot.price,
        snapshot.changePct
      );

      sleepAfterRow_();
      return {
        row,
        symbol,
        outcome: "updated",
        price: snapshot.price,
        changePct: snapshot.changePct
      };
    }

    if (priceOk || changeOk) {
      statusCell.setValue("Partial (kept old) @ " + fmt_(new Date(), "HH:mm:ss"));
      SpreadsheetApp.flush();

      Logger.log(
        "   ⚠️ row=%s | symbol=%s | partial update | price=%s | changePct=%s",
        row,
        symbol,
        priceOk ? snapshot.price : "N/A",
        changeOk ? snapshot.changePct : "N/A"
      );

      sleepAfterRow_();
      return {
        row,
        symbol,
        outcome: "partial",
        price: priceOk ? snapshot.price : "N/A",
        changePct: changeOk ? snapshot.changePct : "N/A"
      };
    }

    statusCell.setValue(buildKeptOldStatus_(lastKnownTs));
    SpreadsheetApp.flush();

    Logger.log("   ⚠️ row=%s | symbol=%s | invalid snapshot | kept old", row, symbol);

    sleepAfterRow_();
    return { row, symbol, outcome: "kept_old" };
  } catch (e) {
    statusCell.setValue(buildKeptOldStatus_(lastKnownTs));
    SpreadsheetApp.flush();

    Logger.log("   ❌ row=%s | symbol=%s | failed=%s | kept old", row, symbol, e.message || e);

    sleepAfterRow_();
    return { row, symbol, outcome: "error", error: e.message || String(e) };
  }
}

function shouldSkipRowByStatus_(statusText, now) {
  const prevText = safeTrim_(statusText);

  if (/^Updating\b/i.test(prevText)) {
    const prevDt = parseStatusCellTime_(prevText, now);

    if (prevDt && now.getTime() - prevDt.getTime() >= 0 && now.getTime() - prevDt.getTime() < 10 * 1000) {
      return { skip: true, writeSkip: false, reason: "Updating < 10s" };
    }
  }

  if (/^Skip\s*@/i.test(prevText)) {
    const prevDt = parseStatusCellTime_(prevText, now);

    if (prevDt && now.getTime() - prevDt.getTime() >= 0 && now.getTime() - prevDt.getTime() < 10 * 1000) {
      return { skip: true, writeSkip: false, reason: "Skip < 10s" };
    }
  }

  if (/^Done\s*@/i.test(prevText)) {
    const prevDt = parseStatusCellTime_(prevText, now);

    if (prevDt && now.getTime() - prevDt.getTime() >= 0 && now.getTime() - prevDt.getTime() < 10 * 1000) {
      return { skip: true, writeSkip: true, reason: "Done < 10s" };
    }
  }

  return { skip: false, writeSkip: false, reason: "" };
}

/********************
 * COMPANY SNAPSHOT — SCS TRADE
 ********************/
function fetchSCSTradeSnapshotWithRetry_(symbol, maxRetries) {
  let last = "Error";

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    Logger.log("      ↻ snapshot retry | symbol=%s | attempt=%s/%s", symbol, attempt, maxRetries);

    try {
      const snapshot = getSCSTradeSnapshot_(symbol);

      if (snapshot && (isFiniteNumber_(snapshot.price) || isFiniteNumber_(snapshot.changePct))) {
        Logger.log(
          "      ✅ snapshot success | symbol=%s | price=%s | change=%s | changePct=%s",
          symbol,
          snapshot.price,
          snapshot.change,
          snapshot.changePct
        );

        return snapshot;
      }

      last = "Invalid snapshot";

      Logger.log(
        "      ⚠️ invalid snapshot | symbol=%s | price=%s | change=%s | changePct=%s",
        symbol,
        snapshot ? snapshot.price : "N/A",
        snapshot ? snapshot.change : "N/A",
        snapshot ? snapshot.changePct : "N/A"
      );
    } catch (e) {
      last = e.message || "Error";
      Logger.log("      ❌ snapshot exception | symbol=%s | message=%s", symbol, last);
    }

    if (attempt < maxRetries) {
      const waitMs = getRetryWaitMs_(attempt, "stock");

      if (waitMs > 0) {
        Logger.log("      ⏳ snapshot retry wait | symbol=%s | ms=%s", symbol, waitMs);
        Utilities.sleep(waitMs);
      }
    }
  }

  Logger.log("      ⛔ snapshot retries exhausted | symbol=%s | last=%s", symbol, last);

  return {
    price: "N/A",
    change: "N/A",
    changePct: "N/A"
  };
}

function getSCSTradeSnapshot_(symbol) {
  const url = CONFIG.urls.scsTrade + encodeURIComponent(symbol);
  const html = fetchHtml_(url);

  return parseSCSTradeSnapshot_(html, symbol);
}

function parseSCSTradeSnapshot_(html, symbol) {
  const text = htmlToPlainText_(html);

  /*
   * Expected SCS company snapshot:
   *
   * Rs. 327.90 4.39 (1.36%)
   *
   * Groups:
   * 1 = current price
   * 2 = absolute change
   * 3 = change percentage
   */

  const match = text.match(
    /Rs\.?\s*([+-]?\d[\d,]*(?:\.\d+)?)\s+([+-]?\d[\d,]*(?:\.\d+)?)\s*\(([+-]?\d+(?:\.\d+)?)\s*%\)/i
  );

  if (match && match[1] && match[3]) {
    const result = {
      price: parseNumber_(match[1]),
      change: parseNumber_(match[2]),
      changePct: parseNumber_(match[3])
    };

    Logger.log(
      "      📊 SCS snapshot parsed | symbol=%s | price=%s | change=%s | changePct=%s",
      symbol || "",
      result.price,
      result.change,
      result.changePct
    );

    return result;
  }

  /*
   * Fallback:
   * If HTML layout spacing changes, still try:
   * - first Rs. value as price
   * - first (...) percentage as changePct
   */
  const priceMatch = text.match(/Rs\.?\s*([+-]?\d[\d,]*(?:\.\d+)?)/i);
  const pctMatch = text.match(/\(([+-]?\d+(?:\.\d+)?)\s*%\)/i);

  const fallback = {
    price: priceMatch && priceMatch[1] ? parseNumber_(priceMatch[1]) : "N/A",
    change: "N/A",
    changePct: pctMatch && pctMatch[1] ? parseNumber_(pctMatch[1]) : "N/A"
  };

  Logger.log(
    "      ⚠️ SCS snapshot fallback parse | symbol=%s | price=%s | changePct=%s",
    symbol || "",
    fallback.price,
    fallback.changePct
  );

  return fallback;
}

/********************
 * INDICES — SCS TRADE
 ********************/
function updateIndicesWithRetry_(sheet, maxRetries) {
  const result = fetchIndicesWithRetry_(maxRetries);

  if (!result || result.error) {
    Logger.log("❌ indices failed | last=%s", result ? result.error : "Unknown");
    return;
  }

  const { kse, kmi, kseCurrent, kmiCurrent } = result;

  if (isFiniteNumber_(kse)) {
    sheet.getRange(CONFIG.logCells.kse100)
      .setValue(kse)
      .setNumberFormat("0.00%");
  }

  if (isFiniteNumber_(kmi)) {
    sheet.getRange(CONFIG.logCells.kmi30)
      .setValue(kmi)
      .setNumberFormat("0.00%");
  }

  if (isFiniteNumber_(kseCurrent)) {
    sheet.getRange(CONFIG.logCells.kse100Current)
      .setValue(kseCurrent)
      .setNumberFormat("#,##0");
  }

  if (isFiniteNumber_(kmiCurrent)) {
    sheet.getRange(CONFIG.logCells.kmi30Current)
      .setValue(kmiCurrent)
      .setNumberFormat("#,##0");
  }

  Logger.log(
    "✅ indices updated | KSE100 pct=%s current=%s | KMI30 pct=%s current=%s",
    kse,
    kseCurrent,
    kmi,
    kmiCurrent
  );
}

function fetchIndicesWithRetry_(maxRetries) {
  let lastError = "Error";

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    Logger.log("   ↻ indices retry | attempt=%s/%s", attempt, maxRetries);

    try {
      const html = fetchHtml_(CONFIG.urls.scsIndices);

      const kseRow = parseSCSIndexRowFromHtml_(html, "KSE 100");
      const kmiRow = parseSCSIndexRowFromHtml_(html, "KMI 30");

      const kse = kseRow ? kseRow.pct : null;
      const kmi = kmiRow ? kmiRow.pct : null;
      const kseCurrent = kseRow ? kseRow.current : null;
      const kmiCurrent = kmiRow ? kmiRow.current : null;

      if (
        isFiniteNumber_(kse) &&
        isFiniteNumber_(kmi) &&
        isFiniteNumber_(kseCurrent) &&
        isFiniteNumber_(kmiCurrent)
      ) {
        return { kse, kmi, kseCurrent, kmiCurrent };
      }

      lastError =
        "Invalid parse " +
        "(KSE100 pct=" + kse +
        ", KSE100 current=" + kseCurrent +
        ", KMI30 pct=" + kmi +
        ", KMI30 current=" + kmiCurrent + ")";

      Logger.log("   ⚠️ indices invalid parse | %s", lastError);
    } catch (e) {
      lastError = e.message || "Error";
      Logger.log("   ❌ indices exception | %s", lastError);
    }

    if (attempt < maxRetries) {
      const waitMs = getRetryWaitMs_(attempt, "indices");

      if (waitMs > 0) {
        Logger.log("   ⏳ indices retry wait | ms=%s", waitMs);
        Utilities.sleep(waitMs);
      }
    }
  }

  return { error: lastError };
}

function parseSCSIndexRowFromHtml_(html, indexName) {
  const text = htmlToPlainText_(html);

  /*
   * Expected SCS index format:
   *
   * KSE 100
   * 170,469.42 2,625.18(1.56 %)
   * Vol: 109,967,132
   *
   * KMI 30
   * 244,778.82 3,737.66(1.55 %)
   * Vol: 43,216,264
   *
   * Groups:
   * 1 = current index value
   * 2 = points change
   * 3 = percentage change
   */

  const labelPattern = indexName.replace(/\s+/g, "\\s+");

  const regex = new RegExp(
    labelPattern +
      "\\s+" +
      "([+-]?\\d[\\d,]*(?:\\.\\d+)?)" +
      "\\s+" +
      "([+-]?\\d[\\d,]*(?:\\.\\d+)?)" +
      "\\s*" +
      "\\(([+-]?\\d+(?:\\.\\d+)?)\\s*%\\)",
    "i"
  );

  const match = text.match(regex);

  if (!match || !match[1] || !match[3]) {
    Logger.log("   ⚠️ index parse failed | index=%s", indexName);
    return null;
  }

  const row = {
    current: parseNumber_(match[1]),
    change: parseNumber_(match[2]),
    pct: parseNumber_(match[3]) / 100
  };

  Logger.log(
    "   📊 SCS index parsed | index=%s | current=%s | change=%s | pct=%s",
    indexName,
    row.current,
    row.change,
    row.pct
  );

  return row;
}

/********************
 * HTTP + ADAPTIVE THROTTLE
 ********************/
function fetchHtml_(url) {
  const sourceKey = getSourceKeyFromUrl_(url);

  sleepAdaptiveBeforeRequest_(sourceKey);

  const resp = UrlFetchApp.fetch(url, CONFIG.requestOptions);
  const code = resp.getResponseCode();
  const text = resp.getContentText();

  if (code !== 200) {
    if (isThrottleHttpCode_(code)) {
      markThrottle_(sourceKey, "HTTP " + code);
    }

    throw new Error("HTTP " + code);
  }

  if (looksThrottledHtml_(text)) {
    markThrottle_(sourceKey, "Throttle-like HTML");
    throw new Error("Throttle-like HTML");
  }

  markSuccess_(sourceKey);

  return text;
}

function getSourceKeyFromUrl_(url) {
  if (url.indexOf("scstrade.com/Default.aspx") !== -1) return "indices";
  if (url.indexOf("scstrade.com") !== -1) return "scs";
  return "scs";
}

function sleepAdaptiveBeforeRequest_(sourceKey) {
  const st = THROTTLE_STATE[sourceKey] || THROTTLE_STATE.scs;
  const delayMs = Math.max(0, st.delayMs | 0);

  if (delayMs > 0) {
    Logger.log("      ⏳ adaptive delay | source=%s | ms=%s", sourceKey, delayMs);
    Utilities.sleep(delayMs);
  }
}

function markThrottle_(sourceKey, reason) {
  const cfg = CONFIG.throttle;
  const st = THROTTLE_STATE[sourceKey] || THROTTLE_STATE.scs;

  st.okStreak = 0;

  const jitter = Math.floor(Math.random() * (cfg.jitterMs + 1));

  const nextDelay = Math.min(
    cfg.maxDelayMs,
    Math.max(
      cfg.minDelayMs,
      Math.floor(st.delayMs * cfg.upFactor) + cfg.basePenaltyMs + jitter
    )
  );

  Logger.log(
    "      🟥 throttle | source=%s | reason=%s | delay=%s→%s",
    sourceKey,
    reason,
    st.delayMs,
    nextDelay
  );

  st.delayMs = nextDelay;
}

function markSuccess_(sourceKey) {
  const cfg = CONFIG.throttle;
  const st = THROTTLE_STATE[sourceKey] || THROTTLE_STATE.scs;

  st.okStreak++;

  if (st.okStreak >= cfg.successStreakToCool && st.delayMs > cfg.minDelayMs) {
    const nextDelay = Math.max(cfg.minDelayMs, Math.floor(st.delayMs * cfg.downFactor));

    Logger.log(
      "      🟩 cool-down | source=%s | streak=%s | delay=%s→%s",
      sourceKey,
      st.okStreak,
      st.delayMs,
      nextDelay
    );

    st.delayMs = nextDelay;
    st.okStreak = 0;
  }
}

function isThrottleHttpCode_(code) {
  return code === 429 || code === 500 || code === 502 || code === 503 || code === 504;
}

function looksThrottledHtml_(html) {
  if (!html) return true;

  const h = String(html).toLowerCase();

  return (
    h.indexOf("too many requests") !== -1 ||
    h.indexOf("rate limit") !== -1 ||
    h.indexOf("temporarily unavailable") !== -1 ||
    h.indexOf("service unavailable") !== -1 ||
    h.indexOf("access denied") !== -1 ||
    h.indexOf("request blocked") !== -1 ||
    h.indexOf("cloudflare") !== -1
  );
}

/********************
 * RETRY WAIT
 ********************/
function getRetryWaitMs_(attempt, type) {
  if (type === "indices") {
    if (attempt === 1) return 800;
    if (attempt === 2) return 2000;
    if (attempt === 3) return 4000;
    return 6000;
  }

  if (attempt === 1) return 1000;
  if (attempt === 2) return 3000;
  if (attempt === 3) return 6000;
  return 10000;
}

/********************
 * MARKET GATING
 ********************/
function isMarketOpenNow_(now) {
  const day = now.getDay();

  if (day === 0 || day === 6) {
    return { isOpen: false, reason: "Weekend" };
  }

  const minsNow = now.getHours() * 60 + now.getMinutes();
  const cfg = day === 5 ? CONFIG.marketHours.fri : CONFIG.marketHours.monThu;

  if (day === 5 && cfg.breakStart && cfg.breakEnd) {
    const bStart = cfg.breakStart.hour * 60 + cfg.breakStart.minute;
    const bEnd = cfg.breakEnd.hour * 60 + cfg.breakEnd.minute;

    if (minsNow >= bStart && minsNow <= bEnd) {
      return { isOpen: false, reason: "Fri break (12:15–14:30)" };
    }
  }

  const openMins = cfg.open.hour * 60 + cfg.open.minute;
  const closeMins = cfg.close.hour * 60 + cfg.close.minute;

  if (minsNow < openMins) {
    return { isOpen: false, reason: day === 5 ? "Fri before open" : "Mon-Thu before open" };
  }

  if (minsNow > closeMins) {
    return { isOpen: false, reason: day === 5 ? "Fri after close" : "Mon-Thu after close" };
  }

  return { isOpen: true, reason: day === 5 ? "Fri window" : "Mon-Thu window" };
}

function isRunAllowedNow_Stateful_(sheet, now) {
  const tsNow = fmt_(now, "yyyy-MM-dd HH:mm:ss");
  const day = now.getDay();
  const stateCell = CONFIG.logCells.afterCloseState;

  const openGate = isMarketOpenNow_(now);

  if (openGate.isOpen) {
    resetAfterCloseState_(sheet, now, stateCell);
    return { isAllowed: true, reason: "Market open window (state reset)" };
  }

  if (day === 0 || day === 6) {
    return { isAllowed: false, reason: "Weekend" };
  }

  const ctx = getAfterCloseContext_(now);

  if (!ctx) {
    return { isAllowed: false, reason: "No after-close context" };
  }

  const { closeKey, closeDt, nextOpenDt } = ctx;

  if (!(now > closeDt && now < nextOpenDt)) {
    return { isAllowed: false, reason: "Outside after-close window" };
  }

  const st = readAfterCloseState_(sheet, stateCell);

  if (!st.closeKey || st.closeKey !== closeKey) {
    st.closeKey = closeKey;
    st.step = 0;
    st.lastRun = "";
    writeAfterCloseState_(sheet, stateCell, st);
  }

  const caught = catchUpAfterCloseStep_(closeDt, nextOpenDt, now, st);

  if (caught.step !== st.step) {
    st.step = caught.step;
    writeAfterCloseState_(sheet, stateCell, st);
  }

  const nextRunDt = caught.nextRunDt;

  if (nextRunDt >= nextOpenDt) {
    return { isAllowed: false, reason: "After-close schedule finished" };
  }

  const tol = getToleranceMins_();
  const diffMins = Math.abs(now.getTime() - nextRunDt.getTime()) / 60000;

  if (diffMins > tol) {
    return {
      isAllowed: false,
      reason: "Not scheduled yet (next=" + fmt_(nextRunDt, "yyyy-MM-dd HH:mm:ss") + ")"
    };
  }

  if (st.lastRun) {
    const lastRunDt = parseLocalTs_(st.lastRun);

    if (lastRunDt && Math.abs(now.getTime() - lastRunDt.getTime()) / 60000 <= tol) {
      return { isAllowed: false, reason: "Already ran in this schedule window" };
    }
  }

  st.lastRun = tsNow;
  st.step = Math.max(0, Number(st.step || 0)) + 1;
  writeAfterCloseState_(sheet, stateCell, st);

  return {
    isAllowed: true,
    reason: "After-close run (scheduled=" + fmt_(nextRunDt, "yyyy-MM-dd HH:mm:ss") + ", stepUsed=" + caught.step + ")"
  };
}

function getToleranceMins_() {
  const n = Number(CONFIG.afterClose.toleranceMins || 5);
  return isFinite(n) && n > 0 ? n : 5;
}

function catchUpAfterCloseStep_(closeDt, nextOpenDt, now, st) {
  const firstMins = CONFIG.afterClose.firstRunAfterCloseMins;
  const tol = getToleranceMins_();

  function nextRunForStep_(s) {
    const extraHours = (s * (s + 1)) / 2;
    return new Date(closeDt.getTime() + firstMins * 60000 + extraHours * 3600000);
  }

  let step = Math.max(0, Number(st.step || 0));
  let nextRunDt = nextRunForStep_(step);

  let safety = 0;

  while (((now.getTime() - nextRunDt.getTime()) / 60000) > tol && safety < 60) {
    step++;
    nextRunDt = nextRunForStep_(step);

    if (nextRunDt >= nextOpenDt) {
      break;
    }

    safety++;
  }

  return { step, nextRunDt };
}

function getAfterCloseContext_(now) {
  const mostRecentClose = getMostRecentTradingClose_(now);

  if (!mostRecentClose) {
    return null;
  }

  const nextOpen = getNextTradingOpenAfter_(mostRecentClose.closeDt);

  if (!nextOpen) {
    return null;
  }

  return {
    closeKey: mostRecentClose.closeKey,
    closeDt: mostRecentClose.closeDt,
    nextOpenDt: nextOpen
  };
}

function getMostRecentTradingClose_(now) {
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const sessToday = getSessionForDate_(today);

  if (!sessToday) {
    return null;
  }

  const nowMins = now.getHours() * 60 + now.getMinutes();
  let closeDate = today;

  if (nowMins <= sessToday.closeMins) {
    closeDate = getPrevTradingDayDate_(today);
  }

  const sess = getSessionForDate_(closeDate);

  if (!sess) {
    return null;
  }

  const closeDt = new Date(
    closeDate.getFullYear(),
    closeDate.getMonth(),
    closeDate.getDate(),
    Math.floor(sess.closeMins / 60),
    sess.closeMins % 60,
    0
  );

  return {
    closeKey: fmt_(closeDt, "yyyy-MM-dd"),
    closeDt
  };
}

function getNextTradingOpenAfter_(closeDt) {
  const closeDate = new Date(closeDt.getFullYear(), closeDt.getMonth(), closeDt.getDate());
  const nextTrade = getNextTradingDayDate_(closeDate);
  const sess = getSessionForDate_(nextTrade);

  if (!sess) {
    return null;
  }

  return new Date(
    nextTrade.getFullYear(),
    nextTrade.getMonth(),
    nextTrade.getDate(),
    Math.floor(sess.openMins / 60),
    sess.openMins % 60,
    0
  );
}

function getSessionForDate_(d) {
  const day = d.getDay();

  if (day === 0 || day === 6) {
    return null;
  }

  const cfg = day === 5 ? CONFIG.marketHours.fri : CONFIG.marketHours.monThu;

  return {
    openMins: cfg.open.hour * 60 + cfg.open.minute,
    closeMins: cfg.close.hour * 60 + cfg.close.minute
  };
}

function getNextTradingDayDate_(d) {
  const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());

  while (true) {
    dt.setDate(dt.getDate() + 1);

    const day = dt.getDay();

    if (day !== 0 && day !== 6) {
      return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    }
  }
}

function getPrevTradingDayDate_(d) {
  const dt = new Date(d.getFullYear(), d.getMonth(), d.getDate());

  while (true) {
    dt.setDate(dt.getDate() - 1);

    const day = dt.getDay();

    if (day !== 0 && day !== 6) {
      return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    }
  }
}

/********************
 * STATE HELPERS
 ********************/
function readAfterCloseState_(sheet, stateCell) {
  try {
    const raw = safeTrim_(sheet.getRange(stateCell).getValue());

    if (!raw) {
      return { closeKey: "", step: 0, lastRun: "" };
    }

    const obj = JSON.parse(raw);

    return {
      closeKey: obj.closeKey || "",
      step: Number(obj.step || 0),
      lastRun: obj.lastRun || ""
    };
  } catch (e) {
    return { closeKey: "", step: 0, lastRun: "" };
  }
}

function writeAfterCloseState_(sheet, stateCell, st) {
  sheet.getRange(stateCell).setValue(JSON.stringify(st));
}

function resetAfterCloseState_(sheet, now, stateCell) {
  writeAfterCloseState_(sheet, stateCell, {
    closeKey: "",
    step: 0,
    lastRun: fmt_(now, "yyyy-MM-dd HH:mm:ss")
  });
}

/********************
 * GENERAL HELPERS
 ********************/
function fmt_(date, pattern) {
  return Utilities.formatDate(date, CONFIG.timezone, pattern);
}

function safeTrim_(v) {
  return v === null || v === undefined ? "" : String(v).trim();
}

function isFiniteNumber_(v) {
  return typeof v === "number" && isFinite(v);
}

function parseNumber_(value) {
  if (value === null || value === undefined) {
    return null;
  }

  const cleaned = String(value).replace(/,/g, "").trim();
  const num = parseFloat(cleaned);

  return isFinite(num) ? num : null;
}

function getLastKnownUpdateTimestampText_(statusText) {
  const s = safeTrim_(statusText);
  const m = s.match(/@\s*(\d{2}:\d{2}:\d{2})\s*$/);

  return m ? m[1] : "";
}

function parseStatusCellTime_(statusText, now) {
  const s = safeTrim_(statusText);
  const m = s.match(/@\s*(\d{2}):(\d{2}):(\d{2})\s*$/);

  if (!m) {
    return null;
  }

  return new Date(now.getFullYear(), now.getMonth(), now.getDate(), +m[1], +m[2], +m[3]);
}

function parseLocalTs_(ts) {
  const m = String(ts || "").match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);

  if (!m) {
    return null;
  }

  return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
}

function buildKeptOldStatus_(lastKnownTs) {
  return lastKnownTs ? "Error (kept old) @ " + lastKnownTs : "Error (kept old)";
}

function getLastSymbolRow_(sheet) {
  const startRow = CONFIG.rows.alhisf.start;
  const configuredEndRow = CONFIG.rows.alhisf.end;
  const numRows = configuredEndRow - startRow + 1;

  const values = sheet.getRange(startRow, CONFIG.columns.symbolIndex, numRows, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (safeTrim_(values[i][0])) {
      return startRow + i;
    }
  }

  return startRow;
}

function sleepAfterRow_() {
  const ms = Number(CONFIG.afterRowDelayMs || 0);

  if (ms > 0) {
    Logger.log("   ⏳ row sleep | ms=%s", ms);
    Utilities.sleep(ms);
  }
}

function htmlToPlainText_(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&#37;/g, "%")
    .replace(/\s+/g, " ")
    .trim();
}
