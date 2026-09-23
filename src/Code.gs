/**
 * HOLDINGS TRACKER — Google Apps Script v8 (Batch I/O & Bug Fixes)
 *
 * Tracks open positions and unrealized gain/loss across multiple brokerage
 * accounts, on monthly snapshot tabs. Paste broker exports or screenshots;
 * the Anthropic API parses them into structured positions.
 *
 * SETUP: see README.md in this repo. In short:
 *   1. Paste this file into Extensions → Apps Script as Code.gs.
 *   2. Add TradeJournal.gs as a second script file (optional but recommended).
 *   3. Reload the sheet, then Holdings → Settings → Set Anthropic API Key.
 *
 * THINGS YOU MAY WANT TO EDIT IN CONFIG BELOW:
 *   • DEFAULT_SECTIONS — the category rows (and their colors) used to group
 *     positions. Rename freely; keep 'Liquid' and 'Small Themes/Other' since
 *     the auto-categorizer falls back to them.
 *   • EXTRA_HOLDINGS_TABS — holdings tabs are named like "Jan26" (MMMyy) by
 *     default; list any custom-named tabs here (e.g. a family portfolio tab).
 *   • MIN_POSITION_VALUE — stock/ETF positions below this current value are
 *     ignored on import and removed by Apply Cleanup. Options are exempt.
 *
 * v8 changes from v7:
 * • Optional TradeJournal: onOpen() no longer throws when TradeJournal.gs is
 *   absent — the Holdings menu builds regardless (the file's own header calls
 *   the journal "optional but recommended"; now it actually is).
 * • Batched Claude categorization: the Bulk Add dialog used to fire one
 *   synchronous API call PER unknown ticker while rendering, hanging or
 *   timing out on large imports. Unknown tickers are now categorized in a
 *   single API call returning a JSON map.
 * • Add Position merges duplicates: adding a ticker that already exists now
 *   merges into the existing row (qty added to the chosen account) instead of
 *   creating a second row for the same ticker.
 * • Account-column detection no longer uses substring matching: headers are
 *   only skipped when they exactly match the script's own generated Total
 *   columns, so real accounts named e.g. "Total Return Fund" work again.
 * • Batched sheet I/O: applyBatchUpdate / applyCostBasisFormulas /
 *   applyRule1And2 now read each column once and flush writes in bulk instead
 *   of one API call per cell — large pastes no longer flirt with the
 *   6-minute Apps Script execution limit.
 * • applyRule1And2 no longer deletes every priced row when a tab has zero
 *   account columns (value can't be computed, so Rule 1 is skipped).
 * • FX prompt now states the conversion direction explicitly
 *   (USD = foreign ÷ rate) instead of leaving it to inference.
 * • Model override: set the Script Property API_MODEL_OVERRIDE to pin a
 *   different Anthropic model without editing code.
 * • Bulk Add dialog alerts clearly when a tab has no sections yet, instead of
 *   silently skipping every position.
 *
 * v7 changes from v6:
 * • Stale Row Bug Fix: Sections are dynamically re-calculated inside loops to prevent rows from inserting above headers.
 * • Historical Memory: Script maps existing base tickers and auto-selects their historical category for new positions.
 * • Batch Auto-Sync: Changing a category for one ticker in the Bulk dialog auto-updates all sibling tickers in the same batch.
 */

// =============================================================================
// CONFIG
// =============================================================================
const CONFIG = {
  HEADER_ROW: 1,
  TICKER_COL: 1, PRICE_COL: 2, AVG_COST_COL: 3,
  GAIN_DOLLAR_COL: 4, GAIN_PCT_COL: 5,
  FIRST_ACCOUNT_COL: 6,
  MIN_POSITION_VALUE: 1000,
  API_MODEL: 'claude-haiku-4-5-20251001',
  API_KEY_PROPERTY: 'ANTHROPIC_API_KEY',
  FX_RATES_PROPERTY: 'FX_RATES',
  FX_REFRESH_TIMESTAMP_PROPERTY: 'FX_REFRESHED_AT',
  FX_AUTO_REFRESH_DAYS: 30,
  HELPER_SHEET_PREFIX: '_costbasis_',
  // Custom-named holdings tabs beyond the MMMyy monthly tabs, e.g. ['Family'].
  EXTRA_HOLDINGS_TABS: [],
  DEFAULT_FX_RATES: {
    USD: 1.0,  EUR: 0.92,  GBP: 0.79,  JPY: 156.0, CNY: 7.20,
    HKD: 7.80, TWD: 32.50, KRW: 1370.0, AUD: 1.51, CAD: 1.37,
    CHF: 0.89, INR: 83.50, SGD: 1.35,  MXN: 17.50, BRL: 5.10
  },
  DEFAULT_SECTIONS: {
    'Technology': '#cfe2f3', 'Healthcare': '#d9ead3',
    'Energy': '#fce5cd', 'Financials': '#d9d2e9',
    'Consumer': '#fff2cc', 'Industrials': '#c9daf8',
    'Index': '#d0e0e3', 'Small Themes/Other': '#ead1dc',
    'Liquid': '#b6d7a8', 'Puts': '#ea9999', 'Crypto': '#ffd966'
  }
};

// =============================================================================
// MENU
// =============================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Holdings')
    .addItem('Update Account(s)…',  'showUpdateAccountDialog')
    .addItem('Add Position…',       'showAddPositionDialog')
    .addSeparator()
    .addItem('Add Section…',        'showAddSectionDialog')
    .addItem('Quick Price Update…', 'showQuickPriceUpdateDialog')
    .addItem('Add Account…',        'showAddAccountDialog')
    .addSeparator()
    .addItem('Apply Cleanup',       'applyCleanupRules')
    .addItem('Recalc Formulas',     'recalculateAllFormulas')
    .addSeparator()
    .addItem('Start New Month…',    'startNewMonth')
    .addSeparator()
    .addSubMenu(ui.createMenu('Settings')
      .addItem('Set Anthropic API Key',  'setApiKey')
      .addItem('Test API Connection',    'testApiConnection')
      .addItem('Clear API Key',          'clearApiKey')
      .addSeparator()
      .addItem('View/Edit FX Rates',     'showFxRatesDialog')
      .addItem('Refresh FX Rates Now',   'refreshFxRatesManual')
      .addItem('Reset FX Rates to defaults', 'resetFxRates')
      .addSeparator()
      .addItem('Show Helper Sheet',      'showHelperSheet')
      .addItem('Hide Helper Sheet',      'hideHelperSheet'))
    .addToUi();
  // TradeJournal.gs is optional — a missing file must not kill the whole menu.
  if (typeof addTradeJournalMenu === 'function') addTradeJournalMenu(ui);
}

// =============================================================================
// API KEY MANAGEMENT — key lives in Script Properties, never in the sheet
// =============================================================================
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Set Anthropic API Key',
    'Paste your Anthropic API key (sk-ant-…).\n\nIt is stored in this script\'s Script Properties — never in the sheet itself.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (!key) { ui.alert('No key entered.'); return; }
  PropertiesService.getScriptProperties().setProperty(CONFIG.API_KEY_PROPERTY, key);
  ui.alert('✓ API key saved.');
}

function clearApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.API_KEY_PROPERTY);
  SpreadsheetApp.getUi().alert('API key cleared.');
}

function testApiConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const txt = callClaudeAPI([{ type: 'text', text: 'Reply with the single word OK.' }]);
    ui.alert('✓ API connection works. Model replied: ' + String(txt).slice(0, 80));
  } catch (e) {
    ui.alert('✗ ' + e.message);
  }
}

// =============================================================================
// SHEET HELPERS
// =============================================================================
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
function showQuickPriceUpdateDialog() {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return;
  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 12px; font-size: 13px; }
  label { display: block; margin: 8px 0 4px; font-weight: 600; }
  textarea { width: 100%; height: 200px; font-family: monospace; font-size: 12px;
             padding: 8px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; }
  button { padding: 8px 16px; font-size: 13px; cursor: pointer; border-radius: 4px; margin-top: 10px; }
  .primary { background: #1a73e8; color: white; border: none; }
  .secondary { background: white; border: 1px solid #ccc; margin-left: 6px; }
  .hint { color: #666; font-size: 11px; margin-top: 4px; }
  #status { margin-top: 10px; font-family: monospace; font-size: 11px;
            white-space: pre-wrap; max-height: 120px; overflow-y: auto; }
</style>
<label>One ticker + price per line — updates column B only:</label>
<textarea id="data" placeholder="AAPL 231.50&#10;XLU 0.285&#10;XLU 1/21/28 70C 0.31&#10;BIL 91.70"></textarea>
<div class="hint">Stocks: TICKER PRICE &nbsp;|&nbsp; Options: TICKER M/D/YY STRIKEC PRICE</div>
<button class="primary" onclick="apply()">Update Prices</button>
<button class="secondary" onclick="google.script.host.close()">Cancel</button>
<div id="status"></div>
<script>
  function apply() {
    const lines = document.getElementById('data').value;
    document.getElementById('status').textContent = 'Updating…';
    google.script.run
      .withSuccessHandler(r => {
        document.getElementById('status').textContent = r.message;
        if (r.ok) setTimeout(() => google.script.host.close(), 1800);
      })
      .withFailureHandler(e => {
        document.getElementById('status').textContent = '✗ ' + e.message;
      })
      .applyQuickPriceUpdate(lines);
  }
<\/script>
  `).setWidth(440).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Price Update');
}

function applyQuickPriceUpdate(text) {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return { ok: false, message: 'Not on a holdings tab' };

  const lines = String(text || '').split('\n').map(l => l.trim()).filter(l => l);
  const log = [], errors = [];

  for (const line of lines) {
    const parts = line.split(/\s+/);
    const price = parseFloat(parts[parts.length - 1]);
    if (isNaN(price) || parts.length < 2) {
      errors.push(`Bad format: "${line}"`); continue;
    }
    const label = canonicalLabel(parts.slice(0, -1).join(' '));
    const row = findTickerRow(sheet, label);
    if (row === -1) {
      errors.push(`Not in sheet: "${label}"`); continue;
    }
    sheet.getRange(row, CONFIG.PRICE_COL).setValue(price);
    log.push(`✓ ${label} → $${price}`);
  }

  if (log.length > 0) {
    recalculateGainLoss();
    recalculateTotals(sheet);
  }

  const msg = log.join('\n') + (errors.length ? '\n\nSkipped:\n' + errors.join('\n') : '');
  return { ok: log.length > 0, message: msg || 'Nothing updated.' };
}
function isHoldingsTab(name) {
  if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\d{2}$/i.test(name)) return true;
  const n = String(name).trim().toLowerCase();
  return CONFIG.EXTRA_HOLDINGS_TABS.some(t => String(t).trim().toLowerCase() === n);
}

function getActiveHoldingsSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!isHoldingsTab(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Only works on a holdings tab. You are on: "' + sheet.getName() + '"');
    return null;
  }
  return sheet;
}

function getAccountColumns(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), CONFIG.FIRST_ACCOUNT_COL);
  const headers = sheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_ACCOUNT_COL, 1, lastCol - CONFIG.FIRST_ACCOUNT_COL + 1).getValues()[0];
  const out = {};

  // Headers this script itself writes — never treat them as accounts.
  // (Previously this used substring matching, which wrongly skipped real
  // accounts named e.g. "Total Return Fund" or "Value Partners".)
  const GENERATED = ['Total Qty', 'Total Value $', '% of Portfolio'];

  headers.forEach((n, i) => {
    const s = String(n || '').trim();
    const lower = s.toLowerCase();
    const isGenerated = GENERATED.some(g => g.toLowerCase() === lower);

    if (s && !isGenerated) {
      out[s] = CONFIG.FIRST_ACCOUNT_COL + i;
    }
  });

  return out;
}

function canonicalLabel(label) {
  const s = String(label).trim();
  const m = s.match(/^([A-Z][A-Z0-9\.]*)\s+(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d+(?:\.\d+)?)([CP])\b/i);
  if (!m) return s.toUpperCase();
  const [, tk, mo, day, yr, strike, cp] = m;
  const yy = yr.length === 4 ? yr.slice(2) : yr;
  return `${tk.toUpperCase()} ${parseInt(mo, 10)}/${parseInt(day, 10)}/${yy} ${strike}${cp.toUpperCase()}`;
}

function isOptionLabel(label) {
  return /\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d+(?:\.\d+)?[CP]\b/i.test(label);
}

function getBaseTicker(label) {
  return String(label).trim().split(/\s+/)[0].toUpperCase();
}

function findTickerRow(sheet, label) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const target = canonicalLabel(label);
  const values = sheet.getRange(2, CONFIG.TICKER_COL, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const cell = String(values[i][0] || '').trim();
    if (!cell) continue;
    if (canonicalLabel(cell) === target) return i + 2;
  }
  return -1;
}

function getAllPositionRows(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, CONFIG.TICKER_COL, last - 1, CONFIG.PRICE_COL).getValues();
  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const t = String(data[i][0] || '').trim();
    const p = data[i][1];
    if (!t) continue;
    if (p === '' || p === null) continue;
    rows.push({ row: i + 2, ticker: t });
  }
  return rows;
}

function getSectionRows(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, CONFIG.TICKER_COL, last - 1, CONFIG.PRICE_COL).getValues();
  const bg = sheet.getRange(2, CONFIG.TICKER_COL, last - 1, 1).getBackgrounds();
  const out = [];
  for (let i = 0; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    const price = data[i][1];
    if (!name) continue;
    const noPrice = (price === '' || price === null);
    const hasColor = bg[i][0] && bg[i][0].toLowerCase() !== '#ffffff' && bg[i][0] !== '';
    const nameLower = name.toLowerCase();
const known = Object.keys(CONFIG.DEFAULT_SECTIONS).some(k => k.toLowerCase() === nameLower);
    if (noPrice && (hasColor || known)) out.push({ row: i + 2, name });
  }
  return out;
}

function getSectionEndRow(sheet, sectionRow) {
  const sections = getSectionRows(sheet);
  for (const s of sections) if (s.row > sectionRow) return s.row - 1;
  return sheet.getLastRow();
}

function getTickerSectionMap(sheet) {
  const map = {};
  const sections = getSectionRows(sheet);
  const rows = getAllPositionRows(sheet);

  for (const r of rows) {
    let currentSec = null;
    for (let i = sections.length - 1; i >= 0; i--) {
      if (sections[i].row < r.row) {
        currentSec = sections[i].name;
        break;
      }
    }
    if (currentSec) {
      map[getBaseTicker(r.ticker)] = currentSec;
    }
  }
  return map;
}

// =============================================================================
// DYNAMIC TOTALS & ROLLUPS
// =============================================================================
function recalculateTotals(sheet) {
  const accounts = getAccountColumns(sheet);
  const acctCols = Object.values(accounts).sort((a, b) => a - b);
  if (acctCols.length === 0) return;

  const firstAcctCol = acctCols[0];
  const lastAcctCol = acctCols[acctCols.length - 1];
  const firstLetter = columnToLetter(firstAcctCol);
  const lastLetter = columnToLetter(lastAcctCol);

  const qtyCol = lastAcctCol + 1;
  const valCol = lastAcctCol + 2;
  const pctCol = lastAcctCol + 3;
  const qtyLetter = columnToLetter(qtyCol);
  const valLetter = columnToLetter(valCol);
  const pctLetter = columnToLetter(pctCol);

  sheet.getRange(CONFIG.HEADER_ROW, qtyCol).setValue('Total Qty').setFontWeight('bold');
  sheet.getRange(CONFIG.HEADER_ROW, valCol).setValue('Total Value $').setFontWeight('bold');
  sheet.getRange(CONFIG.HEADER_ROW, pctCol).setValue('% of Portfolio').setFontWeight('bold');

  const rows = getAllPositionRows(sheet);
  for (const r of rows) {
    const row = r.row;
    sheet.getRange(row, qtyCol).setFormula(`=IFERROR(SUM(${firstLetter}${row}:${lastLetter}${row}), 0)`).setNumberFormat('0.####');
    sheet.getRange(row, valCol).setFormula(`=IFERROR(B${row}*${qtyLetter}${row}, 0)`).setNumberFormat('$#,##0.00');
    sheet.getRange(row, pctCol).setFormula(`=IFERROR(${valLetter}${row}/SUMIFS(${valLetter}:${valLetter}, B:B, ">0"), "")`).setNumberFormat('0.00%');
  }

  const sections = getSectionRows(sheet);
  for (const s of sections) {
    const r = s.row;
    const endR = getSectionEndRow(sheet, r);
    sheet.getRange(r, qtyCol).clearContent();

    if (endR > r) {
      sheet.getRange(r, valCol).setFormula(`=SUM(${valLetter}${r+1}:${valLetter}${endR})`).setNumberFormat('$#,##0.00').setFontWeight('bold');
    } else {
      sheet.getRange(r, valCol).setValue(0).setNumberFormat('$#,##0.00').setFontWeight('bold');
    }
    sheet.getRange(r, pctCol).setFormula(`=IFERROR(${valLetter}${r}/SUMIFS(${valLetter}:${valLetter}, B:B, ">0"), "")`).setNumberFormat('0.00%').setFontWeight('bold');
  }
}

// =============================================================================
// HELPER SHEET (cost basis storage)
// =============================================================================
function helperSheetName(mainSheetName) {
  return CONFIG.HELPER_SHEET_PREFIX + mainSheetName;
}

function getOrCreateHelperSheet(mainSheet) {
  const ss = mainSheet.getParent();
  const name = helperSheetName(mainSheet.getName());
  let helper = ss.getSheetByName(name);
  if (!helper) {
    helper = ss.insertSheet(name);
    helper.hideSheet();
    helper.getRange(1, 1).setValue('Ticker').setFontWeight('bold');
  }
  const lastCol = Math.max(mainSheet.getLastColumn(), CONFIG.FIRST_ACCOUNT_COL);
  const numAccountCols = lastCol - CONFIG.FIRST_ACCOUNT_COL + 1;
  const mainHeaders = mainSheet.getRange(CONFIG.HEADER_ROW, CONFIG.FIRST_ACCOUNT_COL, 1, numAccountCols).getValues()[0];
  helper.getRange(1, CONFIG.FIRST_ACCOUNT_COL, 1, numAccountCols).setValues([mainHeaders]).setFontWeight('bold');
  return helper;
}

function findHelperRow(helper, ticker) {
  const last = helper.getLastRow();
  if (last < 2) return -1;
  const target = canonicalLabel(ticker);
  const tickers = helper.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < tickers.length; i++) {
    if (canonicalLabel(String(tickers[i][0] || '').trim()) === target) return i + 2;
  }
  return -1;
}

function findOrAddHelperRow(helper, ticker) {
  let row = findHelperRow(helper, ticker);
  if (row > 0) return row;
  const newRow = Math.max(helper.getLastRow(), 1) + 1;
  helper.getRange(newRow, 1).setValue(canonicalLabel(ticker));
  return newRow;
}

function getHelperAvgCost(mainSheet, ticker, accountName) {
  const ss = mainSheet.getParent();
  const helper = ss.getSheetByName(helperSheetName(mainSheet.getName()));
  if (!helper) return null;
  const col = getAccountColumns(mainSheet)[accountName];
  if (!col) return null;
  const row = findHelperRow(helper, ticker);
  if (row < 0) return null;
  const v = helper.getRange(row, col).getValue();
  return (typeof v === 'number' && !isNaN(v)) ? v : null;
}

function setHelperAvgCost(mainSheet, ticker, accountName, cost) {
  const helper = getOrCreateHelperSheet(mainSheet);
  const accounts = getAccountColumns(mainSheet);
  const col = accounts[accountName];
  if (!col) throw new Error('Account not found in row 1: ' + accountName);
  const row = findOrAddHelperRow(helper, ticker);
  if (cost === null || cost === undefined || cost === '' || isNaN(cost)) {
    helper.getRange(row, col).clearContent();
  } else {
    helper.getRange(row, col).setValue(Number(cost));
  }
}

function removeFromHelper(mainSheet, ticker) {
  const ss = mainSheet.getParent();
  const helper = ss.getSheetByName(helperSheetName(mainSheet.getName()));
  if (!helper) return;
  const row = findHelperRow(helper, ticker);
  if (row > 0) helper.deleteRow(row);
}


function applyCostBasisFormulas(mainSheet) {
  const helper     = getOrCreateHelperSheet(mainSheet);
  const rows       = getAllPositionRows(mainSheet);
  const accountCols = Object.values(getAccountColumns(mainSheet)).sort((a, b) => a - b);
  if (accountCols.length === 0) return;

  const firstAcct = accountCols[0];
  const lastAcct  = accountCols[accountCols.length - 1];
  const span      = lastAcct - firstAcct + 1;          // columns to read in one block
  const lastRow   = mainSheet.getLastRow();
  if (lastRow < 2) return;

  // ---- Batch read: main-sheet quantities for the whole account block ----
  const mainQty = mainSheet.getRange(2, firstAcct, lastRow - 1, span).getValues();

  // ---- Batch read: column C (avg cost) values, formulas, notes, backgrounds.
  //      Staged in memory and flushed once at the end instead of one API
  //      call per cell per row. ----
  const cRange    = mainSheet.getRange(2, CONFIG.AVG_COST_COL, lastRow - 1, 1);
  const cFormulas = cRange.getFormulas();
  const cVals     = cRange.getValues();
  const cNotes    = cRange.getNotes();
  const cBgs      = cRange.getBackgrounds();

  // ---- Batch read: helper tickers + cost block; build an in-memory lookup ----
  const helperLast = helper.getLastRow();
  const helperTickerVals = helperLast >= 2
    ? helper.getRange(2, 1, helperLast - 1, 1).getValues()
    : [];
  const helperCostVals = helperLast >= 2
    ? helper.getRange(2, firstAcct, helperLast - 1, span).getValues()
    : [];

  // ticker -> { sheetRow, costs[] }  (costs aligned to firstAcct..lastAcct)
  const helperByTicker = {};
  for (let i = 0; i < helperTickerVals.length; i++) {
    const key = canonicalLabel(String(helperTickerVals[i][0] || '').trim());
    if (key) helperByTicker[key] = { sheetRow: i + 2, costs: helperCostVals[i].slice() };
  }

  const WARN_MISSING = missingCount => '⚠️ Missing cost basis for ' + missingCount + ' account(s) holding '
    + 'this position. Those shares default to the weighted-average cost of the '
    + 'accounts that DO have a cost.';
  const ERR_NO_COST = '⛔ No cost basis recorded for any account holding this position. '
    + 'Average cost / gain $ / gain % cannot be computed until at least one account '
    + 'has a cost. Add it via Update Account(s)… or directly in the helper sheet.';

  for (const r of rows) {
    const idx             = r.row - 2;
    const existingFormula = cFormulas[idx][0];
    const existingValue   = cVals[idx][0];
    const canonTicker     = canonicalLabel(r.ticker);
    const qtyRow          = mainQty[r.row - 2] || [];

    // ---- One-time migration of a legacy HARDCODED cost in C into the helper.
    //      Gated on "ticker not yet in helper" so it never re-fires on the
    //      computed values this function itself writes. ----
    if (!existingFormula && typeof existingValue === 'number' && existingValue > 0
        && !helperByTicker[canonTicker]) {
      const helperRow  = findOrAddHelperRow(helper, r.ticker);
      const migrated   = new Array(span).fill('');
      for (let k = 0; k < span; k++) {
        const qty = qtyRow[k];
        if (typeof qty === 'number' && qty > 0) {
          helper.getRange(helperRow, firstAcct + k).setValue(existingValue);
          migrated[k] = existingValue;
        }
      }
      helperByTicker[canonTicker] = { sheetRow: helperRow, costs: migrated };
    }

    // ---- Compute the weighted average over accounts that (a) hold the
    //      position (qty > 0) AND (b) have a numeric cost in the helper. ----
    const costs = helperByTicker[canonTicker] ? helperByTicker[canonTicker].costs : null;
    let heldQty = 0, knownQty = 0, knownWeighted = 0, missingCount = 0;

    for (let k = 0; k < span; k++) {
      const qty = qtyRow[k];
      if (typeof qty === 'number' && qty > 0) {
        heldQty += qty;
        const cost = costs ? costs[k] : '';
        if (typeof cost === 'number' && !isNaN(cost)) {
          knownQty      += qty;
          knownWeighted += qty * cost;
        } else {
          missingCount++;
        }
      }
    }

    // ---- Stage the C value / note / background; flushed in bulk below. ----
    if (heldQty > 0 && knownQty > 0) {
      cVals[idx][0] = knownWeighted / knownQty;        // weighted avg of known accts
      if (missingCount > 0) {
        cNotes[idx][0] = WARN_MISSING(missingCount);
        cBgs[idx][0] = '#fff2cc';                     // yellow: fell back
      } else {
        cNotes[idx][0] = null;
        cBgs[idx][0] = null;
      }
    } else if (heldQty > 0) {
      // Holds the position but no cost recorded in ANY account -> cannot compute.
      cVals[idx][0] = '';
      cNotes[idx][0] = ERR_NO_COST;
      cBgs[idx][0] = '#f4cccc';                       // red: cannot compute
    } else {
      // Not held (qty 0 everywhere) -> nothing to value.
      cVals[idx][0] = '';
      cNotes[idx][0] = null;
      cBgs[idx][0] = null;
    }
  }

  // ---- One write per attribute for the whole column. ----
  cRange.setValues(cVals);
  cRange.setNotes(cNotes);
  cRange.setBackgrounds(cBgs);
}

function showHelperSheet() {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return;
  const helper = getOrCreateHelperSheet(sheet);
  helper.showSheet();
  SpreadsheetApp.getUi().alert(`"${helper.getName()}" is now visible. Use Hide Helper Sheet when done.`);
}
function hideHelperSheet() {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return;
  const ss = sheet.getParent();
  const helper = ss.getSheetByName(helperSheetName(sheet.getName()));
  if (helper) helper.hideSheet();
}

// =============================================================================
// FX RATES
// =============================================================================
function getFxRates() {
  const stored = PropertiesService.getScriptProperties().getProperty(CONFIG.FX_RATES_PROPERTY);
  if (stored) { try { return JSON.parse(stored); } catch (e) {} }
  return CONFIG.DEFAULT_FX_RATES;
}
function setFxRates(rates) {
  PropertiesService.getScriptProperties().setProperty(CONFIG.FX_RATES_PROPERTY, JSON.stringify(rates));
  PropertiesService.getScriptProperties().setProperty(CONFIG.FX_REFRESH_TIMESTAMP_PROPERTY, new Date().toISOString());
}
function resetFxRates() {
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.FX_RATES_PROPERTY);
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.FX_REFRESH_TIMESTAMP_PROPERTY);
  SpreadsheetApp.getUi().alert('FX rates reset to defaults.');
}
function refreshFxRatesSilent() {
  try {
    const resp = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/USD', { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return false;
    const body = JSON.parse(resp.getContentText());
    if (body.result !== 'success' || !body.rates) return false;
    const wanted = Object.keys(CONFIG.DEFAULT_FX_RATES);
    const fresh = { USD: 1.0 };
    for (const code of wanted) if (body.rates[code]) fresh[code] = Number(body.rates[code]);
    setFxRates(fresh);
    return true;
  } catch (e) { return false; }
}
function maybeAutoRefreshFx() {
  const tsStr = PropertiesService.getScriptProperties().getProperty(CONFIG.FX_REFRESH_TIMESTAMP_PROPERTY);
  if (!tsStr) { refreshFxRatesSilent(); return; }
  const ageMs = Date.now() - new Date(tsStr).getTime();
  const thresholdMs = CONFIG.FX_AUTO_REFRESH_DAYS * 24 * 60 * 60 * 1000;
  if (ageMs > thresholdMs) refreshFxRatesSilent();
}
function refreshFxRatesManual() {
  const ui = SpreadsheetApp.getUi();
  if (refreshFxRatesSilent()) {
    const rates = getFxRates();
    const summary = Object.entries(rates).map(([k, v]) => `${k}: ${v}`).join('\n');
    ui.alert(`✓ FX rates refreshed (1 USD = X):\n\n${summary}`);
  } else {
    ui.alert('✗ Could not refresh FX rates. Using existing cached rates.');
  }
}
function showFxRatesDialog() {
  const rates = getFxRates();
  const ts = PropertiesService.getScriptProperties().getProperty(CONFIG.FX_REFRESH_TIMESTAMP_PROPERTY);
  const tsLabel = ts ? `Last refreshed: ${new Date(ts).toLocaleString()}` : 'Using defaults (never refreshed).';
  const rateRows = Object.entries(rates).map(([code, val]) =>
    `<tr><td>${code}</td><td><input type="number" step="0.0001" data-code="${code}" value="${val}" style="width:140px"></td></tr>`
  ).join('');
  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 12px; font-size: 13px; }
  table { border-collapse: collapse; margin: 8px 0; }
  th, td { padding: 4px 10px; text-align: left; border-bottom: 1px solid #eee; }
  th { background: #f8f9fa; font-size: 12px; }
  .meta { color: #5f6368; font-size: 11px; margin-bottom: 8px; }
  button { padding: 8px 16px; margin-right: 6px; font-size: 13px; cursor: pointer; border-radius: 4px; }
  .primary { background: #1a73e8; color: white; border: none; }
  .secondary { background: white; border: 1px solid #ccc; }
  #status { margin-top: 12px; font-size: 12px; }
</style>
<div class="meta">1 USD = X foreign units. ${tsLabel}<br>Auto-refresh: every ${CONFIG.FX_AUTO_REFRESH_DAYS} days when you Parse with AI.</div>
<table><tr><th>Currency</th><th>Rate</th></tr>${rateRows}</table>
<button class="primary" onclick="save()">Save</button>
<button class="secondary" onclick="google.script.host.close()">Cancel</button>
<div id="status"></div>
<script>
  function save() {
    const inputs = document.querySelectorAll('input[data-code]');
    const rates = {};
    for (const inp of inputs) {
      const v = parseFloat(inp.value);
      if (!isNaN(v)) rates[inp.getAttribute('data-code')] = v;
    }
    google.script.run.withSuccessHandler(() => {
      document.getElementById('status').textContent = '✓ Saved.';
      setTimeout(() => google.script.host.close(), 600);
    }).setFxRates(rates);
  }
<\/script>
  `).setWidth(380).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'FX Rates');
}

// =============================================================================
// ANTHROPIC API (Batch Updated)
// =============================================================================
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty(CONFIG.API_KEY_PROPERTY);
  if (!key) throw new Error('No API key set. Holdings → Settings → Set Anthropic API Key first.');
  return key;
}

// Model override: set the Script Property API_MODEL_OVERRIDE to pin a
// different Anthropic model without editing code.
function getApiModel() {
  const override = PropertiesService.getScriptProperties().getProperty('API_MODEL_OVERRIDE');
  return override || CONFIG.API_MODEL;
}

function buildSystemPrompt() {
  const rates = getFxRates();
  const rateLines = Object.entries(rates).filter(([k]) => k !== 'USD')
    .map(([code, val]) => `  ${code}: 1 USD = ${val} ${code}`).join('\n');
  return `You parse multiple brokerage account data blocks and return ONLY a single JSON array, no markdown, no commentary.

You will receive data wrapped in "--- BEGIN DATA FOR ACCOUNT: [MasterAccountName] ---".
Each item MUST use this [MasterAccountName] exactly, ignoring any internal sub-account names found in the raw text.
{"account": "MasterAccountName", "label": "...", "qty": number, "price": number, "avg_cost": number|null}

LABEL FORMAT:
- US stocks: uppercase ticker — "AAPL", "NVDA", "BRK.B"
- STRIP ASTERISKS: Remove any "**" from money market/sweep symbols (e.g., "FDRXX**" becomes "FDRXX", "CORE**" becomes "CORE").
- Options: "TICKER M/D/YY STRIKE[C|P]"
  - No leading zeros: "1/15/27" not "01/15/2027"; 2-digit year
  - Strike: integer if whole ("30"), decimal if not ("7.5")
  - C=Call, P=Put, uppercase
- Foreign-listed stocks: USE THE EXCHANGE SYMBOL AS IT APPEARS. E.g., "322310" (KRX), "3105" (TWSE), "9988.HK"

QTY & ACCOUNT (CRITICAL — AGGREGATE EVERYTHING):
- IGNORE internal sub-account names inside the raw data (like "BrokerageLink", "Roth IRA", etc.).
- SUM quantities for the same ticker across ALL sub-accounts within the block.
- Return only ONE row per ticker, mapped strictly to the [MasterAccountName].
- Options: number of CONTRACTS — do NOT multiply by 100
CASH/SWEEP POSITIONS:
- Named money market funds (FDRXX, SPAXX, FZFXX, CORE**, etc.) where Quantity and Price are blank: use the dollar amount in the Current Value column as \`qty\`, set \`price\` to 1.00. Strip trailing asterisks from the symbol.
- Rows where the first field is "Cash & Cash Investments" or "Cash" with no real ticker: use the first dollar amount in the row as \`qty\`, set \`price\` to 1.00, label it "CASH".

PRICE (always in USD):
- Current market price per share/contract
- Convert foreign currency to USD by DIVIDING by the rate below:
${rateLines}
  Conversion: USD amount = foreign amount ÷ rate (e.g. €920 at 0.92 → $1,000). Never multiply by the rate.

AVG_COST (always in USD, weighted across sub-accounts):
- avg_cost = SUM(qty_i × cost_i) / SUM(qty_i)
- Convert foreign-currency cost to USD using the rates above
- For Schwab CSVs (identifiable by headers "Qty (Quantity)", "Cost Basis", "Mkt Val", "Price" in any order):
  "Cost Basis" is the TOTAL dollars paid for the whole position, not per share.
  Stocks: avg_cost = Cost Basis / Qty.
  Options: avg_cost = Cost Basis / (Qty × 100), because Qty is contracts and Price is per share.
  EXAMPLE: 105 contracts, Cost Basis $4,479.50 → avg_cost = $4,479.50 ÷ (105 × 100) = $0.4266, NOT $42.66.
  VALIDATION: option avg_cost should be similar in magnitude to Price. If avg_cost ≈ 100× Price, you forgot the ×100.
- For Fidelity CSVs with an "Average Cost Basis" column: use it directly for stocks. For options divide by 100.
- For Fidelity CSVs with only a "Cost Basis" column (total, no per-share average): same approach as Schwab above.

SKIP:
- "Not Priced Today", "Positions Total", "Account Total", "Pending Activity", and any row that is a portfolio summary or subtotal.
- Section headers.
- Zero-quantity rows (EXCEPT cash/sweep positions described above).

Return [] if nothing parseable. Output JUST the JSON array.`;
}

function callClaudeAPI(messages) {
  const apiKey = getApiKey();
  const payload = {
    model: getApiModel(),
    max_tokens: 4096,
    system: buildSystemPrompt(),
    messages: [{ role: 'user', content: messages }]
  };
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const raw = resp.getContentText();
  let body;
  try { body = JSON.parse(raw); } catch (e) { throw new Error(`API returned non-JSON (${code}): ${raw.slice(0, 200)}`); }
  if (code !== 200) {
    const msg = (body.error && body.error.message) || raw;
    throw new Error(`API error ${code}: ${msg}`);
  }
  return body.content.filter(c => c.type === 'text').map(c => c.text).join('\n');
}

function parseBrokerBatch(blocks) {
  try {
    maybeAutoRefreshFx();
    const messages = [];
    for (const b of blocks) {
       messages.push({type: 'text', text: `--- BEGIN DATA FOR ACCOUNT: ${b.account} ---`});
       if (b.image) {
         messages.push({type: 'image', source: {type: 'base64', media_type: b.image.type, data: b.image.data}});
         messages.push({type: 'text', text: "Parse the brokerage positions from this screenshot."});
       }
       if (b.text && !b.image) {
         messages.push({type: 'text', text: b.text});
       }
       messages.push({type: 'text', text: `--- END DATA FOR ACCOUNT: ${b.account} ---`});
    }
    const txt = callClaudeAPI(messages);
    return { ok: true, text: jsonToStructuredLinesBatch(txt) };
  } catch (e) { return { ok: false, message: e.message }; }
}

function jsonToStructuredLinesBatch(apiText) {
  let cleaned = String(apiText).replace(/```json\s*/gi, '').replace(/```\s*$/g, '').trim();
  const m = cleaned.match(/\[\s*[\s\S]*\]/);
  if (m) cleaned = m[0];
  let positions;
  try { positions = JSON.parse(cleaned); }
  catch (e) { throw new Error('AI response was not valid JSON. Raw:\n' + apiText.slice(0, 400)); }
  if (!Array.isArray(positions)) throw new Error('AI did not return a JSON array');
  if (positions.length === 0) return '# No positions detected.';
  return positions.map(p => {
    const qty = (typeof p.qty === 'number') ? p.qty : parseFloat(p.qty);
    const price = (typeof p.price === 'number') ? p.price : parseFloat(p.price);
    const ac = (p.avg_cost === null || p.avg_cost === undefined || p.avg_cost === '') ? null
             : ((typeof p.avg_cost === 'number') ? p.avg_cost : parseFloat(p.avg_cost));
    const acStr = ac !== null && !isNaN(ac) ? `  ${ac}` : '';
    return `[${p.account}]  ${p.label}  ${qty}  ${price}${acStr}`;
  }).join('\n');
}

// =============================================================================
// UPDATE ACCOUNT(S) DIALOG
// =============================================================================
function showUpdateAccountDialog() {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return;
  const accounts = Object.keys(getAccountColumns(sheet));
  const sections = getSectionRows(sheet).map(s => s.name);
  if (accounts.length === 0) {
    SpreadsheetApp.getUi().alert('No account columns. Use "Add Account…" first.');
    return;
  }
  const tabName = sheet.getName();

  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 12px; font-size: 13px; color: #202124; }
  label { display: block; margin: 10px 0 4px; font-weight: 600; }
  select, textarea { width: 100%; box-sizing: border-box; font-size: 12px; }
  textarea { height: 100px; padding: 8px; font-family: 'SF Mono', Menlo, monospace; border: 1px solid #ccc; border-radius: 4px; }
  textarea:focus { outline: 2px solid #1a73e8; border-color: transparent; }
  .badge { display: inline-block; background: #e8f0fe; color: #1a73e8; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
  .btnrow { display: flex; gap: 8px; margin-top: 14px; flex-wrap: wrap; }
  button { padding: 8px 14px; font-size: 13px; cursor: pointer; border-radius: 4px; font-weight: 500; }
  .primary { background: #1a73e8; color: white; border: none; } .primary:hover { background: #1557b0; }
  .ai { background: #f4b400; color: #000; border: none; } .ai:hover { background: #d99e00; }
  .secondary { background: white; border: 1px solid #dadce0; } .secondary:hover { background: #f8f9fa; }
  .check { display: flex; align-items: center; gap: 6px; margin-top: 8px; font-weight: normal; font-size: 12px; }

  .block-container { background: #f8f9fa; border: 1px solid #dadce0; padding: 10px; border-radius: 6px; margin-bottom: 12px; }
  .hint { color: #5f6368; font-size: 11px; margin-top: 4px; }
  #status { margin-top: 12px; white-space: pre-wrap; font-family: monospace; font-size: 11px; max-height: 100px; overflow-y: auto; padding: 8px; background: #f8f9fa; border-radius: 4px; min-height: 20px; }
  .working { color: #1a73e8; } .err { color: #d93025; } .ok { color: #188038; }
  #banner { display:none; margin-top:12px; padding:10px 14px; border-radius:6px; background:#e6f4ea; color:#137333; font-weight:600; font-size:13px; border:1px solid #a8d5b5; }
  #banner.warn { background:#fef7e0; color:#b06000; border-color:#f9c945; }
</style>

<div><span class="badge">${tabName}</span></div>

<div id="setup-phase">
  <div id="blocks-container"></div>
  <button class="secondary" id="add-block-btn" onclick="addBlock()" style="width:100%; margin-bottom: 10px;">+ Add Another Account</button>
</div>

<div id="review-phase" style="display:none;">
  <label>Review AI Output <span class="hint">— Check lines before applying</span></label>
  <textarea id="review-data" style="height: 250px;"></textarea>
</div>

<label class="check"><input type="checkbox" id="full" checked /><span>Full snapshot — clear tickers NOT in the data (for included accounts only)</span></label>
<label class="check"><input type="checkbox" id="updateCost" checked /><span>Update broker avg cost in helper sheet</span></label>

<div class="btnrow">
  <button class="ai" id="parse-btn" onclick="parseAIBatch()">🤖 Parse with AI</button>
  <button class="primary" id="apply-btn" onclick="apply()" style="display:none;">Apply to Sheet</button>
  <button class="secondary" onclick="google.script.host.close()">Cancel</button>
</div>

<div id="banner"></div>
<div id="status">Ready.</div>

<script>
  window.pendingNewTickers = [];
  var ACCOUNTS = ${JSON.stringify(accounts)};

  let nextId = 1;
  let pendingImages = {}; // id -> {type, data}

  function setStatus(msg, cls) {
    const el = document.getElementById('status');
    el.textContent = msg; el.className = cls || '';
  }

  function addBlock() {
    const id = nextId++;
    const div = document.createElement('div');
    div.className = 'block-container';
    div.id = 'block-' + id;
    div.innerHTML = \`
      <div style="display:flex; gap:10px; margin-bottom: 6px;">
        <select id="acct-\${id}" style="flex:1;">\${ACCOUNTS.map(a => '<option>'+a+'</option>').join('')}</select>
        <button class="secondary" onclick="document.getElementById('block-\${id}').remove()" style="padding:2px 8px; font-size:11px;">Remove</button>
      </div>
      <textarea id="data-\${id}" placeholder="Paste CSV, lines, or drop screenshot here..." style="height:70px;"></textarea>
      <div id="status-\${id}" class="hint">Ready</div>
    \`;
    document.getElementById('blocks-container').appendChild(div);
    attachEvents(id);
  }

  function loadImage(id, blob, label) {
    const reader = new FileReader();
    reader.onload = function(ev) {
      pendingImages[id] = { type: blob.type, data: ev.target.result.split(',')[1] };
      document.getElementById('data-' + id).value = '[Image ready: ' + (label || blob.type) + ' — click 🤖 Parse with AI]';
      document.getElementById('status-' + id).textContent = '📷 Image queued.';
      document.getElementById('status-' + id).style.color = '#188038';
    };
    reader.readAsDataURL(blob);
  }

  function loadFileText(id, file) {
    const reader = new FileReader();
    reader.onload = function(ev) {
      document.getElementById('data-' + id).value = ev.target.result;
      document.getElementById('status-' + id).textContent = '📄 Text file loaded.';
    };
    reader.readAsText(file);
  }

  function attachEvents(id) {
    const ta = document.getElementById('data-' + id);
    ta.addEventListener('paste', function(e) {
      if (!e.clipboardData) return;
      for (let i = 0; i < e.clipboardData.items.length; i++) {
        let item = e.clipboardData.items[i];
        if (item.type && item.type.indexOf('image/') === 0) {
          e.preventDefault();
          loadImage(id, item.getAsFile(), 'pasted');
          return;
        }
      }
    });
    ta.addEventListener('dragover', e => { e.preventDefault(); ta.style.backgroundColor = '#e8f0fe'; });
    ta.addEventListener('dragleave', e => { e.preventDefault(); ta.style.backgroundColor = ''; });
    ta.addEventListener('drop', e => {
      e.preventDefault(); ta.style.backgroundColor = '';
      const f = e.dataTransfer.files[0];
      if (f) {
         if ((f.type || '').indexOf('image/') === 0) loadImage(id, f, f.name);
         else loadFileText(id, f);
      }
    });
    ta.addEventListener('input', e => {
      if (pendingImages[id] && e.target.value.indexOf('[Image ready') !== 0) {
        delete pendingImages[id];
        document.getElementById('status-' + id).textContent = 'Ready';
        document.getElementById('status-' + id).style.color = '';
      }
    });
  }

  // Init first block
  addBlock();

  let requestedAccountsForBatch = [];

  function parseAIBatch() {
    const reqBlocks = [];
    requestedAccountsForBatch = [];

    const blockDivs = document.querySelectorAll('.block-container');
    for (const div of blockDivs) {
      const id = div.id.replace('block-', '');
      const acct = document.getElementById('acct-' + id).value;
      const txt = document.getElementById('data-' + id).value;
      const img = pendingImages[id];
      if (txt.trim() || img) {
        reqBlocks.push({ account: acct, text: txt, image: img });
        if (!requestedAccountsForBatch.includes(acct)) requestedAccountsForBatch.push(acct);
      }
    }

    if (reqBlocks.length === 0) { setStatus('Paste data or add images first.', 'err'); return; }

    setStatus('🤖 Parsing Batch Data… (may take a few seconds)', 'working');

    google.script.run.withSuccessHandler(r => {
      if (r.ok) {
        document.getElementById('setup-phase').style.display = 'none';
        document.getElementById('review-phase').style.display = 'block';
        document.getElementById('parse-btn').style.display = 'none';
        document.getElementById('apply-btn').style.display = 'inline-block';
        document.getElementById('review-data').value = r.text;
        setStatus('✓ Parsed successfully. Review the lines, then click Apply.', 'ok');
      } else {
        setStatus('✗ ' + r.message, 'err');
      }
    }).withFailureHandler(e => setStatus('✗ ' + e.message, 'err')).parseBrokerBatch(reqBlocks);
  }

  function openBulkAdd() {
    google.script.run.showBulkAddDialog(JSON.stringify(window.pendingNewTickers));
  }

  function apply() {
    const textData = document.getElementById('review-data').value;
    const full = document.getElementById('full').checked;
    const updateCost = document.getElementById('updateCost').checked;

    const btn = document.getElementById('apply-btn');
    btn.disabled = true; btn.textContent = 'Applying…';
    setStatus('Working…', 'working');

    google.script.run
      .withSuccessHandler(function(r) {
        btn.disabled = false; btn.textContent = 'Apply to Sheet';
        if (!r.ok) { setStatus('✗ ' + r.message, 'err'); return; }
        var hasNew = r.newTickers && r.newTickers.length > 0;
        var banner = document.getElementById('banner');
        banner.style.display = 'block';
        if (hasNew) {
          window.pendingNewTickers = r.newTickers;
          banner.className = 'warn';
          banner.innerHTML = '✓ Applied — <strong>' + r.newTickers.length + ' new ticker(s)</strong> not yet in sheet. &nbsp;<button onclick="openBulkAdd()" style="padding:4px 12px;font-size:12px;cursor:pointer;border-radius:3px;background:#1a73e8;color:white;border:none;font-weight:600;">Add All New Positions →<\/button>';
          setStatus(r.message, 'ok');
        } else {
          banner.className = '';
          banner.textContent = '✓ Done! Sheet updated. Closing…';
          setStatus(r.message, 'ok');
          setTimeout(function() { google.script.host.close(); }, 2500);
        }
      })
      .withFailureHandler(function(e) {
        btn.disabled = false; btn.textContent = 'Apply to Sheet';
        setStatus('✗ ' + e.message, 'err');
      })
      .applyBatchUpdate(requestedAccountsForBatch, textData, full, updateCost);
  }
</script>
  `).setWidth(660).setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Accounts (Batch)');
}

function parsePastedBatchData(text) {
  const lines = String(text || '').split('\n').map(l => l.trim()).filter(l => l && !l.startsWith('#'));
  const OPT_RE = /^([A-Z][A-Z0-9\.]*)\s+(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d+(?:\.\d+)?)([CP])\b(.*)$/i;
  const positions = [];
  const errors = [];

  for (const line of lines) {
    const acctMatch = line.match(/^\[(.*?)\]\s+(.*)$/);
    if (!acctMatch) {
      errors.push(`Missing [Account] prefix: "${line}"`);
      continue;
    }
    const account = acctMatch[1].trim();
    const restLine = acctMatch[2].trim();

    let label, rest;
    const m = restLine.match(OPT_RE);
    if (m) {
      label = canonicalLabel(`${m[1]} ${m[2]} ${m[3]}${m[4]}`);
      rest = m[5];
    } else {
      const parts = restLine.split(/[\s,]+/).filter(p => p);
      label = parts[0].toUpperCase();
      rest = parts.slice(1).join(' ');
    }
    const nums = rest.split(/[\s,]+/).filter(p => p).map(p => parseFloat(String(p).replace(/[$,]/g, '')));
    if (nums.length < 2 || nums.slice(0, 2).some(isNaN)) {
      errors.push(`Could not parse data for ${account}: "${restLine}"`);
      continue;
    }
    const qty = nums[0], price = nums[1];
    const avgCost = (nums.length >= 3 && !isNaN(nums[2])) ? nums[2] : null;
    const isOpt = isOptionLabel(label);
    positions.push({ account, label, qty: isOpt ? qty * 100 : qty, price, avgCost, isOption: isOpt });
  }
  return { positions, errors };
}

function applyBatchUpdate(requestedAccounts, pastedText, fullSnapshot, updateAvgCost) {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return { ok: false, message: 'Not on a holdings tab' };
  const accounts = getAccountColumns(sheet);

  const { positions, errors } = parsePastedBatchData(pastedText);
  if (positions.length === 0 && errors.length > 0) {
    return { ok: false, message: 'Could not parse any lines:\n' + errors.join('\n') };
  }
  if (positions.length === 0) return { ok: false, message: 'No positions to apply.' };

  getOrCreateHelperSheet(sheet);

  // ---- Batched reads: one API call per column instead of one per cell. ----
  const lastRow = sheet.getLastRow();
  const colA = lastRow >= 2
    ? sheet.getRange(2, CONFIG.TICKER_COL, lastRow - 1, 1).getValues()
    : [];
  // canonical label -> row (first occurrence wins, same as findTickerRow)
  const tickerRowMap = {};
  for (let i = 0; i < colA.length; i++) {
    const cell = String(colA[i][0] || '').trim();
    if (!cell) continue;
    const key = canonicalLabel(cell);
    if (!(key in tickerRowMap)) tickerRowMap[key] = i + 2;
  }
  const priceVals = lastRow >= 2
    ? sheet.getRange(2, CONFIG.PRICE_COL, lastRow - 1, 1).getValues()
    : [];
  const acctCols = Object.values(accounts).sort((a, b) => a - b);
  const colToK = {};
  acctCols.forEach((c, i) => { colToK[c] = i; });
  let qtyBlock = [];
  if (lastRow >= 2 && acctCols.length > 0) {
    const firstAcct = acctCols[0];
    const span = acctCols[acctCols.length - 1] - firstAcct + 1;
    qtyBlock = sheet.getRange(2, firstAcct, lastRow - 1, span).getValues();
  }
  const touchedCols = {};

  const log = [];
  const seenInAccount = {};
  requestedAccounts.forEach(a => seenInAccount[a] = new Set());
  const newTickersMap = new Map();

  for (const p of positions) {
    const col = accounts[p.account];
    if (!col) {
      errors.push(`Unknown account "${p.account}" ignored.`);
      continue;
    }

    // ALWAYS register in seenInAccount FIRST, before any filtering,
    // so fullSnapshot clearing works correctly even for filtered/new positions.
    if (seenInAccount[p.account]) seenInAccount[p.account].add(p.label);

    const row = tickerRowMap[canonicalLabel(p.label)];
    if (row === undefined) {
      // For new positions: options bypass the value filter (current mkt value of a
      // deep OTM option is irrelevant — what matters is you own it / paid for it).
      // For stocks/ETFs, keep the $1,000 current-value floor.
      const totalValue = p.price * p.qty;
      if (!p.isOption && totalValue < CONFIG.MIN_POSITION_VALUE) {
        log.push(`• Ignored new position ${p.label} (Value $${totalValue.toFixed(2)} is under $${CONFIG.MIN_POSITION_VALUE})`);
        continue;
      }

      const key = `${p.account}-${p.label}`;
      if (!newTickersMap.has(key)) {
        newTickersMap.set(key, { account: p.account, label: p.label, qty: p.qty, price: p.price, avgCost: p.avgCost, isOption: p.isOption });
      }
      continue;
    }

    // Stage the writes in memory; flushed in bulk below.
    priceVals[row - 2][0] = p.price;
    qtyBlock[row - 2][colToK[col]] = p.qty;
    touchedCols[col] = true;
    if (updateAvgCost && p.avgCost !== null) {
      setHelperAvgCost(sheet, p.label, p.account, p.avgCost);
      log.push(`✓ [${p.account}] ${p.label} → $${p.price}, qty ${p.qty}, avg $${p.avgCost}`);
    } else {
      const warning = (p.qty > 0) ? ` (⚠️ No avg cost provided)` : '';
      log.push(`✓ [${p.account}] ${p.label} → $${p.price}, qty ${p.qty}${warning}`);
    }
  }

  if (fullSnapshot) {
    for (let i = 0; i < colA.length; i++) {
      const t = String(colA[i][0] || '').trim();
      const px = priceVals[i] ? priceVals[i][0] : null;
      if (!t || px === '' || px === null) continue; // same filter as getAllPositionRows
      const canon = canonicalLabel(t);
      for (const acct of requestedAccounts) {
        const col = accounts[acct];
        if (!col || !seenInAccount[acct] || seenInAccount[acct].has(canon)) continue;
        const k = colToK[col];
        const cur = qtyBlock[i][k];
        if (cur !== '' && cur !== null && cur !== undefined) {
          qtyBlock[i][k] = '';
          touchedCols[col] = true;
          setHelperAvgCost(sheet, canon, acct, null);
          log.push(`• cleared ${t} from ${acct}`);
        }
      }
    }
  }

  // ---- Batched flush: one setValues per touched column. ----
  if (lastRow >= 2) {
    sheet.getRange(2, CONFIG.PRICE_COL, lastRow - 1, 1).setValues(priceVals);
    for (const cStr of Object.keys(touchedCols)) {
      const c = Number(cStr);
      const k = colToK[c];
      const colVals = qtyBlock.map(qrow => [qrow[k]]);
      sheet.getRange(2, c, lastRow - 1, 1).setValues(colVals);
    }
  }

  applyRule1And2(sheet, log);
  applyCostBasisFormulas(sheet);
  recalculateGainLoss();
  recalculateTotals(sheet);

  // Format Columns D and E to whole numbers
  formatGainLossColumns(sheet);

  const msg = log.join('\n') + (errors.length ? '\n\nWarnings:\n' + errors.join('\n') : '');
  return { ok: true, message: msg, newTickers: Array.from(newTickersMap.values()) };
}

// =============================================================================
// BULK ADD NEW POSITIONS
// =============================================================================
function showBulkAddDialog(tickersJson) {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return;
  const tickers = JSON.parse(tickersJson);
  const sections = getSectionRows(sheet).map(s => s.name);
  const accounts = Object.keys(getAccountColumns(sheet));
  if (!sections.length) {
    SpreadsheetApp.getUi().alert('No sections on this tab yet. Use Holdings → Add Section… to create one first.');
    return;
  }
  const sectionOpts = sections.map(s => `<option>${s}</option>`).join('');

  // Create mapping of historical base tickers to their sections
  const tickerMap = getTickerSectionMap(sheet);

  const LIQUID_TICKERS = ['CASH', 'BIL', 'SHV', 'CORE', 'FDRXX', 'SPAXX', 'FZFXX', 'SGOV'];
  const isLiquidBase = base => LIQUID_TICKERS.includes(base) || base.startsWith('CORE') || base.startsWith('FDRXX');

  // Pre-resolve each ticker: liquid → 'Liquid', known → historical section,
  // otherwise queue it for ONE batched Claude call (not one call per ticker).
  const preResolved = tickers.map(t => {
    const base = getBaseTicker(t.label);
    if (isLiquidBase(base)) return 'Liquid';
    if (tickerMap[base]) return tickerMap[base];
    return null;
  });
  const needClaude = [...new Set(
    tickers.map((t, i) => (preResolved[i] === null ? getBaseTicker(t.label) : null)).filter(Boolean)
  )];
  const claudeMap = getCategoriesFromClaude(needClaude, sections);

  const rows = tickers.map((t, i) => {
    const displayQty = t.isOption ? t.qty / 100 : t.qty;
    const ac = t.avgCost !== null ? t.avgCost : '';
    const rowAcctOpts = accounts.map(a => `<option${a === t.account ? ' selected' : ''}>${a}</option>`).join('');

    // Auto-select category: historical memory → batched Claude suggestion → fallback.
    const base = getBaseTicker(t.label);
    let defaultSec = preResolved[i] || claudeMap[base] || 'Small Themes/Other';
    if (!sections.includes(defaultSec)) defaultSec = sections[0];
    const rowSecOpts = sections.map(s => `<option${s === defaultSec ? ' selected' : ''}>${s}</option>`).join('');

    return `<tr style="border-bottom:1px solid #eee;">
      <td style="padding:5px 6px; font-family:monospace; font-size:12px; font-weight:600;">${t.label}</td>
      <td style="padding:5px 4px;"><select id="acct${i}" style="width:100%; font-size:11px; padding:2px;">${rowAcctOpts}</select></td>
      <td style="padding:5px 4px;"><select id="sec${i}" onchange="syncSection('${base}', this.value)" style="width:100%; font-size:11px; padding:2px;">${rowSecOpts}</select></td>
      <td style="padding:5px 6px; text-align:right; font-size:11px; color:#5f6368;">${displayQty}</td>
      <td style="padding:5px 6px; text-align:right; font-size:11px; color:#5f6368;">$${t.price}</td>
      <td style="padding:5px 4px; text-align:right;"><input id="ac${i}" type="number" step="0.0001" value="${ac}" style="width:70px; font-size:11px; text-align:right;" placeholder="avg cost"></td>
    </tr>`;
  }).join('');

  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family:-apple-system,system-ui,sans-serif; margin:12px; font-size:13px; }
  h3 { margin:0 0 10px; font-size:15px; }
  table { border-collapse:collapse; width:100%; margin:8px 0; }
  th { background:#f8f9fa; padding:5px 6px; text-align:left; font-size:11px; border-bottom:2px solid #e0e0e0; }
  .toprow { display:flex; gap:10px; align-items:center; margin-bottom:10px; flex-wrap:wrap; }
  .toprow label { font-size:12px; font-weight:600; }
  select.big { padding:5px; font-size:12px; border-radius:4px; }
  button { padding:9px 18px; font-size:13px; cursor:pointer; border-radius:4px; font-weight:500; }
  .primary { background:#1a73e8; color:white; border:none; }
  .primary:hover { background:#1557b0; }
  .secondary { background:white; border:1px solid #ccc; }
  #banner { display:none; margin-top:10px; padding:8px 12px; border-radius:5px; background:#e6f4ea; color:#137333; font-weight:600; }
  #banner.err { background:#fce8e6; color:#d93025; }
  #status { margin-top:8px; font-size:11px; font-family:monospace; max-height:80px; overflow-y:auto; white-space:pre-wrap; }
</style>
<h3>Add New Positions (${tickers.length})</h3>
<div class="toprow">
  <label>Set all sections to:</label>
  <select id="bulkSec" class="big">${sectionOpts}</select>
  <button class="secondary" onclick="applyAll()" style="padding:5px 10px;font-size:12px;">Apply to all</button>
</div>
<div style="max-height:320px; overflow-y:auto;">
<table>
  <tr>
    <th>Ticker</th><th>Account</th><th>Section</th><th style="text-align:right;">Qty</th>
    <th style="text-align:right;">Price</th><th style="text-align:right;">Avg Cost</th>
  </tr>
  ${rows}
</table>
</div>
<div style="margin-top:12px; display:flex; gap:8px;">
  <button class="primary" id="addBtn" onclick="addAll()">Add All Positions</button>
  <button class="secondary" onclick="google.script.host.close()">Cancel</button>
</div>
<div id="banner"></div>
<div id="status"></div>
<script>
  var TICKERS = ${JSON.stringify(tickers)};

  function applyAll() {
    var sec = document.getElementById('bulkSec').value;
    TICKERS.forEach(function(t, i) { document.getElementById('sec' + i).value = sec; });
  }

  function syncSection(baseTicker, newValue) {
    TICKERS.forEach(function(t, i) {
      var base = t.label.trim().split(/\\s+/)[0].toUpperCase();
      if (base === baseTicker) {
        document.getElementById('sec' + i).value = newValue;
      }
    });
  }

  function addAll() {
    var btn = document.getElementById('addBtn');
    btn.disabled = true; btn.textContent = 'Adding…';
    var items = TICKERS.map(function(t, i) {
      return {
        label: t.label, qty: t.qty, price: t.price,
        isOption: t.isOption,
        avgCost: document.getElementById('ac' + i).value === '' ? null : parseFloat(document.getElementById('ac' + i).value),
        section: document.getElementById('sec' + i).value,
        account: document.getElementById('acct' + i).value
      };
    });
    google.script.run
      .withSuccessHandler(function(r) {
        btn.disabled = false; btn.textContent = 'Add All Positions';
        var b = document.getElementById('banner');
        b.style.display = 'block';
        b.className = r.ok ? '' : 'err';
        b.textContent = r.ok ? '✓ Done! ' + r.added + ' position(s) added.' : '✗ ' + r.message;
        document.getElementById('status').textContent = r.detail || '';
        if (r.ok) setTimeout(function() { google.script.host.close(); }, 2000);
      })
      .withFailureHandler(function(e) {
        btn.disabled = false; btn.textContent = 'Add All Positions';
        var b = document.getElementById('banner');
        b.style.display = 'block'; b.className = 'err';
        b.textContent = '✗ ' + e.message;
      })
      .applyBulkAdd(JSON.stringify(items));
  }
<\/script>
  `).setWidth(700).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add New Positions');
}

function applyBulkAdd(itemsJson) {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return { ok: false, message: 'Not on a holdings tab' };
  const items = JSON.parse(itemsJson);
  const accounts = getAccountColumns(sheet);
  let added = 0;
  const detail = [];

  for (const item of items) {
    const label = canonicalLabel(item.label);
    const posValue = item.price * item.qty;
    if (!item.isOption && posValue < CONFIG.MIN_POSITION_VALUE) {
      detail.push(`skip ${label}: value $${posValue.toFixed(2)} < $${CONFIG.MIN_POSITION_VALUE}`);
      continue;
    }

    // FETCH SECTIONS INSIDE THE LOOP to prevent indexing bugs when multiple rows are added above a category
    const currentSections = getSectionRows(sheet);
    const section = currentSections.find(s => s.name === item.section);
    if (!section) { detail.push(`skip ${label}: section "${item.section}" not found`); continue; }

    const insertAfter = getSectionEndRow(sheet, section.row);
    sheet.insertRowAfter(insertAfter);
    const newRow = insertAfter + 1;

    sheet.getRange(newRow, 1, 1, sheet.getLastColumn())
      .setFontColor('#000000')
      .setBackground(null)
      .setFontWeight('normal');

    sheet.getRange(newRow, CONFIG.TICKER_COL).setValue(label);
    sheet.getRange(newRow, CONFIG.PRICE_COL).setValue(item.price);
    const col = accounts[item.account];
    if (col) sheet.getRange(newRow, col).setValue(item.qty);
    if (item.avgCost !== null && !isNaN(item.avgCost)) {
      setHelperAvgCost(sheet, label, item.account, item.avgCost);
    }
    added++;
    detail.push(`✓ [${item.account}] ${label} → ${item.section}`);
  }

  applyCostBasisFormulas(sheet);
  recalculateGainLoss();
  recalculateTotals(sheet);
  formatGainLossColumns(sheet);
  return { ok: true, added, detail: detail.join('\n') };
}

// =============================================================================
// ADD POSITION
// =============================================================================
function showAddPositionDialog() {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return;
  const accounts = Object.keys(getAccountColumns(sheet));
  const sections = getSectionRows(sheet).map(s => s.name);
  const tickerMap = getTickerSectionMap(sheet);

  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 12px; font-size: 13px; }
  label { display: block; margin: 8px 0 4px; font-weight: 600; }
  input, select { width: 100%; box-sizing: border-box; padding: 5px; font-size: 13px; }
  button { margin-top: 12px; padding: 8px 16px; font-size: 13px; cursor: pointer; border-radius: 4px; }
  .primary { background: #1a73e8; color: white; border: none; }
  .secondary { background: white; border: 1px solid #ccc; }
  .row { display: flex; gap: 8px; } .row > * { flex: 1; }
  .hint { color: #666; font-size: 11px; margin-top: 2px; }
</style>
<label>Ticker / option label</label>
<input id="tk" placeholder="AAPL  or  ET 1/15/27 30C  or  322310" />
<label>Section</label>
<select id="sec">${sections.map(s => `<option>${s}</option>`).join('')}</select>
<label>Current price USD (Col B)</label>
<input id="px" type="number" step="0.0001" />
<label>Average cost USD — leave blank if unknown</label>
<input id="avg" type="number" step="0.0001" />
<label>Initial account & quantity</label>
<div class="row">
  <select id="acct">${accounts.map(a => `<option>${a}</option>`).join('')}</select>
  <input id="qty" type="number" step="1" placeholder="qty" />
</div>
<div class="hint">For options, enter contracts — multiplied by 100 on write.</div>
<button class="primary" onclick="run()">Add</button>
<button class="secondary" onclick="google.script.host.close()">Cancel</button>
<div id="status" style="margin-top:12px; font-family:monospace; font-size:11px;"></div>
<script>
  var MAP = ${JSON.stringify(tickerMap)};
  document.getElementById('tk').addEventListener('input', function(e) {
     var base = e.target.value.trim().split(/\\s+/)[0].toUpperCase();
     if (MAP[base]) {
        document.getElementById('sec').value = MAP[base];
     }
  });

  function run() {
    var v = {
      ticker: document.getElementById('tk').value,
      section: document.getElementById('sec').value,
      price: parseFloat(document.getElementById('px').value),
      avgCost: document.getElementById('avg').value === '' ? null : parseFloat(document.getElementById('avg').value),
      account: document.getElementById('acct').value,
      qty: parseFloat(document.getElementById('qty').value)
    };
    google.script.run
      .withSuccessHandler(function(r) {
        document.getElementById('status').textContent = (r.ok ? '✓ ' : '✗ ') + r.message;
        if (r.ok) setTimeout(function() { google.script.host.close(); }, 900);
      })
      .applyAddPosition(v);
  }
<\/script>
  `).setWidth(420).setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Position');
}

function applyAddPosition(v) {
  const sheet = getActiveHoldingsSheet();
  if (!sheet) return { ok: false, message: 'Not on a holdings tab' };
  const label = canonicalLabel(v.ticker);
  const isOpt = isOptionLabel(label);
  const storedQty = isOpt ? v.qty * 100 : v.qty;
  if (!v.price || isNaN(v.price) || !storedQty || isNaN(storedQty)) {
    return { ok: false, message: 'Enter a valid price and quantity.' };
  }
  // Options bypass the $1,000 floor (same as the batch import path) — a deep
  // OTM option's market value is irrelevant; what matters is you own it.
  if (!isOpt && v.price * storedQty < CONFIG.MIN_POSITION_VALUE) {
    return { ok: false, message: `Position value $${(v.price*storedQty).toFixed(2)} < $${CONFIG.MIN_POSITION_VALUE} (Rule 5).` };
  }

  // Merge into an existing row instead of creating a duplicate ticker row.
  const existingRow = findTickerRow(sheet, label);
  if (existingRow !== -1) {
    const accounts = getAccountColumns(sheet);
    const col = accounts[v.account];
    sheet.getRange(existingRow, CONFIG.PRICE_COL).setValue(v.price);
    if (col) {
      const cur = sheet.getRange(existingRow, col).getValue();
      sheet.getRange(existingRow, col).setValue((typeof cur === 'number' ? cur : 0) + storedQty);
    }
    let costNote = '';
    if (v.avgCost !== null && !isNaN(v.avgCost)) {
      if (getHelperAvgCost(sheet, label, v.account) === null) {
        setHelperAvgCost(sheet, label, v.account, v.avgCost);
      } else {
        costNote = ' (kept existing avg cost)';
      }
    }
    applyCostBasisFormulas(sheet);
    recalculateGainLoss();
    recalculateTotals(sheet);
    formatGainLossColumns(sheet);
    return { ok: true, message: `Merged into existing ${label} at row ${existingRow} — qty added to ${v.account}${costNote}.` };
  }

  const currentSections = getSectionRows(sheet);
  const section = currentSections.find(s => s.name === v.section);
  if (!section) return { ok: false, message: `Section "${v.section}" not found.` };
  const insertAfter = getSectionEndRow(sheet, section.row);
  sheet.insertRowAfter(insertAfter);
  const newRow = insertAfter + 1;

  sheet.getRange(newRow, 1, 1, sheet.getLastColumn())
    .setFontColor('#000000')
    .setBackground(null)
    .setFontWeight('normal');

  sheet.getRange(newRow, CONFIG.TICKER_COL).setValue(label);
  sheet.getRange(newRow, CONFIG.PRICE_COL).setValue(v.price);
  const accounts = getAccountColumns(sheet);
  const col = accounts[v.account];
  if (col) sheet.getRange(newRow, col).setValue(storedQty);
  if (v.avgCost !== null && !isNaN(v.avgCost)) {
    setHelperAvgCost(sheet, label, v.account, v.avgCost);
  }
  applyCostBasisFormulas(sheet);
  recalculateGainLoss();
  recalculateTotals(sheet);
  formatGainLossColumns(sheet);
  return { ok: true, message: `Added ${label} under ${v.section} at row ${newRow}.` };
}


// =============================================================================
// SECTIONS / ACCOUNTS
// =============================================================================
function showAddSectionDialog() {
  const sheet = getActiveHoldingsSheet(); if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Add Section', 'Section name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = r.getResponseText().trim(); if (!name) return;
  const color = CONFIG.DEFAULT_SECTIONS[name] || '#dddddd';
  const lastRow = sheet.getLastRow();
  sheet.insertRowAfter(lastRow);
  const newRow = lastRow + 1;
  const lastCol = Math.max(sheet.getLastColumn(), CONFIG.FIRST_ACCOUNT_COL);
  const range = sheet.getRange(newRow, 1, 1, lastCol);
  range.setBackground(color); range.setFontWeight('bold');
  sheet.getRange(newRow, CONFIG.TICKER_COL).setValue(name);
  ui.alert(`Section "${name}" added at row ${newRow}.`);
}

function showAddAccountDialog() {
  const sheet = getActiveHoldingsSheet(); if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Add Account Column', 'Account name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = r.getResponseText().trim(); if (!name) return;

  const accts = getAccountColumns(sheet);
  const acctCols = Object.values(accts);
  const lastAcctCol = acctCols.length > 0 ? Math.max(...acctCols) : CONFIG.FIRST_ACCOUNT_COL - 1;
  const newCol = lastAcctCol + 1;

  sheet.insertColumnAfter(lastAcctCol);
  sheet.getRange(CONFIG.HEADER_ROW, newCol).setValue(name).setFontWeight('bold');

  const helper = getOrCreateHelperSheet(sheet);
  helper.insertColumnAfter(lastAcctCol);
  helper.getRange(CONFIG.HEADER_ROW, newCol).setValue(name).setFontWeight('bold');

  applyCostBasisFormulas(sheet);
  recalculateTotals(sheet);
  ui.alert(`Account "${name}" safely added as column ${newCol}.`);
}

// =============================================================================
// CLEANUP + RECALC
// =============================================================================
function applyCleanupRules() {
  const sheet = getActiveHoldingsSheet(); if (!sheet) return;
  const log = [];
  applyRule1And2(sheet, log);
  applyCostBasisFormulas(sheet);
  recalculateGainLoss();
  recalculateTotals(sheet);
  SpreadsheetApp.getUi().alert(log.length ? log.join('\n') : 'Nothing to clean up.');
}

function applyRule1And2(sheet, log) {
  log = log || [];
  const today = new Date(); today.setHours(0,0,0,0);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // ---- Batched reads: labels, prices, and the whole account qty block. ----
  const colA = sheet.getRange(2, CONFIG.TICKER_COL, lastRow - 1, 1).getValues();
  const colB = sheet.getRange(2, CONFIG.PRICE_COL, lastRow - 1, 1).getValues();
  const accounts = getAccountColumns(sheet);
  const acctCols = Object.values(accounts).sort((a, b) => a - b);
  const colToK = {};
  acctCols.forEach((c, i) => { colToK[c] = i; });
  let qtyBlock = [];
  if (acctCols.length > 0) {
    const span = acctCols[acctCols.length - 1] - acctCols[0] + 1;
    qtyBlock = sheet.getRange(2, acctCols[0], lastRow - 1, span).getValues();
  }

  // Same filter as getAllPositionRows: ticker present AND price present.
  const rows = [];
  for (let i = 0; i < colA.length; i++) {
    const t = String(colA[i][0] || '').trim();
    const p = colB[i][0];
    if (!t || p === '' || p === null) continue;
    rows.push({ row: i + 2, ticker: t, price: p });
  }

  const toDelete = [];
  const tickersToRemove = [];

  for (const r of rows) {
    if (isOptionLabel(r.ticker)) {
      const m = r.ticker.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
      if (m) {
        const yr = parseInt(m[3], 10) < 100 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
        const exp = new Date(yr, parseInt(m[1], 10) - 1, parseInt(m[2], 10));
        if (exp < today) {
          toDelete.push(r.row); tickersToRemove.push(r.ticker);
          log.push(`Rule 2: expired ${r.ticker}`); continue;
        }
      }
    }
  }

  // Rule 1 needs account quantities to compute value. With zero account
  // columns the value is unknowable — skip instead of deleting everything
  // (the old code computed value = price × 0 and removed all priced rows).
  if (acctCols.length > 0) {
    for (const r of rows) {
      if (toDelete.indexOf(r.row) !== -1) continue;
      const price = r.price;
      if (price === '' || price === null || isNaN(price)) continue;
      let totalQty = 0;
      for (const c of acctCols) {
        const v = qtyBlock[r.row - 2][colToK[c]];
        if (typeof v === 'number') totalQty += v;
      }
      const value = price * totalQty;
      if (value < CONFIG.MIN_POSITION_VALUE) {
        toDelete.push(r.row); tickersToRemove.push(r.ticker);
        log.push(`Rule 1: ${r.ticker} value $${value.toFixed(2)} — removed`);
      }
    }
  }
  const pairs = toDelete.map((row, i) => ({ row, ticker: tickersToRemove[i] }));
  pairs.sort((a, b) => b.row - a.row);
  for (const p of pairs) {
    sheet.deleteRow(p.row);
    removeFromHelper(sheet, p.ticker);
  }

  // --- SMART BLANK ROW CLEANUP ---
  // Fresh row count: rows were deleted above.
  const sweepLast = sheet.getLastRow();
  if (sweepLast >= 2) {
    const data = sheet.getRange(1, 1, sweepLast, 2).getValues();
    const bgs = sheet.getRange(1, 1, sweepLast, 1).getBackgrounds();
    let deletedCount = 0;

    // Tracks what is immediately beneath the row we are checking
    let nextRowType = 'EOF';

    // Sweep from bottom to top (prevents row index shifting issues)
    for (let i = sweepLast - 1; i >= 1; i--) {
      const rowNum = i + 1;
      const ticker = String(data[i][0]).trim();
      const price = data[i][1];
      const bg = bgs[i][0] ? bgs[i][0].toLowerCase() : '#ffffff';

      const isHeader = (ticker !== '' && (price === '' || price === null) && bg !== '#ffffff' && bg !== '');
      const isBlank = (ticker === '' && (bg === '#ffffff' || bg === ''));

      if (isHeader) {
        nextRowType = 'HEADER';
      } else if (!isBlank) {
        nextRowType = 'TICKER';
      } else {
        // It IS a blank row. Decide whether to keep or delete.
        if (nextRowType === 'HEADER') {
          // Keep exactly ONE blank row directly above a header
          nextRowType = 'BLANK_ABOVE_HEADER';
        } else {
          // Delete if it's trapped between tickers, or stacked above another blank row
          sheet.deleteRow(rowNum);
          deletedCount++;
        }
      }
    }
    if (deletedCount > 0) {
      log.push(`• Cleaned up ${deletedCount} excess empty row(s).`);
    }
  }
}

function recalculateGainLoss() {
  const sheet = getActiveHoldingsSheet(); if (!sheet) return;
  const rows = getAllPositionRows(sheet);
  const accounts = getAccountColumns(sheet);
  const acctCols = Object.values(accounts).sort((a, b) => a - b);
  if (acctCols.length === 0) return;
  const firstA1 = columnToLetter(acctCols[0]);
  const lastA1  = columnToLetter(acctCols[acctCols.length - 1]);
  for (const r of rows) {
    const row = r.row;
    const dCell = sheet.getRange(row, CONFIG.GAIN_DOLLAR_COL);
    const eCell = sheet.getRange(row, CONFIG.GAIN_PCT_COL);
    dCell.setFormula(`=IFERROR(IF(C${row}="","",(B${row}-C${row})*SUM(${firstA1}${row}:${lastA1}${row})),"")`);
    eCell.setFormula(`=IFERROR(IF(C${row}="","",(B${row}-C${row})/C${row}),"")`).setNumberFormat('0.00%');
  }
}

function recalculateAllFormulas() {
  const sheet = getActiveHoldingsSheet(); if (!sheet) return;
  getOrCreateHelperSheet(sheet);
  applyCostBasisFormulas(sheet);
  recalculateGainLoss();
  recalculateTotals(sheet);
  SpreadsheetApp.getUi().alert('✓ Formulas refreshed.\n• Col C: cost basis\n• Col D/E: gains\n• New dynamic totals + category roll-ups built!');
}

// =============================================================================
// NEW MONTH
// =============================================================================
function startNewMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cur = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const m = cur.getName().match(/^([A-Z][a-z]{2})(\d{2})$/i);
  if (!m) { ui.alert('Switch to the current month tab first (tab name must be MMMyy, e.g. "Jan26").'); return; }
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  let idx = months.findIndex(x => x.toLowerCase() === m[1].toLowerCase());
  let yr = parseInt(m[2], 10);
  idx += 1; if (idx > 11) { idx = 0; yr += 1; }
  const newName = months[idx] + (yr < 10 ? '0' + yr : yr);
  if (ss.getSheetByName(newName)) { ui.alert(`Tab "${newName}" already exists.`); return; }

  const copy = cur.copyTo(ss);
  copy.setName(newName);

  const oldHelperName = helperSheetName(cur.getName());
  const oldHelper = ss.getSheetByName(oldHelperName);
  if (oldHelper) {
    const newHelperName = helperSheetName(newName);
    const helperCopy = oldHelper.copyTo(ss);
    helperCopy.setName(newHelperName);
    helperCopy.hideSheet();
  }

  ss.setActiveSheet(copy);
  applyCostBasisFormulas(copy);
  recalculateGainLoss();
  recalculateTotals(copy);
  ui.alert(`Created "${newName}" and its helper sheet.`);
}

function formatGainLossColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Format Column D (Gain/Loss $) to whole dollars
  sheet.getRange(2, CONFIG.GAIN_DOLLAR_COL, lastRow - 1).setNumberFormat("$#,##0;-$#,##0");

  // Format Column E (Gain/Loss %) to whole percentages
  sheet.getRange(2, CONFIG.GAIN_PCT_COL, lastRow - 1).setNumberFormat("0%;-0%");
}

// Categorize MANY tickers in ONE API call (returns {TICKER: category}).
// The old per-ticker version fired a synchronous UrlFetchApp per unknown
// ticker while the Bulk Add dialog rendered — hanging or timing out on
// large imports. Unknown tickers are now resolved up front in a single call.
function getCategoriesFromClaude(tickers, existingCategories) {
  const out = {};
  const uniq = [...new Set((tickers || []).map(t => String(t).toUpperCase()))];
  if (!uniq.length) return out;

  const apiKey = PropertiesService.getScriptProperties().getProperty(CONFIG.API_KEY_PROPERTY);
  if (!apiKey) {
    uniq.forEach(t => { out[t] = 'Small Themes/Other'; });
    return out;
  }

  const prompt = `You are a financial portfolio categorization assistant.
Here are the ONLY allowed categories: [${existingCategories.join(', ')}].

For EACH ticker below, pick the single best-fitting category from that list, based on the company's primary industry and operations.

Tickers: [${uniq.join(', ')}]

Rules:
1. Reply with ONLY a JSON object mapping each ticker to exactly one category name from the list, e.g. {"AAPL":"Technology","XOM":"Energy"}.
2. If a ticker does not clearly fit a specific sector, map it to "Small Themes/Other".
3. No markdown, no commentary — just the JSON object.`;

  try {
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: getApiModel(),
        max_tokens: 1024,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(resp.getContentText());
    const txt = (json.content || []).filter(c => c.type === 'text').map(c => c.text).join('\n');
    const m = txt.match(/\{[\s\S]*\}/);
    const map = m ? JSON.parse(m[0]) : {};
    uniq.forEach(t => {
      const cat = map[t];
      out[t] = existingCategories.includes(cat) ? cat : 'Small Themes/Other';
    });
  } catch (e) {
    Logger.log('Claude batch categorization failed: ' + e);
    uniq.forEach(t => { if (!out[t]) out[t] = 'Small Themes/Other'; });
  }
  return out;
}

// Backward-compatible single-ticker wrapper.
function getCategoryFromClaude(ticker, existingCategories) {
  return getCategoriesFromClaude([ticker], existingCategories)[String(ticker).toUpperCase()]
    || 'Small Themes/Other';
}
