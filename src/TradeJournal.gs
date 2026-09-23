/**
 * TRADE JOURNAL — companion module to the Holdings Tracker (v7)
 *
 * WHAT THIS IS
 *   Holdings Tracker (v7) tracks open positions and UNREALIZED gain/loss.
 *   This module adds REALIZED gain/loss: a running log of CLOSED trades
 *   (specific-lot basis) with a thesis/exit-notes journal, fed by pasting
 *   the same broker "Realized Gain/Loss" / "Closed Positions" exports or
 *   screenshots you already use for Holdings Tracker — same places, same
 *   brokerages, parsed by the same Claude API key.
 *
 * SETUP (one-time, ~2 minutes)
 *   1. Open Extensions → Apps Script on your portfolio spreadsheet.
 *   2. Click the "+" next to Files → Script → name it "TradeJournal" →
 *      paste this entire file in → Save.
 *   3. Open your existing Code.gs (the v7 file) and find the onOpen()
 *      function. Add ONE line right after the Holdings menu's .addToUi():
 *
 *        function onOpen() {
 *          const ui = SpreadsheetApp.getUi();
 *          ui.createMenu('📊 Holdings')
 *            ...
 *            .addToUi();
 *          addTradeJournalMenu(ui);   // <-- ADD THIS LINE
 *        }
 *
 *   4. Reload the spreadsheet. You'll see a new "📉 Trade Journal" menu.
 *   5. Your existing "Set Anthropic API Key" (Holdings → Settings) is
 *      reused automatically — nothing else to configure.
 *
 * DEPENDENCIES (reused from the v7 file — this file defines none of these)
 *   CONFIG.API_MODEL, CONFIG.API_KEY_PROPERTY, getApiKey(), columnToLetter(),
 *   canonicalLabel(), isOptionLabel(), isHoldingsTab(), getAccountColumns()
 *
 * DESIGN NOTES
 *   • Cost-basis method: specific-lot, with ONE exception — lots of the same
 *     ticker that share the SAME open date AND the SAME close date are merged
 *     into one row (summed qty, qty-weighted average prices). Lots opened or
 *     closed on different dates — even a day apart — stay separate rows.
 *     No FIFO/average-cost re-matching beyond that.
 *   • Duplicate-safe re-pastes: on save, any trade whose ticker + open date +
 *     close date + qty already exists in the journal is skipped, so pasting a
 *     broker export that overlaps what you already logged can't double-count.
 *     The paste dialog shows the last journaled close date, so you only need
 *     to capture trades closed on/after that date (inclusive — the overlap
 *     day is safe because of the dedupe).
 *   • "3 months" isn't hardcoded anywhere — Trade Summary lets you pick the
 *     lookback window each time you run it (30/90/180/365 days, YTD, all-time),
 *     so how far back you report is a reporting-time choice, not a fixed rule.
 *   • Input methods: paste text, paste/drop a screenshot, or DRAG-AND-DROP a
 *     .csv / .txt / .tsv file straight onto the paste box — the file's text
 *     is read in the browser and dropped into the box as if you'd pasted it.
 *   • Out of scope for v1 (easy to add later if useful): wash-sale flagging,
 *     Section/category tagging reused from Holdings, screenshot reference
 *     links. Skipped for now to keep the first version tight.
 */

// =============================================================================
// CONFIG
// =============================================================================
const JCONFIG = {
  SHEET_NAME: 'Trade Journal',
  SUMMARY_SHEET_NAME: 'Journal Summary',
  HEADER_ROW: 1,
  COL: {
    ID: 1, TICKER: 2, ACCOUNT: 3, OPEN_DATE: 4, CLOSE_DATE: 5, QTY: 6,
    OPEN_PRICE: 7, CLOSE_PRICE: 8, COST_BASIS: 9, PROCEEDS: 10,
    PNL_DOLLAR: 11, PNL_PCT: 12, HOLD_DAYS: 13, TERM: 14,
    THESIS: 15, EXIT_NOTES: 16
  },
  LONG_TERM_DAYS: 365
};

// =============================================================================
// MENU — call addTradeJournalMenu(ui) from your existing onOpen() (see header)
// =============================================================================
function addTradeJournalMenu(ui) {
  ui.createMenu('📉 Trade Journal')
    .addItem('Log Closed Trades…', 'showLogTradesDialog')
    .addItem('Add Thesis / Exit Notes (AI)…', 'showNotesDialog')
    .addItem('Trade Summary…', 'showJournalSummaryDialog')
    .addSeparator()
    .addItem('Recalc Journal Formulas', 'recalcJournalFormulas')
    .addToUi();
}

// =============================================================================
// SHEET HELPERS
// =============================================================================
function getOrCreateJournalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(JCONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(JCONFIG.SHEET_NAME);
    const headers = [
      'ID', 'Ticker', 'Account', 'Open Date', 'Close Date', 'Qty',
      'Open Price', 'Close Price', 'Cost Basis $', 'Proceeds $',
      'Realized P&L $', 'Realized P&L %', 'Hold Days', 'Term',
      'Thesis (entry reason)', 'Exit Notes / Lesson'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(JCONFIG.COL.THESIS, 260);
    sheet.setColumnWidth(JCONFIG.COL.EXIT_NOTES, 260);
  }
  return sheet;
}

function nextJournalId(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 'J0001';
  const ids = sheet.getRange(2, JCONFIG.COL.ID, last - 1, 1).getValues()
    .map(r => String(r[0] || '')).filter(String);
  let max = 0;
  ids.forEach(id => {
    const n = parseInt(id.replace(/\D/g, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return 'J' + String(max + 1).padStart(4, '0');
}

// Parse a date string as LOCAL time. A bare "YYYY-MM-DD" fed to new Date()
// is treated as UTC midnight, which displays as the PREVIOUS day once it
// lands in a sheet whose timezone is west of UTC (all US timezones).
function parseLocalDate_(s) {
  if (s instanceof Date) return isNaN(s.getTime()) ? '' : s;
  const m = String(s || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : d;
}

// Normalize a cell value or input string to "yyyy-MM-dd" for dedupe keys,
// so a Date object read from the sheet and the string the dialog submits
// compare equal.
function journalDateKey_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const m = String(v || '').trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  const d = parseLocalDate_(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || '');
}

// Latest Close Date in the journal — the watermark for "screenshot from here".
function getLastJournaledClose_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JCONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return '';
  let max = null;
  sheet.getRange(2, JCONFIG.COL.CLOSE_DATE, sheet.getLastRow() - 1, 1).getValues()
    .forEach(r => {
      const d = r[0];
      if (d instanceof Date && !isNaN(d.getTime()) && (!max || d > max)) max = d;
    });
  return max ? Utilities.formatDate(max, Session.getScriptTimeZone(), 'M/d/yyyy') : '';
}

function writeJournalFormulas(sheet, row) {
  const c = JCONFIG.COL;
  const L = columnToLetter; // reused from Holdings v7 file

  sheet.getRange(row, c.COST_BASIS)
    .setFormula(`=IFERROR(${L(c.QTY)}${row}*${L(c.OPEN_PRICE)}${row},"")`)
    .setNumberFormat('$#,##0.00');

  sheet.getRange(row, c.PROCEEDS)
    .setFormula(`=IFERROR(${L(c.QTY)}${row}*${L(c.CLOSE_PRICE)}${row},"")`)
    .setNumberFormat('$#,##0.00');

  sheet.getRange(row, c.PNL_DOLLAR)
    .setFormula(`=IFERROR(${L(c.PROCEEDS)}${row}-${L(c.COST_BASIS)}${row},"")`)
    .setNumberFormat('$#,##0;-$#,##0');

  sheet.getRange(row, c.PNL_PCT)
    .setFormula(`=IFERROR(${L(c.PNL_DOLLAR)}${row}/${L(c.COST_BASIS)}${row},"")`)
    .setNumberFormat('0.00%');

  sheet.getRange(row, c.HOLD_DAYS)
    .setFormula(`=IFERROR(${L(c.CLOSE_DATE)}${row}-${L(c.OPEN_DATE)}${row},"")`)
    .setNumberFormat('0');

  sheet.getRange(row, c.TERM)
    .setFormula(`=IFERROR(IF(${L(c.HOLD_DAYS)}${row}>=${JCONFIG.LONG_TERM_DAYS},"Long","Short"),"")`);

  sheet.getRange(row, c.OPEN_DATE, 1, 2).setNumberFormat('M/d/yyyy');
}

function recalcJournalFormulas() {
  const sheet = getOrCreateJournalSheet();
  const last = sheet.getLastRow();
  for (let row = 2; row <= last; row++) writeJournalFormulas(sheet, row);
  SpreadsheetApp.getUi().alert('✓ Journal formulas refreshed for ' + Math.max(0, last - 1) + ' row(s).');
}

function getKnownAccounts_() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (typeof isHoldingsTab === 'function' && isHoldingsTab(sheet.getName())) {
      const accts = Object.keys(getAccountColumns(sheet));
      if (accts.length) return accts;
    }
  } catch (e) {}
  return ['Unknown'];
}

// =============================================================================
// AI PARSING — independent of v7's callClaudeAPI so v7 stays untouched
// =============================================================================
function buildJournalSystemPrompt() {
  return `You parse REALIZED / CLOSED trade data from brokerage "Realized Gain/Loss" or "Closed Positions" reports (CSV, pasted text, or screenshots) and return ONLY a JSON array, no markdown, no commentary.

Each item: {"ticker": "...", "account": "...", "open_date": "YYYY-MM-DD", "close_date": "YYYY-MM-DD", "qty": number, "open_price": number, "close_price": number}

RULES:
- ticker: uppercase base ticker ("AAPL"). For options use "TICKER M/D/YY STRIKE[C|P]" (no leading zeros, 2-digit year, C=Call, P=Put, uppercase).
- account: use the account hint given for that data block unless the data itself clearly states a different account name, in which case prefer the data's own label.
- qty: number of shares or option contracts closed in this lot (do not multiply option qty by 100).
- open_price / close_price: per-share (or per-contract) price, in USD. If the source only gives a TOTAL cost basis or TOTAL proceeds for the lot, divide by qty (and by 100 for options, since qty is contracts but price is per-share) to get a per-share/contract price.
- MERGE same-day lots: if multiple lots of the same ticker in the same data block share the SAME open date AND the SAME close date, combine them into ONE row — qty = summed qty, open_price and close_price = qty-weighted average prices (max 4 decimals). Do NOT merge lots whose open or close dates differ, even by one day — keep those as separate rows. Never net buys against sells or combine lots across different dates.
- Skip subtotal/summary rows, section headers, and "Total"/"Grand Total"/"Short Term Total"/"Long Term Total" rows.
- If a date is ambiguous, prefer M/D/YYYY (US) parsing. Output dates as YYYY-MM-DD.

Return [] if nothing parseable. Output JUST the JSON array.`;
}

function callClaudeAPIForJournal(messages, systemPromptOverride) {
  const apiKey = getApiKey(); // reused from Holdings v7 file
  const payload = {
    model: getApiModel(), // reused from Holdings file ( honors API_MODEL_OVERRIDE )
    max_tokens: 4096,
    system: systemPromptOverride || buildJournalSystemPrompt(),
    messages: [{ role: 'user', content: messages }]
  };
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const raw = resp.getContentText();
  let body;
  try {
    body = JSON.parse(raw);
  } catch (e) {
    throw new Error('API returned non-JSON (' + code + '): ' + raw.slice(0, 200));
  }
  if (code !== 200) {
    const msg = (body.error && body.error.message) || raw;
    throw new Error('API error ' + code + ': ' + msg);
  }
  return body.content.filter(c => c.type === 'text').map(c => c.text).join('\n');
}

function parseClosedTrades(blocks) {
  try {
    const messages = [];
    for (const b of blocks) {
      messages.push({ type: 'text', text: '--- BEGIN DATA (account hint: ' + (b.account || 'unspecified') + ') ---' });
      if (b.image) {
        messages.push({ type: 'image', source: { type: 'base64', media_type: b.image.type, data: b.image.data } });
        messages.push({ type: 'text', text: 'Parse the realized/closed trades from this screenshot.' });
      }
      if (b.text && !b.image) {
        messages.push({ type: 'text', text: b.text });
      }
      messages.push({ type: 'text', text: '--- END DATA ---' });
    }
    const txt = callClaudeAPIForJournal(messages);
    let cleaned = String(txt).replace(/```json\s*/gi, '').replace(/```\s*$/g, '').trim();
    const m = cleaned.match(/\[\s*[\s\S]*\]/);
    if (m) cleaned = m[0];
    let trades;
    try {
      trades = JSON.parse(cleaned);
    } catch (e) {
      throw new Error('AI response was not valid JSON. Raw:\n' + txt.slice(0, 400));
    }
    if (!Array.isArray(trades)) throw new Error('AI did not return a JSON array');
    return { ok: true, trades };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

// =============================================================================
// LOG CLOSED TRADES DIALOG (paste + AI parse, multi-block, screenshot + CSV drop)
// =============================================================================
function showLogTradesDialog() {
  const accounts = getKnownAccounts_();
  const lastClose = getLastJournaledClose_();

  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 12px; font-size: 13px; color: #202124; }
  label { display: block; margin: 10px 0 4px; font-weight: 600; }
  select, textarea, input { width: 100%; box-sizing: border-box; font-size: 12px; }
  textarea { height: 80px; padding: 8px; font-family: 'SF Mono', Menlo, monospace; border: 1px solid #ccc; border-radius: 4px; }
  textarea.dragover { border: 2px dashed #1a73e8; background: #e8f0fe; }
  .block-container { background: #f8f9fa; border: 1px solid #dadce0; padding: 10px; border-radius: 6px; margin-bottom: 10px; }
  .hint { color: #5f6368; font-size: 11px; margin-top: 4px; }
  .btnrow { display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap; }
  button { padding: 8px 14px; font-size: 13px; cursor: pointer; border-radius: 4px; font-weight: 500; }
  .primary { background: #1a73e8; color: white; border: none; }
  .ai { background: #f4b400; color: #000; border: none; }
  .secondary { background: white; border: 1px solid #dadce0; }
  table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 11px; }
  th { background: #f8f9fa; padding: 4px 6px; text-align: left; border-bottom: 2px solid #e0e0e0; }
  td { padding: 3px 4px; border-bottom: 1px solid #eee; }
  td input { font-size: 11px; padding: 2px 4px; }
  #status { margin-top: 10px; white-space: pre-wrap; font-family: monospace; font-size: 11px; max-height: 80px; overflow-y: auto; }
  .working { color: #1a73e8; } .err { color: #d93025; } .ok { color: #188038; }
</style>

<div id="setup-phase">
  <div id="blocks-container"></div>
  <button class="secondary" onclick="addBlock()" style="width:100%;">+ Add Another Paste (e.g. another brokerage)</button>
</div>
<div class="btnrow">
  <button class="ai" onclick="parseAll()">🤖 Parse with AI</button>
  <button class="secondary" onclick="google.script.host.close()">Cancel</button>
</div>
<div id="status">Ready. Paste your broker's export, or drag a .csv file / screenshot onto the box.${lastClose ? '\\nLast journaled close: ' + lastClose + ' — only capture trades closed on/after that date (already-journaled rows are auto-skipped on save).' : ''}</div>

<div id="review-wrap" style="display:none;">
  <div id="review-table"></div>
  <div class="btnrow">
    <button class="primary" onclick="saveAll()">Save to Journal</button>
    <button class="secondary" onclick="google.script.host.close()">Cancel</button>
  </div>
</div>

<script>
  var ACCOUNTS = ${JSON.stringify(accounts)};
  var nextId = 1;
  var pendingImages = {};
  var TRADES = [];

  function setStatus(msg, cls) {
    var el = document.getElementById('status');
    el.textContent = msg;
    el.className = cls || '';
  }

  function addBlock() {
    var id = nextId++;
    var div = document.createElement('div');
    div.className = 'block-container';
    div.id = 'block-' + id;
    var acctOpts = '';
    for (var i = 0; i < ACCOUNTS.length; i++) acctOpts += '<option>' + ACCOUNTS[i] + '</option>';
    div.innerHTML =
      '<div style="display:flex; gap:10px; margin-bottom:6px;">' +
        '<select id="acct-' + id + '" style="flex:1;">' + acctOpts + '</select>' +
        '<button class="secondary" onclick="document.getElementById(\\'block-' + id + '\\').remove()" style="padding:2px 8px;font-size:11px;">Remove</button>' +
      '</div>' +
      '<textarea id="data-' + id + '" placeholder="Paste data, or drag a .csv file or screenshot here..."></textarea>' +
      '<div id="status-' + id + '" class="hint">Ready</div>';
    document.getElementById('blocks-container').appendChild(div);
    attachEvents(id);
  }

  function loadImage(id, blob) {
    var reader = new FileReader();
    reader.onload = function(ev) {
      pendingImages[id] = { type: blob.type, data: ev.target.result.split(',')[1] };
      document.getElementById('data-' + id).value = '[Image ready - click Parse with AI]';
      document.getElementById('status-' + id).textContent = '📷 Image queued.';
    };
    reader.readAsDataURL(blob);
  }

  // Read a dropped text file (.csv/.txt/.tsv) into the textarea,
  // exactly as if its contents had been pasted.
  function loadTextFile(id, file) {
    var reader = new FileReader();
    reader.onload = function(ev) {
      delete pendingImages[id]; // a text file replaces any queued screenshot
      var ta = document.getElementById('data-' + id);
      ta.value = ev.target.result;
      var lines = String(ev.target.result).split(/\\r?\\n/).filter(function(s){ return s.trim(); }).length;
      document.getElementById('status-' + id).textContent =
        '📄 ' + file.name + ' loaded (' + lines + ' line' + (lines === 1 ? '' : 's') + ').';
    };
    reader.onerror = function() {
      document.getElementById('status-' + id).textContent = '✗ Could not read ' + file.name;
    };
    reader.readAsText(file);
  }

  // Route a dropped/pasted file to the right loader by type AND extension.
  // Extension check matters: on machines with Excel installed, .csv files often
  // report MIME type "application/vnd.ms-excel" (or empty ""), not "text/csv",
  // so a MIME-only check would silently ignore them.
  function isTextLikeFile(f) {
    var type = f.type || '';
    var name = (f.name || '').toLowerCase();
    return type.indexOf('text/') === 0 ||
           type === 'application/vnd.ms-excel' ||
           /\\.(csv|txt|tsv)$/.test(name);
  }

  function handleFile(id, f) {
    if (!f) return;
    if ((f.type || '').indexOf('image/') === 0) { loadImage(id, f); return; }
    if (isTextLikeFile(f)) { loadTextFile(id, f); return; }
    document.getElementById('status-' + id).textContent =
      '✗ Unsupported file: ' + f.name + '. Drop a .csv/.txt/.tsv or an image, or paste the text.';
  }

  function attachEvents(id) {
    var ta = document.getElementById('data-' + id);
    ta.addEventListener('paste', function(e) {
      if (!e.clipboardData) return;
      for (var i = 0; i < e.clipboardData.items.length; i++) {
        var item = e.clipboardData.items[i];
        if (item.kind === 'file') {
          var f = item.getAsFile();
          if (f) { e.preventDefault(); handleFile(id, f); return; }
        }
      }
    });
    ta.addEventListener('dragover', function(e) { e.preventDefault(); ta.classList.add('dragover'); });
    ta.addEventListener('dragleave', function() { ta.classList.remove('dragover'); });
    ta.addEventListener('drop', function(e) {
      e.preventDefault();
      ta.classList.remove('dragover');
      handleFile(id, e.dataTransfer.files[0]);
    });
    ta.addEventListener('input', function(e) {
      if (pendingImages[id] && e.target.value.indexOf('[Image ready') !== 0) {
        delete pendingImages[id];
        document.getElementById('status-' + id).textContent = 'Ready';
      }
    });
  }

  addBlock();

  function parseAll() {
    var blocks = [];
    var divs = document.querySelectorAll('.block-container');
    for (var i = 0; i < divs.length; i++) {
      var id = divs[i].id.replace('block-', '');
      var acct = document.getElementById('acct-' + id).value;
      var txt = document.getElementById('data-' + id).value;
      var img = pendingImages[id];
      if ((txt && txt.trim()) || img) blocks.push({ account: acct, text: txt, image: img });
    }
    if (blocks.length === 0) { setStatus('Paste data, drop a .csv, or add a screenshot first.', 'err'); return; }
    setStatus('🤖 Parsing… (may take a few seconds)', 'working');
    google.script.run
      .withSuccessHandler(onParsed)
      .withFailureHandler(function(e) { setStatus('✗ ' + e.message, 'err'); })
      .parseClosedTrades(blocks);
  }

  function onParsed(r) {
    if (!r.ok) { setStatus('✗ ' + r.message, 'err'); return; }
    if (!r.trades || r.trades.length === 0) { setStatus('No closed trades found in that data.', 'err'); return; }
    TRADES = r.trades;
    var html = '<div class="hint" style="margin-top:8px">Untick "Log" to skip a row. Auto-untick moves under ' +
      '<input id="minpct" type="number" value="0.5" step="0.1" style="width:50px"> % ' +
      '<button class="secondary" onclick="autoUntick()" style="padding:2px 8px;font-size:11px;">Apply</button></div>' +
      '<table><tr><th>Log</th><th>Ticker</th><th>Account</th><th>Open Date</th><th>Close Date</th>' +
      '<th>Qty</th><th>Entry $</th><th>Exit $</th><th>Move %</th><th>Thesis</th><th>Exit notes</th></tr>';
    for (var i = 0; i < TRADES.length; i++) {
      var t = TRADES[i];
      var movePct = t.open_price ? (((t.close_price - t.open_price) / t.open_price) * 100).toFixed(1) + '%' : '—';
      html += '<tr>' +
        '<td style="text-align:center"><input type="checkbox" id="keep' + i + '" checked style="width:auto"></td>' +
        '<td>' + t.ticker + '</td>' +
        '<td>' + t.account + '</td>' +
        '<td><input id="od' + i + '" value="' + (t.open_date || '') + '" style="width:80px"></td>' +
        '<td><input id="cd' + i + '" value="' + (t.close_date || '') + '" style="width:80px"></td>' +
        '<td><input id="q' + i + '" type="number" value="' + t.qty + '" style="width:50px"></td>' +
        '<td><input id="op' + i + '" type="number" step="0.0001" value="' + t.open_price + '" style="width:65px"></td>' +
        '<td><input id="cp' + i + '" type="number" step="0.0001" value="' + t.close_price + '" style="width:65px"></td>' +
        '<td>' + movePct + '</td>' +
        '<td><input id="th' + i + '" placeholder="why you entered"></td>' +
        '<td><input id="ex' + i + '" placeholder="why you exited / lesson"></td>' +
        '</tr>';
    }
    html += '</table>';
    document.getElementById('review-table').innerHTML = html;
    document.getElementById('review-wrap').style.display = 'block';
    setStatus('✓ Parsed ' + TRADES.length + ' trade(s). Check dates/prices, add notes, then save. Already-journaled rows are skipped automatically on save.', 'ok');
  }

  // Untick rows whose |entry→exit move| is below the chosen %.
  // Uses the CURRENT values in the editable price fields, so manual fixes count.
  function autoUntick() {
    var min = parseFloat(document.getElementById('minpct').value) || 0;
    var n = 0;
    for (var i = 0; i < TRADES.length; i++) {
      var op = parseFloat(document.getElementById('op' + i).value);
      var cp = parseFloat(document.getElementById('cp' + i).value);
      if (!op || isNaN(cp)) continue;
      var keep = Math.abs((cp - op) / op) * 100 >= min;
      document.getElementById('keep' + i).checked = keep;
      if (!keep) n++;
    }
    setStatus('Unticked ' + n + ' row(s) moving less than ' + min + '%. Review, then save.', 'ok');
  }

  function saveAll() {
    var items = [];
    for (var i = 0; i < TRADES.length; i++) {
      if (!document.getElementById('keep' + i).checked) continue;
      items.push({
        ticker: TRADES[i].ticker,
        account: TRADES[i].account,
        openDate: document.getElementById('od' + i).value,
        closeDate: document.getElementById('cd' + i).value,
        qty: parseFloat(document.getElementById('q' + i).value),
        openPrice: parseFloat(document.getElementById('op' + i).value),
        closePrice: parseFloat(document.getElementById('cp' + i).value),
        thesis: document.getElementById('th' + i).value,
        exitNotes: document.getElementById('ex' + i).value
      });
    }
    setStatus('Saving…', 'working');
    google.script.run
      .withSuccessHandler(function(r) {
        setStatus(r.message, r.ok ? 'ok' : 'err');
        if (r.ok) setTimeout(function() { google.script.host.close(); }, 1500);
      })
      .withFailureHandler(function(e) { setStatus('✗ ' + e.message, 'err'); })
      .applyLogTrades(JSON.stringify(items));
  }
<\/script>
  `).setWidth(720).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log Closed Trades');
}

function isBadNum_(v) {
  // Catches NaN, but also null/undefined/'' — which is what a blank numeric
  // input field becomes after JSON.stringify/JSON.parse (JSON has no NaN,
  // so it round-trips to null; plain isNaN(null) is FALSE since Number(null)
  // is 0, which would silently let blank rows through).
  return v === null || v === undefined || v === '' || isNaN(Number(v));
}

function applyLogTrades(itemsJson) {
  const items = JSON.parse(itemsJson);
  const sheet = getOrCreateJournalSheet();
  const c = JCONFIG.COL;

  // Key set of already-journaled trades (ticker + open + close + qty), so
  // re-pasting a broker report that overlaps earlier pastes can't double-log.
  const existing = new Set();
  const lastExisting = sheet.getLastRow();
  if (lastExisting >= 2) {
    sheet.getRange(2, 1, lastExisting - 1, c.QTY).getValues().forEach(r => {
      existing.add([
        canonicalLabel(String(r[c.TICKER - 1] || '')),
        journalDateKey_(r[c.OPEN_DATE - 1]),
        journalDateKey_(r[c.CLOSE_DATE - 1]),
        Number(r[c.QTY - 1])
      ].join('|'));
    });
  }

  let added = 0;
  const skipped = [];
  const dupes = [];

  for (const t of items) {
    if (!t.ticker || isBadNum_(t.qty) || isBadNum_(t.openPrice) || isBadNum_(t.closePrice)) {
      skipped.push(t.ticker || '(blank)');
      continue;
    }
    const key = [
      canonicalLabel(t.ticker),
      journalDateKey_(t.openDate),
      journalDateKey_(t.closeDate),
      Number(t.qty)
    ].join('|');
    if (existing.has(key)) { dupes.push(t.ticker); continue; }
    existing.add(key); // also catches duplicates within the same paste

    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, c.ID).setValue(nextJournalId(sheet));
    sheet.getRange(row, c.TICKER).setValue(canonicalLabel(t.ticker));
    sheet.getRange(row, c.ACCOUNT).setValue(t.account || 'Unknown');
    sheet.getRange(row, c.OPEN_DATE).setValue(t.openDate ? parseLocalDate_(t.openDate) : '');
    sheet.getRange(row, c.CLOSE_DATE).setValue(t.closeDate ? parseLocalDate_(t.closeDate) : '');
    sheet.getRange(row, c.QTY).setValue(t.qty);
    sheet.getRange(row, c.OPEN_PRICE).setValue(t.openPrice);
    sheet.getRange(row, c.CLOSE_PRICE).setValue(t.closePrice);
    sheet.getRange(row, c.THESIS).setValue(t.thesis || '');
    sheet.getRange(row, c.EXIT_NOTES).setValue(t.exitNotes || '');
    writeJournalFormulas(sheet, row);
    added++;
  }

  const msg = '✓ Logged ' + added + ' trade(s) to "' + JCONFIG.SHEET_NAME + '".' +
    (dupes.length ? ' Skipped ' + dupes.length + ' already-journaled: ' + dupes.join(', ') + '.' : '') +
    (skipped.length ? ' Skipped ' + skipped.length + ' incomplete row(s): ' + skipped.join(', ') : '');
  return { ok: added > 0 || dupes.length > 0, message: msg };
}

// =============================================================================
// AI NOTES — talk through your trades in one freeform dump (typed or dictated);
// Claude matches each remark to the right journal row and fills Thesis /
// Exit Notes. Trades you don't mention are left untouched.
// =============================================================================
function getTradesForNotes_() {
  const sheet = getOrCreateJournalSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const c = JCONFIG.COL;
  const data = sheet.getRange(2, 1, last - 1, c.EXIT_NOTES).getValues();
  const tz = Session.getScriptTimeZone();
  const fmt = d => (d instanceof Date && !isNaN(d.getTime())) ? Utilities.formatDate(d, tz, 'M/d/yyyy') : '';
  return data.map((r, i) => ({
    row: i + 2,
    id: String(r[c.ID - 1] || ''),
    ticker: String(r[c.TICKER - 1] || ''),
    account: String(r[c.ACCOUNT - 1] || ''),
    open: fmt(r[c.OPEN_DATE - 1]),
    close: fmt(r[c.CLOSE_DATE - 1]),
    pnl: Math.round(Number(r[c.PNL_DOLLAR - 1]) || 0),
    thesis: String(r[c.THESIS - 1] || ''),
    exit: String(r[c.EXIT_NOTES - 1] || '')
  })).filter(t => t.id);
}

function buildNotesSystemPrompt() {
  return 'You match a trader\'s freeform commentary (often voice-dictated, may contain transcription errors) to specific trades in their journal and turn it into journal notes.\n' +
    'Return ONLY a JSON array, no markdown: [{"id": "J0001", "thesis": "...", "exit_notes": "..."}]\n' +
    'RULES:\n' +
    '- Only include trades the commentary clearly refers to — by ticker, company name, journal ID ("J0027"), or unmistakable description (e.g. "the SpaceX play"). Omit every trade not mentioned.\n' +
    '- If the same ticker matches MULTIPLE rows (different accounts or different dates) and the commentary does not say which one, apply the same notes to EACH matching row. If the trader names an account ("in my Roth", "the Schwab one"), a date, or a journal ID, use it to pick only the right row(s).\n' +
    '- thesis = why they ENTERED. exit_notes = why they EXITED and/or the lesson learned. If the commentary only covers one, include only that key.\n' +
    '- Clean up dictation artifacts and filler words, but keep the trader\'s own voice, slang, and meaning. 1-2 sentences per field.\n' +
    '- Never invent reasoning that is not in the commentary.\n' +
    'Return [] if nothing matches.';
}

function matchNotesWithAI(freeText) {
  try {
    const trades = getTradesForNotes_();
    if (!trades.length) return { ok: false, message: 'No trades in the journal yet — log some first.' };
    const list = trades.map(t =>
      t.id + ' | ' + t.ticker + ' | account: ' + t.account + ' | opened ' + t.open + ' | closed ' + t.close + ' | P&L $' + t.pnl +
      (t.thesis ? ' | already has thesis' : '') + (t.exit ? ' | already has exit note' : '')
    ).join('\n');
    const txt = callClaudeAPIForJournal(
      [{ type: 'text', text: 'JOURNAL TRADES:\n' + list + '\n\nTRADER COMMENTARY:\n' + freeText }],
      buildNotesSystemPrompt()
    );
    let cleaned = String(txt).replace(/```json\s*/gi, '').replace(/```\s*$/g, '').trim();
    const m = cleaned.match(/\[\s*[\s\S]*\]/);
    if (m) cleaned = m[0];
    const arr = JSON.parse(cleaned);
    if (!Array.isArray(arr)) throw new Error('AI did not return a JSON array');
    const byId = {};
    trades.forEach(t => { byId[t.id] = t; });
    const items = arr.filter(x => x && byId[x.id]).map(x => ({
      id: x.id,
      ticker: byId[x.id].ticker,
      thesis: x.thesis || '',
      exit_notes: x.exit_notes || ''
    }));
    return { ok: true, items: items };
  } catch (e) {
    return { ok: false, message: e.message };
  }
}

function applyNotes(itemsJson) {
  const items = JSON.parse(itemsJson);
  const sheet = getOrCreateJournalSheet();
  const c = JCONFIG.COL;
  const last = sheet.getLastRow();
  const ids = last < 2 ? [] :
    sheet.getRange(2, c.ID, last - 1, 1).getValues().map(r => String(r[0] || ''));
  let updated = 0;
  const missing = [];
  for (const it of items) {
    const idx = ids.indexOf(String(it.id));
    if (idx < 0) { missing.push(it.id); continue; }
    const row = idx + 2;
    if (it.thesis) sheet.getRange(row, c.THESIS).setValue(it.thesis);
    if (it.exit_notes) sheet.getRange(row, c.EXIT_NOTES).setValue(it.exit_notes);
    updated++;
  }
  return {
    ok: updated > 0,
    message: '✓ Notes saved on ' + updated + ' trade(s).' +
      (missing.length ? ' Unknown ID(s): ' + missing.join(', ') : '')
  };
}

function showNotesDialog() {
  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 12px; font-size: 13px; color: #202124; }
  label { display: block; margin: 4px 0 4px; font-weight: 600; }
  textarea { width: 100%; box-sizing: border-box; height: 150px; padding: 8px; font-size: 13px; border: 1px solid #ccc; border-radius: 4px; }
  .hint { color: #5f6368; font-size: 11px; margin-top: 4px; }
  .btnrow { display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap; }
  button { padding: 8px 14px; font-size: 13px; cursor: pointer; border-radius: 4px; font-weight: 500; }
  .primary { background: #1a73e8; color: white; border: none; }
  .ai { background: #f4b400; color: #000; border: none; }
  .secondary { background: white; border: 1px solid #dadce0; }
  table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 11px; }
  th { background: #f8f9fa; padding: 4px 6px; text-align: left; border-bottom: 2px solid #e0e0e0; }
  td { padding: 3px 4px; border-bottom: 1px solid #eee; }
  td input { width: 100%; box-sizing: border-box; font-size: 11px; padding: 2px 4px; }
  #status { margin-top: 10px; white-space: pre-wrap; font-family: monospace; font-size: 11px; }
  .working { color: #1a73e8; } .err { color: #d93025; } .ok { color: #188038; }
</style>

<label>Talk through your trades (dictate straight into this box)</label>
<textarea id="talk" placeholder="e.g. 'AAPL was a momentum play off the product launch, sold when it stalled. XYZ was a tip from a friend, cut it at a loss, lesson: no more tips without my own research. The SPY puts were an earnings hedge, expired worthless...'"></textarea>
<div class="hint">Mention tickers or company names in any order — Claude figures out which journal rows you mean. Trades you don't mention are left alone. On Mac, double-tap the Fn/globe key to dictate.</div>
<div class="btnrow">
  <button class="ai" onclick="match()">🤖 Match with AI</button>
  <button class="secondary" onclick="google.script.host.close()">Cancel</button>
</div>
<div id="status">Ready.</div>

<div id="review-wrap" style="display:none;">
  <div id="review-table"></div>
  <div class="btnrow">
    <button class="primary" onclick="saveNotes()">Save Notes</button>
    <button class="secondary" onclick="google.script.host.close()">Cancel</button>
  </div>
</div>

<script>
  var MATCHES = [];
  function setStatus(m, c) { var el = document.getElementById('status'); el.textContent = m; el.className = c || ''; }

  function match() {
    var t = document.getElementById('talk').value;
    if (!t.trim()) { setStatus('Say or type something first.', 'err'); return; }
    setStatus('🤖 Matching commentary to trades…', 'working');
    google.script.run
      .withSuccessHandler(onMatched)
      .withFailureHandler(function(e) { setStatus('✗ ' + e.message, 'err'); })
      .matchNotesWithAI(t);
  }

  function esc(s) { return String(s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;'); }

  function onMatched(r) {
    if (!r.ok) { setStatus('✗ ' + r.message, 'err'); return; }
    MATCHES = r.items || [];
    if (!MATCHES.length) { setStatus('Could not match any of that to journal trades. Try including tickers.', 'err'); return; }
    var html = '<table><tr><th>Save</th><th>ID</th><th>Ticker</th><th>Thesis</th><th>Exit notes</th></tr>';
    for (var i = 0; i < MATCHES.length; i++) {
      var m = MATCHES[i];
      html += '<tr>' +
        '<td style="text-align:center"><input type="checkbox" id="k' + i + '" checked style="width:auto"></td>' +
        '<td>' + m.id + '</td><td>' + esc(m.ticker) + '</td>' +
        '<td><input id="t' + i + '" value="' + esc(m.thesis) + '"></td>' +
        '<td><input id="e' + i + '" value="' + esc(m.exit_notes) + '"></td></tr>';
    }
    html += '</table>';
    document.getElementById('review-table').innerHTML = html;
    document.getElementById('review-wrap').style.display = 'block';
    setStatus('✓ Matched ' + MATCHES.length + ' trade(s). Edit wording if needed, untick any, then save.', 'ok');
  }

  function saveNotes() {
    var items = [];
    for (var i = 0; i < MATCHES.length; i++) {
      if (!document.getElementById('k' + i).checked) continue;
      items.push({
        id: MATCHES[i].id,
        thesis: document.getElementById('t' + i).value,
        exit_notes: document.getElementById('e' + i).value
      });
    }
    if (!items.length) { setStatus('Nothing ticked to save.', 'err'); return; }
    setStatus('Saving…', 'working');
    google.script.run
      .withSuccessHandler(function(r) {
        setStatus(r.message, r.ok ? 'ok' : 'err');
        if (r.ok) setTimeout(function() { google.script.host.close(); }, 1500);
      })
      .withFailureHandler(function(e) { setStatus('✗ ' + e.message, 'err'); })
      .applyNotes(JSON.stringify(items));
  }
<\/script>
  `).setWidth(700).setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Thesis / Exit Notes');
}

// =============================================================================
// TRADE SUMMARY — variable lookback window (not hardcoded to any fixed period)
// =============================================================================
function showJournalSummaryDialog() {
  const html = HtmlService.createHtmlOutput(`
<style>
  body { font-family: -apple-system, system-ui, sans-serif; margin: 14px; font-size: 13px; }
  label { display: block; margin: 8px 0 4px; font-weight: 600; }
  select { width: 100%; padding: 6px; font-size: 13px; }
  button { padding: 8px 16px; font-size: 13px; cursor: pointer; border-radius: 4px; margin-top: 14px; margin-right: 6px; }
  .primary { background: #1a73e8; color: white; border: none; }
  .secondary { background: white; border: 1px solid #ccc; }
  #status { margin-top: 12px; font-size: 12px; }
</style>
<label>Report window</label>
<select id="window">
  <option value="30">Last 30 days</option>
  <option value="90" selected>Last 3 months</option>
  <option value="180">Last 6 months</option>
  <option value="365">Last 12 months</option>
  <option value="ytd">Year to date</option>
  <option value="all">All time</option>
</select>
<div>
  <button class="primary" onclick="run()">Generate</button>
  <button class="secondary" onclick="google.script.host.close()">Close</button>
</div>
<div id="status"></div>
<script>
  function run() {
    document.getElementById('status').textContent = 'Working…';
    google.script.run
      .withSuccessHandler(function(r) {
        document.getElementById('status').textContent = r.message;
        if (r.ok) setTimeout(function() { google.script.host.close(); }, 1800);
      })
      .withFailureHandler(function(e) { document.getElementById('status').textContent = '✗ ' + e.message; })
      .buildJournalSummary(document.getElementById('window').value);
  }
<\/script>
  `).setWidth(340).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, 'Trade Summary');
}

function buildJournalSummary(windowKey) {
  const journal = getOrCreateJournalSheet();
  const last = journal.getLastRow();
  if (last < 2) return { ok: false, message: 'No trades logged yet. Use "Log Closed Trades…" first.' };

  const c = JCONFIG.COL;
  const data = journal.getRange(2, 1, last - 1, c.EXIT_NOTES).getValues();

  const now = new Date();
  now.setHours(0, 0, 0, 0);
  let cutoff = null;
  if (windowKey === 'ytd') {
    cutoff = new Date(now.getFullYear(), 0, 1);
  } else if (windowKey !== 'all') {
    cutoff = new Date(now.getTime() - Number(windowKey) * 86400000);
  }

  const rows = data.filter(r => {
    const cd = r[c.CLOSE_DATE - 1];
    if (!(cd instanceof Date) || isNaN(cd.getTime())) return false;
    return !cutoff || cd >= cutoff;
  });

  const pnlIdx = c.PNL_DOLLAR - 1, termIdx = c.TERM - 1, tkrIdx = c.TICKER - 1;
  let wins = 0, losses = 0, grossWin = 0, grossLoss = 0, netPnl = 0, shortPnl = 0, longPnl = 0;
  const bySymbolLoss = {};

  rows.forEach(r => {
    const pnl = Number(r[pnlIdx]) || 0;
    netPnl += pnl;
    if (pnl >= 0) { wins++; grossWin += pnl; } else { losses++; grossLoss += pnl; }
    if (r[termIdx] === 'Long') longPnl += pnl; else shortPnl += pnl;
    const tkr = r[tkrIdx];
    if (pnl < 0) bySymbolLoss[tkr] = (bySymbolLoss[tkr] || 0) + pnl;
  });

  const ss = journal.getParent();
  let summary = ss.getSheetByName(JCONFIG.SUMMARY_SHEET_NAME);
  if (!summary) summary = ss.insertSheet(JCONFIG.SUMMARY_SHEET_NAME);
  summary.clear();

  const label = windowKey === 'all' ? 'All time' : windowKey === 'ytd' ? 'Year to date' : ('Last ' + windowKey + ' days');
  const worst = Object.entries(bySymbolLoss).sort((a, b) => a[1] - b[1]).slice(0, 10);

  const out = [
    ['Trade Summary — ' + label, ''],
    ['Generated', new Date().toLocaleString()],
    ['', ''],
    ['Closed trades', rows.length],
    ['Wins', wins],
    ['Losses', losses],
    ['Win rate', rows.length ? (wins / rows.length) : 0],
    ['Net realized P&L', netPnl],
    ['Gross gains', grossWin],
    ['Gross losses', grossLoss],
    ['Avg win', wins ? (grossWin / wins) : 0],
    ['Avg loss', losses ? (grossLoss / losses) : 0],
    ['Short-term P&L', shortPnl],
    ['Long-term P&L', longPnl],
    ['', ''],
    ['Worst tickers by realized loss', '']
  ];
  worst.forEach(pair => out.push([pair[0], pair[1]]));

  summary.getRange(1, 1, out.length, 2).setValues(out);
  summary.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(13);
  summary.getRange(16, 1, 1, 2).setFontWeight('bold');
  summary.getRange(7, 2, 1, 1).setNumberFormat('0.0%');
  summary.getRange(8, 2, 6, 1).setNumberFormat('$#,##0;-$#,##0');
  if (worst.length) summary.getRange(17, 2, worst.length, 1).setNumberFormat('$#,##0;-$#,##0');
  summary.autoResizeColumns(1, 2);

  return { ok: true, message: '✓ "' + JCONFIG.SUMMARY_SHEET_NAME + '" updated (' + label + ', ' + rows.length + ' trades).' };
}
