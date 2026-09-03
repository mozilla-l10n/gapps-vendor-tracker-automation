/**
 * Vendor Tracker — Reconcile Invoices   (file: reconcile_invoices)
 *
 * What it does: two commands, used together once a month.
 *   "Reconcile invoices…" asks for a month, reads that month's
 *   OC-*.pdf invoices from the shared Drive folder, matches each
 *   invoice to AllJobs row(s) by job name (amounts used for
 *   verification), and writes a "Reconciliation" report tab: one
 *   line per invoice with a checkbox — pre-ticked for clean matches,
 *   unticked with a written reason for anything doubtful. The tab is
 *   a scratchpad: each run overwrites it.
 *   "Apply reconciliation" acts only on ticked lines: fills empty
 *   Invoiced + Invoice Month cells (never overwrites), and writes
 *   the amount to the Jira ticket's invoiced field — filling empty
 *   tickets, verifying matching ones, flagging (not changing) any
 *   that disagree. On months already filled by hand, the same run
 *   verifies PDF ↔ sheet ↔ Jira and surfaces mismatches.
 *
 * Setup (once, in addition to the sync script's setup):
 *   Left sidebar → Services (+) → "Drive API" → version v2 → Add.
 *   (Used to convert PDFs to readable text. Google limits the speed
 *   of conversions, so a run takes a few minutes and waits politely
 *   when told to slow down; very large months may need a second
 *   run after the 6-minute script limit.)
 *
 * Deployment values (invoice folder ID, year, Jira URL and field
 * IDs) live in config.gs — see config.dist.gs and the README.
 * Behavior settings (sheet names, columns) live in RCONFIG below.
 */

var RCONFIG = {
  SHEET_NAME: 'AllJobs',
  REPORT_SHEET: 'Reconciliation',
  COL: { NAME: 1, INVOICED: 12, JIRA: 13, MONTH: 17 }  // A, L, M, Q
};

/* ============================ STEP 1: RECONCILE ============================ */

function reconcileInvoices() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Reconcile invoices',
    'Which month? (e.g. 07 or ' + INVOICE_YEAR + '-07)', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var mm = resp.getResponseText().trim().replace(INVOICE_YEAR + '-', '');
  if (!/^\d{2}$/.test(mm)) { ui.alert('Please enter a two-digit month, e.g. 07'); return; }
  var monthKey = INVOICE_YEAR + '-' + mm;
  var ticketRe = new RegExp(JIRA_PROJECT + '-\\d+');

  // --- find the month's folder (named like "07-July") ---
  var parent = DriveApp.getFolderById(INVOICE_FOLDER_ID);
  var folders = parent.getFolders(), monthFolder = null;
  while (folders.hasNext()) {
    var fo = folders.next();
    if (fo.getName().indexOf(mm + '-') === 0) { monthFolder = fo; break; }
  }
  if (!monthFolder) { ui.alert('No folder starting with "' + mm + '-" found in the ' + INVOICE_YEAR + ' invoice folder.'); return; }

  // --- read + parse every OC-*.pdf ---
  var files = monthFolder.getFiles(), invoices = [];
  while (files.hasNext()) {
    var file = files.next();
    if (!/^OC-.*\.pdf$/i.test(file.getName())) continue;
    var text = pdfToText_(file);
    var subject = (text.match(/Subject\s*\n([\s\S]*?)\n\s*\n?\s*Items/) || [,'???'])[1].replace(/\s+/g, ' ').trim();
    var totalM = text.match(/Total price without (?:VAT|tax)\s+([\d,]+\.\d{2})/);
    var total = totalM ? parseFloat(totalM[1].replace(/,/g, '')) : null;
    // translate epoch-timestamp job names (e.g. 1785158421094) to dates
    subject = subject.replace(/\b(1[67]\d{11})\b/g, function (m) {
      return Utilities.formatDate(new Date(parseInt(m, 10)), 'UTC', 'yyyy-MM-dd');
    });
    invoices.push({ oc: file.getName().replace(/\.pdf$/i, ''), subject: subject, total: total });
  }
  if (!invoices.length) { ui.alert('No OC-*.pdf files found in ' + monthFolder.getName() + '.'); return; }
  invoices.sort(function (a, b) { return a.oc < b.oc ? -1 : 1; });

  // --- load this month's AllJobs rows ---
  var sheet = SpreadsheetApp.getActive().getSheetByName(RCONFIG.SHEET_NAME);
  var last = lastDataRow_(sheet);
  var names = sheet.getRange(2, RCONFIG.COL.NAME, last - 1).getValues();
  var invs  = sheet.getRange(2, RCONFIG.COL.INVOICED, last - 1).getValues();
  var jiras = sheet.getRange(2, RCONFIG.COL.JIRA, last - 1).getDisplayValues();
  var months = sheet.getRange(2, RCONFIG.COL.MONTH, last - 1).getDisplayValues();
  var rows = [];
  for (var i = 0; i < names.length; i++) {
    var rowMonth = String(months[i][0]).trim();
    if (rowMonth !== monthKey && rowMonth !== '') continue;  // this month + not-yet-invoiced rows
    rows.push({ row: i + 2, name: String(names[i][0]), inv: invs[i][0],
                jira: (String(jiras[i][0]).match(ticketRe) || [''])[0],
                monthSet: rowMonth === monthKey, used: false });
  }

  // --- match each invoice to row(s) ---
  var report = [];
  invoices.forEach(function (v) {
    var hits = matchInvoice_(v, rows);
    var sum = hits.reduce(function (s, h) { return s + (typeof h.inv === 'number' ? h.inv : 0); }, 0);
    var allFilled = hits.length && hits.every(function (h) { return h.inv !== '' && h.inv !== null; });
    var status, check;
    if (!hits.length)                    { status = 'NO MATCH — review';        check = false; }
    else if (allFilled && Math.abs(sum - v.total) < 0.01) { status = 'VERIFIED (already filled, amounts agree)'; check = true; }
    else if (allFilled)                  { status = 'MISMATCH: tracker has ' + sum.toFixed(2); check = false; }
    else                                 { status = 'FILL (tracker empty)';     check = true; }
    hits.forEach(function (h) { h.used = true; });
    report.push([check, v.oc, v.total, v.subject.slice(0, 90),
                 hits.map(function (h) { return h.row; }).join(', '),
                 hits.map(function (h) { return h.name.slice(0, 50); }).join(' + '),
                 status,
                 hits.map(function (h) { return h.jira || 'automation'; }).join(', ')]);
  });
  var leftovers = rows.filter(function (r) { return !r.used && r.monthSet; })
                      .map(function (r) { return r.name.slice(0, 60); });

  // --- write the report tab ---
  var ss = SpreadsheetApp.getActive();
  var rep = ss.getSheetByName(RCONFIG.REPORT_SHEET) || ss.insertSheet(RCONFIG.REPORT_SHEET);
  rep.clear();
  rep.getRange(1, 1, 1, 8).setValues([['Apply?', 'Invoice', 'Amount', 'Invoice subject', 'Row(s)', 'Tracker job(s)', 'Status', 'Jira ticket(s)']]).setFontWeight('bold');
  rep.getRange(1, 10).setValue('Month: ' + monthKey);
  if (report.length) {
    rep.getRange(2, 1, report.length, 8).setValues(report);
    rep.getRange(2, 1, report.length, 1).insertCheckboxes();
  }
  if (leftovers.length) {
    rep.getRange(report.length + 3, 2).setValue('Tracker rows for ' + monthKey + ' not covered by any invoice:');
    rep.getRange(report.length + 4, 2, leftovers.length).setValues(leftovers.map(function (n) { return [n]; }));
  }
  ss.setActiveSheet(rep);
  ui.alert('Reconciliation ready', report.length + ' invoice(s) processed.\nReview the Reconciliation tab, tick/untick the Apply boxes, then run "Apply reconciliation".', ui.ButtonSet.OK);
}

/* ============================ STEP 2: APPLY ============================ */

function applyReconciliation() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var rep = ss.getSheetByName(RCONFIG.REPORT_SHEET);
  if (!rep) { ui.alert('Run "Reconcile invoices…" first.'); return; }
  var monthKey = String(rep.getRange(1, 10).getValue()).replace('Month: ', '').trim();
  var sheet = ss.getSheetByName(RCONFIG.SHEET_NAME);
  var data = rep.getRange(2, 1, Math.max(rep.getLastRow() - 1, 1), 8).getValues();
  var ticketRe = new RegExp(JIRA_PROJECT + '-\\d+');

  var filled = 0, jiraOk = 0, jiraFlags = [];
  data.forEach(function (line) {
    if (line[0] !== true || !line[4]) return;           // unchecked or no rows
    var total = line[2];
    var rowNums = String(line[4]).split(',').map(function (s) { return parseInt(s.trim(), 10); }).filter(Boolean);

    // --- fill tracker: first empty row gets the remainder, other empty rows get 0 ---
    var already = 0;
    rowNums.forEach(function (r) {
      var v = sheet.getRange(r, RCONFIG.COL.INVOICED).getValue();
      if (typeof v === 'number') already += v;
    });
    var first = true;
    rowNums.forEach(function (r) {
      var cell = sheet.getRange(r, RCONFIG.COL.INVOICED);
      if (cell.getValue() === '' || cell.getValue() === null) {
        cell.setValue(first ? Math.round((total - already) * 100) / 100 : 0);
        first = false; filled++;
      }
      var mCell = sheet.getRange(r, RCONFIG.COL.MONTH);
      if (String(mCell.getDisplayValue()).trim() === '') mCell.setValue(monthKey);
    });

    // --- write back to Jira (only real tickets; fill-or-verify, never overwrite) ---
    rowNums.forEach(function (r) {
      var key = (String(sheet.getRange(r, RCONFIG.COL.JIRA).getDisplayValue()).match(ticketRe) || [null])[0];
      if (!key) return;
      var amount = sheet.getRange(r, RCONFIG.COL.INVOICED).getValue();
      if (typeof amount !== 'number') return;
      var result = jiraSetInvoiced_(key, amount);
      if (result === 'ok') jiraOk++;
      else if (result) jiraFlags.push(key + ': ' + result);
    });
  });

  ui.alert('Apply done',
    'Tracker cells filled: ' + filled + '\nJira tickets updated/verified: ' + jiraOk +
    (jiraFlags.length ? '\nJira flags:\n  • ' + jiraFlags.join('\n  • ') : ''),
    ui.ButtonSet.OK);
}

/* ============================ HELPERS ============================ */

// PDF -> text via Drive conversion (needs Advanced Drive Service v2 enabled)
// Waits and retries when Google's OCR rate limit pushes back.
function pdfToText_(file) {
  Utilities.sleep(2000);  // gentle pacing between files
  var res = null;
  for (var attempt = 1; attempt <= 5; attempt++) {
    try {
      res = Drive.Files.insert(
        { title: 'tmp-ocr-' + file.getName() },
        file.getBlob(),
        { ocr: true, ocrLanguage: 'en' });
      break;  // success
    } catch (e) {
      if (String(e).indexOf('rate limit') === -1 || attempt === 5) throw e;
      Utilities.sleep(10000 * attempt);  // wait longer each retry: 10s, 20s, 30s…
    }
  }
  var text = DocumentApp.openById(res.id).getBody().getText();
  Drive.Files.remove(res.id);   // delete the temporary document
  return text;
}

function lastDataRow_(sheet) {
  var colA = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  for (var i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

// Match one invoice to tracker rows: exact amount first, then names, then bundles.
function matchInvoice_(v, rows) {
  var free = rows.filter(function (r) { return !r.used; });
  // 1) unique exact amount match on a filled row
  var exact = free.filter(function (r) { return typeof r.inv === 'number' && Math.abs(r.inv - v.total) < 0.01; });
  if (exact.length === 1) return exact;
  // 2) token matching against subject (cost centers, year, quarters and
  //    filler words carry no signal, so they are excluded)
  var stop = { 'the': 1, 'for': 1, 'and': 1, 'q1': 1, 'q2': 1, 'q3': 1, 'q4': 1 };
  stop[INVOICE_YEAR] = 1;
  function toks(s) {
    return String(s).toLowerCase().replace(/[^a-z0-9#\/-]+/g, ' ').split(' ')
      .filter(function (t) { return t.length > 1 && !stop[t] && !/^cc\d+$/.test(t); });
  }
  var st = toks(v.subject);
  var scored = free.map(function (r) {
    var rt = toks(r.name), n = 0;
    rt.forEach(function (t) { if (st.indexOf(t) >= 0) n++; });
    return { r: r, score: n };
  }).filter(function (x) { return x.score >= 2; })
    .sort(function (a, b) { return b.score - a.score; });
  if (!scored.length && exact.length > 1) return [exact[0]]; // amount ties: take first, human reviews
  if (!scored.length) return [];
  var hits = [scored[0].r];
  // 3) bundle: pull in additional scoring rows while sum of filled values < total
  for (var i = 1; i < scored.length; i++) {
    var sum = hits.reduce(function (s, h) { return s + (typeof h.inv === 'number' ? h.inv : 0); }, 0);
    var anyEmpty = hits.some(function (h) { return h.inv === '' || h.inv === null; });
    if (!anyEmpty && Math.abs(sum - v.total) < 0.01) break;
    if (scored[i].score >= 2) hits.push(scored[i].r);
  }
  return hits;
}

// Set Jira invoiced field: fill if empty/0, verify if equal, flag if different.
function jiraSetInvoiced_(key, amount) {
  var props = PropertiesService.getScriptProperties();
  var auth = 'Basic ' + Utilities.base64Encode(props.getProperty('JIRA_EMAIL') + ':' + props.getProperty('JIRA_TOKEN'));
  var get = UrlFetchApp.fetch(JIRA_BASE + '/rest/api/3/issue/' + key + '?fields=' + JIRA_INVOICED_FIELD,
    { headers: { Authorization: auth }, muteHttpExceptions: true });
  if (get.getResponseCode() !== 200) return 'read error ' + get.getResponseCode();
  var current = JSON.parse(get.getContentText()).fields[JIRA_INVOICED_FIELD];
  if (typeof current === 'number' && current !== 0) {
    return Math.abs(current - amount) < 0.01 ? 'ok' : 'has ' + current + ', tracker says ' + amount + ' — NOT changed';
  }
  var body = {}; body[JIRA_INVOICED_FIELD] = amount;
  var put = UrlFetchApp.fetch(JIRA_BASE + '/rest/api/3/issue/' + key,
    { method: 'put', contentType: 'application/json', payload: JSON.stringify({ fields: body }),
      headers: { Authorization: auth }, muteHttpExceptions: true });
  return put.getResponseCode() === 204 ? 'ok' : 'write error ' + put.getResponseCode();
}
