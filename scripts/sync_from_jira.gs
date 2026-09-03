/**
 * Vendor Tracker — Sync from Jira   (file: sync_from_jira)
 *
 * What it does: pulls open + recently created intake tickets into the
 *   AllJobs tab. Creates rows for new Mother Tongue tickets (name,
 *   CC, locales, dates, status "Pending", Jira link, PM). For
 *   tickets already tracked, updates the Jira-owned fields when the
 *   ticket changed: deadline, target locales, Mozilla PM, CC.
 *   Every change is listed in the end-of-run popup as old → new;
 *   nothing changes silently. Never touches sheet-owned fields:
 *   status, word count, estimate, invoiced, notes. Job names are
 *   flagged when they differ, never auto-changed.
 *   Skips HTP-vendor tickets ("Legal…" components — they belong in
 *   the master's HTP Requests tab) and tickets closed without work.
 *
 * Also defines the "L10n Tools" menu for both scripts in this
 * project (this file + reconcile_invoices).
 *
 * Deployment values (Jira URL, project key, field IDs) live in
 * config.gs — see config.dist.gs and the README for setup.
 * Credentials live in Script properties (JIRA_EMAIL, JIRA_TOKEN),
 * never in code.
 *
 * Behavior settings (columns, status value, note text) live in
 * CONFIG below.
 */

var CONFIG = {
  SHEET_NAME: 'AllJobs',
  COL: {            // 1-based column indexes in AllJobs
    NAME: 1,        // A Job name
    CC: 2,          // B CC (text!)
    LOCALE: 4,      // D Target Locale
    INITIATED: 9,   // I Job Initiated
    DEADLINE: 10,   // J Deadlines
    STATUS: 11,     // K Status
    JIRA: 13,       // M Jira Issue
    NOTE: 14,       // N NOTE
    PM: 15          // O Mozilla PM
  },
  HTP_COMPONENT: /^Legal/i,             // components matching this = HTP vendor
  NEW_STATUS: 'Pending',                // must be a value from the Status dropdown
  AUTONOTE: 'Auto-created from Jira — fill word count and estimate'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('L10n Tools')
    .addItem('Sync from Jira', 'syncFromJira')
    .addSeparator()
    .addItem('Reconcile invoices…', 'reconcileInvoices')
    .addItem('Apply reconciliation', 'applyReconciliation')
    .addToUi();
}

function syncFromJira() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('JIRA_EMAIL');
  var token = props.getProperty('JIRA_TOKEN');
  if (!email || !token) {
    ui.alert('Setup needed', 'Add JIRA_EMAIL and JIRA_TOKEN in Project Settings → Script properties.', ui.ButtonSet.OK);
    return;
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { ui.alert('Sheet "' + CONFIG.SHEET_NAME + '" not found.'); return; }
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var ticketRe = new RegExp(JIRA_PROJECT + '-\\d+', 'g');

  // --- 1. True last data row = last non-empty cell in column A ---
  var colA = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  var lastRow = 1;
  for (var i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim() !== '') { lastRow = i + 1; break; }
  }

  // --- 2. Existing ticket IDs -> row number (first row wins) ---
  var existing = {};   // { '<PROJECT>-327': rowNumber }
  if (lastRow >= 2) {
    sheet.getRange(2, CONFIG.COL.JIRA, lastRow - 1, 1).getDisplayValues().forEach(function (r, idx) {
      var m = String(r[0]).match(ticketRe);
      if (m) m.forEach(function (id) { if (!(id in existing)) existing[id] = idx + 2; });
    });
  }

  // --- 3. Fetch tickets from Jira (all open + created in last 30d) ---
  var jql = 'project = ' + JIRA_PROJECT +
            ' AND (statusCategory != Done OR created >= -30d) ORDER BY created ASC';
  var url = JIRA_BASE + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) +
            '&maxResults=50&fields=summary,created,duedate,assignee,status,components,' +
            [JIRA_CC_FIELD, JIRA_NAME_FIELD, JIRA_DEADLINE_FIELD, JIRA_LOCALE_FIELD].join(',');
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(email + ':' + token) },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    ui.alert('Jira error ' + resp.getResponseCode(), resp.getContentText().slice(0, 300), ui.ButtonSet.OK);
    return;
  }
  var issues = JSON.parse(resp.getContentText()).issues || [];

  // --- 4. Diff: append new rows, update tracked rows from Jira ---
  var added = [], skipped = [], updated = [], nameFlags = [], alreadyTracked = 0;

  // helper: set a cell if the Jira value differs; record the change
  function updateCell_(row, col, key, label, newVal, displayNew, forceText) {
    var cell = sheet.getRange(row, col);
    var oldVal = cell.getValue();
    var oldDisp = (oldVal instanceof Date)
        ? Utilities.formatDate(oldVal, tz, 'yyyy-MM-dd')
        : String(cell.getDisplayValue()).trim();
    var newDisp = String(displayNew).trim();
    if (oldDisp === newDisp || newDisp === '') return;   // equal, or Jira empty → leave sheet alone
    if (forceText) cell.setNumberFormat('@');
    cell.setValue(newVal);
    updated.push(key + ' — ' + label + ': ' + (oldDisp === '' ? '(empty)' : oldDisp) + ' → ' + newDisp);
  }

  issues.forEach(function (iss) {
    var key = iss.key;
    var f = iss.fields;
    var dueRaw = f[JIRA_DEADLINE_FIELD] || f.duedate;
    var due = dueRaw ? Utilities.parseDate(String(dueRaw).split('T')[0], tz, 'yyyy-MM-dd') : null;
    var dueDisp = due ? Utilities.formatDate(due, tz, 'yyyy-MM-dd') : '';
    var locales = (f[JIRA_LOCALE_FIELD] || '').toString().trim();
    var pm = f.assignee && f.assignee.displayName ? f.assignee.displayName.split(' ')[0] : '';
    pm = PM_ALIAS[pm] || pm;
    var cc = (f[JIRA_CC_FIELD] || '').toString().replace(/^CC/, '');
    var jiraName = (f[JIRA_NAME_FIELD] || '').toString().trim();

    // ---- already tracked: make Jira-owned fields equal to Jira ----
    if (existing[key]) {
      alreadyTracked++;
      var r = existing[key];
      updateCell_(r, CONFIG.COL.LOCALE,   key, 'locales',  locales, locales);
      if (due) updateCell_(r, CONFIG.COL.DEADLINE, key, 'deadline', due, dueDisp);
      updateCell_(r, CONFIG.COL.PM,       key, 'PM',       pm,      pm);
      updateCell_(r, CONFIG.COL.CC,       key, 'CC',       cc,      cc, true);
      // name: flag only, and only for rows the sync created (auto-note present)
      if (jiraName) {
        var note = String(sheet.getRange(r, CONFIG.COL.NOTE).getDisplayValue());
        var curName = String(sheet.getRange(r, CONFIG.COL.NAME).getDisplayValue()).trim();
        if (note.indexOf('Auto-created from Jira') >= 0 && curName !== jiraName) {
          nameFlags.push(key + ' — ticket name is "' + jiraName + '", sheet says "' + curName + '" (not changed)');
        }
      }
      return;
    }

    // ---- new ticket: route or append ----
    var comps = (f.components || []).map(function (c) { return c.name; }).join(', ');
    if (CONFIG.HTP_COMPONENT.test(comps)) { skipped.push(key + ' (HTP/legal → HTP Requests tab)'); return; }

    var doneNoWork = f.status && f.status.statusCategory && f.status.statusCategory.key === 'done';
    if (doneNoWork) { skipped.push(key + ' (closed, no tracker row)'); return; }

    var name = (jiraName || f.summary || key).toString().trim();
    var createdDate = f.created ? new Date(f.created) : new Date();

    lastRow++;
    var row = lastRow;
    // Force CC to TEXT before writing (master's queries compare as text)
    sheet.getRange(row, CONFIG.COL.CC).setNumberFormat('@');
    sheet.getRange(row, CONFIG.COL.NAME).setValue(name);
    sheet.getRange(row, CONFIG.COL.CC).setValue(cc);
    if (locales) sheet.getRange(row, CONFIG.COL.LOCALE).setValue(locales);
    sheet.getRange(row, CONFIG.COL.INITIATED).setValue(createdDate);
    if (due) sheet.getRange(row, CONFIG.COL.DEADLINE).setValue(due);
    sheet.getRange(row, CONFIG.COL.STATUS).setValue(CONFIG.NEW_STATUS);
    sheet.getRange(row, CONFIG.COL.JIRA).setFormula(
      '=HYPERLINK("' + JIRA_BASE + '/browse/' + key + '","' + key + '")');
    sheet.getRange(row, CONFIG.COL.NOTE).setValue(CONFIG.AUTONOTE);
    if (pm) sheet.getRange(row, CONFIG.COL.PM).setValue(pm);
    added.push(key + ' — ' + name);
    existing[key] = row;
  });

  // --- 5. Report: every change is visible, nothing moves silently ---
  var msg = 'Checked ' + issues.length + ' ticket(s).\n\n' +
            'Added: ' + (added.length ? '\n  • ' + added.join('\n  • ') : 'none') + '\n' +
            'Already tracked: ' + alreadyTracked + '\n' +
            (updated.length ? 'Updated from Jira:\n  • ' + updated.join('\n  • ') + '\n' : '') +
            (nameFlags.length ? 'Name differences (review, not changed):\n  • ' + nameFlags.join('\n  • ') + '\n' : '') +
            (skipped.length ? 'Skipped:\n  • ' + skipped.join('\n  • ') : '');
  ui.alert('Sync from Jira', msg, ui.ButtonSet.OK);
}
