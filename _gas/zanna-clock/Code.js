/**
 * ZANNA CLOCK v1.11.0 — one pay rule, correct breaks, working holiday lookup
 * Apps Script web app bound to a Google Sheet. ALL rules server-side.
 * Mutations wrapped in LockService. State DERIVED from append-only Events log.
 *
 * The header below stopped at v1.5.1 for several releases while the code moved
 * on. What is actually in this file: v1.6 management auth, v1.7 PIN in Script
 * Properties, v1.8 one-tap clock-in, v1.9 the Payroll_App hook, and v1.10 below.
 *
 * v1.11.0 changes (2026-08-17):
 *  - findLeaveTab_ CAN NOW ACTUALLY FIND THE BOOKINGS TAB. It required a column
 *    called `name`; FairLeave's Bookings tab has only `employeeId`, so the
 *    fallback could never match the one tab it exists to find. The lookup
 *    worked solely because getSheetByName('Bookings') hit exactly - rename that
 *    tab and holidays would silently vanish from every display, with no error.
 *  - rebuildPayroll_ CLEARS instead of erroring when Sessions is empty. It
 *    returned early without clearing, so the previous period's figures stayed
 *    on the Payroll tab looking current - straight after a reset, or before
 *    anyone has completed a shift. A missing Sessions TAB is still an error;
 *    an empty one is a legitimate state and now reports 0 rows.
 *
 * v1.10.0 changes (2026-08-17):
 *  - PAYROLL CSV EXPORT RETIRED. doGet no longer serves ?action=payroll, and no
 *    longer reads e.parameter.mgmtPin at all. It paid split shifts wrongly and
 *    it carried the management PIN in a URL. Use Control Panel → Payroll.
 *    payrollCsv_ is still defined below but is unreachable over HTTP.
 *
 * v1.3 changes (2026-08-15):
 *  - ONE PAY RULE. The "Exclude Time" column on the Config tab decides, per
 *    break type, whether that break's minutes come off paid time (YES =
 *    deducted). The Payroll tab, the CSV export and the analytics endpoint all
 *    read it, so they agree by construction. Previously all three disagreed:
 *    the CSV deducted every break, the Payroll tab deducted none, and
 *    analytics used PAY_BREAKS.
 *  - PAY_BREAKS and PAID_BREAK_MAX_MINS are RETIRED. Nothing reads them. Delete
 *    both rows from the Config tab; leaving them in place is what let the two
 *    mechanisms contradict each other unnoticed.
 *  - BREAK PAIRING FIXED. mutate_ set BreakType only on BREAK_START, so a live
 *    BREAK_END wrote a blank, and rebuildSessions_ paired on an exact type
 *    match. Every break taken on the kiosk silently failed to pair and never
 *    reached payroll. BREAK_END now carries the type, and the pairing tolerates
 *    the blank rows written before this fix.
 *  - EVERY break type counts toward payroll, not just BREAK_15 and BREAK_30.
 *  - Unclosed breaks are counted and reported instead of being dropped.
 *  - HOLIDAY LOOKUP FIXED (v1.3.1). FairLeave's Bookings tab has no name
 *    column — it stores an internal employeeId — so the alias lookup found no
 *    name and skipped the whole branch. Setting LEAVE_SHEET_ID alone did
 *    nothing. Bookings now resolve through FairLeave's Employees tab.
 *  - ONE SessionID PER SHIFT (v1.5). mutate_ hardcoded '-001-', so a second
 *    clock-in on the same day reused the first shift's id; rebuildSessions_
 *    took the first CLOCK_IN and first CLOCK_OUT and merged the two, losing
 *    the second shift's hours and attaching its breaks to the wrong shift.
 *    Shifts are now numbered, and the rebuild splits a group into segments so
 *    historical '-001-' rows still resolve into real shifts. Recovered shifts
 *    are reported as "split shifts recovered".
 *  - STAFF KEYED ON EmployeeID (v1.4). The Staff tab gains an EmployeeID
 *    column and the sync matches on it first, so a rename in the Scheduler
 *    updates the row in place and KEEPS the PIN hash. Previously a rename
 *    deactivated the old row and created a nameless new one, forcing the
 *    person to re-enrol at the kiosk. Duplicate active names — which the
 *    kiosk cannot handle at all — are now detected and reported.
 *    Events/Sessions/Payroll stay keyed on NAME: the history records who
 *    someone was at the time. Run upgradeToV14() once.
 *  - Staff sync now DEACTIVATES people who are no longer active in the
 *    Scheduler. It never deletes them — the PIN hash and the exact name that
 *    every Events row references have to survive.
 *
 * UPGRADING FROM v1.2: paste this file over the old code and save. Then:
 *   1. Config tab — delete the PAY_BREAKS and PAID_BREAK_MAX_MINS rows.
 *   2. Config tab — set "Exclude Time" (column E) per break: YES = deducted.
 *   3. Menu → Rebuild analytics (Sessions), then Rebuild payroll.
 * No redeploy is needed unless the kiosk HTML changed. When you DO redeploy:
 * Deploy → Manage deployments → ✏️ edit → Version: New version → Deploy.
 * The /exec URL stays the same. NEVER "New deployment".
 */

// ---------------------------------------------------------------------------
// SETUP / UPGRADE
// ---------------------------------------------------------------------------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = {
    Config:   [['Key', 'Value', 'Notes', 'Comfort Break', 'Exclude Time'],
               ['MGMT_PIN', 'SEE_SCRIPT_PROPERTIES', 'Moved to Project Settings → Script Properties → Mgmt_PIN. Do NOT put a PIN here.', '', ''],
               ['AMBER_PCT', '80', 'Amber when elapsed >= this % of allowance', '', ''],
               ['BREAK_15', '15', '☕ 15 minute break', '', ''],
               ['BREAK_30', '30', '🍽 30 minute break', '', ''],
               ['BREAK_COMFORT', '0', '🚻 Comfort break', 'Yes', 'NO'],
               ['SCHEDULER_SHEET_ID', '', 'Shift Scheduler spreadsheet ID (staff sync)', '', ''],
               ['SCHEDULER_STAFF_TAB', '', 'Optional: exact staff tab name; blank = auto-detect', '', ''],
               ['LEAVE_SHEET_ID', '', 'FairLeave spreadsheet ID (holiday lookup); blank = off', '', '']],
    Staff:    [['Name', 'Department', 'PinHash', 'Salt', 'Active']],
    Devices:  [['Token', 'Label', 'Type', 'Active'],
               [newToken_(), 'Kitchen tablet', 'KIOSK', 'TRUE'],
               [newToken_(), 'Canteen screen', 'DISPLAY_CANTEEN', 'TRUE'],
               [newToken_(), 'Office screen', 'DISPLAY_OFFICE', 'TRUE'],
               [newToken_(), 'Emergency (fire marshal phone)', 'EMERGENCY', 'TRUE']],
    Events:   [['Timestamp', 'Name', 'EventType', 'BreakType', 'SessionID', 'Premises', 'Device']],
    Sessions: [['SessionID', 'Name', 'Department', 'Type', 'BreakType', 'Start', 'End',
                'DurationMins', 'Premises', 'Device', 'Date', 'Status']],
    Overruns: [['Timestamp', 'Name', 'BreakType', 'AllowedMins', 'ActualMins', 'OverByMins']],
    Absences: [['Name', 'From', 'To', 'Type', 'Notes']],
    Audit:    [['Timestamp', 'Actor', 'Action', 'Details']]
  };
  Object.keys(tabs).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, tabs[name].length, tabs[name][0].length).setValues(tabs[name]);
      sh.setFrozenRows(1);
    }
  });
  audit_('setup', 'setup', 'Tabs verified/created');
}

/** One-time migration from v1.0 — safe to re-run. */
function upgradeToV11() {
  setup();                                            // adds Absences tab if missing
  const sh = tab_('Config'), rows = sh.getDataRange().getValues();
  const have = {};
  for (let i = 1; i < rows.length; i++) have[String(rows[i][0])] = i + 1;
  const oldB15 = have['BREAK_15_MINS'] ? String(rows[have['BREAK_15_MINS'] - 1][1]) : '15';
  const oldB30 = have['BREAK_30_MINS'] ? String(rows[have['BREAK_30_MINS'] - 1][1]) : '30';
  if (!have['BREAK_15'])      sh.appendRow(['BREAK_15', oldB15, '☕ 15 minute break']);
  if (!have['BREAK_30'])      sh.appendRow(['BREAK_30', oldB30, '🍽 30 minute break']);
  if (!have['BREAK_COMFORT']) sh.appendRow(['BREAK_COMFORT', '0', '🚻 Comfort break']);
  if (!have['LEAVE_SHEET_ID']) sh.appendRow(['LEAVE_SHEET_ID', '', 'FairLeave spreadsheet ID (holiday lookup); blank = off']);
  if (!String(sh.getRange(1, 4).getValue()).trim()) sh.getRange(1, 4).setValue('Comfort Break');
  const cRow = sh.createTextFinder('BREAK_COMFORT').matchEntireCell(true).findNext();
  if (cRow && !String(sh.getRange(cRow.getRow(), 4).getValue()).trim()) {
    sh.getRange(cRow.getRow(), 4).setValue('Yes');
  }
  ['BREAK_30_MINS', 'BREAK_15_MINS'].forEach(function (k) {
    const r = sh.createTextFinder(k).matchEntireCell(true).findNext();
    if (r) sh.deleteRow(r.getRow());
  });
  audit_('management', 'upgrade', 'Upgraded to v1.1');
}

/** One-time migration to v1.2 — safe to re-run. Adds the Sessions tab. */
function upgradeToV12() {
  setup();                                            // adds Sessions tab if missing
  const r = rebuildSessions_();
  audit_('management', 'upgrade', 'Upgraded to v1.2 (Sessions analytics) — ' + r.sessions + ' rows');
}

/**
 * One-time migration to v1.3 — safe to re-run.
 * Ensures the "Exclude Time" header exists, then rebuilds. It does NOT invent a
 * pay policy: every break stays paid until someone puts YES in that column,
 * because guessing at this silently changes what people are paid.
 */
function upgradeToV13() {
  setup();
  const sh = tab_('Config');
  const header = sh.getRange(1, 1, 1, Math.max(5, sh.getLastColumn())).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  if (header.indexOf('comfort break') < 0 && !String(sh.getRange(1, 4).getValue()).trim()) {
    sh.getRange(1, 4).setValue('Comfort Break');
  }
  if (header.indexOf('exclude time') < 0 && !String(sh.getRange(1, 5).getValue()).trim()) {
    sh.getRange(1, 5).setValue('Exclude Time');
  }
  const excluded = excludedBreaks_();
  const on = Object.keys(excluded).filter(function (k) { return excluded[k]; });
  const r = rebuildSessions_();
  audit_('management', 'upgrade', 'Upgraded to v1.3 — deducted break types: ' +
         (on.length ? on.join(', ') : 'NONE SET') + '; Sessions ' + r.sessions + ' rows');
  return { ok: true, deductedBreakTypes: on, sessions: r.sessions };
}

/** Adds the EmployeeID header if missing, then syncs to populate it. */
function upgradeToV14() {
  var sh = tab_('Staff');
  if (!sh) return { ok: false, error: 'No Staff tab.' };
  var cols = staffCols_(sh);
  if (cols.empId < 0) {
    var col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue('EmployeeID');
    audit_('management', 'upgrade', 'Staff tab gained an EmployeeID column');
  }
  var r = syncScheduler_();
  var msg = 'v1.4: EmployeeID column ready. ' +
            (r.ok ? 'Sync backfilled ' + (r.idsBackfilled || 0) + ' id(s).' : 'Sync failed: ' + r.error);
  Logger.log(msg);
  return { ok: true, message: msg, sync: r };
}

function newToken_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 20);
}

// ---------------------------------------------------------------------------
// MANAGEMENT MENU (in the Sheet)
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⏰ Zanna Clock')
    .addItem('🔄 Sync staff from Scheduler', 'menuSync')
    .addItem('🤒 Record sick absence', 'menuSick')
    .addSeparator()
    .addItem('🧱 Rebuild analytics (Sessions)', 'menuRebuildSessions')
    .addItem('💰 Rebuild payroll', 'menuRebuildPayroll')
    .addSeparator()
    .addItem('🩺 Diagnose Scheduler link', 'menuDiagnoseScheduler')
    .addItem('🩺 Diagnose FairLeave link', 'menuDiagnoseLeave')
    .addItem('🩺 Show break pay rules', 'menuShowPayRules')
    .addItem('🩺 Staff IDs', 'menuStaffIds')
    .addItem('🩺 Which code is live', 'menuCodeAudit')
    .addToUi();
}

function menuSync() {
  var r = syncScheduler_();
  if (!r.ok) return SpreadsheetApp.getUi().alert('Sync failed: ' + r.error);
  var msg = 'Sync complete ✓\nAdded: ' + r.added + '\nUpdated: ' + r.updated +
            '\nRenamed: ' + r.renamed + '\nIDs backfilled: ' + r.idsBackfilled +
            '\nDeactivated: ' + r.deactivated;
  if (r.duplicateNames && r.duplicateNames.length) {
    msg += '\n\n⚠ Two or more active staff share a name:\n  ' + r.duplicateNames.join(', ') +
           '\nThe second cannot clock in. Fix in the Scheduler.';
  }
  if (r.duplicateIds && r.duplicateIds.length) {
    msg += '\n\n⚠ Duplicate EmployeeID in the Scheduler:\n  ' + r.duplicateIds.join(', ');
  }
  SpreadsheetApp.getUi().alert(msg);
}

function menuRebuildSessions() {
  const r = rebuildSessions_();
  SpreadsheetApp.getUi().alert(r.ok
    ? 'Sessions rebuilt ✓\nRows: ' + r.sessions +
      (r.open ? '\nStill open/in-progress: ' + r.open : '') +
      (r.unclosed ? '\nUnclosed breaks (started, never ended): ' + r.unclosed : '') +
      (r.splitShifts ? '\nSplit shifts recovered: ' + r.splitShifts : '')
    : 'Rebuild failed: ' + r.error);
}

function menuRebuildPayroll() {
  const r = rebuildPayroll_();
  SpreadsheetApp.getUi().alert(r.ok
    ? 'Payroll rebuilt ✓\nRows: ' + r.rows +
      (r.unknownBreakTypes && r.unknownBreakTypes.length
        ? '\n\n⚠ Not in Config, treated as PAID:\n' + r.unknownBreakTypes.join(', ') : '')
    : 'Rebuild failed: ' + r.error);
}

/** Shows exactly which breaks are deducted, so the pay rule is never a guess. */
function menuShowPayRules() {
  const ex = excludedBreaks_();
  const keys = Object.keys(ex).sort();
  if (!keys.length) return SpreadsheetApp.getUi().alert('No BREAK_ rows found in Config.');
  const lines = keys.map(function (k) {
    return (ex[k] ? '  DEDUCTED  ' : '  paid      ') + k;
  });
  SpreadsheetApp.getUi().alert(
    'Break pay rules\n\nSet by the "Exclude Time" column on the Config tab.\nYES = deducted from paid time.\n\n' +
    lines.join('\n'));
}

function menuSick() {
  const ui = SpreadsheetApp.getUi();
  const name = ui.prompt('Sick absence', 'Staff name (exactly as on the Staff tab):', ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() !== ui.Button.OK || !name.getResponseText().trim()) return;
  if (!staffRow_(name.getResponseText())) return ui.alert('No active staff member called "' + name.getResponseText() + '"');
  const from = ui.prompt('Sick absence', 'From date (YYYY-MM-DD), blank = today:', ui.ButtonSet.OK_CANCEL);
  if (from.getSelectedButton() !== ui.Button.OK) return;
  const to = ui.prompt('Sick absence', 'To date (YYYY-MM-DD), blank = same day:', ui.ButtonSet.OK_CANCEL);
  if (to.getSelectedButton() !== ui.Button.OK) return;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const f = from.getResponseText().trim() || today;
  const t = to.getResponseText().trim() || f;
  tab_('Absences').appendRow([name.getResponseText().trim(), f, t, 'SICK', '']);
  audit_('management', 'sickAbsence', name.getResponseText() + ' ' + f + '→' + t);
  ui.alert('Recorded: ' + name.getResponseText() + ' sick ' + f + ' → ' + t +
           '\n\nTo end it early or correct it, edit/delete the row on the Absences tab.');
}

function menuCodeAudit() { SpreadsheetApp.getUi().alert(codeAudit()); }

function menuStaffIds() { SpreadsheetApp.getUi().alert(diagnoseStaffIds()); }

function menuDiagnoseScheduler() { SpreadsheetApp.getUi().alert(diagnoseScheduler_()); }
function menuDiagnoseLeave() { SpreadsheetApp.getUi().alert(diagnoseLeave_()); }

// ---------------------------------------------------------------------------
// HTTP ENTRY POINTS
// ---------------------------------------------------------------------------
/**
 * doGet — version probe only. The CSV export was retired here in v1.10.0.
 *
 * The payroll CSV export that used to live here (?action=payroll) is RETIRED.
 * Two reasons, both real:
 *
 *   1. IT PAID THE WRONG MONEY. payrollCsv_ walks the Events log directly and
 *      counts from the first clock-in to the last clock-out of the day. It
 *      predates the per-break "Exclude Time" rules, so a split shift is paid
 *      for the gap between the two shifts — someone working 07:00-11:00 and
 *      16:00-20:00 is paid for the five hours they were at home.
 *      rebuildPayroll_ is the correct calculation, and the Payroll app reads
 *      that and nothing else.
 *
 *   2. THE MANAGEMENT PIN TRAVELLED IN THE URL. Query strings end up in browser
 *      history, referrer headers and any proxy log in between. This function no
 *      longer reads e.parameter.mgmtPin — so an old bookmarked link cannot even
 *      offer the PIN to be checked, let alone leak a working one.
 *
 * The retired link answers with an explanation rather than silence, because
 * whoever finds it will be a manager mid-payroll-run wondering why their
 * bookmark broke, and "nothing happened" sends them looking for a fault.
 *
 * payrollCsv_ itself is left in the file, now unreachable over HTTP. Deleting
 * it is a separate, calmer job.
 */
function doGet(e) {
  var qr = clkQrDoGet(e);
  if (qr) return qr;
  if (e && e.parameter && e.parameter.action === 'payroll') {
    audit_('unknown', 'payrollCsv.retired',
           'Someone opened the retired ?action=payroll link');
    return json_({
      ok: false,
      error: 'retired',
      message: 'This payroll export was retired in v1.10.0. It calculated split ' +
               'shifts incorrectly and carried the management PIN in the URL. ' +
               'Use Control Panel > Payroll instead.'
    });
  }
  return json_({ ok: true, app: 'ZannaClock', version: '1.11.0' });
}

function doPost(e) {
  var qr = clkQrDoPost(e);
  if (qr) return qr;

  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request body' }); }

  try {
    const device = deviceFor_(req.deviceToken);
    if (!device) return json_({ ok: false, error: 'Unknown or inactive device' });

    if (req.action === 'liveStatus')  return json_(liveStatus_(device.type));
    if (req.action === 'emergency')   return json_(emergencyRegister_());

    const kioskActions = ['getStaff', 'setPin', 'startSession', 'startShift', 'clockIn',
                          'clockOut', 'startBreak', 'endBreak'];
    if (kioskActions.indexOf(req.action) >= 0) {
      if (device.type !== 'KIOSK') return json_({ ok: false, error: 'This device cannot perform that action' });
      switch (req.action) {
        case 'getStaff':     return json_(getStaff_());
        case 'setPin':       return json_(setPin_(req.name, req.pin));
        case 'startSession': return json_(startSession_(req.name, req.pin));
        case 'startShift':   return json_(startShift_(req, device));
        case 'clockIn':      return json_(mutate_(req, 'CLOCK_IN', device));
        case 'clockOut':     return json_(mutate_(req, 'CLOCK_OUT', device));
        case 'startBreak':   return json_(mutate_(req, 'BREAK_START', device));
        case 'endBreak':     return json_(mutate_(req, 'BREAK_END', device));
      }
    }

    if (clkMgmtAuth_(req.mgmtPin, req.action)) {
      switch (req.action) {
        case 'resetPin':          return json_(resetPin_(req.name));
        case 'syncScheduler':     return json_(syncScheduler_());
        case 'diagnoseScheduler': return json_({ ok: true, report: diagnoseScheduler_() });
        case 'overrunReport':     return json_(overrunReport_(req.from, req.to));
        case 'rebuildSessions':   return json_(rebuildSessions_());
        case 'rebuildPayroll':    return json_(rebuildPayroll_());
        case 'analytics':         return json_(analytics_(req.from, req.to, req.name, req.dept));
        default:
          // Payroll_App.gs adds its own management ops here. Kept in a separate
          // file so this switch never has to grow again, and so a full-file
          // paste of Code.gs cannot take the payroll app down with it.
          if (typeof clockPayrollHandle_ === 'function') {
            var pay = clockPayrollHandle_(req.action, req);
            if (pay) return json_(pay);
          }
      }
    }
    return json_({ ok: false, error: 'Unknown action or missing permission' });
  } catch (err) {
    audit_('system', 'error', String(err));
    return json_({ ok: false, error: 'Server error: ' + err.message });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// BREAK TYPES — config-driven
// Any Config key starting with BREAK_ is a break. Value = allowed minutes.
// Notes = button label. Row order = UI order.
// Column D "Comfort Break" = Yes → untimed, always on-premises, no overrun,
//   masked on the canteen screen. (Value 0 also implies comfort.)
// Column E "Exclude Time"  = YES → that break's minutes are DEDUCTED from paid
//   time. These are two different questions about the same break; the old code
//   read D when it meant E, which is how payroll came to deduct nothing.
// ---------------------------------------------------------------------------
function breaks_() {
  const rows = tab_('Config').getDataRange().getValues(), out = [];
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]).trim();
    if (key.indexOf('BREAK_') === 0 && String(rows[i][1]).trim() !== '') {
      const mins = Number(rows[i][1]) || 0;
      const comfort = /^y/i.test(String(rows[i][3] || '').trim()) || mins === 0;
      out.push({ key: key, mins: mins, timed: !comfort,
                 label: String(rows[i][2] || '').trim() || key });
    }
  }
  return out;
}

function breakByKey_(key) {
  const list = breaks_();
  for (let i = 0; i < list.length; i++) if (list[i].key === key) return list[i];
  return null;
}

/**
 * Which break types are EXCLUDED from paid time. The single source of truth for
 * pay, read by the Payroll tab, the CSV export and the analytics endpoint.
 * Returns { BREAK_15: false, BREAK_30: true, ... } for every BREAK_* key.
 *
 * The column is resolved by HEADER, not position — the bug this replaces read
 * column D ("Comfort Break") while intending column E ("Exclude Time").
 */
function excludedBreaks_() {
  const rows = tab_('Config').getDataRange().getValues();
  const header = rows[0].map(function (h) { return String(h).trim().toLowerCase(); });
  let col = header.indexOf('exclude time');
  if (col < 0) col = 4;                       // column E, where it lives today
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]).trim();
    if (key.indexOf('BREAK_') !== 0) continue;
    const v = String(rows[i][col] || '').trim().toUpperCase();
    out[key] = (v === 'YES' || v === 'Y' || v === 'TRUE');
  }
  return out;
}

/** Human summary of a break tally: "BREAK_30 × 2 = 60 min". */
function breakSummary_(tally, keys) {
  const parts = keys.map(function (k) {
    return k + ' × ' + tally[k].count + ' = ' + tally[k].mins + ' min';
  });
  return parts.length ? parts.join('; ') : '—';
}

// ---------------------------------------------------------------------------
// AUTH
// ---------------------------------------------------------------------------
function deviceFor_(token) {
  if (!token) return null;
  const rows = tab_('Devices').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(token) && String(rows[i][3]).toUpperCase() === 'TRUE') {
      return { token: rows[i][0], label: rows[i][1], type: String(rows[i][2]).toUpperCase() };
    }
  }
  return null;
}

function hashPin_(salt, pin) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + String(pin));
  return digest.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/** Find a staff member by NAME or by EmployeeID. Active rows only. */
function staffRow_(nameOrId) {
  var sh = tab_('Staff'), rows = sh.getDataRange().getValues();
  var c = staffCols_(sh);
  var needle = String(nameOrId).trim().toLowerCase();
  if (!needle) return null;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][c.active]).toUpperCase() !== 'TRUE') continue;
    var nm = String(rows[i][c.name]).trim().toLowerCase();
    var id = c.empId >= 0 ? String(rows[i][c.empId]).trim().toLowerCase() : '';
    if (nm === needle || (id && id === needle)) {
      return { row: i + 1, name: rows[i][c.name], dept: rows[i][c.dept],
               pinHash: rows[i][c.pin], salt: rows[i][c.salt],
               empId: c.empId >= 0 ? String(rows[i][c.empId]).trim() : '',
               cols: c };
    }
  }
  return null;
}

/**
 * Everyone's status for today in ONE pass over the Events tab.  (v1.8.0)
 *
 * currentStateFor_ re-reads the whole Events tab per person. Calling it fifteen
 * times to colour a list would be fifteen full sheet reads on every kiosk
 * refresh — the opposite of the point. Same state machine, read once.
 */
function todaysStatuses_() {
  var out = {};
  todaysEvents_().forEach(function (e) {
    var k = e.name;
    if (!out[k]) out[k] = { status: 'OUT', breakType: '', shifts: 0 };
    if (e.type === 'CLOCK_IN')    { out[k].status = 'WORKING';  out[k].shifts++; }
    if (e.type === 'CLOCK_OUT')   { out[k].status = 'OUT'; }
    if (e.type === 'BREAK_START') { out[k].status = 'ON_BREAK'; out[k].breakType = e.breakType; }
    if (e.type === 'BREAK_END')   { out[k].status = 'WORKING';  out[k].breakType = ''; }
  });
  return out;
}

function getStaff_() {
  var sh = tab_('Staff'), rows = sh.getDataRange().getValues();
  var c = staffCols_(sh);
  var st = todaysStatuses_();
  var staff = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][c.active]).toUpperCase() === 'TRUE') {
      var nm = String(rows[i][c.name]);
      var s = st[nm] || { status: 'OUT', breakType: '', shifts: 0 };
      staff.push({ name: rows[i][c.name], dept: rows[i][c.dept],
                   empId: c.empId >= 0 ? String(rows[i][c.empId]).trim() : '',
                   hasPin: !!rows[i][c.pin],
                   // v1.8.0 — so the kiosk can colour the list without a second
                   // request. Older kiosks simply ignore these three fields.
                   status: s.status,
                   breakType: s.breakType,
                   shiftsToday: s.shifts });
    }
  }
  staff.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return { ok: true, staff: staff, breaks: breaks_() };
}

function setPin_(name, pin) {
  if (!/^\d{4}$/.test(String(pin))) return { ok: false, error: 'PIN must be exactly 4 digits' };
  return withLock_(function () {
    var s = staffRow_(name);
    if (!s) return { ok: false, error: 'Staff member not found' };
    if (s.pinHash) return { ok: false, error: 'PIN already set — ask management for a reset' };
    var salt = newToken_();
    tab_('Staff').getRange(s.row, s.cols.pin + 1).setValue(hashPin_(salt, pin));
    tab_('Staff').getRange(s.row, s.cols.salt + 1).setValue(salt);
    audit_(s.name, 'setPin', 'First-use PIN set');
    return { ok: true };
  });
}

function resetPin_(name) {
  return withLock_(function () {
    var s = staffRow_(name);
    if (!s) return { ok: false, error: 'Staff member not found' };
    tab_('Staff').getRange(s.row, s.cols.pin + 1).setValue('');
    tab_('Staff').getRange(s.row, s.cols.salt + 1).setValue('');
    audit_('management', 'resetPin', s.name);
    return { ok: true };
  });
}

function startSession_(name, pin) {
  const s = staffRow_(name);
  if (!s) return { ok: false, error: 'Staff member not found' };
  if (!s.pinHash) return { ok: false, needsPin: true };
  if (hashPin_(s.salt, pin) !== s.pinHash) {
    audit_(name, 'pinFail', 'Wrong PIN at kiosk');
    return { ok: false, error: 'Wrong PIN' };
  }
  return { ok: true, state: currentStateFor_(name) };
}

/**
 * ONE-TAP MORNING CLOCK-IN.  (v1.8.0)
 *
 * The kiosk used to need two round trips to start a shift: startSession to
 * check the PIN, then clockIn from the action screen. At 07:00 with fifteen
 * people arriving together that is fifteen queues of two taps and two waits.
 *
 * startShift does both in ONE request: verify the PIN, and if this is the first
 * clock-in of the day, write the CLOCK_IN and tell the kiosk to say "Clocked In"
 * and go straight back to the staff list.
 *
 * It is deliberately NOT a blanket auto-clock-in. It fires only when:
 *   • the person is currently OUT, and
 *   • they have no CLOCK_IN recorded today.
 *
 * Anything else — mid-shift, on a break, or back for a second shift — returns
 * autoClockedIn:false WITHOUT writing anything, and the kiosk shows the normal
 * action screen. Still one round trip, so the fast path costs the slow path
 * nothing.
 *
 * A second shift is left to the normal screen on purpose. Someone returning at
 * 17:00 has a real choice to make and is not part of a morning queue; silently
 * starting a second shift on one tap is how a mis-tap becomes a payroll query.
 * If you want that too, the rule is the `priorShifts === 0` test below.
 */
function startShift_(req, device) {
  var name = req && req.name;
  var s = staffRow_(name);
  if (!s) return { ok: false, error: 'Staff member not found' };
  if (!s.pinHash) return { ok: false, needsPin: true };
  if (hashPin_(s.salt, req.pin) !== s.pinHash) {
    audit_(name, 'pinFail', 'Wrong PIN at kiosk');
    return { ok: false, error: 'Wrong PIN' };
  }

  // Use the canonical name from the Staff row, not what the kiosk sent.
  var who = s.name;
  var st = currentStateFor_(who);
  var priorShifts = todaysEvents_().filter(function (e) {
    return e.name === who && e.type === 'CLOCK_IN';
  }).length;

  var isShiftStart = (st.status === 'OUT' && priorShifts === 0);
  if (!isShiftStart) {
    return { ok: true, autoClockedIn: false, name: who, state: st,
             reason: st.status !== 'OUT' ? 'already_' + st.status.toLowerCase()
                                         : 'second_shift_today' };
  }

  // mutate_ re-checks the PIN and takes the lock. Re-checking is cheap and
  // means there is exactly one place that decides whether a PIN is good.
  var r = mutate_(req, 'CLOCK_IN', device);
  if (!r.ok) return r;

  return { ok: true, autoClockedIn: true, name: who,
           message: 'Clocked In', state: r.state };
}

// ---------------------------------------------------------------------------
// EVENTS — state is always derived, never stored
// ---------------------------------------------------------------------------
function mutate_(req, eventType, device) {
  return withLock_(function () {
    const s = staffRow_(req.name);
    if (!s) return { ok: false, error: 'Staff member not found' };
    if (!s.pinHash || hashPin_(s.salt, req.pin) !== s.pinHash) {
      return { ok: false, error: 'PIN check failed' };
    }
    const st = currentStateFor_(s.name);
    const now = new Date();

    if (eventType === 'CLOCK_IN'  && st.status !== 'OUT')
      return { ok: false, error: 'Already clocked in' };
    if (eventType === 'CLOCK_OUT' && st.status === 'OUT')
      return { ok: false, error: 'Not clocked in' };
    if (eventType === 'CLOCK_OUT' && st.status === 'ON_BREAK')
      return { ok: false, error: 'End your break before clocking out' };
    if (eventType === 'BREAK_START' && st.status !== 'WORKING')
      return { ok: false, error: st.status === 'ON_BREAK' ? 'Already on a break' : 'Clock in first' };
    if (eventType === 'BREAK_END' && st.status !== 'ON_BREAK')
      return { ok: false, error: 'No active break' };

    let breakType = '', premises = '', sessionID = '';
    if (eventType === 'CLOCK_IN') {
      // Number the shift. The sequence used to be hardcoded '001', so a second
      // clock-in on the same day silently reused the first shift's id and the
      // two were merged when Sessions rebuilt.
      const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const priorShifts = todaysEvents_().filter(function (e) {
        return e.name === s.name && e.type === 'CLOCK_IN';
      }).length;
      const seq = ('00' + (priorShifts + 1)).slice(-3);
      sessionID = s.name + '-' + dateStr + '-' + seq + '-SHIFT';
    } else if (eventType === 'BREAK_START') {
      const bk = breakByKey_(String(req.breakType || '').toUpperCase());
      if (!bk) return { ok: false, error: 'Unknown break type' };
      breakType = bk.key;
      premises = !bk.timed ? 'ON'
               : (String(req.premises || '').toUpperCase() === 'OFF' ? 'OFF' : 'ON');
      sessionID = st.sessionID || '';
    } else if (eventType === 'BREAK_END') {
      // Carry the type onto the END row, or rebuildSessions_ cannot pair it.
      breakType = st.breakType || '';
      premises = st.premises || '';
      sessionID = st.sessionID || '';
    } else {
      sessionID = st.sessionID || '';        // CLOCK_OUT
    }

    tab_('Events').appendRow([now, s.name, eventType, breakType, sessionID, premises,
                              (device && device.label) || 'kiosk']);

    if (eventType === 'BREAK_END' && st.breakType) {
      const bk = breakByKey_(st.breakType);
      if (bk && bk.timed) {
        const actual = Math.round((now - new Date(st.since)) / 60000);
        if (actual > bk.mins) {
          tab_('Overruns').appendRow([now, s.name, bk.key, bk.mins, actual, actual - bk.mins]);
          audit_(s.name, 'overrun', bk.key + ' over by ' + (actual - bk.mins) + ' min');
        }
      }
    }
    audit_(s.name, eventType, breakType ? (breakType + '/' + premises) : '');
    return { ok: true, state: currentStateFor_(s.name) };
  });
}

function currentStateFor_(name) {
  const ev = todaysEvents_().filter(function (e) { return e.name === name; });
  let status = 'OUT', breakType = '', premises = '', since = null, clockIn = null, sessionID = '';
  ev.forEach(function (e) {
    if (e.type === 'CLOCK_IN')   { status = 'WORKING'; since = e.ts; clockIn = e.ts; sessionID = e.sessionID; }
    if (e.type === 'CLOCK_OUT')  { status = 'OUT'; since = e.ts; sessionID = ''; }
    if (e.type === 'BREAK_START'){ status = 'ON_BREAK'; breakType = e.breakType; premises = e.premises; since = e.ts; }
    if (e.type === 'BREAK_END')  { status = 'WORKING'; breakType = ''; premises = ''; since = e.ts; }
  });
  return { status: status, breakType: breakType, premises: premises, sessionID: sessionID,
           since: since ? since.toISOString() : null,
           clockIn: clockIn ? clockIn.toISOString() : null };
}

function todaysEvents_() {
  const midnight = new Date(); midnight.setHours(0, 0, 0, 0);
  const rows = tab_('Events').getDataRange().getValues(), out = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const ts = new Date(rows[i][0]);
    if (ts < midnight) break;
    out.unshift({ ts: ts, name: String(rows[i][1]), type: String(rows[i][2]),
                  breakType: String(rows[i][3]), sessionID: String(rows[i][4]),
                  premises: String(rows[i][5]) });
  }
  return out;
}

// ---------------------------------------------------------------------------
// SESSIONS — DERIVED, LINKED via explicit SessionID
// ---------------------------------------------------------------------------
function rebuildSessions_() {
  return withLock_(function () {
    const evRows = tab_('Events').getDataRange().getValues();
    const header = evRows[0];

    const colTimestamp = header.indexOf('Timestamp');
    const colName = header.indexOf('Name');
    const colEventType = header.indexOf('EventType');
    const colBreakType = header.indexOf('BreakType');
    const colSessionID = header.indexOf('SessionID');
    const colPremises = header.indexOf('Premises');
    const colDevice = header.indexOf('Device');

    if (colTimestamp < 0 || colName < 0 || colEventType < 0 || colSessionID < 0) {
      return { ok: false, error: 'Events tab columns missing: Timestamp, Name, EventType, SessionID required.' };
    }

    const staffSheet = tab_('Staff');
    const staffMap = {};
    if (staffSheet) {
      const staffData = staffSheet.getDataRange().getValues();
      const sc = (typeof staffCols_ === 'function') ? staffCols_(staffSheet)
               : { name: 0, dept: 1 };
      for (let i = 1; i < staffData.length; i++) {
        const n = String(staffData[i][sc.name]).trim();
        if (n) staffMap[n] = String(staffData[i][sc.dept]).trim();
      }
    }

    const sessionData = {};
    for (let i = 1; i < evRows.length; i++) {
      const row = evRows[i];
      const sessionID = String(row[colSessionID] || '').trim();
      if (!sessionID) continue;
      const nm = String(row[colName]).trim();
      if (!sessionData[sessionID]) {
        sessionData[sessionID] = { name: nm, dept: staffMap[nm] || '', events: [] };
      }
      sessionData[sessionID].events.push({
        timestamp: new Date(row[colTimestamp]),
        eventType: String(row[colEventType]).trim(),
        breakType: String(row[colBreakType] || '').trim(),
        premises: String(row[colPremises] || '').trim(),
        device: String(row[colDevice] || '').trim()
      });
    }

    const sessionsRows = [];
    const tz = Session.getScriptTimeZone();
    let openCount = 0, unclosedCount = 0, splitCount = 0;

    Object.keys(sessionData).sort().forEach(function (sessionID) {
      const data = sessionData[sessionID];
      const events = data.events.sort(function (a, b) { return a.timestamp - b.timestamp; });

      // Split the group into SHIFT SEGMENTS. Rows written before this fix share
      // one SessionID for every shift that day, so a single group can hold two
      // or more complete shifts. Segmenting recovers them instead of merging.
      const segments = [];
      let cur = null;
      events.forEach(function (evt) {
        if (evt.eventType === 'CLOCK_IN') {
          if (cur) segments.push(cur);            // previous shift never clocked out
          cur = { inEvt: evt, outEvt: null, inner: [] };
        } else if (evt.eventType === 'CLOCK_OUT') {
          if (cur && !cur.outEvt) { cur.outEvt = evt; segments.push(cur); cur = null; }
        } else if (cur) {
          cur.inner.push(evt);
        }
      });
      if (cur) segments.push(cur);
      if (segments.length > 1) splitCount++;

      segments.forEach(function (seg, idx) {
        // Keep the first segment on the original id so existing references still
        // resolve; suffix the rest, which also makes the historical collision
        // visible on the Sessions tab rather than hidden.
        const sid = idx === 0 ? sessionID : sessionID + '#' + (idx + 1);
        const dateStr = Utilities.formatDate(seg.inEvt.timestamp, tz, 'yyyy-MM-dd');

        if (seg.outEvt) {
          sessionsRows.push([sid, data.name, data.dept, 'SHIFT', '',
            Utilities.formatDate(seg.inEvt.timestamp, tz, 'yyyy-MM-dd HH:mm:ss'),
            Utilities.formatDate(seg.outEvt.timestamp, tz, 'yyyy-MM-dd HH:mm:ss'),
            Math.round((seg.outEvt.timestamp - seg.inEvt.timestamp) / 60000),
            '', seg.inEvt.device, dateStr, 'closed']);
        } else {
          sessionsRows.push([sid, data.name, data.dept, 'SHIFT', '',
            Utilities.formatDate(seg.inEvt.timestamp, tz, 'yyyy-MM-dd HH:mm:ss'),
            '', '', '', seg.inEvt.device, dateStr, 'open']);
          openCount++;
        }

        // Pair breaks WITHIN this segment, so a break can never attach to the
        // wrong shift. A blank type on the END row is accepted — rows written
        // before v1.3 have one — and a following BREAK_START ends the search so
        // an unclosed break cannot swallow the next break's end.
        for (let i = 0; i < seg.inner.length; i++) {
          const evt = seg.inner[i];
          if (evt.eventType !== 'BREAK_START') continue;
          let matched = false;
          for (let j = i + 1; j < seg.inner.length; j++) {
            if (seg.inner[j].eventType === 'BREAK_START') break;
            if (seg.inner[j].eventType === 'BREAK_END' &&
                (!seg.inner[j].breakType || seg.inner[j].breakType === evt.breakType)) {
              sessionsRows.push([sid, data.name, data.dept, 'BREAK', evt.breakType,
                Utilities.formatDate(evt.timestamp, tz, 'yyyy-MM-dd HH:mm:ss'),
                Utilities.formatDate(seg.inner[j].timestamp, tz, 'yyyy-MM-dd HH:mm:ss'),
                Math.round((seg.inner[j].timestamp - evt.timestamp) / 60000),
                '', '', Utilities.formatDate(evt.timestamp, tz, 'yyyy-MM-dd'), 'closed']);
              matched = true;
              break;
            }
          }
          if (!matched) unclosedCount++;
        }
      });
    });

    let sh = tab_('Sessions');
    if (!sh) sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Sessions');
    sh.clearContents();
    const header2 = ['SessionID', 'Name', 'Department', 'Type', 'BreakType', 'Start', 'End',
                     'DurationMins', 'Premises', 'Device', 'Date', 'Status'];
    sh.getRange(1, 1, 1, header2.length).setValues([header2]);
    sh.setFrozenRows(1);
    if (sessionsRows.length > 0) {
      sh.getRange(2, 1, sessionsRows.length, header2.length).setValues(sessionsRows);
    }

    audit_('management', 'rebuildSessions', sessionsRows.length + ' rows (' + openCount +
           ' open, ' + unclosedCount + ' unclosed breaks, ' + splitCount + ' split shifts recovered)');
    return { ok: true, sessions: sessionsRows.length, open: openCount,
             unclosed: unclosedCount, splitShifts: splitCount };
  });
}

/** Public wrapper so a time-based trigger can call the rebuild. */
function scheduledRebuild() { rebuildSessions_(); }

/** Install (or reinstall) an hourly Sessions rebuild. Run once from the editor. */
function installSessionsTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scheduledRebuild') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scheduledRebuild').timeBased().everyHours(1).create();
  audit_('management', 'installTrigger', 'Hourly Sessions rebuild installed');
}

// ---------------------------------------------------------------------------
// ABSENCES — holiday from FairLeave (read-only) + sick from Absences tab
// Lookup failures must NEVER block clocking or the register.
// ---------------------------------------------------------------------------
const LEAVE_ALIASES = {
  name:   ['name', 'employee', 'employee name', 'staff', 'staff name', 'full name'],
  empId:  ['employeeid', 'employee id', 'empid', 'staffid', 'staff id', 'person id'],
  from:   ['from', 'start', 'startdate', 'start date', 'fromdate', 'from date', 'first day'],
  to:     ['to', 'end', 'enddate', 'end date', 'todate', 'to date', 'last day'],
  status: ['status', 'state']
};

function aliasMap_(headers, aliases) {
  const map = {};
  headers.forEach(function (h, idx) {
    const clean = String(h).trim().toLowerCase();
    Object.keys(aliases).forEach(function (key) {
      if (map[key] === undefined && aliases[key].indexOf(clean) >= 0) map[key] = idx;
    });
  });
  return map;
}

/**
 * FairLeave employee id → name, read from its Employees tab.
 * Without this a booking cannot be tied to a person on the Clock's Staff tab.
 * Returns {} if the tab or columns are missing — the caller then falls back to
 * whatever name column the bookings themselves carry.
 */
function leaveNameMap_(ss) {
  const map = {};
  const sh = ss.getSheetByName('Employees') || ss.getSheetByName('Staff');
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getDataRange().getValues();
  const hdr = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const idCol = hdr.indexOf('id'), nameCol = hdr.indexOf('name');
  if (idCol < 0 || nameCol < 0) return map;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][idCol] || '').trim();
    const nm = String(data[i][nameCol] || '').trim();
    if (id && nm) map[id] = nm;
  }
  return map;
}

/** name(lower) → 'HOLIDAY' | 'SICK' for today. Never throws. */
function absencesToday_() {
  const out = {};
  const today = new Date(); today.setHours(12, 0, 0, 0);

  // Sick (and any manual absences) from our own Absences tab
  try {
    const sh = tab_('Absences');
    if (sh) {
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const name = String(rows[i][0]).trim();
        if (!name) continue;
        const f = new Date(rows[i][1]); f.setHours(0, 0, 0, 0);
        const t = new Date(rows[i][2] || rows[i][1]); t.setHours(23, 59, 59, 999);
        if (!isNaN(f) && !isNaN(t) && today >= f && today <= t) {
          out[name.toLowerCase()] = String(rows[i][3] || 'SICK').toUpperCase();
        }
      }
    }
  } catch (e) { audit_('system', 'absenceReadFail', String(e)); }

  // Holidays from FairLeave (approved bookings covering today)
  try {
    const id = cfg_('LEAVE_SHEET_ID');
    if (id) {
      const ss = SpreadsheetApp.openById(id);
      const sh = ss.getSheetByName('Bookings') || findLeaveTab_(ss);
      if (sh) {
        const data = sh.getDataRange().getValues();
        const map = aliasMap_(data[0], LEAVE_ALIASES);
        // A name column OR an id column will do — FairLeave has only the id.
        const canName = (map.name !== undefined || map.empId !== undefined);
        if (canName && map.from !== undefined && map.to !== undefined) {
          const idMap = (map.empId !== undefined) ? leaveNameMap_(ss) : {};
          let unresolved = 0;
          for (let r = 1; r < data.length; r++) {
            if (map.status !== undefined &&
                !/approv|active|confirm/i.test(String(data[r][map.status]))) continue;

            let name = map.name !== undefined ? String(data[r][map.name] || '').trim() : '';
            if (!name && map.empId !== undefined) {
              const key = String(data[r][map.empId] || '').trim();
              if (key) { name = idMap[key] || ''; if (!name) unresolved++; }
            }
            if (!name) continue;

            const f = new Date(data[r][map.from]); f.setHours(0, 0, 0, 0);
            const t = new Date(data[r][map.to]);   t.setHours(23, 59, 59, 999);
            if (!isNaN(f) && !isNaN(t) && today >= f && today <= t) {
              if (!out[name.toLowerCase()]) out[name.toLowerCase()] = 'HOLIDAY';
            }
          }
          // Worth knowing about: a booking whose employee id resolves to
          // nobody means the two systems have drifted apart.
          if (unresolved) {
            audit_('system', 'leaveUnresolved',
                   unresolved + ' approved booking(s) had an employee id with no matching FairLeave employee');
          }
        } else {
          audit_('system', 'leaveReadFail',
                 'Bookings tab has no usable columns — detected: ' + JSON.stringify(map));
        }
      }
    }
  } catch (e) { audit_('system', 'leaveReadFail', String(e)); }

  return out;
}

/**
 * Fallback tab finder for the holiday lookup. Reached only when
 * getSheetByName('Bookings') misses.
 *
 * v1.11.0 — this used to REQUIRE a column called `name`. FairLeave's Bookings
 * tab has no name column; it stores an internal `employeeId`. So the fallback
 * could never match the one tab it exists to find, and the holiday lookup
 * worked SOLELY because the exact-name lookup happened to hit. Rename that tab
 * in FairLeave and holidays would stop appearing on every display, silently,
 * with no error anywhere.
 *
 * A person is now identified by name OR by employee id — the same condition
 * absencesToday_ already applies to the rows themselves.
 *
 * Candidates are SCORED rather than taken first-match, because relaxing the
 * test widens what can match. Scoring means a differently-shaped tab that
 * happens to carry a person and two dates cannot quietly win over the real
 * bookings tab.
 */
function findLeaveTab_(ss) {
  var sheets = ss.getSheets();
  var best = null, bestScore = 0;

  for (var i = 0; i < sheets.length; i++) {
    var headers = sheets[i].getLastColumn()
      ? sheets[i].getRange(1, 1, 1, sheets[i].getLastColumn()).getValues()[0] : [];
    var map = aliasMap_(headers, LEAVE_ALIASES);

    // A person, and a date range. Anything less is not a leave booking.
    var hasWho = (map.name !== undefined || map.empId !== undefined);
    if (!hasWho || map.from === undefined || map.to === undefined) continue;

    var score = 1;
    if (/book|leave|holiday|annual|absence/i.test(sheets[i].getName())) score += 4;
    if (map.status !== undefined) score += 2;    // real bookings are approvable
    if (map.empId !== undefined) score += 1;
    if (map.name  !== undefined) score += 1;

    if (score > bestScore) { bestScore = score; best = sheets[i]; }
  }

  if (best) leaveTabFallbackNote_(best.getName());
  return best;
}

/**
 * Records the fallback at most once every six hours.
 *
 * liveStatus_ reaches this path on every office-display refresh, so an audit
 * row per call would bury the Audit tab within a day. Same throttle the weak-PIN
 * nag uses, for the same reason.
 */
function leaveTabFallbackNote_(tabName) {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('clk_leave_fallback')) return;
    cache.put('clk_leave_fallback', '1', 21600);   // 6 h, the cache maximum
    audit_('system', 'leaveTabFallback',
           'No tab called "Bookings" — the holiday lookup fell back to "' + tabName +
           '". Renaming that tab is survivable now, but worth knowing about.');
  } catch (e) { /* never let a diagnostic break the lookup */ }
}

/** Reports what it can see AND whether today's holidays actually resolve. */
function diagnoseLeave_() {
  const id = cfg_('LEAVE_SHEET_ID');
  if (!id) {
    return 'LEAVE_SHEET_ID not set in Config — holiday lookup is OFF.\n\n' +
           'The office display and the emergency register will never show anyone as on holiday.';
  }
  try {
    const ss = SpreadsheetApp.openById(id);
    const lines = ss.getSheets().map(function (sh) {
      const headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
      return sh.getName() + ' | headers: [' + headers.join(', ') + '] | detected: ' +
             JSON.stringify(aliasMap_(headers, LEAVE_ALIASES));
    });
    const idMap = leaveNameMap_(ss);
    const abs = absencesToday_();
    const holidays = Object.keys(abs).filter(function (k) { return abs[k] === 'HOLIDAY'; });
    return lines.join('\n') +
      '\n\nEmployee id → name entries: ' + Object.keys(idMap).length +
      '\nOn holiday today: ' + (holidays.length ? holidays.join(', ') : 'nobody');
  } catch (e) { return 'Cannot open FairLeave sheet: ' + e.message; }
}

function absentStaff_(activeNamesWithEvents) {
  var abs = absencesToday_(), out = [];
  var sh = tab_('Staff'), rows = sh.getDataRange().getValues();
  var c = staffCols_(sh);
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][c.active]).toUpperCase() !== 'TRUE') continue;
    var name = String(rows[i][c.name]).trim();
    var type = abs[name.toLowerCase()];
    if (type && !activeNamesWithEvents[name]) {
      out.push({ name: name, dept: String(rows[i][c.dept] || '(no dept)'), type: type });
    }
  }
  out.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return out;
}

// ---------------------------------------------------------------------------
// LIVE VIEWS
// ---------------------------------------------------------------------------
function liveStatus_(deviceType) {
  const ev = todaysEvents_(), names = {};
  ev.forEach(function (e) { names[e.name] = true; });
  const breaks = [];
  Object.keys(names).forEach(function (n) {
    const st = currentStateFor_(n);
    if (st.status === 'ON_BREAK') {
      const bk = breakByKey_(st.breakType);
      const untimed = !bk || !bk.timed;
      const masked = (deviceType === 'DISPLAY_CANTEEN' && untimed);
      breaks.push({
        name: n,
        dept: (staffRow_(n) || {}).dept || '',
        label: masked ? 'Break' : (bk ? bk.label : 'Break'),
        premises: masked ? '' : st.premises,
        since: st.since,
        allowedMins: (bk && bk.timed) ? bk.mins : null
      });
    }
  });
  breaks.sort(function (a, b) { return new Date(a.since) - new Date(b.since); });
  const res = { ok: true, breaks: breaks, amberPct: Number(cfg_('AMBER_PCT')) || 80,
                serverTime: new Date().toISOString() };
  if (deviceType === 'DISPLAY_OFFICE') res.absences = absentStaff_(names);
  return res;
}

/** In building = clocked in, not out, not on an off-premises break. */
function emergencyRegister_() {
  const ev = todaysEvents_(), names = {};
  ev.forEach(function (e) { names[e.name] = true; });
  const inBuilding = [], offSite = [];
  Object.keys(names).forEach(function (n) {
    const st = currentStateFor_(n);
    if (st.status === 'OUT') return;
    const person = { name: n, dept: (staffRow_(n) || {}).dept || '(no dept)',
                     onBreak: st.status === 'ON_BREAK' };
    if (st.status === 'ON_BREAK' && st.premises === 'OFF') offSite.push(person);
    else inBuilding.push(person);
  });
  const byDept = {};
  inBuilding.forEach(function (p) { (byDept[p.dept] = byDept[p.dept] || []).push(p); });
  Object.keys(byDept).forEach(function (d) {
    byDept[d].sort(function (a, b) { return a.name.localeCompare(b.name); });
  });
  return { ok: true, inBuilding: byDept, count: inBuilding.length, offSite: offSite,
           absent: absentStaff_(names), serverTime: new Date().toISOString() };
}

// ---------------------------------------------------------------------------
// REPORTS — all three read excludedBreaks_(), so they cannot disagree
// ---------------------------------------------------------------------------
function payrollCsv_(from, to) {
  const start = from ? new Date(from) : new Date(new Date().setDate(new Date().getDate() - 14));
  start.setHours(0, 0, 0, 0);
  const end = to ? new Date(to) : new Date();
  end.setHours(23, 59, 59, 999);

  const excluded = excludedBreaks_();
  const rows = tab_('Events').getDataRange().getValues();
  const days = {};

  for (let i = 1; i < rows.length; i++) {
    const ts = new Date(rows[i][0]);
    if (ts < start || ts > end) continue;
    const key = rows[i][1] + '|' + Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const d = days[key] = days[key] ||
      { in: null, out: null, deductedMs: 0, paidBreakMs: 0, bStart: null, bType: '' };
    const type = String(rows[i][2]);
    if (type === 'CLOCK_IN' && !d.in) d.in = ts;
    if (type === 'CLOCK_OUT') d.out = ts;
    if (type === 'BREAK_START') { d.bStart = ts; d.bType = String(rows[i][3] || '').trim(); }
    if (type === 'BREAK_END' && d.bStart) {
      const ms = ts - d.bStart;
      if (excluded[d.bType] === true) d.deductedMs += ms; else d.paidBreakMs += ms;
      d.bStart = null; d.bType = '';
    }
  }

  let csv = 'Name,Date,ClockIn,ClockOut,DeductedBreakMins,PaidBreakMins,NetHours\n';
  Object.keys(days).sort().forEach(function (key) {
    const parts = key.split('|'), d = days[key];
    const fmt = function (t) { return t ? Utilities.formatDate(t, Session.getScriptTimeZone(), 'HH:mm') : ''; };
    const net = (d.in && d.out) ? (((d.out - d.in) - d.deductedMs) / 3600000).toFixed(2) : '';
    csv += [parts[0], parts[1], fmt(d.in), fmt(d.out),
            Math.round(d.deductedMs / 60000), Math.round(d.paidBreakMs / 60000), net].join(',') + '\n';
  });
  return csv;
}

function analytics_(from, to, nameFilter, deptFilter) {
  const start = from ? new Date(from + 'T00:00:00') : new Date(new Date().setDate(new Date().getDate() - 14));
  const end = to ? new Date(to + 'T23:59:59') : new Date();
  const nf = nameFilter ? String(nameFilter).trim().toLowerCase() : '';
  const df = deptFilter ? String(deptFilter).trim().toLowerCase() : '';

  const excluded = excludedBreaks_();

  const sh = tab_('Sessions');
  if (!sh) return { ok: false, error: 'Sessions tab not built yet — run rebuildSessions' };
  const rows = sh.getDataRange().getValues();

  const byDay = {}, breakByType = {}, unknownTypes = {};
  let openRows = 0, unclosedRows = 0;

  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][1]).trim();
    if (!name) continue;
    const dept = String(rows[i][2] || '');
    const type = String(rows[i][3]);
    const start2 = new Date(rows[i][5]);
    if (isNaN(start2) || start2 < start || start2 > end) continue;
    if (nf && name.toLowerCase() !== nf) continue;
    if (df && String(dept).toLowerCase() !== df) continue;

    const status = String(rows[i][11]);
    if (status !== 'closed') { if (status === 'open') openRows++; else unclosedRows++; continue; }

    const mins = Number(rows[i][7]) || 0;
    const date = String(rows[i][10]);
    const k = name + '|' + date;
    const b = byDay[k] = byDay[k] || { name: name, dept: dept, date: date,
      shiftMins: 0, breakMins: 0, deductedBreakMins: 0, paidBreakMins: 0, shifts: 0, breaks: 0 };

    if (type === 'SHIFT') { b.shiftMins += mins; b.shifts++; }
    else if (type === 'BREAK') {
      const bt = String(rows[i][4] || 'BREAK_UNSPECIFIED');
      if (!Object.prototype.hasOwnProperty.call(excluded, bt)) unknownTypes[bt] = true;
      b.breakMins += mins;
      if (excluded[bt] === true) b.deductedBreakMins += mins; else b.paidBreakMins += mins;
      b.breaks++;
      const agg = breakByType[bt] = breakByType[bt] ||
        { breakType: bt, count: 0, totalMins: 0, deducted: excluded[bt] === true };
      agg.count++; agg.totalMins += mins;
    }
  }

  const days = Object.keys(byDay).map(function (k) {
    const b = byDay[k];
    const netMins = Math.max(0, b.shiftMins - b.deductedBreakMins);
    return { name: b.name, dept: b.dept, date: b.date,
             shiftMins: b.shiftMins, breakMins: b.breakMins,
             deductedBreakMins: b.deductedBreakMins, paidBreakMins: b.paidBreakMins,
             netPaidMins: netMins, netPaidHours: +(netMins / 60).toFixed(2),
             shifts: b.shifts, breaks: b.breaks };
  }).sort(function (a, b) { return a.name.localeCompare(b.name) || a.date.localeCompare(b.date); });

  const totals = days.reduce(function (t, d) {
    t.shiftMins += d.shiftMins; t.breakMins += d.breakMins;
    t.deductedBreakMins += d.deductedBreakMins; t.netPaidMins += d.netPaidMins; return t;
  }, { shiftMins: 0, breakMins: 0, deductedBreakMins: 0, netPaidMins: 0 });
  totals.netPaidHours = +(totals.netPaidMins / 60).toFixed(2);

  const breaksList = Object.keys(breakByType).map(function (k) {
    const a = breakByType[k];
    a.avgMins = a.count ? +(a.totalMins / a.count).toFixed(1) : 0;
    return a;
  }).sort(function (a, b) { return b.totalMins - a.totalMins; });

  return {
    ok: true,
    payRule: 'ExcludeTime',
    excludedBreakTypes: Object.keys(excluded).filter(function (k) { return excluded[k]; }),
    range: { from: Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
             to: Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd') },
    days: days, breaks: breaksList, totals: totals,
    warnings: { open: openRows, unclosed: unclosedRows,
                unconfiguredBreakTypes: Object.keys(unknownTypes) },
    serverTime: new Date().toISOString()
  };
}

function overrunReport_(from, to) {
  const start = from ? new Date(from) : new Date(0);
  const end = to ? new Date(to + 'T23:59:59') : new Date();
  const rows = tab_('Overruns').getDataRange().getValues(), out = [];
  for (let i = 1; i < rows.length; i++) {
    const ts = new Date(rows[i][0]);
    if (ts >= start && ts <= end) {
      out.push({ ts: ts.toISOString(), name: rows[i][1], breakType: rows[i][2],
                 allowed: rows[i][3], actual: rows[i][4], overBy: rows[i][5] });
    }
  }
  return { ok: true, overruns: out };
}

// ---------------------------------------------------------------------------
// PAYROLL TAB
// ---------------------------------------------------------------------------
function setupPayrollTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Payroll');
  if (!sh) sh = ss.insertSheet('Payroll');
  const header = ['Employee', 'Week', 'Week Start Date', 'Shift Start', 'Shift End',
                  'Gross Minutes', 'Breaks Deducted', 'Breaks Paid',
                  'Total Deducted Minutes', 'Net Paid Minutes', 'Net Paid Hours', 'Calculation'];
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
}

function rebuildPayroll_() {
  let sh = tab_('Payroll');
  if (!sh) { setupPayrollTab_(); sh = tab_('Payroll'); }

  const sessionsSheet = tab_('Sessions');

  // A MISSING tab is a fault — something is wrong with the workbook.
  if (!sessionsSheet) {
    return { ok: false, error: 'Sessions tab not found. Run rebuildSessions first.' };
  }

  // An EMPTY tab is not. It is the normal state straight after a go-live reset,
  // and every morning until the first person completes a shift.
  //
  // v1.11.0: this used to share the error path above and return WITHOUT
  // clearing, so the previous period's figures sat on the Payroll tab looking
  // like current ones. Wrong figures that look right are worse than an error.
  if (sessionsSheet.getLastRow() <= 1) {
    sh.clearContents();
    setupPayrollTab_();
    audit_('management', 'rebuildPayroll',
           'Sessions is empty — Payroll cleared, 0 rows (not an error)');
    return { ok: true, rows: 0, unknownBreakTypes: [], empty: true };
  }

  const excluded = excludedBreaks_();

  const allSessions = sessionsSheet.getDataRange().getValues();
  const header = allSessions[0];
  const colName = header.indexOf('Name');
  const colType = header.indexOf('Type');
  const colBreakType = header.indexOf('BreakType');
  const colStart = header.indexOf('Start');
  const colEnd = header.indexOf('End');
  const colDurationMins = header.indexOf('DurationMins');
  const colDate = header.indexOf('Date');
  const colStatus = header.indexOf('Status');
  const colDept = header.indexOf('Department');

  if (colName < 0 || colType < 0 || colStart < 0 || colEnd < 0 ||
      colDurationMins < 0 || colDate < 0 || colStatus < 0) {
    return { ok: false, error: 'Sessions tab columns missing or renamed.' };
  }

  const payrollData = {};
  const unknownTypes = {};

  for (let i = 1; i < allSessions.length; i++) {
    const row = allSessions[i];
    if (!row[colName] || String(row[colStatus]).toLowerCase() !== 'closed') continue;

    const name = String(row[colName]).trim();
    const type = String(row[colType]).trim();
    const breakType = String(row[colBreakType] || '').trim();
    const startDate = new Date(row[colStart]);
    const endDate = new Date(row[colEnd]);
    const durationMins = Number(row[colDurationMins]) || 0;
    if (!name || isNaN(startDate)) continue;

    const week = getWeekNumber_(startDate);
    const key = name + '|' + week;

    if (!payrollData[key]) {
      payrollData[key] = {
        name: name, dept: String(row[colDept] || ''), week: week,
        weekStartDate: getWeekStartDate_(startDate),
        shifts: [], breaks: {}
      };
    }

    if (type === 'SHIFT') {
      payrollData[key].shifts.push({ start: startDate, end: endDate, mins: durationMins });
    } else if (type === 'BREAK') {
      // EVERY break type counts. The old version hardcoded BREAK_15 and
      // BREAK_30, so comfort breaks and anything added later were invisible to
      // payroll no matter how they were configured.
      const bt = breakType || 'BREAK_UNSPECIFIED';
      if (!Object.prototype.hasOwnProperty.call(excluded, bt)) unknownTypes[bt] = true;
      const t = payrollData[key].breaks[bt] = payrollData[key].breaks[bt] || { count: 0, mins: 0 };
      t.count++; t.mins += durationMins;
    }
  }

  const payrollRows = [];
  Object.keys(payrollData).sort().forEach(function (key) {
    const data = payrollData[key];
    if (data.shifts.length === 0) return;

    let earliest = data.shifts[0].start, latest = data.shifts[0].end, grossMins = 0;
    data.shifts.forEach(function (s) {
      if (s.start < earliest) earliest = s.start;
      if (s.end > latest) latest = s.end;
      grossMins += s.mins;
    });

    const deductedKeys = [], paidKeys = [];
    let totalDeducted = 0;
    Object.keys(data.breaks).sort().forEach(function (bt) {
      // A break type absent from Config is treated as PAID and flagged.
      // Silently deducting time nobody configured would quietly underpay.
      if (excluded[bt] === true) { deductedKeys.push(bt); totalDeducted += data.breaks[bt].mins; }
      else paidKeys.push(bt);
    });

    const netPaidMins = Math.max(0, grossMins - totalDeducted);
    const netPaidHours = (netPaidMins / 60).toFixed(2);
    let calculation = grossMins + ' min - ' + totalDeducted + ' min = ' +
                      netPaidMins + ' min (' + netPaidHours + ' hrs)';
    const flagged = paidKeys.filter(function (bt) { return excluded[bt] === undefined; });
    if (flagged.length) {
      calculation += '  ⚠ ' + flagged.join(', ') + ' not in Config — treated as paid';
    }

    payrollRows.push([
      data.name,
      'W' + data.week,
      Utilities.formatDate(data.weekStartDate, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(earliest, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      Utilities.formatDate(latest, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      grossMins,
      breakSummary_(data.breaks, deductedKeys),
      breakSummary_(data.breaks, paidKeys),
      totalDeducted,
      netPaidMins,
      netPaidHours,
      calculation
    ]);
  });

  sh.clearContents();
  const header2 = ['Employee', 'Week', 'Week Start Date', 'Shift Start', 'Shift End',
                   'Gross Minutes', 'Breaks Deducted', 'Breaks Paid',
                   'Total Deducted Minutes', 'Net Paid Minutes', 'Net Paid Hours', 'Calculation'];
  sh.getRange(1, 1, 1, header2.length).setValues([header2]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, header2.length).setFontWeight('bold');
  if (payrollRows.length > 0) {
    sh.getRange(2, 1, payrollRows.length, header2.length).setValues(payrollRows);
  }

  const unknownList = Object.keys(unknownTypes);
  audit_('management', 'rebuildPayroll', payrollRows.length + ' rows generated' +
         (unknownList.length ? ' — unconfigured break types treated as paid: ' + unknownList.join(', ') : ''));
  return { ok: true, rows: payrollRows.length, unknownBreakTypes: unknownList };
}

function getWeekNumber_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getWeekStartDate_(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(d.setDate(diff));
}

// ---------------------------------------------------------------------------
// SCHEDULER SYNC
// ---------------------------------------------------------------------------
var HEADER_ALIASES = {
  empId:  ['employeeid', 'emp id', 'empid', 'staffid', 'staff id', 'id'],
  name:   ['name', 'employee', 'employee name', 'staff', 'staff name', 'full name'],
  dept:   ['department', 'dept', 'team', 'section', 'area', 'role'],
  active: ['active', 'enabled', 'status']
};

function headerMap_(headers) {
  const map = {}, unrecognised = [];
  headers.forEach(function (h, idx) {
    const clean = String(h).trim().toLowerCase();
    let hit = false;
    Object.keys(HEADER_ALIASES).forEach(function (key) {
      if (HEADER_ALIASES[key].indexOf(clean) >= 0) { map[key] = idx; hit = true; }
    });
    if (!hit && clean) unrecognised.push(h);
  });
  if (unrecognised.length) audit_('system', 'syncHeaders', 'Unrecognised: ' + unrecognised.join(', '));
  return map;
}

function findStaffTab_(ss) {
  const named = cfg_('SCHEDULER_STAFF_TAB');
  if (named) return ss.getSheetByName(named);
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const map = headerMap_(sheets[i].getRange(1, 1, 1, sheets[i].getLastColumn() || 1).getValues()[0]);
    if (map.name !== undefined) return sheets[i];
  }
  return null;
}

/**
 * Sync staff from the Scheduler, matching on EmployeeID first.
 *
 * Matching by id is what makes a rename survivable: the Clock row is updated
 * in place, keeping the PIN hash, instead of being deactivated and replaced by
 * a nameless new row. Falls back to name so existing rows can be adopted and
 * backfilled the first time this runs.
 */
function syncScheduler_() {
  var id = cfg_('SCHEDULER_SHEET_ID');
  if (!id) return { ok: false, error: 'SCHEDULER_SHEET_ID not set in Config' };
  return withLock_(function () {
    var ss = SpreadsheetApp.openById(id);
    var src = findStaffTab_(ss);
    if (!src) return { ok: false, error: 'Could not find a staff tab — run diagnoseScheduler' };
    var data = src.getDataRange().getValues();
    var map = headerMap_(data[0]);

    var sh = tab_('Staff'), c = staffCols_(sh);
    var existing = sh.getDataRange().getValues();

    var byId = {}, byName = {};
    for (var i = 1; i < existing.length; i++) {
      var nm0 = String(existing[i][c.name]).trim();
      var id0 = c.empId >= 0 ? String(existing[i][c.empId]).trim() : '';
      if (id0) byId[id0] = i + 1;
      if (nm0) byName[nm0.toLowerCase()] = i + 1;
    }

    var added = 0, updated = 0, renamed = 0, deactivated = 0, idsBackfilled = 0;
    var seenActive = {}, seenNames = {}, dupNames = {}, dupIds = {}, seenIds = {};

    for (var r = 1; r < data.length; r++) {
      var name = String(data[r][map.name] || '').trim();
      if (!name) continue;
      var extId = map.empId !== undefined ? String(data[r][map.empId] || '').trim() : '';
      if (map.active !== undefined) {
        var a = String(data[r][map.active]).trim().toLowerCase();
        if (a === 'false' || a === 'no' || a === '0' || a === 'inactive') continue;
      }
      // Two active people sharing a name break the kiosk: staffRow_ returns the
      // first, so the second can never clock in. Detect rather than silently
      // create a row nobody can use.
      var lk = name.toLowerCase();
      if (seenNames[lk]) dupNames[name] = true; else seenNames[lk] = true;
      if (extId) { if (seenIds[extId]) dupIds[extId] = true; else seenIds[extId] = true; }

      // Record BOTH keys. Recording only the id means a row matched by NAME
      // this run — which is every row on the first sync after the upgrade —
      // is invisible to the deactivation pass below, and the whole roster
      // gets switched off.
      if (extId) seenActive['id:' + extId] = true;
      seenActive['nm:' + lk] = true;
      var dept = map.dept !== undefined ? String(data[r][map.dept] || '').trim() : '';

      var row = (extId && byId[extId]) || byName[lk] || 0;
      if (row) {
        sh.getRange(row, c.dept + 1).setValue(dept);
        if (c.empId >= 0 && extId) {
          var storedId = String(existing[row - 1][c.empId] || '').trim();
          if (!storedId) { sh.getRange(row, c.empId + 1).setValue(extId); idsBackfilled++; }
        }
        var storedName = String(existing[row - 1][c.name]).trim();
        if (storedName && storedName !== name) {
          // Matched by id, so this is the same person under a new name. Update
          // in place — the PIN hash and salt stay exactly where they are.
          sh.getRange(row, c.name + 1).setValue(name);
          audit_('management', 'staffRenamed', storedId + ' "' + storedName + '" → "' + name + '" (PIN kept)');
          renamed++;
        }
        updated++;
      } else {
        var blank = [];
        for (var w = 0; w < c.width; w++) blank.push('');
        blank[c.name] = name; blank[c.dept] = dept; blank[c.active] = 'TRUE';
        if (c.empId >= 0) blank[c.empId] = extId;
        sh.appendRow(blank);
        added++;
      }
    }

    // Anyone still TRUE here but no longer active in the Scheduler is switched
    // off. Never deleted: the PIN hash and the exact name every Events row
    // references have to survive, or clocking history stops resolving.
    if (map.active !== undefined && Object.keys(seenActive).length > 0) {
      // RE-READ. The snapshot taken at the top predates the ids backfilled and
      // the names updated above; judging who to deactivate from it switches off
      // people who were just matched successfully.
      var current = sh.getDataRange().getValues();
      for (var k = 1; k < current.length; k++) {
        var nm = String(current[k][c.name]).trim();
        if (!nm || String(current[k][c.active]).toUpperCase() !== 'TRUE') continue;
        var kid = c.empId >= 0 ? String(current[k][c.empId] || '').trim() : '';
        if (seenActive['id:' + kid] || seenActive['nm:' + nm.toLowerCase()]) continue;
        sh.getRange(k + 1, c.active + 1).setValue('FALSE');
        deactivated++;
        audit_('management', 'staffDeactivated', nm + ' — no longer active in the Scheduler');
      }
    }

    var dupN = Object.keys(dupNames), dupI = Object.keys(dupIds);
    if (dupN.length) audit_('system', 'syncDuplicateName',
      'Two or more ACTIVE staff share a name — the second cannot clock in: ' + dupN.join(', '));
    if (dupI.length) audit_('system', 'syncDuplicateId',
      'Duplicate EmployeeID in the Scheduler: ' + dupI.join(', '));

    audit_('management', 'syncScheduler', 'Added ' + added + ', updated ' + updated +
      ', renamed ' + renamed + ', ids backfilled ' + idsBackfilled + ', deactivated ' + deactivated);
    return { ok: true, added: added, updated: updated, renamed: renamed,
             idsBackfilled: idsBackfilled, deactivated: deactivated,
             duplicateNames: dupN, duplicateIds: dupI };
  });
}

function diagnoseScheduler_() {
  const id = cfg_('SCHEDULER_SHEET_ID');
  if (!id) return 'SCHEDULER_SHEET_ID not set in Config';
  const ss = SpreadsheetApp.openById(id);
  return ss.getSheets().map(function (sh) {
    const headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    const map = headerMap_(headers);
    return sh.getName() + ' | headers: [' + headers.join(', ') + '] | detected: ' + JSON.stringify(map);
  }).join('\n');
}

// ---------------------------------------------------------------------------
// PLUMBING
// ---------------------------------------------------------------------------
function tab_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

/** Staff columns resolved by header, falling back to the v1.3 positions. */
function staffCols_(sh) {
  var last = sh.getLastColumn();
  var hdr = last ? sh.getRange(1, 1, 1, last).getValues()[0]
                     .map(function (h) { return String(h).trim().toLowerCase(); }) : [];
  var at = function (names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var k = hdr.indexOf(names[i]);
      if (k >= 0) return k;
    }
    return fallback;
  };
  return {
    name:   at(['name'], 0),
    dept:   at(['department', 'dept'], 1),
    pin:    at(['pinhash'], 2),
    salt:   at(['salt'], 3),
    active: at(['active'], 4),
    empId:  at(['employeeid', 'employee id', 'emp id', 'staffid'], -1),
    width:  Math.max(last, 6)
  };
}

function cfg_(key) {
  const rows = tab_('Config').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][0]) === key) return String(rows[i][1]);
  return '';
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** Public wrapper so the daily trigger can call the sync. */
function scheduledSync() { syncScheduler_(); }

/** Install (or reinstall) a daily staff sync at 5:00 AM. Run once from the editor. */
function installSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scheduledSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scheduledSync').timeBased().atHour(5).everyDays(1).create();
  audit_('management', 'installSyncTrigger', 'Daily sync at 5:00 AM installed');
}

/** Remove the daily sync trigger if needed. */
function removeSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scheduledSync') ScriptApp.deleteTrigger(t);
  });
  audit_('management', 'removeSyncTrigger', 'Daily sync trigger removed');
}

function audit_(actor, action, details) {
  try { tab_('Audit').appendRow([new Date(), actor, action, details]); } catch (e) {}
}

// ---------------------------------------------------------------------------
// EDITOR HELPERS
// ---------------------------------------------------------------------------
function whereAmI() {
  var ss = SpreadsheetApp.getActive();
  Logger.log(ss.getName() + '  |  ' + ss.getId());
}

/** Shows the id ↔ name mapping and anything unmatched. */
function diagnoseStaffIds() {
  var sh = tab_('Staff'), c = staffCols_(sh);
  var rows = sh.getDataRange().getValues();
  var out = ['Staff tab columns: ' + JSON.stringify({
    name: c.name, dept: c.dept, pin: c.pin, salt: c.salt, active: c.active, empId: c.empId })];
  if (c.empId < 0) out.push('⚠ No EmployeeID column — run upgradeToV14()');
  out.push('');
  out.push('id        name            active  pin');
  var noId = 0;
  for (var i = 1; i < rows.length; i++) {
    var nm = String(rows[i][c.name]).trim();
    if (!nm) continue;
    var id = c.empId >= 0 ? String(rows[i][c.empId] || '').trim() : '';
    if (!id) noId++;
    var pad = function (v, n) { v = String(v); while (v.length < n) v += ' '; return v; };
    out.push('  ' + pad(id || '—', 10) + pad(nm, 16) +
             pad(String(rows[i][c.active]), 8) + (rows[i][c.pin] ? 'set' : 'NOT SET'));
  }
  if (noId) out.push('', noId + ' row(s) have no EmployeeID — run a staff sync to backfill.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

function clockDiag() {
  var out = { triggers: [], clockTabs: [], schedulerTabs: [], breakPayRules: {} };
  ScriptApp.getProjectTriggers().forEach(function (t) {
    out.triggers.push(t.getHandlerFunction() + ' | ' + t.getEventType());
  });
  var ss = SpreadsheetApp.getActive();
  out.clockTabs = ss.getSheets().map(function (s) { return s.getName(); });
  try {
    var ex = excludedBreaks_();
    Object.keys(ex).forEach(function (k) { out.breakPayRules[k] = ex[k] ? 'DEDUCTED' : 'paid'; });
  } catch (e) { out.breakPayRules = 'could not read: ' + e.message; }
  try {
    var id = ss.getSheetByName('Config').getDataRange().getValues()
      .filter(function (r) { return String(r[0]).trim() === 'SCHEDULER_SHEET_ID'; })[0][1];
    out.schedulerTabs = SpreadsheetApp.openById(String(id).trim())
      .getSheets().map(function (s) { return s.getName(); });
  } catch (e) { out.schedulerTabs = 'could not open: ' + e.message; }
  Logger.log(JSON.stringify(out, null, 2));
}

/* ---------------------------------------------------------------------------
   ONE-TIME UPGRADE HELPER — dry run for the EmployeeID migration.
   Kept in this file on purpose: it lived in a separate paste and was wiped
   the first time Code.gs was replaced wholesale, which is exactly when it is
   needed. Writes nothing.
   --------------------------------------------------------------------------- */
function previewV14() {
  var out = [], say = function (s) { out.push(s); };
  var pad = function (v, n) { v = String(v); while (v.length < n) v += ' '; return v; };

  // --- Clock side --------------------------------------------------------
  var sh = tab_('Staff');
  if (!sh) { Logger.log('No Staff tab.'); return 'No Staff tab.'; }
  var c = staffCols_(sh);
  var staffRows = sh.getDataRange().getValues();

  say('DRY RUN — nothing is written.');
  say('');
  say('CLOCK Staff tab');
  say('  columns: ' + JSON.stringify({ name: c.name, dept: c.dept, pin: c.pin,
                                       salt: c.salt, active: c.active, empId: c.empId }));
  say('  EmployeeID column: ' + (c.empId >= 0 ? 'present' : 'MISSING — upgradeToV14() will add it'));

  // --- Scheduler side ----------------------------------------------------
  var id = cfg_('SCHEDULER_SHEET_ID');
  if (!id) { say('  SCHEDULER_SHEET_ID not set — cannot preview.'); Logger.log(out.join('\n')); return out.join('\n'); }
  var src, data, map;
  try {
    src = findStaffTab_(SpreadsheetApp.openById(id));
    if (!src) { say('  No staff tab found in the Scheduler.'); Logger.log(out.join('\n')); return out.join('\n'); }
    data = src.getDataRange().getValues();
    map = headerMap_(data[0]);
  } catch (e) {
    say('  Cannot open the Scheduler: ' + e.message);
    Logger.log(out.join('\n')); return out.join('\n');
  }
  say('');
  say('SCHEDULER "' + src.getName() + '"');
  say('  resolved: ' + JSON.stringify(map));
  if (map.empId === undefined) {
    say('  ⚠ No EmployeeID column detected — matching would fall back to NAME only,');
    say('    which is the behaviour v1.4 is meant to replace.');
  }

  // --- Build the same picture the sync would ------------------------------
  var schedActive = [], seenNames = {}, dupNames = {}, seenIds = {}, dupIds = {};
  for (var r = 1; r < data.length; r++) {
    var nm = String(data[r][map.name] || '').trim();
    if (!nm) continue;
    if (map.active !== undefined) {
      var a = String(data[r][map.active]).trim().toLowerCase();
      if (a === 'false' || a === 'no' || a === '0' || a === 'inactive') continue;
    }
    var eid = map.empId !== undefined ? String(data[r][map.empId] || '').trim() : '';
    var lk = nm.toLowerCase();
    if (seenNames[lk]) dupNames[nm] = true; else seenNames[lk] = true;
    if (eid) { if (seenIds[eid]) dupIds[eid] = true; else seenIds[eid] = true; }
    schedActive.push({ id: eid, name: nm, lk: lk });
  }

  var clock = [];
  for (var i = 1; i < staffRows.length; i++) {
    var cn = String(staffRows[i][c.name]).trim();
    if (!cn) continue;
    clock.push({
      row: i + 1, name: cn, lk: cn.toLowerCase(),
      id: c.empId >= 0 ? String(staffRows[i][c.empId] || '').trim() : '',
      active: String(staffRows[i][c.active]).toUpperCase() === 'TRUE',
      hasPin: !!staffRows[i][c.pin]
    });
  }

  // --- What would happen to each CLOCK row --------------------------------
  var backfill = 0, renames = 0, deacts = 0, untouched = 0;
  say('');
  say('EXISTING CLOCK ROWS');
  say('  ' + pad('name', 16) + pad('id now', 9) + pad('action', 34) + 'pin');
  clock.forEach(function (p) {
    var m = null;
    if (p.id) for (var k = 0; k < schedActive.length; k++) if (schedActive[k].id === p.id) { m = schedActive[k]; break; }
    if (!m)   for (var k2 = 0; k2 < schedActive.length; k2++) if (schedActive[k2].lk === p.lk) { m = schedActive[k2]; break; }

    var action;
    if (!m) {
      if (p.active) { action = '→ DEACTIVATED (not in Scheduler)'; deacts++; }
      else          { action = '  already inactive, untouched'; untouched++; }
    } else if (m.name !== p.name) {
      action = '→ RENAMED to "' + m.name + '", PIN kept'; renames++;
    } else if (!p.id && m.id) {
      action = '→ id backfilled: ' + m.id; backfill++;
    } else {
      action = '  no change'; untouched++;
    }
    say('  ' + pad(p.name, 16) + pad(p.id || '—', 9) + pad(action, 34) +
        (p.hasPin ? 'set' : 'NOT SET'));
  });

  // --- Scheduler people with no Clock row ---------------------------------
  var adds = [];
  schedActive.forEach(function (s) {
    var found = clock.some(function (p) {
      return (s.id && p.id === s.id) || p.lk === s.lk;
    });
    if (!found) adds.push(s.id + ' ' + s.name);
  });

  say('');
  say('SUMMARY of what upgradeToV14() would do');
  say('  ids backfilled : ' + backfill);
  say('  renamed        : ' + renames);
  say('  added          : ' + adds.length + (adds.length ? '  (' + adds.join(', ') + ')' : ''));
  say('  DEACTIVATED    : ' + deacts);
  say('  unchanged      : ' + untouched);

  var dn = Object.keys(dupNames), di = Object.keys(dupIds);
  if (dn.length) { say(''); say('  ⚠ Duplicate ACTIVE names in the Scheduler: ' + dn.join(', '));
                   say('    The kiosk cannot handle these — the second person cannot clock in.'); }
  if (di.length) { say('  ⚠ Duplicate EmployeeIDs in the Scheduler: ' + di.join(', ')); }

  say('');
  if (deacts > 0) {
    say('✖ STOP. ' + deacts + ' active staff would be switched off.');
    say('  On a first run every row should match by name, so this should be ZERO.');
    say('  A non-zero figure means names differ between the two sheets — check for');
    say('  spelling, stray spaces, or someone genuinely removed from the Scheduler.');
  } else if (dn.length) {
    say('⚠ Safe to run, but fix the duplicate names in the Scheduler first.');
  } else {
    say('✔ Safe to run. No active staff would be deactivated.');
  }
  say('');
  say('Nothing has been changed. Run upgradeToV14() when you are happy.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ---------------------------------------------------------------------------
   WHICH CODE IS LIVE — run this before debugging any fix that 'doesn't take'.
   Apps Script merges every .gs file and the LAST definition of a duplicated
   name wins, silently. A stale Payroll.gs shadowed rebuildPayroll_ for three
   rounds of debugging before this existed.
   --------------------------------------------------------------------------- */
function codeAudit() {
  var out = [];
  var say = function (s) { out.push(s); };

  var checks = [
    ['rebuildPayroll_',  'excludedBreaks_',
     'v1.3  — deducts per Exclude Time',
     'OLD   — isExcludeTime_, BREAK_15/BREAK_30 only, reads the wrong column'],
    ['rebuildSessions_', 'unclosedCount++',
     'v1.3  — tolerant break pairing, counts unclosed',
     'OLD   — exact type match, kiosk breaks never pair'],
    ['mutate_',          'BREAK_END',
     'v1.3  — stamps BreakType onto BREAK_END',
     'OLD   — writes a blank BreakType on BREAK_END'],
    ['payrollCsv_',      'excludedBreaks_',
     'v1.3  — deducts per Exclude Time (RETIRED v1.10 — unreachable over HTTP)',
     'OLD   — deducts every break, always'],
    ['analytics_',       'excludedBreaks_',
     'v1.3  — deducts per Exclude Time',
     'OLD   — uses the retired PAY_BREAKS'],
    ['absencesToday_',   'leaveNameMap_',
     'v1.3.1 — resolves FairLeave employee ids',
     'OLD   — no id resolution, holidays never show']
  ];

  say('WHICH VERSION IS LIVE');
  var stale = 0;
  checks.forEach(function (c) {
    var name = c[0], marker = c[1];
    var fn = null;
    try { fn = (typeof globalThis !== 'undefined' ? globalThis : this)[name]; } catch (e) {}
    if (typeof fn !== 'function') { say('  MISSING  ' + name); stale++; return; }
    var src = '';
    try { src = String(fn); } catch (e) { src = ''; }
    var isNew = src.indexOf(marker) >= 0;
    if (!isNew) stale++;
    say('  ' + (isNew ? c[2] : c[3]) + '   ← ' + name);
  });

  // A leftover isExcludeTime_ is proof that an old copy is still in the project.
  var hasOld = false;
  try { hasOld = typeof isExcludeTime_ === 'function'; } catch (e) {}
  say('');
  say('isExcludeTime_ still defined: ' + (hasOld ? 'YES — an old copy is still in the project' : 'no'));

  // The Payroll header says which rebuildPayroll_ last wrote it.
  try {
    var ph = tab_('Payroll').getRange(1, 1, 1, 12).getValues()[0];
    say('Payroll header col 7/8   : "' + ph[6] + '" / "' + ph[7] + '"');
    say('  ' + (String(ph[6]) === 'Breaks Deducted'
      ? '→ written by v1.3' : '→ written by the OLD rebuildPayroll_'));
  } catch (e) { say('Payroll header: could not read — ' + e.message); }

  say('');
  if (stale === 0 && !hasOld) {
    say('✔ Everything live is current.');
  } else {
    say('✖ ' + stale + ' function(s) are stale, and/or an old copy is still present.');
    say('');
    say('  Apps Script merges all .gs files, last definition wins. Look at the');
    say('  Files list on the left of the editor. If there is any file besides');
    say('  Code.gs — payroll_function_fixed.gs, Payroll.gs, Code copy.gs —');
    say('  it is almost certainly redefining these. Delete it, or remove the');
    say('  duplicated functions from it, then run this again.');
  }

  Logger.log(out.join('\n'));
  return out.join('\n');
}


// ---------------------------------------------------------------------------
// MANAGEMENT AUTHENTICATION  (v1.6.0)
//
// Before this, both entry points compared the PIN with plain ===.  Three
// separate weaknesses:
//
//   1. === exits at the first differing character, so the time it takes to
//      fail leaks how much of the PIN was right. Guessable one digit at a time.
//   2. No lockout and no rate limit. The endpoint is deployed "Anyone", so a
//      4-digit PIN is 10,000 tries — minutes of scripted requests.
//   3. A failed attempt was never recorded anywhere, so a brute-force attempt
//      left no trace at all.
//
// This mirrors the hardening already in FairLeave v2.3.1 (mgmtAuth_ /
// flConstantEquals_), deliberately using the same shape so the two systems can
// be reasoned about together.
// ---------------------------------------------------------------------------

var CLK_AUTH = {
  PIN_PROP: 'Mgmt_PIN',   // Script Property. NOT the Config tab — see clkMgmtPin_
  MAX_FAILS: 10,          // consecutive failures before the door closes
  LOCKOUT_MINUTES: 15,    // how long it stays closed
  MIN_PIN_LENGTH: 8       // below this, the PIN is reported as weak
};

/** Sentinel left in the Config tab so the old row explains itself. */
var CLK_PIN_MOVED = 'SEE_SCRIPT_PROPERTIES';

/**
 * The management PIN. Script Properties ONLY — never the Config tab.
 *
 * A PIN in a Config cell is readable by anyone with edit access to the
 * workbook, survives in version history, and rides along in every export or
 * copy of the sheet. Script Properties are visible only to someone who can
 * open the Apps Script project.
 *
 * There is deliberately NO fallback to the Config tab. A fallback would mean
 * the sheet value still worked, which is the whole thing being fixed. If the
 * property is missing, management is closed until migrateMgmtPinToProperties()
 * is run — a loud failure, not a quiet one.
 */
function clkMgmtPin_() {
  try {
    return String(PropertiesService.getScriptProperties()
                    .getProperty(CLK_AUTH.PIN_PROP) || '');
  } catch (e) {
    return '';
  }
}

/**
 * The only place a management PIN is checked.
 *
 * A missing PIN is NOT a failed attempt. doPost reaches this line for any
 * action that is not a kiosk or display action, including junk and typos, and
 * counting those would let ordinary noise lock management out.
 */
function clkMgmtAuth_(pin, what) {
  if (pin === undefined || pin === null || String(pin) === '') return false;

  var stored = clkMgmtPin_();
  if (!stored) {
    audit_('system', 'mgmt.noPinSet',
           'Script Property "' + CLK_AUTH.PIN_PROP + '" is not set — management is unreachable. ' +
           'Run migrateMgmtPinToProperties().');
    return false;
  }

  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('clk_mgmt_fails') || 0);
  if (fails >= CLK_AUTH.MAX_FAILS) {
    audit_('unknown', 'mgmt.lockedOut',
           'Management locked for ' + CLK_AUTH.LOCKOUT_MINUTES + ' min after ' +
           CLK_AUTH.MAX_FAILS + ' failures — attempted: ' + (what || '?'));
    return false;
  }

  if (clkConstantEquals_(String(pin), stored)) {
    cache.remove('clk_mgmt_fails');
    // Nag about a weak PIN at most once a day, so the Audit tab stays readable.
    if (clkMgmtPinWeak_() && !cache.get('clk_weak_warned')) {
      cache.put('clk_weak_warned', '1', 21600);   // 6 h, the cache maximum
      audit_('Management', 'mgmt.weakPin',
             CLK_AUTH.PIN_PROP + ' is shorter than ' + CLK_AUTH.MIN_PIN_LENGTH +
             ' characters — lengthen it in Project Settings → Script Properties');
    }
    return true;
  }

  cache.put('clk_mgmt_fails', String(fails + 1), CLK_AUTH.LOCKOUT_MINUTES * 60);
  audit_('unknown', 'mgmt.authFail',
         'Management PIN attempt ' + (fails + 1) + ' of ' + CLK_AUTH.MAX_FAILS +
         ' — action: ' + (what || '?'));
  return false;
}

/**
 * Compares the WHOLE length regardless of where the mismatch is, so the time
 * taken carries no information about how close the guess was.
 * The length is folded into the same accumulator rather than short-circuiting
 * on it, which would leak the PIN's length on its own.
 */
function clkConstantEquals_(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  var n = Math.max(a.length, b.length);
  for (var i = 0; i < n; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

/** True when the PIN is short enough to be worth brute-forcing. */
function clkMgmtPinWeak_() {
  return clkMgmtPin_().length < CLK_AUTH.MIN_PIN_LENGTH;
}

/** Read-only. Reports the state of management auth. Run from the dropdown. */
function clockAuthCheck() {
  var pin = clkMgmtPin_();
  var inSheet = String(cfg_('MGMT_PIN') || '');
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('clk_mgmt_fails') || 0);
  var out = [];

  out.push('ZANNA CLOCK — management authentication');
  out.push('');
  out.push('  PIN source          : Script Property "' + CLK_AUTH.PIN_PROP + '"');
  out.push('  property set        : ' + (pin ? 'yes, ' + pin.length + ' characters' : 'NO'));
  out.push('  minimum recommended : ' + CLK_AUTH.MIN_PIN_LENGTH);
  out.push('  constant-time check : yes');
  out.push('  lockout             : ' + CLK_AUTH.MAX_FAILS + ' failures / ' +
           CLK_AUTH.LOCKOUT_MINUTES + ' minutes');
  out.push('  failures right now  : ' + fails + (fails >= CLK_AUTH.MAX_FAILS ? '  ← LOCKED' : ''));
  out.push('');

  // The Config tab must no longer hold anything that looks like a PIN.
  out.push('CONFIG TAB');
  if (!inSheet) {
    out.push('  MGMT_PIN row        : blank ✔');
  } else if (inSheet === CLK_PIN_MOVED) {
    out.push('  MGMT_PIN row        : "' + CLK_PIN_MOVED + '" ✔  (migrated)');
  } else {
    out.push('  MGMT_PIN row        : ✖ STILL HOLDS A VALUE (' + inSheet.length + ' characters)');
    out.push('    It is no longer used for authentication, but anyone with edit access');
    out.push('    to this workbook can read it — and it is probably your old PIN.');
    out.push('    Run migrateMgmtPinToProperties() to clear it.');
  }
  out.push('');

  if (!pin) {
    out.push('✖ No management PIN is set. Every management action is refused.');
    out.push('  Run migrateMgmtPinToProperties() to move the existing one across,');
    out.push('  or set ' + CLK_AUTH.PIN_PROP + ' by hand in Project Settings.');
  } else if (/^\d+$/.test(pin) && pin.length <= 4) {
    out.push('✖ The PIN is ' + pin.length + ' digits. That is ' +
             Math.pow(10, pin.length) + ' possibilities — the lockout is now the only');
    out.push('  thing standing between the endpoint and a scripted guess. Change it.');
  } else if (clkMgmtPinWeak_()) {
    out.push('⚠ The PIN is shorter than ' + CLK_AUTH.MIN_PIN_LENGTH + ' characters. Lengthen it.');
  } else {
    out.push('✔ PIN length is reasonable and it is stored outside the spreadsheet.');
  }

  out.push('');
  out.push('STILL WORTH KNOWING');
  out.push('  • The payroll CSV export (?action=payroll) was RETIRED in v1.10.0 —');
  out.push('    no management PIN travels in a URL any more. Use the Payroll app.');
  out.push('  • The lockout counter is global, not per-caller — a web app deployed');
  out.push('    "Anyone" has no caller identity to key on. Ten wrong guesses lock');
  out.push('    management out for everyone for ' + CLK_AUTH.LOCKOUT_MINUTES + ' minutes.');
  out.push('    Clear it early with clockAuthUnlock().');
  out.push('  • Failed attempts are written to the Audit tab as mgmt.authFail.');
  out.push('');
  out.push('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/**
 * ONE-TIME. Moves MGMT_PIN out of the Config tab into Script Properties.
 *
 * Run from the Apps Script editor, not the web app — it needs no PIN because
 * only someone who can already open the project can run it.
 *
 * Order matters: the property is written and READ BACK before the sheet cell is
 * cleared, so a failure part-way through never leaves you with no PIN anywhere.
 * The PIN itself is never written to the log.
 */
function migrateMgmtPinToProperties() {
  var props = PropertiesService.getScriptProperties();
  var existing = String(props.getProperty(CLK_AUTH.PIN_PROP) || '');
  var sheetVal = String(cfg_('MGMT_PIN') || '');
  var out = [], say = function (t) { out.push(t); };

  say('MIGRATE  Config tab MGMT_PIN  →  Script Property "' + CLK_AUTH.PIN_PROP + '"');
  say('');

  if (existing && (!sheetVal || sheetVal === CLK_PIN_MOVED)) {
    say('✔ Already migrated. The property holds ' + existing.length + ' characters and');
    say('  the Config tab holds no PIN. Nothing to do.');
    Logger.log(out.join('\n')); return out.join('\n');
  }

  if (existing && sheetVal && sheetVal !== CLK_PIN_MOVED) {
    say('✖ REFUSED — both places hold a value.');
    say('    property   : ' + existing.length + ' characters');
    say('    Config tab : ' + sheetVal.length + ' characters');
    say('');
    say('  I will not guess which one you mean, and overwriting the property could');
    say('  lock you out. Decide, set ' + CLK_AUTH.PIN_PROP + ' by hand in');
    say('  Project Settings, then run this again to clear the sheet.');
    Logger.log(out.join('\n')); return out.join('\n');
  }

  if (!sheetVal || sheetVal === CLK_PIN_MOVED) {
    say('✖ Nothing to migrate: the Config tab has no MGMT_PIN value and the');
    say('  property is not set. Management is currently unreachable.');
    say('');
    say('  Set ' + CLK_AUTH.PIN_PROP + ' by hand:');
    say('    Apps Script → ⚙ Project Settings → Script Properties → Add');
    say('    Use 8+ characters, not 4 digits.');
    Logger.log(out.join('\n')); return out.join('\n');
  }

  // 1. write, 2. read back, 3. only then clear the sheet
  props.setProperty(CLK_AUTH.PIN_PROP, sheetVal);
  var check = String(props.getProperty(CLK_AUTH.PIN_PROP) || '');
  if (check !== sheetVal) {
    say('✖ The property did not save correctly. The Config tab has NOT been touched,');
    say('  so your current PIN still works. Try again.');
    Logger.log(out.join('\n')); return out.join('\n');
  }
  say('  1. Property written and read back  ✔  (' + check.length + ' characters)');

  var sh = tab_('Config');
  var rows = sh.getDataRange().getValues();
  var wrote = false;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'MGMT_PIN') {
      sh.getRange(i + 1, 2).setValue(CLK_PIN_MOVED);
      if (sh.getLastColumn() >= 3) {
        sh.getRange(i + 1, 3).setValue('Moved to Project Settings → Script Properties → ' +
                                       CLK_AUTH.PIN_PROP + '. Do NOT put a PIN here.');
      }
      wrote = true;
      break;
    }
  }
  say('  2. Config tab cell replaced with "' + CLK_PIN_MOVED + '"  ' + (wrote ? '✔' : '— row not found'));

  audit_('management', 'mgmt.pinMigrated',
         'MGMT_PIN moved from the Config tab to Script Property ' + CLK_AUTH.PIN_PROP);

  say('');
  say('✔ Done. The PIN is the same — only where it lives has changed.');
  say('');
  say('IMPORTANT');
  say('  • The value is still in this sheet\'s VERSION HISTORY. If it was ever a');
  say('    real secret, change it now: Project Settings → Script Properties.');
  say('  • It is ' + check.length + ' characters. ' +
      (check.length < CLK_AUTH.MIN_PIN_LENGTH
        ? 'That is below the ' + CLK_AUTH.MIN_PIN_LENGTH + '-character minimum — change it.'
        : 'That is acceptable.'));
  say('  • Run clockAuthCheck() to confirm.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/** Clears the lockout counter after a legitimate lockout. */
function clockAuthUnlock() {
  CacheService.getScriptCache().remove('clk_mgmt_fails');
  audit_('management', 'mgmt.unlock', 'Lockout counter cleared manually');
  var msg = 'Lockout cleared. Management can try again immediately.';
  Logger.log(msg);
  return msg;
}