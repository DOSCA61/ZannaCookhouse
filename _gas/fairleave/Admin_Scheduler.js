/**
 * Admin_Scheduler.gs — Shift Scheduler / Zanna Clock settings for the
 * Data Maintenance console (admin.html v1.2.0).
 *
 * Adds four ops to the existing admin endpoint:
 *   scheduler.settings.get    — read every Config key across the three sheets
 *   scheduler.settings.save   — validate, write back to the row it came from, audit
 *   scheduler.action.run      — sync staff / rebuild sessions / rebuild payroll / install triggers
 *   scheduler.diagnose        — tab names + resolved headers + where each key lives
 *
 * ── WIRING ────────────────────────────────────────────────────────────────
 * Admin.gs v1.2.0 already does both halves of this. If you are wiring it into
 * an older Admin.gs by hand, BOTH are required:
 *
 *   1. In adminHandle(), in the `default:` branch of the op switch:
 *        if (typeof adminSchedulerHandle_ === 'function') {
 *          var s = adminSchedulerHandle_(op, p, 'Management');
 *          if (s) return s;
 *        }
 *        return { ok: false, error: 'unknown_op', op: op };
 *
 *   2. Add to ADMIN_MUTATING:
 *        'scheduler.settings.save', 'scheduler.action.run'
 *
 * Step 2 is not optional. Admin.gs takes the script lock for anything in that
 * array, and this file checks the array to decide whether to lock. Skip it and
 * both files lock — which, since Apps Script locks are not reentrant, hangs
 * every scheduler save until waitLock times out.
 *
 * Redeploy with Manage deployments → ✏️ → New version.
 *
 * ── DESIGN NOTES ──────────────────────────────────────────────────────────
 * • Never assumes schema. Config tab name, header names and column order are all
 *   resolved dynamically; unrecognised shapes are reported, not guessed at.
 * • Every key remembers which sheet it came from, and a save goes back to that
 *   exact row. New keys go to a declared preferred owner — never scattered.
 * • Writes are wrapped in LockService, same as the rest of the system.
 * • Credential-shaped keys are never returned and never written.
 * • A failed audit write never fails the save it was recording.
 */

/* ── Which sheet owns a key we have never seen before ───────────────────── */
var SCHED_PREFERRED_OWNER = {
  BREAK_:                    'clock',
  PAY_BREAKS:                'clock',
  PAID_BREAK_MAX_MINS:       'clock',
  SYNC_HOUR:                 'clock',
  CLOCK_SHEET_ID:            'fairleave',
  // FairLeave, not the clock: adminSchedSS_() reads getConfig_('SCHEDULER_SHEET_ID')
  // from the FairLeave Config tab. Owning it anywhere else would let someone edit
  // this field and change nothing that Admin.gs actually uses.
  SCHEDULER_SHEET_ID:        'fairleave',
  WORK_SATURDAY:             'fairleave',
  NO_SATURDAY_HONOURED:      'scheduler',
  DEFAULT_CONTRACTED_DAYS:   'scheduler',
  DEFAULT_ANNUAL_LEAVE_DAYS: 'scheduler',
  DEFAULT_NO_SATURDAY:       'scheduler',
  EMPLOYEE_ID_PREFIX:        'scheduler',
  EMPLOYEE_ID_DIGITS:        'scheduler'
};

/* ── Server-side validation. The browser checks the same things; this is the
      copy that actually counts, because the browser is not trusted. ───────── */
var SCHED_RULES = {
  PAY_BREAKS:                { oneOf: ['unpaid', 'paid', 'threshold'] },
  PAID_BREAK_MAX_MINS:       { int: true, min: 0,  max: 240 },
  SYNC_HOUR:                 { int: true, min: 0,  max: 23 },
  WORK_SATURDAY:             { bool: true },
  NO_SATURDAY_HONOURED:      { bool: true },
  DEFAULT_NO_SATURDAY:       { bool: true },
  DEFAULT_CONTRACTED_DAYS:   { int: true, min: 0,  max: 7 },
  DEFAULT_ANNUAL_LEAVE_DAYS: { int: true, min: 0,  max: 60 },
  EMPLOYEE_ID_PREFIX:        { re: /^[A-Za-z]{1,4}$/,        msg: 'prefix must be 1–4 letters' },
  EMPLOYEE_ID_DIGITS:        { int: true, min: 1,  max: 6 },
  SCHEDULER_SHEET_ID:        { re: /^[A-Za-z0-9_-]{20,}$/,   msg: 'not a sheet ID — paste the part between /d/ and /edit', allowBlank: true },
  CLOCK_SHEET_ID:            { re: /^[A-Za-z0-9_-]{20,}$/,   msg: 'not a sheet ID', allowBlank: true }
};

// Deliberately broad. A key wrongly treated as a credential is a mild annoyance
// (edit it in the sheet); a credential wrongly treated as an ordinary setting is
// printed into the browser. CLOCK_RELAY_KEY is exactly that case.
var SCHED_SECRET_RE = /(^|_)(PIN|TOKEN|SECRET|PASSWORD|APIKEY|API_KEY|KEY|HASH|SALT|WEBHOOK|CREDENTIAL)(_|$)/i;

/** True if EITHER this file's list or Admin.gs's ADMIN_SECRET_KEY_RE says so.
 *  Two masks that can drift apart is how a credential eventually leaks through
 *  the one that was not updated — so this always takes the union. */
function schedIsSecret_(key) {
  if (SCHED_SECRET_RE.test(key)) return true;
  try { if (typeof ADMIN_SECRET_KEY_RE !== 'undefined' && ADMIN_SECRET_KEY_RE.test(key)) return true; } catch (e) {}
  return false;
}

/** Values starting = + - @ are executed as formulas by Sheets. Admin.gs already
 *  solved this; reuse its helper rather than growing a second one that lags. */
function schedSafe_(v) {
  try { if (typeof adminSafe_ === 'function') return adminSafe_(v); } catch (e) {}
  if (v === null || v === undefined) return '';
  var s = String(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/** Admin.gs takes the script lock for any op listed in ADMIN_MUTATING. Apps
 *  Script script locks are NOT reentrant — taking one twice in a single
 *  execution blocks until waitLock times out and throws. So this file only
 *  locks when it is running WITHOUT that outer lock (standalone, or from the
 *  editor). Adding 'scheduler.settings.save' to ADMIN_MUTATING is what turns
 *  this off; the two must be read together. */
function schedOuterLockHeld_(op) {
  try {
    return typeof ADMIN_MUTATING !== 'undefined' && ADMIN_MUTATING.indexOf(op) > -1;
  } catch (e) { return false; }
}

/* ══ Router ═══════════════════════════════════════════════════════════════ */
function adminSchedulerHandle_(op, p, actor) {
  p = p || {};
  switch (op) {
    case 'scheduler.settings.get':  return schedGet_();
    case 'scheduler.settings.save': return schedSave_(p.settings, actor);
    case 'scheduler.action.run':    return schedRun_(p.job, actor);
    case 'scheduler.diagnose':      return schedDiagnose_();
    default: return null;   // not ours — let Admin.gs carry on
  }
}

/* ══ Sources ══════════════════════════════════════════════════════════════ */
function schedSources_() {
  var own = SpreadsheetApp.getActive();
  var ownCfg = schedReadOne_(own);
  var clockId = (ownCfg.map.CLOCK_SHEET_ID || {}).value || '';
  var schedId = (ownCfg.map.SCHEDULER_SHEET_ID || {}).value || '';

  // Clock first: it owns break/pay config, so its copy of a shared key wins.
  var out = [
    schedOpen_('clock',     'Zanna Clock',     clockId, own),
    schedOpen_('scheduler', 'Shift Scheduler', schedId, null),
    { id: 'fairleave', name: 'FairLeave', ss: own }
  ].filter(function (s) { return !!s.ss; });

  // Label every source with the spreadsheet's REAL Google Sheets filename, not
  // the role we were hoping it would be. Without this, a clock source that fell
  // back to the bound sheet still calls itself "Zanna Clock", and the console
  // tells you a value was "stored in the Zanna Clock Config tab" while it sits
  // in a different file entirely. The role stays on .id; .name is now a fact.
  out.forEach(function (s) {
    try {
      var real = s.ss.getName();
      if (real) s.name = real;
    } catch (e) { /* keep the role name if the file will not identify itself */ }
  });

  // Dedupe by spreadsheet. With CLOCK_SHEET_ID blank the clock falls back to
  // this same sheet — leaving both in would read every key twice, mark all of
  // them "shared", and let two new keys race for the same append row.
  var seenSs = {}, deduped = [];
  out.forEach(function (s) {
    var id;
    try { id = s.ss.getId(); } catch (e) { id = s.id; }
    if (seenSs[id]) { seenSs[id].roles = (seenSs[id].roles || [seenSs[id].id]).concat(s.id); return; }
    seenSs[id] = s;
    deduped.push(s);
  });
  return deduped;
}

function schedOpen_(id, name, sheetId, fallback) {
  // fellBack marks a source that is NOT the spreadsheet it was meant to be.
  // schedGet_ turns this into a visible warning, because silently writing clock
  // settings into the FairLeave sheet looks exactly like a save that did nothing.
  if (!sheetId) {
    return { id: id, name: name, ss: fallback, fellBack: !!fallback,
             error: fallback ? '' : 'not configured' };
  }
  try {
    return { id: id, name: name, ss: SpreadsheetApp.openById(String(sheetId).trim()) };
  } catch (e) {
    // Unreachable is a fact worth surfacing, not a reason to abort the whole read.
    return { id: id, name: name, ss: fallback, fellBack: !!fallback, error: String(e.message || e) };
  }
}

/* Find the Config tab without assuming it is called "Config". */
function schedConfigSheet_(ss) {
  if (!ss) return null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (/^(config|settings|configuration)$/i.test(sheets[i].getName().trim())) return sheets[i];
  }
  return null;
}

/* Resolve Name / Value / Description columns by alias, not by position. */
function schedHeaderMap_(sheet) {
  var last = sheet.getLastColumn();
  if (last < 1) return null;
  var hdr = sheet.getRange(1, 1, 1, last).getValues()[0].map(function (h) {
    return String(h == null ? '' : h).trim().toLowerCase();
  });
  var find = function (aliases) {
    for (var i = 0; i < hdr.length; i++) if (aliases.indexOf(hdr[i]) !== -1) return i + 1;
    return 0;
  };
  var keyCol  = find(['name', 'key', 'setting', 'config', 'parameter', 'property']);
  var valCol  = find(['value', 'val', 'setting value', 'data']);
  var noteCol = find(['description', 'note', 'notes', 'comment', 'purpose']);
  // Legacy Config tabs with no header row at all: A=key, B=value, C=note.
  if (!keyCol && !valCol) { keyCol = 1; valCol = 2; noteCol = 3; }
  if (!keyCol || !valCol) return null;
  return { keyCol: keyCol, valCol: valCol, noteCol: noteCol, headers: hdr };
}

function schedReadOne_(ss) {
  var empty = { map: {}, sheet: null, cols: null, error: '' };
  var sheet = schedConfigSheet_(ss);
  if (!sheet) { empty.error = 'no Config tab'; return empty; }
  var cols = schedHeaderMap_(sheet);
  if (!cols) { empty.error = 'Config tab has no recognisable Name/Value columns'; return empty; }

  var last = sheet.getLastRow();
  var map = {};
  if (last >= 2) {
    var width = Math.max(cols.keyCol, cols.valCol, cols.noteCol || 0);
    var rows = sheet.getRange(2, 1, last - 1, width).getValues();
    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][cols.keyCol - 1] == null ? '' : rows[i][cols.keyCol - 1]).trim();
      if (!key) continue;
      var val = rows[i][cols.valCol - 1];
      map[key] = {
        value: val == null ? '' : String(val).trim(),
        note:  cols.noteCol ? String(rows[i][cols.noteCol - 1] == null ? '' : rows[i][cols.noteCol - 1]).trim() : '',
        row:   i + 2
      };
    }
  }
  return { map: map, sheet: sheet, cols: cols, error: '' };
}

/* ══ GET ══════════════════════════════════════════════════════════════════ */
function schedGet_() {
  var sources = schedSources_();
  var seen = {}, out = [], warnings = [], sheetIds = {};

  sources.forEach(function (src) {
    if (src.error) warnings.push(src.name + ': ' + src.error);
    // One unreadable sheet must not take the whole settings tab down with it.
    var guard = null;
    try { schedReadOne_(src.ss); } catch (e) { guard = String(e.message || e); }
    if (guard) { warnings.push(src.name + ' could not be read: ' + guard); return; }
    // The single most confusing failure this thing can have: you edit a break
    // time, it saves, and the sheet you are looking at never changes — because
    // the clock's settings were never in that sheet to begin with.
    if (src.fellBack && src.id === 'clock') {
      warnings.push('CLOCK_SHEET_ID is not set, so break and pay settings are being read from and written to "' +
        src.name + '" — the spreadsheet this script is bound to. If your Zanna Clock is a different file, ' +
        'set CLOCK_SHEET_ID or nothing you change here will reach it.');
    }
    var r = schedReadOne_(src.ss);
    if (r.error) { warnings.push(src.name + ': ' + r.error); return; }

    Object.keys(r.map).forEach(function (key) {
      var entry = r.map[key];
      if (key === 'SCHEDULER_SHEET_ID') sheetIds[src.id] = entry.value;

      if (seen[key]) {
        // Same key in two sheets. Not an error — the clock and FairLeave each
        // keep their own copy on purpose — but a silent disagreement is a bug
        // that takes days to find, so say so.
        seen[key].shared = true;
        if (seen[key].value !== entry.value && !schedIsSecret_(key)) {
          warnings.push(key + ' differs between ' + seen[key].origin + ' ("' +
            (seen[key].value || '(blank)') + '") and ' + src.name + ' ("' + (entry.value || '(blank)') + '").');
        }
        return;
      }

      var secret = schedIsSecret_(key);
      var rec = {
        key: key,
        value: secret ? '' : entry.value,
        secret: secret,
        note: entry.note,
        origin: src.name,
        originId: src.id,
        shared: false
      };
      seen[key] = rec;
      out.push(rec);
    });
  });

  // Keys the console knows about that no sheet has yet — offered as blanks so
  // they can be set for the first time from here.
  Object.keys(SCHED_PREFERRED_OWNER).forEach(function (k) {
    if (k.slice(-1) === '_' || seen[k]) return;
    var owner = schedOwnerName_(sources, SCHED_PREFERRED_OWNER[k]);
    out.push({ key: k, value: '', secret: false, note: '', origin: owner, originId: SCHED_PREFERRED_OWNER[k], shared: false, isNew: true });
  });

  if (sheetIds.clock && sheetIds.fairleave && sheetIds.clock !== sheetIds.fairleave) {
    warnings.push('The clock and FairLeave point at different Scheduler sheets. One of them is syncing against the wrong roster.');
  }

  out.sort(function (a, b) { return a.key < b.key ? -1 : a.key > b.key ? 1 : 0; });

  return {
    ok: true,
    settings: out,
    source: sources.map(function (s) { return s.name; }).join(' + '),
    timezone: Session.getScriptTimeZone(),
    triggers: schedTriggers_(),
    warning: warnings.join(' '),
    version: (typeof ADMIN_VERSION !== 'undefined') ? ADMIN_VERSION : undefined
  };
}

/**
 * Resolve a ROLE ('clock' | 'scheduler' | 'fairleave') to a surviving source.
 *
 * Dedupe collapses roles that turn out to be the same spreadsheet, and the
 * survivor keeps the first role's id — so a straight `s.id === role` lookup can
 * miss a role that is very much present, just travelling under another name.
 * That produced the worst possible version of this bug: with CLOCK_SHEET_ID
 * unset, 'fairleave' was folded into 'clock', so saving CLOCK_SHEET_ID failed
 * with "no writable Config tab in fairleave" — you could not set the key whose
 * absence caused the collapse.
 */
function schedSourceForRole_(sources, role) {
  var i;
  for (i = 0; i < sources.length; i++) if (sources[i].id === role) return sources[i];
  for (i = 0; i < sources.length; i++) {
    if (sources[i].roles && sources[i].roles.indexOf(role) > -1) return sources[i];
  }
  // Last resort: the spreadsheet this script is bound to always exists and
  // always has a Config tab, or nothing here would have loaded at all.
  var own = null;
  try { own = SpreadsheetApp.getActive().getId(); } catch (e) {}
  for (i = 0; i < sources.length; i++) {
    try { if (own && sources[i].ss.getId() === own) return sources[i]; } catch (e) {}
  }
  return sources.length ? sources[0] : null;
}

function schedOwnerName_(sources, id) {
  var s = schedSourceForRole_(sources, id);
  return s ? s.name : 'this spreadsheet';
}

/* Only this script project's triggers are visible. A job that does not even
   exist in this project cannot be reported as "not installed" — its trigger
   lives in the clock's own project, where we cannot see it. So:
     true      → trigger found here
     false     → the handler exists here but has no trigger (offer to install)
     undefined → not this project's business; the console says "not visible" */
function schedTriggers_() {
  var t = { dailySync: undefined, hourlySessions: undefined };
  var names;
  try {
    names = ScriptApp.getProjectTriggers().map(function (x) { return x.getHandlerFunction(); });
  } catch (e) {
    return t;   // trigger scope not granted — say nothing rather than guess
  }
  var has = function (re) { return names.some(function (n) { return re.test(n); }); };

  if (has(/sync.*(staff|scheduler)|dailySync/i)) t.dailySync = true;
  else if (schedResolveFn_(['syncStaffFromScheduler_', 'syncStaffFromScheduler', 'syncStaff_', 'dailySync_'])) t.dailySync = false;

  if (has(/rebuildSessions|sessionsTrigger/i)) t.hourlySessions = true;
  else if (schedResolveFn_(['rebuildSessions_', 'rebuildSessions'])) t.hourlySessions = false;

  return t;
}

/* Resolve a global function by name without eval. Returns the function or null. */
function schedResolveFn_(names) {
  var g = (typeof globalThis !== 'undefined') ? globalThis : this;
  for (var i = 0; i < names.length; i++) {
    try { if (g && typeof g[names[i]] === 'function') return g[names[i]]; } catch (e) {}
  }
  return null;
}

/* ══ SAVE ═════════════════════════════════════════════════════════════════ */
function schedSave_(changes, actor) {
  if (!changes || !changes.length) return { ok: false, error: 'save_failed', detail: 'nothing to save' };

  var rejected = [];
  changes.forEach(function (c) {
    var msg = schedValidate_(c.key, c.value);
    if (msg) rejected.push(c.key + ': ' + msg);
  });
  // Not 'save_failed': the console's error map would swallow the detail and show
  // a generic "the server refused the save", which tells nobody anything.
  if (rejected.length) return { ok: false, error: 'validation_failed', detail: 'Rejected before anything was written', rejected: rejected };

  var lock = null;
  if (!schedOuterLockHeld_('scheduler.settings.save')) {
    lock = LockService.getScriptLock();
    try { lock.waitLock(20000); }
    catch (e) { return { ok: false, error: 'save_failed', detail: 'the sheet was busy — try again' }; }
  }

  try {
    var sources = schedSources_();
    var readers = {};
    sources.forEach(function (s) { readers[s.id] = { src: s, r: schedReadOne_(s.ss) }; });

    var written = [], failed = [], touchedPay = false, syncHour = null;
    // Do not re-read getLastRow() per new key. Apps Script batches writes, and
    // whether a subsequent getLastRow() reflects an unflushed setValue is not
    // something to bet a Config tab on: if it returns the stale value, two new
    // keys in one save both target the same row and the second silently
    // overwrites the first. A local cursor is correct under either semantics.
    var nextRow = {};

    changes.forEach(function (c) {
      var key = String(c.key || '').trim();
      var value = c.value == null ? '' : String(c.value).trim();
      if (!key) return;
      if (schedIsSecret_(key)) { failed.push(key + ': credentials are not editable here'); return; }

      // TRUE/FALSE are compared as strings by the rules engine, so a lowercase
      // 'true' silently reads as "not TRUE" and flips the working week. Admin.gs
      // learned this the hard way in adminSettingsSave_; same normalisation here.
      if (/^(true|false)$/i.test(value)) value = value.toUpperCase();
      var safeValue = schedSafe_(value);

      // Every sheet that already holds this key gets the new value — not just
      // the first one found. SCHEDULER_SHEET_ID is the case that matters:
      // updating only the clock's copy while adminSchedSS_() reads FairLeave's
      // would look like a successful save that changed nothing.
      var targets = [];
      for (var i = 0; i < sources.length; i++) {
        var rd = readers[sources[i].id];
        if (rd && rd.r.map && rd.r.map[key]) targets.push(sources[i].id);
      }
      if (!targets.length) {
        // New key: send it to its declared owner, resolved through the role map
        // so a deduped role still lands somewhere real.
        var owner = schedSourceForRole_(sources, schedOwnerFor_(key));
        if (owner) targets = [owner.id];
      }

      var okTargets = [], fromVal = null;
      targets.forEach(function (targetId) {
        var rd2 = readers[targetId];
        if (!rd2 || !rd2.r.sheet) {
          var alt = schedSourceForRole_(sources, targetId);
          if (alt) rd2 = readers[alt.id];
        }
        if (!rd2 || !rd2.r.sheet) {
          failed.push(key + ': no writable Config tab in ' + targetId);
          return;
        }
        var existing = rd2.r.map[key];
        var sheet = rd2.r.sheet, cols = rd2.r.cols;
        try {
          if (existing) {
            sheet.getRange(existing.row, cols.valCol).setValue(safeValue);
            if (fromVal === null) fromVal = existing.value;
          } else {
            var cursorKey = targetId + '/' + sheet.getSheetId();
            if (!nextRow[cursorKey]) nextRow[cursorKey] = sheet.getLastRow() + 1;
            var row = nextRow[cursorKey]++;
            sheet.getRange(row, cols.keyCol).setValue(key);
            sheet.getRange(row, cols.valCol).setValue(safeValue);
            if (cols.noteCol) sheet.getRange(row, cols.noteCol).setValue('Added from the admin console');
          }
          okTargets.push(rd2.src.name);
        } catch (e) {
          // One bad row must not abandon the rest of the save half-applied and
          // then throw a non-JSON error at the browser.
          failed.push(key + ' in ' + targetId + ': ' + (e.message || e));
        }
      });
      if (!okTargets.length) return;
      written.push({ key: key, to: okTargets.join(' + '), from: fromVal === null ? '(new)' : fromVal, value: value });
      if (key === 'PAY_BREAKS' || key === 'PAID_BREAK_MAX_MINS' || key.indexOf('BREAK_') === 0) touchedPay = true;
      if (key === 'SYNC_HOUR') syncHour = value;
    });

    SpreadsheetApp.flush();

    written.forEach(function (w) {
      schedAudit_(actor, 'scheduler.setting', w.key + ': "' + w.from + '" → "' + w.value + '" (' + w.to + ')');
    });

    var reinstalled = null;
    if (syncHour !== null) reinstalled = schedReinstallSyncTrigger_(Number(syncHour), actor);

    if (failed.length && !written.length) return { ok: false, error: 'validation_failed', detail: 'Nothing was written', rejected: failed };

    return {
      ok: true,
      saved: written.length,
      rejected: failed.length ? failed : undefined,
      rebuildNeeded: touchedPay || undefined,
      triggerReinstalled: reinstalled || undefined
    };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function schedOwnerFor_(key) {
  if (SCHED_PREFERRED_OWNER[key]) return SCHED_PREFERRED_OWNER[key];
  if (key.indexOf('BREAK_') === 0) return SCHED_PREFERRED_OWNER['BREAK_'];
  return 'fairleave';
}

function schedValidate_(key, value) {
  var v = value == null ? '' : String(value).trim();
  var rule = SCHED_RULES[key];

  if (key.indexOf('BREAK_') === 0 && !rule) {
    if (!/^\d{1,3}$/.test(v)) return 'break minutes must be a whole number 0–999';
    return '';
  }
  if (!rule) return '';                       // unknown key — passthrough, as read
  if (!v) return rule.allowBlank ? '' : (rule.bool || rule.int || rule.oneOf ? 'cannot be blank' : '');

  if (rule.oneOf && rule.oneOf.indexOf(v) === -1) return 'must be one of ' + rule.oneOf.join(', ');
  if (rule.bool && !/^(TRUE|FALSE)$/i.test(v)) return 'must be TRUE or FALSE';
  if (rule.int) {
    if (!/^-?\d+$/.test(v)) return 'must be a whole number';
    var n = Number(v);
    if (rule.min != null && n < rule.min) return 'cannot be below ' + rule.min;
    if (rule.max != null && n > rule.max) return 'cannot be above ' + rule.max;
  }
  if (rule.re && !rule.re.test(v)) return rule.msg || 'is not in the expected format';
  return '';
}

/* ══ ACTIONS ══════════════════════════════════════════════════════════════ */
function schedRun_(job, actor) {
  var JOBS = {
    syncStaff:              ['syncStaffFromScheduler_', 'syncStaffFromScheduler', 'syncStaff_', 'syncStaff'],
    rebuildSessions:        ['rebuildSessions_', 'rebuildSessions'],
    rebuildPayroll:         ['rebuildPayroll_', 'rebuildPayroll'],
    installSyncTrigger:     ['installSyncTrigger'],
    installSessionsTrigger: ['installSessionsTrigger']
  };
  var candidates = JOBS[job];
  if (!candidates) return { ok: false, error: 'job_unavailable', detail: 'Unknown job "' + job + '".' };

  var fn = schedResolveFn_(candidates);
  if (fn) {
    try {
      var res = fn();
      schedAudit_(actor, 'scheduler.run', job + ' run from the admin console');
      return { ok: true, summary: schedSummarise_(job, res), triggers: schedTriggers_() };
    } catch (e) {
      schedAudit_(actor, 'scheduler.run.failed', job + ': ' + (e.message || e));
      return { ok: false, error: 'job_unavailable', detail: job + ' threw: ' + (e.message || e) };
    }
  }

  // Not in this project. The clock's own script may be reachable over HTTP.
  var relayed = schedRelay_(job, actor);
  if (relayed) return relayed;

  return {
    ok: false,
    error: 'job_unavailable',
    detail: 'This script project has no ' + candidates[0] + '. That job lives in the Zanna Clock script. ' +
            'Either run it from the clock sheet menu, or let this console call it by setting three ' +
            'Script Properties on this project (Project Settings → Script Properties): ' +
            'Clock_Exec_Url, Clock_Mgmt_PIN, Clock_Device_Token.'
  };
}

/* ---------------------------------------------------------------------------
   CROSS-PROJECT RELAY TO THE ZANNA CLOCK   (rewritten 2026-08-16)

   Apps Script cannot call another project's functions, so jobs that live in the
   clock have to go over HTTP.

   The previous version posted {action:'relay', job, key} and expected the clock
   to expose a bespoke 'relay' endpoint gated by a shared secret. The clock has
   no such endpoint, so this never worked — and building one would have meant a
   SECOND way into the clock, unauthenticated by device and guarded only by a
   key in a Config cell.

   Instead this uses the clock's EXISTING management API: the same doPost that
   the admin tools already use, behind clkMgmtAuth_ — constant-time comparison,
   ten-failure lockout, every attempt written to the clock's Audit tab. No new
   endpoint, no new attack surface, one code path.

   Three Script Properties on THIS project (Project Settings → Script
   Properties). Deliberately not the Config tab: a Config cell is readable by
   anyone with edit access to the spreadsheet, and this is the clock's
   management PIN.

       Clock_Exec_Url       the clock's /exec URL
       Clock_Mgmt_PIN       the clock's Mgmt_PIN
       Clock_Device_Token   any active token from the clock's Devices tab

   Returns null when it is not configured, so the caller falls through to its
   own "run it from the clock's menu" message.
   --------------------------------------------------------------------------- */

/** Console job name → the action the clock's doPost actually understands. */
var SCHED_CLOCK_JOBS = {
  syncStaff:       'syncScheduler',
  rebuildSessions: 'rebuildSessions',
  rebuildPayroll:  'rebuildPayroll',
  resetClockPin:   'resetPin'          // needs { name: ... } — see schedRelay_'s extra
  // installSyncTrigger and installSessionsTrigger are NOT here. The clock does
  // not expose them over HTTP, and pretending otherwise would report success
  // for something that never ran. Those stay menu-only.
};

/**
 * Build the POST body. Caller-supplied fields go in first and are then
 * overwritten by the three the relay owns, so `extra` can add an argument but
 * can never redirect the call or inject credentials.
 */
function schedPayload_(extra, action, token, pin) {
  var body = {};
  if (extra && typeof extra === 'object') {
    Object.keys(extra).forEach(function (k) { body[k] = extra[k]; });
  }
  body.action = action;
  body.deviceToken = token;
  body.mgmtPin = pin;
  return body;
}

/**
 * Relay one job to the Zanna Clock's management API.
 *
 * `extra` is merged into the POST body for jobs that need an argument — today
 * only resetClockPin, which sends { name: ... }. It is merged FIRST so that
 * action, deviceToken and mgmtPin cannot be overridden by a caller: a console
 * op must never be able to talk the relay into calling a different action or
 * supplying its own PIN.
 */
function schedRelay_(job, actor, extra) {
  var props = PropertiesService.getScriptProperties();
  var url   = String(props.getProperty('Clock_Exec_Url') || '').trim();
  var pin   = String(props.getProperty('Clock_Mgmt_PIN') || '').trim();
  var token = String(props.getProperty('Clock_Device_Token') || '').trim();

  if (!url && !pin && !token) return null;          // not set up at all

  var missing = [];
  if (!url)   missing.push('Clock_Exec_Url');
  if (!pin)   missing.push('Clock_Mgmt_PIN');
  if (!token) missing.push('Clock_Device_Token');
  if (missing.length) {
    return { ok: false, error: 'job_unavailable',
             detail: 'The relay to the clock is half configured. Missing Script ' +
                     'Propert' + (missing.length > 1 ? 'ies' : 'y') + ': ' + missing.join(', ') +
                     '. Set them in Project Settings → Script Properties.' };
  }

  // Check the URL's SHAPE before spending a request on it. A truncated paste
  // returns a bare 404, which reads like the clock is down rather than like a
  // bad setting — and a real /exec URL is ~112 characters, so a short one is
  // detectable without calling anything.
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_\-]{20,}\/exec$/.test(url)) {
    var why = 'Clock_Exec_Url does not look like an Apps Script web app URL (it is ' +
              url.length + ' characters; a real one is about 112).';
    if (/\/dev$/.test(url)) {
      why += ' It ends in /dev — that is the editor-only URL and needs a Google login. Use the /exec one.';
    } else if (!/\/exec$/.test(url)) {
      why += ' It does not end in /exec.';
    } else {
      why += ' It looks truncated — check the whole thing was pasted.';
    }
    return { ok: false, error: 'job_unavailable',
             detail: why + ' Copy it from the clock project: Deploy → Manage deployments → the Web app URL.' };
  }

  var action = SCHED_CLOCK_JOBS[job];
  if (!action) {
    return { ok: false, error: 'job_unavailable',
             detail: '"' + job + '" cannot be run over the relay — the clock does not expose it. ' +
                     'Run it from the clock sheet menu.' };
  }

  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      // text/plain matches what every other client sends and avoids a CORS
      // preflight; the clock parses the body itself either way.
      contentType: 'text/plain;charset=utf-8',
      payload: JSON.stringify(schedPayload_(extra, action, token, pin)),
      muteHttpExceptions: true
    });

    var body = resp.getContentText() || '';
    var data = {};
    try { data = JSON.parse(body); }
    catch (e) {
      // An HTML page here almost always means a Google sign-in wall, i.e. the
      // clock's deployment is not set to "Anyone".
      return { ok: false, error: 'job_unavailable',
               detail: 'The clock returned something that is not JSON (HTTP ' + resp.getResponseCode() +
                       '). Check Clock_Exec_Url ends in /exec and that the clock deployment is ' +
                       'set to "Anyone".' };
    }

    if (data && data.ok) {
      schedAudit_(actor, 'scheduler.run', job + ' relayed to the clock (' + action + ')');
      return { ok: true, summary: schedRelaySummary_(job, data) };
    }

    schedAudit_(actor, 'scheduler.run.failed', job + ' relay refused: ' + (data.error || resp.getResponseCode()));
    var hint = '';
    if (/permission|unknown action/i.test(String(data.error || ''))) {
      hint = ' That message means either the PIN is wrong, or the clock script is older than the ' +
             'action being called. Check Clock_Mgmt_PIN, and that the clock is deployed.';
    } else if (/device/i.test(String(data.error || ''))) {
      hint = ' Clock_Device_Token does not match an ACTIVE row in the clock\'s Devices tab.';
    }
    return { ok: false, error: 'job_unavailable',
             detail: 'The clock refused: ' + (data.error || ('HTTP ' + resp.getResponseCode())) + '.' + hint };

  } catch (e) {
    var m = String(e.message || e);
    // A missing OAuth scope is not a network problem, and calling it one sends
    // people looking at URLs and firewalls. Name it.
    if (/permission to call UrlFetchApp|script\.external_request|authoriz/i.test(m)) {
      return { ok: false, error: 'not_authorised',
               detail: 'This project is not allowed to make outbound requests yet. That is a ' +
                       'permission, not a connection problem.\n\n' +
                       'Project Settings → tick "Show appsscript.json manifest file", open ' +
                       'appsscript.json, and add this to oauthScopes:\n' +
                       '    "https://www.googleapis.com/auth/script.external_request"\n\n' +
                       'Save, run authorizeClockRelay() and approve the prompt, then deploy a ' +
                       'new version.' };
    }
    return { ok: false, error: 'job_unavailable', detail: 'Could not reach the clock: ' + m };
  }
}

/**
 * RUN THIS ONCE FROM THE EDITOR after adding the relay.
 *
 * Apps Script decides which permissions a project needs by reading the code,
 * and only asks for them when a function is run from the editor. Before this
 * file existed nothing here made an outbound request, so the project — and the
 * deployed web app running under it — has no external_request permission. The
 * console reports that as "You do not have permission to call
 * UrlFetchApp.fetch", which sounds like a fault and is really just consent
 * that was never asked for.
 *
 * Pick authorizeClockRelay from the function dropdown and press Run. Approve
 * the prompt. Then deploy a new version so the web app runs under the refreshed
 * authorisation.
 *
 * It performs a real syncStaff, so it doubles as a test: whatever it prints is
 * exactly what the console's "Sync to Clock" button will say.
 */
function authorizeClockRelay() {
  var props = PropertiesService.getScriptProperties();
  var out = [];
  var say = function (t) { out.push(t); };

  say('CLOCK RELAY — authorise and test');
  say('');
  ['Clock_Exec_Url', 'Clock_Mgmt_PIN', 'Clock_Device_Token'].forEach(function (k) {
    var v = String(props.getProperty(k) || '');
    // Never print the values. The URL is not secret but the PIN is, and one
    // habit is easier to keep than an exception.
    say('  ' + k + (v ? '  set (' + v.length + ' chars)' : '  NOT SET'));
  });
  say('');

  var r;
  try {
    r = schedRelay_('syncStaff', 'Management');
  } catch (e) {
    say('✖ ' + (e.message || e));
    if (/permission|authoriz/i.test(String(e.message || e))) {
      say('');
      say('  That is the consent prompt failing or being declined. Run this again');
      say('  and approve it. If no prompt appears at all, open Project Settings,');
      say('  tick "Show appsscript.json manifest file", and check oauthScopes —');
      say('  if that list is present it must include:');
      say('      https://www.googleapis.com/auth/script.external_request');
    }
    Logger.log(out.join('\n'));
    return out.join('\n');
  }

  if (r === null) {
    say('🔒 Not configured — set the three Script Properties above first.');
    say('   (Project Settings → Script Properties)');
  } else if (r.ok) {
    say('✔ ' + r.summary);
    say('');
    say('  The permission is granted and the relay works. Now deploy a NEW');
    say('  VERSION so the web app runs under this authorisation, or the console');
    say('  button will keep failing while this function succeeds.');
  } else {
    say('✖ ' + (r.detail || r.error));
  }
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/** Turn the clock's own reply into something worth reading. */
function schedRelaySummary_(job, data) {
  if (job === 'resetClockPin') {
    return 'Kiosk PIN cleared. They choose a new one themselves the next time ' +
           'they tap their name — nothing to issue and nothing to tell them ' +
           'beyond "set it again".';
  }
  if (job === 'syncStaff') {
    var bits = ['Staff synced to the clock.'];
    if (data.added !== undefined)        bits.push('Added ' + data.added);
    if (data.updated !== undefined)      bits.push('updated ' + data.updated);
    if (data.renamed)                    bits.push('renamed ' + data.renamed);
    if (data.idsBackfilled)              bits.push('ids backfilled ' + data.idsBackfilled);
    bits.push('deactivated ' + (data.deactivated || 0) + '.');
    var msg = bits.join(' ');
    if (data.duplicateNames && data.duplicateNames.length) {
      msg += ' ⚠ Duplicate active names: ' + data.duplicateNames.join(', ') +
             ' — the second person cannot clock in.';
    }
    return msg;
  }
  if (job === 'rebuildSessions') {
    return 'Sessions rebuilt on the clock: ' + (data.sessions || 0) + ' rows' +
           (data.open ? ', ' + data.open + ' still open' : '') +
           (data.unclosed ? ', ' + data.unclosed + ' unclosed break(s)' : '') + '.';
  }
  if (job === 'rebuildPayroll') {
    return 'Payroll rebuilt on the clock: ' + (data.rows || 0) + ' rows' +
           (data.unknownBreakTypes && data.unknownBreakTypes.length
             ? '. ⚠ Break type(s) not in Config, treated as paid: ' + data.unknownBreakTypes.join(', ')
             : '.');
  }
  return data.summary || (job + ' completed on the clock.');
}

function schedSummarise_(job, res) {
  if (res && typeof res === 'object' && res.summary) return res.summary;
  if (typeof res === 'string' && res) return res;
  var m = { syncStaff: 'Staff synced to the clock.', rebuildSessions: 'Sessions rebuilt.',
            rebuildPayroll: 'Payroll rebuilt.', installSyncTrigger: 'Daily sync trigger installed.',
            installSessionsTrigger: 'Hourly sessions trigger installed.' };
  return m[job] || 'Done.';
}

function schedReinstallSyncTrigger_(hour, actor) {
  if (!(hour >= 0 && hour <= 23)) return null;

  var target = null, existing = [];
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (/sync.*(staff|scheduler)|dailySync/i.test(t.getHandlerFunction())) {
        target = t.getHandlerFunction();
        existing.push(t);
      }
    });
  } catch (e) { return null; }
  if (!target) return null;   // nothing here to move — the clock owns its own

  // Create first, delete second. The other order leaves the business with no
  // overnight staff sync at all if the create fails.
  try {
    ScriptApp.newTrigger(target).timeBased().atHour(hour).everyDays(1).create();
  } catch (e) {
    schedAudit_(actor, 'scheduler.trigger.failed',
      'SYNC_HOUR saved as ' + hour + ' but the trigger could not be recreated: ' + (e.message || e) +
      ' — the old trigger is still in place.');
    return null;
  }
  existing.forEach(function (t) { try { ScriptApp.deleteTrigger(t); } catch (e) {} });
  schedAudit_(actor, 'scheduler.trigger', target + ' moved to ' + hour + ':00');
  return hour + ':00';
}

/* ══ DIAGNOSE ═════════════════════════════════════════════════════════════ */
function schedDiagnose_() {
  var out = { timezone: Session.getScriptTimeZone(), sources: [] };
  schedSources_().forEach(function (src) {
    var entry = { id: src.id, name: src.name, error: src.error || '' };
    try {
      entry.spreadsheet = src.ss ? src.ss.getName() : '(none)';
      entry.tabs = src.ss ? src.ss.getSheets().map(function (s) { return s.getName(); }) : [];
      var r = schedReadOne_(src.ss);
      entry.configTab = r.sheet ? r.sheet.getName() : null;
      entry.configHeaders = r.cols ? r.cols.headers : null;
      entry.resolvedColumns = r.cols ? { key: r.cols.keyCol, value: r.cols.valCol, note: r.cols.noteCol } : null;
      entry.keys = Object.keys(r.map).sort();
      entry.readError = r.error || '';
    } catch (e) { entry.error = String(e.message || e); }
    out.sources.push(entry);
  });
  // Guarded, because this is the function people run WHEN THINGS ARE BROKEN.
  // ScriptApp needs the script.scriptapp OAuth scope, which an existing
  // deployment may never have been authorised for — and a diagnostic that dies
  // on its last line, taking the useful 95% of its output with it, is worse
  // than no diagnostic at all.
  try {
    out.triggers = ScriptApp.getProjectTriggers().map(function (t) {
      return t.getHandlerFunction() + ' (' + t.getEventType() + ')';
    });
  } catch (e) {
    out.triggers = 'unavailable — ' + String(e.message || e).split('.')[0] +
      '. Trigger display and the Install buttons stay switched off; everything ' +
      'else on the Scheduler tab works. To enable them, open the Apps Script ' +
      'editor, run runAdminDiagnostic() once and accept the permission prompt, ' +
      'then redeploy as a new version.';
  }
  return { ok: true, diagnostic: out };
}

/* ══ AUDIT ════════════════════════════════════════════════════════════════ */
function schedAudit_(actor, action, details) {
  try {
    if (typeof audit_ === 'function') { audit_(actor, action, details); return; }
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName('Audit');
    if (!sh) return;
    sh.appendRow([new Date(), actor || 'admin', action, details]);
  } catch (e) {
    // An audit failure must never take the save down with it.
    console.error('audit write failed: ' + (e.message || e));
  }
}
