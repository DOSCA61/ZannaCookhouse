/**********************************************************************
 * Admin.gs — Zanna Cookhouse Control Panel · data maintenance backend
 * Companion to FairLeave v2 Code.gs. Paste as a NEW file in the same
 * Apps Script project. Requires ONE line added to doPost (see below).
 *
 * DESIGN: this file deliberately owns almost no logic. It authenticates
 * the console, then delegates to the functions Code.gs already has —
 * route_(), syncFromScheduler_(), readAll_(), addRow_(), updateRow_(),
 * schedTab_(), headerMap_(). That way employee removal still cancels
 * bookings and fails swaps, team deletion still refuses to orphan
 * members, and the rules engine stays the single source of behaviour.
 * Duplicating that here would have quietly diverged from it.
 *
 * WHAT IS ACTUALLY NEW HERE:
 *   - Admin PIN → short-lived session (Code.gs re-checks MGMT_PIN per call)
 *   - Writing employees INTO the Shift Scheduler sheet, then syncing back,
 *     so the Scheduler stays master (locked decision) instead of FairLeave
 *     growing a second, unsynced employee list
 *   - Panel_Apps / Access_Token editing for the control panel tiles
 *   - A read-only Audit view and a browser-side diagnostic
 *
 * v1.2.0 — the console's Scheduler tab. The ops themselves live in
 * Admin_Scheduler.gs; this file only routes to them (see the `default:`
 * branch) and declares which of them mutate, so the lock is taken here
 * and only here. Admin_Scheduler.gs reads that declaration and skips its
 * own lock accordingly — Apps Script script locks are NOT reentrant, and
 * taking one twice in a single execution deadlocks until it times out.
 * If Admin_Scheduler.gs is absent the routing is skipped cleanly and the
 * console shows its "not deployed" banner instead of a server error.
 *
 * v1.3.0 - Reset kiosk PIN, from the Edit employee screen. Previously the only
 * way was to clear PinHash and Salt by hand on the CLOCK's Staff tab - a
 * different spreadsheet, and a job nobody would find without the manual. It is
 * relayed to the clock, so it needs the same three Script Properties as
 * Sync to Clock, and it writes nothing in FairLeave.
 *
 * v1.2.1 - adminNextExtId_ now preserves the sheet's own ID scheme. It used
 * to strip the prefix and the padding, so "+ Add employee" on a roster of
 * E001-E015 generated "16" rather than "E016". Nothing else changed.
 *
 * Script Property required:
 *   Admin_PIN   — 10+ characters. Not a 4-digit PIN.
 *
 * NOT required: the scheduler ID. Code.gs already keeps it in the Config
 * tab as SCHEDULER_SHEET_ID. This file reads it from there, and only
 * falls back to a Script Property if Config is somehow empty.
 *
 * On the lockout: doPost gets no reliable client IP, so the failed-attempt
 * counter has to be global — anyone with the /exec URL could trip it
 * deliberately. It is therefore tuned loose (12 tries per 5 minutes):
 * enough to stop guessing, not enough to be a usable denial-of-service.
 * PIN LENGTH is the real control.
 **********************************************************************/

/**
 * Bump this on every change to this file. It is returned by auth, ping and
 * diagnose, and shown in the console footer — so "is the deployed version the
 * one I just pasted?" is answered by looking, not by assuming.
 */
var ADMIN_VERSION = '1.3.0 (2026-08-17)';

var ADMIN_CFG = {
  PIN_PROP:        'Admin_PIN',
  TOKEN_PROP:      'Access_Token',
  PANEL_PROP:      'Panel_Apps',
  SESSION_MINUTES: 30,
  MAX_FAILS:       12,
  LOCKOUT_MINUTES: 5,
  MIN_PIN_LENGTH:  10
};

/** Config keys never shown to, or writable from, the browser. One definition,
 *  used by both the read mask and the write guard, so they cannot drift.
 *  Admin_Scheduler.gs unions its own list with this one for the same reason. */
var ADMIN_SECRET_KEY_RE = /pin|token|secret|password|apikey|key$/i;

/** Ops that change data — these take the script lock, exactly as doPost does.
 *  Admin_Scheduler.gs checks this array to decide whether to lock internally,
 *  so adding a scheduler op here is what STOPS it double-locking. Do not
 *  remove the two scheduler entries without changing that file too. */
var ADMIN_MUTATING = [
  'employees.save', 'employees.archive', 'employees.regenToken', 'employees.sync',
  'leave.import', 'teams.save', 'teams.delete', 'settings.save',
  'blackouts.save', 'blackouts.delete', 'panel.save', 'panel.rotateToken',
  'scheduler.settings.save', 'scheduler.action.run'
];

/* ═══════════════════════════════════════════════════════════════════
   ENTRY POINT
   In doPost, immediately after the body is parsed into _p:
     if (_p.action === 'admin') return adminOut_(adminHandle(_p));
   ═══════════════════════════════════════════════════════════════════ */
function adminHandle(p) {
  var lock = null;
  try {
    var op = String(p.op || '');

    if (op === 'auth') return adminAuth_(String(p.pin || ''));

    // Sign out has to work on an already-dead session, so it sits above the check.
    if (op === 'signOut') {
      CacheService.getScriptCache().remove('admin_sess_' + String(p.session || ''));
      return { ok: true };
    }

    var actor = adminSession_(String(p.session || ''));
    if (!actor) return { ok: false, error: 'session_expired' };

    if (ADMIN_MUTATING.indexOf(op) > -1) {
      lock = LockService.getScriptLock();
      lock.waitLock(15000);
    }

    switch (op) {
      case 'ping':                 return { ok: true, actor: actor, version: ADMIN_VERSION };
      case 'employees.list':       return adminEmployeesList_();
      case 'employees.save':       return adminEmployeeSave_(p.employee);
      case 'employees.archive':    return adminEmployeeArchive_(p.id);
      case 'employees.regenToken': return adminRegenToken_(p.id);
      case 'employees.resetClockPin': return adminResetClockPin_(p.emp, actor);
      case 'employees.sync':       return adminSync_();
      case 'leave.import':         return adminImportLeave_();
      case 'teams.list':           return adminTeamsList_();
      case 'teams.save':           return adminTeamSave_(p.team);
      case 'teams.delete':         return adminTeamDelete_(p.id);
      case 'settings.get':         return adminSettingsGet_();
      case 'settings.save':        return adminSettingsSave_(p.settings);
      case 'blackouts.list':       return { ok: true, blackouts: readAll_('Blackouts') };
      case 'blackouts.save':       return adminBlackoutSave_(p.blackout);
      case 'blackouts.delete':     return adminBlackoutDelete_(p.id);
      case 'panel.get':            return adminPanelGet_();
      case 'panel.save':           return adminPanelSave_(p.apps);
      case 'panel.rotateToken':    return adminRotateToken_(p.token);
      case 'audit.list':           return adminAuditList_(p.limit);
      case 'diagnose':             return adminDiagnose_();
      default:
        // Scheduler tab ops (Admin_Scheduler.gs). Routed from `default:` so it
        // can never shadow an op above, and guarded by typeof so a project
        // without that file answers 'unknown_op' — which the console reads as
        // "not deployed yet" — rather than throwing a server_error that looks
        // like a broken backend.
        if (typeof adminSchedulerHandle_ === 'function') {
          var sched = adminSchedulerHandle_(op, p, 'Management');
          if (sched) return sched;
        }
        return { ok: false, error: 'unknown_op', op: op };
    }
  } catch (err) {
    return { ok: false, error: 'server_error', detail: String(err && err.message || err) };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function adminOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════════════════════════════════════════════════════════════
   AUTH
   ═══════════════════════════════════════════════════════════════════ */
function adminAuth_(pin) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('admin_fails') || 0);
  if (fails >= ADMIN_CFG.MAX_FAILS) {
    return { ok: false, error: 'locked_out', minutes: ADMIN_CFG.LOCKOUT_MINUTES };
  }

  var stored = String(PropertiesService.getScriptProperties().getProperty(ADMIN_CFG.PIN_PROP) || '');
  if (!stored) return { ok: false, error: 'not_configured' };

  if (!adminConstantEquals_(pin, stored)) {
    cache.put('admin_fails', String(fails + 1), ADMIN_CFG.LOCKOUT_MINUTES * 60);
    audit_('unknown', 'admin.authFail', 'Console sign-in attempt ' + (fails + 1));
    return { ok: false, error: 'bad_pin', version: ADMIN_VERSION, remaining: Math.max(0, ADMIN_CFG.MAX_FAILS - fails - 1) };
  }

  cache.remove('admin_fails');
  var session = Utilities.getUuid().replace(/-/g, '');
  cache.put('admin_sess_' + session, 'admin', ADMIN_CFG.SESSION_MINUTES * 60);
  audit_('Management', 'admin.signIn', 'Data maintenance console opened');

  return {
    ok: true, session: session,
    version: ADMIN_VERSION,
    expiresIn: ADMIN_CFG.SESSION_MINUTES * 60,
    weakPin: stored.length < ADMIN_CFG.MIN_PIN_LENGTH
  };
}

function adminSession_(session) {
  if (!session) return null;
  var cache = CacheService.getScriptCache();
  var who = cache.get('admin_sess_' + session);
  if (!who) return null;
  cache.put('admin_sess_' + session, who, ADMIN_CFG.SESSION_MINUTES * 60); // sliding
  return who;
}

/** Compares the whole length regardless of mismatch position — no early-exit signal. */
function adminConstantEquals_(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

/**
 * Run an existing Code.gs management action with its real side effects.
 * The MGMT PIN never leaves the server — it is read from Config and handed
 * straight to route_(), which is the same path the FairLeave app itself uses.
 */
function adminAsMgmt_(action, req) {
  req = req || {};
  req.action = action;
  req.pin = getConfig_('MGMT_PIN');
  var out = route_(action, req);
  if (out && out.state) delete out.state;   // the console fetches what it needs itself
  return out;
}

/* ═══════════════════════════════════════════════════════════════════
   SCHEDULER ACCESS
   Alias lists below intentionally mirror syncFromScheduler_() exactly,
   so anything this console writes is something that sync can read back.
   ═══════════════════════════════════════════════════════════════════ */
function adminSchedSS_() {
  var id = getConfig_('SCHEDULER_SHEET_ID') ||
           PropertiesService.getScriptProperties().getProperty('Scheduler_Sheet_Id') || '';
  if (!id) throw new Error('SCHEDULER_SHEET_ID is not set in the Config tab.');
  try { return SpreadsheetApp.openById(String(id).trim()); }
  catch (e) { throw new Error('Cannot open the scheduler sheet: ' + e.message); }
}

function adminSchedEmpSheet_() {
  var sh = schedTab_(adminSchedSS_(), ['Employees', 'Staff', 'People']);
  if (!sh) throw new Error('No Employees tab found in the scheduler sheet.');
  return sh;
}

function adminSchedCols_(sh) {
  var H = headerMap_(sh);
  return {
    H: H,
    id:     H.find(['EmployeeID', 'EmpID', 'ID', 'StaffID']),
    name:   H.find(['Name', 'EmployeeName', 'FullName', 'Employee']),
    dept:   H.find(['Department', 'Dept', 'Team', 'Section']),
    active: H.find(['Active', 'IsActive', 'Enabled']),
    ent:    H.find(['AnnualLeaveEntitlement', 'AnnualLeave', 'Entitlement', 'ALDays', 'LeaveDays', 'HolidayEntitlement']),
    email:  H.find(['Email', 'EmailAddress', 'Mail']),
    role:   H.find(['Role', 'JobTitle', 'Position', 'Title']),
    start:  H.find(['StartDate', 'Joined', 'JoinDate', 'HireDate', 'DateStarted'])
  };
}

function adminSchedRows_(sh) {
  var last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues() : [];
}

/**
 * Neutralise spreadsheet formula injection. A value starting = + - @ is executed
 * as a formula by Sheets, so a name like =HYPERLINK("http://evil","Payslip") would
 * go live in a sheet a manager opens. A leading apostrophe forces text.
 */
function adminSafe_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/* ═══════════════════════════════════════════════════════════════════
   EMPLOYEES — Scheduler is master; FairLeave holds leave-specific data
   ═══════════════════════════════════════════════════════════════════ */
function adminEmployeesList_() {
  var fl = readAll_('Employees');                       // includes archived
  var teamName = {};
  readAll_('Teams').forEach(function (t) { teamName[t.id] = t.name; });

  var out = [], schedWarning = null, seenExt = {};
  var entOwner = 'fairleave', emailOwner = 'fairleave';
  var available = { email: false, role: false, start: false };

  fl.forEach(function (e) {
    if (e.extId) seenExt[e.extId] = true;
    out.push({
      id: e.id,
      extId: e.extId || '',
      name: e.name,
      teamId: e.teamId || '',
      team: teamName[e.teamId] || '',
      email: e.email || '',
      entitlement: parseInt(e.entitlement, 10) || 0,
      used: usedDays_(e.id, []),
      active: String(e.active) !== 'FALSE',
      hasToken: !!e.token,
      linked: !!e.extId,
      pendingImport: false
    });
  });

  // Anyone in the Scheduler who has not been synced into FairLeave yet.
  try {
    var sh = adminSchedEmpSheet_();
    var c = adminSchedCols_(sh);
    if (c.ent >= 0) entOwner = 'scheduler';
    if (c.email >= 0) { emailOwner = 'scheduler'; available.email = true; }
    available.role = c.role >= 0;
    available.start = c.start >= 0;
    if (c.id < 0 || c.name < 0) {
      schedWarning = 'The scheduler Employees tab has no recognised ID or Name column. Headers: ' + c.H.raw.join(', ');
    } else {
      adminSchedRows_(sh).forEach(function (r) {
        var ext = String(r[c.id]).trim();
        var nm = String(r[c.name]).trim();
        if (!ext || !nm || seenExt[ext]) return;
        out.push({
          id: '', extId: ext, name: nm,
          teamId: '', team: c.dept >= 0 ? String(r[c.dept]).trim() : '',
          email: c.email >= 0 ? String(r[c.email]).trim() : '',
          entitlement: c.ent >= 0 ? (parseInt(r[c.ent], 10) || 0) : 0,
          used: 0, active: true, hasToken: false,
          linked: true, pendingImport: true
        });
      });
    }
  } catch (err) {
    schedWarning = String(err.message || err);
  }

  return {
    ok: true,
    employees: out,
    teams: readAll_('Teams').map(function (t) { return { id: t.id, name: t.name }; }),
    entitlementOwner: entOwner,
    emailOwner: emailOwner,
    // Fields the Scheduler has no column for. The console hides these rather
    // than showing inputs whose values would go nowhere.
    available: available,
    schedWarning: schedWarning
  };
}

/**
 * Save an employee. The Scheduler is master, so the write goes there first
 * and is then pulled back through the existing sync. Entitlement is written
 * wherever sync will not overwrite it: the scheduler if it has that column,
 * FairLeave otherwise.
 */
function adminEmployeeSave_(emp) {
  if (!emp || !String(emp.name || '').trim()) return { ok: false, error: 'missing_name' };

  var sh, c;
  try { sh = adminSchedEmpSheet_(); c = adminSchedCols_(sh); }
  catch (err) { return { ok: false, error: 'scheduler_unavailable', detail: String(err.message || err) }; }
  if (c.id < 0 || c.name < 0) return { ok: false, error: 'scheduler_headers_unrecognised', headers: c.H.raw };

  var rows = adminSchedRows_(sh);
  var ext = String(emp.extId || '').trim();
  var targetRow = -1;
  for (var i = 0; i < rows.length; i++) {
    if (ext && String(rows[i][c.id]).trim() === ext) { targetRow = i + 2; break; }
  }

  var isNew = targetRow === -1;
  if (isNew) {
    targetRow = sh.getLastRow() + 1;
    if (!ext) ext = adminNextExtId_(rows, c.id);
    sh.getRange(targetRow, c.id + 1).setValue(adminSafe_(ext));
  }

  var written = ['Name'];
  sh.getRange(targetRow, c.name + 1).setValue(adminSafe_(emp.name));
  if (c.dept >= 0 && emp.team !== undefined)   { sh.getRange(targetRow, c.dept + 1).setValue(adminSafe_(emp.team)); written.push('Department'); }
  var emailToFl = null;
  if (emp.email !== undefined) {
    if (c.email >= 0) { sh.getRange(targetRow, c.email + 1).setValue(adminSafe_(emp.email)); written.push('Email'); }
    else emailToFl = adminSafe_(emp.email);   // Scheduler has no Email column — FairLeave owns it
  }
  if (c.role >= 0 && emp.role !== undefined)   { sh.getRange(targetRow, c.role + 1).setValue(adminSafe_(emp.role)); written.push('Role'); }
  if (c.start >= 0 && emp.start)               { sh.getRange(targetRow, c.start + 1).setValue(adminSafe_(emp.start)); written.push('StartDate'); }
  if (c.active >= 0) { sh.getRange(targetRow, c.active + 1).setValue(emp.active === false ? 'FALSE' : 'TRUE'); written.push('Active'); }

  var ent = (emp.entitlement === '' || emp.entitlement === undefined || emp.entitlement === null)
    ? null : (parseInt(emp.entitlement, 10) || 0);
  var entTarget = null;
  if (ent !== null && c.ent >= 0) {
    // The scheduler owns entitlement — writing it in FairLeave would be
    // silently overwritten by the very next sync.
    sh.getRange(targetRow, c.ent + 1).setValue(ent);
    written.push('Entitlement');
    entTarget = 'scheduler';
  }

  // Pull the change back into FairLeave through the existing one-way importer.
  var sync = syncFromScheduler_();

  // Anything the Scheduler had nowhere to store is written to FairLeave instead,
  // after the sync so it is not immediately overwritten by it.
  var flPatch = {};
  if (ent !== null && entTarget === null) flPatch.entitlement = String(ent);
  if (emailToFl !== null) flPatch.email = emailToFl;
  if (Object.keys(flPatch).length) {
    var flEmp = findBy_('Employees', 'extId', ext);
    if (flEmp) {
      updateRow_('Employees', flEmp.id, flPatch);
      if (flPatch.entitlement) entTarget = 'fairleave';
      if (flPatch.email) written.push('Email → FairLeave');
    }
  }

  audit_('Management', isNew ? 'admin.employeeAdd' : 'admin.employeeUpdate',
    emp.name + ' (scheduler ID ' + ext + '); fields: ' + written.join(', '));

  return {
    ok: true, created: isNew, extId: ext, fields: written,
    entitlementWrittenTo: entTarget,
    syncOk: !!(sync && sync.ok),
    syncSummary: sync && sync.summary ? sync.summary : (sync && sync.error) || ''
  };
}

/**
 * Archive: mark inactive in the Scheduler (master), sync, then run FairLeave's
 * own removal so approved bookings are cancelled and pending swaps are failed.
 * Nothing is deleted — bookings, swaps and audit rows are all retained.
 */
function adminEmployeeArchive_(flId) {
  if (!flId) return { ok: false, error: 'missing_id' };
  var e = findBy_('Employees', 'id', flId);
  if (!e) return { ok: false, error: 'not_found' };

  var schedNote = 'no linked scheduler record';
  if (e.extId) {
    try {
      var sh = adminSchedEmpSheet_(), c = adminSchedCols_(sh);
      if (c.active < 0) {
        schedNote = 'scheduler has no Active column — left as is';
      } else {
        var rows = adminSchedRows_(sh), done = false;
        for (var i = 0; i < rows.length; i++) {
          if (String(rows[i][c.id]).trim() === e.extId) {
            sh.getRange(i + 2, c.active + 1).setValue('FALSE');
            schedNote = 'set inactive in scheduler'; done = true; break;
          }
        }
        if (!done) schedNote = 'scheduler row not found';
      }
    } catch (err) { schedNote = 'scheduler unavailable: ' + String(err.message || err); }
  }

  // Existing behaviour: deactivate, clear token, cancel approved bookings, fail swaps.
  var res = adminAsMgmt_('deleteEmployee', { id: flId });
  audit_('Management', 'admin.employeeArchive', e.name + ' — ' + schedNote + '; bookings cancelled, swaps closed');

  return { ok: !!(res && res.ok), scheduler: schedNote };
}

function adminRegenToken_(flId) {
  if (!flId) return { ok: false, error: 'missing_id' };
  var e = findBy_('Employees', 'id', flId);
  if (!e) return { ok: false, error: 'not_found' };
  var res = adminAsMgmt_('regenToken', { id: flId });
  if (!res || !res.ok) return { ok: false, error: 'regen_failed' };
  audit_('Management', 'admin.tokenReissue', e.name + ' — previous private link invalidated');
  return { ok: true, message: 'New link issued for ' + e.name + '. Share it from inside FairLeave using the QR code — it is deliberately not shown here.' };
}

/**
 * Clear someone's kiosk PIN so they enrol again on first use.
 *
 * The PIN hash lives on the CLOCK's Staff tab, not in FairLeave, so this is a
 * relay call - the same path, and the same three Script Properties, as
 * Sync to Clock. If the relay is not configured the console says so plainly
 * rather than reporting a success that never left this project.
 *
 * NOT in ADMIN_MUTATING, deliberately. It writes nothing in this spreadsheet,
 * so the script lock would protect nothing - and holding it across an outbound
 * HTTP request would block every other console action for the length of that
 * request.
 *
 * Identification: the clock's staffRow_ matches on NAME **or** EmployeeID. The
 * id is tried first because it survives a rename; the name is tried only if the
 * clock reports no such person, which happens when the clock's Staff tab has no
 * EmployeeID column to match against. One extra request, and only on failure.
 */
function adminResetClockPin_(emp, actor) {
  emp = emp || {};
  var name = String(emp.name || '').trim();
  var ext  = String(emp.extId || '').trim();

  // Prefer the stored record over whatever the browser sent.
  if (emp.id) {
    var fl = findBy_('Employees', 'id', emp.id);
    if (fl) {
      if (fl.name) name = String(fl.name).trim();
      if (fl.extId) ext = String(fl.extId).trim();
    }
  }
  if (!name && !ext) return { ok: false, error: 'missing_identity' };

  if (typeof schedRelay_ !== 'function') {
    return { ok: false, error: 'relay_unavailable',
             detail: 'Admin_Scheduler.gs is not in this project, so the console cannot ' +
                     'call the clock. Clear PinHash and Salt on the clock Staff tab instead - ' +
                     'both, or neither: a hash with no salt can never match.' };
  }

  var who = ext || name;
  var res = schedRelay_('resetClockPin', actor, { name: who });

  if (res && !res.ok && ext && name && /not found/i.test(String(res.detail || ''))) {
    res = schedRelay_('resetClockPin', actor, { name: name });
    who = name;
  }

  if (!res) {
    // schedRelay_ returns null when none of the three properties is set at all.
    // Say what to set AND what to do instead, because someone standing at a
    // kiosk with a locked-out member of staff needs a way through right now.
    return { ok: false, error: 'relay_unavailable',
             detail: 'The console cannot reach the clock: Clock_Exec_Url, Clock_Mgmt_PIN ' +
                     'and Clock_Device_Token are not set (Project Settings → Script Properties).\n\n' +
                     'To do it by hand meanwhile: open the Zanna Clock spreadsheet, Staff tab, ' +
                     'find their row and clear BOTH PinHash and Salt. Both or neither — a hash ' +
                     'with no salt can never match, and they would be locked out for good.' };
  }
  if (!res.ok) return res;

  audit_(actor || 'Management', 'admin.resetClockPin',
         (name || who) + ' - kiosk PIN cleared; they set a new one at first use');
  return { ok: true, name: name || who, message: res.summary };
}

function adminSync_() {
  var res = syncFromScheduler_();
  return res && res.ok ? { ok: true, summary: res.summary } : { ok: false, error: 'sync_failed', detail: res && res.error };
}

function adminImportLeave_() {
  var res = importLeaveFromScheduler_();
  return res && res.ok ? { ok: true, summary: res.summary } : { ok: false, error: 'import_failed', detail: res && res.error };
}

/**
 * Next scheduler EmployeeID, keeping whatever scheme the sheet already uses.
 *
 * v1.2.1. The previous version did `String(r[idCol]).replace(/\D/g, '')` and
 * returned `String(max + 1)`, throwing away both the prefix and the
 * zero-padding: on a roster of E001-E015 the Add employee form generated "16".
 *
 * That is not immediately broken - every sync matches the exact ID string, so
 * "16" is a valid if ugly ID. The trap is later. Add "E016" by hand afterwards
 * and the sheet holds "16" and "E016": two different people whose IDs look
 * identical at a glance, and which the Clock's duplicate-EmployeeID check will
 * not flag, because as strings they genuinely differ.
 *
 * Now: most common prefix, widest zero-padding seen, highest number + 1.
 *   E001..E015 -> E016    1..15 -> 16    EMP0007 -> EMP0008    (empty) -> E001
 *
 * IDs that are not <letters><digits> are ignored rather than guessed at.
 */
function adminNextExtId_(rows, idCol) {
  var FALLBACK_PREFIX = 'E';   // only used when the sheet has no usable IDs at all
  var FALLBACK_WIDTH  = 3;

  var max = 0, width = 0, tally = {}, bestPrefix = null, bestCount = 0, seen = 0;

  (rows || []).forEach(function (r) {
    var s = String(r[idCol] == null ? '' : r[idCol]).trim();
    var m = s.match(/^([A-Za-z]*)(\d+)$/);
    if (!m) return;

    seen++;
    var pre = m[1].toUpperCase();
    tally[pre] = (tally[pre] || 0) + 1;
    if (tally[pre] > bestCount) { bestCount = tally[pre]; bestPrefix = pre; }

    if (m[2].length > width) width = m[2].length;
    var n = parseInt(m[2], 10);
    if (!isNaN(n) && n > max) max = n;
  });

  if (!seen) { bestPrefix = FALLBACK_PREFIX; width = FALLBACK_WIDTH; }

  var next = String(max + 1);
  while (next.length < width) next = '0' + next;
  return bestPrefix + next;
}

/* ═══════════════════════════════════════════════════════════════════
   TEAMS — FairLeave owns concurrency limits (Teams.max)
   ═══════════════════════════════════════════════════════════════════ */
function adminTeamsList_() {
  var counts = {};
  readAll_('Employees').forEach(function (e) {
    if (String(e.active) !== 'FALSE') counts[e.teamId] = (counts[e.teamId] || 0) + 1;
  });
  return {
    ok: true,
    teams: readAll_('Teams').map(function (t) {
      return { id: t.id, name: t.name, max: parseInt(t.max, 10) || 1, members: counts[t.id] || 0 };
    })
  };
}

function adminTeamSave_(team) {
  if (!team || !String(team.name || '').trim()) return { ok: false, error: 'missing_name' };
  var max = Math.max(1, parseInt(team.max, 10) || 1);
  var res = team.id
    ? adminAsMgmt_('updateTeam', { id: team.id, name: team.name, max: max })
    : adminAsMgmt_('addTeam', { name: team.name, max: max });
  if (!res || !res.ok) return { ok: false, error: 'save_failed', detail: res && res.error };
  audit_('Management', 'admin.teamSave', team.name + ' max off per day = ' + max);
  return { ok: true, created: !team.id };
}

function adminTeamDelete_(id) {
  if (!id) return { ok: false, error: 'missing_id' };
  // route_ refuses if the team still has active members — that guard is kept.
  var res = adminAsMgmt_('deleteTeam', { id: id });
  return res && res.ok ? { ok: true } : { ok: false, error: 'delete_failed', detail: res && res.error };
}

/* ═══════════════════════════════════════════════════════════════════
   SETTINGS — Config tab. MGMT_PIN is masked by ADMIN_SECRET_KEY_RE.
   ═══════════════════════════════════════════════════════════════════ */
function adminSettingsGet_() {
  var known = {
    COMPANY_NAME:       'Name shown across FairLeave and in notification emails.',
    WORK_SATURDAY:      'TRUE = Mon–Sat working week, FALSE = Mon–Fri. Saved as uppercase automatically. Changes how future bookings are counted; existing bookings are not recalculated. Also editable on the Scheduler tab — same row.',
    SCHEDULER_SHEET_ID: 'Spreadsheet ID of the Shift Scheduler. Changing this repoints every sync.',
    CLOCK_SHEET_ID:     'Spreadsheet ID of the Zanna Clock (Events/Sessions/Payroll). Used by the Scheduler tab.',
    MGMT_PIN:           'Management PIN for the FairLeave app itself.'
  };
  return {
    ok: true,
    settings: readAll_('Config').map(function (r) {
      var secret = ADMIN_SECRET_KEY_RE.test(r.key);
      return { key: r.key, value: secret ? '' : r.value, secret: secret, note: known[r.key] || '' };
    })
  };
}

function adminSettingsSave_(settings) {
  if (!settings || !settings.length) return { ok: false, error: 'nothing_to_save' };
  var changed = [], blocked = [];
  settings.forEach(function (s) {
    if (!s || !s.key) return;
    if (ADMIN_SECRET_KEY_RE.test(s.key)) { blocked.push(s.key); return; }
    var val = adminSafe_(s.value);
    // WORK_SATURDAY is compared as a string elsewhere — normalise the case here
    // so a lowercase 'true' can never silently flip the working week again.
    if (/^(true|false)$/i.test(String(val).trim())) val = String(val).trim().toUpperCase();
    setConfig_(s.key, val);
    changed.push(s.key + ' = ' + s.value);
  });
  if (changed.length) audit_('Management', 'admin.settingsUpdate', changed.join('; '));
  return { ok: true, changed: changed, blocked: blocked };
}

/* ═══════════════════════════════════════════════════════════════════
   BLACKOUTS — company-wide only, per the locked decision
   ═══════════════════════════════════════════════════════════════════ */
function adminBlackoutSave_(b) {
  if (!b || !b.start || !b.end) return { ok: false, error: 'missing_dates' };
  if (b.end < b.start) return { ok: false, error: 'end_before_start' };
  var res = adminAsMgmt_('addBlackout', { start: b.start, end: b.end, reason: adminSafe_(b.reason || '') });
  if (!res || !res.ok) return { ok: false, error: 'save_failed', detail: res && res.error };
  audit_('Management', 'admin.blackoutAdd', b.start + ' → ' + b.end + (b.reason ? ' (' + b.reason + ')' : ''));
  return { ok: true };
}

function adminBlackoutDelete_(id) {
  if (!id) return { ok: false, error: 'missing_id' };
  var bo = findBy_('Blackouts', 'id', id);
  var res = adminAsMgmt_('deleteBlackout', { id: id });
  if (!res || !res.ok) return { ok: false, error: 'delete_failed' };
  audit_('Management', 'admin.blackoutDelete', bo ? bo.start + ' → ' + bo.end : String(id));
  return { ok: true };
}

/* ═══════════════════════════════════════════════════════════════════
   CONTROL PANEL CONFIG — Panel_Apps and Access_Token
   ═══════════════════════════════════════════════════════════════════ */
function adminPanelGet_() {
  var props = PropertiesService.getScriptProperties();
  var apps = {}, parseError = null;
  try { apps = JSON.parse(props.getProperty(ADMIN_CFG.PANEL_PROP) || '{}'); }
  catch (e) { parseError = String(e.message); }
  var token = String(props.getProperty(ADMIN_CFG.TOKEN_PROP) || '');
  return {
    ok: true, apps: apps, parseError: parseError,
    tokenLength: token.length, tokenWeak: token.length < 12,
    // 'payroll' added 2026-08-15. A key absent from BOTH this list and the saved
    // map gets no input field in the console, so a new app could never be added
    // through the UI — chicken and egg. New app? Add its key here.
    knownKeys: ['fairleave', 'scheduler', 'lms', 'clockKiosk', 'clockCanteen',
                'clockOffice', 'clockEmergency', 'payroll']
  };
}

function adminPanelSave_(apps) {
  if (!apps || typeof apps !== 'object') return { ok: false, error: 'bad_payload' };
  var clean = {}, rejected = [];
  Object.keys(apps).forEach(function (k) {
    var v = String(apps[k] || '').trim();
    if (!v) return;
    if (!/^https:\/\/[^\s"'<>]+$/i.test(v)) { rejected.push(k); return; }
    clean[k] = v;
  });
  if (rejected.length) return { ok: false, error: 'invalid_urls', rejected: rejected };
  PropertiesService.getScriptProperties().setProperty(ADMIN_CFG.PANEL_PROP, JSON.stringify(clean));
  audit_('Management', 'admin.panelApps', Object.keys(clean).join(', ') || '(all cleared)');
  return { ok: true, apps: clean };
}

function adminRotateToken_(token) {
  token = String(token || '').trim();
  if (token.length < 12) return { ok: false, error: 'too_short', min: 12 };
  if (/\s/.test(token)) return { ok: false, error: 'no_spaces' };
  PropertiesService.getScriptProperties().setProperty(ADMIN_CFG.TOKEN_PROP, token);
  audit_('Management', 'admin.tokenRotate', 'Panel access code changed (length ' + token.length + ')');
  return { ok: true, message: 'Access code updated. Every staff device needs the new code on its next visit.' };
}

/* ═══════════════════════════════════════════════════════════════════
   AUDIT
   ═══════════════════════════════════════════════════════════════════ */
function adminAuditList_(limit) {
  limit = Math.min(Math.max(parseInt(limit, 10) || 100, 1), 500);
  var all = readAll_('Audit');
  var slice = all.slice(Math.max(0, all.length - limit)).reverse();
  return {
    ok: true, total: all.length,
    entries: slice.map(function (r) {
      return {
        when: String(r.timestamp || '').replace('T', ' ').slice(0, 16),
        actor: r.actor || '', action: r.action || '', details: r.details || ''
      };
    })
  };
}

/* ═══════════════════════════════════════════════════════════════════
   DIAGNOSTIC — run this first when something misbehaves.
   Reports what was found rather than assuming anything.
   ═══════════════════════════════════════════════════════════════════ */
function adminDiagnose_() {
  var report = { ok: true, adminVersion: ADMIN_VERSION, fairleave: {}, scheduler: {}, config: {}, properties: {} };
  var props = PropertiesService.getScriptProperties();

  var pin = props.getProperty(ADMIN_CFG.PIN_PROP) || '';
  report.properties.Admin_PIN = pin ? (pin.length + ' chars') : 'MISSING';
  report.properties.Admin_PIN_strength = pin.length >= ADMIN_CFG.MIN_PIN_LENGTH ? 'ok' : 'TOO SHORT — lengthen it';
  report.properties.Access_Token = (props.getProperty(ADMIN_CFG.TOKEN_PROP) || '').length + ' chars';
  report.properties.Panel_Apps = props.getProperty(ADMIN_CFG.PANEL_PROP) ? 'set' : 'MISSING';

  // Is the Scheduler tab's backend actually present in this deployment?
  report.properties.Admin_Scheduler_gs = (typeof adminSchedulerHandle_ === 'function')
    ? 'loaded' : 'MISSING — the console Scheduler tab will show its deploy banner';

  readAll_('Config').forEach(function (r) {
    report.config[r.key] = ADMIN_SECRET_KEY_RE.test(r.key) ? '(hidden)' : r.value;
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  report.fairleave.name = ss.getName();
  report.fairleave.tabs = ss.getSheets().map(function (s) { return s.getName(); });
  ['Config', 'Teams', 'Employees', 'Bookings', 'Blackouts', 'Swaps', 'Audit'].forEach(function (t) {
    var sh = ss.getSheetByName(t);
    report.fairleave[t] = sh ? (Math.max(0, sh.getLastRow() - 1) + ' rows') : 'MISSING';
  });
  var linked = 0, unlinked = 0;
  readAll_('Employees').forEach(function (e) { e.extId ? linked++ : unlinked++; });
  report.fairleave.linkedToScheduler = linked;
  report.fairleave.notLinked = unlinked + (unlinked ? '  ← these will not sync; they have no extId' : '');

  try {
    var sc = adminSchedSS_();
    report.scheduler.name = sc.getName();
    report.scheduler.tabs = sc.getSheets().map(function (s) { return s.getName(); });
    var sh2 = schedTab_(sc, ['Employees', 'Staff', 'People']);
    if (!sh2) {
      report.scheduler.employees = 'NO Employees/Staff/People TAB FOUND';
      report.ok = false;
    } else {
      var c = adminSchedCols_(sh2);
      report.scheduler.employeesTab = sh2.getName();
      report.scheduler.headers = c.H.raw;
      report.scheduler.rows = Math.max(0, sh2.getLastRow() - 1);
      report.scheduler.resolved = {};
      ['id', 'name', 'dept', 'active', 'ent', 'email', 'role', 'start'].forEach(function (k) {
        report.scheduler.resolved[k] = c[k] >= 0 ? c.H.raw[c[k]] : 'UNRESOLVED';
      });
      report.scheduler.entitlementOwner = c.ent >= 0
        ? 'scheduler (console writes entitlement there so sync cannot overwrite it)'
        : 'fairleave (scheduler has no entitlement column)';
    }
    var lv = schedTab_(sc, ['LeaveRequests', 'Leave', 'LeaveLog']);
    report.scheduler.leaveTab = lv ? lv.getName() : 'NOT FOUND — leave push will be skipped';
  } catch (err) {
    report.scheduler.error = String(err.message || err);
    report.ok = false;
  }

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/** Run this one from the editor dropdown. */
function runAdminDiagnostic() { return adminDiagnose_(); }
