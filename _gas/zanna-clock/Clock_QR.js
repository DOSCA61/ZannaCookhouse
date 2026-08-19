/**
 * Clock_QR.gs — QR shift login for Zanna Clock
 * v1.1.0 · targets Zanna Clock v1.11.0 · written 2026-08-17, rewritten 2026-08-18
 *
 * WHAT CHANGED IN v1.1.0, AND WHY IT MATTERS
 * v1.0.0 was written without sight of Code.gs. It guessed how the Clock stores
 * a clock-in and guessed wrong. It is NOT safe to paste — see the notes above
 * section 5. This version calls the Clock's own state machine instead of
 * reimplementing it, so the phone and the kitchen tablet cannot drift apart.
 *
 *   - Events is [Timestamp, Name, EventType, BreakType, SessionID, Premises,
 *     Device] and is keyed on NAME. v1.0.0 looked for EmployeeID and Date
 *     columns; neither exists, so it would have refused every clock-in.
 *   - Code.gs serialises on a SCRIPT lock. v1.0.0 took a DOCUMENT lock, so the
 *     two did not exclude each other.
 *   - The kiosk resets at midnight. v1.0.0 searched 4,000 rows back, so a
 *     forgotten clock-out yesterday blocked this morning — the phone was
 *     stricter than the tablet for the same person.
 *   - The personal token now travels in the URL FRAGMENT, not the query, so it
 *     never reaches Google's request logs or a Referer header. This also
 *     removes the script-injection hole in the phone page at the root: the
 *     server never sees the token, so it cannot embed it.
 *   - The kiosk panel counts down against wall-clock time, so a tablet that
 *     slept cannot sit there showing an expired code that looks healthy.
 *
 * A separate script file on purpose: per the project's 2026-08-15 decision,
 * helpers live in their own .gs so a full-file paste over Code.gs cannot lose
 * them. Apps Script merges every .gs in a project into one namespace, so this
 * runs exactly as if it were part of Code.gs — but only inside the
 * **Zanna Clock** project. Pasted into FairLeave it will find nothing.
 *
 * WHAT THIS ADDS
 *   A kiosk screen shows a QR that changes every 10 minutes. Staff scan it with
 *   the phone camera, and clock in with one tap. Their identity is a personal
 *   token, not a PIN. Breaks stay on the kiosk; phones are away during service.
 *
 * INTEGRATION — exactly one insertion into Code.gs. See clkQrInstallHelp().
 *
 * ES5 only (no let/const/arrow/template literals) so this file behaves the same
 * on the V8 and legacy runtimes.
 */

var CLK_QR_VERSION = '1.3.2';

/* Crockford base32: no I, L, O or U — nothing that can be misread aloud. */
var CLK_QR_ALPHA = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';

/* Script Properties this file owns. Never in a sheet cell — 2026-08-15 decision. */
var CLK_QR_P_SECRET = 'Clock_QR_Secret';   /* signs the rotating site code   */
var CLK_QR_P_PEPPER = 'Clock_Token_Pepper'; /* hashes employee mobile tokens */


/* ==========================================================================
 * 1 · Sheet plumbing
 * Everything here resolves by header name, never by column index, so a column
 * inserted in the Clock sheet by hand does not silently corrupt clock events.
 * ======================================================================== */

function clkQrSS_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** First sheet whose name matches any candidate, case-insensitively. */
function clkQrSheet_(candidates, required) {
  var sheets = clkQrSS_().getSheets(), i, j;
  for (i = 0; i < candidates.length; i++) {
    for (j = 0; j < sheets.length; j++) {
      if (String(sheets[j].getName()).trim().toLowerCase() === candidates[i].toLowerCase()) {
        return sheets[j];
      }
    }
  }
  if (required) throw new Error('Clock_QR: no sheet named ' + candidates.join(' / '));
  return null;
}

/** Normalised header -> 0-based column index. Matches the Clock's own headerMap_ rules. */
function clkQrHeaders_(sheet) {
  var last = sheet.getLastColumn();
  if (last < 1) return {};
  var row = sheet.getRange(1, 1, 1, last).getValues()[0], map = {}, i;
  for (i = 0; i < row.length; i++) {
    var k = String(row[i]).toLowerCase().replace(/[\s_]/g, '');
    if (k && map[k] === undefined) map[k] = i;
  }
  return map;
}

/** First header alias that exists, or -1. */
function clkQrCol_(headers, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var k = aliases[i].toLowerCase().replace(/[\s_]/g, '');
    if (headers[k] !== undefined) return headers[k];
  }
  return -1;
}

/** Append a column and return its 0-based index. */
function clkQrAddCol_(sheet, title) {
  var col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue(title);
  return col - 1;
}

function clkQrStaffSheet_() { return clkQrSheet_(['Staff', 'Employees', 'Staff List', 'People'], true); }
function clkQrEventsSheet_() { return clkQrSheet_(['Events', 'Event Log', 'Clockings'], true); }
function clkQrDevicesSheet_() { return clkQrSheet_(['Devices', 'Device', 'Kiosks'], true); }
function clkQrConfigSheet_() { return clkQrSheet_(['Config', 'Settings', 'Configuration'], false); }
function clkQrAuditSheet_() { return clkQrSheet_(['Audit', 'Audit Log', 'AuditLog', 'Log'], false); }

/** Config tab lookup. Own implementation so this file does not depend on Code.gs. */
function clkQrCfg_(key, dflt) {
  var sh = clkQrConfigSheet_();
  if (!sh || sh.getLastRow() < 1) return dflt;
  var vals = sh.getRange(1, 1, sh.getLastRow(), Math.max(2, sh.getLastColumn())).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === key.toLowerCase()) {
      var v = String(vals[i][1]).trim();
      return v === '' ? dflt : v;
    }
  }
  return dflt;
}

function clkQrCfgBool_(key, dflt) {
  var v = clkQrCfg_(key, null);
  if (v === null) return dflt;
  v = String(v).trim().toLowerCase();
  return v === 'true' || v === 'yes' || v === '1' || v === 'y';
}

function clkQrCfgNum_(key, dflt) {
  var v = Number(clkQrCfg_(key, dflt));
  return isNaN(v) ? dflt : v;
}

/** Write an audit row if the Clock has an Audit tab. Never throws — auditing must not break clocking. */
function clkQrAudit_(event, detail) {
  try {
    var sh = clkQrAuditSheet_();
    if (!sh) return;
    var h = clkQrHeaders_(sh);
    var n = Math.max(sh.getLastColumn(), 3);
    var row = [];
    for (var i = 0; i < n; i++) row.push('');
    var cT = clkQrCol_(h, ['Timestamp', 'Time', 'When', 'DateTime']);
    var cE = clkQrCol_(h, ['Event', 'Action', 'Type']);
    var cD = clkQrCol_(h, ['Detail', 'Details', 'Message', 'Note', 'Notes', 'Data']);
    row[cT >= 0 ? cT : 0] = new Date();
    row[cE >= 0 ? cE : 1] = event;
    row[cD >= 0 ? cD : 2] = detail == null ? '' : String(detail);
    sh.appendRow(row);
  } catch (err) { /* deliberately silent */ }
}


/* ==========================================================================
 * 2 · Secrets
 * ======================================================================== */

function clkQrProp_(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null ? '' : String(v);
}

/**
 * Create the two Script Properties if they are missing. Safe to run twice —
 * it never overwrites an existing value, because rotating the QR secret
 * invalidates every code on screen and rotating the pepper invalidates every
 * employee token at once.
 */
function clkQrEnsureSecrets() {
  var props = PropertiesService.getScriptProperties();
  var made = [];
  if (!props.getProperty(CLK_QR_P_SECRET)) {
    props.setProperty(CLK_QR_P_SECRET, Utilities.getUuid() + Utilities.getUuid());
    made.push(CLK_QR_P_SECRET);
  }
  if (!props.getProperty(CLK_QR_P_PEPPER)) {
    props.setProperty(CLK_QR_P_PEPPER, Utilities.getUuid() + Utilities.getUuid());
    made.push(CLK_QR_P_PEPPER);
  }
  var msg = made.length ? 'Created: ' + made.join(', ') : 'Both already set — nothing changed.';
  Logger.log(msg);
  clkQrAudit_('qr.secrets', msg);
  return msg;
}

function clkQrSecret_() {
  var v = clkQrProp_(CLK_QR_P_SECRET);
  if (!v) throw new Error('Clock_QR: ' + CLK_QR_P_SECRET + ' is not set. Run clkQrEnsureSecrets() once.');
  return v;
}

function clkQrPepper_() {
  var v = clkQrProp_(CLK_QR_P_PEPPER);
  if (!v) throw new Error('Clock_QR: ' + CLK_QR_P_PEPPER + ' is not set. Run clkQrEnsureSecrets() once.');
  return v;
}


/* ==========================================================================
 * 3 · The rotating site code
 *
 * Derived, never stored: code = base32(HMAC-SHA256(secret, device|window)).
 * Nothing to write, nothing to clean up, and a code cannot leak from the sheet
 * because it exists nowhere except on the screen for ten minutes.
 * ======================================================================== */

function clkQrWindowSecs_() {
  var n = clkQrCfgNum_('QR_WINDOW_SECS', 600);
  return (n >= 60 && n <= 3600) ? Math.floor(n) : 600;
}

function clkQrGraceSecs_() {
  var n = clkQrCfgNum_('QR_GRACE_SECS', 90);
  return (n >= 0 && n <= 600) ? Math.floor(n) : 90;
}

function clkQrWindowAt_(epochSecs) { return Math.floor(epochSecs / clkQrWindowSecs_()); }

function clkQrNowSecs_() { return Math.floor(new Date().getTime() / 1000); }

function clkQrB32_(bytes, chars) {
  var bits = 0, value = 0, out = '';
  for (var i = 0; i < bytes.length && out.length < chars; i++) {
    value = (value << 8) | (bytes[i] & 0xff);
    bits += 8;
    while (bits >= 5 && out.length < chars) {
      out += CLK_QR_ALPHA[(value >>> (bits - 5)) & 31];
      bits -= 5;
    }
  }
  return out;
}

/** The code a given device is showing during a given window. */
function clkQrCodeFor_(deviceToken, windowIdx) {
  var sig = Utilities.computeHmacSha256Signature(
    String(deviceToken) + '|' + String(windowIdx), clkQrSecret_());
  return clkQrB32_(sig, 8);
}

/** Tidy up what somebody typed: strip punctuation, fix the letters Crockford drops. */
function clkQrNormaliseCode_(raw) {
  var s = String(raw == null ? '' : raw).toUpperCase().replace(/[^0-9A-Z]/g, '');
  s = s.replace(/[IL]/g, '1').replace(/O/g, '0');
  return s;
}

function clkQrConstEq_(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  for (var i = 0; i < a.length && i < b.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/** Active devices that are allowed to host a QR panel. */
function clkQrDevices_() {
  var sh = clkQrDevicesSheet_(), h = clkQrHeaders_(sh);
  var cTok = clkQrCol_(h, ['Token', 'DeviceToken', 'Device Token', 'Key']);
  var cTyp = clkQrCol_(h, ['Type', 'DeviceType', 'Kind', 'Role']);
  var cAct = clkQrCol_(h, ['Active', 'Enabled', 'IsActive']);
  var cNam = clkQrCol_(h, ['Name', 'Label', 'Device', 'Location', 'Description']);
  if (cTok < 0) throw new Error('Clock_QR: Devices tab has no Token column.');

  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var tok = String(vals[i][cTok]).trim();
    if (!tok) continue;
    var active = cAct < 0 ? true : /^(true|yes|1|y)$/i.test(String(vals[i][cAct]).trim());
    if (!active) continue;
    var type = cTyp < 0 ? 'KIOSK' : String(vals[i][cTyp]).trim().toUpperCase();
    if (type !== 'KIOSK' && type !== 'QR') continue;   /* displays stay read-only */
    out.push({
      token: tok,
      type: type,
      name: cNam < 0 ? type : (String(vals[i][cNam]).trim() || type),
      row: i + 2
    });
  }
  return out;
}

/**
 * Which device — if any — is showing this code right now?
 * The previous window is accepted for QR_GRACE_SECS after rollover, so nobody
 * is rejected because the code turned over between the scan and the tap.
 */
function clkQrResolveCode_(rawCode) {
  var code = clkQrNormaliseCode_(rawCode);
  if (code.length !== 8) return { ok: false, reason: 'BAD_FORMAT' };

  var now = clkQrNowSecs_(), win = clkQrWindowSecs_();
  var cur = Math.floor(now / win);
  var intoWindow = now - cur * win;
  var windows = [cur];
  if (intoWindow < clkQrGraceSecs_()) windows.push(cur - 1);

  var devices = clkQrDevices_(), i, j;
  for (i = 0; i < devices.length; i++) {
    for (j = 0; j < windows.length; j++) {
      if (clkQrConstEq_(code, clkQrCodeFor_(devices[i].token, windows[j]))) {
        return { ok: true, device: devices[i], stale: windows[j] !== cur };
      }
    }
  }
  return { ok: false, reason: 'EXPIRED_OR_UNKNOWN' };
}


/* ==========================================================================
 * 4 · Employee mobile tokens
 *
 * The sheet holds a HASH of the token, never the token. Anyone with edit access
 * to the workbook can read every row of it — which is the whole reason the
 * management PIN was moved to Script Properties on 2026-08-15. The same
 * reasoning applies to a credential that clocks someone in.
 *
 * The consequence, and it is a real one: a personal link can be shown at the
 * moment it is issued and never again. Lost link = reissue.
 * ======================================================================== */

function clkQrTokenNew_() {
  var hex = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');  /* 64 hex = 32 bytes */
  var bytes = [];
  for (var i = 0; i < 32; i += 2) bytes.push(parseInt(hex.substr(i, 2), 16));  /* first 16 bytes */
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');             /* 22 chars */
}

function clkQrTokenHash_(token) {
  return Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(String(token), clkQrPepper_()));
}

/** Make sure the Staff tab has the three columns this feature needs. */
function clkQrEnsureStaffCols_() {
  var sh = clkQrStaffSheet_(), h = clkQrHeaders_(sh);
  var made = [];
  if (clkQrCol_(h, ['MobileTokenHash']) < 0) { clkQrAddCol_(sh, 'MobileTokenHash'); made.push('MobileTokenHash'); }
  h = clkQrHeaders_(sh);
  if (clkQrCol_(h, ['MobileTokenIssued']) < 0) { clkQrAddCol_(sh, 'MobileTokenIssued'); made.push('MobileTokenIssued'); }
  h = clkQrHeaders_(sh);
  if (clkQrCol_(h, ['MobileTokenRevoked']) < 0) { clkQrAddCol_(sh, 'MobileTokenRevoked'); made.push('MobileTokenRevoked'); }
  return made;
}

function clkQrStaffCols_(sh) {
  var h = clkQrHeaders_(sh);
  return {
    id: clkQrCol_(h, ['EmployeeID', 'Employee ID', 'EmpID', 'StaffID', 'ID']),
    name: clkQrCol_(h, ['Name', 'EmployeeName', 'Employee', 'FullName', 'Staff']),
    dept: clkQrCol_(h, ['Department', 'Dept', 'Area']),
    active: clkQrCol_(h, ['Active', 'IsActive', 'Enabled']),
    hash: clkQrCol_(h, ['MobileTokenHash']),
    issued: clkQrCol_(h, ['MobileTokenIssued']),
    revoked: clkQrCol_(h, ['MobileTokenRevoked'])
  };
}

function clkQrStaffRows_() {
  var sh = clkQrStaffSheet_(), c = clkQrStaffCols_(sh);
  if (c.id < 0 || c.name < 0) throw new Error('Clock_QR: Staff tab needs EmployeeID and Name columns.');
  var last = sh.getLastRow();
  if (last < 2) return { sheet: sh, cols: c, rows: [] };
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][c.id]).trim();
    if (!id) continue;
    rows.push({
      row: i + 2,
      id: id,
      name: String(vals[i][c.name]).trim(),
      dept: c.dept < 0 ? '' : String(vals[i][c.dept]).trim(),
      active: c.active < 0 ? true : /^(true|yes|1|y)$/i.test(String(vals[i][c.active]).trim()),
      hash: c.hash < 0 ? '' : String(vals[i][c.hash]).trim(),
      revoked: c.revoked < 0 ? '' : String(vals[i][c.revoked]).trim()
    });
  }
  return { sheet: sh, cols: c, rows: rows };
}

/** Token -> employee, or null. Constant-time over the hash so a timing signal cannot narrow it. */
function clkQrEmployeeForToken_(token) {
  token = String(token == null ? '' : token).trim();
  if (token.length < 16) return null;
  var want = clkQrTokenHash_(token);
  var st = clkQrStaffRows_(), found = null;
  for (var i = 0; i < st.rows.length; i++) {
    var r = st.rows[i];
    if (r.hash && clkQrConstEq_(r.hash, want)) found = r;
  }
  if (!found) return null;
  if (!found.active) return { blocked: 'INACTIVE', emp: found };
  if (found.revoked) return { blocked: 'REVOKED', emp: found };
  return { blocked: '', emp: found };
}


/* ==========================================================================
 * 5 · Writing the clock event
 *
 * The one place this feature touches the Clock's own data. It writes an Events
 * row shaped like the kiosk's, so Sessions, Payroll and the analytics rebuilds
 * carry on working without knowing mobile clock-ins exist.
 *
 * RUN clkQrDiagnose() BEFORE GO-LIVE. It prints the Events headers it matched
 * and the row it would write, without writing one.
 * ======================================================================== */

/* ==========================================================================
 * 5 · Writing the clock event
 *
 * v1.1.0 — this section used to GUESS. It resolved Events columns by header
 * alias and assembled its own row, because it was written without sight of
 * Code.gs. It now calls the Clock's own state machine and writes the same
 * seven-column row the kitchen tablet writes.
 *
 * Everything here depends on Code.gs being in the same Apps Script project.
 * That dependency is the point: one state machine, not two that can drift.
 * ======================================================================== */

/** This file is useless outside the Zanna Clock project. Say so loudly. */
function clkQrRequireClock_() {
  if (typeof currentStateFor_ !== 'function' || typeof staffRow_     !== 'function' ||
      typeof withLock_        !== 'function' || typeof todaysEvents_ !== 'function' ||
      typeof tab_             !== 'function' || typeof audit_        !== 'function') {
    throw new Error('Clock_QR: Code.gs was not found in this project. This file belongs ' +
                    'in the Zanna Clock Apps Script project and nowhere else.');
  }
}

/**
 * Optional Method column on Events, so a phone clock-in is distinguishable
 * from a kiosk one in payroll review.
 *
 * Appending a column is safe: rebuildSessions_ resolves every Events column
 * with indexOf, and mutate_'s positional seven-value append simply leaves the
 * new cell blank on kiosk rows. Nothing else in Code.gs reads it.
 */
function clkQrMethodCol_() {
  var sh = tab_('Events'), last = sh.getLastColumn();
  if (last < 1) return -1;
  var head = sh.getRange(1, 1, 1, last).getValues()[0];
  for (var i = 0; i < head.length; i++) {
    if (String(head[i]).trim().toLowerCase() === 'method') return i;
  }
  return -1;
}

function clkQrEnsureEventCols_() {
  clkQrRequireClock_();
  if (clkQrMethodCol_() >= 0) return [];
  var sh = tab_('Events');
  sh.getRange(1, sh.getLastColumn() + 1).setValue('Method');
  return ['Method'];
}

function clkQrTz_() {
  try { return clkQrSS_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
  catch (e) { return 'Europe/Dublin'; }
}

function clkQrTime_(d) {
  try { return Utilities.formatDate(new Date(d), clkQrTz_(), 'HH:mm'); }
  catch (e) { return ''; }
}

/**
 * Record a mobile clock event.
 *
 * Deliberately mirrors mutate_ in Code.gs minus the PIN check — here the token
 * IS the PIN, and it was verified before we reached this point. Same lock, same
 * guards, same SessionID shape, same positional append, so Sessions, Payroll
 * and the analytics rebuild cannot tell a phone clock-in from a kiosk one
 * except by the Device and Method columns, which is exactly the intent.
 *
 * The SessionID format is not cosmetic. mutate_ numbers each shift; reuse the
 * first shift's id and rebuildSessions_ merges two shifts into one and loses
 * the hours. Copy it exactly or not at all.
 */
function clkQrRecord_(emp, want, deviceName) {
  clkQrRequireClock_();
  return withLock_(function () {
    var s = staffRow_(emp.id || emp.name);
    if (!s) {
      return { ok: false, code: 'NO_STAFF',
               error: 'You are not on the current rota. See a manager.' };
    }

    var who = s.name;                        /* canonical, never what the phone sent */
    var st  = currentStateFor_(who);
    var now = new Date();

    if (want === 'in') {
      if (st.status !== 'OUT') {
        return { ok: false, code: st.status === 'ON_BREAK' ? 'ON_BREAK' : 'ALREADY_IN',
                 error: st.status === 'ON_BREAK'
                   ? 'You are on a break. End it on the kitchen tablet.'
                   : 'You are already clocked in' +
                     (st.clockIn ? ' since ' + clkQrTime_(st.clockIn) : '') + '.' };
      }
    } else {
      if (st.status === 'OUT') {
        return { ok: false, code: 'NOT_IN',
                 error: 'You are not clocked in, so there is nothing to close.' };
      }
      if (st.status === 'ON_BREAK') {
        return { ok: false, code: 'ON_BREAK',
                 error: 'End your break on the kitchen tablet before clocking out.' };
      }
    }

    var eventType = (want === 'in') ? 'CLOCK_IN' : 'CLOCK_OUT';
    var sessionID;
    if (want === 'in') {
      var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var priorShifts = todaysEvents_().filter(function (e) {
        return e.name === who && e.type === 'CLOCK_IN';
      }).length;
      sessionID = who + '-' + dateStr + '-' + ('00' + (priorShifts + 1)).slice(-3) + '-SHIFT';
    } else {
      sessionID = st.sessionID || '';
    }

    if (clkQrCfgBool_('QR_DRY_RUN', false)) {
      clkQrAudit_('qr.dryRun', who + ' ' + eventType + ' via ' + deviceName);
      return { ok: true, dryRun: true, action: eventType, at: clkQrTime_(now), name: who };
    }

    /* Events column order, positionally: Timestamp, Name, EventType, BreakType,
       SessionID, Premises, Device. Breaks never come from a phone, so BreakType
       and Premises are blank — the same blanks mutate_ writes for a CLOCK_IN. */
    var row = [now, who, eventType, '', sessionID, '', deviceName];
    var mCol = clkQrMethodCol_();
    if (mCol >= 0) {
      while (row.length < mCol) row.push('');
      row[mCol] = 'MOBILE_QR';
    }
    tab_('Events').appendRow(row);

    audit_(who, eventType, 'MOBILE_QR · ' + deviceName);
    return { ok: true, action: eventType, at: clkQrTime_(now), name: who };
  });
}

/** Today's state for one person, expressed the way the phone page wants it. */
function clkQrStateFor_(name) {
  clkQrRequireClock_();
  var st = currentStateFor_(name);
  return {
    state: st.status === 'OUT' ? 'OUT' : 'IN',
    onBreak: st.status === 'ON_BREAK',
    since: st.clockIn ? clkQrTime_(st.clockIn) : ''
  };
}


/* ==========================================================================
 * 6 · What the phone calls
 * Reached by google.script.run from the page this file serves, so there is no
 * doPost route to add and no CORS to fight.
 * ======================================================================== */

/** Read-only: who is this token, and are they in or out? */
function clkQrHello(token) {
  try {
    if (!clkQrCfgBool_('QR_ENABLED', true)) return { ok: false, error: 'QR clock-in is switched off.' };
    var hit = clkQrEmployeeForToken_(token);
    if (!hit) {
      clkQrAudit_('qr.badToken', 'clkQrHello');
      return { ok: false, code: 'BAD_TOKEN', error: 'This link is not recognised. Ask for a new one.' };
    }
    if (hit.blocked === 'INACTIVE') return { ok: false, code: 'INACTIVE', error: 'You are not active on the rota.' };
    if (hit.blocked === 'REVOKED') return { ok: false, code: 'REVOKED', error: 'This link has been withdrawn. Ask for a new one.' };

    /* The Clock's own view of today. It resets at midnight, exactly as the
       kitchen tablet does, so a forgotten clock-out yesterday cannot block
       this morning — v1.0.0 walked back 4,000 rows and did block it. */
    var st = clkQrStateFor_(hit.emp.name);
    return {
      ok: true,
      name: hit.emp.name,
      first: String(hit.emp.name).split(/\s+/)[0],
      state: st.state,
      since: st.since,
      onBreak: st.onBreak,
      allowOut: clkQrCfgBool_('QR_ALLOW_OUT', true),
      outNeedsScan: clkQrCfgBool_('QR_OUT_REQUIRES_SCAN', true)
    };
  } catch (err) {
    return { ok: false, error: 'The clock could not be reached: ' + err.message };
  }
}

/** payload: { t: token, s: site code, a: 'in' | 'out' } */
function clkQrSubmit(payload) {
  try {
    if (!clkQrCfgBool_('QR_ENABLED', true)) return { ok: false, error: 'QR clock-in is switched off.' };
    payload = payload || {};
    var want = String(payload.a) === 'out' ? 'out' : 'in';

    var hit = clkQrEmployeeForToken_(payload.t);
    if (!hit) {
      clkQrAudit_('qr.badToken', 'clkQrSubmit');
      return { ok: false, code: 'BAD_TOKEN', error: 'This link is not recognised. Ask for a new one.' };
    }
    if (hit.blocked) return { ok: false, code: hit.blocked, error: 'This link cannot be used. Speak to your manager.' };

    if (want === 'out' && !clkQrCfgBool_('QR_ALLOW_OUT', true)) {
      return { ok: false, code: 'OUT_DISABLED', error: 'Clock out on the kitchen tablet.' };
    }

    var needScan = (want === 'in') || clkQrCfgBool_('QR_OUT_REQUIRES_SCAN', true);
    var deviceName = 'MOBILE';

    if (needScan) {
      var res = clkQrResolveCode_(payload.s);
      if (!res.ok) {
        clkQrAudit_('qr.badCode', hit.emp.id + ' · ' + res.reason);
        return {
          ok: false, code: res.reason,
          error: res.reason === 'BAD_FORMAT'
            ? 'That code does not look right — it is 8 characters.'
            : 'That code has expired. Look at the screen again for the new one.'
        };
      }
      deviceName = res.device.name;
    }

    return clkQrRecord_(hit.emp, want, deviceName);
  } catch (err) {
    return { ok: false, error: 'The clock could not be reached: ' + err.message };
  }
}


/* ==========================================================================
 * 7 · Pages
 * ======================================================================== */

function clkQrExecUrl_() {
  var cfg = clkQrCfg_('QR_EXEC_URL', '');
  if (cfg) return String(cfg).trim();
  try { return ScriptApp.getService().getUrl(); } catch (e) { return ''; }
}

/**
 * Is this the editor's private test URL rather than the deployed one?
 *
 * Called from the editor, ScriptApp.getService().getUrl() hands back the /dev
 * URL. It works perfectly for the person who owns the project and for nobody
 * else — so a link or a QR built from it looks completely fine while you test
 * it, and demands a Google sign-in from every member of staff on the morning.
 * Caught this way on 2026-08-18, one step before issuing thirteen dead links.
 */
function clkQrIsDevUrl_(u) {
  return /\/dev(\?|#|$)/.test(String(u || ''));
}

/** Throw rather than mint anything against a URL only the owner can open. */
function clkQrRequireExecUrl_() {
  var u = clkQrExecUrl_();
  if (!u) {
    throw new Error('Clock_QR: no web app URL. Deploy the project, then put the ' +
                    '/exec URL in the QR_EXEC_URL row on the Config tab.');
  }
  if (clkQrIsDevUrl_(u)) {
    throw new Error('Clock_QR: that is the editor\'s /dev URL, which only works for ' +
                    'you. Links built from it would ask every member of staff to sign ' +
                    'in to Google. Put the deployed /exec URL in the QR_EXEC_URL row ' +
                    'on the Config tab and run this again.');
  }
  return u;
}

/**
 * Safe to place inside a <script> block.
 *
 * JSON.stringify escapes quotes and backslashes but NOT "<", so a value
 * containing </script> would close the element and everything after it would be
 * parsed as markup. This escapes anything outside printable ASCII — which also
 * covers U+2028/U+2029, the two separators JavaScript treats as newlines — plus
 * < > and &, as \uXXXX. The emitted literal is pure ASCII, cannot end the
 * element, and still parses back to exactly the original string.
 *
 * In v1.1.0 the personal token no longer reaches the server at all, so the worst
 * case this used to guard is gone. It stays because defence that costs nothing
 * should not be removed on the strength of an argument about reachability.
 */
function clkQrJs_(v) {
  return JSON.stringify(v == null ? '' : String(v))
    .replace(/[^\x20-\x7e]|[<>&]/g, function (ch) {
      return '\\u' + ('000' + ch.charCodeAt(0).toString(16)).slice(-4);
    });
}

function clkQrEsc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function clkQrJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * The doPost hook — added in v1.2.0 when the phone page moved to GitHub Pages.
 *
 * WHY THIS EXISTS, given the 2026-08-17 note said not to widen doPost.
 * The phone page can no longer live on Apps Script. Its sandbox makes
 * localStorage third-party (Safari discards it, so nobody is remembered) and
 * withholds the camera (so a scanner cannot be built). Both were verified on a
 * real iPhone on 2026-08-18. What was left was asking staff to type eight
 * characters every morning, which is a clock-in method people route around.
 *
 * THIS IS A CHANGE OF TRANSPORT, NOT OF CREDENTIAL. Both halves are still
 * required and still checked by the same code: the personal token proves who,
 * the rotating site code proves where. google.script.run was never gated by a
 * device token either — anyone who could load the page could call it. Nothing
 * became reachable that was not reachable before; it simply arrives over HTTP.
 *
 * Keep it that way. This route accepts a personal token and a site code and
 * NOTHING else. It must not grow into a general-purpose door — if a future
 * change cannot pass the transport-not-credential test above, it does not
 * belong here.
 *
 * Returns null for any body that is not ours, so every existing kiosk and
 * display POST reaches the Clock's own handler untouched.
 */
function clkQrDoPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return null; }
  if (!body || String(body.qp || '') !== 'mclk') return null;

  var act = String(body.action || '');
  if (act === 'hello')  return clkQrJson_(clkQrHello(body.t));
  if (act === 'submit') return clkQrJson_(clkQrSubmit({ t: body.t, s: body.s, a: body.a }));
  return clkQrJson_({ ok: false, error: 'Unknown QR action.' });
}

/**
 * The doGet hook. Returns an HtmlOutput for the pages this file owns, or null
 * so the Clock's existing doGet carries on exactly as before.
 */
function clkQrDoGet(e) {
  var p = e && e.parameter ? String(e.parameter.qp || '') : '';
  if (!p) return null;
  if (p === 'qr') return clkQrDisplayPage_(e.parameter);
  if (p === 'mclk') return clkQrMobilePage_(e.parameter);
  if (p === 'print') return clkQrPrintPage_(e.parameter);
  return null;
}

/* ---- the kiosk QR panel ---- */

function clkQrDisplayPage_(params) {
  var token = String(params.d || '').trim();
  var devices = clkQrDevices_(), dev = null, i;
  for (i = 0; i < devices.length; i++) if (clkQrConstEq_(devices[i].token, token)) dev = devices[i];

  if (!dev) {
    return clkQrHtml_('Zanna Clock', '<div class="pad"><h1>Not a clocking device</h1>' +
      '<p>This screen needs the token of an <b>active</b> KIOSK or QR row on the Devices tab.</p></div>');
  }
  if (!clkQrCfgBool_('QR_ENABLED', true)) {
    return clkQrHtml_('Zanna Clock', '<div class="pad"><h1>QR clock-in is off</h1>' +
      '<p>Set <code>QR_ENABLED</code> to TRUE on the Config tab.</p></div>');
  }

  var now = clkQrNowSecs_(), win = clkQrWindowSecs_();
  var code = clkQrCodeFor_(dev.token, Math.floor(now / win));
  var secsLeft = win - (now % win);
  /* v1.2.0: point the QR at the GitHub Pages phone app when QR_MOBILE_URL is
     set. The code travels in the FRAGMENT there — it is not a secret, but a
     fragment keeps it out of Pages' logs and out of Referer headers for free.
     Blank falls back to the Apps Script page, so an un-migrated install still
     works exactly as it did in v1.1.1. */
  var mob = String(clkQrCfg_('QR_MOBILE_URL', '') || '').trim();
  var url = mob ? (mob + '#s=' + code) : (clkQrExecUrl_() + '?qp=mclk&s=' + code);
  var showCode = clkQrCfgBool_('QR_SHOW_CODE', true);

  var body =
    '<div class="qrwrap">' +
      '<div class="qrhead"><span class="dot"></span> Scan to clock in — ' + clkQrEsc_(dev.name) + '</div>' +
      '<div id="qr" class="qr"></div>' +
      (showCode
        ? '<div class="code"><span>' + clkQrEsc_(code.substring(0, 4)) + '</span><i>–</i><span>' +
          clkQrEsc_(code.substring(4)) + '</span></div>' +
          '<div class="hint">No camera? Open your own clock-in link and type this code.</div>'
        : '') +
      '<div class="bar"><div id="fill" class="fill"></div></div>' +
      '<div class="foot">New code every ' + Math.round(win / 60) + ' minutes · <span id="cd"></span></div>' +
    '</div>' +
    '<script>' + clkQrLibJs_() + '</script>' +
    '<script>' +
      'var PAYLOAD=' + clkQrJs_(url) + ',LEFT=' + secsLeft + ',WIN=' + win + ';' +
      'document.getElementById("qr").innerHTML=ZQR.toSvg(PAYLOAD,{quiet:2});' +
      /* Anchored to WALL-CLOCK time, not to a tick counter.
         A tablet that sleeps stops firing timers, so a counter under-reports how
         long the code has been up: twenty real minutes can pass while LEFT falls
         by two. The old wake-up guard then asked "LEFT<=0", got false, and left
         an expired QR on screen looking perfectly healthy — everyone scanning it
         refused, and the printed characters under it stale too. 06:45 on a
         Monday is exactly when a tablet has been asleep all weekend. */
      'var DEADLINE=Date.now()+LEFT*1000;' +
      'function left(){return Math.round((DEADLINE-Date.now())/1000);}' +
      'function refresh(){location.reload();}' +
      'function tick(){' +
        'var L=left();' +
        'if(L<=0){refresh();return;}' +
        'var m=Math.floor(L/60),s=L%60;' +
        'document.getElementById("cd").textContent=m+":"+(s<10?"0":"")+s;' +
        'document.getElementById("fill").style.width=(L/WIN*100)+"%";' +
        'setTimeout(tick,1000);' +
      '}tick();' +
      'document.addEventListener("visibilitychange",function(){if(!document.hidden&&left()<=0)refresh();});' +
      /* pageshow fires when the page returns from the back-forward cache, which
         visibilitychange does not always cover. */
      'window.addEventListener("pageshow",function(){if(left()<=0)refresh();});' +
    '</script>';

  return clkQrHtml_('Clock in — ' + dev.name, body, true);
}

/* ---- the phone page ---- */

/**
 * The phone page.
 *
 * v1.1.0 — the personal token is NO LONGER read from the query string. It
 * arrives in the URL fragment (#t=...), which browsers never send to the
 * server. Three things follow from that:
 *
 *   - it stays out of Google's request logs and out of Referer headers;
 *   - the server cannot embed it in this page, so the page cannot be made to
 *     leak it, which closes the injection hole in v1.0.0 at the root rather
 *     than patching it;
 *   - the personal link itself becomes the durable credential. That matters on
 *     iOS: Apps Script serves this page inside a sandboxed googleusercontent
 *     iframe, so localStorage is THIRD-PARTY storage and Safari may refuse or
 *     evict it. With the token in the link, adding it to the home screen works
 *     regardless — storage is now an optimisation, not a dependency.
 *
 * The fragment is deliberately NOT stripped after reading. Stripping it would
 * look tidier and would break exactly the case above: reopen the icon on a
 * phone whose storage was evicted and there would be nothing left to identify.
 *
 * The site code `s` stays in the query. It is displayed on a wall — it is not a
 * secret, and clkQrNormaliseCode_ reduces it to [0-9A-Z] before it is used.
 */
function clkQrMobilePage_(params) {
  var s = clkQrNormaliseCode_(params.s || '');

  var body =
    '<div class="m">' +
      '<div id="view" class="card"><div class="spin"></div><p class="muted">Checking…</p></div>' +
      '<div class="brand">Zanna Cookhouse</div>' +
    '</div>' +
    '<script>' +
      'var S=' + clkQrJs_(s) + ',T="",ME=null;' +
      'function store(v){try{localStorage.setItem("zannaTok",v);}catch(e){}}' +
      'function recall(){try{return localStorage.getItem("zannaTok")||"";}catch(e){return "";}}' +
      /* v1.1.1 — READ THE FRAGMENT THROUGH google.script.url.
         Apps Script serves this page inside a sandboxed iframe on a
         googleusercontent origin, so `location.hash` in here is the IFRAME's
         URL, not the address bar's. The personal token is therefore invisible
         to it, and v1.1.0 showed every employee "this phone does not have your
         link" no matter how good the link was. google.script.url.getLocation is
         the only way to see the real fragment, and it is asynchronous — so boot
         waits for it rather than running first and finding nothing.
         The token still never reaches the server: getLocation resolves in the
         browser. Requires IFRAME sandbox mode, which is the only mode now. */
      'function readToken(done){' +
        'var settled=false;' +
        'function finish(v){if(settled)return;settled=true;T=v||recall();if(T)store(T);done();}' +
        /* If getLocation never calls back the page would sit on "Checking…"
           forever, so fall through to stored-token-only after 3 seconds. */
        'setTimeout(function(){finish("");},3000);' +
        'try{' +
          'google.script.url.getLocation(function(loc){' +
            'var hs=(loc&&loc.hash)?String(loc.hash):"";' +
            /* getLocation strips the leading "#", so anchor on start-or-& too. */
            'var m=hs.match(/(?:^|[#&])t=([A-Za-z0-9_-]{16,64})/);' +
            'finish(m?m[1]:"");' +
          '});' +
        '}catch(e){finish("");}' +
      '}' +
      'var V=document.getElementById("view");' +
      'function esc(x){var d=document.createElement("div");d.textContent=x==null?"":x;return d.innerHTML;}' +

      'function needLink(){' +
        'V.innerHTML="<h1>Almost there</h1><p>This phone does not have your personal clock-in link yet.</p>' +
        '<p class=\\"muted\\">Open the link your manager sent you once, on this phone, in this browser. ' +
        'After that, scanning the kitchen code is all you need.</p>";' +
      '}' +

      'function fail(msg){V.innerHTML="<h1>Sorry</h1><p>"+esc(msg)+"</p>"+(T?"<button class=\\"btn ghost\\" onclick=\\"boot()\\">Try again</button>":"");}' +

      'function boot(){' +
        'if(!T){needLink();return;}' +
        'V.innerHTML="<div class=\\"spin\\"></div><p class=\\"muted\\">Checking…</p>";' +
        'google.script.run.withSuccessHandler(function(r){' +
          'if(!r||!r.ok){fail(r&&r.error?r.error:"The clock did not answer.");return;}' +
          'ME=r;render();' +
        '}).withFailureHandler(function(e){fail(String(e&&e.message||e));}).clkQrHello(T);' +
      '}' +

      'function render(){' +
        'var greet=(new Date()).getHours()<12?"Good morning":"Hello";' +
        'var h="<h1>"+greet+", "+esc(ME.first)+"</h1>";' +
        'var goingIn=(ME.state!=="IN");' +
        'if(!goingIn&&!ME.allowOut){' +
          'V.innerHTML=h+"<p>You clocked in at <b>"+esc(ME.since)+"</b>.</p>' +
          '<p class=\\"muted\\">Clock out on the kitchen tablet.</p>";return;' +
        '}' +
        'if(!goingIn)h+="<p>Clocked in since <b>"+esc(ME.since)+"</b>.</p>";' +
        'var needScan=goingIn||ME.outNeedsScan;' +
        'if(needScan&&S.length!==8){' +
          'V.innerHTML=h+"<p>Type the 8-character code on the kitchen screen.</p>' +
          '<input id=\\"code\\" class=\\"code\\" inputmode=\\"latin\\" autocapitalize=\\"characters\\" ' +
          'autocomplete=\\"off\\" spellcheck=\\"false\\" maxlength=\\"9\\" placeholder=\\"XXXX-XXXX\\">' +
          '<button class=\\"btn\\" id=\\"go\\">"+(goingIn?"Clock in":"Clock out")+"</button>";' +
          'var inp=document.getElementById("code");inp.focus();' +
          'inp.addEventListener("input",function(){' +
            'var v=inp.value.toUpperCase().replace(/[^0-9A-Z]/g,"");' +
            'inp.value=v.length>4?v.slice(0,4)+"-"+v.slice(4,8):v;' +
          '});' +
          'document.getElementById("go").onclick=function(){S=inp.value.replace(/[^0-9A-Z]/gi,"").toUpperCase();submit(goingIn?"in":"out");};' +
          'return;' +
        '}' +
        'V.innerHTML=h+"<button class=\\"btn big\\" id=\\"go\\">"+(goingIn?"Clock in":"Clock out")+"</button>' +
        '<p class=\\"muted\\">"+(needScan?"Code accepted from the kitchen screen.":"")+"</p>";' +
        'document.getElementById("go").onclick=function(){submit(goingIn?"in":"out");};' +
      '}' +

      'function submit(a){' +
        'var b=document.getElementById("go");if(b){b.disabled=true;b.textContent="Working…";}' +
        'google.script.run.withSuccessHandler(function(r){' +
          'if(!r||!r.ok){' +
            'V.innerHTML="<h1>Not done</h1><p>"+esc(r&&r.error?r.error:"Unknown problem")+"</p>' +
            '<button class=\\"btn ghost\\" onclick=\\"S=\\x27\\x27;boot()\\">Back</button>";return;' +
          '}' +
          'var inn=(r.action==="CLOCK_IN");' +
          'V.innerHTML="<div class=\\"tick\\">✓</div><h1>"+(inn?"Clocked in":"Clocked out")+"</h1>' +
          '<p class=\\"big\\">"+esc(r.at)+"</p><p class=\\"muted\\">"+esc(r.name)+(r.dryRun?" · TEST MODE, nothing recorded":"")+"</p>' +
          '<p class=\\"muted\\">"+(inn?"Have a good shift. Breaks are on the kitchen tablet.":"See you tomorrow.")+"</p>";' +
        '}).withFailureHandler(function(e){fail(String(e&&e.message||e));}).clkQrSubmit({t:T,s:S,a:a});' +
      '}' +
      'readToken(boot);' +
    '</script>';

  return clkQrHtml_('Zanna Clock', body);
}

/* ---- the printable slip sheet ---- */

/**
 * @param params  optional ?id=<EmployeeID> to pre-aim at one person. Convenience
 *                only — it pre-fills the form, it does not skip the PIN. The
 *                Data Maintenance console deep-links a row this way.
 */
function clkQrPrintPage_(params) {
  var pre = String((params && params.id) || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 32);
  var body =
    '<div class="sheetwrap">' +
      '<div class="sheethead no-print">' +
        '<h1>Clock-in slips</h1>' +
        '<p class="muted">Each slip is a working credential — it signs that person in. ' +
          'Print, hand out, watch one person use it, then destroy the spares. ' +
          'The links cannot be shown again.</p>' +
        '<div id="gate">' +
          '<input id="pin" type="password" inputmode="numeric" autocomplete="off" placeholder="Management PIN">' +
          '<select id="mode">' +
            '<option value="missing">Only people without a link</option>' +
            '<option value="all">Everyone — replaces existing links</option>' +
            '<option value="one">One person…</option>' +
          '</select>' +
          '<input id="who" placeholder="EmployeeID, e.g. 7" value="' + clkQrEsc_(pre) + '"' +
            (pre ? '' : ' style="display:none"') + '>' +
          '<button class="btn" id="go">Generate</button>' +
          '<p class="warn" id="warn"></p>' +
        '</div>' +
        '<div id="acts" style="display:none">' +
          '<button class="btn" id="printBtn">🖨️ Print</button>' +
          '<button class="btn ghost" id="againBtn">Start again</button>' +
          '<p class="warn" id="keep">Print this as many times as you need — the sheet stays until ' +
            'you close the tab. <b>Once it is gone the codes cannot be shown again</b>, ' +
            'only reissued, which cuts off anyone already holding a slip.</p>' +
        '</div>' +
      '</div>' +
      '<div id="sheet"></div>' +
    '</div>' +
    '<script>' + clkQrLibJs_() + '</script>' +
    '<script>' +
      /* Pre-aimed at one person from a Data Maintenance row. */
      (pre ? 'document.getElementById("mode").value="one";' : '') +
      'var G=document.getElementById("gate"),A=document.getElementById("acts"),' +
        'SH=document.getElementById("sheet"),W=document.getElementById("warn");' +
      'function esc(x){var d=document.createElement("div");d.textContent=x==null?"":x;return d.innerHTML;}' +
      'document.getElementById("mode").addEventListener("change",function(){' +
        'document.getElementById("who").style.display=(this.value==="one")?"":"none";' +
        'W.textContent=(this.value==="all")?"This replaces EVERY link. Any phone already set up stops working straight away.":"";' +
      '});' +
      'document.getElementById("go").onclick=function(){' +
        'var b=this;b.disabled=true;b.textContent="Working…";W.textContent="";' +
        'google.script.run.withSuccessHandler(function(r){' +
          'b.disabled=false;b.textContent="Generate";' +
          'if(!r||!r.ok){W.textContent=(r&&r.error)?r.error:"No answer from the clock.";return;}' +
          'if(!r.people.length){W.textContent="Nobody matched — everyone active already has a link.";return;}' +
          'draw(r);' +
        '}).withFailureHandler(function(e){b.disabled=false;b.textContent="Generate";' +
          'W.textContent=String(e&&e.message||e);})' +
        '.clkQrIssueSheetData(document.getElementById("pin").value,' +
          'document.getElementById("mode").value,document.getElementById("who").value);' +
      '};' +
      'function draw(r){' +
        'G.style.display="none";A.style.display="";' +
        'var h=\'<div class="sheettitle">Clock-in slips — \'+esc(r.company)+\' — \'+r.people.length+\' of them</div><div class="grid">\';' +
        'for(var i=0;i<r.people.length;i++){var p=r.people[i];' +
          'h+=\'<div class="slip"><div class="q" id="q\'+i+\'"></div>' +
             '<div class="nm">\'+esc(p.name)+\'</div><div class="id">#\'+esc(p.id)+\'</div>' +
             '<div class="tip">Scan once with your camera, then Add to Home Screen</div></div>\';}' +
        'h+="</div>";SH.innerHTML=h;' +
        'for(var j=0;j<r.people.length;j++){' +
          'document.getElementById("q"+j).innerHTML=ZQR.toSvg(r.people[j].link,{quiet:1});}' +
      '}' +
      'document.getElementById("printBtn").onclick=function(){window.print();};' +
      /* Start again DISCARDS the batch. Nothing anywhere can bring those codes
         back, and anyone already holding a printed slip from this run keeps a
         link that only reissuing can replace — which cuts them off in turn. So
         it asks, once, rather than being a quiet one-click loss. */
      'document.getElementById("againBtn").onclick=function(){' +
        'if(confirm("Discard these codes?\\n\\nThey cannot be shown again. Anyone already holding a printed slip from this batch will keep working, but you will have no way to reprint theirs — only reissue, which cuts them off.\\n\\nPrint first if you have not.")) location.reload();' +
      '};' +
    '</script>';
  return clkQrHtml_('Clock-in slips', body, true);
}

/* ---- shared chrome ---- */

function clkQrHtml_(title, body, isDisplay) {
  var css =
    '*{box-sizing:border-box}' +
    'body{margin:0;font:16px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;' +
      'background:#0f1115;color:#eef1f6;-webkit-text-size-adjust:100%}' +
    'h1{font-size:26px;line-height:1.25;margin:0 0 12px;font-weight:650}' +
    'p{margin:0 0 12px}' +
    '.muted{color:#9aa3b2;font-size:14px}' +
    '.pad{padding:28px}' +
    'code{background:#1b1f29;padding:2px 6px;border-radius:5px}' +

    /* phone */
    '.m{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:22px}' +
    '.card{width:100%;max-width:420px;background:#161a22;border:1px solid #232936;border-radius:18px;padding:26px;text-align:center}' +
    '.btn{display:block;width:100%;margin:18px 0 6px;padding:20px;border:0;border-radius:14px;' +
      'background:#2f6df6;color:#fff;font-size:19px;font-weight:600;cursor:pointer}' +
    '.btn:disabled{opacity:.55}' +
    '.btn.big{padding:30px;font-size:23px}' +
    '.btn.ghost{background:#232936}' +
    'input.code{width:100%;padding:18px;font-size:30px;letter-spacing:.16em;text-align:center;' +
      'border-radius:14px;border:1px solid #2b3242;background:#0f1115;color:#eef1f6;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}' +
    '.tick{font-size:52px;color:#3ad07f;line-height:1}' +
    'p.big{font-size:40px;font-weight:650;margin:6px 0}' +
    '.brand{margin-top:22px;color:#5c6577;font-size:12px;letter-spacing:.12em;text-transform:uppercase}' +
    '.spin{width:26px;height:26px;margin:10px auto;border:3px solid #2b3242;border-top-color:#2f6df6;' +
      'border-radius:50%;animation:sp .8s linear infinite}' +
    '@keyframes sp{to{transform:rotate(360deg)}}' +

    /* kiosk panel */
    '.qrwrap{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:18px;padding:24px}' +
    '.qrhead{font-size:clamp(18px,3.4vw,30px);font-weight:600;display:flex;align-items:center;gap:10px}' +
    '.dot{width:11px;height:11px;border-radius:50%;background:#3ad07f;animation:pulse 2s ease-in-out infinite}' +
    '@keyframes pulse{50%{opacity:.25}}' +
    '.qr{width:min(58vh,58vw);background:#fff;padding:14px;border-radius:16px;line-height:0}' +
    '.qr svg{width:100%;height:auto;display:block}' +
    '.code{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:clamp(30px,6.5vw,64px);' +
      'letter-spacing:.1em;display:flex;gap:10px;align-items:center}' +
    '.code i{color:#5c6577;font-style:normal}' +
    '.hint{color:#9aa3b2;font-size:clamp(12px,1.7vw,16px)}' +
    '.bar{width:min(58vh,58vw);height:6px;background:#232936;border-radius:99px;overflow:hidden}' +
    '.fill{height:100%;background:#2f6df6;transition:width 1s linear}' +
    '.foot{color:#5c6577;font-size:clamp(12px,1.7vw,16px)}' +

    /* slip sheet */
    '.sheetwrap{padding:24px;max-width:1000px;margin:0 auto}' +
    '.sheethead h1{font-size:24px;margin-bottom:8px}' +
    '.sheethead .muted{max-width:60ch}' +
    '#gate,#acts{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:16px 0}' +
    '#gate input,#gate select{padding:12px;border-radius:10px;border:1px solid #2b3242;' +
      'background:#0f1115;color:#eef1f6;font:inherit}' +
    '.btn{display:inline-block;width:auto;margin:0;padding:12px 18px;border:0;border-radius:10px;' +
      'background:#2f6df6;color:#fff;font-size:15px;font-weight:600;cursor:pointer}' +
    '.warn{color:#ffcf6b;font-size:13.5px;flex-basis:100%;margin:0}' +
    '.sheettitle{font-size:13px;color:#9aa3b2;margin:10px 0 14px}' +
    '.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}' +
    '.slip{border:1px solid #232936;border-radius:10px;padding:12px;text-align:center;break-inside:avoid}' +
    '.slip .q{background:#fff;padding:8px;border-radius:8px;line-height:0}' +
    '.slip .q svg{width:100%;height:auto;display:block}' +
    '.slip .nm{font-weight:700;font-size:13.5px;margin-top:8px}' +
    '.slip .id{font-size:11.5px;color:#9aa3b2}' +
    '.slip .tip{font-size:10.5px;color:#5c6577;margin-top:4px}' +
    '@media print{' +
      'body{background:#fff;color:#000}' +
      '.no-print{display:none !important}' +
      '.sheetwrap{padding:0;max-width:none}' +
      '.slip{border:1px solid #999;color:#000}' +
      '.slip .id,.slip .tip{color:#444}' +
      '.sheettitle{color:#444}' +
    '}';

  var html =
    '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">' +
    '<meta name="apple-mobile-web-app-capable" content="yes">' +
    '<meta name="theme-color" content="#0f1115">' +
    '<title>' + clkQrEsc_(title) + '</title><style>' + css + '</style></head><body>' +
    body + '</body></html>';

  var out = HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (!isDisplay) out.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  return out;
}


/* ==========================================================================
 * 8 · Setup, issuing links, diagnostics
 * Run these from the Apps Script editor or the spreadsheet menu.
 * ======================================================================== */

/** One-off. Creates secrets, adds the columns, seeds Config with commented defaults. */
function clkQrSetup() {
  var log = [];
  log.push(clkQrEnsureSecrets());
  var s = clkQrEnsureStaffCols_();
  log.push(s.length ? 'Staff columns added: ' + s.join(', ') : 'Staff columns already present.');
  var ev = clkQrEnsureEventCols_();
  log.push(ev.length ? 'Events columns added: ' + ev.join(', ') : 'Events columns already present.');

  var cfg = clkQrConfigSheet_();
  if (cfg) {
    var defaults = [
      ['QR_ENABLED', 'TRUE', 'Master switch for QR clock-in'],
      ['QR_WINDOW_SECS', '600', 'Seconds a QR code stays on screen'],
      ['QR_GRACE_SECS', '90', 'How long the previous code still works after it rolls over'],
      ['QR_SHOW_CODE', 'TRUE', 'Print the 8-character code under the QR (needed for the no-camera route)'],
      ['QR_ALLOW_OUT', 'TRUE', 'Let phones clock OUT as well as in'],
      ['QR_OUT_REQUIRES_SCAN', 'TRUE', 'Clocking out also needs a fresh code from the screen'],
      ['QR_DRY_RUN', 'FALSE', 'TRUE = go through the motions, write nothing to Events'],
      ['QR_EXEC_URL', '', 'Override the /exec URL put in the QR. Blank = work it out'],
      ['QR_MOBILE_URL', '', 'GitHub Pages phone app, e.g. https://dosca61.github.io/ZannaCookhouse/clock-mobile.html — blank = use the Apps Script page (no camera, no iPhone memory)']
    ];
    var existing = {}, i;
    if (cfg.getLastRow() > 0) {
      var have = cfg.getRange(1, 1, cfg.getLastRow(), 1).getValues();
      for (i = 0; i < have.length; i++) existing[String(have[i][0]).trim().toUpperCase()] = true;
    }
    var added = [];
    for (i = 0; i < defaults.length; i++) {
      if (!existing[defaults[i][0]]) { cfg.appendRow(defaults[i]); added.push(defaults[i][0]); }
    }
    log.push(added.length ? 'Config keys added: ' + added.join(', ') : 'Config keys already present.');
  } else {
    log.push('No Config tab found — defaults will be used.');
  }

  var out = log.join('\n');
  Logger.log(out);
  clkQrAudit_('qr.setup', 'v' + CLK_QR_VERSION);
  return out;
}

/**
 * Issue personal clock-in links for everyone active who has not got one.
 *
 * The links are shown ONCE. Only their hash is kept, so this dialog is the only
 * chance to copy them. Close it without copying and the fix is clkQrReissue().
 */
function clkQrIssueLinks() {
  return clkQrIssue_(false);
}

/** Replace everyone's link — invalidates every phone currently set up. */
function clkQrReissueAll() {
  return clkQrIssue_(true);
}

function clkQrIssue_(force) {
  clkQrEnsureStaffCols_();
  var st = clkQrStaffRows_(), sh = st.sheet, c = st.cols;
  var base = clkQrPersonalBase_();
  var lines = [], i, n = 0;

  for (i = 0; i < st.rows.length; i++) {
    var r = st.rows[i];
    if (!r.active) continue;
    if (r.hash && !force) continue;
    var tok = clkQrTokenNew_();
    sh.getRange(r.row, c.hash + 1).setValue(clkQrTokenHash_(tok));
    sh.getRange(r.row, c.issued + 1).setValue(new Date());
    if (c.revoked >= 0) sh.getRange(r.row, c.revoked + 1).setValue('');
    lines.push(r.id + '\t' + r.name + '\t' + base + '#t=' + tok);
    n++;
  }

  clkQrAudit_('qr.issue', n + ' link(s)' + (force ? ' (reissue all)' : ''));
  var text = n
    ? 'Copy these now — they cannot be shown again.\n\nID\tName\tPersonal clock-in link\n' + lines.join('\n')
    : 'Nothing to issue. Everyone active already has a link. Use clkQrReissueAll() to replace them all, or clkQrReissue("E004") for one person.';

  clkQrShow_('Personal clock-in links', text);
  Logger.log(text);
  return text;
}

/** Replace one person's link. Use when a phone is lost or a link goes astray. */
function clkQrReissue(employeeId) {
  clkQrPersonalBase_();            /* fail before writing a hash we cannot use */
  clkQrEnsureStaffCols_();
  var st = clkQrStaffRows_(), sh = st.sheet, c = st.cols;
  for (var i = 0; i < st.rows.length; i++) {
    if (st.rows[i].id !== String(employeeId).trim()) continue;
    var tok = clkQrTokenNew_();
    sh.getRange(st.rows[i].row, c.hash + 1).setValue(clkQrTokenHash_(tok));
    sh.getRange(st.rows[i].row, c.issued + 1).setValue(new Date());
    if (c.revoked >= 0) sh.getRange(st.rows[i].row, c.revoked + 1).setValue('');
    var link = clkQrPersonalBase_() + '#t=' + tok;
    clkQrAudit_('qr.reissue', st.rows[i].id);
    clkQrShow_('New link for ' + st.rows[i].name, link);
    Logger.log(link);
    return link;
  }
  throw new Error('Clock_QR: no staff row with EmployeeID ' + employeeId);
}

/**
 * Register a token that was minted somewhere else — specifically, FairLeave.
 *
 * Per the 2026-08-18 decision, an employee has ONE personal token covering both
 * holidays and clocking. Whichever system issues it calls this so the Clock can
 * recognise it, and the Clock stores only its own HMAC of it under its own
 * pepper. Two systems, two hashes, one token, and a leak of either hash column
 * is useless against the other system.
 *
 * The plaintext token passes through here and is never written anywhere.
 */
function clkQrAdoptToken(employeeId, token) {
  clkQrEnsureStaffCols_();
  token = String(token == null ? '' : token).trim();
  if (token.length < 16) throw new Error('Clock_QR: that token is too short to be real.');
  var st = clkQrStaffRows_(), sh = st.sheet, c = st.cols;
  for (var i = 0; i < st.rows.length; i++) {
    if (st.rows[i].id !== String(employeeId).trim()) continue;
    sh.getRange(st.rows[i].row, c.hash + 1).setValue(clkQrTokenHash_(token));
    sh.getRange(st.rows[i].row, c.issued + 1).setValue(new Date());
    if (c.revoked >= 0) sh.getRange(st.rows[i].row, c.revoked + 1).setValue('');
    clkQrAudit_('qr.adopt', st.rows[i].id + ' — token issued by another system');
    return st.rows[i].name + ' can now clock in with that token.';
  }
  throw new Error('Clock_QR: no staff row with EmployeeID ' + employeeId);
}

/* ==========================================================================
 * 8b · The printable slip sheet
 *
 * A token is written as a HASH and can never be read back, so a QR of somebody's
 * link can only be drawn at the moment it is issued. This mints and renders in
 * one pass: press once, print, hand out, destroy the spares.
 *
 * Served as an HtmlService page and driven with google.script.run, so it needs
 * no doPost route. That matters — the 2026-08-18 decision says the phone route
 * accepts a personal token and a site code and NOTHING else, and a management
 * action does not belong there.
 *
 * Gated by the Clock's own management PIN through clkMgmtAuth_, so it inherits
 * the constant-time compare, the ten-failure lockout and the mgmt.authFail
 * audit trail rather than inventing a second, weaker door.
 * ======================================================================== */

/**
 * Mint and return links for printing. Management PIN required.
 *
 * mode: 'missing' — only people with no link yet. The safe default: it cannot
 *                   invalidate a link somebody is already using.
 *       'all'     — everyone active, replacing existing links. Every phone
 *                   already set up stops working the moment this runs.
 *       'one'     — a single EmployeeID.
 */
function clkQrIssueSheetData(pin, mode, employeeId) {
  if (!clkMgmtAuth_(pin, 'qr.printSheet')) {
    return { ok: false, error: 'That PIN was not accepted.' };
  }
  var base;
  try { base = clkQrPersonalBase_(); }
  catch (err) { return { ok: false, error: err.message }; }

  clkQrEnsureStaffCols_();
  var st = clkQrStaffRows_(), sh = st.sheet, c = st.cols;
  mode = String(mode || 'missing');
  var want = String(employeeId || '').trim();
  var out = [], i;

  for (i = 0; i < st.rows.length; i++) {
    var r = st.rows[i];
    if (!r.active) continue;
    if (mode === 'one') { if (r.id !== want) continue; }
    else if (mode === 'missing' && r.hash) continue;

    var tok = clkQrTokenNew_();
    sh.getRange(r.row, c.hash + 1).setValue(clkQrTokenHash_(tok));
    sh.getRange(r.row, c.issued + 1).setValue(new Date());
    if (c.revoked >= 0) sh.getRange(r.row, c.revoked + 1).setValue('');
    out.push({ id: r.id, name: r.name, dept: r.dept, link: base + '#t=' + tok });
  }

  clkQrAudit_('qr.printSheet', mode + ' — ' + out.length + ' link(s) issued');
  return { ok: true, mode: mode, people: out,
           company: String(clkQrCfg_('COMPANY', 'Zanna Cookhouse')) };
}

/** Kill a link without removing the person from the rota. */
function clkQrRevoke(employeeId) {
  var st = clkQrStaffRows_(), sh = st.sheet, c = st.cols;
  if (c.revoked < 0) throw new Error('Clock_QR: run clkQrSetup() first.');
  for (var i = 0; i < st.rows.length; i++) {
    if (st.rows[i].id !== String(employeeId).trim()) continue;
    sh.getRange(st.rows[i].row, c.revoked + 1).setValue(new Date());
    clkQrAudit_('qr.revoke', st.rows[i].id);
    return st.rows[i].name + "'s link is revoked.";
  }
  throw new Error('Clock_QR: no staff row with EmployeeID ' + employeeId);
}

/**
 * Where a personal link should point.
 *
 * QR_MOBILE_URL (the GitHub Pages app) when set, otherwise the Apps Script page.
 * Both carry the token in the fragment, so it never reaches a server either way.
 */
function clkQrPersonalBase_() {
  var mob = String(clkQrCfg_('QR_MOBILE_URL', '') || '').trim();
  if (mob) return mob;
  return clkQrRequireExecUrl_() + '?qp=mclk';
}

/** The URL to open on each QR screen. */
function clkQrScreenUrls() {
  var devices = clkQrDevices_(), base = clkQrExecUrl_(), out = [];
  for (var i = 0; i < devices.length; i++) {
    out.push(devices[i].name + ' (' + devices[i].type + ')\n  ' + base + '?qp=qr&d=' + devices[i].token);
  }
  var text = out.length ? out.join('\n\n') : 'No ACTIVE device with type KIOSK or QR on the Devices tab.';
  clkQrShow_('QR screen URLs', text);
  Logger.log(text);
  return text;
}

/**
 * Read-only pre-flight. Reports what it matched and the row it WOULD write.
 * Run this before go-live, and again after any change to the Events tab.
 */
function clkQrDiagnose() {
  var out = ['Clock_QR v' + CLK_QR_VERSION, ''];

  out.push('Secrets');
  out.push('  ' + CLK_QR_P_SECRET + ': ' + (clkQrProp_(CLK_QR_P_SECRET) ? 'set (' + clkQrProp_(CLK_QR_P_SECRET).length + ' chars)' : 'MISSING'));
  out.push('  ' + CLK_QR_P_PEPPER + ': ' + (clkQrProp_(CLK_QR_P_PEPPER) ? 'set (' + clkQrProp_(CLK_QR_P_PEPPER).length + ' chars)' : 'MISSING'));

  out.push('', 'Settings');
  out.push('  QR_ENABLED=' + clkQrCfgBool_('QR_ENABLED', true) +
           '  window=' + clkQrWindowSecs_() + 's  grace=' + clkQrGraceSecs_() + 's');
  out.push('  QR_ALLOW_OUT=' + clkQrCfgBool_('QR_ALLOW_OUT', true) +
           '  QR_OUT_REQUIRES_SCAN=' + clkQrCfgBool_('QR_OUT_REQUIRES_SCAN', true) +
           '  QR_DRY_RUN=' + clkQrCfgBool_('QR_DRY_RUN', false));
  var execUrl = clkQrExecUrl_();
  out.push('  exec URL: ' + (execUrl || 'UNKNOWN — set QR_EXEC_URL'));
  if (clkQrIsDevUrl_(execUrl)) {
    out.push('  ✖ THAT IS THE /dev URL. It works only for you, signed in.');
    out.push('    Every staff link and every QR built from it would demand a Google');
    out.push('    sign-in. Put the deployed /exec URL in the QR_EXEC_URL Config row.');
  }

  var mob = String(clkQrCfg_('QR_MOBILE_URL', '') || '').trim();
  out.push('  phone app: ' + (mob || 'NOT SET — using the Apps Script page'));
  if (!mob) {
    out.push('    On that page Safari discards the stored token and the camera is');
    out.push('    blocked by the sandbox, so staff must type the code every time.');
    out.push('    Set QR_MOBILE_URL to the GitHub Pages app to get scanning.');
  }

  out.push('', 'Devices that can host a QR');
  var devices = clkQrDevices_(), i;
  if (!devices.length) out.push('  NONE — needs an ACTIVE row of type KIOSK or QR');
  for (i = 0; i < devices.length; i++) {
    out.push('  ' + devices[i].name + ' [' + devices[i].type + '] code now: ' +
      clkQrCodeFor_(devices[i].token, clkQrWindowAt_(clkQrNowSecs_())));
  }

  out.push('', 'Staff');
  var st = clkQrStaffRows_(), active = 0, withLink = 0, revoked = 0;
  for (i = 0; i < st.rows.length; i++) {
    if (!st.rows[i].active) continue;
    active++;
    if (st.rows[i].hash) withLink++;
    if (st.rows[i].revoked) revoked++;
  }
  out.push('  ' + active + ' active · ' + withLink + ' with a link · ' + revoked + ' revoked');
  if (withLink < active) out.push('  ' + (active - withLink) + ' still need one — run clkQrIssueLinks()');

  out.push('', 'Code.gs — this file cannot work without it');
  var need = ['currentStateFor_', 'staffRow_', 'withLock_', 'todaysEvents_', 'tab_', 'audit_'];
  var missing = [];
  for (i = 0; i < need.length; i++) {
    var present = false;
    try { present = (typeof this[need[i]] === 'function'); } catch (e2) { present = false; }
    if (!present) {
      try { present = (typeof eval(need[i]) === 'function'); } catch (e3) { present = false; }
    }
    if (!present) missing.push(need[i]);
  }
  out.push(missing.length
    ? '  MISSING: ' + missing.join(', ') + '  ← wrong project, or Code.gs was replaced'
    : '  all six helpers present ✓');

  out.push('', 'Events tab  <-- CHECK THIS');
  var evSh = tab_('Events');
  var head = evSh.getRange(1, 1, 1, evSh.getLastColumn()).getValues()[0];
  out.push('  header: [' + head.join(', ') + ']');
  /* v1.1.0 no longer resolves these by alias. It writes the kiosk's own row,
     positionally, so what matters is that the first seven columns are still
     the seven mutate_ writes — not that some alias happened to match. */
  var expect = ['Timestamp', 'Name', 'EventType', 'BreakType', 'SessionID', 'Premises', 'Device'];
  var shapeOk = true;
  for (i = 0; i < expect.length; i++) {
    var got = String(head[i] || '').trim();
    if (got.toLowerCase() !== expect[i].toLowerCase()) {
      shapeOk = false;
      out.push('  MISMATCH col ' + (i + 1) + ': expected "' + expect[i] + '", found "' + got + '"');
    }
  }
  out.push(shapeOk
    ? '  first seven columns match the kiosk\'s row shape ✓'
    : '  FATAL: the Events shape has changed. mutate_ in Code.gs writes these seven ' +
      'positionally too, so the kiosk is equally affected — fix Code.gs and this together.');
  var mCol = clkQrMethodCol_();
  out.push('  Method column: ' + (mCol < 0 ? 'absent (optional — run clkQrSetup() to add it)'
                                           : 'col ' + (mCol + 1) + ' ✓'));

  out.push('', 'Sample row it would append (nothing written)');
  var sampleName = 'Diagnostic';
  var sample = [new Date(), sampleName, 'CLOCK_IN', '',
                sampleName + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '-001-SHIFT',
                '', devices.length ? devices[0].name : 'KIOSK'];
  if (mCol >= 0) { while (sample.length < mCol) sample.push(''); sample[mCol] = 'MOBILE_QR'; }
  var parts = [];
  for (i = 0; i < head.length; i++) parts.push(head[i] + '=' + (sample[i] === '' || sample[i] === undefined ? '∅' : sample[i]));
  out.push('  ' + parts.join(' · '));

  var text = out.join('\n');
  Logger.log(text);
  clkQrShow_('Clock_QR diagnosis', text);
  return text;
}

/** Prove a code round-trips, without touching Events. */
function clkQrSelfTest() {
  var out = [], devices = clkQrDevices_();
  if (!devices.length) return 'No QR-capable device to test against.';
  var d = devices[0], now = clkQrNowSecs_(), win = clkQrWindowSecs_(), cur = Math.floor(now / win);

  var live = clkQrCodeFor_(d.token, cur);
  out.push('current code ' + live + ' -> ' + (clkQrResolveCode_(live).ok ? 'ACCEPTED ✓' : 'REJECTED ✗'));
  out.push('lower case + dash "' + live.substring(0, 4).toLowerCase() + '-' + live.substring(4).toLowerCase() + '" -> ' +
    (clkQrResolveCode_(live.substring(0, 4).toLowerCase() + '-' + live.substring(4).toLowerCase()).ok ? 'ACCEPTED ✓' : 'REJECTED ✗'));

  var old = clkQrCodeFor_(d.token, cur - 3);
  out.push('code from 3 windows ago -> ' + (clkQrResolveCode_(old).ok ? 'ACCEPTED ✗ (should be rejected)' : 'REJECTED ✓'));
  out.push('junk "ZZZZZZZZ" -> ' + (clkQrResolveCode_('ZZZZZZZZ').ok ? 'ACCEPTED ✗' : 'REJECTED ✓'));
  out.push('too short "ABC" -> ' + (clkQrResolveCode_('ABC').ok ? 'ACCEPTED ✗' : 'REJECTED ✓'));

  var tok = clkQrTokenNew_();
  out.push('fresh token is ' + tok.length + ' chars; unknown token resolves -> ' +
    (clkQrEmployeeForToken_(tok) ? 'FOUND ✗' : 'not found ✓'));

  var text = out.join('\n');
  Logger.log(text);
  clkQrShow_('Clock_QR self-test', text);
  return text;
}

/** The exact edit to make in Code.gs. */
function clkQrInstallHelp() {
  var text =
    'TWO changes to Code.gs — it was one until v1.2.0, so do not stop at the first.\n\n' +
    '1. At the very top of doGet(e):\n\n' +
    '  function doGet(e) {\n' +
    '    var qr = clkQrDoGet(e);      // <-- add\n' +
    '    if (qr) return qr;           // <-- add\n' +
    '    ... everything you already have ...\n' +
    '  }\n\n' +
    '2. At the very top of doPost(e), BEFORE the deviceFor_ check:\n\n' +
    '  function doPost(e) {\n' +
    '    var qr = clkQrDoPost(e);     // <-- add\n' +
    '    if (qr) return qr;           // <-- add\n' +
    '    let req;\n' +
    '    ... everything you already have ...\n' +
    '  }\n\n' +
    'Both hooks return null for anything that is not theirs, so every existing\n' +
    'kiosk, display and management call is untouched. The doPost one must go\n' +
    'ABOVE deviceFor_, because the phone has a person token, not a device token.\n\n' +
    'Then set QR_MOBILE_URL on the Config tab to the GitHub Pages phone app, or\n' +
    'staff will be typing the code by hand forever.\n\n' +
    'Optional — add to the ⏰ Zanna Clock menu in onOpen():\n' +
    "  .addSeparator()\n" +
    "  .addSubMenu(SpreadsheetApp.getUi().createMenu('📱 QR clock-in')\n" +
    "    .addItem('Screen URLs', 'clkQrScreenUrls')\n" +
    "    .addItem('Issue personal links', 'clkQrIssueLinks')\n" +
    "    .addItem('🩺 Diagnose', 'clkQrDiagnose')\n" +
    "    .addItem('🩺 Self-test', 'clkQrSelfTest'))\n\n" +
    'To print clock-in slips: open <your /exec URL>?qp=print in a browser. It asks\n' +
    'for the management PIN, mints the links and draws one QR per person, ready to\n' +
    'guillotine. Each slip is a working credential — hand out, then destroy spares.\n\n' +
    'Then DEPLOY A NEW VERSION. The web app runs the deployed code, not the\n' +
    'editor code — the same trap the Go-Live Runbook flags for the relay.';
  Logger.log(text);
  clkQrShow_('Installing Clock_QR', text);
  return text;
}

/** Show text in a dialog when there is a UI, otherwise fall through to the log. */
function clkQrShow_(title, text) {
  try {
    var html = HtmlService.createHtmlOutput(
      '<div style="font:13px/1.5 ui-monospace,Menlo,monospace;white-space:pre-wrap;' +
      'word-break:break-all;padding:10px">' + clkQrEsc_(text) + '</div>')
      .setWidth(720).setHeight(520);
    SpreadsheetApp.getUi().showModalDialog(html, title);
  } catch (e) { /* no UI — the Logger has it */ }
}


/* ==========================================================================
 * 9 · Embedded QR encoder
 * Inlined on purpose. A CDN that is slow or blocked at 06:45 would mean no QR
 * on the screen and nobody able to clock in.
 * ======================================================================== */

function clkQrLibSrc_() {
  /* zqr — minimal byte-mode QR encoder, EC level M, versions 1-10.
     No dependencies. Returns a boolean matrix. ~ISO/IEC 18004.
     Written for the Zanna Clock kiosk so the QR panel never depends on a CDN. */
  window.ZQR = (function () {
    'use strict';

    /* version -> [ecCodewordsPerBlock, g1Blocks, g1DataCw, g2Blocks, g2DataCw] for EC level M */
    var ECM = {
      1:  [10, 1, 16, 0, 0],
      2:  [16, 1, 28, 0, 0],
      3:  [26, 1, 44, 0, 0],
      4:  [18, 2, 32, 0, 0],
      5:  [24, 2, 43, 0, 0],
      6:  [16, 4, 27, 0, 0],
      7:  [18, 4, 31, 0, 0],
      8:  [22, 2, 38, 2, 39],
      9:  [22, 3, 36, 2, 37],
      10: [26, 4, 43, 1, 44]
    };

    var ALIGN = {
      1: [], 2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30],
      6: [6, 34], 7: [6, 22, 38], 8: [6, 24, 42], 9: [6, 26, 46], 10: [6, 28, 50]
    };

    var REMAINDER = { 1: 0, 2: 7, 3: 7, 4: 7, 5: 7, 6: 7, 7: 0, 8: 0, 9: 0, 10: 0 };

    /* ---- GF(256), primitive polynomial 0x11D ---- */
    var EXP = new Array(512), LOG = new Array(256);
    (function () {
      var x = 1;
      for (var i = 0; i < 255; i++) {
        EXP[i] = x; LOG[x] = i;
        x <<= 1;
        if (x & 0x100) x ^= 0x11d;
      }
      for (var j = 255; j < 512; j++) EXP[j] = EXP[j - 255];
    })();

    function gmul(a, b) {
      if (a === 0 || b === 0) return 0;
      return EXP[LOG[a] + LOG[b]];
    }

    function rsGenerator(n) {
      var g = [1];
      for (var i = 0; i < n; i++) {
        var ng = new Array(g.length + 1);
        for (var k = 0; k < ng.length; k++) ng[k] = 0;
        for (var j = 0; j < g.length; j++) {
          ng[j] ^= gmul(g[j], 1);
          ng[j + 1] ^= gmul(g[j], EXP[i]);
        }
        g = ng;
      }
      return g;
    }

    function rsEncode(data, ecLen) {
      var gen = rsGenerator(ecLen);
      var res = new Array(ecLen);
      for (var i = 0; i < ecLen; i++) res[i] = 0;
      for (var d = 0; d < data.length; d++) {
        var factor = data[d] ^ res[0];
        res.shift();
        res.push(0);
        if (factor !== 0) {
          for (var j = 0; j < ecLen; j++) res[j] ^= gmul(gen[j + 1], factor);
        }
      }
      return res;
    }

    /* ---- UTF-8 ---- */
    function utf8Bytes(str) {
      var out = [], i, c;
      for (i = 0; i < str.length; i++) {
        c = str.charCodeAt(i);
        if (c < 0x80) out.push(c);
        else if (c < 0x800) { out.push(0xc0 | (c >> 6), 0x80 | (c & 63)); }
        else if (c >= 0xd800 && c <= 0xdbff && i + 1 < str.length) {
          var c2 = str.charCodeAt(i + 1);
          var cp = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
          out.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 63), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
          i++;
        } else { out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63)); }
      }
      return out;
    }

    function capacityBytes(v) {
      var s = ECM[v];
      var dataCw = s[1] * s[2] + s[3] * s[4];
      var bits = dataCw * 8 - 4 - (v >= 10 ? 16 : 8);
      return Math.floor(bits / 8);
    }

    function pickVersion(len) {
      for (var v = 1; v <= 10; v++) if (capacityBytes(v) >= len) return v;
      return -1;
    }

    /* ---- bit stream ---- */
    function buildCodewords(bytes, v) {
      var s = ECM[v];
      var dataCw = s[1] * s[2] + s[3] * s[4];
      var bits = [];
      function push(val, n) { for (var i = n - 1; i >= 0; i--) bits.push((val >> i) & 1); }

      push(4, 4);                       /* byte mode */
      push(bytes.length, v >= 10 ? 16 : 8);
      for (var i = 0; i < bytes.length; i++) push(bytes[i], 8);

      var cap = dataCw * 8;
      var term = Math.min(4, cap - bits.length);
      for (var t = 0; t < term; t++) bits.push(0);
      while (bits.length % 8 !== 0) bits.push(0);

      var cw = [];
      for (var b = 0; b < bits.length; b += 8) {
        var val = 0;
        for (var k = 0; k < 8; k++) val = (val << 1) | bits[b + k];
        cw.push(val);
      }
      var pads = [0xec, 0x11], p = 0;
      while (cw.length < dataCw) cw.push(pads[p++ % 2]);
      return cw;
    }

    function interleave(cw, v) {
      var s = ECM[v], ecLen = s[0];
      var blocks = [], ecBlocks = [], pos = 0, i;
      for (i = 0; i < s[1]; i++) { blocks.push(cw.slice(pos, pos + s[2])); pos += s[2]; }
      for (i = 0; i < s[3]; i++) { blocks.push(cw.slice(pos, pos + s[4])); pos += s[4]; }
      for (i = 0; i < blocks.length; i++) ecBlocks.push(rsEncode(blocks[i], ecLen));

      var out = [], maxLen = Math.max(s[2], s[4] || 0), j;
      for (i = 0; i < maxLen; i++)
        for (j = 0; j < blocks.length; j++)
          if (i < blocks[j].length) out.push(blocks[j][i]);
      for (i = 0; i < ecLen; i++)
        for (j = 0; j < ecBlocks.length; j++) out.push(ecBlocks[j][i]);
      return out;
    }

    /* ---- matrix ---- */
    function newMatrix(size) {
      var m = new Array(size), i, j;
      for (i = 0; i < size; i++) { m[i] = new Array(size); for (j = 0; j < size; j++) m[i][j] = null; }
      return m;
    }

    function placeFunction(m, v) {
      var size = m.length, i, j;

      function finder(r, c) {
        for (i = -1; i <= 7; i++) for (j = -1; j <= 7; j++) {
          var rr = r + i, cc = c + j;
          if (rr < 0 || rr >= size || cc < 0 || cc >= size) continue;
          var on = (i >= 0 && i <= 6 && (j === 0 || j === 6)) ||
                   (j >= 0 && j <= 6 && (i === 0 || i === 6)) ||
                   (i >= 2 && i <= 4 && j >= 2 && j <= 4);
          m[rr][cc] = on;
        }
      }
      finder(0, 0); finder(0, size - 7); finder(size - 7, 0);

      for (i = 8; i < size - 8; i++) { m[6][i] = (i % 2 === 0); m[i][6] = (i % 2 === 0); }

      var ap = ALIGN[v], na = ap.length;
      for (i = 0; i < na; i++) for (j = 0; j < na; j++) {
        /* the three corners are occupied by finder patterns */
        if ((i === 0 && j === 0) || (i === 0 && j === na - 1) || (i === na - 1 && j === 0)) continue;
        var r = ap[i], c = ap[j];
        for (var dr = -2; dr <= 2; dr++) for (var dc = -2; dc <= 2; dc++)
          m[r + dr][c + dc] = (Math.max(Math.abs(dr), Math.abs(dc)) !== 1);
      }

      m[size - 8][8] = true; /* dark module */

      /* reserve format areas */
      for (i = 0; i <= 8; i++) {
        if (m[8][i] === null) m[8][i] = false;
        if (m[i][8] === null) m[i][8] = false;
      }
      for (i = 0; i < 8; i++) {
        if (m[8][size - 1 - i] === null) m[8][size - 1 - i] = false;
        if (m[size - 1 - i][8] === null) m[size - 1 - i][8] = false;
      }
      if (v >= 7) {
        for (i = 0; i < 6; i++) for (j = 0; j < 3; j++) {
          m[size - 11 + j][i] = false;
          m[i][size - 11 + j] = false;
        }
      }
    }

    function reservedMask(v, size) {
      var r = newMatrix(size);
      placeFunction(r, v);
      return r;
    }

    function placeData(m, reserved, bits) {
      var size = m.length, idx = 0, up = true;
      for (var col = size - 1; col > 0; col -= 2) {
        if (col === 6) col--;
        for (var n = 0; n < size; n++) {
          var row = up ? size - 1 - n : n;
          for (var k = 0; k < 2; k++) {
            var c = col - k;
            if (reserved[row][c] !== null) continue;
            m[row][c] = idx < bits.length ? bits[idx] === 1 : false;
            idx++;
          }
        }
        up = !up;
      }
    }

    function maskFn(i) {
      switch (i) {
        case 0: return function (r, c) { return (r + c) % 2 === 0; };
        case 1: return function (r) { return r % 2 === 0; };
        case 2: return function (r, c) { return c % 3 === 0; };
        case 3: return function (r, c) { return (r + c) % 3 === 0; };
        case 4: return function (r, c) { return (Math.floor(r / 2) + Math.floor(c / 3)) % 2 === 0; };
        case 5: return function (r, c) { return ((r * c) % 2) + ((r * c) % 3) === 0; };
        case 6: return function (r, c) { return (((r * c) % 2) + ((r * c) % 3)) % 2 === 0; };
        default: return function (r, c) { return (((r + c) % 2) + ((r * c) % 3)) % 2 === 0; };
      }
    }

    function bchFormat(data) {
      var d = data << 10, g = 0x537;
      for (var i = 14; i >= 10; i--) if ((d >> i) & 1) d ^= g << (i - 10);
      return ((data << 10) | d) ^ 0x5412;
    }

    function bchVersion(v) {
      var d = v << 12, g = 0x1f25;
      for (var i = 17; i >= 12; i--) if ((d >> i) & 1) d ^= g << (i - 12);
      return (v << 12) | d;
    }

    function writeFormat(m, mask) {
      var size = m.length;
      var bitsVal = bchFormat((0 << 3) | mask); /* EC level M = 00 */
      var b = [];
      for (var i = 14; i >= 0; i--) b.push((bitsVal >> i) & 1);
      /* b[0] is bit 14 (MSB) ... b[14] is bit 0 */
      function get(n) { return b[14 - n] === 1; }   /* get(n) = bit n, LSB-indexed */

      /* first copy: down column 8, then left along row 8 */
      for (var k = 0; k <= 5; k++) m[k][8] = get(k);
      m[7][8] = get(6);
      m[8][8] = get(7);
      m[8][7] = get(8);
      for (k = 9; k <= 14; k++) m[8][14 - k] = get(k);

      /* second copy: leftwards along row 8 on the right edge, then down column 8 at the bottom */
      for (k = 0; k <= 7; k++) m[8][size - 1 - k] = get(k);
      for (k = 8; k <= 14; k++) m[size - 15 + k][8] = get(k);

      m[size - 8][8] = true;
    }

    function writeVersion(m, v) {
      if (v < 7) return;
      var size = m.length, bitsVal = bchVersion(v);
      for (var i = 0; i < 18; i++) {
        var bit = ((bitsVal >> i) & 1) === 1;
        var r = Math.floor(i / 3), c = i % 3;
        m[size - 11 + c][r] = bit;
        m[r][size - 11 + c] = bit;
      }
    }

    function penalty(m) {
      var size = m.length, score = 0, i, j, run, dark = 0;

      for (i = 0; i < size; i++) {
        run = 1;
        for (j = 1; j < size; j++) {
          if (m[i][j] === m[i][j - 1]) { run++; }
          else { if (run >= 5) score += 3 + (run - 5); run = 1; }
        }
        if (run >= 5) score += 3 + (run - 5);
        run = 1;
        for (j = 1; j < size; j++) {
          if (m[j][i] === m[j - 1][i]) { run++; }
          else { if (run >= 5) score += 3 + (run - 5); run = 1; }
        }
        if (run >= 5) score += 3 + (run - 5);
      }

      for (i = 0; i < size - 1; i++) for (j = 0; j < size - 1; j++) {
        var a = m[i][j];
        if (a === m[i][j + 1] && a === m[i + 1][j] && a === m[i + 1][j + 1]) score += 3;
      }

      var pat1 = [true, false, true, true, true, false, true, false, false, false, false];
      var pat2 = [false, false, false, false, true, false, true, true, true, false, true];
      function matches(get, start, pat) {
        for (var k = 0; k < 11; k++) if (get(start + k) !== pat[k]) return false;
        return true;
      }
      for (i = 0; i < size; i++) {
        for (j = 0; j <= size - 11; j++) {
          var rowGet = (function (r) { return function (x) { return m[r][x]; }; })(i);
          var colGet = (function (c) { return function (x) { return m[x][c]; }; })(i);
          if (matches(rowGet, j, pat1) || matches(rowGet, j, pat2)) score += 40;
          if (matches(colGet, j, pat1) || matches(colGet, j, pat2)) score += 40;
        }
      }

      for (i = 0; i < size; i++) for (j = 0; j < size; j++) if (m[i][j]) dark++;
      var total = size * size;
      score += (Math.ceil(Math.abs(dark * 20 - total * 10) / total) - 1) * 10;
      return score;
    }

    function encode(text) {
      var bytes = utf8Bytes(text);
      var v = pickVersion(bytes.length);
      if (v < 0) throw new Error('ZQR: payload too long (' + bytes.length + ' bytes, max ' + capacityBytes(10) + ')');
      var size = 17 + v * 4;
      var cw = interleave(buildCodewords(bytes, v), v);

      var bits = [];
      for (var i = 0; i < cw.length; i++) for (var k = 7; k >= 0; k--) bits.push((cw[i] >> k) & 1);
      for (var r = 0; r < REMAINDER[v]; r++) bits.push(0);

      var reserved = reservedMask(v, size);
      var best = null, bestScore = Infinity;

      for (var mask = 0; mask < 8; mask++) {
        var m = newMatrix(size);
        placeFunction(m, v);
        placeData(m, reserved, bits);
        var f = maskFn(mask);
        for (var y = 0; y < size; y++) for (var x = 0; x < size; x++)
          if (reserved[y][x] === null && f(y, x)) m[y][x] = !m[y][x];
        writeFormat(m, mask);
        writeVersion(m, v);
        var s = penalty(m);
        if (s < bestScore) { bestScore = s; best = m; }
      }
      return { version: v, size: size, modules: best };
    }

    function toSvg(text, opts) {
      opts = opts || {};
      var q = opts.quiet == null ? 4 : opts.quiet;
      var qr = encode(text);
      var n = qr.size, total = n + q * 2, d = '';
      for (var y = 0; y < n; y++) for (var x = 0; x < n; x++)
        if (qr.modules[y][x]) d += 'M' + (x + q) + ' ' + (y + q) + 'h1v1h-1z';
      return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + total + ' ' + total +
        '" shape-rendering="crispEdges" role="img" aria-label="Clock-in QR code">' +
        '<rect width="' + total + '" height="' + total + '" fill="' + (opts.bg || '#ffffff') + '"/>' +
        '<path d="' + d + '" fill="' + (opts.fg || '#000000') + '"/></svg>';
    }

    return { encode: encode, toSvg: toSvg, capacityBytes: capacityBytes };
  })();


}

function clkQrLibJs_() { return "(" + clkQrLibSrc_.toString() + ")();"; }
function myTestLink() {
  return clkQrReissue('14');
}
