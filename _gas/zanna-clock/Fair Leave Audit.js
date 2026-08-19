/* ===========================================================================
   Zanna Clock — FAIRLEAVE BOOKING AUDIT  (v1.0)

   PASTE AS A NEW FILE.  Name it exactly:  Leave_Audit.gs
   Read-only. It never writes to either spreadsheet.

   ---------------------------------------------------------------------------
   WHY

   "🩺 Diagnose FairLeave link" proves the COLUMNS resolve. It does not prove
   the lookup ever returns anybody, because it can only ask about today — and
   if nobody is on holiday today, an empty answer is indistinguishable from the
   bug it was written to catch. That bug returned an empty list and looked
   perfectly healthy for weeks.

   This asks the same question about ANY date, and lists every booking with the
   name its employee id resolves to, so the whole path is visible at once.

   ---------------------------------------------------------------------------
   HONESTY ABOUT WHAT THIS PROVES

   absencesToday_() hardcodes today's date, so it cannot be pointed at another
   day. leaveAudit() therefore MIRRORS its logic rather than calling it — but
   it reuses the same aliasMap_, LEAVE_ALIASES, leaveNameMap_ and the identical
   status pattern, and then CROSS-CHECKS itself: it runs the mirror for today
   and compares the answer with the real absencesToday_(). If those two agree,
   the mirror is faithful and the answer for any other date can be trusted. If
   they disagree, the report says so and you should believe neither.

   ---------------------------------------------------------------------------
   HOW TO USE

     Set LEAVE_AUDIT_DATE to a date where you KNOW someone was on leave, then
     run leaveAudit() and read the Execution log. Blank = today.

         var LEAVE_AUDIT_DATE = '2026-07-02';
   =========================================================================== */

var LEAVE_AUDIT_DATE = '2026-07-02';        // 'yyyy-MM-dd', or '' for today

/* ---------------------------------------------------------------------------
   This file is a COMPANION to the Zanna Clock's Code.gs. It calls cfg_,
   aliasMap_, leaveNameMap_, findLeaveTab_ and absencesToday_, all of which live
   there. Apps Script only shares a namespace within ONE project, so pasted into
   FairLeave or the Shift Scheduler it dies with a bare "cfg_ is not defined".
   Say so properly instead.
   --------------------------------------------------------------------------- */
function laHome_() {
  var need = [
    ['cfg_',            typeof cfg_            === 'function'],
    ['tab_',            typeof tab_            === 'function'],
    ['aliasMap_',       typeof aliasMap_       === 'function'],
    ['LEAVE_ALIASES',   typeof LEAVE_ALIASES   === 'object' && !!LEAVE_ALIASES],
    ['leaveNameMap_',   typeof leaveNameMap_   === 'function'],
    ['findLeaveTab_',   typeof findLeaveTab_   === 'function'],
    ['absencesToday_',  typeof absencesToday_  === 'function']
  ];
  var missing = [];
  for (var i = 0; i < need.length; i++) if (!need[i][1]) missing.push(need[i][0]);
  if (!missing.length) return '';

  var where = 'unknown';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    where = ss ? ss.getName() : 'no bound spreadsheet';
  } catch (e) {}

  return [
    '✖ WRONG PROJECT.',
    '',
    '  This file belongs in the ZANNA CLOCK Apps Script project — the one whose',
    '  Code.gs holds the kiosk. It is currently bound to: "' + where + '"',
    '',
    '  Missing from this project: ' + missing.join(', '),
    '',
    '  Apps Script shares a namespace across the files of ONE project only, so a',
    '  companion file pasted into a different project cannot see anything.',
    '',
    '  FIX: delete this file here. Open the Zanna Clock spreadsheet',
    '       (1q6oBleuq1CN4xnwFsowBhgZRT9ob3g1QSsQ0o-LiXKI) →',
    '       Extensions → Apps Script → + → Script → paste → run leaveAudit().',
    '',
    '  You are in the right project if the Files list already shows a Code.gs',
    '  containing doPost and rebuildSessions_.'
  ].join('\n');
}

function laPad_(v, n) {
  var s = String(v == null ? '' : v);
  while (s.length < n) s += ' ';
  return s;
}

/** Midday, so daylight-saving shifts cannot push a date onto the wrong day. */
function laDay_(v) {
  var d = (v instanceof Date) ? new Date(v.getTime()) : new Date(v);
  if (isNaN(d)) return null;
  d.setHours(12, 0, 0, 0);
  return d;
}

/**
 * The mirror. Same columns, same id→name map, same status test as
 * absencesToday_ — only the date is a parameter instead of "now".
 * Returns { names: {lower: 'HOLIDAY'}, rows: [...], unresolved: n, ... }
 */
function laMirror_(when) {
  var res = { ok: false, why: '', tab: '', map: null, idMapSize: 0,
              rows: [], names: {}, unresolved: 0, byExactName: false, homeless: false };

  var wrong = laHome_();
  if (wrong) { res.why = wrong; res.homeless = true; return res; }

  var id = cfg_('LEAVE_SHEET_ID');
  if (!id) { res.why = 'LEAVE_SHEET_ID is not set — the holiday lookup is OFF.'; return res; }

  var ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { res.why = 'Cannot open the FairLeave sheet: ' + e.message; return res; }

  var sh = ss.getSheetByName('Bookings');
  res.byExactName = !!sh;
  if (!sh) sh = findLeaveTab_(ss);
  if (!sh) { res.why = 'No Bookings tab found in FairLeave.'; return res; }
  res.tab = sh.getName();

  var data = sh.getDataRange().getValues();
  var map = aliasMap_(data[0], LEAVE_ALIASES);
  res.map = map;

  var canName = (map.name !== undefined || map.empId !== undefined);
  if (!canName || map.from === undefined || map.to === undefined) {
    res.why = 'Bookings tab has no usable columns — detected: ' + JSON.stringify(map);
    return res;
  }

  var idMap = (map.empId !== undefined) ? leaveNameMap_(ss) : {};
  res.idMapSize = Object.keys(idMap).length;

  var target = laDay_(when);
  for (var r = 1; r < data.length; r++) {
    var raw = data[r];
    if (!raw || (!raw[0] && !raw[map.from])) continue;

    var status = map.status !== undefined ? String(raw[map.status] || '') : '(no status column)';
    var counts = map.status === undefined || /approv|active|confirm/i.test(status);

    var key = map.empId !== undefined ? String(raw[map.empId] || '').trim() : '';
    var name = map.name !== undefined ? String(raw[map.name] || '').trim() : '';
    var resolvedFrom = name ? 'name column' : '';
    if (!name && key) {
      name = idMap[key] || '';
      resolvedFrom = name ? 'id → ' + key : '';
      if (!name && counts) res.unresolved++;
    }

    var f = laDay_(raw[map.from]);
    var t = laDay_(raw[map.to]);
    var covers = !!(f && t && target && target >= f && target <= t);

    res.rows.push({ id: String(raw[0] || ''), key: key, name: name, from: f, to: t,
                    fromRaw: raw[map.from], toRaw: raw[map.to],
                    status: status, counts: counts, covers: covers, via: resolvedFrom });

    if (counts && name && covers) res.names[name.toLowerCase()] = 'HOLIDAY';
  }

  res.ok = true;
  return res;
}

function leaveAudit() {
  var wrong = laHome_();
  if (wrong) { Logger.log(wrong); return wrong; }

  var tz = Session.getScriptTimeZone();
  var whenStr = String(LEAVE_AUDIT_DATE || '').trim();
  var when = whenStr ? laDay_(whenStr + 'T12:00:00') : laDay_(new Date());
  if (!when) {
    var bad = 'LEAVE_AUDIT_DATE is not a valid date: "' + whenStr + '". Use yyyy-MM-dd, or leave it blank.';
    Logger.log(bad); return bad;
  }
  var whenLabel = Utilities.formatDate(when, tz, 'yyyy-MM-dd');

  var out = [], say = function (s) { out.push(s); };
  var r = laMirror_(when);

  say('FAIRLEAVE BOOKING AUDIT — read-only');
  say('');
  if (!r.ok) {
    say(r.homeless ? r.why : '✖ ' + r.why);
    Logger.log(out.join('\n'));
    return out.join('\n');
  }

  say('  Bookings tab      : "' + r.tab + '"' +
      (r.byExactName ? '  (found by exact name)' : '  (found by header scan)'));
  say('  Columns resolved  : ' + JSON.stringify(r.map));
  say('  Employees id→name : ' + r.idMapSize);
  say('  Asking about      : ' + whenLabel + (whenStr ? '' : '   (today)'));
  say('');

  // --- every booking, and what the Clock makes of it -----------------------
  say('ALL BOOKINGS');
  say('  ' + laPad_('employee', 22) + laPad_('start', 12) + laPad_('end', 12) +
      laPad_('status', 12) + laPad_('counts?', 9) + 'covers ' + whenLabel);
  if (!r.rows.length) {
    say('  (no booking rows at all)');
  }
  r.rows.forEach(function (b) {
    var who = b.name ? b.name : (b.key ? '?? id ' + b.key : '(no employee)');
    say('  ' + laPad_(who, 22) +
        laPad_(b.from ? Utilities.formatDate(b.from, tz, 'yyyy-MM-dd') : '?' + b.fromRaw, 12) +
        laPad_(b.to   ? Utilities.formatDate(b.to,   tz, 'yyyy-MM-dd') : '?' + b.toRaw, 12) +
        laPad_(b.status, 12) + laPad_(b.counts ? 'yes' : 'no', 9) +
        (b.covers ? 'YES' : '·'));
  });

  var counted = r.rows.filter(function (b) { return b.counts; }).length;
  say('');
  say('RESOLUTION');
  say('  booking rows            : ' + r.rows.length);
  say('  count as approved       : ' + counted);
  say('  employee id → a name    : ' + r.rows.filter(function (b) { return b.counts && b.name; }).length);
  say('  UNRESOLVED ids          : ' + r.unresolved);
  if (r.unresolved) {
    say('    ⚠ An approved booking points at an employee FairLeave does not have.');
    say('      Those people will never show as on holiday. The two systems have drifted.');
  }
  var badDates = r.rows.filter(function (b) { return b.counts && (!b.from || !b.to); }).length;
  if (badDates) {
    say('    ⚠ ' + badDates + ' approved booking(s) have an unreadable start or end date.');
  }

  var who = Object.keys(r.names);
  say('');
  say('ON HOLIDAY on ' + whenLabel + ' : ' + (who.length ? who.join(', ') : 'nobody'));

  // --- self-check ----------------------------------------------------------
  // The mirror is only worth reading if it agrees with the live function on the
  // one date the live function can answer.
  say('');
  say('SELF-CHECK  (is this mirror faithful to the real absencesToday_?)');
  var live = {}, liveErr = '';
  try { live = absencesToday_() || {}; } catch (e) { liveErr = e.message; }
  var liveHols = Object.keys(live).filter(function (k) { return live[k] === 'HOLIDAY'; }).sort();
  var mirrorToday = laMirror_(laDay_(new Date()));
  var mirrorHols = Object.keys(mirrorToday.names || {}).sort();

  say('  live absencesToday_() holidays : ' + (liveErr ? 'ERROR ' + liveErr : (liveHols.join(', ') || 'nobody')));
  say('  mirror, same day               : ' + (mirrorHols.join(', ') || 'nobody'));
  var agree = !liveErr && liveHols.join('|') === mirrorHols.join('|');
  say('');
  if (agree) {
    say('  ✔ They agree, so the date above can be trusted.');
    if (!liveHols.length && !who.length) {
      say('');
      say('  ⚠ BUT both are empty. Agreement on "nobody" proves the code runs, not');
      say('    that it can ever find anyone. Set LEAVE_AUDIT_DATE to a date when');
      say('    someone WAS on approved leave and run this again. If no such date');
      say('    exists yet, approve one booking in FairLeave and re-run.');
    } else if (who.length) {
      say('');
      say('  ✔ END TO END PROVEN: an approved FairLeave booking resolved through the');
      say('    Employees tab to a name the Clock can match. This is the path that');
      say('    silently returned nothing before v1.3.1.');
    }
  } else {
    say('  ✖ They DISAGREE. Trust neither figure.');
    say('    Either absencesToday_ is not the v1.3.1 version, or the Absences tab is');
    say('    contributing rows the mirror does not read. Run codeAudit().');
  }

  say('');
  say('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}