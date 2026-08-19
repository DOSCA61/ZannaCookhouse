/* ===========================================================================
   Zanna Clock — RENAME or MERGE a staff member, history and all.  (v1.1)

   PASTE AS A NEW FILE.  Name it exactly:  Rename_Staff.gs
   (If you already pasted v1.0, replace its contents with this.)
   Do NOT paste into Code.gs. Nothing here shares a name with Code.gs v1.5.1.

   ---------------------------------------------------------------------------
   TWO OPERATIONS. PICK THE RIGHT ONE.

   RENAME — one person, one Staff row, the name on it is wrong.
       previewRename()  →  applyRename()

   MERGE  — one person, TWO Staff rows. This happens when a sync could not
            match a renamed person by name, so it deactivated the old row and
            created a new one. The old row keeps the PIN and all the history;
            the new row has the correct name and nothing else. Left alone,
            Payroll reports them as two different people and the person cannot
            clock in, because the row the kiosk finds has no PIN.
       previewMerge()   →  applyMerge()

   previewRename() refuses if the target name already exists and points you here.

   ---------------------------------------------------------------------------
   WHY EITHER IS NEEDED AT ALL

   The Clock stores a person's NAME in every historical row — Events.Name, the
   SessionID prefix, Overruns.Name, Absences.Name. Change the Staff tab alone
   and all of that history orphans: still there, still in the totals, but
   attributed to a person who no longer exists.

   Both operations rewrite the history to match, then rebuild Sessions and
   Payroll so the derived tabs agree.

   The PIN is never lost. On a rename the row keeps its own PinHash and Salt;
   on a merge they are carried across to the surviving row if it has none.

   The Audit tab is deliberately NOT rewritten. An audit log records what was
   recorded at the time; retroactively editing it would defeat its purpose. The
   rename or merge is written to the Audit tab instead, so the older entries
   stay explicable.

   A merge never deletes a row. The losing row is deactivated and its name is
   stamped with a marker, so it is unmistakably dead but nothing is destroyed.
   Delete it by hand later if you want the sheet tidy.

   ---------------------------------------------------------------------------
   FOR DONAL (2026-08-15) — the Clock has BOTH "Donal" and "Donal S":

       var MERGE_FROM = 'Donal';     // the dead row: PIN + 32 history rows
       var MERGE_INTO = 'Donal S';   // the survivor: matches the Scheduler

       1. previewMerge()   — writes nothing
       2. applyMerge()
       3. previewV14()     — expect DEACTIVATED : 0
       4. upgradeToV14()
   =========================================================================== */

var RENAME_FROM = 'Donal';
var RENAME_TO   = 'Donal S';

var MERGE_FROM  = 'Donal';      // row to retire — its history and PIN move
var MERGE_INTO  = 'Donal S';    // row that survives — should match the Scheduler

/* ------------------------------------------------------------------------- */

function rnEq_(a, b) {
  return String(a == null ? '' : a).trim().toLowerCase() ===
         String(b == null ? '' : b).trim().toLowerCase();
}

/** Header index by any of several accepted labels. -1 if absent. */
function rnCol_(sh, names) {
  var last = sh.getLastColumn();
  if (!last) return -1;
  var hdr = sh.getRange(1, 1, 1, last).getValues()[0]
              .map(function (h) { return String(h).trim().toLowerCase(); });
  for (var i = 0; i < names.length; i++) {
    var k = hdr.indexOf(names[i]);
    if (k >= 0) return k;
  }
  return -1;
}

/**
 * SessionIDs are built as  Name-yyyy-MM-dd-nnn-SHIFT.  Only a LEADING
 * "oldName-" is replaced. A blind string replace would corrupt any id that
 * happened to contain the name elsewhere, and would rewrite ids belonging to
 * someone whose name merely starts the same way.
 */
function rnSid_(sid, from, to) {
  var s = String(sid == null ? '' : sid);
  if (!s) return s;
  var pre = from + '-';
  if (s.length > pre.length && s.slice(0, pre.length).toLowerCase() === pre.toLowerCase()) {
    return to + s.slice(from.length);
  }
  return s;
}

/** Scans every tab that stores a name and reports the work, without doing it. */
function rnScan_(from, to) {
  var r = { from: from, to: to, staff: [], collide: [], events: 0, sids: 0,
            overruns: 0, absences: 0, audit: 0, errors: [], warnings: [] };

  if (!from || !to) { r.errors.push('RENAME_FROM and RENAME_TO must both be set.'); return r; }
  if (rnEq_(from, to)) { r.errors.push('RENAME_FROM and RENAME_TO are the same name.'); return r; }

  // --- Staff ---------------------------------------------------------------
  var sh = tab_('Staff');
  if (!sh) { r.errors.push('No Staff tab.'); return r; }
  var c = staffCols_(sh);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var nm = rows[i][c.name];
    if (rnEq_(nm, from)) {
      r.staff.push({ row: i + 1, name: String(nm),
                     active: String(rows[i][c.active]).toUpperCase() === 'TRUE',
                     hasPin: !!rows[i][c.pin],
                     empId: c.empId >= 0 ? String(rows[i][c.empId] || '').trim() : '' });
    } else if (rnEq_(nm, to)) {
      r.collide.push({ row: i + 1, name: String(nm),
                       active: String(rows[i][c.active]).toUpperCase() === 'TRUE' });
    }
  }
  if (r.staff.length === 0) r.errors.push('No Staff row named "' + from + '".');
  if (r.staff.length > 1)   r.errors.push(r.staff.length + ' Staff rows named "' + from + '" — ambiguous, fix by hand first.');
  if (r.collide.length)     r.errors.push('A Staff row named "' + to + '" already exists (row ' +
                                          r.collide[0].row + '). This is a MERGE, not a rename — the ' +
                                          'same person has two rows. Set MERGE_FROM/MERGE_INTO and run ' +
                                          'previewMerge().');

  // --- Events --------------------------------------------------------------
  var ev = tab_('Events');
  if (ev) {
    var eName = rnCol_(ev, ['name']), eSid = rnCol_(ev, ['sessionid', 'session id']);
    if (eName < 0) r.errors.push('Events tab has no "Name" header.');
    var evRows = ev.getDataRange().getValues();
    for (var e = 1; e < evRows.length; e++) {
      if (eName >= 0 && rnEq_(evRows[e][eName], from)) r.events++;
      if (eSid >= 0 && rnSid_(evRows[e][eSid], from, to) !== String(evRows[e][eSid] || '')) r.sids++;
    }
    if (eSid < 0) r.warnings.push('Events has no "SessionID" header — ids will not be rewritten.');
  } else {
    r.errors.push('No Events tab.');
  }

  // --- Overruns ------------------------------------------------------------
  var ov = tab_('Overruns');
  if (ov) {
    var oName = rnCol_(ov, ['name']);
    if (oName >= 0) {
      var ovRows = ov.getDataRange().getValues();
      for (var o = 1; o < ovRows.length; o++) if (rnEq_(ovRows[o][oName], from)) r.overruns++;
    }
  }

  // --- Absences ------------------------------------------------------------
  var ab = tab_('Absences');
  if (ab) {
    var aName = rnCol_(ab, ['name']);
    if (aName >= 0) {
      var abRows = ab.getDataRange().getValues();
      for (var a = 1; a < abRows.length; a++) if (rnEq_(abRows[a][aName], from)) r.absences++;
    }
  }

  // --- Audit (counted, never rewritten) -------------------------------------
  var au = tab_('Audit');
  if (au) {
    var uName = rnCol_(au, ['actor']);
    if (uName >= 0) {
      var auRows = au.getDataRange().getValues();
      for (var u = 1; u < auRows.length; u++) if (rnEq_(auRows[u][uName], from)) r.audit++;
    }
  }

  return r;
}

/* ===========================================================================
   DRY RUN — writes nothing.
   =========================================================================== */
function previewRename() {
  var r = rnScan_(RENAME_FROM, RENAME_TO);
  var out = [];
  var say = function (s) { out.push(s); };

  say('DRY RUN — nothing is written.');
  say('');
  say('RENAME   "' + r.from + '"   →   "' + r.to + '"');
  say('');

  if (r.staff.length === 1) {
    var p = r.staff[0];
    say('STAFF TAB');
    say('  row ' + p.row + '   active: ' + (p.active ? 'yes' : 'no') +
        '   PIN: ' + (p.hasPin ? 'set (kept)' : 'NOT SET') +
        '   EmployeeID: ' + (p.empId || '—'));
  }

  say('');
  say('HISTORY THAT WOULD BE REWRITTEN');
  say('  Events rows (Name)      : ' + r.events);
  say('  Events rows (SessionID) : ' + r.sids);
  say('  Overruns rows           : ' + r.overruns);
  say('  Absences rows           : ' + r.absences);
  say('');
  say('NOT rewritten');
  say('  Audit rows              : ' + r.audit + '   (a log of what happened — left as recorded)');
  say('  Sessions / Payroll      : derived, rebuilt from Events afterwards');

  r.warnings.forEach(function (w) { say(''); say('  ⚠ ' + w); });

  say('');
  if (r.errors.length) {
    say('✖ STOP. Do not run applyRename().');
    r.errors.forEach(function (e) { say('   • ' + e); });
  } else if (r.events === 0 && r.overruns === 0 && r.absences === 0) {
    say('✔ SAFE. No history exists under "' + r.from + '" — this is a Staff-tab rename only.');
    say('  Run applyRename() when ready.');
  } else {
    say('✔ SAFE. ' + (r.events + r.overruns + r.absences) +
        ' historical row(s) will move with the name, so nothing is orphaned.');
    say('  Run applyRename() when ready.');
  }
  say('');
  say('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ===========================================================================
   APPLY.
   =========================================================================== */
function applyRename() {
  var from = RENAME_FROM, to = RENAME_TO;
  var pre = rnScan_(from, to);
  if (pre.errors.length) {
    var stop = '✖ Refused:\n   • ' + pre.errors.join('\n   • ') +
               '\n\nRun previewRename() for the full picture.';
    Logger.log(stop);
    return stop;
  }

  var out = [], say = function (s) { out.push(s); };
  say('RENAME  "' + from + '"  →  "' + to + '"');
  say('');

  // ---------------------------------------------------------------------------
  // The WRITES take the script lock. The REBUILDS must not — rebuildSessions_
  // takes the same lock itself, and Apps Script script locks are not reentrant,
  // so holding it across the rebuild deadlocks until waitLock times out and
  // throws, leaving the rename applied but Sessions and Payroll stale.
  // ---------------------------------------------------------------------------
  var w = withLock_(function () {
    var r = { nameHits: 0, sidHits: 0, ovHits: 0, abHits: 0, staffRow: 0 };

    // --- Staff -------------------------------------------------------------
    var sh = tab_('Staff');
    var c = staffCols_(sh);
    sh.getRange(pre.staff[0].row, c.name + 1).setValue(to);
    r.staffRow = pre.staff[0].row;

    // --- Events: one read, one write per column -----------------------------
    var ev = tab_('Events');
    var eName = rnCol_(ev, ['name']), eSid = rnCol_(ev, ['sessionid', 'session id']);
    var n = ev.getLastRow() - 1;
    if (n > 0) {
      if (eName >= 0) {
        var col = ev.getRange(2, eName + 1, n, 1).getValues();
        for (var i = 0; i < col.length; i++) {
          if (rnEq_(col[i][0], from)) { col[i][0] = to; r.nameHits++; }
        }
        if (r.nameHits) ev.getRange(2, eName + 1, n, 1).setValues(col);
      }
      if (eSid >= 0) {
        var sids = ev.getRange(2, eSid + 1, n, 1).getValues();
        for (var j = 0; j < sids.length; j++) {
          var next = rnSid_(sids[j][0], from, to);
          if (next !== String(sids[j][0] == null ? '' : sids[j][0])) { sids[j][0] = next; r.sidHits++; }
        }
        if (r.sidHits) ev.getRange(2, eSid + 1, n, 1).setValues(sids);
      }
    }

    r.ovHits = rnRewriteCol_(tab_('Overruns'), ['name'], from, to);
    r.abHits = rnRewriteCol_(tab_('Absences'), ['name'], from, to);
    return r;
  });

  say('  Staff row ' + w.staffRow + ' renamed. PIN and salt untouched.');
  say('  Events: ' + w.nameHits + ' name(s), ' + w.sidHits + ' SessionID(s) rewritten.');
  say('  Overruns: ' + w.ovHits + ' row(s) rewritten.');
  say('  Absences: ' + w.abHits + ' row(s) rewritten.');
  say('');
  say('  Audit: left as recorded (' + pre.audit + ' row(s) still say "' + from + '").');

  // --- Rebuild the derived tabs — OUTSIDE the lock (see note above) ---------
  var rs = rebuildSessions_();
  var rp = rebuildPayroll_();
  say('');
  say('  Rebuilt  Sessions: ' + (rs && rs.sessions) + ' rows   Payroll: ' + (rp && rp.rows) + ' rows');

  audit_('management', 'renameStaff',
         '"' + from + '" → "' + to + '"  (events ' + w.nameHits + ', sids ' + w.sidHits +
         ', overruns ' + w.ovHits + ', absences ' + w.abHits + ')');

  say('');
  say('✔ Done.');
  say('');
  say('NEXT');
  say('  1. Run previewV14()  — it should now report  DEACTIVATED : 0');
  say('  2. Run upgradeToV14()');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/** Rewrites one name column on a simple tab. Returns the number of hits. */
function rnRewriteCol_(sh, headerNames, from, to) {
  if (!sh) return 0;
  var col = rnCol_(sh, headerNames);
  if (col < 0) return 0;
  var n = sh.getLastRow() - 1;
  if (n <= 0) return 0;
  var vals = sh.getRange(2, col + 1, n, 1).getValues();
  var hits = 0;
  for (var i = 0; i < vals.length; i++) {
    if (rnEq_(vals[i][0], from)) { vals[i][0] = to; hits++; }
  }
  if (hits) sh.getRange(2, col + 1, n, 1).setValues(vals);
  return hits;
}

/* ===========================================================================
   MERGE — two Staff rows, one person.
   =========================================================================== */

/** Reads both rows and everything attached to them. Writes nothing. */
function rnScanMerge_(from, into) {
  var r = { from: from, into: into, loser: null, survivor: null,
            losers: 0, survivors: 0,
            events: 0, sids: 0, overruns: 0, absences: 0, audit: 0,
            errors: [], warnings: [] };

  if (!from || !into) { r.errors.push('MERGE_FROM and MERGE_INTO must both be set.'); return r; }
  if (rnEq_(from, into)) { r.errors.push('MERGE_FROM and MERGE_INTO are the same name.'); return r; }

  var sh = tab_('Staff');
  if (!sh) { r.errors.push('No Staff tab.'); return r; }
  var c = staffCols_(sh);
  var rows = sh.getDataRange().getValues();

  var grab = function (i) {
    return { row: i + 1,
             name: String(rows[i][c.name]),
             dept: String(rows[i][c.dept] || ''),
             active: String(rows[i][c.active]).toUpperCase() === 'TRUE',
             hasPin: !!rows[i][c.pin] && !!rows[i][c.salt],
             empId: c.empId >= 0 ? String(rows[i][c.empId] || '').trim() : '' };
  };
  for (var i = 1; i < rows.length; i++) {
    if (rnEq_(rows[i][c.name], from))      { r.losers++;    if (!r.loser)    r.loser = grab(i); }
    else if (rnEq_(rows[i][c.name], into)) { r.survivors++; if (!r.survivor) r.survivor = grab(i); }
  }

  if (r.losers === 0)    r.errors.push('No Staff row named "' + from + '" — nothing to merge.');
  if (r.survivors === 0) r.errors.push('No Staff row named "' + into + '". If the second row does not ' +
                                       'exist, this is a RENAME — use previewRename().');
  if (r.losers > 1)      r.errors.push(r.losers + ' rows named "' + from + '" — ambiguous, fix by hand first.');
  if (r.survivors > 1)   r.errors.push(r.survivors + ' rows named "' + into + '" — ambiguous, fix by hand first.');
  if (r.errors.length) return r;

  if (!r.survivor.hasPin && !r.loser.hasPin) {
    r.warnings.push('NEITHER row has a PIN. After the merge the person still cannot clock in — ' +
                    'set one with Menu → Set PIN.');
  }
  if (r.survivor.hasPin && r.loser.hasPin) {
    r.warnings.push('BOTH rows have a PIN. The surviving row keeps its own; the old one is discarded. ' +
                    'If the person only knows the older PIN, reset it after the merge.');
  }
  if (!r.survivor.active) {
    r.warnings.push('The surviving row "' + into + '" is INACTIVE. The merge will not activate it — ' +
                    'the staff sync owns that. Check the Scheduler lists this person as active.');
  }

  // History still attached to the losing name
  var ev = tab_('Events');
  if (ev) {
    var eName = rnCol_(ev, ['name']), eSid = rnCol_(ev, ['sessionid', 'session id']);
    if (eName < 0) r.errors.push('Events tab has no "Name" header.');
    var evRows = ev.getDataRange().getValues();
    for (var e = 1; e < evRows.length; e++) {
      if (eName >= 0 && rnEq_(evRows[e][eName], from)) r.events++;
      if (eSid >= 0 && rnSid_(evRows[e][eSid], from, into) !== String(evRows[e][eSid] || '')) r.sids++;
    }
    if (eSid < 0) r.warnings.push('Events has no "SessionID" header — ids will not be rewritten.');
  } else {
    r.errors.push('No Events tab.');
  }

  var ov = tab_('Overruns');
  if (ov) {
    var oc = rnCol_(ov, ['name']);
    if (oc >= 0) {
      var ovRows = ov.getDataRange().getValues();
      for (var o = 1; o < ovRows.length; o++) if (rnEq_(ovRows[o][oc], from)) r.overruns++;
    }
  }
  var ab = tab_('Absences');
  if (ab) {
    var ac = rnCol_(ab, ['name']);
    if (ac >= 0) {
      var abRows = ab.getDataRange().getValues();
      for (var a = 1; a < abRows.length; a++) if (rnEq_(abRows[a][ac], from)) r.absences++;
    }
  }
  var au = tab_('Audit');
  if (au) {
    var uc = rnCol_(au, ['actor']);
    if (uc >= 0) {
      var auRows = au.getDataRange().getValues();
      for (var u = 1; u < auRows.length; u++) if (rnEq_(auRows[u][uc], from)) r.audit++;
    }
  }
  return r;
}

/** The marker left on the retired row, so it can never be mistaken for live. */
function rnMergeMark_(name, into) {
  var d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return name + ' [merged ' + d + ' → ' + into + ']';
}

/* ===========================================================================
   DRY RUN — writes nothing.
   =========================================================================== */
function previewMerge() {
  var r = rnScanMerge_(MERGE_FROM, MERGE_INTO);
  var out = [], say = function (s) { out.push(s); };

  say('DRY RUN — nothing is written.');
  say('');
  say('MERGE   "' + r.from + '"   into   "' + r.into + '"');
  say('');

  if (r.loser && r.survivor) {
    var fmt = function (label, p) {
      say('  ' + label);
      say('    row ' + p.row + '   active: ' + (p.active ? 'yes' : 'no') +
          '   PIN: ' + (p.hasPin ? 'set' : 'not set') +
          '   EmployeeID: ' + (p.empId || '—') +
          '   dept: ' + (p.dept || '—'));
    };
    say('STAFF TAB');
    fmt('RETIRED   "' + r.loser.name + '"', r.loser);
    fmt('SURVIVES  "' + r.survivor.name + '"', r.survivor);
  }

  say('');
  say('WHAT WOULD MOVE');
  say('  Events rows (Name)      : ' + r.events);
  say('  Events rows (SessionID) : ' + r.sids);
  say('  Overruns rows           : ' + r.overruns);
  say('  Absences rows           : ' + r.absences);
  if (r.loser && r.survivor) {
    say('  PIN                     : ' +
        (r.survivor.hasPin ? 'survivor keeps its own'
                           : (r.loser.hasPin ? 'copied across from row ' + r.loser.row
                                             : 'NEITHER row has one')));
    say('  Department              : ' +
        (r.survivor.dept ? 'survivor keeps "' + r.survivor.dept + '"'
                         : (r.loser.dept ? 'copied across: "' + r.loser.dept + '"' : 'both blank')));
  }
  say('');
  say('NOT rewritten');
  say('  Audit rows              : ' + r.audit + '   (a log of what happened — left as recorded)');
  say('  Sessions / Payroll      : derived, rebuilt from Events afterwards');
  if (r.loser) {
    say('');
    say('  Row ' + r.loser.row + ' is NOT deleted. It is set inactive and renamed:');
    say('    "' + rnMergeMark_(r.loser.name, r.into) + '"');
  }

  r.warnings.forEach(function (w) { say(''); say('  ⚠ ' + w); });

  say('');
  if (r.errors.length) {
    say('✖ STOP. Do not run applyMerge().');
    r.errors.forEach(function (e) { say('   • ' + e); });
  } else {
    say('✔ SAFE. ' + (r.events + r.overruns + r.absences) +
        ' historical row(s) move onto "' + r.into + '", so Payroll stops treating');
    say('  them as two different people. Run applyMerge() when ready.');
  }
  say('');
  say('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ===========================================================================
   APPLY.
   =========================================================================== */
function applyMerge() {
  var from = MERGE_FROM, into = MERGE_INTO;
  var pre = rnScanMerge_(from, into);
  if (pre.errors.length) {
    var stop = '✖ Refused:\n   • ' + pre.errors.join('\n   • ') +
               '\n\nRun previewMerge() for the full picture.';
    Logger.log(stop);
    return stop;
  }

  var out = [], say = function (s) { out.push(s); };
  say('MERGE  "' + from + '"  into  "' + into + '"');
  say('');

  // Writes under the lock; rebuilds outside it. rebuildSessions_ takes the same
  // script lock, and Apps Script locks are not reentrant.
  var w = withLock_(function () {
    var res = { nameHits: 0, sidHits: 0, ovHits: 0, abHits: 0, pin: '', dept: '' };
    var sh = tab_('Staff');
    var c = staffCols_(sh);

    // --- carry the credentials across, if the survivor has none -------------
    if (!pre.survivor.hasPin && pre.loser.hasPin) {
      var hash = sh.getRange(pre.loser.row, c.pin + 1).getValue();
      var salt = sh.getRange(pre.loser.row, c.salt + 1).getValue();
      sh.getRange(pre.survivor.row, c.pin + 1).setValue(hash);
      sh.getRange(pre.survivor.row, c.salt + 1).setValue(salt);
      res.pin = 'copied from row ' + pre.loser.row + ' — the existing PIN still works';
    } else if (pre.survivor.hasPin) {
      res.pin = 'survivor kept its own';
    } else {
      res.pin = 'NONE — set one with Menu → Set PIN';
    }

    if (!pre.survivor.dept && pre.loser.dept) {
      sh.getRange(pre.survivor.row, c.dept + 1).setValue(pre.loser.dept);
      res.dept = 'copied: "' + pre.loser.dept + '"';
    } else {
      res.dept = pre.survivor.dept ? 'kept: "' + pre.survivor.dept + '"' : 'blank on both';
    }

    // --- move the history ---------------------------------------------------
    var ev = tab_('Events');
    var eName = rnCol_(ev, ['name']), eSid = rnCol_(ev, ['sessionid', 'session id']);
    var n = ev.getLastRow() - 1;
    if (n > 0) {
      if (eName >= 0) {
        var col = ev.getRange(2, eName + 1, n, 1).getValues();
        for (var i = 0; i < col.length; i++) {
          if (rnEq_(col[i][0], from)) { col[i][0] = into; res.nameHits++; }
        }
        if (res.nameHits) ev.getRange(2, eName + 1, n, 1).setValues(col);
      }
      if (eSid >= 0) {
        var sids = ev.getRange(2, eSid + 1, n, 1).getValues();
        for (var j = 0; j < sids.length; j++) {
          var next = rnSid_(sids[j][0], from, into);
          if (next !== String(sids[j][0] == null ? '' : sids[j][0])) { sids[j][0] = next; res.sidHits++; }
        }
        if (res.sidHits) ev.getRange(2, eSid + 1, n, 1).setValues(sids);
      }
    }
    res.ovHits = rnRewriteCol_(tab_('Overruns'), ['name'], from, into);
    res.abHits = rnRewriteCol_(tab_('Absences'), ['name'], from, into);

    // --- retire the losing row LAST, once nothing points at it --------------
    // Not deleted: deleting a Staff row is irreversible and shifts every row
    // number below it. Inactive plus a marker is unambiguous and undoable.
    sh.getRange(pre.loser.row, c.name + 1).setValue(rnMergeMark_(pre.loser.name, into));
    sh.getRange(pre.loser.row, c.active + 1).setValue('FALSE');
    return res;
  });

  say('  Events: ' + w.nameHits + ' name(s), ' + w.sidHits + ' SessionID(s) moved.');
  say('  Overruns: ' + w.ovHits + ' row(s).   Absences: ' + w.abHits + ' row(s).');
  say('  PIN: ' + w.pin);
  say('  Department: ' + w.dept);
  say('  Row ' + pre.loser.row + ' retired (inactive, marked). Not deleted.');
  say('');
  say('  Audit: left as recorded (' + pre.audit + ' row(s) still say "' + from + '").');

  var rs = rebuildSessions_();
  var rp = rebuildPayroll_();
  say('');
  say('  Rebuilt  Sessions: ' + (rs && rs.sessions) + ' rows   Payroll: ' + (rp && rp.rows) + ' rows');

  audit_('management', 'mergeStaff',
         '"' + from + '" → "' + into + '"  (events ' + w.nameHits + ', sids ' + w.sidHits +
         ', overruns ' + w.ovHits + ', absences ' + w.abHits + ', pin: ' + w.pin + ')');

  say('');
  say('✔ Done.');
  say('');
  say('NEXT');
  say('  1. Run previewV14()  — expect  DEACTIVATED : 0');
  say('  2. Run upgradeToV14()');
  Logger.log(out.join('\n'));
  return out.join('\n');
}