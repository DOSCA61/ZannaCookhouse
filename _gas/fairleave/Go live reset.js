/* ===========================================================================
   FairLeave — GO-LIVE RESET  (v1.0)

   PASTE AS A NEW FILE in the FAIRLEAVE project ("FairLeave — Zanna").
   Name it: Go_Live_Reset_FL.gs

   This is the FairLeave half. The Zanna Clock has its own Go_Live_Reset.gs in
   its own project — they cannot see each other, so both must be run.

   ---------------------------------------------------------------------------
   ***  BEFORE YOU RUN ANYTHING  ***

       File → Make a copy

   ---------------------------------------------------------------------------
   HOW TO USE

     1. Make that copy.
     2. Run  previewGoLiveResetFL()   — writes NOTHING.
     3. Set, exactly:   var GO_LIVE_CONFIRM_FL = 'CLEAR ALL DATA';
     4. Run  applyGoLiveResetFL().

   Then, separately, for the Shift Scheduler's leave rows:

     5. Run  previewSchedulerLeave()  — reports every candidate tab and what is
        in it. It clears nothing and guesses at nothing.
     6. Put the exact tab name it shows into SCHED_LEAVE_TAB below.
     7. Run  clearSchedulerLeave().

   Steps 5–7 are separate on purpose. I have never read the Scheduler's leave
   tab, so this file will not act on a name it inferred. You name it; it clears
   that and nothing else.

   ---------------------------------------------------------------------------
   WHAT IS CLEARED         WHAT IS PROTECTED  (hard-blocked)

     Bookings  all rows     Employees  EVERY ROW. Each carries a personal
     Swaps     all rows                access token — the private link each
     Audit     all rows                person uses to book leave. Clearing
                                       this invalidates all fifteen links.
                            Teams      names and the max-off-at-once limits
                            Blackouts  policy, not data (Christmas rush etc.)
                            Config     every setting
   =========================================================================== */

var GO_LIVE_CONFIRM_FL = '';          // set to 'CLEAR ALL DATA' to arm
var SCHED_LEAVE_TAB    = '';          // set from previewSchedulerLeave() output

var GLF_WIPE      = ['Bookings', 'Swaps', 'Audit'];
var GLF_PROTECTED = ['Employees', 'Teams', 'Blackouts', 'Config'];

function glfPad_(v, n) { var s = String(v); while (s.length < n) s += ' '; return s; }
function glfCount_(name) {
  var sh = sheet_(name);
  return sh ? Math.max(0, sh.getLastRow() - 1) : -1;
}

/* ===========================================================================
   DRY RUN
   =========================================================================== */
function previewGoLiveResetFL() {
  var out = [], say = function (s) { out.push(s); };

  say('FAIRLEAVE GO-LIVE RESET — DRY RUN. Nothing is written.');
  say('');
  say('  MAKE A COPY FIRST:  File → Make a copy.  There is no undo.');
  say('');

  say('WILL BE EMPTIED');
  var total = 0;
  GLF_WIPE.forEach(function (name) {
    var n = glfCount_(name);
    if (n < 0) { say('  ' + glfPad_(name, 12) + 'tab not present'); return; }
    total += n;
    say('  ' + glfPad_(name, 12) + glfPad_(n, 6) + 'row' + (n === 1 ? '' : 's'));
  });
  say('  ' + glfPad_('TOTAL', 12) + glfPad_(total, 6) + 'rows destroyed');

  // Spell out what is actually in Bookings — this is the one place where a
  // real, non-test record might be sitting.
  var bookings = readAll_('Bookings') || [];
  if (bookings.length) {
    say('');
    say('THE BOOKINGS, ONE BY ONE');
    var emps = {};
    (readAll_('Employees') || []).forEach(function (e) { emps[e.id] = e.name; });
    bookings.forEach(function (b) {
      say('  ' + glfPad_(emps[b.employeeId] || ('id ' + b.employeeId), 16) +
          glfPad_(b.start + ' → ' + b.end, 26) + b.status);
    });
    var approved = bookings.filter(function (b) { return /approv/i.test(String(b.status)); });
    if (approved.length) {
      say('');
      say('  ⚠ ' + approved.length + ' APPROVED booking(s) above will be destroyed.');
      say('    If any is real leave someone has actually been granted, write it down');
      say('    now — it will have to be re-entered after go-live.');
    }
  }

  say('');
  say('PROTECTED — this file cannot touch these');
  GLF_PROTECTED.forEach(function (name) {
    var n = glfCount_(name);
    say('  ' + glfPad_(name, 12) + (n < 0 ? 'absent' : n + ' rows') +
        (name === 'Employees' ? '   ← personal leave links live in these rows' : ''));
  });

  say('');
  say('THE SHIFT SCHEDULER IS A DIFFERENT SPREADSHEET');
  say('  Its leave rows are not touched by applyGoLiveResetFL().');
  say('  Run previewSchedulerLeave() for that — separate, deliberate step.');

  say('');
  if (GO_LIVE_CONFIRM_FL === 'CLEAR ALL DATA') {
    say('✔ ARMED. applyGoLiveResetFL() will run. Have you made the copy?');
  } else {
    say('🔒 NOT ARMED. To arm, edit the top of this file:');
    say("      var GO_LIVE_CONFIRM_FL = 'CLEAR ALL DATA';");
  }
  say('');
  say('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ===========================================================================
   APPLY
   =========================================================================== */
function applyGoLiveResetFL() {
  if (GO_LIVE_CONFIRM_FL !== 'CLEAR ALL DATA') {
    var no = '🔒 REFUSED — not armed.\n\nRun previewGoLiveResetFL(), make a copy, then set\n' +
             "    var GO_LIVE_CONFIRM_FL = 'CLEAR ALL DATA';";
    Logger.log(no); return no;
  }

  var out = [], say = function (s) { out.push(s); };
  say('FAIRLEAVE GO-LIVE RESET');
  say('');

  GLF_WIPE.forEach(function (name) {
    var sh = sheet_(name);
    if (!sh) { say('  ' + glfPad_(name, 12) + 'tab absent'); return; }
    var n = Math.max(0, sh.getLastRow() - 1);
    var hdr = TABS[name] || sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    sh.clearContents();
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    sh.getRange(1, 1, 1, hdr.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    say('  ' + glfPad_(name, 12) + n + ' rows cleared');
  });

  GLF_PROTECTED.forEach(function (n) { say('  ' + glfPad_(n, 12) + 'untouched'); });

  audit_('Management', 'goLiveReset', 'FairLeave cleared for go-live');

  say('');
  say('✔ Done. Everyone starts with a full entitlement and no history.');
  say('');
  say('CHECK BEFORE MONDAY');
  say('  • Entitlement on the Employees tab is per-person and was NOT reset —');
  say('    it is a setting, not a transaction. Confirm the figures are right.');
  say('  • Personal leave links still work; nobody needs a new one.');
  say('  • The Shift Scheduler is separate: previewSchedulerLeave().');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ===========================================================================
   THE SHIFT SCHEDULER — discovery first, then a named tab.
   =========================================================================== */
function previewSchedulerLeave() {
  var out = [], say = function (s) { out.push(s); };
  var id = getConfig_('SCHEDULER_SHEET_ID');

  say('SHIFT SCHEDULER — LEAVE ROWS. Nothing is written.');
  say('');
  if (!id) {
    say('✖ SCHEDULER_SHEET_ID is not set in FairLeave Config. Cannot look.');
    Logger.log(out.join('\n')); return out.join('\n');
  }

  var ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { say('✖ Cannot open the Scheduler: ' + e.message); Logger.log(out.join('\n')); return out.join('\n'); }

  say('  Spreadsheet: "' + ss.getName() + '"');
  say('');
  say('EVERY TAB IN IT');
  var candidates = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    var rows = Math.max(0, sh.getLastRow() - 1);
    var hdr = sh.getLastColumn()
      ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    var looksLikeLeave = /leave|holiday|absence|timeoff|time off/i.test(name);
    if (looksLikeLeave) candidates.push(name);
    say('  ' + (looksLikeLeave ? '→ ' : '  ') + glfPad_(name, 26) +
        glfPad_(rows + ' rows', 12) + '[' + hdr.join(', ') + ']');
  });

  say('');
  if (!candidates.length) {
    say('⚠ No tab name looks like leave. Read the list above and pick the right one');
    say('  yourself — I will not guess.');
  } else {
    say('LIKELY LEAVE TAB(S): ' + candidates.join(', '));
    candidates.forEach(function (name) {
      var sh = ss.getSheetByName(name);
      var data = sh.getDataRange().getValues();
      if (data.length < 2) return;
      var hdr = data[0].map(function (h) { return String(h).trim().toLowerCase(); });
      var sCol = -1;
      for (var i = 0; i < hdr.length; i++) if (hdr[i].indexOf('status') >= 0) sCol = i;
      say('');
      say('  "' + name + '" — ' + (data.length - 1) + ' data row(s)');
      if (sCol >= 0) {
        var tally = {};
        for (var r = 1; r < data.length; r++) {
          var st = String(data[r][sCol] || '(blank)').trim() || '(blank)';
          tally[st] = (tally[st] || 0) + 1;
        }
        Object.keys(tally).sort().forEach(function (k) {
          say('      ' + glfPad_(k, 14) + tally[k] +
              (/pending/i.test(k) ? '   ← someone is still waiting on these' : ''));
        });
      } else {
        say('      (no status column found)');
      }
      // Rows FairLeave itself pushed are tagged; they will come back on the
      // next sync, so they are not really lost.
      var noteCol = -1;
      for (var j = 0; j < hdr.length; j++) if (hdr[j].indexOf('note') >= 0) noteCol = j;
      if (noteCol >= 0) {
        var pushed = 0;
        for (var q = 1; q < data.length; q++) {
          if (String(data[q][noteCol] || '').indexOf('FairLeave:') === 0) pushed++;
        }
        say('      ' + glfPad_('of which', 14) + pushed + ' were pushed by FairLeave');
      }
    });
  }

  say('');
  say('TO CLEAR ONE');
  say('  1. Put its EXACT name in this file:');
  say("         var SCHED_LEAVE_TAB = '<name from the list above>';");
  say('  2. Set GO_LIVE_CONFIRM_FL if it is not already set.');
  say('  3. Run clearSchedulerLeave().');
  say('');
  say('  Only that one tab is touched. Employees, Schedule, DepartmentHours and');
  say('  everything else in the Scheduler are left completely alone.');
  say('');
  say('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

function clearSchedulerLeave() {
  if (GO_LIVE_CONFIRM_FL !== 'CLEAR ALL DATA') {
    var no = "🔒 REFUSED — set GO_LIVE_CONFIRM_FL = 'CLEAR ALL DATA' first.";
    Logger.log(no); return no;
  }
  if (!SCHED_LEAVE_TAB) {
    var noTab = '🔒 REFUSED — SCHED_LEAVE_TAB is empty.\n\n' +
                'Run previewSchedulerLeave() and put the exact tab name in this file.\n' +
                'This will not guess which tab holds leave.';
    Logger.log(noTab); return noTab;
  }

  var id = getConfig_('SCHEDULER_SHEET_ID');
  if (!id) { Logger.log('✖ SCHEDULER_SHEET_ID not set.'); return '✖ SCHEDULER_SHEET_ID not set.'; }

  var ss, sh;
  try { ss = SpreadsheetApp.openById(id); } catch (e) { return '✖ Cannot open the Scheduler: ' + e.message; }
  sh = ss.getSheetByName(SCHED_LEAVE_TAB);
  if (!sh) {
    var miss = '✖ No tab called "' + SCHED_LEAVE_TAB + '" in "' + ss.getName() + '".\n' +
               'Names must match exactly, including capitals and spaces. ' +
               'Run previewSchedulerLeave() to see the real list.';
    Logger.log(miss); return miss;
  }

  var n = Math.max(0, sh.getLastRow() - 1);
  if (n > 0) {
    // deleteRows, not clearContents — leaves no blank rows behind for the
    // Scheduler's own code to trip over.
    sh.deleteRows(2, n);
  }
  var msg = '✔ Cleared ' + n + ' row(s) from "' + SCHED_LEAVE_TAB + '" in "' + ss.getName() + '".\n' +
            '  Header kept. No other tab touched.';
  audit_('Management', 'goLiveReset.scheduler', SCHED_LEAVE_TAB + ' — ' + n + ' rows');
  Logger.log(msg);
  return msg;
}