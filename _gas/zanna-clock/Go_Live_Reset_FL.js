/* ===========================================================================
   Zanna Clock — GO-LIVE RESET  (v1.0)

   PASTE AS A NEW FILE in the ZANNA CLOCK project.  Name it: Go_Live_Reset.gs

   ---------------------------------------------------------------------------
   ***  BEFORE YOU RUN ANYTHING  ***

       File → Make a copy

   Nothing here is undoable. Apps Script has no undo, and the Sheets version
   history will be a wall of identical "cleared" states. A copy costs ten
   seconds and is the only real safety net. The preview refuses to shut up
   about this because it is the one step people skip.

   ---------------------------------------------------------------------------
   HOW TO USE

     1. Make that copy.
     2. Run  previewGoLiveReset()   — writes NOTHING, lists exactly what dies.
     3. Read it. Row counts should look like the testing you actually did.
     4. Set the confirmation constant below, exactly:
            var GO_LIVE_CONFIRM = 'CLEAR ALL DATA';
     5. Run  applyGoLiveReset().

   The typed confirmation exists because every function in a project sits in
   one dropdown, and applyGoLiveReset() is two rows away from previewV14().
   A misclick should not be able to wipe the clocking history.

   ---------------------------------------------------------------------------
   WHAT IS CLEARED            WHAT IS PROTECTED  (hard-blocked, not just skipped)

     Events      all rows      Staff    names, EmployeeIDs, Department, Active
     Sessions    all rows      Devices  every row — those tokens are baked into
     Payroll     all rows               the live kiosk and display URLs
     Overruns    all rows      Config   every setting
     Absences    all rows
     Audit       all rows

     Staff PinHash + Salt are blanked — you chose "everyone re-enrols". The
     ROWS stay: names, ids and Active are untouched, so the roster, the kiosk
     list and every historical reference to a name survive. Each person picks a
     new 4-digit PIN at the kiosk on first use.

   Headers are rewritten and tabs are never deleted, so formatting, column
   widths and frozen rows survive.
   =========================================================================== */

var GO_LIVE_CONFIRM = '';                    // set to 'CLEAR ALL DATA' to arm

/** Data tabs, and the header each is rebuilt with. */
var GLR_WIPE = {
  Events:   ['Timestamp', 'Name', 'EventType', 'BreakType', 'SessionID', 'Premises', 'Device'],
  Sessions: ['SessionID', 'Name', 'Department', 'Type', 'BreakType', 'Start', 'End',
             'DurationMins', 'Premises', 'Device', 'Date', 'Status'],
  Payroll:  ['Employee', 'Week', 'Week Start Date', 'Shift Start', 'Shift End',
             'Gross Minutes', 'Breaks Deducted', 'Breaks Paid',
             'Total Deducted Minutes', 'Net Paid Minutes', 'Net Paid Hours', 'Calculation'],
  Overruns: ['Timestamp', 'Name', 'BreakType', 'AllowedMins', 'ActualMins', 'OverByMins'],
  Absences: ['Name', 'From', 'To', 'Type', 'Notes'],
  Audit:    ['Timestamp', 'Actor', 'Action', 'Details']
};

/** Touched by nothing in this file, whatever else changes. */
var GLR_PROTECTED = ['Staff', 'Devices', 'Config'];

function glrPad_(v, n) { var s = String(v); while (s.length < n) s += ' '; return s; }

function glrCount_(name) {
  var sh = tab_(name);
  if (!sh) return -1;                          // tab absent
  return Math.max(0, sh.getLastRow() - 1);     // minus the header
}

/* ===========================================================================
   DRY RUN
   =========================================================================== */
function previewGoLiveReset() {
  var out = [], say = function (s) { out.push(s); };

  say('GO-LIVE RESET — DRY RUN. Nothing is written.');
  say('');
  say('┌───────────────────────────────────────────────────────────┐');
  say('│  MAKE A COPY OF THIS SPREADSHEET FIRST:  File → Make a copy │');
  say('│  There is no undo. This is the only way back.              │');
  say('└───────────────────────────────────────────────────────────┘');
  say('');

  say('WILL BE EMPTIED');
  var total = 0;
  Object.keys(GLR_WIPE).forEach(function (name) {
    var n = glrCount_(name);
    if (n < 0) { say('  ' + glrPad_(name, 12) + 'tab not present — skipped'); return; }
    total += n;
    say('  ' + glrPad_(name, 12) + glrPad_(n, 8) + 'row' + (n === 1 ? '' : 's'));
  });
  say('  ' + glrPad_('', 12) + glrPad_('—', 8));
  say('  ' + glrPad_('TOTAL', 12) + glrPad_(total, 8) + 'rows destroyed');

  // --- Staff: rows kept, credentials blanked -------------------------------
  var sh = tab_('Staff');
  var withPin = 0, active = 0, staffRows = 0;
  if (sh) {
    var c = staffCols_(sh), rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (!String(rows[i][c.name] || '').trim()) continue;
      staffRows++;
      if (String(rows[i][c.active]).toUpperCase() === 'TRUE') active++;
      if (rows[i][c.pin]) withPin++;
    }
  }
  say('');
  say('STAFF TAB — rows are KEPT, credentials are blanked');
  say('  staff rows              : ' + staffRows + '  (' + active + ' active)');
  say('  PINs that will be wiped : ' + withPin);
  say('  kept untouched          : Name, Department, Active, EmployeeID');
  if (withPin) {
    say('  → ' + withPin + ' people re-enrol at the kiosk: tap name, choose a new 4-digit PIN.');
  }

  // --- Protected -----------------------------------------------------------
  say('');
  say('PROTECTED — this file cannot touch these');
  GLR_PROTECTED.forEach(function (name) {
    var n = glrCount_(name);
    say('  ' + glrPad_(name, 12) + (n < 0 ? 'absent' : n + ' rows') +
        (name === 'Devices' ? '   ← tokens live inside every kiosk/display URL' : '') +
        (name === 'Config'  ? '   ← break rules, sheet ids, settings' : ''));
  });

  // --- Test data, called out separately ------------------------------------
  if (typeof countTestData === 'function') {
    say('');
    say('(Test_Helpers.gs is present. After this reset there will be nothing left');
    say(' for countTestData() to find — the whole Events tab goes.)');
  }

  say('');
  if (GO_LIVE_CONFIRM === 'CLEAR ALL DATA') {
    say('✔ ARMED. GO_LIVE_CONFIRM is set. applyGoLiveReset() will run.');
    say('  Last check: have you made the copy?');
  } else {
    say('🔒 NOT ARMED. applyGoLiveReset() will refuse.');
    say('  To arm it, edit the top of this file:');
    say("      var GO_LIVE_CONFIRM = 'CLEAR ALL DATA';");
  }
  say('');
  say('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ===========================================================================
   APPLY
   =========================================================================== */
function applyGoLiveReset() {
  if (GO_LIVE_CONFIRM !== 'CLEAR ALL DATA') {
    var no = '🔒 REFUSED — not armed.\n\n' +
             'Run previewGoLiveReset() first, make a copy of the spreadsheet, then set\n' +
             "    var GO_LIVE_CONFIRM = 'CLEAR ALL DATA';\n" +
             'at the top of Go_Live_Reset.gs and run this again.';
    Logger.log(no);
    return no;
  }

  var out = [], say = function (s) { out.push(s); };
  var before = {};
  Object.keys(GLR_WIPE).forEach(function (n) { before[n] = glrCount_(n); });

  say('GO-LIVE RESET');
  say('');

  var w = withLock_(function () {
    var res = { wiped: {}, pins: 0, staffRows: 0 };

    // --- data tabs: clear, then put the header back ------------------------
    Object.keys(GLR_WIPE).forEach(function (name) {
      var sh = tab_(name);
      if (!sh) return;
      var hdr = GLR_WIPE[name];
      sh.clearContents();
      sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
      sh.setFrozenRows(1);
      res.wiped[name] = before[name];
    });

    // --- Staff: blank PinHash and Salt, keep everything else ---------------
    var sh = tab_('Staff');
    if (sh) {
      var c = staffCols_(sh);
      var n = sh.getLastRow() - 1;
      if (n > 0) {
        var pins = sh.getRange(2, c.pin + 1, n, 1).getValues();
        var salts = sh.getRange(2, c.salt + 1, n, 1).getValues();
        for (var i = 0; i < n; i++) {
          if (pins[i][0]) res.pins++;
          pins[i][0] = ''; salts[i][0] = '';
        }
        sh.getRange(2, c.pin + 1, n, 1).setValues(pins);
        sh.getRange(2, c.salt + 1, n, 1).setValues(salts);
        res.staffRows = n;
      }
    }
    return res;
  });

  Object.keys(GLR_WIPE).forEach(function (name) {
    say('  ' + glrPad_(name, 12) +
        (w.wiped[name] === undefined ? 'tab absent' : w.wiped[name] + ' rows cleared'));
  });
  say('');
  say('  Staff        ' + w.staffRows + ' rows kept, ' + w.pins + ' PIN(s) blanked');
  GLR_PROTECTED.forEach(function (n) { say('  ' + glrPad_(n, 12) + 'untouched'); });

  // The Audit tab was just emptied — this is deliberately the first line in it,
  // so the new log opens by saying where it came from.
  audit_('management', 'goLiveReset',
         'Cleared for go-live: ' +
         Object.keys(w.wiped).map(function (k) { return k + '=' + w.wiped[k]; }).join(', ') +
         '; ' + w.pins + ' PINs blanked');

  say('');
  say('✔ Done. The Clock is empty and ready.');
  say('');
  say('WHAT HAPPENS ON MONDAY');
  say('  • Every person taps their name and is asked to CHOOSE a 4-digit PIN.');
  say('    That is expected — tell them, or the first one there will think it is broken.');
  say('  • The first tap of the day clocks them straight in.');
  say('  • Nobody shows green on the kiosk until they clock in.');
  say('');
  say('STILL TO DO BY HAND');
  say('  • FairLeave and the Scheduler are separate projects — run their reset there.');
  say('  • Delete Test_Helpers.gs from this project.');
  say('  • Change Mgmt_PIN in Project Settings → Script Properties (8+ characters).');
  Logger.log(out.join('\n'));
  return out.join('\n');
}
