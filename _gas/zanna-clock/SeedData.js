/* ===========================================================================
   Zanna Clock — seed a full week of realistic clocking data.

   Week: Mon 3 Aug – Sat 8 Aug 2026 (ISO week 32), all 15 staff.
   Paste at the bottom of Code.gs, run seedLastWeek() from the dropdown.
   Remove it all again with removeWeekTestData().

   Every row is tagged Device = "test-week", so cleanup removes exactly what
   this added and nothing else.
   =========================================================================== */

var WEEKTEST_TAG = 'test-week';

/** Two shift patterns, as the DepartmentHours tab implies (07:00–16:00 core). */
var WEEK_STAFF = [
  // name,        pattern  (A = 07:00-16:00, B = 07:30-15:30), works Saturday
  ['Stuart',      'A', true ],
  ['Anna',        'A', true ],
  ['Caroline',    'A', true ],
  ['Sinquela',    'A', true ],
  ['Hanna',       'A', true ],
  ['Jenny',       'A', true ],
  ['Cleonice',    'A', true ],
  ['Marilyn',     'A', true ],
  ['Ciarna',      'A', true ],
  ['Darragh',     'A', true ],
  ['Lilly',       'B', false],
  ['Sajeda',      'B', false],
  ['Calmi',       'B', false],
  ['Justyna',     'B', false],
  ['Donal',       'B', false]
];

var WEEK_PATTERNS = {
  //        in     b15    b15end b30    b30end out     gross
  A: { inH: 7, inM: 0,  b15: 180, b30: 360, outMin: 540 },   // 07:00 → 16:00
  B: { inH: 7, inM: 30, b15: 180, b30: 360, outMin: 480 }    // 07:30 → 15:30
};

/**
 * Deliberate anomalies, so the data exercises the edge cases rather than only
 * the happy path. Keyed by name + day index (0 = Mon).
 *   noMeal   — took the short break only
 *   openShift— never clocked out (forgotten punch)
 *   openBreak— started a break and never ended it
 *   longMeal — meal break ran 35 minutes instead of 30
 */
var WEEK_ANOMALIES = {
  'Cleonice|2': 'noMeal',
  'Darragh|3':  'openShift',
  'Justyna|4':  'openBreak',
  'Marilyn|1':  'longMeal'
};

function seedLastWeek() {
  var tz = Session.getScriptTimeZone();
  var out = [], rows = [];
  var say = function (s) { out.push(s); };

  removeWeekEvents_();

  var expected = {};   // name → { gross, deducted, days }
  var excluded = excludedBreaks_();

  for (var d = 0; d < 6; d++) {                       // Mon..Sat
    var day = new Date(2026, 7, 3 + d);               // 3 Aug 2026 = Monday
    var dateStr = Utilities.formatDate(day, tz, 'yyyy-MM-dd');

    for (var i = 0; i < WEEK_STAFF.length; i++) {
      var name = WEEK_STAFF[i][0];
      var pat  = WEEK_PATTERNS[WEEK_STAFF[i][1]];
      var sat  = WEEK_STAFF[i][2];
      if (d === 5 && !sat) continue;                  // Saturday: core team only

      var anomaly = WEEK_ANOMALIES[name + '|' + d] || '';
      var sid = name + '-' + dateStr + '-001-SHIFT';
      var start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), pat.inH, pat.inM, 0);
      var at = function (m) { return new Date(start.getTime() + m * 60000); };

      var e = expected[name] = expected[name] || { gross: 0, deducted: 0, days: 0 };

      rows.push([at(0), name, 'CLOCK_IN', '', sid, '', WEEKTEST_TAG]);

      // Short break
      rows.push([at(pat.b15), name, 'BREAK_START', 'BREAK_15', sid, 'ON', WEEKTEST_TAG]);
      if (anomaly === 'openBreak') {
        // started, never ended — no BREAK_END row at all
      } else {
        rows.push([at(pat.b15 + 15), name, 'BREAK_END', 'BREAK_15', sid, 'ON', WEEKTEST_TAG]);
        if (excluded['BREAK_15']) e.deducted += 15;
      }

      // Meal break
      if (anomaly !== 'noMeal' && anomaly !== 'openBreak') {
        var mealLen = (anomaly === 'longMeal') ? 35 : 30;
        rows.push([at(pat.b30), name, 'BREAK_START', 'BREAK_30', sid, 'ON', WEEKTEST_TAG]);
        rows.push([at(pat.b30 + mealLen), name, 'BREAK_END', 'BREAK_30', sid, 'ON', WEEKTEST_TAG]);
        if (excluded['BREAK_30']) e.deducted += mealLen;
      }

      // Clock out
      if (anomaly === 'openShift') {
        // forgotten punch — no CLOCK_OUT, so this day earns nothing
      } else {
        rows.push([at(pat.outMin), name, 'CLOCK_OUT', '', sid, '', WEEKTEST_TAG]);
        e.gross += pat.outMin;
        e.days++;
      }
    }
  }

  // One write, not 500 appendRow calls — appendRow in a loop is what makes
  // seeding scripts hit the six-minute execution limit.
  var sh = tab_('Events');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);

  say('SEEDED  Mon 3 Aug – Sat 8 Aug 2026 (ISO week 32)');
  say('  ' + rows.length + ' event rows, ' + WEEK_STAFF.length + ' staff');
  say('  Pattern A 07:00–16:00 (540 min) ×10, incl. Saturday');
  say('  Pattern B 07:30–15:30 (480 min) ×5, Mon–Fri only');
  say('');
  say('DELIBERATE ANOMALIES');
  say('  Cleonice Wed  — short break only, no meal');
  say('  Marilyn  Tue  — meal ran 35 min instead of 30');
  say('  Darragh  Thu  — never clocked out (open shift, earns nothing)');
  say('  Justyna  Fri  — break started, never ended (unclosed)');

  var rs = rebuildSessions_();
  var rp = rebuildPayroll_();
  say('');
  say('REBUILT   Sessions: ' + rs.sessions + ' rows (' + rs.open + ' open, ' +
      rs.unclosed + ' unclosed breaks)   Payroll: ' + rp.rows + ' rows');

  // Compare expected against what Payroll actually produced.
  var pay = tab_('Payroll').getDataRange().getValues();
  var got = {};
  for (var r = 1; r < pay.length; r++) {
    if (String(pay[r][1]) === 'W32') got[String(pay[r][0])] = pay[r];
  }
  say('');
  say('EXPECTED vs PAYROLL (week 32)');
  say('  name        days  gross   deducted   net     status');
  var bad = 0;
  Object.keys(expected).sort().forEach(function (n) {
    var e = expected[n], g = got[n];
    var net = e.gross - e.deducted;
    if (!g) { say('  ' + pad_(n, 12) + 'no payroll row'); bad++; return; }
    var ok = Number(g[5]) === e.gross && Number(g[8]) === e.deducted && Number(g[9]) === net;
    if (!ok) bad++;
    say('  ' + pad_(n, 12) + pad_(e.days, 6) + pad_(e.gross, 8) + pad_(e.deducted, 11) +
        pad_(net, 8) + (ok ? 'ok' : 'MISMATCH got ' + g[5] + '/' + g[8] + '/' + g[9]));
  });
  say('');
  say(bad === 0 ? '✔ Payroll matches expectation for all ' + Object.keys(expected).length + ' staff.'
                : '✖ ' + bad + ' mismatch(es).');
  say('');
  say('Break rules in force: ' + Object.keys(excluded).sort().map(function (k) {
    return k + (excluded[k] ? '=deducted' : '=paid');
  }).join(', '));
  say('Remove it all again with:  removeWeekTestData()');

  Logger.log(out.join('\n'));
  return out.join('\n');
}

function pad_(v, n) {
  var s = String(v);
  while (s.length < n) s += ' ';
  return s;
}

/** Removes only the rows this seeded, then rebuilds. */
function removeWeekTestData() {
  var n = removeWeekEvents_();
  rebuildSessions_();
  rebuildPayroll_();
  var msg = 'Removed ' + n + ' seeded event row(s) and rebuilt.';
  audit_('management', 'weekTestCleanup', msg);
  Logger.log(msg);
  return msg;
}

/** Deletes contiguous blocks bottom-up — far faster than row-by-row. */
function removeWeekEvents_() {
  var sh = tab_('Events');
  var rows = sh.getDataRange().getValues();
  var removed = 0, runEnd = -1;
  for (var i = rows.length - 1; i >= 1; i--) {
    var isTest = String(rows[i][6]) === WEEKTEST_TAG;
    if (isTest && runEnd < 0) runEnd = i;
    if (!isTest && runEnd >= 0) {
      sh.deleteRows(i + 2, runEnd - i);
      removed += runEnd - i;
      runEnd = -1;
    }
  }
  if (runEnd >= 0) { sh.deleteRows(2, runEnd); removed += runEnd; }
  return removed;
}