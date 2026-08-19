/* ===========================================================================
   Zanna Clock — PAYROLL APP BACKEND  (v1.0, needs Code.gs v1.9.0+)

   PASTE AS A NEW FILE in the ZANNA CLOCK project.  Name it:  Payroll_App.gs
   Companion to payroll-v1.0.html.

   ---------------------------------------------------------------------------
   WHY THIS EXISTS

   Running payroll was the last thing forcing anyone to open the spreadsheet.
   A spreadsheet open in front of someone doing payroll is exactly where a
   stray keystroke lands in a cell nobody notices.

   ---------------------------------------------------------------------------
   WHERE THE NUMBERS COME FROM

   Everything here reads the PAYROLL TAB. It never recalculates.

   That matters, because the Clock has more than one way to add up a week:

     rebuildPayroll_()  Events → Sessions → Payroll. Segment-aware, so a split
                        shift counts as two shifts, and break deduction follows
                        the "Exclude Time" column per break type. THIS IS THE
                        ONE PAYROLL USES.

     payrollCsv_()      A separate, older calculation that walks Events directly
                        and takes the FIRST clock-in and LAST clock-out of each
                        day. A person who clocked out at 11:00 and back in at
                        14:00 is paid for the gap. It also predates the
                        per-break-type rules. Reachable at
                        doGet ?action=payroll — DO NOT use it for real payroll.

   Two calculations that can disagree is how the original break-deduction bug
   stayed hidden for weeks. This app deliberately has no calculation of its own
   to become the third.

   ---------------------------------------------------------------------------
   OPS  (all require the management PIN — clkMgmtAuth_ in Code.gs)

     payroll.weeks    list the weeks present, newest first
     payroll.rows     rows for one week, or for a date range
     payroll.process  rebuild Sessions then Payroll from the raw Events
   =========================================================================== */

var PAY_APP_VERSION = '1.0';

/** Called from doPost's management switch (Code.gs v1.9.0 default: branch). */
function clockPayrollHandle_(action, req) {
  switch (action) {
    case 'payroll.weeks':   return payWeeks_();
    case 'payroll.rows':    return payRows_(req);
    case 'payroll.process': return payProcess_();
  }
  return null;                       // not ours — let doPost fall through
}

/** Header-resolved, so a reordered or renamed column fails loudly. */
function payCols_() {
  var sh = tab_('Payroll');
  if (!sh || sh.getLastRow() < 1) return null;
  var hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0]
              .map(function (h) { return String(h).trim().toLowerCase(); });
  var at = function (n) { return hdr.indexOf(n); };
  return {
    hdr: hdr,
    employee: at('employee'), week: at('week'), weekStart: at('week start date'),
    shiftStart: at('shift start'), shiftEnd: at('shift end'),
    gross: at('gross minutes'), deductedList: at('breaks deducted'),
    paidList: at('breaks paid'), deducted: at('total deducted minutes'),
    netMins: at('net paid minutes'), netHours: at('net paid hours'),
    calc: at('calculation')
  };
}

/** dd/MM/yyyy — how rebuildPayroll_ writes Week Start Date — to yyyy-MM-dd. */
function payIso_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  var d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function payReadAll_() {
  var sh = tab_('Payroll');

  // NOT-YET-BUILT and BROKEN are different states and must not share an error.
  // A missing or empty Payroll tab is the normal state of a fresh system, and
  // before this was separated it locked the app at the sign-in screen with
  // "headers not recognised" — which reads like corruption, not "press
  // Process". Only a tab that HAS headers and has the wrong ones is an error.
  if (!sh || sh.getLastRow() < 1) return { ok: true, rows: [], empty: true };

  var c = payCols_();
  var blank = !c || c.hdr.join('').trim() === '';
  if (blank) return { ok: true, rows: [], empty: true };

  if (c.employee < 0 || c.week < 0) {
    return { ok: false, error: 'Payroll tab headers not recognised — found: ' + c.hdr.join(', ') +
             '. Expected at least "Employee" and "Week". If someone has renamed the columns, ' +
             'press Process payroll to rebuild the tab from scratch.' };
  }
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!String(r[c.employee] || '').trim()) continue;
    var calc = String(r[c.calc] || '');
    out.push({
      employee:   String(r[c.employee]),
      week:       String(r[c.week]),
      weekStart:  payIso_(r[c.weekStart]),
      shiftStart: String(r[c.shiftStart] || ''),
      shiftEnd:   String(r[c.shiftEnd] || ''),
      gross:      Number(r[c.gross]) || 0,
      breaksDeducted: String(r[c.deductedList] || ''),
      breaksPaid: String(r[c.paidList] || ''),
      deducted:   Number(r[c.deducted]) || 0,
      netMins:    Number(r[c.netMins]) || 0,
      netHours:   Number(r[c.netHours]) || 0,
      calculation: calc,
      // rebuildPayroll_ appends a warning when a break type is missing from
      // Config and has therefore been treated as PAID. Surface it rather than
      // burying it in a text column nobody reads.
      warning:    calc.indexOf('⚠') >= 0 ? calc.slice(calc.indexOf('⚠')) : ''
    });
  }
  return { ok: true, rows: out };
}

/** Weeks present in the data, newest first, each with its start date. */
function payWeeks_() {
  var all = payReadAll_();
  if (!all.ok) return all;
  var byWeek = {};
  all.rows.forEach(function (r) {
    var k = r.week + '|' + r.weekStart;
    var w = byWeek[k] = byWeek[k] ||
      { week: r.week, weekStart: r.weekStart, staff: 0, netMins: 0, warnings: 0 };
    w.staff++; w.netMins += r.netMins; if (r.warning) w.warnings++;
  });
  var list = Object.keys(byWeek).map(function (k) { return byWeek[k]; });
  if (!list.length) return { ok: true, weeks: [], empty: true, version: PAY_APP_VERSION };
  // Sort by start date, newest first. Sorting on the week NUMBER would put
  // W01 of a new year below W52 of the old one.
  list.sort(function (a, b) { return a.weekStart < b.weekStart ? 1 : a.weekStart > b.weekStart ? -1 : 0; });
  return { ok: true, weeks: list, version: PAY_APP_VERSION };
}

/** Rows for one week, or for a from/to range of week-start dates. */
function payRows_(req) {
  var all = payReadAll_();
  if (!all.ok) return all;

  var week = String((req && req.week) || '').trim();
  var from = String((req && req.from) || '').trim();
  var to   = String((req && req.to)   || '').trim();

  var rows = all.rows;
  var scope = 'all weeks';
  if (week) {
    rows = rows.filter(function (r) { return r.week === week; });
    scope = week + (rows.length ? ' (week beginning ' + rows[0].weekStart + ')' : '');
  } else if (from || to) {
    rows = rows.filter(function (r) {
      if (!r.weekStart) return false;
      if (from && r.weekStart < from) return false;
      if (to   && r.weekStart > to)   return false;
      return true;
    });
    scope = (from || 'the start') + ' → ' + (to || 'now');
  }
  rows.sort(function (a, b) { return a.employee.localeCompare(b.employee); });

  var totals = { staff: rows.length, gross: 0, deducted: 0, netMins: 0, warnings: 0 };
  rows.forEach(function (r) {
    totals.gross += r.gross; totals.deducted += r.deducted;
    totals.netMins += r.netMins; if (r.warning) totals.warnings++;
  });
  totals.netHours = Math.round(totals.netMins / 60 * 100) / 100;

  audit_('management', 'payroll.view', scope + ' — ' + rows.length + ' row(s)');
  return { ok: true, rows: rows, totals: totals, scope: scope, version: PAY_APP_VERSION };
}

/**
 * Recalculate everything from the raw Events log.
 *
 * Safe to run as often as you like: Sessions and Payroll are both DERIVED, so
 * this rebuilds them from scratch rather than adjusting them. Nothing a person
 * clocked is changed — Events is append-only and never written here.
 *
 * Not wrapped in withLock_: rebuildSessions_ takes the script lock itself, and
 * Apps Script locks are not reentrant, so holding one across it would deadlock.
 */
function payProcess_() {
  var rs = rebuildSessions_();
  if (!rs || rs.ok !== true) {
    return { ok: false, error: 'Sessions rebuild failed: ' + ((rs && rs.error) || 'unknown') };
  }
  var rp = rebuildPayroll_();
  if (!rp || rp.ok !== true) {
    return { ok: false, error: 'Payroll rebuild failed: ' + ((rp && rp.error) || 'unknown') };
  }
  audit_('management', 'payroll.process',
         'Sessions ' + rs.sessions + ' rows, Payroll ' + rp.rows + ' rows');
  return {
    ok: true,
    sessions: rs.sessions,
    open: rs.open || 0,
    unclosed: rs.unclosed || 0,
    splitShifts: rs.splitShifts || 0,
    payrollRows: rp.rows,
    unknownBreakTypes: rp.unknownBreakTypes || [],
    version: PAY_APP_VERSION
  };
}

