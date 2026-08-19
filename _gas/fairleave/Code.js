/**********************************************************************
 * FairLeave v2.4.0 — Self-managed holiday system (Apps Script backend)
 * Sheet tabs: Config, Teams, Employees, Bookings, Blackouts, Swaps, Audit
 *
 * SETUP:
 *   1. Create a Google Sheet, open Extensions → Apps Script, paste this file.
 *   2. Run setup() once (authorise when prompted). Optionally run seedDemo().
 *   3. Set the Script Property Mgmt_PIN (Project Settings → Script Properties).
 *      NOT the Config tab — see flMgmtPin_.
 *   4. Deploy → New deployment → Web app:
 *        Execute as: Me | Who has access: Anyone
 *   5. Copy the /exec URL into the FairLeave HTML app settings screen.
 *
 * RULES ENGINE (server-side, race-safe via LockService):
 *   - Blackouts: company-wide hard block
 *   - Team concurrency limit per OPERATING day (Mon–Sat by default)
 *   - Auto-approve when capacity allows — first come, first served
 *   - Swaps: true exchange — releaser picks new dates (or release-only)
 *
 * v2.3 (2026-08-15) — OPERATING days and CHARGEABLE days are now separate.
 *   Saturday is a working day: it still counts toward team capacity, so cover
 *   is protected. It is not a holiday day: it no longer comes off anyone's
 *   balance. Previously isNonWorking_() answered both questions, so the two
 *   could never disagree — and setting WORK_SATURDAY=FALSE would have switched
 *   off Saturday capacity checking as well as Saturday charging.
 *     - isNonChargeable_()  NEW — Sat and Sun never cost entitlement
 *     - operatingDaysIn_()  NEW — days the business runs, for validity checks
 *     - workdaysIn_()       now counts CHARGEABLE days only
 *     - validate_()         range validity now gated on operating days, so a
 *                           Saturday can still be booked off (and costs 0)
 *     - stateFor_()         exposes chargeSaturday so the UI can match the server
 *   WORK_SATURDAY stays TRUE and now means OPERATING days only.
 *
 * NOTE: usedDays_() recomputes from dates on every read rather than storing a
 * day count, so any change to the charging rule applies RETROACTIVELY to
 * existing bookings. That is fine here (one approved booking, on a Thursday)
 * but matters once real leave has accrued.
 *
 * v2.3.1 (2026-08-15) — management auth hardened. MGMT_PIN was compared with
 *   a plain === on an "Anyone" endpoint with no lockout and no rate limit,
 *   guarding every management action. Now constant-time, with a failed-attempt
 *   lockout, and a weak PIN is recorded in the Audit tab on every sign-in.
 *
 * COMPANION FILE: Admin.gs provides the Data Maintenance console
 * (action:'admin'). It is additive — the only change it needs in this
 * file is the one hook line in doPost, marked below.
 **********************************************************************/

var TABS = {
  Config:   ['key','value'],
  Teams:    ['id','name','max'],
  Employees:['id','name','teamId','entitlement','email','token','active','extId'],
  Bookings: ['id','employeeId','start','end','status','note','createdAt'],
  Blackouts:['id','start','end','reason'],
  Swaps:    ['id','requesterId','targetBookingId','reqStart','reqEnd','status','note','createdAt','resolvedAt'],
  Audit:    ['timestamp','actor','action','details']
};

/* ---------------- setup ---------------- */
function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function(name){
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0){
      sh.appendRow(TABS[name]);
      sh.getRange(1,1,1,TABS[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    // keep dates as plain text to avoid Sheets date coercion
    sh.getRange(1,1,Math.max(sh.getMaxRows(),1000),TABS[name].length).setNumberFormat('@');
  });
  // Deliberately NOT seeded with a PIN. This line used to be
  //   setConfig_('MGMT_PIN', getConfig_('MGMT_PIN') || '1234')
  // which put 1234 straight back into the sheet every time setup() ran —
  // including after the PIN had been moved to Script Properties.
  if (getConfig_('MGMT_PIN') !== FL_PIN_MOVED &&
      !PropertiesService.getScriptProperties().getProperty(FL_AUTH.PIN_PROP)) {
    setConfig_('MGMT_PIN', FL_PIN_MOVED);
  }
  setConfig_('COMPANY_NAME', getConfig_('COMPANY_NAME') || 'FairLeave');
  setConfig_('WORK_SATURDAY', getConfig_('WORK_SATURDAY') || 'TRUE');
  setConfig_('SCHEDULER_SHEET_ID', getConfig_('SCHEDULER_SHEET_ID') || '1LV_mjBpITQEzMHdDnd8wX0Pn-TEiYCUns3R08IB8oMk');
  var def = ss.getSheetByName('Sheet1'); if (def && ss.getSheets().length > 7) ss.deleteSheet(def);
}

function seedDemo(){
  var t1 = addRow_('Teams', {id:uid_(), name:'Production', max:'2'});
  var t2 = addRow_('Teams', {id:uid_(), name:'Front of House', max:'1'});
  ['Aoife Byrne','Marek Nowak','Siobhán Kelly','Tomasz Zieliński'].forEach(function(n){
    addRow_('Employees', {id:uid_(), name:n, teamId:t1.id, entitlement:'20', email:'', token:token_(), active:'TRUE'});
  });
  ['Ella Roche','Owen Doyle'].forEach(function(n){
    addRow_('Employees', {id:uid_(), name:n, teamId:t2.id, entitlement:'20', email:'', token:token_(), active:'TRUE'});
  });
  var y = todayISO_().slice(0,4);
  addRow_('Blackouts', {id:uid_(), start:y+'-12-14', end:y+'-12-24', reason:'Christmas rush — all hands'});
}

/* ---------------- HTTP entry ---------------- */
function doGet(e){ return json_({ok:true, service:'FairLeave v2.4.0', time:new Date().toISOString()}); }

function doPost(e){
  var _p = {}; try { _p = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch(_) {}
  if (_p.action === 'getPanelConfig') return panelConfig_(_p);
  if (_p.action === 'admin') return adminOut_(adminHandle(_p));   // Data Maintenance console (Admin.gs)
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch(err){ return json_({ok:false, error:'Bad request body'}); }

  var action = req.action || '';
  var mutating = ['requestHoliday','requestSwap','respondSwap','cancelBooking',
                  'addTeam','updateTeam','deleteTeam','addEmployee','updateEmployee','deleteEmployee',
                  'addBlackout','deleteBlackout','regenToken','syncScheduler','importLeave'].indexOf(action) > -1;
  var lock = null;
  try{
    if (mutating){ lock = LockService.getScriptLock(); lock.waitLock(15000); }
    var out = route_(action, req);
    return json_(out);
  } catch(err){
    return json_({ok:false, error:String(err.message || err)});
  } finally {
    if (lock) lock.releaseLock();
  }
}

function route_(action, req){
  // --- auth resolution ---
  var mgmt = mgmtAuth_(req.pin);   // constant-time + lockout (see mgmtAuth_)
  var me = req.token ? findBy_('Employees','token',String(req.token)) : null;
  if (me && String(me.active) === 'FALSE') me = null;

  switch(action){
    case 'login':
      if (mgmt) {
        audit_('Management','mgmt.signIn','FairLeave management sign-in');
        if (mgmtPinWeak_()) audit_('Management','mgmt.weakPin',
          FL_AUTH.PIN_PROP+' is shorter than '+FL_AUTH.MIN_PIN_LENGTH+' characters — lengthen it in Script Properties');
        return {ok:true, role:'mgmt', weakPin: mgmtPinWeak_(), state: stateFor_(null, true)};
      }
      if (me)   return {ok:true, role:'employee', me: pubEmp_(me), state: stateFor_(me, false)};
      return {ok:false, error:'Login failed — check your link or PIN.'};

    case 'state':
      if (mgmt) return {ok:true, state: stateFor_(null, true)};
      if (me)   return {ok:true, me: pubEmp_(me), state: stateFor_(me, false)};
      return {ok:false, error:'Not authorised.'};

    case 'requestHoliday': {
      if (!me) return {ok:false, error:'Not authorised.'};
      var v = validate_(me.id, req.start, req.end, {});
      if (v.ok){
        var nb = {id:uid_(), employeeId:me.id, start:req.start, end:req.end,
                  status:'approved', note:'', createdAt:String(Date.now())};
        addRow_('Bookings', nb);
        audit_(me.name, 'booking.approved', req.start+' → '+req.end);
        schedPush_(nb, 'approve');
        return {ok:true, approved:true, days:v.days, state: stateFor_(me,false)};
      }
      if (v.conflicts) return {ok:true, approved:false, reason:v.reason,
        conflicts:v.conflicts.map(function(b){ return {id:b.id, name:empName_(b.employeeId), start:b.start, end:b.end}; }),
        state: stateFor_(me,false)};
      return {ok:false, error:v.reason};
    }

    case 'requestSwap': {
      if (!me) return {ok:false, error:'Not authorised.'};
      var tb = findBy_('Bookings','id',req.targetBookingId);
      if (!tb || tb.status !== 'approved') return {ok:false, error:'That booking no longer exists.'};
      var dup = readAll_('Swaps').some(function(w){
        return w.status==='pending' && w.requesterId===me.id && w.targetBookingId===tb.id;
      });
      if (dup) return {ok:false, error:'You already have a pending swap request for that booking.'};
      addRow_('Swaps', {id:uid_(), requesterId:me.id, targetBookingId:tb.id,
                        reqStart:req.reqStart, reqEnd:req.reqEnd, status:'pending', note:'',
                        createdAt:String(Date.now()), resolvedAt:''});
      audit_(me.name, 'swap.requested', 'wants '+req.reqStart+' → '+req.reqEnd+' from '+empName_(tb.employeeId));
      notify_(tb.employeeId, 'Holiday swap request',
        me.name+' is asking you to exchange your holidays '+tb.start+' → '+tb.end+
        ' so they can book '+req.reqStart+' → '+req.reqEnd+'.\n\nOpen FairLeave to accept (you pick new dates) or decline.');
      return {ok:true, state: stateFor_(me,false)};
    }

    case 'respondSwap': {
      if (!me) return {ok:false, error:'Not authorised.'};
      return respondSwap_(me, req);
    }

    case 'cancelBooking': {
      var b = findBy_('Bookings','id',req.bookingId);
      if (!b) return {ok:false, error:'Booking not found.'};
      var allowed = mgmt || (me && b.employeeId === me.id);
      if (!allowed) return {ok:false, error:'Not authorised.'};
      updateRow_('Bookings', b.id, {status:'cancelled', note: mgmt && !(me && b.employeeId===me.id) ? 'Cancelled by management' : b.note});
      failPendingSwapsFor_(b.id, 'Booking was cancelled.');
      schedPush_(b, 'cancel');
      audit_(mgmt ? 'Management' : me.name, 'booking.cancelled', empName_(b.employeeId)+' '+b.start+' → '+b.end);
      return {ok:true, state: stateFor_(me, mgmt)};
    }

    case 'syncScheduler': {
      if (!mgmt) return {ok:false, error:'Management PIN required.'};
      var sr = syncFromScheduler_();
      if (!sr.ok) return sr;
      return {ok:true, summary:sr.summary, state: stateFor_(null, true)};
    }

    case 'importLeave': {
      if (!mgmt) return {ok:false, error:'Management PIN required.'};
      var ir = importLeaveFromScheduler_();
      if (!ir.ok) return ir;
      return {ok:true, summary:ir.summary, state: stateFor_(null, true)};
    }

    /* ----- management-only ----- */
    case 'addTeam':      return mgmtDo_(mgmt, function(){ addRow_('Teams',{id:uid_(),name:req.name,max:String(Math.max(1,req.max|0))}); });
    case 'updateTeam':   return mgmtDo_(mgmt, function(){ updateRow_('Teams',req.id,{name:req.name,max:String(Math.max(1,req.max|0))}); });
    case 'deleteTeam':   return mgmtDo_(mgmt, function(){
                             if (readAll_('Employees').some(function(e){return e.teamId===req.id && e.active!=='FALSE';}))
                               throw new Error('Move its members to another team first.');
                             deleteRow_('Teams',req.id); });
    case 'addEmployee':  return mgmtDo_(mgmt, function(){
                             addRow_('Employees',{id:uid_(),name:req.name,teamId:req.teamId,
                               entitlement:String(req.entitlement|0),email:req.email||'',token:token_(),active:'TRUE'}); });
    case 'updateEmployee':return mgmtDo_(mgmt, function(){
                             updateRow_('Employees',req.id,{name:req.name,teamId:req.teamId,
                               entitlement:String(req.entitlement|0),email:req.email||''}); });
    case 'deleteEmployee':return mgmtDo_(mgmt, function(){
                             updateRow_('Employees',req.id,{active:'FALSE',token:''});
                             readAll_('Bookings').forEach(function(b){
                               if (b.employeeId===req.id && b.status==='approved'){
                                 updateRow_('Bookings',b.id,{status:'cancelled',note:'Employee removed'});
                                 failPendingSwapsFor_(b.id,'Booking holder was removed.');
                               }}); });
    case 'regenToken':   return mgmtDo_(mgmt, function(){ updateRow_('Employees',req.id,{token:token_()}); });
    case 'addBlackout':  return mgmtDo_(mgmt, function(){
                             if (!req.start || !req.end || req.end < req.start) throw new Error('Pick a valid date range.');
                             addRow_('Blackouts',{id:uid_(),start:req.start,end:req.end,reason:req.reason||''}); });
    case 'deleteBlackout':return mgmtDo_(mgmt, function(){ deleteRow_('Blackouts',req.id); });

    default: return {ok:false, error:'Unknown action: '+action};
  }
}

function mgmtDo_(mgmt, fn){
  if (!mgmt) return {ok:false, error:'Management PIN required.'};
  fn();
  audit_('Management','admin.change','');
  return {ok:true, state: stateFor_(null, true)};
}

/* ---------------- swap resolution (true exchange) ---------------- */
function respondSwap_(me, req){
  var w = findBy_('Swaps','id',req.swapId);
  if (!w || w.status !== 'pending') return {ok:false, error:'Swap request not found or already resolved.'};
  var tb = findBy_('Bookings','id',w.targetBookingId);
  if (!tb || tb.status !== 'approved'){
    updateRow_('Swaps', w.id, {status:'failed', note:'Booking no longer exists.', resolvedAt:String(Date.now())});
    return {ok:false, error:'That booking no longer exists — swap closed.'};
  }
  if (tb.employeeId !== me.id) return {ok:false, error:'This swap is not addressed to you.'};

  if (!req.accept){
    updateRow_('Swaps', w.id, {status:'declined', resolvedAt:String(Date.now())});
    audit_(me.name,'swap.declined', 'from '+empName_(w.requesterId));
    notify_(w.requesterId,'Swap declined', me.name+' declined your holiday swap request for '+w.reqStart+' → '+w.reqEnd+'.');
    return {ok:true, state: stateFor_(me,false)};
  }

  // 1) Requester's dates must be valid once my booking is released
  var vA = validate_(w.requesterId, w.reqStart, w.reqEnd, {exclude:[tb.id]});
  if (!vA.ok){
    updateRow_('Swaps', w.id, {status:'failed', note:vA.reason, resolvedAt:String(Date.now())});
    notify_(w.requesterId,'Swap failed','Your swap with '+me.name+' could not complete: '+vA.reason);
    return {ok:false, error:'Swap failed: '+vA.reason, state: stateFor_(me,false)};
  }

  // 2) True exchange: releaser picks new dates (validated with the requester's new booking in place)
  if (!req.releaseOnly){
    if (!req.newStart || !req.newEnd) return {ok:false, error:'Pick your replacement dates (or choose release-only).'};
    var vB = validate_(me.id, req.newStart, req.newEnd, {
      exclude:[tb.id],
      extra:[{employeeId:w.requesterId, start:w.reqStart, end:w.reqEnd}]
    });
    if (!vB.ok) return {ok:false, error:'Your new dates don\'t work: '+vB.reason+' — pick different dates.', keepPending:true};
  }

  // 3) Commit atomically (we hold the script lock)
  updateRow_('Bookings', tb.id, {status:'cancelled', note:'Exchanged via swap'});
  schedPush_(tb, 'cancel');
  var reqBooking = {id:uid_(), employeeId:w.requesterId, start:w.reqStart, end:w.reqEnd,
                    status:'approved', note:'Via swap with '+me.name, createdAt:String(Date.now())};
  addRow_('Bookings', reqBooking);
  schedPush_(reqBooking, 'approve');
  if (!req.releaseOnly){
    var myBooking = {id:uid_(), employeeId:me.id, start:req.newStart, end:req.newEnd,
                     status:'approved', note:'Exchanged via swap', createdAt:String(Date.now())};
    addRow_('Bookings', myBooking);
    schedPush_(myBooking, 'approve');
  }
  updateRow_('Swaps', w.id, {status:'accepted', resolvedAt:String(Date.now()),
    note: req.releaseOnly ? 'Released without rebooking' : 'Exchanged for '+req.newStart+' → '+req.newEnd});
  failPendingSwapsFor_(tb.id, 'Booking was exchanged in another swap.');
  audit_(me.name,'swap.accepted', 'with '+empName_(w.requesterId));
  notify_(w.requesterId,'Swap accepted',
    me.name+' accepted your swap — your holidays '+w.reqStart+' → '+w.reqEnd+' are approved.');
  return {ok:true, state: stateFor_(me,false)};
}

function failPendingSwapsFor_(bookingId, note){
  readAll_('Swaps').forEach(function(w){
    if (w.targetBookingId === bookingId && w.status === 'pending')
      updateRow_('Swaps', w.id, {status:'failed', note:note, resolvedAt:String(Date.now())});
  });
}

/* ---------------- management auth ----------------
 * Same protections Admin.gs already applies to Admin_PIN: constant-time
 * comparison and a failed-attempt lockout. Without these, MGMT_PIN guards
 * every management action on an "Anyone" endpoint with nothing but a plain
 * === against four digits.
 *
 * The counter is global because doPost gets no reliable client IP, so anyone
 * with the /exec URL could trip it deliberately. It is tuned loose for that
 * reason — enough to stop guessing, not enough to be a usable denial of
 * service. PIN LENGTH is the real control.
 * ------------------------------------------------------------------------- */
var FL_AUTH = { PIN_PROP: 'Mgmt_PIN', MAX_FAILS: 12, LOCKOUT_MINUTES: 5, MIN_PIN_LENGTH: 8 };

/** Sentinel left in the Config tab so the old row explains itself. */
var FL_PIN_MOVED = 'SEE_SCRIPT_PROPERTIES';

/**
 * The management PIN. Script Properties ONLY — never the Config tab.
 *
 * A PIN in a Config cell is readable by anyone with edit access to the
 * workbook, survives in version history, and rides along in every export or
 * copy of the sheet. Script Properties are visible only to someone who can
 * open the Apps Script project.
 *
 * There is deliberately NO fallback to the Config tab. A fallback would mean
 * the sheet value still worked, which is the whole thing being fixed.
 *
 * Note this project also holds Admin.gs, which keeps its own Admin_PIN in a
 * separate property. Two different doors, two different keys.
 */
function flMgmtPin_(){
  try {
    return String(PropertiesService.getScriptProperties()
                    .getProperty(FL_AUTH.PIN_PROP) || '');
  } catch (e) { return ''; }
}

/**
 * route_() evaluates management auth on EVERY request, including ordinary
 * employee traffic carrying no PIN at all. A missing PIN must never count as
 * a failed attempt, or normal staff use would lock management out.
 */
function mgmtAuth_(pin){
  if (pin === undefined || pin === null || String(pin) === '') return false;
  var stored = flMgmtPin_();
  if (!stored){
    audit_('system','mgmt.noPinSet',
      'Script Property "'+FL_AUTH.PIN_PROP+'" is not set — management is unreachable. '+
      'Run migrateMgmtPinToProperties().');
    return false;
  }

  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fl_mgmt_fails') || 0);
  if (fails >= FL_AUTH.MAX_FAILS) return false;

  if (flConstantEquals_(String(pin), stored)){
    cache.remove('fl_mgmt_fails');
    return true;
  }
  cache.put('fl_mgmt_fails', String(fails + 1), FL_AUTH.LOCKOUT_MINUTES * 60);
  audit_('unknown','mgmt.authFail','Management PIN attempt '+(fails+1)+' of '+FL_AUTH.MAX_FAILS);
  return false;
}

/** Compares the whole length regardless of mismatch position — no early-exit signal. */
function flConstantEquals_(a, b){
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++){
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

/** True when the PIN is short enough to be worth brute-forcing. */
function mgmtPinWeak_(){
  return flMgmtPin_().length < FL_AUTH.MIN_PIN_LENGTH;
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
function migrateMgmtPinToProperties(){
  var props = PropertiesService.getScriptProperties();
  var existing = String(props.getProperty(FL_AUTH.PIN_PROP) || '');
  var sheetVal = String(getConfig_('MGMT_PIN') || '');
  var out = [], say = function(t){ out.push(t); };

  say('MIGRATE  Config tab MGMT_PIN  →  Script Property "'+FL_AUTH.PIN_PROP+'"');
  say('');

  if (existing && (!sheetVal || sheetVal === FL_PIN_MOVED)){
    say('✔ Already migrated. The property holds '+existing.length+' characters and');
    say('  the Config tab holds no PIN. Nothing to do.');
    Logger.log(out.join('\n')); return out.join('\n');
  }
  if (existing && sheetVal && sheetVal !== FL_PIN_MOVED){
    say('✖ REFUSED — both places hold a value.');
    say('    property   : '+existing.length+' characters');
    say('    Config tab : '+sheetVal.length+' characters');
    say('');
    say('  I will not guess which one you mean, and overwriting the property could');
    say('  lock you out. Decide, set '+FL_AUTH.PIN_PROP+' by hand in Project');
    say('  Settings, then run this again to clear the sheet.');
    Logger.log(out.join('\n')); return out.join('\n');
  }
  if (!sheetVal || sheetVal === FL_PIN_MOVED){
    say('✖ Nothing to migrate: the Config tab has no MGMT_PIN value and the');
    say('  property is not set. Management is currently unreachable.');
    say('');
    say('  Set '+FL_AUTH.PIN_PROP+' by hand:');
    say('    Apps Script → ⚙ Project Settings → Script Properties → Add');
    say('    Use 8+ characters, not 4 digits.');
    Logger.log(out.join('\n')); return out.join('\n');
  }

  props.setProperty(FL_AUTH.PIN_PROP, sheetVal);
  var check = String(props.getProperty(FL_AUTH.PIN_PROP) || '');
  if (check !== sheetVal){
    say('✖ The property did not save correctly. The Config tab has NOT been touched,');
    say('  so your current PIN still works. Try again.');
    Logger.log(out.join('\n')); return out.join('\n');
  }
  say('  1. Property written and read back  ✔  ('+check.length+' characters)');

  setConfig_('MGMT_PIN', FL_PIN_MOVED);
  say('  2. Config tab cell replaced with "'+FL_PIN_MOVED+'"  ✔');

  audit_('management','mgmt.pinMigrated',
    'MGMT_PIN moved from the Config tab to Script Property '+FL_AUTH.PIN_PROP);

  say('');
  say('✔ Done. The PIN is the same — only where it lives has changed.');
  say('');
  say('IMPORTANT');
  say('  • The value is still in this sheet\'s VERSION HISTORY. If it was ever a');
  say('    real secret, change it now: Project Settings → Script Properties.');
  say('  • It is '+check.length+' characters. '+
      (check.length < FL_AUTH.MIN_PIN_LENGTH
        ? 'That is below the '+FL_AUTH.MIN_PIN_LENGTH+'-character minimum — change it.'
        : 'That is acceptable.'));
  say('  • Run fairleaveAuthCheck() to confirm.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/** Read-only. Reports the state of management auth. Run from the dropdown. */
function fairleaveAuthCheck(){
  var pin = flMgmtPin_();
  var inSheet = String(getConfig_('MGMT_PIN') || '');
  var admin = '';
  try { admin = String(PropertiesService.getScriptProperties().getProperty('Admin_PIN') || ''); } catch(e){}
  var fails = Number(CacheService.getScriptCache().get('fl_mgmt_fails') || 0);
  var out = [];

  out.push('FAIRLEAVE — management authentication');
  out.push('');
  out.push('  PIN source          : Script Property "'+FL_AUTH.PIN_PROP+'"');
  out.push('  property set        : '+(pin ? 'yes, '+pin.length+' characters' : 'NO'));
  out.push('  minimum recommended : '+FL_AUTH.MIN_PIN_LENGTH);
  out.push('  constant-time check : yes');
  out.push('  lockout             : '+FL_AUTH.MAX_FAILS+' failures / '+FL_AUTH.LOCKOUT_MINUTES+' minutes');
  out.push('  failures right now  : '+fails+(fails >= FL_AUTH.MAX_FAILS ? '  ← LOCKED' : ''));
  out.push('');
  out.push('CONFIG TAB');
  if (!inSheet)                     out.push('  MGMT_PIN row        : blank ✔');
  else if (inSheet === FL_PIN_MOVED) out.push('  MGMT_PIN row        : "'+FL_PIN_MOVED+'" ✔  (migrated)');
  else {
    out.push('  MGMT_PIN row        : ✖ STILL HOLDS A VALUE ('+inSheet.length+' characters)');
    out.push('    No longer used for authentication, but anyone with edit access to');
    out.push('    this workbook can read it. Run migrateMgmtPinToProperties().');
  }
  out.push('');
  out.push('THE OTHER DOOR IN THIS PROJECT');
  out.push('  Admin.gs (Data Maintenance console) is bound to this same project and');
  out.push('  keeps its own key: Script Property "Admin_PIN" — '+
           (admin ? admin.length+' characters' : 'NOT SET'));
  out.push('  Two different doors, two different keys. Changing one does not change');
  out.push('  the other.');
  out.push('');
  if (!pin) {
    out.push('✖ No management PIN is set. Every management action is refused.');
    out.push('  Run migrateMgmtPinToProperties(), or set '+FL_AUTH.PIN_PROP+' by hand.');
  } else if (/^\d+$/.test(pin) && pin.length <= 4) {
    out.push('✖ The PIN is '+pin.length+' digits — '+Math.pow(10,pin.length)+' possibilities.');
    out.push('  The lockout is the only thing in the way of a scripted guess. Change it.');
  } else if (mgmtPinWeak_()) {
    out.push('⚠ The PIN is shorter than '+FL_AUTH.MIN_PIN_LENGTH+' characters. Lengthen it.');
  } else {
    out.push('✔ PIN length is reasonable and it is stored outside the spreadsheet.');
  }
  out.push('');
  out.push('Nothing has been changed.');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ---------------- rules engine ---------------- */
function validate_(empId, s, e, opts){
  opts = opts || {}; var exclude = opts.exclude || []; var extra = opts.extra || [];
  var t = todayISO_();
  if (!s || !e) return {ok:false, reason:'Pick both a start and an end date.'};
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || !/^\d{4}-\d{2}-\d{2}$/.test(e)) return {ok:false, reason:'Bad date format.'};
  if (e < s) return {ok:false, reason:'End date is before start date.'};
  if (s < t) return {ok:false, reason:'Start date is in the past.'};

  // Chargeable days can legitimately be ZERO — a Saturday off costs no holiday
  // but is still a day the business needs covered. Gate on OPERATING days, or
  // booking a Saturday off would be rejected as "no working days in range".
  var days = workdaysIn_(s,e);
  if (operatingDaysIn_(s,e) === 0) return {ok:false, reason:'No working days in that range.'};

  var bo = readAll_('Blackouts').filter(function(b){ return overlaps_(s,e,b.start,b.end); })[0];
  if (bo) return {ok:false, reason:'Blocked: no holidays '+bo.start+' → '+bo.end+' ('+(bo.reason||'blackout')+').'};

  var emp = findBy_('Employees','id',empId);
  if (!emp) return {ok:false, reason:'Employee not found.'};
  var remaining = (parseInt(emp.entitlement,10)||0) - usedDays_(empId, exclude);
  if (days > remaining) return {ok:false, reason:'Not enough entitlement: needs '+days+' days, '+remaining+' left.'};

  var mine = approved_(exclude).filter(function(b){
    return b.employeeId===empId && overlaps_(s,e,b.start,b.end);
  })[0];
  if (mine) return {ok:false, reason:'You already have approved leave overlapping these dates ('+mine.start+' → '+mine.end+').'};

  var team = findBy_('Teams','id',emp.teamId);
  var max = parseInt(team.max,10)||1;
  var pool = approved_(exclude).concat(extra.map(function(x){
    return {id:'extra', employeeId:x.employeeId, start:x.start, end:x.end, status:'approved'};
  }));
  var empsById = indexBy_(readAll_('Employees'),'id');
  var conflictIds = {}; var fullDay = null;
  eachDay_(s,e).forEach(function(d){
    if (isNonWorking_(d)) return;          // OPERATING days — Saturday included
    var covering = pool.filter(function(b){
      var be = empsById[b.employeeId];
      return be && be.teamId===emp.teamId && b.employeeId!==empId && b.start<=d && d<=b.end;
    });
    if (covering.length >= max){
      fullDay = fullDay || d;
      covering.forEach(function(b){ if (b.id!=='extra') conflictIds[b.id]=true; });
    }
  });
  var keys = Object.keys(conflictIds);
  if (keys.length){
    var all = indexBy_(readAll_('Bookings'),'id');
    return {ok:false, reason:'Team limit reached (max '+max+' off at once) — first full day: '+fullDay+'.',
            conflicts: keys.map(function(k){ return all[k]; })};
  }
  return {ok:true, days:days};
}

function approved_(exclude){
  exclude = exclude || [];
  return readAll_('Bookings').filter(function(b){
    return b.status==='approved' && exclude.indexOf(b.id)===-1;
  });
}
function usedDays_(empId, exclude){
  return approved_(exclude).filter(function(b){ return b.employeeId===empId; })
    .reduce(function(n,b){ return n + workdaysIn_(b.start,b.end); }, 0);
}

/* ---------------- state payloads ---------------- */
function stateFor_(me, isMgmt){
  var emps = readAll_('Employees').filter(function(e){ return e.active!=='FALSE'; });
  return {
    company: getConfig_('COMPANY_NAME'),
    workSaturday: workSaturday_(),   // OPERATING days — capacity, cover, rota
    chargeSaturday: false,           // ENTITLEMENT — Sat/Sun never charged
    teams: readAll_('Teams').map(function(t){ return {id:t.id,name:t.name,max:parseInt(t.max,10)||1}; }),
    employees: emps.map(function(e){ return isMgmt ? mgmtEmp_(e) : pubEmp_(e); }),
    blackouts: readAll_('Blackouts'),
    bookings: readAll_('Bookings'),
    swaps: readAll_('Swaps')
  };
}
function pubEmp_(e){ return {id:e.id,name:e.name,teamId:e.teamId,entitlement:parseInt(e.entitlement,10)||0,used:usedDays_(e.id,[])}; }
function mgmtEmp_(e){ var p=pubEmp_(e); p.email=e.email; p.token=e.token; return p; }

/* ---------------- notifications & audit ---------------- */
function notify_(empId, subject, body){
  try{
    var e = findBy_('Employees','id',empId);
    if (e && e.email) MailApp.sendEmail(e.email, '['+getConfig_('COMPANY_NAME')+'] '+subject, body);
  }catch(err){ /* email is best-effort */ }
}
function audit_(actor, action, details){
  addRow_('Audit',{timestamp:new Date().toISOString(), actor:actor, action:action, details:details});
}

/* ---------------- date utils (string-based) ----------------
 * TWO different questions, deliberately answered by two different functions:
 *
 *   isNonWorking_()     Does the business RUN that day?   → capacity, cover
 *   isNonChargeable_()  Does it cost HOLIDAY?             → entitlement balance
 *
 * Saturday is yes to the first and no to the second. One function answering
 * both is why WORK_SATURDAY could not express the actual policy.
 * ------------------------------------------------------------------------- */
function todayISO_(){ return Utilities.formatDate(new Date(), 'Europe/Dublin', 'yyyy-MM-dd'); }
function pDate_(s){ var p=s.split('-'); return new Date(Date.UTC(+p[0],+p[1]-1,+p[2])); }
function addDays_(s,n){ var d=pDate_(s); d.setUTCDate(d.getUTCDate()+n); return d.toISOString().slice(0,10); }
function eachDay_(s,e){ var out=[],c=s; while(c<=e){ out.push(c); c=addDays_(c,1);} return out; }

/**
 * Is Saturday an OPERATING day? Read through one helper, case-insensitively.
 * A stray lowercase 'true' in Config used to fail the old ==='TRUE' test and
 * silently switch the whole business to Mon–Fri, which is not a failure mode
 * anyone notices until leave balances are wrong.
 */
function workSaturday_(){
  return String(getConfig_('WORK_SATURDAY')).trim().toUpperCase() === 'TRUE';
}

/** Days the business does NOT run. Drives team capacity. */
function isNonWorking_(s){
  var wd = pDate_(s).getUTCDay(); // 0 = Sunday
  if (workSaturday_()) return wd===0;
  return wd===0 || wd===6;
}

/**
 * Days that do NOT cost entitlement. Saturday and Sunday never cost holiday,
 * whatever WORK_SATURDAY says — Saturday is worked, but a Saturday inside
 * someone's leave must not come off a balance sized for a 5-day contract.
 */
function isNonChargeable_(s){
  var wd = pDate_(s).getUTCDay(); // 0 = Sunday, 6 = Saturday
  return wd===0 || wd===6;
}

/** Days that cost entitlement. Used for balances and "days needed". */
function workdaysIn_(s,e){
  return eachDay_(s,e).filter(function(d){ return !isNonChargeable_(d); }).length;
}

/** Days the business operates. Used for capacity and range validity. */
function operatingDaysIn_(s,e){
  return eachDay_(s,e).filter(function(d){ return !isNonWorking_(d); }).length;
}

function overlaps_(aS,aE,bS,bE){ return aS<=bE && bS<=aE; }

/* ---------------- sheet I/O helpers ---------------- */
function sheet_(name){ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function readAll_(name){
  var sh = sheet_(name); var last = sh.getLastRow();
  if (last < 2) return [];
  var head = TABS[name];
  var vals = sh.getRange(2,1,last-1,head.length).getValues();
  return vals.map(function(row){
    var o = {};
    head.forEach(function(h,i){
      var v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v,'Europe/Dublin','yyyy-MM-dd');
      o[h] = String(v);
    });
    return o;
  }).filter(function(o){ return o.id !== '' || name==='Config' || name==='Audit'; });
}
function addRow_(name, obj){
  var head = TABS[name];
  sheet_(name).appendRow(head.map(function(h){ return obj[h]!==undefined ? obj[h] : ''; }));
  return obj;
}
function rowIndexOf_(name, id){
  var sh = sheet_(name); var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2,1,last-1,1).getValues();
  for (var i=0;i<ids.length;i++) if (String(ids[i][0])===String(id)) return i+2;
  return -1;
}
function updateRow_(name, id, patch){
  var r = rowIndexOf_(name,id); if (r<0) return;
  var head = TABS[name]; var sh = sheet_(name);
  var row = sh.getRange(r,1,1,head.length).getValues()[0];
  head.forEach(function(h,i){ if (patch[h]!==undefined) row[i]=patch[h]; });
  sh.getRange(r,1,1,head.length).setValues([row]);
}
function deleteRow_(name, id){
  var r = rowIndexOf_(name,id); if (r>0) sheet_(name).deleteRow(r);
}
function findBy_(name, field, val){
  return readAll_(name).filter(function(o){ return o[field]===String(val); })[0] || null;
}
function indexBy_(arr, key){ var o={}; arr.forEach(function(x){o[x[key]]=x;}); return o; }
function getConfig_(k){ var r=findByCfg_(k); return r ? r.value : ''; }
function setConfig_(k,v){
  var sh = sheet_('Config'); var last = sh.getLastRow();
  for (var i=2;i<=last;i++) if (String(sh.getRange(i,1).getValue())===k){ sh.getRange(i,2).setValue(v); return; }
  sh.appendRow([k,v]);
}
function findByCfg_(k){
  return readAll_('Config').filter(function(o){ return o.key===k; })[0] || null;
}
function empName_(id){ var e=findBy_('Employees','id',id); return e?e.name:'—'; }
function uid_(){ return Utilities.getUuid().replace(/-/g,'').slice(0,10); }
function token_(){ return Utilities.getUuid().replace(/-/g,'').slice(0,16); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/**********************************************************************
 * SHIFT SCHEDULER INTEGRATION (v2.2)
 * Talks to the Zanna-pattern Shift Scheduler spreadsheet:
 *   - Employees + Departments  →  imported INTO FairLeave (scheduler owns them)
 *   - FairLeave approved leave →  pushed INTO scheduler's LeaveRequests
 *     (rows tagged "FairLeave:<bookingId>" in Notes for idempotent sync;
 *      the scheduler's auto-planner already treats Approved leave as a
 *      hard constraint, so pushed leave blocks shifts automatically)
 *   - Existing approved scheduler leave → one-time import as bookings
 * Column headers are resolved dynamically (RequestID/LeaveID, Start/
 * StartDate, etc.) so minor schema differences don't break the sync.
 * All pushes are best-effort: a scheduler problem NEVER blocks a
 * FairLeave booking — failures are recorded in the Audit tab.
 * Run diagnoseScheduler() in the editor to see what was detected.
 *
 * KNOWN GAP (audit item #7): cEnt below does NOT list 'AnnualLeaveDays',
 * which is the column the Zanna Shift Scheduler actually uses. Entitlement
 * therefore never syncs and every employee falls through to the '20' default.
 * Adding the alias flips entitlement ownership to the Scheduler and overwrites
 * all balances on the next sync — do not add it until the real per-person
 * figures are in the Scheduler.
 **********************************************************************/

var FL_TAG = 'FairLeave:';

function schedSS_(){
  var id = getConfig_('SCHEDULER_SHEET_ID');
  if (!id) return null;
  return SpreadsheetApp.openById(id);
}

function schedTab_(ss, names){
  for (var i=0;i<names.length;i++){
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}

/* Resolve a header row into {colName: index} with alias support */
function headerMap_(sh){
  var head = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
  var map = {};
  head.forEach(function(h,i){ map[String(h).toLowerCase().replace(/[\s_]/g,'')] = i; });
  return {raw:head, find:function(aliases){
    for (var i=0;i<aliases.length;i++){
      var k = aliases[i].toLowerCase().replace(/[\s_]/g,'');
      if (map[k] !== undefined) return map[k];
    }
    return -1;
  }};
}

function schedDateISO_(v){
  if (v instanceof Date) return Utilities.formatDate(v,'Europe/Dublin','yyyy-MM-dd');
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // dd/mm/yyyy
  if (m) return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
  return '';
}

/* ---- 1) Import employees + departments from the scheduler ---- */
function syncFromScheduler_(){
  var ss;
  try{ ss = schedSS_(); }catch(e){ return {ok:false, error:'Cannot open scheduler sheet: '+e.message}; }
  if (!ss) return {ok:false, error:'SCHEDULER_SHEET_ID is not set in Config.'};
  var sh = schedTab_(ss, ['Employees','Staff','People']);
  if (!sh) return {ok:false, error:'No Employees tab found in the scheduler sheet.'};

  var H = headerMap_(sh);
  var cId   = H.find(['EmployeeID','EmpID','ID','StaffID']);
  var cName = H.find(['Name','EmployeeName','FullName','Employee']);
  var cDept = H.find(['Department','Dept','Team','Section']);
  var cAct  = H.find(['Active','IsActive','Enabled']);
  var cEnt  = H.find(['AnnualLeaveEntitlement','AnnualLeave','Entitlement','ALDays','LeaveDays','HolidayEntitlement']);
  var cMail = H.find(['Email','EmailAddress','Mail']);
  if (cId<0 || cName<0) return {ok:false, error:'Scheduler Employees tab is missing an ID or Name column. Headers found: '+H.raw.join(', ')};

  var last = sh.getLastRow();
  var rows = last>1 ? sh.getRange(2,1,last-1,sh.getLastColumn()).getValues() : [];

  var flEmps = readAll_('Employees');
  var byExt = {}; flEmps.forEach(function(e){ if(e.extId) byExt[e.extId]=e; });
  var flTeams = readAll_('Teams');
  var teamByName = {}; flTeams.forEach(function(t){ teamByName[t.name.toLowerCase()]=t; });

  var added=0, updated=0, teamsAdded=0, deactivated=0, seen={};

  rows.forEach(function(r){
    var extId = String(r[cId]).trim();
    var name = String(r[cName]).trim();
    if (!extId || !name) return;
    var active = cAct<0 ? true : String(r[cAct]).toLowerCase()!=='false' && String(r[cAct])!=='0' && String(r[cAct]).toLowerCase()!=='no';
    if (!active) return;
    seen[extId] = true;

    var deptName = cDept>=0 && String(r[cDept]).trim() ? String(r[cDept]).trim() : 'General';
    var team = teamByName[deptName.toLowerCase()];
    if (!team){
      team = {id:uid_(), name:deptName, max:'1'};
      addRow_('Teams', team);
      teamByName[deptName.toLowerCase()] = team;
      teamsAdded++;
    }
    var ent = cEnt>=0 && r[cEnt]!=='' ? String(parseInt(r[cEnt],10)||20) : '20';
    var email = cMail>=0 ? String(r[cMail]).trim() : '';

    var existing = byExt[extId];
    if (existing){
      var patch = {name:name, teamId:team.id, active:'TRUE'};
      if (cEnt>=0 && r[cEnt]!=='') patch.entitlement = ent;
      if (email && !existing.email) patch.email = email;
      updateRow_('Employees', existing.id, patch);
      updated++;
    } else {
      addRow_('Employees', {id:uid_(), name:name, teamId:team.id, entitlement:ent,
        email:email, token:token_(), active:'TRUE', extId:extId});
      added++;
    }
  });

  // deactivate FairLeave employees whose scheduler record disappeared/inactive
  flEmps.forEach(function(e){
    if (e.extId && !seen[e.extId] && e.active!=='FALSE'){
      updateRow_('Employees', e.id, {active:'FALSE'});
      deactivated++;
    }
  });

  var summary = 'Sync complete: '+added+' added, '+updated+' updated, '+teamsAdded+' teams created, '+deactivated+' deactivated. Review team "max off at once" limits for any new teams (default 1).';
  audit_('Management','sync.employees', summary);
  return {ok:true, summary:summary};
}

/* ---- 2) Push FairLeave leave into the scheduler's LeaveRequests ---- */
function schedLeaveCols_(sh){
  var H = headerMap_(sh);
  return {
    H:H,
    id:    H.find(['RequestID','LeaveID','ID']),
    emp:   H.find(['EmployeeID','EmpID','StaffID','Employee']),
    start: H.find(['StartDate','Start','From','FromDate']),
    end:   H.find(['EndDate','End','To','ToDate']),
    type:  H.find(['Type','LeaveType','Category']),
    status:H.find(['Status','State','Approval','ApprovalStatus']),
    notes: H.find(['Notes','Note','Comments','Reason'])
  };
}

function schedPush_(booking, action){
  try{
    var ss = schedSS_(); if (!ss) return;
    var sh = schedTab_(ss, ['LeaveRequests','Leave','LeaveLog']);
    if (!sh){ audit_('System','sync.error','No LeaveRequests tab in scheduler sheet'); return; }
    var e = findBy_('Employees','id',booking.employeeId);
    if (!e || !e.extId){ audit_('System','sync.skipped','No scheduler ID for '+(e?e.name:booking.employeeId)); return; }
    var C = schedLeaveCols_(sh);
    if (C.emp<0 || C.start<0 || C.end<0){ audit_('System','sync.error','LeaveRequests headers not recognised: '+C.H.raw.join(', ')); return; }

    var tag = FL_TAG + booking.id;
    var last = sh.getLastRow();
    var found = -1;
    if (last>1){
      var scan = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
      for (var i=0;i<scan.length;i++){
        var idMatch = C.id>=0 && String(scan[i][C.id])===tag;
        var noteMatch = C.notes>=0 && String(scan[i][C.notes]).indexOf(tag)>-1;
        if (idMatch || noteMatch){ found = i+2; break; }
      }
    }

    if (action==='approve'){
      if (found>0){ if (C.status>=0) sh.getRange(found, C.status+1).setValue('Approved'); return; }
      var row = [];
      for (var c=0;c<sh.getLastColumn();c++) row.push('');
      if (C.id>=0)    row[C.id]    = tag;
      row[C.emp]   = e.extId;
      row[C.start] = booking.start;
      row[C.end]   = booking.end;
      if (C.type>=0)  row[C.type]  = 'Annual';
      if (C.status>=0)row[C.status]= 'Approved';
      if (C.notes>=0) row[C.notes] = tag+' · self-managed, auto-approved';
      sh.appendRow(row);
      audit_('System','sync.push', e.name+' '+booking.start+' → '+booking.end+' → scheduler');
    }
    if (action==='cancel' && found>0){
      if (C.status>=0) sh.getRange(found, C.status+1).setValue('Cancelled');
      else sh.deleteRow(found);
      audit_('System','sync.cancel', e.name+' '+booking.start+' → '+booking.end+' cancelled in scheduler');
    }
  }catch(err){
    audit_('System','sync.error', String(err.message||err));
  }
}

/* ---- 3) One-time import of existing approved scheduler leave ---- */
function importLeaveFromScheduler_(){
  var ss;
  try{ ss = schedSS_(); }catch(e){ return {ok:false, error:'Cannot open scheduler sheet: '+e.message}; }
  if (!ss) return {ok:false, error:'SCHEDULER_SHEET_ID is not set in Config.'};
  var sh = schedTab_(ss, ['LeaveRequests','Leave','LeaveLog']);
  if (!sh) return {ok:false, error:'No LeaveRequests tab found in the scheduler sheet.'};
  var C = schedLeaveCols_(sh);
  if (C.emp<0 || C.start<0 || C.end<0) return {ok:false, error:'LeaveRequests headers not recognised. Found: '+C.H.raw.join(', ')};

  var flEmps = readAll_('Employees');
  var byExt = {}; flEmps.forEach(function(e){ if(e.extId) byExt[e.extId]=e; });
  var existing = readAll_('Bookings');

  var last = sh.getLastRow();
  var rows = last>1 ? sh.getRange(2,1,last-1,sh.getLastColumn()).getValues() : [];
  var imported=0, skipped=0;

  rows.forEach(function(r){
    var noteVal = C.notes>=0 ? String(r[C.notes]) : '';
    var idVal = C.id>=0 ? String(r[C.id]) : '';
    if (noteVal.indexOf(FL_TAG)>-1 || idVal.indexOf(FL_TAG)===0) return; // originated here
    var status = C.status>=0 ? String(r[C.status]).toLowerCase() : 'approved';
    if (status.indexOf('approve')<0) return;
    var e = byExt[String(r[C.emp]).trim()];
    if (!e){ skipped++; return; }
    var s = schedDateISO_(r[C.start]), en = schedDateISO_(r[C.end]);
    if (!s || !en) { skipped++; return; }
    var dup = existing.some(function(b){
      return b.employeeId===e.id && b.status==='approved' && overlaps_(s,en,b.start,b.end);
    });
    if (dup){ skipped++; return; }
    addRow_('Bookings', {id:uid_(), employeeId:e.id, start:s, end:en,
      status:'approved', note:'Imported from Shift Scheduler', createdAt:String(Date.now())});
    imported++;
  });

  var summary = 'Import complete: '+imported+' approved leave records imported, '+skipped+' skipped (unknown employee, bad dates, or already in FairLeave).';
  audit_('Management','sync.import', summary);
  return {ok:true, summary:summary};
}

/* ---- Diagnostics: run in the Apps Script editor, check the log ---- */
function diagnoseScheduler(){
  var ss = schedSS_();
  if (!ss){ Logger.log('SCHEDULER_SHEET_ID not set in Config.'); return; }
  Logger.log('Scheduler workbook: '+ss.getName());
  ss.getSheets().forEach(function(sh){ Logger.log('Tab: '+sh.getName()+' ('+sh.getLastRow()+' rows)'); });
  var emp = schedTab_(ss, ['Employees','Staff','People']);
  if (emp) Logger.log('Employees headers: '+emp.getRange(1,1,1,emp.getLastColumn()).getValues()[0].join(' | '));
  var lv = schedTab_(ss, ['LeaveRequests','Leave','LeaveLog']);
  if (lv) Logger.log('LeaveRequests headers: '+lv.getRange(1,1,1,lv.getLastColumn()).getValues()[0].join(' | '));
}

/**
 * Shows exactly what a date range costs versus what it occupies, so the
 * Saturday rule can be checked without creating a real booking.
 * Run from the editor, check the log.
 */
function diagnoseLeaveCounting(){
  var samples = [
    ['2026-08-03','2026-08-07','Mon–Fri'],
    ['2026-08-03','2026-08-08','Mon–Sat'],
    ['2026-08-03','2026-08-09','Mon–Sun'],
    ['2026-08-08','2026-08-08','Saturday only'],
    ['2026-08-09','2026-08-09','Sunday only']
  ];
  var out = ['WORK_SATURDAY = '+getConfig_('WORK_SATURDAY')+'  (operating days)',
             'Saturday/Sunday never cost holiday.', '',
             'range                     label           costs   occupies'];
  samples.forEach(function(s){
    var pad = function(v,n){ v=String(v); while(v.length<n) v+=' '; return v; };
    out.push(pad(s[0]+' → '+s[1],26)+pad(s[2],16)+
             pad(workdaysIn_(s[0],s[1]),8)+operatingDaysIn_(s[0],s[1]));
  });
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* testPanel() removed — it carried a hardcoded '123456' access token. */

/** Control Panel: verify access token, return app URL map. */
function panelConfig_(p) {
  var out = { ok: false };
  try {
    var props = PropertiesService.getScriptProperties();
    var expected = props.getProperty('Access_Token');
    if (expected && p.token && String(p.token).trim() === expected) {
      out = { ok: true, apps: JSON.parse(props.getProperty('Panel_Apps') || '{}') };
    }
  } catch (err) {
    out = { ok: false };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
function whereAmI() {
  var ss = SpreadsheetApp.getActive();
  Logger.log(ss.getName() + '  |  ' + ss.getId());
}