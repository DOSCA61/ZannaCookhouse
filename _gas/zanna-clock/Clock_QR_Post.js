/**
 * Clock_QR_Post.gs — the missing POST route for the phone clock-in page
 * v1.2.0 · targets Zanna Clock v1.11.0 + Clock_QR.gs v1.1.0 · written 2026-08-19
 *
 * WHY THIS FILE EXISTS
 * clock-mobile.html v1.0.1 (GitHub Pages, 18 Aug) POSTs to the Clock's /exec.
 * Clock_QR.gs v1.1.0 serves its phone page from Apps Script itself and talks
 * back over google.script.run, so it deliberately has NO doPost. An Apps Script
 * web app with no doPost does not answer POST at all — the phone waits, gets
 * nothing, and shows its own "Sorry" card. That is the 404 staff reported on
 * 19 Aug. The front-end half of the 18 Aug work shipped; this half did not.
 *
 * Measured 19 Aug against the live deployment, cross-origin from github.io:
 *   GET  -> 200 {"ok":true,"app":"ZannaClock","version":"1.11.0"}  ~1.9s
 *   POST -> no response at all, aborted at 9s
 *
 * WHAT IT DOES NOT DO
 * It does not reimplement anything. It parses the body, checks the route, and
 * hands straight to clkQrHello / clkQrSubmit — the same two functions the
 * Apps Script-served page already calls over google.script.run. One state
 * machine, not two. If those functions are absent this file fails loudly at
 * install rather than quietly at 6am.
 *
 * THE CONTRACT (read off clock-mobile.html v1.0.1, not guessed)
 *   POST body, JSON, sent as text/plain so there is no CORS preflight:
 *     { qp:'mclk', action:'hello',  t:<personal token> }
 *     { qp:'mclk', action:'submit', t:<personal token>, s:<8-char code>, a:'in'|'out' }
 *   Response: the plain object clkQrHello / clkQrSubmit already return.
 *     hello  -> { ok, name, first, state, since, onBreak, allowOut, outNeedsScan }
 *     submit -> { ok, action, at, name, dryRun }
 *     either -> { ok:false, code, error } on refusal
 *
 * INSTALL — TWO STEPS, IN ORDER. DO NOT SKIP STEP 1.
 * Every .gs file in an Apps Script project shares one namespace and a
 * duplicated function name is resolved by taking the last one, silently. If a
 * doPost already exists anywhere in this project, pasting another one will
 * break whatever the first one served, with no warning.
 *
 *   1. Paste this file as a NEW script file named Clock_QR_Post.
 *      Do not paste it over Code.gs. Do not uncomment anything yet.
 *      Run  clkQrPostInstallCheck()  and read the log.
 *
 *   2. Do what the log tells you — one of two things:
 *        (a) no doPost exists  -> uncomment the doPost block at the bottom
 *        (b) a doPost exists   -> leave the block commented and add two lines
 *                                 to the top of the existing doPost:
 *                                     var qr = clkQrDoPost(e);
 *                                     if (qr) return qr;
 *
 *   3. Deploy > Manage deployments > pencil > Version: New version > Deploy.
 *      Use the PENCIL. "New deployment" mints a new /exec URL and every
 *      printed slip, every tile URL and QR_EXEC_URL on the Config tab would
 *      then point at the old one. That is how this system acquired two
 *      deployment IDs in the first place.
 *
 *   4. Verify with  clkQrPostSelfTest()  before you touch a phone.
 */

var CLK_QR_POST_VERSION = '1.2.0';


/* ==========================================================================
 * 1 · The route
 * ======================================================================== */

/**
 * Returns a ContentService JSON response for POSTs this file owns, or null for
 * anything else so an existing doPost can carry on exactly as before.
 *
 * Null-on-no-match is the same shape clkQrDoGet uses for doGet. It is what
 * makes the two-line hook in step 2(b) safe.
 */
function clkQrDoPost(e) {
  var raw = e && e.postData && e.postData.contents;
  if (!raw) return null;

  var body;
  try {
    body = JSON.parse(raw);
  } catch (err) {
    return null;   // not ours — could be a form post, could be noise
  }

  if (!body || String(body.qp || '') !== 'mclk') return null;

  var out;
  try {
    var action = String(body.action || '');

    if (action === 'hello') {
      out = clkQrHello(String(body.t || ''));

    } else if (action === 'submit') {
      out = clkQrSubmit({
        t: String(body.t || ''),
        s: String(body.s || ''),
        a: String(body.a || '')
      });

    } else {
      out = { ok: false, code: 'BAD_ACTION', error: 'Unknown action.' };
    }
  } catch (err) {
    /* The phone shows r.error verbatim under a "Sorry" heading, so this string
       is read by kitchen staff at 6am, not by a developer. Keep it plain, and
       keep the detail in the Audit tab where it is useful. */
    try { clkQrAudit_('qr.postError', String(err && err.message || err)); } catch (ignored) {}
    out = { ok: false, code: 'SERVER', error: 'The clock hit a problem. Try again, or use the kitchen tablet.' };
  }

  return clkQrJson_(out);
}


/** JSON out. Apps Script sets Access-Control-Allow-Origin on /exec responses
 *  when the deployment is "Anyone", which is how the GitHub Pages page can read
 *  this at all. Verified 19 Aug: a cross-origin GET returns a readable body. */
function clkQrJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ==========================================================================
 * 2 · Install check — RUN THIS FIRST
 * ======================================================================== */

/**
 * Reports whether it is safe to uncomment the doPost block, and whether the
 * functions this file depends on actually exist in this project.
 *
 * Run it BEFORE uncommenting anything. Once a doPost is defined in this file,
 * `typeof doPost` is always 'function' and the check can no longer tell you
 * what was there before.
 */
function clkQrPostInstallCheck() {
  var lines = [];
  lines.push('Clock_QR_Post v' + CLK_QR_POST_VERSION + ' — install check');
  lines.push('');

  var deps = ['clkQrHello', 'clkQrSubmit', 'clkQrAudit_'];
  var missing = [];
  for (var i = 0; i < deps.length; i++) {
    var present = false;
    try { present = (eval('typeof ' + deps[i]) === 'function'); } catch (e) { present = false; }
    lines.push((present ? '  OK      ' : '  MISSING ') + deps[i]);
    if (!present) missing.push(deps[i]);
  }

  if (missing.length) {
    lines.push('');
    lines.push('STOP. ' + missing.join(', ') + ' not found in this project.');
    lines.push('Either Clock_QR.gs is not installed here, or this file has been');
    lines.push('pasted into the wrong project. Check the Files list on the left.');
    Logger.log(lines.join('\n'));
    throw new Error('Clock_QR_Post: dependencies missing — ' + missing.join(', '));
  }

  var hasDoPost = false;
  try { hasDoPost = (typeof doPost === 'function'); } catch (e) { hasDoPost = false; }

  lines.push('');
  if (hasDoPost) {
    lines.push('A doPost ALREADY EXISTS in this project.');
    lines.push('');
    lines.push('Do NOT uncomment the block at the bottom of this file.');
    lines.push('Instead put these two lines at the very top of that doPost:');
    lines.push('');
    lines.push('    var qr = clkQrDoPost(e);   // <-- add');
    lines.push('    if (qr) return qr;         // <-- add');
    lines.push('');
    lines.push('clkQrDoPost returns null for every request that is not');
    lines.push('{qp:"mclk"}, so the existing route is untouched.');
  } else {
    lines.push('No doPost in this project.');
    lines.push('');
    lines.push('Safe to uncomment the doPost block at the bottom of this file.');
  }

  lines.push('');
  lines.push('Then: Deploy > Manage deployments > PENCIL > New version > Deploy.');
  lines.push('Then run clkQrPostSelfTest().');

  var msg = lines.join('\n');
  Logger.log(msg);
  return msg;
}


/* ==========================================================================
 * 3 · Self test — RUN AFTER DEPLOYING
 * ======================================================================== */

/**
 * Drives clkQrDoPost with synthetic requests. Writes nothing: an invalid token
 * is refused by clkQrHello before any Events row is considered.
 *
 * This tests the ROUTE, in the editor. It cannot tell you the deployment was
 * updated — only a real POST from outside can do that, and the last line here
 * tells you how.
 */
function clkQrPostSelfTest() {
  var lines = ['Clock_QR_Post v' + CLK_QR_POST_VERSION + ' — self test', ''];

  function run(label, payload) {
    var res = clkQrDoPost({ postData: { contents: JSON.stringify(payload) } });
    if (res === null) { lines.push('  ' + label + ' -> null (not routed)'); return null; }
    var txt = res.getContent();
    lines.push('  ' + label + ' -> ' + txt.slice(0, 120));
    return txt;
  }

  // Should route and be refused politely, not throw.
  run('hello, invalid token ', { qp: 'mclk', action: 'hello',  t: 'SELFTEST_NOT_A_REAL_TOKEN' });
  run('submit, invalid token', { qp: 'mclk', action: 'submit', t: 'SELFTEST_NOT_A_REAL_TOKEN', s: 'AAAABBBB', a: 'in' });
  run('unknown action       ', { qp: 'mclk', action: 'wat',    t: 'x' });

  // Should NOT route — these must come back null so other routes still work.
  run('wrong qp             ', { qp: 'other', action: 'hello', t: 'x' });
  var noQp = clkQrDoPost({ postData: { contents: '{"action":"hello"}' } });
  lines.push('  no qp                -> ' + (noQp === null ? 'null (correct)' : 'ROUTED — WRONG'));
  var junk = clkQrDoPost({ postData: { contents: 'not json at all' } });
  lines.push('  malformed body       -> ' + (junk === null ? 'null (correct)' : 'ROUTED — WRONG'));
  var empty = clkQrDoPost({});
  lines.push('  empty request        -> ' + (empty === null ? 'null (correct)' : 'ROUTED — WRONG'));

  lines.push('');
  lines.push('The first three should be JSON with "ok":false and a readable message.');
  lines.push('The last four MUST say null. A route that swallows other posts is');
  lines.push('worse than one that never worked.');
  lines.push('');
  lines.push('This proves the code. It does NOT prove the deployment was updated —');
  lines.push('the web app serves the last DEPLOYED version, not the last saved edit.');
  lines.push('Finish by opening clock-mobile.html on a real phone with a real slip.');

  var msg = lines.join('\n');
  Logger.log(msg);
  return msg;
}


/* ==========================================================================
 * 4 · The doPost — UNCOMMENT ONLY IF clkQrPostInstallCheck() SAYS SO
 * ======================================================================== */

/*
function doPost(e) {
  var qr = clkQrDoPost(e);
  if (qr) return qr;

  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'Unknown POST route.' }))
    .setMimeType(ContentService.MimeType.JSON);
}
*/
