#!/usr/bin/env node
/**
 * release.mjs — the only way code should ever reach a live Apps Script deployment.
 *
 *   node tools/release.mjs <project> --push            push source only, do NOT deploy
 *   node tools/release.mjs <project> --deploy          push, version, update the EXISTING deployment
 *   node tools/release.mjs <project> --deploy --dry    print every command, run none
 *
 * ── WRITTEN FOR CLASP 3.x ────────────────────────────────────────────────────
 * clasp 3 renamed the deployment commands, and the rename matters more than it
 * looks:
 *
 *     clasp deploy            is now an alias of  create-deployment
 *     clasp redeploy          is now an alias of  update-deployment <deploymentId>
 *
 * So the command that mints a NEW deployment and orphans the old URL — Trap 1,
 * the one that had index.html returning 404 for days — is the one with the
 * friendly, obvious name. The safe operation is the one you have to know to ask
 * for. This script only ever calls update-deployment, against an ID from the
 * registry. `create-deployment` is not reachable from here.
 *
 * ── WHY PUSH AND DEPLOY ARE SEPARATE ─────────────────────────────────────────
 * The web app serves the last DEPLOYED version, not the last saved one. That is
 * Trap 2, and normally it is a nuisance — here it is a safety net. `--push` puts
 * code on the server WITHOUT changing what any member of staff sees, so you can
 * open the editor and look first. Nobody is affected until `--deploy`.
 *
 * ── WHY THERE IS NO --force ANYWHERE ─────────────────────────────────────────
 * `clasp push -f` force-overwrites the remote MANIFEST — appsscript.json, which
 * carries the web app's executeAs and access settings. Overwrite that by accident
 * and you can silently change who is able to reach every endpoint. If clasp asks
 * about the manifest, that is a question worth reading, not suppressing.
 */
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { ROOT, C, die, loadRegistry, getProject, probe } from './lib.mjs';

const [, , key, ...flags] = process.argv;
if (!key || key.startsWith('--')) die('usage: node tools/release.mjs <project> --push|--deploy [--dry]');

const doDeploy = flags.includes('--deploy');
const doPush = flags.includes('--push') || doDeploy;
const dry = flags.includes('--dry');
if (!doPush) die('specify --push or --deploy');

const reg = loadRegistry();
const p = getProject(reg, key);
const cwd = join(ROOT, p.dir);
const WIN = process.platform === 'win32';

function step(label) { console.log(`\n${C.bold('▶ ' + label)}`); }

function sh(cmd, args, { capture = false, inProject = false } = {}) {
  console.log(C.dim(`  $ ${cmd} ${args.join(' ')}${inProject ? `   (in ${p.dir})` : ''}`));
  if (dry) return '';
  return execFileSync(cmd, args, {
    stdio: capture ? ['inherit', 'pipe', 'inherit'] : 'inherit',
    encoding: 'utf8',
    cwd: inProject ? cwd : ROOT,
    shell: WIN,
  });
}

// 0 ── Fail fast, BEFORE anything is pushed. Being refused after a push would
//      leave you half-done and unsure whether the live project changed.
if (doDeploy && p.deploymentId === 'FILL_ME') {
  die(`no deploymentId in the registry for "${key}".\n` +
      `          Get it with:  cd ${p.dir} && clasp list-deployments\n` +
      `          Then put it in _gas/projects.json.\n\n` +
      `          Refusing to deploy without it. In clasp 3, deploying without a\n` +
      `          deployment ID means create-deployment, which mints a NEW URL and\n` +
      `          orphans every device link.\n\n` +
      `          Nothing has been pushed.`);
}

// 1 ── Preflight
step('Preflight');
try {
  execFileSync(process.execPath, [join(ROOT, 'tools', 'preflight.mjs'), key], { stdio: 'inherit', cwd: ROOT });
} catch {
  die('preflight failed — nothing was pushed.');
}

// 2 ── Stamp both sides with one identity
step('Stamp build identity');
sh(process.execPath, [join(ROOT, 'tools', 'stamp.mjs'), key]);

// 3 ── Show exactly what a push would send. This is the Clock_QR.gs check: if a
//      file you expect is absent from this list, STOP.
step('What will be pushed');
sh('clasp', ['status'], { inProject: true });
console.log(C.amber('  Read that list. Every file you expect should be on it.'));

// 4 ── Push. Saved code only — staff still see the deployed version.
step('Push source to the Apps Script project');
console.log(C.dim('  (changes SAVED code only — no member of staff sees anything yet)'));
sh('clasp', ['push'], { inProject: true });

if (!doDeploy) {
  console.log(`\n${C.green('PUSHED, NOT DEPLOYED.')}`);
  console.log(`Open the editor and check it looks right. When you are happy:\n`);
  console.log(`   node tools/release.mjs ${key} --deploy\n`);
  process.exit(0);
}

// 5 ── Create an immutable version. This is your rollback point, and it is why
//      we do not let update-deployment pick a version implicitly.
step('Create an immutable version');
const desc = `release ${new Date().toISOString().slice(0, 16).replace('T', ' ')}`;
let versionNumber = null;
if (!dry) {
  const out = sh('clasp', ['create-version', desc], { capture: true, inProject: true }) || '';
  process.stdout.write(out);
  const m = out.match(/\b(\d+)\b(?![\s\S]*\b\d+\b)/) || out.match(/version\s+(\d+)/i);
  versionNumber = m ? m[1] : null;
  if (!versionNumber) {
    die(`could not read the new version number from clasp's output above.\n` +
        `          Nothing has been deployed — the push succeeded, the deployment is unchanged.\n` +
        `          Find the number with:  cd ${p.dir} && clasp list-versions\n` +
        `          Then finish by hand:\n` +
        `            cd ${p.dir} && clasp update-deployment ${p.deploymentId} -V <version> -d "${desc}"`);
  }
  console.log(`  ${C.green('version ' + versionNumber + ' created')}`);
} else {
  sh('clasp', ['create-version', desc], { inProject: true });
  versionNumber = '<new>';
}

// 6 ── Point the EXISTING deployment at that version. Same ID, same URL.
step(`Update EXISTING deployment ${p.deploymentId.slice(0, 12)}… → version ${versionNumber}`);
sh('clasp', ['update-deployment', p.deploymentId, '-V', String(versionNumber), '-d', desc], { inProject: true });

// 7 ── Verify from outside. What the editor thinks is irrelevant; what the public
//      URL answers is the truth.
step('Verify the live endpoint');
if (dry || p.execUrl === 'FILL_ME') {
  console.log(C.amber('  skipped — no execUrl in the registry (or --dry).'));
} else {
  const res = await probe(p.execUrl);
  if (!res.ok) {
    console.log(`  ${C.red('no response')} — ${res.error}`);
    console.log(C.amber(`  A deploy just happened and the endpoint is not answering.`));
    console.log(C.amber(`  Roll back:  cd ${p.dir} && clasp update-deployment ${p.deploymentId} -V <previous> -d "rollback"`));
    process.exit(2);
  }
  const want = p.versionProbe?.expectContains;
  const hit = want ? res.body.includes(want) : true;
  console.log(`  HTTP ${res.status}`);
  console.log(C.dim(`  ${res.body.slice(0, 300).replace(/\s+/g, ' ')}`));
  console.log(hit ? `  ${C.green('endpoint responded and looks like the right service')}`
                  : `  ${C.red(`expected to see ${want} in the response — it is not there`)}`);
  if (!hit) {
    console.log(C.amber(`  Roll back:  cd ${p.dir} && clasp update-deployment ${p.deploymentId} -V <previous> -d "rollback"`));
    process.exit(2);
  }
}

console.log(`\n${C.green('RELEASED.')}  Same deployment ID, same URL, no device link broken.`);
console.log(C.dim(`Rollback if needed:  cd ${p.dir} && clasp update-deployment ${p.deploymentId} -V <previous> -d "rollback"\n`));
