#!/usr/bin/env node
/**
 * check-exposure.mjs — is your server-side source now readable by the whole internet?
 *
 *   node tools/check-exposure.mjs
 *
 * READ THIS BEFORE YOUR FIRST COMMIT.
 *
 * Putting the Apps Script source in the same repository as the GitHub Pages site
 * is what makes one-commit-changes-both-sides possible, and that is the point of
 * the whole exercise. It also creates a hazard that did not exist before: GitHub
 * Pages serves the repository, so `_gas/fairleave/Code.gs` could become a public
 * URL that anyone can read.
 *
 * There are no secrets in that source — the 2026-08-15 decision moved every
 * credential into Script Properties, which is exactly why that decision matters
 * here. But the source still describes your auth logic, your property names and
 * your action routes. Publishing it hands an attacker the map. Given a public
 * "Anyone" endpoint, do not accept that.
 *
 * TWO DEFENCES, BELT AND BRACES:
 *   1. The folder is named `_gas` with a leading underscore. Jekyll — which is what
 *      GitHub Pages runs by default — does not publish underscore-prefixed
 *      directories.
 *   2. `_config.yml` lists `_gas` and `tools` under `exclude`.
 *
 * Neither is worth anything unverified, and defence 1 silently stops applying if
 * the site is ever switched to a plain static build or a custom Actions workflow.
 * So this script asks the live site, which is the only answer that counts.
 *
 * Run it after your first push to GitHub, and any time the Pages settings change.
 */
import { C, loadRegistry, probe } from './lib.mjs';

const BASE = process.argv[2] || 'https://dosca61.github.io/ZannaCookhouse';
const reg = loadRegistry();

const targets = [];
for (const [key, p] of Object.entries(reg.projects)) {
  for (const f of p.requiredFiles || []) targets.push(`${p.dir}/${f}`);
  targets.push(`${p.dir}/.clasp.json`);
}
targets.push('tools/release.mjs', '_gas/projects.json');

console.log(`\nChecking whether server-side files are publicly served from:\n  ${BASE}\n`);

let exposed = 0;
for (const t of targets) {
  const url = `${BASE}/${t}`;
  const res = await probe(url, 15000);
  if (!res.ok) {
    console.log(`  ${C.amber('?')}  ${t}  ${C.dim('(no response — ' + res.error + ')')}`);
    continue;
  }
  if (res.status === 404) {
    console.log(`  ${C.green('404')}  ${t}  ${C.dim('not published — good')}`);
  } else if (res.status === 200) {
    console.log(`  ${C.red('200')}  ${t}  ${C.red('PUBLICLY READABLE')}`);
    exposed++;
  } else {
    console.log(`  ${C.amber(String(res.status))}  ${t}`);
  }
}

console.log('');
if (exposed) {
  console.log(C.red(`${exposed} file(s) are being served publicly.`));
  console.log(`
Fix it before going further. In order of preference:

  1. Repository Settings → Pages → set the source to a ${C.bold('/docs')} folder or a
     ${C.bold('gh-pages')} branch, and keep the source on main only. This is the clean
     separation and it cannot silently stop working.

  2. Confirm the site is building with Jekyll. If a ${C.bold('.nojekyll')} file exists in
     the repo root, delete it — that file is what turns off the underscore rule
     and the exclude list along with it.

  3. Failing both, move the Apps Script source to a separate PRIVATE repository.
     You lose the single-commit pairing, which is a real loss, but exposed server
     source is a worse trade.
`);
  process.exit(1);
}
console.log(C.green('Nothing server-side is publicly readable. Safe to carry on.\n'));
