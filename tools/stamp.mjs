#!/usr/bin/env node
/**
 * stamp.mjs — write one build identity into the backend and the front-end together.
 *
 *   node tools/stamp.mjs <project>
 *   node tools/stamp.mjs --all
 *
 * This is the fix for Trap 4 (front-end and backend versioned apart, nothing
 * enforcing the pairing).
 *
 * DESIGN NOTE — why this does not edit your existing code.
 * It would be easy to regex `APP_VERSION = '...'` inside Code.gs and rewrite it.
 * That is exactly the kind of clever step that eventually corrupts a file nobody
 * is watching. Instead:
 *
 *   - For each Apps Script project it writes ONE GENERATED FILE, Build.js. Nothing
 *     else is touched. Apps Script merges every .gs into one namespace, so
 *     buildId_() is callable from Code.gs without an import.
 *   - For each front-end page it replaces ONLY the text between two sentinel
 *     comments. If the sentinels are absent it reports the page as unstamped and
 *     changes nothing — it never guesses where to write.
 *
 * To opt a front-end page in, paste this once, anywhere in the <body>:
 *
 *   <!-- BUILD:START -->pending<!-- BUILD:END -->
 *
 * A good home is the existing footer build stamp, per the 2026-08-18 decision that
 * every front-end build must identify itself.
 */
import { writeFileSync, readFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { ROOT, C, die, loadRegistry, getProject, gitDescribe, today } from './lib.mjs';

const reg = loadRegistry();
const arg = process.argv[2];
if (!arg) die('usage: node tools/stamp.mjs <project|--all>');

const stamp = `${today()}+${gitDescribe()}`;
console.log(`\nbuild stamp: ${C.bold(stamp)}\n`);

const keys = arg === '--all' ? Object.keys(reg.projects) : [arg];

for (const key of keys) {
  const p = getProject(reg, key);
  const out = join(ROOT, p.dir, 'Build.js');

  writeFileSync(out, `/**
 * Build.js — GENERATED FILE. Do not edit by hand.
 * Written by tools/stamp.mjs at release time. Any manual change is overwritten.
 *
 * Project : ${p.label}
 * Stamped : ${stamp}
 */
var BUILD_ID = '${stamp}';

/** The build actually running on the server. Call from a diagnostic or a health route. */
function buildId_() {
  return BUILD_ID;
}

/** Run from the editor to see what is deployed vs what is saved. */
function showBuildId() {
  Logger.log('BUILD_ID = ' + BUILD_ID);
  return BUILD_ID;
}
`, 'utf8');

  console.log(`   ${C.green('wrote')}  ${p.dir}/Build.js`);

  // Front-end pages that declare this project as their backend
  const pages = (reg.frontEnd?.pages || []).filter(pg => pg.backend === key);
  for (const pg of pages) {
    const f = join(ROOT, reg.frontEnd.dir || '.', pg.file);
    if (!existsSync(f)) {
      console.log(`   ${C.dim('skip ')}  ${pg.file} — not in this repo`);
      continue;
    }
    const src = readFileSync(f, 'utf8');
    const re = /(<!--\s*BUILD:START\s*-->)([\s\S]*?)(<!--\s*BUILD:END\s*-->)/;
    if (!re.test(src)) {
      console.log(`   ${C.amber('unstamped')}  ${pg.file} — no BUILD:START/END sentinels, left untouched`);
      continue;
    }
    writeFileSync(f, src.replace(re, `$1${stamp}$3`), 'utf8');
    console.log(`   ${C.green('stamped')}  ${pg.file}  ${C.dim(`(→ ${p.label})`)}`);
  }
}

console.log(`
${C.dim('Both sides now carry the same string. If a page ever shows a stamp the server')}
${C.dim('does not report back, the front-end and backend have drifted — which is the')}
${C.dim('failure Trap 4 describes, made visible instead of silent.')}
`);
