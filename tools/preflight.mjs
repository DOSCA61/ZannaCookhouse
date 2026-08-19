#!/usr/bin/env node
/**
 * preflight.mjs — refuse to push to the wrong place, or to push something broken.
 *
 *   node tools/preflight.mjs <project>      e.g.  node tools/preflight.mjs fairleave
 *   node tools/preflight.mjs --all
 *
 * This is the guard that exists because of the four traps and A8. It checks, in order:
 *
 *   1. The folder is pinned to the script ID the registry says it should be (wrong-project guard)
 *   2. That script ID is not on the archived/backup list (A8 guard)
 *   3. Every required file is present locally  — because `clasp push` makes the REMOTE
 *      match the LOCAL folder, so a file missing here gets deleted THERE. This is the
 *      Clock_QR.gs hazard.
 *   4. appsscript.json timezone is Europe/Dublin (hours-are-an-hour-out guard)
 *   5. appsscript.json web app access settings are reported, so you see them before
 *      you can change them by accident
 *
 * Exit code 0 = safe to push. Non-zero = stop.
 */
import { readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { ROOT, C, die, loadRegistry, getProject, readClaspJson, readManifest, isGitDirty } from './lib.mjs';

const reg = loadRegistry();
const arg = process.argv[2];
if (!arg) die('usage: node tools/preflight.mjs <project|--all>');

const keys = arg === '--all' ? Object.keys(reg.projects) : [arg];
let failures = 0;

for (const key of keys) {
  const p = getProject(reg, key);
  console.log(`\n${C.bold(`── ${p.label}  (${key})`)}`);

  const problems = [];
  const notes = [];
  const warns = [];

  // --- 1 & 2. Identity -------------------------------------------------------
  if (p.scriptId === 'FILL_ME') {
    problems.push('registry scriptId is still FILL_ME — fill it in before pushing');
  } else {
    const clasp = readClaspJson(p.dir);
    if (!clasp.scriptId) {
      problems.push('.clasp.json has no scriptId');
    } else if (clasp.scriptId !== p.scriptId) {
      problems.push(
        `WRONG PROJECT PINNED.\n` +
        `             .clasp.json  → ${clasp.scriptId}\n` +
        `             registry     → ${p.scriptId}\n` +
        `             This folder is pointed at a different Apps Script project than the\n` +
        `             registry expects. Do not push. Re-clone, or correct the registry.`
      );
    } else {
      notes.push(`script ID matches registry (…${p.scriptId.slice(-8)})`);
    }

    const archived = (reg.archivedScriptIds || []).find(a => a.scriptId === clasp?.scriptId);
    if (archived) {
      problems.push(`this is an ARCHIVED/BACKUP project — ${archived.reason}`);
    }
  }

  if (p.deploymentId === 'FILL_ME') {
    warns.push('deploymentId not yet filled in — release --deploy will refuse until it is');
  }

  // --- 3. File manifest ------------------------------------------------------
  const dirAbs = join(ROOT, p.dir);
  if (!existsSync(dirAbs)) {
    problems.push(`project folder ${p.dir} does not exist — not cloned yet`);
  } else {
    const present = new Set(readdirSync(dirAbs));
    const missing = (p.requiredFiles || []).filter(f => !present.has(f));
    if (missing.length) {
      problems.push(
        `MISSING FILES: ${missing.join(', ')}\n` +
        `             clasp push makes the remote match this folder, so pushing now would\n` +
        `             DELETE these from the live project. Re-clone before doing anything.`
      );
    } else if (p.requiredFiles?.length) {
      notes.push(`all ${p.requiredFiles.length} required files present`);
    }

    const extra = [...present].filter(f => /\.(gs|js|html)$/i.test(f) && !(p.requiredFiles || []).includes(f));
    if (extra.length) notes.push(C.dim(`other source files here: ${extra.join(', ')}`));
  }

  // --- 4 & 5. Manifest settings ---------------------------------------------
  const man = readManifest(p.dir);
  if (!man) {
    if (existsSync(dirAbs)) problems.push('appsscript.json not found');
  } else {
    const tz = man.json.timeZone;
    if (tz !== p.expectedTimezone) {
      problems.push(`timezone is "${tz}", expected "${p.expectedTimezone}". Hours will be wrong.`);
    } else {
      notes.push(`timezone ${tz}`);
    }
    const w = man.json.webapp;
    if (w) {
      notes.push(`web app: executeAs=${w.executeAs ?? '?'} access=${w.access ?? '?'}`);
      if (w.access === 'ANYONE_ANONYMOUS' || w.access === 'ANYONE') {
        warns.push('deployed to Anyone — expected for this system, but every endpoint is public. Changing this line changes who can reach it.');
      }
    }
  }

  // --- Report ----------------------------------------------------------------
  for (const n of notes) console.log(`   ${C.green('ok')}  ${n}`);
  for (const w of warns) console.log(`   ${C.amber('->')}  ${w}`);
  for (const pr of problems) console.log(`   ${C.red('!!')}  ${pr}`);
  if (problems.length) failures++;
}

if (isGitDirty()) {
  console.log(`\n${C.amber('note')}  working tree has uncommitted changes. Commit before releasing so the`);
  console.log(`      build stamp points at a commit that actually exists.`);
}

console.log('');
if (failures) {
  console.error(`${C.red(`PREFLIGHT FAILED for ${failures} project(s). Nothing was pushed.`)}\n`);
  process.exit(1);
}
console.log(`${C.green('PREFLIGHT PASSED')} — safe to push.\n`);
