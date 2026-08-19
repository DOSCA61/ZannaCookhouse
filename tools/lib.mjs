/**
 * lib.mjs — shared helpers for the Zanna clasp tooling.
 * No dependencies. Node 18+ (uses global fetch).
 */
import { readFileSync, existsSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const REGISTRY = join(ROOT, '_gas', 'projects.json');

export const C = {
  red: s => `\x1b[31m${s}\x1b[0m`,
  green: s => `\x1b[32m${s}\x1b[0m`,
  amber: s => `\x1b[33m${s}\x1b[0m`,
  dim: s => `\x1b[2m${s}\x1b[0m`,
  bold: s => `\x1b[1m${s}\x1b[0m`,
};

export function die(msg) {
  console.error(`\n${C.red('REFUSED')}  ${msg}\n`);
  process.exit(1);
}

export function loadRegistry() {
  if (!existsSync(REGISTRY)) die(`registry not found at ${REGISTRY}`);
  let reg;
  try {
    reg = JSON.parse(readFileSync(REGISTRY, 'utf8'));
  } catch (e) {
    die(`registry is not valid JSON — ${e.message}`);
  }
  return reg;
}

export function getProject(reg, key) {
  const p = reg.projects?.[key];
  if (!p) {
    die(`unknown project "${key}". Known: ${Object.keys(reg.projects || {}).join(', ')}`);
  }
  return { key, ...p };
}

/** Read a project's .clasp.json — this is what pins the folder to a script ID. */
export function readClaspJson(projDir) {
  const f = join(ROOT, projDir, '.clasp.json');
  if (!existsSync(f)) {
    die(`no .clasp.json in ${projDir}. Has this project been cloned yet?\n` +
        `          Run:  cd ${projDir} && clasp clone <scriptId>`);
  }
  try {
    return JSON.parse(readFileSync(f, 'utf8'));
  } catch (e) {
    die(`.clasp.json in ${projDir} is not valid JSON — ${e.message}`);
  }
}

export function readManifest(projDir) {
  // clasp puts appsscript.json either at the project root or under rootDir.
  const candidates = [
    join(ROOT, projDir, 'appsscript.json'),
    join(ROOT, projDir, 'src', 'appsscript.json'),
  ];
  for (const f of candidates) if (existsSync(f)) return { path: f, json: JSON.parse(readFileSync(f, 'utf8')) };
  return null;
}

/** Run a command, streaming output. Returns stdout. Throws on non-zero. */
export function run(cmd, args, opts = {}) {
  return execFileSync(cmd, args, {
    stdio: opts.capture ? ['inherit', 'pipe', 'inherit'] : 'inherit',
    encoding: 'utf8',
    cwd: opts.cwd || ROOT,
    shell: process.platform === 'win32',
  });
}

/** Run git quietly. Returns null rather than throwing, and never leaks to stderr —
 *  this tooling must work in a folder that is not a git repo yet. */
function git(args) {
  try {
    return execFileSync('git', args, {
      encoding: 'utf8', cwd: ROOT, stdio: ['ignore', 'pipe', 'ignore'],
    }).trim();
  } catch {
    return null;
  }
}

export function gitDescribe() {
  const sha = git(['rev-parse', '--short', 'HEAD']);
  if (sha === null) return 'nogit';
  return sha + (git(['status', '--porcelain']) ? '+dirty' : '');
}

export function isGitDirty() {
  const st = git(['status', '--porcelain']);
  return st !== null && st.length > 0;
}

/** Today as yyyy-mm-dd in Europe/Dublin, which is the script timezone. */
export function today() {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Europe/Dublin', year: 'numeric', month: '2-digit', day: '2-digit',
  }).format(new Date());
}

export async function probe(url, timeoutMs = 20000) {
  const ac = new AbortController();
  const t = setTimeout(() => ac.abort(), timeoutMs);
  try {
    const res = await fetch(url, { signal: ac.signal, redirect: 'follow' });
    const body = await res.text();
    return { ok: true, status: res.status, body: body.slice(0, 2000) };
  } catch (e) {
    return { ok: false, error: e.name === 'AbortError' ? `no response in ${timeoutMs}ms` : e.message };
  } finally {
    clearTimeout(t);
  }
}
