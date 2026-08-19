# Zanna Cookhouse — clasp + git migration runbook

**Written 2026-08-19.** Implements action 7 of `Decision_2026-08-19_Platform_Migration.md`
(Option 1: stay on Apps Script, fix the tooling).

**What this achieves:** the source of both Apps Script projects lives in git, and
code reaches a live deployment by exactly one route that cannot orphan a URL.
Traps 1 and 2 close. Trap 4 becomes visible. Trap 3 becomes greppable.

**What it does not achieve, stated plainly:** it does not fix the single-namespace
problem, and it does not remove secrets from spreadsheet cells. Those are accepted
residual risk under the decision record, not silently solved here.

---

## Before you start

**Nothing in Parts 1–3 touches the live system.** Cloning reads. Git is local. The
first command that writes to a live project is in Part 5, and the first command
that changes what staff see is in Part 6. The two are deliberately separate.

**Do Part 0 first. It is not optional.** Cloning the wrong project and later
pushing from it would overwrite live code with a stale copy — the one genuinely
destructive outcome available in this whole procedure.

**Do not copy the workbooks.** Copying a spreadsheet copies its bound script, which
is exactly how the duplicate projects in A8 came to exist. Git is the backup here.

---

## Part 0 — Kill the duplicate-project trap (A8)

There are two identically-named projects for each system. Rename the backups so
they cannot be picked by accident.

1. Open <https://script.google.com/home/my>
2. Find the two **Zanna Holiday (Fair Leave) Management** entries and the two
   **Zanna Clock** entries
3. Identify the live one properly — **open the live spreadsheet → Extensions →
   Apps Script**, and read the script ID out of the URL. Do not identify it by
   modified date from the list; that is a guess.
4. Rename each backup to `ARCHIVE — pre-golive <date>`
5. Record both live script IDs. You need them in Part 2.

Renaming is cosmetic and safe: deployments are addressed by deployment ID, so no
URL changes and nothing breaks.

---

## Part 1 — Local setup *(your machine, your terminal)*

I cannot run these. `clasp login` is a Google OAuth flow against your account, and
the sandbox I can reach on your machine has no network access.

```bash
node --version          # need 18 or newer
npm install -g @google/clasp
clasp login             # opens a browser; approve as admin@richmondit.ie
```

Then enable the API that clasp needs — one switch, once per Google account:

**<https://script.google.com/home/usersettings>** → **Google Apps Script API** → **ON**

---

## Part 2 — Clone both projects *(read-only against the live projects)*

From the root of your `ZannaCookhouse` repo working copy:

```bash
mkdir -p _gas/zanna-clock _gas/fairleave

cd _gas/fairleave
clasp clone 1TNaZ4PgAiuxJxPzOHthh9NVBhOTs4Tduv4z4bFLzLSn6zvgFMANVTbHc
cd ../..

cd _gas/zanna-clock
clasp clone 1aJoYlWavn2NS6adLfImGRPvl61FibCiTQcV4jUdkhcsUExQ_v2OwyVNY
cd ../..
```

**Then verify the clone, before anything else.** Open each project in the editor
and compare its Files list against what landed on disk. They must match exactly.

This matters more than it looks: `clasp push` makes the remote match your folder,
so a file that did not come down would be **deleted** from the live project on the
first push. `Clock_QR.gs` is a separate file precisely so a full-file paste cannot
lose it — this is the one operation that could undo that protection.

---

## Part 3 — Drop in the kit and commit

Copy from this kit into the repo root, preserving paths:

```
_config.yml
.gitignore
CLASP_RUNBOOK.md
_gas/projects.json
tools/lib.mjs
tools/preflight.mjs
tools/stamp.mjs
tools/release.mjs
tools/check-exposure.mjs
```

Fill in `_gas/projects.json` — every `FILL_ME`. Get the deployment IDs with:

```bash
cd _gas/fairleave && clasp list-deployments && cd ../..
cd _gas/zanna-clock && clasp list-deployments && cd ../..
```

You want the deployment currently serving `/exec` — the one whose ID matches the
URL your front-end files already use. Put the `/exec` URLs in `execUrl` too, so
every release is checked against the real endpoint.

Then:

```bash
git add -A
git commit -m "Import Apps Script source for Zanna Clock and FairLeave; add clasp release tooling"
git push
```

**That commit is the first real baseline this system has ever had.**

### Immediately after pushing to GitHub — check nothing leaked

```bash
node tools/check-exposure.mjs
```

GitHub Pages serves the repository, so the Apps Script source could become
publicly readable. Two defences are in place (`_gas` is underscore-prefixed, and
`_config.yml` excludes it), but both are worthless unverified — and defence one
stops applying silently if the site is ever switched away from Jekyll.

There are no credentials in that source; the 2026-08-15 decision moved every one
into Script Properties, which is precisely why that decision earns its keep today.
But the source still describes your auth logic and action routes, and you have a
public "Anyone" endpoint. Do not publish the map.

**Expect every line to read `404`.** If any reads `200`, stop and follow the
remedies the script prints.

---

## Part 4 — Prove the guard works *(still nothing live)*

```bash
node tools/preflight.mjs --all
```

Expect passes on script ID, file manifest and `Europe/Dublin` timezone.

Now prove it actually refuses. Temporarily edit `_gas/fairleave/.clasp.json` and
change one character of the scriptId:

```bash
node tools/preflight.mjs fairleave      # must REFUSE with "WRONG PROJECT PINNED"
```

Put it back. A guard you have never seen fire is a guard you do not know you have.

---

## Part 5 — First push *(writes to the live project — but changes nothing for staff)*

```bash
node tools/release.mjs fairleave --push
```

This pushes **saved** code only. The web app keeps serving the last *deployed*
version, so no member of staff is affected. That is Trap 2 working in your favour
for once, and it is the reason push and deploy are separate commands.

If clasp reports that it will delete a remote file — **stop.** Something did not
come down in the clone. Do not confirm, and do not reach for `--force`.

Open the editor and confirm the code looks right and `Build.gs` has appeared.

---

## Part 6 — First deploy *(this is the live moment)*

**Make the first one a no-op.** Do not combine proving the pipeline with shipping
a change. Push identical code, deploy it, confirm the endpoint still answers
exactly as before.

**Timing:** away from a shift boundary, away from a payroll run.

```bash
node tools/release.mjs fairleave --deploy
```

It preflights, stamps, pushes, deploys **to the existing deployment ID**, then
fetches the live `/exec` and checks the response. Same deployment, same URL, no
device link broken.

Then repeat for `zanna-clock`.

---

## Part 7 — From then on

| Task | Command |
|---|---|
| Check before touching anything | `node tools/preflight.mjs --all` |
| Push code without affecting staff | `node tools/release.mjs <project> --push` |
| Ship it | `node tools/release.mjs <project> --deploy` |
| See what would happen | `node tools/release.mjs <project> --deploy --dry` |
| Confirm source is still not public | `node tools/check-exposure.mjs` |

**Never run bare `clasp deploy`.** In clasp 3 that is an alias of
**`create-deployment`** — it mints a *new* deployment and orphans the old one,
which keeps answering for a while and then 404s everything. That is Trap 1, and
the footgun now has the friendliest name in the command list.

The safe command is **`clasp update-deployment <deploymentId> -V <version>`**
(alias `redeploy`). `release.mjs` only ever calls that one, against an ID from the
registry, so `create-deployment` is not reachable from the tooling.

**Never use `clasp push -f`.** The `-f` flag force-overwrites the remote
*manifest* — `appsscript.json`, which holds the web app's `executeAs` and `access`
settings. Overwrite it by accident and you can silently change who can reach every
endpoint. If clasp asks you about the manifest, read the question.

### Rolling back

Every release creates a numbered, immutable version first, so rollback is one
command:

```bash
cd _gas/<project>
clasp list-versions
clasp update-deployment <deploymentId> -V <previous-version> -d "rollback"
```

### Turning on the front-end half of the build stamp

`stamp.mjs` will not guess where to write in your HTML. Paste this once into each
page's existing footer stamp:

```html
<!-- BUILD:START -->pending<!-- BUILD:END -->
```

From then on the page and its backend carry the same string, and a mismatch tells
you they have drifted — Trap 4 made visible instead of silent.

---

## What to do when it goes wrong

| Symptom | Cause |
|---|---|
| `User has not enabled the Apps Script API` | Part 1, the usersettings switch |
| Preflight: `WRONG PROJECT PINNED` | Working folder points at a different project. Do not push |
| Preflight: `this is an ARCHIVED/BACKUP project` | The A8 guard fired. Re-clone from the live ID |
| Preflight: `MISSING FILES` | Clone is incomplete. Pushing would delete them remotely |
| clasp offers to delete remote files | Same cause. Stop, re-clone, do not `--force` |
| Deployed but behaviour unchanged | Check `showBuildId()` in the editor against the page footer |
| `check-exposure` returns `200` | Move Pages to `/docs` or a `gh-pages` branch |

**Recovering a bad push.** The source is in git, so `git checkout` the last good
commit and push again. And because a push does not deploy, staff were never
affected — provided you have not deployed yet. If you already deployed, the Apps
Script editor keeps a version history you can redeploy from.

---

## Where this leaves the traps

| Trap | Status |
|---|---|
| 1 — Two deployment IDs | **Closed.** `release.mjs` cannot deploy without `-i` |
| 2 — Saving is not deploying | **Closed**, and turned into a safety gap you can use |
| 3 — One namespace across `.gs` files | **Not fixed.** A platform property. But now greppable in git |
| 4 — Front-end / backend drift | **Visible.** One stamp on both sides, once the sentinels are in |
| A8 — Duplicate projects | **Closed** by Part 0 plus the preflight archive list |
