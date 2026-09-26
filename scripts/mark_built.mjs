/**
 * mark_built.mjs — stamp `builtIn` on tracker entries that shipped in this build.
 *
 * WHAT IT DOES: sets `builtIn: <N>` on the entries this build actually carries, so
 * the bench can say which artifact a fix reached. Three conditions, and each one
 * exists to stop a false number being written:
 *
 *   FIXED or VERIFIED   the code has landed. ⚠ VERIFIED WAS EXCLUDED AT FIRST and
 *                       that was wrong: a fix verified on the dev server BEFORE the
 *                       build still ships in it, and excluding it left four entries
 *                       from one day looking permanently unbuilt on the board. The
 *                       thing that actually keeps history out is the next condition.
 *   verify.build != NA  tools, docs, deletions and dev-server faults never ship in
 *                       an artifact, so no build number is true of them.
 *   no builtIn yet      builtIn records the build a fix FIRST shipped in.
 *
 * ⛔ IT DOES NOT BACK-STAMP THINGS AN ARTIFACT CANNOT CARRY, and `verify.build ===
 * 'NA'` is what says so: tools, docs, deletions, dev-server faults and repo-history
 * work are never in a build, so no number is true of them. An early draft ignored that
 * and its own negative test caught it writing a build number onto four such entries.
 * ⭐ A stamp nobody can trust is worse than an empty field, because the field is what
 * the board reports — and the board now reads "not in a build yet" off exactly this.
 *
 * WHY IT EXISTS: entries carried builtIn and the bench rendered it, but nothing wrote
 * it — it had been stamped by hand at lock time, and a practice that lives only in
 * somebody's memory is not a practice: the first build anyone forgets leaves an entry
 * reading a build behind, and nothing reports it. (TOOL-003, 2026-09-25.)
 *
 * ⛔ NEVER FAILS THE BUILD. The artifacts are already cut, signed, installed and
 * fanned out by the time this runs; refusing here would turn a bookkeeping problem
 * into a broken build event. It says what it could not do and exits 0.
 * ⚠ The tracker is gitignored, so there is nothing to commit and a fresh clone
 * legitimately has no docs/bugs.json.
 *
 * CALLED BY: scripts/build_all.sh, step 4 (lock), after the number is committed.
 * Run by hand as: node scripts/mark_built.mjs <build-number>
 */

import { readFileSync, writeFileSync } from 'node:fs'

const TRACKER = 'docs/bugs.json'

/** Does this build carry a number that is TRUE of this entry? */
const carriedByThisBuild = e =>
  ['FIXED', 'VERIFIED'].includes(e.status) && e.verify?.build !== 'NA' && !e.builtIn

const n = Number(process.argv[2])
if (!Number.isInteger(n) || n <= 0) {
  console.log(`[tracker] no valid build number given (${process.argv[2]}) — nothing stamped.`)
  process.exit(0)
}

let raw
try {
  raw = readFileSync(TRACKER, 'utf8')
} catch {
  console.log('[tracker] no docs/bugs.json here — nothing to stamp.')
  process.exit(0)
}

try {
  const t = JSON.parse(raw)
  const due = t.entries.filter(carriedByThisBuild)
  if (!due.length) {
    console.log(`[tracker] build ${n}: nothing awaiting a stamp.`)
    process.exit(0)
  }
  for (const e of due) e.builtIn = n

  writeFileSync(TRACKER, JSON.stringify(t, null, 2) + '\n')

  // Read the receipt off disk, not off the object just written.
  const after = JSON.parse(readFileSync(TRACKER, 'utf8'))
  const missed = due.filter(e => after.entries.find(x => x.id === e.id)?.builtIn !== n)
  if (missed.length) {
    console.log(`[tracker] ⚠ stamp did not land for: ${missed.map(e => e.id).join(', ')}`)
    process.exit(0)
  }
  console.log(`[tracker] build ${n} stamped on ${due.length}: ${due.map(e => e.id).join(', ')}`)
} catch (err) {
  // A malformed tracker is a real problem, but not this script's to resolve, and
  // never a reason to fail a build whose artifacts are already out the door.
  console.log(`[tracker] ⚠ could not stamp build ${n}: ${err.message}`)
}
