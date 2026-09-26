// READS ONLY — a build gate. Exit 1 on a hit. Safe to run anytime.
// check_britishisms.mjs — OpenOrder is AMERICAN. Keep it that way.
//
// ⭐ PORTED FROM FirstLight 2026-09-26, dictionary and reasoning intact, on his
// @decision:gold 2026-09-26 — installed after a British spelling reached a doc
// with nothing in place to catch it. ⛔ Keep the two copies in step by hand: a
// word added to one belongs in the other.
//
// ⭐ WHY. "colour", "behaviour", "recognise" and "artefact" get written here
// without anyone noticing, and the drift is consistent enough that trying harder
// is not a fix. A check that lives outside the author is the only kind that
// works; one that lives in somebody's attention does not exist.
//
// AND IT IS NOT COSMETIC HERE EITHER. This is a public repository for American
// churches. Every comment ships with a clone, the CHANGELOG is the first thing a
// stranger reads, and the app's own words are read by a congregation.
//
//   node scripts/check_britishisms.mjs            # check; exit 1 on any hit
//   node scripts/check_britishisms.mjs --warn     # report, always exit 0
//   node scripts/check_britishisms.mjs --shipped  # only player/visitor-facing
//   node scripts/check_britishisms.mjs --list     # print the dictionary, exit
//
// ⛔ IT NEVER REWRITES ANYTHING. It names the file, the line, the word and the
// American spelling, and stops. (M74's palette is why: substituting silently is
// how a hand-painted asset once got silently rewritten. Report; let the author
// change the word.)
//
// ── ESCAPE HATCHES, when a British spelling is CORRECT ─────────────────────
//   1. Put `britpass` anywhere on the line (in a comment) — that line is skipped.
//   2. Add the file to SKIP_FILES below, with a reason.
// A quotation, a proper noun ("Theatre Royal"), or a character who is actually
// British are all legitimate. The point is to make the choice deliberate.
//
// ── WHERE IT RUNS ──────────────────────────────────────────────────────────
// `scripts/build_all.sh`, in the gate block beside typecheck, so a build cannot
// carry one. ⛔ `build.sh` alone does NOT run it, exactly as it does not run the
// typecheck: the gates belong to the EVENT, not to one artifact.
import { readFileSync, existsSync, statSync } from 'node:fs'
import { execFileSync } from 'node:child_process'

// ── THE DICTIONARY ─────────────────────────────────────────────────────────
// Deliberately EXPLICIT rather than clever. A rule like /\w+ise\b/ would catch
// words I never thought of AND flag `exercise`, `otherwise`, `premise`,
// `noise` — and a false positive here does not annoy someone, it ABORTS A
// BUILD. So: precision over recall. A britishism that slips through costs one
// wrong letter; a false positive costs a ship at 11pm. Miss some. Never lie.
// Adding one is a single entry.

const dict = new Map()
const add = (bad, fix) => dict.set(bad.toLowerCase(), fix)

// -our → -or. Stems inflect cleanly (behaviour+al → behavioral,
// neighbour+hood → neighborhood), so these expand mechanically.
const OUR_SUFFIXES = ['', 's', 'ed', 'ing', 'ings', 'ful', 'fully', 'less',
  'able', 'ably', 'ite', 'ites', 'ist', 'ists', 'ly', 'al', 'ally', 'hood',
  'hoods', 'er', 'ers', 'ous', 'ously']
for (const stem of [
  'colour', 'favour', 'behaviour', 'honour', 'humour', 'labour', 'neighbour',
  'flavour', 'savour', 'vapour', 'rumour', 'armour', 'harbour', 'endeavour',
  'splendour', 'valour', 'odour', 'parlour', 'saviour', 'candour', 'clamour',
  'ardour', 'fervour', 'rigour', 'vigour', 'tumour', 'demeanour', 'arbour',
  'succour', 'misdemeanour',
]) {
  const us = stem.replace('our', 'or')
  for (const s of OUR_SUFFIXES) add(stem + s, us + s)
}
// ⚠ `glamour` is deliberately absent — it is standard American too.

// -ise/-isation → -ize/-ization. Stems end in -is, and the empty suffix is
// NEVER generated, because several stems are real English nouns on their own
// (`emphasis`, `analysis`, `synthesis`). That omission is load-bearing.
const ISE_SUFFIXES = ['e', 'es', 'ed', 'ing', 'ation', 'ations', 'er', 'ers', 'able']
for (const stem of [
  'organis', 'realis', 'recognis', 'apologis', 'memoris', 'minimis', 'maximis',
  'capitalis', 'colonis', 'computeris', 'decentralis', 'digitis', 'equalis',
  'fossilis', 'globalis', 'idolis', 'immunis', 'industrialis', 'liberalis',
  'moralis', 'notaris', 'organis', 'oxidis', 'pasteuris', 'pluralis', 'privatis',
  'randomis', 'secularis', 'socialis', 'sterilis', 'subsidis', 'urbanis',
  'optimis', 'normalis', 'initialis', 'serialis', 'visualis', 'summaris',
  'emphasis', 'prioritis', 'customis', 'synchronis', 'standardis', 'categoris',
  'utilis', 'authoris', 'specialis', 'generalis', 'stabilis', 'sanitis',
  'legitimis', 'familiaris', 'modernis', 'popularis', 'publicis', 'rationalis',
  'characteris', 'criticis', 'sympathis', 'theoris', 'itemis', 'penalis',
  'personalis', 'finalis', 'formalis', 'mobilis', 'neutralis', 'polaris',
  'scrutinis', 'sensitis', 'symbolis', 'vandalis', 'jeopardis', 'dramatis',
  'harmonis', 'centralis', 'localis', 'marginalis', 'materialis', 'mechanis',
  'nationalis', 'patronis', 'revolutionis', 'romanticis', 'satiris',
]) {
  const us = stem.slice(0, -2) + 'iz'
  for (const s of ISE_SUFFIXES) add(stem + s, us + s)
}

// -yse → -yze. Explicit, and note what is MISSING: `analyses` and `paralyses`,
// because those are also the plurals of `analysis` / `paralysis` in American
// English. Flagging them would be wrong roughly half the time.
add('analyse', 'analyze'); add('analysed', 'analyzed'); add('analysing', 'analyzing')
add('paralyse', 'paralyze'); add('paralysed', 'paralyzed'); add('paralysing', 'paralyzing')
add('catalyse', 'catalyze'); add('catalysed', 'catalyzed')

// -re → -er.
for (const [b, a] of [
  ['centre', 'center'], ['centres', 'centers'], ['centred', 'centered'],
  ['centring', 'centering'], ['theatre', 'theater'], ['theatres', 'theaters'],
  ['metre', 'meter'], ['metres', 'meters'], ['litre', 'liter'], ['litres', 'liters'],
  ['fibre', 'fiber'], ['fibres', 'fibers'], ['calibre', 'caliber'],
  ['sombre', 'somber'], ['spectre', 'specter'], ['spectres', 'specters'],
  ['lustre', 'luster'], ['meagre', 'meager'], ['louvre', 'louver'],
  ['manoeuvre', 'maneuver'], ['manoeuvres', 'maneuvers'],
  ['manoeuvred', 'maneuvered'], ['manoeuvring', 'maneuvering'],
  ['kilometre', 'kilometer'], ['kilometres', 'kilometers'],
  ['centimetre', 'centimeter'], ['centimetres', 'centimeters'],
  ['millimetre', 'millimeter'], ['millimetres', 'millimeters'],
]) add(b, a)

// Doubled consonants before a suffix (British doubles an unstressed final L).
for (const [b, a] of [
  ['travelled', 'traveled'], ['travelling', 'traveling'], ['traveller', 'traveler'],
  ['travellers', 'travelers'], ['cancelled', 'canceled'], ['cancelling', 'canceling'],
  ['labelled', 'labeled'], ['labelling', 'labeling'], ['modelled', 'modeled'],
  ['modelling', 'modeling'], ['signalled', 'signaled'], ['signalling', 'signaling'],
  ['fuelled', 'fueled'], ['fuelling', 'fueling'], ['levelled', 'leveled'],
  ['levelling', 'leveling'], ['totalled', 'totaled'], ['equalled', 'equaled'],
  ['rivalled', 'rivaled'], ['marvellous', 'marvelous'], ['jewellery', 'jewelry'],
  ['counsellor', 'counselor'], ['counsellors', 'counselors'],
  ['counselling', 'counseling'], ['dialled', 'dialed'], ['dialling', 'dialing'],
]) add(b, a)

// Dropped L where American keeps it doubled.
for (const [b, a] of [
  ['enrol', 'enroll'], ['enrolment', 'enrollment'], ['fulfil', 'fulfill'],
  ['fulfils', 'fulfills'], ['fulfilment', 'fulfillment'], ['instalment', 'installment'],
  ['instalments', 'installments'], ['skilful', 'skillful'], ['wilful', 'willful'],
  ['appal', 'appall'], ['distil', 'distill'], ['enthral', 'enthrall'],
]) add(b, a)

// -ce/-se and the rest of the usual suspects.
for (const [b, a] of [
  ['defence', 'defense'], ['defences', 'defenses'], ['offence', 'offense'],
  ['offences', 'offenses'], ['pretence', 'pretense'], ['licence', 'license'],
  ['licences', 'licenses'], ['practise', 'practice'], ['practised', 'practiced'],
  ['practising', 'practicing'],
  ['programme', 'program'], ['programmes', 'programs'],
  ['catalogue', 'catalog'], ['catalogues', 'catalogs'], ['catalogued', 'cataloged'],
  ['whilst', 'while'], ['amongst', 'among'],
  ['learnt', 'learned'], ['spelt', 'spelled'], ['dreamt', 'dreamed'],
  ['burnt', 'burned'], ['spilt', 'spilled'],
  ['judgement', 'judgment'], ['judgements', 'judgments'],
  ['acknowledgement', 'acknowledgment'], ['acknowledgements', 'acknowledgments'],
  ['artefact', 'artifact'], ['artefacts', 'artifacts'],
  ['orientated', 'oriented'], ['speciality', 'specialty'],
  ['specialities', 'specialties'],
  ['aluminium', 'aluminum'], ['aeroplane', 'airplane'], ['moustache', 'mustache'],
  ['pyjamas', 'pajamas'], ['plough', 'plow'], ['ploughed', 'plowed'],
  ['storey', 'story'], ['storeys', 'stories'], ['kerb', 'curb'],
  ['tyre', 'tire'], ['tyres', 'tires'],
  ['mould', 'mold'], ['moulds', 'molds'], ['moulded', 'molded'],
  ['moulding', 'molding'], ['mouldy', 'moldy'],
  ['smoulder', 'smolder'], ['smouldered', 'smoldered'],
  ['smouldering', 'smoldering'], ['smoulders', 'smolders'],
  ['draught', 'draft'], ['draughts', 'drafts'], ['gaol', 'jail'],
  ['cheque', 'check'], ['cheques', 'checks'],
  ['sceptic', 'skeptic'], ['sceptics', 'skeptics'], ['sceptical', 'skeptical'],
  ['scepticism', 'skepticism'],
  ['sulphur', 'sulfur'], ['cosy', 'cozy'], ['furore', 'furor'],
  ['maths', 'math'], ['towards', 'toward'],
]) add(b, a)

// ⛔⛔ GREY IS DELIBERATELY NOT IN THIS DICTIONARY. It was in the first version
// and it fired thirty-odd times, including inside proper nouns — Grey Goose,
// Earl Grey — where the British spelling is simply the name of the thing.
//
// ⭐ AND THE STRUCTURAL HALF, which is why it is DELETED rather than waived line
// by line: a rule that has to be excused every time it fires is not a rule, it is
// noise — and noise is how a gate gets switched off entirely. The whole family
// went with it (greys / greyed / greying / greyish / greyscale).
// ⛔ Do not add it back. This is a decision, not an oversight.
//
// ⚠ Other judgment calls, so nobody "fixes" these back in either:
//   dialogue / glamour / acknowledgement — all acceptable American, not flagged.
//   `towards` IS flagged, only because his prose is consistently `toward`.
//   `smelt` is absent on purpose: smelting metal is a real word.
//   `leapt` and `knelt` are standard American and are not flagged.

// Two-word phrases need their own pass.
const PHRASES = [[/\bper cent\b/gi, 'percent']]

// ── FILES ──────────────────────────────────────────────────────────────────
// `git ls-files` is the honest set: it is what ships, and it excludes
// node_modules, build output and every artifact for free — no denylist to rot.
const TEXT_EXT = /\.(ts|tsx|js|mjs|cjs|json|md|html|css|sh|ps1|rs|toml|txt|yml|yaml|xml|svg)$/i

// Third-party or generated content we do not author. A britishism here is
// someone else's and there is nothing to fix.
const SKIP_FILES = new Set([
  'frontend/package-lock.json', // npm's, thousands of packages
  // ⭐ AND THIS FILE. It IS the dictionary — every British spelling in the
  // project is deliberately written down here. Without this line the checker
  // reports several hundred hits against itself the moment it is committed,
  // which is both absurd and the kind of noise that gets a gate switched off.
  'scripts/check_britishisms.mjs',
])
const SKIP_DIRS = [
  'node_modules/', 'dist/', 'build/',
  // ⚠ shadcn components are pasted in and then owned, but they are somebody
  // else's prose until we touch them. Fix one when editing it for another
  // reason; do not sweep them to turn a gate green.
  'frontend/src/components/ui/',
  // ⭐ The publisher's own file, byte-for-byte, and reformatting it would break
  // the provenance that makes every verse checkable against bereanbible.com.
  'bible-text/',
]

// ⛔⛔ AND A NAMED FILE NEVER FAILS, BECAUSE SOME OF THEM ARE QUOTATIONS. A history
// file of somebody's own words reports dozens of hits and every one of them is wrong to
// act on: correcting a quotation to satisfy a checker is editing what a person said.
// ⛔ FILES OUTSIDE GIT ARE NOT NAMED HERE. This script ships publicly, so it does
// not carry a list of the working notes. Pass any extra file as an argument to
// have it checked:  node scripts/check_britishisms.mjs NOTES.md
//
// ── THE TWO TIERS, AND WHY THE LINE IS DRAWN HERE ──────────────────────────
// ⭐⭐ EVERY TRACKED FILE IS OUTWARD HERE, and that is the difference from
// FirstLight. This repository is PUBLIC: every comment ships with a clone, and
// the CHANGELOG is among the first things a stranger reads. There is no tier of
// "ours, privately" inside git.
// ⭐ It is affordable because the backlog was small — 33 hits at the port, all
// comment prose, cleaned in one pass. FirstLight could not do this: its first
// run found 337, most of them in documents quoting him. A gate you cannot get to
// green on the day you install it is a gate that gets switched off.
// ⛔ What stays OUT of the gate is the private layer (EXTRA above), for the same
// reason FirstLight keeps most of its own out: it quotes him.
// ⭐ Everything git tracks is outward, because the repository is public. A file
// named on the command line is reported but never fails: it was asked about
// deliberately, and it may well be somebody's quoted words.
const EXTRA = process.argv.slice(2).filter((a) => !a.startsWith('--'))
const isOutward = (f) => !EXTRA.includes(f)

const args = process.argv.slice(2)
const WARN = args.includes('--warn')
const ALL = args.includes('--all')          // fail on every tier, not just OUTWARD
const SHIPPED_ONLY = args.includes('--shipped')

if (args.includes('--list')) {
  for (const [b, a] of [...dict].sort()) console.log(`${b} → ${a}`)
  console.log(`\n${dict.size} entries`)
  process.exit(0)
}

let files = execFileSync('git', ['ls-files'], { encoding: 'utf8' }).split('\n').filter(Boolean)
files = files.concat(EXTRA)
files = files.filter((f) =>
  TEXT_EXT.test(f) && !SKIP_FILES.has(f) && !SKIP_DIRS.some((d) => f.startsWith(d)))
if (SHIPPED_ONLY) files = files.filter(isOutward)

// ── SCAN ───────────────────────────────────────────────────────────────────
// One regex over the whole dictionary — alternation sorted longest-first so
// `neighbourhood` matches before `neighbour` and the report names the real word.
const WORDS = [...dict.keys()].sort((a, b) => b.length - a.length)
const RE = new RegExp(`\\b(${WORDS.join('|')})\\b`, 'gi')

const hits = []
for (const file of files) {
  if (!existsSync(file) || statSync(file).isDirectory()) continue
  let text
  try { text = readFileSync(file, 'utf8') } catch { continue }   // binary; skip
  if (text.includes('\0')) continue
  text.split('\n').forEach((line, i) => {
    if (line.includes('britpass')) return
    for (const m of line.matchAll(RE)) {
      hits.push({ file, line: i + 1, col: m.index + 1, bad: m[0],
        fix: matchCase(m[0], dict.get(m[0].toLowerCase())), text: line.trim() })
    }
    for (const [re, fix] of PHRASES) {
      for (const m of line.matchAll(re)) {
        hits.push({ file, line: i + 1, col: m.index + 1, bad: m[0], fix, text: line.trim() })
      }
    }
  })
}

// Report the fix in the SAME case as the offence, so it can be pasted straight in.
function matchCase(found, fix) {
  if (found === found.toUpperCase() && found.length > 1) return fix.toUpperCase()
  if (found[0] === found[0].toUpperCase()) return fix[0].toUpperCase() + fix.slice(1)
  return fix
}

// ── REPORT ─────────────────────────────────────────────────────────────────
const outward = hits.filter((h) => isOutward(h.file))
const inward = hits.filter((h) => !isOutward(h.file))
const fatal = ALL ? hits : outward

if (!hits.length) {
  console.log(`✅ britishisms: none (${files.length} files, ${dict.size} words)`)
  process.exit(0)
}

// Print the fatal tier in full — those are the ones somebody has to act on.
const show = fatal.length ? fatal : []
if (show.length) {
  console.error(`\n⛔ ${show.length} BRITISHISM(S) IN OUTWARD-FACING TEXT — a stranger reads these\n`)
  let last = null
  for (const h of show.sort((a, b) => a.file.localeCompare(b.file) || a.line - b.line)) {
    if (h.file !== last) { console.error(`  ${h.file}`); last = h.file }
    const snip = h.text.length > 78 ? h.text.slice(0, 75) + '…' : h.text
    console.error(`    ${String(h.line).padStart(5)}:${String(h.col).padEnd(3)} ${h.bad} → ${h.fix}`)
    console.error(`          ${snip}`)
  }
  console.error('\n  Fix the word, or put `britpass` on the line if the British spelling is')
  console.error('  correct (a quotation, a proper noun, a character who is actually British).')
}

// The verdict LEADS. This prints in the middle of a build, where the eye reads
// the first line and moves on — so "clean" must never sit under a list.
if (!show.length) console.error('\n✅ britishisms: outward-facing text is clean.')

// The inward tier is a COUNT, not a wall of text — visible enough to shrink,
// quiet enough that nobody learns to scroll past it.
if (!ALL && inward.length) {
  const byFile = new Map()
  for (const h of inward) byFile.set(h.file, (byFile.get(h.file) ?? 0) + 1)
  const top = [...byFile].sort((a, b) => b[1] - a[1]).slice(0, 5)
  console.error(`\n  ℹ️  ${inward.length} more in comments and docs (not a build gate). Worst:`)
  for (const [f, n] of top) console.error(`        ${String(n).padStart(4)}  ${f}`)
  console.error('      See them all: node scripts/check_britishisms.mjs --all --warn')
}

console.error('')
process.exit(WARN || !fatal.length ? 0 : 1)
