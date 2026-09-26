/**
 * The church year — the names a Sunday can carry, and which one a date is.
 *
 * Source of truth is The United Methodist Book of Worship: its Christian Year
 * calendar and its eleven Special Sundays. @decision:gold 2026-09-25 · FEATURE-011
 *
 * ⭐⭐ ONE DATE HAS SEVERAL LEGITIMATE NAMES, AND THEY ARE NOT ALTERNATIVES —
 * they are STYLES, and a congregation picks one and stays there. 2026-10-04 is
 * all of "Nineteenth Sunday after Pentecost", "Proper 22", "Season after
 * Pentecost", "Ordinary Time" and "World Communion Sunday" at once. That is why
 * the caller remembers a style rather than asking every week.
 *
 * ⭐ THE TIERS ARE DERIVED FROM THE BOOK'S OWN STRUCTURE, not assigned by taste:
 * a day the Christian Year section names is tier 1, the Special Sundays list is
 * tier 2, and the counted Sundays plus the season labels are tier 3.
 * Precedence for the autofill button: 1 beats 2 beats 3.
 *
 * ⛔ "Ordinary Time" is NOT a Book of Worship or RCL term — the book says
 * "Season after Pentecost" and "Season after the Epiphany". It is here because
 * it is in wide use and is what this project's own services have carried.
 */

export type Tier = 1 | 2 | 3

/** Tier-3 naming styles. The one the caller remembers, per church. */
export type Style = 'numbered' | 'proper' | 'season' | 'ordinary'

export interface DayName {
  name: string
  tier: Tier
  /** Present only on tier 3 — which style this name belongs to. */
  style?: Style
}

const ORDINAL = [
  '', 'First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Seventh',
  'Eighth', 'Ninth', 'Tenth', 'Eleventh', 'Twelfth', 'Thirteenth',
  'Fourteenth', 'Fifteenth', 'Sixteenth', 'Seventeenth', 'Eighteenth',
  'Nineteenth', 'Twentieth', 'Twenty-first', 'Twenty-second', 'Twenty-third',
  'Twenty-fourth', 'Twenty-fifth', 'Twenty-sixth',
]

/* ── date helpers ────────────────────────────────────────────────────────────
   ⚠ Everything here is built on UTC. A local-midnight Date shifts a day either
   side of the international date line and in some DST transitions, and a church
   calendar that is one day out is silently wrong on exactly the weeks that
   matter most. Parse, compute and format in UTC throughout. */

function utc(y: number, m: number, d: number): Date {
  return new Date(Date.UTC(y, m - 1, d))
}
function addDays(d: Date, n: number): Date {
  return new Date(d.getTime() + n * 86400000)
}
function iso(d: Date): string {
  return d.toISOString().slice(0, 10)
}
export function parseISO(s: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s)
  if (!m) return null
  const d = utc(+m[1], +m[2], +m[3])
  return isNaN(d.getTime()) ? null : d
}
/** The Sunday on or after `d`. */
function sundayOnOrAfter(d: Date): Date {
  return addDays(d, (7 - d.getUTCDay()) % 7)
}
/** The Sunday on or before `d`. */
function sundayOnOrBefore(d: Date): Date {
  return addDays(d, -d.getUTCDay())
}
/** The nth (1-based) given weekday of a month. */
function nthWeekday(y: number, month: number, weekday: number, n: number): Date {
  const first = utc(y, month, 1)
  const offset = (weekday - first.getUTCDay() + 7) % 7
  return addDays(first, offset + (n - 1) * 7)
}
/** The last given weekday of a month. */
function lastWeekday(y: number, month: number, weekday: number): Date {
  const last = utc(y, month + 1, 0)
  return addDays(last, -((last.getUTCDay() - weekday + 7) % 7))
}

/**
 * Easter Day, Gregorian — Meeus/Jones/Butcher.
 *
 * ⭐ Everything movable in the year hangs off this one date: Ash Wednesday is
 * Easter −46, Pentecost is Easter +49, and the whole Season after Pentecost is
 * counted from it. Get this wrong and every movable name is wrong together,
 * which is at least loud rather than subtle.
 */
export function easter(year: number): Date {
  const a = year % 19
  const b = Math.floor(year / 100)
  const c = year % 100
  const d = Math.floor(b / 4)
  const e = b % 4
  const f = Math.floor((b + 8) / 25)
  const g = Math.floor((b - f + 1) / 3)
  const h = (19 * a + b - d - g + 15) % 30
  const i = Math.floor(c / 4)
  const k = c % 4
  const l = (32 + 2 * e + 2 * i - h - k) % 7
  const m = Math.floor((a + 11 * h + 22 * l) / 451)
  const month = Math.floor((h + l - 7 * m + 114) / 31)
  const day = ((h + l - 7 * m + 114) % 31) + 1
  return utc(year, month, day)
}

/** The First Sunday of Advent — four Sundays before Christmas Day. */
export function adventOne(year: number): Date {
  const advent4 = sundayOnOrBefore(utc(year, 12, 24))
  return addDays(advent4, -21)
}

/**
 * The Proper number for a Sunday in the Season after Pentecost.
 *
 * ⭐ WHY PROPERS EXIST, since the number looks arbitrary: Easter moves, so the
 * count of Sundays between Pentecost and Advent changes year to year, and
 * counting forward from Pentecost lands on different readings in different
 * years. A Proper is anchored to a DATE RANGE instead, which keeps a given
 * Sunday's readings stable. Proper 4 is the week of May 29 – June 4, and each
 * one after is the next seven days.
 */
function proper(d: Date): number | null {
  const base = utc(d.getUTCFullYear(), 5, 29)
  const weeks = Math.floor((d.getTime() - base.getTime()) / (7 * 86400000))
  const n = 4 + weeks
  return n >= 4 && n <= 29 ? n : null
}

/* ── the fixed and computed days of one year ─────────────────────────────── */

interface Resolved { [isoDate: string]: DayName[] }

function put(map: Resolved, d: Date, n: DayName) {
  const k = iso(d)
  ;(map[k] ||= []).push(n)
}

/**
 * Every named day of one calendar year, keyed by date.
 *
 * ⚠ A church year straddles January, so the Sundays after Christmas and the
 * Season after the Epiphany belong to the year that FOLLOWS the Advent that
 * began them. This builds one CALENDAR year and the caller asks for the
 * calendar year a date falls in, which keeps the seam out of the lookup.
 */
function buildYear(y: number): Resolved {
  const m: Resolved = {}
  const E = easter(y)

  // Advent → Christmas → Epiphany
  const adv1 = adventOne(y)
  for (let i = 0; i < 4; i++) {
    put(m, addDays(adv1, i * 7), { name: `${ORDINAL[i + 1]} Sunday of Advent`, tier: 1 })
  }
  put(m, utc(y, 12, 24), { name: 'Christmas Eve', tier: 1 })
  put(m, utc(y, 12, 25), { name: 'Christmas Day', tier: 1 })
  put(m, utc(y, 12, 25), { name: 'Nativity of the Lord', tier: 1 })
  put(m, sundayOnOrAfter(utc(y, 12, 26)), { name: 'First Sunday after Christmas Day', tier: 1 })
  put(m, utc(y, 12, 31), { name: "New Year's Eve", tier: 1 })
  put(m, utc(y, 1, 1), { name: "New Year's Day", tier: 1 })
  put(m, utc(y, 1, 6), { name: 'Epiphany of the Lord', tier: 1 })

  // Season after the Epiphany — Baptism of the Lord through Transfiguration
  const baptism = sundayOnOrAfter(utc(y, 1, 7))
  put(m, baptism, { name: 'Baptism of the Lord', tier: 1 })
  const ashWednesday = addDays(E, -46)
  const transfiguration = addDays(ashWednesday, -3)
  for (let n = 2; n <= 8; n++) {
    const d = addDays(baptism, (n - 1) * 7)
    if (d >= transfiguration) break
    put(m, d, { name: `${ORDINAL[n]} Sunday after the Epiphany`, tier: 3, style: 'numbered' })
    put(m, d, { name: 'Season after the Epiphany', tier: 3, style: 'season' })
    put(m, d, { name: 'Ordinary Time', tier: 3, style: 'ordinary' })
  }
  put(m, transfiguration, { name: 'Transfiguration Sunday', tier: 1 })

  // Lent and Holy Week
  put(m, ashWednesday, { name: 'Ash Wednesday', tier: 1 })
  for (let n = 1; n <= 5; n++) {
    const d = addDays(E, -49 + n * 7)
    put(m, d, { name: `${ORDINAL[n]} Sunday in Lent`, tier: 3, style: 'numbered' })
    put(m, d, { name: 'Lent', tier: 3, style: 'season' })
  }
  put(m, addDays(E, -7), { name: 'Passion/Palm Sunday', tier: 1 })
  put(m, addDays(E, -7), { name: 'Palm Sunday', tier: 1 })
  put(m, addDays(E, -7), { name: 'Passion Sunday', tier: 1 })
  put(m, addDays(E, -6), { name: 'Monday of Holy Week', tier: 1 })
  put(m, addDays(E, -5), { name: 'Tuesday of Holy Week', tier: 1 })
  put(m, addDays(E, -4), { name: 'Wednesday of Holy Week', tier: 1 })
  put(m, addDays(E, -3), { name: 'Holy Thursday', tier: 1 })
  put(m, addDays(E, -2), { name: 'Good Friday', tier: 1 })
  put(m, addDays(E, -1), { name: 'Holy Saturday', tier: 1 })
  put(m, addDays(E, -1), { name: 'Easter Eve', tier: 1 })

  // Easter through Pentecost
  put(m, E, { name: 'Easter Day', tier: 1 })
  put(m, E, { name: 'Resurrection of the Lord', tier: 1 })
  put(m, E, { name: 'Easter Evening', tier: 1 })
  for (let n = 2; n <= 7; n++) {
    const d = addDays(E, (n - 1) * 7)
    put(m, d, { name: `${ORDINAL[n]} Sunday of Easter`, tier: 3, style: 'numbered' })
    put(m, d, { name: 'Easter', tier: 3, style: 'season' })
  }
  put(m, addDays(E, 39), { name: 'Ascension of the Lord', tier: 1 })
  put(m, addDays(E, 49), { name: 'Day of Pentecost', tier: 1 })

  // Season after Pentecost — Trinity through Christ the King
  const trinity = addDays(E, 56)
  const christTheKing = addDays(adv1, -7)
  put(m, trinity, { name: 'Trinity Sunday', tier: 1 })
  put(m, christTheKing, { name: 'Christ the King Sunday', tier: 1 })
  put(m, christTheKing, { name: 'Reign of Christ Sunday', tier: 1 })
  for (let d = trinity; d <= christTheKing; d = addDays(d, 7)) {
    const n = Math.round((d.getTime() - trinity.getTime()) / (7 * 86400000)) + 1
    if (n >= 2 && n <= 26) {
      put(m, d, { name: `${ORDINAL[n]} Sunday after Pentecost`, tier: 3, style: 'numbered' })
    }
    // ⚠ TRINITY SUNDAY HAS NO PROPER. It is its own day in the RCL and displaces
    // the Proper whose date range it happens to fall in; the Sunday AFTER it
    // picks that range up. Christ the King is the opposite case — it genuinely
    // is Proper 29 — so this skips the first Sunday of the season and nothing else.
    const p = n === 1 ? null : proper(d)
    if (p) put(m, d, { name: `Proper ${p}`, tier: 3, style: 'proper' })
    put(m, d, { name: 'Season after Pentecost', tier: 3, style: 'season' })
    put(m, d, { name: 'Ordinary Time', tier: 3, style: 'ordinary' })
  }
  put(m, utc(y, 11, 1), { name: 'All Saints', tier: 1 })
  put(m, sundayOnOrAfter(utc(y, 11, 1)), { name: 'All Saints Sunday', tier: 1 })
  put(m, nthWeekday(y, 11, 4, 4), { name: 'Thanksgiving', tier: 1 })

  // The eleven Special Sundays. Three are set by each Annual Conference and so
  // have no computable date — they are in PICKER_LIST but never resolve here.
  put(m, addDays(nthWeekday(y, 1, 1, 3), -1), { name: 'Human Relations Day', tier: 2 })
  put(m, addDays(E, -21), { name: 'One Great Hour of Sharing', tier: 2 })
  put(m, addDays(E, -21), { name: 'UMCOR Sunday', tier: 2 })
  put(m, addDays(E, 14), { name: 'Native American Awareness Sunday', tier: 2 })
  put(m, addDays(E, 14), { name: 'Native American Ministries Sunday', tier: 2 })
  put(m, sundayOnOrAfter(utc(y, 4, 23)), { name: 'Heritage Sunday', tier: 2 })
  put(m, trinity, { name: 'Peace with Justice Sunday', tier: 2 })
  put(m, nthWeekday(y, 10, 0, 1), { name: 'World Communion Sunday', tier: 2 })
  put(m, nthWeekday(y, 10, 0, 3), { name: 'Laity Sunday', tier: 2 })
  put(m, lastWeekday(y, 11, 0), { name: 'United Methodist Student Day', tier: 2 })

  return m
}

const cache = new Map<number, Resolved>()
function year(y: number): Resolved {
  let r = cache.get(y)
  if (!r) { r = buildYear(y); cache.set(y, r) }
  return r
}

/**
 * Every name the given date carries, most significant first.
 * Empty for a date that is not a named day (an ordinary weekday).
 */
export function namesForDate(isoDate: string): DayName[] {
  const d = parseISO(isoDate)
  if (!d) return []
  const y = d.getUTCFullYear()
  // A date in early January can belong to the previous year's Advent cycle, so
  // look in both and merge. Duplicates cannot occur: each name is put once.
  const hits = [...(year(y)[isoDate] || []), ...(year(y - 1)[isoDate] || [])]
  return hits.sort((a, b) => a.tier - b.tier)
}

/**
 * What the autofill button writes for this date.
 *
 * ⭐⭐ HIS PRECEDENCE RULE: "a special sunday takes precidence over a generic
 * '*th week of/after*' but less value than a holy day." So tier 1, then tier 2,
 * then tier 3 in the caller's remembered style. @decision:gold 2026-09-25
 *
 * ⚠ The style only ever decides among TIER 3 names. Taking a holy day or a
 * Special Sunday teaches the caller nothing, which is why a week of Palm Sunday
 * does not wipe out a church's standing preference for "Ordinary Time".
 */
export function autofillFor(isoDate: string, style: Style): string | null {
  const names = namesForDate(isoDate)
  if (names.length === 0) return null
  const holy = names.find(n => n.tier === 1)
  if (holy) return holy.name
  const special = names.find(n => n.tier === 2)
  if (special) return special.name
  const styled = names.find(n => n.tier === 3 && n.style === style)
  if (styled) return styled.name
  return names.find(n => n.tier === 3)?.name ?? null
}

/** The style a chosen name belongs to, or null if choosing it teaches nothing. */
export function styleOf(isoDate: string, name: string): Style | null {
  const hit = namesForDate(isoDate).find(n => n.name === name)
  return hit && hit.tier === 3 ? hit.style ?? null : null
}

/**
 * The whole list, for the dropdown. Order is the church year, not the alphabet,
 * so scrolling it reads like a calendar; the field filters as you type, which is
 * what actually makes a list this long usable.
 */
export const PICKER_LIST: string[] = (() => {
  const L: string[] = []
  const seasonRange = (label: (n: number) => string, from: number, to: number) => {
    for (let n = from; n <= to; n++) L.push(label(n))
  }
  seasonRange(n => `${ORDINAL[n]} Sunday of Advent`, 1, 4)
  L.push('Advent')
  L.push('Christmas Eve', 'Christmas Day', 'Nativity of the Lord',
    'First Sunday after Christmas Day', "New Year's Eve", "New Year's Day",
    'Epiphany of the Lord', 'Christmas')
  L.push('Baptism of the Lord')
  seasonRange(n => `${ORDINAL[n]} Sunday after the Epiphany`, 2, 8)
  L.push('Transfiguration Sunday', 'Season after the Epiphany')
  L.push('Ash Wednesday')
  seasonRange(n => `${ORDINAL[n]} Sunday in Lent`, 1, 5)
  L.push('Lent')
  L.push('Passion/Palm Sunday', 'Palm Sunday', 'Passion Sunday',
    'Monday of Holy Week', 'Tuesday of Holy Week', 'Wednesday of Holy Week',
    'Holy Thursday', 'Good Friday', 'Holy Saturday', 'Easter Eve')
  L.push('Easter Day', 'Resurrection of the Lord', 'Easter Evening')
  seasonRange(n => `${ORDINAL[n]} Sunday of Easter`, 2, 7)
  L.push('Ascension of the Lord', 'Easter', 'Day of Pentecost', 'Trinity Sunday')
  seasonRange(n => `${ORDINAL[n]} Sunday after Pentecost`, 2, 26)
  for (let n = 4; n <= 29; n++) L.push(`Proper ${n}`)
  L.push('All Saints', 'All Saints Sunday', 'Thanksgiving',
    'Christ the King Sunday', 'Reign of Christ Sunday',
    'Season after Pentecost', 'Ordinary Time')
  L.push('Human Relations Day', 'One Great Hour of Sharing', 'UMCOR Sunday',
    'Native American Awareness Sunday', 'Native American Ministries Sunday',
    'Heritage Sunday', 'Peace with Justice Sunday', 'Christian Education Sunday',
    'Golden Cross Sunday', 'Rural Life Sunday', 'World Communion Sunday',
    'Laity Sunday', 'United Methodist Student Day')
  return L
})()
