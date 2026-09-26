export interface HymnRef {
  number: string
  title: string
  source: string
}

export interface HymnSearchResult {
  number: string
  title: string
  source: string
  slide_count: number
  file: string
}

export interface OrderOfWorship {
  date: string
  serviceTitle: string
  heroImageFilename: string | null


  praiseHymn1: HymnRef | null
  praiseHymn2: HymnRef | null
  doxology: HymnRef | null
  creed: HymnRef | null
  prayerHymn: HymnRef | null
  liturgicalPrayer: HymnRef | null
  closingHymn: HymnRef | null

  communion: boolean
  scripture: string
  scriptureTranslation: string
  sermonTitle: string
  sermonSubtitle: string
  speakerShortName: string
  offertoryNote: string
}

/** A church profile's people defaults, as /api/churches/{id} returns them. */
export interface ChurchDefaults {
  speakerShort: string
}

export interface Church {
  id: string
  name: string
}

/**
 * A blank service.
 *
 * ⛔ NO NAMES ARE BAKED IN HERE. Who preaches and who leads worship differs per
 * congregation, so the values arrive from the selected church profile. This
 * function used to hardcode two real people, which also meant shipping them in a
 * public repo. @decision:gold 2026-09-24
 */
/** Has anybody actually put anything in this service?
 *
 * ⛔⛔ THE POINT IS TO STOP AN EMPTY FORM OVERWRITING A SAVED SERVICE. @decision:gold
 * 2026-09-26 · BUG-022. A boundary save writes whatever is in the form, so opening a
 * past service or switching church before the boot load has resolved wrote
 * `emptyOrder` over a real service — and it cost his 2026-09-27 bulletin.
 *
 * ⚠ DEFAULTS ARE NOT CONTENT, and that is the whole subtlety: `scriptureTranslation`
 * starts at BSB, `liturgicalPrayer` at 895, and `speakerShortName` is prefilled from
 * the church profile. A form holding only those has had nothing typed into it.
 * ⛔ Do not add a field here without asking whether the app fills it on its own.
 */
export function hasContent(o: OrderOfWorship): boolean {
  const text = [o.serviceTitle, o.scripture, o.sermonTitle, o.sermonSubtitle, o.offertoryNote]
  if (text.some(v => (v ?? '').trim() !== '')) return true
  const picks = [o.praiseHymn1, o.praiseHymn2, o.doxology, o.creed, o.prayerHymn, o.closingHymn]
  if (picks.some(v => v != null)) return true
  if (o.heroImageFilename) return true
  if (o.communion) return true
  // The Lord's Prayer is defaulted, so only a CHANGE to it counts.
  if (o.liturgicalPrayer && o.liturgicalPrayer.number !== '895') return true
  return false
}

/** A stable string over the fields a PERSON filled in, for "are these the same
 *  service?". ⚠ Deliberately excludes `date` and `speakerShortName`: the date is the
 *  same by definition on a church switch, and the speaker is a per-church default, so
 *  including it would make every pair look different and ask a question with no answer. */
export function serviceFingerprint(o: OrderOfWorship): string {
  return JSON.stringify([
    (o.serviceTitle ?? '').trim(), (o.scripture ?? '').trim(),
    o.scriptureTranslation ?? '', (o.sermonTitle ?? '').trim(),
    (o.sermonSubtitle ?? '').trim(), (o.offertoryNote ?? '').trim(),
    o.praiseHymn1, o.praiseHymn2, o.doxology, o.creed, o.prayerHymn, o.closingHymn,
    o.liturgicalPrayer, !!o.communion,
  ])
}

export function emptyOrder(date: string, defaults?: ChurchDefaults): OrderOfWorship {
  return {
    date,
    serviceTitle: '',
    heroImageFilename: null,
    praiseHymn1: null,
    praiseHymn2: null,
    doxology: null,
    creed: null,
    prayerHymn: null,
    liturgicalPrayer: { number: '895', title: "The Lord's Prayer Former Methodist Text", source: 'umh-services' },
    closingHymn: null,
    communion: false,
    scripture: '',
    scriptureTranslation: 'BSB',
    sermonTitle: '',
    sermonSubtitle: '',
    speakerShortName: defaults?.speakerShort ?? '',
    offertoryNote: '',
  }
}
