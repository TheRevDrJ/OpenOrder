/**
 * Is a hymnal configured? — asked once, shared by everyone who needs to know.
 *
 * ⭐⭐ WHY A SHARED CACHE AND NOT A PROP. Seven hymn pickers need this, and threading a
 * boolean through all seven call sites makes the form's markup about configuration
 * rather than about worship. One promise, resolved once, read by whoever asks.
 *
 * ⛔ THE HYMNAL CANNOT SHIP — it is copyrighted, so a fresh install genuinely has none,
 * and that is a normal state rather than an error. What is NOT acceptable is silence:
 * an empty hymn search looks exactly like "no such hymn", so the app read as broken when
 * it was merely unconfigured. (BUG-024.)
 *
 * CALLED BY: HymnPicker (the note in the dropdown) · App (the notice on first load).
 */
import { useEffect, useState } from 'react'

export interface HymnalStatus {
  /** Hymns in the loaded index. 0 means none is configured. */
  count: number
  /** Where the app is currently looking — named so the notice can say it. */
  dir: string
}

let pending: Promise<HymnalStatus> | null = null

/** ⚠ Cached for the life of the page, deliberately: pointing the app at a new hymnal
 *  folder reloads the page (see `handleChangeDir`), so a stale answer cannot outlive the
 *  change that would falsify it. */
export function hymnalStatus(): Promise<HymnalStatus> {
  pending ??= fetch('/api/health')
    .then(r => r.json())
    .then(d => ({ count: d?.hymnal?.count ?? 0, dir: d?.hymnal?.dir ?? '' }))
    // ⛔ A failed health check must not claim the hymnal is missing — that would put a
    // notice in front of someone whose only problem is that the backend is still waking.
    .catch(() => ({ count: -1, dir: '' }))
  return pending
}

/** `null` until the answer arrives, so callers can tell "not yet" from "none". */
export function useHymnalStatus(): HymnalStatus | null {
  const [s, setS] = useState<HymnalStatus | null>(null)
  useEffect(() => { let live = true; hymnalStatus().then(v => live && setS(v)); return () => { live = false } }, [])
  return s
}

/** True only when we KNOW there is no hymnal. ⛔ Not while loading, and not when the
 *  check itself failed (-1). */
export function useHymnalMissing(): boolean {
  const s = useHymnalStatus()
  return s?.count === 0
}
