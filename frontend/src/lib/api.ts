import type { Church, ChurchDefaults, HymnSearchResult, OrderOfWorship } from '@/types'

const BASE = '/api'

export async function searchHymns(query: string): Promise<HymnSearchResult[]> {
  if (!query.trim()) return []
  const res = await fetch(`${BASE}/hymnal/search?q=${encodeURIComponent(query)}`)
  return res.json()
}

export async function getHealth(): Promise<{ status: string; nextSunday: string }> {
  const res = await fetch(`${BASE}/health`)
  return res.json()
}

/** ⭐ A SERVICE IS KEYED ON DATE AND CHURCH (FEATURE-008), so the church travels with
 *  every one of these the way it already travels with generate. */
const withChurch = (church: string, extra?: Record<string, string>) => {
  const q = new URLSearchParams(extra ?? {})
  if (church) q.set('church', church)
  const s = q.toString()
  return s ? `?${s}` : ''
}

export async function listServices(
  church: string,
): Promise<{ date: string; filename: string; own: boolean }[]> {
  const res = await fetch(`${BASE}/services${withChurch(church)}`)
  return res.json()
}

export async function loadService(date: string, church: string): Promise<OrderOfWorship> {
  const res = await fetch(`${BASE}/services/${date}${withChurch(church)}`)
  if (!res.ok) throw new Error('Service not found')
  return res.json()
}

/** `snapshotReason` asks the server to copy the CURRENT saved state aside first —
 *  passed only at the boundaries where work can be lost, never on autosave. */
export async function saveService(
  data: OrderOfWorship, church: string, snapshotReason?: string,
  /** ⭐ Say so when the emptiness is DELIBERATE. The server refuses to write an empty
   *  service over a saved one otherwise (BUG-022) — that guard is what protects the
   *  file from every other writer, not just this form. */
  allowEmpty?: boolean,
): Promise<void> {
  const extra: Record<string, string> = {}
  if (snapshotReason) extra.snapshot_reason = snapshotReason
  if (allowEmpty) extra.allow_empty = 'true'
  const res = await fetch(`${BASE}/services/${data.date}${withChurch(church, Object.keys(extra).length ? extra : undefined)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  })
  // ⛔⛔ A FAILED SAVE MUST THROW. `fetch` resolves on a 500, so without this the
  // status said "Saved" over a write that never happened and the form could then be
  // cleared on the strength of it — the file locked by another program is exactly
  // the case that produces it. @decision:gold 2026-09-25
  if (!res.ok) {
    const detail = await res.json().catch(() => null)
    throw new Error(detail?.detail ?? `Could not save (${res.status})`)
  }
}

export interface Revision { id: string; at: string; reason: string }

export async function listRevisions(date: string, church: string): Promise<Revision[]> {
  const res = await fetch(`${BASE}/services/${date}/snapshots${withChurch(church)}`)
  return res.ok ? res.json() : []
}
export async function revertService(
  date: string, church: string, id: string,
): Promise<OrderOfWorship> {
  const res = await fetch(`${BASE}/services/${date}/revert${withChurch(church)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ id }),
  })
  if (!res.ok) throw new Error('Could not revert')
  return res.json()
}

export async function uploadHeroImage(date: string, file: File, church: string): Promise<string> {
  const formData = new FormData()
  formData.append('file', file)
  const res = await fetch(`${BASE}/services/${date}/hero-image${withChurch(church)}`, {
    method: 'POST',
    body: formData,
  })
  const data = await res.json()
  return data.filename
}

export function downloadUrl(filename: string): string {
  return `${BASE}/download/${encodeURIComponent(filename)}`
}

/** Configured churches. An empty list is valid — it means the neutral default. */
export async function listChurches(): Promise<Church[]> {
  const res = await fetch(`${BASE}/churches`)
  if (!res.ok) return []
  const data = await res.json()
  return data.churches ?? []
}

/** Copy a hero image so the arriving church owns its own copy (FEATURE-016).
 *  ⛔ Returns null when there was nothing to copy — a missing picture must not stop a
 *  church switch. */
export async function carryHeroImage(
  date: string, filename: string, church: string,
): Promise<string | null> {
  const res = await fetch(
    `${BASE}/services/${date}/hero-image/carry?filename=${encodeURIComponent(filename)}&church=${encodeURIComponent(church)}`,
    { method: 'POST' },
  )
  if (!res.ok) return null
  return (await res.json()).filename ?? null
}

export async function getChurchDefaults(id: string): Promise<ChurchDefaults | undefined> {
  const res = await fetch(`${BASE}/churches/${encodeURIComponent(id)}`)
  if (!res.ok) return undefined
  const data = await res.json()
  return data.defaults
}

/** ⭐ The church decides which Word template and which slide palette are used, so
 *  it has to travel with every generate call. */
export function generateUrl(kind: 'bulletin' | 'slides', date: string, church: string): string {
  const q = church ? `?church=${encodeURIComponent(church)}` : ''
  return `${BASE}/generate/${kind}/${date}${q}`
}
