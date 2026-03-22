import type { HymnSearchResult, OrderOfWorship } from '@/types'

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

export async function listServices(): Promise<{ date: string; filename: string }[]> {
  const res = await fetch(`${BASE}/services`)
  return res.json()
}

export async function loadService(date: string): Promise<OrderOfWorship> {
  const res = await fetch(`${BASE}/services/${date}`)
  if (!res.ok) throw new Error('Service not found')
  return res.json()
}

export async function saveService(data: OrderOfWorship): Promise<void> {
  await fetch(`${BASE}/services/${data.date}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  })
}

export async function uploadThemeImage(date: string, file: File): Promise<string> {
  const formData = new FormData()
  formData.append('file', file)
  const res = await fetch(`${BASE}/services/${date}/theme-image`, {
    method: 'POST',
    body: formData,
  })
  const data = await res.json()
  return data.filename
}

export function downloadUrl(filename: string): string {
  return `${BASE}/download/${encodeURIComponent(filename)}`
}
