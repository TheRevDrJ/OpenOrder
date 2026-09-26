import { useState, useEffect, useRef } from 'react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { listChurches } from '@/lib/api'
import type { Church } from '@/types'

interface BibleKeyStatus {
  configured: boolean
  /** A stored key that cannot list right now is still stored — say which. */
  reachable: boolean
  /** Why it could not: a rejected key and a dead connection are different problems. */
  reason?: string
  bibles: { id: string; abbr: string; name: string; lang: string; copyright: string }[]
}

/** One row of the Templates tab: a church's template of one kind. */
interface TemplateRow {
  kind: string
  label: string
  name: string | null
  exists: boolean
  found: string[]
  missing: string[]
  total_expected: number
}

/** What `?confirm=false` reports back: what WOULD be replaced, and the tag diff
 *  against the file being replaced. ⭐ `added`/`removed` are both empty when the two
 *  templates carry the same tags, and the confirm then says nothing about tags. */
interface UploadReport {
  church: string | null
  kind: string
  label: string
  target: string
  replacing: boolean
  found: string[]
  missing: string[]
  outgoing_count: number | null
  added: string[]
  removed: string[]
  total_expected: number
}

// The three configurable locations. Order is deliberate: the one people change
// most often (where their files land) sits at the top.
const FOLDERS = [
  {
    key: 'output_dir',
    label: 'Output Folder',
    hint: 'Where generated bulletins and slides are saved.',
  },
  {
    key: 'data_dir',
    label: 'Data Folder',
    hint: 'Your calendar, saved services, theme images, and bulletin template.',
  },
  {
    key: 'hymnal_dir',
    label: 'Hymnal Folder',
    hint: 'The folder holding your hymnal JSON files.',
  },
] as const

const TABS = [
  { id: 'folders', label: 'Folders' },
  { id: 'templates', label: 'Templates' },
  { id: 'translations', label: 'Translations' },
  { id: 'appearance', label: 'Appearance' },
] as const

type TabId = (typeof TABS)[number]['id']

export function SettingsPanel({
  open,
  onClose,
  onSaved,
  church,
}: {
  open: boolean
  onClose: () => void
  /** Report a file written to disk, so the app can confirm it the same way it
   *  confirms a generated bulletin. */
  onSaved?: (kind: string, data: { filename: string; folder?: string }) => void
  /** Which church is on the SERVICE FORM. ⭐ Used only to pick the row the Templates
   *  tab opens on, because that is usually the one you came here about. ⛔ It does not
   *  bind the tab: a template belongs to a church, and which church is on the form is
   *  a fact about this Sunday. (@decision:gold 2026-09-25 — "If the communion slider
   *  shouldn't count there, the church selection shouldn't count here either.") */
  church?: string
}) {
  const [dirs, setDirs] = useState<Record<string, string>>({})
  const [churchList, setChurchList] = useState<Church[]>([])
  const [templateChurch, setTemplateChurch] = useState('')
  const [templates, setTemplates] = useState<TemplateRow[]>([])
  const [busyKind, setBusyKind] = useState<string | null>(null)
  const [pending, setPending] = useState<{ kind: string; file: File; report: UploadReport } | null>(null)
  const pendingKind = useRef<string>('normal')
  const fileRef = useRef<HTMLInputElement>(null)
  const [bibleKey, setBibleKey] = useState<BibleKeyStatus | null>(null)
  const [keyInput, setKeyInput] = useState('')
  const [savingKey, setSavingKey] = useState(false)
  const [keyError, setKeyError] = useState('')
  const [themeMode, setThemeMode] = useState<'light' | 'dark' | 'system'>(
    () => (localStorage.getItem('theme') as 'light' | 'dark' | 'system') || 'system'
  )
  /** ⭐ Four settings in one scroll was too much to read at once, so they are tabs.
   *  @decision:gold 2026-09-25 ⛔ Opens on Folders every time rather than remembering
   *  the last one — a panel that opens somewhere different each visit makes the
   *  person find their place before they can do anything. */
  const [tab, setTab] = useState<TabId>('folders')

  useEffect(() => {
    if (!open) return
    fetch('/api/settings').then(r => r.json()).then(setDirs)
    fetch('/api/bible/key').then(r => r.json()).then(setBibleKey).catch(() => {})
    // ⭐ The tab owns its own church. It OPENS on whichever is on the form, because
    // that is nearly always the one you came here about — but it is a starting point,
    // not a binding: changing it here does not touch the service.
    listChurches().then(list => {
      setChurchList(list)
      setTemplateChurch(prev => {
        if (prev && list.some(c => c.id === prev)) return prev
        if (church && list.some(c => c.id === church)) return church
        return list[0]?.id ?? ''
      })
    })
  }, [open, church])

  /** Read both rows for one church. ⛔ Always off the server, never patched up from
   *  an upload response — the response describes what was sent, the server knows what
   *  is on disk. */
  async function loadTemplates(churchId: string) {
    const q = churchId ? `?church=${encodeURIComponent(churchId)}` : ''
    const res = await fetch(`/api/templates${q}`)
    if (!res.ok) { setTemplates([]); return }
    const data = await res.json()
    setTemplates(data.templates ?? [])
  }

  useEffect(() => {
    if (!open) return
    // With no churches configured the endpoint answers with the shared template,
    // so an empty id is a legitimate request rather than a reason to skip.
    if (churchList.length > 0 && !templateChurch) return
    loadTemplates(templateChurch)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, templateChurch, churchList.length])

  async function handleChangeDir(key: string, label: string) {
    const pywebview = (window as any).pywebview
    let dir: string | null = null

    if (pywebview?.api?.pick_folder) {
      dir = await pywebview.api.pick_folder()
    } else {
      dir = prompt(`${label} path:`, dirs[key] || '')
    }

    if (dir) {
      dir = dir.replace(/\\\\/g, '/').replace(/\\/g, '/')
      const res = await fetch('/api/settings/dir', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key, path: dir })
      })
      if (res.ok) {
        setDirs(await res.json())
        // The hymnal index is cached server-side; a reload keeps the open
        // form in step with whatever the new folder holds.
        if (key === 'hymnal_dir') window.location.reload()
      } else {
        const err = await res.json()
        alert(err.detail || 'Failed to set folder')
      }
    }
  }

  /** Save or clear the key. ⭐ The server validates before storing, so an error
   *  here means the key is bad and nothing was written. */
  async function handleSaveKey(key: string) {
    setSavingKey(true)
    setKeyError('')
    try {
      const res = await fetch('/api/bible/key', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key }),
      })
      const data = await res.json()
      if (!res.ok) {
        setKeyError(data.detail || 'Could not save that key.')
        return
      }
      setKeyInput('')
      setBibleKey({ configured: data.configured, reachable: true, bibles: data.bibles || [] })
    } catch {
      setKeyError('Could not reach the server.')
    } finally {
      setSavingKey(false)
    }
  }

  const templateQuery = (kind: string) =>
    `?${templateChurch ? `church=${encodeURIComponent(templateChurch)}&` : ''}kind=${kind}`

  async function handleTemplateDownload(kind: string) {
    setBusyKind(kind)
    try {
      const res = await fetch(`/api/template/export${templateQuery(kind)}`, { method: 'POST' })
      const data = await res.json()
      if (res.ok) onSaved?.('Template', data)
      else alert(data.detail || 'Could not save the template')
    } catch {
      alert('Could not reach server')
    } finally {
      setBusyKind(null)
    }
  }

  /** One hidden file input serves both rows, so remember which asked. */
  function pickFor(kind: string) {
    pendingKind.current = kind
    fileRef.current?.click()
  }

  /** PHASE 1 — ask the server what this file WOULD do. ⛔⛔ Nothing is written here:
   *  the endpoint only writes with `confirm`, so a stray click cannot replace a
   *  template and the tag diff is read off the outgoing file while it still exists. */
  async function handleTemplateChosen(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file) return
    const kind = pendingKind.current
    setBusyKind(kind)
    try {
      const body = new FormData()
      body.append('file', file)
      const res = await fetch(`/api/template/upload${templateQuery(kind)}`, { method: 'POST', body })
      const report = await res.json()
      if (!res.ok) { alert(report.detail || 'Could not read that file'); return }
      setPending({ kind, file, report })
    } catch {
      alert('Could not reach server')
    } finally {
      setBusyKind(null)
    }
  }

  /** PHASE 2 — he has read what it will do and said yes. */
  async function confirmReplace() {
    if (!pending) return
    const { kind, file } = pending
    setPending(null)
    setBusyKind(kind)
    try {
      const body = new FormData()
      body.append('file', file)
      const res = await fetch(`/api/template/upload${templateQuery(kind)}&confirm=true`, {
        method: 'POST', body,
      })
      const data = await res.json()
      if (!res.ok) { alert(data.detail || 'Upload failed'); return }
      await loadTemplates(templateChurch)
      onSaved?.('Template', { filename: data.name })
    } catch {
      alert('Could not reach server')
    } finally {
      setBusyKind(null)
    }
  }

  function applyTheme(mode: 'light' | 'dark' | 'system') {
    const html = document.documentElement
    localStorage.setItem('theme', mode)
    setThemeMode(mode)

    if (mode === 'dark') {
      html.classList.add('dark')
    } else if (mode === 'light') {
      html.classList.remove('dark')
    } else {
      // System preference
      if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
        html.classList.add('dark')
      } else {
        html.classList.remove('dark')
      }
    }
  }

  if (!open) return null

  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center" onClick={onClose}>
      <div
        className="bg-card rounded-lg shadow-xl border border-border w-full max-w-lg mx-4 max-h-[80vh] overflow-y-auto"
        onClick={e => e.stopPropagation()}
      >
        <div className="flex items-center justify-between p-4 border-b border-border">
          <h2 className="text-lg font-bold text-foreground">Settings</h2>
          <button onClick={onClose} className="text-muted-foreground hover:text-foreground text-xl leading-none">&times;</button>
        </div>

        <div className="flex border-b border-border px-2" role="tablist" aria-label="Settings sections">
          {TABS.map(({ id, label }) => (
            <button
              key={id}
              role="tab"
              type="button"
              aria-selected={tab === id}
              onClick={() => setTab(id)}
              className={`px-3 py-2 text-sm border-b-2 -mb-px transition-colors ${
                tab === id
                  ? 'border-primary text-foreground font-medium'
                  : 'border-transparent text-muted-foreground hover:text-foreground'
              }`}
            >
              {label}
            </button>
          ))}
        </div>

        <div className="p-4" role="tabpanel">
          {/* Folders — the tab names the section, so no heading here */}
          {tab === 'folders' && (
          <div>
            <div className="space-y-4">
              {FOLDERS.map(({ key, label, hint }) => (
                <div key={key}>
                  <div className="flex items-baseline justify-between gap-2">
                    <span className="text-sm font-medium text-foreground">{label}</span>
                    <Button variant="outline" size="sm" onClick={() => handleChangeDir(key, label)}>
                      Change
                    </Button>
                  </div>
                  <p className="text-xs text-muted-foreground mt-0.5">{hint}</p>
                  <p className="text-sm text-foreground mt-1 font-mono bg-muted rounded px-2 py-1.5 break-all">
                    {dirs[key] || 'Not set'}
                  </p>
                </div>
              ))}
            </div>
          </div>
          )}

          {/* Bulletin templates — one church at a time, both of its templates listed.
              ⭐⭐ THE CHURCH IS CHOSEN HERE, NOT INHERITED FROM THE SERVICE FORM.
              @decision:gold 2026-09-25 — a template is a church's asset, while which
              church sits on the service form is a fact about this Sunday.
              ⛔ NOT CASCADING DROPDOWNS. There are exactly two kinds, both known, so a
              second picker would cost a click and hide one of them. Two rows instead. */}
          {tab === 'templates' && (
          <div>
            {churchList.length > 1 ? (
              <div className="flex items-center gap-2">
                <span className="text-sm text-muted-foreground">Church</span>
                <select
                  value={templateChurch}
                  onChange={e => setTemplateChurch(e.target.value)}
                  className="flex-1 h-9 rounded-md border border-input bg-background px-2 text-sm"
                >
                  {churchList.map(c => (
                    <option key={c.id} value={c.id}>{c.name || c.id}</option>
                  ))}
                </select>
              </div>
            ) : churchList.length === 1 ? (
              <p className="text-sm text-muted-foreground">
                Church <span className="text-foreground font-medium">{churchList[0].name || churchList[0].id}</span>
              </p>
            ) : null}

            <div className="mt-3 space-y-2">
              {templates.map(t => (
                <div key={t.kind} className="rounded-md border border-border p-3">
                  <div className="flex items-baseline justify-between gap-2">
                    <span className="text-sm font-medium text-foreground">{t.label}</span>
                    {t.exists ? (
                      <span className="text-xs text-muted-foreground">
                        {t.found.length}/{t.total_expected} placeholders
                      </span>
                    ) : (
                      <span className="text-xs text-muted-foreground italic">Not set up</span>
                    )}
                  </div>
                  <p className="mt-1 text-xs font-mono text-muted-foreground break-all">
                    {t.name || '—'}
                  </p>
                  {t.exists && t.missing.length > 0 && (
                    <div className="mt-2 flex flex-wrap gap-1">
                      {t.missing.map(m => (
                        <span key={m} className="text-[11px] bg-muted rounded px-1.5 py-0.5 font-mono" title="Not in this template">
                          {m}
                        </span>
                      ))}
                    </div>
                  )}
                  <div className="mt-2 flex items-center gap-2">
                    <Button
                      variant="outline"
                      size="sm"
                      disabled={!t.exists || busyKind === t.kind}
                      title={t.exists ? 'Save a copy to your output folder' : 'Nothing to download yet'}
                      onClick={() => handleTemplateDownload(t.kind)}
                    >
                      Download
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      disabled={busyKind === t.kind}
                      onClick={() => pickFor(t.kind)}
                    >
                      {busyKind === t.kind ? 'Checking…' : t.exists ? 'Replace' : 'Upload'}
                    </Button>
                  </div>
                </div>
              ))}
            </div>
            <input ref={fileRef} type="file" accept=".docx" className="hidden" onChange={handleTemplateChosen} />
            <p className="mt-2 text-xs text-muted-foreground">.docx only</p>
          </div>
          )}

          {/* Bible translations — a church's OWN free API.Bible key.
              ⭐⭐ The key belongs to the congregation, not to us: nothing is
              shared, no quota is pooled, and nobody pays anything, ever.
              ⛔ The key is never read back from the server, so this shows what
              it unlocks rather than the credential itself. */}
          {tab === 'translations' && (
          <div>
            <p className="text-xs text-muted-foreground">
              OpenOrder includes BSB, KJV, ASV and WEB, which are free to use. For other
              translations, get a free key from{' '}
              <a href="https://api.bible" target="_blank" rel="noreferrer" className="text-primary hover:underline">API.Bible</a>
              {' '}and paste it here — it stays on this computer.
            </p>

            {bibleKey?.configured ? (
              <div className="mt-3">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-sm">
                    {bibleKey.reachable
                      ? `Key saved — ${bibleKey.bibles.length} translation${bibleKey.bibles.length === 1 ? '' : 's'} available`
                      : bibleKey.reason?.includes('rejected')
                        ? 'Key saved, but API.Bible is not accepting it'
                        : 'Key saved — API.Bible could not be reached just now'}
                  </span>
                  <Button
                    variant="outline"
                    size="sm"
                    disabled={savingKey}
                    onClick={() => handleSaveKey('')}
                  >
                    Remove
                  </Button>
                </div>
                {bibleKey.bibles.length > 0 && (
                  <div className="mt-2 flex flex-wrap gap-1">
                    {bibleKey.bibles.map(b => (
                      <span key={b.id} title={b.name} className="text-[11px] bg-muted px-1.5 py-0.5 rounded">
                        {b.abbr}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            ) : (
              <div className="mt-3 flex items-center gap-2">
                <Input
                  value={keyInput}
                  onChange={e => { setKeyInput(e.target.value); setKeyError('') }}
                  placeholder="Paste your API.Bible key"
                  autoComplete="off"
                  spellCheck={false}
                  onKeyDown={e => { if (e.key === 'Enter' && keyInput.trim()) handleSaveKey(keyInput) }}
                />
                <Button
                  size="sm"
                  disabled={savingKey || !keyInput.trim()}
                  onClick={() => handleSaveKey(keyInput)}
                >
                  {savingKey ? 'Checking…' : 'Save'}
                </Button>
              </div>
            )}
            {keyError && <p className="mt-2 text-xs text-destructive">{keyError}</p>}
          </div>
          )}

          {/* Appearance */}
          {tab === 'appearance' && (
          <div>
            <div className="flex items-center gap-2">
              {([
                { mode: 'light' as const, label: 'Light', icon: <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/></svg> },
                { mode: 'dark' as const, label: 'Dark', icon: <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></svg> },
                { mode: 'system' as const, label: 'System', icon: <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect width="20" height="14" x="2" y="3" rx="2"/><line x1="8" x2="16" y1="21" y2="21"/><line x1="12" x2="12" y1="17" y2="21"/></svg> },
              ]).map(({ mode, label, icon }) => (
                <button
                  key={mode}
                  onClick={() => applyTheme(mode)}
                  className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md border text-sm transition-colors ${
                    themeMode === mode
                      ? 'border-primary bg-primary/10 text-primary font-medium'
                      : 'border-input bg-background hover:bg-accent text-muted-foreground'
                  }`}
                >
                  {icon}
                  {label}
                </button>
              ))}
            </div>
          </div>
          )}
        </div>
      </div>

      {/* ⛔⛔ REPLACING A TEMPLATE IS DESTRUCTIVE, SO IT IS CONFIRMED, AND THE CONFIRM
          NAMES WHAT IT WILL OVERWRITE. @decision:gold 2026-09-25 The tag diff is
          against the file being replaced, and says NOTHING when the two carry the same
          tags. ⛔ It never refuses over a missing tag: leaving an element out is a
          legitimate choice, so this reports it and the reader decides. */}
      <ConfirmDialog
        open={!!pending}
        title={pending?.report.replacing ? `Replace the ${pending.report.label} template?` : `Upload the ${pending?.report.label} template?`}
        confirmLabel={pending?.report.replacing ? 'Replace' : 'Upload'}
        onCancel={() => setPending(null)}
        onConfirm={confirmReplace}
        body={pending && (
          <>
            <p>
              {pending.report.replacing ? 'This overwrites ' : 'This creates '}
              <span className="font-mono text-foreground">{pending.report.target}</span>
              {churchList.length > 0 && (
                <> for <span className="text-foreground font-medium">
                  {churchList.find(c => c.id === pending.report.church)?.name || pending.report.church}
                </span></>
              )}.
            </p>
            {(pending.report.added.length > 0 || pending.report.removed.length > 0) && (
              <div className="mt-2">
                <p>
                  The one you are replacing has{' '}
                  <span className="text-foreground">{pending.report.outgoing_count}</span> tags; this one has{' '}
                  <span className="text-foreground">{pending.report.found.length}</span>.
                </p>
                {pending.report.removed.length > 0 && (
                  <p className="mt-1">
                    No longer present:{' '}
                    {pending.report.removed.map(t => (
                      <span key={t} className="font-mono text-foreground">{t} </span>
                    ))}
                  </p>
                )}
                {pending.report.added.length > 0 && (
                  <p className="mt-1">
                    Newly present:{' '}
                    {pending.report.added.map(t => (
                      <span key={t} className="font-mono text-foreground">{t} </span>
                    ))}
                  </p>
                )}
              </div>
            )}
          </>
        )}
      />
    </div>
  )
}
