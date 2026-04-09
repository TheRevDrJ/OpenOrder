import { useState, useEffect, useRef } from 'react'
import { Button } from '@/components/ui/button'
import { Label } from '@/components/ui/label'

interface TemplateInfo {
  exists: boolean
  name: string | null
  found: string[]
  missing: string[]
  total_expected: number
}

export function SettingsPanel({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [dataDir, setDataDir] = useState('')
  const [templateInfo, setTemplateInfo] = useState<TemplateInfo | null>(null)
  const [uploading, setUploading] = useState(false)
  const [uploadResult, setUploadResult] = useState<TemplateInfo | null>(null)
  const fileRef = useRef<HTMLInputElement>(null)
  const [darkMode, setDarkMode] = useState(
    document.documentElement.classList.contains('dark')
  )

  useEffect(() => {
    if (open) {
      fetch('/api/settings').then(r => r.json()).then(d => setDataDir(d.data_dir_current || ''))
      fetch('/api/template/info').then(r => r.json()).then(setTemplateInfo)
    }
  }, [open])

  async function handleChangeDir() {
    const pywebview = (window as any).pywebview
    let dir: string | null = null

    if (pywebview?.api?.pick_folder) {
      dir = await pywebview.api.pick_folder()
    } else {
      dir = prompt('Data directory path:', dataDir)
    }

    if (dir) {
      dir = dir.replace(/\\\\/g, '/').replace(/\\/g, '/')
      const res = await fetch('/api/settings/data-dir', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data_dir: dir })
      })
      if (res.ok) {
        window.location.reload()
      } else {
        const err = await res.json()
        alert(err.detail || 'Failed to set directory')
      }
    }
  }

  async function handleTemplateUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setUploading(true)
    setUploadResult(null)

    const formData = new FormData()
    formData.append('file', file)

    try {
      const res = await fetch('/api/template/upload', { method: 'POST', body: formData })
      const data = await res.json()
      if (res.ok) {
        setUploadResult(data)
        setTemplateInfo(data)
      } else {
        alert(data.detail || 'Upload failed')
      }
    } catch {
      alert('Could not reach server')
    } finally {
      setUploading(false)
      e.target.value = ''
    }
  }

  function toggleDarkMode() {
    const html = document.documentElement
    if (html.classList.contains('dark')) {
      html.classList.remove('dark')
      localStorage.setItem('theme', 'light')
      setDarkMode(false)
    } else {
      html.classList.add('dark')
      localStorage.setItem('theme', 'dark')
      setDarkMode(true)
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

        <div className="p-4 space-y-6">
          {/* Data Directory */}
          <div>
            <Label className="text-sm font-semibold text-muted-foreground uppercase tracking-wide">Data Directory</Label>
            <p className="text-sm text-foreground mt-1 font-mono bg-muted rounded px-2 py-1.5 break-all">
              {dataDir || 'Not set'}
            </p>
            <Button variant="outline" size="sm" className="mt-2" onClick={handleChangeDir}>
              Change Directory
            </Button>
          </div>

          {/* Bulletin Template */}
          <div>
            <Label className="text-sm font-semibold text-muted-foreground uppercase tracking-wide">Bulletin Template</Label>
            {templateInfo && (
              <div className="mt-1 text-sm space-y-1">
                <p className="text-foreground">
                  {templateInfo.exists ? (
                    <>
                      <span className="font-medium">{templateInfo.name}</span>
                      <span className="text-muted-foreground ml-2">
                        ({templateInfo.found?.length || 0}/{templateInfo.total_expected} placeholders)
                      </span>
                    </>
                  ) : (
                    <span className="text-destructive">No template found</span>
                  )}
                </p>
                {templateInfo.missing && templateInfo.missing.length > 0 && (
                  <div className="bg-destructive/10 text-destructive rounded p-2 text-xs">
                    <p className="font-medium mb-1">Missing placeholders:</p>
                    {templateInfo.missing.map(p => (
                      <span key={p} className="inline-block bg-destructive/20 rounded px-1.5 py-0.5 mr-1 mb-1 font-mono">{p}</span>
                    ))}
                  </div>
                )}
                {uploadResult && uploadResult.missing?.length === 0 && (
                  <p className="text-green-600 dark:text-green-400 text-xs font-medium">All placeholders found!</p>
                )}
              </div>
            )}
            <div className="mt-2 flex items-center gap-2">
              <Button variant="outline" size="sm" onClick={() => fileRef.current?.click()} disabled={uploading}>
                {uploading ? 'Uploading...' : 'Upload Template'}
              </Button>
              <input ref={fileRef} type="file" accept=".docx" className="hidden" onChange={handleTemplateUpload} />
              <span className="text-xs text-muted-foreground">.docx only</span>
            </div>
          </div>

          {/* Appearance */}
          <div>
            <Label className="text-sm font-semibold text-muted-foreground uppercase tracking-wide">Appearance</Label>
            <div className="mt-2 flex items-center gap-3">
              <button
                onClick={toggleDarkMode}
                className="flex items-center gap-2 px-3 py-1.5 rounded-md border border-input bg-background hover:bg-accent text-sm transition-colors"
              >
                {darkMode ? (
                  <>
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/>
                    </svg>
                    Switch to Light
                  </>
                ) : (
                  <>
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/>
                    </svg>
                    Switch to Dark
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
