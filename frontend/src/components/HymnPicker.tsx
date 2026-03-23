import { useState, useEffect, useRef, useCallback } from 'react'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { searchHymns } from '@/lib/api'
import type { HymnRef, HymnSearchResult } from '@/types'

const SOURCE_LABELS: Record<string, string> = {
  umh: 'UMH',
  'umh-services': 'Services',
  'umh-general': 'General',
  'umh-psalter': 'Psalter',
  tfws: 'TFWS',
}

interface HymnPickerProps {
  label: string
  value: HymnRef | null
  onChange: (hymn: HymnRef | null) => void
  hint?: string
}

export function HymnPicker({ label, value, onChange, hint }: HymnPickerProps) {
  const [showHint, setShowHint] = useState(false)
  const hintRef = useRef<HTMLDivElement>(null)
  const [query, setQuery] = useState('')
  const [results, setResults] = useState<HymnSearchResult[]>([])
  const [open, setOpen] = useState(false)
  const [activeIndex, setActiveIndex] = useState(-1)
  const [dropUp, setDropUp] = useState(false)
  const debounceRef = useRef<ReturnType<typeof setTimeout>>()
  const containerRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  // Decide whether to drop up or down based on viewport position
  const updateDropDirection = useCallback(() => {
    if (!containerRef.current) return
    const rect = containerRef.current.getBoundingClientRect()
    const spaceBelow = window.innerHeight - rect.bottom
    setDropUp(spaceBelow < 280)
  }, [])

  useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current)
    if (!query.trim()) {
      setResults([])
      return
    }
    debounceRef.current = setTimeout(async () => {
      const data = await searchHymns(query)
      setResults(data)
      setOpen(data.length > 0)
      setActiveIndex(-1)
      updateDropDirection()
    }, 200)
    return () => { if (debounceRef.current) clearTimeout(debounceRef.current) }
  }, [query, updateDropDirection])

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  function select(hymn: HymnSearchResult) {
    onChange({ number: hymn.number, title: hymn.title, source: hymn.source })
    setQuery('')
    setResults([])
    setOpen(false)
  }

  function clear() {
    onChange(null)
    setQuery('')
    setTimeout(() => inputRef.current?.focus(), 0)
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (!open || results.length === 0) return
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      setActiveIndex(i => Math.min(i + 1, results.length - 1))
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      setActiveIndex(i => Math.max(i - 1, 0))
    } else if (e.key === 'Enter' && activeIndex >= 0) {
      e.preventDefault()
      select(results[activeIndex])
    } else if (e.key === 'Escape') {
      setOpen(false)
    }
  }

  const dropdownClasses = [
    'absolute z-50 left-0 right-0 bg-popover border border-border rounded-lg shadow-lg max-h-60 overflow-y-auto',
    dropUp ? 'bottom-full mb-1' : 'top-full mt-1',
  ].join(' ')

  return (
    <div ref={containerRef} className="relative">
      <div className="flex items-center gap-1.5 mb-1.5">
        <Label className="text-sm font-medium text-muted-foreground">{label}</Label>
        {hint && (
          <div className="relative" ref={hintRef}>
            <button
              type="button"
              className="text-muted-foreground/50 hover:text-primary transition-colors"
              onMouseEnter={() => setShowHint(true)}
              onMouseLeave={() => setShowHint(false)}
              onClick={() => setShowHint(!showHint)}
              aria-label="Info"
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="w-3.5 h-3.5">
                <circle cx="12" cy="12" r="10"/>
                <path d="M12 16v-4M12 8h.01"/>
              </svg>
            </button>
            {showHint && (
              <div className="absolute z-50 left-0 bottom-full mb-1.5 px-3 py-1.5 text-xs bg-popover text-popover-foreground border border-border rounded-md shadow-md whitespace-nowrap">
                {hint}
                <div className="absolute left-3 top-full w-2 h-2 bg-popover border-r border-b border-border rotate-45 -mt-1" />
              </div>
            )}
          </div>
        )}
      </div>
      {value ? (
        <div className="flex items-center gap-2 rounded-lg border border-input bg-background px-3 py-2 transition-colors hover:border-primary/40">
          <span className="text-xs font-medium bg-primary/10 text-primary px-2 py-0.5 rounded-md">
            {SOURCE_LABELS[value.source] || value.source} {value.number}
          </span>
          <span className="flex-1 text-sm font-medium">{value.title}</span>
          <button
            onClick={clear}
            className="text-muted-foreground hover:text-foreground transition-colors p-0.5"
            title="Clear selection"
          >
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="w-4 h-4">
              <path d="M18 6 6 18M6 6l12 12"/>
            </svg>
          </button>
        </div>
      ) : (
        <Input
          ref={inputRef}
          placeholder="Search by number or title..."
          value={query}
          onChange={e => setQuery(e.target.value)}
          onFocus={() => {
            if (results.length > 0) {
              updateDropDirection()
              setOpen(true)
            }
          }}
          onKeyDown={handleKeyDown}
        />
      )}
      {open && results.length > 0 && (
        <div className={dropdownClasses}>
          {results.map((hymn, i) => (
            <button
              key={`${hymn.source}-${hymn.number}-${i}`}
              className={`w-full text-left px-3 py-2.5 text-sm flex items-center gap-2 transition-colors ${
                i === activeIndex
                  ? 'bg-primary/10 text-primary'
                  : 'hover:bg-accent'
              }`}
              onMouseDown={() => select(hymn)}
            >
              <span className="text-xs font-medium bg-muted px-2 py-0.5 rounded-md shrink-0">
                {SOURCE_LABELS[hymn.source] || hymn.source} {hymn.number}
              </span>
              <span className="truncate">{hymn.title}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  )
}
