import { useState, useRef, useEffect, useCallback } from 'react'
import { createPortal } from 'react-dom'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Tooltip } from '@/components/ui/tooltip'
import {
  PICKER_LIST, namesForDate, autofillFor, styleOf, parseISO,
  type Style, type DayName,
} from '@/lib/church-year'

/**
 * The service title — free text, with the church year one click away.
 *
 * ⭐⭐ WHY A COMBOBOX AND NOT A TOGGLE. @decision:gold 2026-09-25 · FEATURE-011
 * A toggle asks the user to decide HOW they are going to type before they know
 * what they want, and it is a state they can be in wrongly. Here there is no
 * mode: the field is an ordinary text box, typing always wins, and the list is
 * additive. Nothing the dropdown does can stop someone writing "Homecoming".
 *
 * ⭐⭐ THE BUTTON IS WHY THERE IS NO AUTOMATIC FILL. Filling the field on its own
 * made "when may we overwrite what was typed?" a question that had to be
 * guessed at; a button makes it the user's action, so the question disappears.
 * That matters because overriding is not the edge case — a church in a sermon
 * series prints the series, not the lectionary, for weeks at a time.
 *
 * ⚠ THE STYLE MEMORY LIVES IN localStorage, PER CHURCH, AND THAT IS A SEAM.
 * It belongs on the church profile, but there is no write path to `church.json`
 * yet (FEATURE-006 is deferred), and localStorage does not follow the user from
 * the dev server to the packaged app. Move it when a profile write path exists.
 */

const STYLE_KEY = (church: string | null) => `oo.titleStyle.${church || 'default'}`
const DEFAULT_STYLE: Style = 'ordinary'

function readStyle(church: string | null): Style {
  try {
    const v = localStorage.getItem(STYLE_KEY(church))
    if (v === 'numbered' || v === 'proper' || v === 'season' || v === 'ordinary') return v
  } catch { /* private window, blocked storage — the default is fine */ }
  return DEFAULT_STYLE
}
function writeStyle(church: string | null, s: Style) {
  try { localStorage.setItem(STYLE_KEY(church), s) } catch { /* not worth failing over */ }
}

/** "Sunday, October 4, 2026" — formatted in UTC, like everything in church-year. */
function longDate(isoDate: string): string {
  const d = parseISO(isoDate)
  if (!d) return ''
  return d.toLocaleDateString(undefined, {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', timeZone: 'UTC',
  })
}

interface Props {
  value: string
  onChange: (v: string) => void
  /** The service date, ISO. Decides what the button fills and what is pinned on top. */
  date: string
  /** Church id — the style memory is per congregation. */
  church: string | null
}

export function ServiceTitlePicker({ value, onChange, date, church }: Props) {
  const [open, setOpen] = useState(false)
  const [activeIndex, setActiveIndex] = useState(-1)
  /** Has anyone typed since the list opened? Decides whether `value` filters. */
  const [typed, setTyped] = useState(false)
  const containerRef = useRef<HTMLDivElement>(null)
  const dropdownRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  /**
   * ⛔⛔ THE DROPDOWN IS PORTALLED, AND IT HAS TO BE. `Card` sets
   * `overflow-hidden` — its image corner-rounding depends on that — so an
   * absolutely-positioned list inside a card is CLIPPED at the card's edge, and
   * this field sits on the bottom row of Service Information. Positioning alone
   * cannot escape an ancestor's overflow; only leaving the subtree can.
   * @decision:gold 2026-09-25 · FEATURE-011
   * ⚠ HymnPicker has the same latent bug and has only been spared by never
   * sitting at a card's bottom edge.
   */
  const [rect, setRect] = useState<{ left: number; top: number; width: number; up: boolean } | null>(null)

  const forDate: DayName[] = date ? namesForDate(date) : []

  /**
   * ⛔⛔ THE STANDING VALUE IS NOT A FILTER. Filtering on `value` meant opening
   * the picker on a service already titled "Ordinary Time" showed exactly one
   * row — the thing you already had — when the whole reason to open it is to
   * change that. The list narrows only once someone TYPES.
   * @decision:gold 2026-09-25 · FEATURE-011
   */
  const query = typed ? value.trim().toLowerCase() : ''

  // The date's own names sit on top; the rest of the year follows. Filtering
  // applies to both, so typing narrows one long list rather than two.
  const pinned = forDate.map(n => n.name).filter(n => !query || n.toLowerCase().includes(query))
  const pinnedSet = new Set(forDate.map(n => n.name))
  const rest = PICKER_LIST.filter(n => !pinnedSet.has(n) && (!query || n.toLowerCase().includes(query)))
  const rows = [...pinned, ...rest]

  const updateDropDirection = useCallback(() => {
    const el = inputRef.current
    if (!el) return
    const r = el.getBoundingClientRect()
    const below = window.innerHeight - r.bottom
    const up = below < 300 && r.top > below
    setRect({
      left: r.left,
      top: up ? r.top - 6 : r.bottom + 6,
      width: r.width,
      up,
    })
  }, [])

  // The portal is positioned from a rect, so anything that moves the field has
  // to re-measure or the list detaches from it. `true` captures scrolls in any
  // scrolling ancestor, not just the window.
  useEffect(() => {
    if (!open) return
    const onMove = () => updateDropDirection()
    window.addEventListener('scroll', onMove, true)
    window.addEventListener('resize', onMove)
    return () => {
      window.removeEventListener('scroll', onMove, true)
      window.removeEventListener('resize', onMove)
    }
  }, [open, updateDropDirection])

  useEffect(() => {
    function onDocMouseDown(e: MouseEvent) {
      const t = e.target as Node
      // ⚠ Both, because the list is no longer a descendant of the field.
      if (containerRef.current?.contains(t)) return
      if (dropdownRef.current?.contains(t)) return
      setOpen(false)
    }
    document.addEventListener('mousedown', onDocMouseDown)
    return () => document.removeEventListener('mousedown', onDocMouseDown)
  }, [])

  function choose(name: string) {
    onChange(name)
    // ⭐ Only a tier-3 pick teaches the button anything. Taking Palm Sunday must
    // not wipe out a congregation's standing preference for "Ordinary Time".
    const s = date ? styleOf(date, name) : null
    if (s) writeStyle(church, s)
    setOpen(false)
    setActiveIndex(-1)
    setTyped(false)
  }

  function fillFromDate() {
    if (!date) return
    const name = autofillFor(date, readStyle(church))
    if (name) onChange(name)
    inputRef.current?.focus()
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (e.key === 'ArrowDown' && !open) { updateDropDirection(); setOpen(true); return }
    if (!open || rows.length === 0) return
    if (e.key === 'ArrowDown') {
      e.preventDefault(); setActiveIndex(i => Math.min(i + 1, rows.length - 1))
    } else if (e.key === 'ArrowUp') {
      e.preventDefault(); setActiveIndex(i => Math.max(i - 1, 0))
    } else if (e.key === 'Enter' && activeIndex >= 0) {
      e.preventDefault(); choose(rows[activeIndex])
    } else if (e.key === 'Escape') {
      setOpen(false); setActiveIndex(-1); setTyped(false)
    }
  }

  const suggestion = date ? autofillFor(date, readStyle(church)) : null

  return (
    <div ref={containerRef} className="relative">
      <Label htmlFor="serviceTitle" className="mb-1.5 block">Service Title / Season</Label>
      <div className="flex items-center gap-1.5">
        <Input
          ref={inputRef}
          id="serviceTitle"
          placeholder="e.g., Lent"
          autoComplete="off"
          value={value}
          onChange={e => { onChange(e.target.value); setTyped(true); setActiveIndex(-1); if (!open) { updateDropDirection(); setOpen(true) } }}
          onFocus={() => { setTyped(false); updateDropDirection(); setOpen(true) }}
          onKeyDown={handleKeyDown}
        />
        <Tooltip
          side="top"
          align="right"
          content={suggestion
            ? <span>Fill from the date — <span className="font-medium">{suggestion}</span></span>
            : 'Pick a date first'}
        >
        <button
          type="button"
          onClick={fillFromDate}
          disabled={!suggestion}
          aria-label="Fill from the date"
          className="shrink-0 h-9 w-9 flex items-center justify-center rounded-md border border-input text-muted-foreground transition-colors hover:text-primary hover:border-primary/40 disabled:opacity-40 disabled:hover:text-muted-foreground disabled:hover:border-input"
        >
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="w-4 h-4">
            <rect x="3" y="4" width="18" height="18" rx="2" />
            <path d="M16 2v4M8 2v4M3 10h18M8 15h4" />
          </svg>
        </button>
        </Tooltip>
      </div>

      {open && rows.length > 0 && rect && createPortal(
        <div
          ref={dropdownRef}
          style={{
            position: 'fixed',
            left: rect.left,
            width: rect.width,
            ...(rect.up ? { bottom: window.innerHeight - rect.top } : { top: rect.top }),
          }}
          className="z-50 bg-popover border border-border rounded-lg shadow-lg max-h-72 overflow-y-auto"
        >
          {pinned.length > 0 && (
            <div className="px-3 pt-2 pb-1 text-[11px] uppercase tracking-wide text-muted-foreground/70 sticky top-0 bg-popover">
              {longDate(date)}
            </div>
          )}
          {rows.map((name, i) => {
            const isPinned = i < pinned.length
            const showRestHeader = i === pinned.length && pinned.length > 0
            const tier = forDate.find(n => n.name === name)?.tier
            return (
              <div key={name}>
                {showRestHeader && (
                  <div className="px-3 pt-2 pb-1 text-[11px] uppercase tracking-wide text-muted-foreground/70 border-t border-border">
                    The church year
                  </div>
                )}
                <button
                  type="button"
                  className={`w-full text-left px-3 py-2 text-sm flex items-center gap-2 transition-colors ${
                    i === activeIndex ? 'bg-primary/10 text-primary' : 'hover:bg-accent'
                  }`}
                  onMouseDown={() => choose(name)}
                >
                  <span className="flex-1">{name}</span>
                  {isPinned && tier === 1 && (
                    <span className="text-[10px] font-medium bg-primary/10 text-primary px-1.5 py-0.5 rounded shrink-0">holy day</span>
                  )}
                  {isPinned && tier === 2 && (
                    <span className="text-[10px] font-medium bg-muted px-1.5 py-0.5 rounded shrink-0">special Sunday</span>
                  )}
                </button>
              </div>
            )
          })}
        </div>,
        document.body,
      )}
    </div>
  )
}
