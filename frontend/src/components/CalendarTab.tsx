import { useState, useEffect } from 'react'
import { Card, CardHeader, CardTitle, CardContent } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'

const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
const DAY_FULL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

interface Template {
  id: string
  name: string
  dayOfWeek: number
  time: string
  location: string
}

interface Event {
  id: string
  date: string
  time: string
  title: string
  location: string
  description?: string
}

interface Override {
  templateId: string
  date: string
  skip?: boolean
  time?: string
  location?: string
}

interface Note {
  text: string
  enabled: boolean
  bold: boolean
  italic: boolean
  align: 'left' | 'center' | 'right'
}

export function CalendarTab({ serviceDate }: { serviceDate: string }) {
  const [templates, setTemplates] = useState<Template[]>([])
  const [events, setEvents] = useState<Event[]>([])
  const [overrides, setOverrides] = useState<Override[]>([])
  const [loading, setLoading] = useState(true)

  const [editingTemplate, setEditingTemplate] = useState<Template | null>(null)
  const [editingEvent, setEditingEvent] = useState<Event | null>(null)
  const [note, setNote] = useState<Note>({ text: '', enabled: false, bold: false, italic: false, align: 'left' })

  // Compute the 4 dates ahead for each day-of-week
  function upcomingDates(dayOfWeek: number): string[] {
    if (!serviceDate) return []
    const start = new Date(serviceDate + 'T12:00:00')
    const startDow = start.getDay()
    const daysAhead = (dayOfWeek - startDow + 7) % 7
    const first = new Date(start)
    first.setDate(start.getDate() + daysAhead)
    return [0, 1, 2, 3].map(i => {
      const d = new Date(first)
      d.setDate(first.getDate() + i * 7)
      return d.toISOString().slice(0, 10)
    })
  }

  function isSkipped(templateId: string, date: string): boolean {
    return overrides.some(o => o.templateId === templateId && o.date === date && o.skip)
  }

  async function loadAll() {
    setLoading(true)
    const [t, e, o] = await Promise.all([
      fetch('/api/calendar/templates').then(r => r.json()),
      fetch('/api/calendar/events').then(r => r.json()),
      fetch('/api/calendar/overrides').then(r => r.json()),
    ])
    setTemplates(t)
    setEvents(e)
    setOverrides(o)
    setLoading(false)
  }

  async function loadNote() {
    if (!serviceDate) return
    const n = await fetch(`/api/calendar/note/${serviceDate}`).then(r => r.json())
    setNote(n)
  }

  useEffect(() => { loadAll() }, [])
  useEffect(() => { loadNote() }, [serviceDate])

  async function saveNote(updates: Partial<Note>) {
    const newNote = { ...note, ...updates }
    setNote(newNote)
    await fetch(`/api/calendar/note/${serviceDate}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(newNote)
    })
  }

  async function toggleSkip(templateId: string, date: string, currentlySkipped: boolean) {
    await fetch('/api/calendar/overrides/skip', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ templateId, date, skip: !currentlySkipped })
    })
    const o = await fetch('/api/calendar/overrides').then(r => r.json())
    setOverrides(o)
  }

  async function saveTemplate(t: Template) {
    if (t.id && templates.find(x => x.id === t.id)) {
      await fetch(`/api/calendar/templates/${t.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(t)
      })
    } else {
      await fetch('/api/calendar/templates', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(t)
      })
    }
    setEditingTemplate(null)
    loadAll()
  }

  async function deleteTemplate(id: string) {
    if (!confirm('Delete this recurring event? Any per-week overrides will also be removed.')) return
    await fetch(`/api/calendar/templates/${id}`, { method: 'DELETE' })
    loadAll()
  }

  async function saveEvent(e: Event) {
    if (e.id && events.find(x => x.id === e.id)) {
      await fetch(`/api/calendar/events/${e.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(e)
      })
    } else {
      await fetch('/api/calendar/events', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(e)
      })
    }
    setEditingEvent(null)
    loadAll()
  }

  async function deleteEvent(id: string) {
    if (!confirm('Delete this event?')) return
    await fetch(`/api/calendar/events/${id}`, { method: 'DELETE' })
    loadAll()
  }

  // Filter one-off events to those in the next 4 weeks from service date
  const eventsInRange = events.filter(e => {
    if (!serviceDate) return false
    const start = new Date(serviceDate + 'T00:00:00')
    const end = new Date(serviceDate + 'T00:00:00')
    end.setDate(end.getDate() + 28)
    const ev = new Date(e.date + 'T00:00:00')
    return ev >= start && ev < end
  }).sort((a, b) => a.date.localeCompare(b.date))

  function formatDateShort(iso: string): string {
    const d = new Date(iso + 'T12:00:00')
    return `${d.getMonth() + 1}/${d.getDate()}`
  }

  if (loading) {
    return <div className="text-center py-12 text-muted-foreground">Loading calendar...</div>
  }

  return (
    <div className="space-y-6">
      {/* Autosave indicator */}
      <div className="flex items-center gap-2 text-xs text-muted-foreground -mb-3">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-primary">
          <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="m9 11 3 3L22 4"/>
        </svg>
        Changes on this tab save automatically
      </div>

      {/* Recurring Events */}
      <Card className="shadow-sm">
        <CardHeader className="pb-4 flex flex-row items-center justify-between">
          <div>
            <CardTitle className="text-primary">Recurring Events</CardTitle>
            <p className="text-xs text-muted-foreground mt-1">
              Check or uncheck to include in the next 4 weeks
            </p>
          </div>
          <Button
            size="sm"
            variant="outline"
            onClick={() => setEditingTemplate({ id: '', name: '', dayOfWeek: 0, time: '', location: '' })}
          >
            + Add
          </Button>
        </CardHeader>
        <CardContent>
          {templates.length === 0 ? (
            <p className="text-sm text-muted-foreground italic py-4 text-center">
              No recurring events yet. Add Sunday Worship, Sunday School, or any weekly event.
            </p>
          ) : (
            <div className="space-y-3">
              {templates.map(t => {
                const dates = upcomingDates(t.dayOfWeek)
                return (
                  <div key={t.id} className="flex items-center gap-3 py-2 border-b border-border last:border-0">
                    <div className="flex-1 min-w-0">
                      <div className="font-medium text-sm truncate">{t.name}</div>
                      <div className="text-xs text-muted-foreground">
                        {t.time} ({DAY_FULL[t.dayOfWeek]}) {t.location && `· ${t.location}`}
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 flex-shrink-0">
                      {dates.map(d => {
                        const skipped = isSkipped(t.id, d)
                        return (
                          <label
                            key={d}
                            className="flex flex-col items-center cursor-pointer text-xs"
                            title={d}
                          >
                            <input
                              type="checkbox"
                              checked={!skipped}
                              onChange={() => toggleSkip(t.id, d, skipped)}
                              className="w-4 h-4 cursor-pointer accent-primary"
                            />
                            <span className="text-[10px] text-muted-foreground mt-0.5">
                              {formatDateShort(d)}
                            </span>
                          </label>
                        )
                      })}
                    </div>
                    <div className="flex items-center gap-1 flex-shrink-0 ml-2">
                      <button
                        onClick={() => setEditingTemplate(t)}
                        className="p-1 text-muted-foreground hover:text-foreground rounded"
                        title="Edit"
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/>
                        </svg>
                      </button>
                      <button
                        onClick={() => deleteTemplate(t.id)}
                        className="p-1 text-muted-foreground hover:text-destructive rounded"
                        title="Delete"
                      >
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                          <path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                        </svg>
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </CardContent>
      </Card>

      {/* One-off Events */}
      <Card className="shadow-sm">
        <CardHeader className="pb-4 flex flex-row items-center justify-between">
          <div>
            <CardTitle className="text-primary">One-off Events</CardTitle>
            <p className="text-xs text-muted-foreground mt-1">
              Special events in the next 4 weeks
            </p>
          </div>
          <Button
            size="sm"
            variant="outline"
            onClick={() => setEditingEvent({ id: '', date: serviceDate, time: '', title: '', location: '' })}
          >
            + Add
          </Button>
        </CardHeader>
        <CardContent>
          {eventsInRange.length === 0 ? (
            <p className="text-sm text-muted-foreground italic py-4 text-center">
              No one-off events in the next 4 weeks.
            </p>
          ) : (
            <div className="space-y-2">
              {eventsInRange.map(e => (
                <div key={e.id} className="flex items-center gap-3 py-2 border-b border-border last:border-0">
                  <div className="flex-1 min-w-0">
                    <div className="font-medium text-sm">
                      {e.date} · {e.time && <span className="text-primary">{e.time}</span>}
                    </div>
                    <div className="text-sm">{e.title}</div>
                    {e.location && <div className="text-xs text-muted-foreground">{e.location}</div>}
                  </div>
                  <div className="flex items-center gap-1 flex-shrink-0">
                    <button
                      onClick={() => setEditingEvent(e)}
                      className="p-1 text-muted-foreground hover:text-foreground rounded"
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/>
                      </svg>
                    </button>
                    <button
                      onClick={() => deleteEvent(e.id)}
                      className="p-1 text-muted-foreground hover:text-destructive rounded"
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                      </svg>
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>

      {/* Note */}
      <Card className="shadow-sm">
        <CardHeader className="pb-4">
          <CardTitle className="text-primary">Calendar Note</CardTitle>
          <p className="text-xs text-muted-foreground mt-1">
            Optional note printed below the calendar (per-week)
          </p>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex items-center gap-4 flex-wrap">
            <label className="flex items-center gap-2 text-sm cursor-pointer">
              <input
                type="checkbox"
                checked={note.enabled}
                onChange={e => saveNote({ enabled: e.target.checked })}
                className="w-4 h-4 cursor-pointer accent-primary"
              />
              Include in bulletin
            </label>
            <label className={`flex items-center gap-2 text-sm cursor-pointer ${!note.enabled ? 'opacity-50' : ''}`}>
              <input
                type="checkbox"
                checked={note.bold}
                onChange={e => saveNote({ bold: e.target.checked })}
                disabled={!note.enabled}
                className="w-4 h-4 cursor-pointer accent-primary"
              />
              <span className="font-bold">Bold</span>
            </label>
            <label className={`flex items-center gap-2 text-sm cursor-pointer ${!note.enabled ? 'opacity-50' : ''}`}>
              <input
                type="checkbox"
                checked={note.italic}
                onChange={e => saveNote({ italic: e.target.checked })}
                disabled={!note.enabled}
                className="w-4 h-4 cursor-pointer accent-primary"
              />
              <span className="italic">Italic</span>
            </label>
            {/* Justification toggle */}
            <div className={`flex items-center gap-1 ml-auto ${!note.enabled ? 'opacity-50' : ''}`}>
              {(['left', 'center', 'right'] as const).map(a => (
                <button
                  key={a}
                  type="button"
                  onClick={() => saveNote({ align: a })}
                  disabled={!note.enabled}
                  title={`Align ${a}`}
                  className={`p-1.5 rounded border transition-colors ${
                    note.align === a
                      ? 'border-primary bg-primary/10 text-primary'
                      : 'border-input bg-background hover:bg-accent text-muted-foreground'
                  } disabled:cursor-not-allowed`}
                >
                  {a === 'left' && (
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <line x1="21" x2="3" y1="6" y2="6"/><line x1="15" x2="3" y1="12" y2="12"/><line x1="17" x2="3" y1="18" y2="18"/>
                    </svg>
                  )}
                  {a === 'center' && (
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <line x1="21" x2="3" y1="6" y2="6"/><line x1="17" x2="7" y1="12" y2="12"/><line x1="19" x2="5" y1="18" y2="18"/>
                    </svg>
                  )}
                  {a === 'right' && (
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                      <line x1="21" x2="3" y1="6" y2="6"/><line x1="21" x2="9" y1="12" y2="12"/><line x1="21" x2="7" y1="18" y2="18"/>
                    </svg>
                  )}
                </button>
              ))}
            </div>
          </div>
          <textarea
            value={note.text}
            onChange={e => setNote({ ...note, text: e.target.value })}
            onBlur={() => saveNote({})}
            disabled={!note.enabled}
            placeholder='e.g., "Dr. Mellette will be on vacation from April 11–25."'
            className="w-full border border-input rounded-md px-3 py-2 bg-background text-sm min-h-[80px] resize-y disabled:opacity-50"
          />
        </CardContent>
      </Card>

      {/* Template Edit Modal */}
      {editingTemplate && (
        <Modal onClose={() => setEditingTemplate(null)}>
          <h3 className="text-lg font-bold mb-4">
            {editingTemplate.id ? 'Edit Recurring Event' : 'Add Recurring Event'}
          </h3>
          <div className="space-y-3">
            <div>
              <Label>Name</Label>
              <Input
                value={editingTemplate.name}
                onChange={e => setEditingTemplate({ ...editingTemplate, name: e.target.value })}
                placeholder="Sunday Worship"
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Label>Day of Week</Label>
                <select
                  className="w-full border border-input rounded-md px-3 py-2 bg-background text-sm"
                  value={editingTemplate.dayOfWeek}
                  onChange={e => setEditingTemplate({ ...editingTemplate, dayOfWeek: parseInt(e.target.value) })}
                >
                  {DAYS.map((_, i) => <option key={i} value={i}>{DAY_FULL[i]}</option>)}
                </select>
              </div>
              <div>
                <Label>Time</Label>
                <Input
                  value={editingTemplate.time}
                  onChange={e => setEditingTemplate({ ...editingTemplate, time: e.target.value })}
                  placeholder="10:45 AM"
                />
              </div>
            </div>
            <div>
              <Label>Location</Label>
              <Input
                value={editingTemplate.location}
                onChange={e => setEditingTemplate({ ...editingTemplate, location: e.target.value })}
                placeholder="Sanctuary"
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <Button variant="outline" onClick={() => setEditingTemplate(null)}>Cancel</Button>
              <Button onClick={() => saveTemplate(editingTemplate)} disabled={!editingTemplate.name.trim()}>
                Save
              </Button>
            </div>
          </div>
        </Modal>
      )}

      {/* Event Edit Modal */}
      {editingEvent && (
        <Modal onClose={() => setEditingEvent(null)}>
          <h3 className="text-lg font-bold mb-4">
            {editingEvent.id ? 'Edit Event' : 'Add One-off Event'}
          </h3>
          <div className="space-y-3">
            <div>
              <Label>Title</Label>
              <Input
                value={editingEvent.title}
                onChange={e => setEditingEvent({ ...editingEvent, title: e.target.value })}
                placeholder="Mother's Day Brunch"
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Label>Date</Label>
                <Input
                  type="date"
                  value={editingEvent.date}
                  onChange={e => setEditingEvent({ ...editingEvent, date: e.target.value })}
                />
              </div>
              <div>
                <Label>Time</Label>
                <Input
                  value={editingEvent.time}
                  onChange={e => setEditingEvent({ ...editingEvent, time: e.target.value })}
                  placeholder="10:00 AM"
                />
              </div>
            </div>
            <div>
              <Label>Location</Label>
              <Input
                value={editingEvent.location}
                onChange={e => setEditingEvent({ ...editingEvent, location: e.target.value })}
                placeholder="Fellowship Hall"
              />
            </div>
            <div>
              <Label>Description (optional)</Label>
              <Input
                value={editingEvent.description || ''}
                onChange={e => setEditingEvent({ ...editingEvent, description: e.target.value })}
              />
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <Button variant="outline" onClick={() => setEditingEvent(null)}>Cancel</Button>
              <Button
                onClick={() => saveEvent(editingEvent)}
                disabled={!editingEvent.title.trim() || !editingEvent.date}
              >
                Save
              </Button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  )
}

function Modal({ children, onClose }: { children: React.ReactNode; onClose: () => void }) {
  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div
        className="bg-card rounded-lg shadow-xl border border-border w-full max-w-md max-h-[90vh] overflow-y-auto p-6"
        onClick={e => e.stopPropagation()}
      >
        {children}
      </div>
    </div>
  )
}
