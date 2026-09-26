import { useState, useEffect, useRef, type ReactNode } from 'react'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Separator } from '@/components/ui/separator'
import { Switch } from '@/components/ui/switch'
import { HymnPicker } from '@/components/HymnPicker'
import { SettingsPanel } from '@/components/SettingsPanel'
import { CalendarTab } from '@/components/CalendarTab'
import { Toast, type ToastData } from '@/components/Toast'
import { ConfirmDialog } from '@/components/ConfirmDialog'
import { ChoiceDialog } from '@/components/ChoiceDialog'
import { ServiceTitlePicker } from '@/components/ServiceTitlePicker'
import { Tooltip } from '@/components/ui/tooltip'
import { getHealth, saveService, loadService, listServices, uploadHeroImage, downloadUrl, listChurches, getChurchDefaults, generateUrl, listRevisions, revertService, carryHeroImage, type Revision } from '@/lib/api'
import type { OrderOfWorship } from '@/types'
import { emptyOrder, hasContent, serviceFingerprint } from '@/types'
import { switchChoice } from '@/lib/church-switch'
import { hymnalStatus } from '@/lib/hymnal-status'
import type { Church } from '@/types'


function App() {
  const [order, setOrder] = useState<OrderOfWorship>(emptyOrder(''))
  // ⭐⭐ THE FORM SAVES ITSELF. `Save` was a button you had to remember; the calendar
  // tab already autosaved, and this is the same contract. @decision:gold 2026-09-25
  // ⛔ What autosave takes away is the "close without saving" undo, which is why
  // snapshots exist (see `boundarySave`) — the escape is explicit now, not a side
  // effect of not pressing something.
  const [saveState, setSaveState] = useState<'idle' | 'dirty' | 'saving' | 'saved'>('idle')
  const [revisions, setRevisions] = useState<Revision[]>([])
  const [revertOpen, setRevertOpen] = useState(false)
  // ⛔⛔ CLEARING IS DESTRUCTIVE AND MUST ASK. It sat beside "Make Bulletin" — the
  // button pressed every week — so a misclick wiped the form, and the undo was
  // behind clicking the save status, which nobody would find. An action is only
  // "reversible instead of confirmed" when the reversal is DISCOVERABLE.
  // @decision:gold 2026-09-25
  const [confirmClear, setConfirmClear] = useState(false)
  const [pastServices, setPastServices] = useState<{ date: string; filename: string }[]>([])
  const [loadingPast, setLoadingPast] = useState(false)
  const [generating, setGenerating] = useState(false)
  const [generatingSlides, setGeneratingSlides] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [heroPreview, setHeroPreview] = useState<string | null>(null)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [activeTab, setActiveTab] = useState<'service' | 'calendar'>('service')
  // ⭐ The church decides the Word template AND the slide palette, so it travels
  // with every generate call. Remembered because it is the same answer most weeks.
  /** ⭐ The translation list is SERVED, not hardcoded: a church's own API.Bible
   *  key adds to it, so the options are not knowable at build time.
   *  ⚠ Seeded with the four included texts so the picker is never empty, even
   *  before the fetch lands or if the backend is unreachable. */
  const [translations, setTranslations] = useState<{ id: string; name: string; description: string; source?: string }[]>([
    { id: 'BSB', name: 'BSB', description: 'Berean Standard Bible', source: 'included' },
    { id: 'eng_kjv', name: 'KJV', description: 'King James Version', source: 'included' },
    { id: 'eng_asv', name: 'ASV', description: 'American Standard Version', source: 'included' },
    { id: 'ENGWEBP', name: 'WEB', description: 'World English Bible', source: 'included' },
  ])
  // ⛔⛔ BUG-021 — AUTOSAVE IS SUPPRESSED FOR THE DURATION OF A CHURCH SWITCH.
  // `setChurch(B)` re-arms the autosave effect immediately, with B selected and the
  // OUTGOING church's content still in `order`, and `selectChurch` then awaits two
  // calls before replacing the form. The only thing that kept that write from landing
  // was two loopback round trips finishing inside 800ms; miss it and B's saved service
  // is overwritten with A's content while the form shows B's old save — silently, and
  // to the church being arrived at. @decision:gold 2026-09-25
  // ⭐ STATE RATHER THAN A REF ON PURPOSE: a ref would hold the write off and then
  // never re-arm, because nothing would re-run the effect once the switch finished —
  // so a carried-over form would sit unsaved until the next keystroke.
  const [switchingChurch, setSwitchingChurch] = useState(false)

  // ⭐⭐ WAS THIS FORM EMPTIED ON PURPOSE? @decision:gold 2026-09-26 · BUG-022.
  // Clearing is a deliberate emptying and MUST persist — the confirm promises Revert
  // brings the old one back, and the snapshot is what makes that true. An empty form
  // that nobody asked for is the opposite, and must not be written at all. The two are
  // indistinguishable from the payload, so the intent is carried explicitly.
  /** ⭐⭐ THE HYMNAL NOTICE SHOWS ONCE, NOT EVERY LAUNCH. @decision:gold 2026-09-26
   *  A church without a hymnal is in a legitimate, permanent state; repeating the
   *  announcement every morning would be scolding them for a decision they already made.
   *  ⚠ The dismissal is keyed to the FOLDER it was dismissed for, so pointing the app at
   *  a different folder that also turns out to be empty says so again — that is a new
   *  attempt that failed, not the same one being nagged about.
   *  ⛔ The note inside the hymn picker is NOT this and does not stop: it answers "why
   *  did nothing come back", at the moment the question is asked. */
  const [hymnalNotice, setHymnalNotice] = useState<{ dir: string } | null>(null)
  const [emptyOnPurpose, setEmptyOnPurpose] = useState(false)
  // ⛔ BUG-023 — the order object whose save just FAILED. `setSaveState('dirty')` in the
  // failure path re-runs the effect, which re-arms the timer, which fails again: an
  // unbounded 800ms retry loop with no backoff. Holding the failed object lets the
  // status stay honestly "Unsaved" without spinning, and any edit makes a new object
  // and so retries by itself.
  const failedSave = useRef<OrderOfWorship | null>(null)

  /** ⭐⭐ THE CHURCH SWITCH ASKS, EVERY TIME. @decision:gold 2026-09-26 · FEATURE-016
   *  ⛔ And it is not a setting: the same person needs a different answer on different
   *  Sundays, so nobody can answer it in advance.
   *  ⚠ The question is awaited mid-switch, so the resolver is held in a ref and the
   *  switch does not commit the new church until an answer comes back. */
  const [pendingChoice, setPendingChoice] = useState<{
    title: string; body: ReactNode
    options: { id: string; label: string; detail?: string }[]
  } | null>(null)
  const choiceResolver = useRef<((id: string | null) => void) | null>(null)

  function askChoice(spec: NonNullable<typeof pendingChoice>): Promise<string | null> {
    return new Promise(resolve => {
      choiceResolver.current = resolve
      setPendingChoice(spec)
    })
  }

  function answerChoice(id: string | null) {
    setPendingChoice(null)
    const r = choiceResolver.current
    choiceResolver.current = null
    r?.(id)
  }
  const [churches, setChurches] = useState<Church[]>([])
  const [church, setChurch] = useState<string>(() => localStorage.getItem('oo.church') ?? '')

  // ⛔⛔ @decision:gold 2026-09-24 · BUG-010 — generation REFUSES without a church
  // rather than falling through to the neutral default. A deck built with none
  // selected came out wearing the generic palette and the PACKAGED art, which
  // belongs to one congregation, so it looked almost right and the only tell was
  // a color. ⭐ Empty is legitimate only when no church is configured at all.
  const needsChurch = churches.length > 0 && !church

  useEffect(() => {
    listChurches().then(list => {
      setChurches(list)
      // Only auto-pick when there is no ambiguity to resolve.
      if (!church && list.length === 1) selectChurch(list[0].id)
      // ⚠ A remembered id whose profile has since gone must not sit in state
      // looking like a choice — it would be refused at generate time.
      if (church && !list.some(c => c.id === church)) selectChurch('')
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  /** Pull a church's people defaults into any EMPTY field.
   *  ⛔ Never overwrites something already filled in — switching church must not
   *  silently rewrite a service that is part-written. */
  async function applyChurchDefaults(id: string) {
    if (!id) return
    const d = await getChurchDefaults(id)
    if (!d) return
    setOrder(prev => ({
      ...prev,
      speakerShortName: prev.speakerShortName || d.speakerShort,
    }))
  }

  /** Switch church — and bring that church's version of this date with you.
   *
   *  ⛔⛔ THE FORM DOES NOT RELOAD ON ITS OWN, WHICH USED TO BE A SILENT DATA LOSS:
   *  with one service per church, editing one congregation, switching to the other
   *  and generating wrote the first one's content over the second's file. No error, and the
   *  first sign of it was the bulletin. @decision:gold 2026-09-25 · FEATURE-008
   *
   *  ⭐ ONE RULE COVERS BOTH CASES. Save what is being left, then load the incoming
   *  church's service for this date IF IT EXISTS. On an ordinary Sunday it does not,
   *  so the content carries over and generating saves it under the new church —
   *  which is exactly "same service, both congregations", at no cost.
   */
  /** Switch church — saving what is being left, then ASKING what the arriving church's
   *  form should hold.
   *
   *  ⭐⭐ THE SEQUENCE IS FIXED: save the current church first, then ask.
   *  ⛔⛔ AND ONE ANSWER GOVERNS EVERYTHING AFTER IT. The hero image and the speaker
   *  follow that answer; there is no second question about either.
   *  @decision:gold 2026-09-26 · FEATURE-016
   *
   *  ⭐ THE OPTIONS PRUNE THEMSELVES, which is what keeps this from being a nag. With an
   *  empty form and nothing saved there is nothing to decide and it never opens; with a
   *  form that already matches their save, carrying and loading are the same act.
   *
   *  ⚠ THE CHURCH IS NOT COMMITTED UNTIL AN ANSWER COMES BACK, so Cancel leaves
   *  everything exactly as it was — including the picker, which is bound to `church`.
   */
  async function selectChurch(id: string) {
    const leaving = church
    if (id === leaving) return
    // ⛔ The guard goes up FIRST and stays up for the whole switch, question included
    // (BUG-021): the form and the church disagree until this resolves.
    setSwitchingChurch(true)
    try {
      if (leaving && order.date) await boundarySave('before-church-switch', order, leaving)

      const commit = (to: string) => {
        setChurch(to)
        localStorage.setItem('oo.church', to)
      }

      if (!id || !order.date) {
        commit(id)
        await applyChurchDefaults(id)
        return
      }

      const theirs = await loadService(order.date, id).catch(() => null)
      const mine = hasContent(order)
      const sameAlready = theirs != null && serviceFingerprint(order) === serviceFingerprint({ ...emptyOrder(order.date), ...theirs })
      const name = churches.find(c => c.id === id)?.name ?? id

      const load = async () => {
        const t = theirs!
        setOrder({ ...emptyOrder(order.date), ...t })
        setHeroPreview(t.heroImageFilename ? downloadUrl(t.heroImageFilename) + '?t=' + Date.now() : null)
        setSaveState('saved')
        setEmptyOnPurpose(false)
        setToast({ id: Date.now(), message: `Opened ${name}'s service`, detail: order.date })
      }

      const carry = async () => {
        // ⭐ The image is just data, like a hymn pick — but the FILE is named
        // for the other church, so the arriving church gets its own copy or the two
        // would share one picture and neither would know.
        let hero = order.heroImageFilename ?? null
        if (hero) hero = await carryHeroImage(order.date, hero, id)
        const carried = { ...order, heroImageFilename: hero }
        setOrder(carried)
        setHeroPreview(hero ? downloadUrl(hero) + '?t=' + Date.now() : null)
        setEmptyOnPurpose(false)
        // ⚠ Written here rather than left to autosave, and WITH a snapshot reason: this
        // overwrites whatever the arriving church had, so the thing being replaced has
        // to be recoverable from Revert.
        try {
          await saveService(carried, id, theirs ? 'before-church-switch-carry' : undefined)
          setSaveState('saved')
        } catch {
          failedSave.current = carried
          setSaveState('dirty')
        }
        setToast({ id: Date.now(), message: `Saved this service under ${name}`, detail: order.date })
      }

      const clear = async () => {
        const blank = emptyOrder(order.date)
        setOrder(blank)
        setHeroPreview(null)
        setEmptyOnPurpose(true)
        if (theirs) {
          // Snapshot first: they chose to blank a service that existed.
          try {
            await saveService(blank, id, 'before-church-switch-clear', true)
            setSaveState('saved')
          } catch {
            setSaveState('dirty')
          }
        } else {
          setSaveState('idle')
        }
        setToast({ id: Date.now(), message: `Started a blank service for ${name}`, detail: order.date })
      }

      // ⛔ A form that already matches their save makes carrying and loading the same
      // act, so there is nothing to ask. (The empty-and-nothing-saved case is
      // `switchChoice` returning null, below.)
      if (sameAlready) { commit(id); await load(); await applyChurchDefaults(id); return }

      // ⭐ The words live in `lib/church-switch.ts` so the workshop bench reviews THESE
      // strings rather than a copy of them. Here we only draw them.
      const leavingName = leaving ? (churches.find(c => c.id === leaving)?.name ?? leaving) : null
      const spec = switchChoice({ name, leavingName, date: order.date, theirs: !!theirs, mine })
      if (!spec) { commit(id); await applyChurchDefaults(id); return }

      const picked = await askChoice({
        title: spec.title,
        body: <>{spec.lines.map((l, i) => <p key={i} className={i ? 'mt-1' : ''}>{l}</p>)}</>,
        options: spec.options,
      })
      if (picked == null) return          // ⛔ Cancel: the church never moved.

      commit(id)
      if (picked === 'load') await load()
      else if (picked === 'carry') await carry()
      else await clear()

      setPastServices(await listServices(id))
      refreshRevisions(order.date, id)
      await applyChurchDefaults(id)
    } finally {
      // ⚠ FINALLY, or a failed load leaves autosave switched off for the rest of the
      // session and every later edit goes unwritten with the status line saying so.
      setSwitchingChurch(false)
    }
  }

  useEffect(() => {
    hymnalStatus().then(({ count, dir }) => {
      if (count !== 0) return                       // -1 is "could not ask", not "none"
      if (localStorage.getItem('oo.hymnalNoticeFor') === dir) return
      setHymnalNotice({ dir })
    })
  }, [])

  /** ⚠ Dismissing REMEMBERS, whichever way it is dismissed — choosing to go to Settings
   *  counts too, because the next thing that happens there either fixes it or points at
   *  another empty folder, and both answer the notice. */
  function dismissHymnalNotice(goToSettings: boolean) {
    if (hymnalNotice) localStorage.setItem('oo.hymnalNoticeFor', hymnalNotice.dir)
    setHymnalNotice(null)
    if (goToSettings) setSettingsOpen(true)
  }

  useEffect(() => {
    // ⛔ THE CHURCH PREFILL MUST RUN LAST, and that is why this is one awaited
    // chain rather than two effects. loadService's setOrder is a FULL REPLACE, so
    // defaults applied before it resolves are silently wiped — the remembered
    // church would show in the dropdown next to empty fields.
    const remembered = localStorage.getItem('oo.church') ?? ''
    ;(async () => {
      const data = await getHealth().catch(() => null)
      if (!data) return
      setOrder(prev => ({ ...prev, date: data.nextSunday }))
      try {
        const existing = await loadService(data.nextSunday, remembered)
        setOrder({ ...emptyOrder(data.nextSunday), ...existing })
        // ⚠ Say so. Without this the status sits blank after a boot load, which
        // reads as "nothing is saved" on a form that was just restored from disk.
        setSaveState('saved')
        setEmptyOnPurpose(false)   // a loaded form is not a deliberate blank (BUG-022)
        if (existing.heroImageFilename) {
          setHeroPreview(downloadUrl(existing.heroImageFilename) + '?t=' + Date.now())
        }
      } catch {
        // no saved service for that date yet — a blank one is correct
      }
      if (remembered) await applyChurchDefaults(remembered)
      refreshRevisions(data.nextSunday, remembered)
    })()
    listServices(remembered).then(setPastServices)
    refreshTranslations()
  }, [])

  /** ⚠ Re-read after Settings closes: a key added there changes this list, and
   *  nothing else would tell the form about it. */
  function refreshTranslations() {
    fetch('/api/scripture/translations')
      .then(r => r.json())
      .then(list => { if (Array.isArray(list) && list.length) setTranslations(list) })
      .catch(() => {})
  }

  function update<K extends keyof OrderOfWorship>(key: K, value: OrderOfWorship[K]) {
    setOrder(prev => ({ ...prev, [key]: value }))
    setSaveState('dirty')
  }

  /** Save, asking the server to snapshot the CURRENT state first.
   *  ⭐ Used only where work can be lost — before a load, a church switch, or a
   *  generate. ⛔ Never on autosave: a snapshot per keystroke is not a history. */
  /** Returns false if the save did not happen — the caller must not then discard
   *  what is on screen. ⛔ Nothing downstream may assume this succeeded. */
  async function boundarySave(reason: string, o = order, c = church): Promise<boolean> {
    if (!o.date || !c) return true
    // ⛔⛔ AN EMPTY FORM HAS NOTHING TO PRESERVE, SO PRESERVING IT DESTROYS SOMETHING.
    // BUG-022: this fires before a load and before a church switch, and on a page that
    // has not finished loading the form IS empty — so it wrote `emptyOrder` over a real
    // service. ⭐ Returning true is correct rather than lenient: the caller asks "is it
    // safe to move on", and with nothing in the form there is nothing at risk.
    // ⚠ A deliberate clear is the exception and says so.
    if (!hasContent(o) && !emptyOnPurpose) return true
    try {
      await saveService(o, c, reason, !hasContent(o))
      setSaveState('saved')
      setEmptyOnPurpose(false)   // a loaded form is not a deliberate blank (BUG-022)
      failedSave.current = null
      return true
    } catch (e: any) {
      failedSave.current = o
      setSaveState('dirty')
      setErrorMsg(e?.message ?? 'Could not save this service')
      return false
    }
  }

  // ⭐⭐ AUTOSAVE. Debounced so a burst of typing is one write, and gated on a
  // church — a service has nowhere to go without one (same gate as generate).
  useEffect(() => {
    // ⛔ Never mid-switch (BUG-021): `church` has already moved and `order` has not.
    if (switchingChurch) return
    if (saveState !== 'dirty' || !order.date || !church) return
    // ⛔ Never write an empty form nobody asked to empty (BUG-022).
    if (!hasContent(order) && !emptyOnPurpose) return
    // ⛔ And never re-arm for the payload that just failed (BUG-023) — the failure path
    // sets `dirty`, which lands right back here. The status stays honest; an edit makes
    // a new object, so typing is what retries.
    if (failedSave.current === order) return
    const t = setTimeout(async () => {
      setSaveState('saving')
      try {
        await saveService(order, church, undefined, !hasContent(order))
        setSaveState('saved')
        setEmptyOnPurpose(false)   // a loaded form is not a deliberate blank (BUG-022)
        failedSave.current = null
        setPastServices(await listServices(church))
      } catch {
        failedSave.current = order
        setSaveState('dirty')     // ⛔ never claim saved on a failed write
      }
    }, 800)
    return () => clearTimeout(t)
  }, [order, church, saveState, switchingChurch, emptyOnPurpose])

  async function refreshRevisions(date = order.date, c = church) {
    if (!date || !c) return setRevisions([])
    setRevisions(await listRevisions(date, c))
  }

  async function handleLoadDate(date: string, c = church) {
    setLoadingPast(true)
    try {
      if (!await boundarySave('before-load')) return   // keep what is on screen
      const data = await loadService(date, c)
      setOrder({ ...emptyOrder(date), ...data })
      setHeroPreview(data.heroImageFilename
        ? downloadUrl(data.heroImageFilename) + '?t=' + Date.now() : null)
      setSaveState('saved')
      setEmptyOnPurpose(false)   // a loaded form is not a deliberate blank (BUG-022)
    } catch {
      // ⚠ No saved service for that date — a blank form is the right answer, and it
      // is genuinely unsaved until something is typed into it.
      setOrder(emptyOrder(date))
      setHeroPreview(null)
      setSaveState('saved')
      setEmptyOnPurpose(false)   // a loaded form is not a deliberate blank (BUG-022)
    } finally {
      // ⛔ NOT setSaveState HERE. `finally` runs on the early return too, so setting
      // "saved" from here reported a successful save after one had just failed —
      // the exact lie this whole pass is removing.
      setLoadingPast(false)
      refreshRevisions(date, c)
    }
  }

  /** Empty the form for a fresh service, keeping the date and the congregation.
   *
   *  ⭐⭐ AUTOSAVE IS WHY THIS HAS TO EXIST. Starting over used to be "close without
   *  saving"; with the form saving itself there is no way to walk away from what is
   *  in it. @decision:gold 2026-09-25
   *
   *  ⛔ NO CONFIRM DIALOG. It snapshots first, so the answer to "are you sure" is
   *  Revert — and the toast says so at the one moment it is useful. A prompt on a
   *  reversible action is friction pretending to be safety.
   */
  async function handleClear() {
    if (!order.date) return
    // ⛔ The dialog promises the current service is already saved. If it is not,
    // that promise is void and the form is NOT emptied.
    if (!await boundarySave('before-clear')) return
    setOrder(emptyOrder(order.date))
    setHeroPreview(null)
    setEmptyOnPurpose(true)          // ⭐ this emptiness is the point (BUG-022)
    setSaveState('dirty')            // autosave persists the cleared form
    if (church) await applyChurchDefaults(church)
    refreshRevisions()
    setToast({ id: Date.now(), message: 'Form cleared', detail: 'Revert… brings it back' })
  }

  async function handleRevert(id: string) {
    if (!order.date || !church) return
    try {
      const data = await revertService(order.date, church, id)
      setOrder({ ...emptyOrder(order.date), ...data })
      setHeroPreview(data.heroImageFilename
        ? downloadUrl(data.heroImageFilename) + '?t=' + Date.now() : null)
      setSaveState('saved')
      setEmptyOnPurpose(false)   // a loaded form is not a deliberate blank (BUG-022)
      setRevertOpen(false)
      setToast({ id: Date.now(), message: 'Reverted', detail: 'The previous state was kept too' })
      refreshRevisions()
    } catch {
      setErrorMsg('Could not revert to that revision')
    }
  }

  async function handleHeroImage(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file || !order.date) return
    const filename = await uploadHeroImage(order.date, file, church)
    update('heroImageFilename', filename)
    setHeroPreview(URL.createObjectURL(file))
    // Reset input so re-selecting the same file triggers onChange
    e.target.value = ''
  }

  const [errorMsg, setErrorMsg] = useState<string | null>(null)
  /** ⛔ NOT errorMsg. A substitution is not a failure — the deck is complete and
   *  correct, just not in the translation that was asked for. Dressing that in
   *  destructive red would teach people to distrust a document that is fine. */
  const [noticeMsg, setNoticeMsg] = useState<string | null>(null)
  const [toast, setToast] = useState<ToastData | null>(null)

  // The desktop app has no download bar, so a generated file lands silently.
  // This is the receipt: what was made, and where it went.
  function confirmSaved(kind: string, data: { filename: string; folder?: string }) {
    setToast({ id: Date.now(), message: `${kind} saved`, detail: data.folder || data.filename })
  }

  /** ⛔⛔ A SUBSTITUTION MUST NOT PASS AS AN ORDINARY SUCCESS. The deck is
   *  complete and the badge reads BSB, which is honest — but nobody inspects a
   *  badge after they have walked away from the computer, and the translation
   *  they chose is not the one they got. Say it where they are already looking. */
  function reportFallback(fb?: { wanted: string; why: string; tail: string }) {
    if (!fb) return
    setNoticeMsg(`${fb.why} ${fb.tail}`)
  }
  const [scripturePreview, setScripturePreview] = useState<{
    verses: { number: number; text: string }[]
    translation_name: string
    /** What the translation actually returned — may be wider than asked. */
    reference?: string
    /** What was typed into the field. */
    requested?: string
    /** Set when the chosen translation could not be fetched and BSB stood in. */
    fallback?: { wanted: string; why: string; tail: string }
    slides?: unknown[]
  } | null>(null)
  const [loadingScripture, setLoadingScripture] = useState(false)
  const scriptureTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  /** The reference the preview last ran for — decides debounce vs immediate. */
  const lastPreviewRef = useRef('')
  /** Bumped when something outside the order may change the answer. */
  const [previewNonce, setPreviewNonce] = useState(0)

  async function fetchScripturePreview(ref: string, translation: string) {
    if (!ref.trim()) {
      setScripturePreview(null)
      return
    }
    setLoadingScripture(true)
    try {
      const res = await fetch(`/api/scripture/fetch?ref=${encodeURIComponent(ref)}&translation=${encodeURIComponent(translation)}`)
      if (res.ok) {
        const data = await res.json()
        setScripturePreview(data)
      } else {
        setScripturePreview(null)
      }
    } catch {
      setScripturePreview(null)
    } finally {
      setLoadingScripture(false)
    }
  }

  function handleScriptureChange(ref: string) {
    update('scripture', ref)
  }

  function handleTranslationChange(translation: string) {
    update('scriptureTranslation', translation)
  }

  /**
   * ⭐⭐ THE PREVIEW FOLLOWS THE VALUES, NOT THE KEYSTROKES. @decision:gold 2026-09-25
   * It used to be fetched by the two change handlers, so it only ever appeared
   * for someone who had just TYPED — a reloaded page, a service opened from the
   * list, a church switch or a revert all showed a filled-in reference and a
   * translation with no verses under them. Watching the values covers every one
   * of those paths, including ones nobody has written yet.
   *
   * ⚠ TYPING IS DEBOUNCED, PICKING IS NOT. Six hundred milliseconds is right for
   * a reference being typed a character at a time and wrong for a dropdown,
   * where it reads as lag. The delay applies only when the reference itself
   * changed.
   *
   * ⚠ `previewNonce` forces a re-read when nothing in the order changed but the
   * answer might have — adding or removing an API.Bible key, most obviously.
   */
  useEffect(() => {
    const ref = (order.scripture || '').trim()
    if (!ref) { setScripturePreview(null); return }
    const typed = ref !== lastPreviewRef.current
    lastPreviewRef.current = ref
    if (scriptureTimerRef.current) clearTimeout(scriptureTimerRef.current)
    scriptureTimerRef.current = setTimeout(
      () => fetchScripturePreview(ref, order.scriptureTranslation),
      typed ? 600 : 0,
    )
    return () => { if (scriptureTimerRef.current) clearTimeout(scriptureTimerRef.current) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [order.scripture, order.scriptureTranslation, previewNonce])

  async function handleGenerate() {
    if (!order.date) return
    setGenerating(true)
    setErrorMsg(null)
    try {
      await boundarySave('before-generate')
      const res = await fetch(generateUrl('bulletin', order.date, church), { method: 'POST' })
      if (res.ok) {
        const data = await res.json()
        // No window.open: the server has already written the file to the
        // output folder. In the desktop app that call did nothing; in a
        // browser it downloaded a second copy alongside the real one.
        confirmSaved('Bulletin', data)
      } else {
        const err = await res.json()
        setErrorMsg(err.detail || 'Failed to generate bulletin')
      }
    } catch (e: any) {
      console.error('Generate bulletin error:', e)
      setErrorMsg(`Could not reach the server: ${e.message || e}`)
    } finally {
      setGenerating(false)
    }
  }

  async function handleGenerateSlides() {
    if (!order.date) return
    setGeneratingSlides(true)
    setErrorMsg(null)
    setNoticeMsg(null)
    try {
      await boundarySave('before-generate')
      const res = await fetch(generateUrl('slides', order.date, church), { method: 'POST' })
      if (res.ok) {
        const data = await res.json()
        confirmSaved('Presentation', data)
        reportFallback(data.fallback)
      } else {
        const err = await res.json()
        setErrorMsg(err.detail || 'Failed to generate slides')
      }
    } catch (e: any) {
      console.error('Generate slides error:', e)
      setErrorMsg(`Could not reach the server: ${e.message || e}`)
    } finally {
      setGeneratingSlides(false)
    }
  }

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="border-b border-border bg-card">
        <div className="max-w-3xl mx-auto px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <img src="/openorder-logo.svg" alt="OpenOrder" className="h-10 w-auto" />
            <div className="text-center">
              <h1 className="text-2xl font-bold tracking-tight leading-none" style={{ fontFamily: 'Georgia, serif', letterSpacing: '-0.5px' }}>
                <span style={{ color: '#4A90D9' }}>Open</span>
                <span style={{ color: '#F5A623' }}>Order</span>
              </h1>
              <p className="text-[10px] tracking-[3px] text-muted-foreground/60 uppercase font-medium mt-2 leading-none">
                Worship. Simplified.
              </p>
            </div>
          </div>
        </div>
      </header>

      <div className="max-w-3xl mx-auto pt-6 px-4 pb-16">
        {/* Tab navigation
            ⭐⭐ STICKY, AND IT CARRIES THE SAVE STATUS. A status you have to go
            looking for is not reassurance, and this form is long enough that the
            bottom is a scroll away. @decision:gold 2026-09-25
            ⚠ `top-0` works because the page header above is not itself sticky; if
            that ever changes this needs its height as the offset. */}
        <div className="sticky top-0 z-30 -mx-4 px-4 bg-background/95 backdrop-blur
                        flex items-center border-b border-border mb-6">
          <button
            onClick={() => setActiveTab('service')}
            className={`relative px-5 py-2.5 text-sm font-semibold transition-colors -mb-px ${
              activeTab === 'service'
                ? 'text-primary border-b-2 border-primary'
                : 'text-muted-foreground hover:text-foreground border-b-2 border-transparent'
            }`}
          >
            Service
          </button>
          <button
            onClick={() => setActiveTab('calendar')}
            className={`relative px-5 py-2.5 text-sm font-semibold transition-colors -mb-px ${
              activeTab === 'calendar'
                ? 'text-primary border-b-2 border-primary'
                : 'text-muted-foreground hover:text-foreground border-b-2 border-transparent'
            }`}
          >
            Calendar
          </button>

          <div className="ml-auto">
          {/* ⭐⭐ THE SAVE STATUS *IS* THE DOOR TO THE HISTORY — the Google Docs
              pattern. An unlabelled ⋯ told a non-technical reader nothing ("looks
              more like an error than something useful"), and the undo belongs where
              someone looks when they are worried about their work: the words telling
              them whether it is saved. @decision:gold 2026-09-25 */}
          <div className="relative">
            <button
              type="button"
              onClick={async () => {
                if (!order.date || needsChurch) return
                await refreshRevisions(); setRevertOpen(o => !o)
              }}
              disabled={!order.date || needsChurch}
              className="text-sm text-muted-foreground flex items-center gap-1.5 whitespace-nowrap
                         rounded px-2 py-1 -mx-2 hover:bg-accent hover:text-foreground
                         disabled:hover:bg-transparent disabled:cursor-default transition-colors"
              aria-live="polite"
            >
              {needsChurch ? (
                <><span className="size-1.5 rounded-full bg-amber-500" />Choose a church</>
              ) : saveState === 'saving' ? (
                <>Saving…</>
              ) : saveState === 'dirty' ? (
                <><span className="size-1.5 rounded-full bg-amber-500" />Unsaved changes</>
              ) : (
                <>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                       strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" className="text-muted-foreground/70">
                    <path d="M20 6 9 17l-5-5" />
                  </svg>
                  Saved
                </>
              )}
            </button>
            {revertOpen && (
              <div className="absolute left-0 top-full mt-1 z-20 w-80 rounded-md border border-input bg-background shadow-lg p-1">
                <p className="px-3 pt-2 pb-1 text-xs font-medium text-muted-foreground">
                  Earlier versions
                </p>
                {revisions.length === 0 ? (
                  <p className="text-sm text-muted-foreground px-3 pb-2">
                    None yet — one is kept each time you open a service, switch church,
                    or make a document.
                  </p>
                ) : revisions.map(r => (
                  <button
                    key={r.id}
                    onClick={() => handleRevert(r.id)}
                    className="w-full text-left px-3 py-2 text-sm rounded hover:bg-accent"
                  >
                    {new Date(r.at).toLocaleString(undefined,
                      { weekday: 'short', hour: 'numeric', minute: '2-digit' })}
                    <span className="block text-xs text-muted-foreground">
                      {({ 'before-load': 'before opening another service',
                          'before-church-switch': 'before switching church',
                          'before-generate': 'before making a document',
                          'before-clear': 'before starting a new service',
                          'before-revert': 'before the last undo',
                        } as Record<string, string>)[r.reason] ?? r.reason.replace(/-/g, ' ')}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>

          </div>

          {/* ⭐ SETTINGS RIDES ALONG. Everything that is always available now lives
              in the one bar that is always on screen, rather than being split
              between a header that scrolls and a strip that does not. */}
          <button
            onClick={() => setSettingsOpen(true)}
            className="ml-1 p-2 rounded-md hover:bg-accent text-muted-foreground hover:text-foreground transition-colors"
            title="Settings"
            aria-label="Settings"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/>
              <circle cx="12" cy="12" r="3"/>
            </svg>
          </button>
        </div>

        {activeTab === 'calendar' && (
          <CalendarTab key={church} serviceDate={order.date} church={church} />
        )}

        {activeTab === 'service' && <>
        {/* Service Information */}
        <Card className="mb-6 shadow-sm">
          <CardHeader className="pb-4">
            <CardTitle className="text-primary">Service Information</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            {/* ⭐ "WHICH SERVICE AM I WORKING ON" LIVES TOGETHER, AT THE TOP.
                Church, the date, and opening a past one are the same question; Load
                was sitting in the row of verbs at the bottom pretending to be an
                action. @decision:gold 2026-09-25 */}
            {churches.length > 1 && (
              <div className="flex items-end gap-3">
                <div className="flex-1">
                <Label htmlFor="church" className="mb-1.5 block">Church</Label>
                <Tooltip
                  className="w-full"
                  content="Which congregation this service is for. Switching brings that church's version of this date with it."
                >
                <select
                  id="church"
                  className="w-full h-9 rounded-md border border-input bg-transparent px-3 py-1 text-sm shadow-xs"
                  value={church}
                  onChange={e => selectChurch(e.target.value)}
                >
                  {/* ⛔ THIS READ "Default" AND THAT IS WHAT MADE IT DANGEROUS —
                      it looked like a legitimate third option rather than an
                      unset field, so a deck went out under the generic profile
                      without anyone choosing it. */}
                  <option value="">— choose a church —</option>
                  {churches.map(c => (
                    <option key={c.id} value={c.id}>{c.name}</option>
                  ))}
                </select>
                </Tooltip>
                </div>
                <Button
                  variant="outline"
                  className="h-9 font-normal text-muted-foreground"
                  disabled={!order.date || needsChurch}
                  onClick={() => setConfirmClear(true)}
                >
                  New service
                </Button>
                <select
                  aria-label="Open a saved service"
                  className="h-9 rounded-md border border-input bg-transparent px-3 pr-8 text-sm text-muted-foreground shadow-xs cursor-pointer hover:bg-accent transition-colors"
                  value=""
                  onChange={e => { if (e.target.value) handleLoadDate(e.target.value) }}
                  disabled={loadingPast}
                >
                  <option value="">
                    {pastServices.length > 0 ? 'Open a past service…'
                      : !church ? 'Choose a church first'
                      : `No services for ${churches.find(c => c.id === church)?.name ?? 'this church'}`}
                  </option>
                  {pastServices.map(sv => (
                    <option key={sv.date} value={sv.date}>{sv.date}</option>
                  ))}
                </select>
              </div>
            )}
            <div className="grid grid-cols-2 gap-4">
              <div>
                <Label htmlFor="date" className="mb-1.5 block">Date</Label>
                <Input
                  id="date"
                  type="date"
                  value={order.date}
                  onChange={e => {
                    update('date', e.target.value)
                    handleLoadDate(e.target.value)
                  }}
                />
              </div>
              <div>
                <ServiceTitlePicker
                  value={order.serviceTitle}
                  onChange={v => update('serviceTitle', v)}
                  date={order.date}
                  church={church || null}
                />
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Hero Image */}
        <Card className="mb-6 shadow-sm">
          <CardHeader className="pb-4">
            <CardTitle className="text-primary">Hero Image</CardTitle>
          </CardHeader>
          <CardContent>
            <div
              className="border-2 border-dashed border-border rounded-lg p-6 text-center cursor-pointer hover:border-primary/50 transition-colors"
              onClick={() => fileInputRef.current?.click()}
            >
              {heroPreview ? (
                <img src={heroPreview} alt="Hero image" className="max-h-48 mx-auto rounded" />
              ) : (
                <div className="text-muted-foreground">
                  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" className="w-10 h-10 mx-auto mb-2 opacity-50">
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
                    <circle cx="8.5" cy="8.5" r="1.5"/>
                    <polyline points="21 15 16 10 5 21"/>
                  </svg>
                  <p className="text-sm">Click to upload this week's hero image</p>
                </div>
              )}
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                className="hidden"
                onChange={handleHeroImage}
              />
            </div>
          </CardContent>
        </Card>

        {/* Hymns */}
        <Card className="mb-6 shadow-sm">
          <CardHeader className="pb-4">
            <CardTitle className="text-primary">Hymns</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4 pb-8">
            <HymnPicker
              label="Opening Hymn"
              value={order.praiseHymn1}
              onChange={v => update('praiseHymn1', v)}
            />
            <HymnPicker
              label="Offertory Hymn"
              value={order.praiseHymn2}
              onChange={v => update('praiseHymn2', v)}
            />
            <Separator />
            <HymnPicker
              label="Doxology"
              value={order.doxology}
              onChange={v => update('doxology', v)}
              hint="Typically UMH #94 or #95"
            />
            <HymnPicker
              label="Creed"
              value={order.creed}
              onChange={v => update('creed', v)}
              hint="Creeds & Affirmations: UMH #880–889"
            />
            <Separator />
            <HymnPicker
              label="Prayer Hymn"
              value={order.prayerHymn}
              onChange={v => update('prayerHymn', v)}
            />
            <HymnPicker
              label="Liturgical Prayer"
              value={order.liturgicalPrayer}
              onChange={v => update('liturgicalPrayer', v)}
              hint="The Lord's Prayer: UMH #894–896"
            />
            <HymnPicker
              label="Closing Hymn"
              value={order.closingHymn}
              onChange={v => update('closingHymn', v)}
            />
          </CardContent>
        </Card>

        {/* Word and Table */}
        <Card className="mb-6 shadow-sm">
          <CardHeader className="pb-4">
            <CardTitle className="text-primary">Word and Table</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="grid grid-cols-3 gap-4">
              <div>
                <Label htmlFor="scripture" className="mb-1.5 block">Scripture</Label>
                <Input
                  id="scripture"
                  placeholder="e.g., Matthew 4:1-11"
                  value={order.scripture}
                  onChange={e => handleScriptureChange(e.target.value)}
                />
              </div>
              <div>
                <Label htmlFor="translation" className="mb-1.5 block">Translation</Label>
                <select
                  id="translation"
                  className="flex h-9 w-full rounded-md border border-input bg-background px-3 py-1 text-sm shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring"
                  value={order.scriptureTranslation}
                  onChange={e => handleTranslationChange(e.target.value)}
                >
                  {translations.map(t => (
                    <option key={t.id} value={t.id}>
                      {t.name}{t.description ? ` — ${t.description}` : ''}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <Label htmlFor="speakerShortName" className="mb-1.5 block">Speaker</Label>
                <Input
                  id="speakerShortName"
                  placeholder="e.g., Dr. Smith"
                  value={order.speakerShortName}
                  onChange={e => update('speakerShortName', e.target.value)}
                />
              </div>
            </div>

            {/* Scripture preview */}
            {loadingScripture && (
              <div className="text-sm text-muted-foreground italic">Fetching scripture...</div>
            )}
            {scripturePreview && !loadingScripture && (
              <div className="bg-muted/50 rounded-lg p-3 text-sm max-h-48 overflow-y-auto border border-border">
                {/* ⛔⛔ THE SLIDE COUNT COMES FROM THE SERVER, NOT FROM DIVIDING
                    THE VERSE COUNT BY TWO. A paraphrase returns a whole
                    paragraph as ONE verse, so the old arithmetic said "1 slide"
                    for something that generates seven. */}
                <div className="font-medium text-xs text-muted-foreground mb-2">
                  {scripturePreview.translation_name} — {scripturePreview.verses.length}{' '}
                  {scripturePreview.verses.length === 1 ? 'passage' : 'verses'}, will generate{' '}
                  {scripturePreview.slides?.length ?? Math.ceil(scripturePreview.verses.length / 2)} slides
                </div>

                {/* ⛔⛔ THE CHOSEN TRANSLATION COULD NOT BE FETCHED AND THE
                    BUNDLED ONE STOOD IN. Saying nothing would be a quiet lie —
                    the deck would be in a translation nobody picked. Saying it
                    and NOT substituting would be worse: a deck with no reading
                    in it, found on a Sunday. So: both. */}
                {scripturePreview.fallback && (
                  <div className="mb-2 rounded-md border border-amber-500/40 bg-amber-500/10 px-2.5 py-2 text-xs">
                    <span className="font-medium">{scripturePreview.fallback.why}</span>{' '}
                    {scripturePreview.fallback.tail}
                  </div>
                )}

                {/* ⭐⭐ SAY SO WHEN THE TRANSLATION GAVE US MORE THAN WAS ASKED.
                    Some translations — The Message especially — are written in
                    paragraphs and have no verse-level divisions, so the smallest
                    block containing the reference is what comes back. Finding
                    that out from the projector on Sunday is the wrong moment. */}
                {scripturePreview.requested
                  && scripturePreview.reference
                  && scripturePreview.reference !== scripturePreview.requested && (
                  <div className="mb-2 rounded-md border border-amber-500/40 bg-amber-500/10 px-2.5 py-2 text-xs">
                    <span className="font-medium">
                      {scripturePreview.translation_name} doesn't divide these verses.
                    </span>{' '}
                    You asked for {scripturePreview.requested} and this is the smallest
                    portion it has — {scripturePreview.reference}. The slides will show all
                    of it; you can trim them afterward.
                  </div>
                )}
                {scripturePreview.verses.map((v: {number: number, text: string}) => (
                  <p key={v.number} className="mb-1">
                    <span className="font-bold text-primary">{v.number}</span>{' '}
                    {v.text}
                  </p>
                ))}
              </div>
            )}
            <div>
              <Label htmlFor="sermonTitle" className="mb-1.5 block">Sermon Title</Label>
              <Input
                id="sermonTitle"
                placeholder="e.g., The Shadow Mission"
                value={order.sermonTitle}
                onChange={e => update('sermonTitle', e.target.value)}
              />
            </div>
            <div>
              <Label htmlFor="sermonSubtitle" className="mb-1.5 block">Subtitle (optional)</Label>
              <Input
                id="sermonSubtitle"
                placeholder="e.g., (Part 2)"
                value={order.sermonSubtitle}
                onChange={e => update('sermonSubtitle', e.target.value)}
              />
            </div>

            {/* The Table closes this block, the way it closes the movement. */}
            <Switch
              id="communion"
              label="Communion Sunday"
              hint="Adds a Holy Communion slide after the sermon, and uses this church's communion bulletin template when it has one."
              checked={order.communion}
              onCheckedChange={v => update('communion', v)}
            />
          </CardContent>
        </Card>

        {/* Actions
            ⭐⭐ ONE HIERARCHY, THREE TIERS — the two documents ARE the product, so
            they are the only filled buttons and they sit far right where a primary
            action is looked for. "Clear" and "Revert" are housekeeping and drop to a
            quiet menu; the save status is information, not a control, so it sits
            left and subdued. ⛔ Five equal-weight controls in a row made the thing
            the app exists to do look exactly as important as emptying the form.
            @decision:gold 2026-09-25 */}
        <div className="flex items-center gap-3 pt-2">
          <div className="ml-auto flex items-center gap-2">
            <Button
              onClick={handleGenerate}
              disabled={generating || !order.date || needsChurch}
              title={needsChurch ? 'Choose a church first' : undefined}
              className="px-6"
            >
              {generating ? 'Making…' : 'Make Bulletin'}
            </Button>

            <Button
              onClick={handleGenerateSlides}
              disabled={generatingSlides || !order.date || needsChurch}
              title={needsChurch ? 'Choose a church first' : undefined}
              className="px-6"
            >
              {generatingSlides ? 'Making…' : 'Make Slides'}
            </Button>
          </div>
        </div>

        {errorMsg && (
          <div className="mt-4 p-3 rounded-lg bg-destructive/10 text-destructive text-sm flex items-center justify-between">
            <span>{errorMsg}</span>
            <button onClick={() => setErrorMsg(null)} className="ml-4 hover:opacity-70">✕</button>
          </div>
        )}
        {noticeMsg && (
          <div className="mt-4 p-3 rounded-lg border border-amber-500/40 bg-amber-500/10 text-sm flex items-start justify-between">
            <span>{noticeMsg}</span>
            <button onClick={() => setNoticeMsg(null)} className="ml-4 hover:opacity-70 shrink-0">✕</button>
          </div>
        )}
        </>}
      </div>

      {/* ⚠ ONLY IN A PACKAGED BUILD. __APP_BUILD__ compiles to "" under the dev
          server (vite.config.ts), so this footer is absent there rather than
          showing a stale counter.
          ⭐ THE FOOTER, NOT THE HEADER. This is a shippable app, so the number
          belongs out of the way at the bottom (or an About), never beside the
          wordmark. @decision:gold 2026-09-24 */}
      {__APP_BUILD__ && (
        <footer className="max-w-3xl mx-auto px-4 pb-6 pt-2 text-center">
          <p className="text-[10px] text-muted-foreground/40 tabular-nums tracking-wide">
            OpenOrder v{__APP_VERSION__} ({__APP_BUILD__})
          </p>
        </footer>
      )}

      <ConfirmDialog
        open={confirmClear}
        title="Start a new service?"
        body="This empties the form. Your current service is already saved."
        confirmLabel="Start new"
        onConfirm={() => { setConfirmClear(false); handleClear() }}
        onCancel={() => setConfirmClear(false)}
      />
      <SettingsPanel
        open={settingsOpen}
        onClose={() => {
          setSettingsOpen(false)
          refreshTranslations()
          setPreviewNonce(n => n + 1)
        }}
        onSaved={confirmSaved}
        church={church}
      />
      {/* ⛔⛔ A MISSING HYMNAL IS NOT AN ERROR. It is copyrighted, so the app cannot ship
          one and a fresh install genuinely has none. What the notice exists to stop is
          the SILENCE: an empty hymn search reads as "no such hymn", so the app looked
          broken when it was unconfigured. @decision:gold 2026-09-26 · BUG-024 */}
      <ChoiceDialog
        open={!!hymnalNotice}
        title="No hymnal is set up"
        body={
          <>
            {/* ⛔⛔ NAME THE TWO BOOKS. "Point it at your own copy" invited somebody with a
                different hymnal to go looking for an importer that does not exist — the app
                is United Methodist by construction, and the numbers in every hint are UMH
                page numbers. @decision:gold 2026-09-26 */}
            <p>
              OpenOrder works with the <span className="text-foreground">United Methodist
              Hymnal</span> and <span className="text-foreground">The Faith We Sing</span>.
              It cannot import another hymnal.
            </p>
            {/* ⚠ AND IT IS NOT ONLY HYMNS. The Creed, the Doxology and the Lord's Prayer are
                the same picker over the same index, so saying "only hymn search" was false. */}
            <p className="mt-1">
              You can still build a service without it — scripture, the calendar, the sermon
              and the artwork all work, and the bulletin and slides still generate. What you
              cannot choose is anything from the book: hymns, the doxology, creeds and the
              Lord's Prayer.
            </p>
          </>
        }
        options={[
          { id: 'settings', label: 'Choose the hymnal folder…',
            detail: 'Settings → Folders → Hymnal Folder.',
            onPick: () => dismissHymnalNotice(true) },
          { id: 'without', label: 'Continue without hymns',
            detail: 'This notice will not show again.',
            onPick: () => dismissHymnalNotice(false) },
        ]}
        onCancel={() => dismissHymnalNotice(false)}
      />
      <ChoiceDialog
        open={!!pendingChoice}
        title={pendingChoice?.title ?? ''}
        body={pendingChoice?.body}
        options={(pendingChoice?.options ?? []).map(o => ({ ...o, onPick: () => answerChoice(o.id) }))}
        onCancel={() => answerChoice(null)}
      />
      <Toast toast={toast} onDone={() => setToast(null)} />
    </div>
  )
}

export default App
