import type { ReactNode } from 'react'
import { Button } from '@/components/ui/button'

/**
 * A modal that asks WHICH of several things to do — not whether to do one.
 *
 * ⭐⭐ WHY THIS EXISTS RATHER THAN A SETTING. @decision:gold 2026-09-26 · FEATURE-016.
 * Switching church can mean "same service in both congregations, with a tweak" or "two
 * genuinely different services", and which one is true is a fact about that week. His
 * ruling: *"I don't think there's a universal setting here"* — so it is asked at the
 * moment, every time, because nobody can answer it in advance, including him.
 *
 * ⛔ NOT `ConfirmDialog` WITH THREE BUTTONS. That one's shape is safe-vs-destructive,
 * and it puts the destructive choice where the primary action is looked for. Here there
 * is no single safe answer, so the options are listed as peers and read top to bottom.
 *
 * ⚠ THE OPTION SET IS BUILT BY THE CALLER AND IS OFTEN SHORTER THAN THE FULL LIST.
 * With an empty form and nothing saved there is nothing to ask at all, so this never
 * opens — a prompt that appears when it has no question is how a prompt becomes noise.
 *
 * CALLED BY: App.tsx, on a church switch.
 */
export function ChoiceDialog({
  open,
  title,
  body,
  options,
  onCancel,
}: {
  open: boolean
  title: string
  body?: ReactNode
  /** In the order they should be read. `detail` is the consequence in plain words. */
  options: { id: string; label: string; detail?: string; onPick: () => void }[]
  /** ⛔ Cancel must leave the church UNCHANGED — see App.tsx. */
  onCancel: () => void
}) {
  if (!open) return null
  return (
    <div
      className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4"
      onClick={onCancel}
      role="dialog"
      aria-modal="true"
      aria-label={title}
    >
      <div
        className="bg-card rounded-lg shadow-xl border border-border w-full max-w-md p-6"
        onClick={e => e.stopPropagation()}
      >
        <h2 className="text-lg font-semibold">{title}</h2>
        {body && <div className="mt-2 text-sm text-muted-foreground">{body}</div>}
        <div className="mt-4 space-y-2">
          {options.map(o => (
            <button
              key={o.id}
              type="button"
              onClick={o.onPick}
              className="w-full text-left rounded-md border border-input bg-background px-3 py-2 hover:bg-accent transition-colors"
            >
              <div className="text-sm font-medium text-foreground">{o.label}</div>
              {o.detail && <div className="text-xs text-muted-foreground mt-0.5">{o.detail}</div>}
            </button>
          ))}
        </div>
        <div className="mt-4 flex justify-end">
          <Button variant="ghost" onClick={onCancel}>Cancel</Button>
        </div>
      </div>
    </div>
  )
}
