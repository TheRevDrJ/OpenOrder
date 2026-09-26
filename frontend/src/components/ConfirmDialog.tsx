import type { ReactNode } from 'react'
import { Button } from '@/components/ui/button'

/**
 * A yes/no modal for an action that destroys something on screen.
 *
 * ⭐⭐ WHY A MODAL AND NOT AN INLINE PROMPT. @decision:gold 2026-09-25 · Clearing the
 * form was tried as a quiet inline question beside the control, and it read as a
 * puzzle rather than a warning: the reader had to work out which of two verbs
 * matched which outcome. A destructive action deserves the interruption.
 *
 * ⭐ IT SAYS WHAT SURVIVES, NOT JUST WHAT GOES. "Are you sure?" makes a person
 * guess at the stakes; naming the way back is what actually lets them answer.
 *
 * ⛔ NOT `window.confirm` — browser chrome is absent or broken in the
 * packaged app, and it cannot be styled or worded. This matches the overlay already
 * used by the calendar and settings panels rather than introducing a third look.
 */
export function ConfirmDialog({
  open,
  title,
  body,
  confirmLabel,
  onConfirm,
  onCancel,
}: {
  open: boolean
  title: string
  /** ⭐ A node, not just a string: the template-replace confirm has to show a tag
   *  diff as a list, and flattening that into one sentence is how a warning stops
   *  being readable. A plain string still works exactly as before. */
  body?: ReactNode
  confirmLabel: string
  onConfirm: () => void
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
        className="bg-card rounded-lg shadow-xl border border-border w-full max-w-sm p-6"
        onClick={e => e.stopPropagation()}
      >
        <h2 className="text-lg font-semibold">{title}</h2>
        {body && <div className="mt-2 text-sm text-muted-foreground">{body}</div>}
        <div className="mt-6 flex justify-end gap-2">
          {/* ⭐ The safe choice is the plain one and sits first; the destructive
              choice is marked and sits where the primary action is looked for. */}
          <Button variant="ghost" onClick={onCancel}>
            Cancel
          </Button>
          <Button variant="destructive" onClick={onConfirm}>
            {confirmLabel}
          </Button>
        </div>
      </div>
    </div>
  )
}
