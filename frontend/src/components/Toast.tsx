import { useEffect, useRef, useState } from 'react'

export interface ToastData {
  /** Bumped on every new toast so re-firing restarts the timer and animation. */
  id: number
  message: string
  detail?: string
}

/**
 * A brief confirmation that auto-dismisses.
 *
 * The desktop app has no browser download bar, so generating a file used to
 * succeed completely silently — nothing on screen changed and there was no way
 * to tell it had worked, or where the file went. This is that missing receipt.
 */
export function Toast({
  toast,
  onDone,
  duration = 3000,
}: {
  toast: ToastData | null
  onDone: () => void
  duration?: number
}) {
  const [visible, setVisible] = useState(false)

  // Held in a ref so an inline arrow from the caller — which is a new function
  // on every render — can't restart the timers and leave the toast up forever.
  const onDoneRef = useRef(onDone)
  onDoneRef.current = onDone

  useEffect(() => {
    if (!toast) return
    // Next frame, so the element mounts hidden and then transitions in.
    const raf = requestAnimationFrame(() => setVisible(true))
    const hide = setTimeout(() => setVisible(false), duration)
    // Clear only after the fade-out has finished, or it would vanish abruptly.
    const done = setTimeout(() => onDoneRef.current(), duration + 250)
    return () => {
      cancelAnimationFrame(raf)
      clearTimeout(hide)
      clearTimeout(done)
      setVisible(false)
    }
  }, [toast, duration])

  if (!toast) return null

  return (
    <div
      role="status"
      aria-live="polite"
      onClick={() => setVisible(false)}
      className={`fixed bottom-6 right-6 z-[100] max-w-sm cursor-pointer rounded-lg border border-border bg-card px-4 py-3 shadow-lg transition-all duration-200 ${
        visible ? 'translate-y-0 opacity-100' : 'translate-y-2 opacity-0'
      }`}
    >
      <div className="flex items-start gap-3">
        <svg
          className="mt-0.5 shrink-0 text-green-600 dark:text-green-400"
          xmlns="http://www.w3.org/2000/svg"
          width="18"
          height="18"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2.5"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <path d="M20 6 9 17l-5-5" />
        </svg>
        <div className="min-w-0">
          <p className="text-sm font-medium text-foreground">{toast.message}</p>
          {toast.detail && (
            <p className="mt-0.5 break-all font-mono text-xs text-muted-foreground">{toast.detail}</p>
          )}
        </div>
      </div>
    </div>
  )
}
