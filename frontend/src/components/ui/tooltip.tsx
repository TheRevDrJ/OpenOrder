import { useState, useRef, useCallback, useEffect, type ReactNode } from 'react'
import { createPortal } from 'react-dom'

/**
 * A hover tip that is ours, not the browser's.
 *
 * ⛔ NOT the native `title` attribute. @decision:gold 2026-09-25
 * A `title` takes about a second to appear, cannot be styled or positioned, and
 * is browser chrome — which is absent or inconsistent inside the packaged app,
 * where this software actually gets used on a Sunday morning.
 *
 * ⛔⛔ THE BUBBLE IS PORTALLED, FOR THE SAME REASON THE TITLE PICKER'S LIST IS:
 * `Card` sets `overflow-hidden`, so an absolutely-positioned bubble is CLIPPED
 * at the card's edge. The first version of this was positioned inside the card
 * and the first tip anyone hovered came out cut in half.
 *
 * ⚠ AND IT WRAPS. `whitespace-nowrap` is right for two words and wrong for a
 * sentence — it made one long line that ran off the card. Tips here are
 * sentences, so they get a max width and wrap.
 *
 * ⚠ SHOWS ON FOCUS AS WELL AS HOVER. A tip that only answers a mouse is a tip
 * that does not exist for anyone driving the form from the keyboard.
 */
export function Tooltip({
  content,
  children,
  side = 'top',
  align = 'left',
  className = '',
}: {
  content: ReactNode
  children: ReactNode
  side?: 'top' | 'bottom'
  /** Which edge of the trigger the bubble lines up with. */
  align?: 'left' | 'right'
  /**
   * Extra classes for the wrapper. ⚠ The wrapper is an element in the layout,
   * so a full-width control needs `w-full` here or it collapses to its
   * content — the tip must not change what it is attached to.
   */
  className?: string
}) {
  const [pos, setPos] = useState<{ x: number; y: number; flip: boolean } | null>(null)
  const ref = useRef<HTMLDivElement>(null)

  const measure = useCallback(() => {
    const el = ref.current
    if (!el) return
    const r = el.getBoundingClientRect()
    // Flip to the other side when the chosen one has no room. 64px is enough
    // for two wrapped lines plus the gap.
    const wantTop = side === 'top'
    const roomTop = r.top > 64
    const roomBottom = window.innerHeight - r.bottom > 64
    const flip = wantTop ? !roomTop && roomBottom : !roomBottom && roomTop
    const onTop = wantTop !== flip
    setPos({
      x: align === 'left' ? r.left : r.right,
      y: onTop ? r.top - 6 : r.bottom + 6,
      flip: !onTop,
    })
  }, [side, align])

  useEffect(() => {
    if (!pos) return
    const onMove = () => measure()
    window.addEventListener('scroll', onMove, true)
    window.addEventListener('resize', onMove)
    return () => {
      window.removeEventListener('scroll', onMove, true)
      window.removeEventListener('resize', onMove)
    }
  }, [pos, measure])

  if (!content) return <>{children}</>

  const show = () => measure()
  const hide = () => setPos(null)

  return (
    <div
      ref={ref}
      className={`relative inline-flex ${className}`}
      onMouseEnter={show}
      onMouseLeave={hide}
      onFocusCapture={show}
      onBlurCapture={hide}
    >
      {children}
      {pos && createPortal(
        <div
          role="tooltip"
          style={{
            position: 'fixed',
            left: align === 'left' ? pos.x : undefined,
            right: align === 'right' ? window.innerWidth - pos.x : undefined,
            ...(pos.flip ? { top: pos.y } : { bottom: window.innerHeight - pos.y }),
          }}
          className="z-[60] max-w-xs px-3 py-1.5 text-xs leading-snug bg-popover text-popover-foreground border border-border rounded-md shadow-md pointer-events-none"
        >
          {content}
        </div>,
        document.body,
      )}
    </div>
  )
}
