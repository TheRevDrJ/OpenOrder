import { cn } from '@/lib/utils'

/**
 * A toggle switch.
 *
 * ⭐ Hand-built rather than pulling in a dependency: this is the only toggle in
 * the app, and a package would be a new supply-chain edge for twenty lines.
 *
 * ⭐ It IS a real checkbox underneath — the visible track and knob are drawn from
 * its state with `peer-checked:`. That keeps the keyboard behaviour, the focus
 * ring, form semantics and screen-reader announcement that a div-with-onClick
 * throws away, while looking nothing like a checkbox.
 */
export function Switch({
  checked,
  onCheckedChange,
  id,
  label,
  hint,
  className,
}: {
  checked: boolean
  onCheckedChange: (v: boolean) => void
  id: string
  label: string
  hint?: string
  className?: string
}) {
  return (
    <label
      htmlFor={id}
      className={cn(
        'flex items-center gap-3 cursor-pointer select-none rounded-lg',
        'border border-input bg-transparent px-3 py-2.5 transition-colors',
        'hover:bg-accent/40 has-[:focus-visible]:ring-[3px] has-[:focus-visible]:ring-ring/50',
        className,
      )}
    >
      <input
        id={id}
        type="checkbox"
        className="peer sr-only"
        checked={checked}
        onChange={e => onCheckedChange(e.target.checked)}
      />
      <span
        aria-hidden
        className={cn(
          'relative h-6 w-10 shrink-0 rounded-full transition-colors duration-200',
          'bg-input peer-checked:bg-primary',
          "after:absolute after:top-0.5 after:left-0.5 after:h-5 after:w-5 after:rounded-full",
          'after:bg-background after:shadow-sm after:transition-transform after:duration-200',
          'peer-checked:after:translate-x-4',
        )}
      />
      <span className="min-w-0">
        <span className="block text-sm font-medium leading-tight">{label}</span>
        {hint && <span className="block text-xs text-muted-foreground mt-0.5">{hint}</span>}
      </span>
    </label>
  )
}
