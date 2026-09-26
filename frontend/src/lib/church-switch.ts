/**
 * The words the church-switch question asks — and nothing else.
 *
 * ⭐⭐ WHY THIS IS ITS OWN MODULE. The workshop bench reviews this copy, and a bench that
 * renders its own COPY of the words reviews the copy. One function, imported by the app
 * and by the bench, so what is read on the bench is what the app will say.
 * ⛔ The sample states live with the BENCH, not here: example data in shipped source is
 * how a real organization's name ends up in a public repo.
 *
 * ⭐ PLAIN DATA, NO JSX, on purpose: the copy is then reviewable as text, and where the
 * emphasis goes stays a question for the layer that draws it.
 *
 * CALLED BY: App.tsx (selectChurch), and the workshop's switch bench.
 */

export interface SwitchOption {
  id: 'load' | 'carry' | 'clear'
  label: string
  detail?: string
}

export interface SwitchChoice {
  title: string
  /** One paragraph each, in order. */
  lines: string[]
  options: SwitchOption[]
}

export interface SwitchState {
  /** The church being arrived at, by display name. */
  name: string
  /** The church being left, by display name — null on a first selection. */
  leavingName: string | null
  date: string
  /** Does the arriving church already have a saved service for this date? */
  theirs: boolean
  /** Has anybody put anything in the form? */
  mine: boolean
}

/** ⛔ Returns null when there is nothing to ask, and the caller must then ask nothing.
 *  An empty form with nothing saved has one possible outcome, and a prompt offering one
 *  outcome is noise. */
export function switchChoice(s: SwitchState): SwitchChoice | null {
  if (!s.mine && !s.theirs) return null

  const lines = [
    s.theirs
      ? `${s.name} already has a service saved for ${s.date}.`
      : `${s.name} has nothing saved for ${s.date} yet.`,
  ]
  // ⭐ Said ONCE, here, rather than on one of the options. The first draft explained the
  // outgoing church on "load" and the arriving church on the other two, so the frame
  // moved underneath the reader at the moment of deciding. It is true of all three.
  // ⚠ And only when it is true: an empty form had nothing to save.
  if (s.mine && s.leavingName) {
    lines.push(`What is on screen is already saved under ${s.leavingName}.`)
  }

  // ⚠ NO OPTION REPEATS THE CHURCH'S NAME WHILE ITS NEIGHBORS DO. The first draft's
  // "Start blank" was the only unnamed one, which made the most destructive choice read
  // as the most casual. The name is in the title and the first line; here they are a
  // parallel verb set — keep / replace / empty — all about the arriving church.
  const options: SwitchOption[] = []
  if (s.theirs) {
    options.push({
      id: 'load',
      label: `Keep ${s.name}'s saved service`,
      detail: 'Opens the service they already have.',
    })
  }
  if (s.mine) {
    options.push({
      id: 'carry',
      label: s.theirs ? 'Replace it with what is on screen' : `Save what is on screen under ${s.name}`,
      detail: s.theirs ? 'Theirs is kept in Revert.' : undefined,
    })
  }
  options.push({
    id: 'clear',
    label: s.theirs ? 'Empty it' : 'Start with an empty form',
    detail: s.theirs ? 'Theirs is kept in Revert.' : undefined,
  })

  return { title: `Switching to ${s.name}`, lines, options }
}
