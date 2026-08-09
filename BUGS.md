# OpenOrder — bugs and feature requests

One tracker for the project. **Active** is outstanding work and nothing else;
everything decided lives in **Closed**.

## How it works

**The one rule this file exists for:** the Active list answers *"is anything
left?"* **by being looked at**, never by being read and judged. If a status
report ever has to say "there are these two, but they're not really bugs,"
that's a filing error — mark them and move them down.

**Only these four may appear in Active:**
`OPEN` · `IN PROGRESS` · `FIXED (unverified)` · `REOPENED`

Every other status moves the whole entry to Closed immediately: `VERIFIED`,
`WONTFIX`, `BY DESIGN`, `NOT A BUG`, `CANNOT REPRODUCE`.

**The cardinal rule — FIXED ≠ VERIFIED.** Code landing is `FIXED (unverified)`.
It is not done until something proves it: preferably a test that fails before
the fix and passes after. Anything that only exists in the packaged app — the
native window, the icon, install behaviour — **cannot be verified from the dev
server.** It needs a build, Jonathan confirming in the installed app, and the
build number written into the entry. *"It should work now" is not verification,
and neither is reading the diff.*

**Severity is the ship gate:** `CRITICAL` (blocks ship) · `HIGH` (major feature
broken, no workaround) · `MEDIUM` (workaround exists) · `LOW` (cosmetic).
⛔ **No CRITICAL or HIGH may be OPEN / IN PROGRESS / REOPENED when we ship.**

**Categories** (one per entry, filed under where the *fix* lands):
`UI` · `OUTPUT` · `DATA` · `BUILD`

**Feature requests** use `FR-NNN` in the same file. `OPEN` / `IN PROGRESS` /
`BUILT (unverified)` are live work; `DONE` / `UNCERTAIN` / `DECLINED` sit below
the line. A feature is not DONE until he has actually used it.

**IDs are monotonic and never reused.** Grep before claiming one — the tracker
*and* the source, because ids get claimed in code comments too:
```
grep -o "OO-[0-9][0-9][0-9]" BUGS.md | sort -u | tail
grep -rn "OO-0NN" backend/ frontend/src/ scripts/
```

**Entry shape:**
```
## <ID> — <title>
- **Category:** …
- **Severity:** …  **Status:** …  **Reported:** YYYY-MM-DD (by whom)
- **Seen in:** <version/build · dev vs packaged · OS>
- **Report:** repro · expected · actual
- **History:** <what's been tried, why it's still here>
- **Solution:** <commits / what changed> — or "—" while open
- **Test-verified:** <how proven · build # · by whom> — or "NO" while open
```

---

# Active

*(nothing open)*

---

# Feature requests

*(none open)*

---

# Closed

## OO-002 — The Windows icon renders as a circle instead of an oval
- **Category:** BUILD
- **Severity:** LOW  **Status:** VERIFIED  **Reported:** 2026-08-09 (Bob, confirmed by Jonathan)
- **Seen in:** v1.4.0 build 7 · packaged app · Windows (phoenix)
- **Report:** Suspected from source, then confirmed by Jonathan on phoenix:
  *"open order's icon on windows is what we had it set as before. It is round
  instead of oval though."* Expected: the OpenOrder mark, an oval. Actual: the
  420x310 oval squeezed into a square frame, so it renders as a circle.
- **History:** Shares OO-001's first cause (product-logo art used directly as an
  app icon) but **not its fix** — Jonathan's ruling: *"Transparent background is
  fine on windows. It's the norm."* So the macOS answer (an opaque charcoal
  squircle) is explicitly wrong here; Windows keeps the transparent surround and
  only needs the aspect ratio respected.
  A second defect found while fixing it: the old `.ico` carried **one 256px
  frame**, leaving Windows to downscale for the taskbar and Explorer itself.
- **Solution:** new `resources/images/openorder-appicon-win.svg` — the mark at
  its natural aspect on a square transparent canvas, padded above and below
  instead of stretched. `openorder.ico` rebuilt from it with a full size ladder
  (16/24/32/48/64/128/256). Alpha rebuilt from the ellipse geometry rather than
  trusted from the rasterizer, since qlmanage flattens SVG alpha onto white
  (the OO-001 lesson).
- **Second cause, found at build 8 — the fix never reached the exe.** Jonathan:
  *"fixed on the task bar, not in explorer"*, and still round after
  `ie4uinit.exe -show`. Two failed explanations (the artwork; then Explorer's
  cache) meant stop reasoning and measure. Probing the built exe's own icon
  resource showed **one 256px frame, opaque extent 246x248, ratio 0.99 — a
  circle** — while `resources\images\openorder.ico` hashed **identical** on zora
  and the Mac. The source was right all along; the exe was stale.
  **`build_win.ps1` cleared `$Dist` but not `$Work`.** PyInstaller caches
  intermediate products, including the assembled exe with its icon resource
  already embedded, and reuses them when only a *resource* changed. `build.sh`
  had always cleared its workpath; the Windows twin never did. Builds 5-8 all
  shipped the old icon.
  ⚠ *The lesson, and it is the day's recurring one: every check examined the
  `.ico` SOURCE, which was correct throughout. The question was what the EXE
  contained, and it went unasked for two rounds. The taskbar/Explorer split
  wasn't a cache signature at all — it was the two-states pattern again
  (pywebview sets the window icon at runtime from the bundled `.ico`; Explorer
  reads the resource compiled into the exe).*
- **Also fixed:** `e90da7d` → `scripts/build_win.ps1` now clears `$Work\OpenOrder`
  alongside `$Dist\OpenOrder` (v1.4.2).
- **Third round — the defect is fixed; what's left is phoenix's icon cache.**
  Jonathan at build 9: *"still round. I refreshed."* Rather than theorise a
  third time, three measurements, each answering a different question:
  1. **The exe on phoenix is the one I proved correct** — SHA256 identical to
     the copy measured on the Mac (7 frames, all oval).
  2. ⭐ **The Windows shell itself resolves an OVAL for that exact path.**
     `[System.Drawing.Icon]::ExtractAssociatedIcon` on phoenix returned 32x32,
     opaque **28x22, ratio 1.27** — matching the exe's own 32px frame exactly.
     **So the file is right AND the icon API is right; only Explorer's rendered
     view disagrees**, which is the definition of a stale shell cache.
  3. `ie4uinit.exe -show` and `-ClearIconCache` both left
     `iconcache_32.db` (2 MB) and `iconcache_48.db` (4 MB) populated.
  **Done:** the `iconcache*.db` files were deleted (they cleared successfully),
  and a control copy was placed at a **fresh path** —
  `c:\claude\icon-check\OpenOrder-v1.4.2-build9.exe` — which has no cache entry,
  so Explorer must read its icon from the file.
  ⛔ **Not done, deliberately:** restarting Explorer, which is the remaining
  standard remedy. Process-stops are destructive-tier (`rules.md` §12) and he
  didn't ask for one; it's one command whenever he wants it.
- **Test-verified:** **YES** — Jonathan on phoenix, build 9: *"Verified."* The
  control copy at the uncached path settled it, and he named the step that
  finishes the original path: *"I should have restarted explorer before."*
  (2026-08-09, v1.4.2 build 9)

## FR-001 — Build and test OpenOrder on Windows
- **Category:** BUILD
- **Severity:** —  **Status:** DONE  **Reported:** 2026-08-09 (Jonathan)
- **Seen in:** v1.4.0 build 7
- **Request:** *"I want to start building windows. We build on Zora and copy to
  phoenix for testing. Let's start doing the same here."* Follow FirstLight's
  pipeline shape. Later ruling on scope: *"Both every time"* — mac and Windows
  build together on every change, one event, one build number.
- **History:** Stood up end to end on 2026-08-09: bare repo `d:\claude\OpenOrder.git`
  on zora with a working clone, its own Python 3.11 venv and node modules;
  `scripts/build_win.ps1` (pure ASCII) reads the event's build number and never
  bumps it; `scripts/build_win_remote.sh` drives zora with a narrow clean, a HEAD
  match check, the counter scp'd strictly after the pull, full logs to disk and a
  fail-closed receipt check; `scripts/fanout_win.sh` sends the artifact to phoenix
  and back to the Mac with SHA256 verified on every copy; `scripts/build_all.sh`
  runs the whole joint event as one command.
  Two deliberate divergences from FirstLight: the artifact is a PyInstaller onedir
  *folder*, so it travels as one zip and one hash proves the whole thing arrived;
  and the hymnal is **not** bundled — it's copyrighted and gitignored, so the app
  reaches it through the `hymnal_dir` setting instead.
- **Solution:** `22b459d` (the pipeline: `build_win.ps1`, `build_win_remote.sh`,
  `fanout_win.sh`), `26bd6e7` (`build_all.sh` — the joint event as one command),
  `82b82a9` (printf fix for the summary path); machines table and ritual in
  `rules_project.md`.
- **Test-verified:** **YES** — joint build 7 produced mac `1.4.0 (7)` and windows
  `v1.4.0 build 7` from one source state, SHA256 identical across zora, phoenix and
  the Mac; phoenix's exe read back at 08:44:00. Jonathan ran it there and confirmed:
  *"windows confirmed. All good."* (2026-08-09, build 7)

## OO-001 — The app icon was wrong: an oversized circle in the Dock, and it changed on launch
- **Category:** BUILD
- **Severity:** LOW  **Status:** VERIFIED  **Reported:** 2026-08-09 (Jonathan)
- **Seen in:** v1.1.0 build 1 through v1.2.0 build 2 · packaged app only · macOS
- **Report:** *"the icon for the build while running is right. The icon as it sits
  in the dock is wrong… in the dock it's larger than the rest and round not like
  the curved corner icons of everything else."* Expected: a normal macOS app icon —
  a rounded square, sized like its neighbours. Actual: a bare circle, unpadded,
  visually larger than every other Dock icon; and after the first fix, correct in
  the Dock but reverting to the old circle the moment the app launched.
- **History:** Two distinct causes behind one symptom, which is why the first fix
  looked complete and wasn't.
  1. **The artwork had no icon canvas.** `openorder.icns` was built from the
     *product logo* — a 420x310 canvas holding a bare oval on transparency. Squeezed
     into a square icns that gives macOS nothing to draw, so the Dock rendered the
     ellipse itself. Compounded by the rasterizer: `qlmanage` flattens SVG alpha onto
     **white**, which is where the white background came from.
  2. **The running app overrode the bundle icon.** With the artwork fixed, Jonathan
     reported *"it works, but when running it's still this."* `webview.start(icon=…)`
     was being handed `openorder.ico` — the *Windows* icon, same stale art — so macOS
     used the correct bundle icon until launch and the wrong one after.
  The platform behaviour behind both (a runtime icon API is not the bundle icon;
  qlmanage flattens alpha) belongs in `..\bob\build_rules.md`, where it was written
  up so the next project doesn't rediscover it.
- **Solution:** `dac6d0e` (v1.2.0) — new `resources/images/openorder-appicon-mac.svg`
  on Apple's grid (1024 artboard, 824 squircle at radius 185, dark charcoal, mark
  centred), with alpha rebuilt from that geometry rather than trusted from the
  renderer. `22e50fa` (v1.2.1) — macOS passes **no** runtime icon; the bundle's
  `.icns` stands for both states. Windows/Linux keep the `.ico`.
- **Test-verified:** **YES** — icon pixels checked in the *installed* bundle (corner
  alpha 0, charcoal canvas) at build 2; Jonathan confirmed the Dock icon at build 2
  (*"icon looks great in the dock"*) and the running-app icon at build 3
  (*"Yep. Perfect."*). (2026-08-09, builds 2 and 3)
