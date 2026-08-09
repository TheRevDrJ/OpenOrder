# Changelog

What has changed in each release of OpenOrder, newest first.

## 1.4.2 — 2026-08-09 · The Windows icon fix actually reaches the application

*1.4.1 corrected the icon; this is what makes the corrected icon end up in the
program you run.*

- **Fixed:** the Windows build reused cached pieces of the previous build, so a
  changed icon never made it into the finished application. Windows builds now
  rebuild that step from scratch.

## 1.4.1 — 2026-08-09 · The Windows icon is an oval again, at every size

*Matching what the macOS icon got right, in the way Windows expects it.*

- **Fixed:** the Windows application icon was the OpenOrder mark squeezed into a
  square frame, so it appeared as a circle rather than an oval. It now keeps its
  proper shape, on the transparent background Windows expects.
- **Fixed:** the icon carried only one 256-pixel size, leaving Windows to shrink
  it for the taskbar and Explorer. It now ships every standard size.

## 1.4.0 — 2026-08-09 · Windows builds, produced and tested every release

*OpenOrder is now built for Windows and macOS together, from the same source, as
one release.*

- **Added:** a Windows build pipeline. Every release now produces a macOS app and
  a Windows app from the same commit, both carrying the same build number, with
  the Windows build installed onto a test machine automatically.
- **Added:** every Windows build is checksum-verified when it lands, so a
  truncated or corrupted transfer can't be mistaken for a bug in the app.
- **Note for Windows:** the app finds your hymnal through the Hymnal Folder
  setting rather than expecting it inside the application, which keeps hymnal
  lyrics out of the distributed files entirely.

## 1.3.1 — 2026-08-09 · Uploading a template no longer looks like it failed

*A display fix, but an alarming one: it read as though your template had vanished.*

- **Fixed:** after uploading a bulletin template, Settings reported "No template
  found" and disabled Download, even though the upload had succeeded and the file
  was safely stored. Both now reflect the real state.

## 1.3.0 — 2026-08-09 · Download your bulletin template, and keep it through updates

*Template handling that actually works end to end: see the format, edit it, put it
back — and it stays put.*

- **Added:** a **Download Template** button in Settings. It saves a copy of the
  template in use to your output folder, so you can open it, see how it's built,
  and edit it. Uploading was of limited use without it.
- **Fixed:** an uploaded template was stored inside the application itself, so
  **installing an update silently replaced your template with the built-in one.**
  Your template is now kept with your data and survives updates.
- **Added:** Settings shows whether the template in use is the built-in one or
  your own.
- **Changed:** the calendar leaves an extra blank line between weeks, so each
  week separates more clearly on the printed page.

## 1.2.1 — 2026-08-09 · The app keeps its own icon while running

*A cosmetic fix, but a visible one.*

- **Fixed:** launching OpenOrder replaced its Dock icon with a different, older
  one. The icon is now the same whether the app is running or not.

## 1.2.0 — 2026-08-09 · Saved files confirm themselves; a proper macOS icon

*You can tell when a bulletin has been generated, and where it went.*

- **Added:** a brief confirmation when a bulletin or presentation is saved,
  naming the folder it was written to. The desktop app has no browser download
  bar, so generating a file used to happen silently with nothing on screen.
- **Fixed:** the macOS app icon is a real app icon now — the OpenOrder mark on a
  dark rounded square — instead of the bare logo, which the Dock drew as an
  oversized circle among its neighbours.
- **Removed:** a second, redundant download after generating a file, which left a
  duplicate copy behind when running in a browser.

## 1.1.0 — 2026-08-08 · You choose where OpenOrder keeps things

*Three separate folders you control, and settings that survive an update.*

- **Added:** independent folder settings for generated documents (defaults to
  your Downloads folder), your calendar and saved services (defaults to
  `Documents/OpenOrder`), and your hymnal — each with a folder picker in
  Settings.
- **Changed:** settings are now stored in the operating system's standard
  per-user location. Reinstalling or updating the app no longer discards your
  folder choices along with it.
- **Changed:** saved services and uploaded theme images are treated as your
  working files rather than as output, so they stay with your data instead of
  being written to the downloads folder. Only the finished `.docx` and `.pptx`
  land there.
- **Added:** the macOS build now records a version and build number you can read
  in Get Info, verifies its own signature, and installs itself.

## 1.0.0 — 2026-07-11 · OpenOrder runs on macOS as well as Windows

*The first release that isn't Windows-only.*

- **Added:** `setup.sh` and `openorder.sh` for macOS and Linux — one command to
  install, one to start and stop the dev servers — alongside the existing
  Windows scripts.
- **Added:** a macOS desktop app (`OpenOrder.app`) built by `build.sh`.
- **Fixed:** the Windows server manager could stop unrelated programs when
  shutting down. It now only ever acts on OpenOrder's own two ports.
- **Changed:** the frontend and API moved to fixed ports (6800 and 6801) so
  OpenOrder no longer competes with other local development servers.
- **Changed:** line endings are pinned per file type, so the project no longer
  shows spurious whole-file changes when moved between Windows and macOS.
