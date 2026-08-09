# Changelog

What has changed in each release of OpenOrder, newest first.

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
