# Changelog

What has changed in each release of OpenOrder, newest first.

## 1.15.0 — 2026-09-26 · The app says when the hymnal is missing

*The hymn collections are copyrighted, so a fresh install has none until one is
pointed at. Searching simply returned nothing, which reads exactly like "no such
hymn" — the app looked broken when it was only unconfigured.*

- **Added:** a notice on first load when no hymnal is configured, naming the two
  collections OpenOrder reads and where to set the folder. It appears once.
- **Added:** the same explanation inside the picker at the moment a search comes
  back empty, so it is there when the question is actually being asked.
- **Fixed:** the notice now names everything affected. The creed, the doxology and
  the liturgical prayer come from the same collections as the hymns, so a missing
  hymnal costs all four — earlier wording claimed only hymn search was unavailable.
- **Changed:** two checks now have to pass before a release can be built — the type
  check, and a spelling check over every tracked file.

## 1.11.0 — 2026-09-25 · A reading always makes it into the slides

*Choosing a translation that needs the internet meant that, without it, the
slides came out with the scripture missing — a complete-looking deck and no
reading in it, which is the kind of thing you find out on a Sunday.*

- **Added:** when the translation you picked cannot be fetched, the Berean
  Standard Bible stands in rather than the reading being left out, and the app
  says so — both when you enter the reference and again after the slides are
  made. It names the reason plainly: no connection, a key that was not
  accepted, or a translation that does not carry that passage.
- **Fixed:** the verses now appear whenever a service is open, not only while
  you are typing the reference. Reloading the page, opening a past service or
  switching congregation used to leave the reference and translation filled in
  with nothing underneath.
- **Changed:** Settings tells you which problem you have when a key stops
  working — whether it is being refused, or whether the service simply cannot
  be reached right now. One is fixed with a new key, the other by waiting.

## 1.10.0 — 2026-09-25 · Scripture that works with the cable unplugged

*Every translation OpenOrder offered was fetched over the internet, including
the one it used by default. A service put together with no connection simply
had no scripture in it — no warning, just a gap where the reading should be.*

- **Added:** the Berean Standard Bible now travels with the app. All 31,102
  verses are on disk, so the default translation works with no connection at
  all. It is in the public domain, so it costs nothing and is yours to keep.
- **Added:** your own translations. A free key from API.Bible, entered in
  Settings, adds whatever your congregation has picked — the key is yours, the
  key stays on your computer, and nothing is billed to anyone. The translation
  list shows the ones your key actually provides rather than every edition in
  their catalog.
- **Added:** a note when a translation hands back more than you asked for. Some,
  The Message especially, are written in paragraphs rather than verses, so the
  smallest portion containing your reading is what exists. It now says which
  verses you will get before you make the slides.
- **Fixed:** a reading like "John 14:15-17, 25-27" no longer quietly includes
  the verses in between.
- **Changed:** the translation is named once on a reading, on the badge, at the
  start and the end.

## 1.9.0 — 2026-09-25 · The church year, without typing it

*The service title was a blank box, so the name of the Sunday was typed from
memory every week — and a form that only saved when you remembered to press Save
could lose an afternoon to a stray click.*

- **Added:** the service title offers the whole church year. The United Methodist
  Book of Worship's calendar and its eleven Special Sundays are in the list, and a
  button beside the field fills in what that Sunday actually is. A holy day comes
  before a Special Sunday, which comes before the counted Sunday; where only the
  counted names apply it uses the wording your congregation chose last time, so a
  church that says "Ordinary Time" keeps getting it. You can still type anything
  you like — Homecoming, a sermon series, whatever is on the bulletin.
- **Added:** a communion Sunday uses your congregation's communion bulletin
  template when it has one, and its ordinary template when it doesn't.
- **Changed:** starting a new service asks first. It used to empty the form on a
  single click, next to the button you press every week.
- **Changed:** the tabs stay at the top of the page and carry the save status, so
  you can see the work is saved without scrolling to the bottom to check.
- **Fixed:** a save that fails now says so and refuses to let the form be cleared
  or replaced over it. A locked or read-only file explains itself in words.

## 1.8.0 — 2026-09-25 · Two congregations, two services, one Sunday

*Until now a Sunday held one order of worship, which was right while both
congregations sang the same hymns and differed only in letterhead. The first
Sunday they need different sermons, that stops being true.*

- **Added:** a service belongs to a congregation. Each one keeps its own order of
  worship, its own picture and its own communion setting for the same date, and
  the list of past services shows that congregation's.
- **Added:** the form saves itself. The Save button is gone, replaced by a quiet
  line telling you whether the work is saved.
- **Added:** Revert. Earlier versions are kept automatically whenever work could
  be lost — opening a service, switching congregation, producing a document — and
  any of them can be restored. Reverting is itself undoable.
- **Changed:** switching congregation brings that congregation's version of the
  same Sunday with it, and says so. If it doesn't have one yet, what you have
  carries over — which is the ordinary week, where both get the same service.

## 1.7.0 — 2026-09-24 · A second congregation gets everything the first one had

*Support for more than one congregation arrived a version ago, but a second one
still borrowed pieces of the first: its calendar, and the color of it. With this
release a congregation added today produces the same bulletin and the same slides
a single congregation always did — nothing shared, nothing inherited.*

- **Added:** each congregation keeps its own calendar. Recurring events, the
  per-week skips, one-off events and the bulletin note are all its own, and the
  calendar tab edits whichever congregation is selected. A congregation with no
  calendar yet simply has an empty one rather than showing another's.
- **Fixed:** the calendar block in the bulletin was drawn in a fixed olive green,
  which is one congregation's color. It now uses that congregation's own accent —
  the headers, the times, the divider, the note label and the side rule.

## 1.6.0 — 2026-09-24 · Artwork that behaves on a white slide

*The first decks produced for a second congregation showed that backgrounds were
being handled in ways that only ever worked for the artwork already shipped.*

- **Fixed:** artwork with a transparent background came through black, so the
  background control had to be pushed to its limit just to reach white. It is
  now laid onto white, which is what a slide is, and the control starts there.
- **Fixed:** the square watermark behind hymn lyrics was being cropped and
  stretched whenever its background was adjusted. It keeps its shape.
- **Fixed:** background artwork was covering the hymn lyrics, the creed and the
  Lord's Prayer. It had only ever looked right because the artwork was faint
  enough to read through; backgrounds now sit behind the words.
- **Changed:** producing a bulletin or slides without choosing a congregation is
  refused rather than quietly falling back to neutral colors and the packaged
  artwork. The chooser no longer offers a blank option that looked like a choice.

## 1.5.0 — 2026-09-24 · More than one congregation

*OpenOrder was built around a single church. It now serves several from one set
of service details, with each congregation keeping its own letterhead, colors
and artwork.*

- **Added:** congregations. Each one is a folder of its own holding its bulletin
  template, its slide colors, its artwork and who normally preaches there.
  Pick one at the top of the service, and the bulletin and slides follow.
- **Added:** the same service can be produced for each congregation in turn. The
  generated files carry the congregation in their names, so producing the second
  no longer overwrites the first.
- **Added:** a communion Sunday switch, which adds a Holy Communion slide with
  the sermon rather than after it.
- **Changed:** slide titles such as *Pastoral Prayer* are drawn onto the artwork
  rather than being part of it. A picture is now just a picture, and the title
  appears in that congregation's own font and color on a frosted plate whose
  placement, blur and weight are all adjustable.
- **Changed:** the per-service picture is now called the hero image, so that
  *theme* means only a congregation's colors and type. Services saved before
  this keep working untouched.
- **Removed:** the worship leader field and the duplicate speaker field. Both
  were collected and never appeared anywhere; that information belongs in the
  bulletin template, which already carried it.
- **Fixed:** every request from the application to its own backend failed during
  development on machines that resolve local addresses a particular way.

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
  oversized circle among its neighbors.
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
