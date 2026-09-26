#!/usr/bin/env python3
"""Give every legacy date-only service a church, by renaming it.

WHY: FEATURE-008 keys a service on `{date} - {church} - Raw.json`. Everything saved
before 2026-09-25 is `{date} - Raw.json` and belongs to whichever congregation was
the only one at the time — which is why the target is an argument rather than a
guess. @decision:gold 2026-09-25

⛔ RENAMES, NEVER COPIES, and refuses if the destination exists — two files for one
date would be ambiguous about which is real.
⚠ The hero image beside each one moves with it; a service whose art stayed behind
would render a blank slide rather than an error.

USAGE:  python3 scripts/migrate_services_to_church.py <church-id> [--apply]
        Dry run by default. CALLED BY: nobody — run by hand, once.
"""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from backend.app import paths  # noqa: E402

LEGACY = re.compile(r"^(\d{4}-\d{2}-\d{2}) - Raw\.json$")


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    church = sys.argv[1]
    apply = "--apply" in sys.argv
    d = paths.DATA_DIR
    moves: list[tuple[Path, Path]] = []

    for f in sorted(d.glob("* - Raw.json")):
        m = LEGACY.match(f.name)
        if not m:
            continue                                   # already carries a church
        date = m.group(1)
        moves.append((f, d / f"{date} - {church} - Raw.json"))
        # the hero image travels with its service
        for hero in sorted(d.glob(f"{date} - Theme.*")):
            moves.append((hero, d / f"{date} - {church} - Theme{hero.suffix}"))

    if not moves:
        print("Nothing to migrate — no date-only services found.")
        return 0

    clash = [dst for _, dst in moves if dst.exists()]
    if clash:
        print("ABORT: these destinations already exist:")
        for c in clash:
            print("  ", c.name)
        return 1

    for src, dst in moves:
        print(("  MOVE " if apply else "  would move "), src.name, "->", dst.name)
        if apply:
            src.rename(dst)

    # ⛔⛔ THE POINTER MOVES WITH THE FILE OR THE SLIDE GOES BLANK. A service stores
    # its hero by NAME (`heroImageFilename`), resolved as DATA_DIR / that name — so a
    # renamed image with an un-rewritten pointer resolves to nothing, and a missing
    # hero produces an EMPTY SLIDE rather than an error. Renaming alone would have
    # silently blanked the hero on all 17 services.
    renamed = {src.name: dst.name for src, dst in moves if src.suffix != ".json"}
    if apply and renamed:
        fixed = 0
        for _, dst in moves:
            if dst.suffix != ".json":
                continue
            raw = json.loads(dst.read_text(encoding="utf-8"))
            for key in ("heroImageFilename", "themeImageFilename"):
                if raw.get(key) in renamed:
                    raw[key] = renamed[raw[key]]
                    fixed += 1
            dst.write_text(json.dumps(raw, indent=2, ensure_ascii=False) + "\n",
                           encoding="utf-8")
        print(f"  rewrote {fixed} hero pointer(s) to match")

    print(f"\n{len(moves)} file(s) {'moved' if apply else 'would move'}.")
    if not apply:
        print("Dry run. Re-run with --apply to do it.")
    else:
        # read the receipt: every migrated service must still parse
        from backend.app.models import OrderOfWorship
        bad = []
        for _, dst in moves:
            if dst.suffix != ".json":
                continue
            try:
                raw = json.loads(dst.read_text(encoding="utf-8"))
                OrderOfWorship(**raw)
                hero = raw.get("heroImageFilename") or raw.get("themeImageFilename")
                if hero and not (d / hero).is_file():
                    bad.append((dst.name, f"hero missing: {hero}"))
            except Exception as e:                      # pragma: no cover
                bad.append((dst.name, str(e)[:60]))
        print("  every service parses and its hero resolves:",
              "yes" if not bad else bad)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
