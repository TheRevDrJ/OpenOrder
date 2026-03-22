"""
Extract lyrics and slide layout info from .pptx hymn files into JSON.

For each .pptx file, produces a JSON with:
- number: hymn number (parsed from filename)
- title: hymn title (parsed from filename)
- source: which collection it came from
- slides: list of slides, each with lines of text and line count

Skips Dropbox "conflicted copy" files.
"""

import json
import os
import re
import sys
from pathlib import Path
from pptx import Presentation


# Folders to process, with their source labels
FOLDERS = {
    "Hymnal/Hymns by Number": "umh",
    "Hymnal/Other General Services and Acts of Worship": "umh-services",
    "Hymnal/General Services": "umh-general",
    "Hymnal/Psalter": "umh-psalter",
    "The Faith We Sing/Number Index (Cross & Flame background)": "tfws",
}


def parse_hymn_filename(filename: str, source: str) -> dict:
    """Extract hymn number and title from filename."""
    name = Path(filename).stem

    if source == "umh":
        # Format: "057 - O For a Thousand Tongues to Sing"
        match = re.match(r"^(\d+)\s*-\s*(.+)$", name)
        if match:
            return {"number": match.group(1), "title": match.group(2).strip()}
    elif source in ("umh-services", "umh-general"):
        # Format: "881-The Apsotles' Creed, Traditional Version"
        match = re.match(r"^(\d+)-(.+)$", name)
        if match:
            return {"number": match.group(1), "title": match.group(2).strip()}
    elif source == "umh-psalter":
        # Format: "Psalm 001" or "Psalm 009 (11-20)"
        match = re.match(r"^Psalm\s+(\d+)(?:\s*\((.+)\))?$", name)
        if match:
            num = match.group(1)
            suffix = f" ({match.group(2)})" if match.group(2) else ""
            return {"number": num, "title": f"Psalm {num}{suffix}"}
    elif source == "tfws":
        # Format: "2001 - Grace Alone" or similar
        match = re.match(r"^(\d+)\s*-\s*(.+)$", name)
        if match:
            return {"number": match.group(1), "title": match.group(2).strip()}
        # Some might just be number-title without space
        match = re.match(r"^(\d+)(.+)$", name)
        if match:
            return {"number": match.group(1), "title": match.group(2).strip().lstrip("- ")}

    # Fallback
    return {"number": "0", "title": name}


def extract_slide_text(slide) -> list[str]:
    """Extract all text lines from a slide's text frames."""
    lines = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text = paragraph.text.strip()
                if text:
                    lines.append(text)
    return lines


def process_pptx(filepath: str, source: str) -> dict | None:
    """Extract lyrics data from a single .pptx file."""
    filename = os.path.basename(filepath)

    # Skip conflicted copies
    if "conflicted copy" in filename:
        return None

    info = parse_hymn_filename(filename, source)

    try:
        prs = Presentation(filepath)
    except Exception as e:
        print(f"  ERROR reading {filename}: {e}", file=sys.stderr)
        return None

    slides = []
    for slide in prs.slides:
        lines = extract_slide_text(slide)
        if lines:
            slides.append({
                "lines": lines,
                "line_count": len(lines),
            })

    if not slides:
        return None

    return {
        "number": info["number"],
        "title": info["title"],
        "source": source,
        "filename": filename,
        "slide_count": len(slides),
        "slides": slides,
    }


def main():
    script_dir = Path(__file__).parent
    root_dir = script_dir.parent / "Hymnals"
    output_dir = script_dir.parent / "hymnal-json"
    output_dir.mkdir(exist_ok=True)

    # Create subdirs for each source
    for source in set(FOLDERS.values()):
        (output_dir / source).mkdir(exist_ok=True)

    total = 0
    converted = 0
    skipped = 0
    failed = 0

    for folder_rel, source in FOLDERS.items():
        folder = root_dir / folder_rel
        if not folder.exists():
            print(f"WARNING: Folder not found: {folder}")
            continue

        pptx_files = sorted(folder.glob("*.pptx"))
        print(f"\n{source}: {len(pptx_files)} .pptx files in {folder_rel}")

        for filepath in pptx_files:
            total += 1

            if "conflicted copy" in filepath.name:
                skipped += 1
                continue

            result = process_pptx(str(filepath), source)
            if result is None:
                failed += 1
                continue

            # Save JSON
            json_name = filepath.stem + ".json"
            json_path = output_dir / source / json_name
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(result, f, indent=2, ensure_ascii=False)

            converted += 1
            if converted % 50 == 0:
                print(f"  ... {converted} extracted so far")

    # Also build a master index
    index = []
    for source_dir in output_dir.iterdir():
        if source_dir.is_dir():
            for json_file in sorted(source_dir.glob("*.json")):
                with open(json_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                index.append({
                    "number": data["number"],
                    "title": data["title"],
                    "source": data["source"],
                    "slide_count": data["slide_count"],
                    "file": f"{source_dir.name}/{json_file.name}",
                })

    index.sort(key=lambda x: (x["source"], x["number"].zfill(6)))
    index_path = output_dir / "index.json"
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2, ensure_ascii=False)

    print(f"\n=== Done ===")
    print(f"  Converted: {converted}")
    print(f"  Skipped: {skipped}")
    print(f"  Failed: {failed}")
    print(f"  Total: {total}")
    print(f"  Index: {index_path} ({len(index)} entries)")


if __name__ == "__main__":
    main()
