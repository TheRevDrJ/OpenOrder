# OpenOrder

A web-based worship bulletin and presentation slide generator for churches.

OpenOrder streamlines the weekly task of preparing Sunday worship materials. Enter your service details — hymns, scripture, sermon title — and generate a print-ready Word bulletin and PowerPoint presentation in seconds.

## Features

- **Hymn Search** — Search by number or title across the United Methodist Hymnal (UMH) and The Faith We Sing (TFWS)
- **Bulletin Generation** — Produces a formatted `.docx` bulletin ready to print and fold
- **Slide Generation** — Produces a `.pptx` presentation with hymn lyrics, scripture, backgrounds, and static slides
- **Dark Mode** — Light and dark themes for late-night sermon prep
- **Scripture Integration** — Fetches scripture text from the free AO Lab Bible API with local caching. Supports BSB, KJV, ASV, WEB, and more
- **Themeable** — Bulletin layout and styling are driven by themes. Ships with two built-in themes, and the community can create and share more

## Tech Stack

- **Frontend**: React + TypeScript, Tailwind CSS, shadcn/ui
- **Backend**: Python (FastAPI)
- **Document Generation**: python-docx, python-pptx
- **Port**: 8316 (John 3:16)

## Prerequisites

- Python 3.10+
- Node.js 18+
- Hymnal data in JSON format (see [Hymnal Data](#hymnal-data) below)

## Getting Started

### 1. Clone the repo

```bash
git clone https://github.com/TheRevDrJ/OpenOrder.git
cd OpenOrder
```

### 2. Set up the backend

```bash
pip install -r backend/requirements.txt
```

### 3. Set up the frontend

```bash
cd frontend
npm install
cd ..
```

### 4. Add your hymnal data

See [Hymnal Data](#hymnal-data) below for details on providing your own lyrics.

### 5. Run the app

Start both servers:

```bash
# Backend (from project root)
uvicorn backend.app.main:app --host 0.0.0.0 --port 8316

# Frontend (from frontend/)
cd frontend
npm run dev
```

Then open **http://localhost:5173** (dev) or **http://localhost:8316** (API).

## Hymnal Data

**Important**: Hymnal lyrics are copyrighted material and are **not included** in this repository. You must purchase and license your own hymnal slides.

### Where to Buy Hymnal Slides

The official PowerPoint presentation editions are published by the **United Methodist Publishing House (UMPH)** and available from several retailers (as of March 2026):

**The United Methodist Hymnal — Presentation Edition** (USB flash drive with all hymns, psalter, readings, and orders of worship):
- [Cokesbury](https://www.cokesbury.com/The-United-Methodist-Hymnal-Presentation-Edition)
- [Amazon](https://www.amazon.com/United-Methodist-Hymnal-Presentation/dp/1426795378) (ISBN: 9781426795374)
- [Abingdon Press](https://www.abingdonpress.com/product/9781426795374/)
- [Christianbook.com](https://www.christianbook.com/the-united-methodist-hymnal-presentation-edition/9781426795374/pd/795371)
- [Ministry Matters (digital download)](https://www.ministrymatters.com/product/9781630886349)

**The Faith We Sing — Presentation Edition** (USB flash drive with all 284 songs):
- [Cokesbury](https://www.cokesbury.com/The-Faith-We-Sing-Presentation-Edition)
- [Amazon](https://www.amazon.com/Faith-We-Sing-Presentation/dp/1426795386) (ISBN: 9781426795381)
- [Abingdon Press](https://www.abingdonpress.com/product/9781426795381/)
- [Christianbook.com](https://www.christianbook.com/the-faith-we-sing-presentation-edition/umph-publication/9781426795381/pd/795381)

Individual hymn slides are also available from [Digital Songs and Hymns](https://digitalsongsandhymns.com/collections/united-methodist-hymnal).

Please ensure your usage complies with the applicable license. Additional uses beyond projection in worship may require a blanket license from [OneLicense](https://www.onelicense.net) or [CCLI](https://www.ccli.com).

### Converting Slides to JSON

The purchased hymnal slides come in `.ppt` format (older PowerPoint). OpenOrder needs them converted to JSON. The `scripts/` directory includes tools for this:

**Step 1: Convert `.ppt` to `.pptx`**

The Python `python-pptx` library cannot read the legacy `.ppt` format. Use the included PowerShell script to batch-convert via PowerPoint COM automation (requires Microsoft PowerPoint installed):

```powershell
# From the project root — converts all .ppt files to .pptx
.\scripts\convert_ppt_to_pptx.ps1
```

This opens each `.ppt` file in PowerPoint and saves it as `.pptx` in a temporary `pptx-converted/` directory. The script processes all hymnal folders (Hymns by Number, The Faith We Sing, General Services, etc.) and handles errors gracefully — any files that fail to convert are logged.

> **Note**: Requires Microsoft PowerPoint to be installed. The script uses COM automation, so PowerPoint will briefly open and close for each file. Expect ~1-2 minutes per 100 files.

**Step 2: Extract lyrics to JSON**

```bash
python scripts/extract_lyrics.py
```

This reads each `.pptx` file and extracts the lyrics into JSON files in the `hymnal-json/` directory. The extraction is smart about separating metadata from lyrics:

- **Title slides** are identified by shape position and content (hymn title, number, attribution)
- **Offscreen shapes** (negative top position — a common pattern in the purchased slides) are excluded
- **Copyright lines** are preserved as metadata but separated from lyrics
- **Vertical tab characters** are cleaned up and split into proper line breaks
- An **index file** (`hymnal-json/index.json`) is generated for fast hymn search

After extraction, you can verify the output:

```bash
# Check for any slides with suspiciously high line counts (likely metadata bleed)
python -c "
import json, glob
for f in glob.glob('hymnal-json/**/*.json', recursive=True):
    if 'index' in f: continue
    d = json.load(open(f))
    for i, s in enumerate(d.get('slides', [])):
        if s.get('line_count', 0) > 8:
            print(f'{d[\"number\"]} - {d[\"title\"]}: Slide {i+1} has {s[\"line_count\"]} lines')
"
```

### JSON Format

Each hymn is stored as a JSON file with this structure:

```json
{
  "number": "378",
  "title": "Amazing Grace",
  "source": "umh",
  "slides": [
    {
      "lines": [
        "Amazing Grace",
        "378"
      ]
    },
    {
      "lines": [
        "1. Amazing grace! How sweet the sound",
        "that saved a wretch like me!",
        "I once was lost, but now am found;",
        "was blind, but now I see."
      ]
    }
  ]
}
```

## Project Structure

```
OpenOrder/
├── backend/
│   └── app/
│       ├── main.py          # FastAPI routes
│       ├── bulletin.py      # Word document generation
│       ├── slides.py        # PowerPoint slide generation
│       ├── scripture.py     # Bible API fetch + caching
│       ├── models.py        # Pydantic models
│       ├── hymnal.py        # Hymn search API
│       ├── church_config.py # Church-specific defaults
│       └── themes/          # Slide theme definitions
│           ├── __init__.py  # Theme registry
│           └── oakpark.py   # Oak Park UMC theme
├── frontend/
│   └── src/
│       ├── App.tsx          # Main form UI
│       ├── components/      # HymnPicker, ThemeToggle, shadcn/ui
│       └── lib/             # API client, config, utils
├── resources/
│   ├── images/              # Church logos, QR codes
│   └── slides/              # Static slide backgrounds
├── scripts/
│   ├── convert_ppt_to_pptx.ps1  # PowerShell COM conversion
│   └── extract_lyrics.py        # PPTX to JSON extraction
├── hymnal-json/             # (gitignored) Your hymnal data
├── scripture-cache/         # (gitignored) Cached scripture fetches
├── output/                  # (gitignored) Generated files
└── OpenOrder.bat            # Windows launcher (start/stop/restart)
```

## Themes

OpenOrder uses a **theme system** for bulletin layout and styling. Each theme defines how the bulletin looks — fonts, spacing, section order, and formatting.

The app ships with two built-in themes:
- **Oak Park** — A contemporary layout with centered cover page and olive green accents
- **Traditional** — *(coming soon)* A classic church bulletin format

Slide themes live in `backend/app/themes/` and define fonts, colors, shadow effects, and styling. Each theme is a simple Python dict — no GUI builder needed. Community-contributed themes are welcome! If you'd like a custom theme built from your church's existing bulletin, [reach out](https://github.com/TheRevDrJ/OpenOrder/issues) — we're happy to help.

## Roadmap

- [x] Word bulletin generation
- [x] PowerPoint slide generation
- [x] Hymn search across UMH and TFWS
- [x] Scripture integration with Bible API
- [x] Slide theme system (font, colors, shadows, badges)
- [ ] Bulletin theme system with swappable layouts
- [ ] Load/edit past weeks from saved JSON
- [ ] Drag-and-drop service order customization
- [ ] "Bring your own lyrics" for praise songs
- [ ] Multi-church support with configurable profiles

## Get Involved

If you're a pastor interested in bringing OpenOrder to your church, or a developer who wants to contribute, open an issue or reach out. This is a passion project — the goal is to remove barriers, not build a business.

## License

[AGPL-3.0](LICENSE) — use it, modify it, share it. Free forever. If you distribute or run a modified version as a network service, you must share your changes under the same license.

## Links

- **Website:** [openorderproject.org](https://openorderproject.org)
- **GitHub:** [TheRevDrJ/OpenOrder](https://github.com/TheRevDrJ/OpenOrder)

---

Built with [__Ephphatha__](https://github.com/TheRevDrJ). 🙌
