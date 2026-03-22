# OpenOrder

A web-based worship bulletin and presentation slide generator for churches.

OpenOrder streamlines the weekly task of preparing Sunday worship materials. Enter your service details — hymns, scripture, sermon title — and generate a print-ready Word bulletin and PowerPoint presentation in seconds.

## Features

- **Hymn Search** — Search by number or title across the United Methodist Hymnal (UMH) and The Faith We Sing (TFWS)
- **Bulletin Generation** — Produces a formatted `.docx` bulletin ready to print and fold
- **Slide Generation** — *(coming soon)* Produces a `.pptx` presentation with hymn lyrics, backgrounds, and static slides
- **Dark Mode** — Light and dark themes for late-night sermon prep
- **Configurable** — Church name, staff, schedule, and branding are driven by config variables (template builder coming soon)

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

**Important**: Hymnal lyrics are copyrighted material and are **not included** in this repository.

To use OpenOrder, you need to provide your own hymnal data as JSON files in a `hymnal-json/` directory at the project root. Each hymn should be a JSON file with this structure:

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

A conversion pipeline is included in `scripts/` to extract lyrics from existing `.pptx` hymnal files.

## Project Structure

```
OpenOrder/
├── backend/
│   └── app/
│       ├── main.py          # FastAPI routes
│       ├── bulletin.py      # Word document generation
│       ├── models.py        # Pydantic models
│       ├── hymnal.py        # Hymn search API
│       └── church_config.py # Church-specific defaults
├── frontend/
│   └── src/
│       ├── App.tsx          # Main form UI
│       ├── components/      # HymnPicker, ThemeToggle, shadcn/ui
│       └── lib/             # API client, config, utils
├── resources/
│   ├── images/              # Church logos, QR codes
│   └── slides/              # Static slide backgrounds
├── scripts/                 # Hymnal conversion tools
├── hymnal-json/             # (gitignored) Your hymnal data
└── output/                  # (gitignored) Generated files
```

## Roadmap

- [ ] PowerPoint slide generation
- [ ] Bulletin template builder (customize layout, fonts, sections)
- [ ] Load/edit past weeks from saved JSON
- [ ] "Bring your own lyrics" for praise songs
- [ ] Multi-church support with configurable profiles

## License

TBD

---

Built with ☕ and [Claude Code](https://claude.ai/claude-code).
