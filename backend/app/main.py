"""Order of Worship — FastAPI backend."""

import json
from datetime import date, timedelta
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from .hymnal import search_hymns, get_hymn, get_hymn_by_ref
from .models import OrderOfWorship
from .bulletin import generate_bulletin

app = FastAPI(title="Order of Worship")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def next_sunday() -> str:
    """Return the date of the upcoming Sunday (or today if it's Sunday)."""
    today = date.today()
    days_until_sunday = (6 - today.weekday()) % 7
    if days_until_sunday == 0 and today.weekday() != 6:
        days_until_sunday = 7
    return (today + timedelta(days=days_until_sunday)).isoformat()


# --- Health / Info ---

@app.get("/api/health")
def health():
    return {"status": "ok", "nextSunday": next_sunday()}


# --- Hymnal ---

@app.get("/api/hymnal/search")
def hymnal_search(q: str = "", limit: int = 20):
    return search_hymns(q, limit)


@app.get("/api/hymnal/{source}/{number}")
def hymnal_get(source: str, number: str):
    hymn = get_hymn_by_ref(source, number)
    if not hymn:
        raise HTTPException(404, "Hymn not found")
    return hymn


# --- Services (save/load) ---

def _service_path(service_date: str) -> Path:
    return OUTPUT_DIR / f"{service_date} - Raw.json"


@app.get("/api/services")
def list_services():
    """List all saved services."""
    files = sorted(OUTPUT_DIR.glob("* - Raw.json"), reverse=True)
    services = []
    for f in files:
        date_str = f.name.split(" - Raw.json")[0]
        services.append({"date": date_str, "filename": f.name})
    return services


@app.get("/api/services/{service_date}")
def get_service(service_date: str):
    path = _service_path(service_date)
    if not path.exists():
        raise HTTPException(404, "Service not found")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


@app.post("/api/services/{service_date}")
def save_service(service_date: str, data: OrderOfWorship):
    data.date = service_date
    path = _service_path(service_date)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data.model_dump(), f, indent=2, ensure_ascii=False)
    return {"saved": True, "path": str(path)}


# --- Theme image upload ---

@app.post("/api/services/{service_date}/theme-image")
async def upload_theme_image(service_date: str, file: UploadFile):
    ext = Path(file.filename).suffix or ".jpg"
    filename = f"{service_date} - Theme{ext}"
    dest = OUTPUT_DIR / filename
    content = await file.read()
    with open(dest, "wb") as f:
        f.write(content)
    return {"filename": filename}


# --- Generate ---

@app.post("/api/generate/bulletin/{service_date}")
def gen_bulletin(service_date: str):
    path = _service_path(service_date)
    if not path.exists():
        raise HTTPException(404, "Service not found — save first")
    with open(path, "r", encoding="utf-8") as f:
        data = OrderOfWorship(**json.load(f))
    try:
        filepath = generate_bulletin(data)
    except PermissionError:
        raise HTTPException(
            409, "The bulletin file is open in another program (probably Word). Close it and try again."
        )
    return {"filename": filepath.name}


# --- File downloads ---

@app.get("/api/download/{filename}")
def download_file(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        raise HTTPException(404, "File not found")
    return FileResponse(path, filename=filename)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8316)
