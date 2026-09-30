from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from presentation_agent import generate_presentation, ensure_dirs, BASE_DIR, UPLOADS_DIR

app = FastAPI(title="Presentation Agent")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = BASE_DIR / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/health")
def health_check():
    return {"status": "ok", "name": "Presentation Agent"}


@app.get("/")
def home():
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.post("/api/generate")
async def generate(
    files: List[UploadFile] = File(...),
    theme: str = Form("executive"),
    audience: str = Form("leadership team"),
    story: str = Form(""),
):
    ensure_dirs()
    saved_files = []
    for uploaded in files:
        if uploaded.filename:
            destination = UPLOADS_DIR / uploaded.filename
            with destination.open("wb") as handle:
                handle.write(await uploaded.read())
            saved_files.append(destination)

    if not saved_files:
        return {"error": "No files were uploaded."}

    spec = generate_presentation(saved_files, theme=theme, audience=audience, story=story)
    return {
        "status": "success",
        "pptx_url": "/api/download?filename=generated_deck.pptx",
        "summary": {
            "theme": spec["theme"],
            "audience": spec["audience"],
            "story": spec["story"],
            "source_files": spec["source_files"],
        },
        "slides": spec["slides"],
        "cards": spec["cards"],
        "documents": spec["documents"],
        "export_file": spec["export_file"],
    }


@app.get("/api/download")
def download(filename: str):
    export_path = BASE_DIR / "exports" / filename
    if not export_path.exists():
        return {"error": "File not found."}
    return FileResponse(
        str(export_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
