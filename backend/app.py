from pathlib import Path
from typing import List
import shutil
import uuid

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from presentation_agent import generate_presentation, ensure_dirs, BASE_DIR, UPLOADS_DIR, EXPORTS_DIR

app = FastAPI(title="Presentation Agent")

# CORS middleware for all origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = BASE_DIR / "static"

# Mount static files
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/health")
def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "service": "presentation-agent"}


@app.get("/")
def home():
    """Serve the main page."""
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.post("/api/generate")
async def generate_deck(
    files: List[UploadFile] = File(...),
    theme: str = Form("executive"),
    audience: str = Form("leadership"),
    story: str = Form(""),
    generation_type: str = Form("claude")
):
    """
    Generate a presentation from uploaded files.
    
    Args:
        files: List of uploaded files (PDF, Excel, Word, Images, Text, CSV)
        theme: Presentation theme (executive, sales, product, marketing, investor)
        audience: Target audience (leadership, team, investors, customers, etc.)
        story: Custom narrative/story to incorporate
        generation_type: Generation style (claude, openai, meta)
    
    Returns:
        JSON with presentation spec and export file info
    """
    try:
        ensure_dirs()
        
        # Create session directory
        session_id = str(uuid.uuid4())[:8]
        session_dir = UPLOADS_DIR / session_id
        session_dir.mkdir(exist_ok=True)
        
        # Save uploaded files
        saved_files = []
        for file in files:
            if file.filename:
                file_path = session_dir / file.filename
                contents = await file.read()
                with open(file_path, "wb") as f:
                    f.write(contents)
                saved_files.append(file_path)
        
        if not saved_files:
            return JSONResponse(
                status_code=400,
                content={"error": "No files uploaded"}
            )
        
        # Generate presentation
        spec = generate_presentation(
            files=saved_files,
            theme=theme,
            audience=audience,
            story=story,
            generation_type=generation_type
        )
        
        # Add session info
        spec["session_id"] = session_id
        spec["file_count"] = len(saved_files)
        spec["uploaded_files"] = [f.name for f in saved_files]
        
        return spec
    
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )


@app.get("/api/download/{session_id}")
async def download_deck(session_id: str):
    """Download generated presentation."""
    try:
        deck_path = EXPORTS_DIR / "generated_deck.pptx"
        if not deck_path.exists():
            return JSONResponse(
                status_code=404,
                content={"error": "Presentation not found"}
            )
        
        return FileResponse(
            path=str(deck_path),
            filename=f"presentation_{session_id}.pptx",
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )


@app.get("/api/themes")
def get_themes():
    """Get available themes."""
    return {
        "themes": [
            {
                "value": "executive",
                "label": "Executive",
                "description": "Professional & strategic focus",
                "color": "#1E40AF"
            },
            {
                "value": "sales",
                "label": "Sales",
                "description": "Revenue & growth focused",
                "color": "#158F3B"
            },
            {
                "value": "product",
                "label": "Product",
                "description": "Feature & adoption driven",
                "color": "#7C2D12"
            },
            {
                "value": "marketing",
                "label": "Marketing",
                "description": "Creative & campaign focused",
                "color": "#7E22CE"
            },
            {
                "value": "investor",
                "label": "Investor",
                "description": "Financial & opportunity focused",
                "color": "#0F172A"
            }
        ]
    }


@app.get("/api/audiences")
def get_audiences():
    """Get available audiences."""
    return {
        "audiences": [
            {"value": "leadership", "label": "Leadership / Executives"},
            {"value": "team", "label": "Internal Team"},
            {"value": "investors", "label": "Investors / Stakeholders"},
            {"value": "customers", "label": "Customers / Partners"},
            {"value": "board", "label": "Board of Directors"},
            {"value": "public", "label": "General Public"},
            {"value": "employees", "label": "All Employees"}
        ]
    }


@app.get("/api/generation-types")
def get_generation_types():
    """Get available generation types."""
    return {
        "types": [
            {
                "value": "claude",
                "label": "Claude Style",
                "description": "Intelligent, insight-driven strategic narrative"
            },
            {
                "value": "openai",
                "label": "ChatGPT Style",
                "description": "Structured, logical, systematic analysis"
            },
            {
                "value": "meta",
                "label": "Meta/Muse Style",
                "description": "Creative, visual, engaging storytelling"
            }
        ]
    }


@app.post("/api/cleanup/{session_id}")
async def cleanup_session(session_id: str):
    """Clean up session files."""
    try:
        session_dir = UPLOADS_DIR / session_id
        if session_dir.exists():
            shutil.rmtree(session_dir)
        return {"status": "cleaned"}
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )


if __name__ == "__main__":
    import uvicorn
    ensure_dirs()
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        reload=True
    )
