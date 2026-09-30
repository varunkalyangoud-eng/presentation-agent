from collections import Counter
from datetime import datetime
import csv
import json
import re
from pathlib import Path
from typing import Any, Dict, List

from docx import Document
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pypdf import PdfReader
import pytesseract


BASE_DIR = Path(__file__).resolve().parent.parent
UPLOADS_DIR = BASE_DIR / "uploads"
EXPORTS_DIR = BASE_DIR / "exports"


def ensure_dirs() -> None:
    UPLOADS_DIR.mkdir(exist_ok=True)
    EXPORTS_DIR.mkdir(exist_ok=True)


def extract_text_from_txt(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def extract_text_from_docx(path: Path) -> str:
    doc = Document(path)
    return "\n".join(paragraph.text for paragraph in doc.paragraphs)


def extract_text_from_excel(path: Path) -> str:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheets = []
    for sheet in workbook.worksheets:
        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append([str(cell) if cell is not None else "" for cell in row])
        sheets.append("\n".join(", ".join(r) for r in rows[:15]))
    return "\n\n".join(sheets)


def extract_text_from_pdf(path: Path) -> str:
    reader = PdfReader(str(path))
    text_chunks = []
    for page in reader.pages:
        text_chunks.append(page.extract_text() or "")
    return "\n".join(text_chunks)


def extract_text_from_image(path: Path) -> str:
    try:
        img = Image.open(path)
        return pytesseract.image_to_string(img).strip()
    except Exception:
        return ""


def extract_text_from_csv(path: Path) -> str:
    rows = []
    try:
        with path.open("r", encoding="utf-8", errors="ignore", newline="") as handle:
            reader = csv.reader(handle)
            for i, row in enumerate(reader):
                rows.append(", ".join(row))
                if i >= 25:
                    break
    except Exception:
        return ""
    return "\n".join(rows)


def safe_extract_text(file_path: Path) -> str:
    suffix = file_path.suffix.lower()
    if suffix == ".txt":
        return extract_text_from_txt(file_path)
    if suffix == ".docx":
        return extract_text_from_docx(file_path)
    if suffix in {".xlsx", ".xls"}:
        return extract_text_from_excel(file_path)
    if suffix == ".pdf":
        return extract_text_from_pdf(file_path)
    if suffix in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}:
        return extract_text_from_image(file_path)
    if suffix == ".csv":
        return extract_text_from_csv(file_path)
    return ""


def summarize_content(text: str, max_chars: int = 800) -> str:
    cleaned = " ".join(text.split())
    if len(cleaned) <= max_chars:
        return cleaned
    return cleaned[:max_chars].rsplit(" ", 1)[0] + "..."


def build_story_summary(file_names: List[str], theme: str, audience: str, custom_story: str) -> Dict[str, Any]:
    findings = [
        "Strong business momentum and measurable opportunity across the current dataset.",
        "Key operational signals point to growth, efficiency, or risk areas that should be highlighted.",
        "Recommended next steps require clear alignment between leadership, teams, and execution priorities.",
    ]
    if custom_story:
        findings[0] = custom_story

    slides = [
        {
            "title": "Executive Summary",
            "bullets": [
                f"Prepared from {len(file_names)} uploaded files for the {theme} narrative.",
                f"Audience: {audience}",
                "Focus on performance trends, opportunities, and action-oriented recommendations.",
            ],
        },
        {
            "title": "Key Findings",
            "bullets": findings,
        },
        {
            "title": "Recommended Actions",
            "bullets": [
                "Prioritize the biggest growth and efficiency levers.",
                "Validate assumptions with a short review loop across stakeholders.",
                "Translate insights into a concrete dashboard and deliverable timeline.",
            ],
        },
    ]

    return {
        "theme": theme,
        "audience": audience,
        "story": custom_story or "Turn complex data into a persuasive business story with clear decision points.",
        "source_files": file_names,
        "slides": slides,
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
    }


def infer_kpis(documents: List[Dict[str, str]]) -> List[Dict[str, str]]:
    combined = " ".join(doc.get("summary", "") for doc in documents if doc.get("summary"))
    words = re.findall(r"\b[a-zA-Z]{4,}\b", combined.lower())
    counts = Counter(words)
    top_terms = counts.most_common(3)

    deltas = [
        {"label": "Growth", "value": "+24%"},
        {"label": "Revenue", "value": "$4.8M"},
        {"label": "Conversion", "value": "18.6%"},
        {"label": "Efficiency", "value": "92%"},
    ]

    if top_terms:
        deltas[0]["label"] = top_terms[0][0].title()
        deltas[0]["value"] = f"{max(10, min(99, top_terms[0][1] * 8))}%"

    return deltas


def build_advanced_insights(documents: List[Dict[str, str]], theme: str, audience: str) -> Dict[str, Any]:
    kpis = infer_kpis(documents)
    narrative = (
        f"{theme.title()} narrative for {audience}: focus on momentum, risk, and the highest-value actions "
        "with clear business context and measurable impact."
    )
    return {
        "kpis": kpis,
        "narrative": narrative,
        "power_bi_export": {
            "name": "power_bi_exec_dashboard",
            "dataset": "presentation_agent_metrics",
            "tables": [
                {"name": "kpi_snapshot", "columns": ["metric", "value", "theme", "audience"]},
                {"name": "source_summary", "columns": ["document", "summary"]},
            ],
        },
        "tableau_export": {
            "name": "tableau_exec_dashboard",
            "workbook": "Executive Storytelling Dashboard",
            "sheets": [
                {"name": "KPI Snapshot", "fields": ["metric", "value"]},
                {"name": "Source Summary", "fields": ["document", "summary"]},
            ],
        },
        "recommended_actions": [
            "Prioritize performance drivers with the highest revenue and efficiency upside.",
            "Translate insights into leadership-ready decisions and milestones.",
            "Build a tracking cadence for the next reporting cycle.",
        ],
    }


def create_infographic_cards(spec: Dict[str, Any]) -> List[Dict[str, str]]:
    if "advanced_insights" in spec and spec["advanced_insights"].get("kpis"):
        return spec["advanced_insights"]["kpis"]
    return [
        {"label": "Growth", "value": "+24%"},
        {"label": "Revenue", "value": "$4.8M"},
        {"label": "Conversion", "value": "18.6%"},
        {"label": "Efficiency", "value": "92%"},
    ]


def generate_pptx(spec: Dict[str, Any], output_path: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = spec.get("title", "Business Presentation")
    title_slide.placeholders[1].text = f"{spec.get('theme', 'Executive').title()} View | {spec.get('audience', 'Team')}"

    summary_slide = prs.slides.add_slide(prs.slide_layouts[1])
    summary_slide.shapes.title.text = "Executive Summary"
    body = summary_slide.shapes.placeholders[1].text_frame
    body.clear()
    for idx, item in enumerate(spec["slides"][0]["bullets"]):
        paragraph = body.paragraphs[0] if idx == 0 else body.add_paragraph()
        paragraph.text = item
        paragraph.level = 0
        paragraph.font.size = Pt(18)

    findings_slide = prs.slides.add_slide(prs.slide_layouts[5])
    findings_slide.shapes.title.text = "Key Findings"
    for idx, bullet in enumerate(spec["slides"][1]["bullets"]):
        box = findings_slide.shapes.add_textbox(Inches(0.8), Inches(1.4 + idx * 1.1), Inches(11.0), Inches(0.5))
        tf = box.text_frame
        tf.text = f"• {bullet}"
        tf.word_wrap = True
        tf.paragraphs[0].font.size = Pt(20)

    impact_slide = prs.slides.add_slide(prs.slide_layouts[5])
    impact_slide.shapes.title.text = "Impact Snapshot"
    cards = create_infographic_cards(spec)
    for idx, card in enumerate(cards):
        x = Inches(0.7 + (idx % 2) * 5.8)
        y = Inches(1.6 + (idx // 2) * 2.0)
        shape = impact_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(4.7), Inches(1.2))
        shape.fill.fore_color.rgb = 0xEAF2FF
        shape.line.color.rgb = 0x3B82F6
        tf = shape.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = card["label"]
        p.font.size = Pt(18)
        p.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph()
        p2.text = card["value"]
        p2.font.bold = True
        p2.font.size = Pt(28)
        p2.alignment = PP_ALIGN.CENTER

    action_slide = prs.slides.add_slide(prs.slide_layouts[5])
    action_slide.shapes.title.text = "Recommended Actions"
    for idx, item in enumerate(spec["slides"][2]["bullets"]):
        box = action_slide.shapes.add_textbox(Inches(0.9), Inches(1.5 + idx * 1.2), Inches(11.0), Inches(0.7))
        tf = box.text_frame
        tf.text = f"{idx + 1}. {item}"
        tf.paragraphs[0].font.size = Pt(22)

    prs.save(output_path)


def generate_presentation(files: List[Path], theme: str, audience: str, story: str) -> Dict[str, Any]:
    ensure_dirs()
    extracted = []
    for file in files:
        text = safe_extract_text(file)
        extracted.append({
            "name": file.name,
            "summary": summarize_content(text),
            "text": text,
        })

    spec = build_story_summary([f["name"] for f in extracted], theme=theme, audience=audience, custom_story=story)
    spec["title"] = f"{theme.title()} Presentation"
    spec["advanced_insights"] = build_advanced_insights(extracted, theme=theme, audience=audience)
    spec["cards"] = create_infographic_cards(spec)
    spec["documents"] = [{"name": item["name"], "summary": item["summary"]} for item in extracted]
    spec["insights"] = spec["advanced_insights"]
    spec["exports"] = {
        "power_bi": spec["advanced_insights"]["power_bi_export"],
        "tableau": spec["advanced_insights"]["tableau_export"],
    }

    deck_path = EXPORTS_DIR / "generated_deck.pptx"
    generate_pptx(spec, deck_path)
    spec["export_file"] = "generated_deck.pptx"
    return spec


__all__ = [
    "BASE_DIR",
    "UPLOADS_DIR",
    "EXPORTS_DIR",
    "ensure_dirs",
    "generate_presentation",
]
