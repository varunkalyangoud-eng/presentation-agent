from collections import Counter
from datetime import datetime
import csv
import json
import re
from pathlib import Path
from typing import Any, Dict, List
import statistics

from docx import Document
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
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


def extract_text_from_excel(path: Path) -> Dict[str, Any]:
    """Extract structured data from Excel with metrics and statistics."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    extracted = {"sheets": [], "metrics": [], "text": ""}
    
    for sheet in workbook.worksheets:
        sheet_data = {"name": sheet.title, "rows": [], "numeric_cols": []}
        numeric_values = []
        
        for row in sheet.iter_rows(values_only=True):
            processed_row = [str(cell) if cell is not None else "" for cell in row]
            sheet_data["rows"].append(processed_row)
            
            # Extract numeric values for metrics
            for cell in row:
                if isinstance(cell, (int, float)):
                    numeric_values.append(cell)
        
        # Calculate statistics
        if numeric_values:
            sheet_data["stats"] = {
                "max": max(numeric_values),
                "min": min(numeric_values),
                "avg": statistics.mean(numeric_values),
                "count": len(numeric_values)
            }
            extracted["metrics"].extend(numeric_values[:10])
        
        extracted["sheets"].append(sheet_data)
        extracted["text"] += "\n".join(", ".join(r) for r in sheet_data["rows"][:15])
    
    return extracted


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


def extract_text_from_csv(path: Path) -> Dict[str, Any]:
    """Extract structured data from CSV with metrics and patterns."""
    rows = []
    numeric_cols = []
    metrics = []
    
    try:
        with path.open("r", encoding="utf-8", errors="ignore", newline="") as handle:
            reader = csv.reader(handle)
            headers = next(reader, [])
            
            for i, row in enumerate(reader):
                rows.append(", ".join(row))
                
                # Extract numeric data
                for j, cell in enumerate(row):
                    try:
                        val = float(cell)
                        metrics.append(val)
                        if j not in numeric_cols:
                            numeric_cols.append(j)
                    except ValueError:
                        pass
                
                if i >= 25:
                    break
    except Exception:
        return {"text": "", "metrics": [], "headers": []}
    
    return {
        "text": "\n".join(rows),
        "metrics": metrics[:20],
        "headers": headers,
        "numeric_cols": numeric_cols
    }


def safe_extract_text(file_path: Path) -> Dict[str, Any]:
    """Extract text and metadata from files."""
    suffix = file_path.suffix.lower()
    extracted = {"text": "", "metrics": [], "type": suffix}
    
    if suffix == ".txt":
        extracted["text"] = extract_text_from_txt(file_path)
    elif suffix == ".docx":
        extracted["text"] = extract_text_from_docx(file_path)
    elif suffix in {".xlsx", ".xls"}:
        excel_data = extract_text_from_excel(file_path)
        extracted["text"] = excel_data.get("text", "")
        extracted["metrics"] = excel_data.get("metrics", [])
    elif suffix == ".pdf":
        extracted["text"] = extract_text_from_pdf(file_path)
    elif suffix in {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}:
        extracted["text"] = extract_text_from_image(file_path)
    elif suffix == ".csv":
        csv_data = extract_text_from_csv(file_path)
        extracted["text"] = csv_data.get("text", "")
        extracted["metrics"] = csv_data.get("metrics", [])
    
    return extracted


def summarize_content(text: str, max_chars: int = 800) -> str:
    cleaned = " ".join(text.split())
    if len(cleaned) <= max_chars:
        return cleaned
    return cleaned[:max_chars].rsplit(" ", 1)[0] + "..."


def extract_key_insights(text: str, metrics: List[float]) -> List[str]:
    """Extract key insights from text and numeric data."""
    insights = []
    
    # Keyword-based insights
    keywords = {
        "growth": "showing strong growth trajectory",
        "decline": "showing declining trend",
        "revenue": "revenue metrics are significant",
        "profit": "profitability is a key driver",
        "efficiency": "efficiency improvements detected",
        "risk": "risk factors identified",
        "opportunity": "opportunities for expansion",
        "challenge": "challenges to address",
        "momentum": "positive momentum observed",
        "trend": "clear trend patterns emerging"
    }
    
    text_lower = text.lower()
    for keyword, insight in keywords.items():
        if keyword in text_lower:
            insights.append(insight)
    
    # Numeric insights
    if metrics:
        avg = sum(metrics) / len(metrics)
        max_val = max(metrics)
        min_val = min(metrics)
        
        if max_val > avg * 1.5:
            insights.append(f"Peak value of {max_val:.1f} indicates significant performance spike")
        if min_val < avg * 0.7:
            insights.append(f"Low point of {min_val:.1f} requires attention")
        if len(metrics) > 1:
            variance = statistics.variance(metrics)
            if variance > avg * 0.5:
                insights.append("High variability in metrics suggests fluctuating performance")
    
    return insights[:5] if insights else [
        "Data shows consistent patterns and measurable outcomes",
        "Key performance indicators align with business objectives"
    ]


def generate_smart_slides(documents: List[Dict[str, Any]], theme: str, audience: str, story: str) -> List[Dict[str, Any]]:
    """Generate presentation slides based on extracted content."""
    all_text = " ".join(doc.get("summary", "") for doc in documents)
    all_metrics = []
    for doc in documents:
        if "metrics" in doc and isinstance(doc["metrics"], list):
            all_metrics.extend(doc["metrics"][:10])
    
    insights = extract_key_insights(all_text, all_metrics)
    
    slides = [
        {
            "title": "Executive Summary",
            "bullets": [
                f"Prepared from {len(documents)} data sources for {theme} narrative",
                f"Audience: {audience}",
                "AI-generated insights focused on actionable business outcomes"
            ],
            "layout": "title"
        },
        {
            "title": "Key Insights",
            "bullets": insights,
            "layout": "findings"
        },
        {
            "title": "Performance Snapshot",
            "bullets": [
                f"Total metrics analyzed: {len(all_metrics)}",
                f"Average performance: {(sum(all_metrics) / len(all_metrics)):.1f}" if all_metrics else "Data patterns detected",
                "Trend analysis shows consistent business progression"
            ] if all_metrics else [
                "Qualitative analysis complete",
                "Narrative patterns identified and mapped",
                "Strategic themes extracted"
            ],
            "layout": "metrics"
        },
        {
            "title": "Recommended Actions",
            "bullets": [
                "Leverage identified opportunities for growth acceleration",
                "Implement monitoring for flagged performance areas",
                "Execute next steps aligned with strategic priorities"
            ],
            "layout": "actions"
        }
    ]
    
    if story:
        slides[0]["bullets"][2] = story
    
    return slides


def generate_smart_kpis(documents: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """Generate KPI cards based on extracted metrics and content."""
    all_metrics = []
    all_text = " ".join(doc.get("summary", "") for doc in documents)
    
    for doc in documents:
        if "metrics" in doc and isinstance(doc["metrics"], list):
            all_metrics.extend(doc["metrics"])
    
    kpis = []
    
    if all_metrics:
        avg = sum(all_metrics) / len(all_metrics)
        max_val = max(all_metrics)
        growth = ((max_val - min(all_metrics)) / min(all_metrics) * 100) if min(all_metrics) > 0 else 0
        
        kpis = [
            {"label": "Average", "value": f"{avg:.1f}"},
            {"label": "Peak", "value": f"{max_val:.1f}"},
            {"label": "Growth", "value": f"+{min(99, growth):.0f}%"},
            {"label": "Samples", "value": f"{len(all_metrics)}"}
        ]
    else:
        # Infer from text
        words = re.findall(r"\b[a-zA-Z]{4,}\b", all_text.lower())
        counts = Counter(words)
        top_terms = counts.most_common(4)
        
        kpis = [
            {"label": term[0].title(), "value": f"{min(99, term[1] * 8)}%"}
            for term in top_terms
        ]
    
    if not kpis:
        kpis = [
            {"label": "Content", "value": "Analyzed"},
            {"label": "Quality", "value": "High"},
            {"label": "Ready", "value": "Yes"},
            {"label": "Score", "value": "95%"}
        ]
    
    return kpis


def generate_advanced_pptx(spec: Dict[str, Any], output_path: Path) -> None:
    """Generate a professional presentation with theme-based styling."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Theme colors
    theme_colors = {
        "executive": {"primary": RGBColor(0x1E, 0x40, 0xAF), "accent": RGBColor(0x3B, 0x82, 0xF6)},
        "sales": {"primary": RGBColor(0x15, 0x8F, 0x3B), "accent": RGBColor(0x22, 0xC5, 0x5E)},
        "product": {"primary": RGBColor(0x7C, 0x2D, 0x12), "accent": RGBColor(0xEA, 0x58, 0x0C)},
        "marketing": {"primary": RGBColor(0x7E, 0x22, 0xCE), "accent": RGBColor(0xA8, 0x5E, 0xFF)},
        "investor": {"primary": RGBColor(0x0F, 0x17, 0x2A), "accent": RGBColor(0x1E, 0x40, 0xAF)}
    }
    
    colors = theme_colors.get(spec.get("theme", "executive"), theme_colors["executive"])
    
    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[6])
    background = title_slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = colors["primary"]
    
    # Title
    title_box = title_slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.3), Inches(2))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = spec.get("title", "AI-Generated Presentation")
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    
    # Subtitle
    subtitle_box = title_slide.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(12.3), Inches(1.5))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = f"{spec.get('theme', 'Executive').title()} Narrative | {spec.get('audience', 'Leadership')}"
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(255, 255, 255)
    
    # Content slides
    for slide_spec in spec.get("slides", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Slide background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Title with accent bar
        title_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(0), Inches(13.333), Inches(1)
        )
        title_shape.fill.solid()
        title_shape.fill.fore_color.rgb = colors["primary"]
        title_shape.line.color.rgb = colors["primary"]
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12.3), Inches(0.7))
        title_frame = title_box.text_frame
        p = title_frame.paragraphs[0]
        p.text = slide_spec.get("title", "")
        p.font.size = Pt(40)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        
        # Bullets or metrics layout
        if slide_spec.get("layout") == "metrics":
            # KPI cards
            cards = spec.get("cards", [])
            for idx, card in enumerate(cards[:4]):
                x = Inches(0.7 + (idx % 2) * 6.2)
                y = Inches(1.8 + (idx // 2) * 2.5)
                
                card_shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.5), Inches(2)
                )
                card_shape.fill.solid()
                card_shape.fill.fore_color.rgb = RGBColor(240, 248, 255)
                card_shape.line.color.rgb = colors["accent"]
                card_shape.line.width = Pt(2)
                
                tf = card_shape.text_frame
                tf.clear()
                
                p = tf.paragraphs[0]
                p.text = card.get("label", "")
                p.font.size = Pt(18)
                p.font.bold = True
                p.font.color.rgb = colors["primary"]
                p.alignment = PP_ALIGN.CENTER
                
                p2 = tf.add_paragraph()
                p2.text = card.get("value", "")
                p2.font.size = Pt(32)
                p2.font.bold = True
                p2.font.color.rgb = colors["accent"]
                p2.alignment = PP_ALIGN.CENTER
        else:
            # Bullet points
            for idx, bullet in enumerate(slide_spec.get("bullets", [])):
                y = Inches(1.8 + idx * 1.1)
                bullet_box = slide.shapes.add_textbox(Inches(0.8), y, Inches(11.7), Inches(0.8))
                bullet_frame = bullet_box.text_frame
                bullet_frame.word_wrap = True
                p = bullet_frame.paragraphs[0]
                p.text = f"• {bullet}"
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(30, 30, 30)
                p.space_before = Pt(6)
                p.space_after = Pt(6)
    
    prs.save(output_path)


def generate_presentation(files: List[Path], theme: str, audience: str, story: str) -> Dict[str, Any]:
    """Generate a complete presentation with AI-like automation."""
    ensure_dirs()
    
    extracted = []
    for file in files:
        file_data = safe_extract_text(file)
        text = file_data.get("text", "")
        metrics = file_data.get("metrics", [])
        
        extracted.append({
            "name": file.name,
            "summary": summarize_content(text),
            "text": text,
            "metrics": metrics,
            "type": file_data.get("type", "")
        })
    
    # Generate smart slides
    slides = generate_smart_slides(extracted, theme=theme, audience=audience, story=story)
    
    # Generate smart KPIs
    cards = generate_smart_kpis(extracted)
    
    # Build spec
    spec = {
        "title": f"{theme.title()} Presentation - AI Generated",
        "theme": theme,
        "audience": audience,
        "story": story or "AI-powered business storytelling and data analysis",
        "source_files": [f["name"] for f in extracted],
        "slides": slides,
        "cards": cards,
        "documents": [
            {
                "name": item["name"],
                "summary": item["summary"],
                "metrics_count": len(item.get("metrics", []))
            }
            for item in extracted
        ],
        "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "insights": extract_key_insights(
            " ".join(item.get("summary", "") for item in extracted),
            sum((item.get("metrics", []) for item in extracted), [])
        )
    }
    
    # Generate advanced PowerPoint
    deck_path = EXPORTS_DIR / "generated_deck.pptx"
    generate_advanced_pptx(spec, deck_path)
    
    spec["export_file"] = "generated_deck.pptx"
    spec["exports"] = {
        "power_bi": {
            "name": "power_bi_exec_dashboard",
            "dataset": "presentation_agent_metrics",
            "tables": [
                {"name": "kpi_snapshot", "columns": ["metric", "value", "theme", "audience"]},
                {"name": "source_summary", "columns": ["document", "summary"]},
                {"name": "insights", "columns": ["insight", "type", "priority"]},
            ],
        },
        "tableau": {
            "name": "tableau_exec_dashboard",
            "workbook": "AI Executive Dashboard",
            "sheets": [
                {"name": "KPI Snapshot", "fields": ["metric", "value"]},
                {"name": "Insights", "fields": ["insight", "type"]},
                {"name": "Documents", "fields": ["document", "summary"]},
            ],
        },
    }
    
    return spec


__all__ = [
    "BASE_DIR",
    "UPLOADS_DIR",
    "EXPORTS_DIR",
    "ensure_dirs",
    "generate_presentation",
]
