"""
PDF Report Builder
Renders a single evaluation (one photo + its report text) as a PDF: photo near the top, findings
organized below by outcome (Issues Found / Unclear / No Issues Found / Follow-Up). Designed to
fit one page for a typical result — reportlab's Platypus layout only spills onto a second page
when the content genuinely doesn't fit (e.g. a photo with many high-priority defects), which is
exactly the "one page unless there's a lot to say" behavior this needs, without any manual
page-fitting logic.
"""
import html
import io
import re
from datetime import datetime
from pathlib import Path

from PIL import Image as PILImage, ImageOps
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer

import photo_evaluator as pe

NAVY = colors.HexColor("#1C2B3A")
ORANGE = colors.HexColor("#E8590C")
STATUS_COLORS = {
    "Fail": colors.HexColor("#C0392B"),
    "Needs Verification": colors.HexColor("#B7791F"),
    "Pass": colors.HexColor("#2F855A"),
    "N/A": colors.HexColor("#6B7280"),
}

MAX_PHOTO_WIDTH = 6.3 * inch
MAX_PHOTO_HEIGHT = 3.0 * inch


def _styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("TitleCompact", parent=base["Heading1"], fontSize=15, spaceAfter=2, textColor=NAVY),
        "meta": ParagraphStyle("Meta", parent=base["Normal"], fontSize=8.5, textColor=colors.grey, spaceAfter=10),
        "status": ParagraphStyle("Status", parent=base["Heading2"], fontSize=12, spaceAfter=10),
        "section": ParagraphStyle("SectionHeader", parent=base["Heading3"], fontName="Helvetica-Bold",
                                   fontSize=11, spaceBefore=10, spaceAfter=4, textColor=NAVY),
        "body": ParagraphStyle("Body", parent=base["Normal"], fontSize=9, leading=12, alignment=TA_LEFT),
        "bullet": ParagraphStyle("Bullet", parent=base["Normal"], fontSize=9, leading=12,
                                  leftIndent=14, bulletIndent=2, spaceAfter=4),
        "compact": ParagraphStyle("Compact", parent=base["Normal"], fontSize=8.5, leading=11,
                                   textColor=colors.HexColor("#374151")),
    }


def _md_to_rl(text: str) -> str:
    """Escape XML-sensitive characters, then convert **bold** markdown to reportlab's <b> tags."""
    text = html.escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    text = text.replace("\n", "<br/>")
    return text


def _split_items(text: str) -> list:
    """Split a section's raw text into individual findings. The model separates distinct findings
    with a blank line; falls back to per-line splitting for simple lists (No Issues Found,
    Follow-Up bullets) or a single unbroken finding."""
    text = (text or "").strip()
    if not text:
        return []
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    if len(paragraphs) > 1:
        return paragraphs
    lines = [ln.strip().lstrip("-").strip() for ln in text.split("\n") if ln.strip()]
    return lines if lines else [text]


def _fit_photo(photo_path: Path):
    """Return an RLImage flowable sized to fit the page, or a small text notice if the file
    can't be decoded as an image — so a bad file produces a readable PDF instead of a 500."""
    try:
        with PILImage.open(photo_path) as img:
            img = ImageOps.exif_transpose(img)
            img = img.convert("RGB")
            w, h = img.size
            scale = min(MAX_PHOTO_WIDTH / w, MAX_PHOTO_HEIGHT / h)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85)
            buf.seek(0)
            return RLImage(buf, width=w * scale, height=h * scale)
    except Exception:
        return Paragraph("[Photo could not be rendered]", ParagraphStyle(
            "PhotoError", fontSize=9, textColor=colors.grey))


def build_report_pdf(photo_path: Path, photo_display_name: str, report_text: str, notes: str = "") -> bytes:
    """Build the PDF and return its bytes (caller decides whether to save, send_file, or attach
    to an email)."""
    styles = _styles()
    sections = pe.parse_report_sections(report_text)
    status = pe.suggest_status_from_report(report_text)

    story = [_fit_photo(photo_path), Spacer(1, 10)]

    story.append(Paragraph(html.escape(photo_display_name), styles["title"]))
    meta_bits = [f"Evaluated {datetime.now().strftime('%b %d, %Y %I:%M %p')}"]
    if notes:
        meta_bits.append(html.escape(notes))
    story.append(Paragraph(" &nbsp;|&nbsp; ".join(meta_bits), styles["meta"]))

    status_style = ParagraphStyle("StatusColored", parent=styles["status"],
                                   textColor=STATUS_COLORS.get(status, colors.black))
    story.append(Paragraph(f"Status: {status}", status_style))

    issues = _split_items(sections["issues_found"])
    if issues and sections["issues_found"].strip().lower() != "none.":
        story.append(Paragraph("Issues Found", styles["section"]))
        for item in issues:
            story.append(Paragraph(f"&bull; {_md_to_rl(item)}", styles["bullet"]))
    else:
        story.append(Paragraph("<b>Issues Found:</b> None.", styles["body"]))

    unclear = _split_items(sections["unclear"])
    if unclear and sections["unclear"].strip().lower() != "none.":
        story.append(Paragraph("Unclear / Field Team to Verify", styles["section"]))
        for item in unclear:
            story.append(Paragraph(f"&bull; {_md_to_rl(item)}", styles["bullet"]))

    no_issues = _split_items(sections["no_issues"])
    if no_issues:
        story.append(Paragraph("No Issues Found", styles["section"]))
        story.append(Paragraph(html.escape(", ".join(no_issues)), styles["compact"]))

    followup = _split_items(sections["followup"])
    if followup and sections["followup"].strip().lower() != "none.":
        story.append(Paragraph("Follow-Up Recommended", styles["section"]))
        for item in followup:
            story.append(Paragraph(f"&bull; {_md_to_rl(item)}", styles["bullet"]))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=0.5 * inch, bottomMargin=0.5 * inch, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
        title=f"QA/QC Report - {photo_display_name}",
    )
    doc.build(story)
    return buf.getvalue()
