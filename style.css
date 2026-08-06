"""
Submittal Checker
Evaluates files in /submittals against the requirements on a given checklist to flag
whether the correct submittal appears to have been used (e.g. right product, right spec
section, right rating/finish).

Beta scope: works on text-based submittals (.txt/.md extracted spec sheets) and images
(product photos/data sheet screenshots) via Claude vision. PDFs: convert to text first
(see /mnt/skills/public/pdf-reading) — left for the next iteration.
"""
import base64
import mimetypes
import os
from pathlib import Path

import anthropic

import photo_evaluator as pe  # reuse its image resize/EXIF-rotation logic for image submittals

SUBMITTALS_DIR = Path("submittals")
MODEL = "claude-sonnet-4-6"

SYSTEM_PROMPT = """You are a construction submittal reviewer. You are given the requirement \
text from a QA/QC checklist item and one submittal document (text or image). Determine \
whether the submittal appears to match what was specified.

Output:
- **Match Assessment**: Matches / Does Not Match / Unclear — Needs Verification
- **Reasoning**: what in the submittal supports or contradicts the requirement
- **Flags**: anything missing (model number, rating, finish, stamped approval, etc.)
- **Recommendation**: what the user should confirm if unclear

Be conservative — if you can't tell from the content provided, say so and specify exactly \
what additional information is needed.
"""


def list_submittals():
    SUBMITTALS_DIR.mkdir(exist_ok=True)
    return sorted(p for p in SUBMITTALS_DIR.glob("*") if p.is_file())


def check_submittal(submittal_path: Path, requirement_text: str) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is not set. Export it in your shell before running.")

    client = anthropic.Anthropic(api_key=api_key, max_retries=4)  # ride out brief upstream 5xx/502 blips

    content = [{"type": "text", "text": f"Checklist requirement:\n{requirement_text}\n\nSubmittal file: {submittal_path.name}"}]

    if submittal_path.suffix.lower() in {".txt", ".md"}:
        text = submittal_path.read_text(errors="ignore")[:6000]
        content.append({"type": "text", "text": f"Submittal content:\n{text}"})
    elif submittal_path.suffix.lower() in {".png", ".jpg", ".jpeg", ".webp", ".gif"}:
        mime, data = pe._encode_image(submittal_path)
        content.append({"type": "image", "source": {"type": "base64", "media_type": mime, "data": data}})
    else:
        content.append({"type": "text", "text": "(Unsupported file type for beta — convert to .txt or an image first.)"})

    response = client.messages.create(
        model=MODEL,
        max_tokens=1000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": content}],
    )

    out = []
    for block in response.content:
        if block.type == "text":
            out.append(block.text)
    return "\n".join(out) if out else "(No text returned.)"
