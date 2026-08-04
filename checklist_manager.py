"""
Checklist Manager
- Copies the pre-made template into the shared /checklists folder under a new name
- Lists existing checklists (visible to any user pointed at the same shared folder)
- Opens a checklist's items for quick viewing in the terminal
"""
import math
import shutil
import re
from pathlib import Path
from datetime import date
import openpyxl
from openpyxl.styles import Alignment

TEMPLATE_PATH = Path("templates/QAQC_Checklist_Template.xlsx")
CHECKLISTS_DIR = Path("checklists")


def _safe_name(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_\- ]+", "", name).strip().replace(" ", "_")
    return name or "Checklist"


def create_checklist(project: str, location: str, inspector: str) -> Path:
    """Copy the master template into the shared checklists folder as a new, named file."""
    CHECKLISTS_DIR.mkdir(exist_ok=True)
    fname = f"{_safe_name(project)}_{_safe_name(location)}_{date.today().isoformat()}.xlsx"
    dest = CHECKLISTS_DIR / fname
    shutil.copy(TEMPLATE_PATH, dest)

    wb = openpyxl.load_workbook(dest)
    ws = wb["Checklist"]
    ws["C4"] = project
    ws["C5"] = location
    ws["C6"] = inspector
    ws["G4"] = date.today().isoformat()
    ws["G6"] = fname.replace(".xlsx", "")
    wb.save(dest)
    return dest


def list_checklists():
    CHECKLISTS_DIR.mkdir(exist_ok=True)
    return sorted(CHECKLISTS_DIR.glob("*.xlsx"))


def read_checklist_items(path: Path):
    """Return list of dicts for each checklist row (Item #, Category, Requirement, Status, Photo, Notes, Action)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Checklist"]
    items = []
    for row in ws.iter_rows(min_row=10, max_col=8, values_only=False):
        item_no = row[1].value  # column B
        if not item_no:
            continue
        items.append({
            "row": row[0].row,
            "item_no": row[1].value,
            "category": row[2].value,
            "requirement": row[3].value,
            "status": row[4].value,
            "photo_ref": row[5].value,
            "notes": row[6].value,
            "corrective_action": row[7].value,
        })
    return items


def update_checklist_row(path: Path, row_number: int, status=None, photo_ref=None, notes=None, corrective_action=None):
    """Write into the checklist's Status / Photo Reference / Notes / Corrective Action columns
    (E, F, G, H) for a given row. Notes and Corrective Action get wrap-text formatting and the
    row height is expanded so longer AI-generated findings stay readable instead of clipping."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Checklist"]
    if status is not None:
        ws.cell(row=row_number, column=5, value=status)
    if photo_ref is not None:
        ws.cell(row=row_number, column=6, value=photo_ref)
    if notes is not None:
        cell = ws.cell(row=row_number, column=7, value=notes)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    if corrective_action is not None:
        cell = ws.cell(row=row_number, column=8, value=corrective_action)
        cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Rough row-height auto-fit so wrapped Notes/Corrective Action text isn't clipped.
    longest_text = max(len(str(notes or "")), len(str(corrective_action or "")))
    if longest_text:
        chars_per_line = 40  # approx for the Notes column width
        est_lines = max(1, math.ceil(longest_text / chars_per_line))
        needed_height = min(est_lines * 15, 400)  # cap so one giant report can't blow out the sheet
        if (ws.row_dimensions[row_number].height or 15) < needed_height:
            ws.row_dimensions[row_number].height = needed_height

    wb.save(path)
