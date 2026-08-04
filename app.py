"""
Construction QA/QC — Web App entry point.

Run locally:   python app.py
Run in prod:   gunicorn app:app   (see Procfile)

This wraps the existing CLI modules (checklist_manager.py, photo_evaluator.py,
submittal_checker.py) completely unmodified — they don't know or care whether they're being
called from a terminal or a web form.

DATA_DIR (env var) points at a persistent disk. Every data folder the existing modules already
use (library/, checklists/, submittals/, outputs/, templates/, input/) is created under it, by
changing the process's working directory once at startup — since all the existing modules use
relative Paths like Path("library"), this is enough to redirect everything without touching them.
"""
import os
import sys
import secrets
from datetime import datetime
from pathlib import Path

from flask import (Flask, render_template, request, redirect, url_for,
                    send_file, session, flash, abort)
from werkzeug.utils import secure_filename

# Make sure Python can always find the app's own modules (checklist_manager.py etc.) regardless
# of working directory, since we're about to chdir into the persistent data disk below.
sys.path.insert(0, str(Path(__file__).resolve().parent))

# --- Point every relative path used by the existing modules at the persistent disk ---
DATA_DIR = Path(os.environ.get("DATA_DIR", "./data")).resolve()
DATA_DIR.mkdir(parents=True, exist_ok=True)
os.chdir(DATA_DIR)

for folder in ["checklists", "library", "library/photos", "input", "submittals", "outputs", "templates"]:
    Path(folder).mkdir(parents=True, exist_ok=True)

import build_template
import checklist_manager as cm
import photo_evaluator as pe
import submittal_checker as sc

if not Path("templates/QAQC_Checklist_Template.xlsx").exists():
    build_template.build_template()

ALLOWED_IMAGE_EXT = {".jpg", ".jpeg", ".png", ".webp"}

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", secrets.token_hex(32))

# --- Optional lightweight shared-access gate ---------------------------------------------
# Set SITE_PASSWORD as an env var on the server to require a shared passphrase before anyone can
# use the app (recommended once this is on a public URL, to stop random visitors from triggering
# paid API calls). Leave it unset to keep the app fully open.
SITE_PASSWORD = os.environ.get("SITE_PASSWORD")


@app.before_request
def require_site_password():
    if not SITE_PASSWORD:
        return
    if request.endpoint in {"login", "static"}:
        return
    if not session.get("authenticated"):
        return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if request.form.get("password") == SITE_PASSWORD:
            session["authenticated"] = True
            return redirect(request.args.get("next") or url_for("home"))
        flash("Incorrect password.", "error")
    return render_template("login.html")


# --- Home -----------------------------------------------------------------------------------

@app.route("/")
def home():
    return render_template("home.html")


# --- Checklists -------------------------------------------------------------------------------

@app.route("/checklists")
def checklists_list():
    files = cm.list_checklists()
    return render_template("checklists_list.html", files=files)


@app.route("/checklists/new", methods=["GET", "POST"])
def checklists_new():
    if request.method == "POST":
        project = request.form.get("project", "").strip()
        location = request.form.get("location", "").strip()
        inspector = request.form.get("inspector", "").strip()
        if not project:
            flash("Project name is required.", "error")
            return render_template("checklists_new.html")
        path = cm.create_checklist(project, location, inspector)
        flash(f"Created checklist: {path.name}", "success")
        return redirect(url_for("checklists_view", filename=path.name))
    return render_template("checklists_new.html")


@app.route("/checklists/<filename>")
def checklists_view(filename):
    path = _safe_checklist_path(filename)
    items = cm.read_checklist_items(path)
    return render_template("checklists_view.html", filename=filename, items=items)


@app.route("/checklists/<filename>/download")
def checklists_download(filename):
    path = _safe_checklist_path(filename)
    return send_file(path, as_attachment=True, download_name=filename)


def _safe_checklist_path(filename: str) -> Path:
    path = (cm.CHECKLISTS_DIR / secure_filename(filename)).resolve()
    if cm.CHECKLISTS_DIR.resolve() not in path.parents or not path.exists():
        abort(404)
    return path


# --- Evaluate photos --------------------------------------------------------------------------

@app.route("/evaluate", methods=["GET", "POST"])
def evaluate():
    if request.method == "GET":
        diag = pe.diagnose_criteria_setup()
        return render_template("evaluate.html", diag=diag)

    files = [f for f in request.files.getlist("photos") if f and f.filename]
    if not files:
        flash("Choose at least one photo to evaluate.", "error")
        return redirect(url_for("evaluate"))

    notes = request.form.get("notes", "").strip()
    results = []
    for f in files:
        ext = Path(f.filename).suffix.lower()
        if ext not in ALLOWED_IMAGE_EXT:
            results.append({"photo_name": f.filename, "error": f"Unsupported file type: {ext or '(none)'}"})
            continue

        token = secrets.token_hex(4)
        safe_name = f"{token}_{secure_filename(f.filename)}"
        saved_path = Path("input") / safe_name
        f.save(saved_path)

        try:
            report_text = pe.evaluate_photo(saved_path, extra_context=notes)
        except RuntimeError as e:
            results.append({"photo_name": f.filename, "error": str(e)})
            continue

        report_filename = f"{saved_path.stem}_QAQC_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        (Path("outputs") / report_filename).write_text(report_text)

        results.append({
            "photo_name": f.filename,
            "saved_photo_name": safe_name,
            "report_filename": report_filename,
            "report_text": report_text,
            "suggested_status": pe.suggest_status_from_report(report_text),
        })

    return render_template("evaluate_result.html", results=results)


@app.route("/evaluate/save-to-library", methods=["POST"])
def evaluate_save_to_library():
    photo_name = request.form.get("photo")
    report_filename = request.form.get("report")
    photo_path = Path("input") / secure_filename(photo_name or "")
    if not photo_path.exists():
        flash("Original photo no longer available on the server.", "error")
        return redirect(url_for("evaluate"))

    comment = request.form.get("comment", "").strip()
    if not comment:
        report_path = Path("outputs") / secure_filename(report_filename or "")
        comment = report_path.read_text()[:500] if report_path.exists() else ""

    dest = pe.append_to_library(
        photo_path,
        category=request.form.get("category", "").strip(),
        comment=comment,
        severity=request.form.get("severity", "").strip(),
        keywords=request.form.get("keywords", "").strip(),
        added_by=request.form.get("added_by", "").strip(),
    )
    flash(f"Saved to library: {dest.name}", "success")
    return redirect(url_for("evaluate"))


@app.route("/evaluate/link", methods=["GET", "POST"])
def evaluate_link():
    photo_name = request.args.get("photo") or request.form.get("photo")
    report_filename = request.args.get("report") or request.form.get("report")
    report_path = Path("outputs") / secure_filename(report_filename or "")
    if not report_path.exists():
        abort(404)
    report_text = report_path.read_text()

    if request.method == "POST":
        selection = request.form.get("checklist_row", "")
        try:
            checklist_filename, row_str = selection.split("|", 1)
            row_number = int(row_str)
        except ValueError:
            flash("Choose a checklist item first.", "error")
            return redirect(url_for("evaluate_link", photo=photo_name, report=report_filename))

        checklist_path = _safe_checklist_path(checklist_filename)
        status = request.form.get("status", "").strip() or pe.suggest_status_from_report(report_text)
        cm.update_checklist_row(
            checklist_path, row_number,
            status=status, photo_ref=photo_name,
            notes=pe.notes_from_report(report_text),
            corrective_action=pe.corrective_action_from_report(report_text) or None,
        )
        flash(f"Updated {checklist_path.name}, row {row_number} (Status: {status}).", "success")
        return redirect(url_for("checklists_view", filename=checklist_path.name))

    checklists = []
    for path in cm.list_checklists():
        checklists.append({"filename": path.name, "rows": cm.read_checklist_items(path)})
    suggested_status = pe.suggest_status_from_report(report_text)
    return render_template(
        "evaluate_link.html", photo_name=photo_name, report_filename=report_filename,
        checklists=checklists, suggested_status=suggested_status,
    )


# --- Submittal checker -------------------------------------------------------------------------

@app.route("/submittals", methods=["GET", "POST"])
def submittals():
    if request.method == "GET":
        return render_template("submittals.html")

    f = request.files.get("submittal")
    requirement = request.form.get("requirement", "").strip()
    if not f or not f.filename or not requirement:
        flash("A submittal file and the requirement text are both required.", "error")
        return redirect(url_for("submittals"))

    safe_name = f"{secrets.token_hex(4)}_{secure_filename(f.filename)}"
    saved_path = Path("submittals") / safe_name
    f.save(saved_path)

    try:
        report_text = sc.check_submittal(saved_path, requirement)
    except RuntimeError as e:
        flash(str(e), "error")
        return redirect(url_for("submittals"))

    report_filename = f"{saved_path.stem}_submittal_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    (Path("outputs") / report_filename).write_text(report_text)
    return render_template("submittal_result.html", report_text=report_text, filename=f.filename)


# --- Library management (upload the criteria spreadsheet / submittals onto the server) ---------

@app.route("/library", methods=["GET", "POST"])
def library():
    if request.method == "POST":
        criteria_file = request.files.get("criteria_xlsx")
        if criteria_file and criteria_file.filename:
            criteria_file.save(pe.CRITERIA_XLSX)
            flash("Uploaded QA/QC Agent Training spreadsheet.", "success")

        submittal_files = [f for f in request.files.getlist("submittal_files") if f and f.filename]
        for f in submittal_files:
            f.save(Path("submittals") / secure_filename(f.filename))
        if submittal_files:
            flash(f"Uploaded {len(submittal_files)} submittal file(s).", "success")

        return redirect(url_for("library"))

    diag = pe.diagnose_criteria_setup()
    criteria_count = len(pe.load_criteria()) if not diag else 0
    submittal_files = sorted(p.name for p in Path("submittals").glob("*") if p.is_file())
    return render_template(
        "library.html", diag=diag, criteria_count=criteria_count, submittal_files=submittal_files,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
