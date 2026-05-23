"""
app.py
------
Flask web application for the Resume Screening System.

Features:
  - Login / logout (session-based, hardcoded credentials)
  - Upload job description (text or file) + multiple resumes (PDF/DOCX/TXT)
  - TF-IDF cosine similarity scoring
  - Ranked results with filter tabs (All / Selected / Review / Rejected)
  - Accuracy metrics panel
  - Export to Excel

Runs on http://localhost:5000
"""

from __future__ import annotations

import io
import os
import sys
from functools import wraps

from flask import (
    Flask, render_template, request, flash,
    redirect, url_for, session, send_file,
)

# Make Scripts/ importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "Scripts"))
from resume_scorer import score_resumes  # noqa: E402

app = Flask(__name__)
app.secret_key = "resume-screening-dev-key"

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
CREDENTIALS = {"admin": "admin123"}
SELECTED_THRESHOLD = 70.0
REVIEW_THRESHOLD   = 40.0


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------

def extract_text_from_pdf(file_bytes: bytes) -> str:
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
        return "\n".join(pages).strip()
    except Exception:
        return ""


def extract_text_from_docx(file_bytes: bytes) -> str:
    try:
        from docx import Document
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join(p.text for p in doc.paragraphs).strip()
    except Exception:
        return ""


def extract_text(filename: str, file_bytes: bytes) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_bytes)
    elif ext == ".docx":
        return extract_text_from_docx(file_bytes)
    elif ext == ".txt":
        return file_bytes.decode("utf-8", errors="replace").strip()
    return ""


def assign_status(score: float) -> str:
    if score >= SELECTED_THRESHOLD:
        return "Selected"
    elif score >= REVIEW_THRESHOLD:
        return "Review"
    return "Rejected"


def score_band(score: float) -> str:
    if score < 20:   return "0-20"
    elif score < 40: return "20-40"
    elif score < 60: return "40-60"
    elif score < 80: return "60-80"
    else:            return "80-100"


# ---------------------------------------------------------------------------
# Routes — Auth
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("logged_in"):
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if CREDENTIALS.get(username) == password:
            session["logged_in"] = True
            session["username"] = username
            return redirect(url_for("index"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Routes — Main app
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
@login_required
def index():
    return render_template("index.html", username=session.get("username"))


@app.route("/screen", methods=["POST"])
@login_required
def screen():
    # Job description
    jd_text = request.form.get("jd_text", "").strip()
    jd_file = request.files.get("jd_file")
    if jd_file and jd_file.filename:
        jd_text = extract_text(jd_file.filename, jd_file.read()).strip()
    if not jd_text:
        flash("Please provide a job description (text or file).", "error")
        return redirect(url_for("index"))

    # Resumes
    resume_files = request.files.getlist("resumes")
    if not resume_files or all(f.filename == "" for f in resume_files):
        flash("Please upload at least one resume.", "error")
        return redirect(url_for("index"))

    candidates, resume_texts, warnings = [], [], []
    for f in resume_files:
        if not f.filename:
            continue
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in {".pdf", ".docx", ".txt"}:
            warnings.append(f"Skipped '{f.filename}' — unsupported format.")
            continue
        text = extract_text(f.filename, f.read())
        if not text.strip():
            warnings.append(f"'{f.filename}' appears to be empty or image-only — scored 0.")
        candidates.append(f.filename)
        resume_texts.append(text)

    if not candidates:
        flash("No valid resume files found.", "error")
        return redirect(url_for("index"))

    # Score
    try:
        scores = score_resumes(jd_text, resume_texts)
    except ValueError as exc:
        flash(f"Scoring error: {exc}", "error")
        return redirect(url_for("index"))

    # Build results
    results = []
    for name, score in zip(candidates, scores):
        results.append({
            "candidate": name,
            "score": score,
            "status": assign_status(score),
        })
    results.sort(key=lambda r: r["score"], reverse=True)

    # Accuracy / distribution metrics
    bands = {"0-20": 0, "20-40": 0, "40-60": 0, "60-80": 0, "80-100": 0}
    for r in results:
        bands[score_band(r["score"])] += 1

    top = results[0] if results else None
    avg_score = round(sum(r["score"] for r in results) / len(results), 2) if results else 0.0

    return render_template(
        "results.html",
        results=results,
        total=len(results),
        selected=sum(1 for r in results if r["status"] == "Selected"),
        review=sum(1 for r in results if r["status"] == "Review"),
        rejected=sum(1 for r in results if r["status"] == "Rejected"),
        jd_preview=jd_text[:250] + ("…" if len(jd_text) > 250 else ""),
        warnings=warnings,
        bands=bands,
        top=top,
        avg_score=avg_score,
        username=session.get("username"),
    )


@app.route("/export", methods=["POST"])
@login_required
def export():
    """Download results as Excel."""
    import json
    results_json = request.form.get("results_json", "[]")
    results = json.loads(results_json)

    try:
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Screening Results"
        ws.append(["Rank", "Candidate", "Score (%)", "Status"])
        for i, r in enumerate(results, 1):
            ws.append([i, r["candidate"], r["score"], r["status"]])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="ResumeScreeningResults.xlsx",
        )
    except ImportError:
        # openpyxl not installed — return CSV instead
        lines = ["Rank,Candidate,Score (%),Status"]
        for i, r in enumerate(results, 1):
            lines.append(f'{i},"{r["candidate"]}",{r["score"]},{r["status"]}')
        buf = io.BytesIO("\n".join(lines).encode())
        return send_file(
            buf,
            mimetype="text/csv",
            as_attachment=True,
            download_name="ResumeScreeningResults.csv",
        )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
