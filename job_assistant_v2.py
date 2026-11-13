"""
Automotive GPT Job Assistant v2
- Searches Google Jobs via SerpApi in Germany, Sweden, Tunisia, France, Belgium
- Matches jobs against CV + LinkedIn export using OpenAI
- Generates tailored DOCX CV and cover letter drafts
- Stores jobs and application status in SQLite
- Requires explicit human approval before opening an application link
- Does NOT submit applications or automate LinkedIn interactions
"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import sqlite3
import sys
import time
import webbrowser
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv
from docx import Document
from docx.shared import Inches, Pt
from openai import OpenAI
from pypdf import PdfReader

load_dotenv()

ROOT = Path(__file__).resolve().parent
RESULTS = ROOT / "job_results"
RESULTS.mkdir(exist_ok=True)
DB_PATH = RESULTS / "applications.sqlite3"

COUNTRIES = ["Germany", "Sweden", "Tunisia", "France", "Belgium"]
JOB_TITLES = [
    "Senior Automotive Software Tester",
    "Automotive Test Automation Engineer",
    "ECU Validation Engineer",
    "Embedded Software Test Engineer",
    "Automotive Software Verification Engineer",
    "SDV Test Engineer",
    "HIL Test Engineer Automotive",
    "Vehicle Integration Test Engineer",
]
MINIMUM_SCORE = 70
MAX_RESULTS_PER_QUERY = 10

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
SERPAPI_KEY = os.getenv("SERPAPI_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-6-astra")
CV_FILE = os.getenv("CV_FILE", "my_cv.pdf")
LINKEDIN_FILE = os.getenv("LINKEDIN_FILE", "linkedin_profile.pdf")


def require_keys() -> None:
    missing = []
    if not OPENAI_API_KEY:
        missing.append("OPENAI_API_KEY")
    if not SERPAPI_KEY:
        missing.append("SERPAPI_KEY")
    if missing:
        raise SystemExit(
            "Missing API key(s) in .env: " + ", ".join(missing)
            + "\nCopy .env.example to .env and add your keys."
        )


def read_profile_file(filename: str) -> str:
    path = Path(filename)
    if not path.is_absolute():
        path = ROOT / path
    if not path.exists():
        raise FileNotFoundError(
            f"Profile file not found: {path}\n"
            "Add your CV and LinkedIn PDF/DOCX/TXT export, then update .env."
        )

    suffix = path.suffix.lower()
    if suffix == ".pdf":
        reader = PdfReader(str(path))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    elif suffix == ".docx":
        doc = Document(str(path))
        chunks = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                chunks.append(" | ".join(cell.text for cell in row.cells))
        text = "\n".join(chunks)
    elif suffix in {".txt", ".md"}:
        text = path.read_text(encoding="utf-8", errors="replace")
    else:
        raise ValueError(f"Unsupported profile format: {suffix}")

    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if len(text) < 100:
        raise ValueError(
            f"Very little text extracted from {path}. "
            "If this is a scanned PDF, run OCR first or provide DOCX/TXT."
        )
    return text[:50000]


def init_db() -> None:
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                job_id TEXT PRIMARY KEY,
                title TEXT,
                company TEXT,
                location TEXT,
                country TEXT,
                description TEXT,
                url TEXT,
                source TEXT,
                date_found TEXT,
                score INTEGER,
                recommendation TEXT,
                analysis_json TEXT,
                status TEXT DEFAULT 'Found',
                draft_cv TEXT,
                draft_letter TEXT,
                last_updated TEXT
            )
        """)
        con.commit()


def stable_job_id(job: dict[str, Any]) -> str:
    raw = "|".join([
        str(job.get("title", "")),
        str(job.get("company_name", "")),
        str(job.get("location", "")),
        str(job.get("share_link", "")),
    ]).lower()
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def search_google_jobs(title: str, country: str) -> list[dict[str, Any]]:
    params = {
        "engine": "google_jobs",
        "q": title,
        "location": country,
        "hl": "en",
        "api_key": SERPAPI_KEY,
    }
    response = requests.get(
        "https://serpapi.com/search.json",
        params=params,
        timeout=45,
    )
    response.raise_for_status()
    data = response.json()
    if data.get("error"):
        print(f"  Search provider message: {data['error']}")
        return []
    return data.get("jobs_results", [])[:MAX_RESULTS_PER_QUERY]


def find_job_url(job: dict[str, Any]) -> str:
    if job.get("share_link"):
        return job["share_link"]
    for option in job.get("apply_options", []) or []:
        if option.get("link"):
            return option["link"]
    for item in job.get("related_links", []) or []:
        if item.get("link"):
            return item["link"]
    return ""


def make_openai_client() -> OpenAI:
    return OpenAI(api_key=OPENAI_API_KEY)


def analyze_match(client: OpenAI, cv: str, linkedin: str,
                  job: dict[str, Any], country: str) -> dict[str, Any]:
    schema = {
        "type": "object",
        "properties": {
            "match_score": {"type": "integer"},
            "recommendation": {"type": "string"},
            "matching_evidence": {"type": "array", "items": {"type": "string"}},
            "gaps": {"type": "array", "items": {"type": "string"}},
            "risk_flags": {"type": "array", "items": {"type": "string"}},
            "reason": {"type": "string"},
            "tailored_cv_summary": {"type": "string"},
            "cover_letter": {"type": "string"},
            "application_answers": {
                "type": "object",
                "properties": {
                    "motivation": {"type": "string"},
                    "relevant_experience": {"type": "string"}
                },
                "required": ["motivation", "relevant_experience"],
                "additionalProperties": False
            },
            "keywords_to_emphasize": {"type": "array", "items": {"type": "string"}}
        },
        "required": [
            "match_score", "recommendation", "matching_evidence", "gaps",
            "risk_flags", "reason", "tailored_cv_summary", "cover_letter",
            "application_answers", "keywords_to_emphasize"
        ],
        "additionalProperties": False
    }

    payload = {
        "candidate_cv": cv,
        "linkedin_profile": linkedin,
        "job": {
            "title": job.get("title"),
            "company": job.get("company_name"),
            "location": job.get("location"),
            "country_search": country,
            "description": job.get("description", ""),
            "via": job.get("via", ""),
        },
    }

    instructions = """
You are a careful automotive and embedded-software technical recruiter.
Assess the real fit between the candidate and vacancy.
Score 0-100 based on required skills, automotive domain, test methods,
automation, ECU/embedded knowledge, seniority, and location relevance.
Only use facts supported by the supplied CV or LinkedIn profile.
Never invent employers, dates, tools, programming languages, certifications,
qualifications, achievements, metrics, or responsibilities.
If information is absent, put it in gaps or risk_flags.
A tailored summary and cover letter may rephrase true facts but must not
claim experience the candidate does not have. Do not claim work authorization,
language fluency, relocation willingness, or visa status unless stated.
Keep the letter concise, specific to the vacancy, and professional.
Return the requested JSON schema only.
"""

    response = client.responses.create(
        model=OPENAI_MODEL,
        instructions=instructions,
        input=json.dumps(payload, ensure_ascii=False),
        text={
            "format": {
                "type": "json_schema",
                "name": "job_match_analysis",
                "strict": True,
                "schema": schema,
            }
        },
    )
    result = json.loads(response.output_text)
    result["match_score"] = max(0, min(100, int(result["match_score"])))
    return result


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    doc.add_heading(text, level=level)


def save_cv_draft(candidate_cv: str, analysis: dict[str, Any],
                  job: dict[str, Any], job_id: str) -> Path:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)

    doc.add_heading("Tailored CV Draft – Review Required", 0)
    doc.add_paragraph(
        f"Target: {job.get('title', 'Role')} | "
        f"{job.get('company_name', 'Company')} | {job.get('location', '')}"
    )
    doc.add_paragraph(
        "This draft is an editing aid. Verify every statement against your "
        "original CV before sending."
    )
    add_heading(doc, "Professional Summary", 1)
    doc.add_paragraph(analysis["tailored_cv_summary"])
    add_heading(doc, "Existing CV Content – Source", 1)
    doc.add_paragraph(candidate_cv[:18000])
    add_heading(doc, "Keywords to Emphasize (only if accurate)", 1)
    for item in analysis["keywords_to_emphasize"]:
        doc.add_paragraph(item, style="List Bullet")

    path = RESULTS / f"{job_id}_tailored_cv.docx"
    doc.save(path)
    return path


def save_cover_letter(analysis: dict[str, Any],
                      job: dict[str, Any], job_id: str) -> Path:
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(10)
    doc.add_heading("Cover Letter – Review Required", 0)
    doc.add_paragraph(
        f"Application for {job.get('title', 'the position')} at "
        f"{job.get('company_name', 'the company')}"
    )
    doc.add_paragraph(analysis["cover_letter"])
    path = RESULTS / f"{job_id}_cover_letter.docx"
    doc.save(path)
    return path


def store_job(job: dict[str, Any], country: str,
              analysis: dict[str, Any], cv_path: Path,
              letter_path: Path) -> None:
    job_id = stable_job_id(job)
    now = datetime.now(timezone.utc).isoformat()
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
            INSERT INTO jobs (
                job_id, title, company, location, country, description,
                url, source, date_found, score, recommendation, analysis_json,
                status, draft_cv, draft_letter, last_updated
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Found', ?, ?, ?)
            ON CONFLICT(job_id) DO UPDATE SET
                score=excluded.score,
                recommendation=excluded.recommendation,
                analysis_json=excluded.analysis_json,
                draft_cv=excluded.draft_cv,
                draft_letter=excluded.draft_letter,
                last_updated=excluded.last_updated
        """, (
            job_id,
            job.get("title", ""),
            job.get("company_name", ""),
            job.get("location", ""),
            country,
            job.get("description", ""),
            find_job_url(job),
            job.get("via", "Google Jobs"),
            now,
            analysis["match_score"],
            analysis["recommendation"],
            json.dumps(analysis, ensure_ascii=False),
            str(cv_path),
            str(letter_path),
            now,
        ))
        con.commit()


def export_csv() -> Path:
    output = RESULTS / "job_matches.csv"
    with sqlite3.connect(DB_PATH) as con:
        rows = con.execute("""
            SELECT title, company, location, country, score, recommendation,
                   status, url, draft_cv, draft_letter, date_found
            FROM jobs
            ORDER BY score DESC, date_found DESC
        """).fetchall()

    headers = [
        "title", "company", "location", "country", "score",
        "recommendation", "status", "url", "draft_cv",
        "draft_letter", "date_found"
    ]
    with output.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    return output


def show_and_review_matches() -> None:
    with sqlite3.connect(DB_PATH) as con:
        con.row_factory = sqlite3.Row
        rows = con.execute("""
            SELECT * FROM jobs
            WHERE score >= ?
            ORDER BY score DESC, date_found DESC
        """, (MINIMUM_SCORE,)).fetchall()

    if not rows:
        print(f"\nNo jobs currently meet the {MINIMUM_SCORE}% threshold.")
        return

    print(f"\n{len(rows)} matching job(s) at or above {MINIMUM_SCORE}%:")
    for index, row in enumerate(rows, 1):
        print("\n" + "=" * 72)
        print(f"{index}. {row['title']} — {row['company']}")
        print(f"   Location: {row['location']} ({row['country']})")
        print(f"   GPT match: {row['score']}% | Status: {row['status']}")
        print(f"   Reason: {json.loads(row['analysis_json']).get('reason', '')}")
        print(f"   URL: {row['url'] or 'No application URL found'}")
        print(f"   CV draft: {row['draft_cv']}")
        print(f"   Letter: {row['draft_letter']}")

        if row["status"] in {"Applied", "Rejected", "Not interested"}:
            continue

        answer = input(
            "\nOpen this application page for your manual review? [y/N]: "
        ).strip().lower()

        if answer == "y":
            if row["url"]:
                webbrowser.open(row["url"])
            else:
                print("No verified application URL was found.")
                continue

            print(
                "Review the vacancy and submit manually on the employer's "
                "official application page."
            )
            status = input(
                "Update status [A=Applied, I=Interested, N=Not interested, "
                "Enter=leave unchanged]: "
            ).strip().lower()

            mapping = {
                "a": "Applied",
                "i": "Interested",
                "n": "Not interested",
            }
            if status in mapping:
                with sqlite3.connect(DB_PATH) as con:
                    con.execute(
                        "UPDATE jobs SET status=?, last_updated=? WHERE job_id=?",
                        (
                            mapping[status],
                            datetime.now(timezone.utc).isoformat(),
                            row["job_id"],
                        ),
                    )
                    con.commit()


def main() -> None:
    require_keys()
    init_db()

    print("Loading CV and LinkedIn profile...")
    cv = read_profile_file(CV_FILE)
    linkedin = read_profile_file(LINKEDIN_FILE)
    client = make_openai_client()

    total_seen = 0
    for country in COUNTRIES:
        for title in JOB_TITLES:
            print(f"\nSearching {country}: {title}")
            try:
                jobs = search_google_jobs(title, country)
            except requests.RequestException as exc:
                print(f"Search failed: {exc}")
                continue

            for job in jobs:
                total_seen += 1
                job_id = stable_job_id(job)
                with sqlite3.connect(DB_PATH) as con:
                    exists = con.execute(
                        "SELECT 1 FROM jobs WHERE job_id=?", (job_id,)
                    ).fetchone()
                if exists:
                    print("  Already tracked; skipping duplicate.")
                    continue

                try:
                    analysis = analyze_match(client, cv, linkedin, job, country)
                    score = analysis["match_score"]
                    print(
                        f"  {job.get('company_name', 'Unknown company')}: "
                        f"{score}%"
                    )

                    if score >= MINIMUM_SCORE:
                        cv_path = save_cv_draft(cv, analysis, job, job_id)
                        letter_path = save_cover_letter(analysis, job, job_id)
                        store_job(job, country, analysis, cv_path, letter_path)
                    else:
                        # Store lower-scoring matches too, without creating drafts.
                        store_job(
                            job, country, analysis,
                            Path("Not generated"), Path("Not generated")
                        )
                except Exception as exc:
                    print(f"  GPT/processing error: {exc}")

                time.sleep(0.5)

    csv_path = export_csv()
    print(f"\nSearch pass finished. Raw job results inspected: {total_seen}")
    print(f"CSV tracker: {csv_path}")
    print(f"SQLite database: {DB_PATH}")
    show_and_review_matches()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped by user.")
        sys.exit(130)
