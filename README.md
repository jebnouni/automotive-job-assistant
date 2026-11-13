# Automotive GPT Job Assistant v2

## What it does
- Searches Google Jobs via SerpApi for automotive testing roles in Germany, Sweden, Tunisia, France and Belgium.
- Evaluates each vacancy against your CV and LinkedIn export using the OpenAI API.
- Uses a 70% minimum threshold for generating tailored application drafts.
- Creates a tailored CV DOCX draft and cover-letter DOCX draft for qualifying roles.
- Tracks opportunities in SQLite and exports a CSV tracker.
- Asks before opening each application page. You review and submit manually.

## What it does NOT do
- It does not log in to LinkedIn, scrape private LinkedIn pages, send LinkedIn messages, or submit job applications.
- It does not guarantee that every vacancy is current or that every apply URL is an employer URL. Verify the employer, job status, location, eligibility and application page before applying.
- GPT match scores are estimates, not objective hiring probabilities.

## Setup
1. Install Python 3.10 or newer.
2. Extract this ZIP.
3. Copy `.env.example` to `.env`.
4. Add your OpenAI API key and SerpApi key to `.env`.
5. Put your CV in this folder as `my_cv.pdf` and your LinkedIn profile export as `linkedin_profile.pdf`, or edit the file names in `.env`.
6. Open a terminal in this folder and run:

   ```bash
   python -m venv .venv
   # Windows:
   .venv\Scripts\activate
   # macOS/Linux:
   source .venv/bin/activate
   pip install -r requirements.txt
   python job_assistant_v2.py
   ```

## LinkedIn export
On LinkedIn, use Settings & Privacy > Data privacy > Get a copy of your data. You can also save your public profile as PDF. This program reads a PDF/DOCX/TXT export; it does not need your LinkedIn password.

## Output
The `job_results` folder contains:
- `job_matches.csv` — all analyzed matches and statuses
- `applications.sqlite3` — local application tracking database
- `<job_id>_tailored_cv.docx` — CV editing draft
- `<job_id>_cover_letter.docx` — cover-letter draft

The tailored CV draft includes the original extracted CV text as source material plus a GPT-written summary. It is intentionally not a full automatic rewrite of every CV section: review and edit it to preserve your original formatting and ensure all claims are accurate.

## Configuration
Edit `JOB_TITLES`, `COUNTRIES`, `MINIMUM_SCORE` and `MAX_RESULTS_PER_QUERY` near the top of `job_assistant_v2.py`.

## Cost and privacy
SerpApi and OpenAI may charge for API usage. CV and LinkedIn text is sent to the OpenAI API for analysis. Do not use this with confidential employer/customer information unless you are authorized to share it. Keep `.env` private and never upload it to GitHub.
