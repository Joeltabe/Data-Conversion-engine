# TranscriptX — Conversion System

Convert academic transcript PDFs to Excel with a full web UI.

## Quick Setup

```bash
# 1. Install Python dependencies
pip install flask pdfplumber openpyxl

# 2. Run the server
python app.py

# 3. Open your browser
http://localhost:5000
```

## How to Use

1. Set your **School ID** (default: 10)
2. Drop one or more PDF transcript files onto the upload zone
3. Click **Start Conversion** — engine parses all pages automatically
4. Switch to **Review & Edit** to inspect every record
5. Click any student row to open the edit panel and fix marks inline
6. Click **Generate Excel** — file downloads immediately

## Output Format

The Excel file contains three sheets:
- **Results** — Flat, import-ready table: Matricule · Subject Code · Schoolid · Test · Exam · Db1/Acc (academic year) · Semester. Matricule repeated per course row, no merged/title rows.
- **Summary** — One row per student with course counts and validation status
- **Legend** — Colour key and column guide

## CLI Usage (without the web UI)

```bash
# Single PDF
python pdf_to_excel_engine.py results.pdf

# Custom output
python pdf_to_excel_engine.py results.pdf -o Output.xlsx --school-id 10

# Merge multiple PDFs into one Excel
python pdf_to_excel_engine.py acc.pdf fin.pdf mgt.pdf --batch -o merged.xlsx

# Split a combined results PDF into one PDF per student
python pdf_to_excel_engine.py results.pdf --split -o Output.xlsx

# Split to a custom folder (default: <pdf>_split next to the source)
python pdf_to_excel_engine.py results.pdf --split --split-dir "F:\out"

# Skip confirmation prompt
python pdf_to_excel_engine.py results.pdf --no-confirm
```

## Split Combined Results PDFs

The engine recognises the multi-student **"STUDENT RESULTS"** sheets that the
exam office publishes (one student per page, with `NAME AND SURNAME`,
`MATRICULE`, and `First Semester Results For 2025/2026`).

- **CLI:** pass `--split` to also export one PDF per student alongside the Excel.
- **Web UI:** after uploading such a PDF, click **✂ Split PDFs** in the Review
  bar — the app produces a ZIP of `MATRICULE_NAME.pdf` files for download.
- Split files re-parse cleanly as single-student transcripts, so they can be
  uploaded individually or archived per student.

## Project Structure

```
TranscriptSystem/
├── app.py                  ← Flask web server
├── pdf_to_excel_engine.py  ← Core conversion engine (also works standalone)
├── requirements.txt        ← Python dependencies
├── templates/
│   └── index.html          ← Web UI (single-file, no build step)
├── uploads/                ← Temporary PDF storage
└── outputs/                ← Generated Excel files
```

## What the Engine Handles

- Any number of students per PDF
- Any number of courses per semester (varies per student)
- 1, 2, 3 or 4 semesters detected automatically
- Any subject codes, level, specialty, department
- Batch mode — multiple PDFs merged into one Excel
- Fallback text parser if table extraction fails
- Inline mark editing with live total recalculation
- Full validation: CA 0–30, Exam 0–70, total cross-check
