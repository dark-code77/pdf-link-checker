# 🔗 PDF Link Checker

A Python tool that scans any PDF and validates all email and web hyperlinks.
Includes a web UI to run it from your browser — no coding needed to use it.

## What it checks
- Missing hyperlinks (text visible but not clickable)
- Wrong link type (email linked to a web URL, or vice versa)
- Broken addresses split across lines
- Partial linking (only half of a split address is linked)
- Wrong link targets (displayed address doesn't match the actual href)

## Output
Generates a colour-coded Excel report with columns:
`PDF_Link | Hyperlink_Text | Page_Numb | Link_Type | Is_Hyperlinked | Is_Valid | Result`

## How to run

### Install dependencies
pip install flask pypdf pdfplumber openpyxl pyngrok qrcode

### Run the web UI
python app.py             # local only
python app.py --lan       # share on same WiFi
python app.py --public    # share public link (anyone, anywhere)

Then open http://localhost:5050 in your browser.

## Tech used
- Python, Flask, pypdf, pdfplumber, openpyxl
