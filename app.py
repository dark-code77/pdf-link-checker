"""
PDF Link Checker — Web App
===========================
A local web UI to run pdf_link_checker.py through your browser.

Usage:
    pip install flask pypdf pdfplumber openpyxl
    python app.py

Then open:  http://localhost:5050
"""

import re, os, sys, io, json, threading, webbrowser, tempfile
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, List, Dict

from flask import Flask, request, jsonify, send_file, render_template_string

# ── Try importing checker deps ────────────────────────────────────────────────
try:
    import pdfplumber
    from pypdf import PdfReader
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    DEPS_OK = True
except ImportError as e:
    DEPS_OK = False
    MISSING_DEP = str(e)

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()

# ─────────────────────────────────────────────────────────────────────────────
# Core logic (copied from pdf_link_checker.py)
# ─────────────────────────────────────────────────────────────────────────────

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
URL_REGEX   = re.compile(r"(?:https?://|www\.)[a-zA-Z0-9\-._~:/?#\[\]@!$&'()*+,;=%]+")

@dataclass
class LinkAnnotation:
    page_num: int
    uri: str
    rect: tuple

@dataclass
class TextToken:
    page_num: int
    text: str
    x0: float; y0: float; x1: float; y1: float

@dataclass
class Row:
    pdf_link: str
    hyperlink_text: str
    page_num: int
    link_type: str
    is_hyperlinked: str
    is_valid: str
    hyperlink_points_to: str
    result: str

def extract_hyperlinks(pdf_path):
    links = []
    reader = PdfReader(pdf_path)
    for page_idx, page in enumerate(reader.pages):
        if "/Annots" not in page: continue
        annots = page["/Annots"]
        if not annots: continue
        ph = float(page.mediabox.height)
        for ref in annots:
            try:
                a = ref.get_object() if hasattr(ref,"get_object") else ref
                if a.get("/Subtype") != "/Link": continue
                act = a.get("/A")
                if not act or act.get("/S") != "/URI": continue
                uri = str(act.get("/URI",""))
                r = a.get("/Rect")
                if not r: continue
                rx0,ry0,rx1,ry1 = [float(v) for v in r]
                x0,x1 = sorted([rx0,rx1])
                y0 = ph - max(ry0,ry1); y1 = ph - min(ry0,ry1)
                links.append(LinkAnnotation(page_idx, uri, (x0,y0,x1,y1)))
            except: continue
    return links

def extract_tokens(pdf_path):
    tokens = []
    with pdfplumber.open(pdf_path) as pdf:
        for pi, page in enumerate(pdf.pages):
            for w in page.extract_words(x_tolerance=3,y_tolerance=3,keep_blank_chars=False,use_text_flow=False):
                tokens.append(TextToken(pi,w["text"],w["x0"],w["top"],w["x1"],w["bottom"]))
    return tokens

def rect_overlap(a, b, thr=0.3):
    ax0,ay0,ax1,ay1=a; bx0,by0,bx1,by1=b
    ix0,iy0=max(ax0,bx0),max(ay0,by0); ix1,iy1=min(ax1,bx1),min(ay1,by1)
    if ix1<=ix0 or iy1<=iy0: return False
    inter=(ix1-ix0)*(iy1-iy0)
    s=min((ax1-ax0)*(ay1-ay0),(bx1-bx0)*(by1-by0))
    return s>0 and (inter/s)>=thr

def find_link(tok, page_map):
    for ann in page_map.get(tok.page_num,[]):
        if rect_overlap((tok.x0,tok.y0,tok.x1,tok.y1), ann.rect): return ann
    return None

def is_mailto(u): return u.lower().startswith("mailto:")
def is_http(u): return u.lower().startswith(("http://","https://","www."))

def addr_match(displayed, uri):
    d = displayed.strip().rstrip(".,;:!?")
    if is_mailto(uri):
        return d.lower() == uri[7:].split("?")[0].strip().lower()
    if is_http(uri):
        def n(u):
            u=u.lower().rstrip("/")
            for p in ("http://","https://","www."):
                if u.startswith(p): u=u[len(p):]
            return u
        return n(d)==n(uri)
    return False

def group_lines(tokens):
    if not tokens: return []
    lines, cur = [], [tokens[0]]
    for t in tokens[1:]:
        if t.page_num==cur[-1].page_num and abs(t.y0-cur[-1].y0)<4: cur.append(t)
        else: lines.append(cur); cur=[t]
    lines.append(cur); return lines

def scan_line(line):
    full=" ".join(t.text for t in line); results=[]
    for pat in (EMAIL_REGEX, URL_REGEX):
        for m in pat.finditer(full):
            addr=m.group().rstrip(".,;:!?)\"'"); pos=0; cov=[]
            for tok in line:
                s=full.find(tok.text,pos); e=s+len(tok.text)
                if s<m.end() and e>m.start(): cov.append(tok)
                pos=s+1
            if cov: results.append((addr,cov))
    return results

def validate_pdf(pdf_path):
    links     = extract_hyperlinks(pdf_path)
    tokens    = extract_tokens(pdf_path)
    page_map  = {}
    for l in links: page_map.setdefault(l.page_num,[]).append(l)
    lines     = group_lines(tokens)
    rows: List[Row] = []

    for line in lines:
        for addr, toks in scan_line(line):
            page = toks[0].page_num + 1
            is_email = bool(EMAIL_REGEX.fullmatch(addr.strip()))
            link_type = "Email" if is_email else "Web Link"

            found = [find_link(t,page_map) for t in toks]
            found = [f for f in found if f]
            uris  = list({f.uri for f in found})

            is_hyperlinked = "Yes" if uris else "No"
            hyperlink_text = uris[0] if uris else "No Link"
            is_valid = "Yes"; result = "Pass"

            if not uris:
                is_valid = "No"; result = "Fail"
            else:
                for uri in uris:
                    if is_email and is_http(uri):   is_valid="No"; result="Fail"
                    elif not is_email and is_mailto(uri): is_valid="No"; result="Fail"
                    elif not addr_match(addr, uri): is_valid="No"; result="Fail"

            rows.append(Row(addr, hyperlink_text, page, link_type,
                            is_hyperlinked, is_valid, hyperlink_text, result))
    return rows

def export_excel(rows, path):
    wb = Workbook(); ws = wb.active; ws.title = "PDF Link Validation"
    headers = ["PDF_Link","Hyperlink_Text","Page_Numb","Link_Type",
               "Is_Hyperlinked","Is_Valid","Hyperlink_Points_To","Result"]
    ws.append(headers)
    hfill = PatternFill("solid",fgColor="1F3864")
    hfont = Font(bold=True,color="FFFFFF",size=11)
    for c in ws[1]:
        c.fill=hfill; c.font=hfont
        c.alignment=Alignment(horizontal="center",vertical="center")
    pf=PatternFill("solid",fgColor="C6EFCE"); ff=Font(color="276221",bold=True)
    nf=PatternFill("solid",fgColor="FFCCCC"); nff=Font(color="9C0006",bold=True)
    for row in rows:
        ws.append([row.pdf_link,row.hyperlink_text,row.page_num,row.link_type,
                   row.is_hyperlinked,row.is_valid,row.hyperlink_points_to,row.result])
        ri=ws.max_row; rc=ws[f"H{ri}"]
        rc.fill,rc.font = (pf,ff) if row.result=="Pass" else (nf,nff)
    for col,w in zip("ABCDEFGH",[30,40,12,12,14,11,40,10]):
        ws.column_dimensions[col].width=w
    for row in ws.iter_rows(min_row=2,max_row=ws.max_row):
        for idx,c in enumerate(row):
            c.alignment=Alignment(
                horizontal="center" if idx in [2,3,4,5,7] else "left",
                vertical="center", wrap_text=True)
    wb.save(path)

# ─────────────────────────────────────────────────────────────────────────────
# HTML Template
# ─────────────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>PDF Link Checker</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

:root{
  --bg:#0d0f14;
  --surface:#13161e;
  --card:#191d28;
  --border:#252836;
  --accent:#4f8ef7;
  --accent2:#f7c948;
  --pass:#22c55e;
  --fail:#ef4444;
  --text:#e8eaf0;
  --muted:#6b7280;
  --mono:'JetBrains Mono',monospace;
  --sans:'Syne',sans-serif;
}

body{background:var(--bg);color:var(--text);font-family:var(--sans);min-height:100vh;
  background-image:radial-gradient(ellipse 80% 60% at 50% -10%,rgba(79,142,247,.12),transparent);
}

/* ── Layout ── */
.wrap{max-width:1200px;margin:0 auto;padding:32px 24px}

/* ── Header ── */
header{display:flex;align-items:center;gap:16px;margin-bottom:48px}
.logo{width:44px;height:44px;background:var(--accent);border-radius:12px;
  display:grid;place-items:center;font-size:22px;flex-shrink:0}
h1{font-size:clamp(22px,3vw,30px);font-weight:800;letter-spacing:-.5px}
h1 span{color:var(--accent)}
.badge{margin-left:auto;font-family:var(--mono);font-size:11px;
  background:var(--card);border:1px solid var(--border);border-radius:6px;
  padding:4px 10px;color:var(--muted)}

/* ── Drop zone ── */
.drop-zone{
  border:2px dashed var(--border);border-radius:20px;padding:56px 32px;
  text-align:center;cursor:pointer;transition:.25s;position:relative;overflow:hidden;
  background:var(--surface);
}
.drop-zone::before{
  content:'';position:absolute;inset:0;
  background:radial-gradient(circle at 50% 0%,rgba(79,142,247,.06),transparent 70%);
  pointer-events:none;
}
.drop-zone:hover,.drop-zone.drag{border-color:var(--accent);
  background:rgba(79,142,247,.04)}
.drop-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.drop-icon{font-size:52px;margin-bottom:16px;display:block}
.drop-zone h2{font-size:20px;font-weight:700;margin-bottom:8px}
.drop-zone p{color:var(--muted);font-size:14px}
.file-name{margin-top:14px;font-family:var(--mono);font-size:13px;color:var(--accent);
  background:rgba(79,142,247,.1);border-radius:8px;padding:8px 14px;display:none}

/* ── Controls row ── */
.controls{display:flex;gap:12px;margin-top:20px;flex-wrap:wrap;align-items:center}
.ctrl-label{font-size:13px;color:var(--muted);font-weight:600;letter-spacing:.5px;text-transform:uppercase}

label.toggle{display:flex;align-items:center;gap:8px;cursor:pointer;
  font-size:14px;color:var(--text);user-select:none}
input[type=checkbox]{appearance:none;width:38px;height:20px;background:var(--border);
  border-radius:10px;position:relative;cursor:pointer;transition:.2s;flex-shrink:0}
input[type=checkbox]::after{content:'';position:absolute;width:14px;height:14px;
  background:#fff;border-radius:50%;top:3px;left:3px;transition:.2s}
input[type=checkbox]:checked{background:var(--accent)}
input[type=checkbox]:checked::after{left:21px}

/* ── Button ── */
.btn{display:inline-flex;align-items:center;gap:8px;padding:12px 28px;border-radius:12px;
  border:none;cursor:pointer;font-family:var(--sans);font-weight:700;font-size:15px;
  transition:.2s;letter-spacing:.2px}
.btn-primary{background:var(--accent);color:#fff}
.btn-primary:hover{background:#3a7de8;transform:translateY(-1px)}
.btn-primary:disabled{opacity:.45;cursor:not-allowed;transform:none}
.btn-outline{background:transparent;border:2px solid var(--border);color:var(--text)}
.btn-outline:hover{border-color:var(--accent);color:var(--accent)}
.btn-dl{background:var(--accent2);color:#0d0f14}
.btn-dl:hover{background:#e6b83a;transform:translateY(-1px)}

/* ── Progress ── */
.progress-wrap{display:none;margin-top:28px}
.progress-bar-bg{height:6px;background:var(--border);border-radius:3px;overflow:hidden}
.progress-bar{height:100%;background:linear-gradient(90deg,var(--accent),#7eb4fc);
  border-radius:3px;width:0%;transition:.3s}
.progress-msg{font-family:var(--mono);font-size:12px;color:var(--muted);margin-top:8px}

/* ── Stats cards ── */
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:16px;
  margin:32px 0}
.stat-card{background:var(--card);border:1px solid var(--border);border-radius:16px;
  padding:20px 24px;position:relative;overflow:hidden}
.stat-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px}
.stat-card.total::after{background:var(--accent)}
.stat-card.pass::after{background:var(--pass)}
.stat-card.fail::after{background:var(--fail)}
.stat-card.rate::after{background:var(--accent2)}
.stat-card .label{font-size:11px;font-weight:700;letter-spacing:.8px;
  text-transform:uppercase;color:var(--muted);margin-bottom:8px}
.stat-card .value{font-size:36px;font-weight:800;line-height:1}
.stat-card.pass .value{color:var(--pass)}
.stat-card.fail .value{color:var(--fail)}
.stat-card.rate .value{color:var(--accent2)}

/* ── Filter bar ── */
.filter-bar{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.filter-bar span{font-size:13px;color:var(--muted);font-weight:600}
.pill{padding:6px 14px;border-radius:20px;border:1px solid var(--border);
  background:var(--card);color:var(--muted);font-size:13px;cursor:pointer;
  transition:.15s;font-family:var(--sans);font-weight:600}
.pill:hover,.pill.active{background:var(--accent);border-color:var(--accent);color:#fff}
.pill.pass-pill.active{background:var(--pass);border-color:var(--pass)}
.pill.fail-pill.active{background:var(--fail);border-color:var(--fail)}
.search-box{margin-left:auto;padding:7px 14px;border-radius:10px;border:1px solid var(--border);
  background:var(--card);color:var(--text);font-family:var(--mono);font-size:13px;outline:none;
  min-width:200px;transition:.2s}
.search-box:focus{border-color:var(--accent)}

/* ── Table ── */
.table-wrap{background:var(--card);border:1px solid var(--border);border-radius:16px;overflow:hidden}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:#111420;padding:14px 16px;text-align:left;
  font-size:11px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:var(--muted);
  border-bottom:1px solid var(--border);white-space:nowrap}
tbody tr{border-bottom:1px solid rgba(37,40,54,.6);transition:.15s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:rgba(79,142,247,.04)}
td{padding:13px 16px;vertical-align:top;word-break:break-all}
td:first-child{font-family:var(--mono);font-size:12px;color:var(--text)}
.td-uri{font-family:var(--mono);font-size:11px;color:var(--muted);word-break:break-all}
.tag{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:6px;
  font-size:11px;font-weight:700;letter-spacing:.3px;white-space:nowrap}
.tag-email{background:rgba(147,51,234,.15);color:#c084fc}
.tag-web{background:rgba(79,142,247,.15);color:var(--accent)}
.tag-yes{background:rgba(34,197,94,.12);color:var(--pass)}
.tag-no{background:rgba(239,68,68,.12);color:var(--fail)}
.tag-pass{background:rgba(34,197,94,.15);color:var(--pass);font-size:12px;padding:4px 12px}
.tag-fail{background:rgba(239,68,68,.15);color:var(--fail);font-size:12px;padding:4px 12px}
.page-num{font-family:var(--mono);font-size:12px;color:var(--muted)}

/* ── Toolbar ── */
.toolbar{display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap}
.toolbar-title{font-size:17px;font-weight:700;flex:1}
.row-count{font-family:var(--mono);font-size:12px;color:var(--muted)}

/* ── Empty / Error ── */
.empty{padding:64px;text-align:center;color:var(--muted)}
.empty span{font-size:40px;display:block;margin-bottom:12px}
.error-box{background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.25);
  border-radius:12px;padding:16px 20px;color:#fca5a5;font-family:var(--mono);
  font-size:13px;margin-top:16px;display:none}

/* ── Pagination ── */
.pagination{display:flex;justify-content:flex-end;gap:6px;margin-top:16px;align-items:center}
.pg-btn{padding:6px 12px;border-radius:8px;border:1px solid var(--border);
  background:var(--card);color:var(--muted);cursor:pointer;font-size:13px;transition:.15s}
.pg-btn:hover,.pg-btn.active{background:var(--accent);border-color:var(--accent);color:#fff}
.pg-btn:disabled{opacity:.3;cursor:not-allowed}
.pg-info{font-family:var(--mono);font-size:12px;color:var(--muted)}

@media(max-width:640px){
  .stats{grid-template-columns:1fr 1fr}
  .filter-bar{flex-direction:column;align-items:flex-start}
  .search-box{margin-left:0;width:100%}
}
</style>
</head>
<body>
<div class="wrap">

  <!-- Header -->
  <header>
    <div class="logo">🔗</div>
    <div>
      <h1>PDF <span>Link</span> Checker</h1>
      <p style="color:var(--muted);font-size:13px">Validate all email & web addresses in your PDF</p>
    </div>
    <div class="badge">v2.0</div>
  </header>

  <!-- Upload -->
  <div class="drop-zone" id="dropZone">
    <input type="file" id="fileInput" accept=".pdf"/>
    <span class="drop-icon">📄</span>
    <h2>Drop your PDF here</h2>
    <p>or click anywhere in this box to browse</p>
    <div class="file-name" id="fileName"></div>
  </div>

  <!-- Controls -->
  <div class="controls">
    <span class="ctrl-label">Options</span>
    <label class="toggle">
      <input type="checkbox" id="verboseCheck"/>
      Verbose output
    </label>
    <div style="flex:1"></div>
    <button class="btn btn-primary" id="runBtn" disabled onclick="runCheck()">
      ▶ &nbsp;Run Check
    </button>
  </div>

  <!-- Progress -->
  <div class="progress-wrap" id="progressWrap">
    <div class="progress-bar-bg"><div class="progress-bar" id="progressBar"></div></div>
    <div class="progress-msg" id="progressMsg">Uploading…</div>
  </div>

  <!-- Error box -->
  <div class="error-box" id="errorBox"></div>

  <!-- Results (hidden until run) -->
  <div id="results" style="display:none">

    <!-- Stats -->
    <div class="stats" id="statsGrid"></div>

    <!-- Toolbar -->
    <div class="toolbar">
      <div class="toolbar-title">Results <span class="row-count" id="rowCount"></span></div>
      <button class="btn btn-dl" id="dlBtn" onclick="downloadExcel()">
        ⬇&nbsp; Download Excel
      </button>
    </div>

    <!-- Filter bar -->
    <div class="filter-bar">
      <span>Filter:</span>
      <button class="pill active" onclick="setFilter('all',this)">All</button>
      <button class="pill pass-pill" onclick="setFilter('Pass',this)">✅ Pass</button>
      <button class="pill fail-pill" onclick="setFilter('Fail',this)">❌ Fail</button>
      <input class="search-box" id="searchBox" placeholder="Search address…" oninput="applyFilters()"/>
    </div>

    <!-- Table -->
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Page</th>
            <th>PDF Link (Displayed)</th>
            <th>Hyperlink URI</th>
            <th>Link Type</th>
            <th>Has Link?</th>
            <th>Valid?</th>
            <th>Result</th>
          </tr>
        </thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>

    <!-- Pagination -->
    <div class="pagination" id="pagination"></div>
  </div>

</div>

<script>
const PAGE_SIZE = 50;
let allRows = [], filteredRows = [], currentPage = 1, activeFilter = 'all';
let lastReportPath = '';

// ── File input ──────────────────────────────────────────────────────────────
const dz = document.getElementById('dropZone');
const fi = document.getElementById('fileInput');

fi.addEventListener('change', () => handleFile(fi.files[0]));
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});

function handleFile(file) {
  if (!file || !file.name.endsWith('.pdf')) return;
  const fn = document.getElementById('fileName');
  fn.textContent = '📎  ' + file.name;
  fn.style.display = 'inline-block';
  document.getElementById('runBtn').disabled = false;
  document.getElementById('results').style.display = 'none';
  document.getElementById('errorBox').style.display = 'none';
}

// ── Run check ───────────────────────────────────────────────────────────────
async function runCheck() {
  const file = fi.files[0];
  if (!file) return;

  setProgress(true, 5, 'Uploading PDF…');
  document.getElementById('runBtn').disabled = true;
  document.getElementById('errorBox').style.display = 'none';
  document.getElementById('results').style.display = 'none';

  const fd = new FormData();
  fd.append('pdf', file);
  fd.append('verbose', document.getElementById('verboseCheck').checked);

  try {
    setProgress(true, 20, 'Extracting annotations…');
    const res = await fetch('/check', { method:'POST', body: fd });
    setProgress(true, 70, 'Validating links…');
    const data = await res.json();

    if (!res.ok || data.error) {
      showError(data.error || 'Server error');
      return;
    }

    setProgress(true, 95, 'Building report…');
    await new Promise(r => setTimeout(r, 300));
    setProgress(false);

    allRows = data.rows;
    lastReportPath = data.report_path;
    activeFilter = 'all';
    document.getElementById('searchBox').value = '';
    document.querySelectorAll('.pill').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.pill')[0].classList.add('active');

    renderStats(data);
    applyFilters();
    document.getElementById('results').style.display = 'block';

  } catch(e) {
    showError('Connection error: ' + e.message);
  } finally {
    document.getElementById('runBtn').disabled = false;
  }
}

// ── Stats ────────────────────────────────────────────────────────────────────
function renderStats(data) {
  const rate = data.total > 0 ? ((data.pass/data.total)*100).toFixed(1) : '0.0';
  document.getElementById('statsGrid').innerHTML = `
    <div class="stat-card total"><div class="label">Total Found</div><div class="value">${data.total}</div></div>
    <div class="stat-card pass"><div class="label">Passed</div><div class="value">${data.pass}</div></div>
    <div class="stat-card fail"><div class="label">Failed</div><div class="value">${data.fail}</div></div>
    <div class="stat-card rate"><div class="label">Pass Rate</div><div class="value">${rate}%</div></div>
  `;
}

// ── Filtering ────────────────────────────────────────────────────────────────
function setFilter(f, el) {
  activeFilter = f;
  document.querySelectorAll('.pill').forEach(p=>p.classList.remove('active'));
  el.classList.add('active');
  currentPage = 1;
  applyFilters();
}

function applyFilters() {
  const q = document.getElementById('searchBox').value.toLowerCase();
  filteredRows = allRows.filter(r => {
    const matchFilter = activeFilter==='all' || r.result===activeFilter;
    const matchSearch = !q ||
      r.pdf_link.toLowerCase().includes(q) ||
      r.hyperlink_points_to.toLowerCase().includes(q);
    return matchFilter && matchSearch;
  });
  currentPage = 1;
  renderTable();
}

// ── Table ────────────────────────────────────────────────────────────────────
function renderTable() {
  const start = (currentPage-1)*PAGE_SIZE;
  const pageRows = filteredRows.slice(start, start+PAGE_SIZE);
  const body = document.getElementById('tableBody');
  document.getElementById('rowCount').textContent =
    `(${filteredRows.length} of ${allRows.length})`;

  if (!pageRows.length) {
    body.innerHTML = `<tr><td colspan="7"><div class="empty">
      <span>🔍</span>No results match your filter.</div></td></tr>`;
    renderPagination();
    return;
  }

  body.innerHTML = pageRows.map(r => `
    <tr>
      <td class="page-num">${r.page_num}</td>
      <td style="font-family:var(--mono);font-size:12px">${esc(r.pdf_link)}</td>
      <td class="td-uri">${esc(r.hyperlink_points_to)}</td>
      <td><span class="tag ${r.link_type==='Email'?'tag-email':'tag-web'}">
        ${r.link_type==='Email'?'✉':'🌐'} ${r.link_type}</span></td>
      <td><span class="tag ${r.is_hyperlinked==='Yes'?'tag-yes':'tag-no'}">
        ${r.is_hyperlinked==='Yes'?'✓':'✗'} ${r.is_hyperlinked}</span></td>
      <td><span class="tag ${r.is_valid==='Yes'?'tag-yes':'tag-no'}">
        ${r.is_valid==='Yes'?'✓':'✗'} ${r.is_valid}</span></td>
      <td><span class="tag ${r.result==='Pass'?'tag-pass':'tag-fail'}">
        ${r.result==='Pass'?'✅ Pass':'❌ Fail'}</span></td>
    </tr>
  `).join('');

  renderPagination();
}

// ── Pagination ───────────────────────────────────────────────────────────────
function renderPagination() {
  const total = Math.ceil(filteredRows.length/PAGE_SIZE);
  const pg = document.getElementById('pagination');
  if (total<=1){pg.innerHTML='';return;}
  let html = `<span class="pg-info">Page ${currentPage} of ${total}</span>`;
  html+=`<button class="pg-btn" ${currentPage===1?'disabled':''} onclick="goPage(${currentPage-1})">‹</button>`;
  for(let i=1;i<=total;i++){
    if(i===1||i===total||Math.abs(i-currentPage)<=1){
      html+=`<button class="pg-btn ${i===currentPage?'active':''}" onclick="goPage(${i})">${i}</button>`;
    } else if(Math.abs(i-currentPage)===2){
      html+='<span class="pg-info">…</span>';
    }
  }
  html+=`<button class="pg-btn" ${currentPage===total?'disabled':''} onclick="goPage(${currentPage+1})">›</button>`;
  pg.innerHTML=html;
}

function goPage(p){currentPage=p;renderTable();window.scrollTo(0,0);}

// ── Download ─────────────────────────────────────────────────────────────────
async function downloadExcel() {
  if (!lastReportPath) return;
  const res = await fetch('/download?path='+encodeURIComponent(lastReportPath));
  if (!res.ok) { showError('Download failed'); return; }
  const blob = await res.blob();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = lastReportPath.split('/').pop() || 'report.xlsx';
  a.click();
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function setProgress(show, pct, msg) {
  document.getElementById('progressWrap').style.display = show?'block':'none';
  document.getElementById('progressBar').style.width = pct+'%';
  document.getElementById('progressMsg').textContent = msg||'';
}
function showError(msg) {
  const b = document.getElementById('errorBox');
  b.textContent = '❌ ' + msg; b.style.display = 'block';
  setProgress(false);
}
function esc(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>"""

# ─────────────────────────────────────────────────────────────────────────────
# Flask routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    if not DEPS_OK:
        return f"<h2>Missing dependencies</h2><pre>{MISSING_DEP}</pre><p>Run: pip install pypdf pdfplumber openpyxl flask</p>", 500
    return render_template_string(HTML)


@app.route("/check", methods=["POST"])
def check():
    if not DEPS_OK:
        return jsonify(error=f"Missing dependency: {MISSING_DEP}"), 500

    f = request.files.get("pdf")
    if not f or not f.filename.lower().endswith(".pdf"):
        return jsonify(error="Please upload a valid .pdf file"), 400

    tmp_pdf = os.path.join(UPLOAD_FOLDER, "upload_check.pdf")
    f.save(tmp_pdf)

    try:
        rows = validate_pdf(tmp_pdf)
    except Exception as e:
        return jsonify(error=f"PDF processing error: {str(e)}"), 500

    # Export Excel
    base = os.path.splitext(f.filename)[0]
    out_path = os.path.join(UPLOAD_FOLDER, f"{base}_link_report.xlsx")
    try:
        export_excel(rows, out_path)
    except Exception as e:
        return jsonify(error=f"Excel export error: {str(e)}"), 500

    total = len(rows)
    passed = sum(1 for r in rows if r.result == "Pass")

    return jsonify(
        total=total,
        passed=passed,
        **{"pass": passed, "fail": total - passed},
        report_path=out_path,
        rows=[{
            "page_num":          r.page_num,
            "pdf_link":          r.pdf_link,
            "hyperlink_text":    r.hyperlink_text,
            "link_type":         r.link_type,
            "is_hyperlinked":    r.is_hyperlinked,
            "is_valid":          r.is_valid,
            "hyperlink_points_to": r.hyperlink_points_to,
            "result":            r.result,
        } for r in rows]
    )


@app.route("/download")
def download():
    path = request.args.get("path","")
    if not path or not os.path.exists(path):
        return "File not found", 404
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────────────────────────────────────
# Launch
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    PORT = 5050
    url  = f"http://localhost:{PORT}"
    print(f"\n{'─'*52}")
    print(f"  🔗  PDF Link Checker  —  Web UI")
    print(f"{'─'*52}")
    print(f"  Open in browser:  {url}")
    print(f"  Press Ctrl+C to stop.")
    print(f"{'─'*52}\n")

    # Open browser after a short delay
    def _open():
        import time; time.sleep(1.2)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()

    app.run(port=PORT, debug=False)
