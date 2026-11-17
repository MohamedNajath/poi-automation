"""
Dahua POI Report Automation  -  single-file app (Streamlit)

Run:   pip install -r requirements.txt
       streamlit run app.py

Workflow: upload your master sheet (.xlsx preferred) -> upload up to 25 POI PDFs ->
Process -> review / edit -> Export.  The Verification ID is taken from each PDF FILE NAME.
The original master file is never modified; a new copy is produced.
All master-column matching is by header text (see FIELD_COLUMNS / config below), not by column letter.
"""
import csv, io, re, datetime, shutil
from copy import copy
from pathlib import Path

import fitz                      # PyMuPDF
import pdfplumber
from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

MAX_PDFS = 25

# ----------------------------------------------------------------------------- CONFIG
# (field, header regex, band regex).  Headers are lower-cased, whitespace-collapsed.
FIELD_COLUMNS = [
    ("verification_id", r"^verification id$", None),
    ("main",            r"^main$", None),
    ("project",         r"^project\.?$", None),
    ("location",        r"^location$", None),
    ("category",        r"^cat\w*gory$", None),
    ("report_date",     r"^date$", None),
    ("building_status", r"^building status$", None),
    ("remarks_si",      r"^remarks by si$", None),
    ("si_company",      r"^system int\w*$", None),
    ("si_name",         r"^contact person name\d*$", "system"),
    ("si_mobile",       r"^contact person mobile\w*\s*\w*\d*$", "system"),
    ("si_email",        r"^con\w*t email", "system"),
    ("client_name",     r"^contact person name\d*$", "client"),
    ("client_mobile",   r"^contact person mobile\w*\s*\w*\d*$", "client"),
    ("client_email",    r"^con\w*t email", "client"),
]
SCALAR_FIELDS = [f for f, _, _ in FIELD_COLUMNS]
# Category text (from PDF) -> value for the master "Main" column.  Edit freely.
CATEGORY_TO_MAIN = [("RESIDENTIAL", "Residential"), ("COMMERCIAL", "Commercial"), ("SHOWROOM", "Commercial"), ("RETAIL", "Commercial"), ("MALL", "Commercial"), ("OFFICE", "Commercial"), ("SCHOOL", "Educational"),
                    ("EDUCATION", "Educational"), ("BANK", "Financial"), ("FINANCIAL", "Financial")]
BUILDING_STATUS = {"NEW": "NEW", "AS-BUILD": "AS BUILD"}
BRACKET_INDOOR  = {"WALL": r"^wall mount", "CEILING": r"^ceiling mount", "PENDANT": r"^pendant"}
BRACKET_OUTDOOR = {"WALL": r"^outdoor housing \w+ wall", "CEILING": r"^outdoor housing ceiling", "PENDANT": r"^outdoor pendant"}
OUTDOOR_HOUSING = r"^outdoor dh-pfh"
# ------------------------------------------------------------------------------------

flat = lambda s: re.sub(r"\s+", " ", str(s or "")).strip()
def grab(p, t, flags=re.S | re.I):
    m = re.search(p, t, flags); return flat(m.group(1)) if m else ""
EMAIL = r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+"


# ============================================================================ PDF EXTRACTION
def id_from_filename(name):
    m = re.match(r"\s*([A-Za-z]\d{5,8})", Path(name).stem)
    return m.group(1).upper() if m else ""

def norm_model(s):
    s = re.split(r"[(?\n]", str(s or ""))[0].upper().replace(" ", "")
    return re.sub(r"^(DHI|DH)-", "", s).replace("-", "")

def parse_devices(txt):
    """NVR/HDD rows. Works whether each row is one text line or each cell is its own line."""
    m = re.search(r"NVR Details:?(.*?)(?:STORAGE CALC|SITE SURVEY REPORT|\b03:|$)", txt, re.S | re.I)
    sec = re.sub(r"^.*?Qty\b", "", flat(m.group(1)), count=1) if m else ""
    out = []
    for mm in re.finditer(r"(?<!\S)(\d{1,2})\s+([A-Za-z]{2,8})\s+(.+?)\s+(\d{1,3})(?=\s+\d{1,2}\s+[A-Za-z]{2,8}\s|\s*$)", sec):
        sn, dt, mid, q = mm.groups()
        model, desc = (mid, mid) if dt.upper() == "HDD" else (mid.split(" ", 1) + [""])[:2]
        out.append(dict(sn=sn, dtype=dt.upper(), model=model, desc=desc, qty=q, page=0))
    return out

def extract_pdf(data: bytes, filename: str) -> dict:
    w = []
    doc = fitz.open(stream=data, filetype="pdf")
    pages = [p.get_text() for p in doc]
    if sum(len(t.strip()) for t in pages) < 100:
        w.append("No text layer (scanned PDF). OCR is not enabled in this build - fill manually.")
    t1 = pages[0] if pages else ""
    alltxt = "\n".join(pages)
    rec = {"file": filename, "verification_id": id_from_filename(filename) or
           grab(r"ID\s*[-–]\s*([A-Z]\d{5,8})", alltxt)}
    if not rec["verification_id"]: w.append("Verification ID not found in file name")

    rec["project"]  = grab(r"Facility Name:\s*(.*?)\s*Report\s*#", t1)
    rec["report_no"] = grab(r"Report\s*#:\s*(.*?)\s*Location:", t1)
    rec["location"] = grab(r"Location:\s*(.*?)\s*Category:", t1)
    rec["category"] = grab(r"Category:\s*(.*?)\s*Type:", t1)
    boxes = re.findall(r"([\uf0a8\uf0f0\uf0fd\uf0fe☐☒☑✓✔])\s*(As-?\s*Build|New)", t1, re.I)
    checked = [re.sub(r"\s", "", n).upper() for c, n in boxes if c in "\uf0fd\uf0fe☒☑✓✔"]
    rec["building_status"] = BUILDING_STATUS.get(checked[0], "") if checked else ""
    if not rec["building_status"]: w.append("Project type (As-Build / New) not detected")
    cu = rec["category"].upper()
    rec["main"] = next((v for k, v in CATEGORY_TO_MAIN if k in cu), "")

    cb = grab(r"Client.s\s*Details:(.*?)Contractor:", t1)
    cm = re.search(r"Client.s\s*Details:(.*?)Contractor:", t1, re.S | re.I); cb = cm.group(1) if cm else ""
    rec["client_name"]   = grab(r"Name:\s*(.*?)\s*Mobile\s*#", cb)
    rec["client_mobile"] = grab(r"Mobile\s*#:\s*(.*?)\s*Designation", cb)
    rec["client_designation"] = grab(r"Designation:\s*(.*?)\s*Email:", cb)
    em = re.search(EMAIL, re.sub(r"\.\s+(\w)", r".\1", flat(cb.split("Email:")[-1]))); rec["client_email"] = em.group(0) if em else ""

    m = re.search(r"Contractor:\s*(.*?)\s*Contractor.s\s*Details:(.*?)(?:Design\s+Verified|Certified\s*Engineer)", t1, re.S | re.I)
    rec["si_company"] = flat(m.group(1)) if m else ""
    sb = m.group(2) if m else ""
    rec["si_name"]   = grab(r"Name:\s*(.*?)\s*Mobile\s*#", sb)
    rec["si_mobile"] = grab(r"Mobile\s*#:\s*(.*?)\s*Designation", sb)
    rec["contractor_designation"] = grab(r"Designation:\s*(.*?)\s*Email:", sb)
    em = re.search(EMAIL, re.sub(r"\.\s+(\w)", r".\1", flat(sb.split("Email:")[-1]))); rec["si_email"] = em.group(0) if em else ""
    rec["engineer"]   = grab(r"Certified\s*Engineer:\s*(.*?)\s*Certified\s*Technician:", t1)
    rec["technician"] = grab(r"Certified\s*Technician:\s*(.*?)\s*Installation and Configuration", t1)

    rec["report_date"] = ""
    head = t1[:400]
    dm = re.search(r"(\d{1,2})[\s\-/.]+([A-Za-z]{3,9})[\s\-/.]+(\d{4})", head)
    if dm:
        for fmt in ("%d %b %Y", "%d %B %Y"):
            try: rec["report_date"] = datetime.datetime.strptime(" ".join(dm.groups()), fmt).strftime("%d/%m/%Y"); break
            except ValueError: pass
    if not rec["report_date"]:
        dn = re.search(r"\b(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})\b", head)
        if dn:
            try: rec["report_date"] = datetime.date(int(dn.group(3)), int(dn.group(2)), int(dn.group(1))).strftime("%d/%m/%Y")
            except ValueError: pass
    if not rec["report_date"]: w.append("Report date not found")
    elif not (2020 <= int(rec["report_date"][-4:]) <= datetime.date.today().year + 1):
        w.append(f"Report date {rec['report_date']} looks unusual - please check the year")
    cm_ = re.search(r"(?:THE\s+)?SITE IS[^\n]*", alltxt, re.I)
    rec["remarks_si"] = flat(cm_.group(0)).upper() if cm_ else ""
    rec["vendor"] = grab(r"Camera Vendor\s*([A-Za-z]+)", alltxt)

    cams, devs = [], []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for pi, page in enumerate(pdf.pages, 1):
            for tb in page.extract_tables():
                rows = [[flat(c) for c in r if c is not None] for r in tb if r]
                rows = [r for r in rows if r]
                if not rows: continue
                blob = " ".join(" ".join(r) for r in rows).lower(); n = len(rows[0])
                has_data = any(r[0].isdigit() for r in rows)
                is_cam = "camera model" in blob or (n >= 9 and has_data)
                is_dev = not is_cam and ("device type" in blob or (n == 5 and has_data))
                for r in rows:
                    if not r[0].isdigit(): continue
                    if is_cam and len(r) >= 9:
                        cams.append(dict(sn=r[0], location=r[1], height=r[2], distance=r[3], target=r[4],
                                         camera=r[5], lens=r[6], itype=r[7], env=r[8], page=pi))
                    elif is_dev and len(r) >= 5:
                        devs.append(dict(sn=r[0], dtype=r[1], model=r[2], desc=r[3], qty=r[4], page=pi))
    devs = parse_devices(alltxt) or devs
    rec["cameras"], rec["devices"] = cams, devs
    if not cams: w.append("No camera table found")
    if not devs: w.append("No NVR/HDD table found")
    tot = grab(r"Total Cameras\s*(\d+)", alltxt)
    if tot and cams and int(tot) != len(cams): w.append(f"Total Cameras says {tot} but {len(cams)} rows were read")
    rec["warnings"] = w
    return rec


# ============================================================================ MASTER TEMPLATE
def csv_to_workbook(data: bytes) -> Workbook:
    rows = list(csv.reader(io.StringIO(data.decode("utf-8-sig", errors="replace"))))
    wb = Workbook(); ws = wb.active; ws.title = "Master Sheet"
    hr = next((i for i, r in enumerate(rows[:20]) if any(flat(c).lower() == "verification id" for c in r)), 3)
    cam = next((j for j, c in enumerate(rows[hr - 2]) if flat(c).upper() == "CAMERA"), 26) if hr >= 2 else 26
    for r in rows[:hr + 1]: ws.append([c if c != "" else None for c in r])
    for r in rows[hr + 1:]:
        out = []
        for j, c in enumerate(r):
            c = c.strip()
            if c.isdigit() and (j == 0 or (j >= cam and flat(rows[hr][j] if j < len(rows[hr]) else ""))): c = int(c)
            out.append(c if c != "" else None)
        ws.append(out)
    ncol = max(len(r) for r in rows)
    hf = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    starts = [j for j, c in enumerate(rows[hr - 2]) if c.strip()] if hr >= 2 else []
    for k, s in enumerate(starts):
        e = (starts[k + 1] if k + 1 < len(starts) else ncol) - 1
        if e > s: ws.merge_cells(start_row=hr - 1, start_column=s + 1, end_row=hr - 1, end_column=e + 1)
    for rr in range(2, hr + 2):
        for c in ws[rr]:
            c.fill = PatternFill("solid", fgColor="2F5597" if rr < hr + 1 else "1F3864"); c.font = hf
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[hr + 1].height = 60
    for j in range(1, ncol + 1): ws.column_dimensions[L(j)].width = 14
    for k, wd in {"F": 40, "T": 30, "W": 28}.items(): ws.column_dimensions[k].width = wd
    ws.freeze_panes = ws.cell(hr + 2, 3)
    return wb

def load_template(data: bytes, name: str) -> Workbook:
    if name.lower().endswith(".csv"): return csv_to_workbook(data)
    return load_workbook(io.BytesIO(data))

def analyse(ws) -> dict:
    """Locate header row, field columns, model columns. Nothing hard-coded by letter."""
    hr = next((r for r in range(1, 25) for c in range(1, 8)
               if flat(ws.cell(r, c).value).lower() == "verification id"), None)
    if hr is None: raise ValueError("Could not find a 'Verification ID' header in the first 25 rows.")
    ncol = ws.max_column
    hdr = {c: flat(ws.cell(hr, c).value).lower() for c in range(1, ncol + 1)}
    raw = {c: str(ws.cell(hr, c).value or "") for c in range(1, ncol + 1)}
    band, cur = {}, ""
    for c in range(1, ncol + 1):
        v = flat(ws.cell(hr - 2, c).value).lower() if hr >= 3 else ""
        cur = v or cur; band[c] = cur
    cols = {}
    for f, hp, bp in FIELD_COLUMNS:
        for c in range(1, ncol + 1):
            if hdr[c] and re.search(hp, hdr[c]) and (bp is None or re.search(bp, band[c])):
                cols[f] = c; break
    cam = next((c for c in range(1, ncol + 1) if band[c] == "camera"), None)
    models, caps, brk_in, brk_out, housing = {}, {"surv": {}, "ent": {}}, {}, {}, None
    for c in range(cam or ncol + 1, ncol + 1):
        h = hdr[c]
        if not h: continue
        if re.match(r"^(\d+)\s*tb", h):
            cap = int(re.match(r"^(\d+)\s*tb", h).group(1))
            kind = "ent" if "enterprise" in band[c] else ("surv" if ("surv" in band[c] or band[c] == "status") else None)
            if kind: caps[kind].setdefault(cap, c)
            continue
        for k, p in BRACKET_INDOOR.items():
            if re.search(p, h): brk_in.setdefault(k, c)
        for k, p in BRACKET_OUTDOOR.items():
            if re.search(p, h): brk_out.setdefault(k, c)
        if re.search(OUTDOOR_HOUSING, h): housing = c
        models.setdefault(norm_model(raw[c]), c)
    label = {c: f"{L(c)} | {flat(raw[c])[:45]}" for c in range(1, ncol + 1)}
    equip = [c for c in range(cam or ncol + 1, ncol + 1) if hdr[c]]
    return dict(hr=hr, ncol=ncol, cols=cols, models=models, caps=caps, brk_in=brk_in, brk_out=brk_out,
                housing=housing, label=label, equip=equip, band=band)

def resolve_equipment(rec, info):
    out, w = {}, []
    add = lambda c, q=1: out.__setitem__(c, out.get(c, 0) + q)
    for cam in rec["cameras"]:
        for kind in ("camera", "lens"):
            n = norm_model(cam[kind])
            if not n: w.append(f"Camera {cam['sn']}: {kind} model missing")
            elif n in info["models"]: add(info["models"][n])
            else: w.append(f"{kind} '{cam[kind]}' not found in master columns")
        t = cam["itype"].upper()
        key = ("WALL" if "WALL" in t else "CEILING" if "CEIL" in t else "PENDANT" if "PEND" in t
               else "POLE" if "POLE" in t else None)
        outdoor = "OUT" in cam["env"].upper()
        if outdoor:
            if info["housing"]: add(info["housing"])
            else: w.append("Outdoor housing column not found")
        if key == "POLE":
            w.append("Pole mount: the master has no pole bracket column - only the outdoor housing is counted. Add a bracket manually if needed.")
        elif not key: w.append(f"Camera {cam['sn']}: installation type unclear ('{cam['itype']}')")
        else:
            c = (info["brk_out"] if outdoor else info["brk_in"]).get(key)
            if c: add(c)
            else: w.append(f"No bracket column for {'outdoor ' if outdoor else ''}{key}")
    for d in rec["devices"]:
        q = int(d["qty"]) if str(d["qty"]).isdigit() else 1
        if re.search(r"HDD|HARD", d["dtype"].upper()):
            m = re.search(r"(\d+)\s*TB", (d["model"] + " " + d["desc"]).upper())
            txt = (d["dtype"] + d["model"] + d["desc"]).upper()
            kind = "ent" if "ENTERPRISE" in txt else "surv"
            if "SURV" not in txt and "ENTERPRISE" not in txt: w.append("HDD kind not stated - assumed surveillance")
            c = info["caps"][kind].get(int(m.group(1))) if m else None
            if c: add(c, q)
            else: w.append(f"HDD '{d['model']}' has no matching master column")
        else:
            c = info["models"].get(norm_model(d["model"]))
            if c: add(c, q)
            else: w.append(f"Device '{d['model']}' not found in master columns")
    return out, list(dict.fromkeys(w))


# ============================================================================ STATUS / EXPORT
def build_index(ws, info):
    idc = info["cols"].get("verification_id", 2); idx = {}
    for r in range(info["hr"] + 1, ws.max_row + 1):
        v = flat(ws.cell(r, idc).value).upper()
        if v and v != "AA": idx.setdefault(v, []).append(r)
    return idx

def is_placeholder(ws, r, info):
    idc = info["cols"].get("verification_id", 2)
    for c in range(1, info["ncol"] + 1):
        if c in (1, idc): continue
        v = ws.cell(r, c).value
        if v not in (None, "") and not (isinstance(v, str) and (v.startswith("=") or v.startswith("#"))): return False
    return True

def decide(ws, info, index, rec, seen):
    vid = rec["verification_id"].upper()
    if not vid: return "No ID", "Skip"
    if vid in seen: return "Duplicate in this batch", "Skip"
    rows = index.get(vid)
    if not rows: return "New - will append", "Append"
    if is_placeholder(ws, rows[0], info): return f"ID exists (row {rows[0]}, empty) - will fill", "Fill"
    return f"Duplicate - row {rows[0]} already has data", "Skip"

def last_data_row(ws, info):
    r = ws.max_row
    while r > info["hr"] and all(ws.cell(r, c).value in (None, "") for c in range(1, info["ncol"] + 1)): r -= 1
    return r

def export_workbook(tmpl_bytes, tmpl_name, recs, equip, actions, highlight=False):
    wb = load_template(tmpl_bytes, tmpl_name); ws = wb.worksheets[0]
    info = analyse(ws); index = build_index(ws, info); log = []
    Y = PatternFill("solid", fgColor="FFF2CC")
    for rec in recs:
        key = rec["file"]; act = actions.get(key, "Skip"); vid = rec["verification_id"].upper()
        if act == "Skip": log.append((key, vid, "Skipped", "")); continue
        if act in ("Fill", "Update") and index.get(vid): row = index[vid][0]
        else:
            last = last_data_row(ws, info); row = last + 1
            for c in range(1, info["ncol"] + 1):
                s_ = ws.cell(last, c); d = ws.cell(row, c)
                if last > info["hr"]:
                    d._style = copy(s_._style)
                    if isinstance(s_.value, str) and s_.value.startswith("="):
                        d.value = Translator(s_.value, origin=s_.coordinate).translate_formula(d.coordinate)
                else: d.font = Font(name="Arial", size=9)
            if last > info["hr"] and ws.row_dimensions[last].height: ws.row_dimensions[row].height = ws.row_dimensions[last].height
            snc = ws.cell(last, 1).value
            ws.cell(row, 1).value = snc + 1 if isinstance(snc, int) else (1 if last == info["hr"] else None)
            index.setdefault(vid, []).append(row)
        for f, c in info["cols"].items():
            v = rec.get(f)
            if v in (None, ""): continue
            cell = ws.cell(row, c)
            if f == "report_date":
                try: v = datetime.datetime.strptime(v, "%d/%m/%Y"); cell.number_format = "dd/mm/yyyy"
                except ValueError: pass
            cell.value = v
            if highlight: cell.fill = Y
        for c, q in equip.get(key, {}).items():
            if q: ws.cell(row, c).value = int(q); (highlight and setattr(ws.cell(row, c), "fill", Y))
        log.append((key, vid, "Updated" if act in ("Fill", "Update") else "Appended", row))
    out = io.BytesIO(); wb.save(out)
    return out.getvalue(), log

def processing_report(recs, log, equip, info):
    wb = Workbook(); s = wb.active; s.title = "Summary"
    s.append(["File", "Verification ID", "Result", "Excel row", "Warnings", "Mapped equipment"])
    lg = {l[0]: l for l in log}
    for r in recs:
        l = lg.get(r["file"], ("", "", "", ""))
        eq = "; ".join(f"{info['label'][c].split('| ')[1]} x{q}" for c, q in equip.get(r["file"], {}).items() if q)
        s.append([r["file"], r["verification_id"], l[2], l[3], " | ".join(r["warnings"]), eq])
    c = wb.create_sheet("Cameras")
    c.append(["Verification ID", "SN", "Location", "Height (m)", "Distance (m)", "Target (m)", "Camera", "Lens", "Type", "Environment"])
    for r in recs:
        for x in r["cameras"]:
            c.append([r["verification_id"], x["sn"], x["location"], x["height"], x["distance"], x["target"],
                      x["camera"], x["lens"], x["itype"], x["env"]])
    for w_ in wb.worksheets:
        for cell in w_[1]: cell.font = Font(bold=True)
        for col in "ABCDEFGHIJ": w_.column_dimensions[col].width = 26
    o = io.BytesIO(); wb.save(o); return o.getvalue()


# ============================================================================ UI
BASE = Path(__file__).parent
TEMPLATE = BASE / "master_template.xlsx"            # empty master layout (headers only)
STORE = BASE / "output" / "Dahua_POI_Master.xlsx"   # the running master file; one row is added per PDF

def ensure_store():
    STORE.parent.mkdir(exist_ok=True)
    if not STORE.exists(): shutil.copy(TEMPLATE, STORE)

def main():
    import pandas as pd, streamlit as st
    st.set_page_config(page_title="Dahua POI Report Automation", layout="wide")
    ss = st.session_state
    ss.setdefault("up_n", 0)
    ensure_store()
    tmpl = (STORE.read_bytes(), STORE.name)
    ws = load_template(*tmpl).worksheets[0]; info = analyse(ws); index = build_index(ws, info)
    n_rows = last_data_row(ws, info) - info["hr"]
    stamp = datetime.date.today().strftime("%Y_%m_%d")
    xl = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    st.title("Dahua POI Report Automation")
    with st.sidebar:
        st.header("Master Excel")
        st.metric("Projects in master", n_rows)
        st.download_button("Download Master Excel", tmpl[0], f"Dahua_POI_Master_{stamp}.xlsx", xl, use_container_width=True)
        with st.expander("Start a new empty master"):
            if st.checkbox("I understand this deletes all rows added so far"):
                if st.button("Reset master"):
                    shutil.copy(TEMPLATE, STORE); ss.pop("recs", None); ss.pop("last", None); st.rerun()

    if "last" in ss:
        log, rep = ss.last
        done = [l for l in log if l[2] != "Skipped"]
        st.success(f"Added to master: {', '.join(f'{l[1]} (row {l[3]})' for l in done) or 'nothing'}"
                   + (f". Skipped: {', '.join(l[1] or l[0] for l in log if l[2] == 'Skipped')}" if len(done) < len(log) else ""))
        st.download_button("Download Processing Report", rep, f"Processing_Report_{stamp}.xlsx", xl)

    st.subheader("1. Upload POI PDF report(s)")
    st.caption("Upload one PDF at a time or several (max 25). The Verification ID is read from the file name, "
               "and each PDF is added as one new row in the master Excel.")
    pdfs = st.file_uploader(f"POI PDFs", type=["pdf"], accept_multiple_files=True, key=f"pdfs_upload_{ss.up_n}")
    uniq = {}
    for f in pdfs or []: uniq.setdefault((f.name, f.size), f)
    files = list(uniq.values())
    if len(files) > MAX_PDFS: st.error(f"Only the first {MAX_PDFS} files will be processed."); files = files[:MAX_PDFS]

    if st.button("Process", type="primary", disabled=not files):
        ss.pop("last", None)
        recs, equip, seen = [], {}, set()
        bar = st.progress(0.0)
        for i, f in enumerate(files):
            try: rec = extract_pdf(f.getvalue(), f.name)
            except Exception as e:
                rec = {"file": f.name, "verification_id": id_from_filename(f.name), "cameras": [], "devices": [],
                       "warnings": [f"Extraction failed: {e}"]}
            eq, w = resolve_equipment(rec, info) if rec["cameras"] or rec["devices"] else ({}, [])
            rec["warnings"] += w
            rec["status"], rec["action"] = decide(ws, info, index, rec, seen)
            if rec["verification_id"]: seen.add(rec["verification_id"].upper())
            equip[f.name] = eq; recs.append(rec); bar.progress((i + 1) / len(files))
        ss.recs, ss.equip = recs, equip

    if "recs" in ss:
        recs = ss.recs
        st.subheader("2. Review and edit")
        rows = [{"file": r["file"], "action": r["action"], "status": r["status"], **{f: r.get(f, "") for f in SCALAR_FIELDS},
                 "warnings": " | ".join(r["warnings"])} for r in recs]
        df = st.data_editor(pd.DataFrame(rows), hide_index=True, use_container_width=True, key="ed1",
            disabled=["file", "status", "warnings"],
            column_config={"action": st.column_config.SelectboxColumn(options=["Append", "Fill", "Update", "Skip"], required=True)})
        st.caption("Append = add a new row. Skip = leave this PDF out. A Verification ID that is already in the master "
                   "defaults to Skip; choose Update to overwrite that row's non-blank fields.")
        rev = {v: k for k, v in info["label"].items()}
        erows = [{"file": r["file"], "master column": info["label"][c], "qty": q}
                 for r in recs for c, q in ss.equip[r["file"]].items()]
        edf = st.data_editor(pd.DataFrame(erows, columns=["file", "master column", "qty"]), hide_index=True,
            num_rows="dynamic", use_container_width=True, key="ed2",
            column_config={"file": st.column_config.SelectboxColumn(options=[r["file"] for r in recs], required=True),
                           "master column": st.column_config.SelectboxColumn(options=[info["label"][c] for c in info["equip"]], required=True),
                           "qty": st.column_config.NumberColumn(min_value=0, step=1)})

        if st.button("Add to Master Excel", type="primary"):
            final = []
            for r, (_, d) in zip(recs, df.iterrows()):
                r2 = dict(r); r2.update({f: ("" if pd.isna(d[f]) else str(d[f]).strip()) for f in SCALAR_FIELDS}); final.append(r2)
            equip = {r["file"]: {} for r in recs}
            for _, e in edf.dropna().iterrows():
                c = rev[e["master column"]]; equip[e["file"]][c] = equip[e["file"]].get(c, 0) + int(e["qty"])
            actions = dict(zip(df["file"], df["action"]))
            data, log = export_workbook(*tmpl, final, equip, actions)
            STORE.write_bytes(data)
            ss.last = (log, processing_report(final, log, equip, info))
            ss.pop("recs", None); ss.up_n += 1
            st.rerun()

if __name__ == "__main__":
    main()
