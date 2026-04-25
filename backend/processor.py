"""
PDF & Excel processing utilities for the Result Asterisk Portal.

- Parses TC (Tabulation Chart) PDFs: multi-student-per-page tabular dump.
- Parses GS (Provisional Grade Sheet) PDFs: one student per page.
- Parses SEM_X excel sheets: detects highlighted (yellow / blue) cells per
  (roll_no, subject_code) which represent BACK / failed subjects.
- Parses TC_X / GS_X excel sheets to extract structured per-student records
  directly (so admin can upload only an Excel file and we still generate all
  semester TC*/GS* PDFs).
- Generates new TC*/GS* PDFs in the same visual format with " *" appended
  to subject names where a back exists. GS pages also embed a Code-128
  barcode (encoded with the university roll number).
"""
from __future__ import annotations

import io
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import pdfplumber
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)

ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
ROMAN_TO_INT = {r: i + 1 for i, r in enumerate(ROMAN)}

# ---------------------------------------------------------------------------
# Excel parsing — SEM_X sheet -> highlighted (back) subjects per student
# ---------------------------------------------------------------------------

# default openpyxl "no fill" tokens
_NO_FILL = {"00000000", "FFFFFFFF", "0", None, ""}


def _is_highlighted(cell) -> bool:
    f = getattr(cell, "fill", None)
    if not f or not getattr(f, "patternType", None):
        return False
    if f.patternType != "solid":
        return False
    fg = f.fgColor
    if not fg:
        return False
    rgb = fg.rgb if fg.type == "rgb" else None
    if rgb in _NO_FILL:
        return False
    return rgb is not None  # any non-default solid fill counts


def parse_sem_excel(file_bytes: bytes) -> Dict[str, Dict[str, Dict[str, bool]]]:
    """Return {sem_roman: {roll_no: {subject_code: True/False(=back)}}}."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    out: Dict[str, Dict[str, Dict[str, bool]]] = {}
    for sname in wb.sheetnames:
        m = re.match(r"^SEM[_ ]?(I{1,3}|IV|V|VI{1,3}|VIII)$", sname.strip(), re.I)
        if not m:
            continue
        sem = m.group(1).upper()
        ws = wb[sname]
        # Build subject_code -> list of column indices (1-based) it spans
        subj_cols: List[Tuple[str, List[int]]] = []
        last_code: Optional[str] = None
        for c in range(6, ws.max_column + 1):
            v = ws.cell(1, c).value
            if v and str(v).strip():
                code = str(v).strip()
                subj_cols.append((code, [c]))
                last_code = code
            else:
                if subj_cols and last_code:
                    subj_cols[-1][1].append(c)
        sem_map: Dict[str, Dict[str, bool]] = {}
        for r in range(6, ws.max_row + 1):
            roll = ws.cell(r, 2).value
            if not roll:
                continue
            roll = str(roll).strip()
            student_back: Dict[str, bool] = {}
            for code, cols in subj_cols:
                back = any(_is_highlighted(ws.cell(r, c)) for c in cols)
                if back:
                    student_back[code] = True
            if student_back:
                sem_map[roll] = student_back
        out[sem] = sem_map
    return out


# ---------------------------------------------------------------------------
# Excel parsing — TC_X and GS_X sheets -> per-student records
# ---------------------------------------------------------------------------

_SEM_RE = re.compile(r"^(TC|GS)[_ ]?(I{1,3}|IV|V|VI{1,3}|VIII)$", re.I)


def _norm(v) -> str:
    return "" if v is None else str(v).strip()


def _is_grade(g: str) -> bool:
    return bool(g) and bool(re.match(r"^([A-Z][+\-]?|Excellent|Ab|F|P|D|E)$", g))


_HEADER_BLOCK_RE = re.compile(
    r"(?P<program>(Bachelor|Master).+?)\n+(?P<branch>.+?)\n+(?P<sem>[IVX]+)\s+Semester\s+(?P<session>[A-Za-z]+\s+\d{4})",
    re.S,
)


def _extract_sheet_meta(ws) -> Dict[str, str]:
    """Extract program / branch / semester / session from the top header rows
    of a TC_X or GS_X sheet (the first cell that mentions 'Bachelor' or
    'Master')."""
    meta: Dict[str, str] = {"program": "", "branch": "", "semester": "",
                              "exam_session": ""}
    blob_parts: List[str] = []
    for r in range(1, 8):
        for c in range(1, 8):
            v = ws.cell(r, c).value
            if v:
                blob_parts.append(str(v))
    blob = "\n".join(blob_parts)
    blob = re.sub(r"Tabulation Chart for\s*", "", blob)
    blob = re.sub(r"GRADE SHEET\s*", "", blob)
    m = _HEADER_BLOCK_RE.search(blob)
    if m:
        meta["program"] = m.group("program").strip()
        meta["branch"] = re.sub(r"\s{2,}", " ", m.group("branch")).strip()
        meta["semester"] = m.group("sem").strip()
        meta["exam_session"] = m.group("session").strip()
    return meta


def parse_tc_excel_sheet(ws) -> List[dict]:
    """Parse a TC_X sheet -> list of student records.

    Block boundaries: rows starting with 'MM' in column A (case-insensitive).
    """
    records: List[dict] = []
    sheet_meta = _extract_sheet_meta(ws)
    rows = ws.max_row
    block_starts: List[int] = []
    for r in range(1, rows + 1):
        v = _norm(ws.cell(r, 1).value)
        if v.upper() == "MM":
            if _norm(ws.cell(r, 3).value).upper() == "SC":
                block_starts.append(r)
    block_starts.append(rows + 1)

    for i in range(len(block_starts) - 1):
        r0, r1 = block_starts[i], block_starts[i + 1]
        rec: dict = {"subjects": [], **sheet_meta}
        hdr_row = r0 + 1
        if hdr_row >= r1:
            continue
        rec["name"] = _norm(ws.cell(hdr_row, 2).value)
        rec["father_name"] = _norm(ws.cell(hdr_row, 5).value)
        rec["roll_no"] = _norm(ws.cell(hdr_row, 9).value)
        rec["enroll_no"] = _norm(ws.cell(hdr_row, 13).value)
        if not rec["roll_no"]:
            continue
        for rr in range(hdr_row + 2, r1):
            code = _norm(ws.cell(rr, 1).value)
            low = code.lower()
            # ---- Result / summary row ----
            if low.startswith("result"):
                rec["result"] = _norm(ws.cell(rr, 2).value)
                rec["remark"] = _norm(ws.cell(rr, 5).value)
                rec["sgpa"] = _norm(ws.cell(rr, 8).value)
                rec["cgpa"] = _norm(ws.cell(rr, 10).value)
                rec["earned_credits"] = _norm(ws.cell(rr, 12).value)
                rec["cuml_earned_credits"] = _norm(ws.cell(rr, 14).value)
                continue
            if not code or low.startswith("subject") or low.startswith("total"):
                continue
            if not re.match(r"^[A-Z]{2,4}\s?\d{2,4}$", code):
                continue
            name = _norm(ws.cell(rr, 2).value)
            credits = _norm(ws.cell(rr, 6).value)
            ext = _norm(ws.cell(rr, 7).value)
            ses = _norm(ws.cell(rr, 9).value)
            tot = _norm(ws.cell(rr, 11).value)
            grade = _norm(ws.cell(rr, 13).value)
            gp = _norm(ws.cell(rr, 14).value)
            rec["subjects"].append({
                "code": code,
                "name": name,
                "credits": credits,
                "external": ext,
                "sessional": ses,
                "total": tot,
                "grade": grade,
                "grade_points": gp,
                "back": False,
            })
        for k in ("sgpa", "cgpa"):
            v = rec.get(k, "")
            try:
                rec[k] = f"{float(v):.2f}"
            except Exception:
                pass
        if rec["subjects"]:
            records.append(rec)
    return records


def parse_gs_excel_sheet(ws) -> List[dict]:
    """Parse a GS_X sheet -> list of student records.

    Block boundaries: rows where col A starts with 'Name:' AND col F contains 'University Roll'.
    """
    records: List[dict] = []
    sheet_meta = _extract_sheet_meta(ws)
    rows = ws.max_row
    block_starts: List[int] = []
    for r in range(1, rows + 1):
        v = _norm(ws.cell(r, 1).value)
        f = _norm(ws.cell(r, 6).value)
        if v.startswith("Name:") and f.startswith("University Roll"):
            block_starts.append(r)
    block_starts.append(rows + 1)

    for i in range(len(block_starts) - 1):
        r0, r1 = block_starts[i], block_starts[i + 1]
        rec: dict = {"subjects": [], **sheet_meta}
        rec["name"] = _norm(ws.cell(r0, 3).value)
        rec["roll_no"] = _norm(ws.cell(r0, 9).value)
        rec["father_name"] = _norm(ws.cell(r0 + 1, 3).value)
        rec["enroll_no"] = _norm(ws.cell(r0 + 1, 9).value)
        if not rec["roll_no"]:
            continue
        # Subjects start after the "Subject Code" header row
        for rr in range(r0 + 3, r1):
            code = _norm(ws.cell(rr, 1).value)
            name = _norm(ws.cell(rr, 3).value)
            credits = _norm(ws.cell(rr, 8).value)
            grade = _norm(ws.cell(rr, 9).value)
            gp = _norm(ws.cell(rr, 10).value)
            low = code.lower()
            if not code:
                continue
            if low.startswith("total"):
                continue
            if low.startswith("earned"):
                # Earned Credits / SGPA / CGPA row
                rec["earned_credits"] = _norm(ws.cell(rr, 3).value)
                rec["sgpa"] = _norm(ws.cell(rr, 6).value)
                rec["cgpa"] = _norm(ws.cell(rr, 9).value)
                continue
            if low.startswith("result"):
                rec["result"] = _norm(ws.cell(rr, 2).value)
                rec["remark"] = _norm(ws.cell(rr, 6).value)
                continue
            if re.match(r"^[A-Z]{2,4}\s?\d{2,4}$", code):
                rec["subjects"].append({
                    "code": code,
                    "name": name,
                    "credits": credits,
                    "grade": grade,
                    "grade_points": gp,
                    "back": False,
                })
        for k in ("sgpa", "cgpa"):
            v = rec.get(k, "")
            try:
                rec[k] = f"{float(v):.2f}"
            except Exception:
                pass
        if rec["subjects"]:
            records.append(rec)
    return records


def parse_tc_gs_excel(file_bytes: bytes) -> Dict[str, Dict[str, List[dict]]]:
    """Parse all TC_X / GS_X sheets. Returns {'TC': {sem: records}, 'GS': {sem: records}}."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    out: Dict[str, Dict[str, List[dict]]] = {"TC": {}, "GS": {}}
    for sname in wb.sheetnames:
        m = _SEM_RE.match(sname.strip())
        if not m:
            continue
        kind = m.group(1).upper()
        sem = m.group(2).upper()
        ws = wb[sname]
        if kind == "TC":
            out["TC"][sem] = parse_tc_excel_sheet(ws)
        else:
            out["GS"][sem] = parse_gs_excel_sheet(ws)
    return out


# ---------------------------------------------------------------------------
# PDF parsing — TC (multi-student per page)
# ---------------------------------------------------------------------------

_HEADER_PATTERNS = [
    "GOVIND BALLABH PANT INSTITUTE",
    "PAURI GARHWAL",
    "(AN AUTONOMOUS",
    "(AFFILIATED TO",
]


def _detect_program_branch_sem(text: str) -> Tuple[str, str, str]:
    """Return (program, branch, semester_roman) from a TC/GS page header."""
    program = ""
    branch = ""
    sem = ""
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if "Tabulation Chart for" in line:
            program = line.replace("Tabulation Chart for", "").strip()
        elif line.startswith("Bachelor of") or line.startswith("Master of"):
            program = line
        m = re.match(r"^([IVX]+)\s+Semester\s+([A-Za-z]+\s+\d{4})", line)
        if m:
            sem = m.group(1).upper()
            continue
        # Branch is usually the line right after program if it's not the sem line
        if program and not sem and line != program and "Semester" not in line:
            if not any(h in line for h in _HEADER_PATTERNS):
                if not branch:
                    branch = line
    return program, branch, sem


def parse_tc_pdf(file_bytes: bytes) -> List[dict]:
    """Parse Tabulation Chart PDF -> list of student records."""
    students: List[dict] = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        program, branch, sem = "", "", ""
        all_text = []
        for page in pdf.pages:
            text = page.extract_text() or ""
            p, b, s = _detect_program_branch_sem(text)
            program = p or program
            branch = b or branch
            sem = s or sem
            all_text.append(text)
        full = "\n".join(all_text)

        # Each student block starts with "MM N SC SN C ..." marker line.
        blocks = re.split(r"\n(?=MM\s+\d+\s+SC\s+SN\s+C\b)", full)
        for blk in blocks:
            if "University Roll No" not in blk:
                continue
            rec = _parse_tc_block(blk)
            if rec:
                rec["program"] = program
                rec["branch"] = branch
                rec["semester"] = sem
                students.append(rec)
    return students


def _parse_tc_block(block: str) -> Optional[dict]:
    rec: dict = {"subjects": []}

    # Roll & enrollment
    roll = re.search(r"University Roll No:?\s*(\S+)", block)
    enr = re.search(r"University Enrol(?:lment|\.)?\s*No:?\s*(\S+)", block)
    if not roll:
        return None
    rec["roll_no"] = roll.group(1).strip()
    rec["enroll_no"] = enr.group(1).strip() if enr else ""

    # The "Name:" line in the TC PDF collapses Name and Father's Name on a
    # single line: "Name: <Student> <Father> University Roll No: ..."
    # We extract the trailing student name from the MM line if present.
    mm = re.search(r"^MM\s+\d+\s+SC\s+SN\s+C\s+(.+)$", block, re.M)
    student_in_mm = ""
    if mm:
        candidate = mm.group(1).strip()
        if candidate and not re.match(r"^[IVX]+$", candidate):
            student_in_mm = candidate
    name_line = re.search(
        r"Name:\s*(.+?)\s+University Roll No:?\s*\S+",
        block.replace("\n", " "),
    )
    name_block = name_line.group(1).strip() if name_line else ""
    if student_in_mm and student_in_mm in name_block:
        father = name_block.replace(student_in_mm, "", 1).strip()
        rec["name"] = student_in_mm
        rec["father_name"] = father
    else:
        # No reliable split — store full blob in name; admin can correct via GS upload
        parts = name_block.split()
        if len(parts) >= 2:
            mid = len(parts) // 2
            rec["name"] = " ".join(parts[:mid])
            rec["father_name"] = " ".join(parts[mid:])
        else:
            rec["name"] = name_block
            rec["father_name"] = ""

    # Subjects: lines like "AHT 002 Engineering Chemistry 4 45/100 37/50 82/150 E+ 5.5"
    subj_re = re.compile(
        r"^([A-Z]{2,4}\s?\d{2,4})\s+(.+?)\s+(\d+)\s+([\d\-]+/\d+|-)\s+(\d+/\d+)\s+(\d+/\d+)\s+([A-Z][+\-]?|Excellent|F)\s*([\d.]+)?\s*$"
    )
    for line in block.splitlines():
        line = line.strip()
        m = subj_re.match(line)
        if m:
            rec["subjects"].append({
                "code": m.group(1).strip(),
                "name": m.group(2).strip(),
                "credits": m.group(3),
                "external": m.group(4),
                "sessional": m.group(5),
                "total": m.group(6),
                "grade": m.group(7),
                "grade_points": m.group(8) or "",
                "back": False,
            })

    # Result / SGPA / CGPA
    flat = block.replace("\n", " ")
    sg = re.search(r"SGPA:\s*([\d.]+)", flat)
    cg = re.search(r"CGPA:\s*([\d.]+)", flat)
    res = re.search(r"Result:\s*([A-Z ]+?)\s+(?:Remark|SGPA)", flat)
    rem = re.search(r"Remark:\s*(.*?)\s*(?=SGPA|Cuml|Earned|MM\s+\d+|$)", flat)
    rec["sgpa"] = sg.group(1) if sg else ""
    rec["cgpa"] = cg.group(1) if cg else ""
    rec["result"] = res.group(1).strip() if res else ""
    raw_rem = (rem.group(1).strip() if rem else "")
    rec["remark"] = raw_rem[:200]
    return rec if rec.get("subjects") else None


# ---------------------------------------------------------------------------
# PDF parsing — GS (one student per page)
# ---------------------------------------------------------------------------


def parse_gs_pdf(file_bytes: bytes) -> List[dict]:
    students: List[dict] = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        program, branch, sem = "", "", ""
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                ls = line.strip()
                if ls.startswith("Bachelor of") or ls.startswith("Master of"):
                    program = ls
                m = re.match(r"^([IVX]+)\s+Semester\s+([A-Za-z]+\s+\d{4})", ls)
                if m:
                    sem = m.group(1).upper()
            # branch line is between program and semester
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            for i, l in enumerate(lines):
                if l == program and i + 1 < len(lines):
                    nxt = lines[i + 1]
                    if "Semester" not in nxt and not any(h in nxt for h in _HEADER_PATTERNS):
                        branch = nxt
                        break

            rec = _parse_gs_page(text)
            if rec:
                rec["program"] = program
                rec["branch"] = branch
                rec["semester"] = sem
                students.append(rec)
    return students


def _parse_gs_page(text: str) -> Optional[dict]:
    rec: dict = {"subjects": []}
    flat = text.replace("\n", " ")

    m = re.search(
        r"Name:\s*(.+?)\s+University Roll No\.?:\s*(\S+).*?Father'?s?\s*Name:\s*(.+?)\s+University Enrollment No\.?:\s*(\S+)",
        flat,
    )
    if m:
        rec["name"] = m.group(1).strip()
        rec["roll_no"] = m.group(2).strip()
        rec["father_name"] = m.group(3).strip()
        rec["enroll_no"] = m.group(4).strip()
    else:
        roll = re.search(r"University Roll No\.?:\s*(\S+)", flat)
        if not roll:
            return None
        rec["roll_no"] = roll.group(1).strip()
        rec["name"] = ""
        rec["father_name"] = ""
        rec["enroll_no"] = ""

    sl = re.search(r"SL\.?\s*NO\.?:\s*(\S+)", flat)
    rec["sl_no"] = sl.group(1) if sl else ""

    subj_re = re.compile(
        r"^([A-Z]{2,4}\s?\d{2,4})\s+(.+?)\s+(\d+)\s+([A-Z][+\-]?|Excellent|F)\s*([\d.]+)?\s*$"
    )
    for line in text.splitlines():
        ls = line.strip()
        m2 = subj_re.match(ls)
        if m2:
            rec["subjects"].append({
                "code": m2.group(1).strip(),
                "name": m2.group(2).strip(),
                "credits": m2.group(3),
                "grade": m2.group(4),
                "grade_points": m2.group(5) or "",
                "back": False,
            })

    sg = re.search(r"SGPA:\s*([\d.]+)", flat)
    cg = re.search(r"CGPA:\s*([\d.]+)", flat)
    ec = re.search(r"Earned Credits:\s*(\d+)", flat)
    res = re.search(r"Result:\s*([A-Z ]+?)\s+Remark:", flat)
    rem = re.search(r"Remark:\s*(.+?)(?:Prepared by|$)", flat)
    rec["sgpa"] = sg.group(1) if sg else ""
    rec["cgpa"] = cg.group(1) if cg else ""
    rec["earned_credits"] = ec.group(1) if ec else ""
    rec["result"] = res.group(1).strip() if res else ""
    rec["remark"] = (rem.group(1).strip() if rem else "")[:200]
    return rec if rec.get("subjects") else None


# ---------------------------------------------------------------------------
# Apply asterisk markers
# ---------------------------------------------------------------------------


def apply_back_markers(records: List[dict], back_map_for_sem: Dict[str, Dict[str, bool]]) -> Tuple[int, int]:
    """Mutate records: append ' *' to subject name when SEM excel marks it as
    back AND the student has cleared the paper (i.e. grade is not F / Ab /
    Dt). The '*' therefore means *subject cleared after a back paper*.

    Returns (markers_applied, students_matched).
    """
    n = 0
    matched = 0
    record_rolls = {r.get("roll_no", "") for r in records}
    skip_grades = {"F", "AB", "ABS", "DT"}
    for rec in records:
        roll = rec.get("roll_no", "")
        backs = back_map_for_sem.get(roll, {})
        if backs:
            matched += 1
        for s in rec.get("subjects", []):
            code = s["code"].replace(" ", "").upper()
            grade = (s.get("grade") or "").strip().upper()
            for bcode in backs:
                if bcode.replace(" ", "").upper() == code:
                    if grade in skip_grades:
                        # back marker still recorded for analytics, but no *.
                        s["back"] = True
                        s["back_pending"] = True
                        break
                    s["back"] = True
                    if not s["name"].rstrip().endswith("*"):
                        s["name"] = s["name"].rstrip() + " *"
                    n += 1
                    break
    if back_map_for_sem and matched == 0 and record_rolls:
        sample_excel = next(iter(back_map_for_sem.keys()))
        sample_pdf = next(iter(record_rolls))
        import logging as _log
        _log.getLogger("portal").warning(
            "apply_back_markers: 0 of %d Excel rolls matched any PDF roll. "
            "Excel sample=%s vs PDF sample=%s — Excel and PDF may belong to "
            "different batches.",
            len(back_map_for_sem), sample_excel, sample_pdf,
        )
    return n, matched


def apply_non_credit_markers(records: List[dict]) -> int:
    """Append ' $' to subject names whose credits are 0 / blank — these are
    non-credit subjects (e.g. General Proficiency). Idempotent."""
    n = 0
    for rec in records:
        for s in rec.get("subjects", []):
            cred = (s.get("credits") or "").strip()
            try:
                is_zero = int(cred) == 0
            except Exception:
                is_zero = cred in ("", "-", "—", "0")
            if is_zero:
                s["non_credit"] = True
                if "$" not in s["name"]:
                    s["name"] = s["name"].rstrip(" *") + " $"
                    if s.get("back") and not s.get("back_pending"):
                        s["name"] += " *"
                    n += 1
    return n


# ---------------------------------------------------------------------------
# Generate PDF — TC*  (Tabulation Chart with asterisks)
# ---------------------------------------------------------------------------

_INSTITUTE_LINES = [
    "GOVIND BALLABH PANT INSTITUTE OF ENGINEERING AND TECHNOLOGY",
    "PAURI GARHWAL, UTTARAKHAND",
    "(AN AUTONOMOUS INSTITUTE OF GOVT. OF UTTARAKHAND)",
    "(AFFILIATED TO VEER MADHO SINGH BHANDARI UTTARAKHAND TECHNICAL UNIVERSITY)",
]

# Logo paths — extracted once from the original sample PDFs and bundled with the app.
_ASSETS = Path(__file__).parent / "assets"
INSTITUTE_LOGO = _ASSETS / "institute_logo.png"
UTU_LOGO = _ASSETS / "utu_logo.png"


from reportlab.platypus import Image as RLImage  # noqa: E402

# Barcode generation (Code 128) — used on every GS page
import barcode as _barcode_lib  # noqa: E402
from barcode.writer import ImageWriter  # noqa: E402


def _make_barcode_png(code_text: str) -> Optional[io.BytesIO]:
    """Generate a Code-128 barcode PNG for the given text. Returns BytesIO."""
    if not code_text:
        return None
    try:
        Code128 = _barcode_lib.get_barcode_class("code128")
        buf = io.BytesIO()
        # write_text=False keeps the image clean; we render text below ourselves
        Code128(code_text, writer=ImageWriter()).write(
            buf,
            options={"module_width": 0.30, "module_height": 8.0, "write_text": False, "quiet_zone": 1.0},
        )
        buf.seek(0)
        return buf
    except Exception:
        return None


def _logo_header(text_lines, total_width_mm: float, logo_size_mm: float = 22):
    """Build a 3-column row: [institute logo] [stacked text lines] [UTU logo]."""
    text_cell = []
    for txt, style in text_lines:
        text_cell.append(Paragraph(txt, style))

    inst = (
        RLImage(str(INSTITUTE_LOGO), width=logo_size_mm * mm, height=logo_size_mm * mm)
        if INSTITUTE_LOGO.exists()
        else ""
    )
    utu = (
        RLImage(str(UTU_LOGO), width=logo_size_mm * mm, height=logo_size_mm * mm)
        if UTU_LOGO.exists()
        else ""
    )
    text_w = total_width_mm - 2 * (logo_size_mm + 2)
    t = Table(
        [[inst, text_cell, utu]],
        colWidths=[(logo_size_mm + 2) * mm, text_w * mm, (logo_size_mm + 2) * mm],
    )
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("ALIGN", (2, 0), (2, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _styles():
    s = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("title", parent=s["Heading1"], fontName="Helvetica-Bold",
                                 fontSize=11, alignment=1, spaceAfter=2, leading=13),
        "sub": ParagraphStyle("sub", parent=s["Normal"], fontName="Helvetica",
                               fontSize=9, alignment=1, spaceAfter=1, leading=11),
        "small": ParagraphStyle("small", parent=s["Normal"], fontName="Helvetica",
                                 fontSize=8, alignment=1, spaceAfter=1, leading=10),
        "section": ParagraphStyle("section", parent=s["Normal"], fontName="Helvetica-Bold",
                                   fontSize=11, alignment=1, spaceBefore=4, spaceAfter=4,
                                   leading=13),
        "label": ParagraphStyle("label", parent=s["Normal"], fontName="Helvetica",
                                 fontSize=9, alignment=0, leading=11),
        "right": ParagraphStyle("right", parent=s["Normal"], fontName="Helvetica",
                                 fontSize=9, alignment=2),
        "back_subject": ParagraphStyle("back_subject", parent=s["Normal"],
                                        fontName="Helvetica-Bold", fontSize=8,
                                        textColor=colors.HexColor("#92400e"), alignment=0,
                                        leading=10),
        "subject": ParagraphStyle("subject", parent=s["Normal"], fontName="Helvetica",
                                    fontSize=8, alignment=0, leading=10),
    }


def _draw_tc_header(canv, doc, program: str, branch: str, sem: str, session: str):
    """Per-page header for TC: institute logos + name strip + course/branch/semester.

    Drawn directly on the canvas so it appears on every A3 page.
    """
    canv.saveState()
    page_w, _page_h = doc.pagesize
    margin = 12 * mm
    logo_size = 22 * mm
    top_y = _page_h - 6 * mm  # top edge of logo

    # Logos (left = institute, right = UTU) — silently skip if files missing.
    if INSTITUTE_LOGO.exists():
        try:
            canv.drawImage(
                str(INSTITUTE_LOGO), margin, top_y - logo_size,
                width=logo_size, height=logo_size,
                mask="auto", preserveAspectRatio=True,
            )
        except Exception:
            pass
    if UTU_LOGO.exists():
        try:
            canv.drawImage(
                str(UTU_LOGO), page_w - margin - logo_size, top_y - logo_size,
                width=logo_size, height=logo_size,
                mask="auto", preserveAspectRatio=True,
            )
        except Exception:
            pass

    # Centered institute text block (4 lines beside logos)
    cx = page_w / 2
    canv.setFillColor(colors.black)
    canv.setFont("Helvetica-Bold", 12)
    canv.drawCentredString(cx, top_y - 4.5 * mm, _INSTITUTE_LINES[0])
    canv.setFont("Helvetica", 9.5)
    canv.drawCentredString(cx, top_y - 9 * mm, _INSTITUTE_LINES[1])
    canv.setFont("Helvetica", 8.5)
    canv.drawCentredString(cx, top_y - 13 * mm, _INSTITUTE_LINES[2])
    canv.drawCentredString(cx, top_y - 16.5 * mm, _INSTITUTE_LINES[3])

    # Course / branch / semester block — sits just below logos, full width.
    course_y = top_y - logo_size - 5 * mm
    canv.setFont("Helvetica-Bold", 13)
    canv.drawCentredString(
        cx, course_y, f"Tabulation Chart for {program or program_short(program)}"
    )
    canv.setFont("Helvetica-Bold", 11)
    canv.drawCentredString(cx, course_y - 5 * mm, branch or "")
    canv.setFont("Helvetica", 10.5)
    sess_line = f"{sem} Semester"
    if session:
        sess_line += f"   |   {session}"
    canv.drawCentredString(cx, course_y - 9.6 * mm, sess_line)

    # Thin divider beneath header
    canv.setLineWidth(0.4)
    canv.setStrokeColor(colors.HexColor("#9ca3af"))
    canv.line(margin, course_y - 12.5 * mm, page_w - margin, course_y - 12.5 * mm)
    canv.restoreState()


def _draw_tc_footer(canv, doc, program: str, branch: str, sem: str, session: str):
    """Per-page footer for TC: 5-cell signature row + page number.
    (Legend is rendered in-flow, immediately after the last student record.)"""
    canv.saveState()
    page_w, _ = doc.pagesize
    margin = 12 * mm
    # ---- Signature row ----
    y = 14 * mm
    cell_w = (page_w - 2 * margin) / 5.0
    labels = [
        "Prepared by", "Checked by",
        "Examination Controller", "Director",
        "Examination Controller (VMSB UTU)",
    ]
    canv.setLineWidth(0.4)
    canv.setStrokeColor(colors.HexColor("#9ca3af"))
    canv.line(margin, y + 2, page_w - margin, y + 2)
    canv.setFont("Helvetica", 9)
    canv.setFillColor(colors.HexColor("#1c1917"))
    for i, lbl in enumerate(labels):
        x = margin + i * cell_w + cell_w / 2
        canv.drawCentredString(x, y - 6, lbl)
    canv.setFont("Helvetica", 7)
    canv.setFillColor(colors.HexColor("#57534e"))
    canv.drawRightString(page_w - margin, 6 * mm, f"Page {doc.page}")
    canv.restoreState()


def _draw_tc_page(canv, doc, program: str, branch: str, sem: str, session: str):
    """Combined per-page chrome: header (top) + footer (bottom)."""
    _draw_tc_header(canv, doc, program, branch, sem, session)
    _draw_tc_footer(canv, doc, program, branch, sem, session)


def generate_tc_pdf(records: List[dict], program: str = "", branch: str = "",
                     semester_roman: str = "", exam_session: str = "") -> bytes:
    """Tabulation Chart — A3 portrait, 4 students per page max, NO content
    spanning two pages (KeepTogether). Per-page signed footer. No grade
    reference. Programme / branch / semester / session are taken from the
    sheet-level metadata stored on each record (with the function arguments
    only as fallback)."""
    # Resolve metadata from the records themselves where possible.
    if records:
        meta = records[0]
        program = meta.get("program", "") or program
        branch = meta.get("branch", "") or branch
        semester_roman = meta.get("semester", "") or semester_roman
        exam_session = meta.get("exam_session", "") or exam_session

    from reportlab.platypus import (
        BaseDocTemplate, Frame, PageTemplate, KeepTogether,
    )

    buf = io.BytesIO()
    page = A3
    margin = 12 * mm
    footer_h = 28 * mm  # legend + signature band
    header_h = 50 * mm  # institute strip + course/branch/semester (drawn on canvas)
    doc = BaseDocTemplate(
        buf, pagesize=page,
        leftMargin=margin, rightMargin=margin,
        topMargin=header_h, bottomMargin=footer_h,
    )
    frame = Frame(
        margin, footer_h, page[0] - 2 * margin, page[1] - header_h - footer_h,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        showBoundary=0,
    )
    doc.addPageTemplates([PageTemplate(
        id="tc", frames=[frame],
        onPage=lambda c, d: _draw_tc_page(c, d, program, branch, semester_roman, exam_session),
    )])

    st = _styles()
    page_width_mm = page[0] / mm - 24
    col_widths = [26 * mm, 95 * mm, 16 * mm, 28 * mm, 28 * mm, 28 * mm, 22 * mm, 26 * mm]

    story: list = []
    students_per_page = 4

    def build_student_block(rec):
        info = [[
            Paragraph(f"<b>Name:</b> {rec.get('name','')}", st["label"]),
            Paragraph(f"<b>Father's Name:</b> {rec.get('father_name','')}", st["label"]),
            Paragraph(f"<b>University Roll No.:</b> {rec.get('roll_no','')}", st["label"]),
            Paragraph(f"<b>University Enrol. No.:</b> {rec.get('enroll_no','')}", st["label"]),
        ]]
        t_info = Table(info, colWidths=[70 * mm, 70 * mm, 65 * mm, 64 * mm])
        t_info.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]))
        rows = [[
            "Subject\nCode", "Subject Name", "Credits", "External\nMarks",
            "Sessional\nMarks", "Total\nMarks", "Grade", "Grade\nPoints",
        ]]
        for s in rec.get("subjects", []):
            name_para = Paragraph(s["name"], st["back_subject"] if s.get("back") else st["subject"])
            rows.append([
                s["code"], name_para, s.get("credits", ""), s.get("external", ""),
                s.get("sessional", ""), s.get("total", ""), s.get("grade", ""),
                s.get("grade_points", ""),
            ])
        totals = _compute_totals(rec.get("subjects", []))
        rows.append([
            Paragraph("<b>Total Credits / Marks / Total Grade Points</b>", st["label"]),
            "", str(totals["credits"]), totals["external"], totals["sessional"],
            totals["total"], "", f"{totals['grade_points']:.1f}",
        ])
        rows.append([
            Paragraph(f"<b>Result:</b> {rec.get('result','—')}", st["label"]), "",
            Paragraph(f"<b>Remark:</b> {rec.get('remark','—') or '—'}", st["label"]), "",
            Paragraph(f"<b>SGPA:</b> {rec.get('sgpa','—')}", st["label"]), "",
            Paragraph(f"<b>CGPA:</b> {rec.get('cgpa','—')}", st["label"]), "",
        ])
        rows.append([
            Paragraph(f"<b>Earned Credits:</b> {rec.get('earned_credits','—')}", st["label"]),
            "", "", "",
            Paragraph(
                f"<b>Cumulative Earned Credits:</b> "
                f"{rec.get('cuml_earned_credits') or rec.get('earned_credits','—')}",
                st["label"],
            ), "", "", "",
        ])
        t = Table(rows, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("GRID", (0, 0), (-1, -4), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e1b4b")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("ALIGN", (2, 1), (-1, -4), "CENTER"),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("SPAN", (0, -3), (1, -3)),
            ("BACKGROUND", (0, -3), (-1, -3), colors.HexColor("#f5f5f4")),
            ("FONTNAME", (0, -3), (-1, -3), "Helvetica-Bold"),
            ("ALIGN", (2, -3), (-1, -3), "CENTER"),
            ("SPAN", (0, -2), (1, -2)),
            ("SPAN", (2, -2), (3, -2)),
            ("SPAN", (4, -2), (5, -2)),
            ("SPAN", (6, -2), (7, -2)),
            ("FONTNAME", (0, -2), (-1, -2), "Helvetica-Bold"),
            ("BACKGROUND", (0, -2), (-1, -2), colors.HexColor("#fef3c7")),
            ("SPAN", (0, -1), (3, -1)),
            ("SPAN", (4, -1), (7, -1)),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#fafaf9")),
        ]))
        return KeepTogether([t_info, t, Spacer(1, 2.5 * mm)])

    # Legend flowable to drop in after each chunk of 3 students.
    legend_style = ParagraphStyle(
        "tc_legend", fontName="Helvetica-Oblique", fontSize=8.5, leading=11,
        textColor=colors.HexColor("#57534e"), alignment=0,
    )
    def legend_flowable():
        return Paragraph(
            "<b>*</b>&nbsp;&nbsp;subject cleared after back paper "
            "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            "<b>$</b>&nbsp;&nbsp;non-credit subject",
            legend_style,
        )

    students_per_page = 3
    total = len(records)
    for i in range(0, total, students_per_page):
        chunk = records[i:i + students_per_page]
        for rec in chunk:
            story.append(build_student_block(rec))
        # Legend right after the last student of this page
        story.append(Spacer(1, 1.5 * mm))
        story.append(legend_flowable())
        if i + students_per_page < total:
            story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Generate PDF — GS*  (Grade Sheet with asterisks + barcode)
# ---------------------------------------------------------------------------


def _draw_gs_decoration(canv, doc, watermark_path: str = ""):
    """Per-page decorations for GS: institute-logo watermark only (no footer text)."""
    canv.saveState()
    page_w, page_h = doc.pagesize
    if watermark_path and Path(watermark_path).exists():
        try:
            canv.setFillAlpha(0.07)
            wmark_w = 130 * mm
            wmark_h = 130 * mm
            canv.drawImage(
                watermark_path,
                (page_w - wmark_w) / 2, (page_h - wmark_h) / 2,
                width=wmark_w, height=wmark_h, mask="auto",
                preserveAspectRatio=True,
            )
            canv.setFillAlpha(1)
        except Exception:
            pass
    canv.restoreState()


def generate_gs_pdf(records: List[dict], program: str = "", branch: str = "",
                     semester_roman: str = "", exam_session: str = "",
                     batch: str = "",
                     all_sem_summary: Optional[Dict[str, Dict[str, Dict[str, str]]]] = None) -> bytes:
    """Polished A4-portrait Grade Sheet.

    `all_sem_summary` (optional): {roll_no: {SEM: {sgpa, cgpa, earned_credits,
    result}}} — used to render the per-semester SGPA/CGPA history table at
    the bottom of every page.
    """
    if records:
        meta = records[0]
        program = meta.get("program", "") or program
        branch = meta.get("branch", "") or branch
        semester_roman = meta.get("semester", "") or semester_roman
        exam_session = meta.get("exam_session", "") or exam_session

    from reportlab.platypus import (
        BaseDocTemplate, Frame, PageTemplate, KeepTogether,
    )

    buf = io.BytesIO()
    page = A4
    margin = 18 * mm
    doc = BaseDocTemplate(
        buf, pagesize=page,
        leftMargin=margin, rightMargin=margin,
        topMargin=12 * mm, bottomMargin=22 * mm,
    )
    frame = Frame(
        margin, 22 * mm, page[0] - 2 * margin, page[1] - 12 * mm - 22 * mm,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        showBoundary=0,
    )
    doc.addPageTemplates([PageTemplate(
        id="gs", frames=[frame],
        onPage=lambda c, d: _draw_gs_decoration(c, d, str(INSTITUTE_LOGO)),
    )])

    st = _styles()
    page_width_mm = page[0] / mm - 36
    story = []

    for idx, rec in enumerate(records):
        # ---- Top-right barcode (replaces SL.NO.) ----
        bc_top = _make_barcode_png(rec.get("roll_no", ""))
        if bc_top is not None:
            try:
                top_bc = RLImage(bc_top, width=44 * mm, height=11 * mm)
                top_table = Table(
                    [["", top_bc]],
                    colWidths=[(page_width_mm - 50) * mm, 50 * mm],
                )
                top_table.setStyle(TableStyle([
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (-1, 0), (-1, 0), "RIGHT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ]))
                story.append(top_table)
            except Exception:
                pass

        # ---- Institute / UTU strip — institute and university each on its own line ----
        single_line_styles = [
            (
                "<font size='11.5' face='Helvetica-Bold'>GOVIND BALLABH PANT INSTITUTE OF ENGINEERING AND TECHNOLOGY</font>",
                st["sub"],
            ),
            (
                "<font size='8.5'>Pauri Garhwal, Uttarakhand &middot; An Autonomous Institute of the Government of Uttarakhand</font>",
                st["small"],
            ),
            (
                "<font size='9.5' face='Helvetica-Bold'>Affiliated to Veer Madho Singh Bhandari Uttarakhand Technical University</font>",
                st["sub"],
            ),
        ]
        story.append(_logo_header(single_line_styles, total_width_mm=page_width_mm, logo_size_mm=22))
        story.append(Spacer(1, 2 * mm))

        # ---- Title band — programme / branch / semester+session each on its own line ----
        title_band = Table([[Paragraph(
            "<para alignment='center'>"
            "<font size='15' face='Helvetica-Bold' color='white'>GRADE SHEET</font><br/>"
            f"<font size='11' face='Helvetica-Bold' color='white'>{program}</font><br/>"
            f"<font size='10' color='white'>{branch}</font><br/>"
            f"<font size='9.5' color='#cbd5e1'>{semester_roman} Semester &nbsp;&middot;&nbsp; {exam_session}</font>"
            "</para>",
            st["sub"],
        )]], colWidths=[page_width_mm * mm])
        title_band.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#1e1b4b")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]))
        story.append(title_band)
        story.append(Spacer(1, 3 * mm))

        # ---- Student details card ----
        info = Table([
            [Paragraph("<font size='7.5' color='#57534e'><b>NAME</b></font>", st["small"]),
             Paragraph(f"<font size='10.5'><b>{rec.get('name','')}</b></font>", st["label"]),
             Paragraph("<font size='7.5' color='#57534e'><b>UNIVERSITY ROLL NO.</b></font>", st["small"]),
             Paragraph(f"<font size='10.5' face='Helvetica-Bold'>{rec.get('roll_no','')}</font>",
                       st["label"])],
            [Paragraph("<font size='7.5' color='#57534e'><b>FATHER'S NAME</b></font>", st["small"]),
             Paragraph(f"<font size='10'>{rec.get('father_name','')}</font>", st["label"]),
             Paragraph("<font size='7.5' color='#57534e'><b>UNIVERSITY ENROLLMENT NO.</b></font>", st["small"]),
             Paragraph(f"<font size='10' face='Helvetica-Bold'>{rec.get('enroll_no','')}</font>", st["label"])],
        ], colWidths=[34 * mm, 60 * mm, 40 * mm, 40 * mm])
        info.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#9ca3af")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#e5e7eb")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f5f5f4")),
            ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#f5f5f4")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(info)
        story.append(Spacer(1, 3 * mm))

        # ---- Subjects table ----
        rows = [["Subject Code", "Subject Name", "Credits", "Grade", "Grade Points"]]
        total_credits = 0
        total_gp = 0.0
        for s in rec.get("subjects", []):
            name_para = Paragraph(s["name"], st["back_subject"] if s.get("back") else st["subject"])
            rows.append([s["code"], name_para, s.get("credits", ""),
                         s.get("grade", ""), s.get("grade_points", "")])
            try:
                total_credits += int(s.get("credits") or 0)
            except Exception:
                pass
            try:
                total_gp += float(s.get("grade_points") or 0) * int(s.get("credits") or 0)
            except Exception:
                pass
        rows.append([
            Paragraph("<b>Total Credits / Total Grade Points</b>", st["label"]),
            "", str(total_credits), "—", f"{total_gp:.1f}",
        ])
        col_widths = [27 * mm, 78 * mm, 22 * mm, 22 * mm, 25 * mm]
        t = Table(rows, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e1b4b")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, 0), 9.5),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("ALIGN", (2, 1), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2),
              [colors.white, colors.HexColor("#fafaf9")]),
            ("SPAN", (0, -1), (1, -1)),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#fef3c7")),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 3 * mm))

        # ---- Semester-wise Result — current + previous semesters only ----
        roll = rec.get("roll_no", "")
        per_sem = (all_sem_summary or {}).get(roll, {})
        sem_order = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
        try:
            current_idx = sem_order.index(semester_roman)
        except ValueError:
            current_idx = len(sem_order) - 1
        visible_sems = sem_order[:current_idx + 1]
        hist_header = ["Semester"] + visible_sems
        sgpa_row = ["SGPA"]
        cgpa_row = ["CGPA"]
        ec_row = ["Earned Cr."]
        res_row = ["Result"]
        for s in visible_sems:
            cell = per_sem.get(s, {})
            if not cell and s == semester_roman:
                cell = {
                    "sgpa": rec.get("sgpa", ""),
                    "cgpa": rec.get("cgpa", ""),
                    "earned_credits": rec.get("earned_credits", ""),
                    "result": rec.get("result", ""),
                }
            sgpa_row.append(cell.get("sgpa") or "—")
            cgpa_row.append(cell.get("cgpa") or "—")
            ec_row.append(cell.get("earned_credits") or "—")
            res_row.append((cell.get("result") or "—")[:6])
        n_cols = len(visible_sems)
        col_w = (page_width_mm - 24) / max(n_cols, 1)
        history = Table([hist_header, sgpa_row, cgpa_row, ec_row, res_row],
                          colWidths=[24 * mm] + [col_w * mm] * n_cols,
                          repeatRows=1)
        history.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#9ca3af")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#e5e7eb")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e1b4b")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#f5f5f4")),
            ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(Paragraph(
            "<font size='8.5' color='#57534e'><b>SEMESTER-WISE RESULT</b></font>",
            st["label"],
        ))
        story.append(Spacer(1, 1 * mm))
        story.append(history)
        story.append(Spacer(1, 3 * mm))

        # ---- Summary card (SGPA / CGPA / Earned credits / Result) ----
        sgpa = rec.get("sgpa") or "—"
        cgpa = rec.get("cgpa") or "—"
        earned = rec.get("earned_credits") or str(total_credits)
        result_val = rec.get("result") or "—"
        remark_val = rec.get("remark") or ""
        summary_top = Table([[
            Paragraph(f"<font size='7.5' color='#57534e'>SGPA</font><br/>"
                      f"<font size='15' face='Helvetica-Bold'>{sgpa}</font>", st["label"]),
            Paragraph(f"<font size='7.5' color='#57534e'>CGPA</font><br/>"
                      f"<font size='15' face='Helvetica-Bold'>{cgpa}</font>", st["label"]),
            Paragraph(f"<font size='7.5' color='#57534e'>EARNED CREDITS</font><br/>"
                      f"<font size='15' face='Helvetica-Bold'>{earned}</font>", st["label"]),
            Paragraph(f"<font size='7.5' color='#57534e'>RESULT</font><br/>"
                      f"<font size='15' face='Helvetica-Bold'>{result_val}</font>", st["label"]),
        ]], colWidths=[(page_width_mm / 4) * mm] * 4)
        summary_top.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#9ca3af")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#e5e7eb")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fafaf9")),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(summary_top)
        if remark_val:
            story.append(Spacer(1, 1.5 * mm))
            story.append(Paragraph(
                f"<font size='8' color='#57534e'><b>REMARK:</b></font> "
                f"<font size='9'>{remark_val}</font>",
                st["label"],
            ))

        story.append(Spacer(1, 14 * mm))

        # ---- Signature row at the bottom (Director AFTER Examination Controller) ----
        sig = Table([
            ["Prepared by", "Checked by", "Examination Controller", "Director"]
        ], colWidths=[(page_width_mm / 4) * mm] * 4,
            style=TableStyle([
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1c1917")),
                ("TOPPADDING", (0, 0), (-1, -1), 22),
                ("LINEABOVE", (0, 0), (-1, 0), 0.5, colors.HexColor("#9ca3af")),
            ]))
        story.append(KeepTogether([sig]))
        if idx + 1 < len(records):
            story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()


def _branch_short(branch: str) -> str:
    if not branch:
        return "XX"
    parts = re.findall(r"[A-Z]", branch.upper())
    return "".join(parts)[:4] or "XX"


def program_short(program: str) -> str:
    """Return canonical short label for a programme.

    Accepts long names ('Bachelor of Technology', 'Master of Computer
    Applications', 'Master of Technology') as well as the short literal forms
    used by the frontend ('B.Tech', 'M.Tech', 'MCA'). Falls back to 'PROG'.
    """
    if not program:
        return "PROG"
    norm = re.sub(r"[^A-Z]", "", program.upper())  # strip dots/spaces
    if "MASTEROFCOMPUTER" in norm or norm == "MCA":
        return "MCA"
    if "MASTEROFTECH" in norm or norm.startswith("MTECH"):
        return "M. TECH"
    if "BACHELOR" in norm or norm.startswith("BTECH"):
        return "B. TECH"
    if "MASTER" in norm:
        return "M. TECH"
    return "PROG"


# ---------------------------------------------------------------------------
# Grade tables — reference tables embedded at the bottom of TC* / GS*
# Source: GBPIET ordinances (M.Tech) and Academic Council minutes 2025
# (B.Tech & MCA, Section 22.10).
# ---------------------------------------------------------------------------

_GRADE_TABLE_BTECH_MCA_2025 = [
    ("Letter Grade", "Grade Point", "Marks Range"),
    ("O", "10", "≥ 95%"),
    ("A+", "9.5", "90% – < 95%"),
    ("A", "9", "85% – < 90%"),
    ("B+", "8.5", "80% – < 85%"),
    ("B", "8", "75% – < 80%"),
    ("C+", "7.5", "70% – < 75%"),
    ("C", "7", "65% – < 70%"),
    ("D+", "6.5", "60% – < 65%"),
    ("D", "6", "55% – < 60%"),
    ("E+", "5.5", "50% – < 55%"),
    ("E", "5", "45% – < 50%"),
    ("P", "4.5", "40% – < 45%"),
    ("F", "0", "< 40%"),
]

_GRADE_TABLE_MTECH = [
    ("Letter Grade", "Grade Point", "Marks Range"),
    ("O — Outstanding", "10", "≥ 90%"),
    ("A+ — Excellent", "9", "85% – < 90%"),
    ("A — Very Good", "8", "80% – < 85%"),
    ("B — Good", "7", "70% – < 80%"),
    ("C — Average", "6", "60% – < 70%"),
    ("P — Pass", "5", "50% – < 60%"),
    ("F — Fail", "0", "< 50%"),
    ("AB — Absent", "—", "Absent"),
]


def _grade_table_for_program(program: str) -> List[Tuple[str, str, str]]:
    if program_short(program) == "M. TECH":
        return _GRADE_TABLE_MTECH
    return _GRADE_TABLE_BTECH_MCA_2025


def _grade_reference_block(program: str) -> Table:
    rows = _grade_table_for_program(program)
    t = Table(rows, colWidths=[40 * mm, 25 * mm, 50 * mm], hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.4, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e1b4b")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _sum_marks(records_subjects: List[dict], key: str) -> Tuple[int, int]:
    """Sum 'a/b' style marks columns. Returns (sum_a, sum_b)."""
    sa = sb = 0
    for s in records_subjects:
        v = (s.get(key) or "").strip()
        m = re.match(r"^(-?\d+)\s*/\s*(\d+)$", v)
        if m:
            sa += int(m.group(1)) if m.group(1) != "-" else 0
            sb += int(m.group(2))
    return sa, sb


def _compute_totals(subjects: List[dict]) -> dict:
    tc = 0
    tgp = 0.0
    for s in subjects:
        try:
            c = int(s.get("credits") or 0)
        except Exception:
            c = 0
        try:
            gp = float(s.get("grade_points") or 0)
        except Exception:
            gp = 0.0
        tc += c
        tgp += c * gp
    ext_a, ext_b = _sum_marks(subjects, "external")
    ses_a, ses_b = _sum_marks(subjects, "sessional")
    tot_a, tot_b = _sum_marks(subjects, "total")
    return {
        "credits": tc,
        "grade_points": tgp,
        "external": f"{ext_a}/{ext_b}" if ext_b else "",
        "sessional": f"{ses_a}/{ses_b}" if ses_b else "",
        "total": f"{tot_a}/{tot_b}" if tot_b else "",
    }

