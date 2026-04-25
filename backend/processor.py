"""
PDF & Excel processing utilities for the Result Asterisk Portal.

- Parses TC (Tabulation Chart) PDFs: multi-student-per-page tabular dump.
- Parses GS (Provisional Grade Sheet) PDFs: one student per page.
- Parses SEM_X excel sheets: detects highlighted (yellow / blue) cells per
  (roll_no, subject_code) which represent BACK / failed subjects.
- Generates new TC*/GS* PDFs in the same visual format with " *" appended
  to subject names where a back exists.
"""
from __future__ import annotations

import io
import re
from typing import Dict, List, Optional, Tuple

import openpyxl
import pdfplumber
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
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


def apply_back_markers(records: List[dict], back_map_for_sem: Dict[str, Dict[str, bool]]) -> int:
    """Mutate records: append ' *' to subject name when SEM excel marks it as back. Return count."""
    n = 0
    for rec in records:
        roll = rec.get("roll_no", "")
        backs = back_map_for_sem.get(roll, {})
        for s in rec.get("subjects", []):
            code = s["code"].replace(" ", "").upper()
            for bcode in backs:
                if bcode.replace(" ", "").upper() == code:
                    s["back"] = True
                    if not s["name"].rstrip().endswith("*"):
                        s["name"] = s["name"].rstrip() + " *"
                    n += 1
                    break
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


def _styles():
    s = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("title", parent=s["Heading1"], fontName="Helvetica-Bold",
                                 fontSize=11, alignment=1, spaceAfter=2),
        "sub": ParagraphStyle("sub", parent=s["Normal"], fontName="Helvetica",
                               fontSize=9, alignment=1, spaceAfter=1),
        "small": ParagraphStyle("small", parent=s["Normal"], fontName="Helvetica",
                                 fontSize=8, alignment=1, spaceAfter=1),
        "section": ParagraphStyle("section", parent=s["Normal"], fontName="Helvetica-Bold",
                                   fontSize=10, alignment=1, spaceBefore=4, spaceAfter=4),
        "label": ParagraphStyle("label", parent=s["Normal"], fontName="Helvetica",
                                 fontSize=9, alignment=0),
        "back_subject": ParagraphStyle("back_subject", parent=s["Normal"],
                                        fontName="Helvetica-Bold", fontSize=8,
                                        textColor=colors.HexColor("#92400e"), alignment=0),
        "subject": ParagraphStyle("subject", parent=s["Normal"], fontName="Helvetica",
                                    fontSize=8, alignment=0),
    }


def generate_tc_pdf(records: List[dict], program: str, branch: str, semester_roman: str,
                     exam_session: str = "December 2025") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=10 * mm,
                             rightMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm)
    st = _styles()
    story = []

    def header_block():
        story.append(Paragraph(_INSTITUTE_LINES[0], st["title"]))
        story.append(Paragraph(_INSTITUTE_LINES[1], st["sub"]))
        story.append(Paragraph(_INSTITUTE_LINES[2], st["small"]))
        story.append(Paragraph(_INSTITUTE_LINES[3], st["small"]))
        story.append(Paragraph(f"Tabulation Chart for {program}", st["section"]))
        story.append(Paragraph(branch, st["sub"]))
        story.append(Paragraph(f"{semester_roman} Semester {exam_session}", st["sub"]))
        story.append(Spacer(1, 4 * mm))

    header_block()

    students_per_page = 3
    for idx, rec in enumerate(records):
        # Student info
        info = [[
            Paragraph(f"<b>Name:</b> {rec.get('name','')}", st["label"]),
            Paragraph(f"<b>Father's Name:</b> {rec.get('father_name','')}", st["label"]),
            Paragraph(f"<b>University Roll No:</b> {rec.get('roll_no','')}", st["label"]),
            Paragraph(f"<b>University Enrol. No:</b> {rec.get('enroll_no','')}", st["label"]),
        ]]
        t_info = Table(info, colWidths=[70 * mm, 70 * mm, 65 * mm, 70 * mm])
        t_info.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(t_info)

        # Subject table
        rows = [["Subject\nCode", "Subject Name", "Credits", "External\nMarks",
                 "Sessional\nMarks", "Total\nMarks", "Grade", "Grade\nPoints"]]
        for s in rec.get("subjects", []):
            name_para = Paragraph(s["name"], st["back_subject"] if s.get("back") else st["subject"])
            rows.append([
                s["code"], name_para, s["credits"], s.get("external", ""),
                s.get("sessional", ""), s.get("total", ""), s["grade"],
                s.get("grade_points", ""),
            ])
        # Summary row
        rows.append(["Total Credits/ Marks/ Total Grade Points", "", "", "", "", "", "", ""])
        # Result row
        rows.append([
            f"Result: {rec.get('result','')}", "",
            f"Remark: {rec.get('remark','')}", "",
            f"SGPA: {rec.get('sgpa','')}", "",
            f"CGPA: {rec.get('cgpa','')}", "",
        ])
        col_widths = [22 * mm, 95 * mm, 16 * mm, 22 * mm, 22 * mm, 22 * mm, 18 * mm, 20 * mm]
        t = Table(rows, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("GRID", (0, 0), (-1, -3), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e1b4b")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 7.5),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("ALIGN", (2, 1), (-1, -3), "CENTER"),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("SPAN", (0, -2), (-1, -2)),  # summary spans
            ("SPAN", (0, -1), (1, -1)),
            ("SPAN", (2, -1), (3, -1)),
            ("SPAN", (4, -1), (5, -1)),
            ("SPAN", (6, -1), (7, -1)),
            ("BACKGROUND", (0, -2), (-1, -2), colors.HexColor("#f5f5f4")),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ]))
        story.append(t)
        story.append(Spacer(1, 5 * mm))

        if (idx + 1) % students_per_page == 0 and idx + 1 < len(records):
            story.append(PageBreak())
            header_block()

    # Footer (last page)
    story.append(Spacer(1, 10 * mm))
    story.append(Table([["Prepared by", "Checked by (Tabulators)", "Controller (GBPIET)",
                          "Director (GBPIET)", "Controller (UTU)"]],
                        colWidths=[55 * mm] * 5,
                        style=TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                           ("FONTSIZE", (0, 0), (-1, -1), 8)])))

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Generate PDF — GS*  (Provisional Grade Sheet with asterisks)
# ---------------------------------------------------------------------------


def generate_gs_pdf(records: List[dict], program: str, branch: str, semester_roman: str,
                     exam_session: str = "December 2025", batch: str = "") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                             topMargin=14 * mm, bottomMargin=14 * mm)
    st = _styles()
    story = []
    short = _branch_short(branch)
    for idx, rec in enumerate(records):
        sl_no = rec.get("sl_no") or f"{program_short(program)}/{short}/{batch}/{500 + idx + 1}"
        story.append(Paragraph(f"SL. NO.: {sl_no}", st["label"]))
        story.append(Paragraph(_INSTITUTE_LINES[0], st["title"]))
        story.append(Paragraph(_INSTITUTE_LINES[1], st["sub"]))
        story.append(Paragraph(_INSTITUTE_LINES[2], st["small"]))
        story.append(Paragraph(_INSTITUTE_LINES[3], st["small"]))
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph("PROVISIONAL GRADE SHEET", st["section"]))
        story.append(Paragraph(program, st["sub"]))
        story.append(Paragraph(branch, st["sub"]))
        story.append(Paragraph(f"{semester_roman} Semester {exam_session}", st["sub"]))
        story.append(Spacer(1, 4 * mm))

        info = [
            [Paragraph(f"<b>Name:</b> {rec.get('name','')}", st["label"]),
             Paragraph(f"<b>University Roll No.:</b> {rec.get('roll_no','')}", st["label"])],
            [Paragraph(f"<b>Father's Name:</b> {rec.get('father_name','')}", st["label"]),
             Paragraph(f"<b>University Enrollment No.:</b> {rec.get('enroll_no','')}", st["label"])],
        ]
        t_info = Table(info, colWidths=[90 * mm, 84 * mm])
        t_info.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(t_info)
        story.append(Spacer(1, 3 * mm))

        rows = [["Subject Code", "Subject Name", "Credits", "Grade", "Grade Points"]]
        total_credits = 0
        total_gp = 0.0
        for s in rec.get("subjects", []):
            name_para = Paragraph(s["name"], st["back_subject"] if s.get("back") else st["subject"])
            rows.append([s["code"], name_para, s["credits"], s["grade"], s.get("grade_points", "")])
            try:
                total_credits += int(s["credits"])
            except Exception:
                pass
            try:
                total_gp += float(s["grade_points"]) * int(s["credits"])
            except Exception:
                pass
        rows.append(["Total Credits / Total Grade Points", "", str(total_credits), "-", f"{total_gp:.0f}"])
        rows.append([f"Earned Credits: {rec.get('earned_credits','')}", "",
                      f"SGPA: {rec.get('sgpa','')}", f"CGPA: {rec.get('cgpa','')}", ""])
        rows.append([f"Result: {rec.get('result','')}", "",
                      f"Remark: {rec.get('remark','')}", "", ""])
        col_widths = [25 * mm, 80 * mm, 20 * mm, 20 * mm, 29 * mm]
        t = Table(rows, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("GRID", (0, 0), (-1, -4), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1e1b4b")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("ALIGN", (2, 1), (-1, -4), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("SPAN", (0, -3), (1, -3)),
            ("BACKGROUND", (0, -3), (-1, -3), colors.HexColor("#f5f5f4")),
            ("SPAN", (0, -2), (1, -2)),
            ("SPAN", (3, -2), (4, -2)),
            ("FONTNAME", (0, -2), (-1, -2), "Helvetica-Bold"),
            ("SPAN", (0, -1), (1, -1)),
            ("SPAN", (2, -1), (4, -1)),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ]))
        story.append(t)

        story.append(Spacer(1, 12 * mm))
        story.append(Table([["Prepared by", "", "Checked by", "", "Examination Controller"]],
                            colWidths=[35 * mm] * 5,
                            style=TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                                               ("FONTSIZE", (0, 0), (-1, -1), 9)])))
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
    if "Bachelor" in program:
        return "B. TECH"
    if "Master of Computer" in program:
        return "MCA"
    if "Master" in program:
        return "M. TECH"
    return "PROG"
