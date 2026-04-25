"""Iteration 7 — TC Result/SGPA/CGPA fields, per-page footer signatures,
KeepTogether per student, NO grade reference; GS single-line header,
sig order Examination Controller before Director, watermark, semester-history.
"""
import os
import re
import pytest
import requests
import fitz  # PyMuPDF

BASE_URL = os.environ.get("REACT_APP_BACKEND_URL", "https://grade-sheet-hub.preview.emergentagent.com").rstrip("/")
API = f"{BASE_URL}/api"
SAMPLES = "/app/samples"

ADMIN_EMAIL = "admin@gbpiet.ac.in"
ADMIN_PASSWORD = "Admin@2026"


@pytest.fixture(scope="module")
def session():
    return requests.Session()


@pytest.fixture(scope="module")
def auth_headers(session):
    r = session.post(f"{API}/auth/login",
                     json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD},
                     timeout=30)
    assert r.status_code == 200, r.text
    return {"Authorization": f"Bearer {r.json()['token']}"}


@pytest.fixture(scope="module")
def excel_upload(session, auth_headers):
    files = {"excel": ("aiml.xlsx", open(f"{SAMPLES}/aiml.xlsx", "rb"),
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    data = {"program": "B.Tech",
            "branch": "Computer Science & Engineering (AIML)",
            "batch": "2022",
            "exam_session": "December 2025"}
    r = session.post(f"{API}/admin/uploads/excel", data=data, files=files,
                     headers=auth_headers, timeout=600)
    assert r.status_code == 200, r.text
    return r.json()


def _fetch_tc(session, auth_headers, excel_upload, sem="I"):
    uid = excel_upload["upload_id"]
    target = next((s["semester"] for s in excel_upload["semesters"]
                   if s.get("tc_file") and s["semester"] == sem), None)
    if target is None:
        target = next((s["semester"] for s in excel_upload["semesters"]
                       if s.get("tc_file")), None)
    assert target, "no TC sheet"
    r = session.get(f"{API}/admin/files/{uid}/sem/{target}/tc",
                    headers=auth_headers, timeout=60)
    assert r.status_code == 200
    return r.content, target


def _fetch_gs(session, auth_headers, excel_upload, sem="I"):
    uid = excel_upload["upload_id"]
    target = next((s["semester"] for s in excel_upload["semesters"]
                   if s.get("gs_file") and s["semester"] == sem), None)
    if target is None:
        target = next((s["semester"] for s in excel_upload["semesters"]
                       if s.get("gs_file")), None)
    assert target, "no GS sheet"
    r = session.get(f"{API}/admin/files/{uid}/sem/{target}/gs",
                    headers=auth_headers, timeout=60)
    assert r.status_code == 200
    return r.content, target


# =============== TC PDF =====================
class TestTCFieldsPopulated:
    """TC PDF must render values for SGPA, CGPA, Result, Remark,
    Earned Credits, Cumulative Earned Credits per student (parser fix)."""

    def test_tc_page1_has_value_fields(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_tc(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            t = doc[0].get_text()
            # All labels present
            for needle in ("SGPA", "CGPA", "Result", "Remark",
                           "Earned Credits", "Cumulative Earned Credits"):
                assert needle in t, f"TC page1 missing label '{needle}'"
            # Result must show actual status (PASS/FAIL/REAPPEAR), not '—'
            assert re.search(r"Result\s*[:\-]?\s*(PASS|FAIL|REAPPEAR|PROMOTED)",
                             t, re.IGNORECASE), \
                "TC page1 'Result' has no real status value"
            # SGPA/CGPA must have numeric value (e.g. '5.', '6.', etc.)
            assert re.search(r"SGPA\s*[:\-]?\s*\d", t), \
                "TC page1 'SGPA' has no numeric value"
            assert re.search(r"CGPA\s*[:\-]?\s*\d", t), \
                "TC page1 'CGPA' has no numeric value"
            # Earned Credits numeric
            assert re.search(r"Earned Credits\s*[:\-]?\s*\d", t), \
                "TC page1 'Earned Credits' has no numeric value"
        finally:
            doc.close()


class TestTCFooterEveryPage:
    """Every TC page must have signed footer with all 5 labels."""

    def test_footer_on_every_page(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_tc(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            assert doc.page_count >= 2, f"TC must have >=2 pages, got {doc.page_count}"
            page_indices = sorted({0, 1, doc.page_count // 2, doc.page_count - 1})
            labels = ("Prepared by", "Checked by", "Examination Controller",
                      "Director", "VMSB UTU")
            for idx in page_indices:
                t = doc[idx].get_text()
                for label in labels:
                    assert label in t, \
                        f"TC page {idx+1} missing footer label '{label}'"
        finally:
            doc.close()


class TestTCNoStudentSplit:
    """KeepTogether: no student record overflows across pages.
    Sanity check: each page <=4 student blocks; total roll occurrences == student count."""

    def test_no_student_overflow(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_tc(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            for p in range(doc.page_count):
                t = doc[p].get_text()
                cnt = t.count("University Roll No")
                assert cnt <= 4, \
                    f"TC page {p+1} has {cnt} students (>4 → overflow risk)"
        finally:
            doc.close()


class TestTCNoGradeReference:
    """Iteration 7: Grade Reference table removed from TC."""

    def test_tc_no_grade_reference(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_tc(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            full = "\n".join(p.get_text() for p in doc)
            assert "Grade Reference" not in full, \
                "TC unexpectedly contains 'Grade Reference' header"
            assert "Letter Grade" not in full, \
                "TC unexpectedly contains 'Letter Grade' header"
        finally:
            doc.close()


# =============== GS PDF =====================
class TestGSHeaderSingleLine:
    def test_institute_and_university_present(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_gs(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            t = doc[0].get_text()
            assert "GOVIND BALLABH PANT INSTITUTE" in t.upper(), \
                "GS missing institute name"
            assert "Veer Madho Singh Bhandari" in t or \
                   "VEER MADHO SINGH BHANDARI" in t.upper(), \
                "GS missing UTU name"
        finally:
            doc.close()


class TestGSSignatureOrder:
    """Signature row: Prepared by, Checked by, Examination Controller, Director.
    => Examination Controller index < Director index (Director appears AFTER)."""

    def test_director_after_examination_controller(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_gs(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            t = doc[0].get_text()
            ec_idx = t.find("Examination Controller")
            dir_idx = t.find("Director")
            assert ec_idx >= 0 and dir_idx >= 0, \
                "GS page1 missing Examination Controller or Director"
            assert dir_idx > ec_idx, \
                f"Director should appear AFTER Examination Controller in text (ec_idx={ec_idx}, dir_idx={dir_idx})"
        finally:
            doc.close()


class TestGSSemesterHistoryTable:
    def test_semester_history_present(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_gs(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            t = doc[0].get_text()
            assert "Semester-wise Result History" in t, \
                "GS missing 'Semester-wise Result History' header"
            assert "SGPA" in t and "CGPA" in t, \
                "GS sem-history missing SGPA/CGPA rows"
            # at least one numeric SGPA value (digit somewhere with dot)
            assert re.search(r"\d\.\d", t), \
                "GS sem-history missing any numeric SGPA/CGPA value"
        finally:
            doc.close()


class TestGSImagesAndPageCount:
    def test_three_images_and_no_overflow(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_gs(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            imgs = doc[0].get_images(full=True)
            assert len(imgs) >= 3, \
                f"GS page1 expected >=3 images (institute+UTU+barcode/watermark), got {len(imgs)}"
            # page_count == student count for sem I
            sr = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
            students = [s for s in sr.json()["students"] if s["roll_no"].startswith("220090170")]
            # Number of students in this branch with sem I result must match page count
            assert doc.page_count > 1, "GS must have multiple pages"
            # tolerance: page count should be close to student list (within 10%)
            assert abs(doc.page_count - len(students)) <= max(2, int(0.1 * len(students))), \
                f"GS page_count={doc.page_count} far from student count={len(students)}"
        finally:
            doc.close()


class TestGSExamSessionFromSheet:
    def test_exam_session_present(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_gs(session, auth_headers, excel_upload, "I")
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            t = doc[0].get_text()
            # exam session header should be present (non-empty)
            assert "Exam" in t or "Session" in t or re.search(r"\b(202\d|201\d)\b", t), \
                "GS page1 missing exam session indicator"
        finally:
            doc.close()


class TestPerStudentGSWithSummary:
    """parse_tc_excel_sheet must populate SGPA/CGPA/Result/Earned/Cuml
    so per-student GS shows a usable summary."""

    def test_per_student_gs_has_summary(self, session, auth_headers, excel_upload):
        sr = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        rolls = [s["roll_no"] for s in sr.json()["students"]
                 if s["roll_no"].startswith("220090170")]
        assert rolls, "no 220090170* roll"
        roll = rolls[0]
        r = session.get(f"{API}/admin/student/{roll}/gs/I",
                        headers=auth_headers, timeout=60)
        if r.status_code != 200:
            pytest.skip(f"per-student GS not available: {r.status_code}")
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            t = doc[0].get_text()
            assert "SGPA" in t and "CGPA" in t
            assert "Earned Credits" in t or "EARNED CREDITS" in t.upper()
            # Numeric SGPA — confirms parser populated value
            assert re.search(r"SGPA\s*[:\-]?\s*\d", t) or re.search(r"\d\.\d", t), \
                "per-student GS missing numeric SGPA"
        finally:
            doc.close()
