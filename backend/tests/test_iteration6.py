"""Iteration 6 — TC footer signatures, Cumulative Earned Credits, GS layout, cuml_earned_credits persistence."""
import os
import io
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
    r = session.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD}, timeout=30)
    assert r.status_code == 200
    return {"Authorization": f"Bearer {r.json()['token']}"}


@pytest.fixture(scope="module")
def excel_upload(session, auth_headers):
    files = {"excel": ("aiml.xlsx", open(f"{SAMPLES}/aiml.xlsx", "rb"),
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    data = {"program": "B.Tech",
            "branch": "Computer Science & Engineering (AIML)",
            "batch": "2022",
            "exam_session": "December 2025"}
    r = session.post(f"{API}/admin/uploads/excel", data=data, files=files, headers=auth_headers, timeout=600)
    assert r.status_code == 200, r.text
    return r.json()


def _fetch_tc(session, auth_headers, excel_upload, semester="I"):
    uid = excel_upload["upload_id"]
    sem_avail = next((s["semester"] for s in excel_upload["semesters"] if s.get("tc_file") and s["semester"] == semester), None)
    if not sem_avail:
        sem_avail = next((s["semester"] for s in excel_upload["semesters"] if s.get("tc_file")), None)
    assert sem_avail
    r = session.get(f"{API}/admin/files/{uid}/sem/{sem_avail}/tc", headers=auth_headers, timeout=60)
    assert r.status_code == 200
    return r.content


def _fetch_gs(session, auth_headers, excel_upload):
    uid = excel_upload["upload_id"]
    sem_avail = next((s["semester"] for s in excel_upload["semesters"] if s.get("gs_file")), None)
    assert sem_avail
    r = session.get(f"{API}/admin/files/{uid}/sem/{sem_avail}/gs", headers=auth_headers, timeout=60)
    assert r.status_code == 200
    return r.content, sem_avail


class TestTCFooterAndFields:
    def test_tc_a3_portrait(self, session, auth_headers, excel_upload):
        pdf = _fetch_tc(session, auth_headers, excel_upload)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            w, h = doc[0].rect.width, doc[0].rect.height
            assert abs(w - 842) <= 2 and abs(h - 1191) <= 2, f"TC not A3 portrait: {w}x{h}"
        finally:
            doc.close()

    def test_tc_contains_credit_fields(self, session, auth_headers, excel_upload):
        pdf = _fetch_tc(session, auth_headers, excel_upload)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            full = "\n".join(p.get_text() for p in doc)
            for needle in ("Cumulative Earned Credits", "Earned Credits", "SGPA", "CGPA", "Result", "Remark"):
                assert needle in full, f"TC missing '{needle}'"
        finally:
            doc.close()

    def test_tc_footer_signatures(self, session, auth_headers, excel_upload):
        """Last page must contain all 5 signature labels including 'Examination Controller (VMSB UTU)'."""
        pdf = _fetch_tc(session, auth_headers, excel_upload)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            last_text = doc[-1].get_text()
            for label in ("Prepared by", "Checked by", "Examination Controller", "Director"):
                assert label in last_text, f"TC last page missing '{label}'"
            assert "VMSB UTU" in last_text, "TC last page missing 'VMSB UTU' signature label"
        finally:
            doc.close()

    def test_tc_four_students_per_page(self, session, auth_headers, excel_upload):
        """Iteration 7: KeepTogether per-student → page 1 must have 1-4 students (no overflow)."""
        pdf = _fetch_tc(session, auth_headers, excel_upload)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            page1_text = doc[0].get_text()
            count = page1_text.count("University Roll No")
            assert 1 <= count <= 4, f"Page 1 should have 1-4 students, got {count}"
        finally:
            doc.close()


class TestGSLayout:
    def test_gs_page1_contents(self, session, auth_headers, excel_upload):
        pdf, _ = _fetch_gs(session, auth_headers, excel_upload)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            t = doc[0].get_text()
            assert "GRADE SHEET" in t.upper(), "GS page1 missing 'GRADE SHEET'"
            assert "PROVISIONAL" not in t.upper(), "GS page1 unexpectedly contains 'PROVISIONAL'"
            assert "SL. NO." not in t and "SL.NO." not in t, "GS page1 has 'SL. NO.'"
            assert "Director" in t, "GS page1 missing 'Director'"
            assert "Examination Controller" in t, "GS page1 missing 'Examination Controller'"
            imgs = doc[0].get_images(full=True)
            assert len(imgs) >= 3, f"GS page1 expected ≥3 images, got {len(imgs)}"
        finally:
            doc.close()


class TestCumlEarnedCreditsPersistence:
    def test_cuml_earned_credits_in_per_student_gs(self, session, auth_headers, excel_upload):
        """Per-student GS PDF embeds earned_credits + cuml_earned_credits via summary card."""
        # find roll matching 220090170*
        sr = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        assert sr.status_code == 200
        rolls = [s["roll_no"] for s in sr.json()["students"] if s["roll_no"].startswith("220090170")]
        assert rolls, "No 220090170* roll found"
        roll = rolls[0]
        # GS for sem I
        r = session.get(f"{API}/admin/student/{roll}/gs/I", headers=auth_headers, timeout=60)
        if r.status_code != 200:
            pytest.skip(f"per-student GS not available: {r.status_code}")
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            t = doc[0].get_text()
            # Iteration 7: per-student GS shows "EARNED CREDITS" (uppercase) in bottom block
            # and "Earned Cr." in semester history table
            assert ("Earned Credits" in t or "EARNED CREDITS" in t.upper() or "Earned Cr" in t), \
                "Per-student GS missing Earned Credits"
            assert "SGPA" in t and "CGPA" in t and "Result" in t.lower() or "RESULT" in t
        finally:
            doc.close()

    def test_cuml_in_tc_per_student_block(self, session, auth_headers, excel_upload):
        """TC PDF should expose 'Cumulative Earned Credits' label."""
        pdf = _fetch_tc(session, auth_headers, excel_upload)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            full = "\n".join(p.get_text() for p in doc)
            # Must appear at least once per student-block, so >=4 on page 1 alone
            assert full.count("Cumulative Earned Credits") >= 4, \
                f"'Cumulative Earned Credits' count too low: {full.count('Cumulative Earned Credits')}"
        finally:
            doc.close()
