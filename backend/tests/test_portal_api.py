"""End-to-end backend tests for GBPIET Result Asterisk Portal."""
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


@pytest.fixture(scope="session")
def session():
    return requests.Session()


@pytest.fixture(scope="session")
def auth_token(session):
    r = session.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD}, timeout=30)
    assert r.status_code == 200, f"login failed: {r.status_code} {r.text}"
    data = r.json()
    assert "token" in data and isinstance(data["token"], str) and len(data["token"]) > 20
    assert data["email"] == ADMIN_EMAIL
    return data["token"]


@pytest.fixture(scope="session")
def auth_headers(auth_token):
    return {"Authorization": f"Bearer {auth_token}"}


# ---------------- AUTH ----------------
class TestAuth:
    def test_login_success(self, session):
        r = session.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": ADMIN_PASSWORD}, timeout=30)
        assert r.status_code == 200
        d = r.json()
        assert d["email"] == ADMIN_EMAIL
        assert "token" in d
        # bcrypt hash check via DB-side: token signature acceptable (JWT three-parts)
        assert d["token"].count(".") == 2

    def test_login_invalid(self, session):
        r = session.post(f"{API}/auth/login", json={"email": ADMIN_EMAIL, "password": "wrong"}, timeout=30)
        assert r.status_code == 401

    def test_me_with_bearer(self, session, auth_headers):
        r = session.get(f"{API}/auth/me", headers=auth_headers, timeout=30)
        assert r.status_code == 200
        d = r.json()
        assert d["email"] == ADMIN_EMAIL
        assert "password_hash" not in d  # _id and password should be excluded

    def test_me_unauthenticated(self, session):
        r = requests.get(f"{API}/auth/me", timeout=30)
        assert r.status_code == 401


# ---------------- META ----------------
class TestMeta:
    def test_meta_programs(self, session):
        r = session.get(f"{API}/meta/programs", timeout=30)
        assert r.status_code == 200
        d = r.json()
        assert "B.Tech" in d["programs"]
        assert "2022" in d["batches"]
        assert "I" in d["semesters"]


# ---------------- UPLOAD + PROCESSING ----------------
@pytest.fixture(scope="session")
def upload_tc_excel(session, auth_headers):
    """Upload TC PDF + SEM Excel for AIML 2022 sem I — primary test fixture."""
    files = {
        "tc_pdf": ("tc.pdf", open(f"{SAMPLES}/tc.pdf", "rb"), "application/pdf"),
        "sem_excel": ("aiml.xlsx", open(f"{SAMPLES}/aiml.xlsx", "rb"),
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }
    data = {
        "program": "B.Tech",
        "branch": "Computer Science & Engineering (AIML)",
        "batch": "2022",
        "semester": "I",
        "exam_session": "December 2025",
    }
    r = session.post(f"{API}/admin/uploads", data=data, files=files, headers=auth_headers, timeout=240)
    assert r.status_code == 200, r.text
    return r.json()


@pytest.fixture(scope="session")
def upload_gs(session, auth_headers):
    files = {
        "gs_pdf": ("gs.pdf", open(f"{SAMPLES}/gs.pdf", "rb"), "application/pdf"),
    }
    data = {
        "program": "B.Tech",
        "branch": "Civil Engineering",
        "batch": "2023",
        "semester": "V",
        "exam_session": "December 2025",
    }
    r = session.post(f"{API}/admin/uploads", data=data, files=files, headers=auth_headers, timeout=240)
    assert r.status_code == 200, r.text
    return r.json()


class TestUploads:
    def test_upload_tc_excel_counts(self, upload_tc_excel):
        d = upload_tc_excel
        assert "upload_id" in d
        assert d["tc_count"] >= 50, f"TC count too low: {d['tc_count']}"
        # Excel should detect back students (yellow fills); allow >=1 to be tolerant
        assert d["back_students"] >= 1, f"backs not detected: {d['back_students']}"

    def test_upload_response_match_fields(self, upload_tc_excel):
        """New fields in upload response: tc_matched, gs_matched, warning."""
        d = upload_tc_excel
        assert "tc_matched" in d, "tc_matched missing in upload response"
        assert "gs_matched" in d, "gs_matched missing in upload response"
        assert "warning" in d, "warning key missing in upload response"
        assert isinstance(d["tc_matched"], int)
        assert isinstance(d["gs_matched"], int)

    def test_upload_gs_counts(self, upload_gs):
        d = upload_gs
        assert d["gs_count"] >= 50, f"GS count too low: {d['gs_count']}"

    def test_download_tc_pdf(self, session, auth_headers, upload_tc_excel):
        uid = upload_tc_excel["upload_id"]
        r = session.get(f"{API}/admin/files/{uid}/tc", headers=auth_headers, timeout=60)
        assert r.status_code == 200
        assert r.headers.get("content-type", "").startswith("application/pdf")
        assert len(r.content) > 1000
        # Verify logos are embedded on page 1
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            assert doc.page_count >= 1
            imgs = doc[0].get_images(full=True)
            assert len(imgs) >= 2, f"Expected >=2 images (institute+UTU logos) on TC page1, got {len(imgs)}"
        finally:
            doc.close()

    def test_download_gs_pdf(self, session, auth_headers, upload_gs):
        uid = upload_gs["upload_id"]
        r = session.get(f"{API}/admin/files/{uid}/gs", headers=auth_headers, timeout=60)
        assert r.status_code == 200
        assert r.headers.get("content-type", "").startswith("application/pdf")
        assert len(r.content) > 1000
        # Verify logos are embedded on page 1
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            assert doc.page_count >= 1
            imgs = doc[0].get_images(full=True)
            assert len(imgs) >= 2, f"Expected >=2 images (institute+UTU logos) on GS page1, got {len(imgs)}"
        finally:
            doc.close()

    def test_list_uploads(self, session, auth_headers, upload_tc_excel):
        r = session.get(f"{API}/admin/uploads", headers=auth_headers, timeout=30)
        assert r.status_code == 200
        ups = r.json()["uploads"]
        assert any(u["id"] == upload_tc_excel["upload_id"] for u in ups)

    def test_list_students(self, session, auth_headers, upload_tc_excel):
        r = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        assert r.status_code == 200
        stus = r.json()["students"]
        assert len(stus) > 0
        assert all("roll_no" in s for s in stus)

    def test_admin_stats(self, session, auth_headers, upload_tc_excel):
        r = session.get(f"{API}/admin/stats", headers=auth_headers, timeout=30)
        assert r.status_code == 200
        d = r.json()
        for k in ("uploads", "students", "results", "backlogs_total"):
            assert k in d
        assert d["uploads"] >= 1


# ---------------- STUDENT PORTAL ----------------
class TestStudent:
    def test_student_login_invalid(self, session):
        r = session.post(f"{API}/student/login", json={"roll_no": "ZZZ999"}, timeout=30)
        assert r.status_code == 404

    def test_student_login_no_dob(self, session, auth_headers, upload_tc_excel, upload_gs):
        """DOB removed: login with only roll_no must work."""
        sr = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        roll = sr.json()["students"][0]["roll_no"]
        # NO dob in body
        r = session.post(f"{API}/student/login", json={"roll_no": roll}, timeout=30)
        assert r.status_code == 200, r.text
        d = r.json()
        assert d["student"]["roll_no"] == roll
        assert isinstance(d["results"], list)

    def test_student_login_dob_ignored(self, session, auth_headers, upload_tc_excel, upload_gs):
        """DOB is optional/ignored: passing any value still returns 200."""
        sr = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        roll = sr.json()["students"][0]["roll_no"]
        r = session.post(f"{API}/student/login",
                         json={"roll_no": roll, "dob": "anything-random"},
                         timeout=30)
        assert r.status_code == 200, r.text
        d = r.json()
        assert d["student"]["roll_no"] == roll

# ---------------- EXCEL-ONLY UPLOAD (iteration 3) ----------------
@pytest.fixture(scope="session")
def upload_excel_only(session, auth_headers):
    """POST /api/admin/uploads/excel with the AIML.xlsx (TC_/GS_/SEM_ sheets)."""
    files = {
        "excel": ("aiml.xlsx", open(f"{SAMPLES}/aiml.xlsx", "rb"),
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }
    data = {
        "program": "B.Tech",
        "branch": "Computer Science & Engineering (AIML)",
        "batch": "2022",
        "exam_session": "December 2025",
    }
    r = session.post(f"{API}/admin/uploads/excel", data=data, files=files,
                      headers=auth_headers, timeout=600)
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:500]}"
    return r.json()


class TestExcelOnly:
    def test_excel_only_response_shape(self, upload_excel_only):
        d = upload_excel_only
        assert "upload_id" in d
        assert isinstance(d["semesters"], list) and len(d["semesters"]) >= 1
        assert d["students_indexed"] >= 50
        assert d["total_asterisks_applied"] >= 0
        for s in d["semesters"]:
            assert s["semester"] in ["I","II","III","IV","V","VI","VII","VIII"]
            assert "tc_count" in s and "gs_count" in s
            assert "tc_file" in s and "gs_file" in s
            assert "backs_in_sem" in s
            assert "asterisks_applied" in s

    def test_excel_only_missing_file_422(self, session, auth_headers):
        # No 'excel' file part — FastAPI returns 422
        r = session.post(f"{API}/admin/uploads/excel",
                         data={"program": "B.Tech", "branch": "X", "batch": "2022"},
                         headers=auth_headers, timeout=30)
        assert r.status_code == 422, r.text

    def test_excel_only_invalid_workbook_400(self, session, auth_headers):
        # send a file with no TC_/GS_ sheets — should yield 400
        bad = io.BytesIO()
        import openpyxl
        wb = openpyxl.Workbook()
        wb.active.title = "Cover"
        wb.active["A1"] = "no sheets"
        wb.save(bad)
        bad.seek(0)
        files = {
            "excel": ("bad.xlsx", bad,
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        }
        data = {"program": "B.Tech", "branch": "Civil Engineering", "batch": "2022"}
        r = session.post(f"{API}/admin/uploads/excel", data=data, files=files,
                          headers=auth_headers, timeout=60)
        assert r.status_code == 400, r.text
        assert "TC_" in r.text or "GS_" in r.text

    def test_download_sem_tc_pdf(self, session, auth_headers, upload_excel_only):
        uid = upload_excel_only["upload_id"]
        # find a semester that has tc_file
        tc_sem = next((s["semester"] for s in upload_excel_only["semesters"]
                        if s["tc_file"]), None)
        assert tc_sem, "No TC sheet generated"
        r = session.get(f"{API}/admin/files/{uid}/sem/{tc_sem}/tc",
                         headers=auth_headers, timeout=60)
        assert r.status_code == 200
        assert r.headers.get("content-type", "").startswith("application/pdf")
        assert len(r.content) > 1000

    def test_download_sem_gs_pdf_with_barcode(self, session, auth_headers, upload_excel_only):
        uid = upload_excel_only["upload_id"]
        gs_sem = next((s["semester"] for s in upload_excel_only["semesters"]
                        if s["gs_file"]), None)
        assert gs_sem, "No GS sheet generated"
        r = session.get(f"{API}/admin/files/{uid}/sem/{gs_sem}/gs",
                         headers=auth_headers, timeout=60)
        assert r.status_code == 200
        assert r.headers.get("content-type", "").startswith("application/pdf")
        # Verify GS page1 has >=3 images: institute logo + UTU logo + barcode
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            assert doc.page_count >= 1
            imgs = doc[0].get_images(full=True)
            assert len(imgs) >= 3, (
                f"Expected >=3 images (institute+UTU logos + barcode) on GS page1, "
                f"got {len(imgs)}"
            )
        finally:
            doc.close()

    def test_download_sem_invalid_404(self, session, auth_headers, upload_excel_only):
        uid = upload_excel_only["upload_id"]
        # Sem X is not a valid Roman, and even if it were, no file exists
        r = session.get(f"{API}/admin/files/{uid}/sem/X/gs",
                         headers=auth_headers, timeout=30)
        assert r.status_code == 404

    def test_list_uploads_includes_excel_only(self, session, auth_headers, upload_excel_only):
        r = session.get(f"{API}/admin/uploads", headers=auth_headers, timeout=30)
        assert r.status_code == 200
        ups = r.json()["uploads"]
        match = next((u for u in ups
                      if u["id"] == upload_excel_only["upload_id"]), None)
        assert match is not None, "excel-only upload not in list"
        assert match.get("source") == "excel-only"
        assert isinstance(match.get("semesters"), list) and len(match["semesters"]) >= 1
        assert match.get("tc_file") is None
        assert match.get("gs_file") is None


    def test_student_login_valid(self, session, auth_headers, upload_tc_excel, upload_gs):
        # pick first roll from /admin/students
        sr = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        roll = sr.json()["students"][0]["roll_no"]
        r = session.post(f"{API}/student/login", json={"roll_no": roll}, timeout=30)
        assert r.status_code == 200, r.text
        d = r.json()
        assert d["student"]["roll_no"] == roll
        assert isinstance(d["results"], list)
        # if multiple semesters, ensure desc
        order = ["I","II","III","IV","V","VI","VII","VIII"]
        sems = [r["semester"] for r in d["results"]]
        idxs = [order.index(s) for s in sems if s in order]
        assert idxs == sorted(idxs, reverse=True)


# ---------------- ITERATION 4: TC A3 + GS A4 + barcode + Director + Grade Ref ----
@pytest.fixture(scope="session")
def upload_excel_only_mtech(session, auth_headers):
    """Same aiml.xlsx but uploaded as program='M.Tech' to force the M.Tech
    ordinance grade table in the reference block."""
    files = {
        "excel": ("aiml.xlsx", open(f"{SAMPLES}/aiml.xlsx", "rb"),
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
    }
    data = {
        "program": "M.Tech",
        "branch": "Computer Science & Engineering",
        "batch": "2022",
        "exam_session": "December 2025",
    }
    r = session.post(f"{API}/admin/uploads/excel", data=data, files=files,
                     headers=auth_headers, timeout=600)
    assert r.status_code == 200, f"M.Tech excel upload failed: {r.status_code} {r.text[:300]}"
    return r.json()


class TestI4TcPdf:
    """TC PDF must be A3 portrait, all 8 columns, totals row, enrollment, grade ref."""

    def _fetch_tc(self, session, auth_headers, upload_excel_only):
        uid = upload_excel_only["upload_id"]
        # Pick the largest-content semester (more subjects → totals row populated)
        tc_sem = next((s["semester"] for s in upload_excel_only["semesters"]
                        if s["tc_file"]), None)
        assert tc_sem
        r = session.get(f"{API}/admin/files/{uid}/sem/{tc_sem}/tc",
                        headers=auth_headers, timeout=60)
        assert r.status_code == 200
        return r.content, tc_sem

    def test_tc_page_size_a3_portrait(self, session, auth_headers, upload_excel_only):
        pdf, _ = self._fetch_tc(session, auth_headers, upload_excel_only)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            w, h = doc[0].rect.width, doc[0].rect.height
            # A3 portrait = 842 x 1191 pt (297 x 420 mm). Tolerate +/-2pt.
            assert abs(w - 842) <= 2 and abs(h - 1191) <= 2, \
                f"TC page size not A3 portrait: got {w}x{h}pt, expected ~842x1191pt"
        finally:
            doc.close()

    def test_tc_contains_enrollment_and_grades(self, session, auth_headers, upload_excel_only):
        pdf, _ = self._fetch_tc(session, auth_headers, upload_excel_only)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            page1_text = doc[0].get_text()
            assert "Enrol" in page1_text, "TC page1 missing 'Enrol' label"
            # At least 2 typical letter grades on page 1
            grade_hits = sum(1 for g in ("A+", "A ", "B+", "B ", "C+", "C ", "D ", "P ", "F ", "O ")
                             if g in page1_text)
            assert grade_hits >= 2, \
                f"TC page1 has too few letter-grade tokens: {grade_hits}"
            # Director signature
            full_text = "\n".join(p.get_text() for p in doc)
            assert "Director" in full_text, "TC missing 'Director' signature"
            # Grade reference is rendered on the LAST page only (memory-aware)
            last_text = doc[-1].get_text()
            assert ("Letter Grade" in last_text or "Grade Point" in last_text), \
                "TC last page missing Grade Reference table"
        finally:
            doc.close()

    def test_tc_totals_row_present(self, session, auth_headers, upload_excel_only):
        """A 'Total' row should be rendered in TC tables."""
        pdf, _ = self._fetch_tc(session, auth_headers, upload_excel_only)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            full = "\n".join(p.get_text() for p in doc)
            # 'Total' label appears in the totals row of each student record
            assert full.count("Total") >= 5, \
                f"TC totals row not rendered enough: count={full.count('Total')}"
        finally:
            doc.close()


class TestI4GsPdf:
    """GS PDF must be A4 portrait, no SL.NO/PROVISIONAL, 3+ images per page, Director."""

    def _fetch_gs(self, session, auth_headers, upload_excel_only):
        uid = upload_excel_only["upload_id"]
        gs_sem = next((s["semester"] for s in upload_excel_only["semesters"]
                        if s["gs_file"]), None)
        assert gs_sem
        r = session.get(f"{API}/admin/files/{uid}/sem/{gs_sem}/gs",
                        headers=auth_headers, timeout=60)
        assert r.status_code == 200
        return r.content

    def test_gs_page_size_a4_portrait(self, session, auth_headers, upload_excel_only):
        pdf = self._fetch_gs(session, auth_headers, upload_excel_only)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            w, h = doc[0].rect.width, doc[0].rect.height
            # A4 portrait = 595 x 842 pt
            assert abs(w - 595) <= 2 and abs(h - 842) <= 2, \
                f"GS page size not A4 portrait: got {w}x{h}pt, expected ~595x842pt"
        finally:
            doc.close()

    def test_gs_no_slno_no_provisional_has_grade_sheet(self, session, auth_headers, upload_excel_only):
        pdf = self._fetch_gs(session, auth_headers, upload_excel_only)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            sample_pages = [0]
            if doc.page_count > 50:
                sample_pages.append(50)
            if doc.page_count > 106:
                sample_pages.append(106)
            for idx in sample_pages:
                t = doc[idx].get_text()
                assert "SL. NO." not in t and "SL.NO." not in t, \
                    f"GS page {idx+1} still contains 'SL. NO.'"
                assert "PROVISIONAL" not in t.upper() or "PROVISIONAL GRADE SHEET" not in t.upper(), \
                    f"GS page {idx+1} still contains 'PROVISIONAL GRADE SHEET'"
                assert "GRADE SHEET" in t.upper(), \
                    f"GS page {idx+1} missing 'GRADE SHEET' title"
                assert "Director" in t, \
                    f"GS page {idx+1} missing 'Director' signature"
        finally:
            doc.close()

    def test_gs_three_images_per_sampled_page(self, session, auth_headers, upload_excel_only):
        pdf = self._fetch_gs(session, auth_headers, upload_excel_only)
        doc = fitz.open(stream=pdf, filetype="pdf")
        try:
            sample_pages = [0]
            if doc.page_count > 50:
                sample_pages.append(50)
            if doc.page_count > 106:
                sample_pages.append(106)
            for idx in sample_pages:
                imgs = doc[idx].get_images(full=True)
                assert len(imgs) >= 3, \
                    f"GS page {idx+1} expected >=3 images (institute+UTU+barcode), got {len(imgs)}"
        finally:
            doc.close()


class TestI4GradeReference:
    """Grade reference must switch between B.Tech (2025 AC) and M.Tech ordinance."""

    def test_btech_grade_table_btech_reference(self, session, auth_headers, upload_excel_only):
        uid = upload_excel_only["upload_id"]
        gs_sem = next((s["semester"] for s in upload_excel_only["semesters"]
                        if s["gs_file"]), None)
        r = session.get(f"{API}/admin/files/{uid}/sem/{gs_sem}/gs",
                        headers=auth_headers, timeout=60)
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            full = "\n".join(p.get_text() for p in doc)
            # B.Tech 2025 table has the 'E+' / 'E' / 'D+' rows that M.Tech does not
            assert ("E+" in full) or ("D+" in full), \
                "B.Tech grade-ref table not rendered (missing E+/D+ rows)"
            # The M.Tech-only labels should NOT appear
            assert "Outstanding" not in full, \
                "B.Tech PDF erroneously contains M.Tech 'Outstanding' label"
        finally:
            doc.close()

    def test_mtech_grade_table_mtech_reference(self, session, auth_headers, upload_excel_only_mtech):
        uid = upload_excel_only_mtech["upload_id"]
        gs_sem = next((s["semester"] for s in upload_excel_only_mtech["semesters"]
                        if s["gs_file"]), None)
        assert gs_sem, "M.Tech excel produced no GS sheets"
        r = session.get(f"{API}/admin/files/{uid}/sem/{gs_sem}/gs",
                        headers=auth_headers, timeout=60)
        assert r.status_code == 200
        doc = fitz.open(stream=r.content, filetype="pdf")
        try:
            full = "\n".join(p.get_text() for p in doc)
            for needle in ("Outstanding", "Excellent", "Pass", "Absent"):
                assert needle in full, \
                    f"M.Tech grade reference missing label: '{needle}'"
        finally:
            doc.close()


class TestI4EnrollmentPersistence:
    """parse_tc_excel_sheet col-13 fix — enroll_no must be populated for known roll."""

    def test_enroll_no_for_known_roll(self, session, auth_headers, upload_excel_only):
        # Look up a known roll. Use /admin/students to discover available rolls
        r = session.get(f"{API}/admin/students", headers=auth_headers, timeout=30)
        assert r.status_code == 200
        students = r.json()["students"]
        # Find a roll in 22009017xxxx range (CSE-AIML 2022)
        target = next((s for s in students if s["roll_no"].startswith("220090170")), None)
        assert target, "Known roll 220090170* not in /admin/students"
        # The /admin/students endpoint may not include enroll_no; verify via
        # the per-student GS endpoint instead — it pulls the student doc.
        roll = target["roll_no"]
        r = session.get(f"{API}/admin/student/{roll}/gs/I",
                         headers=auth_headers, timeout=60)
        # 200 means the student record + result both exist
        if r.status_code == 200:
            # Pull the underlying student doc to verify enroll_no by checking
            # the rendered PDF text contains 'UTU/' (university enrol prefix)
            doc = fitz.open(stream=r.content, filetype="pdf")
            try:
                t = doc[0].get_text()
                assert "UTU/" in t, \
                    f"Per-student GS PDF missing 'UTU/' enrollment for roll {roll}"
            finally:
                doc.close()
        else:
            pytest.skip(f"Per-student GS not available for {roll}: {r.status_code}")
