"""Result Asterisk Portal — FastAPI backend."""
from dotenv import load_dotenv
from pathlib import Path

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / ".env")

import os
import re
import uuid
import logging
import json
from datetime import datetime, timezone, timedelta
from typing import Dict, List, Optional

import bcrypt
import jwt
from fastapi import (
    APIRouter,
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    Response,
    UploadFile,
    status,
)
from fastapi.responses import StreamingResponse
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
from pydantic import BaseModel, Field

from processor import (
    apply_back_markers,
    apply_non_credit_markers,
    restore_external_from_total,
    generate_gs_pdf,
    generate_tc_pdf,
    parse_gs_pdf,
    parse_sem_excel,
    parse_sem_excel_sheet,
    parse_tc_gs_excel,
    parse_tc_or_gs_named_sheet,
    parse_tc_pdf,
    inspect_workbook_sheets,
    fix_mtech_non_credit_columns,
)

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
log = logging.getLogger("portal")

mongo_url = os.environ["MONGO_URL"]
db_name = os.environ["DB_NAME"]
client = AsyncIOMotorClient(mongo_url)
db = client[db_name]

JWT_SECRET = os.environ["JWT_SECRET"]
JWT_ALG = "HS256"
ACCESS_MIN = 60 * 8  # 8h sessions for the admin
ADMIN_EMAIL = os.environ["ADMIN_EMAIL"]
ADMIN_PASSWORD = os.environ["ADMIN_PASSWORD"]

STORAGE = ROOT_DIR / "storage"
STORAGE.mkdir(exist_ok=True)

PROGRAMS = {
    "B.Tech": [
        "Computer Science & Engineering",
        "Computer Science & Engineering (AIML)",
        "Electronics and Communication Engineering",
        "Electrical Engineering",
        "Biotech",
        "Civil Engineering",
        "Mechanical Engineering",
        "Mechanical Engineering (MFG)",
    ],
    "MCA": ["MCA"],
    "M.Tech": [
        "Computer Science & Engineering",
        "Power System",
        "Geotechnology",
        "Biotechnology",
        "Electronics and Communication Engineering",
        "Production Engineering",
        "Thermal Engineering",
        "Digital Communication",
    ],
}
BATCHES = ["2022", "2023", "2024", "2025"]

app = FastAPI(title="GBPIET Result Asterisk Portal")
api = APIRouter(prefix="/api")


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------
def hash_password(p: str) -> str:
    return bcrypt.hashpw(p.encode(), bcrypt.gensalt()).decode()


def verify_password(p: str, h: str) -> bool:
    try:
        return bcrypt.checkpw(p.encode(), h.encode())
    except Exception:
        return False


def create_token(sub: str, email: str) -> str:
    payload = {
        "sub": sub,
        "email": email,
        "exp": datetime.now(timezone.utc) + timedelta(minutes=ACCESS_MIN),
        "iat": datetime.now(timezone.utc),
        "type": "access",
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALG)


async def get_current_admin(request: Request) -> dict:
    token = request.cookies.get("access_token")
    if not token:
        h = request.headers.get("Authorization", "")
        if h.startswith("Bearer "):
            token = h[7:]
    if not token:
        raise HTTPException(status_code=401, detail="Not authenticated")
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALG])
        if payload.get("type") != "access":
            raise HTTPException(status_code=401, detail="Invalid token type")
        u = await db.users.find_one({"id": payload["sub"]}, {"_id": 0, "password_hash": 0})
        if not u:
            raise HTTPException(status_code=401, detail="User not found")
        return u
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Invalid token")


# ---------------------------------------------------------------------------
# Pydantic models
# ---------------------------------------------------------------------------
class LoginIn(BaseModel):
    email: str
    password: str


class VerifyOtpIn(BaseModel):
    challenge_id: str
    otp: str


class StudentLoginIn(BaseModel):
    roll_no: str
    dob: str = ""  # optional, kept for backwards compatibility


class DobItem(BaseModel):
    roll_no: str
    dob: str


# ---------------------------------------------------------------------------
# OTP / SMTP helpers (admin 2FA)
# ---------------------------------------------------------------------------
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
OTP_RECIPIENT = os.environ.get("OTP_RECIPIENT", "")
OTP_TTL_SECONDS = 600           # 10 minutes
OTP_MAX_ATTEMPTS = 5            # per challenge


def _send_otp_email(otp: str, to_addr: str) -> None:
    """Send the OTP via Gmail SMTP. Raises HTTPException(503) on failure."""
    import smtplib
    from email.message import EmailMessage

    if not (SMTP_USER and SMTP_PASSWORD and to_addr):
        raise HTTPException(503, "SMTP is not configured")

    msg = EmailMessage()
    msg["From"] = f"GBPIET Result Portal <{SMTP_USER}>"
    msg["To"] = to_addr
    msg["Subject"] = "GBPIET Admin Portal — Login OTP"
    msg.set_content(
        "Govind Ballabh Pant Institute of Engineering and Technology\n"
        "Result Portal — Admin Login\n\n"
        f"Your one-time password is: {otp}\n"
        "It is valid for 10 minutes.\n\n"
        "If you did not initiate this login, please ignore this email and "
        "contact the system administrator.\n"
    )
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)
    except Exception as e:
        log.exception("SMTP send failed: %s", e)
        raise HTTPException(503, "Failed to send OTP email — check SMTP "
                                  "credentials / network access.")


def _generate_otp() -> str:
    """6-digit numeric OTP, zero-padded."""
    import secrets
    return f"{secrets.randbelow(1_000_000):06d}"


# ---------------------------------------------------------------------------
# Auth endpoints
# ---------------------------------------------------------------------------
@api.post("/auth/login")
async def login(body: LoginIn):
    """Step 1 of admin 2FA: validate password, mail an OTP, return a
    short-lived challenge_id that the client must redeem with the OTP."""
    email = body.email.strip().lower()
    user = await db.users.find_one({"email": email})
    if not user or not verify_password(body.password, user["password_hash"]):
        raise HTTPException(status_code=401, detail="Invalid credentials")

    otp = _generate_otp()
    challenge_id = str(uuid.uuid4())
    now = datetime.now(timezone.utc)
    await db.otp_challenges.insert_one({
        "challenge_id": challenge_id,
        "user_id": user["id"],
        "email": email,
        "otp_hash": hash_password(otp),
        "expires_at": now + timedelta(seconds=OTP_TTL_SECONDS),
        "attempts": 0,
        "verified": False,
        "created_at": now,
    })
    # Mask the recipient for the response so the OTP is never echoed back.
    masked = OTP_RECIPIENT
    if "@" in masked:
        local, _, domain = masked.partition("@")
        masked = (local[:3] + "•••" + local[-1:] if len(local) > 4 else "••••") + "@" + domain
    _send_otp_email(otp, OTP_RECIPIENT)

    return {
        "otp_required": True,
        "challenge_id": challenge_id,
        "expires_in": OTP_TTL_SECONDS,
        "sent_to": masked,
        "message": "An OTP has been emailed to the authorized recipient.",
    }


@api.post("/auth/verify-otp")
async def verify_otp(body: VerifyOtpIn, response: Response):
    """Step 2 of admin 2FA: redeem the OTP, issue the JWT."""
    chal = await db.otp_challenges.find_one({"challenge_id": body.challenge_id})
    if not chal:
        raise HTTPException(status_code=400, detail="Invalid or expired challenge")
    now = datetime.now(timezone.utc)
    expires_at = chal.get("expires_at")
    if isinstance(expires_at, str):
        try:
            expires_at = datetime.fromisoformat(expires_at)
        except Exception:
            expires_at = now - timedelta(seconds=1)
    # Mongo returns datetimes as naive UTC — coerce to aware before compare.
    if isinstance(expires_at, datetime) and expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
    if chal.get("verified") or (expires_at and now > expires_at):
        await db.otp_challenges.delete_one({"challenge_id": body.challenge_id})
        raise HTTPException(status_code=400, detail="Invalid or expired challenge")
    if chal.get("attempts", 0) >= OTP_MAX_ATTEMPTS:
        await db.otp_challenges.delete_one({"challenge_id": body.challenge_id})
        raise HTTPException(status_code=429, detail="Too many invalid OTP attempts")
    if not verify_password((body.otp or "").strip(), chal["otp_hash"]):
        await db.otp_challenges.update_one(
            {"challenge_id": body.challenge_id},
            {"$inc": {"attempts": 1}},
        )
        raise HTTPException(status_code=401, detail="Invalid OTP")

    # Single-use: delete the challenge and issue the JWT.
    await db.otp_challenges.delete_one({"challenge_id": body.challenge_id})
    user = await db.users.find_one({"id": chal["user_id"]})
    if not user:
        raise HTTPException(status_code=401, detail="User not found")
    token = create_token(user["id"], user["email"])
    response.set_cookie(
        "access_token", token, httponly=True, secure=True, samesite="none",
        max_age=ACCESS_MIN * 60, path="/",
    )
    return {
        "id": user["id"],
        "email": user["email"],
        "name": user.get("name", "Admin"),
        "token": token,
    }


@api.post("/auth/logout")
async def logout(response: Response):
    response.delete_cookie("access_token", path="/")
    return {"ok": True}


@api.get("/auth/me")
async def me(user: dict = Depends(get_current_admin)):
    return user


# ---------------------------------------------------------------------------
# Meta
# ---------------------------------------------------------------------------
@api.get("/meta/programs")
async def meta_programs():
    return {"programs": PROGRAMS, "batches": BATCHES,
            "semesters": ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]}


# ---------------------------------------------------------------------------
# Upload + processing
# ---------------------------------------------------------------------------
@api.post("/admin/uploads")
async def upload_files(
    program: str = Form(...),
    branch: str = Form(...),
    batch: str = Form(...),
    semester: str = Form(...),  # roman: I..VIII
    exam_session: str = Form("December 2025"),
    tc_pdf: Optional[UploadFile] = File(None),
    gs_pdf: Optional[UploadFile] = File(None),
    sem_excel: Optional[UploadFile] = File(None),
    user: dict = Depends(get_current_admin),
):
    if not (tc_pdf or gs_pdf or sem_excel):
        raise HTTPException(400, "Upload at least one file")

    upload_id = str(uuid.uuid4())
    folder = STORAGE / upload_id
    folder.mkdir(parents=True, exist_ok=True)

    # 1. Parse SEM Excel — back map for the relevant semester
    back_map_for_sem: dict = {}
    excel_bytes = None
    if sem_excel:
        excel_bytes = await sem_excel.read()
        (folder / "input_sem.xlsx").write_bytes(excel_bytes)
        full = parse_sem_excel(excel_bytes)
        back_map_for_sem = full.get(semester.upper(), {})

    # 2. Parse + mark + regenerate TC
    tc_records = []
    tc_out_path = None
    tc_matched = 0
    if tc_pdf:
        b = await tc_pdf.read()
        (folder / "input_tc.pdf").write_bytes(b)
        tc_records = parse_tc_pdf(b)
        applied, tc_matched = apply_back_markers(tc_records, back_map_for_sem)
        log.info("TC: parsed %d students, applied %d back markers (matched=%d)",
                  len(tc_records), applied, tc_matched)
        # Resolve program/branch from records or fall back to form values
        prog = (tc_records[0]["program"] if tc_records else "") or program
        br = (tc_records[0]["branch"] if tc_records else "") or branch
        sem = (tc_records[0]["semester"] if tc_records else "") or semester
        out = generate_tc_pdf(tc_records, prog, br, sem, exam_session)
        tc_out_path = folder / "TC_starred.pdf"
        tc_out_path.write_bytes(out)

    # 3. Parse + mark + regenerate GS
    gs_records = []
    gs_out_path = None
    gs_matched = 0
    if gs_pdf:
        b = await gs_pdf.read()
        (folder / "input_gs.pdf").write_bytes(b)
        gs_records = parse_gs_pdf(b)
        applied, gs_matched = apply_back_markers(gs_records, back_map_for_sem)
        log.info("GS: parsed %d students, applied %d back markers (matched=%d)",
                  len(gs_records), applied, gs_matched)
        prog = (gs_records[0]["program"] if gs_records else "") or program
        br = (gs_records[0]["branch"] if gs_records else "") or branch
        sem = (gs_records[0]["semester"] if gs_records else "") or semester
        out = generate_gs_pdf(gs_records, prog, br, sem, exam_session, batch)
        gs_out_path = folder / "GS_starred.pdf"
        gs_out_path.write_bytes(out)

    # 4. Persist results: prefer GS records (one per page = clean), fall back to TC
    chosen_records = gs_records if gs_records else tc_records
    # Build a roll → gs_hash map from whichever set of records was processed
    # by `generate_gs_pdf` (it mutates records in place adding `gs_hash`).
    hash_by_roll = {
        r.get("roll_no", ""): r.get("gs_hash", "")
        for r in gs_records if r.get("gs_hash")
    }
    students_payload = []
    for rec in chosen_records:
        students_payload.append({
            "roll_no": rec.get("roll_no"),
            "name": rec.get("name", ""),
            "father_name": rec.get("father_name", ""),
            "enroll_no": rec.get("enroll_no", ""),
        })
        # upsert student
        await db.students.update_one(
            {"roll_no": rec.get("roll_no")},
            {"$set": {
                "roll_no": rec.get("roll_no"),
                "name": rec.get("name", ""),
                "father_name": rec.get("father_name", ""),
                "enroll_no": rec.get("enroll_no", ""),
                "program": program,
                "branch": branch,
                "batch": batch,
            }},
            upsert=True,
        )
        # upsert per-semester result
        await db.results.update_one(
            {"roll_no": rec.get("roll_no"), "semester": semester.upper()},
            {"$set": {
                "roll_no": rec.get("roll_no"),
                "semester": semester.upper(),
                "program": program,
                "branch": branch,
                "batch": batch,
                "exam_session": exam_session,
                "subjects": rec.get("subjects", []),
                "sgpa": rec.get("sgpa", ""),
                "cgpa": rec.get("cgpa", ""),
                "result": rec.get("result", ""),
                "remark": rec.get("remark", ""),
                "earned_credits": rec.get("earned_credits", ""),
                "gs_hash": rec.get("gs_hash") or hash_by_roll.get(rec.get("roll_no", ""), ""),
                "updated_at": datetime.now(timezone.utc).isoformat(),
            }},
            upsert=True,
        )

    upload_doc = {
        "id": upload_id,
        "program": program,
        "branch": branch,
        "batch": batch,
        "semester": semester.upper(),
        "exam_session": exam_session,
        "tc_count": len(tc_records),
        "gs_count": len(gs_records),
        "back_students_in_sem": len(back_map_for_sem),
        "tc_file": "TC_starred.pdf" if tc_out_path else None,
        "gs_file": "GS_starred.pdf" if gs_out_path else None,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "created_by": user["email"],
    }
    await db.uploads.insert_one(dict(upload_doc))

    return {
        "upload_id": upload_id,
        "tc_count": len(tc_records),
        "gs_count": len(gs_records),
        "back_students": len(back_map_for_sem),
        "tc_matched": tc_matched,
        "gs_matched": gs_matched,
        "warning": (
            "0 of the highlighted-back students from the Excel sheet matched any roll "
            "number in the uploaded PDFs. Make sure the Excel sheet and the PDFs "
            "are for the same batch."
            if back_map_for_sem and (tc_records or gs_records) and (tc_matched + gs_matched) == 0
            else None
        ),
        "tc_url": f"/api/admin/files/{upload_id}/tc" if tc_out_path else None,
        "gs_url": f"/api/admin/files/{upload_id}/gs" if gs_out_path else None,
        "students": students_payload,
    }


@api.post("/admin/uploads/excel")
async def upload_excel_only(
    program: str = Form(...),
    branch: str = Form(...),
    batch: str = Form(...),
    exam_session: str = Form("December 2025"),
    excel: UploadFile = File(...),
    user: dict = Depends(get_current_admin),
):
    """Upload a single Excel containing TC_X / GS_X / SEM_X sheets for all
    semesters of a batch+branch. Generates per-semester TC*/GS* PDFs with
    asterisks (from highlighted SEM_X cells) and barcodes on every GS page.
    """
    raw = await excel.read()
    upload_id = str(uuid.uuid4())
    folder = STORAGE / upload_id
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "input.xlsx").write_bytes(raw)

    # Parse all TC_X / GS_X sheets, plus SEM_X back map (highlighted cells)
    tc_gs = parse_tc_gs_excel(raw)
    sem_back = parse_sem_excel(raw)
    log.info("Excel parsed: TC sems=%s, GS sems=%s, SEM back sems=%s",
              list(tc_gs["TC"].keys()), list(tc_gs["GS"].keys()), list(sem_back.keys()))

    semesters_seen = sorted(
        set(tc_gs["TC"].keys()) | set(tc_gs["GS"].keys()),
        key=lambda r: ["I","II","III","IV","V","VI","VII","VIII"].index(r),
    )
    if not semesters_seen:
        raise HTTPException(
            400,
            "No TC_X or GS_X sheets found in the Excel. Sheet names should be "
            "TC_I, TC_II, ..., GS_I, GS_II, etc.",
        )

    sem_results: List[dict] = []
    rolls_known: dict = {}
    total_back_marks = 0
    # Build a per-roll, per-sem summary of SGPA/CGPA/etc. for embedding into GS.
    all_sem_summary: Dict[str, Dict[str, Dict[str, str]]] = {}
    for sem in semesters_seen:
        for rec in tc_gs["GS"].get(sem, []) or tc_gs["TC"].get(sem, []):
            roll = rec.get("roll_no", "")
            if not roll:
                continue
            all_sem_summary.setdefault(roll, {})[sem] = {
                "sgpa": rec.get("sgpa", ""),
                "cgpa": rec.get("cgpa", ""),
                "earned_credits": rec.get("earned_credits", ""),
                "result": rec.get("result", ""),
            }

    for sem in semesters_seen:
        back_map = sem_back.get(sem, {})
        tc_records = tc_gs["TC"].get(sem, [])
        gs_records = tc_gs["GS"].get(sem, [])
        a1, _ = apply_back_markers(tc_records, back_map)
        a2, _ = apply_back_markers(gs_records, back_map)
        apply_non_credit_markers(tc_records)
        apply_non_credit_markers(gs_records)
        # B.Tech-only TC fix: restore external marks where they were
        # suppressed as "-" but total > sessional (i.e. external marks
        # actually exist and equal total - sessional).
        if (program or "").strip().lower().startswith("b.tech"):
            restore_external_from_total(tc_records)
        total_back_marks += a1 + a2

        tc_out = None
        if tc_records:
            pdf_bytes = generate_tc_pdf(tc_records, program, branch, sem, exam_session)
            tc_path = folder / f"TC_{sem}_starred.pdf"
            tc_path.write_bytes(pdf_bytes)
            tc_out = tc_path.name
        gs_out = None
        if gs_records:
            pdf_bytes = generate_gs_pdf(
                gs_records, program, branch, sem, exam_session, batch,
                all_sem_summary=all_sem_summary,
            )
            gs_path = folder / f"GS_{sem}_starred.pdf"
            gs_path.write_bytes(pdf_bytes)
            gs_out = gs_path.name

        # Persist per-student result for this semester (prefer GS — cleanest)
        # Merge in TC's external/sessional/total marks (GS sheets don't have
        # them) so the student portal can display the same per-subject marks
        # breakdown that the TC PDF shows.
        tc_by_roll = {r.get("roll_no", ""): r for r in tc_records}
        for rec in (gs_records or tc_records):
            roll = rec.get("roll_no")
            if not roll:
                continue
            tc_rec = tc_by_roll.get(roll)
            if tc_rec and rec is not tc_rec:
                tc_subj_by_code = {
                    s.get("code", "").replace(" ", "").upper(): s
                    for s in tc_rec.get("subjects", [])
                }
                for s in rec.get("subjects", []):
                    key = s.get("code", "").replace(" ", "").upper()
                    tcs = tc_subj_by_code.get(key)
                    if tcs:
                        for f in ("external", "sessional", "total"):
                            if not s.get(f):
                                s[f] = tcs.get(f, "")
                        if tcs.get("back_pending") and not s.get("back_pending"):
                            s["back_pending"] = True
            rolls_known[roll] = {
                "roll_no": roll,
                "name": rec.get("name", ""),
                "father_name": rec.get("father_name", ""),
                "enroll_no": rec.get("enroll_no", ""),
            }
            # Cumulative earned credits comes only from TC — copy if missing.
            cuml = rec.get("cuml_earned_credits") or (
                tc_rec.get("cuml_earned_credits") if tc_rec else ""
            )
            await db.results.update_one(
                {"roll_no": roll, "semester": sem},
                {"$set": {
                    "roll_no": roll,
                    "semester": sem,
                    "program": program,
                    "branch": branch,
                    "batch": batch,
                    "exam_session": exam_session,
                    "subjects": rec.get("subjects", []),
                    "sgpa": rec.get("sgpa", ""),
                    "cgpa": rec.get("cgpa", ""),
                    "result": rec.get("result", ""),
                    "remark": rec.get("remark", ""),
                    "earned_credits": rec.get("earned_credits", ""),
                    "cuml_earned_credits": cuml,
                    "gs_hash": rec.get("gs_hash", ""),
                    "updated_at": datetime.now(timezone.utc).isoformat(),
                }},
                upsert=True,
            )
        sem_results.append({
            "semester": sem,
            "tc_count": len(tc_records),
            "gs_count": len(gs_records),
            "tc_file": tc_out,
            "gs_file": gs_out,
            "backs_in_sem": len(back_map),
            "asterisks_applied": a1 + a2,
        })

    # Upsert students
    for r, info in rolls_known.items():
        await db.students.update_one(
            {"roll_no": r},
            {"$set": {**info, "program": program, "branch": branch, "batch": batch}},
            upsert=True,
        )

    upload_doc = {
        "id": upload_id,
        "program": program,
        "branch": branch,
        "batch": batch,
        "semester": "ALL",
        "exam_session": exam_session,
        "tc_count": sum(s["tc_count"] for s in sem_results),
        "gs_count": sum(s["gs_count"] for s in sem_results),
        "back_students_in_sem": sum(s["backs_in_sem"] for s in sem_results),
        "tc_file": None,
        "gs_file": None,
        "semesters": sem_results,
        "source": "excel-only",
        "students_indexed": len(rolls_known),
        "total_asterisks_applied": total_back_marks,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "created_by": user["email"],
    }
    await db.uploads.insert_one(dict(upload_doc))

    return {
        "upload_id": upload_id,
        "semesters": sem_results,
        "students_indexed": len(rolls_known),
        "total_asterisks_applied": total_back_marks,
    }


# ---------------------------------------------------------------------------
# M.Tech-specific upload — branch-keyed sheets
# ---------------------------------------------------------------------------
@api.post("/admin/uploads/mtech/inspect")
async def mtech_inspect(
    excel: UploadFile = File(...),
    user: dict = Depends(get_current_admin),
):
    """Quick endpoint used by the M.Tech upload UI: returns the list of
    sheet names found in the uploaded workbook so the admin can pick which
    SEM_<branch>, TC_<branch> and GS_<branch> tabs to use for the run."""
    raw = await excel.read()
    info = inspect_workbook_sheets(raw)
    return info


@api.post("/admin/uploads/mtech")
async def upload_mtech(
    program: str = Form("Master of Technology"),
    branch: str = Form(...),
    batch: str = Form(...),
    semester: str = Form(...),
    exam_session: str = Form(...),
    sem_sheet: str = Form(...),
    tc_sheet: str = Form(""),
    gs_sheet: str = Form(""),
    student_limit: int = Form(0),
    excel: UploadFile = File(...),
    tc_pdf: Optional[UploadFile] = File(None),
    gs_pdf: Optional[UploadFile] = File(None),
    user: dict = Depends(get_current_admin),
):
    """M.Tech upload pipeline.

    Inputs:
      * **excel**       – workbook with branch-named sheets
      * **sem_sheet**   – name of the ``SEM_<branch>`` sheet to use as the
                          back-paper map for the chosen semester
      * **tc_sheet** / **gs_sheet** (optional) – named ``TC_*`` / ``GS_*``
                          sheets in the same workbook to use as data source
      * **tc_pdf** / **gs_pdf** (optional) – fall back to PDFs when the data
                          isn't available as Excel sheets

    Generates the regenerated TC + GS PDFs for the *chosen branch only* and
    persists per-student results with verification hashes.
    """
    sem = semester.strip().upper()
    raw_excel = await excel.read()
    tc_pdf_bytes = await tc_pdf.read() if tc_pdf else b""
    gs_pdf_bytes = await gs_pdf.read() if gs_pdf else b""

    upload_id = str(uuid.uuid4())
    folder = STORAGE / upload_id
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "input.xlsx").write_bytes(raw_excel)
    if tc_pdf_bytes:
        (folder / "tc_input.pdf").write_bytes(tc_pdf_bytes)
    if gs_pdf_bytes:
        (folder / "gs_input.pdf").write_bytes(gs_pdf_bytes)

    # 1. Back-paper map for this branch (from the chosen SEM_<branch> sheet)
    back_map = parse_sem_excel_sheet(raw_excel, sem_sheet)

    # 2. Source records — prefer a TC/GS sheet inside the workbook; fall back
    # to the uploaded PDFs if no sheet was specified.
    if tc_sheet:
        tc_records = parse_tc_or_gs_named_sheet(raw_excel, tc_sheet, "tc")
    elif tc_pdf_bytes:
        tc_records = parse_tc_pdf(tc_pdf_bytes)
    else:
        tc_records = []
    if gs_sheet:
        gs_records = parse_tc_or_gs_named_sheet(raw_excel, gs_sheet, "gs")
    elif gs_pdf_bytes:
        gs_records = parse_gs_pdf(gs_pdf_bytes)
    else:
        gs_records = []

    if not tc_records and not gs_records:
        raise HTTPException(400, "No TC or GS data found — provide either a "
                                  "tc_sheet/gs_sheet inside the Excel or a "
                                  "tc_pdf/gs_pdf PDF. (TC will be auto-derived "
                                  "from GS if only GS is supplied.)")

    # 3. Apply back markers + non-credit markers
    apply_back_markers(tc_records, back_map)
    apply_back_markers(gs_records, back_map)
    apply_non_credit_markers(tc_records, branch)
    apply_non_credit_markers(gs_records, branch)
    # For all M.Tech branches OTHER than Biotechnology, the Marksheets
    # layout displaces sessional/total marks into the grade column and the
    # letter grade has to be derived from the grade points.
    fix_mtech_non_credit_columns(tc_records, branch)
    fix_mtech_non_credit_columns(gs_records, branch)

    # If TC records weren't supplied (or parsing failed), derive them from GS
    # records so the M.Tech flow always produces both TC and GS PDFs.
    if not tc_records and gs_records:
        import copy as _copy
        tc_records = _copy.deepcopy(gs_records)
        log.info("TC records derived from GS (no TC source supplied) — count=%d",
                 len(tc_records))

    # Optional cap on number of students to render — drives the GS/TC volume
    # so the admin can quickly preview a subset of records.
    if student_limit and student_limit > 0:
        tc_records = tc_records[:student_limit]
        gs_records = gs_records[:student_limit]

    # 4. Generate fresh TC + GS PDFs for this branch
    out_files = {}
    if tc_records:
        try:
            prog = (tc_records[0].get("program") or program)
            br = (tc_records[0].get("branch") or branch)
            sem_roman = (tc_records[0].get("semester") or sem)
            tc_bytes = generate_tc_pdf(tc_records, prog, br, sem_roman, exam_session)
            (folder / f"TC_{sem}.pdf").write_bytes(tc_bytes)
            out_files["tc"] = f"TC_{sem}.pdf"
        except Exception as e:
            log.exception("TC generation failed: %s", e)
    if gs_records:
        try:
            prog = (gs_records[0].get("program") or program)
            br = (gs_records[0].get("branch") or branch)
            sem_roman = (gs_records[0].get("semester") or sem)
            # Build a per-roll all_sem_summary so the GS shows the complete
            # I-IV history for M.Tech students.
            existing = await db.results.find(
                {"roll_no": {"$in": [r.get("roll_no") for r in gs_records]}},
                {"_id": 0},
            ).to_list(2000)
            all_sem: dict = {}
            for r in existing:
                all_sem.setdefault(r.get("roll_no", ""), {})[r.get("semester", "")] = {
                    "sgpa": r.get("sgpa", ""),
                    "cgpa": r.get("cgpa", ""),
                    "earned_credits": r.get("earned_credits", ""),
                    "result": r.get("result", ""),
                }
            for r in gs_records:
                all_sem.setdefault(r.get("roll_no", ""), {})[sem_roman] = {
                    "sgpa": r.get("sgpa", ""),
                    "cgpa": r.get("cgpa", ""),
                    "earned_credits": r.get("earned_credits", ""),
                    "result": r.get("result", ""),
                }
            gs_bytes = generate_gs_pdf(
                gs_records, prog, br, sem_roman, exam_session, batch,
                all_sem_summary=all_sem,
            )
            (folder / f"GS_{sem}.pdf").write_bytes(gs_bytes)
            out_files["gs"] = f"GS_{sem}.pdf"
        except Exception as e:
            log.exception("GS generation failed: %s", e)

    # 5. Persist students & results (with gs_hash)
    chosen = gs_records or tc_records
    hash_by_roll = {
        r.get("roll_no", ""): r.get("gs_hash", "")
        for r in gs_records if r.get("gs_hash")
    }
    # Merge TC's per-subject marks (external / sessional / total) into the
    # chosen records so the Student Portal can render the same marks
    # breakdown that the TC PDF shows. GS sheets typically carry only the
    # grade & GP columns, leaving marks blank.
    if gs_records and tc_records:
        tc_by_roll = {r.get("roll_no", ""): r for r in tc_records}
        for rec in chosen:
            tc_rec = tc_by_roll.get(rec.get("roll_no", ""))
            if not tc_rec or tc_rec is rec:
                continue
            tc_subj = {
                (s.get("code") or "").replace(" ", "").upper(): s
                for s in tc_rec.get("subjects", [])
            }
            for s in rec.get("subjects", []):
                key = (s.get("code") or "").replace(" ", "").upper()
                tcs = tc_subj.get(key)
                if not tcs:
                    continue
                for f in ("external", "sessional", "total"):
                    if not s.get(f):
                        s[f] = tcs.get(f, "")
                if tcs.get("back_pending") and not s.get("back_pending"):
                    s["back_pending"] = True
                if tcs.get("back") and not s.get("back"):
                    s["back"] = True

    branch_disp = branch
    for rec in chosen:
        roll = rec.get("roll_no")
        if not roll:
            continue
        await db.students.update_one(
            {"roll_no": roll},
            {"$set": {
                "roll_no": roll,
                "name": rec.get("name", ""),
                "father_name": rec.get("father_name", ""),
                "enroll_no": rec.get("enroll_no", ""),
                "program": program,
                "branch": branch_disp,
                "batch": batch,
            }},
            upsert=True,
        )
        # Cumulative earned credits — copy from TC if GS row didn't carry it.
        cuml = rec.get("cuml_earned_credits") or ""
        if not cuml and tc_records:
            for tr in tc_records:
                if tr.get("roll_no") == roll:
                    cuml = tr.get("cuml_earned_credits") or ""
                    break
        await db.results.update_one(
            {"roll_no": roll, "semester": sem},
            {"$set": {
                "roll_no": roll,
                "semester": sem,
                "program": program,
                "branch": branch_disp,
                "batch": batch,
                "exam_session": exam_session,
                "subjects": rec.get("subjects", []),
                "sgpa": rec.get("sgpa", ""),
                "cgpa": rec.get("cgpa", ""),
                "result": rec.get("result", ""),
                "remark": rec.get("remark", ""),
                "earned_credits": rec.get("earned_credits", ""),
                "cuml_earned_credits": cuml,
                "gs_hash": rec.get("gs_hash") or hash_by_roll.get(roll, ""),
                "updated_at": datetime.now(timezone.utc).isoformat(),
            }},
            upsert=True,
        )

    # 6. Upload doc bookkeeping
    upload_doc = {
        "id": upload_id,
        "program": program,
        "branch": branch_disp,
        "batch": batch,
        "semester": sem,
        "exam_session": exam_session,
        "tc_count": len(tc_records),
        "gs_count": len(gs_records),
        "tc_file": out_files.get("tc"),
        "gs_file": out_files.get("gs"),
        "source": "mtech",
        "sem_sheet": sem_sheet,
        "tc_sheet": tc_sheet,
        "gs_sheet": gs_sheet,
        "students_indexed": len(chosen),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "created_by": user["email"],
    }
    await db.uploads.insert_one(dict(upload_doc))

    return {
        "upload_id": upload_id,
        "branch": branch_disp,
        "semester": sem,
        "tc_url": f"/api/admin/files/{upload_id}/tc" if out_files.get("tc") else None,
        "gs_url": f"/api/admin/files/{upload_id}/gs" if out_files.get("gs") else None,
        "students_indexed": len(chosen),
    }


@api.get("/admin/files/{upload_id}/sem/{semester}/{kind}")
async def download_sem_file(upload_id: str, semester: str, kind: str,
                              user: dict = Depends(get_current_admin)):
    """Download a per-semester TC* or GS* generated from an Excel-only upload."""
    folder = STORAGE / upload_id
    sem = semester.upper()
    if kind == "tc":
        fp = folder / f"TC_{sem}_starred.pdf"
        fname = f"TC_{sem}_starred.pdf"
    else:
        fp = folder / f"GS_{sem}_starred.pdf"
        fname = f"GS_{sem}_starred.pdf"
    if not fp.exists():
        raise HTTPException(404, "Not found")
    return StreamingResponse(open(fp, "rb"), media_type="application/pdf",
                              headers={"Content-Disposition": f'attachment; filename="{fname}"'})


@api.get("/admin/uploads")
async def list_uploads(user: dict = Depends(get_current_admin)):
    docs = await db.uploads.find({}, {"_id": 0}).sort("created_at", -1).to_list(500)
    return {"uploads": docs}


@api.delete("/admin/uploads/{upload_id}")
async def delete_upload(upload_id: str, user: dict = Depends(get_current_admin)):
    folder = STORAGE / upload_id
    if folder.exists():
        for f in folder.iterdir():
            f.unlink()
        folder.rmdir()
    await db.uploads.delete_one({"id": upload_id})
    return {"ok": True}


@api.get("/admin/files/{upload_id}/{kind}")
async def download_file(upload_id: str, kind: str, user: dict = Depends(get_current_admin)):
    folder = STORAGE / upload_id
    prefix = "TC" if kind == "tc" else "GS"
    # Look for the canonical name first, then any per-semester variant.
    candidates = [folder / f"{prefix}_starred.pdf"] + sorted(folder.glob(f"{prefix}_*.pdf"))
    fp = next((p for p in candidates if p.exists()), None)
    if not fp:
        raise HTTPException(404, "Not found")
    return StreamingResponse(open(fp, "rb"), media_type="application/pdf",
                              headers={"Content-Disposition": f'attachment; filename="{fp.name}"'})


# Per-student GS download (single student PDF) -----------------------------
@api.get("/admin/student/{roll_no}/gs/{semester}")
async def download_student_gs(roll_no: str, semester: str,
                                user: dict = Depends(get_current_admin)):
    res = await db.results.find_one({"roll_no": roll_no, "semester": semester.upper()},
                                      {"_id": 0})
    student = await db.students.find_one({"roll_no": roll_no}, {"_id": 0})
    if not res or not student:
        raise HTTPException(404, "Result not found")
    rec = {
        "roll_no": roll_no,
        "name": student.get("name", ""),
        "father_name": student.get("father_name", ""),
        "enroll_no": student.get("enroll_no", ""),
        "subjects": res.get("subjects", []),
        "sgpa": res.get("sgpa", ""),
        "cgpa": res.get("cgpa", ""),
        "result": res.get("result", ""),
        "remark": res.get("remark", ""),
        "earned_credits": res.get("earned_credits", ""),
        "program": res.get("program", ""),
        "branch": res.get("branch", ""),
        "semester": res.get("semester", ""),
        "exam_session": res.get("exam_session", ""),
    }
    # Build full sem-history for this student
    all_sem: Dict[str, Dict[str, Dict[str, str]]] = {roll_no: {}}
    async for r in db.results.find({"roll_no": roll_no}, {"_id": 0}):
        all_sem[roll_no][r["semester"]] = {
            "sgpa": r.get("sgpa", ""),
            "cgpa": r.get("cgpa", ""),
            "earned_credits": r.get("earned_credits", ""),
            "result": r.get("result", ""),
        }
    pdf = generate_gs_pdf([rec], res["program"], res["branch"], res["semester"],
                           res.get("exam_session", "December 2025"), res["batch"],
                           all_sem_summary=all_sem)
    return StreamingResponse(io_bytes(pdf), media_type="application/pdf",
                              headers={"Content-Disposition": f'attachment; filename="GS_{roll_no}_{semester}.pdf"'})


def io_bytes(b: bytes):
    import io as _io
    return _io.BytesIO(b)


# ---------------------------------------------------------------------------
# Students DOB management (for student portal auth)
# ---------------------------------------------------------------------------
@api.get("/admin/students")
async def list_students(user: dict = Depends(get_current_admin),
                          program: Optional[str] = None,
                          branch: Optional[str] = None,
                          batch: Optional[str] = None):
    q: dict = {}
    if program: q["program"] = program
    if branch: q["branch"] = branch
    if batch: q["batch"] = batch
    docs = await db.students.find(q, {"_id": 0}).sort("roll_no", 1).to_list(2000)
    return {"students": docs}


@api.post("/admin/students/dob")
async def set_dobs(items: List[DobItem], user: dict = Depends(get_current_admin)):
    n = 0
    for it in items:
        r = await db.students.update_one(
            {"roll_no": it.roll_no.strip()},
            {"$set": {"dob": it.dob.strip()}},
        )
        n += r.matched_count
    return {"updated": n}


# ---------------------------------------------------------------------------
# Student portal
# ---------------------------------------------------------------------------
@api.post("/student/login")
async def student_login(body: StudentLoginIn):
    roll = body.roll_no.strip()
    student = await db.students.find_one({"roll_no": roll}, {"_id": 0})
    if not student:
        raise HTTPException(404, "Roll number not found")
    results = await db.results.find({"roll_no": roll}, {"_id": 0}).to_list(20)
    # sort latest first by roman numeral
    order = {r: i for i, r in enumerate(["I","II","III","IV","V","VI","VII","VIII"])}
    results.sort(key=lambda r: order.get(r.get("semester","I"), 0), reverse=True)
    return {"student": student, "results": results}


# ---------------------------------------------------------------------------
# Public Grade-Sheet verification — looks up a GS by its printed hash code
# ---------------------------------------------------------------------------
@api.get("/verify/{gs_hash}")
async def verify_grade_sheet(gs_hash: str):
    h = (gs_hash or "").strip().upper()
    if not h or len(h) < 6:
        raise HTTPException(400, "Invalid hash")
    res = await db.results.find_one({"gs_hash": h}, {"_id": 0})
    if not res:
        raise HTTPException(404, "No grade sheet found for this verification code")
    student = await db.students.find_one(
        {"roll_no": res.get("roll_no", "")}, {"_id": 0}
    ) or {}
    return {
        "gs_hash": h,
        "student": {
            "name": student.get("name", ""),
            "father_name": student.get("father_name", ""),
            "roll_no": student.get("roll_no", res.get("roll_no", "")),
            "enroll_no": student.get("enroll_no", ""),
            "program": student.get("program", res.get("program", "")),
            "branch": student.get("branch", res.get("branch", "")),
            "batch": student.get("batch", res.get("batch", "")),
        },
        "result": {
            "semester": res.get("semester", ""),
            "exam_session": res.get("exam_session", ""),
            "sgpa": res.get("sgpa", ""),
            "cgpa": res.get("cgpa", ""),
            "earned_credits": res.get("earned_credits", ""),
            "cuml_earned_credits": res.get("cuml_earned_credits", ""),
            "result": res.get("result", ""),
            "remark": res.get("remark", ""),
            "subjects": res.get("subjects", []),
        },
    }


# ---------------------------------------------------------------------------
# Stats / overview
# ---------------------------------------------------------------------------
@api.get("/admin/stats")
async def admin_stats(user: dict = Depends(get_current_admin)):
    uploads = await db.uploads.count_documents({})
    students = await db.students.count_documents({})
    results = await db.results.count_documents({})
    backs = 0
    async for r in db.results.find({}, {"subjects": 1}):
        backs += sum(1 for s in r.get("subjects", []) if s.get("back"))
    return {
        "uploads": uploads,
        "students": students,
        "results": results,
        "backlogs_total": backs,
    }


# ---------------------------------------------------------------------------
# Startup: seed admin + indexes
# ---------------------------------------------------------------------------
@app.on_event("startup")
async def startup():
    await db.users.create_index("email", unique=True)
    await db.students.create_index("roll_no", unique=True)
    await db.results.create_index([("roll_no", 1), ("semester", 1)], unique=True)
    await db.uploads.create_index("created_at")

    existing = await db.users.find_one({"email": ADMIN_EMAIL.lower()})
    if not existing:
        await db.users.insert_one({
            "id": str(uuid.uuid4()),
            "email": ADMIN_EMAIL.lower(),
            "name": "Administrator",
            "role": "admin",
            "password_hash": hash_password(ADMIN_PASSWORD),
            "created_at": datetime.now(timezone.utc).isoformat(),
        })
        log.info("Seeded admin user: %s", ADMIN_EMAIL)
    elif not verify_password(ADMIN_PASSWORD, existing["password_hash"]):
        await db.users.update_one(
            {"email": ADMIN_EMAIL.lower()},
            {"$set": {"password_hash": hash_password(ADMIN_PASSWORD)}},
        )
        log.info("Updated admin password for %s", ADMIN_EMAIL)


@app.on_event("shutdown")
async def shutdown():
    client.close()


# ---------------------------------------------------------------------------
# Health & router mounting
# ---------------------------------------------------------------------------
@api.get("/")
async def root():
    return {"service": "GBPIET Result Asterisk Portal", "status": "ok"}


app.include_router(api)
_cors = os.environ.get("CORS_ORIGINS", "*")
_origins = [o.strip() for o in _cors.split(",") if o.strip()]
_kwargs = dict(allow_methods=["*"], allow_headers=["*"], allow_credentials=True)
if "*" in _origins:
    _kwargs["allow_origin_regex"] = ".*"
else:
    _kwargs["allow_origins"] = _origins
app.add_middleware(CORSMiddleware, **_kwargs)
