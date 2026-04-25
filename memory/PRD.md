# GBPIET Result Asterisk Portal — PRD

## Original Problem Statement
> "I need a portal (dashboard) where I can upload generated PDF of TC, GS and excel sheets of SEM_. You have to add an asterisk (*) to the end of the subject name in their TC and GS in PDF of a particular student of a particular semester if for that subject the student has a highlighted cell in sheet SEM_I, SEM_II, SEM_III, SEM_IV, SEM_V, SEM_VI, SEM_VII, SEM_VIII. And where I can upload B.Tech (CSE, CSE-AIML, ECE, EE, Biotech, Civil, Mech, Mech-MFG), MCA, M.Tech (CSE, Power System, Geotechnology, Biotech, ECE, Production, Thermal, Digital Communication) TC, GS and SEM_ Sheets for batches 2022, 2023, 2024, 2025 in different files. You have to extract the details of each PDF and Excel sheets, and generate the updated * files of TC and GS in the same format as I have attached. As well as I need a separate portal (dashboard) for students where the students can insert their roll number and their complete (sem-wise last to first) results (including * subjects) are to be displayed there."

## Architecture
- **Backend** (FastAPI) — `/app/backend/server.py`, `/app/backend/processor.py`
  - JWT-based admin auth (bcrypt, PyJWT, httpOnly cookie + Bearer fallback)
  - PDF parsing (`pdfplumber`) + Excel parsing (`openpyxl`) + PDF generation (`reportlab`)
  - MongoDB (`motor`): `users`, `students`, `results`, `uploads`
- **Frontend** (React 19 + react-router 7 + shadcn-ui + Tailwind)
  - Routes: `/`, `/admin/login`, `/admin/*`, `/student`, `/student/results`

## User Personas
1. **Examination Cell Administrator** — uploads transcripts and SEM sheets, downloads starred output, manages student DOB.
2. **Student** — looks up own results semester-wise (latest first), sees back subjects highlighted with `*`.

## Core Requirements (Static)
- Admin email/password login with seeded credentials.
- Admin upload of TC PDF, GS PDF, SEM Excel per (Programme, Branch, Batch, Semester).
- Detect highlighted (yellow / blue or any non-default solid fill) cells in `SEM_X` sheets → mark matching subjects with `*`.
- Regenerate TC* / GS* PDFs in the original institute format.
- Per-student GS PDF download.
- Student lookup by roll + DOB.
- Semester-wise (latest first) results display with `*` highlighting.

## Implemented (2026-04-25)
- ✅ Admin auth (JWT, seeded admin) at `/api/auth/login`, `/api/auth/me`, `/api/auth/logout`.
- ✅ Upload pipeline `POST /api/admin/uploads` (multipart) — parses TC + GS + SEM xlsx, applies asterisks, persists per-student per-semester results, stores generated PDFs.
- ✅ TC PDF parser: 61 students extracted from sample (multi-student-per-page layout).
- ✅ GS PDF parser: 70 students extracted from sample (one student per page).
- ✅ SEM Excel parser: 20 students with backs detected from sample AIML 2022 batch.
- ✅ TC* / GS* PDF regeneration with reportlab, matches institute layout (header, info table, subject table with back rows in amber, summary, footer).
- ✅ Per-student single-student GS download.
- ✅ Listing / deletion of uploads.
- ✅ Student DOB management (manual entry per row).
- ✅ Student login (roll + DOB).
- ✅ Semester-wise results page with `*` markers, SGPA/CGPA, print friendly.
- ✅ Mismatch warning when SEM excel rolls don't intersect uploaded PDF rolls.
- ✅ End-to-end testing — 14/14 backend tests passing, frontend flows verified (`/app/test_reports/iteration_1.json`).

## What's Mocked
- Nothing is mocked. All endpoints, parsing, persistence, and PDF generation are real.

## Backlog (P1)
- Bulk DOB CSV upload (currently only manual per-row).
- Programme-level batch view (all branches in a programme summarized).
- Cumulative consolidated GS (all 8 semesters in one PDF per student).
- Student forgot-DOB / contact admin flow.
- M.Tech / MCA program tested with real samples (only B.Tech samples in current test set).

## Backlog (P2)
- Batch download (zip of all student GS PDFs).
- Email notification to students when new semester result is uploaded.
- Audit log of admin actions.
- Multi-admin support with roles.

## Test Credentials
See `/app/memory/test_credentials.md`.

## Files
- Backend: `/app/backend/server.py`, `/app/backend/processor.py`
- Frontend: `/app/frontend/src/App.js`, `/app/frontend/src/lib/auth.jsx`, `/app/frontend/src/pages/*`
- Tests: `/app/backend/tests/test_portal_api.py`
- Storage: `/app/backend/storage/<upload_id>/{TC_starred.pdf, GS_starred.pdf, input_*.{pdf,xlsx}}`
