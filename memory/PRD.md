# GBPIET Result Asterisk Portal — PRD

## Original Problem Statement
> Portal where admin can upload TC PDF, GS PDF and SEM Excel sheets; SEM_X highlighted (yellow / blue) cells indicate back subjects; regenerate TC*/GS* PDFs in the original format. Programme list: B.Tech (CSE, CSE-AIML, ECE, EE, Biotech, Civil, Mech, Mech-MFG), MCA, M.Tech (CSE, Power System, Geotechnology, Biotech, ECE, Production, Thermal, Digital Communication). Batches 2022, 2023, 2024, 2025. Separate student portal where students enter roll number to view all-semester results.

## Architecture
- **Backend**: FastAPI + MongoDB + JWT bcrypt admin auth + pdfplumber + openpyxl + reportlab + python-barcode + PyMuPDF
- **Frontend**: React 19 + react-router 7 + shadcn-ui + Tailwind
- **Routes**: `/`, `/admin/login`, `/admin/{overview,upload,files,students}`, `/student`, `/student/results`

## User Personas
1. **Examination Cell Administrator** — manages uploads & exports.
2. **Student** — looks up own results semester-wise (latest first).

## Core Requirements (Static)
- Admin email/password login.
- Two upload modes:
  - **Excel-only** (single file with `SEM_X`, `TC_X`, `GS_X` sheets) → generates ALL semester TC*/GS* PDFs in one shot.
  - **PDF + Excel** (TC PDF + GS PDF + SEM Excel for one semester) → generates that semester's TC*/GS*.
- Detect highlighted (any non-default solid fill) cells in SEM_X → mark matching subjects with `*`.
- Regenerated PDFs match original layout (institute logo left, UTU logo right, centered text).
- Every GS page embeds a Code-128 barcode encoding the student's roll number (independent of SL.NO).
- Per-student GS PDF download.
- Student lookup by roll only (no DOB).
- Semester-wise (latest first) results display with `*` highlighting.

## Implemented Timeline
- **2026-04-25 v1**: Base portal — admin auth, PDF parsers, SEM-Excel back detection, TC*/GS* regeneration, student portal w/ DOB. (i1, 14/14 tests pass)
- **2026-04-25 v1.1**: DOB removed for students; institute & UTU logos embedded in regenerated PDF headers. (i2, 17/17 tests pass)
- **2026-04-25 v1.2**: Excel-only generation pipeline — single Excel → all-semester TC*/GS* PDFs; Code-128 barcode on every GS page. New endpoints `POST /api/admin/uploads/excel`, `GET /api/admin/files/{id}/sem/{sem}/{kind}`. New tabbed Upload UI; Generated Files now expandable per-semester. (i3, 24/24 tests pass)

## What's Mocked
- Nothing.

## Backlog (P1)
- Bulk-batch ZIP export (all per-student GS\* in one zip).
- Cumulative consolidated GS (all 8 semesters combined into one PDF per student).
- Real M.Tech / MCA samples to validate parsers on those formats.
- WhatsApp / email share of personalised result link from student portal.

## Backlog (P2)
- Email notifications on result publication.
- Audit log of admin actions.
- Multi-admin roles.

## Test Credentials
See `/app/memory/test_credentials.md`.

## Key Files
- `/app/backend/server.py` — endpoints
- `/app/backend/processor.py` — parsing + PDF generation + barcodes
- `/app/backend/assets/{institute_logo.png, utu_logo.png}`
- `/app/frontend/src/pages/{AdminUpload, AdminFiles, AdminStudents, StudentLogin, StudentResults}.jsx`
- `/app/backend/tests/test_portal_api.py` — 24 tests
