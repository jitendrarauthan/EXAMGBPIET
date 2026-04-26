import { useRef, useState } from "react";
import { api, fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Input } from "../components/ui/input";
import { Label } from "../components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../components/ui/select";
import { Card } from "../components/ui/card";
import { toast } from "sonner";
import {
  Upload as UploadIcon,
  FileSpreadsheet,
  Loader2,
  Download,
  ArrowRight,
  X,
} from "lucide-react";

const SEMS = ["I", "II", "III", "IV"];

function FileBox({ label, accept, file, setFile, testid }) {
  const ref = useRef(null);
  return (
    <Card
      data-testid={testid}
      onClick={() => ref.current?.click()}
      className="p-5 border-stone-200 rounded-sm shadow-none border-dashed hover:border-indigo-950 cursor-pointer transition-colors"
    >
      <input
        ref={ref}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => setFile(e.target.files?.[0] || null)}
      />
      <div className="flex items-center gap-3">
        <FileSpreadsheet className="w-5 h-5 text-stone-600" />
        <div className="flex-1">
          <p className="text-xs uppercase tracking-[0.15em] text-stone-500 font-semibold">
            {label}
          </p>
          <p className="font-mono text-sm mt-1 truncate">
            {file ? file.name : "Click to choose file"}
          </p>
        </div>
        {file && (
          <Button
            size="icon"
            variant="ghost"
            className="rounded-sm"
            onClick={(e) => {
              e.stopPropagation();
              setFile(null);
            }}
          >
            <X className="w-4 h-4" />
          </Button>
        )}
      </div>
    </Card>
  );
}

export default function AdminUploadMTech() {
  // form state
  const [excel, setExcel] = useState(null);
  const [tcPdf, setTcPdf] = useState(null);
  const [gsPdf, setGsPdf] = useState(null);

  const [program] = useState("Master of Technology");
  const [branch, setBranch] = useState("");
  const [batch, setBatch] = useState("");
  const [semester, setSemester] = useState("I");
  const [examSession, setExamSession] = useState("December 2024");

  // sheet-name picks
  const [semSheet, setSemSheet] = useState("");
  const [tcSheet, setTcSheet] = useState("");
  const [gsSheet, setGsSheet] = useState("");

  // discovered sheet lists
  const [sheets, setSheets] = useState({ sem: [], tc: [], gs: [], all: [] });
  const [inspecting, setInspecting] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [result, setResult] = useState(null);

  const inspect = async () => {
    if (!excel) return toast.error("Choose the M.Tech Excel first");
    setInspecting(true);
    setSheets({ sem: [], tc: [], gs: [], all: [] });
    try {
      const fd = new FormData();
      fd.append("excel", excel);
      const { data } = await api.post("/admin/uploads/mtech/inspect", fd, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      setSheets(data);
      // auto-pick if there's only one
      if (data.sem?.length === 1) setSemSheet(data.sem[0]);
      if (data.tc?.length === 1) setTcSheet(data.tc[0]);
      if (data.gs?.length === 1) setGsSheet(data.gs[0]);
      toast.success(
        `Found ${data.sem.length} SEM_, ${data.tc.length} TC_, ${data.gs.length} GS_ sheets`,
      );
    } catch (err) {
      toast.error(fmtError(err));
    } finally {
      setInspecting(false);
    }
  };

  const submit = async (e) => {
    e.preventDefault();
    if (!excel) return toast.error("Choose the Excel workbook");
    if (!semSheet)
      return toast.error("Pick a SEM_<branch> sheet for back papers");
    if (!tcSheet && !tcPdf)
      return toast.error("Pick a TC_ sheet OR upload a TC PDF");
    if (!gsSheet && !gsPdf)
      return toast.error("Pick a GS_ sheet OR upload a GS PDF");
    if (!branch || !batch)
      return toast.error("Branch and batch are required");

    setSubmitting(true);
    setResult(null);
    try {
      const fd = new FormData();
      fd.append("program", program);
      fd.append("branch", branch);
      fd.append("batch", batch);
      fd.append("semester", semester);
      fd.append("exam_session", examSession);
      fd.append("sem_sheet", semSheet);
      fd.append("tc_sheet", tcSheet || "");
      fd.append("gs_sheet", gsSheet || "");
      fd.append("excel", excel);
      if (tcPdf) fd.append("tc_pdf", tcPdf);
      if (gsPdf) fd.append("gs_pdf", gsPdf);
      const { data } = await api.post("/admin/uploads/mtech", fd, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      setResult(data);
      toast.success(
        `M.Tech ${data.branch} · Sem ${data.semester} generated for ${data.students_indexed} students`,
      );
    } catch (err) {
      toast.error(fmtError(err));
    } finally {
      setSubmitting(false);
    }
  };

  const apiBase = process.env.REACT_APP_BACKEND_URL;

  return (
    <main className="max-w-4xl mx-auto space-y-8">
      <div>
        <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
          M.Tech &middot; Branch-wise upload
        </p>
        <h1 className="font-display text-3xl mt-1">
          Generate TC &amp; GS for a single M.Tech branch
        </h1>
        <p className="text-stone-600 text-sm mt-3 leading-relaxed max-w-2xl">
          The M.Tech workbook contains separate{" "}
          <span className="font-mono">SEM_&lt;branch&gt;</span>,{" "}
          <span className="font-mono">TC_&lt;branch&gt;</span> and{" "}
          <span className="font-mono">GS_&lt;branch&gt;</span> sheets per
          specialisation. Upload the Excel, pick which sheets to use for the
          chosen branch, and the system will produce the regenerated TC + GS
          PDFs only for that branch.
        </p>
      </div>

      <Card className="p-6 md:p-8 rounded-sm border-stone-200">
        <form onSubmit={submit} data-testid="mtech-upload-form" className="space-y-7">
          {/* Step 1: workbook + inspect */}
          <div className="space-y-3">
            <p className="text-xs uppercase tracking-[0.15em] text-stone-500 font-semibold">
              Step 1 — Upload the M.Tech workbook
            </p>
            <FileBox
              label="M.Tech Excel"
              accept=".xlsx,.xls"
              file={excel}
              setFile={(f) => {
                setExcel(f);
                setSheets({ sem: [], tc: [], gs: [], all: [] });
                setSemSheet("");
                setTcSheet("");
                setGsSheet("");
              }}
              testid="mtech-excel-slot"
            />
            <Button
              type="button"
              onClick={inspect}
              disabled={!excel || inspecting}
              data-testid="mtech-inspect-btn"
              variant="outline"
              className="rounded-sm"
            >
              {inspecting ? (
                <Loader2 className="w-4 h-4 animate-spin mr-2" />
              ) : (
                <ArrowRight className="w-4 h-4 mr-2" />
              )}
              Inspect sheets
            </Button>
          </div>

          {/* Step 2: pick sheets */}
          {sheets.all.length > 0 && (
            <div className="space-y-3 border-t border-stone-200 pt-6">
              <p className="text-xs uppercase tracking-[0.15em] text-stone-500 font-semibold">
                Step 2 — Pick the source sheets for this branch
              </p>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <SheetPicker
                  label="SEM_ sheet (back papers)"
                  options={sheets.sem}
                  value={semSheet}
                  onChange={setSemSheet}
                  testid="mtech-sem-sheet"
                  required
                />
                <SheetPicker
                  label="TC_ sheet (optional)"
                  options={sheets.tc}
                  value={tcSheet}
                  onChange={setTcSheet}
                  testid="mtech-tc-sheet"
                />
                <SheetPicker
                  label="GS_ sheet (optional)"
                  options={sheets.gs}
                  value={gsSheet}
                  onChange={setGsSheet}
                  testid="mtech-gs-sheet"
                />
              </div>
              <p className="text-xs text-stone-500 leading-relaxed">
                You can use a TC_ / GS_ sheet from the workbook OR upload a
                separate PDF in step 3 below.
              </p>
            </div>
          )}

          {/* Step 3: optional PDFs */}
          <div className="space-y-3 border-t border-stone-200 pt-6">
            <p className="text-xs uppercase tracking-[0.15em] text-stone-500 font-semibold">
              Step 3 — Or upload TC / GS PDFs (optional)
            </p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <FileBox
                label="TC PDF"
                accept=".pdf"
                file={tcPdf}
                setFile={setTcPdf}
                testid="mtech-tc-pdf-slot"
              />
              <FileBox
                label="GS PDF"
                accept=".pdf"
                file={gsPdf}
                setFile={setGsPdf}
                testid="mtech-gs-pdf-slot"
              />
            </div>
          </div>

          {/* Step 4: metadata */}
          <div className="space-y-3 border-t border-stone-200 pt-6">
            <p className="text-xs uppercase tracking-[0.15em] text-stone-500 font-semibold">
              Step 4 — Branch &amp; semester details
            </p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Field label="Branch (full name)" required>
                <Input
                  value={branch}
                  onChange={(e) => setBranch(e.target.value)}
                  placeholder="Computer Science & Engineering"
                  data-testid="mtech-branch-input"
                  className="rounded-sm"
                  required
                />
              </Field>
              <Field label="Batch" required>
                <Input
                  value={batch}
                  onChange={(e) => setBatch(e.target.value)}
                  placeholder="2023"
                  data-testid="mtech-batch-input"
                  className="rounded-sm"
                  required
                />
              </Field>
              <Field label="Semester" required>
                <Select value={semester} onValueChange={setSemester}>
                  <SelectTrigger
                    data-testid="mtech-semester"
                    className="rounded-sm"
                  >
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    {SEMS.map((s) => (
                      <SelectItem key={s} value={s}>
                        Semester {s}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </Field>
              <Field label="Exam session">
                <Input
                  value={examSession}
                  onChange={(e) => setExamSession(e.target.value)}
                  placeholder="December 2024"
                  data-testid="mtech-session-input"
                  className="rounded-sm"
                />
              </Field>
            </div>
          </div>

          <div className="border-t border-stone-200 pt-6">
            <Button
              type="submit"
              disabled={submitting}
              data-testid="mtech-submit"
              className="rounded-sm bg-indigo-950 hover:bg-indigo-900 h-11 px-8"
            >
              {submitting ? (
                <Loader2 className="w-4 h-4 animate-spin mr-2" />
              ) : (
                <UploadIcon className="w-4 h-4 mr-2" />
              )}
              Generate TC &amp; GS
            </Button>
          </div>
        </form>
      </Card>

      {result && (
        <Card
          data-testid="mtech-result-card"
          className="p-6 md:p-8 rounded-sm border-emerald-200 bg-emerald-50/30 fade-up"
        >
          <p className="text-xs uppercase tracking-[0.2em] text-emerald-800 font-semibold">
            Generated successfully
          </p>
          <h2 className="font-display text-2xl mt-2">
            {result.branch} · Semester {result.semester}
          </h2>
          <p className="text-stone-600 text-sm mt-2 font-mono">
            {result.students_indexed} students indexed
          </p>
          <div className="flex flex-wrap gap-3 mt-6">
            {result.tc_url && (
              <a
                href={`${apiBase}${result.tc_url}`}
                target="_blank"
                rel="noreferrer"
                data-testid="mtech-download-tc"
                className="inline-flex items-center gap-2 rounded-sm border border-stone-300 px-4 py-2 text-sm font-mono hover:bg-stone-100"
              >
                <Download className="w-4 h-4" /> TC PDF
              </a>
            )}
            {result.gs_url && (
              <a
                href={`${apiBase}${result.gs_url}`}
                target="_blank"
                rel="noreferrer"
                data-testid="mtech-download-gs"
                className="inline-flex items-center gap-2 rounded-sm border border-stone-300 px-4 py-2 text-sm font-mono hover:bg-stone-100"
              >
                <Download className="w-4 h-4" /> GS PDF
              </a>
            )}
          </div>
        </Card>
      )}
    </main>
  );
}

const Field = ({ label, required, children }) => (
  <div>
    <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
      {label} {required && <span className="text-red-700">*</span>}
    </Label>
    <div className="mt-1">{children}</div>
  </div>
);

function SheetPicker({ label, options, value, onChange, testid, required }) {
  if (!options || options.length === 0) {
    return (
      <Field label={label} required={required}>
        <p className="text-xs text-stone-500 italic">No matching sheets found</p>
      </Field>
    );
  }
  return (
    <Field label={label} required={required}>
      <Select value={value || ""} onValueChange={onChange}>
        <SelectTrigger data-testid={testid} className="rounded-sm">
          <SelectValue placeholder="Select sheet" />
        </SelectTrigger>
        <SelectContent>
          {options.map((s) => (
            <SelectItem key={s} value={s}>
              {s}
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
    </Field>
  );
}
