import { useEffect, useRef, useState } from "react";
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
  FileText,
  FileSpreadsheet,
  X,
  Loader2,
  CheckCircle2,
  Download,
} from "lucide-react";

const SEMS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"];

function FileSlot({ id, label, accept, file, setFile }) {
  const ref = useRef(null);
  return (
    <Card
      data-testid={`upload-slot-${id}`}
      className="p-5 border-stone-200 rounded-sm shadow-none border-dashed hover:border-indigo-950 transition-colors cursor-pointer"
      onClick={() => ref.current?.click()}
    >
      <input
        ref={ref}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => setFile(e.target.files?.[0] || null)}
      />
      <div className="flex items-start gap-3">
        {accept.includes("pdf") ? (
          <FileText className="w-5 h-5 text-stone-500 mt-0.5" strokeWidth={1.5} />
        ) : (
          <FileSpreadsheet
            className="w-5 h-5 text-stone-500 mt-0.5"
            strokeWidth={1.5}
          />
        )}
        <div className="flex-1 min-w-0">
          <p className="text-xs uppercase tracking-[0.15em] text-stone-500 font-semibold">
            {label}
          </p>
          {file ? (
            <div className="flex items-center justify-between gap-2 mt-1">
              <p className="font-mono text-sm truncate">{file.name}</p>
              <button
                type="button"
                onClick={(e) => {
                  e.stopPropagation();
                  setFile(null);
                }}
                className="text-stone-400 hover:text-stone-700"
                data-testid={`clear-${id}`}
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          ) : (
            <p className="text-stone-400 text-sm mt-1">Click to choose…</p>
          )}
        </div>
      </div>
    </Card>
  );
}

export default function AdminUpload() {
  const [meta, setMeta] = useState({
    programs: {},
    batches: [],
    semesters: SEMS,
  });
  const [program, setProgram] = useState("");
  const [branch, setBranch] = useState("");
  const [batch, setBatch] = useState("");
  const [semester, setSemester] = useState("");
  const [examSession, setExamSession] = useState("December 2025");
  const [tcFile, setTcFile] = useState(null);
  const [gsFile, setGsFile] = useState(null);
  const [excelFile, setExcelFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState(null);

  useEffect(() => {
    api.get("/meta/programs").then((r) => setMeta(r.data));
  }, []);

  const branches = program ? meta.programs[program] || [] : [];

  const submit = async () => {
    if (!program || !branch || !batch || !semester) {
      toast.error("Please fill program, branch, batch and semester.");
      return;
    }
    if (!tcFile && !gsFile && !excelFile) {
      toast.error("Upload at least one file.");
      return;
    }
    const fd = new FormData();
    fd.append("program", program);
    fd.append("branch", branch);
    fd.append("batch", batch);
    fd.append("semester", semester);
    fd.append("exam_session", examSession);
    if (tcFile) fd.append("tc_pdf", tcFile);
    if (gsFile) fd.append("gs_pdf", gsFile);
    if (excelFile) fd.append("sem_excel", excelFile);

    setBusy(true);
    setResult(null);
    try {
      const { data } = await api.post("/admin/uploads", fd, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      setResult(data);
      toast.success(
        `Processed: ${data.tc_count} TC + ${data.gs_count} GS records (${data.back_students} backs)`
      );
    } catch (e) {
      toast.error(fmtError(e));
    } finally {
      setBusy(false);
    }
  };

  const downloadFile = async (kind) => {
    const url = `/admin/files/${result.upload_id}/${kind}`;
    const tok = localStorage.getItem("admin_token");
    const r = await fetch(
      `${process.env.REACT_APP_BACKEND_URL}/api${url}`,
      { headers: { Authorization: `Bearer ${tok}` } }
    );
    const blob = await r.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = kind === "tc" ? "TC_starred.pdf" : "GS_starred.pdf";
    a.click();
  };

  return (
    <div className="p-10 fade-up max-w-5xl">
      <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
        Step 1 · Configure batch
      </p>
      <h1 className="font-display text-4xl mt-1">Upload &amp; process</h1>
      <p className="text-stone-600 text-sm mt-2 max-w-2xl">
        Pick the programme, branch, batch and semester. Then attach the
        original TC PDF, GS PDF and the SEM_ Excel sheet (highlighted yellow /
        blue cells in SEM_ are treated as back subjects).
      </p>

      {/* Filters */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mt-8">
        <div>
          <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
            Programme
          </Label>
          <Select value={program} onValueChange={(v) => { setProgram(v); setBranch(""); }}>
            <SelectTrigger
              data-testid="program-select"
              className="mt-1 rounded-sm"
            >
              <SelectValue placeholder="Choose…" />
            </SelectTrigger>
            <SelectContent>
              {Object.keys(meta.programs).map((p) => (
                <SelectItem key={p} value={p}>
                  {p}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <div>
          <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
            Branch
          </Label>
          <Select value={branch} onValueChange={setBranch} disabled={!program}>
            <SelectTrigger
              data-testid="branch-select"
              className="mt-1 rounded-sm"
            >
              <SelectValue placeholder="Choose…" />
            </SelectTrigger>
            <SelectContent>
              {branches.map((b) => (
                <SelectItem key={b} value={b}>
                  {b}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <div>
          <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
            Batch
          </Label>
          <Select value={batch} onValueChange={setBatch}>
            <SelectTrigger
              data-testid="batch-select"
              className="mt-1 rounded-sm"
            >
              <SelectValue placeholder="Choose…" />
            </SelectTrigger>
            <SelectContent>
              {meta.batches.map((b) => (
                <SelectItem key={b} value={b}>
                  {b}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <div>
          <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
            Semester
          </Label>
          <Select value={semester} onValueChange={setSemester}>
            <SelectTrigger
              data-testid="semester-select"
              className="mt-1 rounded-sm"
            >
              <SelectValue placeholder="Choose…" />
            </SelectTrigger>
            <SelectContent>
              {SEMS.map((s) => (
                <SelectItem key={s} value={s}>
                  Sem {s}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
        <div className="md:col-span-2">
          <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
            Exam session label
          </Label>
          <Input
            value={examSession}
            onChange={(e) => setExamSession(e.target.value)}
            data-testid="exam-session-input"
            className="mt-1 rounded-sm"
          />
        </div>
      </div>

      {/* Files */}
      <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold mt-12">
        Step 2 · Attach files
      </p>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-3">
        <FileSlot
          id="tc"
          label="TC PDF"
          accept="application/pdf,.pdf"
          file={tcFile}
          setFile={setTcFile}
        />
        <FileSlot
          id="gs"
          label="GS PDF"
          accept="application/pdf,.pdf"
          file={gsFile}
          setFile={setGsFile}
        />
        <FileSlot
          id="excel"
          label="SEM_ Excel"
          accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          file={excelFile}
          setFile={setExcelFile}
        />
      </div>

      <Button
        onClick={submit}
        disabled={busy}
        data-testid="upload-submit-btn"
        className="mt-10 rounded-sm bg-indigo-950 hover:bg-indigo-900 h-11 px-6"
      >
        {busy ? (
          <>
            <Loader2 className="w-4 h-4 mr-2 animate-spin" />
            Processing…
          </>
        ) : (
          <>
            <UploadIcon className="w-4 h-4 mr-2" /> Process &amp; generate starred files
          </>
        )}
      </Button>

      {/* Result */}
      {result && (
        <Card
          data-testid="upload-result-card"
          className="p-6 mt-10 border-amber-200 bg-amber-50 rounded-sm shadow-none"
        >
          <div className="flex items-start gap-3">
            <CheckCircle2 className="w-5 h-5 text-amber-700 mt-0.5" />
            <div className="flex-1">
              <h3 className="font-display text-2xl">Done</h3>
              <p className="text-stone-700 text-sm mt-1 font-mono">
                TC: {result.tc_count} students · GS: {result.gs_count} students ·
                Back-marked students in SEM excel: {result.back_students}
              </p>
              {result.warning && (
                <p
                  data-testid="upload-warning"
                  className="text-red-900 text-sm mt-3 bg-red-50 border border-red-200 px-3 py-2 rounded-sm"
                >
                  ⚠ {result.warning}
                </p>
              )}
              <div className="flex gap-3 mt-5">
                {result.tc_url && (
                  <Button
                    onClick={() => downloadFile("tc")}
                    data-testid="download-tc-btn"
                    className="rounded-sm bg-indigo-950 hover:bg-indigo-900"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download TC*.pdf
                  </Button>
                )}
                {result.gs_url && (
                  <Button
                    onClick={() => downloadFile("gs")}
                    data-testid="download-gs-btn"
                    variant="outline"
                    className="rounded-sm border-stone-300"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download GS*.pdf
                  </Button>
                )}
              </div>
            </div>
          </div>
        </Card>
      )}
    </div>
  );
}
