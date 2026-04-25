import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "../components/ui/button";
import { Card } from "../components/ui/card";
import { LogOut, Printer } from "lucide-react";

const ROMAN_ORDER = ["VIII", "VII", "VI", "V", "IV", "III", "II", "I"];

export default function StudentResults() {
  const nav = useNavigate();
  const [data, setData] = useState(null);

  useEffect(() => {
    const raw = sessionStorage.getItem("student_data");
    if (!raw) {
      nav("/student");
      return;
    }
    setData(JSON.parse(raw));
  }, [nav]);

  if (!data) return null;
  const { student, results } = data;
  const ordered = [...results].sort(
    (a, b) => ROMAN_ORDER.indexOf(a.semester) - ROMAN_ORDER.indexOf(b.semester)
  );

  // Sum 'a/b' style mark strings → "totalA/totalB" for the semester.
  const sumTotalMarks = (subjects) => {
    let a = 0,
      b = 0,
      hit = false;
    (subjects || []).forEach((s) => {
      const m = (s.total || "").match(/^(-?\d+)\s*\/\s*(\d+)$/);
      if (m) {
        a += parseInt(m[1], 10);
        b += parseInt(m[2], 10);
        hit = true;
      }
    });
    return hit ? `${a}/${b}` : "—";
  };

  // A subject is highlighted only when it was marked back AND the student
  // actually cleared it (i.e. did not get F / Ab / Dt).
  const isCleared = (s) => s.back && !s.back_pending;

  // Pre-compute running cumulative earned credits keyed by semester roman.
  // Iterate in ascending semester order (I → VIII) regardless of display order.
  const ASC_SEMS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"];
  const cumBySem = {};
  let _runningEC = 0;
  ASC_SEMS.forEach((s) => {
    const r = ordered.find((x) => x.semester === s);
    if (r) {
      _runningEC += parseInt(r.earned_credits, 10) || 0;
      cumBySem[s] = _runningEC;
    }
  });

  return (
    <div className="min-h-screen bg-stone-100">
      <header className="bg-white border-b border-stone-200 print:hidden">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
              Govind Ballabh Pant Institute of Engineering &amp; Technology
            </p>
            <h1 className="font-display text-xl">Results</h1>
          </div>
          <div className="flex gap-2">
            <Button
              size="sm"
              variant="outline"
              onClick={() => window.print()}
              data-testid="print-results"
              className="rounded-sm"
            >
              <Printer className="w-4 h-4 mr-2" /> Print
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => {
                sessionStorage.removeItem("student_data");
                nav("/student");
              }}
              data-testid="student-logout"
              className="rounded-sm"
            >
              <LogOut className="w-4 h-4 mr-2" /> Sign out
            </Button>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-12 space-y-10">
        {/* Legend */}
        <Card
          data-testid="legend-card"
          className="bg-white border border-stone-200 px-6 py-4 rounded-sm shadow-none flex flex-wrap items-center gap-6 text-xs font-mono"
        >
          <div className="flex items-center gap-2">
            <span className="subject-back text-xs">*</span>
            <span className="text-stone-700">subject cleared after back paper</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="px-2 py-0.5 bg-stone-100 border border-stone-200 rounded-sm text-stone-700 font-semibold">$</span>
            <span className="text-stone-700">non-credit subject</span>
          </div>
        </Card>
        {/* Student card */}
        <Card
          data-testid="student-card"
          className="bg-white border border-stone-200 p-8 md:p-10 shadow-sm rounded-sm paper-surface fade-up"
        >
          <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
            Student profile
          </p>
          <h2 className="font-display text-4xl md:text-5xl mt-3 leading-tight">
            {student.name || "Student"}
          </h2>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-6 mt-8 font-mono text-sm">
            <div>
              <p className="text-xs text-stone-500 uppercase tracking-[0.15em]">
                Roll No.
              </p>
              <p className="mt-1">{student.roll_no}</p>
            </div>
            <div>
              <p className="text-xs text-stone-500 uppercase tracking-[0.15em]">
                Enrollment
              </p>
              <p className="mt-1">{student.enroll_no || "—"}</p>
            </div>
            <div>
              <p className="text-xs text-stone-500 uppercase tracking-[0.15em]">
                Branch
              </p>
              <p className="mt-1 text-xs">{student.branch || "—"}</p>
            </div>
            <div>
              <p className="text-xs text-stone-500 uppercase tracking-[0.15em]">
                Batch
              </p>
              <p className="mt-1">{student.batch || "—"}</p>
            </div>
          </div>
        </Card>

        {ordered.length === 0 && (
          <Card className="p-10 text-center bg-white border-stone-200 rounded-sm shadow-none">
            <p className="font-display text-2xl">No results recorded yet</p>
            <p className="text-sm text-stone-500 mt-2">
              The examination cell has not uploaded any results for your roll
              number. Please check back later.
            </p>
          </Card>
        )}

        {/* Semester-wise summary table */}
        {ordered.length > 0 && (
          <Card
            data-testid="semester-summary-card"
            className="bg-white border border-stone-200 p-8 md:p-10 rounded-sm shadow-sm fade-up"
          >
            <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
              Academic summary
            </p>
            <h3 className="font-display text-3xl mt-1">
              Semester-wise overview
            </h3>
            <div className="mt-6 overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-xs uppercase tracking-[0.15em] text-stone-500 border-b border-stone-200">
                    <th className="text-left p-2 font-semibold">Semester</th>
                    <th className="text-left p-2 font-semibold">Session</th>
                    <th className="text-right p-2 font-semibold">SGPA</th>
                    <th className="text-right p-2 font-semibold">CGPA</th>
                    <th className="text-right p-2 font-semibold">Total Marks</th>
                    <th className="text-right p-2 font-semibold">Earned Cr.</th>
                    <th className="text-right p-2 font-semibold">Cum. Earned Cr.</th>
                    <th className="text-left p-2 font-semibold">Result</th>
                    <th className="text-left p-2 font-semibold">Remark</th>
                  </tr>
                </thead>
                <tbody className="font-mono">
                  {ordered.map((r) => {
                      return (
                        <tr
                          key={r.semester}
                          data-testid={`summary-row-${r.semester}`}
                          className="border-t border-stone-100 hover:bg-stone-50"
                        >
                          <td className="p-2 font-semibold">
                            Sem {r.semester}
                          </td>
                          <td className="p-2 text-xs text-stone-600">
                            {r.exam_session}
                          </td>
                          <td className="p-2 text-right">{r.sgpa || "—"}</td>
                          <td className="p-2 text-right">{r.cgpa || "—"}</td>
                          <td className="p-2 text-right">
                            {sumTotalMarks(r.subjects)}
                          </td>
                          <td className="p-2 text-right">
                            {r.earned_credits || "—"}
                          </td>
                          <td className="p-2 text-right">
                            {cumBySem[r.semester] || "—"}
                          </td>
                          <td className="p-2">
                            <span
                              className={
                                r.result?.toLowerCase().includes("pass")
                                  ? "text-emerald-700 font-semibold"
                                  : "text-amber-800 font-semibold"
                              }
                            >
                              {r.result || "—"}
                            </span>
                          </td>
                          <td className="p-2 text-xs text-stone-600">
                            {r.remark || "—"}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>
          </Card>
        )}

        {ordered.map((r, idx) => {
          const hasMarks = (r.subjects || []).some(
            (s) => s.external || s.total
          );
          return (
            <Card
              key={r.semester}
              data-testid={`semester-card-${r.semester}`}
              className="semester-card bg-white border border-stone-200 p-8 md:p-10 rounded-sm shadow-sm fade-up"
              style={{ animationDelay: `${idx * 80}ms` }}
            >
              {/* Print-only mini profile — repeats at the top of every printed semester page */}
              <div className="print-profile hidden print:block mb-6 pb-4 border-b border-stone-300">
                <div className="flex items-baseline justify-between gap-4 flex-wrap">
                  <div>
                    <p className="text-[10px] uppercase tracking-[0.2em] text-stone-500 font-semibold">
                      GBPIET · Govind Ballabh Pant Institute of Engineering &amp; Technology
                    </p>
                    <p className="font-display text-xl mt-1">{student.name || "Student"}</p>
                  </div>
                  <div className="font-mono text-[11px] text-stone-700 grid grid-cols-2 md:grid-cols-4 gap-x-6 gap-y-1">
                    <div>
                      <span className="text-stone-500">Roll:</span>{" "}
                      <span className="font-semibold">{student.roll_no}</span>
                    </div>
                    <div>
                      <span className="text-stone-500">Enrollment:</span>{" "}
                      <span>{student.enroll_no || "—"}</span>
                    </div>
                    <div className="col-span-2">
                      <span className="text-stone-500">Branch:</span>{" "}
                      <span>{student.branch || "—"}</span>
                    </div>
                    <div>
                      <span className="text-stone-500">Batch:</span>{" "}
                      <span>{student.batch || "—"}</span>
                    </div>
                    {student.program && (
                      <div className="col-span-3">
                        <span className="text-stone-500">Programme:</span>{" "}
                        <span>{student.program}</span>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="flex items-end justify-between flex-wrap gap-4">
                <div>
                  <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
                    Semester
                  </p>
                  <h3 className="font-display text-3xl mt-1">
                    {r.semester} Semester
                  </h3>
                  <p className="text-xs text-stone-500 mt-1 font-mono">
                    {r.exam_session}
                  </p>
                </div>
                <div className="text-right font-mono text-sm space-y-1">
                  {r.sgpa && (
                    <p>
                      <span className="text-stone-500">SGPA:</span>{" "}
                      <span className="font-semibold">{r.sgpa}</span>
                    </p>
                  )}
                  {r.cgpa && (
                    <p>
                      <span className="text-stone-500">CGPA:</span>{" "}
                      <span className="font-semibold">{r.cgpa}</span>
                    </p>
                  )}
                  {r.result && (
                    <p>
                      <span className="text-stone-500">Result:</span>{" "}
                      <span className="font-semibold">{r.result}</span>
                    </p>
                  )}
                </div>
              </div>

              <table className="w-full text-sm mt-8 border-t border-stone-200">
                <thead>
                  <tr className="text-xs uppercase tracking-[0.15em] text-stone-500">
                    <th className="text-left p-2 font-semibold w-28">Code</th>
                    <th className="text-left p-2 font-semibold">Subject</th>
                    <th className="text-right p-2 font-semibold w-16">Cr.</th>
                    {hasMarks && (
                      <>
                        <th className="text-right p-2 font-semibold">Ext</th>
                        <th className="text-right p-2 font-semibold">Ses</th>
                        <th className="text-right p-2 font-semibold">Total</th>
                      </>
                    )}
                    <th className="text-right p-2 font-semibold w-16">Grade</th>
                    <th className="text-right p-2 font-semibold w-20">GP</th>
                  </tr>
                </thead>
                <tbody className="font-mono">
                  {(r.subjects || []).map((s, i) => {
                    const cleared = isCleared(s);
                    return (
                      <tr
                        key={i}
                        className={`border-t border-stone-100 ${
                          cleared ? "bg-amber-50" : ""
                        }`}
                      >
                        <td className="p-2">{s.code}</td>
                        <td className="p-2 sans">
                          {cleared ? (
                            <span className="subject-back">{s.name}</span>
                          ) : (
                            s.name
                          )}
                        </td>
                        <td className="p-2 text-right">{s.credits}</td>
                        {hasMarks && (
                          <>
                            <td className="p-2 text-right">{s.external || "—"}</td>
                            <td className="p-2 text-right">{s.sessional || "—"}</td>
                            <td className="p-2 text-right">{s.total || "—"}</td>
                          </>
                        )}
                        <td className="p-2 text-right">{s.grade}</td>
                        <td className="p-2 text-right">{s.grade_points}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {r.remark && (
                <p className="mt-4 text-xs text-stone-500 font-mono">
                  Remark: {r.remark}
                </p>
              )}
            </Card>
          );
        })}
      </main>
    </div>
  );
}
