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
                    <th className="text-right p-2 font-semibold">Earned Cr.</th>
                    <th className="text-left p-2 font-semibold">Result</th>
                    <th className="text-left p-2 font-semibold">Remark</th>
                  </tr>
                </thead>
                <tbody className="font-mono">
                  {ordered.map((r) => {
                      const backCount = (r.subjects || []).filter(
                        (s) => s.back,
                      ).length;
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
                            {r.earned_credits || "—"}
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
                            {backCount > 0 && (
                              <span className="ml-2 text-xs text-amber-900">
                                ({backCount} back*)
                              </span>
                            )}
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
          const backCount = (r.subjects || []).filter((s) => s.back).length;
          return (
            <Card
              key={r.semester}
              data-testid={`semester-card-${r.semester}`}
              className="bg-white border border-stone-200 p-8 md:p-10 rounded-sm shadow-sm fade-up"
              style={{ animationDelay: `${idx * 80}ms` }}
            >
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
                  {backCount > 0 && (
                    <p className="text-amber-900 font-semibold">
                      {backCount} back subject{backCount > 1 ? "s" : ""} *
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
                    {(r.subjects?.[0]?.external) && (
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
                  {(r.subjects || []).map((s, i) => (
                    <tr
                      key={i}
                      className={`border-t border-stone-100 ${
                        s.back ? "bg-amber-50" : ""
                      }`}
                    >
                      <td className="p-2">{s.code}</td>
                      <td className="p-2 sans">
                        {s.back ? (
                          <span className="subject-back">{s.name}</span>
                        ) : (
                          s.name
                        )}
                      </td>
                      <td className="p-2 text-right">{s.credits}</td>
                      {s.external && (
                        <>
                          <td className="p-2 text-right">{s.external}</td>
                          <td className="p-2 text-right">{s.sessional}</td>
                          <td className="p-2 text-right">{s.total}</td>
                        </>
                      )}
                      <td className="p-2 text-right">{s.grade}</td>
                      <td className="p-2 text-right">{s.grade_points}</td>
                    </tr>
                  ))}
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
