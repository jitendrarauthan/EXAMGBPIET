import { useEffect, useMemo, useState } from "react";
import { api, fmtError } from "../lib/api";
import { Card } from "../components/ui/card";
import { Input } from "../components/ui/input";
import { toast } from "sonner";
import { Search } from "lucide-react";

export default function AdminStudents() {
  const [students, setStudents] = useState([]);
  const [loading, setLoading] = useState(true);
  const [q, setQ] = useState("");

  useEffect(() => {
    api
      .get("/admin/students")
      .then((r) => setStudents(r.data.students))
      .catch((e) => toast.error(fmtError(e)))
      .finally(() => setLoading(false));
  }, []);

  const filtered = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (!t) return students;
    return students.filter(
      (s) =>
        (s.roll_no || "").toLowerCase().includes(t) ||
        (s.name || "").toLowerCase().includes(t) ||
        (s.branch || "").toLowerCase().includes(t)
    );
  }, [q, students]);

  const downloadGs = async (roll, sem) => {
    const tok = localStorage.getItem("admin_token");
    const r = await fetch(
      `${process.env.REACT_APP_BACKEND_URL}/api/admin/student/${roll}/gs/${sem}`,
      { headers: { Authorization: `Bearer ${tok}` } }
    );
    if (!r.ok) {
      toast.error("Result not found for that semester");
      return;
    }
    const blob = await r.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `GS_${roll}_${sem}.pdf`;
    a.click();
  };

  return (
    <div className="p-10 fade-up">
      <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
        Master register
      </p>
      <h1 className="font-display text-4xl mt-1">Students</h1>
      <p className="text-stone-600 text-sm mt-2">
        Roll numbers populated automatically from uploaded TC / GS PDFs. Use the
        Quick GS column to export a single-student starred grade sheet for any
        semester on file.
      </p>

      <div className="relative mt-8 max-w-md">
        <Search className="w-4 h-4 absolute left-3 top-3 text-stone-400" />
        <Input
          placeholder="Search by roll, name or branch…"
          value={q}
          onChange={(e) => setQ(e.target.value)}
          data-testid="students-search"
          className="pl-9 rounded-sm"
        />
      </div>

      <Card className="mt-6 rounded-sm border-stone-200 shadow-none overflow-hidden">
        {loading && <p className="p-6 text-stone-500">Loading…</p>}
        {!loading && (
          <div className="overflow-auto max-h-[70vh]">
            <table className="w-full text-sm">
              <thead className="bg-stone-50 border-b border-stone-200 sticky top-0">
                <tr className="text-xs uppercase tracking-[0.15em] text-stone-500">
                  <th className="text-left p-3">Roll No</th>
                  <th className="text-left p-3">Name</th>
                  <th className="text-left p-3">Branch</th>
                  <th className="text-left p-3">Batch</th>
                  <th className="text-left p-3">Quick GS export</th>
                </tr>
              </thead>
              <tbody className="font-mono">
                {filtered.length === 0 && (
                  <tr>
                    <td
                      colSpan={5}
                      className="p-8 text-center text-stone-500 italic"
                    >
                      No students found.
                    </td>
                  </tr>
                )}
                {filtered.map((s) => (
                  <tr
                    key={s.roll_no}
                    className="border-t border-stone-100 hover:bg-stone-50"
                  >
                    <td className="p-3">{s.roll_no}</td>
                    <td className="p-3 sans">{s.name}</td>
                    <td className="p-3 text-stone-600 text-xs">{s.branch}</td>
                    <td className="p-3">{s.batch}</td>
                    <td className="p-3">
                      <div className="flex gap-1 flex-wrap">
                        {["I", "II", "III", "IV", "V", "VI", "VII", "VIII"].map(
                          (sm) => (
                            <button
                              key={sm}
                              onClick={() => downloadGs(s.roll_no, sm)}
                              className="text-xs px-2 py-0.5 border border-stone-200 hover:border-indigo-950 rounded-sm"
                              data-testid={`dl-gs-${s.roll_no}-${sm}`}
                            >
                              {sm}
                            </button>
                          )
                        )}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}
