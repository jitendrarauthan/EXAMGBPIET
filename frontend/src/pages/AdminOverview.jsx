import { useEffect, useState } from "react";
import { api } from "../lib/api";
import { Card } from "../components/ui/card";
import { Upload, FilePlus2, Users, AlertTriangle } from "lucide-react";
import { Link } from "react-router-dom";

export default function AdminOverview() {
  const [stats, setStats] = useState({
    uploads: 0,
    students: 0,
    results: 0,
    backlogs_total: 0,
  });
  const [recents, setRecents] = useState([]);

  useEffect(() => {
    api
      .get("/admin/stats")
      .then((r) => setStats(r.data))
      .catch(() => {});
    api
      .get("/admin/uploads")
      .then((r) => setRecents(r.data.uploads.slice(0, 6)))
      .catch(() => {});
  }, []);

  const cards = [
    { label: "Total uploads", value: stats.uploads, icon: Upload },
    { label: "Students indexed", value: stats.students, icon: Users },
    { label: "Result records", value: stats.results, icon: FilePlus2 },
    {
      label: "Back subjects flagged",
      value: stats.backlogs_total,
      icon: AlertTriangle,
    },
  ];

  return (
    <div className="p-10 fade-up">
      <div className="flex items-end justify-between">
        <div>
          <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
            Examination cell
          </p>
          <h1 className="font-display text-4xl mt-1">Overview</h1>
        </div>
        <Link
          to="/admin/upload"
          data-testid="quick-upload-link"
          className="text-sm bg-indigo-950 hover:bg-indigo-900 text-white px-4 py-2 rounded-sm flex items-center gap-2"
        >
          <Upload className="w-4 h-4" /> New upload
        </Link>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mt-10">
        {cards.map(({ label, value, icon: Icon }) => (
          <Card
            key={label}
            data-testid={`stat-${label.toLowerCase().replace(/ /g, "-")}`}
            className="p-6 rounded-sm border-stone-200 shadow-none"
          >
            <Icon className="w-5 h-5 text-stone-500" strokeWidth={1.5} />
            <p className="text-xs uppercase tracking-[0.15em] text-stone-500 mt-6">
              {label}
            </p>
            <p className="font-mono text-3xl mt-2">{value}</p>
          </Card>
        ))}
      </div>

      <div className="mt-12">
        <h2 className="font-display text-2xl">Recent uploads</h2>
        <div className="mt-4 border border-stone-200 bg-white">
          <table className="w-full text-sm">
            <thead className="border-b border-stone-200 bg-stone-50">
              <tr className="text-xs uppercase tracking-[0.15em] text-stone-500">
                <th className="text-left p-3 font-semibold">Programme</th>
                <th className="text-left p-3 font-semibold">Branch</th>
                <th className="text-left p-3 font-semibold">Batch</th>
                <th className="text-left p-3 font-semibold">Sem</th>
                <th className="text-right p-3 font-semibold">TC</th>
                <th className="text-right p-3 font-semibold">GS</th>
                <th className="text-right p-3 font-semibold">Backs (excel)</th>
                <th className="text-left p-3 font-semibold">Created</th>
              </tr>
            </thead>
            <tbody className="font-mono">
              {recents.length === 0 && (
                <tr>
                  <td
                    colSpan={8}
                    className="p-8 text-center text-stone-500 italic"
                  >
                    No uploads yet — go to{" "}
                    <Link to="/admin/upload" className="underline">
                      Upload
                    </Link>{" "}
                    to begin.
                  </td>
                </tr>
              )}
              {recents.map((u) => (
                <tr key={u.id} className="border-t border-stone-100 hover:bg-stone-50">
                  <td className="p-3">{u.program}</td>
                  <td className="p-3">{u.branch}</td>
                  <td className="p-3">{u.batch}</td>
                  <td className="p-3">{u.semester}</td>
                  <td className="p-3 text-right">{u.tc_count}</td>
                  <td className="p-3 text-right">{u.gs_count}</td>
                  <td className="p-3 text-right">{u.back_students_in_sem}</td>
                  <td className="p-3 text-stone-500 text-xs">
                    {new Date(u.created_at).toLocaleString()}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
