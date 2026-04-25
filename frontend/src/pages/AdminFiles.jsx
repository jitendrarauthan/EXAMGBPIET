import { useEffect, useState } from "react";
import { api, fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Card } from "../components/ui/card";
import { Download, Trash2 } from "lucide-react";
import { toast } from "sonner";

export default function AdminFiles() {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);

  const load = () =>
    api
      .get("/admin/uploads")
      .then((r) => setItems(r.data.uploads))
      .catch((e) => toast.error(fmtError(e)))
      .finally(() => setLoading(false));

  useEffect(() => {
    load();
  }, []);

  const download = async (id, kind) => {
    const tok = localStorage.getItem("admin_token");
    const r = await fetch(
      `${process.env.REACT_APP_BACKEND_URL}/api/admin/files/${id}/${kind}`,
      { headers: { Authorization: `Bearer ${tok}` } }
    );
    if (!r.ok) {
      toast.error("File not found");
      return;
    }
    const blob = await r.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = kind === "tc" ? "TC_starred.pdf" : "GS_starred.pdf";
    a.click();
  };

  const remove = async (id) => {
    if (!window.confirm("Delete this upload and generated files?")) return;
    try {
      await api.delete(`/admin/uploads/${id}`);
      toast.success("Deleted");
      load();
    } catch (e) {
      toast.error(fmtError(e));
    }
  };

  return (
    <div className="p-10 fade-up">
      <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
        Output archive
      </p>
      <h1 className="font-display text-4xl mt-1">Generated files</h1>
      <p className="text-stone-600 text-sm mt-2">
        Every upload produces a starred TC* and / or GS* PDF that can be
        downloaded any time.
      </p>

      <div className="mt-8 space-y-3">
        {loading && <p className="text-stone-500">Loading…</p>}
        {!loading && items.length === 0 && (
          <Card className="p-10 text-center rounded-sm shadow-none border-stone-200">
            <p className="font-display text-xl text-stone-700">No files yet</p>
            <p className="text-sm text-stone-500 mt-1">
              Run an upload to generate starred TC* / GS* files.
            </p>
          </Card>
        )}
        {items.map((u) => (
          <Card
            key={u.id}
            data-testid={`file-row-${u.id}`}
            className="p-5 rounded-sm border-stone-200 shadow-none flex items-center justify-between gap-4"
          >
            <div className="flex-1 min-w-0">
              <p className="text-sm font-medium">
                {u.program} · <span className="text-stone-700">{u.branch}</span>
              </p>
              <p className="text-xs text-stone-500 mt-1 font-mono">
                Batch {u.batch} · Sem {u.semester} · {u.exam_session} · TC=
                {u.tc_count} · GS={u.gs_count} · backs={u.back_students_in_sem}
              </p>
              <p className="text-xs text-stone-400 mt-1 font-mono">
                {new Date(u.created_at).toLocaleString()}
              </p>
            </div>
            <div className="flex items-center gap-2">
              {u.tc_file && (
                <Button
                  onClick={() => download(u.id, "tc")}
                  data-testid={`dl-tc-${u.id}`}
                  size="sm"
                  className="rounded-sm bg-indigo-950 hover:bg-indigo-900"
                >
                  <Download className="w-3.5 h-3.5 mr-1.5" /> TC*
                </Button>
              )}
              {u.gs_file && (
                <Button
                  onClick={() => download(u.id, "gs")}
                  data-testid={`dl-gs-${u.id}`}
                  size="sm"
                  variant="outline"
                  className="rounded-sm border-stone-300"
                >
                  <Download className="w-3.5 h-3.5 mr-1.5" /> GS*
                </Button>
              )}
              <Button
                onClick={() => remove(u.id)}
                data-testid={`del-${u.id}`}
                size="sm"
                variant="ghost"
                className="rounded-sm text-stone-400 hover:text-red-700"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </Button>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}
