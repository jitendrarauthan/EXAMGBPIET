import { useEffect, useState } from "react";
import { api, fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Card } from "../components/ui/card";
import { Download, Trash2, ChevronDown, Sparkles } from "lucide-react";
import { toast } from "sonner";

function downloadFromAuth(url, filename) {
  const tok = localStorage.getItem("admin_token");
  return fetch(`${process.env.REACT_APP_BACKEND_URL}${url}`, {
    headers: { Authorization: `Bearer ${tok}` },
  }).then((r) => {
    if (!r.ok) throw new Error("File not found");
    return r.blob().then((blob) => {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      a.click();
    });
  });
}

export default function AdminFiles() {
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [expanded, setExpanded] = useState({});

  const load = () =>
    api
      .get("/admin/uploads")
      .then((r) => setItems(r.data.uploads))
      .catch((e) => toast.error(fmtError(e)))
      .finally(() => setLoading(false));

  useEffect(() => {
    load();
  }, []);

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
        Every upload produces starred TC* and GS* PDFs that can be downloaded
        any time. Excel-only uploads expand to show per-semester files.
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
        {items.map((u) => {
          const isExcel = u.source === "excel-only";
          const open = expanded[u.id];
          return (
            <Card
              key={u.id}
              data-testid={`file-row-${u.id}`}
              className="rounded-sm border-stone-200 shadow-none overflow-hidden"
            >
              <div className="p-5 flex items-center justify-between gap-4">
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <p className="text-sm font-medium">
                      {u.program} ·{" "}
                      <span className="text-stone-700">{u.branch}</span>
                    </p>
                    {isExcel && (
                      <span className="inline-flex items-center gap-1 text-[10px] uppercase tracking-[0.15em] bg-indigo-950 text-white px-2 py-0.5 rounded-sm">
                        <Sparkles className="w-3 h-3" /> excel-only
                      </span>
                    )}
                  </div>
                  <p className="text-xs text-stone-500 mt-1 font-mono">
                    Batch {u.batch} ·{" "}
                    {isExcel ? "All sems" : `Sem ${u.semester}`} ·{" "}
                    {u.exam_session} · TC={u.tc_count} · GS={u.gs_count}
                  </p>
                  <p className="text-xs text-stone-400 mt-1 font-mono">
                    {new Date(u.created_at).toLocaleString()}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  {isExcel ? (
                    <Button
                      onClick={() =>
                        setExpanded((p) => ({ ...p, [u.id]: !open }))
                      }
                      size="sm"
                      variant="outline"
                      data-testid={`expand-${u.id}`}
                      className="rounded-sm border-stone-300"
                    >
                      <ChevronDown
                        className={`w-3.5 h-3.5 mr-1.5 transition-transform ${
                          open ? "rotate-180" : ""
                        }`}
                      />
                      {open ? "Hide" : "Show"} files
                    </Button>
                  ) : (
                    <>
                      {u.tc_file && (
                        <Button
                          onClick={() =>
                            downloadFromAuth(
                              `/api/admin/files/${u.id}/tc`,
                              "TC_starred.pdf"
                            ).catch((e) => toast.error(e.message))
                          }
                          data-testid={`dl-tc-${u.id}`}
                          size="sm"
                          className="rounded-sm bg-indigo-950 hover:bg-indigo-900"
                        >
                          <Download className="w-3.5 h-3.5 mr-1.5" /> TC*
                        </Button>
                      )}
                      {u.gs_file && (
                        <Button
                          onClick={() =>
                            downloadFromAuth(
                              `/api/admin/files/${u.id}/gs`,
                              "GS_starred.pdf"
                            ).catch((e) => toast.error(e.message))
                          }
                          data-testid={`dl-gs-${u.id}`}
                          size="sm"
                          variant="outline"
                          className="rounded-sm border-stone-300"
                        >
                          <Download className="w-3.5 h-3.5 mr-1.5" /> GS*
                        </Button>
                      )}
                    </>
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
              </div>
              {isExcel && open && (
                <div className="border-t border-stone-200 bg-stone-50 p-4 space-y-1.5">
                  {(u.semesters || []).map((s) => (
                    <div
                      key={s.semester}
                      data-testid={`sem-files-${u.id}-${s.semester}`}
                      className="flex items-center justify-between gap-3 px-2 py-1.5 rounded-sm hover:bg-white"
                    >
                      <div className="font-mono text-xs">
                        <span className="font-semibold">Sem {s.semester}</span>
                        <span className="text-stone-500 ml-2">
                          TC {s.tc_count} · GS {s.gs_count} ·{" "}
                          {s.asterisks_applied} *
                        </span>
                      </div>
                      <div className="flex gap-1.5">
                        {s.tc_file && (
                          <Button
                            onClick={() =>
                              downloadFromAuth(
                                `/api/admin/files/${u.id}/sem/${s.semester}/tc`,
                                `TC_${s.semester}_starred.pdf`
                              ).catch((e) => toast.error(e.message))
                            }
                            size="sm"
                            data-testid={`dl-sem-tc-${u.id}-${s.semester}`}
                            className="rounded-sm bg-indigo-950 hover:bg-indigo-900 h-7 text-xs"
                          >
                            <Download className="w-3 h-3 mr-1" /> TC*
                          </Button>
                        )}
                        {s.gs_file && (
                          <Button
                            onClick={() =>
                              downloadFromAuth(
                                `/api/admin/files/${u.id}/sem/${s.semester}/gs`,
                                `GS_${s.semester}_starred.pdf`
                              ).catch((e) => toast.error(e.message))
                            }
                            size="sm"
                            variant="outline"
                            data-testid={`dl-sem-gs-${u.id}-${s.semester}`}
                            className="rounded-sm border-stone-300 h-7 text-xs"
                          >
                            <Download className="w-3 h-3 mr-1" /> GS*
                          </Button>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          );
        })}
      </div>
    </div>
  );
}
