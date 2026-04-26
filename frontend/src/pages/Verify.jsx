import { useState } from "react";
import { Link } from "react-router-dom";
import { api, fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Input } from "../components/ui/input";
import { Label } from "../components/ui/label";
import { Card } from "../components/ui/card";
import { toast } from "sonner";
import { ShieldCheck, Loader2, ArrowLeft, BadgeCheck } from "lucide-react";
import PortalHeader from "../components/PortalHeader";
import PortalFooter from "../components/PortalFooter";

export default function Verify() {
  const [code, setCode] = useState("");
  const [busy, setBusy] = useState(false);
  const [data, setData] = useState(null);

  const submit = async (e) => {
    e.preventDefault();
    const h = code.trim().toUpperCase();
    if (!h) return toast.error("Enter the verification code from the GS");
    setBusy(true);
    setData(null);
    try {
      const { data } = await api.get(`/verify/${encodeURIComponent(h)}`);
      setData(data);
    } catch (err) {
      toast.error(fmtError(err));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="min-h-screen bg-stone-100 flex flex-col">
      <PortalHeader
        right={
          <Link
            to="/"
            className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold hover:text-stone-800 inline-flex items-center gap-1.5"
            data-testid="verify-home-link"
          >
            <ArrowLeft className="w-3.5 h-3.5" /> Back home
          </Link>
        }
      />

      <main className="max-w-3xl mx-auto px-6 py-14 space-y-10 flex-1 w-full">
        <Card
          data-testid="verify-search-card"
          className="bg-white border border-stone-200 p-8 md:p-10 rounded-sm shadow-sm"
        >
          <ShieldCheck className="w-9 h-9 text-emerald-700" strokeWidth={1.5} />
          <h1 className="font-display text-3xl mt-4">Verify a Grade Sheet</h1>
          <p className="text-stone-600 text-sm mt-2">
            Enter the alphanumeric verification code printed below the barcode
            on the official GBPIET Grade Sheet to confirm its authenticity.
          </p>

          <form
            onSubmit={submit}
            className="mt-6 flex flex-col sm:flex-row gap-3"
            data-testid="verify-form"
          >
            <div className="flex-1">
              <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
                Verification code
              </Label>
              <Input
                value={code}
                onChange={(e) => setCode(e.target.value.toUpperCase())}
                placeholder="e.g. 3B893A55D979"
                data-testid="verify-code-input"
                className="mt-1 rounded-sm font-mono uppercase tracking-wider"
                autoFocus
                required
                maxLength={32}
              />
            </div>
            <Button
              type="submit"
              disabled={busy}
              data-testid="verify-submit"
              className="rounded-sm bg-emerald-700 hover:bg-emerald-600 h-11 mt-1 sm:mt-6 sm:px-8"
            >
              {busy ? <Loader2 className="w-4 h-4 animate-spin" /> : "Verify"}
            </Button>
          </form>
        </Card>

        {data && (
          <Card
            data-testid="verify-result-card"
            className="bg-white border border-emerald-200 p-8 md:p-10 rounded-sm shadow-sm fade-up"
          >
            <div className="flex items-center gap-3">
              <BadgeCheck className="w-7 h-7 text-emerald-600" />
              <p className="text-xs uppercase tracking-[0.2em] text-emerald-800 font-semibold">
                Verified Grade Sheet
              </p>
            </div>
            <p className="font-mono text-xs text-stone-500 mt-2">
              Code: <span className="font-semibold text-stone-700">{data.gs_hash}</span>
            </p>

            <h2 className="font-display text-3xl mt-6">
              {data.student?.name || "—"}
            </h2>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-6 mt-6 font-mono text-sm">
              <Field label="Roll No." value={data.student?.roll_no} />
              <Field label="Enrollment" value={data.student?.enroll_no} />
              <Field label="Father's Name" value={data.student?.father_name} />
              <Field label="Programme" value={data.student?.program} />
              <Field label="Branch" value={data.student?.branch} />
              <Field label="Batch" value={data.student?.batch} />
            </div>

            <div className="mt-10 border-t border-stone-200 pt-6">
              <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
                Result · {data.result?.semester} Semester ·{" "}
                {data.result?.exam_session}
              </p>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-6 mt-4 font-mono text-sm">
                <Field label="SGPA" value={data.result?.sgpa} />
                <Field label="CGPA" value={data.result?.cgpa} />
                <Field label="Earned Credits" value={data.result?.earned_credits} />
                <Field label="Result" value={data.result?.result} />
              </div>
              {data.result?.remark && (
                <p className="mt-4 text-xs text-stone-600 font-mono">
                  Remark: {data.result.remark}
                </p>
              )}
            </div>

            {(data.result?.subjects || []).length > 0 && (
              <div className="mt-8 overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-xs uppercase tracking-[0.15em] text-stone-500 border-b border-stone-200">
                      <th className="text-left p-2 font-semibold">Code</th>
                      <th className="text-left p-2 font-semibold">Subject</th>
                      <th className="text-right p-2 font-semibold">Cr.</th>
                      <th className="text-right p-2 font-semibold">Grade</th>
                      <th className="text-right p-2 font-semibold">GP</th>
                    </tr>
                  </thead>
                  <tbody className="font-mono">
                    {data.result.subjects.map((s, i) => (
                      <tr
                        key={i}
                        className="border-t border-stone-100"
                        data-testid={`verify-subject-row-${i}`}
                      >
                        <td className="p-2">{s.code}</td>
                        <td className="p-2 sans">{s.name}</td>
                        <td className="p-2 text-right">{s.credits}</td>
                        <td className="p-2 text-right">{s.grade}</td>
                        <td className="p-2 text-right">{s.grade_points}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        )}
      </main>
      <PortalFooter />
    </div>
  );
}

const Field = ({ label, value }) => (
  <div>
    <p className="text-xs text-stone-500 uppercase tracking-[0.15em]">
      {label}
    </p>
    <p className="mt-1">{value || "—"}</p>
  </div>
);
