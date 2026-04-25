import { useState } from "react";
import { useNavigate } from "react-router-dom";
import { api, fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Input } from "../components/ui/input";
import { Label } from "../components/ui/label";
import { Card } from "../components/ui/card";
import { toast } from "sonner";
import { GraduationCap, Loader2 } from "lucide-react";

export default function StudentLogin() {
  const nav = useNavigate();
  const [roll, setRoll] = useState("");
  const [dob, setDob] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    if (!roll.trim()) return toast.error("Enter your roll number");
    setBusy(true);
    try {
      const { data } = await api.post("/student/login", {
        roll_no: roll.trim(),
        dob: dob.trim(),
      });
      sessionStorage.setItem("student_data", JSON.stringify(data));
      nav("/student/results");
    } catch (e) {
      toast.error(fmtError(e));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      className="min-h-screen flex items-center justify-center px-6 py-12 relative"
      style={{
        backgroundImage:
          "linear-gradient(rgba(28, 25, 23, 0.78), rgba(28, 25, 23, 0.78)), url('https://images.pexels.com/photos/19193975/pexels-photo-19193975.jpeg?auto=compress&cs=tinysrgb&dpr=2&h=650&w=940')",
        backgroundSize: "cover",
        backgroundPosition: "center",
      }}
    >
      <Card className="w-full max-w-md rounded-sm border-stone-200 shadow-xl bg-white p-10 fade-up">
        <GraduationCap className="w-8 h-8 text-red-900" strokeWidth={1.5} />
        <p className="text-xs uppercase tracking-[0.2em] text-stone-500 mt-6 font-semibold">
          GBPIET · Student Portal
        </p>
        <h1 className="font-display text-3xl mt-1">Check your results</h1>
        <p className="text-stone-600 text-sm mt-2">
          Enter your university roll number and date of birth.
        </p>

        <form onSubmit={submit} className="mt-8 space-y-4" data-testid="student-login-form">
          <div>
            <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
              University Roll No
            </Label>
            <Input
              value={roll}
              onChange={(e) => setRoll(e.target.value)}
              placeholder="e.g. 230090107001"
              data-testid="student-roll-input"
              className="mt-1 rounded-sm font-mono"
              required
            />
          </div>
          <div>
            <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
              Date of Birth
            </Label>
            <Input
              type="date"
              value={dob}
              onChange={(e) => setDob(e.target.value)}
              data-testid="student-dob-input"
              className="mt-1 rounded-sm font-mono"
            />
          </div>
          <Button
            type="submit"
            disabled={busy}
            data-testid="student-login-submit"
            className="w-full mt-2 rounded-sm bg-red-900 hover:bg-red-800 h-11"
          >
            {busy ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              "View results"
            )}
          </Button>
        </form>
        <p className="text-xs text-stone-500 mt-6">
          Your DOB is verified against the institute record. If your DOB has
          not been registered yet, contact the examination cell.
        </p>
      </Card>
    </div>
  );
}
