import { Link } from "react-router-dom";
import { GraduationCap, FileText, ArrowRight, BadgeCheck } from "lucide-react";
import PortalHeader from "../components/PortalHeader";
import PortalFooter from "../components/PortalFooter";

export default function Landing() {
  return (
    <div className="min-h-screen bg-stone-50 flex flex-col">
      <PortalHeader />

      <main className="max-w-6xl mx-auto px-6 py-16 flex-1 w-full">
        <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
          Examination Cell · GBPIET
        </p>
        <h1 className="font-display text-4xl md:text-6xl mt-4 max-w-3xl leading-[1.05]">
          Verified semester-wise results &amp; authenticated grade sheets.
        </h1>
        <p className="text-base md:text-lg text-stone-600 mt-6 max-w-2xl leading-relaxed">
          Students can look up their complete semester-wise records using only
          their roll number. Recruiters &amp; institutions can verify a printed
          GBPIET Grade Sheet by entering the alphanumeric code printed below
          the barcode on the document.
        </p>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mt-14">
          <Link
            to="/student"
            data-testid="enter-student-portal"
            className="group bg-white border border-stone-200 p-8 hover:border-red-900 transition-all"
          >
            <GraduationCap className="w-7 h-7 text-red-900" strokeWidth={1.5} />
            <p className="text-xs uppercase tracking-[0.2em] text-stone-500 mt-6 font-semibold">
              For Students
            </p>
            <h2 className="font-display text-3xl mt-2">Student Portal</h2>
            <p className="text-stone-600 text-sm mt-3 leading-relaxed">
              View your complete semester-wise results. Subjects marked with
              <span className="subject-back mx-1 text-xs">*</span>
              are back papers cleared in a later attempt. Just enter your roll
              number — no password needed.
            </p>
            <div className="mt-8 inline-flex items-center gap-2 text-red-900 font-medium group-hover:gap-3 transition-all">
              Check results <ArrowRight className="w-4 h-4" />
            </div>
          </Link>

          <Link
            to="/verify"
            data-testid="enter-verify-portal"
            className="group bg-white border border-stone-200 p-8 hover:border-emerald-700 transition-all"
          >
            <BadgeCheck className="w-7 h-7 text-emerald-700" strokeWidth={1.5} />
            <p className="text-xs uppercase tracking-[0.2em] text-stone-500 mt-6 font-semibold">
              For Verifiers
            </p>
            <h2 className="font-display text-3xl mt-2">Verify a Grade Sheet</h2>
            <p className="text-stone-600 text-sm mt-3 leading-relaxed">
              Confirm the authenticity of any GBPIET Grade Sheet by entering
              the printed verification code. Each GS carries a unique code
              tied to the student record and updates if the GS is reissued.
            </p>
            <div className="mt-8 inline-flex items-center gap-2 text-emerald-700 font-medium group-hover:gap-3 transition-all">
              Open verifier <ArrowRight className="w-4 h-4" />
            </div>
          </Link>
        </div>

        <div className="mt-20 grid grid-cols-1 md:grid-cols-3 gap-8 border-t border-stone-200 pt-10">
          {[
            [
              "Roll-number lookup",
              "Students authenticate with just their University Roll Number — semester-wise SGPA, CGPA and earned credits available instantly.",
            ],
            [
              "Authenticated GS",
              "Every printed Grade Sheet carries a unique alphanumeric code. Anyone can use the verifier to confirm it was issued by GBPIET.",
            ],
            [
              "Print-friendly",
              "The student portal layout adapts to A4 portrait printing — each semester begins with the student's profile at the top of the page.",
            ],
          ].map(([t, d]) => (
            <div key={t}>
              <FileText className="w-5 h-5 text-stone-700" strokeWidth={1.5} />
              <p className="font-display text-xl mt-3">{t}</p>
              <p className="text-stone-600 text-sm mt-1 leading-relaxed">
                {d}
              </p>
            </div>
          ))}
        </div>
      </main>

      <PortalFooter />
    </div>
  );
}
