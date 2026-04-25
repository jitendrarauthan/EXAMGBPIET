import { Link } from "react-router-dom";
import { Button } from "../components/ui/button";
import { ShieldCheck, GraduationCap, FileText, ArrowRight, BadgeCheck } from "lucide-react";

export default function Landing() {
  return (
    <div className="min-h-screen bg-stone-50">
      {/* Top ribbon */}
      <header className="border-b border-stone-200 bg-white">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-sm bg-indigo-950 flex items-center justify-center text-white font-bold">
              G
            </div>
            <div>
              <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
                GBPIET · Pauri Garhwal
              </p>
              <h1 className="font-display text-lg leading-tight">
                Result Asterisk Portal
              </h1>
            </div>
          </div>
          <span className="text-xs text-stone-500 hidden md:block font-mono">
            v1 · Examination Cell
          </span>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-16">
        <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
          Examination Office · Result Workflow
        </p>
        <h1 className="font-display text-4xl md:text-6xl mt-4 max-w-3xl leading-[1.05]">
          A precise, audited workflow for marking back subjects on Tabulation
          Charts &amp; Grade Sheets.
        </h1>
        <p className="text-base md:text-lg text-stone-600 mt-6 max-w-2xl leading-relaxed">
          Upload the original TC, GS PDFs and the SEM Excel sheet. The portal
          identifies highlighted (back) subjects per student and regenerates
          starred (<span className="font-mono font-semibold">*</span>) versions
          of every transcript — ready for download or distribution.
        </p>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mt-14">
          <Link
            to="/admin/login"
            data-testid="enter-admin-portal"
            className="group bg-white border border-stone-200 p-8 hover:border-indigo-950 transition-all"
          >
            <ShieldCheck className="w-7 h-7 text-indigo-950" strokeWidth={1.5} />
            <p className="text-xs uppercase tracking-[0.2em] text-stone-500 mt-6 font-semibold">
              For the Examination Cell
            </p>
            <h2 className="font-display text-3xl mt-2">Admin Portal</h2>
            <p className="text-stone-600 text-sm mt-3 leading-relaxed">
              Upload TC / GS PDFs and SEM_ Excel sheets per programme, branch
              and batch. Browse generated starred files and audit student
              records.
            </p>
            <div className="mt-8 inline-flex items-center gap-2 text-indigo-950 font-medium group-hover:gap-3 transition-all">
              Sign in <ArrowRight className="w-4 h-4" />
            </div>
          </Link>

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
              are back subjects. Just enter your roll number — no password needed.
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
              tied to the student record.
            </p>
            <div className="mt-8 inline-flex items-center gap-2 text-emerald-700 font-medium group-hover:gap-3 transition-all">
              Open verifier <ArrowRight className="w-4 h-4" />
            </div>
          </Link>
        </div>

        <div className="mt-20 grid grid-cols-1 md:grid-cols-3 gap-8 border-t border-stone-200 pt-10">
          {[
            ["Bulk uploads", "Upload semester PDFs and SEM Excel together; we deduplicate and version each batch."],
            ["Asterisk auditing", "Yellow / blue highlighted cells in SEM_ sheets translate to a precise * marker on the matching subject."],
            ["Per-student exports", "Download single-student GS or full-batch TC PDFs in the original institute layout."],
          ].map(([t, d]) => (
            <div key={t}>
              <FileText className="w-5 h-5 text-stone-700" strokeWidth={1.5} />
              <p className="font-display text-xl mt-3">{t}</p>
              <p className="text-stone-600 text-sm mt-1 leading-relaxed">{d}</p>
            </div>
          ))}
        </div>
      </main>

      <footer className="border-t border-stone-200 mt-20">
        <div className="max-w-6xl mx-auto px-6 py-6 text-xs text-stone-500 flex items-center justify-between font-mono">
          <span>Govind Ballabh Pant Institute of Engineering &amp; Technology</span>
          <span>Pauri Garhwal · Uttarakhand</span>
        </div>
      </footer>
    </div>
  );
}
