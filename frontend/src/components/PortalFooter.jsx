/**
 * Small footer displayed at the bottom of every public-facing portal page
 * (Landing, Student login/results, Verify). Carries the credit line
 * requested by the institute.
 */
export default function PortalFooter() {
  return (
    <footer
      data-testid="portal-footer"
      className="mt-16 border-t border-stone-200 bg-white print:hidden"
    >
      <div className="max-w-7xl mx-auto px-6 py-6 flex flex-col md:flex-row items-center justify-between gap-2 text-xs text-stone-500 font-mono">
        <p>
          © {new Date().getFullYear()} GBPIET, Pauri Garhwal · Examination Cell
        </p>
        <p>
          Designed and maintained by{" "}
          <span className="font-semibold text-stone-700">
            Dr. Jitendra Singh Rauthan
          </span>
        </p>
      </div>
    </footer>
  );
}
