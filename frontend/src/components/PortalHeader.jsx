/**
 * Branded GBPIET header used across public portal pages.
 * Shows the institute logo + name on the left and a slot for right-side
 * navigation actions (e.g. "Back home", quick links).
 */
export default function PortalHeader({ right = null }) {
  return (
    <header
      data-testid="portal-header"
      className="bg-white border-b border-stone-200 print:hidden"
    >
      <div className="max-w-7xl mx-auto px-6 py-3 flex items-center justify-between gap-4">
        <a
          href="/"
          className="flex items-center gap-3"
          data-testid="portal-header-home"
        >
          <img
            src="/gbpiet_logo.png"
            alt="GBPIET"
            className="w-11 h-11 object-contain"
          />
          <div className="leading-tight">
            <p className="font-display text-lg text-stone-900">GBPIET</p>
            <p className="text-[10px] uppercase tracking-[0.18em] text-stone-500 font-semibold">
              Examination Cell · Pauri Garhwal
            </p>
          </div>
        </a>
        <div className="flex items-center gap-3">{right}</div>
      </div>
    </header>
  );
}
