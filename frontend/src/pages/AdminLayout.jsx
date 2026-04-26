import { Outlet, NavLink, useNavigate } from "react-router-dom";
import { useAuth } from "../lib/auth";
import { useEffect } from "react";
import { Button } from "../components/ui/button";
import {
  LayoutDashboard,
  Upload,
  FilePlus2,
  Users,
  LogOut,
} from "lucide-react";

const NAV = [
  { to: "/admin", label: "Overview", icon: LayoutDashboard, end: true },
  { to: "/admin/upload", label: "Upload", icon: Upload },
  { to: "/admin/upload-mtech", label: "Upload (M.Tech)", icon: Upload },
  { to: "/admin/files", label: "Generated Files", icon: FilePlus2 },
  { to: "/admin/students", label: "Students", icon: Users },
];

export default function AdminLayout() {
  const { user, ready, logout } = useAuth();
  const nav = useNavigate();
  const hasToken =
    typeof window !== "undefined" && !!localStorage.getItem("admin_token");

  useEffect(() => {
    if (ready && !user && !hasToken) nav("/admin/login");
  }, [ready, user, hasToken, nav]);

  if (!ready)
    return (
      <div className="min-h-screen flex items-center justify-center text-stone-500">
        Loading…
      </div>
    );
  if (!user && !hasToken) return null;

  return (
    <div className="min-h-screen flex bg-stone-50">
      {/* Sidebar */}
      <aside className="w-64 border-r border-stone-200 bg-white flex flex-col">
        <div className="p-6 border-b border-stone-200">
          <p className="text-xs uppercase tracking-[0.2em] text-stone-500 font-semibold">
            GBPIET
          </p>
          <h1 className="font-display text-xl mt-1 leading-tight">
            Result Asterisk
          </h1>
          <p className="text-xs font-mono text-stone-500 mt-1">Admin Console</p>
        </div>
        <nav className="p-3 space-y-0.5 flex-1">
          {NAV.map(({ to, label, icon: Icon, end }) => (
            <NavLink
              key={to}
              to={to}
              end={end}
              data-testid={`nav-${label.toLowerCase().replace(/ /g, "-")}`}
              className={({ isActive }) =>
                `flex items-center gap-3 px-3 py-2 rounded-sm text-sm transition-colors ${
                  isActive
                    ? "bg-stone-900 text-white"
                    : "text-stone-700 hover:bg-stone-100"
                }`
              }
            >
              <Icon className="w-4 h-4" strokeWidth={1.5} />
              {label}
            </NavLink>
          ))}
        </nav>
        <div className="p-4 border-t border-stone-200">
          <p className="text-xs text-stone-500 font-mono truncate">
            {user.email}
          </p>
          <Button
            variant="outline"
            size="sm"
            onClick={async () => {
              await logout();
              nav("/admin/login");
            }}
            data-testid="admin-logout-btn"
            className="w-full mt-3 rounded-sm"
          >
            <LogOut className="w-3.5 h-3.5 mr-2" /> Sign out
          </Button>
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 overflow-auto">
        <Outlet />
      </main>
    </div>
  );
}
