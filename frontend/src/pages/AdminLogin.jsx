import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "../lib/auth";
import { fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Input } from "../components/ui/input";
import { Label } from "../components/ui/label";
import { toast } from "sonner";
import { Loader2, ShieldCheck } from "lucide-react";

export default function AdminLogin() {
  const { user, ready, login } = useAuth();
  const navigate = useNavigate();
  const [email, setEmail] = useState("admin@gbpiet.ac.in");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (ready && user) navigate("/admin");
  }, [ready, user, navigate]);

  const submit = async (e) => {
    e.preventDefault();
    setLoading(true);
    try {
      await login(email, password);
      toast.success("Welcome back");
      navigate("/admin", { replace: true });
    } catch (err) {
      toast.error(fmtError(err));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen grid grid-cols-1 md:grid-cols-2 bg-stone-50">
      {/* Left visual panel */}
      <div className="hidden md:flex flex-col justify-between bg-indigo-950 text-white p-12 relative overflow-hidden">
        <div className="absolute inset-0 opacity-10 paper-surface" />
        <div className="relative">
          <p className="text-xs uppercase tracking-[0.2em] font-semibold opacity-70">
            GBPIET · Examination Office
          </p>
          <h1 className="font-display text-4xl mt-3 leading-tight">
            Sign in to the
            <br /> Result Asterisk Portal.
          </h1>
        </div>
        <div className="relative">
          <p className="font-display text-xl">“Tamaso ma jyotirgamaya”</p>
          <p className="text-xs opacity-60 mt-1 font-mono">
            Lead us from darkness to light.
          </p>
        </div>
      </div>

      {/* Form */}
      <div className="flex items-center justify-center px-6 py-16">
        <form
          onSubmit={submit}
          className="w-full max-w-sm fade-up"
          data-testid="admin-login-form"
        >
          <ShieldCheck className="w-8 h-8 text-indigo-950" strokeWidth={1.5} />
          <h2 className="font-display text-3xl mt-6">Administrator</h2>
          <p className="text-stone-600 text-sm mt-2">
            Sign in with your institute credentials to manage uploads and
            transcripts.
          </p>

          <div className="space-y-4 mt-8">
            <div>
              <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
                Email
              </Label>
              <Input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                data-testid="admin-email-input"
                className="mt-1 rounded-sm"
                required
              />
            </div>
            <div>
              <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
                Password
              </Label>
              <Input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                data-testid="admin-password-input"
                className="mt-1 rounded-sm"
                required
              />
            </div>
          </div>

          <Button
            type="submit"
            disabled={loading}
            data-testid="admin-login-submit"
            className="w-full mt-8 rounded-sm bg-indigo-950 hover:bg-indigo-900 h-11"
          >
            {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : "Sign in"}
          </Button>

          <p className="text-xs text-stone-500 mt-6 font-mono">
            Default: admin@gbpiet.ac.in · Admin@2026
          </p>
        </form>
      </div>
    </div>
  );
}
