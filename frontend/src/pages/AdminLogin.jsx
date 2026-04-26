import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { useAuth } from "../lib/auth";
import { fmtError } from "../lib/api";
import { Button } from "../components/ui/button";
import { Input } from "../components/ui/input";
import { Label } from "../components/ui/label";
import { toast } from "sonner";
import { Loader2, ShieldCheck, Mail } from "lucide-react";

export default function AdminLogin() {
  const { user, ready, login, verifyOtp } = useAuth();
  const navigate = useNavigate();
  const [email, setEmail] = useState("admin@gbpiet.ac.in");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);

  // OTP step state
  const [stage, setStage] = useState("password"); // "password" | "otp"
  const [challengeId, setChallengeId] = useState("");
  const [sentTo, setSentTo] = useState("");
  const [otp, setOtp] = useState("");

  useEffect(() => {
    if (ready && user) navigate("/admin");
  }, [ready, user, navigate]);

  const submitPassword = async (e) => {
    e.preventDefault();
    setLoading(true);
    try {
      const data = await login(email, password);
      if (data && data.otp_required) {
        setChallengeId(data.challenge_id);
        setSentTo(data.sent_to || "");
        setStage("otp");
        toast.success("OTP sent to your authorized email");
      } else {
        toast.success("Welcome back");
        navigate("/admin", { replace: true });
      }
    } catch (err) {
      toast.error(fmtError(err));
    } finally {
      setLoading(false);
    }
  };

  const submitOtp = async (e) => {
    e.preventDefault();
    setLoading(true);
    try {
      await verifyOtp(challengeId, otp.trim());
      toast.success("Verified — welcome back");
      navigate("/admin", { replace: true });
    } catch (err) {
      toast.error(fmtError(err));
    } finally {
      setLoading(false);
    }
  };

  const restart = () => {
    setStage("password");
    setChallengeId("");
    setSentTo("");
    setOtp("");
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
        {stage === "password" ? (
          <form
            onSubmit={submitPassword}
            className="w-full max-w-sm fade-up"
            data-testid="admin-login-form"
          >
            <ShieldCheck className="w-8 h-8 text-indigo-950" strokeWidth={1.5} />
            <h2 className="font-display text-3xl mt-6">Administrator</h2>
            <p className="text-stone-600 text-sm mt-2">
              Sign in with your institute credentials. A one-time password will
              be emailed to the authorized examination cell address.
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
              {loading ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                "Continue"
              )}
            </Button>
          </form>
        ) : (
          <form
            onSubmit={submitOtp}
            className="w-full max-w-sm fade-up"
            data-testid="admin-otp-form"
          >
            <Mail className="w-8 h-8 text-indigo-950" strokeWidth={1.5} />
            <h2 className="font-display text-3xl mt-6">Verify it’s you</h2>
            <p className="text-stone-600 text-sm mt-2">
              We’ve emailed a 6-digit code to{" "}
              <span className="font-mono text-stone-800">{sentTo || "the authorized address"}</span>.
              Enter it below to complete sign-in.
            </p>

            <div className="space-y-4 mt-8">
              <div>
                <Label className="text-xs uppercase tracking-[0.15em] text-stone-500">
                  One-time password
                </Label>
                <Input
                  type="text"
                  inputMode="numeric"
                  pattern="[0-9]*"
                  maxLength={6}
                  autoFocus
                  value={otp}
                  onChange={(e) => setOtp(e.target.value.replace(/\D/g, ""))}
                  data-testid="admin-otp-input"
                  className="mt-1 rounded-sm font-mono tracking-[0.5em] text-center text-lg"
                  required
                />
              </div>
            </div>

            <Button
              type="submit"
              disabled={loading || otp.length < 4}
              data-testid="admin-otp-submit"
              className="w-full mt-8 rounded-sm bg-indigo-950 hover:bg-indigo-900 h-11"
            >
              {loading ? (
                <Loader2 className="w-4 h-4 animate-spin" />
              ) : (
                "Verify & sign in"
              )}
            </Button>

            <button
              type="button"
              onClick={restart}
              data-testid="admin-otp-back"
              className="w-full mt-4 text-xs text-stone-500 hover:text-stone-800 underline-offset-2 hover:underline"
            >
              Use a different account
            </button>
          </form>
        )}
      </div>
    </div>
  );
}
