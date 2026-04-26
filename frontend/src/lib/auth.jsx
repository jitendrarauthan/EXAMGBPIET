import { createContext, useContext, useEffect, useState } from "react";
import { api } from "./api";

const AuthCtx = createContext(null);

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null); // null = unknown / loading
  const [ready, setReady] = useState(false);

  useEffect(() => {
    // Only check auth on admin routes; student routes don't need it.
    const path =
      typeof window !== "undefined" ? window.location.pathname : "/";
    if (!path.startsWith("/admin")) {
      setUser(false);
      setReady(true);
      return;
    }
    api
      .get("/auth/me")
      .then((r) => setUser(r.data))
      .catch(() => setUser(false))
      .finally(() => setReady(true));
  }, []);

  const login = async (email, password) => {
    const { data } = await api.post("/auth/login", { email, password });
    // 2FA: backend returns an OTP challenge. The caller must redeem it
    // via verifyOtp() before a JWT is issued.
    if (data && data.otp_required) {
      return data; // { otp_required, challenge_id, expires_in, sent_to }
    }
    if (data.token) localStorage.setItem("admin_token", data.token);
    setUser(data);
    setReady(true);
    return data;
  };

  const verifyOtp = async (challenge_id, otp) => {
    const { data } = await api.post("/auth/verify-otp", { challenge_id, otp });
    if (data.token) localStorage.setItem("admin_token", data.token);
    setUser(data);
    setReady(true);
    return data;
  };

  const logout = async () => {
    try {
      await api.post("/auth/logout");
    } catch {}
    localStorage.removeItem("admin_token");
    setUser(false);
  };

  return (
    <AuthCtx.Provider value={{ user, ready, login, verifyOtp, logout, setUser }}>
      {children}
    </AuthCtx.Provider>
  );
}

export const useAuth = () => useContext(AuthCtx);
