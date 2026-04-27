import axios from "axios";

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export const api = axios.create({
  baseURL: API,
  withCredentials: true,
});

// Attach Bearer token if present (admin)
api.interceptors.request.use((config) => {
  const tok = localStorage.getItem("admin_token");
  if (tok) config.headers.Authorization = `Bearer ${tok}`;
  return config;
});

// Auto-redirect to /admin/login when the session expires mid-action.
api.interceptors.response.use(
  (r) => r,
  (err) => {
    const status = err?.response?.status;
    const detail = err?.response?.data?.detail;
    const isExpired =
      status === 401 &&
      typeof detail === "string" &&
      /expired|invalid token|not authenticated|user not found/i.test(detail);
    if (isExpired && typeof window !== "undefined") {
      const onAdmin = window.location.pathname.startsWith("/admin");
      if (onAdmin && window.location.pathname !== "/admin/login") {
        localStorage.removeItem("admin_token");
        // Toast through console — sonner not imported here.
        console.warn("Session expired — redirecting to login.");
        window.location.replace("/admin/login");
      }
    }
    return Promise.reject(err);
  },
);

export const fmtError = (e) => {
  const detail = e?.response?.data?.detail;
  if (detail == null) return e?.message || "Something went wrong";
  if (typeof detail === "string") return detail;
  if (Array.isArray(detail))
    return detail.map((x) => (x?.msg ? x.msg : JSON.stringify(x))).join(" ");
  return JSON.stringify(detail);
};

export default api;
