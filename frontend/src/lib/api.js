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

export const fmtError = (e) => {
  const detail = e?.response?.data?.detail;
  if (detail == null) return e?.message || "Something went wrong";
  if (typeof detail === "string") return detail;
  if (Array.isArray(detail))
    return detail.map((x) => (x?.msg ? x.msg : JSON.stringify(x))).join(" ");
  return JSON.stringify(detail);
};

export default api;
