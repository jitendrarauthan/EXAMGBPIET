import { BrowserRouter, Routes, Route, Navigate } from "react-router-dom";
import "@/App.css";
import { AuthProvider } from "@/lib/auth";
import { Toaster } from "@/components/ui/sonner";

import Landing from "@/pages/Landing";
import AdminLogin from "@/pages/AdminLogin";
import AdminLayout from "@/pages/AdminLayout";
import AdminOverview from "@/pages/AdminOverview";
import AdminUpload from "@/pages/AdminUpload";
import AdminFiles from "@/pages/AdminFiles";
import AdminStudents from "@/pages/AdminStudents";
import StudentLogin from "@/pages/StudentLogin";
import StudentResults from "@/pages/StudentResults";

function App() {
  return (
    <AuthProvider>
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<Landing />} />
          <Route path="/admin/login" element={<AdminLogin />} />
          <Route path="/admin" element={<AdminLayout />}>
            <Route index element={<AdminOverview />} />
            <Route path="upload" element={<AdminUpload />} />
            <Route path="files" element={<AdminFiles />} />
            <Route path="students" element={<AdminStudents />} />
          </Route>
          <Route path="/student" element={<StudentLogin />} />
          <Route path="/student/results" element={<StudentResults />} />
          <Route path="*" element={<Navigate to="/" replace />} />
        </Routes>
      </BrowserRouter>
      <Toaster richColors position="top-right" />
    </AuthProvider>
  );
}

export default App;
