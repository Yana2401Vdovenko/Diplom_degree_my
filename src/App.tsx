import { BrowserRouter as Router, Routes, Route, Navigate } from "react-router-dom";
import { useEffect, useState } from "react";
import Login from "./pages/Login";
import Register from "./pages/Register";
import Dashboard from "./pages/RoleDashboard";
import Workload from "./pages/Workload";
import Teachers from "./pages/Teachers";
import Disciplines from "./pages/Disciplines";
import Reports from "./pages/Reports";
import Department from "./pages/Department";
import Help from "./pages/Help";
import { supabase } from "./utils/supabaseClient";
import GlobalNotificationBell from "./components/GlobalNotificationBell";
import "./styles/GlobalNotifications.css";
import { ensureCurrentUserDepartment } from "./utils/currentDepartment";
import { resolveAuthenticatedRole } from "./utils/authRole";

function ProtectedRoute({ children }: { children: React.ReactNode }) {
  const [loading, setLoading] = useState(true);
  const [authenticated, setAuthenticated] = useState(false);

  useEffect(() => {
    let mounted = true;
    const syncDepartment = (user: Parameters<typeof ensureCurrentUserDepartment>[0]) => {
      if (!user) return;
      // Defer Auth calls outside onAuthStateChange to avoid holding the Auth lock.
      window.setTimeout(() => {
        void resolveAuthenticatedRole(user).then(async (resolved) => {
          if (resolved.role) localStorage.setItem("userRole", resolved.role === "head" ? "Завідувач кафедри" : resolved.role === "teacher" ? "Викладач" : "Лаборант");
          if (!resolved.departmentCode || resolved.role === "laborant") await ensureCurrentUserDepartment(user);
        }).catch((error: unknown) => {
          if (mounted) console.error("Не вдалося синхронізувати роль/кафедру користувача:", error);
        });
      }, 0);
    };
    supabase.auth.getSession().then(({ data }) => {
      setAuthenticated(!!data.session);
      syncDepartment(data.session?.user ?? null);
      setLoading(false);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setAuthenticated(!!session);
      syncDepartment(session?.user ?? null);
    });

    return () => { mounted = false; subscription.unsubscribe(); };
  }, []);

  if (loading) return <div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100vh" }}>Завантаження…</div>;
  if (!authenticated) return <Navigate to="/" replace />;
  return <><GlobalNotificationBell />{children}</>;
}

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Login />} />
        <Route path="/register" element={<Register />} />
        <Route path="/dashboard" element={<ProtectedRoute><Dashboard /></ProtectedRoute>} />
        <Route path="/workload" element={<ProtectedRoute><Workload /></ProtectedRoute>} />
        <Route path="/teachers" element={<ProtectedRoute><Teachers /></ProtectedRoute>} />
        <Route path="/disciplines" element={<ProtectedRoute><Disciplines /></ProtectedRoute>} />
        <Route path="/department" element={<ProtectedRoute><Department /></ProtectedRoute>} />
        <Route path="/reports" element={<ProtectedRoute><Reports /></ProtectedRoute>} />
        <Route path="/help" element={<ProtectedRoute><Help /></ProtectedRoute>} />
        <Route path="*" element={<h2 style={{ textAlign: "center", marginTop: "50px" }}>404 - Сторінку не знайдено</h2>} />
      </Routes>
    </Router>
  );
}

export default App;
