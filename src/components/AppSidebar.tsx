import { FiBookOpen, FiHome, FiHelpCircle, FiLogOut, FiUsers } from "react-icons/fi";
import { FaUniversity } from "react-icons/fa";
import { useNavigate } from "react-router-dom";
import { useEffect, useState } from "react";
import logo from "../assets/images/logo.png";
import "../styles/AppSidebar.css";
import "../styles/UnifiedSidebar.css";
import { clearCurrentUser, currentUserName } from "../utils/currentUser";
import { supabase } from "../utils/supabaseClient";
import { CURRENT_DEPARTMENT_CHANGED_EVENT, ensureCurrentUserDepartment, findDepartmentFacultyCode } from "../utils/currentDepartment";

type Section = "dashboard" | "workload" | "teachers" | "help" | "disciplines" | "department";

export default function AppSidebar({ active }: { active: Section }) {
  const navigate = useNavigate();
  const role = localStorage.getItem("userRole") ?? "Лаборант";
  const name = currentUserName();
  const [departmentName, setDepartmentName] = useState("");
  const [facultyName, setFacultyName] = useState("");
  const go = (path: string) => navigate(path);
  useEffect(() => {
    let active = true;
    const loadDepartment = async () => {
      const { data: auth } = await supabase.auth.getUser();
      if (!auth.user) return;
      try {
        const { departmentCode } = await ensureCurrentUserDepartment(auth.user);
        if (!departmentCode) return;
        const { data: department } = await supabase.from("Кафедра").select("*").eq("Код_кафедри", departmentCode).maybeSingle();
        if (!department) return;
        const departmentRow = department as Record<string, unknown>;
        const facultyCode = findDepartmentFacultyCode(departmentRow);
        const { data: faculty } = facultyCode
          ? await supabase.from("Факультет_ННІ").select("*").eq("Код_факультету_нні", facultyCode).maybeSingle()
          : { data: null };
        const facultyRow = faculty as Record<string, unknown> | null;
        const facultyName = String(facultyRow?.["Назва_факультету_нні"] ?? facultyRow?.["Назва_факультету"] ?? facultyRow?.["Назва"] ?? "").trim();
        const departmentName = String(departmentRow["Назва_кафедри"] ?? "").trim();
        if (active) {
          setDepartmentName(departmentName);
          setFacultyName(facultyName);
        }
      } catch (error) { console.error("Не вдалося завантажити прив’язку користувача:", error); }
    };
    void loadDepartment();
    window.addEventListener(CURRENT_DEPARTMENT_CHANGED_EVENT, loadDepartment);
    return () => { active = false; window.removeEventListener(CURRENT_DEPARTMENT_CHANGED_EVENT, loadDepartment); };
  }, []);
  const initials = name.trim().split(/[\s.@_-]+/).filter(Boolean).slice(0, 2).map((part) => part[0]?.toLocaleUpperCase("uk-UA")).join("") || "К";
  const normalizedDepartmentName = departmentName.replace(/^(?:кафедра\s+)+/i, "").trim();
  const departmentLabel = normalizedDepartmentName ? `Кафедра ${normalizedDepartmentName}` : "Кафедра комп'ютерних наук";
  const facultyLabel = /інформац.*технолог/i.test(facultyName) ? "ФІТ" : facultyName || "Факультет ІТ";

  return <aside className="app-sidebar">
    <div className="app-sidebar-brand"><img src={logo} alt="Герб НУБіП" /><div className="app-sidebar-brand-copy"><b>НУБіП</b><small title={departmentLabel}>{departmentLabel}</small><small className="app-sidebar-faculty" title={facultyLabel}>{facultyLabel}</small></div></div>
    <nav className="app-sidebar-nav">
      <button className={active === "dashboard" ? "active" : ""} onClick={() => go("/dashboard")}><FiHome aria-hidden="true" /> <span>Головна</span></button>
      <span className="app-sidebar-section-label">УПРАВЛІННЯ</span>
      <button className={active === "department" ? "active" : ""} onClick={() => go("/department")}><FaUniversity aria-hidden="true" /> <span>Кафедра</span></button>
      <button className={active === "teachers" ? "active" : ""} onClick={() => go("/teachers")}><FiUsers aria-hidden="true" /> <span>Викладачі</span></button>
      <button className={active === "workload" || active === "disciplines" ? "active" : ""} onClick={() => go("/workload")}><FiBookOpen aria-hidden="true" /> <span>Навантаження</span></button>
      <span className="app-sidebar-section-label">ДОДАТКОВО</span>
      <button className={active === "help" ? "active" : ""} onClick={() => go("/help")}><FiHelpCircle aria-hidden="true" /> <span>Довідка</span></button>
    </nav>
    <div className="app-sidebar-footer">
      <div className="app-sidebar-user"><span className="app-sidebar-avatar" aria-hidden="true">{initials}</span><div className="app-sidebar-user-copy"><b title={name}>{name}</b><small>{role}</small></div>
        <button className="app-sidebar-logout" title="Вийти" aria-label="Вийти" onClick={() => { clearCurrentUser(); void supabase.auth.signOut(); go("/"); }}><FiLogOut aria-hidden="true" /></button>
      </div>
    </div>
  </aside>;
}
