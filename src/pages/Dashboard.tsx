import "../styles/Dashboard.css";
import "../styles/UnifiedSidebar.css";
import logo from "../assets/images/logo.png";
import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { currentUserName, clearCurrentUser } from "../utils/currentUser";
import AppSidebar from "../components/AppSidebar";
import { notify } from "../utils/notifications";
import { supabase } from "../utils/supabaseClient";
import { ensureCurrentUserDepartment, findDepartmentFacultyCode } from "../utils/currentDepartment";
import { FaHome, FaGraduationCap, FaUsers, FaChevronDown, FaSignOutAlt, FaQuestionCircle, FaBriefcase, FaFileAlt, FaClock } from "react-icons/fa";
import { ResponsiveContainer, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, PieChart, Pie, Cell } from "recharts";

type Row = Record<string, unknown>;
type WorkItem = { name: string; hours: number; color: string };
type DisciplineTotal = { code: string; name: string; hours: number };

const fetchAllRows = async (table: string, columns = "*") => {
  const rows: Row[] = [];
  for (let from = 0; ; from += 1000) {
    const result = await supabase.from(table).select(columns).range(from, from + 999);
    if (result.error) return { data: null as Row[] | null, error: result.error };
    const page = (result.data ?? []) as unknown as Row[];
    rows.push(...page);
    if (page.length < 1000) return { data: rows, error: null };
  }
};

const trim = (value: unknown) => String(value ?? "").trim();
const numeric = (value: unknown) => {
  const parsed = Number(trim(value).replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(parsed) ? parsed : 0;
};
const toDbYear = (value: unknown) => {
  const year = trim(value);
  const full = year.match(/^(\d{4})\/(\d{4})$/);
  if (full) return `${full[1]}/${full[2].slice(-2)}`;
  return year.match(/^\d{4}\/\d{2}$/) ? year : "";
};
const toDisplayYear = (value: unknown) => {
  const year = toDbYear(value);
  return year ? `${year.slice(0, 4)}/20${year.slice(-2)}` : "";
};
const totalHours = (row: Row) => {
  const storedTotal = numeric(row["Всього_годин"]);
  if (storedTotal !== 0) return storedTotal;
  return ["Лекції", "Лабораторні_заняття", "Практичні_заняття", "Семінарні_заняття", "Екзамен", "Залік", "Курсовий_проєкт", "Практика", "Дипломне_проєктування", "Рецензії"]
    .reduce((sum, column) => sum + numeric(row[column]), 0);
};
const numberText = (value: number) => new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 2 }).format(value);
const ukrainianCountWord = (value: number, forms: [string, string, string]) => {
  const remainder100 = Math.abs(value) % 100;
  if (remainder100 >= 11 && remainder100 <= 14) return forms[2];
  const remainder10 = Math.abs(value) % 10;
  if (remainder10 === 1) return forms[0];
  if (remainder10 >= 2 && remainder10 <= 4) return forms[1];
  return forms[2];
};
const workKinds: Array<{ name: string; column: string; color: string }> = [
  { name: "Лекції", column: "Лекції", color: "#3b82f6" },
  { name: "Лабораторні", column: "Лабораторні_заняття", color: "#8b5cf6" },
  { name: "Практичні", column: "Практичні_заняття", color: "#10b981" },
  { name: "Семінарські", column: "Семінарні_заняття", color: "#f59e0b" },
  { name: "Іспити й заліки", column: "Екзамен", color: "#ef4444" },
  { name: "Заліки", column: "Залік", color: "#06b6d4" },
  { name: "Курсові проєкти", column: "Курсовий_проєкт", color: "#ec4899" },
  { name: "Практика", column: "Практика", color: "#84cc16" },
  { name: "Дипломне проєктування", column: "Дипломне_проєктування", color: "#f97316" },
  { name: "Рецензії", column: "Рецензії", color: "#64748b" },
];

function Dashboard({ departmentName: defaultDepartmentName = "Кафедра" }: { departmentName?: string }) {
  const navigate = useNavigate();
  const [role, setRole] = useState(() => localStorage.getItem("userRole") ?? "Лаборант");
  const [profileName] = useState(currentUserName);
  const [workloadRows, setWorkloadRows] = useState<Row[]>([]);
  const [teacherRows, setTeacherRows] = useState<Row[]>([]);
  const [disciplineNames, setDisciplineNames] = useState<Record<string, string>>({});
  const [availableYears, setAvailableYears] = useState<string[]>([]);
  const [studyForms, setStudyForms] = useState<Array<{ code: string; name: string }>>([]);
  const [academicYear, setAcademicYear] = useState("");
  const [selectedFormCode, setSelectedFormCode] = useState("");
  const [departmentCode, setDepartmentCode] = useState("");
  const [departmentTitle, setDepartmentTitle] = useState(defaultDepartmentName);
  const [facultyTitle, setFacultyTitle] = useState("");
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    const savedRole = localStorage.getItem("userRole");
    if (savedRole) setRole(savedRole);
    let cancelled = false;
    const load = async () => {
      const [authResult, workloadResult, teachersResult, infoResult, disciplinesResult, departmentsResult, facultyResult, formsResult] = await Promise.all([
        supabase.auth.getUser(),
        fetchAllRows("Завантаження_навантаження"),
        fetchAllRows("Викладач"),
        fetchAllRows("Кафедральна_інформація_про_виклад"),
        fetchAllRows("Дисципліни", "Код_дисципліни,Назва_дисципліни"),
        fetchAllRows("Кафедра"),
        fetchAllRows("Факультет_ННІ"),
        fetchAllRows("Форма_навчання", "Код_форми_навчання,Назва_форми_навчання"),
      ]);
      const failed = [workloadResult, teachersResult, infoResult, disciplinesResult, departmentsResult, facultyResult, formsResult].find((result) => result.error)?.error;
      if (cancelled) return;
      if (failed) {
        setLoadError(failed.message);
        notify(`Не вдалося завантажити показники головної сторінки: ${failed.message}`, "error", "Головна");
        setLoading(false);
        return;
      }
      const rows = workloadResult.data ?? [];
      const infos = infoResult.data ?? [];
      const depts = departmentsResult.data ?? [];
      let binding: { departmentCode: string; facultyCode: string };
      try { binding = await ensureCurrentUserDepartment(authResult.data.user); }
      catch (error) {
        const message = error instanceof Error ? error.message : "Невідома помилка";
        setLoadError(`Не вдалося визначити кафедру користувача: ${message}`);
        notify(`Не вдалося визначити кафедру користувача: ${message}`, "error", "Головна");
        setLoading(false);
        return;
      }
      const selectedDepartment = binding.departmentCode;
      if (!selectedDepartment) {
        setLoadError("Для цього облікового запису не знайдено прив’язану кафедру.");
        setLoading(false);
        return;
      }
      const deptRows = depts.find((item) => trim(item["Код_кафедри"]) === selectedDepartment);
      const facultyCode = findDepartmentFacultyCode(deptRows);
      const facultyRows = facultyResult.data ?? [];
      const facultyRow = facultyRows.find((item) => trim(item["Код_факультету_нні"] ?? item["Код_факультету"]) === facultyCode);
      const resolvedFacultyTitle = trim(facultyRow?.["Назва_факультету_нні"] ?? facultyRow?.["Назва_факультету"] ?? facultyRow?.["Назва"]);
      const dbYears = [...new Set(rows.filter((row) => trim(row["Код_кафедри"]) === selectedDepartment)
        .map((row) => toDisplayYear(row["Рік"])).filter(Boolean))]
        .sort((left, right) => toDbYear(right).localeCompare(toDbYear(left)));
      setDepartmentCode(selectedDepartment);
      setDepartmentTitle(trim(deptRows?.["Назва_кафедри"]) || defaultDepartmentName);
      setFacultyTitle(resolvedFacultyTitle || "Факультет не вказано");
      setWorkloadRows(rows.filter((row) => trim(row["Код_кафедри"]) === selectedDepartment));
      const departmentTeacherCodes = new Set(infos.filter((row) => trim(row["Код_кафедри"]) === selectedDepartment).map((row) => trim(row["Код_викладача"])));
      setTeacherRows((teachersResult.data ?? []).filter((teacher) => departmentTeacherCodes.has(trim(teacher["Код_викладача"]))));
      setDisciplineNames(Object.fromEntries((disciplinesResult.data ?? []).map((row) => [trim(row["Код_дисципліни"]), trim(row["Назва_дисципліни"]) ])));
      setAvailableYears(dbYears);
      setAcademicYear(dbYears[0] ?? "");
      const forms = (formsResult.data ?? []).map((row) => ({ code: trim(row["Код_форми_навчання"]), name: trim(row["Назва_форми_навчання"]) })).filter((form) => form.code);
      setStudyForms(forms);
      setSelectedFormCode(forms[0]?.code ?? "");
      setLoading(false);
    };
    void load();
    return () => { cancelled = true; };
  }, [defaultDepartmentName]);

  const yearRows = useMemo(() => workloadRows.filter((row) =>
    trim(row["Код_кафедри"]) === departmentCode
      && toDbYear(row["Рік"]) === toDbYear(academicYear)
      && (!selectedFormCode || trim(row["Код_форми_навчання"]) === selectedFormCode)), [workloadRows, departmentCode, academicYear, selectedFormCode]);
  const total = yearRows.reduce((sum, row) => sum + totalHours(row), 0);
  const assigned = yearRows.filter((row) => Boolean(trim(row["Код_викладача"]))).reduce((sum, row) => sum + totalHours(row), 0);
  const vacant = Math.max(0, total - assigned);
  const progress = total > 0 ? Math.round((assigned / total) * 100) : 0;
  const disciplineCount = new Set(yearRows.map((row) => trim(row["Код_дисципліни"])).filter(Boolean)).size;
  const chartData: WorkItem[] = workKinds.map((kind) => ({ name: kind.name, hours: yearRows.reduce((sum, row) => sum + numeric(row[kind.column]), 0), color: kind.color })).filter((item) => item.hours > 0);
  const historyData = availableYears.slice().reverse().map((label) => ({ year: label, hours: workloadRows.filter((row) => trim(row["Код_кафедри"]) === departmentCode && toDbYear(row["Рік"]) === toDbYear(label)).reduce((sum, row) => sum + totalHours(row), 0) }));
  const topDisciplineTotals = (rows: Row[]): DisciplineTotal[] => {
    const grouped = new Map<string, DisciplineTotal>();
    rows.forEach((row) => {
      const code = trim(row["Код_дисципліни"]);
      if (!code) return;
      const current = grouped.get(code) ?? { code, name: disciplineNames[code] || "Без назви", hours: 0 };
      current.hours += totalHours(row);
      grouped.set(code, current);
    });
    return [...grouped.values()].sort((a, b) => b.hours - a.hours).slice(0, 8);
  };
  const topDisciplines = topDisciplineTotals(yearRows);

  const handleLogout = () => { clearCurrentUser(); void supabase.auth.signOut(); navigate("/"); };

  return <div className="dashboard-container">
    <AppSidebar active="dashboard" />
    <aside className="sidebar legacy-sidebar">
      <div className="sidebar-brand"><img src={logo} alt="НУБіП Лого" className="sidebar-logo" /><div><h3>НУБіП</h3><p>Система управління</p></div></div>
      <nav className="sidebar-menu">
        <button type="button" className="menu-item active" onClick={() => navigate("/dashboard")}><FaHome /> Головна</button>
        <span className="menu-category">Управління</span>
        <button type="button" className="menu-item" onClick={() => navigate("/workload")}><FaBriefcase /> Навантаження</button>
        <button type="button" className="menu-item" onClick={() => navigate("/teachers")}><FaGraduationCap /> Викладачі</button>
        <button type="button" className="menu-item" onClick={() => navigate("/department")}><FaHome /> Кафедра</button>
        <span className="menu-category">Додатково</span>
        <button type="button" className="menu-item" onClick={() => navigate("/help")}><FaQuestionCircle /> Довідка</button>
      </nav>
      <div className="sidebar-profile"><div className="profile-info"><div className="profile-avatar-placeholder"><FaUsers /></div><div><h4>{profileName}</h4><p>{role}</p></div><FaChevronDown className="profile-arrow" /></div><button className="btn-logout" onClick={handleLogout}><FaSignOutAlt /> Вийти</button></div>
    </aside>

    <main className="main-content">
      <header className="main-header">
        <div className="header-left"><span className="welcome-text">Вітаємо, {profileName}</span><h2>{departmentTitle}</h2><p className="role-sub">{facultyTitle}{role ? ` · ${role}` : ""}</p></div>
        <div className="header-right dashboard-year-picker">
          <label className="academic-year"><select className="date-badge" value={academicYear} onChange={(event) => setAcademicYear(event.target.value)} aria-label="Оберіть навчальний рік"><option value="">Немає навчальних років</option>{availableYears.map((item) => <option key={item} value={item}>{item} навч. рік</option>)}</select></label>
          <label className="academic-year"><select className="date-badge" value={selectedFormCode} onChange={(event) => setSelectedFormCode(event.target.value.trim())} aria-label="Оберіть форму навчання"><option value="">Усі форми</option>{studyForms.map((form) => <option key={form.code} value={form.code}>{form.name}</option>)}</select></label>
        </div>
      </header>

      <section className="metrics-grid">
        <div className="metric-card" onClick={() => navigate("/workload")} style={{ cursor: "pointer" }}><div className="metric-icon orange"><FaBriefcase /></div><div className="metric-info"><p>Загальний обсяг</p><h3>{loading ? "…" : `${numberText(total)} год`}</h3><span className="sub-stat">{academicYear || "Навчальний рік не задано"}</span></div></div>
        <div className="metric-card" onClick={() => navigate("/workload")} style={{ cursor: "pointer" }}><div className="metric-icon blue"><FaUsers /></div><div className="metric-info"><p>Розподілено</p><h3>{loading ? "…" : `${numberText(assigned)} год`}</h3><span className="sub-stat">{progress}% від загального обсягу</span></div></div>
        <div className="metric-card" onClick={() => navigate("/workload")} style={{ cursor: "pointer" }}><div className="metric-icon" style={{ background: "#ecfeff", color: "#0891b2" }}><FaClock /></div><div className="metric-info"><p>Вакантні години</p><h3>{loading ? "…" : `${numberText(vacant)} год`}</h3><span className="sub-stat">Без призначеного викладача</span></div></div>
        <div className="metric-card" onClick={() => navigate("/teachers")} style={{ cursor: "pointer" }}><div className="metric-icon purple"><FaGraduationCap /></div><div className="metric-info"><p>Штатний склад НПП</p><h3>{loading ? "…" : `${teacherRows.length} ${ukrainianCountWord(teacherRows.length, ["особа", "особи", "осіб"])}`}</h3><span className="sub-stat">{loading ? "Завантаження" : `${disciplineCount} ${ukrainianCountWord(disciplineCount, ["дисципліна", "дисципліни", "дисциплін"])} в плані`}</span></div></div>
      </section>

      {!loading && availableYears.length === 0 && <div className="dashboard-empty"><p>{loadError ? "Не вдалося отримати навчальні роки. Перевірте сповіщення або відкрийте імпорт." : "У БД ще немає даних про навчальне навантаження."}</p><button type="button" onClick={() => navigate("/workload")}>Перейти до імпорту</button></div>}

      <section className="charts-grid">
        <div className="chart-card main-chart"><div className="card-header"><h3>Навантаження за навчальними роками</h3></div><div className="chart-wrapper">{historyData.length ? <ResponsiveContainer width="100%" height="100%"><LineChart data={historyData} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E5E7EB" /><XAxis dataKey="year" stroke="#9CA3AF" tickLine={false} style={{ fontSize: 11 }} /><YAxis stroke="#9CA3AF" tickLine={false} style={{ fontSize: 11 }} /><Tooltip formatter={(value) => [`${numberText(Number(value))} год`, "Обсяг"]} /><Line type="monotone" dataKey="hours" name="Навантаження" stroke="#2563EB" strokeWidth={3} dot={{ r: 4 }} /></LineChart></ResponsiveContainer> : <p className="dashboard-chart-empty">Немає даних для графіка.</p>}</div></div>
        <div className="chart-card pie-chart-card"><div className="card-header stack-left"><h3>Структура навантаження</h3><span className="sub-stat">За видами роботи · {academicYear || "—"}</span></div><div className="pie-content">{chartData.length ? <><div className="pie-wrapper"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={chartData} innerRadius={60} outerRadius={80} paddingAngle={3} dataKey="hours" nameKey="name">{chartData.map((item) => <Cell key={item.name} fill={item.color} />)}</Pie><Tooltip formatter={(value) => [`${numberText(Number(value))} год`, ""]} /></PieChart></ResponsiveContainer><div className="pie-center-text"><h4>{numberText(total)}</h4><p>годин</p></div></div><ul className="pie-legend">{chartData.map((item) => <li key={item.name}><div className="legend-left"><span className="dot-legend" style={{ backgroundColor: item.color }} /><span className="legend-label">{item.name}</span></div><strong className="legend-value">{numberText(item.hours)} год.</strong></li>)}</ul></> : <p className="dashboard-chart-empty">Немає рядків навантаження за цей рік.</p>}</div></div>
      </section>

      <section className="details-grid">
        <div className="detail-card"><div className="card-header"><h3>Дисципліни з найбільшими годинами</h3></div>{topDisciplines.length ? <ul className="popular-list top-disciplines-list">{topDisciplines.map((item) => <li key={item.code}><div className="list-info"><span>{item.name}</span><strong>{numberText(item.hours)} год</strong></div><div className="progress-bar"><div className="progress" style={{ width: `${total > 0 ? Math.min(100, Math.round(item.hours / total * 100)) : 0}%`, backgroundColor: "#2563EB" }} /></div></li>)}</ul> : <p className="dashboard-empty-inline">{loading ? "Завантаження…" : "Немає дисциплін за обраний рік."}</p>}</div>
        <div className="detail-card"><h3>Швидкий доступ</h3><div className="quick-actions-grid">
          <button className="action-btn" onClick={() => navigate("/teachers")}><span className="action-icon"><FaGraduationCap /></span><span className="action-content"><strong>Список викладачів</strong><small>Кадровий склад, посади та ставки НПП</small></span></button>
          <button className="action-btn" onClick={() => navigate("/workload")}><span className="action-icon"><FaBriefcase /></span><span className="action-content"><strong>Навантаження кафедри</strong><small>Імпорт Excel, розподіл годин і звіти</small></span></button>
          <button className="action-btn" onClick={() => navigate("/department")}><span className="action-icon"><FaFileAlt /></span><span className="action-content"><strong>Дані кафедри</strong><small>Загальна інформація та реквізити кафедри</small></span></button>
        </div></div>
      </section>
    </main>
  </div>;
}

export default Dashboard;
