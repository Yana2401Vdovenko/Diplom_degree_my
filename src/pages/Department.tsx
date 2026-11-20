import { useEffect, useState } from "react";
import type { FormEvent } from "react";
import { FiAward, FiEdit2, FiHome, FiMail, FiMapPin, FiPhone, FiRefreshCw, FiX } from "react-icons/fi";
import AppSidebar from "../components/AppSidebar";
import { supabase } from "../utils/supabaseClient";
import { ensureCurrentUserDepartment, findDepartmentFacultyCode, setCurrentUserDepartment } from "../utils/currentDepartment";
import { notify } from "../utils/notifications";
import "../styles/DepartmentManager.css";
import "../styles/DepartmentCards.css";

type Department = { code: string; facultyCode: string; facultyName: string; name: string; phone: string; email: string; address: string; building: string; headName: string; headCredentials: string };
type DepartmentForm = Pick<Department, "code" | "facultyCode" | "name" | "phone" | "email" | "address" | "building">;
const asText = (value: unknown) => String(value ?? "").trim();
const mapDepartment = (row: Record<string, unknown>, facultyRows: Record<string, unknown>[]): Department => {
  const facultyCode = findDepartmentFacultyCode(row);
  const faculty = facultyRows.find((item) => asText(item["Код_факультету_нні"] ?? item["Код_факультету"]) === facultyCode);
  return {
    code: asText(row["Код_кафедри"]), facultyCode,
    facultyName: asText(faculty?.["Назва_факультету_нні"] ?? faculty?.["Назва_факультету"] ?? faculty?.["Назва"] ?? row["Назва_факультету_нні"]),
    name: asText(row["Назва_кафедри"]), phone: asText(row["Телефон_кафедри"]), email: asText(row["Email_кафедри"]),
    address: asText(row["Адреса_кафедри"]), building: asText(row["Корпус_кафедри"]), headName: "", headCredentials: "",
  };
};

export default function DepartmentPage() {
  const [role] = useState(() => localStorage.getItem("userRole") ?? "Лаборант");
  const canManage = ["лаборант", "laborant", "admin"].some((item) => role.toLocaleLowerCase("uk-UA").includes(item));
  const [items, setItems] = useState<Department[]>([]);
  const [activeCode, setActiveCode] = useState("");
  const [editing, setEditing] = useState<DepartmentForm | null>(null);
  const [form, setForm] = useState<DepartmentForm | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  const load = async () => {
    setLoading(true);
    const [{ data: authData }, departmentResult, teacherResult, infoResult, facultyResult, positionResult] = await Promise.all([
      supabase.auth.getUser(),
      supabase.from("Кафедра").select("*").order("Назва_кафедри"),
      supabase.from("Викладач").select("*"),
      supabase.from("Кафедральна_інформація_про_виклад").select("*"),
      supabase.from("Факультет_ННІ").select("*"),
      supabase.from("Посада").select("*"),
    ]);
    const error = departmentResult.error ?? teacherResult.error ?? infoResult.error ?? facultyResult.error ?? positionResult.error;
    if (error) {
      notify(`Не вдалося завантажити дані кафедр: ${error.message}`, "error", "Кафедра");
      setLoading(false);
      return;
    }
    const departments = (departmentResult.data ?? []) as Record<string, unknown>[];
    const teachers = (teacherResult.data ?? []) as Record<string, unknown>[];
    const assignments = (infoResult.data ?? []) as Record<string, unknown>[];
    const facultyRows = (facultyResult.data ?? []) as Record<string, unknown>[];
    const positions = (positionResult.data ?? []) as Record<string, unknown>[];
    const mapped = departments.map((row) => mapDepartment(row, facultyRows));
    const withHeads = mapped.map((department) => {
      const headAssignment = assignments.find((assignment) => {
        if (asText(assignment["Код_кафедри"]) !== department.code) return false;
        const positionCode = asText(assignment["Код_посади"]);
        const position = positions.find((item) => asText(item["Код_посади"]) === positionCode);
        return /завідувач кафедри/i.test(asText(position?.["Назва_посади"] ?? position?.["Назва"] ?? position?.["Посада"] ?? assignment["Назва_посади"]));
      });
      const headTeacher = headAssignment && teachers.find((teacher) => asText(teacher["Код_викладача"]) === asText(headAssignment["Код_викладача"]));
      return {
        ...department,
        headName: asText(headTeacher?.["ПІБ"]),
        headCredentials: [asText(headTeacher?.["Науковий_ступінь"]), asText(headTeacher?.["Вчене_звання"])].filter((value) => value && !/^без /i.test(value)).join(", "),
      };
    });
    const user = authData.user;
    const metadata = user?.user_metadata ?? {};
    const email = asText(user?.email).toLocaleLowerCase("uk-UA");
    const teacherCodes = new Set(teachers.filter((teacher) => asText(teacher["Email_викладача"]).toLocaleLowerCase("uk-UA") === email).map((teacher) => asText(teacher["Код_викладача"])));
    const linkedDepartmentCodes = assignments.filter((assignment) => teacherCodes.has(asText(assignment["Код_викладача"]))).map((assignment) => asText(assignment["Код_кафедри"]));
    const metadataDepartments = [metadata["Коди_кафедр"], metadata["department_codes"]].flatMap((value) => Array.isArray(value) ? value.map(asText) : typeof value === "string" ? value.split(",").map(asText) : []);
    let preferredCode = asText(metadata["Код_кафедри"] ?? metadata["department_code"]);
    let allowedCodes = [...new Set([...metadataDepartments, ...linkedDepartmentCodes, preferredCode])].filter((code) => mapped.some((item) => item.code === code));
    if (!allowedCodes.length) {
      try {
        const binding = await ensureCurrentUserDepartment(user);
        preferredCode = binding.departmentCode;
        if (binding.departmentCode) allowedCodes = [binding.departmentCode];
      } catch (bindingError) {
        notify(`Не вдалося визначити кафедру користувача: ${bindingError instanceof Error ? bindingError.message : "невідома помилка"}`, "error", "Кафедра");
      }
    }
    const visible = allowedCodes.length ? withHeads.filter((item) => allowedCodes.includes(item.code)) : [];
    const nextActive = visible.some((item) => item.code === preferredCode) ? preferredCode : visible[0]?.code ?? "";
    setItems(visible);
    setActiveCode(nextActive);
    setLoading(false);
  };

  useEffect(() => { void load(); }, []);
  const setField = (key: keyof DepartmentForm, value: string) => setForm((current) => current ? { ...current, [key]: value } : current);
  const beginEdit = (item: Department) => {
    setEditing({ code: item.code, facultyCode: item.facultyCode, name: item.name, phone: item.phone, email: item.email, address: item.address, building: item.building });
    setForm({ code: item.code, facultyCode: item.facultyCode, name: item.name, phone: item.phone, email: item.email, address: item.address, building: item.building });
  };
  const closeEdit = () => { setEditing(null); setForm(null); };
  const save = async (event: FormEvent) => {
    event.preventDefault();
    if (!editing || !form || !form.name.trim()) { notify("Заповніть назву кафедри.", "error", "Кафедра"); return; }
    setSaving(true);
    const payload = {
      "Назва_кафедри": form.name.trim(),
      "Телефон_кафедри": form.phone.trim() || null,
      "Email_кафедри": form.email.trim() || null,
      "Адреса_кафедри": form.address.trim() || null,
      "Корпус_кафедри": form.building.trim() || null,
    };
    const { error } = await supabase.from("Кафедра").update(payload).eq("Код_кафедри", editing.code);
    setSaving(false);
    if (error) { notify(`Не вдалося зберегти реквізити: ${error.message}`, "error", "Кафедра"); return; }
    closeEdit();
    notify("Дані кафедри оновлено.", "success", "Кафедра");
    await load();
  };
  const activate = async (item: Department) => {
    if (item.code === activeCode) return;
    const { data } = await supabase.auth.getUser();
    try {
      await setCurrentUserDepartment(data.user, item.code, item.facultyCode);
      setActiveCode(item.code);
      notify(`Активну кафедру змінено: ${item.name}`, "success", "Кафедра");
    } catch (error) {
      notify(`Не вдалося обрати кафедру: ${error instanceof Error ? error.message : "невідома помилка"}`, "error", "Кафедра");
    }
  };

  return <div className="app-shell"><AppSidebar active="department" /><main className="department-manager app-shell-main">
    <header className="department-header department-page-heading">
      <div><p>Налаштування</p><h1>{items.length > 1 ? "Мої кафедри" : "Кафедра"}</h1></div>
      <div className="department-page-actions"><span className="department-role-badge">Роль: {role}</span><button type="button" className="department-secondary" onClick={() => void load()} disabled={loading}><FiRefreshCw className={loading ? "department-refreshing" : ""} />Оновити</button></div>
    </header>
    {loading ? <div className="department-card-grid"><div className="department-info-card department-card-loading">Завантаження кафедр…</div></div>
      : items.length === 0 ? <div className="department-card-grid"><div className="department-info-card department-card-loading">Для облікового запису не знайдено прив’язаних кафедр.</div></div>
        : <section className="department-card-grid" aria-label="Прив’язані кафедри">{items.map((item) => {
          const isActive = item.code === activeCode;
          return <article className="department-info-card" key={item.code}>
            <div className="department-info-card-top"><span className="department-code-badge">Код {item.code}</span>
              {isActive ? <span className="department-active-badge">Активна кафедра</span> : <button type="button" className="department-select-button" onClick={() => void activate(item)}>Обрати для роботи</button>}
              {canManage && <button type="button" className="department-card-edit" onClick={() => beginEdit(item)}><FiEdit2 />Редагувати</button>}
            </div>
            <h2>{item.name}</h2><p className="department-faculty-name">{item.facultyName || "Факультет / ННІ не вказано"}</p>
            <div className="department-contact-grid">
              <div><FiPhone /><span><small>Телефон</small><b>{item.phone || "—"}</b></span></div>
              <div><FiMail /><span><small>Email</small><b>{item.email || "—"}</b></span></div>
              <div><FiHome /><span><small>Корпус</small><b>{item.building ? (/корпус/i.test(item.building) ? item.building : `Корпус №${item.building}`) : "—"}</b></span></div>
              <div><FiMapPin /><span><small>Адреса</small><b>{item.address || "—"}</b></span></div>
            </div>
            <div className="department-head-contact"><FiAward /><span><small>Завідувач кафедри</small><b>{item.headName || "Не призначено (додайте викладача з посадою „Завідувач кафедри“)"}</b>{item.headCredentials && <em>{item.headCredentials}</em>}</span></div>
          </article>;
        })}</section>}
    {editing && form && <div className="department-modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget && !saving) closeEdit(); }}>
      <section className="department-edit-modal" role="dialog" aria-modal="true" aria-labelledby="department-edit-title">
        <header><div><h2 id="department-edit-title">Редагувати реквізити кафедри</h2><p>{editing.name}</p></div><button type="button" aria-label="Закрити" onClick={closeEdit} disabled={saving}><FiX /></button></header>
        <form onSubmit={(event) => void save(event)}>
          <label>Код кафедри<input value={form.code} readOnly disabled /></label><small className="department-code-hint">Зміна коду кафедри та факультету доступна лише адміністратору</small>
          <label>Назва кафедри *<input required value={form.name} onChange={(event) => setField("name", event.target.value)} /></label>
          <label>Код факультету / ННІ<input value={form.facultyCode} readOnly disabled /></label><small className="department-code-hint">Зміна коду кафедри та факультету доступна лише адміністратору</small>
          <label>Телефон<input value={form.phone} onChange={(event) => setField("phone", event.target.value)} /></label>
          <label>Email<input type="email" value={form.email} onChange={(event) => setField("email", event.target.value)} /></label>
          <label>Корпус<input value={form.building} onChange={(event) => setField("building", event.target.value)} /></label>
          <label className="department-address-field">Адреса<input value={form.address} onChange={(event) => setField("address", event.target.value)} /></label>
          <footer><button type="button" className="department-secondary" onClick={closeEdit} disabled={saving}>Скасувати</button><button type="submit" className="department-primary" disabled={saving}>{saving ? "Збереження…" : "Зберегти"}</button></footer>
        </form>
      </section>
    </div>}
  </main></div>;
}
