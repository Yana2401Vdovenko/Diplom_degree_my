import { useEffect, useMemo, useRef, useState } from "react";
import { FiChevronDown, FiDownload, FiEdit2, FiPlus, FiRefreshCw } from "react-icons/fi";
import { supabase } from "../utils/supabaseClient";
import { ensureCurrentUserDepartment, findDepartmentFacultyCode } from "../utils/currentDepartment";
import { notify } from "../utils/notifications";
import type { NotificationType } from "../utils/notifications";
import AppSidebar from "../components/AppSidebar";
import SearchableSelect from "../components/SearchableSelect";
import { FilterReset, FilterSearch, FilterSelect } from "../components/TableFilters";
import { exportTeacherRosterReport, type TeacherRosterExportRow } from "../utils/reportExport";
import "../styles/TeacherWorkspace.css";
import "../styles/TeacherOverrides.css";
import "../styles/UnifiedSidebar.css";

type View = "all" | "main" | "internal" | "external" | "parttime";
type Mode = "list" | "add" | "edit" | "restore";
type Teacher = { code: string; rawCode?: string; name: string; position: string; positionCode: string; rate: string; workType: "Основний" | "Сумісник"; workTypeCode: string; workTypeName: string; staffKind: string; staffKindName: string; normHours: number; degree: string; rank: string; phone: string; email: string; birthDate: string; archived?: boolean; departmentCode?: string; departmentName?: string; archiveSummary?: string; archivedDepartmentInfo?: Record<string, unknown>[]; hasArchivedDepartmentInfo?: boolean; archivedTeacherData?: Record<string, unknown>; archiveRowKeyField?: string; archiveRowKeyValue?: string | number; archiveRecordKey?: string };
type Reference = { code: string; name: string };
// Значення звірені з таблицею «Тип_викладача»: 1 — штатний, 2 — внутрішній сумісник.
const workTypeCodes: Record<Teacher["workType"], string> = { "Основний": "1", "Сумісник": "2" };
const emptyTeacher: Teacher = { code: "", name: "", position: "", positionCode: "", rate: "1", workType: "Основний", workTypeCode: "1", workTypeName: "Штатний", staffKind: "НПП", staffKindName: "НПП", normHours: 480, degree: "Без наук. ступеня", rank: "Без вч. звання", phone: "", email: "", birthDate: "" };
const asReference = (row: Record<string, unknown>, codeField: string, nameFields: string[]): Reference => ({ code: String(row[codeField] ?? "").trim(), name: String(nameFields.map((field) => row[field]).find(Boolean) ?? "").trim() });
const classifyStaff = (statusCode: string, rate: string, norms: Record<string, number>): Pick<Teacher, "normHours"> => {
  const baseHours = norms[statusCode] ?? (statusCode === "ПП" ? 900 : 480);
  return { normHours: baseHours * (Number(rate) || 0) };
};

export default function Teachers() {
  const [role] = useState(() => localStorage.getItem("userRole") ?? "Лаборант");
  const canManageTeachers = role.toLocaleLowerCase("uk-UA").includes("лаборант");
  const [teachers, setTeachers] = useState<Teacher[]>([]);
  const [archivedTeachers, setArchivedTeachers] = useState<Teacher[]>([]);
  const [positions, setPositions] = useState<Reference[]>([]);
  const [teacherTypes, setTeacherTypes] = useState<Reference[]>([]);
  const [statuses, setStatuses] = useState<Reference[]>([]);
  const [degreeOptions, setDegreeOptions] = useState<string[]>(["Без наук. ступеня", "Канд. наук", "Д-р наук", "Доктор філософії (PhD)"]);
  const [rankOptions, setRankOptions] = useState<string[]>(["Без вч. звання", "Доцент", "Професор", "Старший дослідник"]);
  const [departmentCode, setDepartmentCode] = useState("");
  const [departmentName, setDepartmentName] = useState("");
  const [facultyName, setFacultyName] = useState("");
  const [departmentHeadName, setDepartmentHeadName] = useState("");
  const [mode, setMode] = useState<Mode>("list");
  const [view, setView] = useState<View>("all");
  const [form, setForm] = useState<Teacher>(emptyTeacher);
  const [selectedCode, setSelectedCode] = useState("");
  const [saving, setSaving] = useState(false);
  const setNotice = (message: string | null) => {
    if (!message?.trim()) return;
    const type: NotificationType = /не вдалося|помил|не підтвердив|конфлікт/i.test(message) ? "error"
      : /але не вдалося|ще не|залишився/i.test(message) ? "warning"
        : /збережено|оновлено|переміщено|поновлено|додано/i.test(message) ? "success" : "info";
    notify(message, type, "Викладачі");
  };

  const loadTeachers = async () => {
    const [authResult, teacherResult, archiveResult, infoResult, departmentResult, facultyResult, positionResult, typeResult, statusResult, normResult] = await Promise.all([
      supabase.auth.getUser(),
      supabase.from("Викладач").select("*").order("ПІБ"),
      supabase.from("Архів").select("*").eq("Таблиця", "Викладач").order("Дата_архівації", { ascending: false }),
      supabase.from("Кафедральна_інформація_про_виклад").select("*"),
      supabase.from("Кафедра").select("*").order("Код_кафедри"),
      supabase.from("Факультет_ННІ").select("*"),
      supabase.from("Посада").select("*"),
      supabase.from("Тип_викладача").select("*"),
      supabase.from("Статус_викладача").select("*"),
      supabase.from("Навантаження").select("*"),
    ]);
    if (teacherResult.error) { setNotice(`Не вдалося завантажити викладачів: ${teacherResult.error.message}`); return; }
    if (archiveResult.error) { setNotice(`Не вдалося завантажити архів: ${archiveResult.error.message}`); return; }
    const lookupError = [positionResult, typeResult, statusResult, normResult].find((result) => result.error)?.error;
    if (lookupError) setNotice(`Не вдалося отримати довідники з Supabase: ${lookupError.message}. Виконайте файл supabase/rls_roles.sql у SQL Editor і увійдіть у систему повторно.`);
    const infos = (infoResult.data ?? []) as Record<string, unknown>[];
    const departments = (departmentResult.data ?? []) as Record<string, unknown>[];
    let currentDepartment = "";
    try {
      currentDepartment = (await ensureCurrentUserDepartment(authResult.data.user)).departmentCode;
    } catch (error) {
      setNotice(`Не вдалося визначити кафедру користувача: ${error instanceof Error ? error.message : "невідома помилка"}`);
      return;
    }
    if (!currentDepartment) { setNotice("Для цього облікового запису не визначена кафедра."); return; }
    const currentDepartmentRow = departments.find((item) => String(item["Код_кафедри"] ?? "").trim() === currentDepartment.trim());
    const currentDepartmentName = String(currentDepartmentRow?.["Назва_кафедри"] ?? "").trim();
    const currentFacultyCode = findDepartmentFacultyCode(currentDepartmentRow);
    const facultyRows = (facultyResult.data ?? []) as Record<string, unknown>[];
    const currentFacultyRow = facultyRows.find((item) => String(item["Код_факультету_нні"] ?? item["Код_факультету"] ?? "").trim() === currentFacultyCode.trim());
    const currentFacultyName = String(currentFacultyRow?.["Назва_факультету_нні"] ?? currentFacultyRow?.["Назва_факультету"] ?? currentFacultyRow?.["Назва"] ?? currentDepartmentRow?.["Назва_факультету"] ?? currentDepartmentRow?.["Факультет"] ?? "").trim();
    setDepartmentName(currentDepartmentName);
    setFacultyName(currentFacultyName);
    const positionList = (positionResult.data ?? []).map((row) => asReference(row as Record<string, unknown>, "Код_посади", ["Назва"])).filter((row) => row.code);
    const positionRows = (positionResult.data ?? []) as Record<string, unknown>[];
    const headAssignment = infos.find((info) => {
      if (String(info["Код_кафедри"] ?? "").trim() !== currentDepartment.trim()) return false;
      const positionCode = String(info["Код_посади"] ?? "").trim();
      const position = positionRows.find((item) => String(item["Код_посади"] ?? "").trim() === positionCode);
      return /завідувач кафедри/i.test(String(position?.["Назва_посади"] ?? position?.["Назва"] ?? info["Назва_посади"] ?? info["Посада"] ?? ""));
    });
    const departmentHead = headAssignment && ((teacherResult.data ?? []) as Record<string, unknown>[]).find((teacher) => String(teacher["Код_викладача"] ?? "").trim() === String(headAssignment["Код_викладача"] ?? "").trim());
    setDepartmentHeadName(String(departmentHead?.["ПІБ"] ?? "").trim());
    const typeList = ((typeResult.data ?? []) as Record<string, unknown>[]).map((row) => ({
      code: String(row["Код_типу"] ?? "").trim(),
      name: String(row["Назва_типу"] ?? "").trim(),
    })).filter((row) => row.code);
    const statusList = ((statusResult.data ?? []) as Record<string, unknown>[]).map((row) => ({
      code: String(row["Код_статусу_викладача"] ?? "").trim(),
      name: String(row["Назва"] ?? "").trim(),
    })).filter((row) => row.code);
    const norms = Object.fromEntries(((normResult.data ?? []) as unknown as Record<string, unknown>[]).map((item) => [String(item["Код_статусу_викладача"] ?? "").trim(), Number(item["Години"] ?? 0)]));
    setPositions(positionList);
    setTeacherTypes(typeList);
    setStatuses(statusList);
    const uniqueTextValues = (base: string[], column: string) => [...new Set([...base, ...((teacherResult.data ?? []) as Record<string, unknown>[]).map((item) => String(item[column] ?? "").trim()).filter(Boolean)])];
    setRankOptions(uniqueTextValues(["Без вч. звання", "Доцент", "Професор", "Старший дослідник"], "Вчене_звання"));
    setDegreeOptions(uniqueTextValues(["Без наук. ступеня", "Канд. наук", "Д-р наук", "Доктор філософії (PhD)"], "Науковий_ступінь"));
    // Поки кафедральна інформація порожня, беремо код із таблиці «Кафедра».
    // Це дає змогу створити перший зв'язувальний запис для викладача.
    setDepartmentCode(currentDepartment);
    const mapTeacher = (row: Record<string, unknown>, infoOverride?: Record<string, unknown>): Teacher => {
      const rawCode = String(row["Код_викладача"] ?? "");
      const code = rawCode.trim();
      const info = infoOverride ?? infos.find((item) => String(item["Код_викладача"] ?? "").trim() === code.trim()) ?? {};
      // Нові викладачі можуть ще не мати кафедрального запису. У такому разі
      // форма має отримати справжнє значення select, а не службове «—».
      const positionCode = String(info["Код_посади"] ?? "").trim();
      const position = positionList.find((item) => item.code.trim() === positionCode)?.name ?? String(info["Посада"] ?? "").trim();
      const typeCode = String(info["Код_типу"] ?? info["Код_типу_викладача"] ?? "").trim();
      const typeName = typeList.find((item) => item.code.trim() === typeCode)?.name
        ?? String(info["Тип_викладача"] ?? (typeCode === "2" ? "Внутрішній сумісник" : typeCode === "3" ? "Зовнішній сумісник" : typeCode === "1" ? "Штатний" : "Не вказано")).trim();
      const statusCode = String(info["Код_статусу_викладача"] ?? info["Категорія"] ?? "").trim();
      const statusName = statusList.find((item) => item.code.trim() === statusCode)?.name || statusCode || "—";
      const rate = String(info["Частка_ставки"] ?? info["Ставка"] ?? "1");
      const teacherDepartmentCode = String(info["Код_кафедри"] ?? "").trim();
      const teacherDepartmentName = String(departments.find((item) => String(item["Код_кафедри"] ?? "").trim() === teacherDepartmentCode)?.["Назва_кафедри"] ?? "").trim();
      return { code, rawCode, name: String(row["ПІБ"] ?? ""), position, positionCode, rate, workType: /суміс/i.test(typeName) ? "Сумісник" : "Основний", workTypeCode: typeCode, workTypeName: typeName || "Не вказано", staffKind: statusCode, staffKindName: statusName, ...classifyStaff(statusCode, rate, norms), degree: String(row["Науковий_ступінь"] ?? "Без наук. ступеня"), rank: String(row["Вчене_звання"] ?? "Без вч. звання"), phone: String(row["Телефон_викладача"] ?? ""), email: String(row["Email_викладача"] ?? ""), birthDate: String(row["Дата_народження"] ?? ""), departmentCode: teacherDepartmentCode, departmentName: teacherDepartmentName };
    };
    setTeachers((teacherResult.data ?? []).map((row) => mapTeacher(row as Record<string, unknown>)).filter((teacher) => teacher.departmentCode === currentDepartment));
    setArchivedTeachers((archiveResult.data ?? []).map((archiveRowValue) => {
      const archiveRow = archiveRowValue as Record<string, unknown>;
      const data = (archiveRow["Дані"] ?? {}) as Record<string, unknown>;
      const rawStoredInfo = data["Кафедральна_інформація_про_виклад"];
      const hasStoredInfo = Array.isArray(rawStoredInfo) || Boolean(rawStoredInfo && typeof rawStoredInfo === "object");
      const storedInfo = Array.isArray(rawStoredInfo)
        ? rawStoredInfo as Record<string, unknown>[]
        : rawStoredInfo && typeof rawStoredInfo === "object" ? [rawStoredInfo as Record<string, unknown>] : [];
      const displayInfo = hasStoredInfo ? storedInfo : (data["Код_кафедри"] || data["Посада"] || data["Ставка"] ? [data] : []);
      const teacher = mapTeacher(data, displayInfo[0] ?? {});
      const summary = displayInfo.map((info) => {
        const infoTeacher = mapTeacher(data, info);
        return [infoTeacher.position, `${infoTeacher.rate} ставки`, infoTeacher.departmentName || infoTeacher.departmentCode].filter(Boolean).join(" · ");
      }).filter(Boolean).join("; ");
      const archiveRowKeyField = ["id", "Код_архіву", "ID", "Ідентифікатор_архіву"].find((field) => archiveRow[field] != null);
      return { ...teacher, archived: true, archivedDepartmentInfo: storedInfo, hasArchivedDepartmentInfo: hasStoredInfo, archivedTeacherData: data, archiveRowKeyField, archiveRowKeyValue: archiveRowKeyField ? archiveRow[archiveRowKeyField] as string | number : undefined, archiveRecordKey: String(archiveRow["Ключ_запису"] ?? teacher.code), archiveSummary: summary || [teacher.position, `${teacher.rate} ставки`, teacher.departmentName || teacher.departmentCode].filter(Boolean).join(" · ") };
    }).filter((teacher) => teacher.code && teacher.departmentCode === currentDepartment));
  };
  // The request resolves asynchronously; the rule does not recognize the state update inside loadTeachers.
  // eslint-disable-next-line react-hooks/set-state-in-effect
  useEffect(() => { void loadTeachers(); }, []);
  const updateForm = (field: keyof Teacher, value: string) => setForm((current) => {
    if (field === "workTypeCode") {
      const type = teacherTypes.find((item) => item.code.trim() === value.trim());
      const name = type?.name || "Не вказано";
      return { ...current, workTypeCode: value.trim(), workTypeName: name, workType: /суміс/i.test(name) ? "Сумісник" : "Основний" };
    }
    if (field === "positionCode") {
      const position = positions.find((item) => item.code.trim() === value.trim());
      return { ...current, positionCode: value.trim(), position: position?.name ?? current.position };
    }
    if (field === "staffKind") {
      const status = statuses.find((item) => item.code.trim() === value.trim());
      return { ...current, staffKind: value.trim(), staffKindName: status?.name || value.trim() || "—" };
    }
    return { ...current, [field]: value };
  });
  const beginEdit = (teacher: Teacher) => {
    const defaultType = teacherTypes.find((item) => item.code.trim() === "1") ?? teacherTypes[0];
    const defaultStatus = statuses.find((item) => item.code.trim() === "НПП") ?? statuses[0];
    setSelectedCode(teacher.code);
    setForm({ ...teacher, positionCode: teacher.positionCode || positions[0]?.code || "", staffKind: teacher.staffKind || defaultStatus?.code || "", staffKindName: teacher.staffKindName || defaultStatus?.name || "НПП", workTypeCode: teacher.workTypeCode || defaultType?.code || "", workTypeName: teacher.workTypeCode ? teacher.workTypeName : defaultType?.name || "Штатний" });
    setMode("edit");
  };
  const selectTeacherForEdit = (code: string) => {
    const teacher = teachers.find((item) => item.code === code.trim());
    if (teacher) beginEdit(teacher);
  };
  const changeFilter = (value: string) => {
    setView(value as View);
    setMode("list");
  };
  const saveTeacher = async () => {
    if (!canManageTeachers) { setNotice("Лише лаборант може додавати або змінювати викладачів."); return; }
    if (!form.name.trim()) { setNotice("Заповніть ПІБ викладача."); return; }
    if (!statuses.some((item) => item.code.trim() === form.staffKind.trim())) { setNotice("Оберіть категорію НПП / ПП із довідника БД."); return; }
    if (!teacherTypes.some((item) => item.code.trim() === form.workTypeCode.trim())) { setNotice("Оберіть тип викладача із довідника БД."); return; }
    if (!Number.isFinite(Number(form.rate)) || Number(form.rate) <= 0) { setNotice("Вкажіть коректну частку ставки, більшу за 0."); return; }
    setSaving(true); setNotice(null);
    const selectedTeacher = mode === "edit" ? teachers.find((teacher) => teacher.code.trim() === selectedCode.trim()) : undefined;
    let code = selectedTeacher?.code.trim() ?? form.code.trim();
    if (mode === "add" || !code) {
      const { data: existingCodes, error: codesError } = await supabase.from("Викладач").select("Код_викладача");
      if (codesError) { setSaving(false); setNotice(`Не вдалося визначити наступний код викладача: ${codesError.message}`); return; }
      const maxId = Math.max(0, ...((existingCodes ?? []) as unknown as Record<string, unknown>[]).map((item) => parseInt(String(item["Код_викладача"] ?? "").trim(), 10) || 0));
      code = String(maxId + 1);
    }
    const rawTeacherId = selectedTeacher?.rawCode ?? code;
    const existingResult = await supabase.from("Кафедральна_інформація_про_виклад").select("*").eq("Код_викладача", rawTeacherId);
    if (existingResult.error) {
      setSaving(false);
      setNotice(`Не вдалося перевірити кафедральну інформацію: ${existingResult.error.message}`);
      return;
    }
    const existingRows = (existingResult.data ?? []) as unknown as Record<string, unknown>[];
    const existing = existingRows.find((row) => String(row["Код_кафедри"] ?? "").trim() === departmentCode.trim()) ?? existingRows[0] ?? null;
    const targetDepartmentCode = String(existing?.["Код_кафедри"] ?? departmentCode).trim();
    if (!targetDepartmentCode) {
      setSaving(false);
      setNotice("Не знайдено код кафедри. Для збереження ставки, НПП/ПП і посади потрібен запис у таблиці «Кафедральна_інформація_про_виклад» із кодом кафедри.");
      return;
    }
    const positionCode = form.positionCode.trim();
    if (!positionCode || !positions.some((item) => item.code.trim() === positionCode)) { setSaving(false); setNotice(`Оберіть посаду з довідника БД. Не знайдено її код для «${form.position || "посади"}».`); return; }
    const teacherPayload = { "ПІБ": form.name.trim(), "Вчене_звання": form.rank, "Науковий_ступінь": form.degree, "Телефон_викладача": form.phone, "Email_викладача": form.email, "Дата_народження": form.birthDate || null };
    const teacherWrite = mode === "edit"
      ? await supabase.from("Викладач").update(teacherPayload).eq("Код_викладача", rawTeacherId).select("Код_викладача")
      : await supabase.from("Викладач").upsert({ "Код_викладача": code, ...teacherPayload }, { onConflict: "Код_викладача" }).select("Код_викладача");
    const savedTeacherRows = (teacherWrite.data ?? []) as unknown as Record<string, unknown>[];
    if (teacherWrite.error) {
      setSaving(false);
      setNotice(`Не вдалося оновити таблицю «Викладач»: ${teacherWrite.error.message}`);
      return;
    }
    if (savedTeacherRows.length === 0) { setSaving(false); setNotice('Таблиця «Викладач» повернула 0 оновлених рядків. Перевірте Код_викладача та права доступу.'); return; }
    const infoPayload: Record<string, unknown> = {
      "Код_кафедри": targetDepartmentCode.trim(),
      "Код_викладача": code.trim(),
      "Код_статусу_викладача": form.staffKind.trim(),
      "Код_посади": positionCode,
      "Частка_ставки": Number(form.rate) || 1,
      "Код_типу": form.workTypeCode.trim() || null,
    };
    const infoWrite = await supabase.from("Кафедральна_інформація_про_виклад")
      .upsert(infoPayload, { onConflict: "Код_викладача,Код_кафедри" }).select("Код_викладача");
    const savedDepartmentRows = (infoWrite.data ?? []) as unknown as Record<string, unknown>[];
    setSaving(false);
    if (infoWrite.error) { setNotice(`Не вдалося оновити таблицю «Кафедральна_інформація_про_виклад»: ${infoWrite.error.message}`); return; }
    if (savedDepartmentRows.length === 0) {
      setNotice('Таблиця «Кафедральна_інформація_про_виклад» повернула 0 оновлених рядків. Перевірте Код_викладача, Код_кафедри та права доступу.');
      return;
    }
    setNotice("Дані викладача збережено."); setMode("list"); setForm(emptyTeacher); await loadTeachers();
  };
  const archiveTeacher = async () => {
    if (!canManageTeachers) { setNotice("Лише лаборант може переміщувати викладачів до архіву."); return; }
    if (!form.code || !window.confirm(`Звільнити викладача «${form.name}» та перемістити його до архіву?`)) return;
    setSaving(true); setNotice(null);
    const code = form.code.trim();
    const [{ data: teacherData, error: teacherReadError }, { data: departmentInfoData, error: infoReadError }] = await Promise.all([
      supabase.from("Викладач").select("*").eq("Код_викладача", code).maybeSingle(),
      supabase.from("Кафедральна_інформація_про_виклад").select("*").eq("Код_викладача", code),
    ]);
    if (teacherReadError || !teacherData) { setSaving(false); setNotice(`Не вдалося прочитати дані викладача перед архівацією: ${teacherReadError?.message ?? "запис не знайдено"}`); return; }
    if (infoReadError) { setSaving(false); setNotice(`Не вдалося прочитати кафедральну інформацію: ${infoReadError.message}`); return; }
    const departmentInfo = (departmentInfoData ?? []) as unknown as Record<string, unknown>[];
    const archivedData = { ...(teacherData as Record<string, unknown>), "Кафедральна_інформація_про_виклад": departmentInfo };
    const archivePayload = { "Таблиця": "Викладач", "Ключ_запису": code, "Дані": archivedData, "Назва": String((teacherData as Record<string, unknown>)["ПІБ"] ?? form.name) };
    let { data: savedArchive, error: archiveError } = await supabase.from("Архів")
      .upsert(archivePayload, { onConflict: "Таблиця,Ключ_запису" }).select("*").maybeSingle();
    if (archiveError || !savedArchive) {
      const { data: existingArchive, error: archiveReadError } = await supabase.from("Архів").select("*")
        .eq("Таблиця", "Викладач").eq("Ключ_запису", code).limit(1).maybeSingle();
      if (archiveReadError) { setSaving(false); setNotice(`Не вдалося зберегти дані в архіві: ${archiveError?.message ?? archiveReadError.message}`); return; }
      if (existingArchive) {
        const { error: updateError } = await supabase.from("Архів").update(archivePayload).eq("Таблиця", "Викладач").eq("Ключ_запису", code);
        if (updateError) archiveError = updateError;
        else {
          const { data, error } = await supabase.from("Архів").select("*").eq("Таблиця", "Викладач").eq("Ключ_запису", code).limit(1).maybeSingle();
          savedArchive = data;
          archiveError = error;
        }
      } else {
        const { data, error } = await supabase.from("Архів").insert(archivePayload).select("*").maybeSingle();
        savedArchive = data;
        archiveError = error;
      }
    }
    if (archiveError || !savedArchive) { setSaving(false); setNotice(`Не вдалося перемістити до архіву: ${archiveError?.message ?? "Supabase не підтвердив збереження."}`); return; }
    const { error: workloadUpdateError } = await supabase.from("Завантаження_навантаження").update({ "Код_викладача": null }).eq("Код_викладача", code);
    if (workloadUpdateError) { setSaving(false); setNotice(`Дані збережено в архіві, але не вдалося зняти викладача з навантаження: ${workloadUpdateError.message}`); return; }
    if (departmentInfo.length) {
      const { error: infoDeleteError } = await supabase.from("Кафедральна_інформація_про_виклад").delete().eq("Код_викладача", code);
      if (infoDeleteError) { setSaving(false); setNotice(`Дані збережено в архіві та навантаження відв’язано, але не вдалося видалити кафедральну інформацію: ${infoDeleteError.message}`); return; }
    }
    const { error: deleteError } = await supabase.from("Викладач").delete().eq("Код_викладача", code);
    setSaving(false);
    if (deleteError) { setNotice(`Архів створено, але не вдалося видалити активний запис: ${deleteError.message}`); return; }
    setNotice("Викладача переміщено до архіву."); setMode("list"); setForm(emptyTeacher); await loadTeachers();
  };
  const restoreTeacher = async (teacher: Teacher) => {
    setSaving(true); setNotice(null);
    const code = teacher.code.trim();
    const fail = (message: string) => { setSaving(false); setNotice(message); };
    const archivedData = teacher.archivedTeacherData ?? {};
    const teacherPayload = {
      "Код_викладача": code,
      "ПІБ": String(archivedData["ПІБ"] ?? teacher.name),
      "Вчене_звання": String(archivedData["Вчене_звання"] ?? teacher.rank),
      "Науковий_ступінь": String(archivedData["Науковий_ступінь"] ?? teacher.degree),
      "Телефон_викладача": archivedData["Телефон_викладача"] == null ? null : String(archivedData["Телефон_викладача"]),
      "Email_викладача": archivedData["Email_викладача"] == null ? null : String(archivedData["Email_викладача"]),
      "Дата_народження": archivedData["Дата_народження"] || teacher.birthDate || null,
    };
    const { data: restoredTeacher, error: restoreError } = await supabase.from("Викладач")
      .upsert(teacherPayload, { onConflict: "Код_викладача" }).select("*").maybeSingle();
    if (restoreError) { fail(`Не вдалося поновити основний запис викладача: ${restoreError.message}`); return; }
    if (!restoredTeacher || String((restoredTeacher as Record<string, unknown>)["Код_викладача"] ?? "").trim() !== code) {
      fail("Supabase не підтвердив відновлення основного запису викладача."); return;
    }

    const archivedInfoRows = teacher.hasArchivedDepartmentInfo
      ? teacher.archivedDepartmentInfo ?? []
      : teacher.archivedTeacherData?.["Код_кафедри"] || teacher.archivedTeacherData?.["Посада"] || teacher.archivedTeacherData?.["Ставка"]
        ? [teacher.archivedTeacherData]
        : [{ "Код_кафедри": teacher.departmentCode || departmentCode, "Посада": teacher.position, "Частка_ставки": teacher.rate, "Код_статусу_викладача": teacher.staffKind, "Код_типу": teacher.workTypeCode || workTypeCodes[teacher.workType] }];
    for (const archivedInfo of archivedInfoRows) {
      const archivedPosition = String(archivedInfo["Код_посади"] ?? "").trim();
      const archivedPositionName = String(archivedInfo["Посада"] ?? teacher.position).trim();
      const archivedTypeName = String(archivedInfo["Тип_викладача"] ?? "").trim();
      const department = String(archivedInfo["Код_кафедри"] ?? teacher.departmentCode ?? departmentCode).trim();
      const infoPayload: Record<string, unknown> = {
        "Код_викладача": code,
        "Код_кафедри": department,
        "Код_посади": archivedPosition || positions.find((item) => item.name.trim() === archivedPositionName)?.code || teacher.positionCode,
        "Частка_ставки": archivedInfo["Частка_ставки"] ?? archivedInfo["Ставка"] ?? teacher.rate,
        "Код_статусу_викладача": String(archivedInfo["Код_статусу_викладача"] ?? archivedInfo["Категорія"] ?? teacher.staffKind).trim(),
        "Код_типу": String(archivedInfo["Код_типу"] ?? archivedInfo["Код_типу_викладача"] ?? teacherTypes.find((item) => item.name === archivedTypeName)?.code ?? (archivedTypeName ? workTypeCodes[/суміс/i.test(archivedTypeName) ? "Сумісник" : "Основний"] : teacher.workTypeCode || workTypeCodes[teacher.workType])).trim(),
      };
      if (!department || !String(infoPayload["Код_посади"] ?? "").trim()) {
        fail("Основний запис відновлено, але архів не містить достатніх даних кафедри або посади для відновлення прив’язки."); return;
      }

      const infoTable = supabase.from("Кафедральна_інформація_про_виклад");
      let { data: savedInfo, error } = await infoTable.upsert(infoPayload, { onConflict: "Код_викладача,Код_кафедри" }).select("*").maybeSingle();
      if (error || !savedInfo) {
        const { data: existingRows, error: lookupError } = await infoTable.select("*").eq("Код_викладача", code);
        if (lookupError) { fail(`Викладача поновлено, але не вдалося перевірити кафедральні дані: ${lookupError.message}`); return; }
        const existingInfo = ((existingRows ?? []) as unknown as Record<string, unknown>[]).find((row) => String(row["Код_кафедри"] ?? "").trim() === department);
        if (existingInfo) {
          const result = await infoTable.update(infoPayload).eq("Код_викладача", code).eq("Код_кафедри", department).select("*").limit(1).maybeSingle();
          savedInfo = result.data;
          error = result.error;
        } else {
          const result = await infoTable.insert(infoPayload).select("*").maybeSingle();
          savedInfo = result.data;
          error = result.error;
        }
      }
      if (error) { fail(`Викладача поновлено, але кафедральну інформацію не вдалося відновити: ${error.message}`); return; }
      if (!savedInfo || String((savedInfo as Record<string, unknown>)["Код_викладача"] ?? "").trim() !== code || String((savedInfo as Record<string, unknown>)["Код_кафедри"] ?? "").trim() !== department) {
        fail("Supabase не підтвердив відновлення кафедральної інформації. Запис в архіві збережено для повторної спроби."); return;
      }
    }

    let archiveDeleteQuery = supabase.from("Архів").delete();
    if (teacher.archiveRowKeyField && teacher.archiveRowKeyValue != null) {
      archiveDeleteQuery = archiveDeleteQuery.eq(teacher.archiveRowKeyField, teacher.archiveRowKeyValue);
    } else {
      archiveDeleteQuery = archiveDeleteQuery.eq("Таблиця", "Викладач").eq("Ключ_запису", teacher.archiveRecordKey ?? code);
    }
    const { data: deletedArchiveRows, error: archiveDeleteError } = await archiveDeleteQuery.select("*");
    if (archiveDeleteError) { fail(`Дані викладача поновлено, але не вдалося видалити запис архіву: ${archiveDeleteError.message}`); return; }
    if (!deletedArchiveRows?.length) {
      let verifyQuery = supabase.from("Архів").select("Таблиця,Ключ_запису");
      if (teacher.archiveRowKeyField && teacher.archiveRowKeyValue != null) verifyQuery = verifyQuery.eq(teacher.archiveRowKeyField, teacher.archiveRowKeyValue);
      else verifyQuery = verifyQuery.eq("Таблиця", "Викладач").eq("Ключ_запису", teacher.archiveRecordKey ?? code);
      const { data: remainingArchiveRows, error: verifyError } = await verifyQuery;
      if (verifyError) { fail(`Не вдалося перевірити видалення архіву: ${verifyError.message}`); return; }
      if (remainingArchiveRows?.length) { fail("Дані викладача поновлено, але запис архіву залишився. Спробуйте поновлення ще раз."); return; }
    }
    setArchivedTeachers((current) => current.filter((item) => teacher.archiveRowKeyField && teacher.archiveRowKeyValue != null
      ? !(item.archiveRowKeyField === teacher.archiveRowKeyField && String(item.archiveRowKeyValue) === String(teacher.archiveRowKeyValue))
      : !(item.code.trim() === code && item.archiveRecordKey === teacher.archiveRecordKey)));
    setSaving(false);
    setNotice(null);
    setNotice(`Викладача «${teacher.name}» поновлено разом із кафедральною інформацією.`);
    setMode("list");
    await loadTeachers();
  };
  const changeMode = (nextMode: Mode) => {
    if ((mode === "add" || mode === "edit") && form.name && !window.confirm("Ви переходите до іншого вікна. Незбережені зміни буде втрачено. Продовжити?")) return;
    setMode(nextMode);
    if (nextMode === "add") {
      const defaultPosition = positions.find((item) => /асистент/i.test(item.name)) ?? positions[0];
      const defaultType = teacherTypes.find((item) => item.code.trim() === "1") ?? teacherTypes[0];
      const defaultStatus = statuses.find((item) => item.code.trim() === "НПП") ?? statuses[0];
      setForm({ ...emptyTeacher, position: defaultPosition?.name ?? "", positionCode: defaultPosition?.code ?? "", workTypeCode: defaultType?.code ?? "", workTypeName: defaultType?.name ?? "", staffKind: defaultStatus?.code ?? "", staffKindName: defaultStatus?.name ?? "" });
    } else setForm(emptyTeacher);
  };

  return <div className={`teacher-workspace ${canManageTeachers ? "can-manage" : "read-only"}`}>
<AppSidebar active="teachers" />
    <main className="teacher-main"><header><div><p>Управління кадровим складом</p><h1>Викладачі кафедри</h1></div><div className="teacher-toolbar"><button type="button" onClick={() => changeMode("add")}><FiPlus /> Додати викладача</button><button type="button" onClick={() => { if ((mode === "add" || mode === "edit") && form.name && !window.confirm("Ви переходите до редагування. Незбережені зміни буде втрачено. Продовжити?")) return; setSelectedCode(""); setForm(emptyTeacher); setMode("edit"); }}><FiEdit2 /> Редагувати</button><button type="button" onClick={() => setMode("restore")}><FiRefreshCw /> Поновити</button></div></header>
      {mode === "list" && <TeacherTable teachers={teachers} view={view} filterValue={view} onFilterChange={changeFilter} departmentName={departmentName} facultyName={facultyName} departmentHeadName={departmentHeadName} />}
      {(mode === "add" || mode === "edit") && <TeacherForm mode={mode} form={form} teachers={teachers} positions={positions} statuses={statuses} teacherTypes={teacherTypes} degreeOptions={degreeOptions} rankOptions={rankOptions} selectedCode={selectedCode} onSelect={selectTeacherForEdit} onChange={updateForm} onSave={saveTeacher} onArchive={archiveTeacher} onCancel={() => changeMode("list")} saving={saving} />}
      {mode === "restore" && <RestorePanel teachers={archivedTeachers} onRestore={restoreTeacher} onBack={() => setMode("list")} saving={saving} />}
    </main>
  </div>;
}

function ComboboxInput({ value, options, onChange, ariaLabel }: { value: string; options: string[]; onChange: (value: string) => void; ariaLabel: string }) {
  const rootRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const filteredOptions = useMemo(() => {
    const query = value.trim().toLocaleLowerCase("uk-UA");
    const unique = [...new Set(options.map((item) => item.trim()).filter(Boolean))];
    return query ? unique.filter((item) => item.toLocaleLowerCase("uk-UA").includes(query)) : unique;
  }, [options, value]);

  useEffect(() => {
    if (!open) return;
    const onPointerDown = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("pointerdown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("pointerdown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [open]);

  return <div ref={rootRef} style={{ position: "relative", width: "100%", marginTop: 5 }}>
    <input aria-label={ariaLabel} aria-expanded={open} aria-autocomplete="list" autoComplete="off" value={value} onFocus={() => setOpen(true)} onChange={(event) => { onChange(event.target.value); setOpen(true); }} style={{ boxSizing: "border-box", width: "100%", height: 40, padding: "0 38px 0 12px", border: "1px solid #cbd5e1", borderRadius: 10, background: "#fff", color: "#0f172a", fontSize: 14, outline: "none" }} />
    <button type="button" aria-label={`Показати підказки: ${ariaLabel}`} aria-expanded={open} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen((current) => !current)} style={{ position: "absolute", top: 1, right: 1, display: "flex", alignItems: "center", justifyContent: "center", width: 36, height: 38, border: 0, borderRadius: "0 9px 9px 0", background: "transparent", color: "#64748b", cursor: "pointer" }}><FiChevronDown size={16} style={{ transform: open ? "rotate(180deg)" : undefined }} /></button>
    {open && <div role="listbox" aria-label={`Підказки: ${ariaLabel}`} style={{ position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 80, maxHeight: 220, overflowY: "auto", border: "1px solid #cbd5e1", borderRadius: 10, background: "#fff", boxShadow: "0 10px 25px rgba(15, 23, 42, 0.14)" }}>
      {filteredOptions.length ? filteredOptions.map((option) => <button key={option} type="button" role="option" aria-selected={option === value} onMouseDown={(event) => event.preventDefault()} onClick={() => { onChange(option); setOpen(false); }} style={{ display: "block", width: "100%", padding: "9px 12px", border: 0, borderBottom: "1px solid #f1f5f9", background: option === value ? "#eff6ff" : "#fff", color: option === value ? "#1d4ed8" : "#1e293b", fontSize: 14, textAlign: "left", cursor: "pointer" }}>{option}</button>) : <div style={{ padding: "10px 12px", color: "#64748b", fontSize: 13 }}>Немає збігів у підказках. Власне значення можна зберегти.</div>}
    </div>}
  </div>;
}

function TeacherTable({ teachers, view, filterValue, onFilterChange, departmentName, facultyName, departmentHeadName }: { teachers: Teacher[]; view: View; filterValue: string; onFilterChange: (value: string) => void; departmentName: string; facultyName: string; departmentHeadName: string }) {
  const [search, setSearch] = useState("");
  const [reportOpen, setReportOpen] = useState(false);
  const viewTeachers = teachers.filter((teacher) => {
    const code = teacher.workTypeCode.trim();
    if (view === "main") return code === "1";
    if (view === "parttime") return code === "2" || code === "3" || teacher.workType === "Сумісник";
    return true;
  });
  const filteredTeachers = viewTeachers.filter((teacher) => !search || teacher.name.toLocaleLowerCase("uk-UA").includes(search.trim().toLocaleLowerCase("uk-UA")));
  const filterLabel = view === "main" ? "Тільки штатні" : view === "parttime" ? "Тільки сумісники" : "Усі викладачі (Штатні + Сумісники)";
  const exportRows: TeacherRosterExportRow[] = filteredTeachers.map((teacher) => ({ name: teacher.name, position: teacher.position, category: teacher.staffKindName, rate: teacher.rate, employment: teacher.workTypeCode ? teacher.workTypeName : "Не вказано", normHours: teacher.normHours, degree: teacher.degree, rank: teacher.rank, phone: teacher.phone, email: teacher.email, birthDate: teacher.birthDate }));
  const reportFile = { title: "Кадровий склад кафедри", subtitle: `${filterLabel} · Показано ${filteredTeachers.length} із ${teachers.length} осіб`, facultyName, departmentName, departmentHeadName, rows: exportRows, fileName: `Кадровий_склад_${(departmentName || "кафедра").replace(/[^\p{L}\p{N}-]+/gu, "_")}` };
  const reset = () => { setSearch(""); onFilterChange("all"); };
  const badgeStyle = (teacher: Teacher): React.CSSProperties => !teacher.workTypeCode.trim()
    ? { background: "#f1f5f9", color: "#64748b" }
    : teacher.workTypeCode.trim() === "1" ? { background: "#dcfce7", color: "#15803d" }
      : teacher.workTypeCode.trim() === "2" ? { background: "#dbeafe", color: "#1d4ed8" } : { background: "#fef3c7", color: "#b45309" };
  return <section className="teacher-card teacher-table-card">
    <div className="teacher-card-title teacher-roster-title"><div className="teacher-roster-heading-line"><h2>Кафедральна інформація про викладачів</h2><span className="teacher-roster-stat">Усі співробітники · Показано <strong>{filteredTeachers.length}</strong> із <strong>{teachers.length}</strong> осіб</span></div></div>
    <div className="teacher-roster-filters">
      <FilterSearch ariaLabel="Пошук викладача за ПІБ" placeholder="Пошук викладача за ПІБ..." value={search} onChange={setSearch} minWidth={280} />
      <FilterSelect ariaLabel="Фільтр штатні або сумісники" value={filterValue === "main" || filterValue === "parttime" ? filterValue : "all"} onChange={(value) => onFilterChange(value)} options={[{ value: "all", label: "Усі викладачі (Штатні + Сумісники)" }, { value: "main", label: "Тільки штатні" }, { value: "parttime", label: "Тільки сумісники" }]} minWidth={260} />
      {(search || filterValue !== "all") && <FilterReset onClick={reset} />}
      <button type="button" className="teacher-roster-export-button" onClick={() => setReportOpen(true)}><FiDownload size={15} aria-hidden="true" />Завантажити звіт</button>
    </div>
    <div className="teacher-table-wrap"><table><colgroup>
      <col style={{ width: 34 }} /><col style={{ width: "18%" }} /><col style={{ width: "9%" }} /><col style={{ width: "6%" }} />
      <col style={{ width: "10%" }} /><col style={{ width: "6%" }} /><col style={{ width: "10%" }} /><col style={{ width: "8%" }} />
      <col style={{ width: 92 }} /><col style={{ width: "9%" }} /><col style={{ width: "12%" }} />
    </colgroup><thead><tr><th>№</th><th>ПІБ</th><th>Посада</th><th>Категорія</th><th>Ставка / зайнятість</th><th>Норма, год</th><th>Науковий ступінь</th><th>Вчене звання</th><th>ДЕНЬ НАР.</th><th>ТЕЛЕФОН</th><th>EMAIL</th></tr></thead><tbody>{filteredTeachers.length ? filteredTeachers.map((teacher, index) => <tr key={teacher.code}><td>{index + 1}</td><td className="teacher-name">{teacher.name}</td><td>{teacher.position || "—"}</td><td>{teacher.staffKindName || "—"}</td><td><span>{teacher.rate}</span><span style={{ display: "block", width: "max-content", maxWidth: "100%", margin: "4px auto 0", padding: "2px 7px", borderRadius: 999, ...badgeStyle(teacher), fontSize: 11, fontWeight: 600 }}>{teacher.workTypeCode ? teacher.workTypeName : "Не вказано"}</span></td><td>{teacher.normHours}</td><td>{teacher.degree || "—"}</td><td>{teacher.rank || "—"}</td><td>{teacher.birthDate || "—"}</td><td className="teacher-phone">{teacher.phone || "—"}</td><td className="teacher-email">{teacher.email || "—"}</td></tr>) : <tr><td colSpan={11}>{teachers.length ? "За обраними фільтрами викладачів не знайдено." : "Викладачів ще немає. Оберіть «Додати викладача»."}</td></tr>}</tbody></table></div>
    {reportOpen && <TeacherRosterReportModal report={reportFile} onClose={() => setReportOpen(false)} />}
  </section>;
}

function TeacherRosterReportModal({ report, onClose }: { report: { title: string; subtitle: string; facultyName: string; departmentName: string; departmentHeadName: string; rows: TeacherRosterExportRow[]; fileName: string }; onClose: () => void }) {
  const [downloading, setDownloading] = useState<"docx" | "xlsx" | "pdf" | "">("");
  const [title, setTitle] = useState(report.title);
  const [subtitle, setSubtitle] = useState(report.subtitle);
  const [rows, setRows] = useState(report.rows);
  const [comment, setComment] = useState("");
  const [reportDate, setReportDate] = useState(new Date().toLocaleDateString("uk-UA"));
  const [departmentSigner, setDepartmentSigner] = useState(report.departmentHeadName);
  const fields: Array<[string, keyof TeacherRosterExportRow]> = [
    ["ПІБ викладача", "name"], ["Посада", "position"], ["Категорія", "category"], ["Ставка", "rate"],
    ["Зайнятість", "employment"], ["Норма, год", "normHours"], ["Науковий ступінь", "degree"],
    ["Вчене звання", "rank"], ["ДЕНЬ НАР.", "birthDate"], ["Телефон", "phone"], ["Email", "email"],
  ];
  const editInput: React.CSSProperties = { boxSizing: "border-box", width: "100%", minWidth: 0, height: 32, padding: "4px 8px", border: "1px solid #cbd5e1", borderRadius: 5, background: "#fff", color: "#0f172a", fontSize: 12.5 };
  const updateRow = (rowIndex: number, field: keyof TeacherRosterExportRow, value: string) => setRows((current) => current.map((row, index) => index === rowIndex
    ? { ...row, [field]: field === "normHours" ? Number(value) || 0 : value } as TeacherRosterExportRow
    : row));
  const download = async (format: "docx" | "xlsx" | "pdf") => {
    setDownloading(format);
    try { await exportTeacherRosterReport({ ...report, title, subtitle, rows, comment, reportDate, departmentSigner }, format); }
    catch (error) { notify(`Не вдалося сформувати звіт: ${error instanceof Error ? error.message : "невідома помилка"}`, "error", "Звіт викладачів"); }
    finally { setDownloading(""); }
  };
  const restore = () => { setTitle(report.title); setSubtitle(report.subtitle); setRows(report.rows); setComment(""); setReportDate(new Date().toLocaleDateString("uk-UA")); setDepartmentSigner(report.departmentHeadName); };
  const columnWidths = [36, 210, 120, 84, 62, 112, 82, 125, 115, 102, 118, 175];
  return <div className="teacher-report-overlay" role="presentation" style={{ alignItems: "flex-start", padding: "20px 2vw" }} onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
    <section className="teacher-report-modal" role="dialog" aria-modal="true" aria-labelledby="teacher-report-title">
      <header><div><h2 id="teacher-report-title">Попередній перегляд та редагування звіту перед завантаженням</h2><p>{title}</p></div><button type="button" aria-label="Закрити" onClick={onClose}>×</button></header>
      <div className="teacher-report-paper">
        <div className="teacher-report-document-heading">
          <strong>Національний університет біоресурсів і природокористування України</strong>
          <span>{report.facultyName || "Факультет не вказано"}</span><span>{report.departmentName || "Кафедра не вказана"}</span>
          <input aria-label="Назва звіту" style={{ ...editInput, height: 38, maxWidth: 900, margin: "8px auto 0", textAlign: "center", fontWeight: 700, fontSize: 16 }} value={title} onChange={(event) => setTitle(event.target.value)} />
          <input aria-label="Підзаголовок звіту" style={{ ...editInput, maxWidth: 900, margin: "5px auto 0", textAlign: "center" }} value={subtitle} onChange={(event) => setSubtitle(event.target.value)} />
        </div>
        <div className="teacher-report-preview-table"><table><colgroup>{columnWidths.map((width, index) => <col key={index} style={{ width }} />)}</colgroup><thead><tr><th>№</th>{fields.map(([label, field]) => <th key={field}>{label}</th>)}</tr></thead><tbody>
          {rows.length ? rows.map((row, rowIndex) => <tr key={`${row.name}-${rowIndex}`}><td>{rowIndex + 1}</td>{fields.map(([label, field]) => <td key={field}><input aria-label={`${label}, рядок ${rowIndex + 1}`} style={editInput} value={String(row[field] ?? "")} onChange={(event) => updateRow(rowIndex, field, event.target.value)} /></td>)}</tr>) : <tr><td colSpan={12} className="teacher-report-empty">За обраними фільтрами викладачів не знайдено.</td></tr>}
        </tbody></table></div>
        <label className="teacher-report-edit-label">Коментар під таблицею<textarea value={comment} onChange={(event) => setComment(event.target.value)} rows={2} /></label>
        <div className="teacher-report-signatures"><label>Завідувач кафедри<input value={departmentSigner} onChange={(event) => setDepartmentSigner(event.target.value)} placeholder="ПІБ підписанта" /></label><label>Дата<input value={reportDate} onChange={(event) => setReportDate(event.target.value)} /></label></div>
      </div>
      <footer><button type="button" className="teacher-report-reset" disabled={Boolean(downloading)} onClick={restore}>Скинути зміни</button><div><button type="button" disabled={Boolean(downloading)} onClick={() => void download("docx")}>{downloading === "docx" ? "Формування…" : "Завантажити Word"}</button><button type="button" disabled={Boolean(downloading)} onClick={() => void download("xlsx")}>{downloading === "xlsx" ? "Формування…" : "Завантажити Excel"}</button><button type="button" disabled={Boolean(downloading)} onClick={() => void download("pdf")}>{downloading === "pdf" ? "Формування…" : "Завантажити PDF"}</button></div></footer>
    </section>
  </div>;
}

function TeacherForm({ mode, form, teachers, positions, statuses, teacherTypes, degreeOptions, rankOptions, selectedCode, onSelect, onChange, onSave, onCancel, onArchive, saving }: { mode: Mode; form: Teacher; teachers: Teacher[]; positions: Reference[]; statuses: Reference[]; teacherTypes: Reference[]; degreeOptions: string[]; rankOptions: string[]; selectedCode: string; onSelect: (code: string) => void; onChange: (field: keyof Teacher, value: string) => void; onSave: () => void; onCancel: () => void; onArchive: () => void; saving: boolean }) {
  const hasPosition = positions.some((item) => item.code.trim() === form.positionCode.trim());
  const hasType = teacherTypes.some((item) => item.code.trim() === form.workTypeCode.trim());
  const hasStatus = statuses.some((item) => item.code.trim() === form.staffKind.trim());
  return <section className="teacher-card teacher-form"><h2>{mode === "add" ? "Додати викладача" : "Редагувати викладача"}</h2>{mode === "edit" && <label>Оберіть викладача<SearchableSelect ariaLabel="Оберіть викладача для редагування" placeholder="— оберіть —" value={selectedCode} onChange={onSelect} options={teachers.map((teacher) => ({ value: teacher.code, label: teacher.name, sublabel: `${teacher.position || "Посаду не вказано"} · ${teacher.rate} ставки${teacher.workTypeName ? ` · ${teacher.workTypeName}` : ""}` }))} /></label>}<div className="teacher-form-grid">
    <label>ПІБ *<input value={form.name} onChange={(event) => onChange("name", event.target.value)} placeholder="Прізвище Ім'я По батькові" /></label>
    <label>Посада<select value={form.positionCode} onChange={(event) => onChange("positionCode", event.target.value)}><option value="">— оберіть посаду —</option>{!hasPosition && form.positionCode && <option value={form.positionCode}>{form.position || form.positionCode}</option>}{positions.map((item) => <option key={item.code} value={item.code}>{item.name}</option>)}</select></label>
    <label>Частка ставки<input type="number" min="0.05" step="0.05" value={form.rate} onChange={(event) => onChange("rate", event.target.value)} /></label>
    <label>Категорія НПП / ПП<select value={form.staffKind} onChange={(event) => onChange("staffKind", event.target.value)}><option value="">— оберіть категорію —</option>{!hasStatus && form.staffKind && <option value={form.staffKind}>{form.staffKindName || form.staffKind}</option>}{statuses.map((item) => <option key={item.code} value={item.code}>{item.name || item.code}</option>)}</select></label>
    <label>Тип викладача / зайнятість<select value={form.workTypeCode} onChange={(event) => onChange("workTypeCode", event.target.value)}><option value="">— не вказано —</option>{!hasType && form.workTypeCode && <option value={form.workTypeCode}>{form.workTypeName || form.workTypeCode}</option>}{teacherTypes.map((item) => <option key={item.code} value={item.code}>{item.name || item.code}</option>)}</select></label>
    <label>Науковий ступінь<ComboboxInput ariaLabel="Науковий ступінь" value={form.degree} options={degreeOptions} onChange={(value) => onChange("degree", value)} /></label>
    <label>Вчене звання<ComboboxInput ariaLabel="Вчене звання" value={form.rank} options={rankOptions} onChange={(value) => onChange("rank", value)} /></label>
    <label>Телефон<input value={form.phone} onChange={(event) => onChange("phone", event.target.value)} /></label><label>Електронна пошта<input type="email" value={form.email} onChange={(event) => onChange("email", event.target.value)} /></label><label>Дата народження<input type="date" value={form.birthDate} onChange={(event) => onChange("birthDate", event.target.value)} /></label>
  </div><div className="teacher-actions"><button className="primary" disabled={saving} onClick={onSave}>{saving ? "Збереження…" : mode === "add" ? "Додати" : "Зберегти"}</button>{mode === "edit" && <button className="danger" disabled={!form.code || saving} onClick={onArchive}>Звільнити</button>}<button onClick={onCancel}>Скасувати зміни</button></div></section>;
}
function RestorePanel({ teachers, onRestore, onBack, saving }: { teachers: Teacher[]; onRestore: (teacher: Teacher) => void; onBack: () => void; saving: boolean }) { return <section className="teacher-card restore"><div><h2>Поновити викладача</h2><p>Оберіть співробітника, якого потрібно повернути на посаду.</p></div>{teachers.length ? <div className="teacher-actions">{teachers.map((teacher) => <button className="primary" disabled={saving} key={teacher.code} onClick={() => onRestore(teacher)}><span>Поновити: {teacher.name}</span><small style={{ display: "block", marginTop: 4, fontWeight: 400, opacity: .85 }}>{teacher.archiveSummary || [teacher.position, `${teacher.rate} ставки`, teacher.departmentName || teacher.departmentCode].filter(Boolean).join(" · ")}</small></button>)}</div> : <p className="empty-archive">Архівних співробітників немає.</p>}<button onClick={onBack}>Повернутися</button></section>; }
