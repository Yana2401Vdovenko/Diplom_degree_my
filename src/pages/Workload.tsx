import * as React from "react";
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import "../styles/Workload.css";
import "../styles/WorkloadRoleActions.css";
import "../styles/UnifiedSidebar.css";
import { useLocation, useNavigate } from "react-router-dom";
import { FiBookOpen, FiChevronDown, FiClipboard, FiDownload, FiEdit2, FiGrid, FiHelpCircle, FiLoader, FiLogOut, FiPlus, FiPlusCircle, FiRotateCcw, FiSearch, FiTrash2, FiUsers } from "react-icons/fi";
import logo from "../assets/images/logo.png";
import { supabase } from "../utils/supabaseClient";
import { currentUserName } from "../utils/currentUser";
import { ensureCurrentUserDepartment, findDepartmentFacultyCode, findDepartmentHeadName } from "../utils/currentDepartment";
import AppSidebar from "../components/AppSidebar";
import SearchableSelect from "../components/SearchableSelect";
import { FilterReset, FilterSearch, FilterSelect, filterBarStyle } from "../components/TableFilters";
import { exportEditableReport, type DisciplineReportRow, type EditableReport, type ExportWorkload } from "../utils/reportExport";
import { notify } from "../utils/notifications";

export interface Teacher {
  id: number;
  code: string;
  name: string;
  position: string;
  rank?: string;
  employmentType?: string;
  rate: number;
  normHours: number;
  departmentCode: string;
}

export interface Discipline {
  id: number;
  name: string;
  hours: number;
  semester: number;
  semesters?: number[];
  studyForms?: StudyForm[];
  teacherAllocations?: Array<{ code: string; name: string; hours: number }>;
  teacherWorkAssignments?: Array<{ code: string; name: string; workTypes: string[] }>;
  teacher_id?: number | null;
  group_code?: string | null;
  study_form?: StudyForm;
  academic_year?: string;
}

export interface Vacancy {
  id: number;
  discipline_name: string;
  group_code: string;
  work_type: string;
  hours: number;
  is_assigned: boolean;
  study_form?: StudyForm;
  academic_year?: string;
}

export interface WorkloadRecord {
  id: number;
  teacher_id: number;
  teacher_code?: string | null;
  discipline_name: string;
  discipline_code?: string;
  abbreviation?: string;
  groups: string;
  academic_year: string;
  study_form: StudyForm;
  department_code?: string;
  form_code?: string;
  level_code?: string;
  total_hours?: number;
  semester?: number;
  lectures?: number;
  labs?: number;
  practices?: number;
  exams?: number;
  credits?: number;
  course_projects?: number;
  seminars?: number;
  practice_hours?: number;
  diploma_project?: number;
  reviews?: number;
}

type StudyForm = "Денна" | "Заочна" | "Усі форми (Денна + Заочна)";
type FormSelection = "all" | "day" | "extramural";
type Tab = "adjust" | "report" | "form56" | "vacancies" | "disciplines";
type ImportResult = { teachers: number; disciplines: number; workloads: number; details?: number; detailsWarning?: string };
type PlanImportRow = { sourceRow: number; id?: number; discipline: string; disciplineCode?: string; department: string; target: string; abbreviations: string[]; levelKind: "bachelor" | "master"; levelCode?: string; semester: number; semesterHalf?: 1 | 2; lecture: number; practice: number; seminar: number; lab: number; exam: number; credit: number; courseProject: number; practiceHours?: number; diplomaProject?: number; reviews?: number; total: number; teacherCode?: string | null; form_code?: string; study_form?: Exclude<StudyForm, "Усі форми (Денна + Заочна)"> };
type RawPlanImportRow = PlanImportRow & { rawTotal?: number; rawExam?: number; rawCredit?: number };
type SplitWorkKind = "lecture" | "lab" | "practice" | "seminar" | "exam" | "credit" | "courseProject" | "practiceHours" | "diplomaProject" | "reviews";
type SplitPart = { teacherCode: string; hours: number };
type SplitAllocation = Partial<Record<SplitWorkKind, SplitPart[]>>;
const fullAllocationForTeacher = (row: PlanImportRow, teacherCode: string): SplitAllocation => {
  const amounts: Array<[SplitWorkKind, number]> = [
    ["lecture", row.lecture], ["lab", row.lab], ["practice", row.practice], ["seminar", row.seminar], ["exam", row.exam],
    ["credit", row.credit], ["courseProject", row.courseProject], ["practiceHours", row.practiceHours ?? 0],
    ["diplomaProject", row.diplomaProject ?? 0], ["reviews", row.reviews ?? 0],
  ];
  return Object.fromEntries(amounts.filter(([, amount]) => amount > 0).map(([key, hours]) => [key, [{ teacherCode, hours }]])) as SplitAllocation;
};
type WorkloadHourKey = "lectures" | "labs" | "practices" | "seminars" | "exams" | "credits" | "course_projects" | "practice_hours" | "diploma_project" | "reviews";
type WorkloadHourDraft = Record<WorkloadHourKey, number>;
const workloadHourFields: Array<{ key: WorkloadHourKey; column: string; label: string }> = [
  { key: "lectures", column: "Лекції", label: "Лекції" }, { key: "labs", column: "Лабораторні_заняття", label: "Лабораторні" },
  { key: "practices", column: "Практичні_заняття", label: "Практичні" }, { key: "seminars", column: "Семінарні_заняття", label: "Семінарські" },
  { key: "exams", column: "Екзамен", label: "Екзамен" }, { key: "credits", column: "Залік", label: "Залік" },
  { key: "course_projects", column: "Курсовий_проєкт", label: "Курсовий проєкт" }, { key: "practice_hours", column: "Практика", label: "Практика" },
  { key: "diploma_project", column: "Дипломне_проєктування", label: "Дипломні / ЕК" }, { key: "reviews", column: "Рецензії", label: "Рецензії" },
];
const teacherReportColumnPercentages = [2.5, 18, 7, 3.5, 3.8, 3.8, 4.5, 4.5, 3.8, 4.2, 4.3, 5.6, 4.3, 3.8, 5.2, 14.2, 7] as const;
const emptyWorkloadHourDraft = (): WorkloadHourDraft => Object.fromEntries(workloadHourFields.map(({ key }) => [key, 0])) as WorkloadHourDraft;
const hoursFromPlanRow = (row: PlanImportRow): WorkloadHourDraft => ({ lectures: row.lecture, labs: row.lab, practices: row.practice, seminars: row.seminar, exams: row.exam, credits: row.credit, course_projects: row.courseProject, practice_hours: row.practiceHours ?? 0, diploma_project: row.diplomaProject ?? 0, reviews: row.reviews ?? 0 });
const buildDisciplineReportRows = (disciplines: Discipline[], planRows: PlanImportRow[], teachers: Teacher[]): DisciplineReportRow[] => {
  const workTypes: Array<[keyof WorkloadHourDraft, string]> = [
    ["lectures", "Лекції"], ["labs", "Лабораторні"], ["practices", "Практичні"], ["seminars", "Семінари"],
    ["course_projects", "Курсовий"], ["practice_hours", "Практика"], ["diploma_project", "ДП / ЕК"], ["exams", "Екзамен"], ["credits", "Залік"], ["reviews", "Рецензії"],
  ];
  return disciplines.flatMap((discipline) => {
    const disciplineRows = planRows.filter((row) => trimText(row.disciplineCode) === trimText(discipline.group_code));
    const groups = new Map<string, { specialty: string; semester: number; form: string; rows: PlanImportRow[] }>();
    const add = (specialty: string, semester: number, form: string, row?: PlanImportRow) => {
      const key = [specialty, semester, form].join("\u0000");
      const group = groups.get(key) ?? { specialty, semester, form, rows: [] };
      if (row) group.rows.push(row);
      groups.set(key, group);
    };
    disciplineRows.forEach((row) => {
      const specialties = row.abbreviations.length ? row.abbreviations : [row.target || "—"];
      specialties.forEach((specialty) => add(specialty, number(row.semester), row.study_form || "Денна", row));
    });
    if (!disciplineRows.length) {
      const specialties = ["—"];
      const semesters = discipline.semesters?.length ? discipline.semesters : [discipline.semester || 0];
      const forms = discipline.studyForms?.length ? discipline.studyForms.filter((form) => form !== "Усі форми (Денна + Заочна)") : [discipline.study_form || "Денна"];
      specialties.forEach((specialty) => semesters.forEach((semester) => forms.forEach((form) => add(specialty, semester, form))));
    }
    return [...groups.values()].map((group) => {
      const teacherTypes = new Map<string, Set<string>>();
      group.rows.forEach((row) => {
        const teacherCode = trimText(row.teacherCode);
        if (!teacherCode) return;
        const teacherName = teachers.find((teacher) => trimText(teacher.code) === teacherCode)?.name ?? teacherCode;
        const hours = hoursFromPlanRow(row);
        workTypes.forEach(([key, label]) => {
          if (number(hours[key]) <= 0) return;
          const kinds = teacherTypes.get(teacherName) ?? new Set<string>();
          kinds.add(label);
          teacherTypes.set(teacherName, kinds);
        });
      });
      if (!group.rows.length) {
        discipline.teacherWorkAssignments?.filter((assignment) => assignment.name !== "Не призначено").forEach((assignment) => teacherTypes.set(assignment.name, new Set(assignment.workTypes.filter((kind) => !/^(екзамен|залік)$/i.test(kind.trim())))));
      }
      const hasExam = group.rows.some((row) => number(row.exam) > 0);
      const hasCredit = group.rows.some((row) => number(row.credit) > 0);
      return {
        specialty: group.specialty,
        discipline: discipline.name,
        semester: group.semester > 0 ? `${group.semester} семестр` : "",
        forms: group.form ? [group.form] : [],
        control: hasExam && hasCredit ? "Екзамен, Залік" : hasExam ? "Екзамен" : hasCredit ? "Залік" : "—",
        teachers: [...teacherTypes.entries()].map(([name, kinds]) => ({ name, workTypes: [...kinds].filter((kind) => !/^(екзамен|залік)$/i.test(kind.trim())) })),
      };
    });
  });
};
const totalForHourDraft = (draft: WorkloadHourDraft) => roundHours(workloadHourFields.reduce((sum, { key }) => sum + number(draft[key]), 0));
const planRowForExport = (row: PlanImportRow, assignedTeacher = ""): ExportWorkload => {
  const values = hoursFromPlanRow(row);
  return {
    teacher_id: row.id ?? row.sourceRow,
    discipline_name: row.discipline,
    groups: row.abbreviations.join(", ") || row.target || "—",
    semester: row.semester,
    study_form: row.study_form || "Денна",
    total_hours: row.total,
    lectures: values.lectures,
    labs: values.labs,
    practices: values.practices,
    seminars: values.seminars,
    exams: values.exams,
    credits: values.credits,
    course_projects: values.course_projects,
    practice_hours: values.practice_hours,
    diploma_project: values.diploma_project,
    reviews: values.reviews,
    assigned_teacher: assignedTeacher,
  };
};
type ImportOption = { code: string; name: string; programCode?: string; specialtyCode?: string; facultyCode?: string; programName?: string; specialtyName?: string };
type ImportReferences = { departments: ImportOption[]; forms: ImportOption[]; levels: ImportOption[]; abbreviations: ImportOption[] };
type AbbreviationDraft = { abbreviation: string; specialty: string; program: string };
const formModeForCode = (code: string, forms: ImportOption[]): Exclude<FormSelection, "all"> =>
  /заоч/i.test(forms.find((item) => item.code.trim() === code.trim())?.name ?? "") ? "extramural" : "day";
const formCodesForSelection = (selection: string, forms: ImportOption[]) => {
  const day = forms.find((item) => !/заоч/i.test(item.name))?.code.trim();
  const extramural = forms.find((item) => /заоч/i.test(item.name))?.code.trim();
  if (selection === "all") return [day, extramural].filter((code): code is string => Boolean(code));
  if (selection === "day") return day ? [day] : [];
  if (selection === "extramural") return extramural ? [extramural] : [];
  return selection.trim() ? [selection.trim()] : [];
};

const distributionRowKey = (row: PlanImportRow) => String(row.id ?? row.sourceRow);
const abbreviationDefaults: Record<string, { specialty: string; program: string }> = {
  "кн": { specialty: "Комп'ютерні науки", program: "Комп'ютерні науки" },
  "іпз": { specialty: "Інженерія програмного забезпечення", program: "Інженерія програмного забезпечення" },
  "пзіс": { specialty: "Інженерія програмного забезпечення", program: "Програмне забезпечення інформаційних систем" },
  "кі": { specialty: "Комп'ютерна інженерія", program: "Комп'ютерна інженерія" },
  "кіб": { specialty: "Кібербезпека та захист інформації", program: "Кібербезпека та захист інформації" },
  "кб": { specialty: "Кібербезпека та захист інформації", program: "Кібербезпека та захист інформації" },
  "іст": { specialty: "Інформаційні системи та технології", program: "Інформаційні системи та технології" },
  "акіт": { specialty: "Автоматизація, комп'ютерно-інтегровані технології та робототехніка", program: "Автоматизація, комп'ютерно-інтегровані технології та робототехніка" },
  "екк": { specialty: "Економічна кібернетика", program: "Економічна кібернетика" },
  "іуст": { specialty: "Комп'ютерні науки", program: "" },
  "робот": { specialty: "Автоматизація, комп'ютерно-інтегровані технології та робототехніка", program: "" },
  "шіір": { specialty: "Комп'ютерні науки", program: "" },
  "шііз": { specialty: "Комп'ютерні науки", program: "" },
};
const specialtyCodesByAbbreviation: Record<string, string> = {
  "кн": "F3", "іуст": "F3", "шіір": "F3", "шііз": "F3",
  "іпз": "F2", "пзіс": "F2", "кі": "F7",
  "кіб": "F5", "кб": "F5", "іст": "F6",
  "акіт": "G7", "робот": "G7", "екк": "C1", "ек": "C1",
};
const officialSpecialtyCode = (name: string, abbreviation: string) => {
  const aliasCode = specialtyCodesByAbbreviation[normalizedName(abbreviation)];
  if (aliasCode) return aliasCode;
  const normalized = normalizedName(name).replace(/’/g, "'");
  if (normalized.includes("комп'ютерні науки")) return "F3";
  if (normalized.includes("інженерія програмного забезпечення")) return "F2";
  if (normalized.includes("комп'ютерна інженерія")) return "F7";
  if (normalized.includes("кібербезпек")) return "F5";
  if (normalized.includes("інформаційні системи та технології")) return "F6";
  if (normalized.includes("автоматизація") && normalized.includes("робототехніка")) return "G7";
  if (normalized.includes("економік")) return "C1";
  if (normalized.includes("прикладна математика")) return "F1";
  if (normalized.includes("системний аналіз")) return "F4";
  if (normalized.includes("менеджмент")) return "D3";
  if (normalized.includes("маркетинг")) return "D5";
  return "";
};

const TAB_ITEMS: { key: Tab; label: string; icon: React.ReactNode; subtitle: string }[] = [
  { key: "adjust", label: "Розподіл годин", icon: <FiPlusCircle />, subtitle: "Переглядайте та коригуйте навантаження конкретного викладача." },
  { key: "report", label: "Звіт викладача", icon: <FiUsers />, subtitle: "Сформуйте зрозумілий звіт з усіма видами навчальної роботи." },
  { key: "form56", label: "Зведення кафедри", icon: <FiClipboard />, subtitle: "Порівнюйте обсяг роботи всіх викладачів за обраний рік." },
  { key: "vacancies", label: "Вакантні години", icon: <FiSearch />, subtitle: "Знайдіть години, які ще потрібно розподілити." },
  { key: "disciplines", label: "Дисципліни", icon: <FiBookOpen />, subtitle: "Переглядайте дисципліни та відповідальних викладачів." },
];

const number = (value: unknown) => Number(value ?? 0);
const roundHours = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100;
const parseHoursInput = (value: string) => {
  const parsed = Number(value.replace(",", ".").trim());
  return Number.isFinite(parsed) ? roundHours(parsed) : 0;
};
const normalizeColumn = (value: string) => value.toLocaleLowerCase("uk-UA").replace(/[\s_\-./]/g, "");
const trimText = (value: unknown) => String(value ?? "").trim();
const toDbYear = (yearStr: unknown) => {
  const year = trimText(yearStr);
  const fullYear = year.match(/^(\d{4})\/(\d{2})(\d{2})$/);
  if (fullYear) return `${fullYear[1]}/${fullYear[3]}`;
  const shortYear = year.match(/^(\d{4})\/(\d{2})$/);
  return shortYear ? `${shortYear[1]}/${shortYear[2]}` : year;
};
const toDisplayYear = (dbYearStr: unknown) => {
  const year = trimText(dbYearStr);
  const shortYear = year.match(/^(\d{4})\/(\d{2})$/);
  return shortYear ? `${shortYear[1]}/20${shortYear[2]}` : year;
};
const normalizedName = (value: unknown) => trimText(value).toLocaleLowerCase("uk-UA");
const referenceMatch = (value: string) => value.toLocaleLowerCase("uk-UA").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\p{L}\p{N}]/gu, "");
const masterProgramAbbreviations = new Set(["РОБОТ", "ІУСТ", "ПЗІС", "ШІІР", "ШІР"].map(normalizedName));
const importErrorText = (error: unknown, tableName = "Завантаження_навантаження") => {
  const message = error instanceof Error ? error.message : error && typeof error === "object" && "message" in error ? String(error.message) : String(error ?? "Невідома помилка.");
  return /has_permission|role_permissions/i.test(message) && !message.includes("DROP POLICY IF EXISTS")
    ? message + '\nЯкщо помилка спричинена зламаною політикою has_permission, SQL для видалення лише цієї політики: DROP POLICY IF EXISTS has_permission ON public."' + tableName + '";'
    : message;
};
const parseStreamAbbreviations = (value: string) => [...new Set(value.replace(/ІПЗМаг/gi, "ПЗІС").split(/[+,;]+/).map((part) => {
  let clean = part
    .replace(/\([^)]*(?:\d|с\.?\s*т\.?|ст\.?)[^)]*\)/gi, " ")
    .replace(/^\s*\d+\s*/, "")
    .replace(/\s*\d+\s*$/, "")
    .replace(/маг\s*$/i, "")
    .trim();
  if (/ст$/i.test(clean) && !/і(?:у)?ст$/i.test(clean)) clean = clean.replace(/ст$/i, "");
  return clean.replace(/[^\p{L}\p{N}]/gu, "").trim();
}).filter((part) => [...part].length >= 2))];
const formatAbbreviation = (value: string, isMaster: boolean) => parseStreamAbbreviations(value)
  .map((item) => isMaster && normalizedName(item) === "іпз" ? "ПЗІС" : item)
  .join(", ");
const getDisplayAbbreviation = (row: Record<string, unknown>) => {
  const source = trimText(row["Абревіатура"]);
  return formatAbbreviation(source, trimText(row["Код_рівня_навчання"]) === "М") || source;
};
const isMasterLevel = (code: unknown, name = "", abbreviation = "") => /^0*2$/.test(trimText(code))
  || /^м$/i.test(trimText(code))
  || /маг/i.test(name)
  || /маг/i.test(abbreviation)
  || parseStreamAbbreviations(abbreviation).some((item) => masterProgramAbbreviations.has(normalizedName(item)));
const workloadComponentsTotal = (row: WorkloadRecord) => [row.lectures, row.labs, row.practices, row.seminars, row.exams, row.credits, row.course_projects, row.practice_hours, row.diploma_project, row.reviews]
  .reduce<number>((sum, value) => sum + number(value), 0);
const workloadTotal = (row: WorkloadRecord) => {
  const storedTotal = number(row.total_hours);
  return storedTotal === 0 ? workloadComponentsTotal(row) : storedTotal;
};
const hasMissingReviewsColumn = (error: { message?: string } | null | undefined) =>
  Boolean(error?.message && /Рецензії/i.test(error.message) && /column|schema cache|could not find/i.test(error.message));
const formatHours = (value: number): string => new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 2 }).format(Number(value) || 0);
const hours = formatHours;
const largestRemainderIntegers = (values: number[]) => {
  const normalized = values.map((value) => Math.max(0, roundHours(value)));
  const result = normalized.map(Math.floor);
  let remaining = Math.ceil(roundHours(normalized.reduce((sum, value) => sum + value, 0))) - result.reduce((sum, value) => sum + value, 0);
  const order = normalized.map((value, index) => ({ index, fraction: value - Math.floor(value) }))
    .sort((a, b) => b.fraction - a.fraction || a.index - b.index);
  for (const item of order) {
    if (remaining <= 0) break;
    if (item.fraction <= 0) break;
    result[item.index] += 1;
    remaining -= 1;
  }
  return result;
};

export default function Workload() {
  const navigate = useNavigate();
  const location = useLocation();
  const [activeTab, setActiveTab] = useState<Tab | null>(null);
  const [year, setYear] = useState("");
  const [availableYears, setAvailableYears] = useState<string[]>([]);
  useEffect(() => {
    document.body.classList.add("workload-page-active");
    return () => { document.body.classList.remove("workload-page-active", "workload-fullscreen-active"); };
  }, []);
  const [teacherId, setTeacherId] = useState<number | "">("");
  const [teachers, setTeachers] = useState<Teacher[]>([]);
  const [disciplines, setDisciplines] = useState<Discipline[]>([]);
  const [workloads, setWorkloads] = useState<WorkloadRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const assigningId: number | null = null;
  const [role] = useState(() => localStorage.getItem("userRole") ?? "Лаборант");
  const [profileName] = useState(currentUserName);
  const canEditDisciplines = role.toLocaleLowerCase("uk-UA").includes("лаборант");
  const [selectedImportFile, setSelectedImportFile] = useState<string | null>(null);
  const [importing, setImporting] = useState(false);
  const [importMessage, setImportMessageState] = useState<string | null>(null);
  const [planPreview, setPlanPreview] = useState<PlanImportRow[]>([]);
  const [importYear, setImportYear] = useState("");
  const [importDepartment, setImportDepartment] = useState("");
  const [importForm, setImportForm] = useState("");
  const [importReferences, setImportReferences] = useState<ImportReferences>({ departments: [], forms: [], levels: [], abbreviations: [] });
  const [currentDepartment, setCurrentDepartment] = useState("");
  const [departmentName, setDepartmentName] = useState("");
  const [facultyName, setFacultyName] = useState("");
  const [departmentHeadName, setDepartmentHeadName] = useState("");
  const [distributionRows, setDistributionRows] = useState<PlanImportRow[]>([]);
  const [distributionAssignments, setDistributionAssignments] = useState<Record<string, string>>({});
  const [distributionFormCode, setDistributionFormCode] = useState("");
  const selectedFormCodes = formCodesForSelection(distributionFormCode, importReferences.forms);
  const showAllForms = distributionFormCode === "all";
  const [newAbbreviationDrafts, setNewAbbreviationDrafts] = useState<AbbreviationDraft[]>([]);
  const [suggestionRow, setSuggestionRow] = useState<{ sourceRow: number; teacherCode: string } | null>(null);
  const [importVersion, setImportVersion] = useState(0);
  const [isFullscreen, setIsFullscreen] = useState(false);
  useEffect(() => {
    document.body.classList.toggle("workload-fullscreen-active", isFullscreen);
    return () => { document.body.classList.remove("workload-fullscreen-active"); };
  }, [isFullscreen]);
  const importInputRef = useRef<HTMLInputElement>(null);
  const coreDataLoadedRef = useRef(false);
  useEffect(() => {
    if (!isFullscreen) return;
    const exitFullscreen = (event: KeyboardEvent) => { if (event.key === "Escape") setIsFullscreen(false); };
    window.addEventListener("keydown", exitFullscreen);
    return () => window.removeEventListener("keydown", exitFullscreen);
  }, [isFullscreen]);
  useLayoutEffect(() => {
    if (location.pathname === "/workload") setActiveTab(null);
  }, [location.pathname]);
  const setImportMessage = (message: string | null) => {
    setImportMessageState(message);
    if (!message?.trim()) return;
    const type = /помил|не вдалося|не знайдено|error|відрізняється/i.test(message)
      ? "error"
      : /завершено|очищено|збережено|призначено|імпортовано|застосовано/i.test(message) ? "success" : "info";
    notify(message, type, "Навантаження");
  };
  const setError = (message: string | null) => {
    if (message?.trim()) notify(message, "error", "Навантаження");
  };

  const isMissingReviewsColumn = (error: { message?: string } | null | undefined) =>
    Boolean(error?.message && /Рецензії/i.test(error.message) && /column|schema cache|could not find/i.test(error.message));
  const retryWithoutReviews = () => notify("Supabase не має колонки «Рецензії» у схемі. Запис повторено без цього виду годин; інші години збережено.", "warning", "Навантаження");
  const updateWorkloadRow = async (id: number, payload: Record<string, unknown>) => {
    let result = await supabase.from("Завантаження_навантаження").update(payload).eq("id", id);
    if (isMissingReviewsColumn(result.error)) {
      const retryPayload = { ...payload };
      delete retryPayload["Рецензії"];
      result = await supabase.from("Завантаження_навантаження").update(retryPayload).eq("id", id);
      if (!result.error) retryWithoutReviews();
    }
    return result;
  };
  const insertWorkloadRows = async (payload: Record<string, unknown> | Record<string, unknown>[]) => {
    let result = await supabase.from("Завантаження_навантаження").insert(payload);
    if (isMissingReviewsColumn(result.error)) {
      const stripReviews = (row: Record<string, unknown>) => { const next = { ...row }; delete next["Рецензії"]; return next; };
      const retryPayload = Array.isArray(payload) ? payload.map(stripReviews) : stripReviews(payload);
      result = await supabase.from("Завантаження_навантаження").insert(retryPayload);
      if (!result.error) retryWithoutReviews();
    }
    return result;
  };

  useEffect(() => {
    const loadCoreData = async () => {
      if (!coreDataLoadedRef.current) setLoading(true);
      setError(null);
      const [authResult, teacherResult, disciplineResult, teacherInfoResult, normResult, departmentResult, facultyResult, abbreviationResult, specialtyResult, programResult, formResult, levelResult, positionResult, teacherTypeResult] = await Promise.all([
        supabase.auth.getUser(),
        supabase.from("Викладач").select("*").order("ПІБ"),
        supabase.from("Дисципліни").select("*"),
        supabase.from("Кафедральна_інформація_про_виклад").select("*"),
        supabase.from("Навантаження").select("*"),
        supabase.from("Кафедра").select("*"),
        supabase.from("Факультет_ННІ").select("*"),
        supabase.from("Абревіатура").select("*"),
        supabase.from("Спеціальність").select("*"),
        supabase.from("ОПП").select("*"),
        supabase.from("Форма_навчання").select("*"),
        supabase.from("Рівень_навчання").select("*"),
        supabase.from("Посада").select("*"),
        supabase.from("Тип_викладача").select("*"),
      ]);
      const departments = (departmentResult.data ?? []) as unknown as Record<string, unknown>[];
      const teacherInfos = (teacherInfoResult.data ?? []) as unknown as Record<string, unknown>[];
      let profileDepartment = "";
      try {
        profileDepartment = (await ensureCurrentUserDepartment(authResult.data.user)).departmentCode;
      } catch (error) {
        setError(error instanceof Error ? `Не вдалося визначити кафедру користувача: ${error.message}` : "Не вдалося визначити кафедру користувача.");
      }
      if (!profileDepartment) {
        setError("Для цього облікового запису не знайдено прив’язану кафедру. Перевірте профіль користувача та дані кафедральної прив’язки.");
        setLoading(false);
        return;
      }
      if (profileDepartment !== currentDepartment.trim()) {
        setCurrentDepartment(profileDepartment);
        localStorage.setItem("workload-current-department", profileDepartment);
      }
      const activeDepartment = profileDepartment;
      const yearResult = await supabase.from("Завантаження_навантаження").select("Рік,Код_кафедри").eq("Код_кафедри", activeDepartment);
      const initialWorkloadResult = await supabase.from("Завантаження_навантаження").select("id,Рік,Код_кафедри,Код_форми_навчання,Код_рівня_навчання,Код_дисципліни,Код_викладача,Абревіатура,Семестр,Лекції,Лабораторні_заняття,Практичні_заняття,Семінарні_заняття,Екзамен,Залік,Курсовий_проєкт,Практика,Дипломне_проєктування,Рецензії,Всього_годин").eq("Рік", toDbYear(year)).eq("Код_кафедри", activeDepartment);
      const workloadResult = isMissingReviewsColumn(initialWorkloadResult.error)
        ? await supabase.from("Завантаження_навантаження").select("id,Рік,Код_кафедри,Код_форми_навчання,Код_рівня_навчання,Код_дисципліни,Код_викладача,Абревіатура,Семестр,Лекції,Лабораторні_заняття,Практичні_заняття,Семінарні_заняття,Екзамен,Залік,Курсовий_проєкт,Практика,Дипломне_проєктування,Всього_годин").eq("Рік", toDbYear(year)).eq("Код_кафедри", activeDepartment)
        : initialWorkloadResult;
      if (yearResult.error) setError(yearResult.error.message);
      else {
        const actualYears = [...new Set(((yearResult.data ?? []) as unknown as Record<string, unknown>[])
          .filter((row) => trimText(row["Код_кафедри"]) === activeDepartment)
          .map((row) => toDisplayYear(trimText(row["Рік"]))).filter(Boolean))]
          .sort((left, right) => toDbYear(right).localeCompare(toDbYear(left)));
        setAvailableYears(actualYears);
        setYear((current) => actualYears.includes(toDisplayYear(current)) ? toDisplayYear(current) : actualYears[0] ?? "");
      }
      const firstError = teacherResult.error ?? disciplineResult.error ?? workloadResult.error ?? teacherInfoResult.error ?? positionResult.error;
      if (firstError) setError(firstError.message);
      else {
        const positionRows = (positionResult.data ?? []) as unknown as Record<string, unknown>[];
        const teacherTypeRows = (teacherTypeResult.data ?? []) as unknown as Record<string, unknown>[];
        const norms = Object.fromEntries(((normResult.data ?? []) as unknown as Record<string, unknown>[]).map((item) => [trimText(item["Код_статусу_викладача"]), number(item["Години"])]));
        const existingDepartment = departments.find((item) => trimText(item["Код_кафедри"]) === activeDepartment);
        const facultyRows = (facultyResult.data ?? []) as unknown as Record<string, unknown>[];
        const facultyCode = findDepartmentFacultyCode(existingDepartment);
        const facultyRecord = facultyRows.find((item) => trimText(item["Код_факультету_нні"] ?? item["Код_факультету"]) === facultyCode);
        const mappedFacultyName = trimText(facultyRecord?.["Назва_факультету_нні"] ?? facultyRecord?.["Назва_факультету"] ?? facultyRecord?.["Назва"]);
        const nextTeachers = (teacherResult.data ?? []).map((teacher, index) => {
          const code = trimText(teacher["Код_викладача"]);
          const info = teacherInfos.find((item) => trimText(item["Код_викладача"]) === code) ?? {};
          const statusCode = trimText(info["Код_статусу_викладача"]);
          const positionCode = trimText(info["Код_посади"]);
          const positionFromReference = positionRows.find((item) => trimText(item["Код_посади"]) === positionCode);
          const position = trimText(info["Посада"])
            || trimText(positionFromReference?.["Назва_посади"] ?? positionFromReference?.["Назва"] ?? positionFromReference?.["Посада"])
            || "—";
          const typeCode = trimText(info["Код_типу_викладача"] ?? info["Код_типу"]);
          const typeRow = teacherTypeRows.find((item) => trimText(item["Код_типу_викладача"] ?? item["Код_типу"]) === typeCode);
          const employmentType = trimText(typeRow?.["Назва_типу"]) || trimText(typeRow?.["Назва_типу_викладача"]) || trimText(typeRow?.["Тип_викладача"]) || trimText(typeRow?.["Назва"]) || trimText(info["Тип_викладача"]) || "—";
          return { id: index + 1, code, name: String(teacher["ПІБ"] ?? "Без імені"), position, rank: trimText(teacher["Вчене_звання"]) || "—", employmentType, rate: number(info["Частка_ставки"] ?? 1), normHours: number(norms[statusCode]) || 600, departmentCode: trimText(info["Код_кафедри"]) };
        }).filter((teacher) => trimText(teacher.departmentCode) === activeDepartment) as Teacher[];
        setTeachers(nextTeachers);
        setDepartmentName(trimText(existingDepartment?.["Назва_кафедри"]) || "Кафедра");
        setFacultyName(mappedFacultyName || "Факультет не вказано");
        setDepartmentHeadName(findDepartmentHeadName(existingDepartment, nextTeachers));
        const specialtyRows = (specialtyResult.data ?? []) as unknown as Record<string, unknown>[];
        const programRows = (programResult.data ?? []) as unknown as Record<string, unknown>[];
        const programOptions = programRows.map((item) => ({ code: trimText(item["Код_опп"]), name: trimText(item["Назва_опп"]), specialtyCode: trimText(item["Код_спеціальності"]) }));
        const specialtyOptions = specialtyRows.map((item) => ({ code: trimText(item["Код_спеціальності"]), name: trimText(item["Назва_спеціальності"]) }));
        const abbreviationOptions = ((abbreviationResult.data ?? []) as unknown as Record<string, unknown>[]).map((item) => {
          const programCode = trimText(item["Код_опп"]);
          const specialtyCode = trimText(item["Код_спеціальності"]);
          return { code: trimText(item["Код_абревіатури"]), name: trimText(item["Назва_абревіатури"]), programCode, specialtyCode, programName: programOptions.find((entry) => entry.code === programCode)?.name, specialtyName: specialtyOptions.find((entry) => entry.code === specialtyCode)?.name };
        }).filter((item) => item.code && item.name);
        setImportReferences({
          departments: departments.filter((item) => trimText(item["Код_кафедри"]) === activeDepartment).map((item) => ({ code: trimText(item["Код_кафедри"]), name: trimText(item["Назва_кафедри"]), facultyCode: findDepartmentFacultyCode(item) })).filter((item) => item.code),
          forms: ((formResult.data ?? []) as unknown as Record<string, unknown>[]).map((item) => ({ code: trimText(item["Код_форми_навчання"]), name: trimText(item["Назва_форми_навчання"]) })).filter((item) => item.code),
          levels: ((levelResult.data ?? []) as unknown as Record<string, unknown>[]).map((item) => ({ code: trimText(item["Код_рівня_навчання"]), name: trimText(item["Назва_рівня_навчання"]) })).filter((item) => item.code),
          abbreviations: abbreviationOptions,
        });
        const disciplineRows = (disciplineResult.data ?? []) as unknown as Record<string, unknown>[];
        const formRows = (formResult.data ?? []) as unknown as Record<string, unknown>[];
        const workloadFormName = (code: string) => trimText(formRows.find((item) => trimText(item["Код_форми_навчання"]) === code)?.["Назва_форми_навчання"]);
        const levelNameByCode = new Map(((levelResult.data ?? []) as unknown as Record<string, unknown>[]).map((item) => [trimText(item["Код_рівня_навчання"]), trimText(item["Назва_рівня_навчання"])]));
        const mappedRecords = ((workloadResult.data ?? []) as unknown as Record<string, unknown>[])
          .filter((workload) => toDbYear(workload["Рік"]) === toDbYear(year) && trimText(workload["Код_кафедри"]) === activeDepartment)
          .map((workload) => {
            const teacherCode = trimText(workload["Код_викладача"]);
            const disciplineCode = trimText(workload["Код_дисципліни"]);
            const levelCode = trimText(workload["Код_рівня_навчання"]);
            const abbreviation = getDisplayAbbreviation(workload);
            const teacher = nextTeachers.find((item) => item.code === teacherCode);
            const discipline = disciplineRows.find((item) => trimText(item["Код_дисципліни"]) === disciplineCode);
            return {
              id: number(workload.id), teacher_id: teacher?.id ?? 0, teacher_code: teacherCode || null,
              discipline_code: disciplineCode, discipline_name: trimText(discipline?.["Назва_дисципліни"]) || "Без назви",
              groups: abbreviation, abbreviation,
              academic_year: trimText(workload["Рік"]), department_code: trimText(workload["Код_кафедри"]), form_code: trimText(workload["Код_форми_навчання"]),
              level_code: levelCode, study_form: /заоч/i.test(workloadFormName(trimText(workload["Код_форми_навчання"]))) ? "Заочна" as const : "Денна" as const,
              semester: number(workload["Семестр"]), lectures: number(workload["Лекції"]), labs: number(workload["Лабораторні_заняття"]),
              practices: number(workload["Практичні_заняття"]), seminars: number(workload["Семінарні_заняття"]), exams: number(workload["Екзамен"]),
              credits: number(workload["Залік"]), course_projects: number(workload["Курсовий_проєкт"]), practice_hours: number(workload["Практика"]),
              diploma_project: number(workload["Дипломне_проєктування"]), reviews: number(workload["Рецензії"]),
              total_hours: workload["Всього_годин"] == null ? undefined : number(workload["Всього_годин"]),
            } as WorkloadRecord;
          });
        setWorkloads(mappedRecords);
        const loadedForms = formRows.map((item) => ({ code: trimText(item["Код_форми_навчання"]), name: trimText(item["Назва_форми_навчання"]) })).filter((item) => item.code);
        const existingSelection: FormSelection = ["all", "day", "extramural"].includes(distributionFormCode)
          ? distributionFormCode as FormSelection
          : formModeForCode(distributionFormCode || loadedForms[0]?.code || "", loadedForms);
        if (existingSelection !== distributionFormCode) setDistributionFormCode(existingSelection);
        const activeFormCodes = formCodesForSelection(existingSelection, loadedForms);
        const recordsByDiscipline = new Map<string, WorkloadRecord[]>();
        mappedRecords.filter((record) => activeFormCodes.includes(trimText(record.form_code))).forEach((record) => {
          const code = trimText(record.discipline_code);
          if (!code) return;
          const group = recordsByDiscipline.get(code) ?? [];
          group.push(record);
          recordsByDiscipline.set(code, group);
        });
        const mappedDisciplines = disciplineRows.filter((discipline) => trimText(discipline["Код_кафедри"]) === activeDepartment).map((discipline, index) => {
          const code = trimText(discipline["Код_дисципліни"]);
          const activeRecords = recordsByDiscipline.get(code) ?? [];
          const semesters = [...new Set(activeRecords.map((record) => number(record.semester)).filter((semester) => semester > 0))].sort((a, b) => a - b);
          const teacherHours = new Map<string, number>();
          activeRecords.forEach((record) => {
            const teacherCode = trimText(record.teacher_code);
            if (teacherCode) teacherHours.set(teacherCode, (teacherHours.get(teacherCode) ?? 0) + workloadTotal(record));
          });
          const teacherAllocations = [...teacherHours].map(([teacherCode, allocatedHours]) => ({
            code: teacherCode,
            name: nextTeachers.find((teacher) => teacher.code === teacherCode)?.name ?? teacherCode,
            hours: allocatedHours,
          }));
          const teacherWorkTypes = new Map<string, Set<string>>();
          activeRecords.forEach((record) => {
            const teacherCode = trimText(record.teacher_code);
            if (!teacherCode) return;
            workloadHourFields.forEach(({ key, label }) => {
              if (number(record[key]) <= 0) return;
              const workTypes = teacherWorkTypes.get(teacherCode) ?? new Set<string>();
              workTypes.add(label);
              teacherWorkTypes.set(teacherCode, workTypes);
            });
          });
          const teacherWorkAssignments = [...teacherWorkTypes].map(([teacherCode, workTypes]) => ({
            code: teacherCode,
            name: nextTeachers.find((teacher) => teacher.code === teacherCode)?.name ?? teacherCode,
            workTypes: [...workTypes],
          }));
          return {
            id: index + 1,
            name: trimText(discipline["Назва_дисципліни"]) || "Без назви",
            hours: activeRecords.reduce((sum, record) => sum + workloadTotal(record), 0),
            semester: semesters[0] ?? 0,
            semesters,
            studyForms: [...new Set(activeRecords.map((record) => record.study_form))],
            teacherAllocations,
            teacherWorkAssignments,
            group_code: code,
          } as Discipline;
        }).sort((left, right) => {
          const leftPriority = Number(left.hours > 0) + Number(Boolean(left.teacherAllocations?.length));
          const rightPriority = Number(right.hours > 0) + Number(Boolean(right.teacherAllocations?.length));
          return rightPriority - leftPriority || left.name.localeCompare(right.name, "uk-UA");
        });
        setDisciplines(mappedDisciplines);
        const selectedRecords = mappedRecords.filter((record) => activeFormCodes.includes(trimText(record.form_code)));
        const byIdRows = selectedRecords.map((record) => ({ sourceRow: record.id, id: record.id, discipline: record.discipline_name, disciplineCode: record.discipline_code, form_code: record.form_code, study_form: record.study_form === "Заочна" ? "Заочна" as const : "Денна" as const,
          department: record.department_code ?? "", target: record.abbreviation ?? "", abbreviations: parseStreamAbbreviations(record.abbreviation ?? ""),
          levelKind: isMasterLevel(record.level_code, levelNameByCode.get(trimText(record.level_code)), record.abbreviation ?? "") ? "master" as const : "bachelor" as const, levelCode: record.level_code, semester: number(record.semester),
          lecture: number(record.lectures), practice: number(record.practices), seminar: number(record.seminars), lab: number(record.labs), exam: number(record.exams), credit: number(record.credits), courseProject: number(record.course_projects),
          practiceHours: number(record.practice_hours), diplomaProject: number(record.diploma_project), reviews: number(record.reviews), total: workloadTotal(record), teacherCode: record.teacher_code }));
        setDistributionRows(byIdRows);
        setDistributionAssignments(Object.fromEntries(byIdRows.filter((row) => row.teacherCode).map((row) => [distributionRowKey(row), row.teacherCode as string])));
        setTeacherId((current) => current || nextTeachers[0]?.id || "");
      }
      coreDataLoadedRef.current = true;
      setLoading(false);
    };
    void loadCoreData();
  }, [year, currentDepartment, importVersion, distributionFormCode]);

  useEffect(() => { setSuggestionRow(null); }, [year, currentDepartment, importVersion, distributionFormCode]);

  const selectedTeacher = teachers.find((teacher) => teacher.id === teacherId);
  const teacherLoads = workloads.filter((row) => row.teacher_id === teacherId && selectedFormCodes.includes(trimText(row.form_code)));
  const activeStudyForm: StudyForm = showAllForms ? "Усі форми (Денна + Заочна)" : distributionFormCode === "extramural" ? "Заочна" : "Денна";
  const previewRowsForLevel = planPreview;
  const importTotalsByLevel = (["bachelor", "master"] as const).map((levelKind) => ({
    levelKind,
    rows: planPreview.filter((row) => row.levelKind === levelKind),
  })).filter((group) => group.rows.length).map((group) => ({
    ...group,
    totals: group.rows.reduce((sum, row) => ({
      lectures: sum.lectures + row.lecture, laboratories: sum.laboratories + row.lab, practices: sum.practices + row.practice,
      seminars: sum.seminars + row.seminar, exams: sum.exams + row.exam, credits: sum.credits + row.credit,
      courseProjects: sum.courseProjects + row.courseProject, practiceHours: sum.practiceHours + (row.practiceHours ?? 0),
      diplomaProject: sum.diplomaProject + (row.diplomaProject ?? 0), reviews: sum.reviews + (row.reviews ?? 0), total: sum.total + row.total,
    }), { lectures: 0, laboratories: 0, practices: 0, seminars: 0, exams: 0, credits: 0, courseProjects: 0, practiceHours: 0, diplomaProject: 0, reviews: 0, total: 0 }),
  }));
  const importSemesterTotals = ([1, 2] as const).map((half) => ({
    half,
    total: planPreview.filter((row) => (row.semesterHalf ?? (row.semester % 2 === 1 ? 1 : 2)) === half).reduce((sum, row) => sum + row.total, 0),
  }));
  const importWorkbookTotal = planPreview.reduce((sum, row) => sum + row.total, 0);
  const previewAbbreviations = [...new Set(previewRowsForLevel.flatMap((row) => row.abbreviations))];
  const recognizedAbbreviations = previewAbbreviations.map((name) => ({
    name,
    record: importReferences.abbreviations.find((option) => option.name.trim().toLocaleLowerCase("uk-UA") === name.trim().toLocaleLowerCase("uk-UA")),
  }));
  const recognizedProgramCount = new Set(recognizedAbbreviations.map((item) => trimText(item.record?.programCode)).filter(Boolean)).size;

  const assignVacancy = async (vacancy: Vacancy) => {
    if (distributionRows.some((row) => row.sourceRow === vacancy.id)) setActiveTab("adjust");
    else setError("Детальний перелік вакансій для цього навантаження відсутній. Імпортуйте Excel файл.");
  };

  const changeTab = (tab: Tab) => {
    setActiveTab(tab);
  };

  const selectImportFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setSelectedImportFile(file.name);
    setPlanPreview([]);
    setImporting(true);
    setImportMessage(null);
    setError(null);
    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!firstSheet) throw new Error("У файлі не знайдено жодного аркуша.");
      const rawRows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: "" });
      if (!rawRows.length) throw new Error("Обраний аркуш не містить рядків з даними.");
      const title = rawRows.slice(0, 7).flat().map(String).join(" ");
      if (!title.toLocaleLowerCase("uk-UA").includes("робочих навчальних планів")) {
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, { defval: "" });
        const result = await importWorkbookRows(rows);
        setImportMessage("Імпорт завершено: викладачів — " + result.teachers + ", дисциплін — " + result.disciplines + ", підсумкових записів навантаження — " + result.workloads + ".");
        setSelectedImportFile(null);
        setImportVersion((version) => version + 1);
        return;
      }
      const numeric = (value: unknown) => {
        const parsed = typeof value === "number" ? value : Number(trimText(value).replace(/\s/g, "").replace(",", "."));
        return Number.isFinite(parsed) ? Math.ceil(parsed) : 0;
      };
      const exactNumeric = (value: unknown) => {
        const parsed = typeof value === "number" ? value : Number(trimText(value).replace(/\s/g, "").replace(",", "."));
        return Number.isFinite(parsed) ? parsed : 0;
      };
      const semesterColumn = rawRows.slice(3, 6).flatMap((row) => row.map((cell, column) => ({ cell: trimText(cell).toLocaleLowerCase("uk-UA"), column })))
        .find(({ cell }) => /^(номер\s+)?семестр(у)?$/.test(cell))?.column;
      const reviewsColumn = rawRows.slice(3, 6).flatMap((row) => row.map((cell, column) => ({ cell: trimText(cell).toLocaleLowerCase("uk-UA"), column })))
        .find(({ cell }) => /реценз/i.test(cell))?.column ?? 28;
      const parsedRaw: RawPlanImportRow[] = rawRows.slice(6).flatMap((row, index) => {
        const discipline = trimText(row[1]);
        const reviews = numeric(row[reviewsColumn]);
        const rawExcelTotal = exactNumeric(row[13]);
        const rawExam = exactNumeric(row[22]);
        const rawCredit = exactNumeric(row[23]);
        const rawWorkTotal = exactNumeric(row[15]) + exactNumeric(row[17]) + exactNumeric(row[19]) + exactNumeric(row[21])
          + rawExam + rawCredit + exactNumeric(row[26]) + exactNumeric(row[25]) + exactNumeric(row[27]) + exactNumeric(row[reviewsColumn]);
        const rawTotal = rawExcelTotal === 0 ? rawWorkTotal : rawExcelTotal;
        const target = trimText(row[3]).replace(/ІПЗМаг/gi, "ПЗІС");
        const semester = numeric(row[semesterColumn ?? 31]);
        const semesterMarker = semesterColumn === undefined ? "" : [row[semesterColumn - 1], row[semesterColumn + 1]]
          .map((value) => trimText(value).toLocaleLowerCase("uk-UA"))
          .find((value) => /^(?:н|п|непарн(?:ий|а)|парн(?:ий|а))$/.test(value)) ?? "";
        const semesterHalf: 1 | 2 | undefined = /^(?:н|непарн)/.test(semesterMarker) ? 1
          : /^(?:п|парн)/.test(semesterMarker) ? 2
            : semester > 0 ? semester % 2 === 1 ? 1 : 2 : undefined;
        const abbreviationParts = parseStreamAbbreviations(target);
        const rowText = row.map(trimText).join(" ");
        const isMaster = /магістр|магістрат|ос\s*маг|(?:^|\W)мр(?:$|\W)/i.test(rowText) || abbreviationParts.some((name) => masterProgramAbbreviations.has(normalizedName(name)));
        const abbreviations = abbreviationParts.map((name) => isMaster && normalizedName(name) === "іпз" ? "ПЗІС" : name);
        const serviceLine = /всього|разом|ітого|за семестр|за курс|у тому числі|примітка|завідувач|декан/i.test(discipline);
        if (!discipline || serviceLine || rawTotal <= 0 || (semester <= 0 && abbreviations.length === 0)) return [];
        return [{
          sourceRow: index + 7, discipline, department: trimText(row[2]), target,
          abbreviations, levelKind: isMaster ? "master" as const : "bachelor" as const, semester, semesterHalf, total: rawTotal,
          lecture: numeric(row[15]), practice: numeric(row[17]), seminar: numeric(row[19]), lab: numeric(row[21]),
          exam: rawExam, credit: rawCredit, practiceHours: numeric(row[25]), courseProject: numeric(row[26]),
          diplomaProject: numeric(row[27]), reviews, rawTotal, rawExam, rawCredit,
        }];
      });
      const allocatedTotals = largestRemainderIntegers(parsedRaw.map((row) => row.rawTotal ?? 0));
      const allocatedExams = largestRemainderIntegers(parsedRaw.map((row) => row.rawExam ?? 0));
      const allocatedCredits = largestRemainderIntegers(parsedRaw.map((row) => row.rawCredit ?? 0));
      const parsed: PlanImportRow[] = parsedRaw.map((row, index) => {
        const cleanRow = { ...row, total: allocatedTotals[index], exam: allocatedExams[index], credit: allocatedCredits[index] };
        delete cleanRow.rawTotal;
        delete cleanRow.rawExam;
        delete cleanRow.rawCredit;
        return cleanRow;
      });
      if (!parsed.length) throw new Error("У файлі не знайдено дисциплін із ненульовою кількістю годин.");
      const [departmentsResult, formsResult, levelsResult, abbreviationsResult, specialtiesResult, programsResult] = await Promise.all([
        supabase.from("Кафедра").select("*"), supabase.from("Форма_навчання").select("*"),
        supabase.from("Рівень_навчання").select("*"), supabase.from("Абревіатура").select("*"),
        supabase.from("Спеціальність").select("*"), supabase.from("ОПП").select("*"),
      ]);
      const refError = departmentsResult.error ?? formsResult.error ?? levelsResult.error ?? abbreviationsResult.error ?? specialtiesResult.error ?? programsResult.error;
      if (refError) throw new Error("Не вдалося завантажити довідники: " + refError.message);
      const asOptions = (data: unknown[] | null, codeField: string, nameField: string): ImportOption[] => (data ?? []).map((value) => {
        const row = value as Record<string, unknown>;
        return { code: trimText(row[codeField]), name: trimText(row[nameField]) };
      }).filter((option) => option.code && option.name);
      const specialties = asOptions(specialtiesResult.data, "Код_спеціальності", "Назва_спеціальності").map((option) => ({ ...option, facultyCode: trimText((specialtiesResult.data ?? []).map((row) => row as Record<string, unknown>).find((row) => trimText(row["Код_спеціальності"]) === option.code)?.["Код_факультету_нні"]) }));
      const programs = asOptions(programsResult.data, "Код_опп", "Назва_опп").map((option) => {
        const row = (programsResult.data ?? []).map((value) => value as Record<string, unknown>).find((item) => trimText(item["Код_опп"]) === option.code) ?? {};
        return { ...option, specialtyCode: trimText(row["Код_спеціальності"]), facultyCode: trimText(row["Код_факультету_нні"]) };
      });
      const references: ImportReferences = {
        departments: asOptions(departmentsResult.data, "Код_кафедри", "Назва_кафедри").map((option) => ({ ...option, facultyCode: trimText((departmentsResult.data ?? []).map((value) => value as Record<string, unknown>).find((row) => trimText(row["Код_кафедри"]) === option.code)?.["Код_факультету_нні"]) })),
        forms: asOptions(formsResult.data, "Код_форми_навчання", "Назва_форми_навчання"),
        levels: asOptions(levelsResult.data, "Код_рівня_навчання", "Назва_рівня_навчання"),
        abbreviations: (abbreviationsResult.data ?? []).map((value) => {
          const row = value as Record<string, unknown>;
          const programCode = trimText(row["Код_опп"]);
          const specialtyCode = trimText(row["Код_спеціальності"]);
          return { code: trimText(row["Код_абревіатури"]), name: trimText(row["Назва_абревіатури"]), programCode, specialtyCode, facultyCode: trimText(row["Код_факультету_нні"]), programName: programs.find((item) => item.code === programCode)?.name, specialtyName: specialties.find((item) => item.code === specialtyCode)?.name };
        }).filter((option) => option.code && option.name),
      };
      setImportReferences(references);
      const existingAbbreviations = new Set(references.abbreviations.map((item) => normalizedName(item.name)));
      setNewAbbreviationDrafts([...new Map(parsed.flatMap((row) => row.abbreviations).filter((name) => !existingAbbreviations.has(normalizedName(name))).map((abbreviation) => {
        const defaults = abbreviationDefaults[normalizedName(abbreviation)] ?? { specialty: "", program: "" };
        return [normalizedName(abbreviation), { abbreviation, specialty: defaults.specialty, program: defaults.program } as AbbreviationDraft] as const;
      }))].map(([, draft]) => draft));
      setImportYear(toDisplayYear(title.match(/20\d{2}\s*\/\s*20\d{2}/)?.[0].replace(/\s/g, "") ?? year));
      const fileDepartment = parsed.find((row) => row.department)?.department.trim() ?? "";
      setImportDepartment(references.departments.some((option) => option.code.trim() === fileDepartment) ? fileDepartment : "");
      const formLabel = title.toLocaleLowerCase("uk-UA");
      const detectedForm = references.forms.find((option) => formLabel.includes(option.name.toLocaleLowerCase("uk-UA")));
      setImportForm(detectedForm?.code.trim() ?? "");
      setPlanPreview(parsed);
      const workbookTotal = parsed.reduce((sum, row) => sum + row.total, 0);
      setImportMessage("Розпізнано " + parsed.length + " рядків · Загальний обсяг: " + hours(workbookTotal) + " год");
    } catch (importError) {
      setImportMessage(importError instanceof Error ? importError.message : "Не вдалося обробити файл.");
    } finally {
      setImporting(false);
      event.target.value = "";
    }
  };

  const saveNewAbbreviationMemory = async (departmentCode: string) => {
    if (!newAbbreviationDrafts.length) return;
    const department = importReferences.departments.find((option) => option.code.trim() === departmentCode.trim());
    if (!department?.facultyCode?.trim()) throw new Error("Для обраної кафедри не знайдено Код_факультету_нні. Нові спеціальності та ОПП не записано.");
    const [specialtyResult, programResult, abbreviationResult] = await Promise.all([
      supabase.from("Спеціальність").select("*"), supabase.from("ОПП").select("*"), supabase.from("Абревіатура").select("*"),
    ]);
    const readError = specialtyResult.error ?? programResult.error ?? abbreviationResult.error;
    if (readError) throw new Error("Не вдалося перевірити довідники спеціальностей/ОПП: " + readError.message);
    const specialties = (specialtyResult.data ?? []) as unknown as Record<string, unknown>[];
    const programs = (programResult.data ?? []) as unknown as Record<string, unknown>[];
    const abbreviations = (abbreviationResult.data ?? []) as unknown as Record<string, unknown>[];
    const fieldChecks: Array<["Спеціальність" | "ОПП" | "Абревіатура", string[]]> = [
      ["Спеціальність", ["Код_спеціальності", "Назва_спеціальності", "Код_факультету_нні"]],
      ["ОПП", ["Код_опп", "Назва_опп", "Код_спеціальності", "Код_факультету_нні"]],
      ["Абревіатура", ["Код_абревіатури", "Назва_абревіатури", "Код_опп", "Код_спеціальності", "Код_факультету_нні"]],
    ];
    for (const [table, columns] of fieldChecks) {
      const probe = await supabase.from(table).select(columns.join(",")).limit(0);
      if (probe.error) throw new Error("У таблиці «" + table + "» не підтверджено наявні колонки (" + columns.join(", ") + "): " + probe.error.message + ". Довідники не записано.");
    }
    const nextSpecialtyCode = () => {
      for (let index = 1; index <= 99; index += 1) {
        const candidate = ("S" + String(index).padStart(2, "0")).trim().slice(0, 3);
        if (!usedSpecialties.has(candidate.toLocaleUpperCase("uk-UA"))) { usedSpecialties.add(candidate.toLocaleUpperCase("uk-UA")); return candidate; }
      }
      throw new Error("Не залишилося вільних кодів спеціальності S01–S99.");
    };
    const nextProgramCode = () => {
      const candidates = [
        ...Array.from({ length: 99 }, (_, index) => String(index + 1).padStart(2, "0")),
        ..."ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("").flatMap((letter) => Array.from({ length: 9 }, (_, index) => letter + String(index + 1))),
      ];
      for (const candidateValue of candidates) {
        const candidate = candidateValue.trim().slice(0, 2);
        if (!usedPrograms.has(candidate)) { usedPrograms.add(candidate); return candidate; }
      }
      throw new Error("Не залишилося вільних 2-символьних кодів для ОПП.");
    };
    const nextAbbreviationCode = () => {
      for (const symbol of "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") {
        const candidate = symbol.trim().slice(0, 1);
        if (!usedAbbreviations.has(candidate)) { usedAbbreviations.add(candidate); return candidate; }
      }
      throw new Error("Не залишилося вільних 1-символьних кодів для абревіатур.");
    };
    const usedSpecialties = new Set(specialties.map((row) => trimText(row["Код_спеціальності"]).toLocaleUpperCase("uk-UA")));
    const usedPrograms = new Set(programs.map((row) => trimText(row["Код_опп"])));
    const usedAbbreviations = new Set(abbreviations.map((row) => trimText(row["Код_абревіатури"])));
    const facultyCode = department.facultyCode.trim();
    let added = 0;
    for (const draft of newAbbreviationDrafts) {
      const abbreviation = draft.abbreviation.trim();
      const existingAbbreviation = abbreviations.find((row) => normalizedName(row["Назва_абревіатури"]) === normalizedName(abbreviation));
      if (existingAbbreviation) continue;
      const matchingSpecialtyByName = specialties.find((row) => referenceMatch(trimText(row["Назва_спеціальності"])) === referenceMatch(draft.specialty));
      const mappedCode = officialSpecialtyCode(draft.specialty, abbreviation).trim().slice(0, 3);
      const matchingSpecialtyByCode = mappedCode
        ? specialties.find((row) => trimText(row["Код_спеціальності"]).trim().slice(0, 3).toLocaleUpperCase("uk-UA") === mappedCode.toLocaleUpperCase("uk-UA"))
        : undefined;
      const existingSpecialty = matchingSpecialtyByCode ?? matchingSpecialtyByName;
      let specialtyCode = trimText(existingSpecialty?.["Код_спеціальності"]).trim().slice(0, 3) || mappedCode;
      if (!specialtyCode) specialtyCode = nextSpecialtyCode();
      specialtyCode = specialtyCode.trim().slice(0, 3);
      if (!existingSpecialty) {
        const specialtyPayload = { "Код_спеціальності": specialtyCode, "Назва_спеціальності": draft.specialty.trim(), "Код_факультету_нні": facultyCode };
        const { error } = await supabase.from("Спеціальність").insert(specialtyPayload);
        if (error) throw new Error("Не вдалося зберегти спеціальність «" + draft.specialty + "»: " + error.message);
        specialties.push(specialtyPayload);
        usedSpecialties.add(specialtyCode.toLocaleUpperCase("uk-UA"));
      }
      const matchingProgram = programs.find((row) => referenceMatch(trimText(row["Назва_опп"])) === referenceMatch(draft.program));
      const programCode = (trimText(matchingProgram?.["Код_опп"]) || nextProgramCode()).trim().slice(0, 2);
      if (!matchingProgram) {
        const programPayload = { "Код_опп": programCode, "Назва_опп": draft.program.trim(), "Код_спеціальності": specialtyCode, "Код_факультету_нні": facultyCode };
        const { error } = await supabase.from("ОПП").insert(programPayload);
        if (error) throw new Error("Не вдалося зберегти ОПП «" + draft.program + "»: " + error.message);
        programs.push(programPayload);
        usedPrograms.add(programCode);
      }
      const abbreviationCode = nextAbbreviationCode().trim().slice(0, 1);
      const abbreviationPayload = { "Код_абревіатури": abbreviationCode, "Назва_абревіатури": abbreviation, "Код_опп": programCode, "Код_спеціальності": specialtyCode, "Код_факультету_нні": facultyCode };
      const { error } = await supabase.from("Абревіатура").insert(abbreviationPayload);
      if (error) throw new Error("Не вдалося зберегти абревіатуру «" + abbreviation + "»: " + error.message);
      abbreviations.push(abbreviationPayload);
      usedAbbreviations.add(abbreviationCode);
      added += 1;
    }
    return added;
  };

  const refreshAvailableYears = async (departmentCode: string) => {
    const department = trimText(departmentCode);
    if (!department) {
      setAvailableYears([]);
      setYear("");
      return [] as string[];
    }
    const { data, error: yearsError } = await supabase.from("Завантаження_навантаження")
      .select("Рік,Код_кафедри").eq("Код_кафедри", department);
    if (yearsError) throw new Error("Не вдалося оновити список навчальних років: " + yearsError.message);
    const years = [...new Set(((data ?? []) as unknown as Record<string, unknown>[])
      .filter((row) => trimText(row["Код_кафедри"]) === department)
      .map((row) => toDisplayYear(trimText(row["Рік"]))).filter(Boolean))]
      .sort((left, right) => toDbYear(right).localeCompare(toDbYear(left)));
    setAvailableYears(years);
    setYear((current) => years.includes(toDisplayYear(current)) ? toDisplayYear(current) : years[0] ?? "");
    return years;
  };

  const confirmStudyPlanImport = async () => {
    const selectedYear = toDbYear(importYear);
    const selectedDepartment = importDepartment.trim();
    const selectedForm = importForm.trim();
    const recordsForImport = planPreview;
    if (!planPreview.length || !selectedYear || !selectedDepartment || !selectedForm) {
      setImportMessage("Оберіть навчальний рік, кафедру та форму навчання.");
      return;
    }
    if (!/^\d{4}\/(?:\d{2}|\d{4})$/.test(trimText(importYear))) {
      setImportMessage("Вкажіть навчальний рік у форматі 2027/2028 або 2027/28.");
      return;
    }
    if (planPreview.some((row) => row.department.trim() && row.department.trim() !== selectedDepartment)) {
      setImportMessage("Код кафедри в Excel відрізняється від вибраної кафедри. Перевірте файл або параметр кафедри.");
      return;
    }
    if (newAbbreviationDrafts.some((draft) => !draft.specialty.trim() || !draft.program.trim())) {
      setImportMessage("Заповніть назву спеціальності та ОПП для кожної нової абревіатури.");
      return;
    }
    setImporting(true);
    setImportMessage("Перевіряю наявні записи для цього року та форми…");
    try {
      let overwriteCleanupWarning = "";
      const { data: existingData, error: existingError } = await supabase.from("Завантаження_навантаження")
        .select("id,Рік,Код_кафедри,Код_форми_навчання")
        .eq("Рік", selectedYear)
        .eq("Код_кафедри", selectedDepartment)
        .eq("Код_форми_навчання", selectedForm);
      if (existingError) throw new Error("Не вдалося перевірити наявні записи перед імпортом: " + existingError.message);
      const existingRows = ((existingData ?? []) as unknown as Record<string, unknown>[]).filter((row) =>
        toDbYear(row["Рік"]) === selectedYear
        && trimText(row["Код_кафедри"]) === selectedDepartment
        && trimText(row["Код_форми_навчання"]) === selectedForm);
      if (existingRows.length && !window.confirm("Для кафедри " + selectedDepartment + " уже є " + existingRows.length + " записів за " + toDisplayYear(selectedYear) + " («" + selectedForm + "»). Перезаписати тільки цей рік і форму навчання? Інші роки залишаться без змін.")) {
        setImportMessage("Імпорт скасовано. Дані в базі не змінено.");
        return;
      }
      const existingIds = existingRows.map((row) => number(row.id)).filter((id) => id > 0);
      if (existingRows.length) {
        if (existingIds.length !== existingRows.length) throw new Error("Не вдалося безпечно визначити всі id записів для перезапису. Дані не видалено.");
      }
      setImportMessage("Зберігаю довідники та рядки навчального плану…");
      await saveNewAbbreviationMemory(selectedDepartment);
      const dictionaryColumns = { code: "Код_дисципліни", department: "Код_кафедри", name: "Назва_дисципліни" };
      const { data: disciplineData, error: disciplineReadError } = await supabase.from("Дисципліни").select("Код_дисципліни,Код_кафедри,Назва_дисципліни");
      if (disciplineReadError) throw new Error("Не вдалося прочитати довідник дисциплін: " + disciplineReadError.message);
      const disciplineRows = (disciplineData ?? []) as unknown as Record<string, unknown>[];
      const usedCodes = new Set(disciplineRows.map((item) => trimText(item[dictionaryColumns.code])));
      const codeByName = new Map(disciplineRows.filter((item) => trimText(item[dictionaryColumns.department]) === selectedDepartment)
        .map((item) => [normalizedName(item[dictionaryColumns.name]), trimText(item[dictionaryColumns.code])]));
      const missingNames = [...new Map(recordsForImport.map((row) => row.discipline.trim()).filter(Boolean)
        .filter((name) => !codeByName.has(normalizedName(name))).map((name) => [normalizedName(name), name])).values()];
      const newDisciplinePayloads: Record<string, string>[] = [];
      let sequence = 1;
      for (const name of missingNames) {
        let code = "";
        while (sequence <= 36 ** 2) {
          const candidate = "P" + sequence.toString(36).toUpperCase().padStart(2, "0").slice(-2);
          sequence += 1;
          if (!usedCodes.has(candidate)) { code = candidate; break; }
        }
        if (!code) throw new Error("Не вдалося згенерувати код для нової дисципліни «" + name + "».");
        usedCodes.add(code);
        codeByName.set(normalizedName(name), code);
        newDisciplinePayloads.push({ [dictionaryColumns.code]: code, [dictionaryColumns.department]: selectedDepartment, [dictionaryColumns.name]: name });
      }
      if (newDisciplinePayloads.length) {
        const { error: insertDisciplineError } = await supabase.from("Дисципліни").insert(newDisciplinePayloads);
        if (insertDisciplineError) throw new Error("Не вдалося додати дисципліни: " + insertDisciplineError.message);
      }
      const levelCodes: Record<"bachelor" | "master", string> = {
        bachelor: trimText(importReferences.levels.find((item) => /бакалавр/i.test(item.name))?.code) || "Б",
        master: trimText(importReferences.levels.find((item) => /магістр|маг/i.test(item.name))?.code) || "М",
      };
      const payloads = recordsForImport.map((row) => {
        const disciplineCode = codeByName.get(normalizedName(row.discipline));
        if (!disciplineCode) throw new Error("Не знайдено код дисципліни «" + row.discipline + "».");
        return {
          "Рік": selectedYear, "Код_кафедри": selectedDepartment, "Код_форми_навчання": selectedForm,
          "Код_рівня_навчання": levelCodes[row.levelKind], "Код_дисципліни": disciplineCode, "Код_викладача": null,
          "Абревіатура": formatAbbreviation(row.abbreviations.join(", "), trimText(levelCodes[row.levelKind]) === "М" || row.levelKind === "master"), "Семестр": Math.round(row.semester),
          "Лекції": row.lecture, "Лабораторні_заняття": row.lab, "Практичні_заняття": row.practice,
          "Семінарні_заняття": row.seminar, "Екзамен": row.exam, "Залік": row.credit,
          "Курсовий_проєкт": row.courseProject, "Практика": row.practiceHours ?? 0,
          "Дипломне_проєктування": row.diplomaProject ?? 0, "Рецензії": row.reviews ?? 0,
          "Всього_годин": row.total,
        };
      });
      const { error: insertWorkloadError } = await supabase.from("Завантаження_навантаження").insert(payloads);
      if (insertWorkloadError) throw new Error(importErrorText(insertWorkloadError));
      if (existingIds.length) {
        const { error: deleteExistingError } = await supabase.from("Завантаження_навантаження").delete().in("id", existingIds);
        if (deleteExistingError) {
          overwriteCleanupWarning = " Попередні рядки цього самого року й форми не вдалося видалити: " + deleteExistingError.message;
        }
      }
      setImportMessage("Імпорт завершено: записано рядків навантаження — " + payloads.length + "; сума за колонкою «Всього» — " + hours(payloads.reduce((sum, row) => sum + number(row["Всього_годин"]), 0)) + " год; додано дисциплін — " + newDisciplinePayloads.length + "." + overwriteCleanupWarning);
      setCurrentDepartment(selectedDepartment);
      localStorage.setItem("workload-current-department", selectedDepartment);
      const displayYear = toDisplayYear(selectedYear);
      setDistributionFormCode(formModeForCode(selectedForm, importReferences.forms));
      await refreshAvailableYears(selectedDepartment);
      setYear(displayYear);
      setNewAbbreviationDrafts([]);
      setPlanPreview([]);
      setSelectedImportFile(null);
      setImportVersion((version) => version + 1);
    } catch (importError) {
      setImportMessage(importErrorText(importError));
    } finally { setImporting(false); }
  };

  const assignDistributionRow = async (row: PlanImportRow, teacherCode: string) => {
    if (!row.id) { setImportMessage("Цей рядок ще не має id у Supabase. Оновіть сторінку й спробуйте знову."); return; }
    const { error: updateError } = await supabase.from("Завантаження_навантаження").update({ "Код_викладача": teacherCode || null }).eq("id", row.id);
    if (updateError) { setImportMessage(importErrorText(updateError)); return; }
    setDistributionRows((current) => current.map((item) => item.id === row.id ? { ...item, teacherCode: teacherCode || null } : item));
    setDistributionAssignments((current) => { const next = { ...current }; if (teacherCode) next[distributionRowKey(row)] = teacherCode; else delete next[distributionRowKey(row)]; return next; });
    setWorkloads((current) => current.map((item) => item.id === row.id ? { ...item, teacher_code: teacherCode || null, teacher_id: teachers.find((teacher) => teacher.code === teacherCode)?.id ?? 0 } : item));
    if (teacherCode) setSuggestionRow({ sourceRow: row.sourceRow, teacherCode });
  };

  const clearDistributionAssignments = async () => {
    const assignedIds = distributionRows
      .filter((row) => row.id && trimText(row.teacherCode))
      .map((row) => Number(row.id));
    if (!assignedIds.length) {
      setImportMessage("У поточному плані та формі немає призначень для очищення.");
      return;
    }
    const formLabel = activeStudyForm === "Усі форми (Денна + Заочна)" ? "усіх форм" : activeStudyForm.toLocaleLowerCase("uk-UA") + " форми";
    if (!window.confirm(`Ви впевнені, що хочете скинути призначення викладачів для поточного плану (${toDisplayYear(year)}) та ${formLabel}? Самі дисципліни та години навчального плану залишаться без змін.`)) return;
    const { error } = await supabase.from("Завантаження_навантаження")
      .update({ "Код_викладача": null })
      .in("id", assignedIds);
    if (error) {
      setImportMessage(`Не вдалося очистити призначення: ${error.message}`);
      return;
    }
    setDistributionAssignments({});
    setDistributionRows((current) => current.map((row) => ({ ...row, teacherCode: null })));
    setWorkloads((current) => current.map((row) => assignedIds.includes(row.id) ? { ...row, teacher_code: null, teacher_id: 0 } : row));
    setSuggestionRow(null);
    setImportMessage(`Очищено призначення для ${assignedIds.length} рядків. Дисципліни та години залишилися без змін.`);
  };

  const removeTeacherWorkload = async (record: WorkloadRecord) => {
    if (!record.id) return;
    const { error } = await supabase.from("Завантаження_навантаження").update({ "Код_викладача": null }).eq("id", record.id);
    if (error) { setImportMessage("Не вдалося зняти дисципліну: " + error.message); return; }
    setImportMessage("Дисципліну знято з викладача; години повернуто у вакансії.");
    setImportVersion((current) => current + 1);
  };

  const splitDistributionRow = async (row: PlanImportRow, allocation: SplitAllocation, preserveOtherRows = false) => {
    if (!row.id || !row.disciplineCode) { setImportMessage("Не вдалося знайти id рядка або код дисципліни в Supabase."); return; }
    const kinds: Array<{ key: SplitWorkKind; column: string; amount: number }> = [
      { key: "lecture", column: "Лекції", amount: row.lecture }, { key: "lab", column: "Лабораторні_заняття", amount: row.lab },
      { key: "practice", column: "Практичні_заняття", amount: row.practice }, { key: "seminar", column: "Семінарні_заняття", amount: row.seminar },
      { key: "exam", column: "Екзамен", amount: row.exam }, { key: "credit", column: "Залік", amount: row.credit },
      { key: "courseProject", column: "Курсовий_проєкт", amount: row.courseProject }, { key: "practiceHours", column: "Практика", amount: row.practiceHours ?? 0 },
      { key: "diplomaProject", column: "Дипломне_проєктування", amount: row.diplomaProject ?? 0 }, { key: "reviews", column: "Рецензії", amount: row.reviews ?? 0 },
    ];
    const grouped = new Map<string, Record<string, number>>();
    for (const kind of kinds) {
      if (!kind.amount) continue;
      const parts = allocation[kind.key] ?? [{ teacherCode: row.teacherCode ?? "", hours: kind.amount }];
      const amountTotal = roundHours(parts.reduce((sum, part) => sum + Number(part.hours), 0));
      if (parts.some((part) => !Number.isFinite(Number(part.hours)) || Number(part.hours) < 0) || Math.abs(roundHours(amountTotal) - roundHours(kind.amount)) > 0.01) {
        setImportMessage(`Не вдалося зберегти поділ «${kind.column}»: сума годин має дорівнювати ${kind.amount}.`);
        return;
      }
      for (const part of parts) {
        const hoursForTeacher = roundHours(Number(part.hours));
        if (!hoursForTeacher) continue;
        const teacherCode = trimText(part.teacherCode);
        const teacherGroup = grouped.get(teacherCode) ?? {};
        teacherGroup[kind.column] = roundHours((teacherGroup[kind.column] ?? 0) + hoursForTeacher);
        grouped.set(teacherCode, teacherGroup);
      }
    }
    if (!grouped.size) { setImportMessage("У цьому рядку немає годин для поділу."); return; }
    const shared = { "Рік": toDbYear(year), "Код_кафедри": currentDepartment.trim(), "Код_форми_навчання": trimText(row.form_code),
      "Код_рівня_навчання": row.levelCode || (row.levelKind === "master" ? "М" : "Б"), "Код_дисципліни": row.disciplineCode,
      "Абревіатура": row.abbreviations.join(", "), "Семестр": row.semester };
    const emptyWork = Object.fromEntries(kinds.map((kind) => [kind.column, 0]));
    const payloads = [...grouped.entries()].map(([teacherCode, work]) => ({
      ...shared,
      "Код_викладача": teacherCode || null, ...emptyWork, ...work,
      "Всього_годин": roundHours(Object.values(work).reduce((sum, amount) => sum + amount, 0)),
    }));
    const [first, ...rest] = payloads;
    const { error: updateError } = await updateWorkloadRow(row.id, first);
    if (updateError) { setImportMessage(importErrorText(updateError)); return; }
    if (rest.length) {
      const { error: insertError } = await insertWorkloadRows(rest);
      if (insertError) { setImportMessage(importErrorText(insertError)); return; }
    }
    const previousSplitIds = preserveOtherRows ? [] : distributionRows.filter((item) => item.id && item.id !== row.id && trimText(item.form_code) === trimText(row.form_code) && item.disciplineCode === row.disciplineCode
      && item.target === row.target && item.semester === row.semester && (item.levelCode || item.levelKind) === (row.levelCode || row.levelKind)).map((item) => item.id as number);
    if (previousSplitIds.length) {
      const { error: removePreviousError } = await supabase.from("Завантаження_навантаження").delete().in("id", previousSplitIds);
      if (removePreviousError) { setImportMessage(importErrorText(removePreviousError)); return; }
    }
    setImportMessage("Поділ годин збережено в базі даних.");
    setImportVersion((value) => value + 1);
  };

  const copyAssignmentToDiscipline = async (selected: PlanImportRow) => {
    if (!selected.disciplineCode) { setImportMessage("Не знайдено код дисципліни для копіювання призначення."); return; }
    const levelKey = (row: PlanImportRow) => row.levelCode || row.levelKind;
    const groupKey = (row: PlanImportRow) => `${trimText(row.form_code)}\u0000${row.target}\u0000${row.semester}\u0000${levelKey(row)}`;
    const sameSourceGroup = (row: PlanImportRow) => row.disciplineCode === selected.disciplineCode && groupKey(row) === groupKey(selected);
    const sourceRows = distributionRows.filter(sameSourceGroup);
    const assignedSourceRows = sourceRows.filter((row) => trimText(row.teacherCode));
    if (!assignedSourceRows.length) {
      setImportMessage("Спочатку оберіть викладача для цього рядка, а потім натисніть „На всі“");
      return;
    }
    const workKinds: Array<{ column: string; amount: (row: PlanImportRow) => number }> = [
      { column: "Лекції", amount: (row) => row.lecture }, { column: "Лабораторні_заняття", amount: (row) => row.lab },
      { column: "Практичні_заняття", amount: (row) => row.practice }, { column: "Семінарні_заняття", amount: (row) => row.seminar },
      { column: "Екзамен", amount: (row) => row.exam }, { column: "Залік", amount: (row) => row.credit },
      { column: "Курсовий_проєкт", amount: (row) => row.courseProject }, { column: "Практика", amount: (row) => row.practiceHours ?? 0 },
      { column: "Дипломне_проєктування", amount: (row) => row.diplomaProject ?? 0 }, { column: "Рецензії", amount: (row) => row.reviews ?? 0 },
    ];
    const allTargets = distributionRows.filter((row) => row.disciplineCode === selected.disciplineCode && row.id);
    const targetGroups = [...new Set(allTargets.map(groupKey))].filter((key) => key !== groupKey(selected));
    let updatedCount = 0;
    for (const key of targetGroups) {
      const [targetForm, target, semesterText, level] = key.split("\u0000");
      const groupRows = allTargets.filter((row) => trimText(row.form_code) === targetForm && row.target === target && row.semester === number(semesterText) && levelKey(row) === level);
      const base = groupRows[0];
      if (!base?.id) continue;
      const teacherWork = new Map<string, Record<string, number>>();
      for (const kind of workKinds) {
        const amount = groupRows.reduce((sum, row) => sum + kind.amount(row), 0);
        if (!amount) continue;
        const weights = new Map<string, number>();
        for (const source of assignedSourceRows) weights.set(trimText(source.teacherCode), (weights.get(trimText(source.teacherCode)) ?? 0) + kind.amount(source));
        const rankedWeights = [...weights.entries()].filter(([, weight]) => weight > 0).sort((a, b) => b[1] - a[1]);
        const shares = rankedWeights.length ? rankedWeights : [[trimText(assignedSourceRows[0].teacherCode), 1] as [string, number]];
        const totalWeight = shares.reduce((sum, [, weight]) => sum + weight, 0);
        let allocated = 0;
        shares.forEach(([teacherCode, weight], index) => {
          const part = index === shares.length - 1 ? roundHours(amount - allocated) : roundHours(amount * weight / totalWeight);
          allocated = roundHours(allocated + part);
          if (!part) return;
          const payload = teacherWork.get(teacherCode) ?? {};
          payload[kind.column] = roundHours((payload[kind.column] ?? 0) + part);
          teacherWork.set(teacherCode, payload);
        });
      }
      if (!teacherWork.size) continue;
      const shared = { "Рік": toDbYear(year), "Код_кафедри": currentDepartment.trim(), "Код_форми_навчання": trimText(base.form_code),
        "Код_рівня_навчання": base.levelCode || (base.levelKind === "master" ? "М" : "Б"), "Код_дисципліни": base.disciplineCode,
        "Абревіатура": base.abbreviations.join(", "), "Семестр": base.semester };
      const empty = Object.fromEntries(workKinds.map((kind) => [kind.column, 0]));
      const payloads = [...teacherWork.entries()].map(([teacherCode, work]) => ({
        ...shared, "Код_викладача": teacherCode || null, ...empty, ...work,
        "Всього_годин": roundHours(Object.values(work).reduce((sum, amount) => sum + amount, 0)),
      }));
      const [first, ...rest] = payloads;
      const { error: updateError } = await supabase.from("Завантаження_навантаження").update(first).eq("id", base.id);
      if (updateError) { setImportMessage(importErrorText(updateError)); return; }
      if (rest.length) {
        const { error: insertError } = await supabase.from("Завантаження_навантаження").insert(rest);
        if (insertError) { setImportMessage(importErrorText(insertError)); return; }
      }
      const oldIds = groupRows.slice(1).map((row) => row.id).filter((id): id is number => Boolean(id));
      if (oldIds.length) {
        const { error: deleteError } = await supabase.from("Завантаження_навантаження").delete().in("id", oldIds);
        if (deleteError) { setImportMessage(importErrorText(deleteError)); return; }
      }
      updatedCount += groupRows.length;
    }
    setImportMessage(`Призначення застосовано до всіх рядків дисципліни (${updatedCount} шт.)`);
    setImportVersion((value) => value + 1);
  };

  const assignSameDiscipline = async (sourceRow?: PlanImportRow, requestedTeacherCode?: string) => {
    const selected = sourceRow ?? distributionRows.find((row) => row.sourceRow === suggestionRow?.sourceRow);
    const teacherCode = requestedTeacherCode ?? suggestionRow?.teacherCode;
    if (!selected?.disciplineCode) return;
    if (!teacherCode) { setImportMessage("Спочатку оберіть викладача для цієї дисципліни."); return; }
    const targets = distributionRows.filter((row) => row.disciplineCode === selected.disciplineCode && row.id);
    for (const row of targets) {
      const { error: updateError } = await supabase.from("Завантаження_навантаження").update({ "Код_викладача": teacherCode }).eq("id", row.id as number);
      if (updateError) { setImportMessage(importErrorText(updateError)); return; }
    }
    setImportVersion((value) => value + 1);
  };

  const autoAssignRemaining = async () => {
    const eligibleTeachers = teachers.filter((teacher) => teacher.departmentCode.trim() === currentDepartment.trim() && teacher.normHours > 0 && teacher.rate > 0);
    if (!eligibleTeachers.length) { setImportMessage("Не знайдено викладачів кафедри з доступною нормою годин."); return; }
    const used = Object.fromEntries(eligibleTeachers.map((teacher) => [teacher.code, workloads.filter((row) => row.teacher_code === teacher.code).reduce((sum, row) => sum + workloadTotal(row), 0)])) as Record<string, number>;
    const remaining = distributionRows.filter((row) => !row.teacherCode);
    const toAssign = new Map<string, number[]>();
    for (const row of remaining) {
      const available = eligibleTeachers.filter((teacher) => used[teacher.code] + row.total <= teacher.normHours * teacher.rate);
      const chosen = available.sort((a, b) => (used[a.code] / (a.normHours * a.rate)) - (used[b.code] / (b.normHours * b.rate)))[0];
      if (!chosen) continue;
      if (row.id) toAssign.set(chosen.code, [...(toAssign.get(chosen.code) ?? []), row.id]);
      used[chosen.code] += row.total;
    }
    for (const [teacherCode, ids] of toAssign) {
      const { error: updateError } = await supabase.from("Завантаження_навантаження").update({ "Код_викладача": teacherCode }).in("id", ids);
      if (updateError) { setImportMessage(importErrorText(updateError)); return; }
    }
    setSuggestionRow(null);
    const left = remaining.length - [...toAssign.values()].reduce((sum, ids) => sum + ids.length, 0);
    setImportMessage(left ? "Автоматично призначено години в межах ставок. Без перевищення норми залишилось вакантних рядків: " + left + "." : "Усі вакантні рядки розподілено без перевищення норми годин.");
    setImportVersion((value) => value + 1);
  };

  const clearImportedData = async () => {
    const keyYear = toDbYear(year);
    const keyDepartment = trimText(currentDepartment);
    const keyForms = selectedFormCodes.map(trimText).filter(Boolean);
    if (!keyYear || !keyDepartment || !keyForms.length) { setImportMessage("Оберіть навчальний рік і форму навчання."); return; }
    const selectedFormsLabel = showAllForms ? "Денна та Заочна" : (distributionFormCode === "extramural" ? "Заочна" : "Денна");
    if (!window.confirm("Видалити імпортоване навантаження за " + toDisplayYear(keyYear) + " для кафедри " + keyDepartment + " («" + selectedFormsLabel + "»)?")) return;
    setImporting(true);
    setImportMessage("Шукаю рядки імпорту…");
    try {
      const { data, error: selectError } = await supabase.from("Завантаження_навантаження").select("id,Рік,Код_кафедри,Код_форми_навчання,Код_дисципліни")
        .eq("Рік", keyYear).eq("Код_кафедри", keyDepartment).in("Код_форми_навчання", keyForms);
      if (selectError) throw new Error(importErrorText(selectError));
      const matches = ((data ?? []) as unknown as Record<string, unknown>[]).filter((row) =>
        toDbYear(row["Рік"]) === keyYear && trimText(row["Код_кафедри"]) === keyDepartment && keyForms.includes(trimText(row["Код_форми_навчання"])));
      const ids = matches.map((row) => number(row.id)).filter(Boolean);
      if (ids.length) {
        const { error: deleteError } = await supabase.from("Завантаження_навантаження").delete().in("id", ids);
        if (deleteError) throw new Error(importErrorText(deleteError));
      }
      const { data: departmentDisciplines, error: departmentDisciplineError } = await supabase.from("Дисципліни").select("*").eq("Код_кафедри", keyDepartment);
      if (departmentDisciplineError) throw new Error("Навантаження очищено, але не вдалося перевірити дисципліни: " + departmentDisciplineError.message);
      const departmentDisciplineRows = (departmentDisciplines ?? []) as unknown as Record<string, unknown>[];
      const removeCodes = [...new Set([...departmentDisciplineRows.map((row) => trimText(row["Код_дисципліни"])), ...matches.map((row) => trimText(row["Код_дисципліни"]))].filter(Boolean))];
      const removeDisciplines = removeCodes.length > 0 && window.confirm("Також видалити дисципліни цієї кафедри, на які більше немає посилань у навантаженні? (" + removeCodes.length + " перевірити)");
      setDistributionRows([]);
      setDistributionAssignments({});
      setWorkloads((current) => current.filter((row) => !(toDbYear(row.academic_year) === keyYear && trimText(row.department_code) === keyDepartment && keyForms.includes(trimText(row.form_code)))));
      let disciplineMessage = "";
      if (removeDisciplines) {
        let removed = 0;
        for (const code of removeCodes) {
          const { data: refs, error: referenceError } = await supabase.from("Завантаження_навантаження").select("id").eq("Код_дисципліни", code).limit(1);
          if (referenceError) continue;
          if (!refs?.length) {
            const { error: deleteError } = await supabase.from("Дисципліни").delete().eq("Код_дисципліни", code).eq("Код_кафедри", keyDepartment);
            if (!deleteError) removed += 1;
          }
        }
        disciplineMessage = " Видалено дисциплін: " + removed + " із " + removeCodes.length + " (записи з посиланнями можуть залишитися).";
      }
      setImportMessage(matches.length ? "Навантаження очищено за " + toDisplayYear(keyYear) + " для кафедри " + keyDepartment + " і форми «" + selectedFormsLabel + "»." + disciplineMessage : "У БД не знайдено навантаження за " + toDisplayYear(keyYear) + " для кафедри " + keyDepartment + " і форми «" + selectedFormsLabel + "»." + disciplineMessage);
      await refreshAvailableYears(keyDepartment);
      setImportVersion((version) => version + 1);
    } catch (clearError) {
      setImportMessage(importErrorText(clearError));
    } finally {
      setImporting(false);
    }
  };

  const importWorkbookRows = async (rows: Record<string, unknown>[]): Promise<ImportResult> => {
    const read = (row: Record<string, unknown>, ...names: string[]) => {
      const key = Object.keys(row).find((column) => names.some((name) => normalizeColumn(column) === normalizeColumn(name)));
      return key ? row[key] : "";
    };
    const text = (value: unknown) => String(value ?? "").trim();
    const teachersToUpsert = rows.map((row) => ({
      "Код_викладача": text(read(row, "Код_викладача", "код викладача", "teacher_id")),
      "ПІБ": text(read(row, "ПІБ", "Викладач", "teacher_name")),
      "Вчене_звання": text(read(row, "Вчене_звання", "Вчене звання")),
      "Науковий_ступінь": text(read(row, "Науковий_ступінь", "Науковий ступінь")),
      "Телефон_викладача": text(read(row, "Телефон_викладача", "Телефон")),
      "Email_викладача": text(read(row, "Email_викладача", "Email", "E-mail")),
    })).filter((teacher) => teacher["Код_викладача"] && teacher["ПІБ"]);
    const disciplinesToUpsert = rows.map((row) => ({
      "Код_дисципліни": text(read(row, "Код_дисципліни", "код дисципліни", "discipline_id")),
      "Код_кафедри": text(read(row, "Код_кафедри", "код кафедри", "department_id")),
      "Назва_дисципліни": text(read(row, "Назва_дисципліни", "Дисципліна", "Назва дисципліни")),
    })).filter((discipline) => discipline["Код_дисципліни"] && discipline["Код_кафедри"] && discipline["Назва_дисципліни"]);
    const workloadColumns = ["Рік", "Код_кафедри", "Код_форми_навчання", "Код_рівня_навчання", "Лекції", "Лабораторні_заняття", "Практичні_заняття", "Семінарські_заняття", "Екзамен", "Залік", "Курсовий_проект", "Практика", "Дипломне_проєктування", "Рецензії", "Всього_годин"];
    const workloadHourColumns = new Set(["Лекції", "Лабораторні_заняття", "Практичні_заняття", "Семінарські_заняття", "Екзамен", "Залік", "Курсовий_проект", "Практика", "Дипломне_проєктування", "Рецензії", "Всього_годин"]);
    const workloadsToInsert: Record<string, unknown>[] = rows.map((row) => {
      const workload = Object.fromEntries(workloadColumns.map((column) => {
        const value = column === "Рецензії" ? read(row, "Рецензії", "Рецензування") : read(row, column);
        if (column === "Рік") return [column, toDbYear(value)];
        if (["Код_кафедри", "Код_форми_навчання", "Код_рівня_навчання"].includes(column)) return [column, text(value)];
        if (workloadHourColumns.has(column) && text(value)) {
          const parsed = Number(text(value).replace(/\s/g, "").replace(",", "."));
          return [column, Number.isFinite(parsed) ? Math.ceil(parsed) : value];
        }
        return [column, value];
      })) as Record<string, unknown>;
      return Object.fromEntries(Object.entries(workload).filter(([, value]) => text(value) !== "")) as Record<string, unknown>;
    }).filter((workload) => text(workload["Рік"]) && text(workload["Код_кафедри"]));

    if (!teachersToUpsert.length && !disciplinesToUpsert.length && !workloadsToInsert.length) {
      throw new Error("Не знайдено потрібних колонок. Додайте у файл, наприклад: «Код_викладача», «ПІБ», «Код_дисципліни», «Назва_дисципліни», «Рік».");
    }
    if (teachersToUpsert.length) {
      const { error: teacherError } = await supabase.from("Викладач").upsert(teachersToUpsert, { onConflict: "Код_викладача" });
      if (teacherError) throw new Error(`Викладачі: ${teacherError.message}`);
    }
    if (disciplinesToUpsert.length) {
      const { error: disciplineError } = await supabase.from("Дисципліни").upsert(disciplinesToUpsert, { onConflict: "Код_дисципліни" });
      if (disciplineError) throw new Error(`Дисципліни: ${disciplineError.message}`);
    }
    if (workloadsToInsert.length) {
      const { error: workloadError } = await supabase.from("Завантаження_навантаження").insert(workloadsToInsert);
      if (workloadError) throw new Error(`Навантаження: ${workloadError.message}`);
    }
    return { teachers: teachersToUpsert.length, disciplines: disciplinesToUpsert.length, workloads: workloadsToInsert.length };
  };

  return (
    <div className={`dashboard-container ${isFullscreen ? "workload-fullscreen-shell" : ""}`}>
      {!isFullscreen && <AppSidebar active="workload" />}
      <aside className="sidebar legacy-sidebar" style={{ display: isFullscreen ? "none" : undefined }}>
        <div className="sidebar-brand"><img src={logo} alt="НУБіП" className="sidebar-logo" /><div><h3>НУБіП</h3><p>Система управління</p></div></div>
        <nav className="sidebar-menu">
          <button type="button" className="menu-item" onClick={() => navigate("/dashboard")}><FiGrid /> Головна</button>
          <span className="menu-category">Управління</span>
          <button type="button" className="menu-item active"><FiBookOpen /> Навантаження</button>
          <button type="button" className="menu-item" onClick={() => navigate("/teachers")}><FiUsers /> Викладачі</button>
          <button type="button" className="menu-item" onClick={() => navigate("/department")}><FiGrid /> Кафедра</button>
          <span className="menu-category">Додатково</span>
          <button type="button" className="menu-item" onClick={() => navigate("/help")}><FiHelpCircle /> Довідка</button>
        </nav>
        <div className="sidebar-profile"><div className="profile-info"><div className="profile-avatar-placeholder"><FiUsers /></div><div><h4>{profileName}</h4><p>{role}</p></div><FiChevronDown className="profile-arrow" /></div><button className="btn-logout" onClick={() => { localStorage.removeItem("userRole"); localStorage.removeItem("userDisplayName"); localStorage.removeItem("userEmail"); void supabase.auth.signOut(); navigate("/"); }}><FiLogOut /> Вийти</button></div>
      </aside>
      <main className="main-content" style={{ flex: "1 1 0%", minWidth: 0, maxWidth: "100%", overflowX: isFullscreen ? "hidden" : "clip", boxSizing: "border-box", ...(isFullscreen ? { position: "relative", zIndex: 1, display: "flex", flexDirection: "column", width: "100%", height: "100%", minHeight: 0, margin: 0, padding: 0, overflow: "hidden", background: "transparent" } : { width: "auto" }) }}>
        <input ref={importInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={selectImportFile} />
        {!isFullscreen && <header className="mb-7 flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="header-left"><span className="welcome-text">Вітаємо, {profileName}</span><h2>{departmentName || "Кафедра"}</h2><p className="role-sub">{facultyName}{role ? ` · ${role}` : ""}</p></div>
          </div>
          <div className="header-right" style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}><button type="button" className="btn-import-workload" disabled={importing} onClick={() => importInputRef.current?.click()}>{importing ? <FiLoader aria-hidden="true" className="workload-import-spinner" /> : <FiClipboard />} {importing ? (selectedImportFile ? "Імпорт…" : "Обробка…") : "Імпорт Excel"}</button><button type="button" className="btn-import-workload" disabled={importing || !year.trim() || !currentDepartment.trim()} onClick={() => void clearImportedData()}>Очистити дані</button><label className="academic-year"><select aria-label="Форма навчання" className="date-badge" value={distributionFormCode} onChange={(event) => setDistributionFormCode(event.target.value as FormSelection)}><option value="all">Усі форми (Денна + Заочна)</option><option value="day">Денна</option><option value="extramural">Заочна</option></select></label><label className="academic-year">
            <select className="date-badge" value={year} onChange={(event) => setYear(toDisplayYear(event.target.value))}>
              {!year && <option value="">Немає навчальних років</option>}
              {availableYears.map((item) => <option key={item} value={item}>{item} навч. рік</option>)}
            </select>
          </label></div>
        </header>}

        {planPreview.length > 0 && <div role="presentation" style={{ position: "fixed", inset: 0, zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, background: "rgba(15, 23, 42, 0.45)" }}>
          <section role="dialog" aria-modal="true" aria-labelledby="import-preview-title" style={{ maxWidth: 860, width: "100%", maxHeight: "calc(100vh - 48px)", overflowY: "auto", boxSizing: "border-box", background: "#fff", borderRadius: 16, padding: 28, boxShadow: "0 20px 40px rgba(0,0,0,0.15)", color: "#0f172a", fontFamily: "Arial, sans-serif" }}>
            <header style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16, marginBottom: 20, paddingBottom: 16, borderBottom: "1px solid #e2e8f0" }}>
              <div style={{ minWidth: 0 }}><h2 id="import-preview-title" style={{ margin: 0, fontSize: 20, lineHeight: 1.35, fontWeight: 700 }}>Підтвердження імпорту</h2><p style={{ margin: "6px 0 0", color: "#64748b", fontSize: 13, overflowWrap: "anywhere" }}>{selectedImportFile}</p></div>
              <button type="button" aria-label="Закрити попередній перегляд" disabled={importing} style={{ flex: "0 0 auto", width: 36, height: 36, border: "1px solid #e2e8f0", borderRadius: 8, background: "#fff", color: "#64748b", cursor: "pointer", fontSize: 18 }} onClick={() => { setPlanPreview([]); setSelectedImportFile(null); setImportMessage(null); }}>✕</button>
            </header>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 16, marginBottom: 20 }}>
              {[
                <label key="year" style={{ display: "block", minWidth: 0, color: "#334155", fontSize: 13, fontWeight: 600 }}>Навчальний рік<select style={{ display: "block", boxSizing: "border-box", width: "100%", height: 40, marginTop: 6, padding: "0 12px", borderRadius: 8, border: "1px solid #cbd5e1", background: "#fff", color: "#0f172a", fontSize: 14 }} value={availableYears.includes(toDisplayYear(importYear)) ? toDisplayYear(importYear) : "__new_year__"} onChange={(event) => { const value = event.target.value; setImportYear(value === "__new_year__" ? (availableYears.includes(toDisplayYear(importYear)) ? "" : importYear) : toDisplayYear(value)); }}>{availableYears.map((option) => <option key={option} value={option}>{option}</option>)}<option value="__new_year__">+ Новий навчальний рік…</option></select>{!availableYears.includes(toDisplayYear(importYear)) && <input aria-label="Новий навчальний рік" value={importYear} onChange={(event) => setImportYear(event.target.value)} placeholder="2027/2028" style={{ display: "block", boxSizing: "border-box", width: "100%", height: 38, marginTop: 8, padding: "0 10px", borderRadius: 8, border: "1px solid #cbd5e1", fontSize: 14 }} />}</label>,
                <label key="department" style={{ display: "block", minWidth: 0, color: "#334155", fontSize: 13, fontWeight: 600 }}>Кафедра<select style={{ display: "block", boxSizing: "border-box", width: "100%", height: 40, marginTop: 6, padding: "0 12px", borderRadius: 8, border: "1px solid #cbd5e1", background: "#fff", color: "#0f172a", fontSize: 14 }} value={importDepartment} onChange={(event) => setImportDepartment(event.target.value.trim())}><option value="">Оберіть кафедру</option>{importReferences.departments.map((option) => <option key={option.code.trim()} value={option.code.trim()}>{option.name} ({option.code.trim()})</option>)}</select></label>,
                <label key="form" style={{ display: "block", minWidth: 0, color: "#334155", fontSize: 13, fontWeight: 600 }}>Форма навчання<select style={{ display: "block", boxSizing: "border-box", width: "100%", height: 40, marginTop: 6, padding: "0 12px", borderRadius: 8, border: "1px solid #cbd5e1", background: "#fff", color: "#0f172a", fontSize: 14 }} value={importForm} onChange={(event) => setImportForm(event.target.value.trim())}><option value="">Оберіть форму</option>{importReferences.forms.map((option) => <option key={option.code.trim()} value={option.code.trim()}>{option.name} ({option.code.trim()})</option>)}</select></label>,
              ]}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 16, marginBottom: 20 }}>
              {[["Всього рядків", planPreview.length], ["Унікальних дисциплін", new Set(planPreview.map((row) => normalizedName(row.discipline))).size], ["Розпізнано ОПП", recognizedProgramCount]].map(([label, value]) => <div key={String(label)} style={{ boxSizing: "border-box", minWidth: 0, padding: "14px 16px", background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 10 }}><p style={{ margin: 0, color: "#64748b", fontSize: 12 }}>{label}</p><p style={{ margin: "6px 0 0", color: label === "Розпізнано ОПП" ? "#047857" : "#0f172a", fontSize: 20, fontWeight: 700 }}>{value}</p></div>)}
            </div>
            {recognizedAbbreviations.some((item) => item.record) && <details style={{ marginBottom: 20, padding: "10px 12px", border: "1px solid #e2e8f0", borderRadius: 8 }}><summary style={{ cursor: "pointer", color: "#334155", fontSize: 13, fontWeight: 600 }}>Розпізнані абревіатури ОПП</summary><div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>{recognizedAbbreviations.filter((item) => item.record).map(({ name, record }) => <span key={name} title={"ОПП: " + (record?.programName || record?.programCode || "—") + "; спеціальність: " + (record?.specialtyName || record?.specialtyCode || "—")} style={{ padding: "4px 10px", borderRadius: 999, background: "#dcfce7", color: "#166534", fontSize: 12 }}>{name}</span>)}</div></details>}
            {newAbbreviationDrafts.length > 0 && <div style={{ marginBottom: 20, padding: 14, border: "1px solid #e2e8f0", borderRadius: 10, background: "#fff" }}><p style={{ margin: "0 0 10px", color: "#334155", fontSize: 13, fontWeight: 700 }}>Нові абревіатури — уточніть назви перед збереженням</p><div style={{ maxWidth: "100%", overflowX: "auto" }}><table style={{ width: "100%", minWidth: 650, borderCollapse: "collapse", textAlign: "left", fontSize: 12 }}><thead style={{ background: "#f8fafc", color: "#475569" }}><tr>{["Абревіатура", "Назва спеціальності", "Назва ОПП"].map((label) => <th key={label} style={{ padding: "9px 10px", borderBottom: "1px solid #e2e8f0" }}>{label}</th>)}</tr></thead><tbody>{newAbbreviationDrafts.map((draft) => <tr key={draft.abbreviation} style={{ borderBottom: "1px solid #f1f5f9" }}><th style={{ padding: "9px 10px", whiteSpace: "nowrap" }}>{draft.abbreviation}<button type="button" title="Прибрати зі списку" aria-label={"Прибрати " + draft.abbreviation} onClick={() => setNewAbbreviationDrafts((current) => current.filter((item) => item.abbreviation !== draft.abbreviation))} style={{ marginLeft: 8, width: 22, height: 22, border: "1px solid #fecaca", borderRadius: 6, background: "#fff", color: "#b91c1c", cursor: "pointer", lineHeight: 1 }}>✕</button></th>{(["specialty", "program"] as const).map((field) => <td key={field} style={{ padding: "7px 10px" }}><input value={draft[field]} onChange={(event) => setNewAbbreviationDrafts((current) => current.map((item) => item.abbreviation === draft.abbreviation ? { ...item, [field]: event.target.value } : item))} placeholder={field === "specialty" ? "Введіть назву спеціальності" : "Введіть назву ОПП"} style={{ boxSizing: "border-box", width: "100%", minWidth: 220, height: 36, padding: "0 10px", border: "1px solid #cbd5e1", borderRadius: 7, fontSize: 12 }} /></td>)}</tr>)}</tbody></table></div></div>}
            <div style={{ maxWidth: "100%", overflowX: "auto", marginBottom: 20, border: "1px solid #e2e8f0", borderRadius: 8 }}>
              <table style={{ width: "100%", minWidth: 980, borderCollapse: "collapse", textAlign: "left", fontSize: 12 }}><thead style={{ background: "#f8fafc", color: "#475569" }}><tr>{["Рівень", "Лекції", "Лабораторні", "Практичні", "Семінарські", "Екзамен", "Залік", "Курсовий проєкт", "Практика", "Дипломне проєктування", "Рецензії", "Всього викладача"].map((label) => <th key={label} style={{ padding: "10px 12px", borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap" }}>{label}</th>)}</tr></thead>
                <tbody>{importTotalsByLevel.map((group) => <tr key={group.levelKind} style={{ borderBottom: "1px solid #e2e8f0" }}><th style={{ padding: "10px 12px", textAlign: "left", whiteSpace: "nowrap" }}>{group.levelKind === "master" ? "Магістр" : "Бакалавр"}</th><td style={{ padding: "10px 12px" }}>{hours(group.totals.lectures)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.laboratories)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.practices)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.seminars)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.exams)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.credits)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.courseProjects)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.practiceHours)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.diplomaProject)}</td><td style={{ padding: "10px 12px" }}>{hours(group.totals.reviews)}</td><td style={{ padding: "10px 12px", fontWeight: 700 }}>{hours(group.totals.total)}</td></tr>)}
                <tr style={{ background: "#f1f5f9", fontWeight: 700 }}><th style={{ padding: "10px 12px", textAlign: "left" }}>Разом</th>{(["lectures", "laboratories", "practices", "seminars", "exams", "credits", "courseProjects", "practiceHours", "diplomaProject", "reviews", "total"] as const).map((field) => <td key={field} style={{ padding: "10px 12px" }}>{field === "total" ? hours(importWorkbookTotal) : hours(importTotalsByLevel.reduce((sum, group) => sum + group.totals[field], 0))}</td>)}</tr></tbody>
              </table>
            </div>
            <p style={{ margin: "-8px 0 20px", color: "#475569", fontSize: 13, lineHeight: 1.5 }}>
              Підсумок за колонкою «Всього, годин» Excel: І семестр (непарні 1, 3, 5, 7) — <strong>{hours(importSemesterTotals[0].total)}</strong> год · ІІ семестр (парні 2, 4, 6, 8) — <strong>{hours(importSemesterTotals[1].total)}</strong> год · Разом — <strong>{hours(importWorkbookTotal)}</strong> год
            </p>
            <details style={{ marginBottom: 20, border: "1px solid #e2e8f0", borderRadius: 8 }}>
              <summary style={{ cursor: "pointer", padding: "10px 12px", color: "#334155", fontSize: 13, fontWeight: 600 }}>Перевірити всі розпізнані дисципліни ({new Set(planPreview.map((row) => normalizedName(row.discipline))).size})</summary>
              <div style={{ maxHeight: 240, overflowY: "auto", borderTop: "1px solid #e2e8f0" }}><table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left", fontSize: 12 }}><thead style={{ background: "#f8fafc" }}><tr>{["Дисципліна", "Потік / абревіатури", "Рівень", "Всього годин"].map((label) => <th key={label} style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>{label}</th>)}</tr></thead><tbody>{[...new Map(planPreview.map((row) => [normalizedName(row.discipline), row])).values()].map((row) => <tr key={row.sourceRow} style={{ borderTop: "1px solid #f1f5f9" }}><td style={{ padding: "10px 12px" }}>{row.discipline}</td><td style={{ padding: "10px 12px" }}>{row.target || "—"}</td><td style={{ padding: "10px 12px", whiteSpace: "nowrap" }}>{row.levelKind === "master" ? "Магістр" : "Бакалавр"}</td><td style={{ padding: "10px 12px" }}>{hours(row.total)}</td></tr>)}</tbody></table></div>
            </details>
            {importMessage && <p style={{ margin: "0 0 16px", color: "#475569", fontSize: 13, lineHeight: 1.5 }} role="status">{importMessage}</p>}
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 12, marginTop: 24, paddingTop: 16, borderTop: "1px solid #e2e8f0" }}><button type="button" disabled={importing} style={{ padding: "10px 20px", border: "1px solid #cbd5e1", borderRadius: 8, background: "#fff", color: "#334155", fontSize: 14, fontWeight: 600, cursor: "pointer" }} onClick={() => { setPlanPreview([]); setSelectedImportFile(null); setImportMessage(null); }}>Скасувати</button><button type="button" disabled={importing || !importYear.trim() || !importDepartment.trim() || !importForm.trim()} style={{ padding: "10px 20px", border: 0, borderRadius: 8, background: "#2563eb", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", opacity: importing ? 0.65 : 1 }} onClick={() => void confirmStudyPlanImport()}>{importing ? "Імпорт пачки…" : "Підтвердити та імпортувати пачкою"}</button></div>
          </section>
        </div>}

        {!isFullscreen && <nav className="window-tabs-navigation" aria-label="Розділи навантаження">
          {TAB_ITEMS.map(({ key, label, icon }) => (
            <button key={key} onClick={() => changeTab(key)} className={`tab-navigation-btn ${activeTab === key ? "active-tab" : ""}`}><span>{icon}</span>{label}</button>
          ))}
        </nav>}


        {activeTab === null && <div className="rounded-xl border border-slate-200 bg-white px-6 py-10 text-center text-sm text-slate-500 shadow-sm">Оберіть потрібну вкладку вище для перегляду або розподілу навантаження</div>}
        {loading && activeTab !== null && <div className="rounded-xl bg-white p-10 text-center text-sm text-slate-500 shadow-sm">Завантаження даних…</div>}
        {activeTab === "adjust" && <DistributionTab rows={distributionRows} teachers={teachers.filter((teacher) => teacher.departmentCode.trim() === currentDepartment.trim())} workloads={workloads} assignments={distributionAssignments} year={year} studyForm={activeStudyForm} organization={{ facultyName, departmentName, departmentHeadName }} onAssign={assignDistributionRow} onSplit={splitDistributionRow} onCopyAssignment={copyAssignmentToDiscipline} onClearAssignments={clearDistributionAssignments} suggestion={suggestionRow} onAssignSameDiscipline={assignSameDiscipline} onAutoAssign={autoAssignRemaining} fullscreen={isFullscreen} setFullscreen={setIsFullscreen} />}
        {!loading && activeTab === "report" && <ReportTab teachers={teachers.filter((teacher) => teacher.departmentCode.trim() === currentDepartment.trim())} teacherId={teacherId} setTeacherId={setTeacherId} year={year} studyForm={activeStudyForm} records={teacherLoads} allRecords={workloads} allPlanRows={distributionRows} selectedTeacher={selectedTeacher} facultyName={facultyName} departmentName={departmentName} departmentHeadName={departmentHeadName} vacantRows={distributionRows.filter((row) => !row.teacherCode && row.total > 0)} onRemoveRow={removeTeacherWorkload} onSplitVacancy={(row, allocation) => splitDistributionRow(row, allocation, true)} onSplitRecord={(row, allocation) => splitDistributionRow(row, allocation)} onAssignSameDiscipline={assignSameDiscipline} fullscreen={isFullscreen} setFullscreen={setIsFullscreen} />}
        {!loading && activeTab === "form56" && <Form56Tab teachers={teachers.filter((teacher) => teacher.departmentCode.trim() === currentDepartment.trim())} records={workloads} year={year} selectedForm={distributionFormCode as FormSelection} departmentName={departmentName} facultyName={facultyName} departmentHeadName={departmentHeadName} fullscreen={isFullscreen} setFullscreen={setIsFullscreen} />}
        {!loading && activeTab === "vacancies" && <VacanciesTab year={year} studyForm={activeStudyForm} rows={distributionRows.filter((row) => !row.teacherCode && row.total > 0)} loading={false} assigningId={assigningId} onAssign={(row) => void assignVacancy({ id: row.id ?? row.sourceRow, discipline_name: row.discipline, group_code: row.target || row.abbreviations.join("+"), work_type: "Навчальний план", hours: row.total, is_assigned: false })} organization={{ facultyName, departmentName, departmentHeadName }} fullscreen={isFullscreen} setFullscreen={setIsFullscreen} />}
        {!loading && activeTab === "disciplines" && <DisciplinesTab disciplines={disciplines} teachers={teachers} planRows={distributionRows} canEdit={canEditDisciplines} departmentCode={currentDepartment} year={year} formSelection={distributionFormCode} forms={importReferences.forms} levels={importReferences.levels} abbreviations={importReferences.abbreviations} organization={{ facultyName, departmentName, departmentHeadName }} onChanged={() => setImportVersion((version) => version + 1)} fullscreen={isFullscreen} setFullscreen={setIsFullscreen} />}
      </main>
    </div>
  );
}


function Card({ children, style, className = "" }: { children: React.ReactNode; style?: React.CSSProperties; className?: string }) { return <section className={`custom-card-window rounded-[10px] border border-slate-200 bg-white p-5 sm:p-6 ${className}`} style={style}>{children}</section>; }
function EmptyRow({ columns, message }: { columns: number; message: string }) { return <tr><td colSpan={columns} className="px-4 py-10 text-center text-sm text-slate-500">{message}</td></tr>; }

function FullscreenTabBar({ title, setFullscreen, action }: { title: string; setFullscreen: (value: boolean | ((current: boolean) => boolean)) => void; action?: React.ReactNode }) {
  return <div className="workload-fullscreen-tab-bar"><strong>{title}</strong><div className="workload-fullscreen-tab-actions">{action}<button type="button" onClick={() => setFullscreen(false)}>Згорнути вікно</button></div></div>;
}

function DownloadReportButton({ onClick }: { onClick: () => void }) {
  return <button type="button" className="workload-download-report-button" onClick={onClick}><FiDownload aria-hidden="true" />Завантажити звіт</button>;
}

function InlineHoursSummary({ total, assigned, vacant }: { total?: number; assigned?: number; vacant?: number }) {
  return <span className="workload-inline-hours-summary">{total !== undefined && <> · Загальний обсяг: <strong><span style={{ whiteSpace: "nowrap" }}>{hours(total)} год</span></strong></>}{assigned !== undefined && <> · Розподілено: <strong><span style={{ whiteSpace: "nowrap" }}>{hours(assigned)} год</span></strong></>}{vacant !== undefined && <> · Вакантні: <strong><span style={{ whiteSpace: "nowrap" }}>{hours(vacant)} год</span></strong></>}</span>;
}

function FullscreenToggleButton({ onClick }: { onClick: () => void }) {
  return <button type="button" className="workload-inline-fullscreen-button" onClick={onClick} title="Розгорнути вкладку на весь екран">⛶ На весь екран</button>;
}

function DistributionTab(props: { rows: PlanImportRow[]; teachers: Teacher[]; workloads: WorkloadRecord[]; assignments: Record<string, string>; year: string; studyForm: StudyForm; organization: EditableReport["organization"]; onAssign: (row: PlanImportRow, teacherCode: string) => void; onSplit: (row: PlanImportRow, allocation: SplitAllocation) => void; onCopyAssignment: (row: PlanImportRow) => void; onClearAssignments: () => Promise<void>; suggestion: { sourceRow: number; teacherCode: string } | null; onAssignSameDiscipline: (row?: PlanImportRow, teacherCode?: string) => void; onAutoAssign: () => void; fullscreen: boolean; setFullscreen: (value: boolean | ((current: boolean) => boolean)) => void }) {
  const [query, setQuery] = useState("");
  const [abbreviation, setAbbreviation] = useState("");
  const [semester, setSemester] = useState("");
  const [assignmentStatus, setAssignmentStatus] = useState("all");
  const [selectedRow, setSelectedRow] = useState<PlanImportRow | null>(null);
  const [reportPreviewOpen, setReportPreviewOpen] = useState(false);
  const [allocation, setAllocation] = useState<SplitAllocation>({});
  const [quickTeacher, setQuickTeacher] = useState("");
  const fullscreen = props.fullscreen;
  const tableScrollRef = useRef<HTMLDivElement>(null);
  const scrollPositionRef = useRef({ top: 0, left: 0 });
  const kinds: Array<{ key: SplitWorkKind; label: string; short: string; amount: (row: PlanImportRow) => number }> = [
    { key: "lecture", label: "Лекції", short: "Лекції", amount: (row) => row.lecture },
    { key: "lab", label: "Лабораторні", short: "Лаб", amount: (row) => row.lab },
    { key: "practice", label: "Практичні", short: "Прак", amount: (row) => row.practice },
    { key: "seminar", label: "Семінарські", short: "Семінари", amount: (row) => row.seminar },
    { key: "exam", label: "Екзамен", short: "Іспит", amount: (row) => row.exam },
    { key: "credit", label: "Залік", short: "Залік", amount: (row) => row.credit },
    { key: "courseProject", label: "Курсовий проєкт", short: "Курсова", amount: (row) => row.courseProject },
    { key: "practiceHours", label: "Практика", short: "Практика", amount: (row) => row.practiceHours ?? 0 },
    { key: "diplomaProject", label: "Дипломне проєктування / ЕК", short: "Дипломні/ЕК", amount: (row) => row.diplomaProject ?? 0 },
    { key: "reviews", label: "Рецензії", short: "Рецензії", amount: (row) => row.reviews ?? 0 },
  ];
  const abbreviations = [...new Set(props.rows.flatMap((row) => row.abbreviations))];
  const semesters = [...new Set(props.rows.map((row) => row.semester).filter(Boolean))].sort((a, b) => a - b);
  const filtered = props.rows.filter((row) => {
    const teacherCode = trimText(props.assignments[distributionRowKey(row)] ?? row.teacherCode);
    const teacherName = props.teachers.find((teacher) => trimText(teacher.code) === teacherCode)?.name ?? "";
    const searchText = `${row.discipline} ${row.abbreviations.join(" ")} ${teacherName}`.toLocaleLowerCase("uk-UA");
    return (!query || searchText.includes(query.toLocaleLowerCase("uk-UA")))
      && (!abbreviation || row.abbreviations.includes(abbreviation)) && (!semester || String(row.semester) === semester)
      && (assignmentStatus === "all" || (assignmentStatus === "assigned" ? Boolean(teacherCode) : !teacherCode));
  });
  const totalHours = filtered.reduce((sum, row) => sum + number(row.total), 0);
  const assignedHours = filtered.filter((row) => trimText(props.assignments[distributionRowKey(row)] ?? row.teacherCode)).reduce((sum, row) => sum + number(row.total), 0);
  const vacantHours = Math.max(0, totalHours - assignedHours);
  const norm = (teacher: Teacher) => teacher.normHours * teacher.rate;
  const load = (teacher: Teacher) => props.workloads.filter((row) => row.teacher_code === teacher.code).reduce((sum, row) => sum + workloadTotal(row), 0);
  useEffect(() => {
    if (!fullscreen) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (selectedRow) setSelectedRow(null);
      else props.setFullscreen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [fullscreen, selectedRow, props.setFullscreen]);
  const rememberTableScroll = () => {
    const container = tableScrollRef.current;
    if (container) scrollPositionRef.current = { top: container.scrollTop, left: container.scrollLeft };
  };
  useLayoutEffect(() => {
    const container = tableScrollRef.current;
    if (!container) return;
    container.scrollTop = scrollPositionRef.current.top;
    container.scrollLeft = scrollPositionRef.current.left;
  }, [props.rows]);
  const openSplit = (row: PlanImportRow) => {
    const groupRows = props.rows.filter((item) => trimText(item.form_code) === trimText(row.form_code) && item.disciplineCode === row.disciplineCode && item.target === row.target && item.semester === row.semester
      && (item.levelCode || item.levelKind) === (row.levelCode || row.levelKind));
    const group = groupRows.length ? groupRows : [row];
    const sum = (read: (item: PlanImportRow) => number) => group.reduce((total, item) => total + read(item), 0);
    const groupedRow: PlanImportRow = { ...group[0], sourceRow: group[0].sourceRow, id: group[0].id,
      lecture: sum((item) => item.lecture), lab: sum((item) => item.lab), practice: sum((item) => item.practice), seminar: sum((item) => item.seminar),
      exam: sum((item) => item.exam), credit: sum((item) => item.credit), courseProject: sum((item) => item.courseProject),
      practiceHours: sum((item) => item.practiceHours ?? 0), diplomaProject: sum((item) => item.diplomaProject ?? 0), reviews: sum((item) => item.reviews ?? 0),
      total: sum((item) => item.total), teacherCode: group.find((item) => item.teacherCode)?.teacherCode ?? null };
    setSelectedRow(groupedRow);
    setQuickTeacher(group.find((item) => item.teacherCode)?.teacherCode ?? "");
    setAllocation(Object.fromEntries(kinds.map((kind) => [kind.key, group.map((item) => ({ teacherCode: item.teacherCode ?? "", hours: kind.amount(item) })).filter((part) => part.hours > 0)])) as SplitAllocation);
  };
  const pill = (background: string, color: string): React.CSSProperties => ({ display: "inline-flex", alignItems: "center", padding: "1px 6px", borderRadius: 999, background, color, fontSize: 11, lineHeight: 1.2, whiteSpace: "nowrap" });
  return <Card className={`workload-distribution-card ${fullscreen ? "workload-fullscreen-card" : ""}`} style={fullscreen ? { width: "100%", maxWidth: "100%", minWidth: 0, boxSizing: "border-box", overflow: "hidden", display: "flex", flexDirection: "column", flex: "1 1 auto", height: "100%", minHeight: 0, padding: 0 } : { width: "100%", maxWidth: "100%", minWidth: 0, boxSizing: "border-box", overflow: "visible", maxHeight: "none", height: "auto", minHeight: "auto", padding: 0 }}>
    {fullscreen && <FullscreenTabBar title="Розподіл годин" setFullscreen={props.setFullscreen} action={<><DownloadReportButton onClick={() => setReportPreviewOpen(true)} /><button type="button" className="workload-clear-assignments-button" onClick={() => void props.onClearAssignments()}>Очистити розподіл</button></>} />}
    {!fullscreen && <div className="workload-card-header"><div className="workload-header-title-line"><h2>Розподіл годин</h2><span>Показано <strong>{filtered.length}</strong> із <strong>{props.rows.length}</strong> записів<InlineHoursSummary total={totalHours} assigned={assignedHours} vacant={vacantHours} /></span></div><div className="workload-card-header-actions"><DownloadReportButton onClick={() => setReportPreviewOpen(true)} /><button type="button" className="workload-clear-assignments-button" onClick={() => void props.onClearAssignments()}>Очистити розподіл</button><button type="button" onClick={() => props.setFullscreen(true)} title="Розгорнути вкладку на весь екран">⛶ На весь екран</button></div></div>}
    <div className={`workload-distribution-filters ${fullscreen ? "is-fullscreen" : ""}`} style={filterBarStyle}>
      <FilterSearch ariaLabel="Пошук дисципліни або викладача" placeholder="Дисципліна або викладач" value={query} onChange={setQuery} />
      <FilterSelect ariaLabel="Фільтр спеціальності / ОПП" value={abbreviation} onChange={setAbbreviation} options={[{ value: "", label: "Усі спеціальності / ОПП" }, ...abbreviations.map((item) => ({ value: item, label: item }))]} />
      <FilterSelect ariaLabel="Фільтр семестру" value={semester} onChange={setSemester} options={[{ value: "", label: "Усі семестри" }, ...semesters.map((item) => ({ value: String(item), label: `Семестр ${item}` }))]} />
      <FilterSelect ariaLabel="Фільтр призначень" value={assignmentStatus} onChange={setAssignmentStatus} options={[{ value: "all", label: "Усі призначення" }, { value: "assigned", label: "Призначені" }, { value: "vacant", label: "Вакантні" }]} />
      {(query || abbreviation || semester || assignmentStatus !== "all") && <FilterReset onClick={() => { setQuery(""); setAbbreviation(""); setSemester(""); setAssignmentStatus("all"); }} />}
    </div>
    {props.suggestion && <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 14, padding: 12, border: "1px solid #bfdbfe", borderRadius: 10, background: "#eff6ff", fontSize: 13 }}><span>Є перше ручне призначення.</span><button type="button" onClick={() => void props.onAssignSameDiscipline()} className="rounded-lg border border-blue-200 bg-white px-3 py-2 font-semibold text-blue-800">На всі рядки дисципліни</button><button type="button" onClick={() => void props.onAutoAssign()} className="rounded-lg bg-blue-600 px-3 py-2 font-semibold text-white">Авторозподілити вакантні години</button></div>}
    <div ref={tableScrollRef} onScroll={rememberTableScroll} className={`workload-table-scroll ${fullscreen ? "is-fullscreen" : ""}`} style={{ width: "100%", maxWidth: "100%", minWidth: 0, margin: fullscreen ? 0 : "0 0 12px", display: "block", flex: fullscreen ? "1 1 auto" : undefined, minHeight: fullscreen ? 0 : undefined, height: fullscreen ? "100%" : "auto", overflowX: "auto", overflowY: fullscreen ? "auto" : "visible", maxHeight: "none", paddingBottom: fullscreen ? 32 : 0, boxSizing: "border-box" }}><table style={{ width: "100%", maxWidth: "100%", minWidth: 0, tableLayout: "fixed", borderCollapse: "collapse", fontSize: 12 }}><colgroup><col style={{ width: 210 }}/>{Array.from({ length: 10 }, (_, index) => <col key={index} style={{ width: 44 }}/>)}<col style={{ width: 82 }}/><col style={{ width: 185 }}/><col style={{ width: 126 }}/></colgroup><thead><tr>{["Дисципліна", "Лек.", "Лаб.", "Практ.", "Семін.", "Екз.", "Залік", "Курс.", "Практ.", "ДП/ЕК", "Рец.", "Всього, год", "Викладач", "Дії"].map((title, index) => <th key={title} className={index >= 1 && index <= 10 ? "workload-hour-header" : index === 11 ? "workload-total-hours-header" : undefined} title={index === 8 ? "Практика" : undefined} style={{ width: index >= 1 && index <= 10 ? 44 : undefined, minWidth: index >= 1 && index <= 10 ? 44 : undefined, maxWidth: index >= 1 && index <= 10 ? 44 : undefined, padding: index >= 1 && index <= 10 ? "10px 2px" : "8px 6px", background: "#eef2f6", color: "#334155", textAlign: index === 0 || index > 10 ? "left" : "center", borderBottom: "1px solid #e2e8f0", borderRight: "1px solid #f1f5f9", verticalAlign: "middle", height: 36, whiteSpace: "nowrap", fontSize: 11 }}>{title}</th>)}</tr></thead><tbody>{filtered.map((row) => {
      const related = props.rows.filter((item) => trimText(item.form_code) === trimText(row.form_code) && item.disciplineCode === row.disciplineCode && item.target === row.target && item.semester === row.semester && item.teacherCode);
      const values = [row.lecture, row.lab, row.practice, row.seminar, row.exam, row.credit, row.courseProject, row.practiceHours ?? 0, row.diplomaProject ?? 0, row.reviews ?? 0];
      return <tr key={row.id ?? row.sourceRow}>
        <td className="workload-discipline-cell" style={{ width: 210, minWidth: 210, maxWidth: 210, position: "sticky", left: 0, zIndex: 2, background: "#fff", overflow: "hidden" }}><strong title={row.discipline}>{row.discipline}</strong><div className="workload-inline-tags" title={[...row.abbreviations, row.semester > 0 ? `Семестр ${row.semester}` : "", row.study_form || "Денна"].filter(Boolean).join(" · ")}>{row.abbreviations.map((item) => <span key={item} style={pill("#dbeafe", "#1d4ed8")}>{item}</span>)}{row.semester > 0 && <span style={pill("#f1f5f9", "#475569")}>Семестр {row.semester}</span>}<span style={pill(row.study_form === "Заочна" ? "#f3e8ff" : "#dbeafe", row.study_form === "Заочна" ? "#7e22ce" : "#1d4ed8")}>{row.study_form || "Денна"}</span></div></td>
        {values.map((value, index) => <td key={index} className={value === 0 ? "workload-hour-zero" : "workload-hour-value"} style={{ width: 42, padding: "6px 4px", borderBottom: "1px solid #e2e8f0", borderRight: "1px solid #f1f5f9", verticalAlign: "middle", textAlign: "center", whiteSpace: "nowrap", fontSize: 12 }}>{hours(value)}</td>)}
        <td className="workload-total-hours-cell" style={{ width: 82, minWidth: 82, padding: "6px 4px", borderBottom: "1px solid #e2e8f0", borderRight: "1px solid #f1f5f9", textAlign: "center", verticalAlign: "middle", whiteSpace: "nowrap" }}><strong>{hours(row.total)}</strong></td>
        <td className="workload-teacher-cell" style={{ width: 185, minWidth: 170, maxWidth: 185, padding: "6px 5px", borderBottom: "1px solid #e2e8f0", borderRight: "1px solid #f1f5f9", verticalAlign: "middle", overflow: "visible" }}>{related.length > 1 ? <div className="workload-assigned-teachers">{related.map((part) => <span key={part.id} title={`${props.teachers.find((teacher) => teacher.code === part.teacherCode)?.name ?? part.teacherCode} · ${hours(part.total)} год`}>{props.teachers.find((teacher) => teacher.code === part.teacherCode)?.name ?? part.teacherCode} · {hours(part.total)} год</span>)}</div> : <SearchableSelect ariaLabel={`Викладач для ${row.discipline}`} placeholder="Не призначено" value={props.assignments[distributionRowKey(row)] ?? ""} triggerHeight={34} options={[{ value: "", label: "Не призначено" }, ...props.teachers.map((teacher) => ({ value: teacher.code, label: teacher.name, sublabel: `${hours(load(teacher))} / ${hours(norm(teacher))} год ліміту` }))]} onChange={(value) => { rememberTableScroll(); void props.onAssign(row, value); }} />}</td>
        <td className="workload-actions-cell" style={{ width: 126, minWidth: 126, maxWidth: 126, padding: "6px 12px 6px 4px", borderBottom: "1px solid #e2e8f0", verticalAlign: "middle", textAlign: "right", whiteSpace: "nowrap", position: "sticky", right: 0, background: "#fff", zIndex: 2 }}><button type="button" onClick={() => openSplit(row)} title="Детальний розподіл видів робіт">Поділити</button><button type="button" onClick={() => props.onCopyAssignment(row)} title="Застосувати до всіх рядків цієї дисципліни" aria-label="Застосувати до всієї дисципліни">На всі</button></td>
      </tr>;
    })}{!filtered.length && <tr><td colSpan={14} style={{ padding: 32, color: "#64748b", textAlign: "center" }}>Немає рядків за обраними фільтрами.</td></tr>}</tbody></table></div>
    {selectedRow && <DistributionSplitModal row={selectedRow} teachers={props.teachers} quickTeacher={quickTeacher} onQuickTeacherChange={setQuickTeacher} allocation={allocation} onAllocationChange={setAllocation} onQuickAssign={() => { if (quickTeacher) setAllocation(fullAllocationForTeacher(selectedRow, quickTeacher)); }} onApplyAll={() => { if (quickTeacher) { rememberTableScroll(); void props.onAssignSameDiscipline(selectedRow, quickTeacher); setSelectedRow(null); } }} onCancel={() => setSelectedRow(null)} onSave={() => { rememberTableScroll(); void props.onSplit(selectedRow, allocation); }} />}
    {reportPreviewOpen && <ReportPreviewModal report={{ title: `Розподіл навчального навантаження · ${props.year}`, subtitle: `${props.studyForm} · ${filtered.length} записів`, rows: filtered.map((row) => planRowForExport(row, props.teachers.find((teacher) => teacher.code === trimText(props.assignments[distributionRowKey(row)] ?? row.teacherCode))?.name ?? "")), fileName: `Розподіл_годин_${props.year.replace("/", "-")}`, organization: props.organization, kind: "plan" }} onClose={() => setReportPreviewOpen(false)} />}
  </Card>;
}

function DistributionSplitModal(props: { row: PlanImportRow; teachers: Teacher[]; quickTeacher: string; onQuickTeacherChange: (code: string) => void; allocation: SplitAllocation; onAllocationChange: (allocation: SplitAllocation) => void; onQuickAssign: () => void; onApplyAll: () => void; onCancel: () => void; onSave: () => void; vacancyRows?: PlanImportRow[]; onVacancyChange?: (row: PlanImportRow) => void; limitSummary?: { current: number; added: number; limit: number }; showApplyAll?: boolean }) {
  const [hourText, setHourText] = useState<Record<string, string>>({});
  const kinds: Array<{ key: SplitWorkKind; label: string; amount: (row: PlanImportRow) => number }> = [
    { key: "lecture", label: "Лекції", amount: (row) => row.lecture }, { key: "lab", label: "Лабораторні", amount: (row) => row.lab },
    { key: "practice", label: "Практичні", amount: (row) => row.practice }, { key: "seminar", label: "Семінарські", amount: (row) => row.seminar },
    { key: "exam", label: "Екзамен", amount: (row) => row.exam }, { key: "credit", label: "Залік", amount: (row) => row.credit },
    { key: "courseProject", label: "Курсовий проєкт", amount: (row) => row.courseProject }, { key: "practiceHours", label: "Практика", amount: (row) => row.practiceHours ?? 0 },
    { key: "diplomaProject", label: "Дипломне проєктування / ЕК", amount: (row) => row.diplomaProject ?? 0 },
    { key: "reviews", label: "Рецензії", amount: (row) => row.reviews ?? 0 },
  ];
  const button: React.CSSProperties = { height: 38, padding: "0 16px", borderRadius: 8, lineHeight: 1, fontSize: 13.5, fontWeight: 500, display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 6, whiteSpace: "nowrap", transition: "all 0.15s ease", cursor: "pointer" };
  const primary: React.CSSProperties = { ...button, color: "#fff", background: "#2563eb", border: "1px solid #2563eb" };
  const secondary: React.CSSProperties = { ...button, color: "#374151", background: "#fff", border: "1px solid #d1d5db" };
  const auxiliary: React.CSSProperties = { ...button, color: "#1d4ed8", background: "#eff6ff", border: "1px solid #bfdbfe" };
  const partsFor = (key: SplitWorkKind, amount: number) => props.allocation[key] ?? [{ teacherCode: props.row.teacherCode ?? "", hours: amount }];
  const updateParts = (key: SplitWorkKind, parts: SplitPart[]) => props.onAllocationChange({ ...props.allocation, [key]: parts });
  const valid = kinds.filter((kind) => kind.amount(props.row) > 0).every((kind) => {
    const parts = partsFor(kind.key, kind.amount(props.row));
    return parts.length > 0 && parts.every((part) => Number.isFinite(Number(part.hours)) && Number(part.hours) >= 0)
      && Math.abs(roundHours(parts.reduce((sum, part) => sum + Number(part.hours), 0)) - roundHours(kind.amount(props.row))) <= 0.01;
  });
  return <div role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) props.onCancel(); }} style={{ position: "fixed", inset: 0, zIndex: 1100, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, background: "rgba(15,23,42,.5)" }}><section role="dialog" aria-modal="true" aria-labelledby="distribution-dialog-title" style={{ width: "100%", maxWidth: 960, maxHeight: "90vh", overflowY: "auto", padding: 24, boxSizing: "border-box", borderRadius: 16, background: "#fff", boxShadow: "0 24px 60px rgba(15,23,42,.25)" }}>
    <header style={{ display: "flex", justifyContent: "space-between", gap: 16, paddingBottom: 16, borderBottom: "1px solid #e2e8f0" }}><div><h2 id="distribution-dialog-title" style={{ margin: 0, color: "#0f172a", fontSize: 20 }}>{props.vacancyRows ? "Додати навантаження" : props.row.discipline}</h2>{props.vacancyRows && <div style={{ width: "min(620px, 80vw)", marginTop: 10 }}><SearchableSelect ariaLabel="Вакантна дисципліна" placeholder="Оберіть дисципліну" value={String(props.row.id ?? "")} onChange={(value) => { const row = props.vacancyRows?.find((item) => String(item.id) === value); if (row) props.onVacancyChange?.(row); }} options={(props.vacancyRows ?? []).map((row) => ({ value: String(row.id ?? row.sourceRow), label: row.discipline, sublabel: `${row.target || "без абревіатури"} · семестр ${row.semester || "—"} · ${hours(row.total)} год вакантно` }))} /></div>}<div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", marginTop: 9 }}>{props.row.abbreviations.map((item) => <span key={item} style={{ padding: "4px 8px", borderRadius: 999, background: "#dbeafe", color: "#1d4ed8", fontSize: 11 }}>{item}</span>)}{props.row.semester > 0 && <span style={{ padding: "4px 8px", borderRadius: 999, background: "#f1f5f9", color: "#475569", fontSize: 11 }}>Семестр {props.row.semester}</span>}<strong style={{ marginLeft: 5 }}>Всього: {hours(props.row.total)} год</strong></div></div><button type="button" className="workload-modal-secondary" aria-label="Закрити" onClick={props.onCancel} style={{ ...secondary, width: 38, padding: 0 }}>✕</button></header>
    <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "16px 0", borderBottom: "1px solid #e2e8f0" }}><div style={{ flex: 1, minWidth: 220 }}><SearchableSelect ariaLabel="Викладач для всіх видів роботи" placeholder="Обрати викладача" value={props.quickTeacher} onChange={props.onQuickTeacherChange} options={[{ value: "", label: "Не обрано" }, ...props.teachers.map((teacher) => ({ value: teacher.code, label: teacher.name, sublabel: `${teacher.position} · ${teacher.rate} ставки` }))]} /></div><button type="button" className="workload-modal-primary" disabled={!props.quickTeacher} onClick={() => { setHourText({}); props.onQuickAssign(); }} style={{ ...primary, opacity: props.quickTeacher ? 1 : .55, cursor: props.quickTeacher ? "pointer" : "not-allowed" }}>Призначити всім видам</button></div>
    {props.limitSummary && <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginTop: 14, padding: "12px 14px", border: `1px solid ${props.limitSummary.current + props.limitSummary.added > props.limitSummary.limit ? "#fecaca" : "#bbf7d0"}`, borderRadius: 10, background: props.limitSummary.current + props.limitSummary.added > props.limitSummary.limit ? "#fff1f2" : "#f0fdf4", fontSize: 13 }}><span>Уже розподілено: <strong>{hours(props.limitSummary.current)} год</strong></span><span>Додається цьому викладачу: <strong>{hours(props.limitSummary.added)} год</strong></span><span>Річний ліміт: <strong>{hours(props.limitSummary.limit)} год</strong></span><strong style={{ color: props.limitSummary.current + props.limitSummary.added > props.limitSummary.limit ? "#be123c" : "#15803d" }}>{props.limitSummary.current + props.limitSummary.added > props.limitSummary.limit ? `Перевищення: ${hours(props.limitSummary.current + props.limitSummary.added - props.limitSummary.limit)} год` : `Залишок: ${hours(Math.max(0, props.limitSummary.limit - props.limitSummary.current - props.limitSummary.added))} год`}</strong></div>}
    <div style={{ display: "grid", gap: 10, paddingTop: 16 }}>{kinds.filter((kind) => kind.amount(props.row) > 0).map((kind) => {
      const amount = kind.amount(props.row);
      const parts = partsFor(kind.key, amount);
      const sum = roundHours(parts.reduce((total, part) => total + Number(part.hours || 0), 0));
       const validAmount = parts.every((part) => Number.isFinite(Number(part.hours)) && Number(part.hours) >= 0) && Math.abs(roundHours(sum) - roundHours(amount)) <= 0.01;
      const setPart = (index: number, changes: Partial<SplitPart>) => {
         const next = parts.map((part, currentIndex) => currentIndex === index ? { ...part, ...changes, hours: changes.hours === undefined ? part.hours : roundHours(Number(changes.hours)) } : part);
         if (next.length === 2 && index === 0 && changes.hours !== undefined) next[1] = { ...next[1], hours: Math.max(0, roundHours(amount - Number(changes.hours))) };
          if (changes.hours !== undefined && next.length === 2 && index === 0) setHourText((current) => { const updated = { ...current }; delete updated[`${kind.key}:1`]; return updated; });
        updateParts(kind.key, next);
      };
      const setHalf = () => {
        const firstTeacher = parts[0]?.teacherCode || props.row.teacherCode || props.teachers[0]?.code || "";
        const secondTeacher = parts[1]?.teacherCode || props.teachers.find((teacher) => teacher.code !== firstTeacher)?.code || "";
         const firstHours = roundHours(amount / 2);
          setHourText((current) => Object.fromEntries(Object.entries(current).filter(([key]) => !key.startsWith(`${kind.key}:`))));
         updateParts(kind.key, [{ teacherCode: firstTeacher, hours: firstHours }, { teacherCode: secondTeacher, hours: roundHours(amount - firstHours) }]);
      };
      return <div key={kind.key} style={{ padding: "12px 14px", background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 10 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 8 }}><div style={{ display: "flex", alignItems: "center", gap: 8 }}><strong>{kind.label}</strong><span style={{ padding: "4px 8px", borderRadius: 999, background: "#dbeafe", color: "#1d4ed8", fontSize: 12 }}>{hours(amount)} год</span></div><div style={{ display: "flex", alignItems: "center", gap: 6 }}><button type="button" className="workload-modal-split" onClick={() => updateParts(kind.key, [...parts, { teacherCode: "", hours: 0 }])} style={{ ...auxiliary, height: 30, padding: "0 10px", fontSize: 12, fontWeight: 600, borderRadius: 6 }}>+ Поділити</button><button type="button" className="workload-modal-split" onClick={setHalf} style={{ ...auxiliary, height: 30, padding: "0 10px", fontSize: 12, fontWeight: 600, borderRadius: 6 }}>50/50</button></div></div>
        <div style={{ display: "grid", gap: 7 }}>{parts.map((part, index) => { const inputKey = `${kind.key}:${index}`; return <div key={index} style={{ display: "grid", gridTemplateColumns: "minmax(160px,1fr) 120px auto", alignItems: "center", gap: 8 }}><SearchableSelect ariaLabel={`${kind.label}: викладач ${index + 1}`} placeholder="Нерозподілено" value={part.teacherCode} onChange={(value) => setPart(index, { teacherCode: value })} triggerHeight={36} options={[{ value: "", label: "Нерозподілено" }, ...props.teachers.map((teacher) => ({ value: teacher.code, label: teacher.name, sublabel: `${teacher.position} · ${teacher.rate} ставки` }))]} /><label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "#475569" }}><input type="text" inputMode="decimal" value={hourText[inputKey] ?? String(part.hours).replace(".", ",")} onChange={(event) => { const value = event.target.value; setHourText((current) => ({ ...current, [inputKey]: value })); setPart(index, { hours: parseHoursInput(value) }); }} onBlur={() => setHourText((current) => { const updated = { ...current }; delete updated[inputKey]; return updated; })} style={{ boxSizing: "border-box", width: 76, height: 36, borderRadius: 8, border: "1px solid #d1d5db", padding: "0 8px" }}/>год</label>{parts.length > 1 && <button type="button" aria-label="Прибрати частину" onClick={() => { setHourText({}); updateParts(kind.key, parts.filter((_, partIndex) => partIndex !== index)); }} style={{ ...secondary, width: 34, height: 34, padding: 0 }}>✕</button>}</div>; })}</div>
        <p style={{ margin: "8px 0 0", color: validAmount ? "#15803d" : "#b91c1c", fontSize: 12, fontWeight: 600 }}>{validAmount ? `Розподілено: ${hours(sum)} / ${hours(amount)} год` : sum < amount ? `Залишилось: ${hours(amount - sum)} год` : `Перевищено на ${hours(sum - amount)} год`}</p>
      </div>;
    })}</div>
    <footer style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, marginTop: 20, paddingTop: 16, borderTop: "1px solid #e5e7eb" }}>{props.showApplyAll !== false && <button type="button" className="workload-modal-auxiliary" onClick={props.onApplyAll} style={auxiliary}>Застосувати на всю дисципліну</button>}<div style={{ display: "flex", gap: 10, marginLeft: "auto" }}><button type="button" className="workload-modal-secondary" onClick={props.onCancel} style={{ ...secondary, minWidth: 110 }}>Скасувати</button><button type="button" className="workload-modal-primary" disabled={!valid || (props.vacancyRows !== undefined && !props.quickTeacher)} onClick={() => { props.onSave(); props.onCancel(); }} style={{ ...primary, minWidth: 150, opacity: valid ? 1 : .55, cursor: valid ? "pointer" : "not-allowed" }}>Зберегти розподіл</button></div></footer>
  </section></div>;
}

function ReportTab(props: { teachers: Teacher[]; teacherId: number | ""; setTeacherId: (id: number | "") => void; year: string; studyForm: StudyForm; records: WorkloadRecord[]; allRecords: WorkloadRecord[]; allPlanRows: PlanImportRow[]; selectedTeacher?: Teacher; facultyName: string; departmentName: string; departmentHeadName: string; vacantRows: PlanImportRow[]; onRemoveRow: (row: WorkloadRecord) => Promise<void>; onSplitVacancy: (row: PlanImportRow, allocation: SplitAllocation) => Promise<void>; onSplitRecord: (row: PlanImportRow, allocation: SplitAllocation) => Promise<void>; onAssignSameDiscipline: (row?: PlanImportRow, teacherCode?: string) => void; fullscreen: boolean; setFullscreen: (value: boolean | ((current: boolean) => boolean)) => void }) {
  const teacherName = props.selectedTeacher?.name ?? "Викладач";
  const teacherPosition = props.selectedTeacher?.position || "—";
  const reportTitle = `Педагогічне навантаження на ${props.year} навчальний рік`;
  const reportSubtitle = `${teacherName} · ${teacherPosition} (${props.selectedTeacher?.rate ?? 0} ставки) · ${props.studyForm === "Усі форми (Денна + Заочна)" ? props.studyForm : `${props.studyForm} форма`}`;
  const fileName = `Педагогічне_навантаження_${teacherName.replace(/\s+/g, "_")}_${props.year.replace("/", "-")}`;
  const [editingRow, setEditingRow] = useState<WorkloadRecord | null>(null);
  const [editingDistributionRow, setEditingDistributionRow] = useState<PlanImportRow | null>(null);
  const [editingDistributionIds, setEditingDistributionIds] = useState<number[]>([]);
  const [adding, setAdding] = useState(false);
  const [reportPreviewOpen, setReportPreviewOpen] = useState(false);
  const [selectedVacancy, setSelectedVacancy] = useState<PlanImportRow | null>(null);
  const [quickTeacher, setQuickTeacher] = useState(props.selectedTeacher?.code ?? "");
  const [allocation, setAllocation] = useState<SplitAllocation>({});
  useEffect(() => { setQuickTeacher(props.selectedTeacher?.code ?? ""); }, [props.teacherId, props.selectedTeacher?.code]);
  const planRowsForRecord = (record: WorkloadRecord) => {
    const recordCode = trimText(record.discipline_code);
    const recordName = normalizedName(record.discipline_name);
    const recordForms = new Set([record.abbreviation, record.groups].flatMap((value) => trimText(value).split(/[;,]/).map((part) => trimText(part))).filter(Boolean));
    return props.allPlanRows.filter((row) => {
      const sameDiscipline = recordCode
        ? trimText(row.disciplineCode) === recordCode
        : normalizedName(row.discipline) === recordName;
      const rowForms = [row.target, ...row.abbreviations].map(trimText).filter(Boolean);
      const sameSpecialty = !recordForms.size || rowForms.some((value) => recordForms.has(value));
      const sameLevel = !record.level_code || !row.levelCode || trimText(row.levelCode) === trimText(record.level_code);
      return sameDiscipline && sameSpecialty && number(row.semester) === number(record.semester)
        && trimText(row.form_code) === trimText(record.form_code) && sameLevel;
    });
  };
  const openDistributionEditor = (record: WorkloadRecord) => {
    const sourceRows = planRowsForRecord(record);
    if (!sourceRows.length) {
      notify(`Не вдалося знайти планові години для «${record.discipline_name}». Оновіть дані навантаження та спробуйте ще раз.`, "error", "Звіт викладача");
      return;
    }
    const base = sourceRows[0];
    const sum = (read: (row: PlanImportRow) => number) => roundHours(sourceRows.reduce((total, row) => total + read(row), 0));
    const groupedRow: PlanImportRow = {
      ...base,
      lecture: sum((row) => row.lecture), lab: sum((row) => row.lab), practice: sum((row) => row.practice), seminar: sum((row) => row.seminar),
      exam: sum((row) => row.exam), credit: sum((row) => row.credit), courseProject: sum((row) => row.courseProject),
      practiceHours: sum((row) => row.practiceHours ?? 0), diplomaProject: sum((row) => row.diplomaProject ?? 0), reviews: sum((row) => row.reviews ?? 0),
      total: sum((row) => row.total), teacherCode: record.teacher_code ?? null,
    };
    const workKinds: Array<[SplitWorkKind, (row: PlanImportRow) => number]> = [
      ["lecture", (row) => row.lecture], ["lab", (row) => row.lab], ["practice", (row) => row.practice], ["seminar", (row) => row.seminar],
      ["exam", (row) => row.exam], ["credit", (row) => row.credit], ["courseProject", (row) => row.courseProject], ["practiceHours", (row) => row.practiceHours ?? 0],
      ["diplomaProject", (row) => row.diplomaProject ?? 0], ["reviews", (row) => row.reviews ?? 0],
    ];
    const initialAllocation = Object.fromEntries(workKinds.map(([key, amount]) => [key, sourceRows.map((row) => ({ teacherCode: trimText(row.teacherCode), hours: amount(row) })).filter((part) => part.hours > 0)])) as SplitAllocation;
    setEditingRow(record);
    setAdding(false);
    setEditingDistributionRow(groupedRow);
    setEditingDistributionIds(sourceRows.map((row) => row.id).filter((id): id is number => id != null));
    setQuickTeacher(record.teacher_code ?? "");
    setAllocation(initialAllocation);
  };
  const displayedRecords = props.records;
  const limitTotalHours = props.records.reduce((sum, row) => sum + workloadTotal(row), 0);
  const firstSemester = props.records.filter((row) => number(row.semester) > 0 && number(row.semester) % 2 === 1).reduce((sum, row) => sum + workloadTotal(row), 0);
  const secondSemester = props.records.filter((row) => number(row.semester) > 0 && number(row.semester) % 2 === 0).reduce((sum, row) => sum + workloadTotal(row), 0);
  const maxAllowed = (props.selectedTeacher?.normHours ?? 0) * (props.selectedTeacher?.rate ?? 0);
  const allocationTeacher = props.teachers.find((teacher) => teacher.code === quickTeacher) ?? props.selectedTeacher;
  const allocationTeacherRows = props.allRecords.filter((row) => row.teacher_id === allocationTeacher?.id && (!editingDistributionRow || !editingDistributionIds.includes(row.id)));
  const allocationTeacherCurrent = allocationTeacherRows.reduce((sum, row) => sum + workloadTotal(row), 0);
  const allocationTeacherLimit = (allocationTeacher?.normHours ?? 0) * (allocationTeacher?.rate ?? 0);
  const proposedHours = Object.values(allocation).flatMap((parts) => parts ?? []).filter((part) => part.teacherCode === quickTeacher).reduce((sum, part) => sum + number(part.hours), 0);
  const overLimit = limitTotalHours > maxAllowed;
  const otherTeacherNames = (row: WorkloadRecord) => [...new Set(props.allRecords.filter((item) => item.teacher_id != null && String(item.teacher_id) !== String(props.teacherId)
    && (trimText(item.discipline_code) === trimText(row.discipline_code) || (!trimText(row.discipline_code) && trimText(item.discipline_name) === trimText(row.discipline_name)))
    && number(item.semester) === number(row.semester)
    && trimText(item.groups) === trimText(row.groups)
    && trimText(item.academic_year) === trimText(row.academic_year)
    && trimText(item.form_code) === trimText(row.form_code)
    && trimText(item.study_form) === trimText(row.study_form))
    .map((item) => props.teachers.find((teacher) => teacher.id === item.teacher_id)?.name)
    .filter((name): name is string => Boolean(name)))].join(", ") || "—";
  const exportRows: ExportWorkload[] = displayedRecords.map((row) => ({ ...row, groups: props.studyForm === "Усі форми (Денна + Заочна)" ? `${row.groups || "—"} · ${row.study_form}` : row.groups, notes: otherTeacherNames(row), study_form: row.study_form }));
  const openAdd = () => {
    const first = props.vacantRows.find((row) => row.id);
    setEditingRow(null);
    setEditingDistributionRow(null);
    setEditingDistributionIds([]);
    setAdding(true);
    setSelectedVacancy(first ?? null);
    setQuickTeacher(props.selectedTeacher?.code ?? "");
    setAllocation(first ? fullAllocationForTeacher(first, props.selectedTeacher?.code ?? "") : {});
  };
  const closeEditor = () => { setEditingRow(null); setEditingDistributionRow(null); setEditingDistributionIds([]); setAdding(false); setSelectedVacancy(null); setAllocation({}); };
  return <Card className={`workload-report-card ${props.fullscreen ? "workload-fullscreen-card" : ""}`} style={{ width: "100%", boxSizing: "border-box", padding: 0 }}>
    {props.fullscreen && <FullscreenTabBar title="Звіт викладача" setFullscreen={props.setFullscreen} action={<DownloadReportButton onClick={() => setReportPreviewOpen(true)} />} />}
     {!props.fullscreen && <div className="workload-card-header"><div className="workload-header-title-line"><h2>Звіт викладача</h2><span>Показано <strong>{displayedRecords.length}</strong> із <strong>{props.records.length}</strong> рядків · {props.year} · {props.studyForm}</span></div><div className="workload-card-header-actions"><DownloadReportButton onClick={() => setReportPreviewOpen(true)} /><FullscreenToggleButton onClick={() => props.setFullscreen(true)} /></div></div>}
    <div className="workload-report-controls">
      <div className="workload-report-teacher-row">
        <label className="workload-report-teacher-select">Викладач<SearchableSelect ariaLabel="Викладач" placeholder="Оберіть викладача" value={String(props.teacherId)} onChange={(value) => props.setTeacherId(value ? Number(value) : "")} options={props.teachers.map((teacher) => ({ value: String(teacher.id), label: teacher.name, sublabel: `${teacher.position} · ${teacher.rate} ставки${teacher.employmentType ? ` · ${teacher.employmentType}` : ""}` }))} /></label>
        <button type="button" className="workload-add-assignment-button" disabled={!props.selectedTeacher || !props.vacantRows.length} onClick={openAdd}><FiPlus size={15} aria-hidden="true" />Додати навантаження</button>
      </div>
    </div>
    <section aria-label="Контроль ліміту навантаження" className={`rounded-xl border ${overLimit ? "border-rose-300 bg-rose-50" : "border-emerald-200 bg-emerald-50"}`} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 14, flexWrap: "wrap", width: "100%", padding: "12px 18px", boxSizing: "border-box" }}>
      <strong style={{ minWidth: 175, color: "#334155", fontSize: 13 }}>Контроль ліміту навантаження</strong>
      <div className="workload-limit-metrics" style={{ alignItems: "center", gap: 12, flex: "1 1 700px", fontSize: 12 }}>
        <div><span className="text-slate-500">Посада та ставка</span><div className="font-semibold text-slate-800">{teacherPosition} · {props.selectedTeacher?.rate ?? 0} ставки</div></div>
        <div><span className="text-slate-500">Річний ліміт</span><div className="font-semibold text-slate-800">{hours(maxAllowed)} год</div></div>
        <div><span className="text-slate-500">Розподілено</span><div className="font-semibold text-slate-800">{hours(limitTotalHours)} год · І сем: {hours(firstSemester)} · ІІ сем: {hours(secondSemester)}</div></div>
        <div><span className="text-slate-500">Статус / залишок</span><div className={`font-bold ${overLimit ? "text-rose-700" : "text-emerald-700"}`}>{overLimit ? `Перевищення ліміту на ${hours(limitTotalHours - maxAllowed)} год!` : `Можна додати: ${hours(Math.max(0, maxAllowed - limitTotalHours))} год`}</div></div>
      </div>
    </section>
    <div className="workload-report-document" style={{ width: "100%", minWidth: 0 }}>
      <div className="workload-tab-table-scroll workload-teacher-report-table-scroll" style={{ width: "100%", maxWidth: "100%", minWidth: 0, overflow: "visible" }}><table className="workload-teacher-report-table" style={{ width: "100%", maxWidth: "100%", minWidth: 0, tableLayout: "fixed", borderCollapse: "collapse", fontSize: 12, color: "#1e293b" }}>
        <colgroup>{teacherReportColumnPercentages.map((width, index) => <col key={index} style={{ width: `${width}%` }}/>)}</colgroup>
        <thead style={{ background: "#f8fafc", color: "#475569" }}><tr>{["№", "Навчальна дисципліна", "СПЕЦ. / ОПП", "СЕМ.", ...workloadHourFields.map(({ key }) => ({ lectures: "ЛЕК.", labs: "ЛАБ.", practices: "ПРАКТ.", seminars: "СЕМІН.", exams: "ЕКЗ.", credits: "ЗАЛІК", course_projects: "КУРС.", practice_hours: "ПРАКТ-КА", diploma_project: "ДП/ЕК", reviews: "РЕЦ." }[key])), "ВСЬОГО", "ПРИМІТКИ", "ДІЇ"].map((heading, index) => <th key={`${heading}-${index}`} style={{ whiteSpace: index === 1 || index === 2 || index === 15 ? "normal" : "nowrap", wordBreak: "normal", overflowWrap: index === 1 || index === 2 || index === 15 ? "break-word" : "normal", fontSize: index >= 3 && index <= 14 ? 9.5 : 10, fontWeight: 700, lineHeight: 1.15, padding: index === 1 || index === 15 ? "10px 6px" : "10px 1px", borderBottom: "1px solid #e2e8f0", textAlign: index === 1 || index === 15 ? "left" : "center", verticalAlign: "middle" }}>{heading}</th>)}</tr></thead>
         <tbody className="divide-y divide-slate-100">{displayedRecords.length ? displayedRecords.map((row, index) => <tr key={row.id} style={{ borderBottom: "1px solid #f1f5f9" }}><td style={{ textAlign: "center", whiteSpace: "nowrap" }}>{index + 1}</td><td style={{ fontWeight: 600, wordBreak: "normal", overflowWrap: "break-word" }}>{row.discipline_name}</td><td style={{ whiteSpace: "normal", wordBreak: "normal", overflowWrap: "break-word" }}>{row.groups || "—"}{props.studyForm === "Усі форми (Денна + Заочна)" && <span style={{ display: "inline-flex", marginLeft: 5, padding: "3px 7px", borderRadius: 999, background: row.study_form === "Заочна" ? "#f3e8ff" : "#dbeafe", color: row.study_form === "Заочна" ? "#7e22ce" : "#1d4ed8", fontSize: 10 }}>{row.study_form}</span>}</td><td style={{ textAlign: "center", whiteSpace: "nowrap" }}>{row.semester || "—"}</td>{workloadHourFields.map(({ key }) => <td key={key} className={number(row[key]) === 0 ? "workload-hour-zero" : "workload-hour-value"} style={{ textAlign: "center", whiteSpace: "nowrap", fontSize: 12.5 }}>{hours(number(row[key]))}</td>)}<td style={{ textAlign: "center", whiteSpace: "nowrap", fontWeight: 700 }}>{hours(workloadTotal(row))}</td><td style={{ whiteSpace: "normal", wordBreak: "normal", overflowWrap: "break-word", lineHeight: 1.35, fontSize: 12.5, color: "#475569" }}>{otherTeacherNames(row) === "—" ? <span style={{ color: "#94a3b8" }}>—</span> : otherTeacherNames(row)}</td><td style={{ textAlign: "center", whiteSpace: "nowrap" }}><div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: 4 }}><button type="button" aria-label={`Редагувати розподіл годин: ${row.discipline_name}`} title="Редагувати розподіл годин" onClick={() => openDistributionEditor(row)} style={{ width: 28, height: 30, display: "inline-flex", alignItems: "center", justifyContent: "center", padding: 0, border: "1px solid #cbd5e1", borderRadius: 7, background: "#fff", color: "#334155", cursor: "pointer" }}><FiEdit2 size={14}/></button><button type="button" aria-label={`Зняти години: ${row.discipline_name}`} title="Зняти години" onClick={() => { if (window.confirm(`Зняти навантаження «${row.discipline_name}» і повернути години у вакансії?`)) void props.onRemoveRow(row); }} style={{ width: 28, height: 30, display: "inline-flex", alignItems: "center", justifyContent: "center", padding: 0, border: "1px solid #fecaca", borderRadius: 7, background: "#fff7f7", color: "#b91c1c", cursor: "pointer" }}><FiTrash2 size={14}/></button></div></td></tr>) : <tr><td colSpan={17} style={{ padding: 24, textAlign: "center", color: "#64748b" }}>{props.records.length ? "За обраними фільтрами дисциплін немає." : "За обраною формою навчання цьому викладачу ще не призначено дисциплін"}</td></tr>}</tbody>{displayedRecords.length > 0 && <tfoot><tr style={{ background: "#f8fafc", fontWeight: 700 }}><td colSpan={4} style={{ padding: "10px 8px", borderTop: "2px solid #cbd5e1" }}>ВСЬОГО ГОДИН</td>{workloadHourFields.map(({ key }) => <td key={key} style={{ padding: "10px 4px", textAlign: "center", fontWeight: 700, background: "#f8fafc", borderTop: "2px solid #cbd5e1" }}>{hours(displayedRecords.reduce((sum, row) => sum + number(row[key]), 0))}</td>)}<td style={{ padding: "10px 4px", textAlign: "center", fontWeight: 700, background: "#f8fafc", borderTop: "2px solid #cbd5e1" }}>{hours(displayedRecords.reduce((sum, row) => sum + workloadTotal(row), 0))}</td><td style={{ background: "#f8fafc", borderTop: "2px solid #cbd5e1" }} /><td style={{ background: "#f8fafc", borderTop: "2px solid #cbd5e1" }} /></tr></tfoot>}
      </table></div>
    </div>
    {editingRow && editingDistributionRow && <DistributionSplitModal row={editingDistributionRow} teachers={props.teachers} quickTeacher={quickTeacher} onQuickTeacherChange={setQuickTeacher} allocation={allocation} onAllocationChange={setAllocation} onQuickAssign={() => { if (quickTeacher) setAllocation(fullAllocationForTeacher(editingDistributionRow, quickTeacher)); }} onApplyAll={() => { if (quickTeacher) { void props.onAssignSameDiscipline(editingDistributionRow, quickTeacher); closeEditor(); } }} onCancel={closeEditor} onSave={() => { void props.onSplitRecord(editingDistributionRow, allocation); }} limitSummary={{ current: allocationTeacherCurrent, added: proposedHours, limit: allocationTeacherLimit }} />}
    {adding && selectedVacancy && <DistributionSplitModal row={selectedVacancy} vacancyRows={props.vacantRows} onVacancyChange={(row) => { setSelectedVacancy(row); setAllocation(fullAllocationForTeacher(row, quickTeacher)); }} teachers={props.teachers} quickTeacher={quickTeacher} onQuickTeacherChange={setQuickTeacher} allocation={allocation} onAllocationChange={setAllocation} onQuickAssign={() => { if (quickTeacher) setAllocation(fullAllocationForTeacher(selectedVacancy, quickTeacher)); }} onApplyAll={() => {}} showApplyAll={false} limitSummary={{ current: allocationTeacherCurrent, added: proposedHours, limit: allocationTeacherLimit }} onCancel={closeEditor} onSave={() => { if (allocationTeacher) void props.onSplitVacancy(selectedVacancy, allocation); }} />}
    {reportPreviewOpen && <ReportPreviewModal report={{ title: reportTitle, subtitle: reportSubtitle, rows: exportRows, fileName, organization: { facultyName: props.facultyName, departmentName: props.departmentName, departmentHeadName: props.departmentHeadName }, kind: "teacher" }} onClose={() => setReportPreviewOpen(false)} />}
  </Card>;
}

function ReportPreviewModal({ report, onClose }: { report: Omit<EditableReport, "comment" | "reportDate" | "teacherSigner" | "departmentSigner">; onClose: () => void }) {
  const [title, setTitle] = useState(report.title);
  const [subtitle, setSubtitle] = useState(report.subtitle);
  const [rows, setRows] = useState(report.rows);
  const [comment, setComment] = useState("");
  const [reportDate, setReportDate] = useState(new Date().toLocaleDateString("uk-UA"));
  const [teacherSigner, setTeacherSigner] = useState("");
  const [departmentSigner, setDepartmentSigner] = useState(report.organization.departmentHeadName || "");
  const [busy, setBusy] = useState(false);
  const columns: Array<[string, keyof ExportWorkload]> = [["Лек.", "lectures"], ["Лаб.", "labs"], ["Практ.", "practices"], ["Семін.", "seminars"], ["Екз.", "exams"], ["Залік", "credits"], ["Курс.", "course_projects"], ["Практ-ка", "practice_hours"], ["ДП / ЕК", "diploma_project"], ["Рец.", "reviews"]];
  const saveFile = async (format: "docx" | "xlsx" | "pdf") => {
    setBusy(true);
    try { await exportEditableReport({ ...report, title, subtitle, rows, comment, reportDate, teacherSigner, departmentSigner }, format); }
    catch (error) { notify(error instanceof Error ? error.message : "Не вдалося сформувати звіт.", "error"); }
    finally { setBusy(false); }
  };
  const restore = () => { setTitle(report.title); setSubtitle(report.subtitle); setRows(report.rows); setComment(""); setReportDate(new Date().toLocaleDateString("uk-UA")); setTeacherSigner(""); setDepartmentSigner(report.organization.departmentHeadName || ""); };
  const inputStyle: React.CSSProperties = { boxSizing: "border-box", width: "100%", minWidth: 110, height: 32, padding: "0 8px", border: "1px solid #cbd5e1", borderRadius: 6, background: "#fff", color: "#0f172a", font: "inherit" };
  const rowTable = (group: ExportWorkload[], section?: string) => <section key={section || "teacher"} style={{ marginTop: 14 }}>
    {section && <h3 style={{ margin: "0 0 7px", fontSize: 14, fontWeight: 700 }}>{section}</h3>}
    <div style={{ maxWidth: "100%", overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}><thead><tr>{["№", "Навчальна дисципліна", "Спеціальність / ОПП", "Семестр", "Форма", ...columns.map(([name]) => name), "Разом, год", "Примітки"].map((heading) => <th key={heading} style={{ padding: "7px 5px", borderBottom: "1px solid #cbd5e1", textAlign: "left", whiteSpace: "nowrap", background: "#f1f5f9" }}>{heading}</th>)}</tr></thead><tbody>{group.map((row, index) => <tr key={`${row.teacher_id}-${row.study_form}-${row.discipline_name}-${index}`}>
      <td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{index + 1}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.discipline_name}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.groups || "—"}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0", textAlign: "center" }}>{row.semester || "—"}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.study_form || "—"}</td>
      {columns.map(([heading, key]) => <td key={heading} style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0", textAlign: "right" }}>{number(row[key]) ? hours(number(row[key])) : "—"}</td>)}<td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0", textAlign: "right", fontWeight: 700 }}>{hours(workloadExportTotal(row))}</td><td style={{ minWidth: 140, padding: "4px", borderBottom: "1px solid #e2e8f0" }}><input aria-label={`Примітка: ${row.discipline_name}`} style={inputStyle} value={row.notes || ""} onChange={(event) => setRows((current) => current.map((item) => item === row ? { ...item, notes: event.target.value } : item))} /></td>
    </tr>)}</tbody><tfoot><tr style={{ background: "#f1f5f9", fontWeight: 700 }}><td colSpan={5} style={{ padding: "7px 5px", whiteSpace: "nowrap" }}>Всього</td>{columns.map(([heading, key]) => <td key={heading} style={{ padding: "7px 5px", textAlign: "right" }}>{hours(group.reduce((sum, row) => sum + number(row[key]), 0))}</td>)}<td style={{ padding: "7px 5px", textAlign: "right" }}>{hours(group.reduce((sum, row) => sum + workloadExportTotal(row), 0))}</td><td /></tr></tfoot></table></div></section>;
  const planTable = (group: ExportWorkload[]) => <section style={{ marginTop: 14 }}>
    <div style={{ maxWidth: "100%", overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}><thead><tr>{["№", "Дисципліна", "Спеціальність / ОПП", "Семестр", "Форма", ...columns.map(([name]) => name), "Всього, год", ...(report.hideAssignedTeacher ? [] : ["Призначений викладач"])].map((heading) => <th key={heading} style={{ padding: "7px 5px", borderBottom: "1px solid #cbd5e1", textAlign: "left", whiteSpace: "nowrap", background: "#f1f5f9" }}>{heading}</th>)}</tr></thead>
      <tbody>{group.map((row, index) => <tr key={`${row.teacher_id}-${row.discipline_name}-${row.study_form}-${index}`}>
        <td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{index + 1}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.discipline_name || "—"}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.groups || "—"}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0", textAlign: "center" }}>{row.semester || "—"}</td><td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.study_form || "—"}</td>
        {columns.map(([heading, key]) => <td key={heading} style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0", textAlign: "right" }}>{number(row[key]) ? hours(number(row[key])) : "—"}</td>)}<td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0", textAlign: "right", fontWeight: 700 }}>{hours(workloadExportTotal(row))}</td>{!report.hideAssignedTeacher && <td style={{ padding: "6px 5px", borderBottom: "1px solid #e2e8f0" }}>{row.assigned_teacher || "—"}</td>}
      </tr>)}</tbody><tfoot><tr style={{ background: "#f1f5f9", fontWeight: 700 }}><td colSpan={5} style={{ padding: "7px 5px" }}>Всього</td>{columns.map(([heading, key]) => <td key={heading} style={{ padding: "7px 5px", textAlign: "right" }}>{hours(group.reduce((sum, row) => sum + number(row[key]), 0))}</td>)}<td style={{ padding: "7px 5px", textAlign: "right" }}>{hours(group.reduce((sum, row) => sum + workloadExportTotal(row), 0))}</td>{!report.hideAssignedTeacher && <td />}</tr></tfoot></table></div>
  </section>;
  const disciplinesTable = () => <section style={{ marginTop: 14 }}><div style={{ maxWidth: "100%", overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}><thead><tr>{["№ з/п", "Спеціальність / ОПП", "Дисципліна", "Семестр та форма навчання", "Форма проведення (контроль)", "Викладачі"].map((heading) => <th key={heading} style={{ padding: "9px 8px", borderBottom: "1px solid #cbd5e1", textAlign: "left", whiteSpace: "nowrap", background: "#f1f5f9", color: "#334155" }}>{heading}</th>)}</tr></thead><tbody>{(report.disciplineRows ?? []).length ? (report.disciplineRows ?? []).map((row, index) => <tr key={`${row.specialty}-${row.discipline}-${row.semester}-${row.forms.join()}-${index}`}><td style={{ padding: "9px 8px", borderBottom: "1px solid #e2e8f0", verticalAlign: "top" }}>{index + 1}</td><td style={{ padding: "9px 8px", borderBottom: "1px solid #e2e8f0", verticalAlign: "top" }}>{row.specialty || "—"}</td><td style={{ padding: "9px 8px", borderBottom: "1px solid #e2e8f0", verticalAlign: "top", fontWeight: 600 }}>{row.discipline}</td><td style={{ padding: "9px 8px", borderBottom: "1px solid #e2e8f0", verticalAlign: "top" }}><span>{row.semester || "—"}</span><div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginTop: 4 }}>{row.forms.map((form) => <span key={form} style={{ padding: "2px 7px", borderRadius: 999, background: form === "Заочна" ? "#f3e8ff" : "#dbeafe", color: form === "Заочна" ? "#7e22ce" : "#1d4ed8", fontSize: 11 }}>{form}</span>)}</div></td><td style={{ padding: "9px 8px", borderBottom: "1px solid #e2e8f0", verticalAlign: "top" }}>{row.control}</td><td style={{ padding: "9px 8px", borderBottom: "1px solid #e2e8f0", verticalAlign: "top" }}>{row.teachers.length ? <div style={{ display: "grid", gap: 5 }}>{row.teachers.map((teacher) => <div key={teacher.name}><strong>{teacher.name}</strong>{teacher.workTypes.length > 0 && <span> ({teacher.workTypes.filter((kind) => !/^(екзамен|залік)$/i.test(kind.trim())).map((kind) => kind.toLocaleLowerCase("uk-UA")).join(", ")})</span>}</div>)}</div> : <span style={{ color: "#94a3b8" }}>Не призначено</span>}</td></tr>) : <tr><td colSpan={6} style={{ padding: 24, color: "#64748b", textAlign: "center" }}>Дисциплін за обраними параметрами немає.</td></tr>}</tbody></table></div></section>;
  const form56Table = () => {
    const kinds = columns;
    const selection = report.form56Selection ?? "all";
    const teacherIds = [...new Set(rows.map((row) => row.teacher_id))];
    const grouped = teacherIds.map((teacherId) => {
      const teacherRows = rows.filter((row) => row.teacher_id === teacherId);
      const base = teacherRows[0] ?? { teacher_id: teacherId, discipline_name: "", groups: "" } as ExportWorkload;
      const day = teacherRows.find((row) => row.study_form === "Денна") || { ...base, study_form: "Денна", groups: "денна", ...Object.fromEntries(kinds.map(([, key]) => [key, 0])) } as ExportWorkload;
      const extramural = teacherRows.find((row) => row.study_form === "Заочна") || { ...base, study_form: "Заочна", groups: "заочна", ...Object.fromEntries(kinds.map(([, key]) => [key, 0])) } as ExportWorkload;
      const overall = { ...day, study_form: "Всього", ...Object.fromEntries(kinds.map(([, key]) => [key, number(day[key]) + number(extramural[key])])) } as ExportWorkload;
      return { day, extramural, overall };
    });
    const sumFor = (source: ExportWorkload[], key: keyof ExportWorkload) => source.reduce((sum, row) => sum + number(row[key]), 0);
    const allDay = grouped.map(({ day }) => day);
    const allExtramural = grouped.map(({ extramural }) => extramural);
    const totals = (source: ExportWorkload[]) => kinds.map(([, key]) => sumFor(source, key));
    const rowTotal = (source: ExportWorkload[]) => source.reduce((sum, row) => sum + workloadExportTotal(row), 0);
    const footerRows: Array<[string, ExportWorkload[]]> = selection === "all"
      ? [["Всього за денною формою", allDay], ["Всього за заочною формою", allExtramural], ["Загалом по кафедрі (Денна + Заочна)", [...allDay, ...allExtramural]]]
      : selection === "day" ? [["Всього за денною формою", allDay]] : [["Всього за заочною формою", allExtramural]];
    const cell = (content: React.ReactNode, style: React.CSSProperties = {}, key?: React.Key, rowSpan?: number) => <td key={key} rowSpan={rowSpan} style={{ padding: "6px 8px", borderBottom: "1px solid #e2e8f0", ...style }}>{content}</td>;
    return <section style={{ marginTop: 14 }}><div style={{ maxWidth: "100%", overflowX: "auto" }}><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10, color: "#0f172a" }}>
      <thead><tr>{["№", "Викладач", "Посада", "Ставка", "Форма навчання", ...kinds.map(([label]) => label), "Разом, год"].map((heading) => <th key={heading} style={{ padding: "6px 8px", borderBottom: "1px solid #cbd5e1", background: "#f1f5f9", whiteSpace: "nowrap", textAlign: "left" }}>{heading}</th>)}</tr></thead>
      <tbody>{grouped.map(({ day, extramural, overall }, index) => {
        const selectedRows = selection === "day" ? [day] : selection === "extramural" ? [extramural] : [day, extramural, overall];
        return <React.Fragment key={day.teacher_id}>{selectedRows.map((row, rowIndex) => {
          const formLabel = row.study_form === "Денна" ? "Денна" : row.study_form === "Заочна" ? "Заочна" : "Всього";
          const isOverall = formLabel === "Всього";
          const badgeColor = formLabel === "Денна" ? { background: "#dbeafe", color: "#1d4ed8" } : formLabel === "Заочна" ? { background: "#f3e8ff", color: "#7e22ce" } : { background: "#e2e8f0", color: "#334155" };
          const rowSpan = selection === "all" ? 3 : 1;
          return <tr key={`${day.teacher_id}-${formLabel}`} style={isOverall ? { background: "#f8fafc", fontWeight: 700, borderBottom: "2px solid #cbd5e1" } : undefined}>
            {rowIndex === 0 && <>{cell(index + 1, { textAlign: "center", verticalAlign: "middle" }, `number-${day.teacher_id}`, rowSpan)}{cell(day.discipline_name, { fontWeight: 600, verticalAlign: "middle" }, `name-${day.teacher_id}`, rowSpan)}{cell(day.position || "—", { verticalAlign: "middle" }, `position-${day.teacher_id}`, rowSpan)}{cell(day.rate ?? "—", { textAlign: "center", verticalAlign: "middle" }, `rate-${day.teacher_id}`, rowSpan)}</>}
            {cell(<span style={{ display: "inline-flex", padding: "2px 7px", borderRadius: 999, background: badgeColor.background, color: badgeColor.color }}>{formLabel}</span>, { whiteSpace: "nowrap", textAlign: "left", fontWeight: isOverall ? 700 : 500 })}
            {kinds.map(([label, key]) => cell(hours(number(row[key])), { textAlign: "right", whiteSpace: "nowrap", background: isOverall ? "#f8fafc" : undefined }, `${row.teacher_id}-${formLabel}-${label}`))}
            {cell(hours(workloadExportTotal(row)), { textAlign: "right", fontWeight: 700, whiteSpace: "nowrap", background: isOverall ? "#f8fafc" : undefined })}
          </tr>;
        })}</React.Fragment>;
      })}</tbody>
    <tfoot>{footerRows.map(([label, list]) => <tr key={label} style={{ background: "#eef2f6", fontWeight: 700 }}><td colSpan={5} style={{ padding: "6px 8px", borderTop: "2px solid #94a3b8" }}>{label}</td>{totals(list).map((value, valueIndex) => <td key={valueIndex} style={{ padding: "6px 8px", borderTop: "2px solid #94a3b8", textAlign: "right" }}>{hours(value)}</td>)}<td style={{ padding: "6px 8px", borderTop: "2px solid #94a3b8", textAlign: "right" }}>{hours(rowTotal(list))}</td></tr>)}</tfoot>
    </table></div></section>;
  };
  return <div role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }} style={{ position: "fixed", inset: 0, zIndex: 1400, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, background: "rgba(15,23,42,.55)" }}><section role="dialog" aria-modal="true" aria-labelledby="report-preview-title" style={{ width: "min(1440px, 100%)", maxHeight: "94vh", display: "flex", flexDirection: "column", overflow: "hidden", borderRadius: 16, background: "#f1f5f9", boxShadow: "0 24px 70px rgba(15,23,42,.3)" }}>
    <header style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "16px 22px", borderBottom: "1px solid #e2e8f0", background: "#fff" }}><strong id="report-preview-title" style={{ fontSize: 17 }}>Попередній перегляд та редагування звіту перед завантаженням</strong><button type="button" aria-label="Закрити" onClick={onClose} style={{ width: 34, height: 34, border: "1px solid #cbd5e1", borderRadius: 8, background: "#fff", cursor: "pointer" }}>✕</button></header>
    <div style={{ overflow: "auto", padding: 22 }}><article style={{ maxWidth: 1250, margin: "0 auto", padding: 32, background: "#fff", boxShadow: "0 2px 10px #0f172a12", color: "#0f172a" }}>
      <div style={{ marginBottom: 18, textAlign: "center" }}><p style={{ margin: "0 0 8px", color: "#334155", fontWeight: 700 }}>Національний університет біоресурсів і природокористування України</p><p style={{ margin: "3px 0" }}>{report.organization.facultyName}</p><p style={{ margin: "3px 0" }}>{report.organization.departmentName}</p>
        <input aria-label="Заголовок звіту" style={{ ...inputStyle, height: 40, maxWidth: 850, marginTop: 12, textAlign: "center", fontWeight: 700, fontSize: 17 }} value={title} onChange={(event) => setTitle(event.target.value)} />
        <input aria-label="Підзаголовок звіту" style={{ ...inputStyle, maxWidth: 850, marginTop: 8, textAlign: "center" }} value={subtitle} onChange={(event) => setSubtitle(event.target.value)} />
      </div>
      {report.kind === "disciplines" ? disciplinesTable() : report.kind === "form56" ? form56Table() : report.kind === "plan" ? planTable(rows) : rowTable(rows)}
      <label style={{ display: "block", marginTop: 18, fontSize: 12, fontWeight: 700 }}>Коментар під таблицею<textarea value={comment} onChange={(event) => setComment(event.target.value)} rows={2} style={{ ...inputStyle, height: "auto", padding: 8, marginTop: 5, resize: "vertical" }} /></label>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0,1fr))", gap: 12, marginTop: 18, paddingTop: 14, borderTop: "1px solid #cbd5e1" }}><label style={{ fontSize: 12 }}>Завідувач кафедри<input value={departmentSigner} onChange={(event) => setDepartmentSigner(event.target.value)} style={{ ...inputStyle, marginTop: 5 }} /></label>{report.kind === "teacher" && <label style={{ fontSize: 12 }}>Викладач<input value={teacherSigner} onChange={(event) => setTeacherSigner(event.target.value)} style={{ ...inputStyle, marginTop: 5 }} /></label>}<label style={{ fontSize: 12 }}>Дата<input value={reportDate} onChange={(event) => setReportDate(event.target.value)} style={{ ...inputStyle, marginTop: 5 }} /></label></div>
    </article></div>
    <footer style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "14px 22px", borderTop: "1px solid #e2e8f0", background: "#fff" }}><button type="button" onClick={restore} style={{ padding: "9px 14px", border: "1px solid #cbd5e1", borderRadius: 8, background: "#fff", cursor: "pointer" }}>Скинути зміни</button><div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}><button disabled={busy} type="button" onClick={() => void saveFile("docx")} className="btn-import-workload">Завантажити Word</button><button disabled={busy} type="button" onClick={() => void saveFile("xlsx")} className="btn-import-workload">Завантажити Excel</button><button disabled={busy} type="button" onClick={() => void saveFile("pdf")} className="btn-import-workload">Завантажити PDF</button></div></footer>
  </section></div>;
}

const workloadExportTotal = (row: ExportWorkload) => {
  const storedTotal = number(row.total_hours);
  return storedTotal !== 0 ? storedTotal
    : [row.lectures, row.labs, row.practices, row.seminars, row.exams, row.credits, row.course_projects, row.practice_hours, row.diploma_project, row.reviews].reduce<number>((sum, item) => sum + number(item), 0);
};

function Form56Tab({ teachers, records, year, selectedForm, departmentName, facultyName, departmentHeadName, fullscreen, setFullscreen }: { teachers: Teacher[]; records: WorkloadRecord[]; year: string; selectedForm: FormSelection; departmentName: string; facultyName: string; departmentHeadName: string; fullscreen: boolean; setFullscreen: (value: boolean | ((current: boolean) => boolean)) => void }) {
  const [reportPreviewOpen, setReportPreviewOpen] = useState(false);
  const [teacherQuery, setTeacherQuery] = useState("");
  const [positionFilter, setPositionFilter] = useState("");
  const [employmentFilter, setEmploymentFilter] = useState("");
  const [loadFilter, setLoadFilter] = useState("");
  const kinds = workloadHourFields;
  const forms: StudyForm[] = ["Денна", "Заочна"];
  const positions = [...new Set(teachers.map((teacher) => trimText(teacher.position)).filter(Boolean))].sort((a, b) => a.localeCompare(b, "uk"));
  const employmentTypes = [...new Set(teachers.map((teacher) => trimText(teacher.employmentType)).filter(Boolean))].sort((a, b) => a.localeCompare(b, "uk"));
  const filteredTeachers = teachers.filter((teacher) => {
    const matchesLoad = !loadFilter || (loadFilter === "assigned" ? records.some((row) => row.teacher_id === teacher.id) : !records.some((row) => row.teacher_id === teacher.id));
    return (!teacherQuery || teacher.name.toLocaleLowerCase("uk-UA").includes(teacherQuery.toLocaleLowerCase("uk-UA")))
      && (!positionFilter || trimText(teacher.position) === positionFilter)
      && (!employmentFilter || trimText(teacher.employmentType) === employmentFilter) && matchesLoad;
  });
  const tableRows = filteredTeachers.map((teacher) => {
    const dayRows = records.filter((row) => row.teacher_id === teacher.id && row.study_form === "Денна");
    const extramuralRows = records.filter((row) => row.teacher_id === teacher.id && row.study_form === "Заочна");
    const aggregate = (form: StudyForm, source: WorkloadRecord[]): ExportWorkload => ({
      teacher_id: teacher.id, discipline_name: teacher.name, groups: form.toLocaleLowerCase("uk-UA"), study_form: form,
      position: teacher.position, rate: teacher.rate,
      lectures: source.reduce((sum, row) => sum + number(row.lectures), 0), labs: source.reduce((sum, row) => sum + number(row.labs), 0),
      practices: source.reduce((sum, row) => sum + number(row.practices), 0), seminars: source.reduce((sum, row) => sum + number(row.seminars), 0),
      exams: source.reduce((sum, row) => sum + number(row.exams), 0), credits: source.reduce((sum, row) => sum + number(row.credits), 0),
      course_projects: source.reduce((sum, row) => sum + number(row.course_projects), 0), practice_hours: source.reduce((sum, row) => sum + number(row.practice_hours), 0),
      diploma_project: source.reduce((sum, row) => sum + number(row.diploma_project), 0), reviews: source.reduce((sum, row) => sum + number(row.reviews), 0),
      total_hours: source.reduce((sum, row) => sum + workloadTotal(row), 0),
    });
    const day = aggregate("Денна", dayRows);
    const extramural = aggregate("Заочна", extramuralRows);
    const overall = { ...aggregate("Денна", [...dayRows, ...extramuralRows]), study_form: "Всього" as StudyForm };
    return { teacher, day, extramural, overall };
  });
  const assignedRecords = records.filter((row) => row.teacher_id > 0);
  const kindHeadings: Record<WorkloadHourKey, string> = { lectures: "Лек.", labs: "Лаб.", practices: "Практ.", seminars: "Семін.", exams: "Екз.", credits: "Залік", course_projects: "Курс.", practice_hours: "Практ-ка", diploma_project: "ДП / ЕК", reviews: "Рец." };
  const title = `Форма 56 — виконання навчальної роботи`;
  const exportRows = tableRows.flatMap(({ day, extramural }) => selectedForm === "day" ? [day] : selectedForm === "extramural" ? [extramural] : [day, extramural]);
  const visibleRows = tableRows.map(({ teacher, day, extramural, overall }) => ({ teacher, rows: selectedForm === "day" ? [day] : selectedForm === "extramural" ? [extramural] : [day, extramural, overall] }));
  const summaryRows = (row: ExportWorkload) => kinds.map((kind) => number(row[kind.key]));
  const formSummary = (form: StudyForm) => exportRows.filter((row) => row.study_form === form);
  const totalOf = (row: ExportWorkload) => summaryRows(row).reduce((sum, value) => sum + value, 0);
  const cellStyle: React.CSSProperties = { boxSizing: "border-box", padding: "12px 6px", borderBottom: "1px solid #e2e8f0", textAlign: "right", whiteSpace: "nowrap", fontSize: 14, fontVariantNumeric: "tabular-nums", color: "#0f172a", fontWeight: 500 };
  const headStyle: React.CSSProperties = { boxSizing: "border-box", height: 44, padding: "10px 6px", whiteSpace: "nowrap", fontSize: 12.5, fontWeight: 700, color: "#334155", background: "#f1f5f9", borderBottom: "1px solid #e2e8f0" };
  const tableHead = <thead><tr><th style={{ ...headStyle, width: 36, textAlign: "center" }}>№</th><th style={{ ...headStyle, minWidth: 150, textAlign: "left" }}>Викладач</th><th style={{ ...headStyle, width: 120, textAlign: "left" }}>Посада</th><th style={{ ...headStyle, width: 60, textAlign: "center" }}>Ставка</th><th style={{ ...headStyle, width: 84, textAlign: "left" }}>Форма</th>{kinds.map((kind) => <th key={kind.key} style={{ ...headStyle, width: 44, minWidth: 44, textAlign: "right" }}>{kindHeadings[kind.key]}</th>)}<th style={{ ...headStyle, width: 78, textAlign: "right", fontWeight: 800 }}>Разом, год</th></tr></thead>;
  return <Card className={`workload-report-card ${fullscreen ? "workload-fullscreen-card" : ""}`} style={{ width: "100%", boxSizing: "border-box", padding: 0 }}>
    {fullscreen && <FullscreenTabBar title="Зведення кафедри (Форма 56)" setFullscreen={setFullscreen} action={<DownloadReportButton onClick={() => setReportPreviewOpen(true)} />} />}
     {!fullscreen && <div className="workload-card-header"><div className="workload-header-title-line"><h2>{title}</h2><span>Показано <strong>{filteredTeachers.length}</strong> із <strong>{teachers.length}</strong> викладачів · {year} · {selectedForm === "all" ? "Усі форми" : selectedForm === "day" ? "Денна" : "Заочна"}</span></div><div className="workload-card-header-actions"><DownloadReportButton onClick={() => setReportPreviewOpen(true)} /><FullscreenToggleButton onClick={() => setFullscreen(true)} /></div></div>}
     <div style={filterBarStyle}>
       <FilterSearch ariaLabel="Пошук викладача" placeholder="Пошук викладача" value={teacherQuery} onChange={setTeacherQuery} />
       <FilterSelect ariaLabel="Фільтр посади" value={positionFilter} onChange={setPositionFilter} options={[{ value: "", label: "Усі посади" }, ...positions.map((item) => ({ value: item, label: item }))]} />
       <FilterSelect ariaLabel="Фільтр типу зайнятості" value={employmentFilter} onChange={setEmploymentFilter} options={[{ value: "", label: "Усі типи зайнятості" }, ...employmentTypes.map((item) => ({ value: item, label: item }))]} />
       <FilterSelect ariaLabel="Фільтр наявності навантаження" value={loadFilter} onChange={setLoadFilter} options={[{ value: "", label: "Будь-яке навантаження" }, { value: "assigned", label: "Є розподіл" }, { value: "vacant", label: "Без розподілу" }]} />
       {(teacherQuery || positionFilter || employmentFilter || loadFilter) && <FilterReset onClick={() => { setTeacherQuery(""); setPositionFilter(""); setEmploymentFilter(""); setLoadFilter(""); }} />}
     </div>
    <div className="workload-tab-table-scroll workload-form56-scroll"><table style={{ width: "100%", minWidth: 0, tableLayout: "auto", borderCollapse: "collapse", fontSize: 14, color: "#0f172a", fontVariantNumeric: "tabular-nums" }}>{tableHead}<tbody>{visibleRows.map(({ teacher, rows }, index) => <React.Fragment key={teacher.id}>
      {rows.map((row, rowIndex) => { const isOverall = row.study_form === "Всього"; const isFirst = rowIndex === 0; const selectedBadge = row.study_form === "Денна" ? { background: "#dbeafe", color: "#1d4ed8" } : row.study_form === "Заочна" ? { background: "#f3e8ff", color: "#7e22ce" } : { background: "#e2e8f0", color: "#334155" }; const sharedCell = selectedForm === "all"; return <tr key={`${teacher.id}-${row.study_form}`} style={isOverall ? { background: "#f8fafc", fontWeight: 600, borderBottom: "2px solid #cbd5e1" } : undefined}>
        {isFirst && <><td rowSpan={sharedCell ? 3 : 1} style={{ ...cellStyle, width: 36, textAlign: "center", verticalAlign: "middle" }}>{index + 1}</td><td rowSpan={sharedCell ? 3 : 1} style={{ ...cellStyle, width: 150, minWidth: 150, textAlign: "left", verticalAlign: "middle", whiteSpace: "normal", overflowWrap: "anywhere", fontWeight: 600 }}>{teacher.name}</td><td rowSpan={sharedCell ? 3 : 1} style={{ width: 120, minWidth: 120, ...cellStyle, textAlign: "left", verticalAlign: "middle", whiteSpace: "normal" }}>{teacher.position || "—"}</td><td rowSpan={sharedCell ? 3 : 1} style={{ ...cellStyle, width: 60, textAlign: "center", verticalAlign: "middle" }}>{teacher.rate}</td></>}
        <td style={{ ...cellStyle, textAlign: "left", fontWeight: 600, background: isOverall ? "#f8fafc" : undefined }}><span style={{ display: "inline-flex", padding: "2px 7px", borderRadius: 999, background: selectedBadge.background, color: selectedBadge.color }}>{row.study_form === "Всього" ? "Всього" : row.study_form}</span></td>{summaryRows(row).map((value, valueIndex) => <td key={kinds[valueIndex].key} className={value === 0 ? "workload-hour-zero" : "workload-hour-value"} style={{ ...cellStyle, background: isOverall ? "#f8fafc" : undefined, fontWeight: isOverall ? 600 : undefined }}>{hours(value)}</td>)}<td style={{ ...cellStyle, background: isOverall ? "#f8fafc" : undefined, fontWeight: isOverall ? 800 : 700 }}>{hours(totalOf(row))}</td>
      </tr>; })}
    </React.Fragment>)}</tbody><tfoot>
      {(selectedForm === "all" ? forms : [selectedForm === "day" ? "Денна" : "Заочна"] as StudyForm[]).map((form) => { const formRows = formSummary(form); const label = form === "Денна" ? "Всього за денною формою" : "Всього за заочною формою"; return <tr key={form} style={{ background: "#eef2f6", fontWeight: 700, borderTop: "2px solid #94a3b8" }}><td colSpan={5} style={{ ...cellStyle, textAlign: "left", fontWeight: 700 }}>{label}</td>{kinds.map((kind) => { const value = formRows.reduce((sum, row) => sum + number(row[kind.key]), 0); return <td key={kind.key} className={value === 0 ? "workload-hour-zero" : "workload-hour-value"} style={cellStyle}>{hours(value)}</td>; })}<td style={{ ...cellStyle, fontWeight: 800 }}>{hours(formRows.reduce((sum, row) => sum + totalOf(row), 0))}</td></tr>; })}
      {selectedForm === "all" && <tr style={{ background: "#eef2f6", fontWeight: 800, borderTop: "2px solid #94a3b8" }}><td colSpan={5} style={{ ...cellStyle, textAlign: "left", fontWeight: 800 }}>Загалом по кафедрі (Денна + Заочна)</td>{kinds.map((kind) => { const value = exportRows.reduce((sum, row) => sum + number(row[kind.key]), 0); return <td key={kind.key} style={{ ...cellStyle, fontWeight: 700 }}>{hours(value)}</td>; })}<td style={{ ...cellStyle, fontWeight: 800 }}>{hours(exportRows.reduce((sum, row) => sum + totalOf(row), 0))}</td></tr>}
    </tfoot></table>
    {!assignedRecords.length && <div style={{ marginTop: 20, padding: "24px 16px", border: "1px solid #e2e8f0", borderRadius: 10, background: "#f8fafc", color: "#64748b", textAlign: "center", fontSize: 13 }}>За цей рік і кафедру немає розподіленого навантаження.</div>}
    {reportPreviewOpen && <ReportPreviewModal report={{ title: `Звіт кафедри «${departmentName}» про виконання навчальної роботи (Форма №56) за ${year} навчальний рік`, subtitle: `Зведений звіт по всіх викладачах за ${year} навчальний рік`, rows: exportRows, fileName: `Форма_56_${year.replace("/", "-")}`, organization: { facultyName, departmentName, departmentHeadName }, kind: "form56", form56Selection: selectedForm }} onClose={() => setReportPreviewOpen(false)} />}</div>
  </Card>;
}
function VacanciesTab(props: { year: string; studyForm: StudyForm; rows: PlanImportRow[]; loading: boolean; assigningId: number | null; onAssign: (row: PlanImportRow) => void; organization: EditableReport["organization"]; fullscreen: boolean; setFullscreen: (value: boolean | ((current: boolean) => boolean)) => void }) {
  const [reportPreviewOpen, setReportPreviewOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [specialty, setSpecialty] = useState("");
  const [semester, setSemester] = useState("");
  const abbreviations = [...new Set(props.rows.flatMap((row) => row.abbreviations))];
  const semesters = [...new Set(props.rows.map((row) => row.semester).filter(Boolean))].sort((a, b) => a - b);
  const filteredRows = props.rows.filter((row) => (!query || `${row.discipline} ${row.target} ${row.abbreviations.join(" ")}`.toLocaleLowerCase("uk-UA").includes(query.toLocaleLowerCase("uk-UA")))
    && (!specialty || row.abbreviations.includes(specialty) || row.target === specialty) && (!semester || String(row.semester) === semester));
  const reportRows = filteredRows.map((row) => planRowForExport(row));
  return <Card className={`workload-report-card ${props.fullscreen ? "workload-fullscreen-card" : ""}`} style={{ width: "100%", boxSizing: "border-box", padding: 0 }}>
    {props.fullscreen && <FullscreenTabBar title="Вакантні години" setFullscreen={props.setFullscreen} action={<DownloadReportButton onClick={() => setReportPreviewOpen(true)} />} />}
     {!props.fullscreen && <div className="workload-card-header"><div className="workload-header-title-line"><h2>Вакантні години</h2><span>Показано <strong>{filteredRows.length}</strong> із <strong>{props.rows.length}</strong> рядків · {props.year} · {props.studyForm}</span></div><div className="workload-card-header-actions"><DownloadReportButton onClick={() => setReportPreviewOpen(true)} /><FullscreenToggleButton onClick={() => props.setFullscreen(true)} /></div></div>}
     <div style={filterBarStyle}><FilterSearch ariaLabel="Пошук вакантної дисципліни" placeholder="Пошук дисципліни" value={query} onChange={setQuery} /><FilterSelect ariaLabel="Фільтр спеціальності / ОПП" value={specialty} onChange={setSpecialty} options={[{ value: "", label: "Усі спеціальності / ОПП" }, ...abbreviations.map((item) => ({ value: item, label: item }))]} /><FilterSelect ariaLabel="Фільтр семестру" value={semester} onChange={setSemester} options={[{ value: "", label: "Усі семестри" }, ...semesters.map((item) => ({ value: String(item), label: `Семестр ${item}` }))]} />{(query || specialty || semester) && <FilterReset onClick={() => { setQuery(""); setSpecialty(""); setSemester(""); }} />}</div>
     <div className="workload-tab-table-scroll rounded-lg border border-slate-200"><table className="min-w-[850px] w-full border-collapse text-left text-sm"><thead className="sticky top-0 z-[2] bg-slate-50 text-xs uppercase text-slate-500"><tr><th className="px-4 py-3">Дисципліна</th><th className="px-4 py-3">Спеціальність / ОПП</th><th className="px-4 py-3">Семестр</th><th className="px-4 py-3">Структура годин</th><th className="px-4 py-3">Всього, год</th><th className="px-4 py-3">Дія</th></tr></thead><tbody className="divide-y divide-slate-100">{props.loading ? <EmptyRow columns={6} message="Завантаження вакантних годин…" /> : filteredRows.length ? filteredRows.map((row) => <tr key={row.id ?? row.sourceRow}><td className="px-4 py-3 font-medium">{row.discipline}</td><td className="px-4 py-3">{row.target || "—"}<span style={{ display: "inline-flex", marginLeft: 7, padding: "3px 8px", borderRadius: 999, background: row.study_form === "Заочна" ? "#f3e8ff" : "#dbeafe", color: row.study_form === "Заочна" ? "#7e22ce" : "#1d4ed8", fontSize: 11 }}>{row.study_form || "Денна"}</span></td><td className="px-4 py-3">{row.semester || "—"}</td><td className="px-4 py-3">{workloadHourFields.map(({ key, label }) => ({ key, label, value: hoursFromPlanRow(row)[key] })).filter((item) => item.value > 0).map((item) => `${item.label}: ${hours(item.value)}`).join(" · ") || "—"}</td><td className="px-4 py-3 font-bold text-amber-700">{hours(row.total)}</td><td className="px-4 py-3"><button type="button" disabled={props.assigningId === row.id} className="workload-vacancy-assign-button rounded-lg bg-blue-50 px-3 py-2 text-xs font-bold text-blue-700 hover:bg-blue-100 disabled:opacity-50" onClick={() => props.onAssign(row)}>{props.assigningId === row.id ? "Збереження…" : "Перейти до розподілу"}</button></td></tr>) : <EmptyRow columns={6} message="Нерозподілених годин за обраними фільтрами немає." />}</tbody></table></div>
    {reportPreviewOpen && <ReportPreviewModal report={{ title: `Вакантні години · ${props.year}`, subtitle: props.studyForm, rows: reportRows, fileName: `Вакантні_години_${props.year.replace("/", "-")}`, organization: props.organization, kind: "plan", hideAssignedTeacher: true }} onClose={() => setReportPreviewOpen(false)} />}
  </Card>;
}

function DisciplinesTab(props: { disciplines: Discipline[]; teachers: Teacher[]; planRows: PlanImportRow[]; canEdit: boolean; departmentCode: string; year: string; formSelection: string; forms: ImportOption[]; levels: ImportOption[]; abbreviations: ImportOption[]; organization: EditableReport["organization"]; onChanged: () => void; fullscreen: boolean; setFullscreen: (value: boolean | ((current: boolean) => boolean)) => void }) {
  type DialogMode = "add" | "edit" | "archive" | null;
  type ArchivedDisciplineEntry = { archive: Record<string, unknown>; code: string; name: string; archivedAt: string; workloadRows: Record<string, unknown>[] };
  const [dialog, setDialog] = useState<DialogMode>(null);
  const [reportPreviewOpen, setReportPreviewOpen] = useState(false);
  const [showArchived, setShowArchived] = useState(false);
  const [archivedDisciplines, setArchivedDisciplines] = useState<ArchivedDisciplineEntry[]>([]);
  const [loadingArchive, setLoadingArchive] = useState(false);
  const [archiveRefresh, setArchiveRefresh] = useState(0);
  const [disciplineSearch, setDisciplineSearch] = useState("");
  const [specialtyFilter, setSpecialtyFilter] = useState("");
  const [semesterFilter, setSemesterFilter] = useState("");
  const [controlFilter, setControlFilter] = useState("");
  const [assignmentFilter, setAssignmentFilter] = useState("");
  const [disciplineCode, setDisciplineCode] = useState("");
  const [planRowId, setPlanRowId] = useState("");
  const [saving, setSaving] = useState(false);
  const [draft, setDraft] = useState<WorkloadHourDraft>(emptyWorkloadHourDraft());
  const [form, setForm] = useState({ name: "", abbreviation: "", semester: "1", formCode: "", levelCode: "" });
  const inputStyle: React.CSSProperties = { width: "100%", height: 38, padding: "0 10px", border: "1px solid #cbd5e1", borderRadius: 8, background: "#fff", color: "#0f172a" };
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      setLoadingArchive(true);
      const { data, error } = await supabase.from("Архів").select("*").eq("Таблиця", "Дисципліни").order("Дата_архівації", { ascending: false });
      if (cancelled) return;
      if (error) {
        notify(`Не вдалося завантажити архів дисциплін: ${error.message}`, "error", "Дисципліни");
        setArchivedDisciplines([]);
      } else {
        const department = trimText(props.departmentCode);
        const entries = ((data ?? []) as unknown as Record<string, unknown>[]).flatMap((archive): ArchivedDisciplineEntry[] => {
          let stored = archive["Дані"];
          if (typeof stored === "string") { try { stored = JSON.parse(stored); } catch { stored = {}; } }
          const saved = stored && typeof stored === "object" ? stored as Record<string, unknown> : {};
          const rawDiscipline = saved.discipline && typeof saved.discipline === "object" ? saved.discipline as Record<string, unknown> : {};
          // Older archive entries stored rows from the real "Навантаження"
          // table under planRows; prefer that for backward compatibility.
          const allWorkload = Array.isArray(saved.workloadRows) ? saved.workloadRows as Record<string, unknown>[] : Array.isArray(saved.planRows) ? saved.planRows as Record<string, unknown>[] : Array.isArray(saved.workload) ? saved.workload as Record<string, unknown>[] : [];
          const workloadRows = allWorkload.filter((row) => trimText(row["Код_кафедри"]) === department);
          const archivedDepartment = trimText(rawDiscipline["Код_кафедри"] ?? rawDiscipline.departmentCode ?? allWorkload[0]?.["Код_кафедри"]);
          if (!department || archivedDepartment !== department) return [];
          const code = trimText(rawDiscipline["Код_дисципліни"] ?? rawDiscipline.group_code ?? archive["Ключ_запису"]);
          if (!code) return [];
          return [{ archive, code, name: trimText(rawDiscipline["Назва_дисципліни"] ?? rawDiscipline.name ?? archive["Назва"]) || "Без назви", archivedAt: trimText(archive["Дата_архівації"]), workloadRows }];
        });
        setArchivedDisciplines(entries);
      }
      setLoadingArchive(false);
    })();
    return () => { cancelled = true; };
  }, [props.departmentCode, archiveRefresh]);
  const selectedDiscipline = props.disciplines.find((item) => trimText(item.group_code) === disciplineCode);
  const relatedRows = props.planRows.filter((row) => trimText(row.disciplineCode) === disciplineCode && row.id);
  const selectedPlanRow = relatedRows.find((row) => String(row.id) === planRowId) ?? relatedRows[0];
  const defaultFormCode = () => {
    const dayCode = props.forms.find((item) => !/заоч/i.test(item.name))?.code ?? "";
    const extramuralCode = props.forms.find((item) => /заоч/i.test(item.name))?.code ?? "";
    return props.formSelection === "extramural" ? extramuralCode : props.formSelection === "all" ? dayCode || props.forms[0]?.code || "" : dayCode || props.forms[0]?.code || "";
  };
  const openDialog = (mode: Exclude<DialogMode, null>) => {
    setDialog(mode);
    setDisciplineCode(mode === "add" ? "" : trimText(props.disciplines[0]?.group_code));
    setPlanRowId("");
    if (mode === "add") {
      setForm({ name: "", abbreviation: "", semester: "1", formCode: defaultFormCode(), levelCode: props.levels.find((item) => /бакалавр/i.test(item.name))?.code ?? props.levels[0]?.code ?? "" });
      setDraft(emptyWorkloadHourDraft());
    }
  };
  useEffect(() => {
    if (!dialog || dialog === "add" || !selectedDiscipline) return;
    const row = relatedRows[0];
    setPlanRowId(row?.id ? String(row.id) : "");
    setForm({ name: selectedDiscipline.name, abbreviation: row?.target ?? row?.abbreviations?.[0] ?? "", semester: String(row?.semester || selectedDiscipline.semester || 1), formCode: row?.form_code ?? defaultFormCode(), levelCode: row?.levelCode ?? props.levels[0]?.code ?? "" });
    setDraft(row ? hoursFromPlanRow(row) : emptyWorkloadHourDraft());
  // The selected catalog entry intentionally resets the editor fields.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dialog, disciplineCode]);

  const save = async () => {
    const name = form.name.trim();
    const department = trimText(props.departmentCode);
    if (!name || !department) { notify("Вкажіть назву дисципліни та кафедру.", "warning", "Дисципліни"); return; }
    if (!form.formCode.trim()) { notify("Оберіть форму навчання.", "warning", "Дисципліни"); return; }
    setSaving(true);
    try {
      const hourPayload = Object.fromEntries(workloadHourFields.map(({ key, column }) => [column, roundHours(Math.max(0, number(draft[key])))]));
      const total = totalForHourDraft(draft);
      const abbreviation = form.abbreviation.trim();
      const level = trimText(form.levelCode) || (isMasterLevel("", "", abbreviation) ? props.levels.find((item) => /магістр/i.test(item.name))?.code : props.levels.find((item) => /бакалавр/i.test(item.name))?.code) || "";
      const workloadPayload: Record<string, unknown> = {
        "Рік": toDbYear(props.year), "Код_кафедри": department, "Код_форми_навчання": form.formCode.trim(), "Код_рівня_навчання": trimText(level),
        "Код_дисципліни": disciplineCode, "Код_викладача": dialog === "edit" ? trimText(selectedPlanRow?.teacherCode) || null : null, "Абревіатура": abbreviation, "Семестр": Math.max(1, Math.round(number(form.semester))),
        ...hourPayload, "Всього_годин": total,
      };
      if (dialog === "add") {
        const { data: codeRows, error: codeError } = await supabase.from("Дисципліни").select("Код_дисципліни");
        if (codeError) throw codeError;
        const used = new Set(((codeRows ?? []) as unknown as Record<string, unknown>[]).map((row) => trimText(row["Код_дисципліни"]).toLocaleUpperCase("uk-UA")));
        let nextCode = "";
        for (let i = 1; i <= 36 * 36; i += 1) {
          const candidate = `P${i.toString(36).toUpperCase().padStart(2, "0")}`;
          if (!used.has(candidate)) { nextCode = candidate; break; }
        }
        if (!nextCode) throw new Error("Не вдалося згенерувати вільний код дисципліни.");
        const duplicate = props.disciplines.some((item) => normalizedName(item.name) === normalizedName(name));
        if (duplicate) throw new Error("Дисципліна з такою назвою вже є в активному списку кафедри.");
        const { error: disciplineError } = await supabase.from("Дисципліни").insert({ "Код_дисципліни": nextCode, "Код_кафедри": department, "Назва_дисципліни": name });
        if (disciplineError) throw disciplineError;
        workloadPayload["Код_дисципліни"] = nextCode;
        let insertResult = await supabase.from("Завантаження_навантаження").insert(workloadPayload);
        if (hasMissingReviewsColumn(insertResult.error)) {
          const noReviews: Record<string, unknown> = { ...workloadPayload }; delete noReviews["Рецензії"];
          insertResult = await supabase.from("Завантаження_навантаження").insert(noReviews);
          if (!insertResult.error) notify("У базі немає колонки «Рецензії»; записано інші види годин.", "warning", "Дисципліни");
        }
        if (insertResult.error) {
          await supabase.from("Дисципліни").delete().eq("Код_дисципліни", nextCode).eq("Код_кафедри", department);
          throw insertResult.error;
        }
        notify("Дисципліну додано до плану.", "success", "Дисципліни");
      } else if (dialog === "edit") {
        if (!disciplineCode || !selectedDiscipline) throw new Error("Оберіть дисципліну для редагування.");
        const { error: nameError } = await supabase.from("Дисципліни").update({ "Назва_дисципліни": name }).eq("Код_дисципліни", disciplineCode).eq("Код_кафедри", department);
        if (nameError) throw nameError;
        if (selectedPlanRow?.id) {
          const updatePayload = { ...workloadPayload, "Код_дисципліни": disciplineCode };
          let updateResult = await supabase.from("Завантаження_навантаження").update(updatePayload).eq("id", selectedPlanRow.id);
          if (hasMissingReviewsColumn(updateResult.error)) {
            const noReviews: Record<string, unknown> = { ...updatePayload }; delete noReviews["Рецензії"];
            updateResult = await supabase.from("Завантаження_навантаження").update(noReviews).eq("id", selectedPlanRow.id);
            if (!updateResult.error) notify("У базі немає колонки «Рецензії»; записано інші види годин.", "warning", "Дисципліни");
          }
          if (updateResult.error) throw updateResult.error;
        }
        notify("Зміни дисципліни збережено.", "success", "Дисципліни");
      }
      setDialog(null);
      props.onChanged();
    } catch (error) {
      notify(error instanceof Error ? error.message : "Не вдалося зберегти дисципліну.", "error", "Дисципліни");
    } finally { setSaving(false); }
  };

  const archiveSelected = async () => {
    const code = trimText(disciplineCode);
    const selected = props.disciplines.find((item) => trimText(item.group_code) === code);
    if (!code || !selected) { notify("Оберіть дисципліну для архівації.", "warning", "Дисципліни"); return; }
    if (!window.confirm(`Перемістити «${selected.name}» та пов’язані рядки навантаження в архів?`)) return;
    setSaving(true);
    try {
      const department = trimText(props.departmentCode);
      const { data: disciplineCandidates, error: disciplineReadError } = await supabase.from("Дисципліни").select("*").eq("Код_кафедри", department).ilike("Код_дисципліни", `${code}%`).limit(10);
      if (disciplineReadError) throw disciplineReadError;
      const disciplineRow = (disciplineCandidates ?? []).find((row) => trimText((row as Record<string, unknown>)["Код_дисципліни"]) === code);
      if (!disciplineRow) throw new Error("Не вдалося знайти повний запис дисципліни в базі даних.");
      const disciplineData = disciplineRow as unknown as Record<string, unknown>;
      const exactCode = String(disciplineData["Код_дисципліни"] ?? code);
      // The workload plan and assignments are stored in this table. Keep the
      // untrimmed bpchar key from the discipline row for the database lookup.
      const { data: workloadRows, error: workloadReadError } = await supabase
        .from("Завантаження_навантаження")
        .select("*")
        .eq("Код_дисципліни", exactCode);
      if (workloadReadError) throw workloadReadError;
      const archivedAt = new Date().toISOString();
      const archivePayload = {
        "Таблиця": "Дисципліни",
        "Ключ_запису": trimText(disciplineData["Код_дисципліни"]),
        "Дані": {
          discipline: disciplineData,
          workloadRows: workloadRows ?? [],
        },
        "Назва": trimText(disciplineData["Назва_дисципліни"]),
        "Дата_архівації": archivedAt,
      };
      const { data: savedArchive, error: archiveError } = await supabase.from("Архів").insert(archivePayload).select("id").single();
      if (archiveError) throw archiveError;
      if ((workloadRows ?? []).length) {
        const { error: workloadDeleteError } = await supabase.from("Завантаження_навантаження").delete().eq("Код_дисципліни", exactCode);
        if (workloadDeleteError) {
          if (savedArchive?.id) await supabase.from("Архів").delete().eq("id", savedArchive.id);
          throw workloadDeleteError;
        }
      }
      const { error: disciplineDeleteError } = await supabase.from("Дисципліни").delete().eq("Код_дисципліни", exactCode).eq("Код_кафедри", department);
      if (disciplineDeleteError) throw disciplineDeleteError;
      setArchivedDisciplines((current) => [{ archive: { ...archivePayload, id: savedArchive?.id }, code, name: selected.name, archivedAt, workloadRows: (workloadRows ?? []) as Record<string, unknown>[] }, ...current.filter((item) => item.code !== code)]);
      setArchiveRefresh((version) => version + 1);
      notify("Дисципліну переміщено в архів", "success", "Дисципліни");
      setDialog(null);
      props.onChanged();
    } catch (error) {
      const dbError = error as { message?: string; details?: string; hint?: string; code?: string } | null;
      const reason = dbError?.message || dbError?.details || (error == null ? "Невідома помилка Supabase" : JSON.stringify(error));
      console.error("Не вдалося архівувати дисципліну:", error);
      notify(`Не вдалося архівувати дисципліну: ${reason}`, "error", "Дисципліни");
    } finally { setSaving(false); }
  };

  const restoreArchivedDiscipline = async (entry: ArchivedDisciplineEntry) => {
    const department = trimText(props.departmentCode);
    let archiveData: unknown = entry.archive["Дані"];
    if (typeof archiveData === "string") { try { archiveData = JSON.parse(archiveData); } catch { archiveData = {}; } }
    const stored = archiveData && typeof archiveData === "object" ? archiveData as Record<string, unknown> : {};
    const rawDiscipline = stored.discipline && typeof stored.discipline === "object" ? stored.discipline as Record<string, unknown> : {};
    // planRows was used by an earlier archive format; keep accepting it so
    // archived records made before this fix remain restorable.
    const rawWorkload = Array.isArray(stored.workloadRows) ? stored.workloadRows as Record<string, unknown>[] : Array.isArray(stored.planRows) ? stored.planRows as Record<string, unknown>[] : Array.isArray(stored.workload) ? stored.workload as Record<string, unknown>[] : entry.workloadRows;
    if (!department || !entry.code) { notify("Не вдалося визначити кафедру або код дисципліни для відновлення.", "error", "Дисципліни"); return; }
    setSaving(true);
    try {
      const disciplinePayload = {
        "Код_дисципліни": trimText(rawDiscipline["Код_дисципліни"] ?? rawDiscipline.group_code ?? entry.code),
        "Код_кафедри": trimText(rawDiscipline["Код_кафедри"] ?? rawDiscipline.departmentCode ?? department),
        "Назва_дисципліни": trimText(rawDiscipline["Назва_дисципліни"] ?? rawDiscipline.name ?? entry.name),
      };
      const { error: disciplineError } = await supabase.from("Дисципліни").upsert(disciplinePayload, { onConflict: "Код_дисципліни" });
      if (disciplineError) throw disciplineError;

      const restoreRows = async (rows: Record<string, unknown>[]) => {
        const matchingRows = rows.filter((row) => trimText(row["Код_дисципліни"]) === entry.code);
        const rowsWithId = matchingRows.filter((row) => row.id != null);
        const rowsWithoutId = matchingRows.filter((row) => row.id == null).map((row) => { const copy = { ...row }; delete copy.id; return copy; });
        if (rowsWithId.length) {
          const { error } = await supabase.from("Завантаження_навантаження").upsert(rowsWithId, { onConflict: "id" });
          if (error) throw error;
        }
        if (rowsWithoutId.length) {
          const { error } = await supabase.from("Завантаження_навантаження").insert(rowsWithoutId);
          if (error) throw error;
        }
      };
      await restoreRows(rawWorkload);

      if (!entry.archive.id) throw new Error("В архівному записі відсутній UUID id; відновлення виконане, але запис архіву не видалено.");
      const { data: deletedArchive, error: archiveDeleteError } = await supabase.from("Архів").delete().eq("id", entry.archive.id).select("id").maybeSingle();
      if (archiveDeleteError) throw archiveDeleteError;
      if (!deletedArchive) throw new Error("Дисципліну відновлено, але Supabase не підтвердив видалення запису з архіву.");

      setArchivedDisciplines((current) => current.filter((item) => item !== entry));
      setArchiveRefresh((version) => version + 1);
      setShowArchived(false);
      notify("Дисципліну відновлено з архіву", "success", "Дисципліни");
      props.onChanged();
    } catch (error) {
      notify(error instanceof Error ? error.message : "Не вдалося відновити дисципліну з архіву.", "error", "Дисципліни");
    } finally { setSaving(false); }
  };

  const allReportRows: DisciplineReportRow[] = buildDisciplineReportRows(props.disciplines, props.planRows, props.teachers);
  const specialtyOptions = [...new Set(allReportRows.map((row) => row.specialty).filter(Boolean))].sort((a, b) => a.localeCompare(b, "uk"));
  const semesterOptions = [...new Set(allReportRows.map((row) => row.semester).filter(Boolean))].sort((a, b) => a.localeCompare(b, "uk", { numeric: true }));
  const controlOptions = [...new Set(allReportRows.map((row) => row.control).filter((item) => item && item !== "—"))];
  const reportRows = allReportRows.filter((row) => {
    const teacherNames = row.teachers.map((teacher) => teacher.name).join(" ");
    const hasTeachers = row.teachers.some((teacher) => teacher.name !== "Не призначено");
    const searchText = `${row.discipline} ${row.specialty} ${teacherNames}`.toLocaleLowerCase("uk-UA");
    return (!disciplineSearch || searchText.includes(disciplineSearch.toLocaleLowerCase("uk-UA")))
      && (!specialtyFilter || row.specialty === specialtyFilter)
      && (!semesterFilter || row.semester === semesterFilter)
      && (!controlFilter || row.control === controlFilter)
      && (!assignmentFilter || (assignmentFilter === "assigned" ? hasTeachers : !hasTeachers));
  });
  const reportTitle = `Дисципліни кафедри · ${toDisplayYear(props.year)}`;
  const reportSubtitle = `${props.formSelection === "all" ? "Усі форми" : props.formSelection === "extramural" ? "Заочна" : "Денна"} · ${reportRows.length} із ${allReportRows.length} записів`;
  const fileName = `dyscypliny-${toDbYear(props.year)}`;
  const actionButtons = <div className="workload-card-header-actions"><button type="button" className="workload-discipline-neutral-button" aria-pressed={!showArchived} onClick={() => setShowArchived(false)}>Активні ({props.disciplines.length})</button><button type="button" className="workload-discipline-neutral-button" aria-pressed={showArchived} onClick={() => setShowArchived(true)}>Архів ({archivedDisciplines.length})</button><DownloadReportButton onClick={() => setReportPreviewOpen(true)} />{props.canEdit && !showArchived && <><button type="button" className="workload-discipline-add-button" onClick={() => openDialog("add")}>+ Додати дисципліну</button><button type="button" className="workload-discipline-neutral-button" onClick={() => openDialog("edit")}>Редагувати</button><button type="button" className="workload-discipline-archive-button" onClick={() => openDialog("archive")}>В архів</button></>}<FullscreenToggleButton onClick={() => props.setFullscreen(true)} /></div>;
  return <Card className={`workload-report-card ${props.fullscreen ? "workload-fullscreen-card" : ""}`} style={{ width: "100%", boxSizing: "border-box", padding: 0 }}>{props.fullscreen && <FullscreenTabBar title="Дисципліни" setFullscreen={props.setFullscreen} action={actionButtons} />}{!props.fullscreen && <div className="workload-card-header"><div><h2>Дисципліни</h2><span>{showArchived ? <>Архівних дисциплін: <strong>{archivedDisciplines.length}</strong></> : <>Показано <strong>{reportRows.length}</strong> із <strong>{allReportRows.length}</strong> записів</>}</span></div>{actionButtons}</div>}{!showArchived && <div style={filterBarStyle}><FilterSearch ariaLabel="Пошук дисципліни або викладача" placeholder="Дисципліна або викладач" value={disciplineSearch} onChange={setDisciplineSearch} /><FilterSelect ariaLabel="Фільтр спеціальності / ОПП" value={specialtyFilter} onChange={setSpecialtyFilter} options={[{ value: "", label: "Усі спеціальності / ОПП" }, ...specialtyOptions.map((item) => ({ value: item, label: item }))]} /><FilterSelect ariaLabel="Фільтр семестру" value={semesterFilter} onChange={setSemesterFilter} options={[{ value: "", label: "Усі семестри" }, ...semesterOptions.map((item) => ({ value: item, label: item }))]} /><FilterSelect ariaLabel="Фільтр форми контролю" value={controlFilter} onChange={setControlFilter} options={[{ value: "", label: "Усі форми контролю" }, ...controlOptions.map((item) => ({ value: item, label: item }))]} /><FilterSelect ariaLabel="Фільтр призначення викладача" value={assignmentFilter} onChange={setAssignmentFilter} options={[{ value: "", label: "Усі призначення" }, { value: "assigned", label: "Призначені" }, { value: "vacant", label: "Не призначені" }]} />{(disciplineSearch || specialtyFilter || semesterFilter || controlFilter || assignmentFilter) && <FilterReset onClick={() => { setDisciplineSearch(""); setSpecialtyFilter(""); setSemesterFilter(""); setControlFilter(""); setAssignmentFilter(""); }} />}</div>}<div className="workload-tab-table-scroll rounded-lg border border-slate-200" style={{ display: showArchived ? "none" : undefined }}><table className="w-full border-collapse text-left text-sm"><thead className="sticky top-0 z-[2] bg-slate-50 text-xs uppercase text-slate-500"><tr><th className="px-4 py-3">СПЕЦІАЛЬНІСТЬ (ОПП)</th><th className="px-4 py-3">ДИСЦИПЛІНА</th><th className="px-4 py-3">СЕМЕСТР ТА ФОРМА НАВЧАННЯ</th><th className="px-4 py-3">ФОРМА ПРОВЕДЕННЯ (КОНТРОЛЬ)</th><th className="px-4 py-3">ВИКЛАДАЧІ</th></tr></thead><tbody className="divide-y divide-slate-100">{reportRows.length ? reportRows.map((row, index) => <tr key={`${row.specialty}-${row.discipline}-${row.semester}-${row.forms.join()}-${index}`}><td className="px-4 py-3">{row.specialty || "—"}</td><td className="px-4 py-3 font-medium">{row.discipline}</td><td className="px-4 py-3"><span>{row.semester || "—"}</span><span style={{ display: "inline-flex", marginLeft: 7, padding: "3px 8px", borderRadius: 999, background: row.forms[0] === "Заочна" ? "#f3e8ff" : "#dbeafe", color: row.forms[0] === "Заочна" ? "#7e22ce" : "#1d4ed8", fontSize: 11 }}>{row.forms[0] || "—"}</span></td><td className="px-4 py-3">{row.control}</td><td className="px-4 py-3">{row.teachers.length ? <div style={{ display: "grid", gap: 5 }}>{row.teachers.map((assignment) => <div key={assignment.name}><strong>{assignment.name}</strong>{assignment.workTypes.length > 0 && <span> ({assignment.workTypes.map((kind) => kind.toLocaleLowerCase("uk-UA")).join(", ")})</span>}</div>)}</div> : <span style={{ color: "#94a3b8" }}>Не призначено</span>}</td></tr>) : <EmptyRow columns={5} message="Дисциплін за обраними фільтрами немає." />}</tbody></table></div>
    {showArchived && <div className="workload-tab-table-scroll" style={{ marginTop: 0 }}><table className="w-full border-collapse text-left text-sm"><thead className="sticky top-0 z-[2] bg-slate-50 text-xs uppercase text-slate-500"><tr><th className="px-4 py-3">Дисципліна</th><th className="px-4 py-3">Ключ_запису</th><th className="px-4 py-3">Дата архівації</th><th className="px-4 py-3">Дія</th></tr></thead><tbody className="divide-y divide-slate-100">{loadingArchive ? <EmptyRow columns={4} message="Завантаження архіву дисциплін…" /> : archivedDisciplines.length ? archivedDisciplines.map((entry) => <tr key={entry.code}><td className="px-4 py-3 font-semibold">{entry.name}</td><td className="px-4 py-3">{entry.code}</td><td className="px-4 py-3">{entry.archivedAt ? new Date(entry.archivedAt).toLocaleString("uk-UA") : "—"}</td><td className="px-4 py-3"><button type="button" disabled={saving} onClick={() => void restoreArchivedDiscipline(entry)} style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "5px 10px", border: "1px solid #bfdbfe", borderRadius: 8, background: "#eff6ff", color: "#2563eb", fontSize: 12.5, fontWeight: 600, whiteSpace: "nowrap", cursor: saving ? "wait" : "pointer" }}><FiRotateCcw size={14} />{saving ? "Відновлення…" : "Відновити з архіву"}</button></td></tr>) : <EmptyRow columns={4} message="Архівних дисциплін цієї кафедри немає." />}</tbody></table></div>}

    {dialog && <div role="presentation" className="workload-discipline-modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget && !saving) setDialog(null); }}><section role="dialog" aria-modal="true" aria-labelledby="discipline-dialog-title" className="workload-discipline-modal"><header><div><h3 id="discipline-dialog-title">{dialog === "add" ? "Додати дисципліну" : dialog === "edit" ? "Редагувати дисципліну" : "Перемістити дисципліну в архів"}</h3><p>{dialog === "add" ? `Новий запис для ${toDisplayYear(props.year)} навчального року` : "Працюємо з даними поточної кафедри та навчального року"}</p></div><button type="button" aria-label="Закрити" onClick={() => setDialog(null)} disabled={saving}>✕</button></header>
      {dialog !== "add" && <label style={{ display: "block", marginBottom: 12 }}>Дисципліна<SearchableSelect ariaLabel="Дисципліна" placeholder="Оберіть дисципліну" value={disciplineCode} onChange={setDisciplineCode} options={props.disciplines.map((item) => ({ value: item.group_code ?? "", label: item.name, sublabel: item.semesters?.length ? `Семестр ${item.semesters.join(", ")}${item.studyForms?.length ? ` · ${item.studyForms.join(", ")}` : ""}` : item.group_code ?? undefined }))} /></label>}
      {dialog === "edit" && relatedRows.length > 1 && <label style={{ display: "block", marginBottom: 12 }}>Запис плану (форма / семестр / абревіатура)<SearchableSelect ariaLabel="Запис плану" placeholder="Оберіть запис плану" value={planRowId} onChange={setPlanRowId} options={relatedRows.map((row) => ({ value: String(row.id), label: `${row.study_form} · семестр ${row.semester}`, sublabel: row.abbreviations.join(", ") || "без абревіатури" }))} /></label>}
      {dialog !== "archive" && <><div className="workload-discipline-form-grid"><label>Назва дисципліни<input style={inputStyle} value={form.name} onChange={(event) => setForm((current) => ({ ...current, name: event.target.value }))} /></label><label>Абревіатура / ОПП<select style={inputStyle} value={form.abbreviation} onChange={(event) => { const option = props.abbreviations.find((item) => item.name === event.target.value); setForm((current) => ({ ...current, abbreviation: event.target.value, levelCode: isMasterLevel("", option?.programName ?? "", option?.name ?? "") ? props.levels.find((level) => /магістр/i.test(level.name))?.code ?? current.levelCode : current.levelCode })); }}><option value="">Без абревіатури</option>{props.abbreviations.map((item) => <option key={item.code} value={item.name}>{item.name}{item.specialtyName ? ` — ${item.specialtyName}` : ""}{item.programName ? ` / ${item.programName}` : ""}</option>)}</select></label><label>Семестр<input style={inputStyle} type="number" min="1" max="12" value={form.semester} onChange={(event) => setForm((current) => ({ ...current, semester: event.target.value }))} /></label><label>Форма навчання<select style={inputStyle} value={form.formCode} onChange={(event) => setForm((current) => ({ ...current, formCode: event.target.value }))}>{props.forms.map((item) => <option key={item.code} value={item.code}>{item.name}</option>)}</select></label><label>Рівень навчання<select style={inputStyle} value={form.levelCode} onChange={(event) => setForm((current) => ({ ...current, levelCode: event.target.value }))}>{props.levels.map((item) => <option key={item.code} value={item.code}>{item.name}</option>)}</select></label></div>
        <div className="workload-discipline-hours-grid">{workloadHourFields.map(({ key, label }) => <label key={key}>{label}<input style={inputStyle} type="number" min="0" step="0.1" value={draft[key]} onChange={(event) => setDraft((current) => ({ ...current, [key]: number(event.target.value) }))} /></label>)}</div><p className="workload-discipline-total">Всього: {hours(totalForHourDraft(draft))} год</p></>}
      <footer><button type="button" className="workload-discipline-neutral-button" onClick={() => setDialog(null)} disabled={saving}>Скасувати</button>{dialog === "archive" ? <button type="button" className="workload-discipline-archive-button" onClick={() => void archiveSelected()} disabled={saving || !disciplineCode}>{saving ? "Архівація…" : "Підтвердити архівацію"}</button> : <button type="button" className="workload-discipline-add-button" onClick={() => void save()} disabled={saving}>{saving ? "Збереження…" : dialog === "add" ? "Додати дисципліну" : "Зберегти зміни"}</button>}</footer>
    </section></div>}
    {reportPreviewOpen && <ReportPreviewModal report={{ title: reportTitle, subtitle: reportSubtitle, rows: [], disciplineRows: reportRows, fileName, organization: props.organization, kind: "disciplines" }} onClose={() => setReportPreviewOpen(false)} />}
  </Card>;
}
