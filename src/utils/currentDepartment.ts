import { supabase } from "./supabaseClient";

type Row = Record<string, unknown>;
type AuthProfile = { id?: string; email?: string | null; user_metadata?: Record<string, unknown> | null } | null | undefined;

const text = (value: unknown) => String(value ?? "").trim();

/** Resolves the signed-in user's department using existing profile and department rows. */
export function resolveCurrentDepartmentCode(
  profile: AuthProfile,
  teachers: Row[],
  departmentInfo: Row[],
  departments: Row[],
): string {
  const email = text(profile?.email).toLocaleLowerCase("uk-UA");
  const teacher = email
    ? teachers.find((row) => text(row["Email_викладача"]).toLocaleLowerCase("uk-UA") === email)
    : undefined;
  const teacherCode = text(teacher?.["Код_викладача"]);
  const linkedCodes = teacherCode
    ? departmentInfo.filter((row) => text(row["Код_викладача"]) === teacherCode).map((row) => text(row["Код_кафедри"]))
    : [];
  const departmentEmailMatch = email
    ? departments.find((row) => text(row["Email_кафедри"]).toLocaleLowerCase("uk-UA") === email)
    : undefined;
  const metadata = profile?.user_metadata ?? {};
  const departmentHint = text(metadata["Код_кафедри"] ?? metadata["department_code"] ?? metadata["department"]);
  const matchedDepartment = departments.find((row) =>
    text(row["Код_кафедри"]) === departmentHint || text(row["Назва_кафедри"]).toLocaleLowerCase("uk-UA") === departmentHint.toLocaleLowerCase("uk-UA"),
  );
  const metadataCode = text(matchedDepartment?.["Код_кафедри"]);
  if (metadataCode && (!linkedCodes.length || linkedCodes.includes(metadataCode))) return metadataCode;
  const emailDepartmentCode = text(departmentEmailMatch?.["Код_кафедри"]);
  if (emailDepartmentCode) return emailDepartmentCode;
  return linkedCodes.find((code) => departments.some((row) => text(row["Код_кафедри"]) === code)) ?? "";
}

const bindingPromises = new Map<string, Promise<{ departmentCode: string; facultyCode: string }>>();
export const CURRENT_DEPARTMENT_CHANGED_EVENT = "current-department-changed";

/** Persists a user's selected department so every page resolves the same active scope. */
export async function setCurrentUserDepartment(profile: AuthProfile, departmentCode: string, facultyCode: string): Promise<void> {
  if (!profile) throw new Error("Користувача не авторизовано.");
  const code = text(departmentCode);
  if (!code) throw new Error("Не вказано код кафедри.");
  const identity = text(profile.id) || text(profile.email).toLocaleLowerCase("uk-UA");
  const metadata = profile.user_metadata ?? {};
  const { error } = await supabase.auth.updateUser({ data: { ...metadata, Код_кафедри: code, Код_факультету_нні: text(facultyCode) } });
  if (error) throw error;
  bindingPromises.delete(identity);
  if (typeof window !== "undefined") window.dispatchEvent(new Event(CURRENT_DEPARTMENT_CHANGED_EVENT));
}

/** Assigns legacy Auth accounts to an existing department and persists the link in Auth metadata. */
export function ensureCurrentUserDepartment(profile: AuthProfile): Promise<{ departmentCode: string; facultyCode: string }> {
  if (!profile) return Promise.resolve({ departmentCode: "", facultyCode: "" });
  const identity = text(profile.id) || text(profile.email).toLocaleLowerCase("uk-UA");
  const existing = bindingPromises.get(identity);
  if (existing) return existing;
  const task = (async () => {
  const [departmentResult, teacherResult, infoResult] = await Promise.all([
    supabase.from("Кафедра").select("*").order("Назва_кафедри"),
    supabase.from("Викладач").select("Код_викладача,Email_викладача"),
    supabase.from("Кафедральна_інформація_про_виклад").select("Код_викладача,Код_кафедри"),
  ]);
  if (departmentResult.error) throw departmentResult.error;
  const departments = (departmentResult.data ?? []) as unknown as Row[];
  const teachers = (teacherResult.data ?? []) as unknown as Row[];
  const info = (infoResult.data ?? []) as unknown as Row[];
  const resolved = resolveCurrentDepartmentCode(profile, teachers, info, departments)
    || text(departments[0]?.["Код_кафедри"]);
  const department = departments.find((row) => text(row["Код_кафедри"]) === resolved);
  const facultyCode = findDepartmentFacultyCode(department);
  const metadata = profile.user_metadata ?? {};
  if (resolved && (text(metadata["Код_кафедри"]) !== resolved || text(metadata["Код_факультету_нні"]) !== facultyCode)) {
    const { error } = await supabase.auth.updateUser({ data: { ...metadata, Код_кафедри: resolved, Код_факультету_нні: facultyCode } });
    if (error) throw error;
  }
  return { departmentCode: resolved, facultyCode };
  })();
  bindingPromises.set(identity, task);
  void task.catch(() => bindingPromises.delete(identity));
  return task;
}

export function findDepartmentFacultyCode(department: Row | undefined): string {
  return text(department?.["Код_факультету_нні"] ?? department?.["Код_факультету"]);
}

export function findDepartmentHeadName(department: Row | undefined, teachers: Array<{ code?: string; departmentCode?: string; position?: string; name?: string }>): string {
  const explicitHead = department && Object.entries(department).find(([key, value]) => /завідувач|керівник/i.test(key) && text(value));
  if (explicitHead) {
    const headValue = text(explicitHead[1]);
    return teachers.find((teacher) => text(teacher.code) === headValue)?.name ?? headValue;
  }
  return teachers.find((teacher) => /завідувач кафедри/i.test(text(teacher.position)) && text(teacher.departmentCode) === text(department?.["Код_кафедри"]))?.name ?? "";
}
