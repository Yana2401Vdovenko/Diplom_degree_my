import { supabase } from "./supabaseClient";

type AuthUser = {
  id?: string;
  email?: string | null;
  app_metadata?: Record<string, unknown>;
  user_metadata?: Record<string, unknown> | null;
};
export type ResolvedRole = "teacher" | "laborant" | "head" | null;

const text = (value: unknown) => String(value ?? "").trim();
const roleAliases: Record<string, ResolvedRole> = {
  teacher: "teacher", викладач: "teacher", laborant: "laborant", лаборант: "laborant",
  head: "head", "завідувач кафедри": "head", admin: null, адміністратор: null,
};

export function normalizeRole(user: AuthUser | null | undefined): ResolvedRole {
  const raw = text(user?.app_metadata?.role ?? user?.user_metadata?.role_code ?? user?.user_metadata?.role);
  return roleAliases[raw.toLocaleLowerCase("uk-UA")] ?? null;
}

/** Resolves staff roles from the actual staff assignment, while preserving the existing lab assistant binding. */
export async function resolveAuthenticatedRole(user: AuthUser): Promise<{ role: ResolvedRole; departmentCode: string; departmentCodes: string[] }> {
  const metadata = user.user_metadata ?? {};
  const existingRole = normalizeRole(user);
  const metadataDepartment = text(metadata["Код_кафедри"] ?? metadata.department_code);
  if (existingRole === "laborant") return { role: existingRole, departmentCode: metadataDepartment, departmentCodes: metadataDepartment ? [metadataDepartment] : [] };

  const email = text(user.email);
  if (!email) return { role: existingRole, departmentCode: metadataDepartment, departmentCodes: metadataDepartment ? [metadataDepartment] : [] };
  const { data: teacher, error: teacherError } = await supabase.from("Викладач").select("*").ilike("Email_викладача", email).limit(1).maybeSingle();
  if (teacherError || !teacher) return { role: existingRole, departmentCode: metadataDepartment, departmentCodes: metadataDepartment ? [metadataDepartment] : [] };

  const teacherCode = text((teacher as Record<string, unknown>)["Код_викладача"]);
  const [{ data: assignments }, { data: positions }, { data: departments }] = await Promise.all([
    supabase.from("Кафедральна_інформація_про_виклад").select("*").eq("Код_викладача", teacherCode),
    supabase.from("Посада").select("*"),
    supabase.from("Кафедра").select("*")
  ]);
  const assignmentRows = (assignments ?? []) as Record<string, unknown>[];
  const positionRows = (positions ?? []) as Record<string, unknown>[];
  const departmentRows = (departments ?? []) as Record<string, unknown>[];
  const positionName = (assignment: Record<string, unknown>) => {
    const code = text(assignment["Код_посади"]);
    const position = positionRows.find((row) => text(row["Код_посади"]) === code);
    return text(position?.["Назва_посади"] ?? position?.Назва ?? position?.Посада ?? assignment["Назва_посади"]);
  };
  const departmentCodes = [...new Set(assignmentRows.map((row) => text(row["Код_кафедри"])).filter(Boolean))];
  const headAssignment = assignmentRows.find((row) => /завідувач кафедри/i.test(positionName(row)));
  const preferredCode = headAssignment
    ? text(headAssignment["Код_кафедри"])
    : departmentCodes.includes(metadataDepartment) ? metadataDepartment : departmentCodes[0] ?? "";
  const role: ResolvedRole = headAssignment ? "head" : assignmentRows.length ? "teacher" : existingRole;
  const validCodes = departmentCodes.filter((code) => departmentRows.some((row) => text(row["Код_кафедри"]) === code));
  const resolvedCodes = validCodes.length ? validCodes : departmentCodes;
  const selectedDepartment = departmentRows.find((row) => text(row["Код_кафедри"]) === preferredCode);
  const facultyCode = text(selectedDepartment?.["Код_факультету_нні"] ?? selectedDepartment?.["Код_факультету"]);

  if (role) {
    const { error } = await supabase.auth.updateUser({ data: {
      role,
      role_code: role,
      ...(preferredCode ? { "Код_кафедри": preferredCode, department_code: preferredCode } : {}),
      ...(facultyCode ? { "Код_факультету_нні": facultyCode } : {}),
      ...(resolvedCodes.length ? { "Коди_кафедр": resolvedCodes, department_codes: resolvedCodes } : {}),
    } });
    if (error) console.error("Не вдалося зберегти визначену роль/кафедри користувача:", error.message);
  }
  return { role, departmentCode: preferredCode, departmentCodes: resolvedCodes };
}
