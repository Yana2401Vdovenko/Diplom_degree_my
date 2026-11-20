import { AlignmentType, Document, PageOrientation, Packer, Paragraph, Table, TableCell, TableLayoutType, TableRow, TextRun, VerticalAlign, WidthType } from "docx";
import pdfMake from "pdfmake/build/pdfmake";
import pdfFonts from "pdfmake/build/vfs_fonts";
import * as XLSX from "xlsx";

pdfMake.addVirtualFileSystem(pdfFonts);

const DOCX_PAGE_WIDTH_DXA = 15400;
const DOCX_PAGE_MARGIN_DXA = 720;
const DOCX_PAGE_CONTENT_WIDTH_MM = DOCX_PAGE_WIDTH_DXA / 1440 * 25.4;
const docxPageProperties = () => ({
  size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
  margin: { top: DOCX_PAGE_MARGIN_DXA, bottom: DOCX_PAGE_MARGIN_DXA, left: DOCX_PAGE_MARGIN_DXA, right: DOCX_PAGE_MARGIN_DXA },
});
const proportionalDxaWidths = (weights: number[]) => {
  const weightTotal = weights.reduce((sum, weight) => sum + weight, 0) || 1;
  let assigned = 0;
  return weights.map((weight, index) => {
    const width = index === weights.length - 1 ? DOCX_PAGE_WIDTH_DXA - assigned : Math.round(DOCX_PAGE_WIDTH_DXA * weight / weightTotal);
    assigned += width;
    return width;
  });
};
const setLandscapeSheet = (sheet: XLSX.WorkSheet) => {
  (sheet as XLSX.WorkSheet & { "!pageSetup"?: { orientation: "landscape"; paperSize: number; fitToWidth: number; fitToHeight: number } })["!pageSetup"] = {
    orientation: "landscape", paperSize: 9, fitToWidth: 1, fitToHeight: 0,
  };
};
const mmToPt = (millimeters: number) => millimeters * 72 / 25.4;
const PDF_A4_LANDSCAPE_WIDTH_MM = 297;
const PDF_PAGE_MARGIN_MM = 12.7;
const pdfContentWidthMm = PDF_A4_LANDSCAPE_WIDTH_MM - PDF_PAGE_MARGIN_MM * 2;
const pdfColumnWidths = (weights: number[]) => {
  const sum = weights.reduce((total, value) => total + value, 0) || 1;
  // pdfmake adds cell padding and vertical borders outside numeric column widths.
  // Reserve that space first so the rendered table stays inside the printable A4 width.
  const horizontalPaddingMm = weights.length * 2 * 0.8;
  const verticalBordersMm = (weights.length + 1) * 0.35 * 25.4 / 72;
  const columnContentWidthMm = Math.max(0, pdfContentWidthMm - horizontalPaddingMm - verticalBordersMm - 0.5);
  return weights.map((value) => mmToPt(columnContentWidthMm * value / sum));
};
const pdfCompactTableLayout = {
  hLineWidth: () => 0.35, vLineWidth: () => 0.35,
  hLineColor: () => "#cbd5e1", vLineColor: () => "#cbd5e1",
  paddingLeft: () => mmToPt(0.8), paddingRight: () => mmToPt(0.8), paddingTop: () => mmToPt(0.8), paddingBottom: () => mmToPt(0.8),
};
const pdfHeaderCell = (text: string) => ({
  text: String(text), bold: true, fontSize: 8.5, lineHeight: 1.2, fillColor: [241, 245, 249], color: [15, 23, 42],
  margin: [0, 0, 0, 0],
});
const pdfLineHeightCell = (cell: any) => cell && typeof cell === "object"
  ? { ...cell, ...(cell.text === undefined ? { text: "" } : {}), lineHeight: 1.2 }
  : { text: String(cell ?? ""), lineHeight: 1.2 };
const reportColumnWeights = (kind: EditableReport["kind"], hideAssignedTeacher = false): number[] => {
  if (kind === "disciplines") return [3, 18, 25, 16, 15, 23];
  if (kind === "form56") return [3, 23, 9, 5, 7, ...Array(10).fill(4), 13];
  if (kind === "plan" && hideAssignedTeacher) return [420, 5350, 1050, 550, 750, ...Array(10).fill(640), 880];
  return [420, 2600, 1050, 550, 750, ...Array(10).fill(640), 880, 2750];
};
const workloadPdfWidths = (kind: EditableReport["kind"], hideAssignedTeacher = false): number[] => {
  const weights = reportColumnWeights(kind, hideAssignedTeacher);
  // The discipline column was too wide in the PDF view; give the report columns
  // on the right more room while keeping the Word layout unchanged.
  if (kind === "plan" && !hideAssignedTeacher && weights.length > 1) weights[1] *= 0.82;
  return pdfColumnWidths(weights);
};
const docxDefaultStyles = () => ({
  default: {
    document: {
      run: { font: "Times New Roman", color: "000000" },
      paragraph: { indent: { left: 0, right: 0, firstLine: 0, hanging: 0 } },
    },
  },
});
const compactPdfHead = (headers: string[]) => headers.map((header) => header
  .replace(/Спеціальність\s*\/\s*ОПП/gi, "Спец. / ОПП")
  .replace(/^Семестр$/i, "Сем.")
  .replace(/^Разом,\s*год$/i, "Всього, год"));

export type ExportTeacher = { id: number; name: string; position: string; rate: number };
export type TeacherRosterExportRow = {
  name: string;
  position: string;
  category: string;
  rate: string | number;
  employment: string;
  normHours: number;
  degree: string;
  rank: string;
  phone: string;
  email: string;
  birthDate: string;
};
export type ExportWorkload = {
  teacher_id: number;
  position?: string;
  rate?: number;
  study_form?: string;
  semester?: number;
  assigned_teacher?: string;
  discipline_name: string;
  groups: string;
  lectures?: number;
  labs?: number;
  practices?: number;
  exams?: number;
  credits?: number;
  course_projects?: number;
  seminars?: number;
  practice_hours?: number;
  reviews?: number;
  diploma_project?: number;
  total_hours?: number;
  notes?: string;
};
export type ExportOrganization = { facultyName?: string; departmentName?: string; departmentHeadName?: string };
export type DisciplineReportRow = {
  specialty: string;
  discipline: string;
  semester: string;
  forms: string[];
  control: string;
  teachers: Array<{ name: string; workTypes: string[] }>;
};

const value = (item: unknown) => Number(item ?? 0);
const total = (row: ExportWorkload): number => {
  const storedTotal = value(row.total_hours);
  return storedTotal !== 0 ? storedTotal
    : [row.lectures, row.labs, row.practices, row.seminars, row.exams, row.credits, row.course_projects, row.practice_hours, row.diploma_project, row.reviews].reduce<number>((sum, item) => sum + value(item), 0);
};
const fmt = (value: number) => new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 2 }).format(value);
const formatHours = (hours: number): string => new Intl.NumberFormat("uk-UA", { maximumFractionDigits: 2 }).format(Number(hours) || 0);

const saveBlob = (blob: Blob, fileName: string) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
};

export type EditableReport = {
  title: string;
  subtitle: string;
  rows: ExportWorkload[];
  fileName: string;
  organization: ExportOrganization;
  kind: "teacher" | "form56" | "plan" | "disciplines";
  disciplineRows?: DisciplineReportRow[];
  hideAssignedTeacher?: boolean;
  form56Selection?: "day" | "extramural" | "all";
  comment: string;
  reportDate: string;
  teacherSigner: string;
  departmentSigner: string;
};

const reportColumns: Array<[string, keyof ExportWorkload]> = [
  ["Лек.", "lectures"], ["Лаб.", "labs"], ["Практ.", "practices"], ["Семін.", "seminars"],
  ["Екз.", "exams"], ["Залік", "credits"], ["Курс.", "course_projects"], ["Практ-ка", "practice_hours"],
  ["ДП / ЕК", "diploma_project"], ["Рец.", "reviews"],
];
const totalRow = (rows: ExportWorkload[]) => reportColumns.map(([, key]) => rows.reduce((sum, row) => sum + value(row[key]), 0));
const exportHeader = (report: EditableReport) => [
  "Національний університет біоресурсів і природокористування України",
  report.organization.facultyName || "",
  report.organization.departmentName || "",
  report.title,
  report.subtitle,
].filter(Boolean);
const signatureLines = (report: EditableReport) => [
  `Завідувач кафедри: ${report.departmentSigner || "____________________"}`,
  ...(report.kind === "teacher" ? [`Викладач: ${report.teacherSigner || "____________________"}`] : []),
  `Дата: ${report.reportDate || "____________________"}`,
];
const tableRowsForExport = (rows: ExportWorkload[], kind: EditableReport["kind"] = "teacher", hideAssignedTeacher = false) => rows.map((row, index) => kind === "plan" ? [
  String(index + 1), row.discipline_name || "—", row.groups || "—", row.semester ? String(row.semester) : "—", row.study_form || "—",
  ...reportColumns.map(([, key]) => formatHours(value(row[key]))), formatHours(total(row)), ...(hideAssignedTeacher ? [] : [row.assigned_teacher || "—"]),
] : [
  String(index + 1), row.discipline_name || "—", row.groups || "—", row.semester ? String(row.semester) : "—", row.study_form || "—",
  ...reportColumns.map(([, key]) => formatHours(value(row[key]))), formatHours(total(row)), row.notes || "",
]);
const tableHeadForExport = (kind: EditableReport["kind"] = "teacher", hideAssignedTeacher = false) => kind === "plan"
  ? ["№", "Дисципліна", "Спеціальність / ОПП", "Семестр", "Форма", ...reportColumns.map(([label]) => label), "Всього, год", ...(hideAssignedTeacher ? [] : ["Призначений викладач"])]
  : ["№", "Навчальна дисципліна", "Спеціальність / ОПП", "Семестр", "Форма", ...reportColumns.map(([label]) => label), "Разом, год", "Примітки"];
const totalTableRow = (rows: ExportWorkload[], kind: EditableReport["kind"], hideAssignedTeacher = false) => kind === "plan"
  ? ["Всього", "", "", "", "", ...totalRow(rows).map(formatHours), formatHours(rows.reduce((sum, row) => sum + total(row), 0)), ...(hideAssignedTeacher ? [] : [""])]
  : ["Всього", "", "", "", "", ...totalRow(rows).map(formatHours), formatHours(rows.reduce((sum, row) => sum + total(row), 0)), ""];
const disciplineReportHead = ["№ з/п", "Спеціальність / ОПП", "Дисципліна", "Семестр та форма навчання", "Форма проведення (контроль)", "Викладачі"];
const disciplineReportRows = (rows: DisciplineReportRow[] = []) => rows.map((row, index) => [
  String(index + 1), row.specialty || "—", row.discipline || "—",
  [row.semester, row.forms.join(", ")].filter(Boolean).join(" · ") || "—",
  row.control || "—",
  row.teachers.filter((teacher) => teacher.name !== "Не призначено").length ? row.teachers.filter((teacher) => teacher.name !== "Не призначено").map((teacher) => {
    const kinds = teacher.workTypes.filter((kind) => !/^(екзамен|залік)$/i.test(kind.trim())).map((kind) => kind.toLocaleLowerCase("uk-UA"));
    return kinds.length ? `${teacher.name} (${kinds.join(", ")})` : teacher.name;
  }).join("\n") : "Не призначено",
]);
const form56Head = ["№", "Викладач", "Посада", "Ставка", "Форма навчання", ...reportColumns.map(([label]) => label), "Разом, год"];
const form56Groups = (rows: ExportWorkload[]) => [...new Set(rows.map((row) => row.teacher_id))].map((teacherId) => {
  const teacherRows = rows.filter((row) => row.teacher_id === teacherId);
  const base = teacherRows[0];
  const forForm = (form: "Денна" | "Заочна") => teacherRows.find((row) => row.study_form === form) || {
    ...base, study_form: form, groups: form.toLocaleLowerCase("uk-UA"), total_hours: 0, ...Object.fromEntries(reportColumns.map(([, key]) => [key, 0])),
  } as ExportWorkload;
  const day = forForm("Денна");
  const extramural = forForm("Заочна");
  const overall = { ...day, study_form: "Всього", total_hours: total(day) + total(extramural), ...Object.fromEntries(reportColumns.map(([, key]) => [key, value(day[key]) + value(extramural[key])])) } as ExportWorkload;
  return { day, extramural, overall };
});
const form56Totals = (rows: ExportWorkload[]) => reportColumns.map(([, key]) => rows.reduce((sum, row) => sum + value(row[key]), 0));
const form56Overall = (rows: ExportWorkload[]) => rows.reduce((sum, row) => sum + total(row), 0);
const form56Data = (groups: ReturnType<typeof form56Groups>, selection: NonNullable<EditableReport["form56Selection"]> = "all") => groups.flatMap(({ day, extramural, overall }, index) => {
  const single = selection === "day" ? [day] : selection === "extramural" ? [extramural] : [day, extramural, overall];
  return single.map((row, rowIndex) => [
    rowIndex === 0 ? String(index + 1) : "", rowIndex === 0 ? day.discipline_name || "—" : "", rowIndex === 0 ? day.position || "—" : "", rowIndex === 0 ? day.rate == null ? "—" : fmt(value(day.rate)) : "",
    row.study_form === "Денна" ? "Денна" : row.study_form === "Заочна" ? "Заочна" : "Всього", ...reportColumns.map(([, key]) => formatHours(value(row[key]))), formatHours(total(row)),
  ]);
});
const form56FootRows = (dayRows: ExportWorkload[], extramuralRows: ExportWorkload[], selection: NonNullable<EditableReport["form56Selection"]> = "all") => {
  const day = ["Всього за денною формою", ...form56Totals(dayRows).map(formatHours), formatHours(form56Overall(dayRows))];
  const extramural = ["Всього за заочною формою", ...form56Totals(extramuralRows).map(formatHours), formatHours(form56Overall(extramuralRows))];
  const all = ["Загалом по кафедрі (Денна + Заочна)", ...form56Totals([...dayRows, ...extramuralRows]).map(formatHours), formatHours(form56Overall([...dayRows, ...extramuralRows]))];
  return selection === "day" ? [day] : selection === "extramural" ? [extramural] : [day, extramural, all];
};

export async function exportEditableReport(report: EditableReport, format: "docx" | "xlsx" | "pdf") {
  if (format === "xlsx") {
    const sheetRows: unknown[][] = [...exportHeader(report).map((text) => [text]), []];
    let merges: XLSX.Range[] = [];
    const headerRowIndex = sheetRows.length;
    if (report.kind === "disciplines") {
      sheetRows.push(disciplineReportHead, ...disciplineReportRows(report.disciplineRows));
    } else if (report.kind === "form56") {
      sheetRows.push(form56Head);
      const groups = form56Groups(report.rows);
      const startRow = sheetRows.length;
      sheetRows.push(...form56Data(groups, report.form56Selection));
      groups.forEach((_, index) => {
        if (!report.form56Selection || report.form56Selection === "all") {
          const first = startRow + index * 3;
          for (let column = 0; column < 4; column += 1) merges.push({ s: { r: first, c: column }, e: { r: first + 2, c: column } });
        }
      });
      const dayRows = report.rows.filter((row) => row.study_form === "Денна");
      const extramuralRows = report.rows.filter((row) => row.study_form === "Заочна");
      form56FootRows(dayRows, extramuralRows, report.form56Selection).forEach(([label, ...values]) => sheetRows.push([label, "", "", "", "", ...values]));
    } else {
      sheetRows.push(tableHeadForExport(report.kind, report.hideAssignedTeacher), ...tableRowsForExport(report.rows, report.kind, report.hideAssignedTeacher));
      sheetRows.push(totalTableRow(report.rows, report.kind, report.hideAssignedTeacher));
    }
    if (report.comment.trim()) sheetRows.push([], [report.comment]);
    sheetRows.push([], ...signatureLines(report).map((line) => [line]));
    const sheet = XLSX.utils.aoa_to_sheet(sheetRows);
    setLandscapeSheet(sheet);
    const columnCount = Math.max(...sheetRows.map((row) => row.length));
    exportHeader(report).forEach((_, rowIndex) => {
      if (columnCount > 1) merges.push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: columnCount - 1 } });
    });
    if (merges.length) sheet["!merges"] = merges;
    const rowValues = (columnIndex: number) => sheetRows.map((row) => String(row[columnIndex] ?? ""));
    const preferredWidths = report.kind === "disciplines" ? [5, 24, 40, 24, 26, 54]
      : report.kind === "form56" ? [5, 36, 22, 8, 14, ...Array(10).fill(9), 11]
        : report.kind === "plan" ? [5, 42, 28, 9, 14, ...Array(10).fill(9), 11, ...(report.hideAssignedTeacher ? [] : [32])]
          : [5, 42, 28, 9, 14, ...Array(10).fill(9), 11, 32];
    sheet["!cols"] = Array.from({ length: columnCount }, (_, columnIndex) => {
      const maxLength = Math.max(0, ...rowValues(columnIndex).map((text) => text.split("\n").reduce((max, line) => Math.max(max, line.length), 0)));
      const preferred = preferredWidths[columnIndex] ?? 16;
      const width = columnIndex === 0 ? 5 : Math.min(preferred, Math.max(8, maxLength + 4));
      return { wch: width };
    });
    // SheetJS community preserves these style definitions in workbook cells;
    // borders/fills provide a clear table hierarchy in compatible viewers.
    const border = { top: { style: "thin", color: { rgb: "D1D5DB" } }, bottom: { style: "thin", color: { rgb: "D1D5DB" } }, left: { style: "thin", color: { rgb: "D1D5DB" } }, right: { style: "thin", color: { rgb: "D1D5DB" } } };
    for (let rowIndex = headerRowIndex; rowIndex < sheetRows.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < (sheetRows[rowIndex]?.length ?? 0); columnIndex += 1) {
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
        const cell = sheet[address] as (XLSX.CellObject & { s?: Record<string, unknown> }) | undefined;
        if (!cell) continue;
        const isHeader = rowIndex === headerRowIndex;
        const isTotal = String(sheetRows[rowIndex]?.[0] ?? "").toLocaleLowerCase("uk-UA").includes("всього") || String(sheetRows[rowIndex]?.[0] ?? "").toLocaleLowerCase("uk-UA").includes("загалом");
        cell.s = { alignment: { vertical: "center", horizontal: columnIndex > 4 && columnIndex < columnCount - 1 ? "right" : "left", wrapText: true }, border, ...(isHeader ? { fill: { fgColor: { rgb: "F1F5F9" } }, font: { bold: true, color: { rgb: "1E293B" } } } : {}), ...(isTotal ? { font: { bold: true, color: { rgb: "0F172A" } }, fill: { fgColor: { rgb: "E2E8F0" } } } : {}) };
      }
    }
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, report.kind === "form56" ? "Форма 56" : report.kind === "disciplines" ? "Дисципліни" : "Навантаження");
    XLSX.writeFile(book, `${report.fileName}.xlsx`, { cellStyles: true });
    return;
  }

  if (format === "docx") {
    const paragraphs: Array<Paragraph | Table> = [
      ...exportHeader(report).map((text, index) => new Paragraph({ alignment: "center", spacing: { line: 240 }, children: [new TextRun({ text, bold: index === 0 || index === 3, font: "Times New Roman", size: index === 3 ? 24 : 20 })] })),
      new Paragraph({ text: "" }),
    ];
    const widthsDxa = proportionalDxaWidths(reportColumnWeights(report.kind, report.hideAssignedTeacher));
    const docxCell = (text: unknown, column: number, options: { bold?: boolean; header?: boolean; span?: number; align?: "left" | "center" | "right" } = {}) => {
      const span = options.span ?? 1;
      const width = widthsDxa.slice(column, column + span).reduce((sum, item) => sum + item, 0);
      return new TableCell({ columnSpan: span > 1 ? span : undefined, width: { size: width, type: WidthType.DXA }, margins: { top: 45, bottom: 45, left: 45, right: 45 }, shading: options.header ? { fill: "F1F5F9" } : undefined, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: options.align === "right" ? AlignmentType.RIGHT : options.align === "center" ? AlignmentType.CENTER : AlignmentType.LEFT, indent: { left: 0, right: 0, firstLine: 0, hanging: 0 }, spacing: { before: 0, after: 0, line: 240 }, children: [new TextRun({ text: String(text ?? ""), bold: options.bold, font: "Times New Roman", size: report.kind === "form56" ? 19 : 17 })] })] });
    };
    const docxRow = (cells: TableCell[], header = false) => new TableRow({ tableHeader: header, cantSplit: true, children: cells });
    const buildSimpleTable = (headers: string[], bodyRows: unknown[][], totals?: unknown[], widths = widthsDxa) => new Table({
      width: { size: DOCX_PAGE_WIDTH_DXA, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
      rows: [docxRow(headers.map((text, index) => docxCell(text, index, { bold: true, header: true })), true),
        ...bodyRows.map((row) => docxRow(row.map((text, index) => docxCell(text, index, { align: index > 4 && index < row.length - 1 ? "right" : "left" })))),
        ...(totals ? [docxRow([docxCell(totals[0], 0, { bold: true, span: 5 }), ...totals.slice(5).map((text, index) => docxCell(text, index + 5, { bold: true, align: "right" }))])] : [])],
    });
    if (report.kind === "disciplines") {
      const table = buildSimpleTable(disciplineReportHead.map(String), disciplineReportRows(report.disciplineRows).map((row) => row.map(String)), undefined, widthsDxa);
      paragraphs.push(table);
    } else if (report.kind === "form56") {
      const groups = form56Groups(report.rows);
      const dayRows = report.rows.filter((row) => row.study_form === "Денна");
      const extramuralRows = report.rows.filter((row) => row.study_form === "Заочна");
      const header = docxRow(form56Head.map((text, index) => docxCell(String(text), index, { bold: true, header: true })), true);
      const dataRows = groups.flatMap(({ day, extramural, overall }, index) => {
        const selection = report.form56Selection ?? "all";
        const selectedRows = selection === "day" ? [day] : selection === "extramural" ? [extramural] : [day, extramural, overall];
        return selectedRows.map((row, rowIndex) => {
          const rowSpan = selection === "all" ? 3 : 1;
          const firstCells = rowIndex === 0 ? [String(index + 1), String(day.discipline_name || "—"), String(day.position || "—"), String(day.rate == null ? "—" : fmt(value(day.rate)))].map((text, column) => new TableCell({ rowSpan, width: { size: widthsDxa[column], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 45, bottom: 45, left: 45, right: 45 }, children: [new Paragraph({ indent: { left: 0, right: 0, firstLine: 0, hanging: 0 }, spacing: { before: 0, after: 0, line: 240 }, children: [new TextRun({ text, font: "Times New Roman", size: 19 })] })] })) : [];
          const formLabel = row.study_form === "Денна" ? "Денна" : row.study_form === "Заочна" ? "Заочна" : "Всього";
          const cells = [formLabel, ...reportColumns.map(([, key]) => formatHours(value(row[key]))), formatHours(total(row))];
          return docxRow([...firstCells, ...cells.map((text, cellIndex) => docxCell(text, cellIndex + 4, { align: cellIndex > 0 ? "right" : "left", bold: selection === "all" && rowIndex === 2 }))]);
        });
      });
      const footRows = form56FootRows(dayRows, extramuralRows, report.form56Selection).map(([label, ...values]) => docxRow([docxCell(label, 0, { bold: true, span: 5 }), ...values.map((text, index) => docxCell(text, index + 5, { bold: true, align: "right" }))]));
      paragraphs.push(new Table({ width: { size: DOCX_PAGE_WIDTH_DXA, type: WidthType.DXA }, columnWidths: widthsDxa, layout: TableLayoutType.FIXED, rows: [header, ...dataRows, ...footRows] }));
    } else {
      paragraphs.push(buildSimpleTable(tableHeadForExport(report.kind, report.hideAssignedTeacher), tableRowsForExport(report.rows, report.kind, report.hideAssignedTeacher), totalTableRow(report.rows, report.kind, report.hideAssignedTeacher), widthsDxa));
    }
    if (report.comment.trim()) paragraphs.push(new Paragraph({ children: [new TextRun({ text: report.comment })] }));
    paragraphs.push(new Paragraph({ text: "" }), ...signatureLines(report).map((line) => new Paragraph({ children: [new TextRun({ text: line, size: 20 })] })));
    const doc = new Document({ styles: docxDefaultStyles(), sections: [{ properties: { page: docxPageProperties() }, children: paragraphs }] });
    saveBlob(await Packer.toBlob(doc), `${report.fileName}.docx`);
    return;
  }

  const content: any[] = [
    ...exportHeader(report).map((text, index) => ({
      text,
      alignment: "center",
      bold: index === 0 || index === 3,
      fontSize: index === 3 ? 12 : 10,
      lineHeight: index === 3 ? 1 : 1.2,
      margin: [0, 0, 0, 0],
    })),
    { text: "", margin: [0, 0, 0, 8] },
  ];
  const addTable = (rows: ExportWorkload[], subtotalLabel = "Всього") => {
    // The DOCX layout starts the table immediately after the report header.
    const totals = totalTableRow(rows, report.kind, report.hideAssignedTeacher).map((cell, index) => index === 0 ? subtotalLabel : cell);
    const widths = workloadPdfWidths(report.kind, report.hideAssignedTeacher);
    const totalColumnIndex = totals.length - (report.kind === "plan" && !report.hideAssignedTeacher ? 2 : 1);
    content.push({ table: { headerRows: 1, keepWithHeaderRows: 1, dontBreakRows: true, widths, body: [compactPdfHead(tableHeadForExport(report.kind, report.hideAssignedTeacher)).map(pdfHeaderCell), ...tableRowsForExport(rows, report.kind, report.hideAssignedTeacher).map((row) => row.map(pdfLineHeightCell)), totals.map((cell, index) => pdfLineHeightCell(index === 0 || index === totalColumnIndex ? { text: String(cell), bold: true, noWrap: index === 0 } : cell))] }, layout: pdfCompactTableLayout, fontSize: 8.5, margin: [0, 0, 0, 12] });
  };
  if (report.kind === "disciplines") {
    content.push({ table: { headerRows: 1, keepWithHeaderRows: 1, dontBreakRows: true, widths: workloadPdfWidths("disciplines"), body: [compactPdfHead(disciplineReportHead).map(pdfHeaderCell), ...disciplineReportRows(report.disciplineRows).map((row) => row.map((text) => ({ text: String(text), lineHeight: 1.2 })))] }, layout: pdfCompactTableLayout, fontSize: 8.5, margin: [0, 0, 0, 12] });
  } else if (report.kind === "form56") {
    const groups = form56Groups(report.rows);
    const dayRows = report.rows.filter((row) => row.study_form === "Денна");
    const extramuralRows = report.rows.filter((row) => row.study_form === "Заочна");
    const body: any[][] = [compactPdfHead(form56Head.map((header) => header.replace("Форма навчання", "Форма"))).map((header) => ({ ...pdfHeaderCell(header), fontSize: 9.5 }))];
    groups.forEach(({ day, extramural, overall }, index) => {
      const selection = report.form56Selection ?? "all";
      const selectedRows = selection === "day" ? [day] : selection === "extramural" ? [extramural] : [day, extramural, overall];
      const span = selection === "all" ? 3 : 1;
      selectedRows.forEach((row, rowIndex) => {
      body.push([
        ...(rowIndex === 0 ? [{ text: String(index + 1), rowSpan: span, alignment: "center" }, { text: day.discipline_name || "—", rowSpan: span }, { text: day.position || "—", rowSpan: span }, { text: day.rate == null ? "—" : fmt(value(day.rate)), rowSpan: span, alignment: "center" }] : [{}, {}, {}, {}]),
        row.study_form === "Денна" ? "Денна" : row.study_form === "Заочна" ? "Заочна" : { text: "Всього", bold: true }, ...reportColumns.map(([, key]) => formatHours(value(row[key]))), formatHours(total(row)),
      ]);
      });
    });
    form56FootRows(dayRows, extramuralRows, report.form56Selection).forEach(([label, ...values]) => body.push([{ text: String(label), colSpan: 5, bold: true }, "", "", "", "", ...values.map((text) => ({ text: String(text), bold: true }))]));
      content.push({ table: { headerRows: 1, keepWithHeaderRows: 1, dontBreakRows: true, widths: workloadPdfWidths("form56"), body: body.map((row) => row.map(pdfLineHeightCell)) }, layout: pdfCompactTableLayout, fontSize: 9.5, margin: [0, 0, 0, 12] });
  } else addTable(report.rows);
  if (report.comment.trim()) content.push({ text: report.comment, margin: [0, 10, 0, 0] });
  content.push({ text: signatureLines(report).join("     "), style: "signatures" });
  pdfMake.createPdf({ pageOrientation: "landscape", pageSize: "A4", pageMargins: [mmToPt(12.7), mmToPt(12.7), mmToPt(12.7), mmToPt(12.7)], content,
    styles: { centerHeader: { fontSize: 11, bold: true, alignment: "center", margin: [0, 0, 0, 4] }, section: { fontSize: 10, bold: true, margin: [0, 10, 0, 6] }, total: { fontSize: 10, bold: true, margin: [0, 10, 0, 0] }, signatures: { fontSize: 9, margin: [0, 36, 0, 0] } }, defaultStyle: { font: "Roboto", fontSize: 9 } }).download(`${report.fileName}.pdf`);
}

const teacherRosterHeaders = ["№", "ПІБ викладача", "Посада", "Категорія (НПП/ПП)", "Ставка", "Зайнятість", "Норма, год", "Науковий ступінь", "Вчене звання", "ДЕНЬ НАР.", "Телефон", "Email"];
const teacherRosterColumnWeights = [3, 15, 9, 10, 6, 8, 6, 12, 10, 6, 8, 12];
const teacherRosterData = (rows: TeacherRosterExportRow[]) => rows.map((row, index) => [String(index + 1), row.name || "—", row.position || "—", row.category || "—", String(row.rate ?? "—"), row.employment || "Не вказано", String(row.normHours || 0), row.degree || "—", row.rank || "—", row.birthDate || "—", row.phone || "—", row.email || "—"]);

export async function exportTeacherRosterReport(report: {
  title: string;
  subtitle?: string;
  facultyName: string;
  departmentName: string;
  rows: TeacherRosterExportRow[];
  fileName: string;
  comment?: string;
  reportDate?: string;
  departmentSigner?: string;
}, format: "docx" | "xlsx" | "pdf") {
  const header = ["Національний університет біоресурсів і природокористування України", report.facultyName || "Факультет не вказано", report.departmentName || "Кафедра не вказана", report.title, ...(report.subtitle ? [report.subtitle] : [])];
  const rows = teacherRosterData(report.rows);
  const endingRows: unknown[][] = [
    ...(report.comment?.trim() ? [[`Коментар: ${report.comment.trim()}`]] : []),
    [`Завідувач кафедри: ${report.departmentSigner?.trim() || "____________________"}    Підпис: ____________________`],
    [`Дата: ${report.reportDate?.trim() || "____________________"}`],
  ];
  if (format === "xlsx") {
    const sheetRows: unknown[][] = [...header.map((line) => [line]), [], teacherRosterHeaders, ...(rows.length ? rows : [["", "За обраними фільтрами викладачів не знайдено.", ...Array(10).fill("")]]), [], ...endingRows];
    const sheet = XLSX.utils.aoa_to_sheet(sheetRows);
    setLandscapeSheet(sheet);
    const firstFooterRow = header.length + 2 + Math.max(rows.length, 1) + 1;
    sheet["!merges"] = [
      ...header.map((_, row) => ({ s: { r: row, c: 0 }, e: { r: row, c: teacherRosterHeaders.length - 1 } })),
      ...endingRows.map((_, index) => ({ s: { r: firstFooterRow + index, c: 0 }, e: { r: firstFooterRow + index, c: teacherRosterHeaders.length - 1 } })),
    ];
    sheet["!cols"] = [5, 36, 22, 20, 10, 23, 12, 30, 26, 14, 18, 34].map((wch) => ({ wch }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Викладачі");
    XLSX.writeFile(workbook, `${report.fileName}.xlsx`);
    return;
  }
  if (format === "docx") {
    const widths = proportionalDxaWidths(teacherRosterColumnWeights);
    const cell = (text: string, index: number, isHeader = false) => new TableCell({
      width: { size: widths[index], type: WidthType.DXA },
      shading: isHeader ? { fill: "F1F5F9" } : undefined,
      margins: { top: 45, bottom: 45, left: 45, right: 45 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: index === 0 || index === 4 || index === 6 ? AlignmentType.CENTER : AlignmentType.LEFT, indent: { left: 0, right: 0, firstLine: 0, hanging: 0 }, spacing: { before: 0, after: 0, line: 240 }, children: [new TextRun({ text: String(text), bold: isHeader, font: "Times New Roman", size: 19 })] })],
    });
    const table = new Table({
      width: { size: DOCX_PAGE_WIDTH_DXA, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED,
      rows: [new TableRow({ tableHeader: true, children: teacherRosterHeaders.map((text, index) => cell(text, index, true)) }),
        ...(rows.length ? rows : [["", "За обраними фільтрами викладачів не знайдено.", ...Array(10).fill("")]]).map((row) => new TableRow({ cantSplit: true, children: row.map((text, index) => cell(text, index)) }))],
    });
    const paragraphs = [
      ...header.map((text, index) => new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: index === 0 || index === 3, font: "Times New Roman", size: index === 3 ? 26 : index === 0 ? 22 : 20 })] })),
      new Paragraph({ text: "" }), table,
      ...(report.comment?.trim() ? [new Paragraph({ text: `Коментар: ${report.comment.trim()}`, spacing: { before: 180 } })] : []),
      new Paragraph({ text: `Завідувач кафедри: ${report.departmentSigner?.trim() || "____________________"}    Підпис: ____________________`, spacing: { before: 500 } }),
      new Paragraph({ text: `Дата: ${report.reportDate?.trim() || "____________________"}` }),
    ];
    const document = new Document({ styles: docxDefaultStyles(), sections: [{ properties: { page: docxPageProperties() }, children: paragraphs }] });
    saveBlob(await Packer.toBlob(document), `${report.fileName}.docx`);
    return;
  }
  const pdfTeacherHeaders = [...teacherRosterHeaders];
  pdfTeacherHeaders[3] = "Категорія";
  const body = [pdfTeacherHeaders.map((text) => ({ text, bold: true, fillColor: [241, 245, 249], color: [15, 23, 42], fontSize: 9.5, lineHeight: 1.2, margin: [0, 0, 0, 0] })), ...(rows.length ? rows : [["", "За обраними фільтрами викладачів не знайдено.", ...Array(10).fill("")]]).map((row) => row.map((text) => ({ text: String(text), fontSize: 9.5, lineHeight: 1.2, margin: [0, 0, 0, 0] })) )];
  const teacherPdfWidths = pdfColumnWidths(teacherRosterColumnWeights);
  pdfMake.createPdf({
      pageOrientation: "landscape", pageSize: "A4", pageMargins: [mmToPt(12.7), mmToPt(12.7), mmToPt(12.7), mmToPt(12.7)],
    content: [
      ...header.map((text, index) => ({ text, alignment: "center", bold: index === 0 || index === 3, fontSize: index === 3 ? 13 : index === 0 ? 11 : 10, lineHeight: index === 3 ? 1 : 1.2, margin: [0, 0, 0, 0] })),
      { text: "", margin: [0, 0, 0, 8] },
      { table: { headerRows: 1, keepWithHeaderRows: 1, dontBreakRows: true, widths: teacherPdfWidths, body: [body[0], ...body.slice(1).map((row) => row.map((cell, index) => index === 11 && typeof cell.text === "string" ? { ...cell, text: cell.text.replace(/([._@+\-])/g, "$1\u200b") } : cell))] }, layout: pdfCompactTableLayout, fontSize: 9.5 },
      ...(report.comment?.trim() ? [{ text: `Коментар: ${report.comment.trim()}`, margin: [0, 9, 0, 0] }] : []),
      { text: `Завідувач кафедри: ${report.departmentSigner?.trim() || "____________________"}    Підпис: ____________________`, margin: [0, 28, 0, 0] },
      { text: `Дата: ${report.reportDate?.trim() || "____________________"}`, margin: [0, 8, 0, 0] },
    ],
  }).download(`${report.fileName}.pdf`);
}

export async function exportWordReport(title: string, subtitle: string, rows: ExportWorkload[], fileName: string, organization: ExportOrganization = {}) {
  const lines = rows.length
    ? rows.map((row, index) => `${index + 1}. ${row.discipline_name || "Навчальне навантаження"}; ${row.groups || "—"}; форма: ${row.study_form || "—"}; лекції: ${formatHours(value(row.lectures))}; лабораторні: ${formatHours(value(row.labs))}; практичні: ${formatHours(value(row.practices))}; семінарські: ${formatHours(value(row.seminars))}; екзамен: ${formatHours(value(row.exams))}; залік: ${formatHours(value(row.credits))}; курсовий проєкт: ${formatHours(value(row.course_projects))}; практика: ${formatHours(value(row.practice_hours))}; дипломні / ЕК: ${formatHours(value(row.diploma_project))}; рецензії: ${formatHours(value(row.reviews))}; примітки: ${row.notes || "—"}; разом: ${formatHours(total(row))} год.`)
    : ["За обраними параметрами записів немає."];
  const overall = rows.reduce<number>((sum, row) => sum + total(row), 0);
    const documentFile = new Document({
    styles: docxDefaultStyles(),
    sections: [{
      properties: { page: docxPageProperties() },
      children: [
        new Paragraph({ alignment: "center", children: [new TextRun({ text: "Національний університет біоресурсів і природокористування України", bold: true, size: 24 })] }),
        new Paragraph({ alignment: "center", children: [new TextRun({ text: organization.facultyName || "Факультет не вказано", size: 22 })] }),
        ...(organization.departmentName ? [new Paragraph({ alignment: "center", children: [new TextRun({ text: organization.departmentName, size: 22 })] })] : []),
        new Paragraph({ alignment: "center", children: [new TextRun({ text: title, bold: true, size: 28 })] }),
        new Paragraph({ alignment: "center", children: [new TextRun({ text: subtitle, size: 22 })] }),
        ...lines.map((line) => new Paragraph({ children: [new TextRun({ text: line, size: 22 })] })),
        new Paragraph({ children: [new TextRun({ text: `Загальний обсяг: ${formatHours(overall)} год.`, bold: true, size: 24 })] }),
        new Paragraph({ children: [new TextRun({ text: `\n\nПідпис НПП: ____________________                    Завідувач кафедри${organization.departmentHeadName ? ` ${organization.departmentHeadName}` : ""}: ____________________                    Дата: ____________`, size: 22 })] }),
      ],
    }],
  });
  saveBlob(await Packer.toBlob(documentFile), `${fileName}.docx`);
}

export function exportPdfReport(title: string, subtitle: string, rows: ExportWorkload[], fileName: string, organization: ExportOrganization = {}) {
  const overall = rows.reduce<number>((sum, row) => sum + total(row), 0);
  const body = rows.length ? rows.map((row, index) => [String(index + 1), row.discipline_name || "Навчальне навантаження", row.groups || "—", row.study_form || "—", formatHours(value(row.lectures)), formatHours(value(row.labs)), formatHours(value(row.practices)), formatHours(value(row.seminars)), formatHours(value(row.exams)), formatHours(value(row.credits)), formatHours(value(row.course_projects)), formatHours(value(row.practice_hours)), formatHours(value(row.diploma_project)), formatHours(value(row.reviews)), formatHours(total(row))]) : [["", "За обраними параметрами записів немає.", ...Array(13).fill("")]];
  pdfMake.createPdf({
    content: [
      { text: "Національний університет біоресурсів і природокористування України", style: "institution" },
      { text: organization.facultyName || "Факультет не вказано", style: "faculty" },
      ...(organization.departmentName ? [{ text: organization.departmentName, style: "faculty" }] : []),
      { text: title, style: "title" }, { text: subtitle, style: "subtitle" },
      { table: { headerRows: 1, keepWithHeaderRows: 1, dontBreakRows: true, widths: pdfColumnWidths([420, 2600, 1800, 1200, ...Array(10).fill(900), 1180]), body: [["№", "Дисципліна", "Група", "Форма", "Лек.", "Лаб.", "Практ.", "Сем.", "Екз.", "Залік", "Курс.", "Практика", "Дипл./ЕК", "Рецензії", "Всього, год"].map(pdfHeaderCell), ...body.map((row) => row.map((text) => ({ text, lineHeight: 1.2 })))] }, layout: pdfCompactTableLayout, fontSize: 8.5 },
      { text: `Загальний обсяг: ${formatHours(overall)} год.`, style: "total" },
      { text: `Підпис НПП: ____________________                         Завідувач кафедри${organization.departmentHeadName ? ` ${organization.departmentHeadName}` : ""}: ____________________                         Дата: ____________`, style: "signatures" },
    ],
    styles: { institution: { fontSize: 12, bold: true, alignment: "center", margin: [0, 0, 0, 4] }, faculty: { fontSize: 11, alignment: "center", margin: [0, 0, 0, 14] }, title: { fontSize: 14, bold: true, alignment: "center", margin: [0, 0, 0, 6] }, subtitle: { fontSize: 10, alignment: "center", color: "#526681", margin: [0, 0, 0, 16] }, total: { fontSize: 11, bold: true, alignment: "right", margin: [0, 14, 0, 0] }, signatures: { fontSize: 10, margin: [0, 36, 0, 0] } },
    pageOrientation: "landscape", pageSize: "A4", pageMargins: [mmToPt(12.7), mmToPt(12.7), mmToPt(12.7), mmToPt(12.7)],
    defaultStyle: { font: "Roboto", fontSize: 9 },
  }).download(`${fileName}.pdf`);
}
