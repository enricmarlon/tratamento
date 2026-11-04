import ExcelJS from "exceljs";

function cleanHeader(header: string) {
  return header.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

export function safeExcelFileName(value: string) {
  const name = value.trim().replace(/\.xlsx$/i, "").replace(/[<>:\"/\\|?*\u0000-\u001f]/g, "_").replace(/[. ]+$/g, "");
  return `${name || "contatos"}.xlsx`;
}

export async function createExcelBlob(headersToExport: string[], rowsToExport: string[][], courseText?: string, courseColumnIndex = 4) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Contatos tratados");
  const headers = [...headersToExport];
  const rows = rowsToExport.map((row) => [...row]);

  if (courseText !== undefined) {
    const existingCourseIndex = headers.findIndex((header) => cleanHeader(header) === "curso");
    const targetCourseIndex = Math.max(0, Math.min(courseColumnIndex, headers.length));
    if (existingCourseIndex < 0) {
      headers.splice(targetCourseIndex, 0, "curso");
      rows.forEach((row) => row.splice(targetCourseIndex, 0, courseText));
    } else if (existingCourseIndex !== targetCourseIndex) {
      const [courseHeader] = headers.splice(existingCourseIndex, 1);
      const adjustedTargetIndex = Math.max(0, Math.min(courseColumnIndex, headers.length));
      headers.splice(adjustedTargetIndex, 0, courseHeader);
      rows.forEach((row) => {
        row.splice(existingCourseIndex, 1);
        row.splice(adjustedTargetIndex, 0, "");
      });
    }
    const finalCourseIndex = existingCourseIndex >= 0 && existingCourseIndex !== targetCourseIndex
      ? Math.max(0, Math.min(courseColumnIndex, headers.length - 1))
      : targetCourseIndex;
    rows.forEach((row) => { row[finalCourseIndex] = courseText; });
    courseColumnIndex = finalCourseIndex;
  }

  worksheet.addRow(headers);
  rows.forEach((row) => worksheet.addRow(row));
  if (courseText !== undefined) {
    rows.forEach((_, index) => {
      worksheet.getCell(index + 2, courseColumnIndex + 1).alignment = { vertical: "top", wrapText: true };
    });
  }
  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.columns = headers.map((header) => ({ width: cleanHeader(header) === "curso" && courseText !== undefined ? 42 : Math.min(Math.max(header.length + 4, 16), 32) }));

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}