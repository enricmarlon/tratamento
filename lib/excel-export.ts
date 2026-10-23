import ExcelJS from "exceljs";

function cleanHeader(header: string) {
  return header.trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

export function safeExcelFileName(value: string) {
  const name = value.trim().replace(/\.xlsx$/i, "").replace(/[<>:\"/\\|?*\u0000-\u001f]/g, "_").replace(/[. ]+$/g, "");
  return `${name || "contatos"}.xlsx`;
}

export async function createExcelBlob(headersToExport: string[], rowsToExport: string[][], courseText?: string) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Contatos tratados");
  const headers = [...headersToExport];
  const rows = rowsToExport.map((row) => [...row]);

  if (courseText !== undefined) {
    const existingCourseIndex = headers.findIndex((header) => cleanHeader(header) === "curso");
    if (existingCourseIndex < 0) {
      headers.splice(4, 0, "curso");
      rows.forEach((row) => row.splice(4, 0, courseText));
    } else if (existingCourseIndex !== 4) {
      const [courseHeader] = headers.splice(existingCourseIndex, 1);
      headers.splice(4, 0, courseHeader);
      rows.forEach((row) => {
        const [courseValue] = row.splice(existingCourseIndex, 1);
        row.splice(4, 0, courseValue ?? "");
      });
    }
    rows.forEach((row) => { row[4] = courseText; });
  }

  worksheet.addRow(headers);
  rows.forEach((row) => worksheet.addRow(row));
  if (courseText !== undefined) {
    rows.forEach((_, index) => {
      worksheet.getCell(index + 2, 5).alignment = { vertical: "top", wrapText: true };
    });
  }
  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.columns = headers.map((header) => ({ width: cleanHeader(header) === "curso" && courseText !== undefined ? 42 : Math.min(Math.max(header.length + 4, 16), 32) }));

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}