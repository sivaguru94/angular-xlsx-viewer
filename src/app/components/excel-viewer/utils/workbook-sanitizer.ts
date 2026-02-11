/**
 * Sanitizes Univer workbook data produced by LuckyExcel to fix
 * invalid values that cause runtime warnings.
 *
 * Fixes applied:
 * - Negative or zero column widths → default 72px
 * - Negative or zero default column width → 72px
 * - Negative or zero row heights → default 20px
 */
export function sanitizeWorkbookData(data: any): void {
  if (!data?.sheets) return;

  const sheets = data.sheets;
  const sheetEntries = typeof sheets === 'object' ? Object.values(sheets) : sheets;

  for (const sheet of sheetEntries) {
    if (!sheet) continue;

    if (sheet.columnData) {
      for (const key of Object.keys(sheet.columnData)) {
        const col = sheet.columnData[key];
        if (col && typeof col.w === 'number' && col.w <= 0) {
          col.w = 72;
        }
      }
    }

    if (sheet.defaultColumnWidth !== undefined && sheet.defaultColumnWidth <= 0) {
      sheet.defaultColumnWidth = 72;
    }

    if (sheet.rowData) {
      for (const key of Object.keys(sheet.rowData)) {
        const row = sheet.rowData[key];
        if (row && typeof row.h === 'number' && row.h <= 0) {
          row.h = 20;
        }
      }
    }
  }
}
