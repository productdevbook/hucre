// ── Sheet Conversion Utilities ──────────────────────────────────────
// Helper functions to convert Sheet data into objects or arrays.

import type { CellValue, Sheet } from "./_types";

/**
 * Convert sheet rows to an array of objects using a row as headers.
 *
 * @param sheet - The sheet to convert
 * @param options.headerRow - 0-based row index to use as headers (default: 0)
 * @returns Array of objects keyed by header values
 */
export function sheetToObjects<T = Record<string, CellValue>>(
  sheet: Sheet,
  options?: { headerRow?: number },
): T[] {
  const headerRowIdx = options?.headerRow ?? 0;

  if (sheet.rows.length <= headerRowIdx) {
    return [];
  }

  const headerRow = sheet.rows[headerRowIdx]!;
  const headers = headerRow.map((h) => {
    if (h === null || h === undefined) return "";
    return String(h).trim();
  });

  const result: T[] = [];
  for (let i = headerRowIdx + 1; i < sheet.rows.length; i++) {
    const row = sheet.rows[i]!;
    const obj: Record<string, CellValue> = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]!] = j < row.length ? (row[j] ?? null) : null;
    }
    result.push(obj as T);
  }

  return result;
}

/**
 * Convert sheet rows to a 2D array with headers extracted from the first row.
 *
 * @param sheet - The sheet to convert
 * @returns Object with `headers` (string[]) and `data` (remaining rows)
 */
export function sheetToArrays(sheet: Sheet): {
  headers: string[];
  data: CellValue[][];
} {
  if (sheet.rows.length === 0) {
    return { headers: [], data: [] };
  }

  const headerRow = sheet.rows[0]!;
  const headers = headerRow.map((h) => {
    if (h === null || h === undefined) return "";
    return String(h).trim();
  });

  const data = sheet.rows.slice(1);
  return { headers, data };
}
