import type { Sheet, CellValue } from "../_types";

export interface JsonExportOptions {
  /** Which row is headers (0-based). Default: 0 */
  headerRow?: number;
  /** Output format. Default: "objects" */
  format?: "objects" | "arrays" | "columns";
  /** Pretty print. Default: false */
  pretty?: boolean;
}

/**
 * Custom replacer for JSON.stringify that converts Date objects to ISO strings.
 */
function dateReplacer(_key: string, value: unknown): unknown {
  if (value instanceof Date) {
    return value.toISOString();
  }
  return value;
}

/**
 * Export a sheet as a JSON string.
 *
 * Formats:
 * - `"objects"` (default): `[{Name:"Widget", Price:9.99}, ...]`
 * - `"arrays"`: `{headers:["Name","Price"], data:[["Widget",9.99], ...]}`
 * - `"columns"`: `{Name:["Widget","Gadget"], Price:[9.99,24.5]}` (columnar)
 */
export function toJson(sheet: Sheet, options?: JsonExportOptions): string {
  const headerRowIdx = options?.headerRow ?? 0;
  const format = options?.format ?? "objects";
  const pretty = options?.pretty ?? false;
  const indent = pretty ? 2 : undefined;

  const rows = sheet.rows;

  if (rows.length === 0) {
    if (format === "arrays") {
      return JSON.stringify({ headers: [], data: [] }, dateReplacer, indent);
    }
    if (format === "columns") {
      return JSON.stringify({}, dateReplacer, indent);
    }
    return JSON.stringify([], dateReplacer, indent);
  }

  // Extract headers
  const rawHeaders = rows[headerRowIdx];
  if (!rawHeaders) {
    if (format === "arrays") {
      return JSON.stringify({ headers: [], data: [] }, dateReplacer, indent);
    }
    if (format === "columns") {
      return JSON.stringify({}, dateReplacer, indent);
    }
    return JSON.stringify([], dateReplacer, indent);
  }

  const headers = rawHeaders.map((h) => {
    if (h === null || h === undefined) return "";
    return String(h).trim();
  });

  // Data rows (everything after the header row)
  const dataRows = rows.slice(headerRowIdx + 1);

  if (format === "arrays") {
    const data: CellValue[][] = dataRows.map((row) => {
      const result: CellValue[] = [];
      for (let j = 0; j < headers.length; j++) {
        result.push(j < row.length ? (row[j] ?? null) : null);
      }
      return result;
    });
    return JSON.stringify({ headers, data }, dateReplacer, indent);
  }

  if (format === "columns") {
    const columns: Record<string, CellValue[]> = {};
    for (const header of headers) {
      columns[header] = [];
    }
    for (const row of dataRows) {
      for (let j = 0; j < headers.length; j++) {
        columns[headers[j]!]!.push(j < row.length ? (row[j] ?? null) : null);
      }
    }
    return JSON.stringify(columns, dateReplacer, indent);
  }

  // Default: "objects"
  const objects: Record<string, CellValue>[] = [];
  for (const row of dataRows) {
    const obj: Record<string, CellValue> = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]!] = j < row.length ? (row[j] ?? null) : null;
    }
    objects.push(obj);
  }
  return JSON.stringify(objects, dateReplacer, indent);
}
