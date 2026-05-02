// ── Ergonomic API ───────────────────────────────────────────────────
// Unified high-level functions that wrap the format-specific readers/writers.
// Auto-detects format from content (magic bytes) for reading, and dispatches
// to the correct writer based on the `format` option for writing.
// ─────────────────────────────────────────────────────────────────────

import type {
  Workbook,
  ReadOptions,
  WriteOptions,
  WriteOutput,
  CellValue,
  ReadInput,
  TableDefinition,
  TableColumn,
} from "./_types";
import { readXlsx } from "./xlsx/reader";
import { writeXlsx } from "./xlsx/writer";
import { readOds } from "./ods/reader";
import { writeOds } from "./ods/writer";
import { UnsupportedFormatError } from "./errors";
import { readInputToUint8Array } from "./_input";

// ── Format Detection ────────────────────────────────────────────────

/**
 * Detect whether a ZIP archive is XLSX or ODS by inspecting the first
 * local file entry. ODS archives store "mimetype" as the first file
 * with content "application/vnd.oasis.opendocument.spreadsheet".
 * XLSX archives are also ZIP but never have "mimetype" as the first entry.
 */
function detectFormat(data: Uint8Array): "xlsx" | "ods" {
  // Both XLSX and ODS start with PK (ZIP magic: 0x504B0304)
  if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4b) {
    throw new UnsupportedFormatError("unknown (not a ZIP archive)");
  }

  // Read the first local file header to get the filename
  // Local file header: offset 26 = filename length (2 bytes LE), offset 30+ = filename
  if (data.length < 30) {
    throw new UnsupportedFormatError("unknown (ZIP too short)");
  }

  const filenameLen = data[26]! | (data[27]! << 8);
  if (data.length < 30 + filenameLen) {
    throw new UnsupportedFormatError("unknown (ZIP truncated)");
  }

  const decoder = new TextDecoder("utf-8");
  const firstName = decoder.decode(data.subarray(30, 30 + filenameLen));

  if (firstName === "mimetype") {
    // Read the extra field length to find where file data starts
    const extraLen = data[28]! | (data[29]! << 8);
    const dataOffset = 30 + filenameLen + extraLen;

    // Read the uncompressed size from the local header (offset 22, 4 bytes LE)
    const uncompSize = data[22]! | (data[23]! << 8) | (data[24]! << 16) | (data[25]! << 24);

    if (uncompSize > 0 && data.length >= dataOffset + uncompSize) {
      const mimeContent = decoder.decode(data.subarray(dataOffset, dataOffset + uncompSize));
      if (mimeContent.trim() === "application/vnd.oasis.opendocument.spreadsheet") {
        return "ods";
      }
    }

    // Even if we couldn't read the content, "mimetype" as first entry is ODS convention
    return "ods";
  }

  // Default: assume XLSX for any other ZIP
  return "xlsx";
}

// ── Public API ──────────────────────────────────────────────────────

/**
 * Read any supported spreadsheet file. Auto-detects format from content.
 * Supports: XLSX, ODS (CSV uses parseCsv separately since it's string input).
 *
 * Input can be Uint8Array, ArrayBuffer, or ReadableStream&lt;Uint8Array&gt;.
 * ReadableStream input is buffered fully before format detection runs.
 */
export async function read(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = await readInputToUint8Array(input);
  const format = detectFormat(data);

  if (format === "ods") {
    return readOds(data, options);
  }
  return readXlsx(data, options);
}

/**
 * Write a workbook to the specified format.
 */
export async function write(
  options: WriteOptions & { format?: "xlsx" | "ods" },
): Promise<WriteOutput> {
  const format = options.format ?? "xlsx";
  if (format === "ods") {
    return writeOds(options);
  }
  return writeXlsx(options);
}

/**
 * Quick helper: read a file and get the first sheet as array of objects.
 * Assumes first row is headers.
 */
export async function readObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: ReadInput,
  options?: ReadOptions,
): Promise<T[]> {
  const workbook = await read(input, options);

  if (workbook.sheets.length === 0) {
    return [];
  }

  const sheet = workbook.sheets[0]!;
  const rows = sheet.rows;

  if (rows.length === 0) {
    return [];
  }

  // First row is headers
  const headerRow = rows[0]!;
  const headers = headerRow.map((h) => {
    if (h === null || h === undefined) return "";
    return String(h).trim();
  });

  if (headers.length === 0) {
    return [];
  }

  const data: T[] = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]!;
    const obj: Record<string, CellValue> = {};
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j]!;
      if (key === "") continue;
      obj[key] = j < row.length ? (row[j] ?? null) : null;
    }
    data.push(obj as T);
  }

  return data;
}

/** Options for writeObjects table generation */
export interface WriteObjectsTableOption {
  /** Table name (must be unique in workbook) */
  name: string;
  /** Table style (e.g. "TableStyleMedium2") */
  style?: string;
  /** Show totals row */
  showTotalRow?: boolean;
  /** Show auto-filter. Default: true */
  showAutoFilter?: boolean;
  /** Show banded rows. Default: true */
  showRowStripes?: boolean;
  /** Totals per column key: { revenue: "sum", margin: "average" } */
  totals?: Record<
    string,
    "sum" | "average" | "count" | "min" | "max" | "countNums" | "stdDev" | "var"
  >;
}

/**
 * Quick helper: write an array of objects to a spreadsheet format.
 * Infers column headers from the keys of the first object.
 */
export async function writeObjects(
  data: Array<Record<string, CellValue>>,
  options?: {
    sheetName?: string;
    format?: "xlsx" | "ods";
    /** Wrap output in a native Excel table (ListObject) */
    table?: WriteObjectsTableOption;
  },
): Promise<WriteOutput> {
  const sheetName = options?.sheetName ?? "Sheet1";
  const format = options?.format ?? "xlsx";

  if (data.length === 0) {
    return write({
      sheets: [{ name: sheetName, rows: [] }],
      format,
    });
  }

  // Infer columns from first object's keys
  const keys = Object.keys(data[0]!);

  // Build rows: header row + data rows
  const rows: CellValue[][] = [];

  // Header row
  rows.push(keys);

  // Data rows
  for (const item of data) {
    const row: CellValue[] = keys.map((key) => {
      const val = item[key];
      return val === undefined ? null : val;
    });
    rows.push(row);
  }

  // Build Excel table if requested
  let tables: TableDefinition[] | undefined;
  if (options?.table) {
    const t = options.table;
    const colCount = keys.length;
    const rowCount = data.length + 1; // +1 for header
    const endCol = colToLetterSimple(colCount - 1);
    const range = `A1:${endCol}${rowCount + (t.showTotalRow ? 1 : 0)}`;

    const tableColumns: TableColumn[] = keys.map((key) => {
      const totalFn = t.totals?.[key];
      return {
        name: key,
        ...(totalFn ? { totalFunction: totalFn } : {}),
      };
    });

    tables = [
      {
        name: t.name,
        displayName: t.name,
        range,
        columns: tableColumns,
        style: t.style,
        showAutoFilter: t.showAutoFilter,
        showRowStripes: t.showRowStripes,
        showTotalRow: t.showTotalRow,
      },
    ];
  }

  return write({
    sheets: [{ name: sheetName, rows, tables }],
    format,
  });
}

/** Simple column index to letter (0-based) */
function colToLetterSimple(col: number): string {
  let result = "";
  let n = col;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}
