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
  ColumnDef,
} from "./_types";
import { readXlsx } from "./xlsx/reader";
import { writeXlsx } from "./xlsx/writer";
import { readOds } from "./ods/reader";
import { writeOds } from "./ods/writer";
import { UnsupportedFormatError } from "./errors";

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

// ── Helpers ─────────────────────────────────────────────────────────

function toUint8Array(input: ReadInput): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  throw new UnsupportedFormatError(
    "ReadableStream input is not supported by the unified read() API. Use readXlsx/readOds directly.",
  );
}

// ── Public API ──────────────────────────────────────────────────────

/**
 * Read any supported spreadsheet file. Auto-detects format from content.
 * Supports: XLSX, ODS (CSV uses parseCsv separately since it's string input).
 */
export async function read(input: ReadInput, options?: ReadOptions): Promise<Workbook> {
  const data = toUint8Array(input);
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

/**
 * Quick helper: write an array of objects to a spreadsheet format.
 *
 * When `columns` is provided, supports value accessors (dot-path, functions),
 * transforms, formulas, summary rows, conditional styles, column groups, and more.
 * When omitted, infers columns from the first object's keys.
 */
export async function writeObjects<T extends Record<string, unknown> = Record<string, CellValue>>(
  data: T[],
  options?: {
    sheetName?: string;
    format?: "xlsx" | "ods";
    columns?: ColumnDef<T>[];
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

  // If columns provided, use data+columns path for full ColumnDef support.
  // Otherwise, infer columns from first object's keys.
  const columns: ColumnDef[] = options?.columns
    ? (options.columns as ColumnDef[])
    : Object.keys(data[0]!).map((key) => ({ key, header: key }));

  return write({
    sheets: [{ name: sheetName, data: data as Array<Record<string, unknown>>, columns }],
    format,
  });
}
