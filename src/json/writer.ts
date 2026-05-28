// ── JSON Writer ──────────────────────────────────────────────────────

import type { CellValue, Workbook } from "../_types"

export interface JsonWriteOptions {
  /** Pretty-print with 2-space indent. Default: false. */
  pretty?: boolean
  /** Indent string when `pretty` is true. Default: "  ". */
  indent?: string
  /** Convert `Date` cells to ISO strings. Default: true. */
  isoDates?: boolean
}

/**
 * Serialize an array of row objects to a JSON string.
 */
export function writeJson(data: Record<string, CellValue>[], options?: JsonWriteOptions): string {
  const pretty = options?.pretty ?? false
  const indent = options?.indent ?? "  "
  const isoDates = options?.isoDates ?? true
  return JSON.stringify(data, isoDates ? dateReplacer : undefined, pretty ? indent : undefined)
}

/**
 * Serialize an array of row objects to NDJSON / JSON Lines.
 * One JSON object per line, terminated by `\n`.
 */
export function writeNdjson(
  data: Record<string, CellValue>[],
  options?: { isoDates?: boolean },
): string {
  const isoDates = options?.isoDates ?? true
  if (data.length === 0) return ""
  const replacer = isoDates ? dateReplacer : undefined
  return data.map((row) => JSON.stringify(row, replacer)).join("\n") + "\n"
}

/**
 * Convert a Workbook (e.g. from `readXlsx`) to a JSON string.
 *
 * - Single-sheet workbooks: emit `data` as `[{...}, ...]`
 * - Multi-sheet workbooks: emit `{ "Sheet1": [...], "Sheet2": [...] }`
 *
 * Use `sheet` to pick a specific sheet by index or name.
 */
export interface WorkbookToJsonOptions extends JsonWriteOptions {
  /** Sheet to emit. If omitted, all sheets are emitted as an object. */
  sheet?: number | string
  /** 0-based header row index. Default: 0. */
  headerRow?: number
}

export function workbookToJson(wb: Workbook, options?: WorkbookToJsonOptions): string {
  const headerRow = options?.headerRow ?? 0

  if (options?.sheet !== undefined) {
    const sheet =
      typeof options.sheet === "number"
        ? wb.sheets[options.sheet]
        : wb.sheets.find((s) => s.name === options.sheet)
    if (!sheet) {
      throw new Error(
        typeof options.sheet === "number"
          ? `Sheet index ${options.sheet} out of range`
          : `Sheet "${options.sheet}" not found`,
      )
    }
    return writeJson(sheetToRowObjects(sheet.rows, headerRow), options)
  }

  if (wb.sheets.length === 1) {
    return writeJson(sheetToRowObjects(wb.sheets[0]!.rows, headerRow), options)
  }

  const all: Record<string, Record<string, CellValue>[]> = {}
  for (const sheet of wb.sheets) {
    all[sheet.name] = sheetToRowObjects(sheet.rows, headerRow)
  }

  const pretty = options?.pretty ?? false
  const indent = options?.indent ?? "  "
  const isoDates = options?.isoDates ?? true
  return JSON.stringify(all, isoDates ? dateReplacer : undefined, pretty ? indent : undefined)
}

function sheetToRowObjects(rows: CellValue[][], headerRowIdx: number): Record<string, CellValue>[] {
  if (rows.length <= headerRowIdx) return []
  const headerRow = rows[headerRowIdx]!
  const headers = headerRow.map((h) => (h === null || h === undefined ? "" : String(h).trim()))

  const result: Record<string, CellValue>[] = []
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i]!
    const obj: Record<string, CellValue> = {}
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]!] = j < row.length ? (row[j] ?? null) : null
    }
    result.push(obj)
  }
  return result
}

function dateReplacer(_key: string, value: unknown): unknown {
  if (value instanceof Date) return value.toISOString()
  return value
}
