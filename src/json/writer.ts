// ── JSON Writer ──────────────────────────────────────────────────────

import { ParseError } from "../errors"
import { isCellError } from "../cell-error"
import type { CellValue, Workbook } from "../_types"
import { unflattenRow } from "./unflatten"

/**
 * `Date` values always serialize as ISO strings.
 *
 * There used to be an `isoDates` option here, and it never did anything:
 * `JSON.stringify` calls `Date.prototype.toJSON` *before* consulting the
 * replacer, so a replacer testing `value instanceof Date` is never
 * reached. `isoDates: false` produced byte-identical output. Removed
 * before v1 rather than frozen — and there is no honest alternative
 * behaviour to give it, since JSON cannot carry a Date at all.
 */
/** JSON carries values; an error is written as its token, as CSV does. */
function errorReplacer(_key: string, value: unknown): unknown {
  return isCellError(value) ? value.error : value
}

export interface JsonWriteOptions {
  /** Pretty-print with 2-space indent. Default: false. */
  pretty?: boolean
  /** Indent string when `pretty` is true. Default: "  ". */
  indent?: string
  /**
   * Rebuild dot-path keys into nested objects — the inverse of the reader's
   * `flatten`. Default: **false**.
   *
   * Opt-in, not the default, and the asymmetry is on purpose. `writeJson`
   * takes any flat row set, most of which never went through `flatten`: a
   * CSV read, a sheet read, a hand-built array. Spreadsheet headers contain
   * dots routinely — `Q1.2024`, `v1.2`, `Rate.%` — and turning those into
   * nested objects by default would be a new silent mangling introduced by
   * the fix for a silent mangling. It also cannot be proven safe from the
   * flat data alone, because `flatten` does not escape dots that were
   * already in a key.
   *
   * Turn it on when you flattened on the way in and want the nesting back:
   * `writeJson(parseJson(text).data, { unflatten: true })`.
   *
   * See {@link unflattenRow} for the collision and numeric-segment rules.
   */
  unflatten?: boolean
}

/** Apply the `unflatten` option, or hand the rows straight through. */
function prepare(
  data: Record<string, CellValue>[],
  options?: JsonWriteOptions,
): readonly unknown[] {
  return options?.unflatten ? data.map(unflattenRow) : data
}

/**
 * Serialize an array of row objects to a JSON string.
 */
export function writeJson(data: Record<string, CellValue>[], options?: JsonWriteOptions): string {
  const pretty = options?.pretty ?? false
  const indent = options?.indent ?? "  "
  return JSON.stringify(prepare(data, options), errorReplacer, pretty ? indent : undefined)
}

/**
 * Serialize an array of row objects to NDJSON / JSON Lines.
 * One JSON object per line, terminated by `\n`.
 */
export function writeNdjson(
  data: Record<string, CellValue>[],
  options?: Pick<JsonWriteOptions, "unflatten">,
): string {
  if (data.length === 0) return ""
  return (
    prepare(data, options)
      .map((row) => JSON.stringify(row, errorReplacer))
      .join("\n") + "\n"
  )
}

/**
 * Convert a Workbook (e.g. from `readXlsx`) to a JSON string.
 *
 * Use `sheet` to pick a specific sheet by index or name, and `shape` to
 * decide whether the output shape may depend on how many sheets there are.
 */
export interface WorkbookToJsonOptions extends JsonWriteOptions {
  /** Sheet to emit. If omitted, all sheets are emitted. */
  sheet?: number | string
  /** 0-based header row index. Default: 0. */
  headerRow?: number
  /**
   * Output shape when no `sheet` is picked. Default: `"auto"`.
   *
   * - `"auto"` — a one-sheet workbook emits a bare `[{...}]`; any other count
   *   emits `{ "Sheet1": [...], "Sheet2": [...] }`. Convenient, but the shape
   *   is a function of the *data*, so a consumer written against a one-sheet
   *   export breaks the day a second sheet appears.
   * - `"sheets"` — always the keyed object, whatever the sheet count. Pick
   *   this when something downstream has to parse the result.
   *
   * `jsonToWorkbook` reads both.
   */
  shape?: "auto" | "sheets"
}

export function workbookToJson(wb: Workbook, options?: WorkbookToJsonOptions): string {
  const headerRow = options?.headerRow ?? 0

  if (options?.sheet !== undefined) {
    const sheet =
      typeof options.sheet === "number"
        ? wb.sheets[options.sheet]
        : wb.sheets.find((s) => s.name === options.sheet)
    if (!sheet) {
      throw new ParseError(
        typeof options.sheet === "number"
          ? `Sheet index ${options.sheet} out of range`
          : `Sheet "${options.sheet}" not found`,
      )
    }
    return writeJson(sheetToRowObjects(sheet.rows, headerRow), options)
  }

  if ((options?.shape ?? "auto") === "auto" && wb.sheets.length === 1) {
    return writeJson(sheetToRowObjects(wb.sheets[0]!.rows, headerRow), options)
  }

  // Null-prototype for the same reason flatten.ts uses one: a sheet may
  // legally be named `__proto__`, and on a plain object that key hits the
  // prototype setter and the sheet vanishes from the output entirely.
  const all: Record<string, unknown[]> = Object.create(null)
  for (const sheet of wb.sheets) {
    const rows = sheetToRowObjects(sheet.rows, headerRow)
    all[sheet.name] = options?.unflatten ? rows.map(unflattenRow) : rows
  }

  const pretty = options?.pretty ?? false
  const indent = options?.indent ?? "  "
  return JSON.stringify(all, errorReplacer, pretty ? indent : undefined)
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
