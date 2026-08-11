// ── CLI commands ────────────────────────────────────────────────────
// The command definitions live here rather than beside `runMain` so a
// test can import and run them. `src/cli.ts` is the bin: it does nothing
// but hand `mainCommand` to citty. Before the split the whole CLI sat in
// one module ending in a bare `runMain(main)`, which meant importing it
// *ran* it — reading process.argv and possibly calling process.exit. That
// is why it was the only module in the tree at 0% coverage. See #399.
// ─────────────────────────────────────────────────────────────────────

import { defineCommand } from "citty"
import { consola } from "consola"
import { readFileSync, writeFileSync } from "node:fs"
import { createRequire } from "node:module"
import { extname } from "node:path"
import { readXlsx } from "../xlsx/reader"
import { writeXlsx } from "../xlsx/writer"
import { readXlsb } from "../xlsx/xlsb/reader"
import { readXls } from "../xls/reader"
import { readOds } from "../ods/reader"
import { writeOds } from "../ods/writer"
import { parseCsv } from "../csv/reader"
import { writeCsv } from "../csv/writer"
import { validateWithSchema } from "../_schema"
import type { Workbook, CellValue, WriteOptions, SchemaDefinition } from "../_types"

// ── Errors ──────────────────────────────────────────────────────────

/**
 * A CLI failure with a message already fit for a user.
 *
 * The commands throw this instead of calling `process.exit(1)` directly.
 * citty's `runMain` catches it, prints it and exits non-zero, so the
 * behaviour at the terminal is unchanged — but a test can now assert on
 * the failure instead of watching the test runner die.
 */
export class CliError extends Error {
  override readonly name = "CliError"
}

// ── Helpers ─────────────────────────────────────────────────────────

export type Format = "xlsx" | "ods" | "csv" | "xls" | "xlsb"

/** The formats hucre can write; `.xls` and `.xlsb` are read-only. */
export type WritableFormat = Exclude<Format, "xls" | "xlsb">

/** Text formats carry their separator in the extension, not the format. */
const DELIMITERS: Record<string, string> = { ".csv": ",", ".tsv": "\t" }

const READ_ONLY_FORMATS = new Set<Format>(["xls", "xlsb"])

const SUPPORTED = ".xlsx, .ods, .csv, .tsv (read-only: .xls, .xlsb)"

export function detectFormatFromExtension(filePath: string): Format {
  const ext = extname(filePath).toLowerCase()
  switch (ext) {
    case ".xlsx":
      return "xlsx"
    case ".ods":
      return "ods"
    case ".csv":
    case ".tsv":
      return "csv"
    // The legacy binary readers ship in the library, so the CLI can open
    // these too — `hucre convert legacy.xls out.xlsx` is the whole reason
    // a read-only reader is useful at the terminal. Input only; see
    // detectOutputFormat.
    case ".xls":
      return "xls"
    case ".xlsb":
      return "xlsb"
    default:
      throw new CliError(`Unsupported file extension: ${ext || "(none)"}. Supported: ${SUPPORTED}`)
  }
}

/**
 * Detect the format of a file we are about to *write*.
 *
 * `.xls` and `.xlsb` are readable but not writable, and "unsupported
 * extension" would be a lie about a file the CLI just opened happily —
 * so name the actual reason.
 */
export function detectOutputFormat(filePath: string): WritableFormat {
  const format = detectFormatFromExtension(filePath)
  if (READ_ONLY_FORMATS.has(format)) {
    throw new CliError(
      `Cannot write .${format}: it is a read-only format in hucre — ` +
        "readable as input, but there is no writer for it. " +
        "Write .xlsx, .ods, .csv or .tsv instead.",
    )
  }
  return format as WritableFormat
}

/**
 * `.tsv` is tab-separated. Naming the file `.tsv` and filling it with
 * commas — which is what this used to do, because both extensions mapped
 * to the format "csv" and the writer took its default delimiter — makes a
 * file that lies about itself. The read side never had the bug: parseCsv
 * sniffs the delimiter from the content.
 */
export function delimiterForExtension(filePath: string): string {
  return DELIMITERS[extname(filePath).toLowerCase()] ?? ","
}

export async function readFile(filePath: string, encoding?: string): Promise<Workbook> {
  const format = detectFormatFromExtension(filePath)
  const data = readFileSync(filePath)
  const input = new Uint8Array(data)

  switch (format) {
    case "xlsx":
      return readXlsx(input)
    case "xlsb":
      return readXlsb(input)
    case "xls":
      return readXls(input)
    case "ods":
      return readOds(input)
    case "csv": {
      // The bytes go to parseCsv, which reads the byte-order mark. This
      // used to decode UTF-8 unconditionally, so a CSV out of a Turkish or
      // Central European Excel — windows-1254, windows-1250 — converted to
      // mojibake, silently, from the one place in the library that has the
      // bytes and could know better. `--encoding` covers what no mark can
      // say. See #475.
      const rows = parseCsv(input, {
        delimiter: delimiterForExtension(filePath),
        encoding,
      })
      return {
        sheets: [{ name: "Sheet1", rows }],
      }
    }
  }
}

export function formatCellValue(value: CellValue): string {
  if (value === null || value === undefined) return ""
  if (value instanceof Date) return value.toISOString()
  return String(value)
}

// ── Convert Command ─────────────────────────────────────────────────

export const convertCommand = defineCommand({
  meta: {
    name: "convert",
    description:
      "Convert between spreadsheet formats (cell values only — styles, " +
      "merges, formulas, charts and images are not carried over)",
  },
  args: {
    input: {
      type: "positional",
      description: "Input file path",
      required: true,
    },
    output: {
      type: "positional",
      description: "Output file path",
      required: true,
    },
    encoding: {
      type: "string",
      description:
        "Character encoding of a CSV/TSV input (e.g. windows-1254). " +
        "Default: the file's byte-order mark, or utf-8.",
    },
  },
  async run({ args }) {
    const inputPath = args.input as string
    const outputPath = args.output as string
    const outputFormat = detectOutputFormat(outputPath)

    consola.start(`Reading ${inputPath}...`)
    const workbook = await readFile(inputPath, args.encoding as string | undefined)
    consola.success(`Read ${workbook.sheets.length} sheet(s)`)

    consola.start(`Writing ${outputPath}...`)

    if (outputFormat === "csv") {
      // CSV: use first sheet only
      const sheet = workbook.sheets[0]
      if (!sheet) {
        throw new CliError("No sheets found in input file")
      }
      // Every row goes through writeCsv, including the first. It used to
      // be pulled out as `headers` and stringified separately, so a Date
      // in row 0 came out ISO while the same Date in row 1 came out in
      // writeCsv's format — one column, two formats, decided by which
      // row the value happened to land in.
      const csv = writeCsv(sheet.rows, { delimiter: delimiterForExtension(outputPath) })
      writeFileSync(outputPath, csv, "utf-8")
    } else {
      // XLSX or ODS
      const writeOptions: WriteOptions = {
        sheets: workbook.sheets.map((sheet) => ({
          name: sheet.name,
          rows: sheet.rows,
        })),
        properties: workbook.properties,
      }

      let output: Uint8Array
      if (outputFormat === "ods") {
        output = await writeOds(writeOptions)
      } else {
        output = await writeXlsx(writeOptions)
      }

      writeFileSync(outputPath, output)
    }

    consola.success(`Written to ${outputPath}`)
  },
})

// ── Inspect Command ─────────────────────────────────────────────────

export const inspectCommand = defineCommand({
  meta: {
    name: "inspect",
    description: "Inspect a spreadsheet file",
  },
  args: {
    file: {
      type: "positional",
      description: "File to inspect",
      required: true,
    },
    sheet: {
      type: "string",
      description: "Sheet index to show detailed data (0-based)",
    },
    encoding: {
      type: "string",
      description:
        "Character encoding of a CSV/TSV input (e.g. windows-1254). " +
        "Default: the file's byte-order mark, or utf-8.",
    },
  },
  async run({ args }) {
    const filePath = args.file as string

    consola.start(`Inspecting ${filePath}...`)
    const workbook = await readFile(filePath, args.encoding as string | undefined)

    consola.info(`Sheets: ${workbook.sheets.length}`)

    for (let i = 0; i < workbook.sheets.length; i++) {
      const sheet = workbook.sheets[i]!
      const rowCount = sheet.rows.length
      const colCount = sheet.rows.reduce((max, row) => Math.max(max, row.length), 0)

      // Count cell types
      const typeCounts: Record<string, number> = {}
      for (const row of sheet.rows) {
        for (const cell of row) {
          let type: string
          if (cell === null || cell === undefined) type = "empty"
          else if (typeof cell === "string") type = "string"
          else if (typeof cell === "number") type = "number"
          else if (typeof cell === "boolean") type = "boolean"
          else if (cell instanceof Date) type = "date"
          else type = "unknown"

          typeCounts[type] = (typeCounts[type] ?? 0) + 1
        }
      }

      const typeStr = Object.entries(typeCounts)
        .map(([t, c]) => `${t}: ${c}`)
        .join(", ")

      consola.log(`  [${i}] "${sheet.name}" - ${rowCount} rows x ${colCount} cols (${typeStr})`)
    }

    if (workbook.properties) {
      const props = workbook.properties
      if (props.title) consola.log(`  Title: ${props.title}`)
      if (props.creator) consola.log(`  Creator: ${props.creator}`)
      if (props.created) consola.log(`  Created: ${props.created.toISOString()}`)
    }

    // Show detailed sheet data if --sheet is specified
    if (args.sheet !== undefined) {
      const sheetIdx = Number(args.sheet)
      if (!Number.isInteger(sheetIdx) || sheetIdx < 0 || sheetIdx >= workbook.sheets.length) {
        throw new CliError(
          `Invalid sheet index: ${args.sheet}. Valid range: 0-${workbook.sheets.length - 1}`,
        )
      }

      const sheet = workbook.sheets[sheetIdx]!
      consola.info(`\nSheet "${sheet.name}" (first 10 rows):`)

      const previewRows = sheet.rows.slice(0, 10)
      if (previewRows.length === 0) {
        consola.log("  (empty sheet)")
      } else {
        // Build column widths for formatting
        const maxCols = previewRows.reduce((max, row) => Math.max(max, row.length), 0)
        const colWidths: number[] = Array.from({ length: maxCols }, () => 0)

        const formatted = previewRows.map((row) =>
          Array.from({ length: maxCols }, (_, j) => {
            const val = j < row.length ? formatCellValue(row[j]!) : ""
            const str = val.length > 40 ? `${val.substring(0, 37)}...` : val
            if (str.length > colWidths[j]!) colWidths[j] = str.length
            return str
          }),
        )

        for (let i = 0; i < formatted.length; i++) {
          const row = formatted[i]!
          const line = row.map((val, j) => val.padEnd(colWidths[j]! + 2)).join("")
          consola.log(`  ${String(i).padStart(3)}| ${line}`)
        }

        if (sheet.rows.length > 10) {
          consola.log(`  ... and ${sheet.rows.length - 10} more rows`)
        }
      }
    }
  },
})

// ── Validate Command ────────────────────────────────────────────────

export const validateCommand = defineCommand({
  meta: {
    name: "validate",
    description: "Validate a spreadsheet against a JSON schema",
  },
  args: {
    file: {
      type: "positional",
      description: "Spreadsheet file to validate",
      required: true,
    },
    schema: {
      type: "string",
      description: "Path to JSON schema file",
      required: true,
    },
    sheet: {
      type: "string",
      description: "Sheet index to validate (0-based, default: 0)",
      default: "0",
    },
  },
  async run({ args }) {
    const filePath = args.file as string
    const schemaPath = args.schema as string
    const sheetIdx = Number(args.sheet ?? "0")

    if (!Number.isInteger(sheetIdx)) {
      throw new CliError(`Invalid sheet index: ${args.sheet}`)
    }

    consola.start(`Validating ${filePath} with schema ${schemaPath}...`)

    // Read schema
    const schemaJson = readFileSync(schemaPath, "utf-8")
    let schema: SchemaDefinition
    try {
      schema = JSON.parse(schemaJson) as SchemaDefinition
    } catch {
      throw new CliError(`Invalid JSON schema file: ${schemaPath}`)
    }

    // Read spreadsheet
    const workbook = await readFile(filePath, args.encoding as string | undefined)

    if (sheetIdx < 0 || sheetIdx >= workbook.sheets.length) {
      throw new CliError(
        `Invalid sheet index: ${sheetIdx}. File has ${workbook.sheets.length} sheet(s)`,
      )
    }

    const sheet = workbook.sheets[sheetIdx]!
    // headerRow is 0-based since v1 — the first row (#365).
    const result = validateWithSchema(sheet.rows, schema, { headerRow: 0 })

    if (result.errors.length === 0) {
      consola.success(`Valid! ${result.data.length} row(s) passed validation.`)
      return
    }

    consola.error(`Found ${result.errors.length} error(s) in ${result.data.length} row(s):`)
    for (const err of result.errors.slice(0, 20)) {
      consola.log(`  Row ${err.row}, Column "${err.column}": ${err.message}`)
    }
    if (result.errors.length > 20) {
      consola.log(`  ... and ${result.errors.length - 20} more errors`)
    }
    throw new CliError(`Validation failed with ${result.errors.length} error(s)`)
  },
})

// ── Main Command ────────────────────────────────────────────────────

export const pkgVersion: string = (() => {
  try {
    const require = createRequire(import.meta.url)
    const pkg = require("../../package.json") as { version?: string }
    return pkg.version ?? "0.0.0"
  } catch {
    return "0.0.0"
  }
})()

export const mainCommand = defineCommand({
  meta: {
    name: "hucre",
    version: pkgVersion,
    description:
      "Spreadsheet Swiss Army knife. Convert, inspect, and validate XLSX, CSV, " +
      "and ODS files; reads legacy .xls and .xlsb as input.",
  },
  subCommands: {
    convert: convertCommand,
    inspect: inspectCommand,
    validate: validateCommand,
  },
})
