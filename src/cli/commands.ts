// ── CLI commands ────────────────────────────────────────────────────
// The command definitions live here rather than beside `runMain` so a
// test can import and run them. `src/cli.ts` is the bin: it does nothing
// but hand `mainCommand` to citty. Before the split the whole CLI sat in
// one module ending in a bare `runMain(main)`, which meant importing it
// *ran* it — reading process.argv and possibly calling process.exit. That
// is why it was the only module in the tree at 0% coverage. See #399.
// ─────────────────────────────────────────────────────────────────────

import { toWriteOptions } from "../write-model"
import { isCellError } from "../cell-error"
import { defineCommand } from "citty"
import { consola } from "consola"
import { readFileSync, writeFileSync } from "node:fs"
import { createRequire } from "node:module"
import { extname } from "node:path"
import { readXlsx } from "../xlsx/reader"
import { readXlsb } from "../xlsx/xlsb/reader"
import { readXls } from "../xls/reader"
import { readOds } from "../ods/reader"
import { parseCsv } from "../csv/reader"
import { validateWithSchema } from "../_schema"
import { read, write } from "../defter"
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

export type Format =
  | "xlsx"
  | "ods"
  | "csv"
  | "xls"
  | "xlsb"
  | "json"
  | "ndjson"
  | "xml"
  | "html"
  | "markdown"

/** The formats hucre can write; `.xls` and `.xlsb` are read-only. */
export type WritableFormat = Exclude<Format, "xls" | "xlsb">

/** Text formats carry their separator in the extension, not the format. */
const DELIMITERS: Record<string, string> = { ".csv": ",", ".tsv": "\t" }

const READ_ONLY_FORMATS = new Set<Format>(["xls", "xlsb"])

/**
 * The stdin/stdout convention. `hucre convert - out.xlsx` reads the
 * pipe; `hucre convert in.csv -` writes to it. Most of what a CLI is for
 * was unreachable without this. See #469.
 */
export const STDIO = "-"

const SUPPORTED =
  ".xlsx, .ods, .csv, .tsv, .json, .ndjson, .jsonl, .xml, .html, .md " + "(read-only: .xls, .xlsb)"

/**
 * Map a file extension to a format.
 *
 * JSON, NDJSON, XML, HTML and Markdown all had readers and/or writers in
 * the library and none was reachable from the terminal. See #469.
 */
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
    case ".json":
      return "json"
    case ".ndjson":
    case ".jsonl":
      return "ndjson"
    case ".xml":
      return "xml"
    case ".html":
    case ".htm":
      return "html"
    case ".md":
    case ".markdown":
      return "markdown"
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
        `Write one of: ${SUPPORTED.replace(/ \(read-only.*$/, "")}.`,
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

/**
 * Read every byte from stdin.
 *
 * `readFileSync(0)` is the whole of it — the pipe is a file descriptor
 * and Node reads it to EOF. See #469.
 */
export function readStdin(): Uint8Array {
  return new Uint8Array(readFileSync(0))
}

export async function readFile(filePath: string, encoding?: string): Promise<Workbook> {
  // From a pipe there is no extension to go on, so the content decides —
  // which is what `read()` is for, and it now covers the text formats
  // too (#469).
  if (filePath === STDIO) return read(readStdin())

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
    // The remaining text formats have no extension-specific handling the
    // library does not already do — `read()` sniffs the same bytes and
    // dispatches to the same reader.
    case "json":
    case "ndjson":
    case "xml":
    case "html":
      return read(input)
    case "markdown":
      throw new CliError(
        "Markdown is output only in hucre — there is no `fromMarkdown`, " +
          "and there will not be. Use .csv or .json to bring data back in.",
      )
  }
}

export function formatCellValue(value: CellValue): string {
  if (value === null || value === undefined) return ""
  if (value instanceof Date) return value.toISOString()
  if (isCellError(value)) return value.error
  return String(value)
}

// ── Convert Command ─────────────────────────────────────────────────

export const convertCommand = defineCommand({
  meta: {
    name: "convert",
    description:
      "Convert between spreadsheet formats. Everything the authoring model " +
      "carries — styles, merges, formulas, validations — goes with it; " +
      "text formats take values only.",
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
    to: {
      type: "string",
      description:
        "Output format when writing to stdout (`-`), which has no " +
        "extension to read. E.g. `--to csv`.",
    },
    bom: {
      type: "boolean",
      description:
        "Start CSV/TSV output with a UTF-8 byte-order mark. Excel needs " +
        "it to read a UTF-8 CSV as UTF-8 on a non-UTF-8 locale; without " +
        "it the accented characters arrive as mojibake.",
    },
  },
  async run({ args }) {
    const inputPath = args.input as string
    const outputPath = args.output as string
    // Writing to a pipe has no extension either, and guessing would be
    // worse than asking: `--to` names the format. See #469.
    const outputFormat =
      outputPath === STDIO
        ? stdoutFormat(args.to as string | undefined)
        : detectOutputFormat(outputPath)

    // Progress goes to stderr, always — with `-` as the output, stdout is
    // the file, and a "Reading..." line in it would corrupt the result.
    const toStdout = outputPath === STDIO
    const say = (fn: (m: string) => void, message: string): void => {
      if (!toStdout) fn(message)
    }

    say(consola.start, `Reading ${inputPath === STDIO ? "stdin" : inputPath}...`)
    const workbook = await readFile(inputPath, args.encoding as string | undefined)
    say(consola.success, `Read ${workbook.sheets.length} sheet(s)`)

    say(consola.start, `Writing ${outputPath}...`)

    const output = await renderWorkbook(workbook, outputFormat, outputPath, args.bom === true)
    if (toStdout) writeFileSync(1, output)
    else writeFileSync(outputPath, output)

    say(consola.success, `Written to ${outputPath}`)
  },
})

/**
 * Resolve the format for `-` output. There is no extension to read, and
 * defaulting to a binary format would spray a ZIP into a terminal.
 */
function stdoutFormat(to: string | undefined): WritableFormat {
  if (!to) {
    throw new CliError(
      "Writing to stdout needs `--to` to say which format " +
        "(e.g. `--to csv`), because there is no extension to read.",
    )
  }
  // `.x` so the same table decides, and the same error names the same
  // supported list.
  const format = detectFormatFromExtension(`out.${to.toLowerCase()}`)
  if (READ_ONLY_FORMATS.has(format)) {
    throw new CliError(`Cannot write ${to}: it is a read-only format in hucre.`)
  }
  return format as WritableFormat
}

/**
 * Render a workbook to the bytes of one format.
 *
 * CSV keeps its own path because the delimiter comes from the output
 * *extension* rather than the format — `.tsv` is tab-separated, and a
 * file named `.tsv` full of commas lies about itself (#365). Everything
 * else goes through the library's own `write()`, which is the function
 * that is supposed to know how to do this.
 */
async function renderWorkbook(
  workbook: Workbook,
  format: WritableFormat,
  outputPath: string,
  bom: boolean,
): Promise<Uint8Array> {
  if (format === "csv" && !workbook.sheets[0]) {
    throw new CliError("No sheets found in input file")
  }

  // The whole authoring model, not `{ name, rows }`: an xlsx → xlsx
  // conversion used to drop every style, merge and formula on the floor.
  const writeOptions: WriteOptions = toWriteOptions(workbook)

  // Every row goes through writeCsv, including the first. It used to be
  // pulled out as `headers` and stringified separately, so a Date in row 0
  // came out ISO while the same Date in row 1 came out in writeCsv's
  // format — one column, two formats, decided by which row the value
  // happened to land in.
  return write({
    ...writeOptions,
    format,
    csv: { delimiter: delimiterForExtension(outputPath), bom },
  })
}

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
          else if (isCellError(cell)) type = "error"
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
    headerRow: {
      type: "string",
      description: "0-based index of the header row; -1 for no header row (default: 0)",
      default: "0",
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
    const schemaPath = args.schema as string
    const sheetIdx = Number(args.sheet ?? "0")
    const headerRow = Number(args.headerRow ?? "0")

    if (!Number.isInteger(sheetIdx)) {
      throw new CliError(`Invalid sheet index: ${args.sheet}`)
    }
    if (!Number.isInteger(headerRow) || headerRow < -1) {
      throw new CliError(`Invalid header row: ${args.headerRow}`)
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
    const result = validateWithSchema(sheet.rows, schema, { headerRow })

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
