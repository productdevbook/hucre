import type { CellValue, CsvReadOptions } from "../_types"
import { rowsToObjects } from "../_objects"
import { unescapeFormula } from "./formula"
import { inferType } from "../_infer"
import { decodeCsvInput, type CsvInput } from "./encoding"

// ── Public API ───────────────────────────────────────────────────────

/**
 * Detect and strip BOM (Byte Order Mark) from a string.
 * Handles UTF-8 (EF BB BF), UTF-16 LE (FF FE), UTF-16 BE (FE FF).
 */
export function stripBom(input: string): string {
  if (input.length === 0) return input
  const first = input.charCodeAt(0)
  // UTF-8 BOM: U+FEFF, UTF-16 BE BOM: U+FEFF
  if (first === 0xfeff) return input.slice(1)
  // UTF-16 LE BOM: U+FFFE
  if (first === 0xfffe) return input.slice(1)
  return input
}

/**
 * Auto-detect the delimiter from the first few lines of CSV input.
 * Tests comma, semicolon, tab, and pipe. Picks the one with the most
 * consistent non-zero count across lines.
 */
export function detectDelimiter(input: string): string {
  const candidates = [",", ";", "\t", "|"]
  // Grab up to 10 lines (ignoring quoted fields for speed — good enough for detection)
  const sampleLines = getSampleLines(input, 10)

  if (sampleLines.length === 0) return ","

  let bestDelimiter = ","
  let bestScore = -1

  for (const delim of candidates) {
    const counts = sampleLines.map((line) => countUnquoted(line, delim))
    const nonZero = counts.filter((c) => c > 0)
    if (nonZero.length === 0) continue

    // Consistency = how many lines have the same count as the first non-zero
    const mode = nonZero[0]!
    const consistent = nonZero.filter((c) => c === mode).length
    // Score: prefer higher consistency, then higher count
    const score = consistent * 1000 + mode

    if (score > bestScore) {
      bestScore = score
      bestDelimiter = delim
    }
  }

  return bestDelimiter
}

/**
 * Parse a CSV string into a 2D array of cell values.
 */
export function parseCsv(input: CsvInput, options?: CsvReadOptions): CellValue[][] {
  const opts = normalizeReadOptions(options)

  // Bytes are decoded here, honouring the byte-order mark; a string is
  // whatever the caller already decoded. See ./encoding.ts.
  let text = decodeCsvInput(input, options?.encoding)

  if (opts.skipBom) {
    text = stripBom(text)
  }

  // Skip the first N lines before parsing
  const skipLines = options?.skipLines
  if (skipLines && skipLines > 0) {
    let linesSkipped = 0
    let pos = 0
    while (linesSkipped < skipLines && pos < text.length) {
      const ch = text[pos]!
      if (ch === "\r") {
        linesSkipped++
        if (pos + 1 < text.length && text[pos + 1] === "\n") {
          pos += 2
        } else {
          pos++
        }
      } else if (ch === "\n") {
        linesSkipped++
        pos++
      } else {
        pos++
      }
    }
    text = text.slice(pos)
  }

  if (text.length === 0) return []

  const delimiter = opts.delimiter ?? detectDelimiter(text)
  const quote = opts.quote
  const escape = opts.escape

  let rows: string[][]
  let firstFieldQuoted: boolean[]
  if (options?.fastMode) {
    rows = parseFast(text, delimiter)
    // Fast mode does no quote handling, so no field is ever "quoted".
    firstFieldQuoted = Array.from({ length: rows.length }, () => false)
  } else {
    const parsed = parseRaw(text, delimiter, quote, escape)
    rows = parsed.rows
    firstFieldQuoted = parsed.firstFieldQuoted
  }

  // Filter comments — only physically-unquoted leading comment chars count,
  // so a quoted field like "#not a comment" is preserved.
  const commentChar = opts.comment
  let filtered: CellValue[][] = commentChar
    ? rows.filter((row, idx) => {
        if (row.length === 0) return true
        if (firstFieldQuoted[idx]) return true
        const firstVal = row[0]
        if (typeof firstVal === "string" && firstVal.startsWith(commentChar)) {
          return false
        }
        return true
      })
    : rows

  // Skip empty rows
  if (opts.skipEmptyRows) {
    filtered = filtered.filter(
      (row) => row.length > 0 && !row.every((cell) => cell === null || cell === ""),
    )
  }

  // Undo the writer's formula escape before anything else looks at the
  // values, so type inference and header names see what was written rather
  // than `'` + the value (#408).
  if (opts.unescapeFormulae) {
    filtered = filtered.map((row) =>
      row.map((v) => (typeof v === "string" ? unescapeFormula(v) : v)),
    )
  }

  // `header: true` only marks the first row — it still comes back, and only
  // names columns for transformValue. `skipHeaderRow` is the opt-in that
  // consumes it, honoured by streamCsvRows and, until #408, ignored here.
  // The row is captured before it goes, because transformValue names its
  // columns from it either way, and it drops before maxRows so that limit
  // counts data rows in both readers.
  // `transformHeader` used to be honoured by parseCsvObjects alone, where
  // it renames the object keys — even though CsvReadOptions says every one
  // of its options "means the same thing in all three" readers. Here and in
  // streamCsvRows it rewrites the header row itself, which then names the
  // columns `transformValue` sees. See #439 §V.
  const transformHeader = options?.transformHeader
  if (opts.header && transformHeader && filtered.length > 0) {
    filtered = [
      filtered[0]!.map((value, index) => transformHeader(String(value ?? ""), index)),
      ...filtered.slice(1),
    ]
  }

  const headerRow = opts.header && filtered.length > 0 ? filtered[0]! : null
  if (opts.header && opts.skipHeaderRow && filtered.length > 0) {
    filtered = filtered.slice(1)
  }

  // Limit to maxRows data rows
  if (opts.maxRows !== undefined && opts.maxRows >= 0 && filtered.length > opts.maxRows) {
    filtered = filtered.slice(0, opts.maxRows)
  }

  // Type inference
  if (opts.typeInference) {
    const preserveLeadingZeros = opts.preserveLeadingZeros
    filtered = filtered.map((row) => row.map((v) => inferType(v, preserveLeadingZeros)))
  }

  // transformValue callback — applied after type inference
  const transformValue = options?.transformValue
  if (transformValue) {
    // When we don't have headers we pass column index as the header name;
    // `headerRow` above is the first row when `header` is set.
    filtered = filtered.map((row, rowIdx) =>
      row.map((val, colIdx) => {
        const header = headerRow ? String(headerRow[colIdx] ?? colIdx) : String(colIdx)
        return transformValue(val, header, rowIdx, colIdx)
      }),
    )
  }

  // onRow callback — called for each row after all processing
  const onRow = options?.onRow
  if (onRow) {
    for (let i = 0; i < filtered.length; i++) {
      onRow(filtered[i]!, i)
    }
  }

  return filtered
}

/**
 * Result shape for {@link parseCsvObjects}, mirroring `XlsxObjectsResult`
 * and `OdsObjectsResult`.
 *
 * Named rather than inline so callers can annotate a variable, a function
 * return, or a Promise with it — the anonymous shape could not be spelled
 * at all (#365).
 */
export interface CsvObjectsResult<T extends Record<string, CellValue> = Record<string, CellValue>> {
  data: T[]
  headers: string[]
}

/**
 * Parse CSV with a header row, returning an array of objects
 * and the detected headers.
 */
export function parseCsvObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: CsvInput,
  options?: CsvReadOptions & { header: true },
): CsvObjectsResult<T> {
  // Pass through without transformValue/transformHeader to parseCsv — we handle them here
  const { transformHeader, transformValue, ...restOptions } = options ?? {}
  const rows = parseCsv(input, {
    ...restOptions,
    header: false,
    transformValue: undefined,
    transformHeader: undefined,
  })

  // `skipEmptyRows` is a `parseCsv` option applied above, so the row set
  // handed over here is already filtered — projecting it must not filter
  // a second time.
  return rowsToObjects<T>(rows, {
    headerRow: 0,
    skipEmptyRows: false,
    transformHeader,
    transformValue,
  })
}

// ── Fast parser (no quote handling) ──────────────────────────────────

function parseFast(input: string, delimiter: string): string[][] {
  const rows: string[][] = []
  const lines = input.split(/\r\n|\r|\n/)

  // Drop trailing empty line from trailing newline
  if (lines.length > 0 && lines[lines.length - 1] === "") {
    lines.pop()
  }

  for (const line of lines) {
    rows.push(line.split(delimiter))
  }

  return rows
}

// ── Core parser (RFC 4180) ───────────────────────────────────────────

function parseRaw(
  input: string,
  delimiter: string,
  quote: string,
  escape: string,
): { rows: string[][]; firstFieldQuoted: boolean[] } {
  const rows: string[][] = []
  // Tracks whether the FIRST field of each row was (at least partially) quoted.
  // Used so comment filtering only applies to physically-unquoted leading #.
  const firstFieldQuoted: boolean[] = []
  let currentRow: string[] = []
  let currentField = ""
  let inQuoted = false
  // Whether the field currently being built was opened with a quote char.
  let fieldWasQuoted = false
  let rowFirstQuoted = false
  let i = 0
  const len = input.length

  const pushRow = () => {
    rows.push(currentRow)
    firstFieldQuoted.push(rowFirstQuoted)
  }

  while (i < len) {
    const ch = input[i]!

    if (inQuoted) {
      // Check for escape sequence (doubled quote or escape+quote)
      if (ch === escape && i + 1 < len && input[i + 1] === quote) {
        currentField += quote
        i += 2
        continue
      }
      // End of quoted field
      if (ch === quote) {
        inQuoted = false
        i++
        continue
      }
      // Any other character inside quotes
      currentField += ch
      i++
      continue
    }

    // Not in quoted field

    // Check for delimiter
    if (startsWith(input, delimiter, i)) {
      if (currentRow.length === 0) rowFirstQuoted = fieldWasQuoted
      currentRow.push(currentField)
      currentField = ""
      fieldWasQuoted = false
      i += delimiter.length
      continue
    }

    // Check for line endings
    if (ch === "\r" || ch === "\n") {
      if (currentRow.length === 0) rowFirstQuoted = fieldWasQuoted
      currentRow.push(currentField)
      currentField = ""
      fieldWasQuoted = false
      pushRow()
      currentRow = []
      rowFirstQuoted = false
      // Consume \r\n as single line break
      if (ch === "\r" && i + 1 < len && input[i + 1] === "\n") {
        i += 2
      } else {
        i++
      }
      continue
    }

    // Start of quoted field (only at the start of a field)
    if (ch === quote && currentField === "") {
      inQuoted = true
      fieldWasQuoted = true
      i++
      continue
    }

    // Regular character
    currentField += ch
    i++
  }

  // Handle last field/row.
  // Don't add a trailing empty row from a trailing newline, BUT preserve a
  // final row whose single field was an explicit quoted-empty field ("").
  if (currentField !== "" || currentRow.length > 0 || fieldWasQuoted) {
    if (currentRow.length === 0) rowFirstQuoted = fieldWasQuoted
    currentRow.push(currentField)
    pushRow()
  }

  return { rows, firstFieldQuoted }
}

// ── Helpers ──────────────────────────────────────────────────────────

function normalizeReadOptions(options?: CsvReadOptions) {
  return {
    skipBom: options?.skipBom !== false,
    delimiter: options?.delimiter,
    quote: options?.quote ?? '"',
    escape: options?.escape ?? '"',
    typeInference: options?.typeInference ?? false,
    preserveLeadingZeros: options?.preserveLeadingZeros !== false,
    skipEmptyRows: options?.skipEmptyRows ?? false,
    comment: options?.comment,
    header: options?.header ?? false,
    skipHeaderRow: options?.skipHeaderRow ?? false,
    unescapeFormulae: options?.unescapeFormulae ?? false,
    maxRows: options?.maxRows,
  }
}

/**
 * Get up to `n` sample lines from input, splitting on unquoted newlines.
 * Used for delimiter detection.
 */
function getSampleLines(input: string, n: number): string[] {
  const lines: string[] = []
  let current = ""
  let inQuoted = false
  for (let i = 0; i < input.length && lines.length < n; i++) {
    const ch = input[i]!
    if (inQuoted) {
      if (ch === '"' && i + 1 < input.length && input[i + 1] === '"') {
        current += ch
        i++
        continue
      }
      if (ch === '"') {
        inQuoted = false
        current += ch
        continue
      }
      current += ch
      continue
    }
    if (ch === '"') {
      inQuoted = true
      current += ch
      continue
    }
    if (ch === "\n" || ch === "\r") {
      if (current.length > 0) {
        lines.push(current)
        current = ""
      }
      if (ch === "\r" && i + 1 < input.length && input[i + 1] === "\n") {
        i++
      }
      continue
    }
    current += ch
  }
  if (current.length > 0 && lines.length < n) {
    lines.push(current)
  }
  return lines
}

/**
 * Count occurrences of `delimiter` outside of quoted fields in a single line.
 */
function countUnquoted(line: string, delimiter: string): number {
  let count = 0
  let inQuoted = false
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!
    if (inQuoted) {
      if (ch === '"' && i + 1 < line.length && line[i + 1] === '"') {
        i++
        continue
      }
      if (ch === '"') {
        inQuoted = false
        continue
      }
      continue
    }
    if (ch === '"') {
      inQuoted = true
      continue
    }
    if (startsWith(line, delimiter, i)) {
      count++
      i += delimiter.length - 1
    }
  }
  return count
}

function startsWith(str: string, prefix: string, offset: number): boolean {
  if (offset + prefix.length > str.length) return false
  for (let i = 0; i < prefix.length; i++) {
    if (str[offset + i] !== prefix[i]) return false
  }
  return true
}
