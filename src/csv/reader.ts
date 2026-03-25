import type { CellValue, CsvReadOptions } from "../_types";

// ── Public API ───────────────────────────────────────────────────────

/**
 * Detect and strip BOM (Byte Order Mark) from a string.
 * Handles UTF-8 (EF BB BF), UTF-16 LE (FF FE), UTF-16 BE (FE FF).
 */
export function stripBom(input: string): string {
  if (input.length === 0) return input;
  const first = input.charCodeAt(0);
  // UTF-8 BOM: U+FEFF, UTF-16 BE BOM: U+FEFF
  if (first === 0xfeff) return input.slice(1);
  // UTF-16 LE BOM: U+FFFE
  if (first === 0xfffe) return input.slice(1);
  return input;
}

/**
 * Auto-detect the delimiter from the first few lines of CSV input.
 * Tests comma, semicolon, tab, and pipe. Picks the one with the most
 * consistent non-zero count across lines.
 */
export function detectDelimiter(input: string): string {
  const candidates = [",", ";", "\t", "|"];
  // Grab up to 10 lines (ignoring quoted fields for speed — good enough for detection)
  const sampleLines = getSampleLines(input, 10);

  if (sampleLines.length === 0) return ",";

  let bestDelimiter = ",";
  let bestScore = -1;

  for (const delim of candidates) {
    const counts = sampleLines.map((line) => countUnquoted(line, delim));
    const nonZero = counts.filter((c) => c > 0);
    if (nonZero.length === 0) continue;

    // Consistency = how many lines have the same count as the first non-zero
    const mode = nonZero[0]!;
    const consistent = nonZero.filter((c) => c === mode).length;
    // Score: prefer higher consistency, then higher count
    const score = consistent * 1000 + mode;

    if (score > bestScore) {
      bestScore = score;
      bestDelimiter = delim;
    }
  }

  return bestDelimiter;
}

/**
 * Parse a CSV string into a 2D array of cell values.
 */
export function parseCsv(input: string, options?: CsvReadOptions): CellValue[][] {
  const opts = normalizeReadOptions(options);

  if (opts.skipBom) {
    input = stripBom(input);
  }

  if (input.length === 0) return [];

  const delimiter = opts.delimiter ?? detectDelimiter(input);
  const quote = opts.quote;
  const escape = opts.escape;

  const rows = parseRaw(input, delimiter, quote, escape);

  // Filter comments
  const commentChar = opts.comment;
  let filtered = commentChar
    ? rows.filter((row) => {
        if (row.length === 0) return true;
        const firstVal = row[0];
        if (typeof firstVal === "string" && firstVal.startsWith(commentChar)) {
          return false;
        }
        return true;
      })
    : rows;

  // Skip empty rows
  if (opts.skipEmptyRows) {
    filtered = filtered.filter(
      (row) => row.length > 0 && !row.every((cell) => cell === null || cell === ""),
    );
  }

  // Limit to maxRows data rows
  if (opts.maxRows !== undefined && opts.maxRows >= 0 && filtered.length > opts.maxRows) {
    filtered = filtered.slice(0, opts.maxRows);
  }

  // Type inference
  if (opts.typeInference) {
    const preserveLeadingZeros = opts.preserveLeadingZeros;
    return filtered.map((row) => row.map((v) => inferType(v, preserveLeadingZeros)));
  }

  return filtered;
}

/**
 * Parse CSV with a header row, returning an array of objects
 * and the detected headers.
 */
export function parseCsvObjects<T extends Record<string, CellValue> = Record<string, CellValue>>(
  input: string,
  options?: CsvReadOptions & { header: true },
): { data: T[]; headers: string[] } {
  const rows = parseCsv(input, { ...options, header: false });

  if (rows.length === 0) {
    return { data: [], headers: [] };
  }

  const headerRow = rows[0]!;
  const headers = headerRow.map((h) => {
    if (h === null) return "";
    return String(h).trim();
  });

  const data: T[] = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]!;
    const obj: Record<string, CellValue> = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]!] = j < row.length ? row[j]! : null;
    }
    data.push(obj as T);
  }

  return { data, headers };
}

// ── Core parser (RFC 4180) ───────────────────────────────────────────

function parseRaw(input: string, delimiter: string, quote: string, escape: string): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = "";
  let inQuoted = false;
  let i = 0;
  const len = input.length;

  while (i < len) {
    const ch = input[i]!;

    if (inQuoted) {
      // Check for escape sequence (doubled quote or escape+quote)
      if (ch === escape && i + 1 < len && input[i + 1] === quote) {
        currentField += quote;
        i += 2;
        continue;
      }
      // End of quoted field
      if (ch === quote) {
        inQuoted = false;
        i++;
        continue;
      }
      // Any other character inside quotes
      currentField += ch;
      i++;
      continue;
    }

    // Not in quoted field

    // Check for delimiter
    if (startsWith(input, delimiter, i)) {
      currentRow.push(currentField);
      currentField = "";
      i += delimiter.length;
      continue;
    }

    // Check for line endings
    if (ch === "\r") {
      currentRow.push(currentField);
      currentField = "";
      rows.push(currentRow);
      currentRow = [];
      // Consume \r\n as single line break
      if (i + 1 < len && input[i + 1] === "\n") {
        i += 2;
      } else {
        i++;
      }
      continue;
    }

    if (ch === "\n") {
      currentRow.push(currentField);
      currentField = "";
      rows.push(currentRow);
      currentRow = [];
      i++;
      continue;
    }

    // Start of quoted field (only at the start of a field)
    if (ch === quote && currentField === "") {
      inQuoted = true;
      i++;
      continue;
    }

    // Regular character
    currentField += ch;
    i++;
  }

  // Handle last field/row
  // Don't add a trailing empty row from a trailing newline
  if (currentField !== "" || currentRow.length > 0) {
    currentRow.push(currentField);
    rows.push(currentRow);
  }

  return rows;
}

// ── Type inference ───────────────────────────────────────────────────

const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;

function inferType(value: CellValue, preserveLeadingZeros: boolean): CellValue {
  if (value === null) return null;
  if (typeof value !== "string") return value;

  const trimmed = value.trim();
  if (trimmed === "") return value;

  // Boolean detection
  const lower = trimmed.toLowerCase();
  if (lower === "true" || lower === "yes") return true;
  if (lower === "false" || lower === "no") return false;

  // ISO 8601 date detection (must come before number to avoid matching partial numbers)
  if (ISO_DATE_RE.test(trimmed)) {
    const d = new Date(trimmed);
    if (!Number.isNaN(d.getTime())) return d;
  }

  // Leading-zero preservation: keep strings like "0123", "007", "00" as strings.
  // Exceptions: "0.xxx" decimals are still parsed.
  if (preserveLeadingZeros && trimmed.length > 1 && trimmed[0] === "0" && trimmed[1] !== ".") {
    return value;
  }

  // Number detection
  const asNumber = parseNumber(trimmed);
  if (asNumber !== null) return asNumber;

  return value;
}

function parseNumber(s: string): number | null {
  // Handle locale-aware numbers like "1,234.56" or "1,234"
  // Strip commas that are thousands separators (followed by 3 digits)
  const stripped = s.replace(/,(\d{3})/g, "$1");
  // Now try parsing
  if (stripped === "" || stripped === "-" || stripped === "+") return null;
  // Must look like a number (avoid parsing random strings)
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(stripped)) return null;
  const n = Number(stripped);
  if (Number.isNaN(n)) return null;
  if (!Number.isFinite(n)) return null;
  return n;
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
    maxRows: options?.maxRows,
  };
}

/**
 * Get up to `n` sample lines from input, splitting on unquoted newlines.
 * Used for delimiter detection.
 */
function getSampleLines(input: string, n: number): string[] {
  const lines: string[] = [];
  let current = "";
  let inQuoted = false;
  for (let i = 0; i < input.length && lines.length < n; i++) {
    const ch = input[i]!;
    if (inQuoted) {
      if (ch === '"' && i + 1 < input.length && input[i + 1] === '"') {
        current += ch;
        i++;
        continue;
      }
      if (ch === '"') {
        inQuoted = false;
        current += ch;
        continue;
      }
      current += ch;
      continue;
    }
    if (ch === '"') {
      inQuoted = true;
      current += ch;
      continue;
    }
    if (ch === "\n" || ch === "\r") {
      if (current.length > 0) {
        lines.push(current);
        current = "";
      }
      if (ch === "\r" && i + 1 < input.length && input[i + 1] === "\n") {
        i++;
      }
      continue;
    }
    current += ch;
  }
  if (current.length > 0 && lines.length < n) {
    lines.push(current);
  }
  return lines;
}

/**
 * Count occurrences of `delimiter` outside of quoted fields in a single line.
 */
function countUnquoted(line: string, delimiter: string): number {
  let count = 0;
  let inQuoted = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]!;
    if (inQuoted) {
      if (ch === '"' && i + 1 < line.length && line[i + 1] === '"') {
        i++;
        continue;
      }
      if (ch === '"') {
        inQuoted = false;
        continue;
      }
      continue;
    }
    if (ch === '"') {
      inQuoted = true;
      continue;
    }
    if (startsWith(line, delimiter, i)) {
      count++;
      i += delimiter.length - 1;
    }
  }
  return count;
}

function startsWith(str: string, prefix: string, offset: number): boolean {
  if (offset + prefix.length > str.length) return false;
  for (let i = 0; i < prefix.length; i++) {
    if (str[offset + i] !== prefix[i]) return false;
  }
  return true;
}
