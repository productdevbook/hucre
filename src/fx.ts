// ── Formula Helpers ─────────────────────────────────────────────────
// Pure functions returning Excel formula strings.
// No classes, no AST — just string builders.

// ── Aggregation ────────────────────────────────────────────────────

/** SUM formula */
export function sum(...refs: string[]): string {
  return `SUM(${refs.join(",")})`;
}

/** AVERAGE formula */
export function average(...refs: string[]): string {
  return `AVERAGE(${refs.join(",")})`;
}

/** COUNT formula (numbers only) */
export function count(...refs: string[]): string {
  return `COUNT(${refs.join(",")})`;
}

/** COUNTA formula (non-empty cells) */
export function countA(...refs: string[]): string {
  return `COUNTA(${refs.join(",")})`;
}

/** MIN formula */
export function min(...refs: string[]): string {
  return `MIN(${refs.join(",")})`;
}

/** MAX formula */
export function max(...refs: string[]): string {
  return `MAX(${refs.join(",")})`;
}

// ── Math ───────────────────────────────────────────────────────────

/** ROUND(expr, digits) */
export function round(expr: string, digits: number): string {
  return `ROUND(${expr},${digits})`;
}

/** ABS(expr) */
export function abs(expr: string): string {
  return `ABS(${expr})`;
}

/** INT(expr) */
export function int(expr: string): string {
  return `INT(${expr})`;
}

/** MOD(number, divisor) */
export function mod(number: string, divisor: string): string {
  return `MOD(${number},${divisor})`;
}

/** Safe division: IF(denominator=0, fallback, numerator/denominator) */
export function safeDiv(
  numerator: string,
  denominator: string,
  fallback: string | number = 0,
): string {
  return `IF(${denominator}=0,${String(fallback)},${numerator}/${denominator})`;
}

/** Percentage: numerator/denominator formatted as fraction */
export function pct(numerator: string, denominator: string): string {
  return safeDiv(numerator, denominator, 0);
}

// ── Logic ──────────────────────────────────────────────────────────

/** IF(condition, thenValue, elseValue) */
export function iif(
  condition: string,
  thenValue: string | number,
  elseValue: string | number = '""',
): string {
  return `IF(${condition},${String(thenValue)},${String(elseValue)})`;
}

/** AND(conditions...) */
export function and(...conditions: string[]): string {
  return `AND(${conditions.join(",")})`;
}

/** OR(conditions...) */
export function or(...conditions: string[]): string {
  return `OR(${conditions.join(",")})`;
}

/** NOT(condition) */
export function not(condition: string): string {
  return `NOT(${condition})`;
}

/** IFERROR(expr, fallback) */
export function ifError(expr: string, fallback: string | number): string {
  return `IFERROR(${expr},${String(fallback)})`;
}

/** IFNA(expr, fallback) */
export function ifNa(expr: string, fallback: string | number): string {
  return `IFNA(${expr},${String(fallback)})`;
}

// ── Text ───────────────────────────────────────────────────────────

/** CONCATENATE(parts...) */
export function concat(...parts: string[]): string {
  return `CONCATENATE(${parts.join(",")})`;
}

/** TEXTJOIN(delimiter, ignoreEmpty, refs...) */
export function textJoin(delimiter: string, ignoreEmpty: boolean, ...refs: string[]): string {
  return `TEXTJOIN("${delimiter}",${ignoreEmpty ? "TRUE" : "FALSE"},${refs.join(",")})`;
}

/** TEXT(value, format) */
export function text(value: string, format: string): string {
  return `TEXT(${value},"${format}")`;
}

// ── Lookup ─────────────────────────────────────────────────────────

/** VLOOKUP(lookupValue, tableArray, colIndex, exactMatch) */
export function vlookup(
  lookupValue: string,
  tableArray: string,
  colIndex: number,
  exactMatch = true,
): string {
  return `VLOOKUP(${lookupValue},${tableArray},${colIndex},${exactMatch ? "FALSE" : "TRUE"})`;
}

/** HLOOKUP(lookupValue, tableArray, rowIndex, exactMatch) */
export function hlookup(
  lookupValue: string,
  tableArray: string,
  rowIndex: number,
  exactMatch = true,
): string {
  return `HLOOKUP(${lookupValue},${tableArray},${rowIndex},${exactMatch ? "FALSE" : "TRUE"})`;
}

/** INDEX(array, rowNum, colNum?) */
export function index(array: string, rowNum: string, colNum?: string): string {
  return colNum ? `INDEX(${array},${rowNum},${colNum})` : `INDEX(${array},${rowNum})`;
}

/** MATCH(lookupValue, lookupArray, matchType) — default exact match */
export function match(lookupValue: string, lookupArray: string, matchType: 0 | 1 | -1 = 0): string {
  return `MATCH(${lookupValue},${lookupArray},${matchType})`;
}

// ── Conditional aggregation ────────────────────────────────────────

/** SUMIF(range, criteria, sumRange?) */
export function sumIf(range: string, criteria: string, sumRange?: string): string {
  return sumRange ? `SUMIF(${range},${criteria},${sumRange})` : `SUMIF(${range},${criteria})`;
}

/** COUNTIF(range, criteria) */
export function countIf(range: string, criteria: string): string {
  return `COUNTIF(${range},${criteria})`;
}

/** AVERAGEIF(range, criteria, averageRange?) */
export function averageIf(range: string, criteria: string, averageRange?: string): string {
  return averageRange
    ? `AVERAGEIF(${range},${criteria},${averageRange})`
    : `AVERAGEIF(${range},${criteria})`;
}

// ── Column reference helper ────────────────────────────────────────

/**
 * Returns a function that creates a cell reference for a given row.
 * Useful in ColumnDef.formula callbacks.
 *
 * @example
 * ```ts
 * const C = fx.col("C")
 * // formula: (row) => `${C(row)}*2`  → "C5*2" for row 5
 * ```
 */
export function col(letter: string): (row: number) => string {
  return (row) => `${letter}${row}`;
}
