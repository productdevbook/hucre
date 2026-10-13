/**
 * The error tokens Excel can store in a cell (ECMA-376 §18.18.11
 * `ST_CellType` `e`, plus the ones dynamic arrays and data types added).
 * `CellError.error` is typed `string` rather than this union so a token
 * hucre has not met yet still round-trips; this is the documented set.
 */
export type CellErrorCode =
  | "#NULL!"
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#NUM!"
  | "#N/A"
  | "#GETTING_DATA"
  | "#SPILL!"
  | "#CALC!"
  | "#FIELD!"
  | "#CONNECT!"
  | "#BLOCKED!"
  | "#UNKNOWN!"

/**
 * An error value in a cell — `#N/A`, `#DIV/0!`, and the rest.
 *
 * A plain object rather than a class, like everything else in the model,
 * so it survives `structuredClone` and `JSON.stringify`. It exists
 * because a string could not carry the distinction: v1 read `#N/A` back
 * as the text `"#N/A"`, and wrote any string that spelled an error token
 * as an error — so the text `"#N/A"` typed into a cell became `t="e"`,
 * and an error read from one file was indistinguishable from text in the
 * next.
 */
export interface CellError {
  readonly error: string
}

/** Build a {@link CellError}. */
export function cellError(code: CellErrorCode | string): CellError {
  return { error: code }
}

/** Whether a value is a {@link CellError}. */
export function isCellError(value: unknown): value is CellError {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as { error?: unknown }).error === "string"
  )
}
