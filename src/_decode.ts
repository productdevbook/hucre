// ── Decoding a package part ─────────────────────────────────────────
//
// One JavaScript string cannot hold an arbitrarily large part. V8's
// ceiling is 0x1fffffe8 = 536,870,888 characters, about 512 MB, and a
// worksheet above it cannot be turned into a string at all — the
// buffered readers die before any parsing begins, with
//
//   Error: Cannot create a string longer than 0x1fffffe8 characters
//
// which is not a `ParseError`, names no part, and arrives after however
// long the decompression took. Instrument logs at Excel's row limit
// reach it: three files in a corpus of ~600 were 56–99 MB compressed and
// 607 MB expanded. See #503.
//
// This is not "hucre cannot read big files" — `streamXlsxRows` reads
// exactly those files, in 30s at a flat 944 MB. It is that the buffered
// reader had a hard ceiling it did not know about and reported hitting
// it in the least useful way available.

import { ParseError } from "./errors"

/**
 * The largest string this runtime will build, as far as we can tell.
 *
 * V8's limit and the number in the error it throws. Other engines differ
 * — JavaScriptCore's is larger — which is why the check is a `catch`
 * rather than a comparison: a guess about the ceiling would refuse files
 * a runtime could actually handle.
 */
export const MAX_STRING_LENGTH = 0x1fffffe8

/**
 * The message for a part too large to become a string.
 *
 * Separated from the decode so it can be tested: reproducing the
 * condition needs a part over 512 MB, which is larger than this
 * repository, so the message and the decision are what a test can reach.
 * It follows the `maxTotalCells` error — name the bound, the
 * measurement, and the way out.
 */
export function tooLargeToDecode(path: string, byteLength: number): ParseError {
  return new ParseError(
    `Part "${path}" is ${byteLength.toLocaleString("en-US")} bytes, over the ` +
      `${MAX_STRING_LENGTH.toLocaleString("en-US")}-character maximum string length ` +
      `this runtime supports, so it cannot be read into memory whole.\n` +
      `  - streamXlsxRows(input) reads a worksheet of any size, a row at a time.\n` +
      `  - The workbook is not damaged; this is a limit of the buffered reader.`,
  )
}

/**
 * Is this the string-length ceiling, and not some other failure?
 *
 * Three signals, because no one of them holds everywhere:
 *
 * - **Node throws a plain `Error`** carrying `code: "ERR_STRING_TOO_LONG"`.
 *   `TextDecoder.decode` goes through Node's internal encoding layer,
 *   which wraps V8's failure rather than passing it through. The first
 *   version of this checked `instanceof RangeError` — which is what a
 *   plain string concatenation throws — so on Node the guard never fired
 *   and the raw error reached the caller exactly as before. See #516.
 * - **Other engines throw a `RangeError`**, so that stays.
 * - **The byte length is a necessary condition either way.** A string
 *   cannot exceed {@link MAX_STRING_LENGTH} characters unless the buffer
 *   exceeds it in bytes, since UTF-8 uses at least one byte per
 *   character. So a failure on a buffer that large is this failure,
 *   whatever the engine chose to call it — which is the backstop that
 *   does not depend on guessing at engine internals.
 */
function isStringLengthCeiling(error: unknown, byteLength: number): boolean {
  if (error instanceof RangeError) return true
  if ((error as { code?: unknown } | null)?.code === "ERR_STRING_TOO_LONG") return true
  return byteLength > MAX_STRING_LENGTH
}

/**
 * Decode a package part as UTF-8, or say why it cannot be.
 *
 * `path` is only used for the message, so a caller that has not resolved
 * one can pass whatever it knows.
 */
export function decodePart(data: Uint8Array, path: string): string {
  try {
    return new TextDecoder("utf-8").decode(data)
  } catch (error) {
    // Anything that is not the ceiling is not ours to reinterpret.
    if (isStringLengthCeiling(error, data.length)) throw tooLargeToDecode(path, data.length)
    throw error
  }
}
