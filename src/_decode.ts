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
 * The error a part over the string ceiling produces.
 *
 * A caller that can do something about the ceiling — the worksheet
 * reader, which parses the part as a stream instead (#503) — has to
 * recognise this one error among all the `ParseError`s a decode can
 * raise. Matching its message would be the same mistake #516 was: the
 * text is for people, and a reworded sentence would silently turn the
 * recovery off. A class cannot be reworded, and it is how every other
 * distinguishable condition in this library is told apart — see
 * `errors.ts`.
 *
 * Unexported, and a `ParseError` first: nothing outside needs to name it
 * yet, and every existing `catch (e) { e instanceof ParseError }` keeps
 * working.
 */
class PartTooLargeError extends ParseError {}

/** Whether `error` is the "part is over the string ceiling" `ParseError`. */
export function isPartTooLargeToDecode(error: unknown): boolean {
  return error instanceof PartTooLargeError
}

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
  return new PartTooLargeError(
    `Part "${path}" is ${byteLength.toLocaleString("en-US")} bytes, over the ` +
      `${MAX_STRING_LENGTH.toLocaleString("en-US")}-character maximum string length ` +
      `this runtime supports, so it cannot be read into memory whole.\n` +
      `  - streamXlsxRows(input) reads a worksheet of any size, a row at a time.\n` +
      `  - The workbook is not damaged; this is a limit of the buffered reader.`,
  )
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
    // The ceiling does not announce itself the same way twice. V8 raises
    // `RangeError` for a plain string operation, but Node's TextDecoder
    // goes through its own encoding layer and throws a *plain* `Error`
    // carrying `code: "ERR_STRING_TOO_LONG"` — so guarding on RangeError
    // alone converted nothing on the one runtime this was reported
    // against, and every file in #503 kept getting the raw V8 error. See
    // #516.
    //
    // The length test is the backstop for whatever the next engine does.
    // It is only consulted once the decode has already failed, so it
    // cannot refuse a part a roomier engine would have decoded — it just
    // stops the classification depending on which wrapper a runtime
    // happened to choose. A short part that fails is still a genuine
    // decode error and is passed through untouched.
    const isCeiling =
      error instanceof RangeError ||
      (error as { code?: string } | null)?.code === "ERR_STRING_TOO_LONG" ||
      data.length > MAX_STRING_LENGTH
    if (isCeiling) throw tooLargeToDecode(path, data.length)
    throw error
  }
}
