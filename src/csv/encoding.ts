// ── CSV input decoding ───────────────────────────────────────────────
//
// Every other reader in the library takes bytes. CSV took a string, so
// the byte→string step — which is where the actual difficulty of CSV
// lives — was the caller's alone, and the CLI resolved it by assuming
// UTF-8. A CSV written by Excel on a Turkish, Polish or Greek Windows is
// windows-1254 / 1250 / 1253, and came through as mojibake. See #475.
//
// What this does is decode, not guess. A byte-order mark is a statement
// the file makes about itself, so it is honoured; anything else needs the
// caller to say, because telling windows-1254 from windows-1252 by byte
// frequency is a research problem that is wrong often enough to be
// dangerous, and it has no place in a zero-dependency library.

import { InvalidArgumentError } from "../errors"

/** What a byte-order mark says the file is, if it carries one. */
export type BomEncoding = "utf-8" | "utf-16le" | "utf-16be"

/**
 * Read the byte-order mark, if there is one.
 *
 * `utf-16le` is the one worth knowing about: it is what Excel's "Save as
 * Unicode Text" produces, and decoded as UTF-8 it becomes a run of
 * NUL-separated letters that a CSV parser reads as data rather than
 * rejecting.
 */
export function detectBom(bytes: Uint8Array): { encoding: BomEncoding; length: number } | null {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { encoding: "utf-8", length: 3 }
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return { encoding: "utf-16le", length: 2 }
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return { encoding: "utf-16be", length: 2 }
  }
  return null
}

/** Anything a synchronous CSV reader can be handed. */
export type CsvInput = string | Uint8Array | ArrayBuffer

function toBytes(input: Uint8Array | ArrayBuffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input)
}

/**
 * Turn CSV input into a string.
 *
 * A string passes through untouched — the caller has already decided.
 * Bytes are decoded with, in order: the `encoding` the caller named, the
 * encoding the byte-order mark declares, or UTF-8.
 *
 * `TextDecoder` does the work, so the set of names accepted is the WHATWG
 * Encoding Standard's — every label in it, including the legacy
 * single-byte ones, in every runtime hucre supports.
 */
export function decodeCsvInput(input: CsvInput, encoding?: string): string {
  if (typeof input === "string") return input

  const bytes = toBytes(input)
  const bom = detectBom(bytes)

  // A named encoding wins over the mark. The caller knows something we do
  // not, and a file can carry a mark that is simply wrong.
  const label = encoding ?? bom?.encoding ?? "utf-8"

  let decoder: TextDecoder
  try {
    decoder = new TextDecoder(label)
  } catch (error) {
    throw new InvalidArgumentError(
      `Unknown encoding "${label}". Pass a label from the WHATWG Encoding Standard ` +
        `— "utf-8", "utf-16le", "windows-1254", "iso-8859-9" and so on — or omit ` +
        "`encoding` to use the byte-order mark, or UTF-8 when there is none.",
      { cause: error },
    )
  }

  // TextDecoder strips a UTF-8 BOM itself, but not a UTF-16 one when the
  // caller named the encoding rather than letting the mark speak — and it
  // strips nothing when the mark disagrees with the label. Cutting it here
  // makes the three paths behave the same.
  return decoder.decode(bom ? bytes.subarray(bom.length) : bytes)
}
