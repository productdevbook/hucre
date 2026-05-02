// ── ReadInput Normalization ────────────────────────────────────────────
//
// Helpers for normalizing the {@link ReadInput} union type into the byte
// buffer that ZIP-based readers (XLSX, ODS, XLSB) need.
//
// `ReadableStream<Uint8Array>` input must be fully buffered because every
// supported container format stores its central directory or directory
// equivalent at the end of the file — true streaming is not possible
// without random access. Buffering happens once and is shared by every
// reader, so the unified `read()` API and `readXlsx`/`readOds` direct
// entry points all accept streams uniformly.
// ──────────────────────────────────────────────────────────────────────

import type { ReadInput } from "./_types";
import { EncryptedFileError, ParseError } from "./errors";

// ── OLE2 / Compound File Binary container detection ──────────────────
//
// Office password-protected XLSX / XLSM / ODS files are not ZIP archives —
// they are OLE2 (a.k.a. Compound File Binary, CFB) containers carrying
// `\EncryptionInfo` and `\EncryptedPackage` streams (MS-OFFCRYPTO
// §2.3.4.x). They start with a fixed 8-byte magic header that has no
// overlap with the ZIP / OOXML / ODF signatures, so a quick byte sniff
// is enough to tell an encrypted container apart from a plain workbook.
//
// We don't actually decrypt the package here — full encryption support
// is tracked in #156. Surfacing the situation early as a typed
// `EncryptedFileError` saves callers from a confusing
// `"not a valid ZIP archive"` ParseError several layers down.

/** OLE2 / CFB magic header: `D0 CF 11 E0 A1 B1 1A E1`. */
const OLE2_MAGIC = Object.freeze([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const);

/**
 * Whether `data` starts with the OLE2 / CFB compound-document magic
 * bytes — the envelope Office uses for password-protected XLSX, XLSM,
 * and ODS files. Returns `false` for shorter buffers and for any other
 * leading bytes (plain ZIP archives, XML, etc.).
 */
export function isOle2Container(data: Uint8Array): boolean {
  if (data.length < OLE2_MAGIC.length) return false;
  for (let i = 0; i < OLE2_MAGIC.length; i++) {
    if (data[i] !== OLE2_MAGIC[i]) return false;
  }
  return true;
}

/**
 * Throw {@link EncryptedFileError} when `data` is an OLE2 / CFB
 * compound-document container — the envelope Office uses for
 * password-protected XLSX / ODS workbooks. No-op for plain ZIP
 * archives. `format` is recorded on the error so callers can
 * distinguish XLSX vs. ODS encryption paths once decryption is wired
 * up (see #156).
 */
export function assertNotEncrypted(data: Uint8Array, format: "xlsx" | "ods"): void {
  if (isOle2Container(data)) {
    throw new EncryptedFileError(format);
  }
}

/**
 * Drain a {@link ReadableStream} of byte chunks into a single
 * {@link Uint8Array}. Allocates only one extra buffer when the stream
 * yields more than one chunk.
 */
export async function bufferReadableStream(
  stream: ReadableStream<Uint8Array>,
): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let totalLen = 0;

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (value) {
      chunks.push(value);
      totalLen += value.length;
    }
  }

  if (chunks.length === 0) return new Uint8Array(0);
  if (chunks.length === 1) return chunks[0]!;

  const result = new Uint8Array(totalLen);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

/**
 * Detect whether a value is a ReadableStream of bytes. Avoids relying on
 * `instanceof ReadableStream` so the check works across realms (Node
 * worker threads, browser iframes, undici, etc.) where multiple
 * `ReadableStream` constructors may exist.
 */
function isReadableStream(value: unknown): value is ReadableStream<Uint8Array> {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as ReadableStream<Uint8Array>).getReader === "function"
  );
}

/**
 * Normalize a {@link ReadInput} into a {@link Uint8Array}. Buffers any
 * `ReadableStream<Uint8Array>` input fully. Throws {@link ParseError}
 * for unsupported input shapes.
 */
export async function readInputToUint8Array(input: ReadInput): Promise<Uint8Array> {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (isReadableStream(input)) return bufferReadableStream(input);
  throw new ParseError(
    "Unsupported input type. Expected Uint8Array, ArrayBuffer, or ReadableStream<Uint8Array>.",
  );
}
