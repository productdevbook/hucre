import type { SheetImage } from "./_types"
import { InvalidArgumentError } from "./errors"

/**
 * Decode a base64 string to a `Uint8Array`.
 *
 * `atob` is a Web API and is in every runtime this library claims —
 * Node 24, Deno, Bun, browsers, Workers. This used to prefer
 * `globalThis.Buffer` when it was there, reached through an `any` cast,
 * which was the one place in `src/` outside the CLI that touched a Node
 * API. `tsconfig.json` sets `"types": []` so that reaching for one is a
 * compile error; the cast is what got past it. See CLAUDE.md.
 *
 * The two do not agree on invalid input, and that is worth having. Given
 * `"YWJj###"`, `Buffer.from(…, "base64")` silently returns the three
 * bytes it could read; `atob` throws. For an image that means a workbook
 * carrying a truncated or garbage PNG, which Excel refuses to open with
 * no hint as to why — a failure a long way from the bad string that
 * caused it. Failing here instead, and saying so, is the better trade.
 */
function base64ToUint8Array(base64: string): Uint8Array {
  // Strip data URI prefix if present (e.g. "data:image/png;base64,...")
  const clean = base64.includes(",") ? base64.slice(base64.indexOf(",") + 1) : base64

  let binary: string
  try {
    binary = atob(clean)
  } catch {
    throw new InvalidArgumentError(
      "Image data is not valid base64. It decoded to nothing a runtime would accept — " +
        "check for a truncated string, a URL-safe alphabet (`-` and `_` rather than `+` and `/`), " +
        "or bytes that were never encoded at all.",
    )
  }

  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i)
  }
  return bytes
}

/** Create a SheetImage from a base64 string */
export function imageFromBase64(
  base64: string,
  type: SheetImage["type"],
  anchor: SheetImage["anchor"],
): SheetImage {
  const data = base64ToUint8Array(base64)
  return { data, type, anchor }
}
