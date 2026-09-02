import { HucreError } from "../errors"
import type { CsvReadOptions, CellValue } from "../_types"
import { parseCsv } from "./reader"

/**
 * Fetch a CSV from a URL and parse it (requires the fetch API).
 *
 * The response's own `Content-Type; charset=` wins when it names one,
 * because the server is asserting something. Otherwise the bytes are
 * handed to `parseCsv`, which reads the byte-order mark — this used to go
 * through `response.text()`, which assumes UTF-8 in that case and turns a
 * UTF-16 export into NUL-separated letters. `options.encoding` overrides
 * both. See #475.
 */
export async function fetchCsv(url: string, options?: CsvReadOptions): Promise<CellValue[][]> {
  const response = await fetch(url)
  if (!response.ok) throw new HucreError(`Failed to fetch: ${response.status}`)

  const declared = charsetOf(response.headers.get("content-type"))
  const bytes = new Uint8Array(await response.arrayBuffer())

  return parseCsv(bytes, { ...options, encoding: options?.encoding ?? declared })
}

/** The `charset=` of a Content-Type header, if it carries one. */
function charsetOf(contentType: string | null): string | undefined {
  if (!contentType) return undefined
  const match = /;\s*charset\s*=\s*"?([^";]+)"?/i.exec(contentType)
  return match ? match[1]!.trim() : undefined
}
