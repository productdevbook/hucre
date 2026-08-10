import { describe, expect, it, vi } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { writeXlsxStream } from "../src/xlsx/stream-writer"

// ═══════════════════════════════════════════════════════════════════════
// The CI matrix added in this PR found this on its first run.
//
// `deflate-raw` is the format a ZIP entry needs, and "does this runtime
// have CompressionStream" is a different question. Node 18 has the
// constructor and rejects `deflate-raw` — it shipped with `gzip` and
// `deflate` only, and the raw format arrived in Node 20.
//
// Four modules each memoized their own flag and every one probed the
// constructor's existence. The buffered writer survived because it also
// wrapped the construction in a try/catch; the streaming writer threw
// ERR_INVALID_ARG_VALUE on the first chunk, so `writeXlsxStream` did not
// work on the Node version the package claims as its floor.
// ═══════════════════════════════════════════════════════════════════════

async function drain(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  let total = 0
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
    total += value.length
  }
  const out = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    out.set(chunk, offset)
    offset += chunk.length
  }
  return out
}

const ROWS = [
  ["Name", "Amount"],
  ["Ada", 1234.5],
]

/**
 * Stand in for Node 18: the constructor exists and refuses `deflate-raw`.
 * The modules cache their answer, so this runs in an isolated module
 * registry — otherwise the first test in the file would fix the answer
 * for the rest.
 */
function withoutDeflateRaw<T>(body: () => Promise<T>): Promise<T> {
  const real = globalThis.CompressionStream
  class Node18CompressionStream {
    constructor(format: string) {
      if (format === "deflate-raw") {
        throw new TypeError("The argument 'format' is invalid. Received 'deflate-raw'")
      }
      return new real(format as CompressionFormat)
    }
  }
  vi.stubGlobal("CompressionStream", Node18CompressionStream)
  return body().finally(() => vi.unstubAllGlobals())
}

describe("a runtime whose CompressionStream lacks deflate-raw", () => {
  it("still writes a readable workbook through the streaming writer", async () => {
    const bytes = await withoutDeflateRaw(async () => {
      vi.resetModules()
      const { writeXlsxStream: fresh } = await import("../src/xlsx/stream-writer")
      return drain(fresh(ROWS, { name: "S" }))
    })

    // Before the fix this threw ERR_INVALID_ARG_VALUE on the first chunk.
    const wb = await readXlsx(bytes)
    expect(wb.sheets[0]!.rows).toEqual(ROWS)
  })

  it("still writes a readable workbook through the buffered writer", async () => {
    const bytes = await withoutDeflateRaw(async () => {
      vi.resetModules()
      const { writeXlsx: fresh } = await import("../src/xlsx/writer")
      return fresh({ sheets: [{ name: "S", rows: ROWS }] })
    })

    const wb = await readXlsx(bytes)
    expect(wb.sheets[0]!.rows).toEqual(ROWS)
  })
})

describe("a runtime that does support deflate-raw", () => {
  it("compresses, and the two writers agree on the content", async () => {
    const streamed = await drain(writeXlsxStream(ROWS, { name: "S" }))
    const buffered = await writeXlsx({ sheets: [{ name: "S", rows: ROWS }] })

    expect((await readXlsx(streamed)).sheets[0]!.rows).toEqual(ROWS)
    expect((await readXlsx(buffered)).sheets[0]!.rows).toEqual(ROWS)
  })
})
