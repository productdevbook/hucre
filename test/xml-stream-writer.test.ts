import { describe, expect, it } from "vitest"
import { writeXml, writeXmlStream } from "../src/xml/data-writer"
import { readXml } from "../src/xml/data-reader"
import { ParseError } from "../src/errors"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #467 — XML was the one format with no streaming on either side, which
// made it the odd one out of five.
//
// The writer streams cleanly: a prologue, one rendered element per row,
// a closer. The *reader* does not, and this does not pretend otherwise —
// `src/xml/parser.ts` is push-based, so a streaming reader needs a
// pull-based row scanner rather than a wrapper around what is there.
// That is its own change, and the matrix still shows the gap.
// ═══════════════════════════════════════════════════════════════════════

async function drain(stream: ReadableStream<Uint8Array>): Promise<string> {
  const chunks: Uint8Array[] = []
  const reader = stream.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    chunks.push(value)
  }
  const total = chunks.reduce((n, c) => n + c.length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const c of chunks) {
    out.set(c, at)
    at += c.length
  }
  return new TextDecoder().decode(out)
}

const ROWS: Array<Record<string, CellValue>> = [
  { name: "Widget", qty: 3, ok: true },
  { name: "Gadget", qty: 7, ok: false },
]

describe("writeXmlStream produces the same document writeXml does", () => {
  it("byte for byte, on the same rows", async () => {
    // The assertion that makes this a streaming *version* rather than a
    // second implementation with its own opinions.
    expect(await drain(writeXmlStream(ROWS))).toBe(writeXml(ROWS))
  })

  it("with every option that changes the shape", async () => {
    for (const options of [
      { rootTag: "records", rowTag: "record" },
      { declaration: false },
      { pretty: true },
      { pretty: true, indent: "    " },
      { attrPrefix: "$", textKey: "_" },
    ]) {
      expect(await drain(writeXmlStream(ROWS, options)), JSON.stringify(options)).toBe(
        writeXml(ROWS, options),
      )
    }
  })

  it("on no rows at all", async () => {
    expect(await drain(writeXmlStream([]))).toBe(writeXml([]))
  })

  it("with attributes and mixed content", async () => {
    const rows = [{ "@id": 1, "#text": "hello", nested: "x" }]

    expect(await drain(writeXmlStream(rows))).toBe(writeXml(rows))
  })

  it("escaping whatever needs it", async () => {
    const rows = [{ text: '<a & "b">', "@attr": "q'uote" }]

    expect(await drain(writeXmlStream(rows))).toBe(writeXml(rows))
  })
})

describe("it reads back", () => {
  it("through readXml, with the values intact", async () => {
    const { data, headers } = readXml(await drain(writeXmlStream(ROWS)))

    // XML carries no types, so values come back as text — the same as
    // `readXml` of a `writeXml` document, which is the comparison that
    // matters here.
    expect(headers).toEqual(["name", "qty", "ok"])
    expect(data[0]).toEqual({ name: "Widget", qty: "3", ok: "true" })
    expect(data[1]!.name).toBe("Gadget")
    expect(data).toEqual(readXml(writeXml(ROWS)).data)
  })
})

describe("the streaming properties", () => {
  it("pulls rows lazily rather than draining the source first", async () => {
    // Rows big enough that the 64 KB chunk boundary lands well before
    // the source is exhausted; otherwise the first chunk is the only
    // chunk and this would prove nothing.
    let produced = 0
    const filler = "x".repeat(200)
    function* rows(): Generator<Record<string, CellValue>> {
      for (let i = 0; i < 5000; i++) {
        produced++
        yield { i, filler }
      }
    }

    const reader = writeXmlStream(rows()).getReader()
    await reader.read()
    const afterFirstChunk = produced
    await reader.cancel()

    expect(afterFirstChunk).toBeLessThan(5000)
  })

  it("takes an async source", async () => {
    async function* rows(): AsyncGenerator<Record<string, CellValue>> {
      yield { a: 1 }
      yield { a: 2 }
    }

    expect(readXml(await drain(writeXmlStream(rows()))).data).toHaveLength(2)
  })

  it("rejects a bad tag before any bytes go out", async () => {
    // Half a response is worse than no response: validate up front.
    await expect(drain(writeXmlStream(ROWS, { rowTag: "not a tag" }))).rejects.toThrow(ParseError)
    await expect(drain(writeXmlStream(ROWS, { rootTag: "1bad" }))).rejects.toThrow(ParseError)
  })
})
