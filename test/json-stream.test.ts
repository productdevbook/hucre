import { describe, expect, it } from "vitest"
import { NdjsonStreamWriter, readNdjsonStream } from "../src/json"

function streamFromString(s: string): ReadableStream<Uint8Array> {
  const enc = new TextEncoder()
  return new ReadableStream({
    start(controller) {
      controller.enqueue(enc.encode(s))
      controller.close()
    },
  })
}

function chunkedStream(parts: string[]): ReadableStream<Uint8Array> {
  const enc = new TextEncoder()
  let i = 0
  return new ReadableStream({
    pull(controller) {
      if (i < parts.length) {
        controller.enqueue(enc.encode(parts[i]!))
        i++
      } else {
        controller.close()
      }
    },
  })
}

describe("NdjsonStreamWriter", () => {
  it("buffers writes into NDJSON output", () => {
    const w = new NdjsonStreamWriter()
    w.write({ a: 1 })
    w.write({ a: 2 })
    expect(w.toString()).toBe('{"a":1}\n{"a":2}\n')
  })

  it("throws when writing after end()", () => {
    const w = new NdjsonStreamWriter()
    w.end()
    expect(() => w.write({ a: 1 })).toThrow()
  })

  it("emits a ReadableStream that can be drained", async () => {
    const w = new NdjsonStreamWriter()
    w.write({ a: 1 })
    w.write({ a: 2 })
    w.end()

    const reader = w.toStream().getReader()
    const dec = new TextDecoder()
    let out = ""
    while (true) {
      const { value, done } = await reader.read()
      if (done) break
      out += dec.decode(value)
    }
    expect(out).toBe('{"a":1}\n{"a":2}\n')
  })

  it("converts Date values to ISO strings", () => {
    const w = new NdjsonStreamWriter()
    const d = new Date("2025-04-25T00:00:00Z")
    w.write({ at: d })
    expect(w.toString().trim()).toBe(`{"at":"${d.toISOString()}"}`)
  })
})

describe("readNdjsonStream", () => {
  it("yields parsed rows from a stream", async () => {
    const stream = streamFromString('{"a":1}\n{"a":2}\n{"a":3}\n')
    const rows: unknown[] = []
    for await (const row of readNdjsonStream(stream)) {
      rows.push(row)
    }
    expect(rows).toEqual([{ a: 1 }, { a: 2 }, { a: 3 }])
  })

  it("handles split-across-chunk lines", async () => {
    const stream = chunkedStream(['{"a":', '1}\n{"a"', ":2}\n"])
    const rows: unknown[] = []
    for await (const row of readNdjsonStream(stream)) rows.push(row)
    expect(rows).toEqual([{ a: 1 }, { a: 2 }])
  })

  it("handles trailing line without newline", async () => {
    const stream = streamFromString('{"a":1}\n{"a":2}')
    const rows: unknown[] = []
    for await (const row of readNdjsonStream(stream)) rows.push(row)
    expect(rows).toEqual([{ a: 1 }, { a: 2 }])
  })

  it("flattens rows when flattenRows: true", async () => {
    const stream = streamFromString('{"a":{"b":1}}\n{"a":{"b":2}}\n')
    const rows: Record<string, unknown>[] = []
    for await (const row of readNdjsonStream(stream, { flattenRows: true })) {
      rows.push(row)
    }
    expect(rows).toEqual([{ "a.b": 1 }, { "a.b": 2 }])
  })

  it("throws on invalid line by default", async () => {
    const stream = streamFromString('{"a":1}\n{bad}\n')
    await expect(async () => {
      for await (const _ of readNdjsonStream(stream)) {
        // consume
      }
    }).rejects.toThrow(/line 2/)
  })

  it("calls onError and skips when provided", async () => {
    const stream = streamFromString('{"a":1}\n{bad}\n{"a":2}\n')
    const errs: number[] = []
    const rows: unknown[] = []
    for await (const row of readNdjsonStream(stream, {
      onError: (_l, ln) => errs.push(ln),
    })) {
      rows.push(row)
    }
    expect(rows).toEqual([{ a: 1 }, { a: 2 }])
    expect(errs).toEqual([2])
  })
})
