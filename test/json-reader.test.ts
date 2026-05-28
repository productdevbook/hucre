import { describe, expect, it } from "vitest"
import { parseJson, parseNdjson, parseValue } from "../src/json"

describe("parseJson", () => {
  it("reads a top-level array of objects", () => {
    const r = parseJson('[{"a":1,"b":2},{"a":3,"b":4}]')
    expect(r.headers).toEqual(["a", "b"])
    expect(r.data).toEqual([
      { a: 1, b: 2 },
      { a: 3, b: 4 },
    ])
  })

  it("reads { rows: [...] } single-array-property shape", () => {
    const r = parseJson('{"products":[{"sku":"P1"},{"sku":"P2"}]}')
    expect(r.headers).toEqual(["sku"])
    expect(r.data).toEqual([{ sku: "P1" }, { sku: "P2" }])
  })

  it("treats a top-level object as a single row", () => {
    const r = parseJson('{"a":1,"b":2}')
    expect(r.data).toEqual([{ a: 1, b: 2 }])
  })

  it("reads from an explicit rowsAt path", () => {
    const r = parseJson('{"meta":{},"data":{"rows":[{"a":1},{"a":2}]}}', {
      rowsAt: "data.rows",
    })
    expect(r.data).toEqual([{ a: 1 }, { a: 2 }])
  })

  it("flattens nested objects with dot-path keys", () => {
    const r = parseJson('[{"sku":"P1","pricing":{"cost":100,"retail":180},"tags":["wood","oak"]}]')
    expect(r.headers).toEqual(["sku", "pricing.cost", "pricing.retail", "tags"])
    expect(r.data[0]).toEqual({
      sku: "P1",
      "pricing.cost": 100,
      "pricing.retail": 180,
      tags: "wood, oak",
    })
  })

  it("respects flatten: false by stringifying nested objects", () => {
    const r = parseJson('[{"sku":"P1","pricing":{"cost":100}}]', { flatten: false })
    expect(r.data[0]).toEqual({
      sku: "P1",
      pricing: '{"cost":100}',
    })
  })

  it("joins primitive arrays with arrayJoin separator", () => {
    const r = parseJson('[{"tags":["a","b","c"]}]', { arrayJoin: "|" })
    expect(r.data[0]!.tags).toBe("a|b|c")
  })

  it("union of keys across rows becomes headers", () => {
    const r = parseJson('[{"a":1},{"b":2}]')
    expect(r.headers).toEqual(["a", "b"])
    expect(r.data).toEqual([
      { a: 1, b: null },
      { a: null, b: 2 },
    ])
  })

  it("accepts Uint8Array input", () => {
    const buf = new TextEncoder().encode('[{"x":1}]')
    expect(parseJson(buf).data).toEqual([{ x: 1 }])
  })

  it("returns empty result for empty input", () => {
    expect(parseJson("")).toEqual({ data: [], headers: [] })
    expect(parseJson("   ")).toEqual({ data: [], headers: [] })
  })

  it("throws ParseError on invalid JSON", () => {
    expect(() => parseJson("{not json")).toThrow(/Invalid JSON/)
  })

  it("throws ParseError when top-level is null", () => {
    expect(() => parseJson("null")).toThrow(/object or an array/)
  })

  it("throws when rowsAt path is missing", () => {
    expect(() => parseJson('{"a":1}', { rowsAt: "missing.path" })).toThrow(/No data found/)
  })

  it("respects maxRows", () => {
    const r = parseJson('[{"a":1},{"a":2},{"a":3},{"a":4}]', { maxRows: 2 })
    expect(r.data).toHaveLength(2)
  })

  it("applies transformHeader", () => {
    const r = parseJson('[{"firstName":"A","lastName":"B"}]', {
      transformHeader: (h) => h.toLowerCase(),
    })
    expect(r.headers).toEqual(["firstname", "lastname"])
    expect(r.data[0]).toEqual({ firstname: "A", lastname: "B" })
  })

  it("applies transformValue", () => {
    const r = parseJson('[{"price":10},{"price":20}]', {
      transformValue: (v, h) => (h === "price" && typeof v === "number" ? v * 2 : v),
    })
    expect(r.data).toEqual([{ price: 20 }, { price: 40 }])
  })
})

describe("parseValue", () => {
  it("works with already-parsed JSON values", () => {
    const r = parseValue([{ a: 1 }, { a: 2 }])
    expect(r.data).toEqual([{ a: 1 }, { a: 2 }])
  })
})

describe("parseNdjson", () => {
  it("parses one object per line", () => {
    const r = parseNdjson('{"a":1}\n{"a":2}\n{"a":3}\n')
    expect(r.headers).toEqual(["a"])
    expect(r.data).toEqual([{ a: 1 }, { a: 2 }, { a: 3 }])
  })

  it("skips blank lines", () => {
    const r = parseNdjson('{"a":1}\n\n{"a":2}\n')
    expect(r.data).toEqual([{ a: 1 }, { a: 2 }])
  })

  it("handles CRLF line endings", () => {
    const r = parseNdjson('{"a":1}\r\n{"a":2}\r\n')
    expect(r.data).toEqual([{ a: 1 }, { a: 2 }])
  })

  it("throws ParseError with line number on invalid line", () => {
    expect(() => parseNdjson('{"a":1}\n{bad}\n')).toThrow(/line 2/)
  })

  it("calls onError and skips invalid lines when provided", () => {
    const errs: number[] = []
    const r = parseNdjson('{"a":1}\n{bad}\n{"a":2}\n', {
      onError: (_l, ln) => errs.push(ln),
    })
    expect(r.data).toEqual([{ a: 1 }, { a: 2 }])
    expect(errs).toEqual([2])
  })

  it("flattens nested objects per row", () => {
    const r = parseNdjson('{"a":{"b":1}}\n{"a":{"b":2}}\n')
    expect(r.headers).toEqual(["a.b"])
    expect(r.data).toEqual([{ "a.b": 1 }, { "a.b": 2 }])
  })
})
