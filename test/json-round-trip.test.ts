import { valuesOf } from "./_stream"
import { describe, expect, it } from "vitest"
import {
  jsonToWorkbook,
  parseJson,
  parseNdjson,
  unflattenRow,
  unflattenRows,
  workbookToJson,
  writeJson,
  writeNdjson,
} from "../src/json"
import { NdjsonStreamWriter, streamNdjsonRows } from "../src/json/stream"
import { parseCsv } from "../src/csv/reader"
import { streamCsvRows } from "../src/csv/stream"
import type { Workbook } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #409 — the JSON layer read and wrote different documents. Three losses,
// each of which survived because a fixture with one sheet, no nesting and
// no dates never shows any of them.
// ═══════════════════════════════════════════════════════════════════════

const enc = new TextEncoder()

function streamOf(...chunks: string[]): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(enc.encode(chunk))
      controller.close()
    },
  })
}

async function drain<T>(iter: AsyncGenerator<T, void, undefined>): Promise<T[]> {
  const out: T[] = []
  for await (const row of iter) out.push(row)
  return out
}

// ── 1. flatten now has an inverse ───────────────────────────────────

describe("unflattenRow", () => {
  it("rebuilds the nesting that flatten collapsed", () => {
    expect({ ...unflattenRow({ sku: "P1", "pricing.cost": 100, "pricing.retail": 180 }) }).toEqual({
      sku: "P1",
      pricing: { cost: 100, retail: 180 },
    })
  })

  it("rebuilds several levels", () => {
    expect(JSON.parse(JSON.stringify(unflattenRow({ "a.b.c.d": 1 })))).toEqual({
      a: { b: { c: { d: 1 } } },
    })
  })

  it("leaves a key without dots exactly where it is", () => {
    expect({ ...unflattenRow({ a: 1, b: null }) }).toEqual({ a: 1, b: null })
  })

  it("treats every dot as a separator, because flatten does not escape them", () => {
    // `{"a.b": 1}` and `{a: {b: 1}}` flatten to the same cell — flatten even
    // merges them, last one winning — so a literal-dot key cannot be told
    // apart here and comes back nested.
    expect(JSON.parse(JSON.stringify(unflattenRow({ "a.b": 1 })))).toEqual({ a: { b: 1 } })
  })

  it("keeps numeric segments as object keys rather than guessing an array", () => {
    // flatten never emits an index: primitive arrays are joined into one
    // cell and arrays of objects are JSON-encoded. An array here would be a
    // shape the flat form never meant.
    expect(JSON.parse(JSON.stringify(unflattenRow({ "a.0": 1, "a.1": 2 })))).toEqual({
      a: { "0": 1, "1": 2 },
    })
  })

  it("keeps a conflicting key flat instead of dropping either value", () => {
    expect(JSON.parse(JSON.stringify(unflattenRow({ a: 1, "a.b": 2 })))).toEqual({
      a: 1,
      "a.b": 2,
    })
    expect(JSON.parse(JSON.stringify(unflattenRow({ "a.b.c": 1, "a.b": 2 })))).toEqual({
      a: { b: { c: 1 } },
      "a.b": 2,
    })
  })

  it("preserves Date cells rather than descending into them", () => {
    const d = new Date("2024-01-15T10:30:00.000Z")
    expect((unflattenRow({ "meta.at": d }).meta as Record<string, unknown>).at).toBe(d)
  })

  it("rebuilds a __proto__ path as data without polluting Object.prototype", () => {
    // The library already shipped one prototype-pollution bug (fillTemplate),
    // and flatten deliberately keeps "__proto__" as an ordinary column, so a
    // "__proto__.x" key is a shape that really reaches this code.
    const out = unflattenRow({ "__proto__.polluted": true })
    expect(({} as Record<string, unknown>).polluted).toBeUndefined()
    // Serialized rather than compared to a literal: `{__proto__: …}` in an
    // object literal sets the prototype instead of a key, so the expectation
    // would quietly be `{}`.
    expect(JSON.stringify(out)).toBe('{"__proto__":{"polluted":true}}')
  })

  it("rebuilds a constructor.prototype path as data too", () => {
    const out = unflattenRow({ "constructor.prototype.x": 1 })
    expect(({} as Record<string, unknown>).x).toBeUndefined()
    expect(JSON.parse(JSON.stringify(out))).toEqual({ constructor: { prototype: { x: 1 } } })
  })

  it("maps every row via unflattenRows", () => {
    expect(JSON.parse(JSON.stringify(unflattenRows([{ "a.b": 1 }, { "a.b": 2 }])))).toEqual([
      { a: { b: 1 } },
      { a: { b: 2 } },
    ])
  })
})

describe("writeJson unflatten", () => {
  const nested = '[{"sku":"P1","pricing":{"cost":100}}]'

  it("is off by default, so an unrelated dotted header stays a column name", () => {
    // "Q1.2024" is a spreadsheet header, not a path. Nesting it by default
    // would be a new silent mangling introduced by the fix for one.
    expect(writeJson([{ "Q1.2024": 5 }])).toBe('[{"Q1.2024":5}]')
  })

  it("restores the nesting parseJson flattened", () => {
    const { data } = parseJson(nested)
    expect(data[0]).toEqual({ sku: "P1", "pricing.cost": 100 })
    expect(JSON.parse(writeJson(data, { unflatten: true }))).toEqual([
      { sku: "P1", pricing: { cost: 100 } },
    ])
  })

  it("round-trips a nested document byte for byte", () => {
    expect(writeJson(parseJson(nested).data, { unflatten: true })).toBe(nested)
  })

  it("still pretty-prints", () => {
    expect(writeJson([{ "a.b": 1 }], { unflatten: true, pretty: true })).toBe(
      '[\n  {\n    "a": {\n      "b": 1\n    }\n  }\n]',
    )
  })

  it("does not resurrect a joined primitive array, which is not recoverable", () => {
    // `"1, 2"` and the literal string "1, 2" are the same cell by then, so
    // splitting it back would mangle ordinary text containing a comma.
    const { data } = parseJson('[{"arr":[1,2]}]')
    expect(JSON.parse(writeJson(data, { unflatten: true }))).toEqual([{ arr: "1, 2" }])
  })
})

describe("writeNdjson unflatten", () => {
  it("restores nesting per line", () => {
    const { data } = parseNdjson('{"a":{"b":1}}\n{"a":{"b":2}}\n')
    expect(writeNdjson(data, { unflatten: true })).toBe('{"a":{"b":1}}\n{"a":{"b":2}}\n')
  })

  it("is off by default", () => {
    expect(writeNdjson([{ "a.b": 1 }])).toBe('{"a.b":1}\n')
  })
})

describe("NdjsonStreamWriter unflatten", () => {
  it("restores nesting on the streaming path too", () => {
    const w = new NdjsonStreamWriter({ unflatten: true })
    w.addObject({ "user.name": "Ada" })
    expect(w.finish()).toBe('{"user":{"name":"Ada"}}\n')
  })

  it("is off by default", () => {
    const w = new NdjsonStreamWriter()
    w.addObject({ "user.name": "Ada" })
    expect(w.finish()).toBe('{"user.name":"Ada"}\n')
  })
})

// ── 2. dates are revived, on the same rule CSV uses ─────────────────

describe("typeInference revives dates", () => {
  const iso = "2024-01-15T10:30:00.000Z"

  it("hands ISO strings back as strings by default", () => {
    expect(parseJson(`[{"d":"${iso}"}]`).data[0]!.d).toBe(iso)
  })

  it("revives them as Date when asked", () => {
    const value = parseJson(`[{"d":"${iso}"}]`, { typeInference: true }).data[0]!.d
    expect(value).toBeInstanceOf(Date)
    expect((value as Date).toISOString()).toBe(iso)
  })

  it("closes the writeJson → parseJson loop for Date cells", () => {
    const at = new Date(iso)
    const back = parseJson(writeJson([{ at }]), { typeInference: true })
    expect(back.data[0]!.at).toEqual(at)
  })

  it("revives a date nested inside a flattened path", () => {
    const back = parseJson(`[{"meta":{"at":"${iso}"}}]`, { typeInference: true })
    expect(back.data[0]!["meta.at"]).toBeInstanceOf(Date)
  })

  it("accepts exactly what CSV accepts, and rejects exactly what CSV rejects", () => {
    // Same option name, same rule — the readers used to give two answers to
    // the same question.
    const cases = ["2024-01-15", "2024-01-15T10:30:00Z", "2024-13-45", "3/4/2021", "2024", "hello"]
    for (const raw of cases) {
      const viaCsv = parseCsv(`v\n"${raw}"`, { typeInference: true, header: true })[1]![0]
      const viaJson = parseJson(`[{"v":${JSON.stringify(raw)}}]`, { typeInference: true }).data[0]!
        .v
      expect([raw, viaJson instanceof Date]).toEqual([raw, viaCsv instanceof Date])
    }
  })

  it("does not coerce numbers or booleans out of JSON strings", () => {
    // CSV has to guess those out of text; JSON already carries them, so a
    // string is a string by the author's choice.
    const back = parseJson('[{"n":"007","b":"true"}]', { typeInference: true })
    expect(back.data[0]).toEqual({ n: "007", b: "true" })
  })

  it("works for parseNdjson", () => {
    const back = parseNdjson(`{"d":"${iso}"}\n`, { typeInference: true })
    expect(back.data[0]!.d).toBeInstanceOf(Date)
  })

  it("works for streamNdjsonRows, with and without flattenRows", async () => {
    const flat = await drain(
      streamNdjsonRows(streamOf(`{"a":{"d":"${iso}"}}\n`), {
        typeInference: true,
        flattenRows: true,
      }),
    )
    expect(flat[0]!.values["a.d"]).toBeInstanceOf(Date)

    const raw = await drain(
      streamNdjsonRows(streamOf(`{"a":{"d":"${iso}"}}\n`), { typeInference: true }),
    )
    expect((raw[0]!.values.a as unknown as { d: unknown }).d).toBeInstanceOf(Date)
  })

  it("leaves streamNdjsonRows rows untouched without the option", async () => {
    const rows = await drain(streamNdjsonRows(streamOf(`{"d":"${iso}"}\n`)))
    expect(rows[0]!.values.d).toBe(iso)
  })
})

describe("the shared ISO rule did not change CSV", () => {
  it("still infers dates in parseCsv and streamCsvRows alike", async () => {
    const rows = parseCsv("when\n2024-01-15", { typeInference: true, header: true })
    expect(rows[1]![0]).toBeInstanceOf(Date)
    const streamed = await valuesOf(
      streamCsvRows("when\n2024-01-15", { typeInference: true, header: true }),
    )
    expect(streamed[1]![0]).toBeInstanceOf(Date)
  })
})

// ── 3. workbook round trip at any sheet count ───────────────────────

const wb = (...names: string[]): Workbook => ({
  sheets: names.map((name, i) => ({ name, rows: [[`c${i}`], [i]] })),
})

describe("workbookToJson shape", () => {
  it("still emits a bare array for one sheet and a keyed object beyond that", () => {
    expect(workbookToJson(wb("S1"))).toBe('[{"c0":0}]')
    expect(workbookToJson(wb("S1", "S2"))).toBe('{"S1":[{"c0":0}],"S2":[{"c1":1}]}')
  })

  it("emits the keyed object at every sheet count under shape: sheets", () => {
    // The stable contract: a consumer written against a one-sheet export no
    // longer breaks the day a second sheet appears.
    expect(workbookToJson(wb("S1"), { shape: "sheets" })).toBe('{"S1":[{"c0":0}]}')
    expect(workbookToJson(wb(), { shape: "sheets" })).toBe("{}")
  })

  it("keeps a sheet legally named __proto__ in the output", () => {
    // A plain-object accumulator drops it: the key hits the prototype setter.
    const json = workbookToJson(wb("__proto__", "S2"))
    expect(json).toBe('{"__proto__":[{"c0":0}],"S2":[{"c1":1}]}')
    expect(jsonToWorkbook(json).sheets.map((s) => s.name)).toEqual(["__proto__", "S2"])
  })

  it("applies unflatten to every sheet", () => {
    const workbook: Workbook = {
      sheets: [
        { name: "A", rows: [["u.name"], ["Ada"]] },
        { name: "B", rows: [["u.name"], ["Grace"]] },
      ],
    }
    expect(JSON.parse(workbookToJson(workbook, { unflatten: true }))).toEqual({
      A: [{ u: { name: "Ada" } }],
      B: [{ u: { name: "Grace" } }],
    })
  })
})

describe("jsonToWorkbook", () => {
  for (const names of [[], ["S1"], ["S1", "S2"], ["S1", "S2", "S3"]]) {
    it(`round-trips a ${names.length}-sheet workbook through the auto shape`, () => {
      const source = wb(...names)
      const back = jsonToWorkbook(workbookToJson(source))
      // The auto shape carries no name for a lone sheet, which is the price
      // of a bare array; every cell survives either way.
      expect(back.sheets.map((s) => s.rows)).toEqual(source.sheets.map((s) => s.rows))
      if (names.length !== 1) expect(back.sheets.map((s) => s.name)).toEqual(names)
    })

    it(`round-trips a ${names.length}-sheet workbook through shape: sheets, names included`, () => {
      const source = wb(...names)
      const back = jsonToWorkbook(workbookToJson(source, { shape: "sheets" }))
      expect(back).toEqual(source)
    })
  }

  it("names a bare array's sheet Sheet1, or whatever sheetName says", () => {
    expect(jsonToWorkbook('[{"a":1}]').sheets[0]!.name).toBe("Sheet1")
    expect(jsonToWorkbook('[{"a":1}]', { sheetName: "Data" }).sheets[0]!.name).toBe("Data")
  })

  it("reads a plain object as a one-row sheet", () => {
    expect(jsonToWorkbook('{"a":1,"b":2}').sheets[0]!.rows).toEqual([
      ["a", "b"],
      [1, 2],
    ])
  })

  it("returns an empty workbook for empty input", () => {
    expect(jsonToWorkbook("")).toEqual({ sheets: [] })
    expect(jsonToWorkbook("{}")).toEqual({ sheets: [] })
  })

  it("emits no rows at all for an empty sheet rather than a phantom header", () => {
    expect(jsonToWorkbook('{"S1":[]}').sheets[0]!.rows).toEqual([])
  })

  it("accepts bytes and already-parsed values", () => {
    expect(jsonToWorkbook(enc.encode('[{"a":1}]')).sheets[0]!.rows).toEqual([["a"], [1]])
    expect(jsonToWorkbook([{ a: 1 }]).sheets[0]!.rows).toEqual([["a"], [1]])
  })

  it("honours the read options, so dates survive the workbook round trip", () => {
    const at = new Date("2024-01-15T10:30:00.000Z")
    const source: Workbook = { sheets: [{ name: "S", rows: [["at"], [at]] }] }
    const back = jsonToWorkbook(workbookToJson(source, { shape: "sheets" }), {
      typeInference: true,
    })
    expect(back.sheets[0]!.rows[1]![0]).toEqual(at)
  })

  it("throws ParseError on malformed JSON and on a top-level primitive", () => {
    expect(() => jsonToWorkbook("{nope")).toThrow(/Invalid JSON/)
    expect(() => jsonToWorkbook(42)).toThrow(/object or an array/)
  })
})

describe("parseJson on a multi-sheet document", () => {
  const doc = '{"S1":[{"a":1}],"S2":[{"b":2}]}'

  it("names the problem instead of returning one row of stringified sheets", () => {
    expect(() => parseJson(doc)).toThrow(/multi-sheet workbook/)
    expect(() => parseJson(doc)).toThrow(/jsonToWorkbook/)
  })

  it("still reads an object whose array columns are not rows as one row", () => {
    // `{a: [1,2], b: [3,4]}` is a row with two list-valued columns, not a
    // workbook, and the guard has to tell them apart.
    expect(parseJson('{"a":[1,2],"b":[3,4]}').data).toEqual([{ a: "1, 2", b: "3, 4" }])
  })

  it("leaves the single-array-property shape alone", () => {
    expect(parseJson('{"products":[{"sku":"P1"}]}').data).toEqual([{ sku: "P1" }])
  })

  it("offers rowsAt as the documented escape hatch, both ways", () => {
    expect(parseJson(doc, { rowsAt: "S1" }).data).toEqual([{ a: 1 }])
    expect(parseJson(doc, { rowsAt: "" }).data).toEqual([{ S1: '[{"a":1}]', S2: '[{"b":2}]' }])
  })
})
