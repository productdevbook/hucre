import { describe, expect, it } from "vitest"
import { collectHeaders, flattenValue } from "../src/json/flatten"
import { parseJson, parseNdjson, parseValue } from "../src/json/reader"
import { workbookToJson, writeJson, writeNdjson } from "../src/json/writer"
import { streamNdjsonRows } from "../src/json/stream"
import { toJson } from "../src/export/json"
import { toHtml } from "../src/export/html"
import type { Cell, CellValue, Sheet, Workbook } from "../src/_types"

// ── Helpers ──────────────────────────────────────────────────────────

const enc = new TextEncoder()

/** A one-chunk-per-string ReadableStream, so chunk boundaries are explicit. */
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

function sheet(rows: CellValue[][], extra?: Partial<Sheet>): Sheet {
  return { name: "Sheet1", rows, ...extra }
}

/** A sheet whose cells carry styles, for the HTML exporter. */
function styledSheet(rows: CellValue[][], styles: Record<string, Partial<Cell>>): Sheet {
  const cells = new Map<string, Cell>()
  for (const [key, cell] of Object.entries(styles)) {
    cells.set(key, cell as Cell)
  }
  return { name: "Sheet1", rows, cells }
}

// ═══════════════════════════════════════════════════════════════════════
// json/flatten — the documented flattening rules
// ═══════════════════════════════════════════════════════════════════════

describe("flattenValue — scalars and absent values", () => {
  it("drops a top-level null instead of inventing a column for it", async () => {
    // There is no key to hang the value on, so the row is empty rather than
    // carrying a "" column.
    expect({ ...flattenValue(null) }).toEqual({})
    expect({ ...flattenValue(undefined) }).toEqual({})
  })

  it("keeps a null property as a null cell", () => {
    expect({ ...flattenValue({ name: "Ada", nickname: null }) }).toEqual({
      name: "Ada",
      nickname: null,
    })
  })

  it("normalizes undefined properties to null", () => {
    // JSON.parse never produces undefined, but `parseValue` accepts objects
    // built in JS, where a missing-but-present key is common.
    expect({ ...flattenValue({ a: undefined }) }).toEqual({ a: null })
  })

  it("preserves a Date rather than stringifying it", () => {
    const d = new Date("2024-03-01T00:00:00Z")
    expect(flattenValue({ when: d }).when).toBe(d)
  })

  it("stringifies values JSON cannot carry", () => {
    // bigint and functions reach the writer from hand-built JS objects.
    const out = flattenValue({ big: 10n as unknown as CellValue })
    expect(out.big).toBe("10")
  })
})

describe("flattenValue — arrays", () => {
  it("renders an empty array as an empty cell", () => {
    expect(flattenValue({ tags: [] }).tags).toBe("")
  })

  it("joins an all-primitive array with arrayJoin", () => {
    const out = flattenValue({ tags: ["a", 1, true, null] })
    expect(out.tags).toBe("a, 1, true, ")
  })

  it("honours a custom arrayJoin", () => {
    expect(flattenValue({ tags: [1, 2, 3] }, { arrayJoin: " | " }).tags).toBe("1 | 2 | 3")
  })

  it("falls back to JSON for an array of objects", () => {
    // A tabular row has no room for a nested table, so the array is kept
    // verbatim rather than being lossily flattened.
    const out = flattenValue({ items: [{ id: 1 }, { id: 2 }] })
    expect(out.items).toBe(`[{"id":1},{"id":2}]`)
  })

  it("falls back to JSON for a mixed array", () => {
    expect(flattenValue({ mixed: [1, { a: 2 }] }).mixed).toBe(`[1,{"a":2}]`)
  })
})

describe("flattenValue — nested objects", () => {
  it("joins nested keys with dots", () => {
    const out = flattenValue({ user: { name: "Ada", address: { city: "London" } } })
    expect({ ...out }).toEqual({ "user.name": "Ada", "user.address.city": "London" })
  })

  it("renders an empty nested object as an empty cell", () => {
    expect(flattenValue({ meta: {} }).meta).toBe("")
  })

  it("produces no columns at all for an empty top-level object", () => {
    expect({ ...flattenValue({}) }).toEqual({})
  })

  it("serializes nested objects as JSON when flatten is off", () => {
    const out = flattenValue({ id: 1, user: { name: "Ada" } }, { flatten: false })
    expect({ ...out }).toEqual({ id: 1, user: `{"name":"Ada"}` })
  })

  it("stops descending at maxDepth and keeps the remainder as JSON", () => {
    // Guards against a deeply nested (or cyclic-looking) document producing
    // thousands of columns.
    const value = { a: { b: { c: { d: 1 } } } }
    expect({ ...flattenValue(value, { maxDepth: 2 }) }).toEqual({ "a.b": `{"c":{"d":1}}` })
  })

  it("does not prefix a dot when the outer key is the empty string", () => {
    // `{"": {...}}` is legal JSON and appears in API payloads keyed by a
    // possibly-blank identifier.
    expect({ ...flattenValue({ "": { a: 1 } }) }).toEqual({ a: 1 })
  })

  it("stores __proto__ as an ordinary column", () => {
    // The output object has a null prototype precisely so this key cannot
    // hit the prototype setter and vanish.
    const out = flattenValue(JSON.parse(`{"__proto__": 1, "id": 2}`))
    expect(Object.keys(out)).toEqual(["__proto__", "id"])
    expect(out["__proto__"]).toBe(1)
  })
})

describe("collectHeaders", () => {
  it("unions keys across rows in first-seen order", () => {
    const headers = collectHeaders([
      { b: 1, a: 2 },
      { c: 3, a: 4 },
    ])
    expect(headers).toEqual(["b", "a", "c"])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// json/reader
// ═══════════════════════════════════════════════════════════════════════

describe("parseJson — row extraction", () => {
  it("treats an empty rowsAt path as 'the whole document'", () => {
    const result = parseJson(`[{"a":1}]`, { rowsAt: "" })
    expect(result.data).toEqual([{ a: 1 }])
  })

  it("wraps a single object found at rowsAt as one row", () => {
    const result = parseJson(`{"result":{"id":7,"name":"x"}}`, { rowsAt: "result" })
    expect(result.data).toEqual([{ id: 7, name: "x" }])
  })

  it("rejects a rowsAt path that resolves to a primitive", () => {
    expect(() => parseJson(`{"count":3}`, { rowsAt: "count" })).toThrow(/not an object or array/)
  })

  it("rejects a rowsAt path that resolves to nothing", () => {
    expect(() => parseJson(`{"a":1}`, { rowsAt: "missing.deep" })).toThrow(/No data found/)
  })

  it("turns a null array element into an empty row", () => {
    // A sparse API response (`[{...}, null, {...}]`) must not shift the rows
    // that follow it.
    const result = parseJson(`[{"a":1},null,{"a":3}]`)
    expect(result.data).toEqual([{ a: 1 }, { a: null }, { a: 3 }])
  })

  it("wraps primitive and array elements in a 'value' column", () => {
    const result = parseJson(`[1,"two",[3,4]]`)
    expect(result.headers).toEqual(["value"])
    expect(result.data).toEqual([{ value: 1 }, { value: "two" }, { value: "3, 4" }])
  })

  it("returns nothing for whitespace-only input", () => {
    expect(parseJson("   \n  ")).toEqual({ data: [], headers: [] })
  })

  it("accepts UTF-8 bytes as well as a string", () => {
    expect(parseJson(enc.encode(`[{"a":1}]`)).data).toEqual([{ a: 1 }])
  })

  it("rejects a top-level null", () => {
    expect(() => parseValue(null)).toThrow(/must be an object or an array of objects/)
  })

  it("rejects a top-level primitive", () => {
    expect(() => parseValue(42)).toThrow(/must be an object or an array of objects/)
  })
})

describe("parseNdjson", () => {
  it("returns nothing for whitespace-only input", () => {
    expect(parseNdjson("\n \n")).toEqual({ data: [], headers: [] })
  })

  it("hands malformed lines to onError and keeps going", () => {
    const seen: Array<[string, number]> = []
    const result = parseNdjson(`{"a":1}\nnot json\n{"a":3}\n`, {
      onError: (line, lineNumber) => seen.push([line, lineNumber]),
    })
    expect(seen).toEqual([["not json", 2]])
    expect(result.data).toEqual([{ a: 1 }, { a: 3 }])
  })

  it("throws with the offending line number when no onError is given", () => {
    expect(() => parseNdjson(`{"a":1}\noops\n`)).toThrow(/line 2/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// json/stream
// ═══════════════════════════════════════════════════════════════════════

describe("streamNdjsonRows", () => {
  it("skips blank lines between records", () => {
    // Producers that flush per-record often emit a stray newline.
    return drain(streamNdjsonRows(streamOf(`{"a":1}\n`, `\n`, `  \n`, `{"a":2}\n`))).then(
      (rows) => {
        expect(rows).toEqual([{ a: 1 }, { a: 2 }])
      },
    )
  })

  it("flattens a trailing record that has no terminating newline", async () => {
    // The last line of a file frequently lacks its "\n"; it must take the
    // same flattening path as every other line.
    const rows = await drain(
      streamNdjsonRows(
        streamOf(`{"id":1,"user":{"name":"Ada"}}\n`, `{"id":2,"user":{"name":"G"}}`),
        {
          flattenRows: true,
        },
      ),
    )
    expect(rows.map((r) => ({ ...r }))).toEqual([
      { id: 1, "user.name": "Ada" },
      { id: 2, "user.name": "G" },
    ])
  })

  it("yields a trailing array record untouched even with flattenRows on", async () => {
    // Flattening only applies to objects; an array line is passed through.
    const rows = await drain(streamNdjsonRows(streamOf(`[1,2]`), { flattenRows: true }))
    expect(rows).toEqual([[1, 2]])
  })

  it("reports a malformed trailing line through onError", async () => {
    const seen: number[] = []
    const rows = await drain(
      streamNdjsonRows(streamOf(`{"a":1}\n`, `{oops`), {
        onError: (_line, lineNumber) => seen.push(lineNumber),
      }),
    )
    expect(seen).toEqual([2])
    expect(rows).toEqual([{ a: 1 }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// json/writer
// ═══════════════════════════════════════════════════════════════════════

describe("writeJson / writeNdjson", () => {
  it("pretty-prints with a custom indent", () => {
    expect(writeJson([{ a: 1 }], { pretty: true, indent: "\t" })).toBe('[\n\t{\n\t\t"a": 1\n\t}\n]')
  })

  it("emits nothing at all for an empty NDJSON set", () => {
    // Not "\n" — an empty file is a valid empty NDJSON document.
    expect(writeNdjson([])).toBe("")
  })
})

describe("workbookToJson", () => {
  const wb: Workbook = {
    sheets: [
      {
        name: "First",
        rows: [
          ["a", "b"],
          [1, 2],
        ],
      },
      { name: "Second", rows: [["x"], [9]] },
    ],
  }

  it("selects a sheet by name", () => {
    expect(JSON.parse(workbookToJson(wb, { sheet: "Second" }))).toEqual([{ x: 9 }])
  })

  it("selects a sheet by index", () => {
    expect(JSON.parse(workbookToJson(wb, { sheet: 1 }))).toEqual([{ x: 9 }])
  })

  it("names the missing sheet in the error", () => {
    expect(() => workbookToJson(wb, { sheet: "Nope" })).toThrow(`Sheet "Nope" not found`)
    expect(() => workbookToJson(wb, { sheet: 5 })).toThrow("Sheet index 5 out of range")
  })

  it("keys multi-sheet output by sheet name, and pretty-prints it on request", () => {
    const out = workbookToJson(wb, { pretty: true })
    expect(out).toContain('"First"')
    expect(out).toContain("\n  ")
    expect(JSON.parse(out)).toEqual({ First: [{ a: 1, b: 2 }], Second: [{ x: 9 }] })
  })

  it("returns an empty array when the header row is past the end of the sheet", () => {
    const single: Workbook = { sheets: [{ name: "One", rows: [["a"], [1]] }] }
    expect(workbookToJson(single, { headerRow: 5 })).toBe("[]")
  })

  it("uses an empty key for a blank header cell and pads short rows with null", () => {
    const single: Workbook = {
      sheets: [{ name: "One", rows: [["a", null, " c "], [1]] }],
    }
    expect(JSON.parse(workbookToJson(single))).toEqual([{ a: 1, "": null, c: null }])
  })

  it("normalizes an in-row null to null rather than dropping the key", () => {
    const single: Workbook = {
      sheets: [
        {
          name: "One",
          rows: [
            ["a", "b"],
            [1, null],
          ],
        },
      ],
    }
    expect(JSON.parse(workbookToJson(single))).toEqual([{ a: 1, b: null }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// export/json — toJson formats
// ═══════════════════════════════════════════════════════════════════════

describe("toJson — empty and header-less sheets", () => {
  it("returns the shape matching the requested format for an empty sheet", () => {
    expect(toJson(sheet([]))).toBe("[]")
    expect(toJson(sheet([]), { format: "arrays" })).toBe(`{"headers":[],"data":[]}`)
    expect(toJson(sheet([]), { format: "columns" })).toBe("{}")
  })

  it("returns the same empty shapes when headerRow points past the last row", () => {
    // `headerRow: 3` on a two-row sheet is a caller mistake, not a crash.
    const s = sheet([["a"], [1]])
    expect(toJson(s, { headerRow: 3 })).toBe("[]")
    expect(toJson(s, { headerRow: 3, format: "arrays" })).toBe(`{"headers":[],"data":[]}`)
    expect(toJson(s, { headerRow: 3, format: "columns" })).toBe("{}")
  })
})

describe("toJson — formats", () => {
  const s = sheet([
    ["Name", null, " Price "],
    ["Widget", "x", 9.99],
    ["Gadget", null, null],
    ["Sprocket"],
  ])

  it("pads short rows with null in objects format", () => {
    expect(JSON.parse(toJson(s))).toEqual([
      { Name: "Widget", "": "x", Price: 9.99 },
      { Name: "Gadget", "": null, Price: null },
      { Name: "Sprocket", "": null, Price: null },
    ])
  })

  it("pads short rows with null in arrays format", () => {
    expect(JSON.parse(toJson(s, { format: "arrays" }))).toEqual({
      headers: ["Name", "", "Price"],
      data: [
        ["Widget", "x", 9.99],
        ["Gadget", null, null],
        ["Sprocket", null, null],
      ],
    })
  })

  it("pads short rows with null in columns format", () => {
    expect(JSON.parse(toJson(s, { format: "columns" }))).toEqual({
      Name: ["Widget", "Gadget", "Sprocket"],
      "": ["x", null, null],
      Price: [9.99, null, null],
    })
  })

  it("serializes Date cells as ISO strings in every format", () => {
    // JSON has no date type; the exporter pins the representation instead of
    // leaving it to the ambient locale.
    const withDate = sheet([["When"], [new Date("2024-03-01T12:00:00Z")]])
    expect(JSON.parse(toJson(withDate))).toEqual([{ When: "2024-03-01T12:00:00.000Z" }])
    expect(JSON.parse(toJson(withDate, { format: "arrays" })).data).toEqual([
      ["2024-03-01T12:00:00.000Z"],
    ])
    expect(JSON.parse(toJson(withDate, { format: "columns" })).When).toEqual([
      "2024-03-01T12:00:00.000Z",
    ])
  })

  it("emits null for an unparseable Date instead of throwing", () => {
    // `JSON.stringify` calls Date.prototype.toJSON, which returns null for
    // an invalid Date — so this never reaches `toISOString()` and never
    // throws the RangeError #364 fixed in the HTML and ODS writers.
    expect(toJson(sheet([["When"], [new Date("nope")]]))).toBe(`[{"When":null}]`)
  })

  it("indents by two spaces when pretty is set", () => {
    expect(toJson(sheet([["a"], [1]]), { pretty: true })).toContain('\n    "a": 1')
  })

  it("honours headerRow to skip a title row", () => {
    const titled = sheet([["Q1 report"], ["Name", "Qty"], ["Widget", 2]])
    expect(JSON.parse(toJson(titled, { headerRow: 1 }))).toEqual([{ Name: "Widget", Qty: 2 }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// export/html — style translation
// ═══════════════════════════════════════════════════════════════════════

describe("toHtml — inline styles", () => {
  function cssFor(style: Cell["style"]): string {
    const html = toHtml(styledSheet([["v"]], { "0,0": { value: "v", style } }), { styles: true })
    return html.match(/style="([^"]*)"/)?.[1] ?? ""
  }

  it("combines underline and strikethrough into one text-decoration", () => {
    // Two separate `text-decoration` declarations would have the second win,
    // silently dropping the underline.
    expect(cssFor({ font: { underline: true, strikethrough: true } })).toBe(
      "text-decoration:underline line-through",
    )
  })

  it("emits strikethrough on its own when there is no underline", () => {
    expect(cssFor({ font: { strikethrough: true } })).toBe("text-decoration:line-through")
  })

  it("maps alignment, wrapping and font metadata", () => {
    const css = cssFor({
      font: { size: 12, name: "Inter" },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
    })
    expect(css).toBe(
      "font-size:12pt;font-family:Inter;text-align:center;vertical-align:center;white-space:pre-wrap",
    )
  })

  it("omits text-align for the 'general' horizontal alignment", () => {
    // "general" means "let the browser decide", which is the default anyway.
    expect(cssFor({ alignment: { horizontal: "general" } })).toBe("")
  })

  it("emits no colour for a theme or indexed Color the reader left unresolved", () => {
    // XLSX theme colours arrive as `{ theme, tint }` with no rgb; CSS has
    // nothing to say about them, so the declaration is skipped entirely
    // rather than emitting `color:#undefined`.
    expect(
      cssFor({
        font: { bold: true, color: { theme: 1 } },
        fill: { type: "pattern", pattern: "solid", fgColor: { indexed: 64 } },
      }),
    ).toBe("font-weight:bold")
  })

  it("maps each Excel border style onto its CSS width and pattern", () => {
    expect(cssFor({ border: { top: { style: "dashed" } } })).toBe("border-top:1px dashed #000000")
    expect(cssFor({ border: { right: { style: "dotted" } } })).toBe(
      "border-right:1px dotted #000000",
    )
    // `double` changes the pattern but not the width — Excel's "double"
    // is a hairline pair, not a thick rule.
    expect(cssFor({ border: { bottom: { style: "double" } } })).toBe(
      "border-bottom:1px double #000000",
    )
    expect(cssFor({ border: { left: { style: "mediumDashed", color: { rgb: "FF0000" } } } })).toBe(
      "border-left:2px dashed #FF0000",
    )
    expect(cssFor({ border: { top: { style: "thick" } } })).toBe("border-top:3px solid #000000")
  })
})

describe("toHtml — merged header cells", () => {
  it("skips the cells a header-row merge covers", () => {
    // A "2024" heading spanning two columns must produce one <th>, not one
    // <th> plus an empty duplicate.
    const s = sheet(
      [
        ["2024", null],
        [1, 2],
      ],
      { merges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }] },
    )
    const html = toHtml(s, { hasHeaderRow: true })
    expect(html.match(/<th\s/g)!.length).toBe(1)
    expect(html).toContain(`<th scope="col" colspan="2">2024</th>`)
  })
})
