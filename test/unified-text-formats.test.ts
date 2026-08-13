import { describe, expect, it } from "vitest"
import { read, write } from "../src/defter"
import { detectTextFormat } from "../src/_sniff"
import { writeXlsx } from "../src/xlsx/writer"
import { parseCsv } from "../src/csv/reader"
import { UnsupportedFormatError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// #469 — the library reads and/or writes nine formats. The two functions
// that exist to be format-agnostic covered four and two of them.
//
// `read(csvBytes)` answered "Unsupported format: unknown (not a ZIP
// archive)" — a correct error for a function that cannot do the job,
// under a README heading that says **Unified API**. A CSV file read off
// disk is a `Uint8Array` like any other.
//
// The detection decides rather than guesses: every format here announces
// itself in its first non-whitespace character or two, which is a far
// weaker claim than the delimiter auto-detection the CSV reader already
// makes. CSV is the fallback, because "text that is not any of the
// others" is what CSV actually is.
// ═══════════════════════════════════════════════════════════════════════

const enc = (s: string): Uint8Array => new TextEncoder().encode(s)
const dec = (b: Uint8Array): string => new TextDecoder().decode(b)

const CSV = "name,qty\nWidget,3\nGadget,7\n"
const JSON_ARRAY = '[{"name":"Widget","qty":3},{"name":"Gadget","qty":7}]'
const NDJSON = '{"name":"Widget","qty":3}\n{"name":"Gadget","qty":7}\n'
const XML = '<?xml version="1.0"?><rows><row><name>Widget</name><qty>3</qty></row></rows>'
const HTML = "<table><tr><th>name</th></tr><tr><td>Widget</td></tr></table>"

describe("detectTextFormat", () => {
  it("names each format from its opening", () => {
    expect(detectTextFormat(enc(CSV))).toBe("csv")
    expect(detectTextFormat(enc(JSON_ARRAY))).toBe("json")
    expect(detectTextFormat(enc(NDJSON))).toBe("ndjson")
    expect(detectTextFormat(enc(XML))).toBe("xml")
    expect(detectTextFormat(enc(HTML))).toBe("html")
  })

  it("tells one JSON document from many", () => {
    // Both open with `{`. A whole document that parses is JSON; one that
    // does not, whose lines each do, is NDJSON — which is exactly the
    // shape that makes the whole fail.
    expect(detectTextFormat(enc('{"a":1}'))).toBe("json")
    expect(detectTextFormat(enc('{"a":1}\n{"a":2}'))).toBe("ndjson")

    // A pretty-printed object spans lines and is still one document.
    expect(detectTextFormat(enc('{\n  "a": 1\n}'))).toBe("json")
  })

  it("tells HTML from XML", () => {
    expect(detectTextFormat(enc("<!DOCTYPE html><html><body></body></html>"))).toBe("html")
    expect(detectTextFormat(enc("<html>"))).toBe("html")
    expect(detectTextFormat(enc("<table><tr><td>a</td></tr></table>"))).toBe("html")
    expect(detectTextFormat(enc('<?xml version="1.0"?><a/>'))).toBe("xml")
    expect(detectTextFormat(enc("<rows><row/></rows>"))).toBe("xml")
  })

  it("skips a byte-order mark before deciding", () => {
    const bom = (s: string): Uint8Array => new Uint8Array([0xef, 0xbb, 0xbf, ...enc(s)])

    expect(detectTextFormat(bom(CSV))).toBe("csv")
    expect(detectTextFormat(bom(JSON_ARRAY))).toBe("json")
  })

  it("says nothing about bytes that are not text", () => {
    // The important return. Without it, binary rubbish would reach the
    // CSV parser, which would cheerfully make a sheet of it.
    expect(detectTextFormat(new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7]))).toBeNull()
    expect(detectTextFormat(new Uint8Array([]))).toBeNull()
    expect(detectTextFormat(enc("   \n  "))).toBeNull()
  })
})

describe("read() dispatches to the text formats", () => {
  it("CSV", async () => {
    const wb = await read(enc(CSV))

    expect(wb.sheets[0]!.rows).toEqual([
      ["name", "qty"],
      ["Widget", "3"],
      ["Gadget", "7"],
    ])
  })

  it("JSON", async () => {
    const wb = await read(enc(JSON_ARRAY))

    expect(wb.sheets[0]!.rows[0]).toEqual(["name", "qty"])
    expect(wb.sheets[0]!.rows[1]).toEqual(["Widget", 3])
  })

  it("NDJSON, with the header row put back", async () => {
    // A workbook is a grid. The record readers hand back
    // `{ data, headers }`, and dropping the names would lose them.
    const wb = await read(enc(NDJSON))

    expect(wb.sheets[0]!.rows).toEqual([
      ["name", "qty"],
      ["Widget", 3],
      ["Gadget", 7],
    ])
  })

  it("XML", async () => {
    const wb = await read(enc(XML))

    expect(wb.sheets[0]!.rows[0]).toEqual(["name", "qty"])
    expect(wb.sheets[0]!.rows[1]![0]).toBe("Widget")
  })

  it("HTML", async () => {
    const wb = await read(enc(HTML))

    expect(wb.sheets[0]!.rows).toEqual([["name"], ["Widget"]])
  })

  it("still reads the container formats it always did", async () => {
    const bytes = await writeXlsx({ sheets: [{ name: "S", rows: [["a", 1]] }] })

    expect((await read(bytes)).sheets[0]!.rows).toEqual([["a", 1]])
  })

  it("still refuses bytes that are not a spreadsheet at all", async () => {
    await expect(read(new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]))).rejects.toThrow(
      UnsupportedFormatError,
    )
  })
})

describe("write() covers what the library can write", () => {
  const sheets = [
    {
      name: "S",
      rows: [
        ["name", "qty"],
        ["Widget", 3],
      ],
    },
  ]

  it("the text formats each produce their own shape", async () => {
    expect(dec(await write({ sheets, format: "csv" }))).toContain("name,qty")
    expect(dec(await write({ sheets, format: "tsv" }))).toContain("name\tqty")
    expect(dec(await write({ sheets, format: "json" }))).toContain('"name"')
    expect(
      dec(await write({ sheets, format: "ndjson" }))
        .trim()
        .split("\n"),
    ).toHaveLength(1)
    expect(dec(await write({ sheets, format: "xml" }))).toContain("<name>")
    expect(dec(await write({ sheets, format: "html" }))).toContain("<table")
    expect(dec(await write({ sheets, format: "markdown" }))).toContain("| name")
  })

  it("returns bytes for every format, so the caller does not branch", async () => {
    for (const format of ["xlsx", "ods", "csv", "json", "html"] as const) {
      expect(await write({ sheets, format }), format).toBeInstanceOf(Uint8Array)
    }
  })

  it("still defaults to xlsx", async () => {
    const bytes = await write({ sheets })

    expect((await read(bytes)).sheets[0]!.rows[1]).toEqual(["Widget", 3])
  })

  it("round-trips through the text formats it can read back", async () => {
    for (const format of ["csv", "json", "ndjson", "xml"] as const) {
      const wb = await read(await write({ sheets, format }))

      expect(wb.sheets[0]!.rows[0], format).toEqual(["name", "qty"])
      expect(String(wb.sheets[0]!.rows[1]![0]), format).toBe("Widget")
    }
  })

  it("says so rather than crashing on a workbook with no sheets", async () => {
    await expect(write({ sheets: [], format: "csv" })).rejects.toThrow(UnsupportedFormatError)
  })
})

describe("the header-row convention is the same in both directions", () => {
  it("a grid written and read back keeps its names", async () => {
    const sheets = [
      {
        name: "S",
        rows: [
          ["a", "b"],
          [1, 2],
        ],
      },
    ]
    const csv = dec(await write({ sheets, format: "csv" }))

    expect(parseCsv(csv)[0]).toEqual(["a", "b"])
  })

  it("an unnamed first-row cell gets a positional name rather than being dropped", async () => {
    const sheets = [
      {
        name: "S",
        rows: [
          ["a", null],
          [1, 2],
        ],
      },
    ]

    expect(dec(await write({ sheets, format: "json" }))).toContain("column2")
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Every text writer takes an options bag; `write` called all seven with
// none. So the entry #469 added precisely so one call could reach all
// nine formats was the only way to reach seven of them that could not
// configure any of them — including `bom: true`, which #475 documents as
// the answer to Excel opening a UTF-8 CSV as its system code page.
// ═══════════════════════════════════════════════════════════════════════

describe("text-format options reach their writer", () => {
  const sheet = {
    name: "S",
    rows: [
      ["Şehir", "Ürün"],
      ["İzmir", 3],
    ],
  }

  it("csv: delimiter and bom", async () => {
    const bytes = (await write({
      sheets: [sheet],
      format: "csv",
      csv: { delimiter: ";", bom: true },
    })) as Uint8Array

    expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
    expect(dec(bytes)).toContain("Şehir;Ürün")
  })

  it("csv: escapeFormulae", async () => {
    const out = dec(
      (await write({
        sheets: [{ name: "S", rows: [["=1+1"]] }],
        format: "csv",
        csv: { escapeFormulae: true },
      })) as Uint8Array,
    )
    expect(out).toContain("'=1+1")
  })

  it("tsv: bom, with the tab still the delimiter", async () => {
    const bytes = (await write({
      sheets: [sheet],
      format: "tsv",
      tsv: { bom: true },
    })) as Uint8Array
    expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
    expect(dec(bytes)).toContain("Şehir\tÜrün")
  })

  it("json: pretty and indent", async () => {
    const out = dec(
      (await write({
        sheets: [sheet],
        format: "json",
        json: { pretty: true, indent: "    " },
      })) as Uint8Array,
    )
    expect(out).toContain("\n    ")
  })

  it("xml: rootTag and rowTag", async () => {
    // ASCII headers: `writeXml` rejects a non-ASCII element name, which
    // XML 1.0 §2.3 allows. Tracked separately.
    const out = dec(
      (await write({
        sheets: [
          {
            name: "S",
            rows: [
              ["city", "qty"],
              ["Izmir", 3],
            ],
          },
        ],
        format: "xml",
        xml: { rootTag: "cities", rowTag: "city_row" },
      })) as Uint8Array,
    )
    expect(out).toContain("<cities>")
    expect(out).toContain("<city_row>")
  })

  it("html: caption and header row", async () => {
    const out = dec(
      (await write({
        sheets: [sheet],
        format: "html",
        html: { caption: "Şehirler", hasHeaderRow: true },
      })) as Uint8Array,
    )
    expect(out).toContain("<caption>Şehirler</caption>")
    expect(out).toContain("<thead>")
  })

  it("markdown: alignment", async () => {
    const out = dec(
      (await write({
        sheets: [sheet],
        format: "markdown",
        markdown: { alignment: ["right", "right"] },
      })) as Uint8Array,
    )
    expect(out).toContain("--:")
  })

  it("leaves the defaults alone when no bag is passed", async () => {
    const out = dec((await write({ sheets: [sheet], format: "csv" })) as Uint8Array)
    expect(out.startsWith("Şehir,Ürün")).toBe(true)
  })
})
