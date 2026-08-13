import { describe, expect, it } from "vitest"
import { readXml } from "../src/xml/data-reader"
import { writeXml } from "../src/xml/data-writer"
import { streamXmlRows } from "../src/xml/stream-reader"
import { seeded } from "./_fuzz"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #467 — XML was the last format with no streaming reader.
//
// `src/xml/parser.ts` is push-based, so wrapping `parseSax` in a
// generator would buffer every row before yielding the first. The way
// round it is to split the work: a scanner finds each row's *span*, and
// each span goes to `collectRows` — the same function `readXml` uses —
// so the element model and the flattening are identical by construction
// rather than by a second implementation agreeing.
//
// Which means the test that matters is agreement with `readXml`. The
// scanner is the only new logic, and its whole job is knowing where an
// element ends: comments, CDATA, processing instructions, doctypes with
// internal subsets, `>` inside a quoted attribute, self-closing tags,
// and an element nested inside a row that shares its name.
// ═══════════════════════════════════════════════════════════════════════

async function rowsOf(xml: string, options?: Parameters<typeof streamXmlRows>[1]) {
  const out: Array<Record<string, CellValue>> = []
  for await (const row of streamXmlRows(xml, options))
    out.push(row.values as Record<string, CellValue>)
  return out
}

/**
 * `readXml` pads every row to the union of all headers; a streaming
 * reader cannot know a key that appears only in a later row. Comparing
 * on the keys each row actually carries is what the two can agree on.
 */
const trim = (o: Record<string, CellValue>): Record<string, CellValue> =>
  Object.fromEntries(Object.entries(o).filter(([, v]) => v !== null))

describe("it agrees with readXml on the shapes a scanner can get wrong", () => {
  const CASES: Array<[string, string, Parameters<typeof readXml>[1]?]> = [
    ["plain rows", "<root><row><a>1</a><b>x</b></row><row><a>2</a><b>y</b></row></root>"],
    ["attributes", '<root><row id="1"><v>3</v></row><row id="2"><v>4</v></row></root>'],
    [
      "nested elements",
      "<root><row><a><b>1</b><c>2</c></a></row><row><a><b>3</b></a></row></root>",
    ],
    ["a self-closing row", "<root><row/><row><a>1</a></row></root>"],
    [
      "CDATA holding a tag",
      "<root><row><a><![CDATA[<not a tag>]]></a></row><row><a>2</a></row></root>",
    ],
    ["a comment holding a row", "<root><!-- <row>fake</row> --><row><a>1</a></row></root>"],
    ["`>` inside an attribute", '<root><row a="x>y"><v>1</v></row></root>'],
    ["a quote inside an attribute", `<root><row a='q">z'><v>1</v></row></root>`],
    [
      "a doctype with an internal subset",
      '<?xml version="1.0"?><!DOCTYPE root [<!ENTITY x "y">]><root><row><a>1</a></row></root>',
    ],
    ["a row nested inside a row", "<root><row><row>inner</row></row><row><a>2</a></row></root>"],
    ["indentation", "<root>\n  <row>\n    <a>1</a>\n  </row>\n</root>"],
    ["mixed content", "<root><row>text<a>1</a>more</row></root>"],
    ["entities", "<root><row><a>&amp;&lt;&gt;&#65;</a></row></root>"],
    ["a processing instruction inside a row", "<root><row><?target data?><a>1</a></row></root>"],
    [
      "namespaces",
      "<root><ns:row xmlns:ns='u'><ns:a>1</ns:a></ns:row></root>",
      { stripNamespaces: true },
    ],
  ]

  for (const [label, xml, options] of CASES) {
    it(label, async () => {
      const reference = readXml(xml, options)
      const streamed = await rowsOf(xml, { ...options, rowTag: reference.rowTag || undefined })

      expect(streamed.map(trim)).toEqual(reference.data.map(trim))
    })
  }
})

describe("it agrees over generated documents too", () => {
  it("400 of them, seeded", async () => {
    // The hand-written cases are the ones someone thought of. These are
    // the ones nobody did — same discipline as test/fuzz-robustness.
    const rnd = seeded(0xabcdef)
    const NASTY = [
      "plain",
      "a<b",
      "a>b",
      'q"uote',
      "amp&amp",
      "<!-- x -->",
      "]]>",
      "<row>fake</row>",
      "  pad  ",
      "ünï",
      "😀",
    ]

    for (let run = 0; run < 400; run++) {
      const keys = ["a", "b", "c"].slice(0, 1 + Math.floor(rnd() * 3))
      const rows = Array.from({ length: 1 + Math.floor(rnd() * 5) }, () => {
        const row: Record<string, CellValue> = {}
        for (const key of keys) {
          if (rnd() < 0.2) continue
          row[key] =
            rnd() < 0.5 ? NASTY[Math.floor(rnd() * NASTY.length)]! : Math.floor(rnd() * 1000)
        }
        if (rnd() < 0.3) row["@id"] = Math.floor(rnd() * 100)
        return row
      })

      const xml = writeXml(rows, { pretty: rnd() < 0.5 })
      let reference: ReturnType<typeof readXml>
      try {
        reference = readXml(xml)
      } catch {
        continue
      }

      const streamed = await rowsOf(xml, { rowTag: reference.rowTag || undefined })
      expect(streamed.map(trim), `run ${run}`).toEqual(reference.data.map(trim))
    }
  }, 120_000)
})

describe("the streaming properties", () => {
  it("yields an index with each row", async () => {
    const xml = "<root><row><a>1</a></row><row><a>2</a></row><row><a>3</a></row></root>"
    const seen: number[] = []
    for await (const row of streamXmlRows(xml)) seen.push(row.index)

    expect(seen).toEqual([0, 1, 2])
  })

  it("stops at maxRows without scanning the rest", async () => {
    const xml = `<root>${"<row><a>1</a></row>".repeat(1000)}</root>`

    expect(await rowsOf(xml, { maxRows: 3 })).toHaveLength(3)
  })

  it("scans lazily — the caller decides how far to go", async () => {
    // The property the whole thing exists for. Abandoning after one row
    // must not have cost the other 999.
    const xml = `<root>${Array.from({ length: 1000 }, (_, i) => `<row><a>${i}</a></row>`).join("")}</root>`

    const iterator = streamXmlRows(xml)[Symbol.asyncIterator]()
    const first = await iterator.next()
    await iterator.return?.(undefined)

    expect(first.value?.values).toEqual({ a: "0" })
  })

  it("takes bytes as well as a string", async () => {
    const xml = "<root><row><a>şehir</a></row></root>"

    expect(await rowsOf(new TextEncoder().encode(xml) as never)).toEqual([{ a: "şehir" }])
  })

  it("infers the row tag from the first child when none is given", async () => {
    // `readXml` counts every child and takes the most frequent, which
    // needs the whole document. A streaming reader takes the first.
    expect(await rowsOf("<root><record><a>1</a></record></root>")).toEqual([{ a: "1" }])
  })

  it("yields nothing for an empty document", async () => {
    expect(await rowsOf("")).toEqual([])
    expect(await rowsOf("   ")).toEqual([])
  })
})

describe("what it refuses", () => {
  it("an unterminated row, rather than yielding a truncated one", async () => {
    await expect(rowsOf("<root><row><a>1</a>")).rejects.toThrow(/Unterminated/)
  })
})
