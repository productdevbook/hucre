import { describe, expect, it } from "vitest"
import { streamXmlRows } from "../src/xml/stream-reader"
import { readXml } from "../src/xml/data-reader"
import { HucreError } from "../src/errors"
import { seeded, XML_MUTATORS } from "./_fuzz"
import type { CellValue } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// `streamXmlRows` (#467) was the last reader added and the only one with
// a hand-rolled scanner: it finds the *span* of each row element itself
// and hands the span to `collectRows`, the same function `readXml` uses.
// Two things that arrangement has to be held to, and neither was tested.
//
// **It has to survive a broken document.** The scanner tracks quotes,
// comments, CDATA, doctypes with internal subsets and self-closing tags
// by hand, and any of those going wrong on malformed input is a raw
// error or a hang rather than a `ParseError`.
//
// **It has to agree with `readXml`, except where it provably cannot.**
// There are two such places, and they have the same cause — knowing the
// answer means having read the whole document:
//
//   - the row tag, when the caller does not name one. Documented already.
//   - **the key set.** `readXml` returns a rectangle: the union of every
//     row's keys, gaps filled with `null`. A streaming reader would have
//     to reach the last row to know the union, so each row carries only
//     the keys it had.
//
// The second was not written down anywhere, and it is the one that bites:
// code moved from `readXml` reads `values.note` and gets `undefined`
// where it used to get `null`.
// ═══════════════════════════════════════════════════════════════════════

/** Comments, CDATA, a doctype with a subset, `>` inside an attribute,
 *  a nested element sharing the row's name, and an empty row. */
const SOURCE = `<?xml version="1.0"?>
<!DOCTYPE records [<!ENTITY x "y">]>
<records>
  <!-- a comment with < and > in it -->
  <record id="1"><name>Widget</name><qty>3</qty><note><![CDATA[raw <b> text]]></note></record>
  <record id="2"><name>Gadget &amp; co</name><qty>7</qty><nested><record>inner</record></nested></record>
  <record id="3" attr="has > gt"><name>padded</name><qty>0</qty></record>
  <record/>
</records>`

async function stream(xml: string): Promise<Array<Record<string, CellValue>>> {
  const out: Array<Record<string, CellValue>> = []
  for await (const row of streamXmlRows(xml, { rowTag: "record" })) out.push(row.values)
  return out
}

describe("the scanner reads what the parser reads", () => {
  it("finds every row, including the empty one", async () => {
    const rows = await stream(SOURCE)

    expect(rows).toHaveLength(4)
    expect(rows[0]!.name).toBe("Widget")
    expect(rows[1]!.name).toBe("Gadget & co")
  })

  it("is not fooled by a > inside an attribute", async () => {
    const rows = await stream(SOURCE)

    expect(rows[2]!["@attr"]).toBe("has > gt")
    expect(rows[2]!["@id"]).toBe("3")
  })

  it("keeps CDATA verbatim and does not start a row inside one", async () => {
    const rows = await stream(SOURCE)

    expect(rows[0]!.note).toBe("raw <b> text")
  })

  it("does not let a nested <record> open a second row", async () => {
    // The scanner tracks depth for exactly this: an element inside a row
    // that shares its name must not start another.
    const rows = await stream(SOURCE)

    expect(rows).toHaveLength(4)
    expect(rows[1]!["nested.record"]).toBe("inner")
  })

  it("agrees with readXml on every value both of them have", async () => {
    const streamed = await stream(SOURCE)
    const buffered = readXml(SOURCE, { rowTag: "record" }).data

    expect(streamed).toHaveLength(buffered.length)
    for (let i = 0; i < streamed.length; i++) {
      for (const [key, value] of Object.entries(streamed[i]!)) {
        expect(value, `row ${i} key ${key}`).toEqual(buffered[i]![key])
      }
    }
  })
})

describe("the one difference, which is not a defect", () => {
  it("a row carries only its own keys, where readXml pads to the union", async () => {
    const streamed = await stream(SOURCE)
    const buffered = readXml(SOURCE, { rowTag: "record" }).data

    // Row 0 has no `nested.record`; row 1 has no `note`.
    expect("nested.record" in streamed[0]!).toBe(false)
    expect(buffered[0]!["nested.record"]).toBeNull()
    expect("note" in streamed[1]!).toBe(false)
    expect(buffered[1]!.note).toBeNull()
  })

  it("and an empty row is an empty object, not a row of nulls", async () => {
    const streamed = await stream(SOURCE)
    const buffered = readXml(SOURCE, { rowTag: "record" }).data

    expect(streamed[3]).toEqual({})
    expect(Object.keys(buffered[3]!).length).toBeGreaterThan(0)
    expect(Object.values(buffered[3]!).every((v) => v === null)).toBe(true)
  })

  it("so `?? null` is what moves code across", async () => {
    const streamed = await stream(SOURCE)

    expect(streamed[0]!["nested.record"] ?? null).toBeNull()
  })
})

describe("a broken document is a typed error, not a crash", () => {
  it("survives every mutator without a raw throw or a hang", async () => {
    // The bar `src/limits.ts` already sets: a typed error rather than
    // killing the process. Seeded, and bounded to seconds.
    const raw: string[] = []
    const rnd = seeded(0xf00d)

    for (let run = 0; run < 40; run++) {
      for (const [label, mutate] of XML_MUTATORS) {
        try {
          await stream(mutate(SOURCE, rnd))
        } catch (error) {
          if (!(error instanceof HucreError) && raw.length < 10) {
            const e = error as Error
            raw.push(`${label} run ${run}: ${e.constructor.name}: ${e.message.slice(0, 60)}`)
          }
        }
      }
    }

    expect(raw).toEqual([])
  })

  it("an unterminated row is a ParseError naming the tag", async () => {
    await expect(stream("<records><record><name>x</name></records>")).rejects.toThrow(
      /Unterminated <record>/,
    )
  })

  it("empty and whitespace-only input yield nothing", async () => {
    expect(await stream("")).toEqual([])
    expect(await stream("   \n  ")).toEqual([])
  })
})
