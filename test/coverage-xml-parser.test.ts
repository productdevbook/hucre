import { describe, expect, it } from "vitest"
import { parseSax, parseSaxStream, parseXml, SAX_TEXT_FLUSH_CHARS } from "../src/xml/parser"
import type { SaxHandlers } from "../src/xml/parser"
import { readXml } from "../src/xml/data-reader"
import { XmlError, ParseError } from "../src/errors"

const enc = new TextEncoder()

// ── Helpers ──────────────────────────────────────────────────────────

/** Feed the streaming parser a fixed sequence of chunks, one per pull. */
function chunkStream(chunks: (string | Uint8Array)[]): ReadableStream<Uint8Array> {
  const bytes = chunks.map((c) => (typeof c === "string" ? enc.encode(c) : c))
  let i = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (i >= bytes.length) {
        controller.close()
        return
      }
      controller.enqueue(bytes[i++])
    },
  })
}

/** Record every SAX event the streaming parser emits for `chunks`. */
async function streamEvents(chunks: (string | Uint8Array)[]): Promise<{
  open: Array<[string, Record<string, string>]>
  close: string[]
  text: string[]
  cdata: string[]
}> {
  const out = {
    open: [] as Array<[string, Record<string, string>]>,
    close: [] as string[],
    text: [] as string[],
    cdata: [] as string[],
  }
  const handlers: SaxHandlers = {
    onOpenTag: (tag, attrs) => out.open.push([tag, attrs]),
    onCloseTag: (tag) => out.close.push(tag),
    onText: (text) => out.text.push(text),
    onCData: (text) => out.cdata.push(text),
  }
  await parseSaxStream(chunkStream(chunks), handlers)
  return out
}

/** Collect only the attribute map of the first element in `xml`. */
function attrsOf(xml: string): Record<string, string> {
  let found: Record<string, string> = {}
  let first = true
  parseSax(xml, {
    onOpenTag(_tag, attrs) {
      if (first) {
        found = attrs
        first = false
      }
    },
  })
  return found
}

// ═══════════════════════════════════════════════════════════════════════
// Attribute parsing — the lenient paths
// ═══════════════════════════════════════════════════════════════════════

describe("parseSax — attribute syntax the spec does not bless", () => {
  it("stops at an attribute list that begins with '='", () => {
    // Nothing readable as a name: bail out rather than loop forever.
    expect(attrsOf(`<c ="orphan"/>`)).toEqual({})
  })

  it("records a valueless attribute as an empty string", () => {
    // HTML-style boolean attributes turn up in hand-edited OOXML parts.
    expect(attrsOf(`<c hidden/>`)).toEqual({ hidden: "" })
  })

  it("records a valueless attribute that is followed by a real one", () => {
    expect(attrsOf(`<c hidden r="A1"/>`)).toEqual({ hidden: "", r: "A1" })
  })

  it("ignores a trailing '=' with no value behind it", () => {
    expect(attrsOf(`<c r=/>`)).toEqual({})
  })

  it("accepts an unquoted attribute value", () => {
    expect(attrsOf(`<c r=A1/>`)).toEqual({ r: "A1" })
  })

  it("terminates an unquoted value at the next space", () => {
    expect(attrsOf(`<c r=A1 t=s/>`)).toEqual({ r: "A1", t: "s" })
  })
})

// ═══════════════════════════════════════════════════════════════════════
// parseSax — truncated markup
// ═══════════════════════════════════════════════════════════════════════

describe("parseSax — truncated markup", () => {
  it("reports a document that ends on a bare '<'", () => {
    expect(() => parseSax("<a></a><", {})).toThrow(XmlError)
    expect(() => parseSax("<a></a><", {})).toThrow(/Unterminated opening tag/)
  })

  it("reports a declaration that never closes", () => {
    expect(() => parseSax(`<!DOCTYPE workbook`, {})).toThrow(/Unterminated declaration/)
  })

  it("reports a closing tag with no matching open tag", () => {
    expect(() => parseXml(`</row>`)).toThrow(/Unexpected closing tag: no matching open tag/)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// parseSaxStream — constructs split across chunk boundaries
// ═══════════════════════════════════════════════════════════════════════

describe("parseSaxStream — chunk boundaries", () => {
  it("strips a BOM that arrives in the first chunk", async () => {
    const bom = new Uint8Array([0xef, 0xbb, 0xbf])
    const events = await streamEvents([bom, `<root><a>1</a></root>`])
    expect(events.open.map((o) => o[0])).toEqual(["root", "a"])
  })

  it("holds back a lone '<' at the end of a chunk", async () => {
    const events = await streamEvents([`<root><`, `a r="1"/></root>`])
    expect(events.open).toEqual([
      ["root", {}],
      ["a", { r: "1" }],
    ])
  })

  it("rejoins a comment split across chunks", async () => {
    const events = await streamEvents([`<root><!-- part one`, ` part two --><a/></root>`])
    expect(events.open.map((o) => o[0])).toEqual(["root", "a"])
    expect(events.text).toEqual([])
  })

  it("rejoins a CDATA section split across chunks", async () => {
    const events = await streamEvents([`<root><![CDATA[first`, `second]]></root>`])
    expect(events.cdata).toEqual(["firstsecond"])
  })

  it("holds back a partial '<!' marker too short to classify", async () => {
    // Fewer than 9 characters remain, so it could still become <![CDATA[.
    const events = await streamEvents([`<root><![C`, `DATA[body]]></root>`])
    expect(events.cdata).toEqual(["body"])
  })

  it("rejoins a DOCTYPE declaration split across chunks", async () => {
    const events = await streamEvents([`<!DOCTYPE workbook SYSTEM`, ` "wb.dtd"><root/>`])
    expect(events.open.map((o) => o[0])).toEqual(["root"])
  })

  it("rejoins a processing instruction split across chunks", async () => {
    const events = await streamEvents([`<?xml version="1.0"`, ` encoding="UTF-8"?><root/>`])
    expect(events.open.map((o) => o[0])).toEqual(["root"])
  })

  it("rejoins a closing tag split across chunks", async () => {
    const events = await streamEvents([`<root><a>x</a`, `></root>`])
    expect(events.close).toEqual(["a", "root"])
  })

  it("rejoins an opening tag split inside an attribute value", async () => {
    const events = await streamEvents([`<root><c r="A`, `1" t="s"/></root>`])
    expect(events.open[1]).toEqual(["c", { r: "A1", t: "s" }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// parseSaxStream — truncated tail
// ═══════════════════════════════════════════════════════════════════════

describe("parseSaxStream — a source that stops mid-construct", () => {
  // The streaming parser has no error contract for a truncated document —
  // it drops the unfinished construct and returns normally, leaving the
  // caller (e.g. the XLSX row reader) to notice the missing close tags.

  it("drops an unterminated comment", async () => {
    const events = await streamEvents([`<root><a/><!-- cut off here`])
    expect(events.open.map((o) => o[0])).toEqual(["root", "a"])
  })

  it("drops an unterminated CDATA section", async () => {
    const events = await streamEvents([`<root><a/><![CDATA[cut off here`])
    expect(events.cdata).toEqual([])
  })

  it("drops an unterminated declaration", async () => {
    const events = await streamEvents([`<root><a/><!DOCTYPE cut off here`])
    expect(events.open.map((o) => o[0])).toEqual(["root", "a"])
  })

  it("drops a '<!' marker that never became anything", async () => {
    const events = await streamEvents([`<root><a/><!X`])
    expect(events.open.map((o) => o[0])).toEqual(["root", "a"])
  })

  it("drops an unterminated processing instruction", async () => {
    const events = await streamEvents([`<root><a/><?php echo`])
    expect(events.open.map((o) => o[0])).toEqual(["root", "a"])
  })

  it("drops an unterminated closing tag", async () => {
    const events = await streamEvents([`<root><a/></roo`])
    expect(events.close).toEqual(["a"])
  })

  it("drops an unterminated opening tag", async () => {
    const events = await streamEvents([`<root><c r="A1`])
    expect(events.open.map((o) => o[0])).toEqual(["root"])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// parseSaxStream — very long text runs
// ═══════════════════════════════════════════════════════════════════════

describe("parseSaxStream — text runs past the flush threshold", () => {
  // A text run carried across every chunk is re-copied each time, which is
  // quadratic. Past SAX_TEXT_FLUSH_CHARS the parser emits what it safely
  // can and keeps only the tail — so handlers must accumulate, not assign.
  const LONG = "x".repeat(SAX_TEXT_FLUSH_CHARS + 40000)

  it("splits one long run into several onText calls", async () => {
    const events = await streamEvents([`<r>${LONG}`, `tail</r>`])
    expect(events.text.length).toBeGreaterThan(1)
    expect(events.text.join("")).toBe(`${LONG}tail`)
  })

  it("never cuts inside an entity reference", async () => {
    // The chunk ends mid-"&amp;": splitting there would emit a literal
    // "&am" and then "p;" instead of a single "&".
    const events = await streamEvents([`<r>${LONG}&am`, `p;tail</r>`])
    expect(events.text.join("")).toBe(`${LONG}&tail`)
  })

  it("keeps an astral character whole across the split", async () => {
    const events = await streamEvents([`<r>${LONG}\u{1F600}`, `tail</r>`])
    expect(events.text.join("")).toBe(`${LONG}\u{1F600}tail`)
  })

  it("does not split a run that is merely long, not enormous", async () => {
    const short = "y".repeat(1000)
    const events = await streamEvents([`<r>${short}`, `tail</r>`])
    expect(events.text).toEqual([`${short}tail`])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// readXml — row detection and options
// ═══════════════════════════════════════════════════════════════════════

describe("readXml — row tag selection", () => {
  it("returns nothing when the caller pins an empty row tag", () => {
    const result = readXml(`<feed><item><a>1</a></item></feed>`, { rowTag: "" })
    expect(result).toEqual({ data: [], headers: [], rowTag: "" })
  })

  it("prefers a later tag that repeats more often than the first", () => {
    // Feeds often open with a single <header> element before the records.
    const xml = `<feed><header>meta</header><item><a>1</a></item><item><a>2</a></item></feed>`
    const result = readXml(xml)
    expect(result.rowTag).toBe("item")
    expect(result.data).toEqual([{ a: "1" }, { a: "2" }])
  })

  it("throws when the document has a root but no child elements", () => {
    expect(() => readXml(`<feed>text only</feed>`)).toThrow(ParseError)
  })

  it("ignores CDATA that sits outside any row", () => {
    const xml = `<feed><![CDATA[ignore me]]><item><a>1</a></item><item><a>2</a></item></feed>`
    expect(readXml(xml).data).toEqual([{ a: "1" }, { a: "2" }])
  })

  it("stores a text-only row under the text key", () => {
    const result = readXml(`<feed><sku>A-1</sku><sku>A-2</sku></feed>`)
    expect(result.headers).toEqual(["#text"])
    expect(result.data).toEqual([{ "#text": "A-1" }, { "#text": "A-2" }])
  })

  it("fills missing columns with null after transformHeader renames them", () => {
    // Rows in a real feed are heterogeneous; the renamed header still has to
    // exist on every row.
    const xml = `<feed><item><a>1</a></item><item><b>2</b></item></feed>`
    const result = readXml(xml, { transformHeader: (h) => h.toUpperCase() })
    expect(result.headers).toEqual(["A", "B"])
    expect(result.data).toEqual([
      { A: "1", B: null },
      { A: null, B: "2" },
    ])
  })
})

describe("readXml — mixed content", () => {
  it("captures text that sits alongside child elements", () => {
    const xml = `<feed><item>leading note<a>1</a></item></feed>`
    expect(readXml(xml).data).toEqual([{ a: "1", "#text": "leading note" }])
  })

  it("captures mixed text inside a nested element under its dot path", () => {
    const xml = `<feed><item><desc>note<b>bold</b></desc></item></feed>`
    expect(readXml(xml).data).toEqual([{ "desc.b": "bold", "desc.#text": "note" }])
  })
})

// ═══════════════════════════════════════════════════════════════════════
// readXml — flatten: false
// ═══════════════════════════════════════════════════════════════════════

describe("readXml — flatten:false serialisation", () => {
  it("keeps top-level child names and stringifies their subtrees", () => {
    const xml = `<feed><item><specs><spec>1</spec><spec>2</spec><spec>3</spec></specs></item></feed>`
    const result = readXml(xml, { flatten: false })
    expect(JSON.parse(String(result.data[0].specs))).toEqual({ spec: ["1", "2", "3"] })
  })

  it("nests repeated grandchildren as arrays", () => {
    const xml = `<feed><item><specs><group><spec>1</spec><spec>2</spec><spec>3</spec></group></specs></item></feed>`
    const result = readXml(xml, { flatten: false })
    expect(JSON.parse(String(result.data[0].specs))).toEqual({ group: { spec: ["1", "2", "3"] } })
  })

  it("keeps an attribute-carrying leaf as an object with its text", () => {
    const xml = `<feed><item><price currency="EUR">10.50</price></item></feed>`
    const result = readXml(xml, { flatten: false })
    expect(JSON.parse(String(result.data[0].price))).toEqual({
      "@currency": "EUR",
      "#text": "10.50",
    })
  })

  it("keeps an empty attribute-carrying leaf as attributes alone", () => {
    const xml = `<feed><item><price currency="EUR"/></item></feed>`
    const result = readXml(xml, { flatten: false })
    expect(JSON.parse(String(result.data[0].price))).toEqual({ "@currency": "EUR" })
  })

  it("wraps a nested element that carries attributes in an object", () => {
    const xml = `<feed><item><specs><spec code="w">10</spec></specs></item></feed>`
    const result = readXml(xml, { flatten: false })
    expect(JSON.parse(String(result.data[0].specs))).toEqual({
      spec: { "@code": "w", "#text": "10" },
    })
  })

  it("emits null for an empty top-level child", () => {
    const xml = `<feed><item><a/><b>1</b></item></feed>`
    expect(readXml(xml, { flatten: false }).data[0]).toEqual({ a: null, b: "1" })
  })

  it("captures mixed text on the row root", () => {
    const xml = `<feed><item>loose text<a>1</a></item></feed>`
    expect(readXml(xml, { flatten: false }).data[0]).toEqual({ a: "1", "#text": "loose text" })
  })

  it("strips namespace prefixes from the stringified child names", () => {
    const xml = `<feed xmlns:g="urn:g"><item><g:specs><g:spec>1</g:spec></g:specs></item></feed>`
    const result = readXml(xml, { flatten: false, stripNamespaces: true })
    expect(Object.keys(result.data[0])).toEqual(["specs"])
  })
})
