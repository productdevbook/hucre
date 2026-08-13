// `writeXml` turns a column heading into an element name, and used to
// accept only ASCII ones. XML 1.0 §2.3's `NameStartChar` begins at #xC0,
// so every accented or non-Latin heading was refused — a spreadsheet
// whose columns are not named in English could not be written to XML at
// all. It threw rather than mangling, so nothing silently escaped; the
// format was simply unavailable.

import { describe, expect, it } from "vitest"
import { writeXml, writeXmlStream } from "../src/xml/data-writer"
import { readXml } from "../src/xml/data-reader"
import { ParseError } from "../src/errors"

const drain = async (stream: ReadableStream<Uint8Array>): Promise<string> => {
  let out = ""
  const decoder = new TextDecoder()
  for await (const chunk of stream as unknown as AsyncIterable<Uint8Array>) {
    out += decoder.decode(chunk, { stream: true })
  }
  return out + decoder.decode()
}

describe("element names XML allows", () => {
  const accepted = [
    ["Şehir", "Turkish"],
    ["Ünvan", "Turkish, U+00DC"],
    ["Größe", "German"],
    ["café", "French"],
    ["naïve", "French, combining-range neighbour"],
    ["名前", "Japanese"],
    ["日付", "Japanese"],
    ["Ελλάδα", "Greek"],
    ["Кириллица", "Cyrillic"],
    ["ns:Şehir", "prefixed"],
    ["a·b", "U+00B7, a NameChar but not a NameStartChar"],
    ["_x", "underscore start"],
    ["a-b", "hyphen"],
  ] as const

  for (const [name, why] of accepted) {
    it(`accepts ${name} (${why})`, () => {
      const xml = writeXml([{ [name]: 1 }], { rootTag: "rows", rowTag: "row" })
      expect(xml).toContain(`<${name}>`)
    })
  }

  it("round-trips a non-ASCII name through the reader", () => {
    const xml = writeXml([{ Şehir: "İzmir", Ürün: 3 }])
    const { data, headers } = readXml(xml)
    // Values come back as text — `readXml` does no type inference by
    // default. The name is what this is about.
    expect(data[0]).toEqual({ Şehir: "İzmir", Ürün: "3" })
    expect(headers).toEqual(["Şehir", "Ürün"])
  })

  it("leaves the dot doing its own job", () => {
    // A dot in a key is this writer's nesting separator, not part of the
    // name — unchanged by any of this.
    expect(writeXml([{ "a.b": 1 }])).toContain("<a><b>1</b></a>")
  })

  it("takes one as rootTag and rowTag", () => {
    const xml = writeXml([{ a: 1 }], { rootTag: "şehirler", rowTag: "şehir" })
    expect(xml).toContain("<şehirler>")
    expect(xml).toContain("<şehir>")
  })
})

describe("element names XML forbids", () => {
  const rejected = [
    ["1st", "leading digit"],
    ["-x", "leading hyphen"],
    [".x", "leading dot"],
    ["·x", "U+00B7 leading — a NameChar, not a NameStartChar"],
    ["a b", "space"],
    ["a<b", "angle bracket"],
    ["a:b:c", "two colons"],
    ["", "empty"],
    ["́x", "leading combining acute — NameChar only"],
  ] as const

  for (const [name, why] of rejected) {
    it(`rejects ${JSON.stringify(name)} (${why})`, () => {
      expect(() => writeXml([{ [name]: 1 }])).toThrow(ParseError)
    })
  }
})

describe("the streaming writer agrees", () => {
  it("accepts the same names", async () => {
    const out = await drain(writeXmlStream([{ Şehir: "İzmir" }], { rootTag: "kayıtlar" }))
    expect(out).toContain("<Şehir>")
    expect(out).toContain("<kayıtlar>")
  })

  it("rejects the same names", async () => {
    // The generator is lazy, so the name is checked on the first pull
    // rather than at the call — draining is what asks the question.
    await expect(drain(writeXmlStream([{ a: 1 }], { rootTag: "1st" }))).rejects.toThrow(ParseError)
  })
})
