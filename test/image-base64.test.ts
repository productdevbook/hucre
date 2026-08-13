import { describe, expect, it } from "vitest"
import { imageFromBase64 } from "../src/image"
import { InvalidArgumentError } from "../src/errors"

// ═══════════════════════════════════════════════════════════════════════
// `src/image.ts` was the least-covered file in the library at 45%, and
// for a reason worth fixing rather than papering over: it had two decode
// paths, `globalThis.Buffer` when it was there and `atob` otherwise, so
// the tests only ever ran one of them.
//
// It was also the single place in `src/` outside the CLI that reached for
// a Node API. CLAUDE.md: "src/ outside the CLI uses Web APIs only", and
// `tsconfig.json` sets `"types": []` so that reaching for one is a
// compile error. An `any` cast got past it.
//
// `atob` is a Web API and is in every runtime this library claims. One
// path now, and it is the portable one.
//
// The two disagree on invalid input, which is the part worth keeping.
// Given "YWJj###", `Buffer.from(…, "base64")` silently returns the three
// bytes it managed; `atob` throws. Silence there means a workbook
// carrying a garbage PNG that Excel refuses to open, a long way from the
// bad string that caused it.
// ═══════════════════════════════════════════════════════════════════════

/** A 1x1 transparent PNG. */
const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="

const ANCHOR = { type: "oneCell", from: { row: 0, col: 0 } } as const

describe("valid base64 decodes to the bytes it encoded", () => {
  it("a real PNG, header and all", () => {
    const image = imageFromBase64(PNG, "png", ANCHOR)

    // The PNG signature: 89 50 4E 47 0D 0A 1A 0A.
    expect([...image.data.slice(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    expect(image.type).toBe("png")
  })

  it("through a data URI prefix", () => {
    const bare = imageFromBase64(PNG, "png", ANCHOR)
    const prefixed = imageFromBase64(`data:image/png;base64,${PNG}`, "png", ANCHOR)

    expect(prefixed.data).toEqual(bare.data)
  })

  it("every byte value, not just the printable ones", () => {
    // 0..255 encoded. A decoder that goes through a string has to keep
    // each charCode as a byte; this is what notices if it does not.
    const all = Uint8Array.from({ length: 256 }, (_, i) => i)
    let binary = ""
    for (const b of all) binary += String.fromCharCode(b)
    const b64 = btoa(binary)

    expect([...imageFromBase64(b64, "png", ANCHOR).data]).toEqual([...all])
  })

  it("an empty string is empty, not an error", () => {
    expect(imageFromBase64("", "png", ANCHOR).data).toHaveLength(0)
  })

  it("unpadded and whitespace-broken input, which atob forgives", () => {
    // `atob` strips ASCII whitespace and tolerates missing padding, so a
    // base64 blob wrapped across lines still decodes.
    expect([...imageFromBase64("YQ", "png", ANCHOR).data]).toEqual([0x61])
    expect([...imageFromBase64("YWJj\ndGVzdA==", "png", ANCHOR).data]).toEqual([
      0x61, 0x62, 0x63, 0x74, 0x65, 0x73, 0x74,
    ])
  })
})

describe("invalid base64 says so, rather than making bytes up", () => {
  it("throws InvalidArgumentError, not a DOMException", () => {
    // The runtime's own error is an `InvalidCharacterError` naming
    // nothing about images or spreadsheets.
    expect(() => imageFromBase64("not!valid!base64", "png", ANCHOR)).toThrow(InvalidArgumentError)
  })

  it("and the message names what to look for", () => {
    expect(() => imageFromBase64("YWJj###", "png", ANCHOR)).toThrow(/not valid base64/)
    expect(() => imageFromBase64("YWJj###", "png", ANCHOR)).toThrow(/URL-safe alphabet/)
  })

  it("including a URL-safe alphabet, which is the likely mistake", () => {
    // base64url swaps `+/` for `-_`. `Buffer.from` accepts it silently
    // and produces different bytes; this refuses it by name.
    expect(() => imageFromBase64("a-b_c-d_", "png", ANCHOR)).toThrow(InvalidArgumentError)
  })

  it("and over-padding, which used to decode to something", () => {
    expect(() => imageFromBase64("YQ===", "png", ANCHOR)).toThrow(InvalidArgumentError)
  })
})
