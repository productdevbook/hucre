import { describe, expect, it } from "vitest"
import { parseCsv } from "../src/csv/reader"
import { streamCsvRows } from "../src/csv/stream"

describe("parseCsv: trailing quoted-empty row", () => {
  it('preserves a final row that is a single quoted-empty field ("")', () => {
    expect(parseCsv('a,b\n""')).toEqual([["a", "b"], [""]])
  })

  it("still drops a bare trailing newline", () => {
    expect(parseCsv("a,b\n")).toEqual([["a", "b"]])
  })
})

describe("parseCsv: comments only apply to unquoted leading #", () => {
  it("drops a physically-unquoted leading-# row", () => {
    expect(parseCsv("#real comment\nx", { comment: "#" })).toEqual([["x"]])
  })

  it("preserves a quoted field that starts with #", () => {
    expect(parseCsv('"#not a comment",x', { comment: "#" })).toEqual([["#not a comment", "x"]])
  })
})

describe("streamCsvRows: option parity", () => {
  it("honors maxRows", () => {
    expect([...streamCsvRows("a\nb\nc\nd", { maxRows: 2 })]).toEqual([["a"], ["b"]])
  })

  it("preserves the trailing quoted-empty row", () => {
    expect([...streamCsvRows('a,b\n""')]).toEqual([["a", "b"], [""]])
  })

  it("does not treat a quoted leading-# field as a comment", () => {
    expect([...streamCsvRows('"#not a comment",x', { comment: "#" })]).toEqual([
      ["#not a comment", "x"],
    ])
  })

  it("preserves leading zeros under type inference by default", () => {
    expect([...streamCsvRows("0123", { typeInference: true })]).toEqual([["0123"]])
  })

  it("honors skipLines", () => {
    expect([...streamCsvRows("skip\na\nb", { skipLines: 1 })]).toEqual([["a"], ["b"]])
  })
})
