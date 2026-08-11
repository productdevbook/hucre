import { describe, expect, it } from "vitest"
import { fieldsOf } from "./_reflect"
import { readFileSync } from "node:fs"
import type { Sheet, WriteOptions, WriteSheet, Workbook } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #412 — docs/PARITY.md is the v1 statement of what hucre reads versus
// what it writes. A parity document that drifts is worse than none: it
// converts "you have to check" into "you were told wrong".
//
// So the read-only and write-only lists are derived from the types here,
// not transcribed. Add a field to `Sheet`/`Workbook` with no `Write*`
// counterpart, or the reverse, and this fails until PARITY.md names it.
//
// The per-field XLSX guarantees are held by test/xlsx-write-read-parity
// .test.ts; this file guards the prose around them.
// ═══════════════════════════════════════════════════════════════════════

const parity = (): string => readFileSync(new URL("../docs/PARITY.md", import.meta.url), "utf-8")
const readme = (): string => readFileSync(new URL("../README.md", import.meta.url), "utf-8")

const writeFields = new Set([...fieldsOf("WriteOptions"), ...fieldsOf("WriteSheet")])
const readFields = new Set([...fieldsOf("Workbook"), ...fieldsOf("Sheet")])

const readOnly = [...readFields].filter((f) => !writeFields.has(f)).sort()
const writeOnly = [...writeFields].filter((f) => !readFields.has(f)).sort()

describe("every read-only field is named in PARITY.md", () => {
  it("finds some, so the derivation is actually working", () => {
    // A green suite because the extraction silently returned nothing is
    // the failure mode this guards against.
    expect(readOnly.length).toBeGreaterThan(5)
  })

  for (const field of readOnly) {
    it(`names ${field}`, () => {
      expect(parity()).toContain(field)
    })
  }
})

describe("every write-only field is named in PARITY.md", () => {
  it("finds some", () => {
    expect(writeOnly.length).toBeGreaterThan(2)
  })

  for (const field of writeOnly) {
    it(`names ${field}`, () => {
      expect(parity()).toContain(field)
    })
  }
})

describe("the counts it quotes are the real ones", () => {
  it("states the chart-kind split correctly", () => {
    const source = readFileSync(new URL("../src/xlsx/chart/types.ts", import.meta.url), "utf-8")
    const read = /export type ChartKind =([\s\S]*?)\n\n/.exec(source)![1]
    const write = /export type WriteChartKind =([^\n]*)/.exec(source)![1]

    const readCount = [...read.matchAll(/"/g)].length / 2
    const writeCount = [...write.matchAll(/"/g)].length / 2

    const text = parity()
    expect(text).toContain(`${readCount} chart kinds are read`)
    expect(text).toContain(`**${writeCount} can be authored**`)
  })
})

describe("the two XLSX write paths are stated, not implied", () => {
  it("says which one preserves unmodelled parts", () => {
    const text = parity()
    expect(text).toMatch(/openXlsx.*saveXlsx/)
    expect(text).toContain("byte-for-byte")
    // The specific consequence people get bitten by.
    expect(text).toMatch(/drops macros silently/i)
  })

  it("is reachable from the README", () => {
    expect(readme()).toContain("docs/PARITY.md")
  })
})

describe("the formats with no writer at all", () => {
  it("are stated as read-only", () => {
    const text = parity()
    expect(text).toContain("XLS and XLSB — read only")
    expect(text).toContain("Markdown — write only")
  })
})

// Referenced so the type imports are load-bearing: if `Sheet`,
// `Workbook`, `WriteSheet` or `WriteOptions` is renamed, this file stops
// compiling rather than quietly deriving empty lists from a regex that
// no longer matches anything.
export type _Pinned = [Sheet, Workbook, WriteSheet, WriteOptions]
