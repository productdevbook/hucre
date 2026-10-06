import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import { readXls } from "../src/xls/reader"
import { readXlsb } from "../src/xlsx/xlsb/reader"
import { readOds } from "../src/ods/reader"
import { readXlsx } from "../src/xlsx/reader"
import { fieldsOf } from "./_reflect"

// ═══════════════════════════════════════════════════════════════════════
// v1 had one `ReadOptions` for four readers and a table in its doc
// comment saying which reader ignored what. `readXls(bytes, { password })`
// compiled and did nothing. v2 gives each reader its own options type,
// and this file is what keeps the type equal to the behaviour: every
// field a reader's type declares must be read by that reader's source,
// so a field can no longer be accepted and dropped.
// ═══════════════════════════════════════════════════════════════════════

const source = (path: string): string => readFileSync(new URL(path, import.meta.url), "utf-8")

const READERS: Array<{ iface: string; files: string[] }> = [
  { iface: "XlsxReadOptions", files: ["../src/xlsx/reader.ts", "../src/xlsx/worksheet.ts"] },
  { iface: "OdsReadOptions", files: ["../src/ods/reader.ts"] },
  { iface: "XlsbReadOptions", files: ["../src/xlsx/xlsb/reader.ts"] },
  { iface: "XlsReadOptions", files: ["../src/xls/reader.ts"] },
]

describe("each reader's options type names only what that reader reads", () => {
  for (const { iface, files } of READERS) {
    it(`${iface} — every field is read by its reader`, () => {
      const fields = fieldsOf(iface)
      expect(fields.length).toBeGreaterThan(2)
      const text = files.map(source).join("\n")
      for (const field of fields) {
        // `options?.field` or `options.field` — the reader must look at it.
        const used = new RegExp(`options\\??\\.${field}\\b`).test(text)
        expect(used, `${iface}.${field} is declared but never read`).toBe(true)
      }
    })
  }
})

describe("an option the reader does not honour is a type error", () => {
  const bytes = new Uint8Array(0)
  it("readXls takes no password, sheets, readStyles, maxRows or range", () => {
    // @ts-expect-error — .xls is not encrypted with the Agile scheme
    void (() => readXls(bytes, { password: "x" }))
    // @ts-expect-error — the BIFF reader walks every sheet
    void (() => readXls(bytes, { sheets: [0] }))
    // @ts-expect-error — the BIFF reader surfaces no styles
    void (() => readXls(bytes, { readStyles: true }))
    // @ts-expect-error
    void (() => readXls(bytes, { maxRows: 1 }))
    // @ts-expect-error
    void (() => readXls(bytes, { range: "A1:B2" }))
    expect(true).toBe(true)
  })
  it("readXlsb takes no sheets, readStyles, maxRows, range or sparse", () => {
    // @ts-expect-error
    void (() => readXlsb(bytes, { sheets: [0] }))
    // @ts-expect-error
    void (() => readXlsb(bytes, { readStyles: true }))
    // @ts-expect-error
    void (() => readXlsb(bytes, { maxRows: 1 }))
    // @ts-expect-error
    void (() => readXlsb(bytes, { range: "A1:B2" }))
    // @ts-expect-error
    void (() => readXlsb(bytes, { sparse: true }))
    expect(true).toBe(true)
  })
  it("readOds takes no dateSystem, password or sparse", () => {
    // @ts-expect-error — ODS stores ISO dates; there is no 1900/1904 system
    void (() => readOds(bytes, { dateSystem: "1904" }))
    // @ts-expect-error — ODS encryption is not implemented (#156)
    void (() => readOds(bytes, { password: "x" }))
    // @ts-expect-error
    void (() => readOds(bytes, { sparse: true }))
    expect(true).toBe(true)
  })
  it("readXlsx takes every option", () => {
    void (() =>
      readXlsx(bytes, {
        sheets: [0],
        dateSystem: "auto",
        readStyles: true,
        password: "x",
        maxRows: 1,
        range: "A1:B2",
        sparse: false,
        maxInputBytes: 1,
        maxTotalCells: 1,
        maxDecompressedBytes: 1,
        maxSpinCount: 1,
        onWarning: () => {},
      }))
    expect(true).toBe(true)
  })
})
