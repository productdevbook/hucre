import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { readXlsx } from "../src/xlsx/reader"
import { ZipReader } from "../src/zip/reader"
import { ZipWriter } from "../src/zip/writer"

// #622: openXlsx -> saveXlsx of an Excel-made protected template dropped the
// password digest, so the saved file was protected with no password (or a
// different one). The digest is a one-way value; it is read and written back
// verbatim, never turned into `password`.

async function part(bytes: Uint8Array, path: string): Promise<string> {
  return new TextDecoder().decode(await new ZipReader(bytes).extract(path))
}

async function rewrite(
  bytes: Uint8Array,
  path: string,
  edit: (xml: string) => string,
): Promise<Uint8Array> {
  const all = await new ZipReader(bytes).extractAll()
  const zw = new ZipWriter()
  for (const [name, data] of all) {
    zw.add(
      name,
      name === path ? new TextEncoder().encode(edit(new TextDecoder().decode(data))) : data,
    )
  }
  return zw.build()
}

const SHEET_ATTRS =
  'algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000"'
const BOOK_ATTRS =
  'workbookAlgorithmName="SHA-512" workbookHashValue="aGFzaA==" workbookSaltValue="c2FsdA==" workbookSpinCount="100000"'

async function template(): Promise<Uint8Array> {
  const base = await writeXlsx({
    workbookProtection: { lockStructure: true },
    sheets: [{ name: "S", rows: [["x"]], protection: { sheet: true } }],
  })
  const a = await rewrite(base, "xl/worksheets/sheet1.xml", (x) =>
    x.replace("<sheetProtection ", `<sheetProtection password="CC3D" ${SHEET_ATTRS} `),
  )
  return rewrite(a, "xl/workbook.xml", (x) =>
    x.replace(
      "<workbookProtection ",
      `<workbookProtection workbookPassword="CC3D" ${BOOK_ATTRS} revisionsPassword="ABCD" revisionsAlgorithmName="SHA-512" revisionsHashValue="cmV2" revisionsSaltValue="c2FsdA==" revisionsSpinCount="5" `,
    ),
  )
}

describe("password digests survive open -> save (#622)", () => {
  it("keeps the sheet digest", async () => {
    const out = await saveXlsx(await openXlsx(await template()))
    const xml = await part(out, "xl/worksheets/sheet1.xml")
    expect(xml).toContain('password="CC3D"')
    expect(xml).toContain('algorithmName="SHA-512"')
    expect(xml).toContain('hashValue="aGFzaA=="')
    expect(xml).toContain('saltValue="c2FsdA=="')
    expect(xml).toContain('spinCount="100000"')
  })

  it("keeps the workbook and revisions digests", async () => {
    const out = await saveXlsx(await openXlsx(await template()))
    const xml = await part(out, "xl/workbook.xml")
    expect(xml).toContain('workbookPassword="CC3D"')
    expect(xml).toContain('workbookAlgorithmName="SHA-512"')
    expect(xml).toContain('workbookHashValue="aGFzaA=="')
    expect(xml).toContain('workbookSaltValue="c2FsdA=="')
    expect(xml).toContain('workbookSpinCount="100000"')
    expect(xml).toContain('revisionsPassword="ABCD"')
    expect(xml).toContain('revisionsHashValue="cmV2"')
    expect(xml).toContain('revisionsSpinCount="5"')
  })

  it("reads the digest into `digest`, never into `password`", async () => {
    const wb = await readXlsx(await template())
    expect(wb.sheets[0]!.protection?.password).toBeUndefined()
    expect(wb.sheets[0]!.protection?.digest?.hashValue).toBe("aGFzaA==")
    expect(wb.workbookProtection?.digest?.spinCount).toBe(100000)
  })

  it("lets an explicit plaintext password replace a carried digest", async () => {
    const wb = await openXlsx(await template())
    wb.sheets[0]!.protection!.password = "new"
    const xml = await part(await saveXlsx(wb), "xl/worksheets/sheet1.xml")
    expect(xml).not.toContain("hashValue")
    expect(xml).not.toContain("CC3D")
    expect(xml).toMatch(/password="[0-9A-F]{4}"/)
  })
})
