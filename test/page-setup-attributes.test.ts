import { describe, expect, it } from "vitest"
import { writeXlsx } from "../src/xlsx/writer"
import { readXlsx } from "../src/xlsx/reader"
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip"
import { ZipReader } from "../src/zip/reader"
import type { PageSetup } from "../src/_types"

// ═══════════════════════════════════════════════════════════════════════
// #470 — `PageSetup` modelled 15 of CT_PageSetup's attributes. Nothing
// claimed otherwise, but print settings were the thinnest part of an
// otherwise complete sheet model, and PARITY.md listed page setup under
// "read + write, at parity".
//
// The register at the bottom is the part that lasts: `Required<PageSetup>`
// means the next attribute added to the type fails `tsc` here until it is
// also written, read, and proven to survive a roundtrip.
// ═══════════════════════════════════════════════════════════════════════

const decoder = new TextDecoder("utf-8")

async function sheetXml(bytes: Uint8Array): Promise<string> {
  return decoder.decode(await new ZipReader(bytes).extract("xl/worksheets/sheet1.xml"))
}

async function write(pageSetup: PageSetup): Promise<Uint8Array> {
  return writeXlsx({ sheets: [{ name: "S", rows: [["a"]], pageSetup }] })
}

/** Write, read back, and return just the page setup. */
async function roundtrip(pageSetup: PageSetup): Promise<PageSetup> {
  const wb = await readXlsx(await write(pageSetup))
  return wb.sheets[0]!.pageSetup ?? {}
}

describe("the three the issue called out first", () => {
  it("firstPageNumber starts numbering somewhere other than 1", async () => {
    const xml = await sheetXml(await write({ firstPageNumber: 7 }))

    expect(xml).toContain('firstPageNumber="7"')
    // The number alone does nothing in Excel — the flag is what turns it
    // on, so a `firstPageNumber` that printed 1 anyway would be a field
    // that looks set and is not.
    expect(xml).toContain('useFirstPageNumber="1"')

    expect(await roundtrip({ firstPageNumber: 7 })).toMatchObject({
      firstPageNumber: 7,
      useFirstPageNumber: true,
    })
  })

  it("lets a caller set the flag off explicitly", async () => {
    // Implied, not forced: a file can carry the number with the flag
    // clear, and that has to survive being opened and saved.
    const xml = await sheetXml(await write({ firstPageNumber: 7, useFirstPageNumber: false }))

    expect(xml).toContain('useFirstPageNumber="0"')
  })

  it("pageOrder picks the direction pages run", async () => {
    expect(await sheetXml(await write({ pageOrder: "overThenDown" }))).toContain(
      'pageOrder="overThenDown"',
    )
    expect((await roundtrip({ pageOrder: "overThenDown" })).pageOrder).toBe("overThenDown")
  })

  it("does not write pageOrder when it is the default", async () => {
    // `downThenOver` is the CT_PageSetup default; emitting it is noise
    // in the diff of every file hucre touches.
    expect(await sheetXml(await write({ pageOrder: "downThenOver" }))).not.toContain("pageOrder")
  })

  it("paperWidth/paperHeight express a size that has no code", async () => {
    const custom: PageSetup = { paperWidth: "210mm", paperHeight: "297mm" }
    const xml = await sheetXml(await write(custom))

    expect(xml).toContain('paperWidth="210mm"')
    expect(xml).toContain('paperHeight="297mm"')
    expect(await roundtrip(custom)).toMatchObject(custom)
  })

  it("keeps paperSize alongside a custom size rather than instead of it", async () => {
    // Excel reads the explicit dimensions in preference when both are
    // present, so dropping the code would lose information a reader of
    // the file might still want.
    const xml = await sheetXml(
      await write({ paperSize: "a4", paperWidth: "1m", paperHeight: "2m" }),
    )

    expect(xml).toContain('paperSize="9"')
    expect(xml).toContain('paperWidth="1m"')
  })
})

describe("the rest of CT_PageSetup", () => {
  it("writes and reads the flags", async () => {
    const flags: PageSetup = { blackAndWhite: true, draft: true }
    const xml = await sheetXml(await write(flags))

    expect(xml).toContain('blackAndWhite="1"')
    expect(xml).toContain('draft="1"')
    expect(await roundtrip(flags)).toMatchObject(flags)
  })

  it("writes and reads the enumerations", async () => {
    const enums: PageSetup = { cellComments: "atEnd", errors: "dash" }
    const xml = await sheetXml(await write(enums))

    expect(xml).toContain('cellComments="atEnd"')
    expect(xml).toContain('errors="dash"')
    expect(await roundtrip(enums)).toMatchObject(enums)
  })

  it("writes and reads the numbers", async () => {
    const nums: PageSetup = { copies: 3, horizontalDpi: 300, verticalDpi: 300 }
    const xml = await sheetXml(await write(nums))

    expect(xml).toContain('copies="3"')
    expect(xml).toContain('horizontalDpi="300"')
    expect(await roundtrip(nums)).toMatchObject(nums)
  })

  it("elides every default, so a bare sheet emits no pageSetup", async () => {
    const xml = await sheetXml(
      await write({ cellComments: "none", errors: "displayed", pageOrder: "downThenOver" }),
    )

    expect(xml).not.toContain("<pageSetup")
  })

  it("refuses a value that is not the integer it claims to be", async () => {
    // `copies="NaN"` reaching the model would serialize back out as the
    // literal string NaN — a file hucre made worse by opening it.
    const bytes = await write({ copies: 3 })
    const patched = new TextEncoder().encode(
      (await sheetXml(bytes)).replace('copies="3"', 'copies="oops"'),
    )

    const { ZipWriter } = await import("../src/zip/writer")
    const zw = new ZipWriter()
    for (const [name, data] of await new ZipReader(bytes).extractAll()) {
      zw.add(name, name === "xl/worksheets/sheet1.xml" ? patched : data)
    }

    const wb = await readXlsx(await zw.build())
    expect(wb.sheets[0]!.pageSetup?.copies).toBeUndefined()
  })
})

describe("openXlsx -> saveXlsx keeps them", () => {
  it("carries every attribute through the roundtrip path", async () => {
    const full: PageSetup = {
      orientation: "landscape",
      firstPageNumber: 12,
      pageOrder: "overThenDown",
      blackAndWhite: true,
      draft: true,
      cellComments: "asDisplayed",
      errors: "NA",
      copies: 2,
      horizontalDpi: 1200,
      verticalDpi: 1200,
      paperWidth: "8.5in",
      paperHeight: "11in",
    }

    const saved = await saveXlsx(await openXlsx(await write(full)))
    const wb = await readXlsx(saved)

    expect(wb.sheets[0]!.pageSetup).toMatchObject(full)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// The register. `Required<PageSetup>` is the point: adding a field to the
// type breaks `tsc` here until it is listed, and listing it without
// wiring the writer or reader breaks the roundtrip assertion below.
// ═══════════════════════════════════════════════════════════════════════

/** Every field of PageSetup, with a value that is not its default. */
const EVERY_FIELD: Required<PageSetup> = {
  paperSize: "a3",
  orientation: "landscape",
  fitToPage: true,
  fitToWidth: 2,
  fitToHeight: 3,
  scale: 80,
  margins: { top: 1, right: 1, bottom: 1, left: 1, header: 0.5, footer: 0.5 },
  printArea: "$A$1:$D$50",
  printTitlesRow: "$1:$1",
  printTitlesColumn: "$A:$A",
  showGridLines: true,
  showRowColHeaders: true,
  horizontalCentered: true,
  verticalCentered: true,
  firstPageNumber: 5,
  useFirstPageNumber: true,
  pageOrder: "overThenDown",
  blackAndWhite: true,
  draft: true,
  cellComments: "atEnd",
  errors: "blank",
  copies: 4,
  horizontalDpi: 300,
  verticalDpi: 300,
  paperWidth: "210mm",
  paperHeight: "297mm",
  usePrinterDefaults: false,
}

describe("register: every PageSetup field survives a roundtrip", () => {
  it("writes and reads back all of them", async () => {
    const wb = await readXlsx(await write(EVERY_FIELD))
    const got = wb.sheets[0]!.pageSetup

    expect(got).toBeDefined()
    for (const [key, want] of Object.entries(EVERY_FIELD)) {
      expect(got![key as keyof PageSetup], `PageSetup.${key}`).toEqual(want)
    }
  })
})
