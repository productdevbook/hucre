import { describe, expect, it } from "vitest"
import { readXlsx } from "../src/xlsx/reader"
import { parseWorksheet, parseWorksheetStream, type WorksheetContext } from "../src/xlsx/worksheet"
import { MAX_STRING_LENGTH } from "../src/_decode"
import { SAX_TEXT_FLUSH_CHARS } from "../src/xml/parser"
import { XmlError } from "../src/errors"
import { ZipReader } from "../src/zip/reader"

// ═══════════════════════════════════════════════════════════════════════
// #503 — a worksheet part over V8's 0x1fffffe8-character ceiling cannot
// become a string, so the buffered reader could not parse it however much
// memory the machine had. #514 and #518 made that legible (a `ParseError`
// naming the part and the bound); neither made the file load. This is the
// capability: the part is parsed from the ZIP entry's stream instead.
//
// Two things need testing, and they are different things.
//
//  1. That a part past the ceiling reads at all. Nothing smaller
//     reproduces it — the trigger is a buffer over 512 MB, which is
//     larger than this repository, so the part is built rather than
//     committed (`test/huge-part-error.test.ts` set the precedent: a
//     large *buffer* is affordable, a large *fixture* is not).
//
//  2. That the streaming driver and the buffered one agree. They share
//     one handler set by construction, so the risk is not a missing
//     field, it is the chunk boundary: a tag, an entity or a text run
//     split across two chunks. That is what the byte-at-a-time case
//     below is for.
// ═══════════════════════════════════════════════════════════════════════

const ctx: WorksheetContext = {
  sharedStrings: [{ text: "shared one" }, { text: "shared two" }],
  styles: null,
  readStyles: false,
  dateSystem: "1900",
}

/**
 * A worksheet exercising the handlers that carry text across calls —
 * values, formulas, inline strings and their runs, data validations,
 * conditional-formatting formulas, header/footer — plus the elements the
 * finalisation reads. Every one of these accumulates with `+=`, which is
 * exactly what a split run depends on.
 */
const RICH_SHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><tabColor rgb="FFFF0000"/></sheetPr>
  <sheetViews><sheetView showGridLines="0" zoomScale="125" tabSelected="1"><pane xSplit="1" ySplit="2" topLeftCell="B3" state="frozen"/></sheetView></sheetViews>
  <cols><col min="1" max="3" width="18.5" customWidth="1"/><col min="4" max="4" hidden="1"/></cols>
  <sheetData>
    <row r="1" ht="24" customHeight="1">
      <c r="A1" t="inlineStr"><is><r><rPr><b/><sz val="12"/></rPr><t>bold run </t></r><r><t>and a plain one &amp; an entity</t></r></is></c>
      <c r="B1" t="s"><v>0</v></c>
      <c r="C1"><f>SUM(A2:A3)</f><v>7</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2" t="str"><f t="shared" si="0" ref="B2:B3">A2*2</f><v>6</v></c>
      <c r="C2" t="b"><v>1</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>4</v></c>
      <c r="B3" t="inlineStr"><is><t xml:space="preserve">  spaced  </t></is></c>
      <c r="C3" t="e"><v>#DIV/0!</v></c>
    </row>
  </sheetData>
  <autoFilter ref="A1:C3"/>
  <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
  <conditionalFormatting sqref="A2:A3"><cfRule type="cellIs" operator="greaterThan" dxfId="0" priority="1"><formula>2</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="list" sqref="C2:C3" allowBlank="1"><formula1>"yes,no"</formula1></dataValidation></dataValidations>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
  <pageSetup orientation="landscape" paperSize="9"/>
  <headerFooter><oddHeader>&amp;LLeft&amp;CCentre</oddHeader><oddFooter>&amp;RPage &amp;P</oddFooter></headerFooter>
  <rowBreaks count="1"><brk id="2" max="16383" man="1"/></rowBreaks>
</worksheet>`

/** The bytes of `xml`, handed out `chunk` at a time. */
function streamOf(xml: string, chunk: number): ReadableStream<Uint8Array> {
  const bytes = new TextEncoder().encode(xml)
  let offset = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= bytes.length) {
        controller.close()
        return
      }
      const end = Math.min(offset + chunk, bytes.length)
      controller.enqueue(bytes.subarray(offset, end))
      offset = end
    },
  })
}

describe("the streaming worksheet driver agrees with the buffered one", () => {
  const buffered = parseWorksheet(RICH_SHEET, "Sheet1", ctx)

  it("produces the same Sheet when the part arrives whole", async () => {
    const streamed = await parseWorksheetStream(streamOf(RICH_SHEET, 1 << 20), "Sheet1", ctx)

    expect(streamed).toEqual(buffered)
  })

  it("produces the same Sheet when every byte is its own chunk", async () => {
    // The worst boundary case there is: every tag, every attribute value,
    // every entity and every text run is split. If `parseSaxStream`'s
    // hold-back were wrong anywhere, this is where it shows — and a
    // chunking bug that only bit on a 589 MB file would otherwise need a
    // 589 MB file to find.
    const streamed = await parseWorksheetStream(streamOf(RICH_SHEET, 1), "Sheet1", ctx)

    expect(streamed).toEqual(buffered)
  })

  it("agrees on a sheet read sparsely, where there is no grid to compare", async () => {
    const sparseCtx = { ...ctx, sparse: true }

    expect(await parseWorksheetStream(streamOf(RICH_SHEET, 7), "S", sparseCtx)).toEqual(
      parseWorksheet(RICH_SHEET, "S", sparseCtx),
    )
  })
})

describe("the two drivers agree about failure, not only about success", () => {
  // Sharing the handlers makes the two drivers agree on what a good
  // document means. It does not, on its own, make them agree on what a
  // bad one means — and until #503 gave the buffered parser a streaming
  // twin, only one of them had an opinion. `processSaxBuffer` returned
  // the unfinished remainder and stopped; `parseSax` threw. A worksheet
  // truncated by a crashed producer therefore parsed to a short `Sheet`
  // with no error on exactly the files that can only be read streaming.
  const TRUNCATED: Array<[string, string]> = [
    ["an opening tag", `<worksheet><sheetData><row r="1"><c r="A1" t="inl`],
    ["a closing tag", `<worksheet><sheetData><row r="1"></row></sheetDa`],
    ["a comment", `<worksheet><sheetData></sheetData><!-- unfinished`],
    ["a CDATA section", `<worksheet><sheetData><row r="1"><c r="A1"><![CDATA[x`],
    ["a processing instruction", `<worksheet><sheetData></sheetData><?mso-application`],
  ]

  for (const [what, xml] of TRUNCATED) {
    it(`rejects a document ending inside ${what}, as the buffered parser does`, async () => {
      expect(() => parseWorksheet(xml, "S", ctx)).toThrow(XmlError)

      await expect(parseWorksheetStream(streamOf(xml, 8), "S", ctx)).rejects.toThrow(XmlError)
    })
  }
})

describe("a text run flushed mid-CRLF", () => {
  // `parseSaxStream` flushes a text run that has outgrown
  // SAX_TEXT_FLUSH_CHARS rather than carrying it across another chunk,
  // and `normalizeEol` runs on each piece it emits. The cut lands at the
  // chunk boundary, so a boundary falling between a CR and its LF split
  // one line ending into two: the first piece ended "\r" and normalised
  // to "\n", the second began "\n", and the handler's `+=` accumulated
  // both. A cell gained a blank line it never had.
  //
  // Reaching it needs a run over 256 KiB *and* a boundary on the CRLF,
  // which is why the byte-at-a-time case above cannot find it — that one
  // splits everything, but never lets a run grow long enough to flush.
  const PREFIX = `<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>`

  for (const run of [SAX_TEXT_FLUSH_CHARS + 1, SAX_TEXT_FLUSH_CHARS * 2]) {
    it(`keeps one newline when the break falls on the CRLF of a ${run}-char run`, async () => {
      const value = "a".repeat(run) + "\r\n" + "b".repeat(1024)
      const xml = `${PREFIX}${value}</t></is></c></row></sheetData></worksheet>`
      // All ASCII, so a character index is a byte index: end the first
      // chunk immediately after the CR.
      const chunk = PREFIX.length + run + 1

      const streamed = await parseWorksheetStream(streamOf(xml, chunk), "S", ctx)

      expect(streamed.rows[0]![0]).toBe(parseWorksheet(xml, "S", ctx).rows[0]![0])
      expect(String(streamed.rows[0]![0])).not.toContain("\n\n")
    })
  }
})

// ── Building a workbook whose worksheet part is over the ceiling ──────

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Huge" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const HUGE_HEAD = `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c><c r="B1"><v>1</v></c></row>`
const HUGE_TAIL = `<row r="2"><c r="A2" t="inlineStr"><is><t>last</t></is></c><c r="B2"><v>2</v></c></row></sheetData></worksheet>`

/** An entry whose bytes are written straight into the archive buffer. */
interface Entry {
  name: string
  size: number
  fill: (out: Uint8Array) => void
  /**
   * What to write as the entry's uncompressed size, when that should not
   * be the truth. `0` models a central directory that declares nothing,
   * which is what `readWorksheet`'s fallback branches on — not a full
   * data-descriptor archive: bit 3 of the general-purpose flag stays
   * clear and no descriptor is appended, so the `hasDataDescriptor`
   * recovery paths in `extractEntry`/`extractEntryStream` are not what
   * this exercises.
   */
  declares?: number
}

function textEntry(name: string, text: string): Entry {
  const bytes = new TextEncoder().encode(text)
  return { name, size: bytes.length, fill: (out) => out.set(bytes) }
}

/**
 * A STORE-only ZIP, assembled in one allocation.
 *
 * Written here rather than with `ZipWriter` for two reasons the test
 * depends on: the entry is not compressed, so the archive can be built
 * without spending a minute deflating half a gigabyte, and each entry is
 * filled in place, so the half-gigabyte part exists once rather than
 * twice. CRC-32 is left at zero — the readers skip verification for a
 * zero CRC, and computing one over 537 MB would be the slowest thing in
 * this file by far.
 */
function storeZip(entries: Entry[]): { bytes: Uint8Array; declare: (size: number) => void } {
  const nameBytes = entries.map((e) => new TextEncoder().encode(e.name))
  const localSize = entries.reduce((n, e, i) => n + 30 + nameBytes[i]!.length + e.size, 0)
  const centralSize = entries.reduce((n, _e, i) => n + 46 + nameBytes[i]!.length, 0)
  const out = new Uint8Array(localSize + centralSize + 22)
  const view = new DataView(out.buffer)

  const offsets: number[] = []
  const sizeFieldOffsets: number[] = []
  const centralSizeOffsets: number[] = []
  let pos = 0
  for (const [i, entry] of entries.entries()) {
    const name = nameBytes[i]!
    offsets.push(pos)
    view.setUint32(pos, 0x04034b50, true)
    view.setUint16(pos + 4, 20, true)
    view.setUint32(pos + 14, 0, true) // CRC-32 — see above
    view.setUint32(pos + 18, entry.size, true)
    sizeFieldOffsets.push(pos + 22)
    view.setUint32(pos + 22, entry.declares ?? entry.size, true)
    view.setUint16(pos + 26, name.length, true)
    out.set(name, pos + 30)
    entry.fill(out.subarray(pos + 30 + name.length, pos + 30 + name.length + entry.size))
    pos += 30 + name.length + entry.size
  }

  const centralStart = pos
  for (const [i, entry] of entries.entries()) {
    const name = nameBytes[i]!
    view.setUint32(pos, 0x02014b50, true)
    view.setUint16(pos + 4, 20, true)
    view.setUint16(pos + 6, 20, true)
    view.setUint32(pos + 16, 0, true) // CRC-32
    view.setUint32(pos + 20, entry.size, true)
    centralSizeOffsets.push(pos + 24)
    view.setUint32(pos + 24, entry.declares ?? entry.size, true)
    view.setUint16(pos + 28, name.length, true)
    view.setUint32(pos + 42, offsets[i]!, true)
    out.set(name, pos + 46)
    pos += 46 + name.length
  }

  const lastLocal = sizeFieldOffsets[sizeFieldOffsets.length - 1]!
  const lastCentral = centralSizeOffsets[centralSizeOffsets.length - 1]!

  view.setUint32(pos, 0x06054b50, true)
  view.setUint16(pos + 8, entries.length, true)
  view.setUint16(pos + 10, entries.length, true)
  view.setUint32(pos + 12, pos - centralStart, true)
  view.setUint32(pos + 16, centralStart, true)

  return {
    bytes: out,
    // Rewrite the final entry's declared uncompressed size in place. The
    // half-gigabyte body is the expensive part of this fixture and it is
    // identical either way, so the two cases share one archive rather
    // than building it twice: two live 537 MB buffers in a suite that
    // runs test files in parallel starves the time budgets of unrelated
    // tests, which is a flaky suite for no gain.
    declare(size: number) {
      view.setUint32(lastLocal, size, true)
      view.setUint32(lastCentral, size, true)
    },
  }
}

/**
 * A workbook of two rows with half a gigabyte of ignorable whitespace
 * between them, so the worksheet part is past the ceiling and nothing
 * else about it is expensive.
 *
 * Built once and shared: see `declare` on the return of `storeZip`.
 */
function hugeWorkbook(): { bytes: Uint8Array; declare: (size: number) => void; size: number } {
  const padding = MAX_STRING_LENGTH + 1024 - HUGE_HEAD.length - HUGE_TAIL.length
  const head = new TextEncoder().encode(HUGE_HEAD)
  const tail = new TextEncoder().encode(HUGE_TAIL)
  const size = head.length + padding + tail.length

  const zip = storeZip([
    textEntry("[Content_Types].xml", CONTENT_TYPES),
    textEntry("_rels/.rels", ROOT_RELS),
    textEntry("xl/workbook.xml", WORKBOOK),
    textEntry("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
    {
      name: "xl/worksheets/sheet1.xml",
      size,
      fill: (out) => {
        out.fill(0x20)
        out.set(head, 0)
        out.set(tail, out.length - tail.length)
      },
    },
  ])

  return { ...zip, size }
}

describe("a workbook whose worksheet part is past the string ceiling", () => {
  // Before this change both cases threw the #514 `ParseError` naming the
  // part and the bound, and the eleven files in #503 stayed unreadable.
  const expected = [
    ["first", 1],
    ["last", 2],
  ]
  const huge = hugeWorkbook()

  it("reads, instead of reporting that it cannot be read", async () => {
    // The ZIP declares the size, so the streaming path is chosen before
    // anything is decompressed.
    huge.declare(huge.size)

    const wb = await readXlsx(huge.bytes)

    expect(wb.sheets).toHaveLength(1)
    expect(wb.sheets[0]!.name).toBe("Huge")
    expect(wb.sheets[0]!.rows).toEqual(expected)
  }, 120_000)

  it("reads when the ZIP declares no size and the ceiling is met head-on", async () => {
    // With no size in the central directory there is nothing to decide
    // on up front: the buffered read runs, fails at the ceiling, and that
    // error is the signal to retry as a stream. The slow path, and the
    // one that has to work for a file whose ZIP is less forthcoming than
    // Excel's.
    huge.declare(0)

    const wb = await readXlsx(huge.bytes)

    expect(wb.sheets[0]!.rows).toEqual(expected)
  }, 120_000)
})

describe("a declared size the compressed body could not have produced", () => {
  // The declared size decides whether a part is read whole — and so
  // CRC-32 checked — or streamed, which has no whole entry to check.
  // Nothing verifies the field, so a small entry claiming to be enormous
  // would take the streaming route and skip the checksum: corruption
  // buying its way out of verification by lying about its size. A STORE
  // entry expands not at all, so the claim is refused and the buffered,
  // verified path runs.
  const tiny = `<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`

  const { bytes: zip } = storeZip([
    textEntry("[Content_Types].xml", CONTENT_TYPES),
    textEntry("_rels/.rels", ROOT_RELS),
    textEntry("xl/workbook.xml", WORKBOOK),
    textEntry("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
    { ...textEntry("xl/worksheets/sheet1.xml", tiny), declares: MAX_STRING_LENGTH + 1 },
  ])

  it("is not reported as a declared size", () => {
    expect(new ZipReader(zip).declaredSize("xl/worksheets/sheet1.xml")).toBeUndefined()
  })

  it("still reads, by the buffered path it was trying to escape", async () => {
    expect((await readXlsx(zip)).sheets[0]!.rows).toEqual([[1]])
  })
})
