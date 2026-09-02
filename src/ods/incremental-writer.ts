// ── Incremental ODS Writer ──────────────────────────────────────────
//
// The last empty cell in the streaming matrix. `writeOdsStream` covers
// the constant-memory case and carries values only, because ODF puts
// `<office:automatic-styles>` *before* the body: a style first seen on
// row 900,000 has nowhere to be declared once the body has gone out.
//
// A buffering writer does not have that problem. It holds the serialized
// rows until `finish()`, so the style block can be written from
// everything it saw. That is the trade this class exists to make, and it
// is the same one XLSX already offers: `writeXlsxStream` for constant
// memory, `XlsxStreamWriter` for a buffer you can style. See #467.

import { isInlineCell } from "../_inline-cells"
import { isCellError } from "../cell-error"
import type {
  CellValue,
  CellStyle,
  WorkbookProperties,
  CellInput,
  Cell,
  SpreadsheetStreamWriter,
} from "../_types"
import { validateSheetNames } from "../_validate"
import { xmlElement, xmlSelfClose } from "../xml/writer"
import { ZipWriter } from "../zip/writer"
import {
  MIMETYPE,
  cellToOds,
  createStyleCollector,
  getOrCreateStyleName,
  writeManifestXml,
  writeMetaXml,
  writeSettingsXml,
  writeStylesXml,
} from "./writer"
import type { CellContext } from "./writer"

const encoder = /* @__PURE__ */ new TextEncoder()

/** A cell that brings its own formatting, or just a value. */

export interface OdsStreamWriterOptions {
  /** Sheet name. Excel's limits apply — LibreOffice enforces them too. */
  name?: string
  /**
   * Column definitions. `header` writes a first row; `width` and `style`
   * are carried, unlike in `writeOdsStream` where the body has already
   * gone out by the time a style is known.
   */
  columns?: Array<{ header?: string; key?: string; width?: number; style?: CellStyle }>
  /** Document properties written to `meta.xml`. */
  properties?: WorkbookProperties
}

/**
 * Incremental ODS writer.
 *
 * Rows are serialized to XML as they arrive — no workbook object model
 * is built — but every serialized row is retained until {@link finish}
 * assembles the archive, so peak memory scales with the data. For
 * constant-memory output use `writeOdsStream` instead, and accept that
 * it carries values only.
 *
 * ```ts
 * const writer = new OdsStreamWriter({
 *   name: "Report",
 *   columns: [{ header: "Name", width: 20 }, { header: "Qty" }],
 * })
 * writer.addRow(["Widget", { value: 3, style: { font: { bold: true } } }])
 * const bytes = await writer.finish()
 * ```
 *
 * Implements the same `addRow` / `addObject` / `finish` / `toStream`
 * vocabulary as the other incremental writers, so a format-agnostic
 * helper written against `SpreadsheetStreamWriter` takes it unchanged.
 */
export class OdsStreamWriter implements SpreadsheetStreamWriter {
  private sheetName: string
  private columns: OdsStreamWriterOptions["columns"]
  private properties: WorkbookProperties | undefined
  private collector = createStyleCollector()
  private rowFragments: string[] = []
  private maxCols = 0
  private done = false

  constructor(options?: OdsStreamWriterOptions) {
    this.sheetName = options?.name ?? "Sheet1"
    validateSheetNames([{ name: this.sheetName }])
    this.columns = options?.columns
    this.properties = options?.properties

    // A header row is written immediately, the same as XlsxStreamWriter
    // does, so `addRow` starts at the first data row either way.
    const headers = this.columns?.map((c) => c.header)
    if (headers?.some((h) => h !== undefined)) {
      this.addRow(headers.map((h) => h ?? null))
    }
  }

  /** Append a row of positional values, each optionally styled. */
  addRow(values: CellInput[]): void {
    if (this.done) {
      throw new Error("Cannot write to OdsStreamWriter after finish()")
    }
    if (values.length > this.maxCols) this.maxCols = values.length

    const cells: string[] = []
    for (let i = 0; i < values.length; i++) {
      const raw = values[i]
      const styled = isInlineCell(raw) ? raw : undefined
      const value = styled ? (styled.value ?? null) : (raw as CellValue)

      // A cell's own style wins over its column's, which is the same
      // precedence XlsxStreamWriter uses.
      const style = styled?.style ?? this.columns?.[i]?.style
      const ctx: CellContext = {}
      if (style) {
        const name = getOrCreateStyleName(this.collector, style)
        if (name) ctx.styleName = name
      }
      if (styled?.formula !== undefined) {
        ctx.cellOverride = { formula: styled.formula }
      }

      cells.push(cellToOds(value, ctx, this.collector))
    }

    this.rowFragments.push(`<table:table-row>${cells.join("")}</table:table-row>`)
  }

  /**
   * Append a row from an object, projected through the column order.
   *
   * Needs `columns` with `key` accessors, for the same reason
   * `XlsxStreamWriter.addObject` does: an object's values have no
   * position without one.
   */
  addObject(item: Record<string, CellInput>): void {
    if (!this.columns) {
      throw new Error("addObject requires columns with key accessors")
    }
    this.addRow(this.columns.map((c) => (c.key ? (item[c.key] ?? null) : null)))
  }

  /** Finalize and return the ODS document. */
  async finish(): Promise<Uint8Array> {
    this.done = true

    const zip = new ZipWriter()
    // mimetype MUST be first and MUST be stored uncompressed.
    zip.add("mimetype", encoder.encode(MIMETYPE), { compress: false })
    zip.add("META-INF/manifest.xml", encoder.encode(writeManifestXml()))
    zip.add("content.xml", encoder.encode(this.contentXml()))
    zip.add("meta.xml", encoder.encode(writeMetaXml(this.properties)))
    zip.add("styles.xml", encoder.encode(writeStylesXml()))
    zip.add("settings.xml", encoder.encode(writeSettingsXml()))
    return zip.build()
  }

  /**
   * Assemble content.xml.
   *
   * The order is the point: the style block is built *here*, from
   * everything the rows produced, and ODF requires it before the body.
   * Data styles come before the cell styles that reference them.
   */
  private contentXml(): string {
    const columnElements: string[] = []
    const count = Math.max(this.maxCols, this.columns?.length ?? 0)
    for (let i = 0; i < count; i++) {
      const width = this.columns?.[i]?.width
      columnElements.push(
        width === undefined
          ? xmlSelfClose("table:table-column")
          : xmlSelfClose("table:table-column", {
              "table:style-name": this.columnStyleName(i, width),
            }),
      )
    }

    const table = xmlElement("table:table", { "table:name": this.sheetName }, [
      ...columnElements,
      ...this.rowFragments,
    ])
    const body = xmlElement(
      "office:body",
      undefined,
      xmlElement("office:spreadsheet", undefined, table),
    )

    const styleParts = [
      ...this.collector.dataStyleElements.values(),
      ...this.collector.styleElements.values(),
      ...this.collector.textStyleElements.values(),
      ...this.columnStyles.values(),
    ]

    return xmlDocumentContent([
      xmlSelfClose("office:scripts"),
      xmlElement("office:font-face-decls", undefined, ""),
      xmlElement("office:automatic-styles", undefined, styleParts.length > 0 ? styleParts : ""),
      body,
    ])
  }

  private columnStyles = new Map<string, string>()

  /** A `<style:style>` for one column's width, made once per column. */
  private columnStyleName(index: number, width: number): string {
    const name = `co${index + 1}`
    if (!this.columnStyles.has(name)) {
      // ODF column widths are a physical measure; the same
      // 7px-per-character approximation the buffered writer uses.
      const inches = (width * 7 + 5) / 96
      this.columnStyles.set(
        name,
        `<style:style style:name="${name}" style:family="table-column">` +
          `<style:table-column-properties style:column-width="${inches.toFixed(4)}in"/>` +
          "</style:style>",
      )
    }
    return name
  }
}

/** The `office:document-content` shell, with the namespaces ODF wants. */
function xmlDocumentContent(children: string[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
    ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' +
    ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
    ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' +
    ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' +
    ' xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"' +
    ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' +
    ' xmlns:xlink="http://www.w3.org/1999/xlink"' +
    ' xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"' +
    ' xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"' +
    ' office:version="1.3">' +
    children.join("") +
    "</office:document-content>"
  )
}
