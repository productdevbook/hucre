import { ZipWriter } from "../../src/zip/writer"

const enc = new TextEncoder()

export const NS = [
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"`,
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"`,
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`,
  `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`,
  `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"`,
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`,
  `xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"`,
  `xmlns:dc="http://purl.org/dc/elements/1.1/"`,
  `xmlns:xlink="http://www.w3.org/1999/xlink"`,
  `xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"`,
].join(" ")

/** A complete `content.xml` around a `<office:spreadsheet>` body. */
export function contentXml(body: string, automaticStyles = ""): string {
  const styles = automaticStyles
    ? `<office:automatic-styles>${automaticStyles}</office:automatic-styles>`
    : ""
  return (
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content ${NS} office:version="1.3">` +
    styles +
    `<office:body><office:spreadsheet>${body}</office:spreadsheet></office:body>` +
    `</office:document-content>`
  )
}

/** A complete `meta.xml` around the `<office:meta>` children. */
export function metaXml(inner: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta ${NS} office:version="1.3">` +
    `<office:meta>${inner}</office:meta>` +
    `</office:document-meta>`
  )
}

export interface OdsParts {
  /** Full content.xml. Omit to leave the entry out of the archive. */
  content?: string
  meta?: string
  /** Defaults to the spreadsheet media type. */
  mimetype?: string | null
}

export async function odsFile(parts: OdsParts): Promise<Uint8Array> {
  const zip = new ZipWriter()
  if (parts.mimetype !== null) {
    zip.add(
      "mimetype",
      enc.encode(parts.mimetype ?? "application/vnd.oasis.opendocument.spreadsheet"),
      { compress: false },
    )
  }
  if (parts.content !== undefined) zip.add("content.xml", enc.encode(parts.content))
  if (parts.meta !== undefined) zip.add("meta.xml", enc.encode(parts.meta))
  return await zip.build()
}

/** Minimal independent ODF package; the raw spreadsheet XML stays in the test. */
export function odsFromContent(body: string, styles = ""): Promise<Uint8Array> {
  return odsFile({ content: contentXml(body, styles) })
}
