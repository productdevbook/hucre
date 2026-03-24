// ── ODS Writer ──────────────────────────────────────────────────────
// Generates valid OpenDocument Spreadsheet (.ods) files.

import type { WriteOptions, WriteOutput, CellValue, WorkbookProperties } from "../_types";
import { ZipWriter } from "../zip/writer";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";

const encoder = /* @__PURE__ */ new TextEncoder();

// ── ODS Namespaces ──────────────────────────────────────────────────

const NS_OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const NS_TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const NS_TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const NS_STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NS_FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const NS_META = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const NS_DC = "http://purl.org/dc/elements/1.1/";

const MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet";

// ── Helpers ─────────────────────────────────────────────────────────

function formatOdsDate(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function formatOdsDateValue(date: Date): string {
  // ODS date values use ISO 8601 without time zone: YYYY-MM-DD
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `${y}-${m}-${d}T${hh}:${mm}:${ss}`;
}

function cellToOds(value: CellValue): string {
  if (value === null || value === undefined) {
    return xmlSelfClose("table:table-cell");
  }

  if (typeof value === "string") {
    return xmlElement(
      "table:table-cell",
      { "office:value-type": "string" },
      xmlElement("text:p", undefined, xmlEscape(value)),
    );
  }

  if (typeof value === "number") {
    return xmlElement(
      "table:table-cell",
      { "office:value-type": "float", "office:value": String(value) },
      xmlElement("text:p", undefined, String(value)),
    );
  }

  if (typeof value === "boolean") {
    return xmlElement(
      "table:table-cell",
      {
        "office:value-type": "boolean",
        "office:boolean-value": value ? "true" : "false",
      },
      xmlElement("text:p", undefined, value ? "TRUE" : "FALSE"),
    );
  }

  if (value instanceof Date) {
    const dateStr = formatOdsDateValue(value);
    return xmlElement(
      "table:table-cell",
      {
        "office:value-type": "date",
        "office:date-value": dateStr,
      },
      xmlElement("text:p", undefined, dateStr),
    );
  }

  return xmlSelfClose("table:table-cell");
}

// ── content.xml ─────────────────────────────────────────────────────

function writeContentXml(options: WriteOptions): string {
  const { sheets } = options;

  const tableElements: string[] = [];

  for (const sheet of sheets) {
    const rowElements: string[] = [];

    // Resolve rows from rows or data
    let rows: CellValue[][] = [];
    if (sheet.rows) {
      rows = sheet.rows;
    } else if (sheet.data && sheet.columns) {
      // Generate header row + data rows from objects
      const keys = sheet.columns.map((c) => c.key ?? c.header ?? "");
      const hasHeaders = sheet.columns.some((c) => c.header);

      if (hasHeaders) {
        const headerRow = sheet.columns.map((c) => c.header ?? c.key ?? "");
        rows.push(headerRow);
      }

      for (const item of sheet.data) {
        const row = keys.map((k) => (k in item ? item[k] : null));
        rows.push(row);
      }
    }

    for (const row of rows) {
      const cellElements: string[] = [];
      for (const cell of row) {
        cellElements.push(cellToOds(cell));
      }
      rowElements.push(xmlElement("table:table-row", undefined, cellElements));
    }

    tableElements.push(xmlElement("table:table", { "table:name": sheet.name }, rowElements));
  }

  const spreadsheetBody = xmlElement("office:spreadsheet", undefined, tableElements);
  const body = xmlElement("office:body", undefined, spreadsheetBody);

  return xmlDocument(
    "office:document-content",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:table": NS_TABLE,
      "xmlns:text": NS_TEXT,
      "xmlns:style": NS_STYLE,
      "xmlns:fo": NS_FO,
      "office:version": "1.2",
    },
    body,
  );
}

// ── meta.xml ────────────────────────────────────────────────────────

function writeMetaXml(props?: WorkbookProperties): string {
  const children: string[] = [];

  if (props?.title) {
    children.push(xmlElement("dc:title", undefined, xmlEscape(props.title)));
  }
  if (props?.subject) {
    children.push(xmlElement("dc:subject", undefined, xmlEscape(props.subject)));
  }
  if (props?.creator) {
    children.push(xmlElement("meta:initial-creator", undefined, xmlEscape(props.creator)));
  }
  if (props?.description) {
    children.push(xmlElement("dc:description", undefined, xmlEscape(props.description)));
  }
  if (props?.keywords) {
    children.push(xmlElement("meta:keyword", undefined, xmlEscape(props.keywords)));
  }
  if (props?.created) {
    children.push(xmlElement("meta:creation-date", undefined, formatOdsDate(props.created)));
  }

  const modified = props?.modified ?? new Date();
  children.push(xmlElement("dc:date", undefined, formatOdsDate(modified)));

  children.push(xmlElement("meta:generator", undefined, "defter"));

  const metaContent = xmlElement("office:meta", undefined, children);

  return xmlDocument(
    "office:document-meta",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:meta": NS_META,
      "xmlns:dc": NS_DC,
      "office:version": "1.2",
    },
    metaContent,
  );
}

// ── styles.xml ──────────────────────────────────────────────────────

function writeStylesXml(): string {
  return xmlDocument(
    "office:document-styles",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:style": NS_STYLE,
      "xmlns:fo": NS_FO,
      "office:version": "1.2",
    },
    "",
  );
}

// ── manifest.xml ────────────────────────────────────────────────────

function writeManifestXml(): string {
  const NS_MANIFEST = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

  const entries: string[] = [];
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "/",
      "manifest:version": "1.2",
      "manifest:media-type": MIMETYPE,
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "content.xml",
      "manifest:media-type": "text/xml",
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "meta.xml",
      "manifest:media-type": "text/xml",
    }),
  );
  entries.push(
    xmlSelfClose("manifest:file-entry", {
      "manifest:full-path": "styles.xml",
      "manifest:media-type": "text/xml",
    }),
  );

  return xmlDocument(
    "manifest:manifest",
    {
      "xmlns:manifest": NS_MANIFEST,
      "manifest:version": "1.2",
    },
    entries,
  );
}

// ── Main Writer ─────────────────────────────────────────────────────

/**
 * Write a workbook to ODS format.
 * Returns a Uint8Array containing the ZIP archive.
 */
export async function writeOds(options: WriteOptions): Promise<WriteOutput> {
  const zip = new ZipWriter();

  // mimetype MUST be the first entry and MUST be stored uncompressed
  zip.add("mimetype", encoder.encode(MIMETYPE), { compress: false });

  // META-INF/manifest.xml
  zip.add("META-INF/manifest.xml", encoder.encode(writeManifestXml()));

  // content.xml — main spreadsheet data
  zip.add("content.xml", encoder.encode(writeContentXml(options)));

  // meta.xml — document metadata
  zip.add("meta.xml", encoder.encode(writeMetaXml(options.properties)));

  // styles.xml — minimal style definitions
  zip.add("styles.xml", encoder.encode(writeStylesXml()));

  return zip.build();
}
