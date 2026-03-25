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
const NS_NUMBER = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const NS_SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const NS_META = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
const NS_DC = "http://purl.org/dc/elements/1.1/";

const MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet";

// ── Helpers ─────────────────────────────────────────────────────────

function formatOdsDate(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function formatOdsDateValue(date: Date): string {
  // ODS date values use ISO 8601 without time zone: YYYY-MM-DDTHH:MM:SS
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

// ── Row serialization with trailing-empty-cell optimization ─────────

function rowToOds(row: CellValue[]): string {
  const cellElements: string[] = [];

  // Find the last non-null cell index to avoid emitting trailing empty cells
  let lastNonNull = row.length - 1;
  while (lastNonNull >= 0 && (row[lastNonNull] === null || row[lastNonNull] === undefined)) {
    lastNonNull--;
  }

  // Emit cells up to and including the last non-null cell,
  // collapsing consecutive null/undefined cells with number-columns-repeated
  let i = 0;
  while (i <= lastNonNull) {
    const cell = row[i];

    if (cell === null || cell === undefined) {
      // Count consecutive empty cells
      let count = 1;
      while (
        i + count <= lastNonNull &&
        (row[i + count] === null || row[i + count] === undefined)
      ) {
        count++;
      }
      if (count > 1) {
        cellElements.push(
          xmlSelfClose("table:table-cell", {
            "table:number-columns-repeated": String(count),
          }),
        );
      } else {
        cellElements.push(xmlSelfClose("table:table-cell"));
      }
      i += count;
    } else {
      cellElements.push(cellToOds(cell));
      i++;
    }
  }

  return xmlElement("table:table-row", undefined, cellElements);
}

// ── content.xml ─────────────────────────────────────────────────────

function writeContentXml(options: WriteOptions): string {
  const { sheets } = options;

  const tableElements: string[] = [];

  for (const sheet of sheets) {
    const children: string[] = [];

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

    // Determine column count (max width across all rows)
    let colCount = 0;
    for (const row of rows) {
      if (row.length > colCount) colCount = row.length;
    }

    // Emit table:table-column element to declare column count
    if (colCount > 0) {
      if (colCount > 1) {
        children.push(
          xmlSelfClose("table:table-column", {
            "table:number-columns-repeated": String(colCount),
          }),
        );
      } else {
        children.push(xmlSelfClose("table:table-column"));
      }
    }

    // Emit rows
    for (const row of rows) {
      children.push(rowToOds(row));
    }

    tableElements.push(xmlElement("table:table", { "table:name": sheet.name }, children));
  }

  const spreadsheetBody = xmlElement("office:spreadsheet", undefined, tableElements);
  const body = xmlElement("office:body", undefined, spreadsheetBody);

  // Build content sections in order per ODS spec:
  // office:scripts, office:font-face-decls, office:automatic-styles, office:body
  const contentParts: string[] = [];
  contentParts.push(xmlSelfClose("office:scripts"));
  contentParts.push(xmlElement("office:font-face-decls", undefined, ""));
  contentParts.push(xmlElement("office:automatic-styles", undefined, ""));
  contentParts.push(body);

  return xmlDocument(
    "office:document-content",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:table": NS_TABLE,
      "xmlns:text": NS_TEXT,
      "xmlns:style": NS_STYLE,
      "xmlns:fo": NS_FO,
      "xmlns:number": NS_NUMBER,
      "xmlns:svg": NS_SVG,
      "office:version": "1.2",
    },
    contentParts,
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
  // ODS spec requires these child elements even if empty:
  // office:font-face-decls, office:styles, office:automatic-styles, office:master-styles
  const children: string[] = [];
  children.push(xmlElement("office:font-face-decls", undefined, ""));
  children.push(xmlElement("office:styles", undefined, ""));
  children.push(xmlElement("office:automatic-styles", undefined, ""));
  children.push(xmlElement("office:master-styles", undefined, ""));

  return xmlDocument(
    "office:document-styles",
    {
      "xmlns:office": NS_OFFICE,
      "xmlns:style": NS_STYLE,
      "xmlns:text": NS_TEXT,
      "xmlns:table": NS_TABLE,
      "xmlns:fo": NS_FO,
      "xmlns:number": NS_NUMBER,
      "xmlns:svg": NS_SVG,
      "office:version": "1.2",
    },
    children,
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

  // styles.xml — style definitions
  zip.add("styles.xml", encoder.encode(writeStylesXml()));

  return zip.build();
}
