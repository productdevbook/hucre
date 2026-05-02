// ── Streaming ODS Reader ─────────────────────────────────────────────
// Yields rows one at a time from an ODS file via SAX parsing.

import type { CellValue } from "../_types";
import { ParseError, ZipError } from "../errors";
import { assertNotEncrypted } from "../_input";
import { ZipReader } from "../zip/reader";
import { parseSax } from "../xml/parser";

// ── Helpers ──────────────────────────────────────────────────────────

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data);
}

function toUint8Array(input: Uint8Array | ArrayBuffer): Uint8Array {
  if (input instanceof Uint8Array) return input;
  return new Uint8Array(input);
}

// ── Row parser via SAX ──────────────────────────────────────────────

interface OdsStreamRow {
  /** 0-based row index */
  index: number;
  /** Cell values for this row */
  values: CellValue[];
}

function* parseContentRows(xml: string): Generator<OdsStreamRow, void, undefined> {
  const completedRows: OdsStreamRow[] = [];

  let inBody = false;
  let inSpreadsheet = false;
  let inTable = false;
  let inRow = false;
  let inCell = false;
  let inP = false;

  let currentRowIndex = -1;
  let cellRepeat = 1;
  let rowRepeat = 1;
  let currentCells: CellValue[] = [];
  let cellText = "";
  let cellValueType = "";
  let cellValue = "";
  let cellBoolValue = "";
  let cellDateValue = "";

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

      switch (local) {
        case "body":
          inBody = true;
          break;
        case "spreadsheet":
          if (inBody) inSpreadsheet = true;
          break;
        case "table":
          if (inSpreadsheet) {
            inTable = true;
            currentRowIndex = -1;
          }
          break;
        case "table-row":
          if (inTable) {
            inRow = true;
            rowRepeat = Number(attrs["table:number-rows-repeated"] ?? "1");
            currentCells = [];
          }
          break;
        case "table-cell":
          if (inRow) {
            inCell = true;
            cellRepeat = Number(attrs["table:number-columns-repeated"] ?? "1");
            cellText = "";
            cellValueType = attrs["office:value-type"] ?? attrs["calcext:value-type"] ?? "";
            cellValue = attrs["office:value"] ?? "";
            cellBoolValue = attrs["office:boolean-value"] ?? "";
            cellDateValue = attrs["office:date-value"] ?? "";
          }
          break;
        case "covered-table-cell":
          if (inRow) {
            const repeat = Number(attrs["table:number-columns-repeated"] ?? "1");
            for (let i = 0; i < repeat; i++) {
              currentCells.push(null);
            }
          }
          break;
        case "p":
          if (inCell) inP = true;
          break;
      }
    },

    onText(text) {
      if (inP && inCell) {
        cellText += text;
      }
    },

    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

      switch (local) {
        case "p":
          inP = false;
          break;
        case "table-cell":
          if (inCell) {
            const value = resolveCellValue(
              cellValueType,
              cellValue,
              cellBoolValue,
              cellDateValue,
              cellText,
            );
            for (let i = 0; i < cellRepeat; i++) {
              currentCells.push(value);
            }
            inCell = false;
          }
          break;
        case "table-row":
          if (inRow) {
            // Trim trailing nulls
            while (currentCells.length > 0 && currentCells[currentCells.length - 1] === null) {
              currentCells.pop();
            }

            if (currentCells.length > 0) {
              // Cap row repeat to avoid memory explosion for empty trailing rows
              const effectiveRepeat = Math.min(rowRepeat, 1);
              for (let r = 0; r < (currentCells.length > 0 ? rowRepeat : effectiveRepeat); r++) {
                currentRowIndex++;
                completedRows.push({
                  index: currentRowIndex,
                  values: r === 0 ? currentCells : [...currentCells],
                });
              }
            } else {
              currentRowIndex += rowRepeat;
            }
            inRow = false;
          }
          break;
        case "table":
          inTable = false;
          break;
        case "spreadsheet":
          inSpreadsheet = false;
          break;
        case "body":
          inBody = false;
          break;
      }
    },
  });

  for (const row of completedRows) {
    yield row;
  }
}

function resolveCellValue(
  valueType: string,
  value: string,
  boolValue: string,
  dateValue: string,
  text: string,
): CellValue {
  switch (valueType) {
    case "float":
    case "currency":
    case "percentage":
      if (value) return Number(value);
      return null;
    case "boolean":
      if (boolValue === "true") return true;
      if (boolValue === "false") return false;
      return null;
    case "date":
      if (dateValue) {
        const d = new Date(dateValue);
        if (!Number.isNaN(d.getTime())) return d;
      }
      return null;
    case "string":
      return text || "";
    default:
      return text || null;
  }
}

// ── Main streaming reader ───────────────────────────────────────────

/**
 * Create an async iterable that yields rows one at a time from an ODS file.
 * Unzips and parses content.xml with SAX, yielding rows as they are parsed.
 */
export async function* streamOdsRows(
  input: Uint8Array | ArrayBuffer,
): AsyncGenerator<OdsStreamRow, void, undefined> {
  const data = toUint8Array(input);

  // Detect password-protected ODF workbooks (OLE2/CFB envelope) up
  // front so streamers fail fast with a typed `EncryptedFileError`
  // instead of a generic ZIP ParseError. Decryption is tracked in #156.
  assertNotEncrypted(data, "ods");

  // 1. Open ZIP archive
  let zip: ZipReader;
  try {
    zip = new ZipReader(data);
  } catch (err) {
    if (err instanceof ZipError) throw err;
    throw new ParseError("Failed to open ODS file: not a valid ZIP archive", undefined, {
      cause: err,
    });
  }

  // 2. Validate mimetype
  if (!zip.has("mimetype")) {
    throw new ParseError("Invalid ODS: missing 'mimetype' entry.");
  }

  // 3. Parse content.xml
  if (!zip.has("content.xml")) {
    throw new ParseError("Invalid ODS: missing content.xml");
  }
  const contentXml = decodeUtf8(await zip.extract("content.xml"));

  // 4. Yield rows via SAX
  yield* parseContentRows(contentXml);
}
