import { describe, expect, it } from "vitest"

// ── Entry points ─────────────────────────────────────────────────────
// These are the paths `package.json#exports` publishes. Importing the
// barrel itself is the point of this file: no other test does, which is
// exactly how `hucre/xlsx` shipped without the exports its own README
// documented (#361).

import * as root from "../src/index"
import * as xlsx from "../src/xlsx"
import * as csv from "../src/csv"
import * as ods from "../src/ods"
import * as json from "../src/json"
import * as xml from "../src/xml"
import * as ooxml from "../src/ooxml"

// ── Type-level surface ───────────────────────────────────────────────
// Values can be asserted at runtime; types cannot. Importing them here
// means `tsc --noEmit` fails if any disappears from its documented path.

import type {
  Cell,
  CellStyle,
  CellValue,
  ChartAxisCrosses,
  ChartColor,
  ChartErrorBars,
  ChartTrendline,
  ColumnDef,
  ConditionalRule,
  ConditionalRuleType,
  DataValidation,
  MergeRange,
  PaperSize,
  ReadObjectsOptions,
  ReadObjectsResult,
  ReadOptions,
  Sheet,
  SheetObjectsResult,
  SheetTextBox,
  Sparkline,
  TableColumn,
  TableDefinition,
  ValidationOperator,
  ValidationType,
  Workbook,
  WriteOptions,
  WriteSheet,
} from "../src/index"
import type {
  CsvObjectsResult,
  CsvReadOptions,
  CsvStreamWriterOptions,
  CsvWriteOptions,
} from "../src/csv"
import type { OdsObjectsResult } from "../src/ods"
import type { JsonReadResult, NdjsonStreamWriterOptions } from "../src/json"
import type {
  XlsxObjectsResult,
  XlsxStreamSheet,
  XlsxStreamWriterOptions,
  XlsxWriteStreamWorkbookOptions,
} from "../src/xlsx"

/** Referencing each type keeps the imports load-bearing rather than unused. */
type _SurfaceCheck = [
  Cell,
  CellStyle,
  CellValue,
  ChartAxisCrosses,
  ChartColor,
  ChartErrorBars,
  ChartTrendline,
  ColumnDef,
  ConditionalRule,
  ConditionalRuleType,
  // Every `*Objects` reader's result type is nameable, and they are all
  // the same `{ data, headers }` shape (#365 item 6).
  CsvObjectsResult,
  CsvReadOptions,
  CsvStreamWriterOptions,
  CsvWriteOptions,
  DataValidation,
  JsonReadResult,
  MergeRange,
  NdjsonStreamWriterOptions,
  OdsObjectsResult,
  PaperSize,
  ReadObjectsOptions,
  ReadObjectsResult,
  ReadOptions,
  Sheet,
  SheetObjectsResult,
  SheetTextBox,
  Sparkline,
  TableColumn,
  TableDefinition,
  ValidationOperator,
  ValidationType,
  Workbook,
  WriteOptions,
  WriteSheet,
  XlsxObjectsResult,
  XlsxStreamSheet,
  XlsxStreamWriterOptions,
  XlsxWriteStreamWorkbookOptions,
]

/**
 * #365 item 6: every `*Objects` reader returns the *same* `{ data, headers }`
 * type, not merely a similar one. `tsc --noEmit` is the assertion.
 */
type Assert<T extends true> = T
type Mutual<A, B> = [A] extends [B] ? ([B] extends [A] ? true : false) : false

type _ObjectsResultsAgree = [
  Assert<Mutual<ReadObjectsResult, XlsxObjectsResult>>,
  Assert<Mutual<ReadObjectsResult, OdsObjectsResult>>,
  Assert<Mutual<ReadObjectsResult, CsvObjectsResult>>,
  Assert<Mutual<ReadObjectsResult, SheetObjectsResult>>,
  Assert<Mutual<ReadObjectsResult, JsonReadResult>>,
]

// ── Helpers ──────────────────────────────────────────────────────────

function names(mod: Record<string, unknown>): string[] {
  return Object.keys(mod).sort()
}

function expectExports(mod: Record<string, unknown>, expected: string[], label: string): void {
  for (const name of expected) {
    expect(names(mod), `${label} is missing ${name}`).toContain(name)
  }
}

// ═══════════════════════════════════════════════════════════════════════
// Every symbol the README tells people to import from a subpath must
// actually be there. `hucre/xlsx` and `hucre/ods` both failed this.
// ═══════════════════════════════════════════════════════════════════════

describe("hucre/xlsx", () => {
  it("exports the documented read/write API", () => {
    expectExports(
      xlsx,
      ["readXlsx", "writeXlsx", "readXlsxObjects", "writeXlsxObjects", "openXlsx", "saveXlsx"],
      "hucre/xlsx",
    )
  })

  it("exports the other formats it can read", () => {
    expectExports(xlsx, ["readXlsb", "readXls"], "hucre/xlsx")
  })

  it("exports the streaming API", () => {
    expectExports(
      xlsx,
      [
        "streamXlsxRows",
        "XlsxStreamWriter",
        "writeXlsxStream",
        "writeXlsxStreamSheets",
        "XLSX_MAX_ROWS_PER_SHEET",
      ],
      "hucre/xlsx",
    )
  })

  it("exports the sizing, theme, and cell helpers", () => {
    expectExports(
      xlsx,
      [
        "calculateColumnWidth",
        "measureValueWidth",
        "calculateRowHeight",
        "parseCellRef",
        "colToLetter",
        "cellRef",
        "rangeRef",
        "link",
        "hashSheetPassword",
      ],
      "hucre/xlsx",
    )
  })
})

describe("hucre/ods", () => {
  it("exports more than readOds/writeOds", () => {
    // This entry point shipped with exactly two symbols while the README
    // documented five.
    expectExports(
      ods,
      ["readOds", "writeOds", "streamOdsRows", "readOdsObjects", "writeOdsObjects"],
      "hucre/ods",
    )
  })
})

describe("hucre/csv", () => {
  it("exports the documented API", () => {
    expectExports(
      csv,
      [
        "parseCsv",
        "parseCsvObjects",
        "writeCsv",
        "writeCsvObjects",
        "detectDelimiter",
        "stripBom",
        "formatCsvValue",
        "fetchCsv",
        "streamCsvRows",
        "writeCsvStream",
        "CsvStreamWriter",
      ],
      "hucre/csv",
    )
  })
})

describe("hucre/json", () => {
  it("exports the documented API", () => {
    expectExports(
      json,
      [
        "parseJson",
        "parseValue",
        "parseNdjson",
        "writeJson",
        "writeNdjson",
        "workbookToJson",
        "jsonToWorkbook",
        "NdjsonStreamWriter",
        "streamNdjsonRows",
        "flattenValue",
        "collectHeaders",
        "unflattenRow",
        "unflattenRows",
      ],
      "hucre/json",
    )
  })
})

describe("hucre/ooxml", () => {
  it("exports the low-level OOXML part parsers", () => {
    expectExports(
      ooxml,
      [
        "parseChart",
        "cloneChart",
        "chartKindToWriteKind",
        "addChart",
        "getCharts",
        "parsePivotTable",
        "parsePivotCacheDefinition",
        "attachPivotCacheFields",
        "parseSlicers",
        "parseSlicerCache",
        "parseTimelines",
        "parseTimelineCache",
        "parsePersons",
        "parseThreadedComments",
        "parseExternalLink",
        "parseCellImages",
        "assembleCellImages",
        "parseThemeColors",
        "resolveThemeColor",
      ],
      "hucre/ooxml",
    )
  })

  it("keeps the raw-XML parsers off the root", () => {
    // v1 re-exported them from the root, deprecated; v2 removes that.
    // The model-level chart helpers stay at the root — they take a
    // `Chart`, not an XML string.
    const modelLevel = new Set(["cloneChart", "chartKindToWriteKind", "addChart", "getCharts"])
    for (const name of names(ooxml)) {
      if (modelLevel.has(name)) {
        expect(names(root), `root lost ${name}`).toContain(name)
      } else {
        expect(names(root), `root still exports ${name}`).not.toContain(name)
      }
    }
  })
})

describe("hucre/xml", () => {
  it("exports the documented API", () => {
    expectExports(xml, ["readXml", "writeXml"], "hucre/xml")
  })
})

describe("hucre (root)", () => {
  it("exports the high-level API", () => {
    expectExports(root, ["read", "write", "readObjects", "writeObjects"], "hucre")
  })

  it("re-exports every format's entry-point values", () => {
    // A subpath symbol missing from the root would send people to the
    // wrong import; the reverse already happened.
    for (const [label, mod] of [
      ["hucre/xlsx", xlsx],
      ["hucre/csv", csv],
      ["hucre/ods", ods],
      ["hucre/json", json],
      ["hucre/xml", xml],
    ] as const) {
      for (const name of names(mod)) {
        expect(names(root), `root is missing ${name} from ${label}`).toContain(name)
      }
    }
  })

  it("exports the error hierarchy", () => {
    expectExports(
      root,
      [
        "HucreError",
        "ParseError",
        "ZipError",
        "XmlError",
        "ValidationError",
        "UnsupportedFormatError",
        "EncryptedFileError",
      ],
      "hucre",
    )
  })

  it("exports the resource limits", () => {
    // #471 — these were reachable from nowhere, so a caller who wanted to
    // name the bound in their own message had to hard-code 20,000,000 and
    // hope it did not move.
    expectExports(
      root,
      [
        "MAX_COL_INDEX",
        "MAX_DECOMPRESSED_BYTES",
        "MAX_INPUT_BYTES",
        "MAX_REPEAT_COUNT",
        "MAX_ROW_INDEX",
        "MAX_SPAN_CELLS",
        "MAX_SPIN_COUNT",
        "MAX_TOTAL_CELLS",
      ],
      "hucre",
    )

    // Quotable means the value, not just the name.
    expect(root.MAX_TOTAL_CELLS).toBe(20_000_000)
    expect(root.MAX_ROW_INDEX).toBe(1_048_575)
    expect(root.MAX_COL_INDEX).toBe(16_383)
  })
})

// ═══════════════════════════════════════════════════════════════════════
// Snapshot of the whole surface. v1 freezes this — a diff here means a
// deliberate decision, not a drive-by change.
// ═══════════════════════════════════════════════════════════════════════

describe("export surface snapshot", () => {
  it("root", () => {
    expect(names(root)).toMatchSnapshot()
  })

  it("hucre/xlsx", () => {
    expect(names(xlsx)).toMatchSnapshot()
  })

  it("hucre/csv", () => {
    expect(names(csv)).toMatchSnapshot()
  })

  it("hucre/ods", () => {
    expect(names(ods)).toMatchSnapshot()
  })

  it("hucre/json", () => {
    expect(names(json)).toMatchSnapshot()
  })

  it("hucre/xml", () => {
    expect(names(xml)).toMatchSnapshot()
  })

  it("hucre/ooxml", () => {
    expect(names(ooxml)).toMatchSnapshot()
  })
})
