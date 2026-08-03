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
  ReadOptions,
  Sheet,
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
import type { CsvReadOptions, CsvWriteOptions } from "../src/csv"
import type { OdsStreamRow } from "../src/ods"
import type { XlsxObjectsResult } from "../src/xlsx"

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
  CsvReadOptions,
  CsvWriteOptions,
  DataValidation,
  MergeRange,
  OdsStreamRow,
  PaperSize,
  ReadOptions,
  Sheet,
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
      ["streamXlsxRows", "XlsxStreamWriter", "writeXlsxStream", "XLSX_MAX_ROWS_PER_SHEET"],
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
        "parseThemeColors",
        "resolveThemeColor",
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
        "NdjsonStreamWriter",
        "streamNdjsonRows",
        "readNdjsonStream",
        "flattenValue",
        "collectHeaders",
      ],
      "hucre/json",
    )
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
        "DefterError",
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
})
