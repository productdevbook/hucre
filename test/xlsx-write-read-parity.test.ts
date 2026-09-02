// ── Write → read parity over the whole WriteSheet / WriteOptions surface ──
//
// The invariant: anything `writeXlsx` accepts, `readXlsx` gives back.
// It was broken in eight places at once (#407) — each silently, with no
// error and no warning, just `undefined` where a value had been.
//
// Rather than eight regression tests, this file registers every field of
// `WriteSheet` and `WriteOptions` exactly once. The registers are typed as
// mapped types over `keyof Required<...>`, so adding a field to either
// interface fails `tsc` until it is registered here — as a probe that
// round-trips, or as a deliberate one-way entry with the reason it is one.
// That, not the eight fixes, is what stops the ninth.
//
// Each probe gets its own sheet, named after the field it covers, so no
// two fields can interact and a failure names the field directly.

import { describe, expect, it } from "vitest"
import { readXlsx, writeXlsx } from "../src/xlsx"
import { ZipReader } from "../src/zip"
import type { Sheet, WriteOptions, WriteSheet, Workbook } from "../src/_types"

// ── Register shapes ──────────────────────────────────────────────────

/** A field that must survive write → read. */
interface Probe<T> {
  /** The value written into the fixture. */
  value: NonNullable<T>
  /** Pull the same information back out of the parsed workbook. */
  read: (sheet: Sheet, workbook: Workbook) => unknown
  /**
   * What `read` must return. Defaults to `value`; set it where the read
   * model spells the same information differently (a `data[]` sheet comes
   * back as `rows[][]`, a `SheetChart` as a `Chart`, and so on).
   */
  expected?: unknown
  /** Extra fields the probe needs to be meaningful (e.g. `data` needs `columns`). */
  with?: Partial<WriteSheet>
}

/**
 * A field that deliberately does not come back, and why. Every entry here
 * is a decision someone made, not a gap someone tolerated — if a reason
 * reads like "not implemented yet", it belongs in an issue, not here.
 */
interface OneWay {
  oneWay: string
}

type Entry<T> = Probe<T> | OneWay

function isOneWay<T>(entry: Entry<T>): entry is OneWay {
  return "oneWay" in entry
}

const BASE_ROWS = [
  ["Region", "Amount"],
  ["North", 10],
  ["South", 20],
]

const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
])

// ── WriteSheet register ──────────────────────────────────────────────

const SHEET_FIELDS: { [K in keyof Required<WriteSheet>]: Entry<WriteSheet[K]> } = {
  // Every probe's sheet is named after its own key, so locating the sheet
  // by name is itself the check that `name` survived.
  name: { value: "name", read: (sheet) => sheet.name },

  columns: {
    value: [
      { header: "Region", width: 18, hidden: false, outlineLevel: 1 },
      { header: "Amount", width: 12, numFmt: "0.000", style: { font: { bold: true } } },
    ],
    // This sheet supplies its rows as `rows[][]`, so it also pins the
    // other half of #407: `style` and `numFmt` used to apply on the
    // `data[]` path only, silently doing nothing here.
    read: (sheet) => ({
      cols: sheet.columns?.map((c) => ({ width: c.width, outlineLevel: c.outlineLevel })),
      bodyStyle: sheet.cells?.get("1,1")?.style,
    }),
    // `header` is data, not column metadata: it is written into row 0 on
    // the `data[]` path and read back as a cell there. `<cols>` carries
    // width and outline level, and that is what comes back.
    expected: {
      cols: [
        { width: 18, outlineLevel: 1 },
        { width: 12, outlineLevel: undefined },
      ],
      bodyStyle: { numFmt: "0.000", font: { bold: true } },
    },
  },

  rows: { value: BASE_ROWS, read: (sheet) => sheet.rows },

  data: {
    value: [
      { region: "North", amount: 10 },
      { region: "South", amount: 20 },
    ],
    with: {
      columns: [
        { header: "Region", key: "region" },
        { header: "Amount", key: "amount" },
      ],
    },
    // Object rows are a write-side convenience; the file stores a grid, so
    // the reader hands back the grid the objects were flattened into.
    read: (sheet) => sheet.rows,
    expected: BASE_ROWS,
  },

  cells: {
    value: new Map([
      ["1,0", { value: "North", comment: { text: "a note", author: "hucre" } }],
      ["1,1", { formula: "SUM(B2:B3)", formulaDynamic: true }],
    ]),
    read: (sheet) => ({
      comment: sheet.cells?.get("1,0")?.comment?.text,
      formula: sheet.cells?.get("1,1")?.formula,
      dynamic: sheet.cells?.get("1,1")?.formulaDynamic,
    }),
    expected: { comment: "a note", formula: "SUM(B2:B3)", dynamic: true },
  },

  merges: {
    value: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }],
    read: (sheet) => sheet.merges,
  },

  dataValidations: {
    value: [{ type: "list", values: ["North", "South"], range: "A2:A3", allowBlank: true }],
    read: (sheet) => sheet.dataValidations?.map((v) => ({ type: v.type, range: v.range })),
    expected: [{ type: "list", range: "A2:A3" }],
  },

  conditionalRules: {
    value: [
      {
        type: "cellIs",
        priority: 1,
        operator: "greaterThan",
        formula: "15",
        range: "B2:B3",
        style: { font: { bold: true, color: { rgb: "9C0006" } } },
      },
    ],
    read: (sheet) => sheet.conditionalRules,
  },

  autoFilter: { value: { range: "A1:B3" }, read: (sheet) => sheet.autoFilter },

  freezePane: { value: { rows: 1, columns: 1 }, read: (sheet) => sheet.freezePane },

  splitPane: { value: { xSplit: 2000, ySplit: 1000 }, read: (sheet) => sheet.splitPane },

  images: {
    value: [
      {
        data: PNG_1X1,
        type: "png",
        anchor: { from: { row: 4, col: 0 }, to: { row: 9, col: 3 } },
        width: 120,
        height: 80,
        altText: "a chart of nothing",
        title: "Figure 1",
      },
    ],
    read: (sheet) =>
      sheet.images?.map((i) => ({
        type: i.type,
        anchor: i.anchor,
        width: i.width,
        height: i.height,
        altText: i.altText,
        title: i.title,
        bytes: i.data.length,
      })),
    expected: [
      {
        type: "png",
        anchor: { from: { row: 4, col: 0 }, to: { row: 9, col: 3 } },
        width: 120,
        height: 80,
        altText: "a chart of nothing",
        title: "Figure 1",
        bytes: PNG_1X1.length,
      },
    ],
  },

  protection: {
    value: { password: "s3cret", sheet: true, formatCells: true, sort: true },
    read: (sheet) => ({
      sheet: sheet.protection?.sheet,
      formatCells: sheet.protection?.formatCells,
      sort: sheet.protection?.sort,
      // Declared one-way below, and asserted here so the claim stays true:
      // the file stores a hash, and a hash is not the password.
      password: sheet.protection?.password,
    }),
    expected: { sheet: true, formatCells: true, sort: true, password: undefined },
  },

  pageSetup: {
    value: {
      paperSize: "a4",
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      scale: 90,
      printArea: "$A$1:$B$3",
      printTitlesRow: "$1:$1",
      printTitlesColumn: "$A:$A",
      horizontalCentered: true,
      margins: { top: 1, bottom: 1, left: 0.5, right: 0.5, header: 0.3, footer: 0.3 },
    },
    read: (sheet) => sheet.pageSetup,
  },

  headerFooter: {
    value: { oddHeader: "&Ltitle", oddFooter: "&Cpage &P" },
    read: (sheet) => sheet.headerFooter,
  },

  view: {
    value: { showGridLines: false, zoomScale: 125, rightToLeft: true, tabColor: { rgb: "FF0000" } },
    read: (sheet) => sheet.view,
  },

  hidden: { value: true, read: (sheet) => sheet.hidden },

  veryHidden: { value: true, read: (sheet) => sheet.veryHidden },

  tables: {
    value: [
      {
        name: "Sales",
        range: "A1:B3",
        columns: [{ name: "Region" }, { name: "Amount" }],
        style: "TableStyleMedium2",
        showRowStripes: true,
        showAutoFilter: false,
      },
    ],
    read: (sheet) =>
      sheet.tables?.map((t) => ({
        name: t.name,
        range: t.range,
        style: t.style,
        showRowStripes: t.showRowStripes,
        showAutoFilter: t.showAutoFilter,
      })),
    expected: [
      {
        name: "Sales",
        range: "A1:B3",
        style: "TableStyleMedium2",
        showRowStripes: true,
        showAutoFilter: false,
      },
    ],
  },

  rowBreaks: { value: [1], read: (sheet) => sheet.rowBreaks },

  colBreaks: { value: [1], read: (sheet) => sheet.colBreaks },

  rowDefs: {
    value: new Map([[1, { height: 30, hidden: false, outlineLevel: 1 }]]),
    read: (sheet) => ({
      height: sheet.rowDefs?.get(1)?.height,
      outlineLevel: sheet.rowDefs?.get(1)?.outlineLevel,
    }),
    expected: { height: 30, outlineLevel: 1 },
  },

  defaultRowHeight: {
    // Not 15: that is Excel's own default, and the reader deliberately
    // does not surface it — every sheet carries it whether or not the
    // author meant anything by it.
    value: 24,
    read: (sheet) => sheet.defaultRowHeight,
  },

  defaultColWidth: {
    value: 18,
    read: (sheet) => sheet.defaultColWidth,
  },

  outlineProperties: {
    value: { summaryBelow: false, summaryRight: false },
    read: (sheet) => sheet.outlineProperties,
  },

  backgroundImage: {
    value: PNG_1X1,
    read: (sheet) => sheet.backgroundImage?.length,
    expected: PNG_1X1.length,
  },

  sparklines: {
    value: [
      { location: "C2", dataRange: "sparklines!A2:B2", type: "column", color: { rgb: "376092" } },
    ],
    read: (sheet) => sheet.sparklines,
  },

  textBoxes: {
    value: [
      {
        text: "a caption",
        anchor: { from: { row: 4, col: 0 }, to: { row: 7, col: 3 } },
        width: 200,
        height: 60,
        altText: "caption alt",
        title: "Caption",
      },
    ],
    read: (sheet) => sheet.textBoxes,
    expected: [
      {
        text: "a caption",
        anchor: { from: { row: 4, col: 0 }, to: { row: 7, col: 3 } },
        width: 200,
        height: 60,
        altText: "caption alt",
        title: "Caption",
        // The writer always paints a shape and always sizes its text; the
        // reader reports both, so a text box with no `style` comes back
        // carrying the writer's defaults rather than nothing.
        style: { fontSize: 11, fillColor: "FFFFFF", borderColor: "000000" },
      },
    ],
  },

  charts: {
    value: [
      {
        type: "column",
        title: "Sales by region",
        anchor: { from: { row: 4, col: 0 }, to: { row: 14, col: 6 } },
        series: [{ name: "Amount", values: "charts!$B$2:$B$3", categories: "charts!$A$2:$A$3" }],
        altText: "column chart of sales",
        frameTitle: "Sales frame",
      },
    ],
    // Charts read back as `Chart` (an inspection view of chartN.xml), not
    // as the `SheetChart` that authored them.
    read: (sheet) =>
      sheet.charts?.map((c) => ({
        kinds: c.kinds,
        title: c.title,
        anchor: c.anchor,
        altText: c.altText,
        frameTitle: c.frameTitle,
      })),
    expected: [
      {
        kinds: ["bar"],
        title: "Sales by region",
        anchor: { from: { row: 4, col: 0 }, to: { row: 14, col: 6 } },
        altText: "column chart of sales",
        frameTitle: "Sales frame",
      },
    ],
  },

  pivotTables: {
    value: [
      {
        name: "ByRegion",
        sourceRange: "A1:B3",
        targetCell: "D1",
        rows: ["Region"],
        values: [{ field: "Amount", function: "sum" }],
      },
    ],
    // `targetCell` is the anchor; the reader reports the rendered extent
    // the writer computed from it (header + one row per Region + total).
    read: (sheet) => sheet.pivotTables?.map((p) => ({ name: p.name, location: p.location })),
    expected: [{ name: "ByRegion", location: "D1:E3" }],
  },

  a11y: {
    oneWay:
      "Authoring metadata with no cell in the file to live in. `summary` is " +
      "promoted into `properties.description` by writeXlsx (asserted in " +
      "a11y.test.ts) and reads back there; `headerRow` has no OOXML home at " +
      "all — the nearest thing, a table's headerRowCount, means something " +
      "narrower. Neither is a field the reader can honestly reconstruct.",
  },
}

// ── WriteOptions register ────────────────────────────────────────────

const OPTION_FIELDS: { [K in keyof Required<WriteOptions>]: Entry<WriteOptions[K]> } = {
  sheets: {
    oneWay: "The container for everything above, covered field by field by SHEET_FIELDS.",
  },

  properties: {
    value: { title: "Parity", creator: "hucre", company: "ACME", custom: { Reviewed: "yes" } },
    read: (_sheet, wb) => ({
      title: wb.properties?.title,
      creator: wb.properties?.creator,
      company: wb.properties?.company,
      custom: wb.properties?.custom,
    }),
    expected: {
      title: "Parity",
      creator: "hucre",
      company: "ACME",
      custom: { Reviewed: "yes" },
    },
  },

  namedRanges: {
    value: [{ name: "Budget", range: "name!$A$1:$B$3", comment: "the numbers" }],
    read: (_sheet, wb) => wb.namedRanges,
  },

  defaultFont: {
    value: { name: "Georgia", size: 13 },
    read: (_sheet, wb) => wb.defaultFont,
    expected: { name: "Georgia", size: 13 },
  },

  dateSystem: { value: "1904", read: (_sheet, wb) => wb.dateSystem },

  activeSheet: { value: 2, read: (_sheet, wb) => wb.activeSheet },

  workbookProtection: {
    value: { lockStructure: true, lockWindows: true, password: "s3cret" },
    read: (_sheet, wb) => wb.workbookProtection,
    // Same one-way hash as sheet protection; the flags come back, the
    // password does not.
    expected: { lockStructure: true, lockWindows: true },
  },

  stringMode: {
    oneWay:
      "A storage choice, not a property of the workbook: `shared` and " +
      "`inline` describe two encodings of the identical cell values, and " +
      "the reader resolves both to the same strings. There is no field to " +
      "surface, and a reader that reported one would be reporting on the " +
      "file's compression, not its content.",
  },

  vbaProject: {
    oneWay:
      "`Workbook` has no VBA field and should not grow one — handing " +
      "callers a macro binary is a different feature from reading a " +
      "spreadsheet. The part survives the openXlsx → saveXlsx path " +
      "verbatim, which is where preserving it matters (roundtrip.ts).",
  },

  encryption: {
    oneWay:
      "A property of the container, not of the model. The round trip is " +
      "`readXlsx(bytes, { password })` returning the same workbook, which " +
      "encryption.test.ts asserts; there is no encryption field on " +
      "`Workbook` because a decrypted workbook is just a workbook.",
  },
}

// ── One-way register, restated as its own assertion ──────────────────
//
// The two passwords are the entries the issue singled out as correctly
// one-way. They are not in the registers above as `oneWay` because the
// fields they sit on (`protection`, `workbookProtection`) do round-trip —
// it is one key inside each that does not.

const ONE_WAY_PASSWORDS = [
  {
    field: "WriteSheet.protection.password",
    why: "Stored as the legacy 16-bit sheet-protection hash. The hash is what the format defines; the password is not recoverable from it, by design.",
  },
  {
    field: "WriteOptions.workbookProtection.password",
    why: "Same hash, same reason — <workbookProtection workbookPassword> holds a digest, not the secret.",
  },
]

// ── Fixture ──────────────────────────────────────────────────────────

// The registers are declared as mapped types so `tsc` enforces that every
// field is covered. Iterating them is a different job: `Object.entries`
// collapses the per-key value types into a union no loop body can satisfy,
// so walk them through one erased view. The exhaustiveness check lives in
// the declarations above and is unaffected.
type AnyEntry = Entry<unknown>

function entriesOf(register: Record<string, unknown>): Array<[string, AnyEntry]> {
  return Object.entries(register) as Array<[string, AnyEntry]>
}

const SHEET_ENTRIES = entriesOf(SHEET_FIELDS)
const OPTION_ENTRIES = entriesOf(OPTION_FIELDS)

function buildFixture(): WriteOptions {
  const sheets: WriteSheet[] = []

  for (const [key, entry] of SHEET_ENTRIES) {
    if (isOneWay(entry)) continue
    sheets.push({
      name: key,
      rows: BASE_ROWS,
      ...entry.with,
      [key]: entry.value,
    } as WriteSheet)
  }

  const options: Record<string, unknown> = { sheets }
  for (const [key, entry] of OPTION_ENTRIES) {
    if (isOneWay(entry)) continue
    options[key] = entry.value
  }
  return options as unknown as WriteOptions
}

describe("xlsx write → read parity", () => {
  it("gives back everything the writer was given", async () => {
    const fixture = buildFixture()
    const bytes = await writeXlsx(fixture)
    const workbook = await readXlsx(bytes, { readStyles: true })

    const byName = new Map(workbook.sheets.map((s) => [s.name, s]))
    const failures: string[] = []

    for (const [key, entry] of SHEET_ENTRIES) {
      if (isOneWay(entry)) continue
      const sheet = byName.get(key)
      if (!sheet) {
        failures.push(`WriteSheet.${key}: sheet "${key}" is missing from the parsed workbook`)
        continue
      }
      const actual = entry.read(sheet, workbook)
      const expected = "expected" in entry ? entry.expected : entry.value
      try {
        expect(actual).toEqual(expected)
      } catch {
        failures.push(
          `WriteSheet.${key}: wrote ${JSON.stringify(expected)}, read ${JSON.stringify(actual)}`,
        )
      }
    }

    // Option-level probes read off the workbook; the sheet argument is the
    // first one purely so the two registers share a signature.
    const anySheet = workbook.sheets[0]
    for (const [key, entry] of OPTION_ENTRIES) {
      if (isOneWay(entry)) continue
      const actual = entry.read(anySheet, workbook)
      const expected = "expected" in entry ? entry.expected : entry.value
      try {
        expect(actual).toEqual(expected)
      } catch {
        failures.push(
          `WriteOptions.${key}: wrote ${JSON.stringify(expected)}, read ${JSON.stringify(actual)}`,
        )
      }
    }

    expect(failures).toEqual([])
  })

  it("states a reason for every field that does not come back", () => {
    const oneWays = [
      ...SHEET_ENTRIES.map(([k, e]) => [`WriteSheet.${k}`, e] as const),
      ...OPTION_ENTRIES.map(([k, e]) => [`WriteOptions.${k}`, e] as const),
    ].filter(([, e]) => isOneWay(e))

    for (const [field, entry] of oneWays) {
      const reason = (entry as OneWay).oneWay
      // A one-way entry earns its place by explaining itself. A bare
      // "TODO" or an empty string is a gap wearing a register's clothes.
      expect(reason.length, `${field} has no stated reason`).toBeGreaterThan(40)
      expect(reason).not.toMatch(/\bTODO\b|not implemented|for now/i)
    }

    // Pin the shape of the register so a field cannot be quietly demoted
    // from "round-trips" to "one-way" without the change showing up here.
    expect(oneWays.map(([f]) => f)).toEqual([
      "WriteSheet.a11y",
      "WriteOptions.sheets",
      "WriteOptions.stringMode",
      "WriteOptions.vbaProject",
      "WriteOptions.encryption",
    ])
  })

  it("keeps sheet and workbook passwords one-way, and says so", async () => {
    for (const entry of ONE_WAY_PASSWORDS) {
      expect(entry.why.length).toBeGreaterThan(40)
    }

    const bytes = await writeXlsx({
      workbookProtection: { lockStructure: true, password: "s3cret" },
      sheets: [{ name: "S", rows: [["x"]], protection: { sheet: true, password: "s3cret" } }],
    })
    const workbook = await readXlsx(bytes)
    const zip = new ZipReader(bytes)
    const sheetXml = new TextDecoder().decode(await zip.extract("xl/worksheets/sheet1.xml"))
    const bookXml = new TextDecoder().decode(await zip.extract("xl/workbook.xml"))

    // Not merely absent — absent *because* the file holds a digest. Asserting
    // the digest is present keeps "one-way" distinguishable from "dropped":
    // the protection is really in the file, only the secret is not.
    expect(workbook.sheets[0].protection?.password).toBeUndefined()
    expect(sheetXml).toMatch(/<sheetProtection[^>]*password="[0-9A-F]{4}"/)
    expect(sheetXml).not.toContain("s3cret")

    expect(workbook.workbookProtection?.lockStructure).toBe(true)
    expect(bookXml).toMatch(/<workbookProtection[^>]*workbookPassword="[0-9A-F]{4}"/)
    expect(bookXml).not.toContain("s3cret")
  })
})
