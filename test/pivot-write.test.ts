import { describe, it, expect } from "vitest";
import { writeXlsx } from "../src/xlsx/writer";
import { readXlsx } from "../src/xlsx/reader";
import { ZipReader } from "../src/zip/reader";
import { resolvePivotSource, writePivotTable } from "../src/xlsx/pivot-writer";
import type { WritePivotTable, WriteSheet } from "../src/_types";

const decoder = new TextDecoder("utf-8");

async function extract(buf: Uint8Array, path: string): Promise<string> {
  const z = new ZipReader(buf);
  return decoder.decode(await z.extract(path));
}

function has(buf: Uint8Array, path: string): boolean {
  return new ZipReader(buf).has(path);
}

// ── resolvePivotSource ────────────────────────────────────────────────

describe("resolvePivotSource", () => {
  const baseRows = [
    ["Region", "Product", "Revenue"],
    ["EU", "A", 100],
    ["US", "B", 200],
    ["EU", "B", 50],
  ] as const;

  it("collects field names from the header row", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    };
    const r = resolvePivotSource(pivot, "Data", baseRows);
    expect(r.fieldNames).toEqual(["Region", "Product", "Revenue"]);
    expect(r.dataRows).toHaveLength(3);
    expect(r.sheetName).toBe("Data");
  });

  it("auto-derives the source ref when not supplied", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    };
    const r = resolvePivotSource(pivot, "Data", baseRows);
    expect(r.ref).toBe("A1:C4");
  });

  it("honours an explicit sourceRange", () => {
    const pivot: WritePivotTable = {
      name: "P",
      sourceRange: "Sheet1!$A$1:$C$4",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    };
    const r = resolvePivotSource(pivot, "Data", baseRows);
    expect(r.ref).toBe("Sheet1!$A$1:$C$4");
  });

  it("throws when a referenced field is missing", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["NotAField"],
      values: [{ field: "Revenue" }],
    };
    expect(() => resolvePivotSource(pivot, "Data", baseRows)).toThrow(/NotAField/);
  });

  it("throws when the source has fewer than two rows", () => {
    const pivot: WritePivotTable = {
      name: "P",
      values: [{ field: "Revenue" }],
    };
    expect(() => resolvePivotSource(pivot, "Data", [["Header"]])).toThrow(/header.*data row/i);
  });

  it("pads short data rows with null and trims long rows to the header width", () => {
    const pivot: WritePivotTable = {
      name: "P",
      rows: ["Region"],
      values: [{ field: "Revenue" }],
    };
    const rows = [["Region", "Revenue"], ["EU", 100, "extra"], ["US"]];
    const r = resolvePivotSource(pivot, "Data", rows);
    expect(r.dataRows).toEqual([
      ["EU", 100],
      ["US", null],
    ]);
  });
});

// ── writePivotTable (unit) ────────────────────────────────────────────

describe("writePivotTable", () => {
  function build(pivot: WritePivotTable, rows: ReadonlyArray<ReadonlyArray<unknown>>) {
    const resolved = resolvePivotSource(
      pivot,
      "Data",
      rows as ReadonlyArray<ReadonlyArray<import("../src/_types").CellValue>>,
    );
    return writePivotTable(pivot, resolved, 0);
  }

  it("emits a complete cache definition with worksheetSource and cacheFields", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 200],
      ],
    );
    expect(parts.cacheDefinitionXml).toContain("<pivotCacheDefinition");
    expect(parts.cacheDefinitionXml).toContain('<cacheSource type="worksheet">');
    expect(parts.cacheDefinitionXml).toContain('ref="A1:B3"');
    expect(parts.cacheDefinitionXml).toContain('sheet="Data"');
    expect(parts.cacheDefinitionXml).toContain('<cacheFields count="2">');
    expect(parts.cacheDefinitionXml).toContain('<cacheField name="Region"');
    expect(parts.cacheDefinitionXml).toContain('<cacheField name="Revenue"');
    expect(parts.cacheDefinitionXml).toContain('refreshOnLoad="1"');
  });

  it("collects unique string values into sharedItems for row/col fields", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 200],
        ["EU", 50],
      ],
    );
    // Region cacheField gets two items in declaration order
    const regionBlock = parts.cacheDefinitionXml.match(
      /<cacheField name="Region"[^>]*>[\s\S]*?<\/cacheField>/,
    )?.[0];
    expect(regionBlock).toBeDefined();
    expect(regionBlock).toContain('<sharedItems count="2">');
    expect(regionBlock).toContain('<s v="EU"/>');
    expect(regionBlock).toContain('<s v="US"/>');
  });

  it("describes numeric data fields with min / max / containsNumber", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 250],
      ],
    );
    const revBlock = parts.cacheDefinitionXml.match(
      /<cacheField name="Revenue"[^>]*>[\s\S]*?<\/cacheField>/,
    )?.[0];
    expect(revBlock).toContain('containsNumber="1"');
    expect(revBlock).toContain('minValue="100"');
    expect(revBlock).toContain('maxValue="250"');
    expect(revBlock).toContain('containsInteger="1"');
  });

  it("emits <r> records that map string values to shared-item indexes", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 200],
        ["EU", 50],
      ],
    );
    expect(parts.cacheRecordsXml).toContain('count="3"');
    // EU=0, US=1
    expect(parts.cacheRecordsXml).toContain('<x v="0"/><n v="100"/>');
    expect(parts.cacheRecordsXml).toContain('<x v="1"/><n v="200"/>');
    expect(parts.cacheRecordsXml).toContain('<x v="0"/><n v="50"/>');
  });

  it("emits <m/> for blank cells", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", null],
        ["US", 200],
      ],
    );
    expect(parts.cacheRecordsXml).toContain('<x v="0"/><m/>');
  });

  it("places fields on row, column, and data axes by name", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        columns: ["Product"],
        values: [{ field: "Revenue", function: "sum" }],
      },
      [
        ["Region", "Product", "Revenue"],
        ["EU", "A", 100],
        ["US", "B", 200],
      ],
    );
    expect(parts.pivotTableXml).toMatch(/<pivotField axis="axisRow"[^>]*>/);
    expect(parts.pivotTableXml).toMatch(/<pivotField axis="axisCol"[^>]*>/);
    expect(parts.pivotTableXml).toMatch(/<pivotField dataField="1"[^>]*\/>/);
    expect(parts.pivotTableXml).toContain('<rowFields count="1"><field x="0"/>');
    expect(parts.pivotTableXml).toContain('<colFields count="1"><field x="1"/>');
    expect(parts.pivotTableXml).toContain('<dataFields count="1">');
    expect(parts.pivotTableXml).toContain('name="Sum of Revenue"');
    expect(parts.pivotTableXml).toContain('fld="2"');
  });

  it("auto-labels data fields per Excel's convention", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [
          { field: "Revenue", function: "average" },
          { field: "Revenue", function: "max" },
        ],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 200],
      ],
    );
    expect(parts.pivotTableXml).toContain('name="Average of Revenue"');
    expect(parts.pivotTableXml).toContain('subtotal="average"');
    expect(parts.pivotTableXml).toContain('name="Max of Revenue"');
    expect(parts.pivotTableXml).toContain('subtotal="max"');
  });

  it("honours an explicit displayName override", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue", displayName: "Total Sales" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 200],
      ],
    );
    expect(parts.pivotTableXml).toContain('name="Total Sales"');
  });

  it("declares the targetCell location and applies a sensible default size", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
        targetCell: "C5",
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
        ["US", 200],
      ],
    );
    // C5 is col=2, row=4. Default 4-cell area → C5:D7 or wider for data fields.
    expect(parts.pivotTableXml).toMatch(/<location ref="C5:[A-Z]+\d+"/);
  });

  it("emits a pivotTable rels file targeting the cache definition", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
      ],
    );
    expect(parts.pivotTableRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"',
    );
    expect(parts.pivotTableRels).toContain('Target="../pivotCache/pivotCacheDefinition1.xml"');
  });

  it("emits a cacheDefinition rels file targeting the cache records", () => {
    const parts = build(
      {
        name: "P",
        rows: ["Region"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Region", "Revenue"],
        ["EU", 100],
      ],
    );
    expect(parts.cacheDefinitionRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"',
    );
    expect(parts.cacheDefinitionRels).toContain('Target="pivotCacheRecords1.xml"');
  });

  it("uses PivotStyleLight16 by default and honours overrides", () => {
    const def = build({ name: "P", rows: ["R"], values: [{ field: "V" }] }, [
      ["R", "V"],
      ["x", 1],
    ]);
    expect(def.pivotTableXml).toContain('name="PivotStyleLight16"');

    const override = build(
      { name: "P", styleName: "PivotStyleDark1", rows: ["R"], values: [{ field: "V" }] },
      [
        ["R", "V"],
        ["x", 1],
      ],
    );
    expect(override.pivotTableXml).toContain('name="PivotStyleDark1"');
  });

  it("infers numeric type from data even when the column is also a row field", () => {
    // A field that's used on the row axis but happens to hold integers
    // still gets shared-items so row-item indexes work. This documents
    // that the row-axis type wins over the numeric inference.
    const parts = build(
      {
        name: "P",
        rows: ["Year"],
        values: [{ field: "Revenue" }],
      },
      [
        ["Year", "Revenue"],
        [2020, 100],
        [2021, 200],
      ],
    );
    // Year is numeric, so no sharedItems with `<s/>` entries — it's
    // serialized as <n v=".."/> in records.
    expect(parts.cacheRecordsXml).toContain('<n v="2020"/>');
    expect(parts.cacheRecordsXml).toContain('<n v="2021"/>');
  });
});

// ── End-to-end through writeXlsx ──────────────────────────────────────

describe("writeXlsx — pivot tables", () => {
  function buildSheet(): { dataSheet: WriteSheet; pivotSheet: WriteSheet } {
    const dataSheet: WriteSheet = {
      name: "Data",
      rows: [
        ["Region", "Product", "Revenue"],
        ["EU", "A", 100],
        ["EU", "B", 50],
        ["US", "A", 200],
        ["US", "B", 75],
      ],
    };
    const pivotSheet: WriteSheet = {
      name: "Pivot",
      rows: [],
      pivotTables: [
        {
          name: "SalesPivot",
          sourceSheet: "Data",
          rows: ["Region"],
          columns: ["Product"],
          values: [{ field: "Revenue", function: "sum" }],
        },
      ],
    };
    return { dataSheet, pivotSheet };
  }

  it("emits all five pivot parts in the ZIP", async () => {
    const { dataSheet, pivotSheet } = buildSheet();
    const buf = await writeXlsx({ sheets: [dataSheet, pivotSheet] });

    expect(has(buf, "xl/pivotTables/pivotTable1.xml")).toBe(true);
    expect(has(buf, "xl/pivotTables/_rels/pivotTable1.xml.rels")).toBe(true);
    expect(has(buf, "xl/pivotCache/pivotCacheDefinition1.xml")).toBe(true);
    expect(has(buf, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels")).toBe(true);
    expect(has(buf, "xl/pivotCache/pivotCacheRecords1.xml")).toBe(true);
  });

  it("declares overrides for every pivot part in [Content_Types].xml", async () => {
    const { dataSheet, pivotSheet } = buildSheet();
    const buf = await writeXlsx({ sheets: [dataSheet, pivotSheet] });
    const ct = await extract(buf, "[Content_Types].xml");

    expect(ct).toContain("/xl/pivotTables/pivotTable1.xml");
    expect(ct).toContain(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    );
    expect(ct).toContain("/xl/pivotCache/pivotCacheDefinition1.xml");
    expect(ct).toContain(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
    );
    expect(ct).toContain("/xl/pivotCache/pivotCacheRecords1.xml");
    expect(ct).toContain(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
    );
  });

  it("wires <pivotCaches> into workbook.xml and the cache rel into workbook.xml.rels", async () => {
    const { dataSheet, pivotSheet } = buildSheet();
    const buf = await writeXlsx({ sheets: [dataSheet, pivotSheet] });
    const wbXml = await extract(buf, "xl/workbook.xml");
    const wbRels = await extract(buf, "xl/_rels/workbook.xml.rels");

    expect(wbXml).toContain("<pivotCaches>");
    expect(wbXml).toMatch(/<pivotCache cacheId="0"[^>]*r:id="rIdPivot1"/);
    expect(wbRels).toContain('Id="rIdPivot1"');
    expect(wbRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"',
    );
    expect(wbRels).toContain('Target="pivotCache/pivotCacheDefinition1.xml"');
  });

  it("declares the pivotTable relationship in the host sheet's rels", async () => {
    const { dataSheet, pivotSheet } = buildSheet();
    const buf = await writeXlsx({ sheets: [dataSheet, pivotSheet] });
    const rels = await extract(buf, "xl/worksheets/_rels/sheet2.xml.rels");
    expect(rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"',
    );
    expect(rels).toContain('Target="../pivotTables/pivotTable1.xml"');
  });

  it("does not create pivot wiring on workbooks without pivots", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Plain",
          rows: [
            ["a", "b"],
            [1, 2],
          ],
        },
      ],
    });
    const ct = await extract(buf, "[Content_Types].xml");
    expect(ct).not.toContain("/xl/pivotTables/");
    expect(ct).not.toContain("/xl/pivotCache/");
    const wb = await extract(buf, "xl/workbook.xml");
    expect(wb).not.toContain("<pivotCaches>");
  });

  it("supports a pivot that sources data from its own sheet", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Self",
          rows: [
            ["Region", "Revenue"],
            ["EU", 100],
            ["US", 200],
          ],
          pivotTables: [
            {
              name: "SelfPivot",
              rows: ["Region"],
              values: [{ field: "Revenue" }],
            },
          ],
        },
      ],
    });
    const cacheDef = await extract(buf, "xl/pivotCache/pivotCacheDefinition1.xml");
    expect(cacheDef).toContain('sheet="Self"');
    expect(cacheDef).toContain('ref="A1:B3"');
  });

  it("re-reading the workbook recovers the pivot model", async () => {
    const { dataSheet, pivotSheet } = buildSheet();
    const buf = await writeXlsx({ sheets: [dataSheet, pivotSheet] });
    const wb = await readXlsx(buf);

    expect(wb.pivotCaches).toHaveLength(1);
    expect(wb.pivotCaches?.[0].fieldNames).toEqual(["Region", "Product", "Revenue"]);

    const pivotHost = wb.sheets.find((s) => s.name === "Pivot");
    expect(pivotHost?.pivotTables).toHaveLength(1);
    const pt = pivotHost!.pivotTables![0];
    expect(pt.name).toBe("SalesPivot");
    expect(pt.fields[0]).toMatchObject({ name: "Region", axis: "row" });
    expect(pt.fields[1]).toMatchObject({ name: "Product", axis: "col" });
    expect(pt.fields[2]).toMatchObject({
      name: "Revenue",
      axis: "data",
      function: "sum",
      displayName: "Sum of Revenue",
    });
  });

  it("writes multiple pivot tables with sequentially numbered indices", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Data",
          rows: [
            ["Region", "Revenue"],
            ["EU", 100],
            ["US", 200],
          ],
        },
        {
          name: "Pivot1",
          pivotTables: [
            {
              name: "P1",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [{ field: "Revenue" }],
            },
          ],
        },
        {
          name: "Pivot2",
          pivotTables: [
            {
              name: "P2",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [{ field: "Revenue", function: "average" }],
            },
          ],
        },
      ],
    });
    expect(has(buf, "xl/pivotTables/pivotTable1.xml")).toBe(true);
    expect(has(buf, "xl/pivotTables/pivotTable2.xml")).toBe(true);
    expect(has(buf, "xl/pivotCache/pivotCacheDefinition1.xml")).toBe(true);
    expect(has(buf, "xl/pivotCache/pivotCacheDefinition2.xml")).toBe(true);
    expect(has(buf, "xl/pivotCache/pivotCacheRecords1.xml")).toBe(true);
    expect(has(buf, "xl/pivotCache/pivotCacheRecords2.xml")).toBe(true);

    const wbXml = await extract(buf, "xl/workbook.xml");
    expect(wbXml).toMatch(/<pivotCache cacheId="0"[^>]*r:id="rIdPivot1"/);
    expect(wbXml).toMatch(/<pivotCache cacheId="1"[^>]*r:id="rIdPivot2"/);
  });

  it("supports object-style sheet data via columns + data", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Data",
          columns: [
            { header: "Region", key: "region" },
            { header: "Revenue", key: "revenue" },
          ],
          data: [
            { region: "EU", revenue: 100 },
            { region: "US", revenue: 200 },
          ],
        },
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "P",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [{ field: "Revenue" }],
            },
          ],
        },
      ],
    });
    const cacheDef = await extract(buf, "xl/pivotCache/pivotCacheDefinition1.xml");
    expect(cacheDef).toContain('<cacheField name="Region"');
    expect(cacheDef).toContain('<cacheField name="Revenue"');
  });

  it("throws when sourceSheet does not exist in the workbook", async () => {
    await expect(
      writeXlsx({
        sheets: [
          {
            name: "Pivot",
            pivotTables: [
              {
                name: "Bad",
                sourceSheet: "Missing",
                rows: ["Region"],
                values: [{ field: "Revenue" }],
              },
            ],
          },
        ],
      }),
    ).rejects.toThrow(/Missing/);
  });

  it("supports every aggregation function", async () => {
    const buf = await writeXlsx({
      sheets: [
        {
          name: "Data",
          rows: [
            ["Region", "Revenue"],
            ["EU", 100],
            ["US", 200],
          ],
        },
        {
          name: "Pivot",
          pivotTables: [
            {
              name: "Multi",
              sourceSheet: "Data",
              rows: ["Region"],
              values: [
                { field: "Revenue", function: "sum" },
                { field: "Revenue", function: "average" },
                { field: "Revenue", function: "count" },
                { field: "Revenue", function: "max" },
                { field: "Revenue", function: "min" },
              ],
            },
          ],
        },
      ],
    });
    const pt = await extract(buf, "xl/pivotTables/pivotTable1.xml");
    // sum is the OOXML default and is omitted from the attribute list.
    expect(pt).not.toMatch(/subtotal="sum"/);
    expect(pt).toContain('subtotal="average"');
    expect(pt).toContain('subtotal="count"');
    expect(pt).toContain('subtotal="max"');
    expect(pt).toContain('subtotal="min"');
  });
});
