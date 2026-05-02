import { describe, it, expect } from "vitest";
import {
  parsePivotCacheDefinition,
  parsePivotTable,
  attachPivotCacheFields,
} from "../src/xlsx/pivot-reader";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";

const encoder = new TextEncoder();

// ── parsePivotCacheDefinition ──────────────────────────────────────

describe("parsePivotCacheDefinition", () => {
  it("returns undefined when the root element is wrong", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<notACache xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`;
    expect(parsePivotCacheDefinition(xml)).toBeUndefined();
  });

  it("parses worksheet source range and field names in declaration order", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                      r:id="rId1" recordCount="3">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C4" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Product" numFmtId="0"/>
    <cacheField name="Revenue" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>`;
    const cache = parsePivotCacheDefinition(xml);
    expect(cache).toBeDefined();
    expect(cache!.sourceType).toBe("worksheet");
    expect(cache!.sourceRef).toBe("A1:C4");
    expect(cache!.sourceSheet).toBe("Data");
    expect(cache!.fieldNames).toEqual(["Region", "Product", "Revenue"]);
  });

  it("falls back to the worksheetSource name attribute when ref is absent", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource name="SalesTable"/>
  </cacheSource>
  <cacheFields count="1"><cacheField name="X"/></cacheFields>
</pivotCacheDefinition>`;
    const cache = parsePivotCacheDefinition(xml);
    expect(cache!.sourceRef).toBe("SalesTable");
    expect(cache!.sourceSheet).toBeUndefined();
  });

  it("uses an empty string when a cacheField is missing its name attribute", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheFields count="2">
    <cacheField name="Has"/>
    <cacheField/>
  </cacheFields>
</pivotCacheDefinition>`;
    const cache = parsePivotCacheDefinition(xml);
    expect(cache!.fieldNames).toEqual(["Has", ""]);
  });

  it("ignores an unknown cacheSource type", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="hypothetical"/>
</pivotCacheDefinition>`;
    const cache = parsePivotCacheDefinition(xml);
    expect(cache!.sourceType).toBeUndefined();
  });
});

// ── parsePivotTable ────────────────────────────────────────────────

describe("parsePivotTable", () => {
  it("returns undefined when the root has no name attribute", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`;
    expect(parsePivotTable(xml)).toBeUndefined();
  });

  it("parses cacheId, location, and field axes", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     name="SalesPivot" cacheId="0" applyNumberFormats="0"
                     applyBorderFormats="0" applyFontFormats="0"
                     applyPatternFormats="0" applyAlignmentFormats="0"
                     applyWidthHeightFormats="1" dataCaption="Values"
                     updatedVersion="6" minRefreshableVersion="3">
  <location ref="A3:D20" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" showAll="0"/>
    <pivotField axis="axisCol" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <dataFields count="1">
    <dataField name="Sum of Revenue" fld="2" baseField="0" baseItem="0" subtotal="sum"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1"
                       showColHeaders="1" showRowStripes="0" showColStripes="0"
                       showLastColumn="1"/>
</pivotTableDefinition>`;
    const pivot = parsePivotTable(xml);
    expect(pivot).toBeDefined();
    expect(pivot!.name).toBe("SalesPivot");
    expect(pivot!.cacheId).toBe(0);
    expect(pivot!.location).toBe("A3:D20");
    expect(pivot!.firstHeaderRow).toBe(1);
    expect(pivot!.firstDataRow).toBe(2);
    expect(pivot!.firstDataCol).toBe(1);
    expect(pivot!.fields).toHaveLength(3);
    expect(pivot!.fields[0].axis).toBe("row");
    expect(pivot!.fields[1].axis).toBe("col");
    expect(pivot!.fields[2].axis).toBe("data");
    expect(pivot!.fields[2].function).toBe("sum");
    expect(pivot!.fields[2].displayName).toBe("Sum of Revenue");
    expect(pivot!.styleName).toBe("PivotStyleLight16");
    expect(pivot!.dataCaption).toBe("Values");
  });

  it("treats fields without an axis attribute and no dataField flag as hidden", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     name="P" cacheId="0">
  <location ref="A1:A1"/>
  <pivotFields count="1">
    <pivotField showAll="0"/>
  </pivotFields>
</pivotTableDefinition>`;
    const pivot = parsePivotTable(xml);
    expect(pivot!.fields[0].axis).toBe("hidden");
  });

  it("maps axisValues / dataField=true / countA appropriately", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     name="P" cacheId="0">
  <location ref="A1:A1"/>
  <pivotFields count="2">
    <pivotField axis="axisValues"/>
    <pivotField dataField="true"/>
  </pivotFields>
  <dataFields count="2">
    <dataField fld="0" subtotal="countA"/>
    <dataField fld="1" subtotal="average"/>
  </dataFields>
</pivotTableDefinition>`;
    const pivot = parsePivotTable(xml);
    expect(pivot!.fields[0].axis).toBe("data");
    expect(pivot!.fields[0].function).toBe("count"); // countA collapses to count
    expect(pivot!.fields[1].axis).toBe("data");
    expect(pivot!.fields[1].function).toBe("average");
  });

  it("falls back to cacheId 0 when the attribute is missing or unparseable", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     name="P" cacheId="oops">
  <location ref="A1:A1"/>
</pivotTableDefinition>`;
    const pivot = parsePivotTable(xml);
    expect(pivot!.cacheId).toBe(0);
  });
});

// ── attachPivotCacheFields ─────────────────────────────────────────

describe("attachPivotCacheFields", () => {
  it("overlays cache field names and leaves unmatched indexes alone", () => {
    const pivot = parsePivotTable(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     name="P" cacheId="0">
  <location ref="A1:A1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow"/>
    <pivotField axis="axisCol"/>
    <pivotField dataField="1"/>
  </pivotFields>
</pivotTableDefinition>`)!;
    expect(pivot.fields.map((f) => f.name)).toEqual(["field1", "field2", "field3"]);
    attachPivotCacheFields(pivot, {
      cacheId: 0,
      fieldNames: ["Region", "Product"], // shorter than pivot.fields
    });
    expect(pivot.fields.map((f) => f.name)).toEqual(["Region", "Product", "field3"]);
  });
});

// ── Integration: end-to-end XLSX with pivot table ──────────────────

async function buildXlsxWithPivotTable(): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>`),
  );

  z.add(
    "_rels/.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/workbook.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data" sheetId="1" r:id="rId1"/>
    <sheet name="Pivot" sheetId="2" r:id="rId2"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="0" r:id="rId3"/>
  </pivotCaches>
</workbook>`),
  );

  z.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`),
  );

  // Source data sheet
  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Region</t></is></c><c r="B1" t="inlineStr"><is><t>Revenue</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>EU</t></is></c><c r="B2"><v>100</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>US</t></is></c><c r="B3"><v>200</v></c></row>
  </sheetData>
</worksheet>`),
  );

  // Sheet that hosts the pivot table — its rels file declares the pivotTable rel.
  z.add(
    "xl/worksheets/sheet2.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
  );
  z.add(
    "xl/worksheets/_rels/sheet2.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/pivotTables/pivotTable1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                     name="SalesPivot" cacheId="0" applyNumberFormats="0"
                     applyBorderFormats="0" applyFontFormats="0"
                     applyPatternFormats="0" applyAlignmentFormats="0"
                     applyWidthHeightFormats="1" dataCaption="Values">
  <location ref="A3:B5" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <dataFields count="1">
    <dataField name="Sum of Revenue" fld="1" subtotal="sum"/>
  </dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16"/>
</pivotTableDefinition>`),
  );
  z.add(
    "xl/pivotTables/_rels/pivotTable1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/pivotCache/pivotCacheDefinition1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                      r:id="rId1" recordCount="2">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B3" sheet="Data"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Region" numFmtId="0"/>
    <cacheField name="Revenue" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>`),
  );
  z.add(
    "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/pivotCache/pivotCacheRecords1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
  <r><s v="EU"/><n v="100"/></r>
  <r><s v="US"/><n v="200"/></r>
</pivotCacheRecords>`),
  );

  return await z.build();
}

async function zipExtractText(buf: Uint8Array, path: string): Promise<string> {
  const zr = new ZipReader(buf);
  return new TextDecoder("utf-8").decode(await zr.extract(path));
}

function zipHas(buf: Uint8Array, path: string): boolean {
  const zr = new ZipReader(buf);
  return zr.has(path);
}

// ── Reader integration ────────────────────────────────────────────

describe("readXlsx — pivot table integration", () => {
  it("attaches workbook.pivotCaches with parsed source range and field names", async () => {
    const buf = await buildXlsxWithPivotTable();
    const wb = await readXlsx(buf);
    expect(wb.pivotCaches).toBeDefined();
    expect(wb.pivotCaches).toHaveLength(1);
    const cache = wb.pivotCaches![0];
    expect(cache.cacheId).toBe(0);
    expect(cache.sourceType).toBe("worksheet");
    expect(cache.sourceRef).toBe("A1:B3");
    expect(cache.sourceSheet).toBe("Data");
    expect(cache.fieldNames).toEqual(["Region", "Revenue"]);
    expect(cache.hasRecords).toBe(true);
  });

  it("attaches sheet.pivotTables with names overlaid from the cache", async () => {
    const buf = await buildXlsxWithPivotTable();
    const wb = await readXlsx(buf);
    const pivotSheet = wb.sheets.find((s) => s.name === "Pivot");
    expect(pivotSheet?.pivotTables).toHaveLength(1);
    const pt = pivotSheet!.pivotTables![0];
    expect(pt.name).toBe("SalesPivot");
    expect(pt.cacheId).toBe(0);
    expect(pt.location).toBe("A3:B5");
    expect(pt.fields).toHaveLength(2);
    expect(pt.fields[0]).toMatchObject({ name: "Region", axis: "row" });
    expect(pt.fields[1]).toMatchObject({
      name: "Revenue",
      axis: "data",
      function: "sum",
      displayName: "Sum of Revenue",
    });
    expect(pt.styleName).toBe("PivotStyleLight16");
  });

  it("does not set pivotCaches when the workbook has none", async () => {
    const z = new ZipWriter();
    z.add(
      "[Content_Types].xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    );
    z.add(
      "_rels/.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/workbook.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Main" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    );
    z.add(
      "xl/_rels/workbook.xml.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/worksheets/sheet1.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
    );

    const wb = await readXlsx(await z.build());
    expect(wb.pivotCaches).toBeUndefined();
    for (const sheet of wb.sheets) {
      expect(sheet.pivotTables).toBeUndefined();
    }
  });
});

// ── Roundtrip ──────────────────────────────────────────────────────

describe("openXlsx -> saveXlsx — pivot table roundtrip", () => {
  it("preserves the pivot table, cache definition, and records on roundtrip", async () => {
    const buf = await buildXlsxWithPivotTable();
    const wb = await openXlsx(buf);
    const saved = await saveXlsx(wb);

    // All four pivot parts survive intact.
    expect(zipHas(saved, "xl/pivotTables/pivotTable1.xml")).toBe(true);
    expect(zipHas(saved, "xl/pivotTables/_rels/pivotTable1.xml.rels")).toBe(true);
    expect(zipHas(saved, "xl/pivotCache/pivotCacheDefinition1.xml")).toBe(true);
    expect(zipHas(saved, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels")).toBe(true);
    expect(zipHas(saved, "xl/pivotCache/pivotCacheRecords1.xml")).toBe(true);
  });

  it("re-emits the references in workbook.xml, workbook.xml.rels, and content types", async () => {
    const buf = await buildXlsxWithPivotTable();
    const wb = await openXlsx(buf);
    const saved = await saveXlsx(wb);

    const wbXml = await zipExtractText(saved, "xl/workbook.xml");
    expect(wbXml).toContain("<pivotCaches>");
    expect(wbXml).toMatch(/<pivotCache cacheId="0" [^>]*r:id="rId\d+"/);

    const wbRels = await zipExtractText(saved, "xl/_rels/workbook.xml.rels");
    expect(wbRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"',
    );
    expect(wbRels).toContain('Target="pivotCache/pivotCacheDefinition1.xml"');

    const ct = await zipExtractText(saved, "[Content_Types].xml");
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

  it("re-declares the pivot table relationship in the host sheet's rels", async () => {
    const buf = await buildXlsxWithPivotTable();
    const wb = await openXlsx(buf);
    const saved = await saveXlsx(wb);

    const sheet2Rels = await zipExtractText(saved, "xl/worksheets/_rels/sheet2.xml.rels");
    expect(sheet2Rels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"',
    );
    expect(sheet2Rels).toContain('Target="../pivotTables/pivotTable1.xml"');
  });

  it("re-reading the saved workbook recovers the same pivot table model", async () => {
    const buf = await buildXlsxWithPivotTable();
    const wb = await openXlsx(buf);
    const saved = await saveXlsx(wb);
    const reread = await readXlsx(saved);

    expect(reread.pivotCaches).toHaveLength(1);
    expect(reread.pivotCaches?.[0].fieldNames).toEqual(["Region", "Revenue"]);

    const pivotSheet = reread.sheets.find((s) => s.name === "Pivot");
    expect(pivotSheet?.pivotTables).toHaveLength(1);
    const pt = pivotSheet!.pivotTables![0];
    expect(pt.name).toBe("SalesPivot");
    expect(pt.fields[1].function).toBe("sum");
  });

  it("emits clean rels when the workbook has no pivot tables", async () => {
    // A zero-pivot workbook should not gain any pivot-cache references.
    const z = new ZipWriter();
    z.add(
      "[Content_Types].xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    );
    z.add(
      "_rels/.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/workbook.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    );
    z.add(
      "xl/_rels/workbook.xml.rels",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    );
    z.add(
      "xl/worksheets/sheet1.xml",
      encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`),
    );

    const wb = await openXlsx(await z.build());
    const saved = await saveXlsx(wb);
    const wbXml = await zipExtractText(saved, "xl/workbook.xml");
    expect(wbXml).not.toContain("<pivotCaches>");
    const wbRels = await zipExtractText(saved, "xl/_rels/workbook.xml.rels");
    expect(wbRels).not.toContain("pivotCacheDefinition");
    const ct = await zipExtractText(saved, "[Content_Types].xml");
    expect(ct).not.toContain("/xl/pivotCache/");
    expect(ct).not.toContain("/xl/pivotTables/");
  });
});
