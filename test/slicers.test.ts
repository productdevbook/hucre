import { describe, it, expect } from "vitest";
import {
  parseSlicers,
  parseSlicerCache,
  parseTimelines,
  parseTimelineCache,
} from "../src/xlsx/slicer-reader";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");

// ── parseSlicers ─────────────────────────────────────────────────

describe("parseSlicers", () => {
  it("returns an empty array when no slicer elements are present", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>`;
    expect(parseSlicers(xml)).toEqual([]);
  });

  it("parses required attributes and skips entries missing name or cache", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Region" cache="Slicer_Region" caption="Region"/>
  <slicer name="OnlyName"/>
  <slicer cache="Slicer_NoName"/>
  <slicer name="Year" cache="Slicer_Year" caption="Year" columnCount="2"
          style="SlicerStyleLight2" sortOrder="ascending" rowHeight="241300"/>
</slicers>`;
    const slicers = parseSlicers(xml);
    expect(slicers).toHaveLength(2);
    expect(slicers[0]).toEqual({
      name: "Region",
      cache: "Slicer_Region",
      caption: "Region",
    });
    expect(slicers[1]).toEqual({
      name: "Year",
      cache: "Slicer_Year",
      caption: "Year",
      columnCount: 2,
      style: "SlicerStyleLight2",
      sortOrder: "ascending",
      rowHeight: 241300,
    });
  });

  it("ignores numeric attributes that fail to parse", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Bogus" cache="C" columnCount="abc" rowHeight=""/>
</slicers>`;
    const [slicer] = parseSlicers(xml);
    expect(slicer.columnCount).toBeUndefined();
    expect(slicer.rowHeight).toBeUndefined();
  });
});

// ── parseSlicerCache ─────────────────────────────────────────────

describe("parseSlicerCache", () => {
  it("returns undefined when the root element is not a slicerCacheDefinition", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<root/>`;
    expect(parseSlicerCache(xml)).toBeUndefined();
  });

  it("returns undefined when the cache has no name", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`;
    expect(parseSlicerCache(xml)).toBeUndefined();
  });

  it("parses pivot-table-sourced caches", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       name="Slicer_Region" sourceName="Region">
  <pivotTables>
    <pivotTable tabId="0" name="PivotTable1"/>
    <pivotTable tabId="1" name="PivotTable2"/>
    <pivotTable name="MissingTab"/>
  </pivotTables>
</slicerCacheDefinition>`;
    const cache = parseSlicerCache(xml);
    expect(cache).toEqual({
      name: "Slicer_Region",
      sourceName: "Region",
      pivotTables: [
        { tabId: 0, name: "PivotTable1" },
        { tabId: 1, name: "PivotTable2" },
      ],
    });
  });

  it("parses table-sourced caches via the x15 extension", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                       name="Slicer_Status" sourceName="Status">
  <extLst>
    <ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}">
      <x15:tableSlicerCache tableId="1" column="Status" name="Status"/>
    </ext>
  </extLst>
</slicerCacheDefinition>`;
    const cache = parseSlicerCache(xml);
    expect(cache?.name).toBe("Slicer_Status");
    expect(cache?.tableSource).toEqual({ name: "Status", column: "Status" });
    expect(cache?.pivotTables).toBeUndefined();
  });
});

// ── parseTimelines ───────────────────────────────────────────────

describe("parseTimelines", () => {
  it("parses required attributes and visibility flags", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="OrderDate" cache="NativeTimeline_OrderDate" caption="Order Date"
            level="months" showHeader="1" showSelectionLabel="1" showTimeLevel="0"
            showHorizontalScrollbar="true" style="TimeSlicerStyleLight2"/>
  <timeline name="Incomplete"/>
</timelines>`;
    const timelines = parseTimelines(xml);
    expect(timelines).toHaveLength(1);
    expect(timelines[0]).toEqual({
      name: "OrderDate",
      cache: "NativeTimeline_OrderDate",
      caption: "Order Date",
      level: "months",
      showHeader: true,
      showSelectionLabel: true,
      showTimeLevel: false,
      showHorizontalScrollbar: true,
      style: "TimeSlicerStyleLight2",
    });
  });
});

// ── parseTimelineCache ───────────────────────────────────────────

describe("parseTimelineCache", () => {
  it("requires a name", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"/>`;
    expect(parseTimelineCache(xml)).toBeUndefined();
  });

  it("parses pivot-table sources", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         name="NativeTimeline_OrderDate" sourceName="OrderDate">
  <pivotTables>
    <pivotTable tabId="0" name="PivotTable1"/>
  </pivotTables>
</timelineCacheDefinition>`;
    const cache = parseTimelineCache(xml);
    expect(cache).toEqual({
      name: "NativeTimeline_OrderDate",
      sourceName: "OrderDate",
      pivotTables: [{ tabId: 0, name: "PivotTable1" }],
    });
  });
});

// ── End-to-end: full XLSX with slicers & timelines ───────────────

/**
 * Build a minimal XLSX containing a worksheet that points at one slicer
 * file and one timeline file, plus the matching workbook-level cache
 * definitions.
 */
async function buildXlsxWithSlicersAndTimelines(): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>
  <Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>
  <Override PartName="/xl/timelines/timeline1.xml" ContentType="application/vnd.ms-excel.timeline+xml"/>
  <Override PartName="/xl/timelineCaches/timelineCache1.xml" ContentType="application/vnd.ms-excel.timelineCache+xml"/>
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
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
  );

  z.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache" Target="timelineCaches/timelineCache1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>1</v></c></row>
  </sheetData>
</worksheet>`),
  );

  z.add(
    "xl/worksheets/_rels/sheet1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/timeline" Target="../timelines/timeline1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/slicers/slicer1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicers xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <slicer name="Region" cache="Slicer_Region" caption="Region" columnCount="1"
          style="SlicerStyleLight1" rowHeight="241300"/>
</slicers>`),
  );

  z.add(
    "xl/slicerCaches/slicerCache1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<slicerCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                       name="Slicer_Region" sourceName="Region">
  <pivotTables>
    <pivotTable tabId="0" name="PivotTable1"/>
  </pivotTables>
</slicerCacheDefinition>`),
  );

  z.add(
    "xl/timelines/timeline1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="OrderDate" cache="NativeTimeline_OrderDate" caption="Order Date" level="months"/>
</timelines>`),
  );

  z.add(
    "xl/timelineCaches/timelineCache1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<timelineCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                         name="NativeTimeline_OrderDate" sourceName="OrderDate">
  <pivotTables>
    <pivotTable tabId="0" name="PivotTable1"/>
  </pivotTables>
</timelineCacheDefinition>`),
  );

  return await z.build();
}

describe("readXlsx — slicer & timeline integration", () => {
  it("attaches workbook.slicerCaches, workbook.timelineCaches, and per-sheet slicers/timelines", async () => {
    const buf = await buildXlsxWithSlicersAndTimelines();
    const wb = await readXlsx(buf);

    expect(wb.slicerCaches).toEqual([
      {
        name: "Slicer_Region",
        sourceName: "Region",
        pivotTables: [{ tabId: 0, name: "PivotTable1" }],
      },
    ]);
    expect(wb.timelineCaches).toEqual([
      {
        name: "NativeTimeline_OrderDate",
        sourceName: "OrderDate",
        pivotTables: [{ tabId: 0, name: "PivotTable1" }],
      },
    ]);

    const sheet = wb.sheets[0];
    expect(sheet.slicers).toHaveLength(1);
    expect(sheet.slicers?.[0]).toEqual({
      name: "Region",
      cache: "Slicer_Region",
      caption: "Region",
      columnCount: 1,
      style: "SlicerStyleLight1",
      rowHeight: 241300,
    });
    expect(sheet.timelines).toHaveLength(1);
    expect(sheet.timelines?.[0]).toEqual({
      name: "OrderDate",
      cache: "NativeTimeline_OrderDate",
      caption: "Order Date",
      level: "months",
    });
  });

  it("preserves slicer & timeline parts and re-emits all references on roundtrip", async () => {
    const buf = await buildXlsxWithSlicersAndTimelines();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);

    const zip = new ZipReader(out);

    // The four bodies must survive byte-for-byte.
    expect(zip.has("xl/slicers/slicer1.xml")).toBe(true);
    expect(zip.has("xl/slicerCaches/slicerCache1.xml")).toBe(true);
    expect(zip.has("xl/timelines/timeline1.xml")).toBe(true);
    expect(zip.has("xl/timelineCaches/timelineCache1.xml")).toBe(true);

    // [Content_Types].xml must declare every part as an Override or
    // Excel will refuse to load them.
    const ct = decoder.decode(await zip.extract("[Content_Types].xml"));
    expect(ct).toContain("/xl/slicers/slicer1.xml");
    expect(ct).toContain("/xl/slicerCaches/slicerCache1.xml");
    expect(ct).toContain("/xl/timelines/timeline1.xml");
    expect(ct).toContain("/xl/timelineCaches/timelineCache1.xml");
    expect(ct).toContain("application/vnd.ms-excel.slicer+xml");
    expect(ct).toContain("application/vnd.ms-excel.slicerCache+xml");
    expect(ct).toContain("application/vnd.ms-excel.timeline+xml");
    expect(ct).toContain("application/vnd.ms-excel.timelineCache+xml");

    // Workbook rels must carry slicerCache and timelineCache rels.
    const wbRels = decoder.decode(await zip.extract("xl/_rels/workbook.xml.rels"));
    expect(wbRels).toContain(
      'Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache"',
    );
    expect(wbRels).toContain('Target="slicerCaches/slicerCache1.xml"');
    expect(wbRels).toContain(
      'Type="http://schemas.microsoft.com/office/2011/relationships/timelineCache"',
    );
    expect(wbRels).toContain('Target="timelineCaches/timelineCache1.xml"');

    // workbook.xml must declare extLst pointing at both caches.
    const wbXml = decoder.decode(await zip.extract("xl/workbook.xml"));
    expect(wbXml).toContain("<extLst>");
    expect(wbXml).toContain("<x14:slicerCaches>");
    expect(wbXml).toContain("<x15:timelineCachePivotCaches>");

    // Sheet rels must declare slicer + timeline relationships.
    const sheetRels = decoder.decode(await zip.extract("xl/worksheets/_rels/sheet1.xml.rels"));
    expect(sheetRels).toContain(
      'Type="http://schemas.microsoft.com/office/2007/relationships/slicer"',
    );
    expect(sheetRels).toContain('Target="../slicers/slicer1.xml"');
    expect(sheetRels).toContain(
      'Type="http://schemas.microsoft.com/office/2011/relationships/timeline"',
    );
    expect(sheetRels).toContain('Target="../timelines/timeline1.xml"');
  });

  it("re-reading the saved workbook returns the same model", async () => {
    const buf = await buildXlsxWithSlicersAndTimelines();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const reread = await readXlsx(out);

    expect(reread.slicerCaches).toHaveLength(1);
    expect(reread.slicerCaches?.[0].name).toBe("Slicer_Region");
    expect(reread.timelineCaches).toHaveLength(1);
    expect(reread.timelineCaches?.[0].name).toBe("NativeTimeline_OrderDate");

    expect(reread.sheets[0].slicers?.[0].name).toBe("Region");
    expect(reread.sheets[0].timelines?.[0].name).toBe("OrderDate");
  });

  it("does not set slicer/timeline fields when the workbook has none", async () => {
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
    expect(wb.slicerCaches).toBeUndefined();
    expect(wb.timelineCaches).toBeUndefined();
    expect(wb.sheets[0].slicers).toBeUndefined();
    expect(wb.sheets[0].timelines).toBeUndefined();
  });
});
