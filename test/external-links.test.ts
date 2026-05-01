import { describe, it, expect } from "vitest";
import { parseExternalLink } from "../src/xlsx/external-link-reader";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";

const encoder = new TextEncoder();

// ── parseExternalLink: standalone ──────────────────────────────────

describe("parseExternalLink", () => {
  it("returns an empty link when given a nearly-empty externalLink XML", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1"/>
</externalLink>`;
    const link = parseExternalLink(xml);
    expect(link.target).toBe("");
    expect(link.sheetNames).toEqual([]);
    expect(link.sheetData).toEqual([]);
    expect(link.definedNames).toBeUndefined();
  });

  it("resolves the target path and TargetMode from the rels XML", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1"/>
</externalLink>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
    Target="External.xlsx" TargetMode="External"/>
</Relationships>`;
    const link = parseExternalLink(xml, relsXml);
    expect(link.target).toBe("External.xlsx");
    expect(link.targetMode).toBe("External");
  });

  it("parses sheetNames in declaration order", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames>
      <sheetName val="Summary"/>
      <sheetName val="Data"/>
      <sheetName val="Hidden &amp; Notes"/>
    </sheetNames>
  </externalBook>
</externalLink>`;
    const link = parseExternalLink(xml);
    expect(link.sheetNames).toEqual(["Summary", "Data", "Hidden & Notes"]);
  });

  it("parses cached numeric, string, boolean, error, and inline-string cells", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames><sheetName val="Sheet1"/></sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0">
        <row r="1">
          <cell r="A1" t="n"><v>42.5</v></cell>
          <cell r="B1" t="b"><v>1</v></cell>
          <cell r="C1" t="e"><v>#REF!</v></cell>
          <cell r="D1" t="str"><v>Hello</v></cell>
          <cell r="E1" t="s"><v>3</v></cell>
        </row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>`;
    const link = parseExternalLink(xml);
    expect(link.sheetData).toHaveLength(1);
    expect(link.sheetData[0].sheetId).toBe(0);
    const cells = link.sheetData[0].cells;
    expect(cells).toHaveLength(5);
    expect(cells[0]).toEqual({ ref: "A1", type: "n", value: 42.5 });
    expect(cells[1]).toEqual({ ref: "B1", type: "b", value: true });
    expect(cells[2]).toEqual({ ref: "C1", type: "e", value: "#REF!" });
    expect(cells[3]).toEqual({ ref: "D1", type: "str", value: "Hello" });
    // shared-string indices stay numeric — the resolved string lives in
    // the external workbook, which the reader cannot dereference here.
    expect(cells[4]).toEqual({ ref: "E1", type: "s", value: 3 });
  });

  it("parses defined names with sheetId scoping", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames><sheetName val="S"/></sheetNames>
    <definedNames>
      <definedName name="Total" refersTo="[1]S!$B$10"/>
      <definedName name="LocalTotal" refersTo="[1]S!$B$11" sheetId="0"/>
    </definedNames>
  </externalBook>
</externalLink>`;
    const link = parseExternalLink(xml);
    expect(link.definedNames).toHaveLength(2);
    expect(link.definedNames?.[0]).toEqual({ name: "Total", refersTo: "[1]S!$B$10" });
    expect(link.definedNames?.[1]).toEqual({
      name: "LocalTotal",
      refersTo: "[1]S!$B$11",
      sheetId: 0,
    });
  });

  it("skips cells without an `r=` reference rather than throwing", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <externalBook>
    <sheetNames><sheetName val="S"/></sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0">
        <row r="1">
          <cell t="n"><v>1</v></cell>
          <cell r="A1" t="n"><v>2</v></cell>
        </row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>`;
    const link = parseExternalLink(xml);
    expect(link.sheetData[0].cells).toHaveLength(1);
    expect(link.sheetData[0].cells[0].ref).toBe("A1");
  });

  it("falls back to type=n when an unknown cell type is encountered", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <externalBook>
    <sheetNames><sheetName val="S"/></sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0">
        <row r="1">
          <cell r="A1" t="bogus"><v>9</v></cell>
        </row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>`;
    const link = parseExternalLink(xml);
    expect(link.sheetData[0].cells[0].type).toBe("n");
    expect(link.sheetData[0].cells[0].value).toBe(9);
  });
});

// ── End-to-end: full XLSX with external links ──────────────────────

/**
 * Build a minimal but valid XLSX containing one worksheet and one
 * external workbook reference. Anything not strictly required is
 * stripped down to the bare bones the reader actually inspects.
 */
async function buildXlsxWithExternalLink(): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
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
    <sheet name="Main" sheetId="1" r:id="rId1"/>
  </sheets>
  <externalReferences>
    <externalReference r:id="rId2"/>
  </externalReferences>
</workbook>`),
  );

  z.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="n"><v>10</v></c></row>
  </sheetData>
</worksheet>`),
  );

  z.add(
    "xl/externalLinks/externalLink1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames>
      <sheetName val="Lookup"/>
    </sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0">
        <row r="1">
          <cell r="A1" t="n"><v>123</v></cell>
          <cell r="B1" t="str"><v>label</v></cell>
        </row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>`),
  );

  z.add(
    "xl/externalLinks/_rels/externalLink1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
    Target="../External.xlsx" TargetMode="External"/>
</Relationships>`),
  );

  return await z.build();
}

describe("readXlsx — external link integration", () => {
  it("attaches workbook.externalLinks with parsed target and cached values", async () => {
    const buf = await buildXlsxWithExternalLink();
    const wb = await readXlsx(buf);
    expect(wb.externalLinks).toBeDefined();
    expect(wb.externalLinks).toHaveLength(1);
    const link = wb.externalLinks![0];
    expect(link.target).toBe("../External.xlsx");
    expect(link.targetMode).toBe("External");
    expect(link.sheetNames).toEqual(["Lookup"]);
    expect(link.sheetData[0].cells).toEqual([
      { ref: "A1", type: "n", value: 123 },
      { ref: "B1", type: "str", value: "label" },
    ]);
  });

  it("preserves the externalLink files and re-emits the references on roundtrip", async () => {
    const buf = await buildXlsxWithExternalLink();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);

    // The externalLink body and its rels file must survive byte-for-byte.
    const zip = new ZipReader(out);
    expect(zip.has("xl/externalLinks/externalLink1.xml")).toBe(true);
    expect(zip.has("xl/externalLinks/_rels/externalLink1.xml.rels")).toBe(true);

    // The regenerated workbook.xml must declare the externalReference and
    // workbook.xml.rels must carry the externalLink relationship — without
    // these Excel silently drops the link on next open.
    const wbXml = new TextDecoder("utf-8").decode(await zip.extract("xl/workbook.xml"));
    expect(wbXml).toContain("<externalReferences>");
    expect(wbXml).toMatch(/<externalReference [^>]*r:id="rId\d+"/);

    const wbRels = new TextDecoder("utf-8").decode(await zip.extract("xl/_rels/workbook.xml.rels"));
    expect(wbRels).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"',
    );
    expect(wbRels).toContain('Target="externalLinks/externalLink1.xml"');

    // [Content_Types].xml must declare the externalLink override or
    // Excel will refuse to load the part.
    const ct = new TextDecoder("utf-8").decode(await zip.extract("[Content_Types].xml"));
    expect(ct).toContain("/xl/externalLinks/externalLink1.xml");
    expect(ct).toContain(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
    );
  });

  it("re-reading the saved workbook returns the same external link", async () => {
    const buf = await buildXlsxWithExternalLink();
    const wb = await openXlsx(buf);
    const out = await saveXlsx(wb);
    const reread = await readXlsx(out);
    expect(reread.externalLinks).toHaveLength(1);
    expect(reread.externalLinks?.[0].target).toBe("../External.xlsx");
    expect(reread.externalLinks?.[0].sheetData[0].cells[0]).toEqual({
      ref: "A1",
      type: "n",
      value: 123,
    });
  });

  it("does not set externalLinks when the workbook has none", async () => {
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
    expect(wb.externalLinks).toBeUndefined();
  });
});
