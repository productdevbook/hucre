import { describe, it, expect } from "vitest";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";
import {
  parseCellImages,
  assembleCellImages,
  REL_CELL_IMAGES,
} from "../src/xlsx/cell-images-reader";
import { writeContentTypes } from "../src/xlsx/content-types-writer";
import { writeWorkbookRels } from "../src/xlsx/workbook-writer";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");

// ── Test fixtures ────────────────────────────────────────────────────

/** Minimal valid PNG-like bytes (signature + arbitrary tail). */
function fakePng(seed: number, size = 64): Uint8Array {
  const data = new Uint8Array(size);
  data[0] = 0x89;
  data[1] = 0x50;
  data[2] = 0x4e;
  data[3] = 0x47;
  data[4] = 0x0d;
  data[5] = 0x0a;
  data[6] = 0x1a;
  data[7] = 0x0a;
  for (let i = 8; i < size; i++) data[i] = (i + seed) % 256;
  return data;
}

function fakeJpeg(seed: number, size = 64): Uint8Array {
  const data = new Uint8Array(size);
  data[0] = 0xff;
  data[1] = 0xd8;
  data[2] = 0xff;
  data[3] = 0xe0;
  for (let i = 4; i < size; i++) data[i] = (i + seed) % 256;
  return data;
}

const SAMPLE_CELL_IMAGES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="ID_FIRST" descr="A small swatch"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
  </etc:cellImage>
  <etc:cellImage>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="ID_SECOND"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId2"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
  </etc:cellImage>
</etc:cellImages>`;

const SAMPLE_CELL_IMAGES_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpeg"/>
</Relationships>`;

/**
 * Build a minimal but valid XLSX whose workbook declares a WPS-style
 * `xl/cellimages.xml` with two image entries. The cell formulas
 * `=_xlfn.DISPIMG("ID_FIRST", 1)` are not required for hucre to
 * resolve the binaries — that lookup happens in the application —
 * but we include one so the file looks like real WPS output.
 */
async function buildXlsxWithCellImages(opts?: {
  /** Override the second image bytes — used to verify dedupe doesn't merge. */
  secondPng?: Uint8Array;
}): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
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
  <sheets><sheet name="Catalog" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
  );

  z.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_CELL_IMAGES}" Target="cellimages.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="str"><f>_xlfn.DISPIMG("ID_FIRST",1)</f><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>`),
  );

  z.add("xl/cellimages.xml", encoder.encode(SAMPLE_CELL_IMAGES_XML));
  z.add("xl/_rels/cellimages.xml.rels", encoder.encode(SAMPLE_CELL_IMAGES_RELS));

  z.add("xl/media/image1.png", fakePng(1));
  z.add("xl/media/image2.jpeg", opts?.secondPng ?? fakeJpeg(2));

  return z.build();
}

// ── parseCellImages ─────────────────────────────────────────────────

describe("parseCellImages", () => {
  it("returns one ParsedCellImageRef per <etc:cellImage> with required attrs", () => {
    const refs = parseCellImages(SAMPLE_CELL_IMAGES_XML);
    expect(refs).toHaveLength(2);
    expect(refs[0]).toEqual({
      id: "ID_FIRST",
      embedRId: "rId1",
      description: "A small swatch",
    });
    // Optional descr is omitted entirely when absent — not empty string.
    expect(refs[1]).toEqual({ id: "ID_SECOND", embedRId: "rId2" });
  });

  it("skips entries missing <xdr:pic>", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage/>
  <etc:cellImage>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1" name="ID_OK"/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
    </xdr:pic>
  </etc:cellImage>
</etc:cellImages>`;
    const refs = parseCellImages(xml);
    expect(refs).toHaveLength(1);
    expect(refs[0].id).toBe("ID_OK");
  });

  it("skips entries missing the DISPIMG name or the embed rId", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="1"/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
    </xdr:pic>
  </etc:cellImage>
  <etc:cellImage>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="ID_NO_BLIP"/></xdr:nvPicPr>
      <xdr:blipFill/>
    </xdr:pic>
  </etc:cellImage>
</etc:cellImages>`;
    expect(parseCellImages(xml)).toHaveLength(0);
  });
});

// ── assembleCellImages ─────────────────────────────────────────────

describe("assembleCellImages", () => {
  it("merges parsed refs with a media map and preserves the description", () => {
    const refs = parseCellImages(SAMPLE_CELL_IMAGES_XML);
    const data1 = fakePng(1);
    const data2 = fakeJpeg(2);
    const media = new Map([
      ["rId1", { data: data1, type: "png" as const }],
      ["rId2", { data: data2, type: "jpeg" as const }],
    ]);
    const out = assembleCellImages(refs, media);
    expect(out).toEqual([
      { id: "ID_FIRST", data: data1, type: "png", description: "A small swatch" },
      { id: "ID_SECOND", data: data2, type: "jpeg" },
    ]);
  });

  it("drops refs whose embed rId has no media entry", () => {
    const refs = parseCellImages(SAMPLE_CELL_IMAGES_XML);
    const media = new Map([["rId1", { data: fakePng(1), type: "png" as const }]]);
    const out = assembleCellImages(refs, media);
    expect(out.map((c) => c.id)).toEqual(["ID_FIRST"]);
  });

  it("dedupes by id (first occurrence wins)", () => {
    const refs = [
      { id: "DUP", embedRId: "rId1" },
      { id: "DUP", embedRId: "rId2" },
    ];
    const data1 = fakePng(1);
    const data2 = fakePng(99);
    const media = new Map([
      ["rId1", { data: data1, type: "png" as const }],
      ["rId2", { data: data2, type: "png" as const }],
    ]);
    const out = assembleCellImages(refs, media);
    expect(out).toHaveLength(1);
    expect(out[0].data).toBe(data1);
  });
});

// ── readXlsx integration ────────────────────────────────────────────

describe("readXlsx — cellimages integration", () => {
  it("attaches workbook.cellImages with binaries and types resolved", async () => {
    const buf = await buildXlsxWithCellImages();
    const wb = await readXlsx(buf);
    expect(wb.cellImages).toHaveLength(2);

    const first = wb.cellImages?.find((c) => c.id === "ID_FIRST");
    expect(first?.type).toBe("png");
    expect(first?.description).toBe("A small swatch");
    expect(first?.data.byteLength).toBeGreaterThan(8);
    expect(first?.data[0]).toBe(0x89); // PNG signature

    const second = wb.cellImages?.find((c) => c.id === "ID_SECOND");
    expect(second?.type).toBe("jpeg");
    expect(second?.description).toBeUndefined();
    expect(second?.data[0]).toBe(0xff); // JPEG SOI
  });

  it("omits workbook.cellImages when the part is absent", async () => {
    // Build a minimal XLSX without cellimages — just verify the field is absent.
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`),
    );
    const wb = await readXlsx(await z.build());
    expect(wb.cellImages).toBeUndefined();
  });
});

// ── saveXlsx roundtrip ─────────────────────────────────────────────

describe("saveXlsx — cellimages roundtrip", () => {
  it("re-declares cellimages.xml in workbook rels and content types", async () => {
    const buf = await buildXlsxWithCellImages();
    const rt = await openXlsx(buf);
    const out = await saveXlsx(rt);
    const zip = new ZipReader(out);

    // Body parts survive in raw entries.
    expect(zip.has("xl/cellimages.xml")).toBe(true);
    expect(zip.has("xl/_rels/cellimages.xml.rels")).toBe(true);
    expect(zip.has("xl/media/image1.png")).toBe(true);
    expect(zip.has("xl/media/image2.jpeg")).toBe(true);

    // workbook.xml.rels declares the WPS cellimage relationship.
    const wbRels = decoder.decode(await zip.extract("xl/_rels/workbook.xml.rels"));
    expect(wbRels).toContain(REL_CELL_IMAGES);
    expect(wbRels).toContain('Target="cellimages.xml"');

    // [Content_Types].xml carries the override.
    const ct = decoder.decode(await zip.extract("[Content_Types].xml"));
    expect(ct).toContain("/xl/cellimages.xml");
  });

  it("re-reading a saved workbook recovers cellImages with intact bytes", async () => {
    const buf = await buildXlsxWithCellImages();
    const rt = await openXlsx(buf);
    const out = await saveXlsx(rt);
    const reread = await readXlsx(out);
    expect(reread.cellImages).toHaveLength(2);

    const original1 = fakePng(1);
    const got = reread.cellImages?.find((c) => c.id === "ID_FIRST");
    expect(got?.data.byteLength).toBe(original1.byteLength);
    expect(got?.data[0]).toBe(original1[0]);
    expect(got?.data[63]).toBe(original1[63]);
  });

  it("does not declare a cellimages relationship when the part is absent", async () => {
    // Minimal XLSX with no cellimages — saving must not invent the rel.
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
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>`),
    );
    const buf = await z.build();

    const rt = await openXlsx(buf);
    const out = await saveXlsx(rt);
    const zip = new ZipReader(out);

    const wbRels = decoder.decode(await zip.extract("xl/_rels/workbook.xml.rels"));
    expect(wbRels).not.toContain(REL_CELL_IMAGES);
    const ct = decoder.decode(await zip.extract("[Content_Types].xml"));
    expect(ct).not.toContain("/xl/cellimages.xml");
  });
});

// ── content-types & workbook-rels writer unit tests ────────────────

describe("writeContentTypes — hasCellImages", () => {
  it("emits the override only when hasCellImages is true", () => {
    const without = writeContentTypes({ sheetCount: 1, hasSharedStrings: false });
    expect(without).not.toContain("/xl/cellimages.xml");

    const withFlag = writeContentTypes({
      sheetCount: 1,
      hasSharedStrings: false,
      hasCellImages: true,
    });
    expect(withFlag).toContain('PartName="/xl/cellimages.xml"');
    expect(withFlag).toContain(
      'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"',
    );
  });
});

describe("writeWorkbookRels — hasCellImages", () => {
  it("appends the cellimage relationship after sheets/styles/theme/externalLinks", () => {
    const rels = writeWorkbookRels(
      2,
      false,
      false,
      false,
      false,
      [{ rId: "rId4", target: "externalLinks/externalLink1.xml" }],
      true,
    );
    expect(rels).toContain(REL_CELL_IMAGES);
    expect(rels).toContain('Target="cellimages.xml"');
    // Shouldn't collide with the sheet/styles/theme/externalLink rIds.
    // Order is rId1=sheet1, rId2=sheet2, rId3=styles, (no sharedStrings),
    // rId4=theme, rId5=externalLink (caller picks rId4 already used —
    // verify we still emit a unique trailing id for the cellimage).
    const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
    const cellImagesId = ids[ids.length - 1];
    // Whatever id ended up assigned, it must be strictly greater than
    // every other rId — that's how nextRid is bumped past the others.
    for (const id of ids.slice(0, -1)) {
      expect(cellImagesId).toBeGreaterThan(id);
    }
  });

  it("does not emit anything cellimage-related when hasCellImages is false/undefined", () => {
    const rels = writeWorkbookRels(1, false);
    expect(rels).not.toContain(REL_CELL_IMAGES);
  });
});
