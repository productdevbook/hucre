import { describe, it, expect } from "vitest";
import { ZipWriter } from "../src/zip/writer";
import { ZipReader } from "../src/zip/reader";
import { readXlsx } from "../src/xlsx/reader";
import { openXlsx, saveXlsx } from "../src/xlsx/roundtrip";
import { parsePersons, parseThreadedComments } from "../src/xlsx/threaded-comments-reader";

const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8");

// ── parsePersons ────────────────────────────────────────────────────

describe("parsePersons", () => {
  it("parses required attributes (id + displayName)", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="{11111111-1111-1111-1111-111111111111}" displayName="Alice"/>
  <person id="{22222222-2222-2222-2222-222222222222}" displayName="Bob"
          userId="bob@example.com" providerId="AD"/>
</personList>`;
    const persons = parsePersons(xml);
    expect(persons).toHaveLength(2);
    expect(persons[0]).toEqual({
      id: "{11111111-1111-1111-1111-111111111111}",
      displayName: "Alice",
    });
    expect(persons[1]).toEqual({
      id: "{22222222-2222-2222-2222-222222222222}",
      displayName: "Bob",
      userId: "bob@example.com",
      providerId: "AD",
    });
  });

  it("skips entries missing id or displayName", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="No id"/>
  <person id="abc"/>
  <person id="ok" displayName="OK"/>
</personList>`;
    const persons = parsePersons(xml);
    expect(persons).toHaveLength(1);
    expect(persons[0].displayName).toBe("OK");
  });
});

// ── parseThreadedComments ───────────────────────────────────────────

describe("parseThreadedComments", () => {
  it("parses a thread root + one reply with mentions and a done flag", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="B5" dT="2026-04-01T10:00:00Z"
    personId="{p1}" id="{c1}">
    <text>Looks off — should this be in EUR?</text>
    <mentions>
      <mention mentionpersonId="{p2}" mentionId="{m1}" startIndex="0" length="10"/>
    </mentions>
  </threadedComment>
  <threadedComment dT="2026-04-01T10:05:00Z" personId="{p2}" id="{c2}"
    parentId="{c1}" done="1">
    <text>Fixed in v2.</text>
  </threadedComment>
</ThreadedComments>`;
    const comments = parseThreadedComments(xml);
    expect(comments).toHaveLength(2);

    expect(comments[0]).toMatchObject({
      id: "{c1}",
      ref: "B5",
      personId: "{p1}",
      date: "2026-04-01T10:00:00Z",
      text: "Looks off — should this be in EUR?",
    });
    expect(comments[0].mentions).toEqual([
      {
        mentionPersonId: "{p2}",
        mentionId: "{m1}",
        startIndex: 0,
        length: 10,
      },
    ]);

    expect(comments[1]).toMatchObject({
      id: "{c2}",
      personId: "{p2}",
      parentId: "{c1}",
      done: true,
      text: "Fixed in v2.",
    });
    // Replies omit ref — confirm it's actually undefined, not empty string.
    expect(comments[1].ref).toBeUndefined();
  });

  it("treats `done=true` and `done=1` identically and treats other values as not done", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="{p}" id="{a}" done="1"><text>A</text></threadedComment>
  <threadedComment ref="A2" personId="{p}" id="{b}" done="true"><text>B</text></threadedComment>
  <threadedComment ref="A3" personId="{p}" id="{c}" done="0"><text>C</text></threadedComment>
  <threadedComment ref="A4" personId="{p}" id="{d}"><text>D</text></threadedComment>
</ThreadedComments>`;
    const cs = parseThreadedComments(xml);
    expect(cs.map((c) => c.done)).toEqual([true, true, undefined, undefined]);
  });

  it("skips entries missing required GUIDs", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" id="missing-personId"><text>X</text></threadedComment>
  <threadedComment ref="A2" personId="{p}"><text>missing id</text></threadedComment>
  <threadedComment ref="A3" personId="{p}" id="{ok}"><text>OK</text></threadedComment>
</ThreadedComments>`;
    expect(parseThreadedComments(xml)).toHaveLength(1);
  });

  it("returns an empty mentions list as undefined rather than []", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="{p}" id="{c}"><text>plain</text></threadedComment>
</ThreadedComments>`;
    expect(parseThreadedComments(xml)[0].mentions).toBeUndefined();
  });
});

// ── End-to-end fixture ──────────────────────────────────────────────

async function buildXlsxWithThreadedComments(): Promise<Uint8Array> {
  const z = new ZipWriter();

  z.add(
    "[Content_Types].xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/>
  <Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/worksheets/sheet1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData>
</worksheet>`),
  );

  z.add(
    "xl/worksheets/_rels/sheet1.xml.rels",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>
</Relationships>`),
  );

  z.add(
    "xl/persons/person.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person id="{p1}" displayName="Alice" providerId="AD"/>
  <person id="{p2}" displayName="Bob"/>
</personList>`),
  );

  z.add(
    "xl/threadedComments/threadedComment1.xml",
    encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" dT="2026-04-01T10:00:00Z" personId="{p1}" id="{c1}"><text>Hello</text></threadedComment>
  <threadedComment dT="2026-04-01T10:05:00Z" personId="{p2}" id="{c2}" parentId="{c1}"><text>Hi back</text></threadedComment>
</ThreadedComments>`),
  );

  return await z.build();
}

describe("readXlsx — threaded comments integration", () => {
  it("attaches workbook.persons and sheet.threadedComments", async () => {
    const buf = await buildXlsxWithThreadedComments();
    const wb = await readXlsx(buf);
    expect(wb.persons).toEqual([
      { id: "{p1}", displayName: "Alice", providerId: "AD" },
      { id: "{p2}", displayName: "Bob" },
    ]);
    expect(wb.sheets[0].threadedComments).toHaveLength(2);
    expect(wb.sheets[0].threadedComments?.[0].text).toBe("Hello");
    expect(wb.sheets[0].threadedComments?.[1].parentId).toBe("{c1}");
  });
});

describe("saveXlsx — threaded comments roundtrip", () => {
  it("preserves the threadedComments + persons parts and re-declares all references", async () => {
    const buf = await buildXlsxWithThreadedComments();
    const rt = await openXlsx(buf);
    const out = await saveXlsx(rt);
    const zip = new ZipReader(out);

    // Body parts must still be in the ZIP (they live in raw entries).
    expect(zip.has("xl/persons/person.xml")).toBe(true);
    expect(zip.has("xl/threadedComments/threadedComment1.xml")).toBe(true);

    // workbook.xml.rels must declare the person directory.
    const wbRels = decoder.decode(await zip.extract("xl/_rels/workbook.xml.rels"));
    expect(wbRels).toContain("http://schemas.microsoft.com/office/2017/10/relationships/person");
    expect(wbRels).toContain('Target="persons/person.xml"');

    // Per-sheet rels must declare the threadedComments part.
    const sheetRels = decoder.decode(await zip.extract("xl/worksheets/_rels/sheet1.xml.rels"));
    expect(sheetRels).toContain(
      "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
    );
    expect(sheetRels).toContain('Target="../threadedComments/threadedComment1.xml"');

    // Content types must declare overrides for both parts.
    const ct = decoder.decode(await zip.extract("[Content_Types].xml"));
    expect(ct).toContain("/xl/persons/person.xml");
    expect(ct).toContain("/xl/threadedComments/threadedComment1.xml");
    expect(ct).toContain("application/vnd.ms-excel.threadedcomments+xml");
    expect(ct).toContain("application/vnd.ms-excel.person+xml");
  });

  it("re-reading the saved workbook recovers the same threaded structure", async () => {
    const buf = await buildXlsxWithThreadedComments();
    const rt = await openXlsx(buf);
    const out = await saveXlsx(rt);
    const reread = await readXlsx(out);
    expect(reread.persons?.[0].displayName).toBe("Alice");
    expect(reread.sheets[0].threadedComments?.[1].parentId).toBe("{c1}");
  });
});
