import { describe, expect, it } from "vitest";
import { parseXml, parseSax, decodeOoxmlEscapes } from "../src/xml/parser";
import type { XmlElement } from "../src/xml/parser";

// ── Helper ────────────────────────────────────────────────────────

function firstChild(el: XmlElement): XmlElement {
  for (const c of el.children) {
    if (typeof c !== "string") return c;
  }
  throw new Error("No element child found");
}

function elementChildren(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => typeof c !== "string");
}

// ── Basic Parsing ─────────────────────────────────────────────────

describe("parseXml — basic", () => {
  it("parses a simple self-closing element", () => {
    const el = parseXml("<root/>");
    expect(el.tag).toBe("root");
    expect(el.local).toBe("root");
    expect(el.prefix).toBe("");
    expect(el.children).toHaveLength(0);
  });

  it("parses a self-closing element with space before slash", () => {
    const el = parseXml("<root />");
    expect(el.tag).toBe("root");
    expect(el.children).toHaveLength(0);
  });

  it("parses element with attributes", () => {
    const el = parseXml('<root attr="value" count="42"/>');
    expect(el.attrs.attr).toBe("value");
    expect(el.attrs.count).toBe("42");
  });

  it("parses element with single-quoted attributes", () => {
    const el = parseXml("<root attr='hello world'/>");
    expect(el.attrs.attr).toBe("hello world");
  });

  it("parses an open/close tag pair", () => {
    const el = parseXml("<root></root>");
    expect(el.tag).toBe("root");
    expect(el.children).toHaveLength(0);
  });

  it("parses text content", () => {
    const el = parseXml("<root>Hello World</root>");
    expect(el.text).toBe("Hello World");
    expect(el.children).toHaveLength(1);
    expect(el.children[0]).toBe("Hello World");
  });

  it("parses nested elements", () => {
    const el = parseXml("<root><child><grandchild/></child></root>");
    const child = firstChild(el);
    expect(child.tag).toBe("child");
    const grandchild = firstChild(child);
    expect(grandchild.tag).toBe("grandchild");
  });

  it("parses multiple children", () => {
    const el = parseXml("<root><a/><b/><c/></root>");
    const kids = elementChildren(el);
    expect(kids).toHaveLength(3);
    expect(kids[0].tag).toBe("a");
    expect(kids[1].tag).toBe("b");
    expect(kids[2].tag).toBe("c");
  });

  it("parses mixed content (text + elements)", () => {
    const el = parseXml("<root>before<child/>after</root>");
    expect(el.children).toHaveLength(3);
    expect(el.children[0]).toBe("before");
    expect((el.children[1] as XmlElement).tag).toBe("child");
    expect(el.children[2]).toBe("after");
    expect(el.text).toBe("beforeafter");
  });

  it("handles empty text nodes (whitespace only)", () => {
    const el = parseXml("<root>  \n  </root>");
    expect(el.text).toBe("  \n  ");
  });
});

// ── Namespaces ────────────────────────────────────────────────────

describe("parseXml — namespaces", () => {
  it("parses namespaced tags", () => {
    const el = parseXml('<x:row xmlns:x="http://example.com"/>');
    expect(el.tag).toBe("x:row");
    expect(el.local).toBe("row");
    expect(el.prefix).toBe("x");
  });

  it("parses deeply namespaced tags", () => {
    const el = parseXml('<spreadsheetml:worksheet xmlns:spreadsheetml="http://example.com"/>');
    expect(el.tag).toBe("spreadsheetml:worksheet");
    expect(el.local).toBe("worksheet");
    expect(el.prefix).toBe("spreadsheetml");
  });

  it("parses namespace declarations as attributes", () => {
    const el = parseXml('<root xmlns="http://default.ns" xmlns:r="http://rel.ns"/>');
    expect(el.attrs.xmlns).toBe("http://default.ns");
    expect(el.attrs["xmlns:r"]).toBe("http://rel.ns");
  });

  it("handles elements with mixed namespaced and non-namespaced children", () => {
    const xml = '<root xmlns:x="http://ex.com"><x:a/><b/><x:c/></root>';
    const el = parseXml(xml);
    const kids = elementChildren(el);
    expect(kids[0].prefix).toBe("x");
    expect(kids[0].local).toBe("a");
    expect(kids[1].prefix).toBe("");
    expect(kids[1].local).toBe("b");
    expect(kids[2].prefix).toBe("x");
    expect(kids[2].local).toBe("c");
  });
});

// ── Entity Decoding ───────────────────────────────────────────────

describe("parseXml — entities", () => {
  it("decodes &amp;", () => {
    const el = parseXml("<r>A &amp; B</r>");
    expect(el.text).toBe("A & B");
  });

  it("decodes &lt; and &gt;", () => {
    const el = parseXml("<r>&lt;tag&gt;</r>");
    expect(el.text).toBe("<tag>");
  });

  it("decodes &quot; and &apos;", () => {
    const el = parseXml("<r>&quot;hello&apos;</r>");
    expect(el.text).toBe("\"hello'");
  });

  it("decodes decimal numeric entities", () => {
    const el = parseXml("<r>&#65;&#66;&#67;</r>");
    expect(el.text).toBe("ABC");
  });

  it("decodes hex numeric entities", () => {
    const el = parseXml("<r>&#x41;&#x42;&#x43;</r>");
    expect(el.text).toBe("ABC");
  });

  it("decodes entities in attribute values", () => {
    const el = parseXml('<r attr="a &amp; b"/>');
    expect(el.attrs.attr).toBe("a & b");
  });

  it("decodes entities in attribute values with special chars", () => {
    const el = parseXml('<r formula="A1&amp;B1&lt;&gt;"/>');
    expect(el.attrs.formula).toBe("A1&B1<>");
  });

  it("preserves unknown entities as-is", () => {
    const el = parseXml("<r>&unknown;</r>");
    expect(el.text).toBe("&unknown;");
  });
});

// ── CDATA ─────────────────────────────────────────────────────────

describe("parseXml — CDATA", () => {
  it("parses CDATA section", () => {
    const el = parseXml("<r><![CDATA[Hello <world> & friends]]></r>");
    expect(el.text).toBe("Hello <world> & friends");
  });

  it("preserves special characters in CDATA", () => {
    const el = parseXml('<r><![CDATA[<tag attr="val"> & ]]></r>');
    expect(el.text).toBe('<tag attr="val"> & ');
  });

  it("handles empty CDATA", () => {
    const el = parseXml("<r><![CDATA[]]></r>");
    expect(el.text).toBe("");
  });

  it("handles CDATA with XML-like content", () => {
    const el = parseXml('<r><![CDATA[<?xml version="1.0"?>]]></r>');
    expect(el.text).toBe('<?xml version="1.0"?>');
  });
});

// ── Comments ──────────────────────────────────────────────────────

describe("parseXml — comments", () => {
  it("skips XML comments", () => {
    const el = parseXml("<root><!-- this is a comment --><child/></root>");
    const kids = elementChildren(el);
    expect(kids).toHaveLength(1);
    expect(kids[0].tag).toBe("child");
  });

  it("skips comments between elements", () => {
    const el = parseXml("<root><a/><!-- comment --><b/></root>");
    const kids = elementChildren(el);
    expect(kids).toHaveLength(2);
    expect(kids[0].tag).toBe("a");
    expect(kids[1].tag).toBe("b");
  });

  it("skips comments with dashes inside", () => {
    const el = parseXml("<root><!-- a-b-c --><child/></root>");
    const kids = elementChildren(el);
    expect(kids).toHaveLength(1);
  });
});

// ── Processing Instructions ───────────────────────────────────────

describe("parseXml — processing instructions", () => {
  it("skips XML declaration", () => {
    const el = parseXml('<?xml version="1.0" encoding="UTF-8"?><root/>');
    expect(el.tag).toBe("root");
  });

  it("skips XML declaration with standalone", () => {
    const el = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>');
    expect(el.tag).toBe("root");
  });

  it("skips other processing instructions", () => {
    const el = parseXml("<?mso-application progid='Excel.Sheet'?><root/>");
    expect(el.tag).toBe("root");
  });
});

// ── Complex OOXML ─────────────────────────────────────────────────

describe("parseXml — OOXML worksheet", () => {
  const worksheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="n"><v>42</v></c>
      <c r="C1" t="b"><v>1</v></c>
      <c r="D1" t="str"><f>A1&amp;B1</f><v>Hello42</v></c>
    </row>
    <row r="2" spans="1:4">
      <c r="A2" t="s"><v>1</v></c>
      <c r="B2"><v>3.14</v></c>
    </row>
  </sheetData>
</worksheet>`;

  it("parses worksheet root element", () => {
    const el = parseXml(worksheetXml);
    expect(el.tag).toBe("worksheet");
    expect(el.local).toBe("worksheet");
    expect(el.attrs.xmlns).toBe("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
  });

  it("finds sheetData with rows", () => {
    const el = parseXml(worksheetXml);
    const sheetData = elementChildren(el).find((c) => c.local === "sheetData")!;
    expect(sheetData).toBeDefined();
    const rows = elementChildren(sheetData);
    expect(rows).toHaveLength(2);
  });

  it("parses row attributes", () => {
    const el = parseXml(worksheetXml);
    const sheetData = elementChildren(el).find((c) => c.local === "sheetData")!;
    const rows = elementChildren(sheetData);
    expect(rows[0].attrs.r).toBe("1");
    expect(rows[1].attrs.r).toBe("2");
    expect(rows[1].attrs.spans).toBe("1:4");
  });

  it("parses cell references and types", () => {
    const el = parseXml(worksheetXml);
    const sheetData = elementChildren(el).find((c) => c.local === "sheetData")!;
    const row1 = elementChildren(sheetData)[0];
    const cells = elementChildren(row1);
    expect(cells).toHaveLength(4);

    expect(cells[0].attrs.r).toBe("A1");
    expect(cells[0].attrs.t).toBe("s");
    expect(cells[1].attrs.r).toBe("B1");
    expect(cells[1].attrs.t).toBe("n");
    expect(cells[2].attrs.t).toBe("b");
    expect(cells[3].attrs.t).toBe("str");
  });

  it("parses cell values", () => {
    const el = parseXml(worksheetXml);
    const sheetData = elementChildren(el).find((c) => c.local === "sheetData")!;
    const row1 = elementChildren(sheetData)[0];
    const cells = elementChildren(row1);

    const v0 = elementChildren(cells[0]).find((c) => c.local === "v")!;
    expect(v0.text).toBe("0");

    const v1 = elementChildren(cells[1]).find((c) => c.local === "v")!;
    expect(v1.text).toBe("42");
  });

  it("parses formula cells", () => {
    const el = parseXml(worksheetXml);
    const sheetData = elementChildren(el).find((c) => c.local === "sheetData")!;
    const row1 = elementChildren(sheetData)[0];
    const cell = elementChildren(row1)[3]; // D1

    const f = elementChildren(cell).find((c) => c.local === "f")!;
    expect(f.text).toBe("A1&B1"); // & decoded from &amp;

    const v = elementChildren(cell).find((c) => c.local === "v")!;
    expect(v.text).toBe("Hello42");
  });
});

describe("parseXml — OOXML styles", () => {
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="2">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FF0000"/>
      <name val="Arial"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFF00"/>
      </patternFill>
    </fill>
  </fills>
</styleSheet>`;

  it("parses styleSheet root", () => {
    const el = parseXml(stylesXml);
    expect(el.tag).toBe("styleSheet");
  });

  it("parses numFmts", () => {
    const el = parseXml(stylesXml);
    const numFmts = elementChildren(el).find((c) => c.local === "numFmts")!;
    expect(numFmts.attrs.count).toBe("1");
    const fmt = elementChildren(numFmts)[0];
    expect(fmt.attrs.numFmtId).toBe("164");
    expect(fmt.attrs.formatCode).toBe("#,##0.00");
  });

  it("parses fonts with nested elements", () => {
    const el = parseXml(stylesXml);
    const fonts = elementChildren(el).find((c) => c.local === "fonts")!;
    expect(fonts.attrs.count).toBe("2");
    const fontList = elementChildren(fonts);
    expect(fontList).toHaveLength(2);

    // First font
    const f0Kids = elementChildren(fontList[0]);
    const sz = f0Kids.find((c) => c.local === "sz")!;
    expect(sz.attrs.val).toBe("11");

    // Second font — has bold self-closing
    const f1Kids = elementChildren(fontList[1]);
    const bold = f1Kids.find((c) => c.local === "b")!;
    expect(bold).toBeDefined();
    expect(bold.children).toHaveLength(0);
  });

  it("parses fills with nested structure", () => {
    const el = parseXml(stylesXml);
    const fills = elementChildren(el).find((c) => c.local === "fills")!;
    const fill = elementChildren(fills)[0];
    const patternFill = elementChildren(fill).find((c) => c.local === "patternFill")!;
    expect(patternFill.attrs.patternType).toBe("solid");
    const fgColor = elementChildren(patternFill).find((c) => c.local === "fgColor")!;
    expect(fgColor.attrs.rgb).toBe("FFFF00");
  });
});

// ── SAX Parser ────────────────────────────────────────────────────

describe("parseSax", () => {
  it("fires onOpenTag and onCloseTag", () => {
    const openTags: string[] = [];
    const closeTags: string[] = [];
    parseSax("<root><child/></root>", {
      onOpenTag: (tag) => openTags.push(tag),
      onCloseTag: (tag) => closeTags.push(tag),
    });
    expect(openTags).toEqual(["root", "child"]);
    expect(closeTags).toEqual(["child", "root"]);
  });

  it("fires onText for text content", () => {
    const texts: string[] = [];
    parseSax("<r>Hello</r>", {
      onText: (t) => texts.push(t),
    });
    expect(texts).toContain("Hello");
  });

  it("fires onCData for CDATA sections", () => {
    const cdata: string[] = [];
    parseSax("<r><![CDATA[test data]]></r>", {
      onCData: (t) => cdata.push(t),
    });
    expect(cdata).toEqual(["test data"]);
  });

  it("provides parsed attributes", () => {
    let captured: Record<string, string> = {};
    parseSax('<r id="123" type="test"/>', {
      onOpenTag: (_tag, attrs) => {
        captured = attrs;
      },
    });
    expect(captured.id).toBe("123");
    expect(captured.type).toBe("test");
  });

  it("handles self-closing tags correctly", () => {
    const openTags: string[] = [];
    const closeTags: string[] = [];
    parseSax('<col min="1" max="1" width="10"/>', {
      onOpenTag: (tag) => openTags.push(tag),
      onCloseTag: (tag) => closeTags.push(tag),
    });
    expect(openTags).toEqual(["col"]);
    expect(closeTags).toEqual(["col"]);
  });

  it("handles namespaced tags", () => {
    const tags: string[] = [];
    parseSax('<x:row r="1"><x:c/></x:row>', {
      onOpenTag: (tag) => tags.push(tag),
    });
    expect(tags).toEqual(["x:row", "x:c"]);
  });

  it("decodes entities in text", () => {
    const texts: string[] = [];
    parseSax("<r>a &amp; b &lt; c</r>", {
      onText: (t) => texts.push(t),
    });
    expect(texts).toContain("a & b < c");
  });
});

// ── Excel _xHHHH_ Escapes ────────────────────────────────────────

describe("decodeOoxmlEscapes", () => {
  it("decodes _x000D_ to carriage return", () => {
    expect(decodeOoxmlEscapes("Hello_x000D_World")).toBe("Hello\rWorld");
  });

  it("decodes _x000A_ to newline", () => {
    expect(decodeOoxmlEscapes("Line1_x000A_Line2")).toBe("Line1\nLine2");
  });

  it("decodes _x0009_ to tab", () => {
    expect(decodeOoxmlEscapes("Col1_x0009_Col2")).toBe("Col1\tCol2");
  });

  it("decodes multiple escapes", () => {
    expect(decodeOoxmlEscapes("A_x000D__x000A_B")).toBe("A\r\nB");
  });

  it("returns original string if no escapes", () => {
    const str = "Hello World";
    expect(decodeOoxmlEscapes(str)).toBe(str);
  });

  it("handles uppercase hex", () => {
    expect(decodeOoxmlEscapes("_x004A_")).toBe("J");
  });

  it("handles lowercase hex", () => {
    expect(decodeOoxmlEscapes("_x004a_")).toBe("J");
  });
});

// ── Whitespace Handling ───────────────────────────────────────────

describe("parseXml — whitespace", () => {
  it("preserves whitespace in text nodes", () => {
    const el = parseXml("<r>  hello  </r>");
    expect(el.text).toBe("  hello  ");
  });

  it("captures whitespace-only text between elements", () => {
    const el = parseXml("<r>\n  <child/>\n</r>");
    // Whitespace text nodes are present
    expect(el.children.length).toBeGreaterThan(1);
  });

  it("handles tabs and newlines in text", () => {
    const el = parseXml("<r>line1\n\tline2</r>");
    expect(el.text).toBe("line1\n\tline2");
  });
});

// ── Error Cases ───────────────────────────────────────────────────

describe("parseXml — errors", () => {
  it("throws on empty document", () => {
    expect(() => parseXml("")).toThrow();
  });

  it("throws on unterminated comment", () => {
    expect(() => parseXml("<root><!-- unclosed")).toThrow(/comment/i);
  });

  it("throws on unterminated CDATA", () => {
    expect(() => parseXml("<root><![CDATA[unclosed")).toThrow(/cdata/i);
  });

  it("throws on unterminated processing instruction", () => {
    expect(() => parseXml("<?xml unclosed")).toThrow(/processing instruction/i);
  });

  it("throws on unterminated opening tag", () => {
    expect(() => parseXml('<root attr="val')).toThrow(/unterminated/i);
  });

  it("throws on unterminated closing tag", () => {
    expect(() => parseXml("<root></root")).toThrow(/unterminated/i);
  });

  it("includes position info in errors", () => {
    try {
      parseXml("<root>\n  <!-- unclosed");
      expect.fail("Should have thrown");
    } catch (e) {
      expect((e as Error).message).toMatch(/line \d+/);
      expect((e as Error).message).toMatch(/column \d+/);
    }
  });
});

// ── Performance ───────────────────────────────────────────────────

describe("parseXml — performance", () => {
  it("handles 10,000+ elements efficiently", () => {
    const rows: string[] = [];
    for (let i = 0; i < 10_000; i++) {
      rows.push(`<row r="${i + 1}"><c r="A${i + 1}" t="n"><v>${i}</v></c></row>`);
    }
    const xml = `<sheetData>${rows.join("")}</sheetData>`;

    const start = performance.now();
    const el = parseXml(xml);
    const elapsed = performance.now() - start;

    const rowEls = elementChildren(el);
    expect(rowEls).toHaveLength(10_000);
    expect(rowEls[0].attrs.r).toBe("1");
    expect(rowEls[9999].attrs.r).toBe("10000");

    // Should parse in a reasonable time (< 1 second)
    expect(elapsed).toBeLessThan(1000);
  });

  it("handles 10,000+ elements via SAX efficiently", () => {
    const rows: string[] = [];
    for (let i = 0; i < 10_000; i++) {
      rows.push(`<row r="${i + 1}"><c r="A${i + 1}" t="n"><v>${i}</v></c></row>`);
    }
    const xml = `<sheetData>${rows.join("")}</sheetData>`;

    let rowCount = 0;
    const start = performance.now();
    parseSax(xml, {
      onOpenTag: (tag) => {
        if (tag === "row") rowCount++;
      },
    });
    const elapsed = performance.now() - start;

    expect(rowCount).toBe(10_000);
    expect(elapsed).toBeLessThan(1000);
  });
});

// ── Edge Cases ────────────────────────────────────────────────────

describe("parseXml — edge cases", () => {
  it("handles DOCTYPE declaration", () => {
    const el = parseXml('<!DOCTYPE html><root attr="val"/>');
    expect(el.tag).toBe("root");
    expect(el.attrs.attr).toBe("val");
  });

  it("handles attributes with > in quoted values via SAX", () => {
    const attrs: Record<string, string> = {};
    parseSax('<r formula="IF(A1>0,1,0)"/>', {
      onOpenTag: (_tag, a) => Object.assign(attrs, a),
    });
    expect(attrs.formula).toBe("IF(A1>0,1,0)");
  });

  it("handles tag names with hyphens and dots", () => {
    const el = parseXml("<my-tag><sub.tag/></my-tag>");
    expect(el.tag).toBe("my-tag");
    const child = firstChild(el);
    expect(child.tag).toBe("sub.tag");
  });

  it("handles XML with BOM", () => {
    const bom = "\uFEFF";
    const el = parseXml(`${bom}<?xml version="1.0"?><root/>`);
    expect(el.tag).toBe("root");
  });

  it("handles deeply nested elements", () => {
    let xml = "";
    const depth = 100;
    for (let i = 0; i < depth; i++) xml += `<level${i}>`;
    for (let i = depth - 1; i >= 0; i--) xml += `</level${i}>`;

    const el = parseXml(xml);
    expect(el.tag).toBe("level0");

    // Walk to the deepest level
    let current = el;
    for (let i = 1; i < depth; i++) {
      current = firstChild(current);
      expect(current.tag).toBe(`level${i}`);
    }
  });

  it("handles OOXML content types", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>`;

    const el = parseXml(xml);
    expect(el.tag).toBe("Types");
    const kids = elementChildren(el);
    expect(kids).toHaveLength(3);
    expect(kids[0].tag).toBe("Default");
    expect(kids[0].attrs.Extension).toBe("rels");
    expect(kids[2].tag).toBe("Override");
    expect(kids[2].attrs.PartName).toBe("/xl/workbook.xml");
  });

  it("handles shared strings XML", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t xml:space="preserve"> </t></si>
</sst>`;

    const el = parseXml(xml);
    expect(el.tag).toBe("sst");
    expect(el.attrs.count).toBe("3");

    const items = elementChildren(el);
    expect(items).toHaveLength(3);

    const t0 = firstChild(items[0]);
    expect(t0.text).toBe("Hello");

    const t2 = firstChild(items[2]);
    expect(t2.attrs["xml:space"]).toBe("preserve");
    expect(t2.text).toBe(" ");
  });

  it("handles relationship XML", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    const el = parseXml(xml);
    expect(el.tag).toBe("Relationships");
    const rels = elementChildren(el);
    expect(rels).toHaveLength(2);
    expect(rels[0].attrs.Id).toBe("rId1");
    expect(rels[0].attrs.Target).toBe("worksheets/sheet1.xml");
  });
});
