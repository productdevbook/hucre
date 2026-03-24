import { describe, expect, it } from "vitest";
import {
  xmlElement,
  xmlSelfClose,
  xmlEscape,
  xmlEscapeAttr,
  xmlDeclaration,
  xmlDocument,
} from "../src/xml/writer";
import { parseXml } from "../src/xml/parser";
import type { XmlElement } from "../src/xml/parser";

// ── Helpers ───────────────────────────────────────────────────────

function elementChildren(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => typeof c !== "string");
}

// ── xmlEscape ─────────────────────────────────────────────────────

describe("xmlEscape", () => {
  it("escapes ampersand", () => {
    expect(xmlEscape("a & b")).toBe("a &amp; b");
  });

  it("escapes less-than", () => {
    expect(xmlEscape("a < b")).toBe("a &lt; b");
  });

  it("escapes greater-than", () => {
    expect(xmlEscape("a > b")).toBe("a &gt; b");
  });

  it("does not escape quotes in text content", () => {
    expect(xmlEscape('a "b" c')).toBe('a "b" c');
  });

  it("handles multiple special characters", () => {
    expect(xmlEscape("<a & b>")).toBe("&lt;a &amp; b&gt;");
  });

  it("returns same string if no escaping needed", () => {
    const str = "hello world 123";
    expect(xmlEscape(str)).toBe(str);
  });

  it("handles empty string", () => {
    expect(xmlEscape("")).toBe("");
  });
});

// ── xmlEscapeAttr ─────────────────────────────────────────────────

describe("xmlEscapeAttr", () => {
  it("escapes ampersand", () => {
    expect(xmlEscapeAttr("a & b")).toBe("a &amp; b");
  });

  it("escapes less-than", () => {
    expect(xmlEscapeAttr("a < b")).toBe("a &lt; b");
  });

  it("escapes greater-than", () => {
    expect(xmlEscapeAttr("a > b")).toBe("a &gt; b");
  });

  it("escapes double quotes", () => {
    expect(xmlEscapeAttr('a "b" c')).toBe("a &quot;b&quot; c");
  });

  it("escapes tab characters", () => {
    expect(xmlEscapeAttr("a\tb")).toBe("a&#9;b");
  });

  it("escapes newline characters", () => {
    expect(xmlEscapeAttr("a\nb")).toBe("a&#10;b");
  });

  it("escapes carriage return characters", () => {
    expect(xmlEscapeAttr("a\rb")).toBe("a&#13;b");
  });

  it("handles multiple special characters", () => {
    expect(xmlEscapeAttr('<"hello">')).toBe("&lt;&quot;hello&quot;&gt;");
  });

  it("returns same string if no escaping needed", () => {
    const str = "hello world";
    expect(xmlEscapeAttr(str)).toBe(str);
  });
});

// ── xmlSelfClose ──────────────────────────────────────────────────

describe("xmlSelfClose", () => {
  it("builds a simple self-closing tag", () => {
    expect(xmlSelfClose("br")).toBe("<br/>");
  });

  it("builds with string attributes", () => {
    expect(xmlSelfClose("col", { min: "1", max: "3" })).toBe('<col min="1" max="3"/>');
  });

  it("builds with numeric attributes", () => {
    expect(xmlSelfClose("col", { min: 1, max: 3, width: 10.5 })).toBe(
      '<col min="1" max="3" width="10.5"/>',
    );
  });

  it("builds with boolean attributes", () => {
    expect(xmlSelfClose("col", { hidden: true, collapsed: false })).toBe(
      '<col hidden="true" collapsed="false"/>',
    );
  });

  it("skips undefined attributes", () => {
    expect(xmlSelfClose("col", { min: "1", max: undefined })).toBe('<col min="1"/>');
  });

  it("skips null attributes", () => {
    expect(xmlSelfClose("col", { min: "1", max: null as unknown as undefined })).toBe(
      '<col min="1"/>',
    );
  });

  it("escapes attribute values", () => {
    expect(xmlSelfClose("item", { name: 'a "b" & c' })).toBe(
      '<item name="a &quot;b&quot; &amp; c"/>',
    );
  });

  it("handles no attributes", () => {
    expect(xmlSelfClose("br")).toBe("<br/>");
    expect(xmlSelfClose("br", {})).toBe("<br/>");
    expect(xmlSelfClose("br", undefined)).toBe("<br/>");
  });

  it("handles namespace attributes", () => {
    expect(
      xmlSelfClose("root", {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      }),
    ).toBe('<root xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>');
  });
});

// ── xmlElement ────────────────────────────────────────────────────

describe("xmlElement", () => {
  it("builds element with text content", () => {
    expect(xmlElement("v", undefined, "42")).toBe("<v>42</v>");
  });

  it("builds element with attributes and text", () => {
    expect(xmlElement("c", { r: "A1", t: "s" }, "<v>0</v>")).toBe('<c r="A1" t="s"><v>0</v></c>');
  });

  it("builds self-closing when no children", () => {
    expect(xmlElement("empty")).toBe("<empty/>");
    expect(xmlElement("empty", undefined, undefined)).toBe("<empty/>");
    expect(xmlElement("empty", undefined, "")).toBe("<empty/>");
  });

  it("builds element with array of children", () => {
    const children = ["<v>1</v>", "<v>2</v>"];
    expect(xmlElement("row", undefined, children)).toBe("<row><v>1</v><v>2</v></row>");
  });

  it("builds element with attributes and array children", () => {
    const children = [xmlSelfClose("c", { r: "A1" }), xmlSelfClose("c", { r: "B1" })];
    expect(xmlElement("row", { r: "1" }, children)).toBe('<row r="1"><c r="A1"/><c r="B1"/></row>');
  });

  it("handles complex nested structure", () => {
    const cell = xmlElement("c", { r: "A1", t: "s" }, xmlElement("v", undefined, "0"));
    expect(cell).toBe('<c r="A1" t="s"><v>0</v></c>');
  });

  it("escapes text content is caller responsibility", () => {
    // xmlElement does NOT auto-escape children — they're expected to be pre-built XML
    // Use xmlEscape for text content
    const el = xmlElement("t", undefined, xmlEscape("A & B"));
    expect(el).toBe("<t>A &amp; B</t>");
  });
});

// ── xmlDeclaration ────────────────────────────────────────────────

describe("xmlDeclaration", () => {
  it("produces default declaration", () => {
    expect(xmlDeclaration()).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  });

  it("supports custom standalone value", () => {
    expect(xmlDeclaration({ standalone: "no" })).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="no"?>',
    );
  });
});

// ── xmlDocument ───────────────────────────────────────────────────

describe("xmlDocument", () => {
  it("builds a full document with declaration", () => {
    const doc = xmlDocument("root", undefined, xmlElement("child", undefined, "text"));
    expect(doc).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root><child>text</child></root>',
    );
  });

  it("builds self-closing document when no children", () => {
    const doc = xmlDocument("empty");
    expect(doc).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><empty/>');
  });

  it("supports declaration: false", () => {
    const doc = xmlDocument("root", undefined, "content", {
      declaration: false,
    });
    expect(doc).toBe("<root>content</root>");
  });

  it("builds document with namespace attributes", () => {
    const doc = xmlDocument(
      "worksheet",
      {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      },
      xmlElement("sheetData"),
    );
    expect(doc).toContain('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    expect(doc).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    );
  });
});

// ── Complex OOXML Output ──────────────────────────────────────────

describe("writer — complex OOXML", () => {
  it("generates worksheet XML", () => {
    const cells = [
      xmlElement("c", { r: "A1", t: "s" }, xmlElement("v", undefined, "0")),
      xmlElement("c", { r: "B1", t: "n" }, xmlElement("v", undefined, "42")),
      xmlElement("c", { r: "C1", t: "str" }, [
        xmlElement("f", undefined, xmlEscape("A1&B1")),
        xmlElement("v", undefined, xmlEscape("Hello42")),
      ]),
    ];
    const row = xmlElement("row", { r: "1" }, cells);
    const sheetData = xmlElement("sheetData", undefined, row);
    const doc = xmlDocument(
      "worksheet",
      {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      },
      sheetData,
    );

    expect(doc).toContain("<?xml");
    expect(doc).toContain("<worksheet xmlns=");
    expect(doc).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(doc).toContain('<c r="B1" t="n"><v>42</v></c>');
    expect(doc).toContain("<f>A1&amp;B1</f>");
    expect(doc).toContain("</worksheet>");
  });

  it("generates styles XML", () => {
    const numFmt = xmlSelfClose("numFmt", {
      numFmtId: 164,
      formatCode: "#,##0.00",
    });
    const numFmts = xmlElement("numFmts", { count: 1 }, numFmt);

    const font = xmlElement("font", undefined, [
      xmlSelfClose("sz", { val: 11 }),
      xmlSelfClose("color", { theme: 1 }),
      xmlSelfClose("name", { val: "Calibri" }),
    ]);
    const fonts = xmlElement("fonts", { count: 1 }, font);

    const doc = xmlDocument(
      "styleSheet",
      {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      },
      [numFmts, fonts],
    );

    expect(doc).toContain('<numFmt numFmtId="164" formatCode="#,##0.00"/>');
    expect(doc).toContain('<sz val="11"/>');
    expect(doc).toContain('<color theme="1"/>');
    expect(doc).toContain('<name val="Calibri"/>');
  });

  it("generates content types XML", () => {
    const defaults = [
      xmlSelfClose("Default", {
        Extension: "rels",
        ContentType: "application/vnd.openxmlformats-package.relationships+xml",
      }),
      xmlSelfClose("Default", {
        Extension: "xml",
        ContentType: "application/xml",
      }),
    ];
    const overrides = [
      xmlSelfClose("Override", {
        PartName: "/xl/workbook.xml",
        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
      }),
    ];
    const doc = xmlDocument(
      "Types",
      {
        xmlns: "http://schemas.openxmlformats.org/package/2006/content-types",
      },
      [...defaults, ...overrides],
    );

    expect(doc).toContain('<Default Extension="rels"');
    expect(doc).toContain('<Override PartName="/xl/workbook.xml"');
  });

  it("generates shared strings XML", () => {
    const strings = ["Hello", "World", "A & B"];
    const siElements = strings.map((s) =>
      xmlElement("si", undefined, xmlElement("t", undefined, xmlEscape(s))),
    );
    const doc = xmlDocument(
      "sst",
      {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        count: strings.length,
        uniqueCount: strings.length,
      },
      siElements,
    );

    expect(doc).toContain("<t>Hello</t>");
    expect(doc).toContain("<t>World</t>");
    expect(doc).toContain("<t>A &amp; B</t>");
    expect(doc).toContain('count="3"');
  });
});

// ── Round-trip Tests ──────────────────────────────────────────────

describe("round-trip — write then parse", () => {
  it("round-trips a simple element", () => {
    const xml = xmlDocument("root", { id: "1" }, xmlElement("child", undefined, "text"));
    const el = parseXml(xml);
    expect(el.tag).toBe("root");
    expect(el.attrs.id).toBe("1");
    const child = elementChildren(el)[0];
    expect(child.tag).toBe("child");
    expect(child.text).toBe("text");
  });

  it("round-trips escaped text", () => {
    const text = 'A & B < C > D "E"';
    const xml = xmlElement("r", undefined, xmlEscape(text));
    const el = parseXml(xml);
    expect(el.text).toBe(text);
  });

  it("round-trips escaped attribute values", () => {
    const val = 'IF(A1>0,"yes","no")';
    const xml = xmlSelfClose("r", { formula: val });
    const el = parseXml(xml);
    expect(el.attrs.formula).toBe(val);
  });

  it("round-trips a worksheet fragment", () => {
    const cells = [
      xmlElement("c", { r: "A1", t: "s" }, xmlElement("v", undefined, "0")),
      xmlElement("c", { r: "B1", t: "n" }, xmlElement("v", undefined, "42")),
    ];
    const row = xmlElement("row", { r: "1" }, cells);
    const sheetData = xmlElement("sheetData", undefined, row);
    const xml = xmlDocument(
      "worksheet",
      {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      },
      sheetData,
    );

    const ws = parseXml(xml);
    expect(ws.tag).toBe("worksheet");

    const sd = elementChildren(ws).find((c) => c.tag === "sheetData")!;
    const r = elementChildren(sd)[0];
    expect(r.attrs.r).toBe("1");

    const c = elementChildren(r);
    expect(c).toHaveLength(2);
    expect(c[0].attrs.r).toBe("A1");
    expect(c[0].attrs.t).toBe("s");

    const v0 = elementChildren(c[0]).find((x) => x.tag === "v")!;
    expect(v0.text).toBe("0");

    const v1 = elementChildren(c[1]).find((x) => x.tag === "v")!;
    expect(v1.text).toBe("42");
  });

  it("round-trips numeric and boolean attributes", () => {
    const xml = xmlSelfClose("item", { count: 42, visible: true, hidden: false });
    const el = parseXml(xml);
    expect(el.attrs.count).toBe("42");
    expect(el.attrs.visible).toBe("true");
    expect(el.attrs.hidden).toBe("false");
  });

  it("round-trips namespace declarations", () => {
    const xml = xmlDocument(
      "root",
      {
        xmlns: "http://default.ns",
        "xmlns:r": "http://rel.ns",
      },
      xmlSelfClose("child"),
    );

    const el = parseXml(xml);
    expect(el.attrs.xmlns).toBe("http://default.ns");
    expect(el.attrs["xmlns:r"]).toBe("http://rel.ns");
  });

  it("round-trips complex styles", () => {
    const font = xmlElement("font", undefined, [
      xmlSelfClose("b"),
      xmlSelfClose("sz", { val: 14 }),
      xmlSelfClose("color", { rgb: "FF0000" }),
      xmlSelfClose("name", { val: "Arial" }),
    ]);
    const fonts = xmlElement("fonts", { count: 1 }, font);
    const xml = xmlDocument(
      "styleSheet",
      {
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      },
      fonts,
    );

    const ss = parseXml(xml);
    const fontsEl = elementChildren(ss).find((c) => c.tag === "fonts")!;
    expect(fontsEl.attrs.count).toBe("1");

    const fontEl = elementChildren(fontsEl)[0];
    const fontKids = elementChildren(fontEl);
    expect(fontKids.find((c) => c.tag === "b")).toBeDefined();
    expect(fontKids.find((c) => c.tag === "sz")!.attrs.val).toBe("14");
    expect(fontKids.find((c) => c.tag === "color")!.attrs.rgb).toBe("FF0000");
    expect(fontKids.find((c) => c.tag === "name")!.attrs.val).toBe("Arial");
  });
});
