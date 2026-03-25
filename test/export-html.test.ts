import { describe, expect, it } from "vitest";
import { toHtml } from "../src/export/html";
import type { Sheet, Cell, CellStyle, MergeRange } from "../src/_types";

/** Helper to create a minimal sheet */
function makeSheet(rows: Sheet["rows"], overrides?: Partial<Sheet>): Sheet {
  return {
    name: "Sheet1",
    rows,
    ...overrides,
  };
}

describe("toHtml", () => {
  it("basic table structure (table, tbody, tr, td)", () => {
    const sheet = makeSheet([
      ["A", "B"],
      ["C", "D"],
    ]);
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain("<table>");
    expect(html).toContain("</table>");
    expect(html).toContain("<tbody>");
    expect(html).toContain("</tbody>");
    expect(html).toContain("<tr>");
    expect(html).toContain("<td>");
    expect(html).toContain("A");
    expect(html).toContain("D");
    // Should have 2 <tr> elements
    expect(html.match(/<tr>/g)?.length).toBe(2);
    // Should have 4 <td> elements
    expect(html.match(/<td/g)?.length).toBe(4);
  });

  it("header row (thead, th)", () => {
    const sheet = makeSheet([
      ["Name", "Value"],
      ["foo", 42],
    ]);
    const html = toHtml(sheet, { headerRow: true, classes: false });
    expect(html).toContain("<thead>");
    expect(html).toContain("</thead>");
    expect(html).toContain('<th scope="col">Name</th>');
    expect(html).toContain('<th scope="col">Value</th>');
    expect(html).toContain("<tbody>");
    expect(html).toContain("<td>foo</td>");
    expect(html).toContain("<td>42</td>");
    // Only 1 tr in tbody (data row)
    const tbodyContent = html.slice(html.indexOf("<tbody>"), html.indexOf("</tbody>"));
    expect(tbodyContent.match(/<tr>/g)?.length).toBe(1);
  });

  it("cell type classes (num, bool, date, null)", () => {
    const d = new Date(Date.UTC(2024, 0, 15));
    const sheet = makeSheet([[42, true, d, null]]);
    const html = toHtml(sheet);
    expect(html).toContain('class="hucre-num"');
    expect(html).toContain('class="hucre-bool"');
    expect(html).toContain('class="hucre-date"');
    expect(html).toContain('class="hucre-null"');
  });

  it('HTML escaping (< > & " in cell values)', () => {
    const sheet = makeSheet([['<script>alert("xss")</script>', "A & B", "x > y", 'say "hello"']]);
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain("&lt;script&gt;");
    expect(html).toContain("&amp; B");
    expect(html).toContain("x &gt; y");
    expect(html).toContain("say &quot;hello&quot;");
    // Must NOT contain raw < or > from cell content
    expect(html).not.toContain("<script>");
  });

  it("inline styles from CellStyle (bold, color, background, alignment)", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "Bold",
      type: "string",
      style: {
        font: { bold: true, color: { rgb: "FF0000" } },
        fill: { type: "pattern", pattern: "solid", fgColor: { rgb: "FFFF00" } },
        alignment: { horizontal: "center" },
      },
    });
    const sheet = makeSheet([["Bold"]], { cells });
    const html = toHtml(sheet, { styles: true, classes: false });
    expect(html).toContain("font-weight:bold");
    expect(html).toContain("color:#FF0000");
    expect(html).toContain("background-color:#FFFF00");
    expect(html).toContain("text-align:center");
  });

  it("inline styles: border", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "Bordered",
      type: "string",
      style: {
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thick", color: { rgb: "FF0000" } },
        },
      },
    });
    const sheet = makeSheet([["Bordered"]], { cells });
    const html = toHtml(sheet, { styles: true, classes: false });
    expect(html).toContain("border-top:1px solid #000000");
    expect(html).toContain("border-bottom:3px solid #FF0000");
  });

  it("inline styles: italic, underline, strikethrough, font-size, font-family", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "Styled",
      type: "string",
      style: {
        font: {
          italic: true,
          underline: true,
          size: 14,
          name: "Arial",
        },
      },
    });
    const sheet = makeSheet([["Styled"]], { cells });
    const html = toHtml(sheet, { styles: true, classes: false });
    expect(html).toContain("font-style:italic");
    expect(html).toContain("text-decoration:underline");
    expect(html).toContain("font-size:14pt");
    expect(html).toContain("font-family:Arial");
  });

  it("merged cells with colspan", () => {
    const merges: MergeRange[] = [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }];
    const sheet = makeSheet(
      [
        ["Merged", "skip", "skip"],
        ["A", "B", "C"],
      ],
      { merges },
    );
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain('colspan="3"');
    // The merged cell should appear once, the hidden cells should not produce <td>
    const firstRow = html.slice(html.indexOf("<tr>"), html.indexOf("</tr>") + 5);
    expect(firstRow.match(/<td/g)?.length).toBe(1);
  });

  it("merged cells with rowspan", () => {
    const merges: MergeRange[] = [{ startRow: 0, startCol: 0, endRow: 1, endCol: 0 }];
    const sheet = makeSheet(
      [
        ["Merged", "B1"],
        ["skip", "B2"],
      ],
      { merges },
    );
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain('rowspan="2"');
  });

  it("merged cells with both colspan and rowspan", () => {
    const merges: MergeRange[] = [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }];
    const sheet = makeSheet(
      [
        ["Merged", "skip", "C1"],
        ["skip", "skip", "C2"],
      ],
      { merges },
    );
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain('colspan="2"');
    expect(html).toContain('rowspan="2"');
  });

  it("date formatting as ISO date string", () => {
    const d = new Date(Date.UTC(2024, 5, 15));
    const sheet = makeSheet([[d]]);
    const html = toHtml(sheet);
    expect(html).toContain("2024-06-15");
  });

  it("null cells produce empty td", () => {
    const sheet = makeSheet([[null, "hello", null]]);
    const html = toHtml(sheet, { classes: false });
    // Find td elements — null cells should have empty content
    const tds = html.match(/<td[^>]*>(.*?)<\/td>/g) || [];
    expect(tds.length).toBe(3);
    expect(tds[0]).toContain("></td>"); // empty content (may have class attr)
    expect(tds[1]).toContain("hello");
    expect(tds[2]).toContain("></td>");
  });

  it("null cells get hucre-null class when classes enabled", () => {
    const sheet = makeSheet([[null]]);
    const html = toHtml(sheet, { classes: true });
    expect(html).toContain('class="hucre-null"');
  });

  it("empty sheet", () => {
    const sheet = makeSheet([]);
    const html = toHtml(sheet);
    expect(html).toBe("<table></table>");
  });

  it("includeStyleTag option", () => {
    const sheet = makeSheet([["A"]]);
    const html = toHtml(sheet, { includeStyleTag: true });
    expect(html).toContain("<style>");
    expect(html).toContain("border-collapse:collapse");
    expect(html).toContain("</style>");
    expect(html).toContain('<table class="hucre-table">');
    expect(html).toContain("prefers-color-scheme:dark");
  });

  it("includeStyleTag with empty sheet", () => {
    const sheet = makeSheet([]);
    const html = toHtml(sheet, { includeStyleTag: true });
    expect(html).toContain("<style>");
    expect(html).toContain('<table class="hucre-table"></table>');
  });

  it("custom classPrefix", () => {
    const sheet = makeSheet([[42, null]]);
    const html = toHtml(sheet, { classPrefix: "sp" });
    expect(html).toContain('class="sp-num"');
    expect(html).toContain('class="sp-null"');
    expect(html).not.toContain("hucre");
  });

  it("includeStyleTag uses custom classPrefix", () => {
    const sheet = makeSheet([[42]]);
    const html = toHtml(sheet, { includeStyleTag: true, classPrefix: "sp" });
    expect(html).toContain(".sp-num");
    expect(html).toContain(".sp-table");
    expect(html).toContain('class="sp-table"');
  });

  it("dark mode CSS has correct structure", () => {
    const sheet = makeSheet([
      ["Name", "Price"],
      ["Widget", 9.99],
    ]);
    const html = toHtml(sheet, { includeStyleTag: true, headerRow: true });
    // Light mode styles
    expect(html).toContain("color:#1a1a1a");
    expect(html).toContain("background:#fff");
    expect(html).toContain("border:1px solid #e0e0e0");
    expect(html).toContain("background:#f5f5f5");
    // Dark mode media query
    expect(html).toContain("@media(prefers-color-scheme:dark)");
    expect(html).toContain("color:#e0e0e0");
    expect(html).toContain("background:#1a1a1a");
    expect(html).toContain("border-color:#333");
    // Hover
    expect(html).toContain("tr:hover td");
  });

  it("without includeStyleTag no dark/light CSS generated", () => {
    const sheet = makeSheet([["A"]]);
    const html = toHtml(sheet, { includeStyleTag: false });
    expect(html).not.toContain("<style>");
    expect(html).not.toContain("prefers-color-scheme");
    expect(html).not.toContain("hucre-table");
  });

  it("no classes option (classes: false)", () => {
    const sheet = makeSheet([[42, true, null]]);
    const html = toHtml(sheet, { classes: false });
    expect(html).not.toContain("class=");
  });

  it("large sheet (100 rows)", () => {
    const rows: Sheet["rows"] = [];
    for (let r = 0; r < 100; r++) {
      rows.push([`Row ${r}`, r, r % 2 === 0]);
    }
    const sheet = makeSheet(rows);
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain("Row 0");
    expect(html).toContain("Row 99");
    expect(html.match(/<tr>/g)?.length).toBe(100);
  });

  it("boolean values rendered as 'true'/'false'", () => {
    const sheet = makeSheet([[true, false]]);
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain(">true<");
    expect(html).toContain(">false<");
  });

  it("number values rendered correctly", () => {
    const sheet = makeSheet([[0, -1, 3.14, 1000000]]);
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain(">0<");
    expect(html).toContain(">-1<");
    expect(html).toContain(">3.14<");
    expect(html).toContain(">1000000<");
  });

  it("styles: false does not add style attribute", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "Styled",
      type: "string",
      style: { font: { bold: true } },
    });
    const sheet = makeSheet([["Styled"]], { cells });
    const html = toHtml(sheet, { styles: false, classes: false });
    expect(html).not.toContain("style=");
  });

  it("strings get no special class", () => {
    const sheet = makeSheet([["hello"]]);
    const html = toHtml(sheet, { classes: true });
    // The td should not have any class
    expect(html).toContain("<td>hello</td>");
  });

  it("single-cell merge does not produce colspan/rowspan", () => {
    const merges: MergeRange[] = [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }];
    const sheet = makeSheet([["A"]], { merges });
    const html = toHtml(sheet, { classes: false });
    expect(html).not.toContain("colspan");
    expect(html).not.toContain("rowspan");
  });

  it("header row with classes and styles combined", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "Name",
      type: "string",
      style: { font: { bold: true } },
    });
    const sheet = makeSheet(
      [
        ["Name", "Age"],
        ["Alice", 30],
      ],
      { cells },
    );
    const html = toHtml(sheet, { headerRow: true, styles: true, classes: true });
    expect(html).toContain("<th");
    expect(html).toContain("font-weight:bold");
    expect(html).toContain('class="hucre-num"');
  });

  it("dashed and dotted border styles", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "X",
      type: "string",
      style: {
        border: {
          left: { style: "dashed", color: { rgb: "00FF00" } },
          right: { style: "dotted" },
        },
      },
    });
    const sheet = makeSheet([["X"]], { cells });
    const html = toHtml(sheet, { styles: true, classes: false });
    expect(html).toContain("border-left:1px dashed #00FF00");
    expect(html).toContain("border-right:1px dotted");
  });

  it("medium border width", () => {
    const cells = new Map<string, Cell>();
    cells.set("0,0", {
      value: "X",
      type: "string",
      style: {
        border: {
          top: { style: "medium", color: { rgb: "000000" } },
        },
      },
    });
    const sheet = makeSheet([["X"]], { cells });
    const html = toHtml(sheet, { styles: true, classes: false });
    expect(html).toContain("border-top:2px solid #000000");
  });

  it("apostrophe in cell value is escaped", () => {
    const sheet = makeSheet([["it's a test"]]);
    const html = toHtml(sheet, { classes: false });
    expect(html).toContain("it&#39;s a test");
  });
});
