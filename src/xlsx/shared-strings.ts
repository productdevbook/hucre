// ── Shared Strings Parser ─────────────────────────────────────────────
// Parses xl/sharedStrings.xml — the Shared String Table (SST).

import type { XmlElement } from "../xml/parser";
import type { RichTextRun, FontStyle } from "../_types";
import { parseXml, decodeOoxmlEscapes } from "../xml/parser";

export interface SharedString {
  text: string;
  richText?: RichTextRun[];
}

/**
 * Parse xl/sharedStrings.xml into an array of SharedString entries.
 * Each `<si>` element is one entry in the table.
 */
export function parseSharedStrings(xml: string): SharedString[] {
  const doc = parseXml(xml);
  const strings: SharedString[] = [];

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "si") {
      strings.push(parseSiElement(child));
    }
  }

  return strings;
}

/** Parse a single <si> (String Item) element */
function parseSiElement(si: XmlElement): SharedString {
  // Check if this is a simple string (<t>) or rich text (<r> elements)
  const rElements: XmlElement[] = [];
  let simpleText: string | null = null;

  for (const child of si.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "t") {
      simpleText = extractText(child);
    } else if (local === "r") {
      rElements.push(child);
    }
  }

  // Simple string case
  if (rElements.length === 0) {
    return { text: decodeOoxmlEscapes(simpleText ?? "") };
  }

  // Rich text case — multiple <r> elements
  const richText: RichTextRun[] = [];
  const textParts: string[] = [];

  for (const r of rElements) {
    let runText = "";
    let font: FontStyle | undefined;

    for (const rChild of r.children) {
      if (typeof rChild === "string") continue;
      const rLocal = rChild.local || rChild.tag;

      if (rLocal === "t") {
        runText = extractText(rChild);
      } else if (rLocal === "rPr") {
        font = parseRunProperties(rChild);
      }
    }

    const decodedText = decodeOoxmlEscapes(runText);
    textParts.push(decodedText);
    richText.push(font ? { text: decodedText, font } : { text: decodedText });
  }

  return {
    text: textParts.join(""),
    richText,
  };
}

/** Extract text content from a <t> element, respecting xml:space="preserve" */
function extractText(tElement: XmlElement): string {
  return tElement.text ?? "";
}

/** Parse <rPr> (Run Properties) into a FontStyle */
function parseRunProperties(rPr: XmlElement): FontStyle {
  const font: FontStyle = {};

  for (const child of rPr.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    switch (local) {
      case "b":
        font.bold = child.attrs["val"] !== "0" && child.attrs["val"] !== "false";
        break;
      case "i":
        font.italic = child.attrs["val"] !== "0" && child.attrs["val"] !== "false";
        break;
      case "u": {
        const val = child.attrs["val"];
        if (val === "double") font.underline = "double";
        else if (val === "singleAccounting") font.underline = "singleAccounting";
        else if (val === "doubleAccounting") font.underline = "doubleAccounting";
        else font.underline = true;
        break;
      }
      case "strike":
        font.strikethrough = child.attrs["val"] !== "0" && child.attrs["val"] !== "false";
        break;
      case "sz":
        if (child.attrs["val"]) font.size = Number(child.attrs["val"]);
        break;
      case "rFont":
        if (child.attrs["val"]) font.name = child.attrs["val"];
        break;
      case "color":
        font.color = {};
        if (child.attrs["rgb"]) font.color.rgb = child.attrs["rgb"].replace(/^FF/, "");
        if (child.attrs["theme"]) font.color.theme = Number(child.attrs["theme"]);
        if (child.attrs["tint"]) font.color.tint = Number(child.attrs["tint"]);
        if (child.attrs["indexed"]) font.color.indexed = Number(child.attrs["indexed"]);
        break;
      case "vertAlign":
        if (child.attrs["val"] === "superscript" || child.attrs["val"] === "subscript") {
          font.vertAlign = child.attrs["val"];
        }
        break;
      case "family":
        if (child.attrs["val"]) font.family = Number(child.attrs["val"]);
        break;
      case "charset":
        if (child.attrs["val"]) font.charset = Number(child.attrs["val"]);
        break;
      case "scheme":
        if (
          child.attrs["val"] === "major" ||
          child.attrs["val"] === "minor" ||
          child.attrs["val"] === "none"
        ) {
          font.scheme = child.attrs["val"];
        }
        break;
    }
  }

  return font;
}
