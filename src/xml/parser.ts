import { XmlError } from "../errors";

// ── Public Types ──────────────────────────────────────────────────

export interface XmlElement {
  tag: string;
  /** Local name without namespace prefix */
  local: string;
  /** Namespace prefix (e.g. "x" from "x:row") */
  prefix: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text?: string;
}

export type XmlNode = XmlElement | string;

/** SAX-style callbacks */
export interface SaxHandlers {
  onOpenTag?: (tag: string, attrs: Record<string, string>) => void;
  onCloseTag?: (tag: string) => void;
  onText?: (text: string) => void;
  onCData?: (text: string) => void;
}

// ── Constants ─────────────────────────────────────────────────────

const ENTITY_MAP: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

// ── Entity Decoding ───────────────────────────────────────────────

function decodeEntities(text: string): string {
  if (text.indexOf("&") === -1) return text;

  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (match, ref: string) => {
    if (ref.charCodeAt(0) === 35 /* # */) {
      // Numeric entity
      const code =
        ref.charCodeAt(1) === 120 /* x */ || ref.charCodeAt(1) === 88 /* X */
          ? parseInt(ref.slice(2), 16)
          : parseInt(ref.slice(1), 10);
      if (Number.isNaN(code)) return match;
      return String.fromCodePoint(code);
    }
    return ENTITY_MAP[ref] ?? match;
  });
}

/**
 * Decode OOXML `_xHHHH_` escape sequences used by Excel in shared strings.
 * Example: `_x000D_` → `\r`, `_x000A_` → `\n`
 */
export function decodeOoxmlEscapes(text: string): string {
  if (text.indexOf("_x") === -1) return text;

  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_match, hex: string) => {
    return String.fromCharCode(parseInt(hex, 16));
  });
}

// ── Attribute Parsing ─────────────────────────────────────────────

function parseAttrs(raw: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  let i = 0;
  const len = raw.length;

  while (i < len) {
    // Skip whitespace
    while (i < len && isWhitespace(raw.charCodeAt(i))) i++;
    if (i >= len) break;

    // Read attribute name
    const nameStart = i;
    while (i < len && raw.charCodeAt(i) !== 61 /* = */ && !isWhitespace(raw.charCodeAt(i))) i++;
    const name = raw.slice(nameStart, i);
    if (!name) break;

    // Skip whitespace around =
    while (i < len && isWhitespace(raw.charCodeAt(i))) i++;
    if (i >= len || raw.charCodeAt(i) !== 61 /* = */) {
      // Boolean attribute (no value) — store with empty string
      attrs[name] = "";
      continue;
    }
    i++; // skip =
    while (i < len && isWhitespace(raw.charCodeAt(i))) i++;

    if (i >= len) break;

    // Read attribute value
    const quote = raw.charCodeAt(i);
    if (quote === 34 /* " */ || quote === 39 /* ' */) {
      i++; // skip opening quote
      const valStart = i;
      while (i < len && raw.charCodeAt(i) !== quote) i++;
      attrs[name] = decodeEntities(raw.slice(valStart, i));
      i++; // skip closing quote
    } else {
      // Unquoted value (technically not valid XML, but handle gracefully)
      const valStart = i;
      while (i < len && !isWhitespace(raw.charCodeAt(i))) i++;
      attrs[name] = decodeEntities(raw.slice(valStart, i));
    }
  }

  return attrs;
}

// ── Tag Name Splitting ────────────────────────────────────────────

function splitTagName(tag: string): { local: string; prefix: string } {
  const colon = tag.indexOf(":");
  if (colon === -1) return { local: tag, prefix: "" };
  return { prefix: tag.slice(0, colon), local: tag.slice(colon + 1) };
}

// ── Character Helpers ─────────────────────────────────────────────

function isWhitespace(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13;
}

// ── SAX Parser ────────────────────────────────────────────────────

/**
 * SAX-style streaming XML parser. Calls handlers as elements are encountered.
 * No DOM construction — minimal memory footprint.
 */
export function parseSax(xml: string, handlers: SaxHandlers): void {
  // Strip UTF-8 BOM (U+FEFF) if present at the start of the input
  const input = xml.charCodeAt(0) === 0xfeff ? xml.slice(1) : xml;
  const len = input.length;
  let i = 0;

  function error(msg: string): never {
    // Calculate line/column for error reporting
    let line = 1;
    let col = 1;
    for (let j = 0; j < i && j < len; j++) {
      if (input.charCodeAt(j) === 10 /* \n */) {
        line++;
        col = 1;
      } else {
        col++;
      }
    }
    throw new XmlError(`${msg} at line ${line}, column ${col}`);
  }

  while (i < len) {
    if (input.charCodeAt(i) === 60 /* < */) {
      // Possible tag
      const next = i + 1 < len ? input.charCodeAt(i + 1) : 0;

      if (next === 33 /* ! */) {
        // Comment or CDATA
        if (input.slice(i, i + 4) === "<!--") {
          // Comment: skip to -->
          const end = input.indexOf("-->", i + 4);
          if (end === -1) error("Unterminated comment");
          i = end + 3;
          continue;
        }
        if (input.slice(i, i + 9) === "<![CDATA[") {
          // CDATA section
          const end = input.indexOf("]]>", i + 9);
          if (end === -1) error("Unterminated CDATA section");
          const text = input.slice(i + 9, end);
          handlers.onCData?.(text);
          handlers.onText?.(text);
          i = end + 3;
          continue;
        }
        // DOCTYPE or other declaration — skip
        const end = input.indexOf(">", i + 2);
        if (end === -1) error("Unterminated declaration");
        i = end + 1;
        continue;
      }

      if (next === 63 /* ? */) {
        // Processing instruction: <?...?>
        const end = input.indexOf("?>", i + 2);
        if (end === -1) error("Unterminated processing instruction");
        i = end + 2;
        continue;
      }

      if (next === 47 /* / */) {
        // Closing tag: </tagName>
        const end = input.indexOf(">", i + 2);
        if (end === -1) error("Unterminated closing tag");
        const tag = input.slice(i + 2, end).trim();
        handlers.onCloseTag?.(tag);
        i = end + 1;
        continue;
      }

      // Opening tag
      // Find end of tag — need to handle > inside attribute values
      let j = i + 1;
      let inQuote = 0;
      while (j < len) {
        const c = input.charCodeAt(j);
        if (inQuote) {
          if (c === inQuote) inQuote = 0;
        } else if (c === 34 /* " */ || c === 39 /* ' */) {
          inQuote = c;
        } else if (c === 62 /* > */) {
          break;
        }
        j++;
      }
      if (j >= len) error("Unterminated opening tag");

      const selfClosing = input.charCodeAt(j - 1) === 47; /* / */
      const tagContent = input.slice(i + 1, selfClosing ? j - 1 : j);

      // Split tag name from attributes
      let spaceIdx = 0;
      const tcLen = tagContent.length;
      while (spaceIdx < tcLen && !isWhitespace(tagContent.charCodeAt(spaceIdx))) spaceIdx++;
      const tag = tagContent.slice(0, spaceIdx);
      const attrStr = spaceIdx < tcLen ? tagContent.slice(spaceIdx + 1) : "";
      const attrs = attrStr ? parseAttrs(attrStr) : {};

      handlers.onOpenTag?.(tag, attrs);
      if (selfClosing) {
        handlers.onCloseTag?.(tag);
      }

      i = j + 1;
      continue;
    }

    // Text content
    const textStart = i;
    while (i < len && input.charCodeAt(i) !== 60 /* < */) i++;
    const rawText = input.slice(textStart, i);
    if (rawText) {
      const decoded = decodeEntities(rawText);
      handlers.onText?.(decoded);
    }
  }
}

// ── Streaming SAX Parser ─────────────────────────────────────────

/**
 * SAX parser that consumes a ReadableStream<Uint8Array> in chunks.
 * Calls handlers incrementally as XML constructs are completed.
 * Handles chunk boundaries that split tags or text content.
 */
export async function parseSaxStream(
  stream: ReadableStream<Uint8Array>,
  handlers: SaxHandlers,
): Promise<void> {
  const reader = stream.getReader();
  const decoder = new TextDecoder("utf-8");
  let buf = "";
  let bomStripped = false;

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });

    if (!bomStripped) {
      if (buf.charCodeAt(0) === 0xfeff) buf = buf.slice(1);
      bomStripped = true;
    }

    buf = processSaxBuffer(buf, handlers, false);
  }

  // Flush remaining decoder state
  buf += decoder.decode();
  if (buf.length > 0) {
    processSaxBuffer(buf, handlers, true);
  }
}

/**
 * Process as much XML as possible from the buffer.
 * Returns the unprocessed remainder (incomplete tag/text at chunk boundary).
 * When `final` is true, all remaining content must be processable.
 */
function processSaxBuffer(buf: string, handlers: SaxHandlers, final: boolean): string {
  let i = 0;
  const len = buf.length;

  while (i < len) {
    if (buf.charCodeAt(i) === 60 /* < */) {
      const next = i + 1 < len ? buf.charCodeAt(i + 1) : -1;

      if (next === -1 && !final) {
        // Incomplete: just '<' at end of chunk
        return buf.slice(i);
      }

      if (next === 33 /* ! */) {
        // Comment or CDATA
        if (buf.slice(i, i + 4) === "<!--") {
          const end = buf.indexOf("-->", i + 4);
          if (end === -1) {
            return final ? "" : buf.slice(i);
          }
          i = end + 3;
          continue;
        }
        if (buf.slice(i, i + 9) === "<![CDATA[") {
          const end = buf.indexOf("]]>", i + 9);
          if (end === -1) {
            return final ? "" : buf.slice(i);
          }
          const text = buf.slice(i + 9, end);
          handlers.onCData?.(text);
          handlers.onText?.(text);
          i = end + 3;
          continue;
        }
        // Could be incomplete CDATA/comment marker
        if (!final && len - i < 9) {
          return buf.slice(i);
        }
        // DOCTYPE or other declaration — skip
        const end = buf.indexOf(">", i + 2);
        if (end === -1) {
          return final ? "" : buf.slice(i);
        }
        i = end + 1;
        continue;
      }

      if (next === 63 /* ? */) {
        // Processing instruction: <?...?>
        const end = buf.indexOf("?>", i + 2);
        if (end === -1) {
          return final ? "" : buf.slice(i);
        }
        i = end + 2;
        continue;
      }

      if (next === 47 /* / */) {
        // Closing tag: </tagName>
        const end = buf.indexOf(">", i + 2);
        if (end === -1) {
          return final ? "" : buf.slice(i);
        }
        const tag = buf.slice(i + 2, end).trim();
        handlers.onCloseTag?.(tag);
        i = end + 1;
        continue;
      }

      // Opening tag — find end, handling > inside attribute values
      let j = i + 1;
      let inQuote = 0;
      while (j < len) {
        const c = buf.charCodeAt(j);
        if (inQuote) {
          if (c === inQuote) inQuote = 0;
        } else if (c === 34 /* " */ || c === 39 /* ' */) {
          inQuote = c;
        } else if (c === 62 /* > */) {
          break;
        }
        j++;
      }
      if (j >= len) {
        // Tag not complete in this chunk
        return final ? "" : buf.slice(i);
      }

      const selfClosing = buf.charCodeAt(j - 1) === 47; /* / */
      const tagContent = buf.slice(i + 1, selfClosing ? j - 1 : j);

      let spaceIdx = 0;
      const tcLen = tagContent.length;
      while (spaceIdx < tcLen && !isWhitespace(tagContent.charCodeAt(spaceIdx))) spaceIdx++;
      const tag = tagContent.slice(0, spaceIdx);
      const attrStr = spaceIdx < tcLen ? tagContent.slice(spaceIdx + 1) : "";
      const attrs = attrStr ? parseAttrs(attrStr) : {};

      handlers.onOpenTag?.(tag, attrs);
      if (selfClosing) {
        handlers.onCloseTag?.(tag);
      }

      i = j + 1;
      continue;
    }

    // Text content — read until '<' or end of buffer
    const textStart = i;
    while (i < len && buf.charCodeAt(i) !== 60 /* < */) i++;

    if (i >= len && !final) {
      // Text might continue in next chunk — hold it
      return buf.slice(textStart);
    }

    const rawText = buf.slice(textStart, i);
    if (rawText) {
      const decoded = decodeEntities(rawText);
      handlers.onText?.(decoded);
    }
  }

  return "";
}

// ── DOM-style Parser ──────────────────────────────────────────────

/**
 * Parse an XML string into a tree of XmlElement nodes.
 * Uses parseSax internally for a single-pass parse.
 */
export function parseXml(xml: string): XmlElement {
  const root: XmlElement = {
    tag: "",
    local: "",
    prefix: "",
    attrs: {},
    children: [],
  };

  const stack: XmlElement[] = [root];

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      const { local, prefix } = splitTagName(tag);
      const el: XmlElement = { tag, local, prefix, attrs, children: [] };
      stack[stack.length - 1].children.push(el);
      stack.push(el);
    },
    onCloseTag(_tag) {
      if (stack.length <= 1) {
        throw new XmlError("Unexpected closing tag: no matching open tag");
      }
      const el = stack.pop()!;
      // Collect direct text from text-only children
      if (el.children.length > 0) {
        const texts: string[] = [];
        for (const child of el.children) {
          if (typeof child === "string") texts.push(child);
        }
        if (texts.length > 0) {
          el.text = texts.join("");
        }
      }
    },
    onText(text) {
      // Push text as a child node
      stack[stack.length - 1].children.push(text);
    },
  });

  // The root wrapper should have exactly one real child (the document element)
  if (root.children.length === 0) {
    throw new XmlError("Empty document: no root element found");
  }

  // Find the first element child (skip text nodes like whitespace)
  for (const child of root.children) {
    if (typeof child !== "string") return child;
  }

  throw new XmlError("No root element found in document");
}
