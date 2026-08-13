// ── Types ─────────────────────────────────────────────────────────

export interface XmlWriterOptions {
  /** XML declaration. Default: true */
  declaration?: boolean
  /** Standalone attribute. Default: "yes" */
  standalone?: string
}

// ── Escaping ──────────────────────────────────────────────────────

/** Encode a code point as the OOXML `_xHHHH_` escape (uppercase, 4 hex digits) */
function xEscape(code: number): string {
  return `_x${code.toString(16).toUpperCase().padStart(4, "0")}_`
}

/**
 * Characters illegal in XML 1.0 (cannot be represented even as numeric
 * character references): the C0 control range except tab (0x09), LF (0x0A)
 * and CR (0x0D). OOXML encodes these as `_xHHHH_` in text content. CR is also
 * encoded so it survives the XML line-ending normalization parsers apply on
 * read (a literal CR in text content is folded to LF), preserving round-trips.
 */
function isIllegalXmlChar(code: number): boolean {
  return (
    (code >= 0x00 && code <= 0x08) ||
    code === 0x0b ||
    code === 0x0c ||
    (code >= 0x0e && code <= 0x1f)
  )
}

/**
 * How a carriage return is spelled in text content.
 *
 * A literal CR does not survive a parse: XML 1.0 §2.11 folds CR and CRLF
 * to LF before the application ever sees them, so carrying one takes an
 * escape. There are two, and they belong to different formats:
 *
 * - `"ooxml"` — `_x000D_`, Excel's convention. Correct for XLSX, where
 *   the reader decodes it back. **Meaningless in any other format**: to
 *   a consumer that does not know the convention it is seven literal
 *   characters, which is what LibreOffice showed when the ODS writer
 *   inherited this spelling.
 * - `"charRef"` — `&#13;`, which is simply XML. A character reference is
 *   not subject to §2.11 (that covers literal CR in the source), so it
 *   survives a conforming parse anywhere.
 *
 * `"charRef"` is the right answer for everything except OOXML, and is
 * only not the default because XLSX is the larger caller and Excel
 * expects its own spelling.
 */
export type CrSpelling = "ooxml" | "charRef"

/** Escape text content for safe embedding in XML */
export function xmlEscape(text: string, cr: CrSpelling = "ooxml"): string {
  let result = ""
  let last = 0

  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i)
    let replacement: string | undefined
    switch (code) {
      case 38: // &
        replacement = "&amp;"
        break
      case 60: // <
        replacement = "&lt;"
        break
      case 62: // >
        replacement = "&gt;"
        break
      case 13: // CR — encode so it round-trips through XML newline normalization
        replacement = cr === "charRef" ? "&#13;" : xEscape(13)
        break
      default:
        // Strip/encode XML-1.0-illegal control chars so output is always
        // well-formed. NOTE: a literal `_xHHHH_` already present in user data
        // is NOT itself re-escaped here. Excel disambiguates by escaping a
        // leading underscore as `_x005F_`; we deliberately skip that to avoid
        // mangling existing data, accepting the rare ambiguity edge.
        if (isIllegalXmlChar(code)) {
          replacement = xEscape(code)
        }
    }
    if (replacement) {
      result += text.slice(last, i) + replacement
      last = i + 1
    }
  }

  if (last === 0) return text
  return result + text.slice(last)
}

/** Escape attribute value for safe embedding in XML */
export function xmlEscapeAttr(text: string): string {
  let result = ""
  let last = 0

  for (let i = 0; i < text.length; i++) {
    let replacement: string | undefined
    switch (text.charCodeAt(i)) {
      case 38: // &
        replacement = "&amp;"
        break
      case 60: // <
        replacement = "&lt;"
        break
      case 62: // >
        replacement = "&gt;"
        break
      case 34: // "
        replacement = "&quot;"
        break
      case 39: // '
        replacement = "&apos;"
        break
      case 9: // tab
        replacement = "&#9;"
        break
      case 10: // newline
        replacement = "&#10;"
        break
      case 13: // carriage return
        replacement = "&#13;"
        break
      default:
        // XML-1.0-illegal control chars cannot be a numeric ref either;
        // encode them via the OOXML `_xHHHH_` convention to stay well-formed.
        if (isIllegalXmlChar(text.charCodeAt(i))) {
          replacement = xEscape(text.charCodeAt(i))
        }
    }
    if (replacement) {
      result += text.slice(last, i) + replacement
      last = i + 1
    }
  }

  if (last === 0) return text
  return result + text.slice(last)
}

// ── Attribute Serialization ───────────────────────────────────────

type AttrValue = string | number | boolean | undefined | null

function serializeAttrs(attrs: Record<string, AttrValue> | undefined): string {
  if (!attrs) return ""

  let result = ""
  const keys = Object.keys(attrs)

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i]
    const val = attrs[key]

    // Skip undefined and null
    if (val === undefined || val === null) continue

    if (typeof val === "boolean") {
      // Boolean attributes: true → "true", false → "false"
      // In XML, all attributes need values
      result += ` ${key}="${val ? "true" : "false"}"`
    } else {
      result += ` ${key}="${xmlEscapeAttr(String(val))}"`
    }
  }

  return result
}

// ── Element Builders ──────────────────────────────────────────────

/** Build a self-closing XML element string */
export function xmlSelfClose(tag: string, attrs?: Record<string, AttrValue>): string {
  return `<${tag}${serializeAttrs(attrs)}/>`
}

/**
 * Build an OOXML `<t>` element, declaring `xml:space="preserve"` when the
 * text has whitespace an XML consumer is entitled to collapse.
 *
 * Every `<t>` in the package has to make this decision — shared strings,
 * inline strings, and each rich-text run — and the check used to be
 * copy-pasted at each site. The inline-string branch of the worksheet
 * writer was the copy that never got it, so `writeXlsx` with
 * `stringMode: "inline"` emitted `<t>  padded  </t>` and Excel trimmed
 * the padding. One function, so there is nothing left to forget.
 */
export function xmlTextElement(value: string): string {
  const escaped = xmlEscape(value)
  const needsPreserve =
    value.length > 0 &&
    (value[0] === " " ||
      value[value.length - 1] === " " ||
      value.includes("\n") ||
      value.includes("\t"))
  return needsPreserve
    ? `<t xml:space="preserve">${escaped}</t>`
    : xmlElement("t", undefined, escaped)
}

/** Build an XML element string with optional children */
export function xmlElement(
  tag: string,
  attrs?: Record<string, AttrValue>,
  children?: string | string[],
): string {
  const attrStr = serializeAttrs(attrs)

  if (children === undefined || children === null) {
    return `<${tag}${attrStr}/>`
  }

  const content = Array.isArray(children) ? children.join("") : children

  if (!content) {
    return `<${tag}${attrStr}/>`
  }

  return `<${tag}${attrStr}>${content}</${tag}>`
}

// ── Document Builders ─────────────────────────────────────────────

/** Generate XML declaration header */
export function xmlDeclaration(options?: XmlWriterOptions): string {
  const standalone = options?.standalone ?? "yes"
  return `<?xml version="1.0" encoding="UTF-8" standalone="${standalone}"?>`
}

/** Build a complete XML document with declaration and root element */
export function xmlDocument(
  rootTag: string,
  attrs?: Record<string, AttrValue>,
  children?: string | string[],
  options?: XmlWriterOptions,
): string {
  const parts: string[] = []

  const includeDecl = options?.declaration !== false
  if (includeDecl) {
    parts.push(xmlDeclaration(options))
  }

  if (children === undefined || children === null) {
    parts.push(xmlSelfClose(rootTag, attrs))
  } else {
    parts.push(xmlElement(rootTag, attrs, children))
  }

  return parts.join("")
}
