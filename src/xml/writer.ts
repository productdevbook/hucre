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

/** Escape text content for safe embedding in XML */
export function xmlEscape(text: string): string {
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
        replacement = xEscape(13)
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
