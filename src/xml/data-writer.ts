// ── XML Data Writer ──────────────────────────────────────────────────
// Serialize an array of row objects to XML. Keys starting with `attrPrefix`
// (default `@`) are emitted as XML attributes; everything else becomes a
// child element. Nested dot-paths (e.g. `Pricing.Cost`) are reconstructed
// into a tree.

import type { CellValue } from "../_types"
import { ParseError } from "../errors"

export interface XmlWriteOptions {
  /** Root element tag. Default: "root". */
  rootTag?: string
  /** Per-row element tag. Default: "row". */
  rowTag?: string
  /** Prefix marking a key as an XML attribute. Default: "@". */
  attrPrefix?: string
  /** Mixed-content text key. Default: "#text". */
  textKey?: string
  /** Emit `<?xml version="1.0" encoding="UTF-8"?>` declaration. Default: true. */
  declaration?: boolean
  /** Pretty-print with indentation. Default: false. */
  pretty?: boolean
  /** Indent string when `pretty` is true. Default: "  ". */
  indent?: string
}

// ── What may be an element or attribute name ────────────────────────
//
// XML 1.0 §2.3, the `NameStartChar` and `NameChar` productions, minus
// the colon — which is spelled separately below so a name may carry one
// prefix and not a colon anywhere it likes. That is `QName` from
// Namespaces in XML §4, and it is what an XML consumer will accept.
//
// This used to be `/^[A-Za-z_][\w.-]*…/` — ASCII only. `NameStartChar`
// runs from #xC0, so **every** accented or non-Latin heading was
// refused: `Şehir`, `Größe`, `café`, `名前`. A spreadsheet whose column
// names are not English could not be written to XML at all; it threw.
// The rejected names were valid XML, and the ones the production really
// does forbid — a leading digit, a space, `<` — are still rejected.
//
// The `u` flag is required: #x10000–#xEFFFF is above the BMP, and
// without it the surrogate halves are matched separately and a name made
// of astral characters slips through as two non-matching units.
const NAME_START = "A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF"
const NAME_START_2 = "\\u0370-\\u037D\\u037F-\\u1FFF\\u200C-\\u200D\\u2070-\\u218F\\u2C00-\\u2FEF"
const NAME_START_3 = "\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD\\u{10000}-\\u{EFFFF}"
const START = `[${NAME_START}${NAME_START_2}${NAME_START_3}]`
const REST = `[${NAME_START}${NAME_START_2}${NAME_START_3}\\-.0-9\\u00B7\\u0300-\\u036F\\u203F-\\u2040]`

/** One `NCName`, optionally prefixed by another — i.e. a `QName`. */
const VALID_NAME_RE = new RegExp(`^${START}${REST}*(?::${START}${REST}*)?$`, "u")

function escapeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}

function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;")
}

function valueToString(value: CellValue): string {
  if (value === null || value === undefined) return ""
  // See #364 — an unparseable Date threw a raw RangeError mid-write.
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? "" : value.toISOString()
  return String(value)
}

interface TreeNode {
  attrs: Record<string, string>
  text?: string
  children: Map<string, TreeNode>
}

function makeNode(): TreeNode {
  return { attrs: {}, children: new Map() }
}

/**
 * Serialize an array of flat objects to XML.
 *
 * Keys with the `attrPrefix` (default `@`) become XML attributes on the
 * containing element. Dot-separated keys (e.g. `Pricing.Cost`) reconstruct
 * a nested element tree. The `textKey` (default `#text`) emits text content
 * inside an element that also has attributes or children.
 *
 * Throws {@link ParseError} when a key cannot be serialized as a valid XML
 * element name.
 */
export function writeXml(data: Record<string, CellValue>[], options?: XmlWriteOptions): string {
  const rootTag = options?.rootTag ?? "root"
  const rowTag = options?.rowTag ?? "row"
  const attrPrefix = options?.attrPrefix ?? "@"
  const textKey = options?.textKey ?? "#text"
  const declaration = options?.declaration ?? true
  const pretty = options?.pretty ?? false
  const indent = options?.indent ?? "  "

  validateName(rootTag, "rootTag")
  validateName(rowTag, "rowTag")

  const parts: string[] = []
  if (declaration) {
    parts.push('<?xml version="1.0" encoding="UTF-8"?>')
    if (pretty) parts.push("\n")
  }

  const rowDepth = 1
  const sep = pretty ? "\n" : ""
  const pad = (d: number): string => (pretty ? indent.repeat(d) : "")

  parts.push(`<${rootTag}>`)
  parts.push(sep)

  for (const row of data) {
    const tree = buildTree(row, attrPrefix, textKey)
    parts.push(pad(rowDepth))
    parts.push(renderElement(rowTag, tree, pretty, indent, rowDepth))
    parts.push(sep)
  }

  parts.push(`</${rootTag}>`)
  if (pretty) parts.push("\n")
  return parts.join("")
}

function validateName(name: string, label: string): void {
  if (!VALID_NAME_RE.test(name)) {
    throw new ParseError(`Invalid XML name for ${label}: "${name}"`)
  }
}

function buildTree(row: Record<string, CellValue>, attrPrefix: string, textKey: string): TreeNode {
  const root = makeNode()

  for (const [rawKey, rawVal] of Object.entries(row)) {
    if (rawVal === undefined) continue
    insert(root, rawKey, rawVal, attrPrefix, textKey)
  }

  return root
}

// `json/unflatten.ts` reconstructs a tree from dot-paths too. The two are
// intentionally separate — see the note at the top of that file for why.
function insert(
  node: TreeNode,
  key: string,
  value: CellValue,
  attrPrefix: string,
  textKey: string,
): void {
  const path = key.split(".")
  let current = node

  for (let i = 0; i < path.length; i++) {
    const segment = path[i]!
    const isLast = i === path.length - 1

    if (segment.startsWith(attrPrefix)) {
      const attrName = segment.slice(attrPrefix.length)
      validateName(attrName, `attribute "${segment}"`)
      current.attrs[attrName] = valueToString(value)
      return
    }

    if (segment === textKey) {
      current.text = valueToString(value)
      return
    }

    validateName(segment, `element "${segment}"`)

    if (isLast) {
      let child = current.children.get(segment)
      if (!child) {
        child = makeNode()
        current.children.set(segment, child)
      }
      child.text = valueToString(value)
      return
    }

    let child = current.children.get(segment)
    if (!child) {
      child = makeNode()
      current.children.set(segment, child)
    }
    current = child
  }
}

function renderElement(
  tag: string,
  node: TreeNode,
  pretty: boolean,
  indent: string,
  depth: number,
): string {
  const sep = pretty ? "\n" : ""
  const pad = (d: number): string => (pretty ? indent.repeat(d) : "")

  let attrStr = ""
  for (const [name, val] of Object.entries(node.attrs)) {
    attrStr += ` ${name}="${escapeAttr(val)}"`
  }

  const hasChildren = node.children.size > 0
  const text = node.text ?? ""
  const hasText = text !== ""

  if (!hasChildren && !hasText) {
    return `<${tag}${attrStr}/>`
  }

  if (!hasChildren) {
    return `<${tag}${attrStr}>${escapeText(text)}</${tag}>`
  }

  const inner: string[] = []
  for (const [childTag, childNode] of node.children) {
    inner.push(pad(depth + 1))
    inner.push(renderElement(childTag, childNode, pretty, indent, depth + 1))
    inner.push(sep)
  }

  if (hasText) {
    inner.push(pad(depth + 1))
    inner.push(escapeText(text))
    inner.push(sep)
  }

  return `<${tag}${attrStr}>${sep}${inner.join("")}${pad(depth)}</${tag}>`
}

// ── True Streaming XML Writer ────────────────────────────────────────

const TEXT_ENCODER = /* @__PURE__ */ new TextEncoder()

/**
 * Write an XML document as a byte stream, pulling rows from `rows` only
 * as the consumer reads.
 *
 * XML was the one format with no streaming on either side, which made it
 * the odd one out of five. See #467.
 *
 * ```ts
 * return new Response(writeXmlStream(rowCursor, { rowTag: "record" }), {
 *   headers: { "content-type": "application/xml; charset=utf-8" },
 * })
 * ```
 *
 * Peak memory is independent of the row count: each row is rendered,
 * encoded and enqueued on its own, and nothing is retained. The
 * declaration and the root element are written around them, so the
 * result is the same document {@link writeXml} produces from the same
 * rows — there is a test asserting exactly that.
 *
 * The *reader* is still not streaming: `src/xml/parser.ts` is push-based,
 * so a streaming reader needs a pull-based row scanner rather than a
 * wrapper around what is there. That is its own change.
 */
export function writeXmlStream(
  rows: AsyncIterable<Record<string, CellValue>> | Iterable<Record<string, CellValue>>,
  options?: XmlWriteOptions,
): ReadableStream<Uint8Array> {
  const chunks = xmlStreamChunks(rows, options)

  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      try {
        const { done, value } = await chunks.next()
        if (done) {
          controller.close()
          return
        }
        controller.enqueue(value)
      } catch (err) {
        controller.error(err)
      }
    },
    async cancel(reason) {
      await chunks.return?.(reason)
    },
  })
}

/** Render rows into ~64 KB encoded chunks, pulling lazily. */
async function* xmlStreamChunks(
  rows: AsyncIterable<Record<string, CellValue>> | Iterable<Record<string, CellValue>>,
  options?: XmlWriteOptions,
): AsyncGenerator<Uint8Array> {
  const rootTag = options?.rootTag ?? "root"
  const rowTag = options?.rowTag ?? "row"
  const attrPrefix = options?.attrPrefix ?? "@"
  const textKey = options?.textKey ?? "#text"
  const declaration = options?.declaration ?? true
  const pretty = options?.pretty ?? false
  const indent = options?.indent ?? "  "

  // Validated up front, so a bad tag fails before any bytes go out
  // rather than half way through a response.
  validateName(rootTag, "rootTag")
  validateName(rowTag, "rowTag")

  const sep = pretty ? "\n" : ""
  const pad = pretty ? indent : ""

  const CHUNK_BYTES = 64 * 1024
  let pending: string[] = []
  let pendingBytes = 0

  const push = function* (text: string): Generator<Uint8Array> {
    pending.push(text)
    pendingBytes += text.length
    if (pendingBytes >= CHUNK_BYTES) {
      yield TEXT_ENCODER.encode(pending.join(""))
      pending = []
      pendingBytes = 0
    }
  }

  if (declaration) {
    yield* push('<?xml version="1.0" encoding="UTF-8"?>')
    if (pretty) yield* push("\n")
  }
  yield* push(`<${rootTag}>`)
  yield* push(sep)

  for await (const row of rows) {
    yield* push(pad)
    yield* push(renderElement(rowTag, buildTree(row, attrPrefix, textKey), pretty, indent, 1))
    yield* push(sep)
  }

  yield* push(`</${rootTag}>`)
  if (pretty) yield* push("\n")

  if (pending.length > 0) yield TEXT_ENCODER.encode(pending.join(""))
}
