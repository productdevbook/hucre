// ── Streaming XML reader ────────────────────────────────────────────
//
// XML was the last format with no streaming reader. `readXml` builds
// every row before returning one, so a feed larger than memory could not
// be walked at all. See #467.
//
// The obstacle was that `src/xml/parser.ts` is push-based: `parseSax`
// runs to completion and calls back, so wrapping it in a generator would
// buffer every row before yielding the first — the opposite of the
// point.
//
// What this does instead is split the work in two, so almost none of it
// is new code:
//
//   1. A scanner walks the source and finds the *span* of each row
//      element. It understands only what it must to know where an
//      element ends — comments, CDATA, processing instructions,
//      doctypes, quoted attributes, self-closing tags, and elements
//      nested inside a row that share its name.
//   2. Each span is handed to `collectRows`, the same function `readXml`
//      uses, so the element model and the flattening are identical *by
//      construction* rather than by a second implementation agreeing.
//
// The scanner's failure mode is loud: a mis-cut span is not well-formed
// XML and the parse throws, rather than quietly yielding a wrong row.

import { readInputToUint8Array } from "../_input"
import type { CellValue, ReadInput, StreamRow } from "../_types"
import { ParseError } from "../errors"
import { collectRows, elementToFlat, splitTag } from "./data-reader"
import type { XmlReadOptions } from "./data-reader"

export interface XmlStreamReadOptions extends Pick<
  XmlReadOptions,
  "rowTag" | "stripNamespaces" | "attrPrefix" | "flatten" | "textKey" | "maxRows"
> {}

/**
 * Read an XML feed a row at a time.
 *
 * ```ts
 * for await (const row of streamXmlRows(bytes, { rowTag: "record" })) {
 *   console.log(row.index, row.values)
 * }
 * ```
 *
 * Peak memory tracks one row rather than the whole document — the source
 * itself is still held, the same as `streamOdsRows`, because a `<row>`
 * cannot be parsed before its closing tag arrives and the format gives
 * no index to seek by.
 *
 * `rowTag` is worth passing. Without it the tag is taken from the first
 * child of the root, which is what a feed of uniform records has;
 * `readXml` instead counts every child and takes the most frequent,
 * which needs the whole document and so is not something a streaming
 * reader can do.
 *
 * **`values` holds only the keys that row had.** `readXml` returns a
 * rectangle: it collects the union of every row's keys and fills the
 * gaps with `null`, so a record with no `<note>` still has `note: null`.
 * Knowing that union means having read the last row, so a streaming
 * reader cannot do it either — a record with no `<note>` yields an
 * object with no `note` key, and an empty `<record/>` yields `{}`.
 *
 * Read `values.note ?? null` rather than `values.note` if you are moving
 * code from `readXml`, and reach for `readXml` when you want the table
 * shape more than the memory.
 */
export async function* streamXmlRows<T = Record<string, CellValue>>(
  input: ReadInput | string,
  options?: XmlStreamReadOptions,
): AsyncGenerator<StreamRow<T>, void, undefined> {
  const xml =
    typeof input === "string"
      ? input
      : new TextDecoder("utf-8").decode(await readInputToUint8Array(input))
  if (xml.trim() === "") return

  const stripNs = options?.stripNamespaces ?? false
  const flatOpts = {
    attrPrefix: options?.attrPrefix ?? "@",
    flatten: options?.flatten ?? true,
    textKey: options?.textKey ?? "#text",
    stripNs,
  }
  const limit = options?.maxRows ?? Infinity

  let rowTag = options?.rowTag
  let index = 0

  for (const span of scanRowSpans(xml, rowTag, stripNs)) {
    if (index >= limit) return
    // The first span settles the tag when the caller did not name one.
    rowTag ??= span.tag

    // Wrapped so the row sits at depth 2, which is where `collectRows`
    // looks — the same position it occupies under the real root.
    const [element] = collectRows(
      `<x>${span.text}</x>`,
      stripNs ? splitTag(span.tag).local : span.tag,
      stripNs,
    )
    if (!element) continue

    const values: Record<string, CellValue> = {}
    elementToFlat(element, flatOpts, "", values)
    yield { index, sheet: 0, values: values as T }
    index++
  }
}

interface RowSpan {
  tag: string
  text: string
}

/**
 * Find the source span of every row element, without building any of it.
 *
 * A generator so the caller pulls: nothing beyond the current row is
 * scanned until the next `next()`.
 */
function* scanRowSpans(
  xml: string,
  rowTag: string | undefined,
  stripNs: boolean,
): Generator<RowSpan> {
  const matches = (tag: string): boolean => {
    if (rowTag === undefined) return true
    return stripNs ? splitTag(tag).local === rowTag : tag === rowTag
  }

  let i = 0
  let depth = 0
  /** Where the current row started, and how deep we were when it did. */
  let rowStart = -1
  let rowTagName = ""
  let rowDepth = -1

  while (i < xml.length) {
    const lt = xml.indexOf("<", i)
    if (lt < 0) break

    // ── The things that are not elements ─────────────────────────────
    if (xml.startsWith("<!--", lt)) {
      const end = xml.indexOf("-->", lt + 4)
      i = end < 0 ? xml.length : end + 3
      continue
    }
    if (xml.startsWith("<![CDATA[", lt)) {
      const end = xml.indexOf("]]>", lt + 9)
      i = end < 0 ? xml.length : end + 3
      continue
    }
    if (xml.startsWith("<?", lt)) {
      const end = xml.indexOf("?>", lt + 2)
      i = end < 0 ? xml.length : end + 2
      continue
    }
    if (xml.startsWith("<!", lt)) {
      // A doctype, which may carry an internal subset in brackets.
      i = skipDoctype(xml, lt)
      continue
    }

    // ── An element ───────────────────────────────────────────────────
    const closing = xml.charCodeAt(lt + 1) === 47 /* / */
    const tagEnd = findTagEnd(xml, lt)
    if (tagEnd < 0) break

    const selfClosing = !closing && xml.charCodeAt(tagEnd - 1) === 47 /* / */
    const tag = readTagName(xml, lt + (closing ? 2 : 1))

    if (closing) {
      depth--
      if (rowStart >= 0 && depth === rowDepth && tag === rowTagName) {
        yield { tag: rowTagName, text: xml.slice(rowStart, tagEnd + 1) }
        rowStart = -1
      }
    } else if (!selfClosing) {
      // A row opens only at depth 1 — a direct child of the root — and
      // only when we are not already inside one, so an element nested in
      // a row that shares its name does not start a second.
      if (rowStart < 0 && depth === 1 && matches(tag)) {
        rowStart = lt
        rowTagName = tag
        rowDepth = depth
      }
      depth++
    } else if (rowStart < 0 && depth === 1 && matches(tag)) {
      // `<row/>` — a whole row in one tag.
      yield { tag, text: xml.slice(lt, tagEnd + 1) }
    }

    i = tagEnd + 1
  }

  if (rowStart >= 0) {
    throw new ParseError(`Unterminated <${rowTagName}> element in XML input`)
  }
}

/**
 * The index of the `>` that ends the tag starting at `start`.
 *
 * `>` inside a quoted attribute value does not end anything, which is
 * the one place a naive `indexOf(">")` goes wrong on real documents.
 */
function findTagEnd(xml: string, start: number): number {
  let quote = 0
  for (let i = start + 1; i < xml.length; i++) {
    const c = xml.charCodeAt(i)
    if (quote !== 0) {
      if (c === quote) quote = 0
      continue
    }
    if (c === 34 /* " */ || c === 39 /* ' */) {
      quote = c
      continue
    }
    if (c === 62 /* > */) return i
  }
  return -1
}

/** The tag name starting at `start`, up to whitespace or the tag's end. */
function readTagName(xml: string, start: number): string {
  let i = start
  while (i < xml.length) {
    const c = xml.charCodeAt(i)
    if (c === 32 || c === 9 || c === 10 || c === 13 || c === 62 /* > */ || c === 47 /* / */) break
    i++
  }
  return xml.slice(start, i)
}

/**
 * Skip a `<!DOCTYPE …>`, including an internal subset.
 *
 * The subset is bracketed and may itself contain `>`, so the bracket has
 * to be tracked rather than scanning for the first `>`.
 */
function skipDoctype(xml: string, start: number): number {
  let i = start + 2
  let bracket = 0
  let quote = 0
  while (i < xml.length) {
    const c = xml.charCodeAt(i)
    if (quote !== 0) {
      if (c === quote) quote = 0
    } else if (c === 34 || c === 39) {
      quote = c
    } else if (c === 91 /* [ */) {
      bracket++
    } else if (c === 93 /* ] */) {
      bracket--
    } else if (c === 62 /* > */ && bracket <= 0) {
      return i + 1
    }
    i++
  }
  return xml.length
}
