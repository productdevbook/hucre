import { describe, expect, it } from "vitest"
import { parseXml } from "../src/xml/parser"
import { resolveTitleDefRPr, resolveTxPrDefRPr } from "../src/xlsx/chart/text"
import {
  parseAxisLabelBold,
  parseAxisLabelFontSize,
  parseAxisTitleBold,
  parseAxisTitleFontSize,
} from "../src/xlsx/chart/axis"
import type { XmlElement } from "../src/xml/parser"

// ═══════════════════════════════════════════════════════════════════════
// #466 — `<c:txPr><a:p><a:pPr><a:defRPr>` was written out at the call
// site 43 times across five files; axis.ts alone had 14, so reading one
// axis re-walked the same subtree once per attribute and adding a
// property meant another copy of the walk.
//
// The interesting part is not that it was repeated but what each copy
// had to restate: the walk is *scoped*, not a search. A `<a:defRPr>`
// inside the axis's `<c:title><c:tx><c:rich>` must not reach the
// tick-label readers, and vice versa. That invariant now lives in one
// place, so it is worth testing there.
// ═══════════════════════════════════════════════════════════════════════

const el = (xml: string): XmlElement => parseXml(xml)

/** An axis with typography on the tick labels, the title, or both. */
function axis(opts: { label?: string; title?: string }): XmlElement {
  const txPr = opts.label
    ? `<c:txPr><a:p><a:pPr><a:defRPr ${opts.label}/></a:pPr></a:p></c:txPr>`
    : ""
  const title = opts.title
    ? `<c:title><c:tx><c:rich><a:p><a:pPr><a:defRPr ${opts.title}/></a:pPr></a:p></c:rich></c:tx></c:title>`
    : ""
  return el(`<c:catAx xmlns:c="c" xmlns:a="a">${title}${txPr}</c:catAx>`)
}

describe("the two walks land on the right defRPr", () => {
  it("each finds its own host's", () => {
    const both = axis({ label: 'sz="1000"', title: 'sz="2000"' })

    expect(resolveTxPrDefRPr(both)?.attrs.sz).toBe("1000")
    expect(resolveTitleDefRPr(both)?.attrs.sz).toBe("2000")
  })

  it("neither reaches into the other, which is the whole point", () => {
    // A search would find the wrong one here; a walk cannot.
    expect(resolveTxPrDefRPr(axis({ title: 'sz="2000"' }))).toBeUndefined()
    expect(resolveTitleDefRPr(axis({ label: 'sz="1000"' }))).toBeUndefined()
  })

  it("stops at the first missing link rather than guessing", () => {
    // A malformed chain surfaces as absence, not a fabricated value.
    expect(resolveTxPrDefRPr(el('<c:catAx xmlns:c="c"/>'))).toBeUndefined()
    expect(
      resolveTxPrDefRPr(el('<c:catAx xmlns:c="c" xmlns:a="a"><c:txPr/></c:catAx>')),
    ).toBeUndefined()
    expect(
      resolveTxPrDefRPr(el('<c:catAx xmlns:c="c" xmlns:a="a"><c:txPr><a:p/></c:txPr></c:catAx>')),
    ).toBeUndefined()
  })
})

describe("the readers built on it keep their scoping", () => {
  it("tick-label readers do not see the title's typography", () => {
    const titleOnly = axis({ title: 'sz="2000" b="1"' })

    expect(parseAxisLabelFontSize(titleOnly)).toBeUndefined()
    expect(parseAxisLabelBold(titleOnly)).toBeUndefined()
    expect(parseAxisTitleFontSize(titleOnly)).toBe(20)
    expect(parseAxisTitleBold(titleOnly)).toBe(true)
  })

  it("title readers do not see the tick labels'", () => {
    const labelOnly = axis({ label: 'sz="1000" b="1"' })

    expect(parseAxisTitleFontSize(labelOnly)).toBeUndefined()
    expect(parseAxisTitleBold(labelOnly)).toBeUndefined()
    expect(parseAxisLabelFontSize(labelOnly)).toBe(10)
    expect(parseAxisLabelBold(labelOnly)).toBe(true)
  })

  it("both read their own when both are present", () => {
    const both = axis({ label: 'sz="1000"', title: 'sz="2000"' })

    expect(parseAxisLabelFontSize(both)).toBe(10)
    expect(parseAxisTitleFontSize(both)).toBe(20)
  })
})
