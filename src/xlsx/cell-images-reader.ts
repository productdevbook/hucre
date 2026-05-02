// ── Cell-Embedded Images Reader (WPS DISPIMG) ─────────────────────
// Parses `xl/cellimages.xml`, the WPS Office part that backs the
// `=_xlfn.DISPIMG("<id>", 1)` formula. Recent Excel versions also
// recognize this part for round-tripping, so a lot of real-world
// XLSX files (especially those produced by WPS / Kingsoft Office)
// carry it.
//
// Layout:
//   xl/cellimages.xml                 — list of `<etc:cellImage>` entries
//   xl/_rels/cellimages.xml.rels      — image rId → media path
//   xl/media/imageN.{png|jpeg|...}    — actual binaries
//
// The workbook-level relationship is declared in `xl/_rels/workbook.xml.rels`
// with type `http://www.wps.cn/officeDocument/2017/relationships/cellimage`.
//
// Reference: WPS Office et-custom-data namespace
//   http://www.wps.cn/officeDocument/2017/etCustomData

import type { CellImage, SheetImage } from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement } from "../xml/parser";

/** Relationship type that points from `workbook.xml.rels` at `cellimages.xml`. */
export const REL_CELL_IMAGES = "http://www.wps.cn/officeDocument/2017/relationships/cellimage";

/** A single entry parsed from `xl/cellimages.xml` minus the binary. */
export interface ParsedCellImageRef {
  /** DISPIMG id — value of the `name` attribute on `xdr:cNvPr`. */
  id: string;
  /** Image-relationship rId pointing into `cellimages.xml.rels`. */
  embedRId: string;
  /** Optional `descr` attribute on `xdr:cNvPr`. */
  description?: string;
}

/**
 * Parse `xl/cellimages.xml` into a list of references. The caller is
 * responsible for resolving each `embedRId` against
 * `xl/_rels/cellimages.xml.rels` and pulling the binary from the
 * package, which is what `assembleCellImages` below does.
 */
export function parseCellImages(xml: string): ParsedCellImageRef[] {
  const root = parseXml(xml);
  const out: ParsedCellImageRef[] = [];
  for (const child of childElements(root)) {
    if (child.local !== "cellImage") continue;
    const ref = parseCellImageEntry(child);
    if (ref) out.push(ref);
  }
  return out;
}

/**
 * Pure-data assembly step: combine parsed references with resolved
 * media bytes. Filters out entries whose media is missing so callers
 * never see half-populated `CellImage` records, and dedupes by id
 * (first occurrence wins).
 */
export function assembleCellImages(
  refs: readonly ParsedCellImageRef[],
  media: ReadonlyMap<string, { data: Uint8Array; type: SheetImage["type"] }>,
): CellImage[] {
  const out: CellImage[] = [];
  const seen = new Set<string>();
  for (const ref of refs) {
    if (seen.has(ref.id)) continue;
    const m = media.get(ref.embedRId);
    if (!m) continue;
    const entry: CellImage = { id: ref.id, data: m.data, type: m.type };
    if (ref.description) entry.description = ref.description;
    out.push(entry);
    seen.add(ref.id);
  }
  return out;
}

// ── Internals ─────────────────────────────────────────────────────

/**
 * Pull `id` (DISPIMG name) and `embedRId` out of one `<etc:cellImage>` /
 * `<xdr:pic>` block. Returns `undefined` when the entry lacks either —
 * those rows are unreferenceable so we drop them on read.
 */
function parseCellImageEntry(el: XmlElement): ParsedCellImageRef | undefined {
  const pic = findChild(el, "pic");
  if (!pic) return undefined;

  const nvPicPr = findChild(pic, "nvPicPr");
  const cNvPr = nvPicPr ? findChild(nvPicPr, "cNvPr") : undefined;
  const id = cNvPr?.attrs.name;
  if (!id) return undefined;

  const blipFill = findChild(pic, "blipFill");
  const blip = blipFill ? findChild(blipFill, "blip") : undefined;
  const embedRId = blip?.attrs["r:embed"] ?? blip?.attrs.embed;
  if (!embedRId) return undefined;

  const ref: ParsedCellImageRef = { id, embedRId };
  const description = cNvPr?.attrs.descr;
  if (description) ref.description = description;
  return ref;
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}
