// ── Cell Metadata (dynamic arrays) ───────────────────────────────────
// xl/metadata.xml is the part the `cm` attribute on `<c>` indexes into.
// Without it a `cm` points at nothing, which is what hucre shipped
// before #423 (and it wrote `cm` on `<f>`, where the schema has no such
// attribute at all).
//
// Spec trail, since none of this is guessable from the schema alone:
//
//   • ISO/IEC 29500-1 §18.9 (Metadata) defines the part — `metadata`,
//     `metadataTypes` / `metadataType` (§18.9.10), `futureMetadata`
//     (§18.9.4), `cellMetadata`, `bk` (metadata block) and `rc`
//     (metadata record: @t selects the metadata type, @v the block
//     inside that type's futureMetadata).
//   • ISO/IEC 29500-1 §18.3.1.4 (c) calls @cm "a zero-based index" and
//     names no collection. [MS-OI29500] §18.3.1.4 note (a) corrects
//     both halves for Office: "Office specifies that @cm is a one-based
//     index into the cellMetadata collection in the metadata part."
//     hucre writes and reads what Office does, not what the base
//     standard's prose says.
//   • The dynamic-array payload is a Microsoft extension on top:
//     [MS-XLSX] "Metadata" pairs futureMetadata @name="XLDAPR" with the
//     ext @uri {BDBB8CDC-FA1E-496E-A857-3C3F30C029C3} carrying
//     `<xda:dynamicArrayProperties>`.
//
// The concrete attribute set on `<metadataType>` below is not derivable
// from the spec — it is behavioural policy Excel writes verbatim — so it
// is copied from XlsxWriter's `metadata.py`, a producer whose output
// Excel is known to accept. It is *not* verified against an Excel-made
// file here: this repo ships no XLSX fixtures to compare against.

import { parseXml } from "../xml/parser"
import type { XmlElement } from "../xml/parser"
import { childElements, findChild, parseIntSafe } from "../xml/tree"
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer"

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
/** Microsoft dynamic-array namespace, introduced with Excel 365 spilling. */
const NS_XDA = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"

/** Metadata type name Excel reserves for dynamic-array properties. */
export const XLDAPR = "XLDAPR"

/** ext URI that carries `<xda:dynamicArrayProperties>` ([MS-XLSX] Metadata). */
export const XLDAPR_EXT_URI = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"

/** Path of the part inside the XLSX archive. */
export const METADATA_PART_PATH = "xl/metadata.xml"

/** Content type for the Override in [Content_Types].xml. */
export const METADATA_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"

/** Workbook relationship type for the part. */
export const METADATA_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"

/**
 * The `cm` value hucre puts on every dynamic-array cell. One-based, and
 * hucre only ever emits one cellMetadata block, so it is always 1.
 */
export const DYNAMIC_ARRAY_CM = 1

/**
 * Emit the fixed metadata part that backs `cm="1"`.
 *
 * The shape is the same for every workbook regardless of how many cells
 * spill: all of them share the single "this is a dynamic array" record,
 * exactly as Excel and XlsxWriter do. `minSupportedVersion="120000"` is
 * the value Office requires for its own metadata types ([MS-OI29500]
 * §18.9.10 states the requirement for XLMDX; XLDAPR is written with the
 * same value by every producer we could inspect).
 */
export function writeMetadataXml(): string {
  const metadataType = xmlSelfClose("metadataType", {
    name: XLDAPR,
    minSupportedVersion: 120000,
    copy: 1,
    pasteAll: 1,
    pasteValues: 1,
    merge: 1,
    splitFirst: 1,
    rowColShift: 1,
    clearFormats: 1,
    clearComments: 1,
    assign: 1,
    coerce: 1,
    cellMeta: 1,
  })

  // futureMetadata block 0 for XLDAPR: the properties `rc/@v` points at.
  const futureBlock = xmlElement("bk", undefined, [
    xmlElement("extLst", undefined, [
      xmlElement("ext", { uri: XLDAPR_EXT_URI }, [
        xmlSelfClose("xda:dynamicArrayProperties", { fDynamic: 1, fCollapsed: 0 }),
      ]),
    ]),
  ])

  // cellMetadata block 1 (one-based, per [MS-OI29500] §18.3.1.4): type 1
  // is the XLDAPR entry declared above, value 0 its futureMetadata block.
  const cellBlock = xmlElement("bk", undefined, [xmlSelfClose("rc", { t: 1, v: 0 })])

  return xmlDocument("metadata", { xmlns: NS_SPREADSHEET, "xmlns:xda": NS_XDA }, [
    xmlElement("metadataTypes", { count: 1 }, [metadataType]),
    xmlElement("futureMetadata", { name: XLDAPR, count: 1 }, [futureBlock]),
    xmlElement("cellMetadata", { count: 1 }, [cellBlock]),
  ])
}

/**
 * Resolve which `cm` indexes in this package mean "dynamic array".
 *
 * Returns the set of one-based cellMetadata indexes whose record
 * resolves, through `rc/@t` → metadataType and `rc/@v` → futureMetadata
 * block, to an XLDAPR entry. A `cm` outside the set points at some other
 * kind of cell metadata and says nothing about spilling.
 */
export function parseDynamicArrayCellMetadata(xml: string): Set<number> {
  const dynamic = new Set<number>()
  const root = parseXml(xml)
  const meta = root.local === "metadata" ? root : findChild(root, "metadata")
  if (!meta) return dynamic

  // `rc/@t` is a one-based index into the metadataTypes collection.
  const typeNames: string[] = []
  const types = findChild(meta, "metadataTypes")
  if (types) {
    for (const type of childElements(types)) {
      if (type.local === "metadataType") typeNames.push(type.attrs["name"] ?? "")
    }
  }

  // `rc/@v` is a zero-based index into the futureMetadata blocks of the
  // record's own type. Excel writes fDynamic="0" for an array formula
  // that is *not* spilling, so the block has the final say.
  const blocksByType = new Map<string, boolean[]>()
  for (const future of childElements(meta)) {
    if (future.local !== "futureMetadata") continue
    const name = future.attrs["name"] ?? ""
    const blocks = blocksByType.get(name) ?? []
    for (const bk of childElements(future)) {
      if (bk.local === "bk") blocks.push(blockIsDynamicArray(bk))
    }
    blocksByType.set(name, blocks)
  }

  const cellMetadata = findChild(meta, "cellMetadata")
  if (!cellMetadata) return dynamic

  let index = 0
  for (const bk of childElements(cellMetadata)) {
    if (bk.local !== "bk") continue
    index++ // one-based
    for (const rc of childElements(bk)) {
      if (rc.local !== "rc") continue
      if (typeNames[parseIntSafe(rc.attrs["t"], 0) - 1] !== XLDAPR) continue
      const blocks = blocksByType.get(XLDAPR)
      const v = parseIntSafe(rc.attrs["v"], -1)
      // With no futureMetadata block to consult, the type name is all we
      // have — and XLDAPR exists for exactly one purpose.
      if (!blocks || v < 0 || v >= blocks.length || blocks[v]) dynamic.add(index)
    }
  }

  return dynamic
}

/** Whether a futureMetadata `<bk>` carries dynamic-array properties. */
function blockIsDynamicArray(bk: XmlElement): boolean {
  for (const extLst of childElements(bk)) {
    if (extLst.local !== "extLst") continue
    for (const ext of childElements(extLst)) {
      if (ext.local !== "ext") continue
      // GUIDs are case-insensitive; [MS-XLSX] prints this one uppercase,
      // Excel and XlsxWriter both write it lowercase.
      if ((ext.attrs["uri"] ?? "").toLowerCase() !== XLDAPR_EXT_URI) continue
      for (const props of childElements(ext)) {
        if (props.local !== "dynamicArrayProperties") continue
        const fDynamic = props.attrs["fDynamic"]
        return fDynamic !== "0" && fDynamic !== "false"
      }
    }
  }
  // No extension at all: the block asserts nothing, so defer to the type.
  return true
}
