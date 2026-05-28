// ── FeaturePropertyBag (Excel 2024 checkboxes) ────────────────────────
// Microsoft 2022 OOXML extension that backs the Excel 2024 native cell
// checkbox. The on-the-wire shape was reverse-engineered from XlsxWriter's
// fixtures; see `test/xlsx-checkbox.test.ts` for round-trip coverage.
//
// Spec status: this is a Microsoft-only extension that post-dates ECMA-376
// and is not (yet) part of the published OOXML spec. Apart from this writer
// hucre does not otherwise model the property bag.

import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer"

/** Namespace used by `xfpb:`-qualified elements inside cellXf and dxf extLst. */
export const FPB_NS = "http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"

/** Cell-XF complement extension URI (per-xf, used by checkboxes). */
export const FPB_XF_EXT_URI = "{C7286773-470A-42A8-94C5-96B5CB345126}"

/** DXF complement extension URI — unused today but reserved for the
 *  conditional-formatting + checkbox combination XlsxWriter supports. */
export const FPB_DXF_EXT_URI = "{0417FA29-78FA-4A13-93AC-8FF0FAFDF519}"

/** Workbook relationship type for the part. */
export const FPB_REL_TYPE =
  "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag"

/** Content-type for the part. */
export const FPB_CONTENT_TYPE = "application/vnd.ms-excel.featurepropertybag+xml"

/** Path inside the XLSX archive. */
export const FPB_PART_PATH = "xl/featurePropertyBag/featurePropertyBag.xml"

/**
 * Emit the fixed property-bag chain that Excel 2024 expects when at least
 * one cell uses a checkbox. The structure is identical regardless of how
 * many checkbox cells are in the workbook — XlsxWriter and rust_xlsxwriter
 * both emit this exact shape, and Excel only ever resolves it to a single
 * `<xfpb:xfComplement i="0"/>` entry on the relevant cellXfs.
 */
export function writeFeaturePropertyBagXml(): string {
  const bags = [
    xmlSelfClose("bag", { type: "Checkbox" }),
    xmlElement("bag", { type: "XFControls" }, [xmlElement("bagId", { k: "CellControl" }, "0")]),
    xmlElement("bag", { type: "XFComplement" }, [xmlElement("bagId", { k: "XFControls" }, "1")]),
    xmlElement("bag", { type: "XFComplements", extRef: "XFComplementsMapperExtRef" }, [
      xmlElement("a", { k: "MappedFeaturePropertyBags" }, [xmlElement("bagId", undefined, "2")]),
    ]),
  ]
  return xmlDocument("FeaturePropertyBags", { xmlns: FPB_NS }, bags)
}
