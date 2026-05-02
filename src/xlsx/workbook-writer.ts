// ── Workbook XML Writer ──────────────────────────────────────────────
// Generates xl/workbook.xml, xl/_rels/workbook.xml.rels, and _rels/.rels

import type { WriteSheet, NamedRange } from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";
import { hashSheetPassword } from "./password";

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_WORKSHEET =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_SHARED_STRINGS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_WORKBOOK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * A pivotCache wiring entry emitted in workbook.xml. `cacheId` is the
 * stable handle pivot tables reference; `rId` resolves through
 * `xl/_rels/workbook.xml.rels` to the cache definition path.
 */
export interface PivotCacheRef {
  cacheId: number;
  rId: string;
}

/** Generate xl/workbook.xml */
export function writeWorkbookXml(
  sheets: WriteSheet[],
  namedRanges?: NamedRange[],
  dateSystem?: "1900" | "1904",
  activeSheet?: number,
  workbookProtection?: { lockStructure?: boolean; lockWindows?: boolean; password?: string },
  externalLinkRels?: ReadonlyArray<{ rId: string }>,
  pivotCacheRefs?: ReadonlyArray<PivotCacheRef>,
  slicerCacheRels?: ReadonlyArray<{ rId: string }>,
  timelineCacheRels?: ReadonlyArray<{ rId: string }>,
): string {
  const sheetElements: string[] = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const attrs: Record<string, string | number> = {
      name: sheet.name,
      sheetId: i + 1,
      "r:id": `rId${i + 1}`,
    };
    if (sheet.hidden) {
      attrs["state"] = "hidden";
    } else if (sheet.veryHidden) {
      attrs["state"] = "veryHidden";
    }
    sheetElements.push(xmlSelfClose("sheet", attrs));
  }

  const parts: string[] = [];

  // workbookPr with date1904 attribute when using 1904 date system
  if (dateSystem === "1904") {
    parts.push(xmlSelfClose("workbookPr", { date1904: 1 }));
  }

  // bookViews — tells Excel which sheet tab is active
  const activeTab = activeSheet ?? 0;
  parts.push(
    xmlElement("bookViews", undefined, [
      xmlSelfClose("workbookView", {
        xWindow: 0,
        yWindow: 0,
        windowWidth: 16384,
        windowHeight: 8192,
        activeTab,
      }),
    ]),
  );

  // workbookProtection — lock structure and/or windows
  if (workbookProtection) {
    const protAttrs: Record<string, string | number> = {};
    if (workbookProtection.lockStructure) protAttrs["lockStructure"] = 1;
    if (workbookProtection.lockWindows) protAttrs["lockWindows"] = 1;
    if (workbookProtection.password) {
      protAttrs["workbookPassword"] = hashSheetPassword(workbookProtection.password);
    }
    parts.push(xmlSelfClose("workbookProtection", protAttrs));
  }

  parts.push(xmlElement("sheets", undefined, sheetElements));

  // ── Defined Names (named ranges + print area/titles) ──
  if (namedRanges && namedRanges.length > 0) {
    // Build sheet name → index map for resolving scoped named ranges
    const sheetIndexMap = new Map<string, number>();
    for (let i = 0; i < sheets.length; i++) {
      sheetIndexMap.set(sheets[i].name, i);
    }

    const dnElements: string[] = [];

    for (const nr of namedRanges) {
      const attrs: Record<string, string | number> = {
        name: nr.name,
      };

      // Resolve scope: if scope is a sheet name, convert to localSheetId (0-based index)
      if (nr.scope !== undefined) {
        const idx = sheetIndexMap.get(nr.scope);
        if (idx !== undefined) {
          attrs["localSheetId"] = idx;
        }
      }

      if (nr.comment) {
        attrs["comment"] = nr.comment;
      }

      dnElements.push(xmlElement("definedName", attrs, xmlEscape(nr.range)));
    }

    parts.push(xmlElement("definedNames", undefined, dnElements));
  }

  // ── externalReferences — ECMA-376 §18.2.2 places the block after
  // definedNames and before calcPr. Excel tolerates other orders, but
  // the spec order is what we emit so generated files validate clean.
  if (externalLinkRels && externalLinkRels.length > 0) {
    const refChildren = externalLinkRels.map((r) =>
      xmlSelfClose("externalReference", { "r:id": r.rId }),
    );
    parts.push(xmlElement("externalReferences", undefined, refChildren));
  }

  // ── calcPr — tells Excel to recalculate all formulas on open ──
  parts.push(xmlSelfClose("calcPr", { calcId: 0, fullCalcOnLoad: 1 }));

  // ── pivotCaches — wires cacheId values to workbook-rel rIds. ECMA-376
  // §18.2.18 puts this block after calcPr; Excel won't recognise pivot
  // tables on roundtrip without it.
  if (pivotCacheRefs && pivotCacheRefs.length > 0) {
    const cacheChildren = pivotCacheRefs.map((c) =>
      xmlSelfClose("pivotCache", { cacheId: c.cacheId, "r:id": c.rId }),
    );
    parts.push(xmlElement("pivotCaches", undefined, cacheChildren));
  }

  // ── extLst — slicer caches (x14) and timeline caches (x15) ──
  // These extension blocks point Excel at the slicerCacheN.xml and
  // timelineCacheN.xml parts via rIds declared in workbook.xml.rels.
  // Without them Excel treats the cache parts as orphans and drops the
  // associated slicers / timelines on next open.
  const extElements: string[] = [];

  if (slicerCacheRels && slicerCacheRels.length > 0) {
    const slicerRefs = slicerCacheRels.map((r) =>
      xmlSelfClose("x14:slicerCache", { "r:id": r.rId }),
    );
    const slicerCachesEl = xmlElement("x14:slicerCaches", undefined, slicerRefs);
    extElements.push(
      xmlElement(
        "ext",
        {
          uri: "{BBE1A952-AA13-448E-AADC-164F8A28A991}",
          "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
        },
        [slicerCachesEl],
      ),
    );
  }

  if (timelineCacheRels && timelineCacheRels.length > 0) {
    const timelineRefs = timelineCacheRels.map((r) =>
      xmlSelfClose("x15:timelineCachePivotCache", { "r:id": r.rId }),
    );
    const timelineCachesEl = xmlElement("x15:timelineCachePivotCaches", undefined, timelineRefs);
    extElements.push(
      xmlElement(
        "ext",
        {
          uri: "{7E03D99C-DC04-49D9-9315-930204A7B6E9}",
          "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
        },
        [timelineCachesEl],
      ),
    );
  }

  if (extElements.length > 0) {
    parts.push(xmlElement("extLst", undefined, extElements));
  }

  return xmlDocument("workbook", { xmlns: NS_SPREADSHEET, "xmlns:r": NS_R }, parts);
}

const REL_VBA_PROJECT = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
const REL_FEATURE_PROPERTY_BAG =
  "http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag";

const REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person";
const REL_EXTERNAL_LINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
const REL_PIVOT_CACHE_DEFINITION =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
const REL_TIMELINE_CACHE = "http://schemas.microsoft.com/office/2011/relationships/timelineCache";

/** A relationship description for an externalLink emitted in workbook.xml.rels. */
export interface ExternalLinkRel {
  rId: string;
  /** Path relative to the workbook directory, e.g. "externalLinks/externalLink1.xml". */
  target: string;
}

/**
 * A workbook-level pivot cache definition relationship. The rId is
 * shared with the matching `<pivotCache cacheId="..." r:id="..."/>` in
 * workbook.xml so the two stay in lockstep.
 */
export interface PivotCacheRel {
  rId: string;
  /** Path relative to xl/, e.g. "pivotCache/pivotCacheDefinition1.xml". */
  target: string;
}

/** A workbook-level relationship to a slicerCache or timelineCache part. */
export interface CacheRel {
  rId: string;
  /** Path relative to the workbook directory, e.g. "slicerCaches/slicerCache1.xml". */
  target: string;
}

/** Generate xl/_rels/workbook.xml.rels */
export function writeWorkbookRels(
  sheetCount: number,
  hasSharedStrings: boolean,
  hasMacros?: boolean,
  hasFeaturePropertyBag?: boolean,
  hasPersons?: boolean,
  externalLinkRels?: ReadonlyArray<ExternalLinkRel>,
  pivotCacheRels?: ReadonlyArray<PivotCacheRel>,
  slicerCacheRels?: ReadonlyArray<CacheRel>,
  timelineCacheRels?: ReadonlyArray<CacheRel>,
): string {
  const children: string[] = [];

  // Worksheet relationships
  for (let i = 1; i <= sheetCount; i++) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${i}`,
        Type: REL_WORKSHEET,
        Target: `worksheets/sheet${i}.xml`,
      }),
    );
  }

  // Styles relationship
  let nextRid = sheetCount + 1;
  children.push(
    xmlSelfClose("Relationship", {
      Id: `rId${nextRid}`,
      Type: REL_STYLES,
      Target: "styles.xml",
    }),
  );
  nextRid++;

  // Shared strings relationship (if present)
  if (hasSharedStrings) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRid}`,
        Type: REL_SHARED_STRINGS,
        Target: "sharedStrings.xml",
      }),
    );
    nextRid++;
  }

  // Theme relationship
  children.push(
    xmlSelfClose("Relationship", {
      Id: `rId${nextRid}`,
      Type: REL_THEME,
      Target: "theme/theme1.xml",
    }),
  );
  nextRid++;

  // VBA project relationship (for macro-enabled workbooks)
  if (hasMacros) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRid}`,
        Type: REL_VBA_PROJECT,
        Target: "vbaProject.bin",
      }),
    );
    nextRid++;
  }

  // FeaturePropertyBag relationship (Excel 2024 checkboxes)
  if (hasFeaturePropertyBag) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRid}`,
        Type: REL_FEATURE_PROPERTY_BAG,
        Target: "featurePropertyBag/featurePropertyBag.xml",
      }),
    );
    nextRid++;
  }

  // Threaded-comments person directory (Excel 365)
  if (hasPersons) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRid}`,
        Type: REL_PERSON,
        Target: "persons/person.xml",
      }),
    );
    nextRid++;
  }

  // External link relationships (caller supplies pre-assigned rIds)
  if (externalLinkRels) {
    for (const link of externalLinkRels) {
      children.push(
        xmlSelfClose("Relationship", {
          Id: link.rId,
          Type: REL_EXTERNAL_LINK,
          Target: link.target,
        }),
      );
    }
  }

  // Pivot cache definition relationships (caller supplies pre-assigned rIds
  // — the rIds also appear in workbook.xml's <pivotCaches> block, so the
  // two must agree).
  if (pivotCacheRels) {
    for (const cache of pivotCacheRels) {
      children.push(
        xmlSelfClose("Relationship", {
          Id: cache.rId,
          Type: REL_PIVOT_CACHE_DEFINITION,
          Target: cache.target,
        }),
      );
    }
  }

  // Slicer cache relationships (Excel 2010+)
  if (slicerCacheRels) {
    for (const r of slicerCacheRels) {
      children.push(
        xmlSelfClose("Relationship", {
          Id: r.rId,
          Type: REL_SLICER_CACHE,
          Target: r.target,
        }),
      );
    }
  }

  // Timeline cache relationships (Excel 2013+)
  if (timelineCacheRels) {
    for (const r of timelineCacheRels) {
      children.push(
        xmlSelfClose("Relationship", {
          Id: r.rId,
          Type: REL_TIMELINE_CACHE,
          Target: r.target,
        }),
      );
    }
  }

  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, children);
}

/** Generate _rels/.rels (with optional docProps) */
export function writeRootRels(options?: {
  hasCoreProps?: boolean;
  hasAppProps?: boolean;
  hasCustomProps?: boolean;
}): string {
  const children: string[] = [];

  children.push(
    xmlSelfClose("Relationship", {
      Id: "rId1",
      Type: REL_WORKBOOK,
      Target: "xl/workbook.xml",
    }),
  );

  let nextRId = 2;

  if (options?.hasCoreProps) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRId++}`,
        Type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        Target: "docProps/core.xml",
      }),
    );
  }

  if (options?.hasAppProps) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRId++}`,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        Target: "docProps/app.xml",
      }),
    );
  }

  if (options?.hasCustomProps) {
    children.push(
      xmlSelfClose("Relationship", {
        Id: `rId${nextRId++}`,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
        Target: "docProps/custom.xml",
      }),
    );
  }

  return xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, children);
}
