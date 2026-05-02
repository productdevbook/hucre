// ── Content Types Writer ──────────────────────────────────────────────
// Generates [Content_Types].xml for an XLSX package.

import { xmlDocument, xmlSelfClose } from "../xml/writer";

const NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

const CT_RELS = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML = "application/xml";
const CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKBOOK_MACRO = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
const CT_VBA_PROJECT = "application/vnd.ms-office.vbaProject";
const CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SHARED_STRINGS =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";
const CT_THREADED_COMMENTS = "application/vnd.ms-excel.threadedcomments+xml";
const CT_PERSON = "application/vnd.ms-excel.person+xml";
const CT_EXTERNAL_LINK =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
const CT_SLICER = "application/vnd.ms-excel.slicer+xml";
const CT_SLICER_CACHE = "application/vnd.ms-excel.slicerCache+xml";
const CT_TIMELINE = "application/vnd.ms-excel.timeline+xml";
const CT_TIMELINE_CACHE = "application/vnd.ms-excel.timelineCache+xml";

/** Image extension → content type mapping */
const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
  webp: "image/webp",
};

export interface ContentTypesOptions {
  sheetCount: number;
  hasSharedStrings: boolean;
  /** 1-based indices of drawings (e.g. [1, 3] means drawing1.xml and drawing3.xml exist) */
  drawingIndices?: number[];
  /** Set of image extensions used (e.g. new Set(["png", "jpeg"])) */
  imageExtensions?: Set<string>;
  /** 1-based indices of comments (e.g. [1, 2] means comments1.xml and comments2.xml exist) */
  commentIndices?: number[];
  /** 1-based indices of tables (e.g. [1, 2, 3] means table1.xml, table2.xml, table3.xml exist) */
  tableIndices?: number[];
  /**
   * 1-based indices of sheets that have a threadedComments part. Each
   * entry adds an `Override` for `/xl/threadedComments/threadedCommentN.xml`.
   */
  threadedCommentSheetIndices?: number[];
  /** Whether `xl/persons/person.xml` is present. */
  hasPersons?: boolean;
  /**
   * 1-based indices of external link parts. Each entry adds an
   * `Override` for `/xl/externalLinks/externalLinkN.xml`.
   */
  externalLinkIndices?: number[];
  /**
   * 1-based indices of per-sheet slicer parts. Each entry adds an
   * `Override` for `/xl/slicers/slicerN.xml`.
   */
  slicerIndices?: number[];
  /**
   * 1-based indices of workbook-level slicer cache parts. Each entry
   * adds an `Override` for `/xl/slicerCaches/slicerCacheN.xml`.
   */
  slicerCacheIndices?: number[];
  /**
   * 1-based indices of per-sheet timeline parts. Each entry adds an
   * `Override` for `/xl/timelines/timelineN.xml`.
   */
  timelineIndices?: number[];
  /**
   * 1-based indices of workbook-level timeline cache parts. Each entry
   * adds an `Override` for `/xl/timelineCaches/timelineCacheN.xml`.
   */
  timelineCacheIndices?: number[];
  /** Whether docProps/core.xml is present */
  hasCoreProps?: boolean;
  /** Whether docProps/app.xml is present */
  hasAppProps?: boolean;
  /** Whether docProps/custom.xml is present */
  hasCustomProps?: boolean;
  /** Whether VBA macros are present (xl/vbaProject.bin). Uses XLSM content types. */
  hasMacros?: boolean;
  /** Whether Excel 2024 checkbox FeaturePropertyBag is present. */
  hasFeaturePropertyBag?: boolean;
}

/** Generate [Content_Types].xml for XLSX */
export function writeContentTypes(
  sheetCountOrOptions: number | ContentTypesOptions,
  hasSharedStrings?: boolean,
): string {
  // Support both old and new call signatures
  let opts: ContentTypesOptions;
  if (typeof sheetCountOrOptions === "number") {
    opts = {
      sheetCount: sheetCountOrOptions,
      hasSharedStrings: hasSharedStrings ?? false,
    };
  } else {
    opts = sheetCountOrOptions;
  }

  const children: string[] = [];

  // Default extension mappings
  children.push(xmlSelfClose("Default", { Extension: "rels", ContentType: CT_RELS }));
  children.push(xmlSelfClose("Default", { Extension: "xml", ContentType: CT_XML }));

  // Default extension for VBA binary (needed for macro-enabled workbooks)
  if (opts.hasMacros) {
    children.push(xmlSelfClose("Default", { Extension: "bin", ContentType: CT_VBA_PROJECT }));
  }

  // Default extensions for image types
  if (opts.imageExtensions) {
    for (const ext of opts.imageExtensions) {
      const ct = IMAGE_CONTENT_TYPES[ext];
      if (ct) {
        children.push(xmlSelfClose("Default", { Extension: ext, ContentType: ct }));
      }
    }
  }

  // Override for workbook (use macro-enabled content type when VBA present)
  children.push(
    xmlSelfClose("Override", {
      PartName: "/xl/workbook.xml",
      ContentType: opts.hasMacros ? CT_WORKBOOK_MACRO : CT_WORKBOOK,
    }),
  );

  // Override for each worksheet
  for (let i = 1; i <= opts.sheetCount; i++) {
    children.push(
      xmlSelfClose("Override", {
        PartName: `/xl/worksheets/sheet${i}.xml`,
        ContentType: CT_WORKSHEET,
      }),
    );
  }

  // Override for styles
  children.push(
    xmlSelfClose("Override", {
      PartName: "/xl/styles.xml",
      ContentType: CT_STYLES,
    }),
  );

  // Override for theme
  children.push(
    xmlSelfClose("Override", {
      PartName: "/xl/theme/theme1.xml",
      ContentType: CT_THEME,
    }),
  );

  // Override for shared strings (if present)
  if (opts.hasSharedStrings) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/xl/sharedStrings.xml",
        ContentType: CT_SHARED_STRINGS,
      }),
    );
  }

  // Override for each drawing
  if (opts.drawingIndices) {
    for (const idx of opts.drawingIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/drawings/drawing${idx}.xml`,
          ContentType: CT_DRAWING,
        }),
      );
    }
  }

  // Default extension for VML (needed for comment shapes)
  if (opts.commentIndices && opts.commentIndices.length > 0) {
    children.push(xmlSelfClose("Default", { Extension: "vml", ContentType: CT_VML }));
  }

  // Override for each comments file
  if (opts.commentIndices) {
    for (const idx of opts.commentIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/comments${idx}.xml`,
          ContentType: CT_COMMENTS,
        }),
      );
    }
  }

  // Override for each table
  if (opts.tableIndices) {
    for (const idx of opts.tableIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/tables/table${idx}.xml`,
          ContentType: CT_TABLE,
        }),
      );
    }
  }

  // Override for each threadedComments part (Excel 365)
  if (opts.threadedCommentSheetIndices) {
    for (const idx of opts.threadedCommentSheetIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/threadedComments/threadedComment${idx}.xml`,
          ContentType: CT_THREADED_COMMENTS,
        }),
      );
    }
  }

  // Override for the workbook-wide persons directory
  if (opts.hasPersons) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/xl/persons/person.xml",
        ContentType: CT_PERSON,
      }),
    );
  }

  // Override for each external link
  if (opts.externalLinkIndices) {
    for (const idx of opts.externalLinkIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/externalLinks/externalLink${idx}.xml`,
          ContentType: CT_EXTERNAL_LINK,
        }),
      );
    }
  }

  // Override for each slicer
  if (opts.slicerIndices) {
    for (const idx of opts.slicerIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/slicers/slicer${idx}.xml`,
          ContentType: CT_SLICER,
        }),
      );
    }
  }

  // Override for each slicer cache
  if (opts.slicerCacheIndices) {
    for (const idx of opts.slicerCacheIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/slicerCaches/slicerCache${idx}.xml`,
          ContentType: CT_SLICER_CACHE,
        }),
      );
    }
  }

  // Override for each timeline
  if (opts.timelineIndices) {
    for (const idx of opts.timelineIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/timelines/timeline${idx}.xml`,
          ContentType: CT_TIMELINE,
        }),
      );
    }
  }

  // Override for each timeline cache
  if (opts.timelineCacheIndices) {
    for (const idx of opts.timelineCacheIndices) {
      children.push(
        xmlSelfClose("Override", {
          PartName: `/xl/timelineCaches/timelineCache${idx}.xml`,
          ContentType: CT_TIMELINE_CACHE,
        }),
      );
    }
  }

  // Override for FeaturePropertyBag (Excel 2024 checkboxes)
  if (opts.hasFeaturePropertyBag) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/xl/featurePropertyBag/featurePropertyBag.xml",
        ContentType: "application/vnd.ms-excel.featurepropertybag+xml",
      }),
    );
  }

  // Override for docProps
  if (opts.hasCoreProps) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/docProps/core.xml",
        ContentType: "application/vnd.openxmlformats-package.core-properties+xml",
      }),
    );
  }
  if (opts.hasAppProps) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/docProps/app.xml",
        ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
      }),
    );
  }
  if (opts.hasCustomProps) {
    children.push(
      xmlSelfClose("Override", {
        PartName: "/docProps/custom.xml",
        ContentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
      }),
    );
  }

  return xmlDocument("Types", { xmlns: NS_CONTENT_TYPES }, children);
}
