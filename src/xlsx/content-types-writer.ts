// ── Content Types Writer ──────────────────────────────────────────────
// Generates [Content_Types].xml for an XLSX package.

import { xmlDocument, xmlSelfClose } from "../xml/writer";

const NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

const CT_RELS = "application/vnd.openxmlformats-package.relationships+xml";
const CT_XML = "application/xml";
const CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SHARED_STRINGS =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

/** Image extension → content type mapping */
const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
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
  /** Whether docProps/core.xml is present */
  hasCoreProps?: boolean;
  /** Whether docProps/app.xml is present */
  hasAppProps?: boolean;
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

  // Default extensions for image types
  if (opts.imageExtensions) {
    for (const ext of opts.imageExtensions) {
      const ct = IMAGE_CONTENT_TYPES[ext];
      if (ct) {
        children.push(xmlSelfClose("Default", { Extension: ext, ContentType: ct }));
      }
    }
  }

  // Override for workbook
  children.push(
    xmlSelfClose("Override", {
      PartName: "/xl/workbook.xml",
      ContentType: CT_WORKBOOK,
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

  return xmlDocument("Types", { xmlns: NS_CONTENT_TYPES }, children);
}
