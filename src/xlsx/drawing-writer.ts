// ── Drawing Writer ────────────────────────────────────────────────────
// Generates xl/drawings/drawingN.xml and related relationship files
// for embedding images, text boxes, and native charts in XLSX
// worksheets.

import type { SheetChart, SheetImage, SheetTextBox } from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose, xmlEscape } from "../xml/writer";

// ── Namespaces ───────────────────────────────────────────────────────

const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

// ── Constants ────────────────────────────────────────────────────────

/** Default image width in EMU (English Metric Units). 1 inch = 914400 EMU */
const DEFAULT_WIDTH_EMU = 3000000; // ~3.28 inches
/** Default image height in EMU */
const DEFAULT_HEIGHT_EMU = 2000000; // ~2.19 inches
/** Pixels to EMU conversion factor (96 DPI assumption) */
const EMU_PER_PIXEL = 9525;

// ── Types ────────────────────────────────────────────────────────────

export interface DrawingImage {
  /** Path inside the ZIP: xl/media/imageN.ext */
  path: string;
  /** Raw image bytes */
  data: Uint8Array;
  /** MIME content type */
  contentType: string;
}

export interface DrawingResult {
  /** The drawing XML content (xl/drawings/drawingN.xml) */
  drawingXml: string;
  /** The drawing relationships XML (xl/drawings/_rels/drawingN.xml.rels) */
  drawingRels: string;
  /** Image files to add to the ZIP */
  images: DrawingImage[];
  /**
   * Charts attached to this drawing, in declaration order. Each entry
   * carries the global chart number Excel uses for the part name
   * (`xl/charts/chart{n}.xml`). The drawing's rels file already
   * references them — this list lets `writer.ts` register the parts
   * in `[Content_Types].xml` and ZIP the chart bodies.
   */
  charts: { globalChartIndex: number }[];
}

// ── Extension / Content Type Mapping ─────────────────────────────────

const IMAGE_EXTENSIONS: Record<string, string> = {
  png: "png",
  jpeg: "jpeg",
  gif: "gif",
  svg: "svg",
  webp: "webp",
};

const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
  svg: "image/svg+xml",
  webp: "image/webp",
};

// ── Writer ───────────────────────────────────────────────────────────

/** Default textbox width in EMU */
const DEFAULT_TEXTBOX_WIDTH_EMU = 2000000; // ~2.19 inches
/** Default textbox height in EMU */
const DEFAULT_TEXTBOX_HEIGHT_EMU = 500000; // ~0.55 inches

/**
 * Generate drawing XML, drawing relationships, and image entries
 * for a worksheet's images, text boxes, and charts.
 *
 * @param images - Array of SheetImage objects to embed
 * @param imageStartIndex - Global image counter start (for unique image file names across sheets)
 * @param textBoxes - Optional array of SheetTextBox objects to embed
 * @param charts - Optional array of SheetChart objects to embed
 * @param chartStartIndex - Global chart counter start (1-based, for unique chartN.xml names across sheets)
 * @returns DrawingResult with XML, rels, image files, and chart registry
 */
export function writeDrawing(
  images: SheetImage[],
  imageStartIndex: number,
  textBoxes?: SheetTextBox[],
  charts?: SheetChart[],
  chartStartIndex: number = 1,
): DrawingResult {
  const drawingImages: DrawingImage[] = [];
  const relElements: string[] = [];
  const anchorElements: string[] = [];

  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    const imageIndex = imageStartIndex + i;
    const rId = `rId${i + 1}`;
    const ext = IMAGE_EXTENSIONS[img.type] ?? "png";
    const contentType = IMAGE_CONTENT_TYPES[img.type] ?? "image/png";
    const mediaPath = `xl/media/image${imageIndex}.${ext}`;
    const relTarget = `../media/image${imageIndex}.${ext}`;

    // Add image file entry
    drawingImages.push({
      path: mediaPath,
      data: img.data,
      contentType,
    });

    // Add relationship entry
    relElements.push(
      xmlSelfClose("Relationship", {
        Id: rId,
        Type: REL_IMAGE,
        Target: relTarget,
      }),
    );

    // Calculate dimensions in EMU
    const widthEmu = img.width ? img.width * EMU_PER_PIXEL : DEFAULT_WIDTH_EMU;
    const heightEmu = img.height ? img.height * EMU_PER_PIXEL : DEFAULT_HEIGHT_EMU;

    // Build twoCellAnchor element
    const fromCol = img.anchor.from.col;
    const fromRow = img.anchor.from.row;
    const toCol = img.anchor.to?.col ?? fromCol + 3;
    const toRow = img.anchor.to?.row ?? fromRow + 5;

    const fromElement = xmlElement("xdr:from", undefined, [
      xmlElement("xdr:col", undefined, String(fromCol)),
      xmlElement("xdr:colOff", undefined, "0"),
      xmlElement("xdr:row", undefined, String(fromRow)),
      xmlElement("xdr:rowOff", undefined, "0"),
    ]);

    const toElement = xmlElement("xdr:to", undefined, [
      xmlElement("xdr:col", undefined, String(toCol)),
      xmlElement("xdr:colOff", undefined, "0"),
      xmlElement("xdr:row", undefined, String(toRow)),
      xmlElement("xdr:rowOff", undefined, "0"),
    ]);

    const cNvPrAttrs: Record<string, string | number> = {
      id: i + 2,
      name: `Picture ${i + 1}`,
    };
    if (img.title) cNvPrAttrs.title = img.title;
    if (img.altText) cNvPrAttrs.descr = img.altText;

    const nvPicPr = xmlElement("xdr:nvPicPr", undefined, [
      xmlSelfClose("xdr:cNvPr", cNvPrAttrs),
      xmlElement("xdr:cNvPicPr", undefined, [xmlSelfClose("a:picLocks", { noChangeAspect: 1 })]),
    ]);

    const blipFill = xmlElement("xdr:blipFill", undefined, [
      xmlSelfClose("a:blip", { "r:embed": rId }),
      xmlElement("a:stretch", undefined, [xmlSelfClose("a:fillRect")]),
    ]);

    const spPr = xmlElement("xdr:spPr", undefined, [
      xmlElement("a:xfrm", undefined, [
        xmlSelfClose("a:off", { x: 0, y: 0 }),
        xmlSelfClose("a:ext", { cx: widthEmu, cy: heightEmu }),
      ]),
      xmlElement("a:prstGeom", { prst: "rect" }, [xmlSelfClose("a:avLst")]),
    ]);

    const pic = xmlElement("xdr:pic", undefined, [nvPicPr, blipFill, spPr]);

    const anchor = xmlElement("xdr:twoCellAnchor", undefined, [
      fromElement,
      toElement,
      pic,
      xmlSelfClose("xdr:clientData"),
    ]);

    anchorElements.push(anchor);
  }

  // ── Text Box Anchors ──
  if (textBoxes && textBoxes.length > 0) {
    // cNvPr id must be unique across all shapes in the drawing
    let shapeId = images.length + 2;

    for (let t = 0; t < textBoxes.length; t++) {
      const tb = textBoxes[t];

      const fromCol = tb.anchor.from.col;
      const fromRow = tb.anchor.from.row;
      const toCol = tb.anchor.to?.col ?? fromCol + 3;
      const toRow = tb.anchor.to?.row ?? fromRow + 3;

      const widthEmu = tb.width ? tb.width * EMU_PER_PIXEL : DEFAULT_TEXTBOX_WIDTH_EMU;
      const heightEmu = tb.height ? tb.height * EMU_PER_PIXEL : DEFAULT_TEXTBOX_HEIGHT_EMU;

      const fromElement = xmlElement("xdr:from", undefined, [
        xmlElement("xdr:col", undefined, String(fromCol)),
        xmlElement("xdr:colOff", undefined, "0"),
        xmlElement("xdr:row", undefined, String(fromRow)),
        xmlElement("xdr:rowOff", undefined, "0"),
      ]);

      const toElement = xmlElement("xdr:to", undefined, [
        xmlElement("xdr:col", undefined, String(toCol)),
        xmlElement("xdr:colOff", undefined, "0"),
        xmlElement("xdr:row", undefined, String(toRow)),
        xmlElement("xdr:rowOff", undefined, "0"),
      ]);

      const cNvPrAttrs: Record<string, string | number> = {
        id: shapeId++,
        name: `TextBox ${t + 1}`,
      };
      if (tb.title) cNvPrAttrs.title = tb.title;
      if (tb.altText) cNvPrAttrs.descr = tb.altText;

      const nvSpPr = xmlElement("xdr:nvSpPr", undefined, [
        xmlSelfClose("xdr:cNvPr", cNvPrAttrs),
        xmlElement("xdr:cNvSpPr", { txBox: 1 }, []),
      ]);

      // Shape properties
      const spPrChildren: string[] = [
        xmlElement("a:xfrm", undefined, [
          xmlSelfClose("a:off", { x: 0, y: 0 }),
          xmlSelfClose("a:ext", { cx: widthEmu, cy: heightEmu }),
        ]),
        xmlElement("a:prstGeom", { prst: "rect" }, [xmlSelfClose("a:avLst")]),
      ];

      // Fill
      const fillColor = tb.style?.fillColor ?? "FFFFFF";
      spPrChildren.push(
        xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: fillColor })]),
      );

      // Border
      const borderColor = tb.style?.borderColor ?? "000000";
      spPrChildren.push(
        xmlElement("a:ln", undefined, [
          xmlElement("a:solidFill", undefined, [xmlSelfClose("a:srgbClr", { val: borderColor })]),
        ]),
      );

      const spPr = xmlElement("xdr:spPr", undefined, spPrChildren);

      // Text body
      const fontSize = tb.style?.fontSize ?? 11;
      const fontSizeHundredths = fontSize * 100; // DrawingML uses hundredths of a point

      const rPrAttrs: Record<string, string | number> = { lang: "en-US", sz: fontSizeHundredths };
      if (tb.style?.bold) {
        rPrAttrs["b"] = 1;
      }

      const rPrChildren: string[] = [];
      if (tb.style?.color) {
        rPrChildren.push(
          xmlElement("a:solidFill", undefined, [
            xmlSelfClose("a:srgbClr", { val: tb.style.color }),
          ]),
        );
      }

      const rPr =
        rPrChildren.length > 0
          ? xmlElement("a:rPr", rPrAttrs, rPrChildren)
          : xmlSelfClose("a:rPr", rPrAttrs);

      const txBody = xmlElement("xdr:txBody", undefined, [
        xmlSelfClose("a:bodyPr", { wrap: "square" }),
        xmlElement("a:p", undefined, [
          xmlElement("a:r", undefined, [rPr, xmlElement("a:t", undefined, xmlEscape(tb.text))]),
        ]),
      ]);

      const sp = xmlElement("xdr:sp", undefined, [nvSpPr, spPr, txBody]);

      const anchor = xmlElement("xdr:twoCellAnchor", undefined, [
        fromElement,
        toElement,
        sp,
        xmlSelfClose("xdr:clientData"),
      ]);

      anchorElements.push(anchor);
    }
  }

  // ── Chart Anchors ──
  // Each chart becomes an <xdr:graphicFrame> with a <c:chart> reference
  // pointing at a relationship in this drawing's rels file. Charts get
  // their own rId namespace inside the drawing — we continue the rId
  // counter past the image relationships.
  const chartList: { globalChartIndex: number }[] = [];
  if (charts && charts.length > 0) {
    let shapeId = images.length + (textBoxes?.length ?? 0) + 2;

    for (let c = 0; c < charts.length; c++) {
      const chart = charts[c];
      const globalChartIndex = chartStartIndex + c;
      // Drawing-local rId. Image rIds are 1..images.length, so charts
      // start at images.length + 1.
      const rId = `rId${images.length + c + 1}`;

      relElements.push(
        xmlSelfClose("Relationship", {
          Id: rId,
          Type: REL_CHART,
          Target: `../charts/chart${globalChartIndex}.xml`,
        }),
      );

      const fromCol = chart.anchor.from.col;
      const fromRow = chart.anchor.from.row;
      const toCol = chart.anchor.to?.col ?? fromCol + 8;
      const toRow = chart.anchor.to?.row ?? fromRow + 15;

      const fromElement = xmlElement("xdr:from", undefined, [
        xmlElement("xdr:col", undefined, String(fromCol)),
        xmlElement("xdr:colOff", undefined, "0"),
        xmlElement("xdr:row", undefined, String(fromRow)),
        xmlElement("xdr:rowOff", undefined, "0"),
      ]);

      const toElement = xmlElement("xdr:to", undefined, [
        xmlElement("xdr:col", undefined, String(toCol)),
        xmlElement("xdr:colOff", undefined, "0"),
        xmlElement("xdr:row", undefined, String(toRow)),
        xmlElement("xdr:rowOff", undefined, "0"),
      ]);

      const cNvPrAttrs: Record<string, string | number> = {
        id: shapeId++,
        name: `Chart ${c + 1}`,
      };
      if (chart.frameTitle) cNvPrAttrs.title = chart.frameTitle;
      if (chart.altText) cNvPrAttrs.descr = chart.altText;

      const nvGraphicFramePr = xmlElement("xdr:nvGraphicFramePr", undefined, [
        xmlSelfClose("xdr:cNvPr", cNvPrAttrs),
        xmlElement("xdr:cNvGraphicFramePr", undefined, []),
      ]);

      const xfrm = xmlElement("xdr:xfrm", undefined, [
        xmlSelfClose("a:off", { x: 0, y: 0 }),
        xmlSelfClose("a:ext", { cx: 0, cy: 0 }),
      ]);

      const graphic = xmlElement(
        "a:graphic",
        undefined,
        xmlElement("a:graphicData", { uri: NS_C }, [
          xmlSelfClose("c:chart", { "xmlns:c": NS_C, "xmlns:r": NS_R, "r:id": rId }),
        ]),
      );

      const graphicFrame = xmlElement("xdr:graphicFrame", { macro: "" }, [
        nvGraphicFramePr,
        xfrm,
        graphic,
      ]);

      const anchor = xmlElement("xdr:twoCellAnchor", { editAs: "oneCell" }, [
        fromElement,
        toElement,
        graphicFrame,
        xmlSelfClose("xdr:clientData"),
      ]);

      anchorElements.push(anchor);
      chartList.push({ globalChartIndex });
    }
  }

  // Build drawing XML
  const drawingXml = xmlDocument(
    "xdr:wsDr",
    {
      "xmlns:xdr": NS_XDR,
      "xmlns:a": NS_A,
      "xmlns:r": NS_R,
    },
    anchorElements,
  );

  // Build drawing relationships XML
  const drawingRels = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, relElements);

  return {
    drawingXml,
    drawingRels,
    images: drawingImages,
    charts: chartList,
  };
}
