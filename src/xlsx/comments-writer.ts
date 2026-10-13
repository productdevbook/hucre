import { cellEntries } from "../cell-store"
import type { CellStore } from "../_types"
// ── Comments & VML Writer ─────────────────────────────────────────────
// Generates xl/commentsN.xml and xl/drawings/vmlDrawingN.vml for XLSX.

import type { Cell, CellComment } from "../_types"
import { xmlDocument, xmlElement, xmlEscape } from "../xml/writer"
import { cellRef, serializeRichTextRuns } from "./worksheet-writer"

// ── Types ────────────────────────────────────────────────────────────

export interface CommentsResult {
  commentsXml: string
  vmlXml: string
  comments: CommentEntry[]
}

interface CommentEntry extends CellComment {
  ref: string
  row: number
  col: number
  author: string
}

// ── Constants ────────────────────────────────────────────────────────

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// ── Main Writer ──────────────────────────────────────────────────────

/**
 * Collect all cells with comments and generate comments.xml + VML drawing.
 * Returns null if no cells have comments.
 */
export function writeComments(
  cells: CellStore<Partial<Cell>>,
  sheetIndex: number,
): CommentsResult | null {
  // Collect comments from cells
  const commentEntries: CommentEntry[] = []

  for (const [row, col, cell] of cellEntries(cells)) {
    if (!cell.comment) continue

    commentEntries.push({
      ...cell.comment,
      ref: cellRef(row, col),
      row,
      col,
      author: cell.comment.author ?? "",
    })
  }

  if (commentEntries.length === 0) return null

  // Sort by row then column for deterministic output
  commentEntries.sort((a, b) => a.row - b.row || a.col - b.col)

  // Build author list (unique, preserving insertion order)
  const authorMap = new Map<string, number>()
  for (const entry of commentEntries) {
    if (!authorMap.has(entry.author)) {
      authorMap.set(entry.author, authorMap.size)
    }
  }

  return {
    commentsXml: buildCommentsXml(commentEntries, authorMap),
    vmlXml: buildVmlDrawing(commentEntries, sheetIndex),
    comments: commentEntries,
  }
}

// ── Comments XML Builder ─────────────────────────────────────────────

function buildCommentsXml(entries: CommentEntry[], authorMap: Map<string, number>): string {
  // Build <authors> section
  const authorsXml = xmlElement(
    "authors",
    undefined,
    [...authorMap.keys()].map((author) => xmlElement("author", undefined, xmlEscape(author))),
  )

  // Build <commentList> section
  const commentElements = entries.map((entry) => {
    const authorId = authorMap.get(entry.author) ?? 0
    // Notes use the same run encoding as cells, including font properties
    // and xml:space. Flattening here silently discarded comment.richText.
    const textXml = xmlElement(
      "text",
      undefined,
      serializeRichTextRuns(entry.richText?.length ? entry.richText : [{ text: entry.text }]),
    )
    return xmlElement("comment", { ref: entry.ref, authorId }, textXml)
  })
  const commentListXml = xmlElement("commentList", undefined, commentElements)

  return xmlDocument("comments", { xmlns: NS_SPREADSHEET }, [authorsXml, commentListXml])
}

// ── VML Drawing Builder ─────────────────────────────────────────────

function buildVmlDrawing(
  entries: Array<{ ref: string; row: number; col: number }>,
  sheetIndex: number,
): string {
  // Fold static markup into one prefix instead of retaining each fragment.
  const parts: string[] = [
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"' +
      ' xmlns:o="urn:schemas-microsoft-com:office:office"' +
      ' xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      '<o:shapelayout v:ext="edit">' +
      `<o:idmap v:ext="edit" data="${sheetIndex + 1}"/>` +
      "</o:shapelayout>" +
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"' +
      ' path="m,l,21600r21600,l21600,xe">' +
      '<v:stroke joinstyle="miter"/>' +
      '<v:path gradientshapeok="t" o:connecttype="rect"/>' +
      "</v:shapetype>",
  ]

  // One fragment per note keeps VML construction from retaining a dozen
  // attribute/text fragments for every comment in a large streamed sheet.
  // Generate a shape for each comment
  const baseShapeId = (sheetIndex + 1) * 1024 + 1
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i]
    const shapeId = baseShapeId + i

    // Calculate anchor position:
    // Anchor format: leftCol, leftColOffset, topRow, topRowOffset, rightCol, rightColOffset, bottomRow, bottomRowOffset
    // Position the comment box to the right of and below the cell
    const anchorCol = entry.col + 1
    const anchorRow = entry.row
    const rightCol = anchorCol + 2
    const bottomRow = anchorRow + 4

    // Calculate margin-left based on column position (approximate: 48pt per column)
    const marginLeft = (entry.col + 1) * 48
    const marginTop = entry.row * 15

    parts.push(
      `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202"` +
        ` style="position:absolute;margin-left:${marginLeft}pt;margin-top:${marginTop}pt;` +
        `width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden"` +
        ` fillcolor="#ffffe1" o:insetmode="auto">` +
        '<v:fill color2="#ffffe1"/>' +
        '<v:shadow on="t" color="black" obscured="t"/>' +
        '<v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto">' +
        '<div style="text-align:left"/>' +
        "</v:textbox>" +
        '<x:ClientData ObjectType="Note">' +
        "<x:MoveWithCells/>" +
        "<x:SizeWithCells/>" +
        `<x:Anchor>${anchorCol},15,${anchorRow},2,${rightCol},31,${bottomRow},4</x:Anchor>` +
        "<x:AutoFill>False</x:AutoFill>" +
        `<x:Row>${entry.row}</x:Row>` +
        `<x:Column>${entry.col}</x:Column>` +
        "</x:ClientData>" +
        "</v:shape>",
    )
  }

  parts.push("</xml>")

  return parts.join("")
}
