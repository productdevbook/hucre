// ── Comments & VML Writer ─────────────────────────────────────────────
// Generates xl/commentsN.xml and xl/drawings/vmlDrawingN.vml for XLSX.

import type { Cell } from "../_types"
import { xmlDocument, xmlElement, xmlEscape } from "../xml/writer"
import { cellRef } from "./worksheet-writer"

// ── Types ────────────────────────────────────────────────────────────

export interface CommentsResult {
  commentsXml: string
  vmlXml: string
  comments: Array<{ ref: string; row: number; col: number; author: string; text: string }>
}

// ── Constants ────────────────────────────────────────────────────────

const NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// ── Main Writer ──────────────────────────────────────────────────────

/**
 * Collect all cells with comments and generate comments.xml + VML drawing.
 * Returns null if no cells have comments.
 */
export function writeComments(
  cells: Map<string, Partial<Cell>>,
  sheetIndex: number,
): CommentsResult | null {
  // Collect comments from cells
  const commentEntries: Array<{
    ref: string
    row: number
    col: number
    author: string
    text: string
  }> = []

  for (const [key, cell] of cells) {
    if (!cell.comment) continue

    const [rowStr, colStr] = key.split(",")
    const row = parseInt(rowStr, 10)
    const col = parseInt(colStr, 10)
    const ref = cellRef(row, col)
    const author = cell.comment.author ?? ""
    const text = cell.comment.text

    commentEntries.push({ ref, row, col, author, text })
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

  // Generate comments.xml
  const commentsXml = buildCommentsXml(commentEntries, authorMap)

  // Generate VML drawing
  const vmlXml = buildVmlDrawing(commentEntries, sheetIndex)

  return { commentsXml, vmlXml, comments: commentEntries }
}

// ── Comments XML Builder ─────────────────────────────────────────────

function buildCommentsXml(
  entries: Array<{ ref: string; author: string; text: string }>,
  authorMap: Map<string, number>,
): string {
  // Build <authors> section
  const authorElements: string[] = []
  for (const [authorName] of authorMap) {
    authorElements.push(xmlElement("author", undefined, xmlEscape(authorName)))
  }
  const authorsXml = xmlElement("authors", undefined, authorElements)

  // Build <commentList> section
  const commentElements: string[] = []
  for (const entry of entries) {
    const authorId = authorMap.get(entry.author) ?? 0
    const textXml = xmlElement("text", undefined, [
      xmlElement("r", undefined, [xmlElement("t", undefined, xmlEscape(entry.text))]),
    ])
    commentElements.push(xmlElement("comment", { ref: entry.ref, authorId }, [textXml]))
  }
  const commentListXml = xmlElement("commentList", undefined, commentElements)

  return xmlDocument("comments", { xmlns: NS_SPREADSHEET }, [authorsXml, commentListXml])
}

// ── VML Drawing Builder ─────────────────────────────────────────────

function buildVmlDrawing(
  entries: Array<{ ref: string; row: number; col: number }>,
  sheetIndex: number,
): string {
  const parts: string[] = []

  // XML prologue (VML is not standard XML — uses a custom <xml> root)
  parts.push(
    '<xml xmlns:v="urn:schemas-microsoft-com:vml"',
    ' xmlns:o="urn:schemas-microsoft-com:office:office"',
    ' xmlns:x="urn:schemas-microsoft-com:office:excel">',
  )

  // Shape layout
  parts.push(
    '<o:shapelayout v:ext="edit">',
    `<o:idmap v:ext="edit" data="${sheetIndex + 1}"/>`,
    "</o:shapelayout>",
  )

  // Shape type definition (standard comment shape type)
  parts.push(
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"',
    ' path="m,l,21600r21600,l21600,xe">',
    '<v:stroke joinstyle="miter"/>',
    '<v:path gradientshapeok="t" o:connecttype="rect"/>',
    "</v:shapetype>",
  )

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
      `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202"`,
      ` style="position:absolute;margin-left:${marginLeft}pt;margin-top:${marginTop}pt;`,
      `width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden"`,
      ` fillcolor="#ffffe1" o:insetmode="auto">`,
      '<v:fill color2="#ffffe1"/>',
      '<v:shadow on="t" color="black" obscured="t"/>',
      '<v:path o:connecttype="none"/>',
      '<v:textbox style="mso-direction-alt:auto">',
      '<div style="text-align:left"/>',
      "</v:textbox>",
      '<x:ClientData ObjectType="Note">',
      "<x:MoveWithCells/>",
      "<x:SizeWithCells/>",
      `<x:Anchor>${anchorCol},15,${anchorRow},2,${rightCol},31,${bottomRow},4</x:Anchor>`,
      "<x:AutoFill>False</x:AutoFill>",
      `<x:Row>${entry.row}</x:Row>`,
      `<x:Column>${entry.col}</x:Column>`,
      "</x:ClientData>",
      "</v:shape>",
    )
  }

  parts.push("</xml>")

  return parts.join("")
}
