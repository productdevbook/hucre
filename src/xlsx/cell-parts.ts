// Links and comments live outside <sheetData>. Retain only this metadata,
// never its source rows, and share the buffered package serializers.
import type { Cell } from "../_types"
import { createCellStore, setCell } from "../cell-store"
import { xmlDocument, xmlSelfClose } from "../xml/writer"
import { createHyperlinkCollector } from "./worksheet-writer"
import { writeComments } from "./comments-writer"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
const REL_HYPERLINK = `${REL}hyperlink`

export class SheetCellParts {
  private links = createHyperlinkCollector()
  private comments = createCellStore<Partial<Cell>>()

  get hasComments(): boolean {
    return this.comments.size > 0
  }

  add(row: number, col: number, cell: Partial<Cell>): void {
    if (cell.hyperlink) this.links.add(row, col, cell.hyperlink)
    if (cell.comment) {
      // Incremental callers may reuse their cell object before finish().
      setCell(this.comments, row, col, { comment: structuredClone(cell.comment) })
    }
  }

  toXml(): string {
    return (
      this.links.toXml() +
      (this.hasComments ? xmlSelfClose("legacyDrawing", { "r:id": "rIdVML" }) : "")
    )
  }

  *entries(index: number): Generator<{ path: string; xml: string }> {
    const relationships = this.links.relationships.map((link) =>
      xmlSelfClose("Relationship", {
        Id: link.id,
        Type: REL_HYPERLINK,
        Target: link.target,
        TargetMode: "External",
      }),
    )
    const comments = writeComments(this.comments, index - 1)
    if (comments) {
      yield { path: `xl/comments${index}.xml`, xml: comments.commentsXml }
      yield { path: `xl/drawings/vmlDrawing${index}.vml`, xml: comments.vmlXml }
      for (const [id, type, target] of [
        ["rIdVML", "vmlDrawing", `../drawings/vmlDrawing${index}.vml`],
        ["rIdComments", "comments", `../comments${index}.xml`],
      ]) {
        relationships.push(
          xmlSelfClose("Relationship", { Id: id, Type: `${REL}${type}`, Target: target }),
        )
      }
    }
    if (relationships.length) {
      yield {
        path: `xl/worksheets/_rels/sheet${index}.xml.rels`,
        xml: xmlDocument(
          "Relationships",
          { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
          relationships,
        ),
      }
    }
  }
}
