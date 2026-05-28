// ── Comments Reader ──────────────────────────────────────────────────
// Parses xl/commentsN.xml into a map of cell references to CellComment.

import type { CellComment } from "../_types"
import { parseSax } from "../xml/parser"

/**
 * Parse xl/commentsN.xml and return a map of cell reference (e.g. "A1")
 * to CellComment objects.
 */
export function parseComments(xml: string): Map<string, CellComment> {
  const comments = new Map<string, CellComment>()
  const authors: string[] = []

  // SAX parsing state
  let inAuthors = false
  let inAuthor = false
  let inCommentList = false
  let inComment = false
  let inText = false
  let _inR = false
  let inT = false

  let authorText = ""
  let currentRef = ""
  let currentAuthorId = -1
  let currentText = ""

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag

      switch (local) {
        case "authors":
          inAuthors = true
          break
        case "author":
          if (inAuthors) {
            inAuthor = true
            authorText = ""
          }
          break
        case "commentList":
          inCommentList = true
          break
        case "comment":
          if (inCommentList) {
            inComment = true
            currentRef = attrs["ref"] ?? ""
            currentAuthorId = attrs["authorId"] ? Number(attrs["authorId"]) : -1
            currentText = ""
          }
          break
        case "text":
          if (inComment) {
            inText = true
          }
          break
        case "r":
          if (inText) {
            _inR = true
          }
          break
        case "t":
          if (inText) {
            inT = true
          }
          break
      }
    },

    onText(text) {
      if (inAuthor) {
        authorText += text
      } else if (inT) {
        currentText += text
      }
    },

    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag

      switch (local) {
        case "authors":
          inAuthors = false
          break
        case "author":
          if (inAuthor) {
            authors.push(authorText)
            inAuthor = false
          }
          break
        case "commentList":
          inCommentList = false
          break
        case "comment":
          if (inComment && currentRef) {
            const comment: CellComment = { text: currentText }
            if (currentAuthorId >= 0 && currentAuthorId < authors.length) {
              const authorName = authors[currentAuthorId]
              if (authorName) {
                comment.author = authorName
              }
            }
            comments.set(currentRef, comment)
            inComment = false
          }
          break
        case "text":
          inText = false
          break
        case "r":
          _inR = false
          break
        case "t":
          inT = false
          break
      }
    },
  })

  return comments
}
