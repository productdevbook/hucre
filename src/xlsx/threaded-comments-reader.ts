// ── Threaded Comments Reader ──────────────────────────────────────
// Parses Excel 365's modern comment system:
//   xl/persons/person.xml             — workbook-wide person directory
//   xl/threadedComments/threadedComment{N}.xml — per-sheet comment threads
//
// Schema: http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments
// Reference: [MS-XLSX] Threaded Comments
//   https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940

import type { ThreadedComment, ThreadedCommentMention, ThreadedCommentPerson } from "../_types";
import { parseXml } from "../xml/parser";
import type { XmlElement, XmlNode } from "../xml/parser";

/** Parse `xl/persons/person.xml` into a list of persons. */
export function parsePersons(xml: string): ThreadedCommentPerson[] {
  const root = parseXml(xml);
  const persons: ThreadedCommentPerson[] = [];
  for (const child of childElements(root)) {
    if (child.local !== "person") continue;
    const id = child.attrs.id;
    const displayName = child.attrs.displayName;
    if (!id || displayName === undefined) continue;
    const entry: ThreadedCommentPerson = { id, displayName };
    if (child.attrs.userId) entry.userId = child.attrs.userId;
    if (child.attrs.providerId) entry.providerId = child.attrs.providerId;
    persons.push(entry);
  }
  return persons;
}

/** Parse a single `xl/threadedComments/threadedCommentN.xml` part. */
export function parseThreadedComments(xml: string): ThreadedComment[] {
  const root = parseXml(xml);
  const comments: ThreadedComment[] = [];
  for (const child of childElements(root)) {
    if (child.local !== "threadedComment") continue;
    const id = child.attrs.id;
    const personId = child.attrs.personId;
    if (!id || !personId) continue;

    const entry: ThreadedComment = {
      id,
      personId,
      text: readChildText(child, "text"),
    };
    if (child.attrs.ref) entry.ref = child.attrs.ref;
    if (child.attrs.parentId) entry.parentId = child.attrs.parentId;
    if (child.attrs.dT) entry.date = child.attrs.dT;
    if (child.attrs.done === "1" || child.attrs.done === "true") entry.done = true;

    const mentionsEl = findChild(child, "mentions");
    if (mentionsEl) {
      const mentions: ThreadedCommentMention[] = [];
      for (const m of childElements(mentionsEl)) {
        if (m.local !== "mention") continue;
        const mp = m.attrs.mentionpersonId;
        const mid = m.attrs.mentionId;
        if (!mp || !mid) continue;
        mentions.push({
          mentionPersonId: mp,
          mentionId: mid,
          startIndex: parseIntSafe(m.attrs.startIndex, 0),
          length: parseIntSafe(m.attrs.length, 0),
        });
      }
      if (mentions.length > 0) entry.mentions = mentions;
    }

    comments.push(entry);
  }
  return comments;
}

// ── Internals ─────────────────────────────────────────────────────

function childElements(el: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string") out.push(c);
  }
  return out;
}

function findChild(el: XmlElement, localName: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c !== "string" && c.local === localName) return c;
  }
  return undefined;
}

function readChildText(el: XmlElement, localName: string): string {
  const child = findChild(el, localName);
  if (!child) return "";
  let text = "";
  for (const c of child.children as XmlNode[]) {
    if (typeof c === "string") text += c;
  }
  return text;
}

function parseIntSafe(s: string | undefined, fallback: number): number {
  if (s === undefined) return fallback;
  const n = parseInt(s, 10);
  return Number.isNaN(n) ? fallback : n;
}
