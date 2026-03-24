export { parseXml, parseSax, decodeOoxmlEscapes } from "./parser";
export type { XmlElement, XmlNode, SaxHandlers } from "./parser";
export {
  xmlElement,
  xmlSelfClose,
  xmlEscape,
  xmlEscapeAttr,
  xmlDeclaration,
  xmlDocument,
} from "./writer";
export type { XmlWriterOptions } from "./writer";
