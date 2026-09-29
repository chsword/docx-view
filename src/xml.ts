import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element, Node } from '@xmldom/xmldom';

export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
export const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const TRANSPARENT_WORD_WRAPPERS = new Set(['sdt', 'sdtContent', 'customXml', 'ins', 'del', 'moveFrom', 'moveTo']);
export const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const MAX_XML_TEXT_LENGTH = 1_000_000;
const INVALID_XML_TEXT_RE = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff\ud800-\udfff]/u;

export function parseXml(xml: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new Error('DTD and entity declarations are not supported.');
  }
  return new DOMParser({
    onError: (level: string, message: string) => { throw new Error(`Invalid XML (${level}): ${message}`); },
  }).parseFromString(xml, 'application/xml');
}

export function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

export function selectAlternateContentBranch(element: Element): Element | undefined {
  const choices = Array.from(element.getElementsByTagNameNS(MC_NS, 'Choice'));
  return choices.find((choice) => (choice.getAttribute('Requires') ?? '').split(/\s+/).includes('wps'))
    ?? choices[0]
    ?? Array.from(element.getElementsByTagNameNS(MC_NS, 'Fallback'))[0];
}

export function children(node: Node, localName?: string, namespace = WORD_NS): Element[] {
  const result: Element[] = [];
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 1) {
      const element = child as Element;
      if ((!localName || element.localName === localName) && element.namespaceURI === namespace) {
        result.push(element);
      }
    }
  }
  return result;
}

export function isTransparentWordWrapper(element: Element): boolean {
  return element.namespaceURI === WORD_NS && TRANSPARENT_WORD_WRAPPERS.has(element.localName ?? '');
}

export function childrenThroughTransparent(node: Node, localName: string, namespace = WORD_NS): Element[] {
  const result: Element[] = [];
  for (let child = node.firstChild; child; child = child.nextSibling) {
    if (child.nodeType !== 1) continue;
    const element = child as Element;
    if (element.namespaceURI !== namespace) continue;
    if (element.localName === localName) {
      result.push(element);
    } else if (isTransparentWordWrapper(element)) {
      result.push(...childrenThroughTransparent(element, localName, namespace));
    }
  }
  return result;
}

export function descendants(node: Element | Document, localName: string): Element[] {
  const result: Element[] = [];
  const stack: (Node | null)[] = [node.firstChild];
  while (stack.length) {
    const current = stack[stack.length - 1]!;
    if (!current) {
      stack.pop();
      continue;
    }
    stack[stack.length - 1] = current.nextSibling;
    if (current.nodeType === 1) {
      const element = current as Element;
      if (element.namespaceURI === WORD_NS && (localName === '*' || element.localName === localName)) result.push(element);
    }
    if (current.firstChild) stack.push(current.firstChild);
  }
  return result;
}

export function wordElement(document: Document, name: string): Element {
  return document.createElementNS(WORD_NS, `w:${name}`);
}

export function wordValue(element: Element | undefined): string | undefined {
  return element?.getAttributeNS(WORD_NS, 'val') ?? undefined;
}

export function setWordValue(element: Element, value: string): void {
  element.setAttributeNS(WORD_NS, 'w:val', value);
}

export function validatePath(path: string): void {
  if (typeof path !== 'string' || !path || path.startsWith('/') || /[\\\u0000-\u001f]/.test(path) ||
      path.split('/').some(segment => !segment || segment === '.' || segment === '..')) {
    throw new Error(`Invalid package part path: ${path}`);
  }
}

export function assertText(text: unknown, name = 'text'): asserts text is string {
  if (typeof text !== 'string' || text.length > MAX_XML_TEXT_LENGTH ||
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/u.test(text) ||
      /[\ud800-\udfff]/u.test(text)) {
    throw new Error(`${name} must be valid XML text of at most 1,000,000 characters.`);
  }
}

export function sanitizeTextWithInfo(text: string): { text: string; truncated: boolean; truncatedAt?: number } {
  if (!INVALID_XML_TEXT_RE.test(text)) {
    if (text.length <= MAX_XML_TEXT_LENGTH) return { text, truncated: false };
    let end = MAX_XML_TEXT_LENGTH;
    const tail = text.charCodeAt(end - 1);
    const next = text.charCodeAt(end);
    if (tail >= 0xd800 && tail <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) end--;
    return { text: text.slice(0, end), truncated: true, truncatedAt: end };
  }
  let result = '';
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index);
    if ((code >= 0x0000 && code <= 0x0008) || code === 0x000b || code === 0x000c ||
        (code >= 0x000e && code <= 0x001f) || code === 0xfffe || code === 0xffff) continue;
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(index + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        if (result.length + 2 > MAX_XML_TEXT_LENGTH) return { text: result, truncated: true, truncatedAt: result.length };
        result += text.slice(index, index + 2);
        index++;
      }
      continue;
    }
    if (code >= 0xdc00 && code <= 0xdfff) continue;
    if (result.length + 1 > MAX_XML_TEXT_LENGTH) return { text: result, truncated: true, truncatedAt: result.length };
    result += text.charAt(index);
  }
  return { text: result, truncated: false };
}

export function sanitizeText(text: string): string {
  return sanitizeTextWithInfo(text).text;
}

export function isValidXmlCharCode(value: number): boolean {
  return Number.isSafeInteger(value) &&
    (value === 0x9 || value === 0xa || value === 0xd ||
      (value >= 0x20 && value <= 0xd7ff) ||
      (value >= 0xe000 && value <= 0xfffd));
}
