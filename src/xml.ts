import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element, Node } from '@xmldom/xmldom';

export const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
export const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
export const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function parseXml(xml: string): Document {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new Error('DTD and entity declarations are not supported.');
  }
  return new DOMParser({
    onError: (level, message) => { throw new Error(`Invalid XML (${level}): ${message}`); },
  }).parseFromString(xml, 'application/xml');
}

export function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
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

export function descendants(node: Element | Document, localName: string): Element[] {
  return Array.from(node.getElementsByTagNameNS(WORD_NS, localName));
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
  if (typeof text !== 'string' || text.length > 1_000_000 ||
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/u.test(text) ||
      /[\ud800-\udfff]/u.test(text)) {
    throw new Error(`${name} must be valid XML text of at most 1,000,000 characters.`);
  }
}
