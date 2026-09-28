import type { Document, Element } from '@xmldom/xmldom';
import { WORD_NS, children, wordElement } from './xml.js';

export const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
export const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
export const COMMENTS_EXTENDED_TYPE = 'application/vnd.ms-word.commentsExtended+xml';
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

export interface ParsedCommentEntry {
  id: number;
  element: Element;
  paraId?: string;
}

export interface ParsedCommentExEntry {
  paraId: string;
  paraIdParent?: string;
  done?: boolean;
  element: Element;
}

export function defaultCommentsXml(): string {
  return `<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}"/>`;
}

export function defaultCommentsExtendedXml(): string {
  return `<w15:commentsEx xmlns:w15="${W15_NS}"/>`;
}

export function commentReferenceStyle(): 'CommentReference' {
  return 'CommentReference';
}

export function commentParagraphStyle(): 'CommentText' {
  return 'CommentText';
}

export function parseCommentEntries(document: Document | null): ParsedCommentEntry[] {
  if (!document?.documentElement || document.documentElement.namespaceURI !== WORD_NS || document.documentElement.localName !== 'comments') {
    return [];
  }
  return children(document.documentElement, 'comment')
    .map((element): ParsedCommentEntry | null => {
      const id = Number(element.getAttributeNS(WORD_NS, 'id') ?? element.getAttribute('w:id'));
      if (!Number.isSafeInteger(id) || id < 0) return null;
      const paragraphs = children(element, 'p');
      const lastParagraph = paragraphs.at(-1);
      const paraId = lastParagraph?.getAttributeNS(W14_NS, 'paraId') ?? lastParagraph?.getAttribute('w14:paraId') ?? undefined;
      return { id, element, paraId };
    })
    .filter((entry): entry is ParsedCommentEntry => Boolean(entry));
}

export function parseCommentExEntries(document: Document | null): ParsedCommentExEntry[] {
  if (!document?.documentElement || document.documentElement.namespaceURI !== W15_NS || document.documentElement.localName !== 'commentsEx') {
    return [];
  }
  return children(document.documentElement, 'commentEx', W15_NS)
    .map((element): ParsedCommentExEntry | null => {
      const paraId = element.getAttributeNS(W15_NS, 'paraId') ?? element.getAttribute('w15:paraId') ?? '';
      if (!/^[0-9a-f]{8}$/i.test(paraId)) return null;
      const parent = element.getAttributeNS(W15_NS, 'paraIdParent') ?? element.getAttribute('w15:paraIdParent') ?? undefined;
      const doneValue = element.getAttributeNS(W15_NS, 'done') ?? element.getAttribute('w15:done') ?? undefined;
      return {
        paraId: paraId.toUpperCase(),
        paraIdParent: /^[0-9a-f]{8}$/i.test(parent ?? '') ? parent!.toUpperCase() : undefined,
        done: doneValue === undefined ? undefined : ['1', 'true', 'on'].includes(doneValue),
        element,
      };
    })
    .filter((entry): entry is ParsedCommentExEntry => Boolean(entry));
}

export function ensureCommentParagraphParaId(paragraph: Element, paraId: string): void {
  paragraph.setAttributeNS(W14_NS, 'w14:paraId', paraId.toUpperCase());
}

export function makeCommentAnnotationRun(document: Document): Element {
  const run = wordElement(document, 'r');
  const props = wordElement(document, 'rPr');
  const style = wordElement(document, 'rStyle');
  style.setAttributeNS(WORD_NS, 'w:val', commentReferenceStyle());
  props.appendChild(style);
  run.appendChild(props);
  run.appendChild(wordElement(document, 'annotationRef'));
  return run;
}
