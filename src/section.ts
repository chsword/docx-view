import type { Document, Element } from '@xmldom/xmldom';
import type { SectionInfo, SectionType } from './types.js';
import { WORD_NS, children } from './xml.js';

export const SECTION_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc',
  'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
  'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange',
] as const;

const SECTION_TYPES = new Set<SectionType>(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn']);

function wordNumber(element: Element, name: string): number | undefined {
  const raw = element.getAttributeNS(WORD_NS, name);
  if (!raw || !/^-?\d+$/.test(raw)) return undefined;
  return Number(raw);
}

function sectionType(sectPr: Element): SectionType {
  const value = children(sectPr, 'type')[0]?.getAttributeNS(WORD_NS, 'val') ?? '';
  return SECTION_TYPES.has(value as SectionType) ? value as SectionType : 'nextPage';
}

function pageSetup(sectPr: Element): Pick<SectionInfo, 'pageWidth' | 'pageHeight' | 'orientation' | 'margins' | 'columns'> {
  const size = children(sectPr, 'pgSz')[0];
  const margins = children(sectPr, 'pgMar')[0];
  const cols = children(sectPr, 'cols')[0];
  const pageWidth = Math.max(0, wordNumber(size ?? sectPr, 'w') ?? 11906);
  const pageHeight = Math.max(0, wordNumber(size ?? sectPr, 'h') ?? 16838);
  const orientation = (size?.getAttributeNS(WORD_NS, 'orient') ?? '') === 'landscape' ? 'landscape' : 'portrait';
  const count = Math.max(1, wordNumber(cols ?? sectPr, 'num') ?? 1);
  const equalWidth = !['0', 'false', 'off'].includes((cols?.getAttributeNS(WORD_NS, 'equalWidth') ?? '').toLowerCase());
  const widths = children(cols ?? sectPr, 'col').map(col => Math.max(0, wordNumber(col, 'w') ?? 0));
  return {
    pageWidth,
    pageHeight,
    orientation,
    margins: {
      top: Math.max(0, wordNumber(margins ?? sectPr, 'top') ?? 0),
      right: Math.max(0, wordNumber(margins ?? sectPr, 'right') ?? 0),
      bottom: Math.max(0, wordNumber(margins ?? sectPr, 'bottom') ?? 0),
      left: Math.max(0, wordNumber(margins ?? sectPr, 'left') ?? 0),
      header: Math.max(0, wordNumber(margins ?? sectPr, 'header') ?? 0),
      footer: Math.max(0, wordNumber(margins ?? sectPr, 'footer') ?? 0),
      gutter: Math.max(0, wordNumber(margins ?? sectPr, 'gutter') ?? 0),
    },
    columns: {
      count,
      space: Math.max(0, wordNumber(cols ?? sectPr, 'space') ?? 720),
      equalWidth,
      widths: widths.length ? widths : undefined,
    },
  };
}

function pageNumbering(sectPr: Element): SectionInfo['pageNumbering'] {
  const node = children(sectPr, 'pgNumType')[0];
  if (!node) return undefined;
  const start = wordNumber(node, 'start');
  const format = node.getAttributeNS(WORD_NS, 'fmt') ?? undefined;
  return start === undefined && !format ? undefined : { start, format };
}

function references(sectPr: Element, name: 'headerReference' | 'footerReference', resolveRelationship: (id: string) => string | undefined):
SectionInfo['headers'] {
  const result: SectionInfo['headers'] = {};
  for (const reference of children(sectPr, name)) {
    const type = reference.getAttributeNS(WORD_NS, 'type');
    const id = reference.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
      ?? reference.getAttribute('r:id');
    if (!id || !type || !['default', 'first', 'even'].includes(type)) continue;
    const part = resolveRelationship(id);
    if (part) result[type as 'default' | 'first' | 'even'] = part;
  }
  return result;
}

export interface SectionDescriptor {
  startParagraph: number;
  endParagraph: number;
  sectPr: Element;
  source: 'paragraph' | 'body';
  paragraph?: Element;
}

export function collectSections(mainDocument: Document): SectionDescriptor[] {
  const body = children(mainDocument.documentElement!, 'body')[0]!;
  const paragraphs = Array.from(body.getElementsByTagNameNS(WORD_NS, 'p'));
  const sections: SectionDescriptor[] = [];
  let start = 0;
  for (const [index, paragraph] of paragraphs.entries()) {
    const sectPr = children(children(paragraph, 'pPr')[0] ?? paragraph, 'sectPr')[0];
    if (!sectPr) continue;
    sections.push({ startParagraph: start, endParagraph: index, sectPr, source: 'paragraph', paragraph });
    start = index + 1;
  }
  const tail = children(body, 'sectPr')[0];
  if (tail) sections.push({ startParagraph: start, endParagraph: Math.max(start, paragraphs.length) - 1, sectPr: tail, source: 'body' });
  return sections;
}

export function readSections(mainDocument: Document, resolveRelationship: (id: string) => string | undefined): SectionInfo[] {
  return collectSections(mainDocument).map((section, index) => ({
    index,
    startParagraph: section.startParagraph,
    endParagraph: section.endParagraph,
    type: sectionType(section.sectPr),
    ...pageSetup(section.sectPr),
    pageNumbering: pageNumbering(section.sectPr),
    titlePage: !!children(section.sectPr, 'titlePg')[0],
    headers: references(section.sectPr, 'headerReference', resolveRelationship),
    footers: references(section.sectPr, 'footerReference', resolveRelationship),
  }));
}
