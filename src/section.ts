import type { Document, Element } from '@xmldom/xmldom';
import type { SectionInfo, SectionProperties, SectionType } from './types.js';
import { WORD_NS, children, descendants } from './xml.js';
import { readBorderSide } from './styles.js';

export const SECTION_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc',
  'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg',
  'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange',
] as const;

const SECTION_TYPES = new Set<SectionType>(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn']);
const DOC_GRID_TYPES = new Set<NonNullable<SectionInfo['docGrid']>['type']>([
  'default', 'lines', 'linesAndChars', 'snapToChars',
]);

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

function docGrid(sectPr: Element): SectionInfo['docGrid'] {
  const grid = children(sectPr, 'docGrid')[0];
  if (!grid) return undefined;
  const rawType = grid.getAttributeNS(WORD_NS, 'type') || 'default';
  const type = DOC_GRID_TYPES.has(rawType as NonNullable<SectionInfo['docGrid']>['type'])
    ? rawType as NonNullable<SectionInfo['docGrid']>['type']
    : 'default';
  const linePitch = wordNumber(grid, 'linePitch');
  const charSpace = wordNumber(grid, 'charSpace');
  return {
    type,
    ...(linePitch !== undefined ? { linePitch } : {}),
    ...(charSpace !== undefined ? { charSpace } : {}),
  };
}

function lineNumbering(sectPr: Element): SectionInfo['lineNumbering'] {
  const node = children(sectPr, 'lnNumType')[0];
  if (!node) return undefined;
  const countBy = wordNumber(node, 'countBy');
  const start = wordNumber(node, 'start');
  const distance = wordNumber(node, 'distance');
  const rawRestart = node.getAttributeNS(WORD_NS, 'restart') ?? '';
  const restart = ['continuous', 'newPage', 'newSection'].includes(rawRestart)
    ? rawRestart as NonNullable<SectionInfo['lineNumbering']>['restart'] : undefined;
  return {
    ...(countBy !== undefined ? { countBy } : {}),
    ...(start !== undefined ? { start } : {}),
    ...(distance !== undefined ? { distance } : {}),
    ...(restart !== undefined ? { restart } : {}),
  };
}

function pageBorders(sectPr: Element): SectionInfo['pageBorders'] {
  const node = children(sectPr, 'pgBorders')[0];
  if (!node) return undefined;
  const rawDisplay = node.getAttributeNS(WORD_NS, 'display') ?? '';
  const rawOffset = node.getAttributeNS(WORD_NS, 'offsetFrom') ?? '';
  const display = ['allPages', 'firstPage', 'notFirstPage'].includes(rawDisplay)
    ? rawDisplay as NonNullable<SectionInfo['pageBorders']>['display'] : undefined;
  const offsetFrom = ['page', 'text'].includes(rawOffset)
    ? rawOffset as NonNullable<SectionInfo['pageBorders']>['offsetFrom'] : undefined;
  const sides = {
    top: readBorderSide(children(node, 'top')[0]),
    left: readBorderSide(children(node, 'left')[0]),
    bottom: readBorderSide(children(node, 'bottom')[0]),
    right: readBorderSide(children(node, 'right')[0]),
  };
  return {
    ...(display !== undefined ? { display } : {}),
    ...(offsetFrom !== undefined ? { offsetFrom } : {}),
    ...Object.fromEntries(Object.entries(sides).filter(([, value]) => value !== undefined)),
  } as SectionInfo['pageBorders'];
}

function verticalAlignment(sectPr: Element): SectionInfo['verticalAlignment'] {
  const value = children(sectPr, 'vAlign')[0]?.getAttributeNS(WORD_NS, 'val') ?? '';
  return ['top', 'center', 'both', 'bottom'].includes(value)
    ? value as SectionInfo['verticalAlignment'] : undefined;
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
  const paragraphs = descendants(body, 'p').filter((paragraph) => {
    let ancestor = paragraph.parentNode as Element | null;
    while (ancestor && ancestor !== body) {
      if (ancestor.localName === 'txbxContent') return false;
      ancestor = ancestor.parentNode as Element | null;
    }
    return true;
  });
  const indices = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
  const boundaries: Array<{ index: number; paragraph: Element; sectPr: Element }> = [];
  const walk = (parent: Element, sectionScope: boolean): void => {
    for (const child of children(parent)) {
      if (child.localName === 'p') {
        const index = indices.get(child);
        if (index === undefined) continue;
        if (sectionScope) {
          const sectPr = children(children(child, 'pPr')[0] ?? child, 'sectPr')[0];
          if (sectPr) boundaries.push({ index, paragraph: child, sectPr });
        }
        continue;
      }
      if (child.localName === 'tbl') {
        for (const row of children(child, 'tr')) {
          for (const cell of children(row, 'tc')) walk(cell, false);
        }
        continue;
      }
      if (['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '')) walk(child, sectionScope);
    }
  };
  walk(body, true);
  const sections: SectionDescriptor[] = [];
  let start = 0;
  for (const boundary of boundaries) {
    sections.push({
      startParagraph: start,
      endParagraph: boundary.index,
      sectPr: boundary.sectPr,
      source: 'paragraph',
      paragraph: boundary.paragraph,
    });
    start = boundary.index + 1;
  }
  const tail = children(body, 'sectPr')[0];
  if (tail) {
    const endParagraph = start <= paragraphs.length - 1 ? paragraphs.length - 1 : start - 1;
    sections.push({ startParagraph: start, endParagraph, sectPr: tail, source: 'body' });
  }
  return sections;
}

/**
 * 一个 `sectPr` 自己的版面属性，不含它在文档里的位置与页眉页脚引用。`readSections()` 用它，
 * `w:sectPrChange` 里的旧值快照（CT_SectPrBase，本来就没有页眉页脚引用）也用它读。
 */
export function readSectionProperties(sectPr: Element): SectionProperties {
  const grid = docGrid(sectPr);
  const lines = lineNumbering(sectPr);
  const borders = pageBorders(sectPr);
  const alignment = verticalAlignment(sectPr);
  return {
    type: sectionType(sectPr),
    ...pageSetup(sectPr),
    ...(grid ? { docGrid: grid } : {}),
    pageNumbering: pageNumbering(sectPr),
    ...(lines ? { lineNumbering: lines } : {}),
    ...(borders ? { pageBorders: borders } : {}),
    ...(alignment ? { verticalAlignment: alignment } : {}),
    titlePage: !!children(sectPr, 'titlePg')[0],
  };
}

export function readSections(mainDocument: Document, resolveRelationship: (id: string) => string | undefined): SectionInfo[] {
  return collectSections(mainDocument).map((section, index) => ({
    index,
    startParagraph: section.startParagraph,
    endParagraph: section.endParagraph,
    ...readSectionProperties(section.sectPr),
    headers: references(section.sectPr, 'headerReference', resolveRelationship),
    footers: references(section.sectPr, 'footerReference', resolveRelationship),
  }));
}
