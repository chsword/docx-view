import type {
  DocumentBlock,
  ParagraphInfo,
  SectionInfo,
  TableRowInfo,
} from './types.js';

export interface MeasureContext {
  defaultTabStopTwips: number;
}

export interface LineBox {
  heightPx: number;
  startOffset: number;
  endOffset: number;
}

export interface LayoutMeasurer {
  measureParagraph(paragraph: ParagraphInfo, widthPx: number, context: MeasureContext): LineBox[];
  measureTableRow(row: TableRowInfo, widthPx: number, context: MeasureContext): number;
}

export type FlowItem =
  | { type: 'line'; paragraph: number; line: LineBox }
  | { type: 'tableRow'; table: number; row: number; heightPx: number }
  | { type: 'break'; kind: 'page' | 'section' };

export interface PageBox {
  index: number;
  number: number;
  section: number;
  items: FlowItem[];
  contentHeightPx: number;
}

const TWIPS_TO_PX = 96 / 1440;

function usableSize(section: SectionInfo): { width: number; height: number } {
  const width = (section.pageWidth - section.margins.left - section.margins.right) * TWIPS_TO_PX;
  const height = (section.pageHeight - section.margins.top - section.margins.bottom) * TWIPS_TO_PX;
  return { width: Math.max(0, width), height: Math.max(0, height) };
}

/**
 * Lay out the block stream without touching the DOM. Tables are deliberately
 * kept together here; splitting rows and columns is deferred to A3.
 */
export function paginate(
  blocks: DocumentBlock[],
  sections: SectionInfo[],
  measurer: LayoutMeasurer,
  context: MeasureContext,
): PageBox[] {
  if (blocks.length === 0) return [];

  const layoutSections = sections.length > 0 ? sections : [{
    index: 0,
    startParagraph: 0,
    endParagraph: Number.MAX_SAFE_INTEGER,
    isImplicit: true,
    type: 'nextPage',
    pageWidth: 11906,
    pageHeight: 16838,
    orientation: 'portrait',
    margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 0, footer: 0, gutter: 0 },
    columns: { count: 1, space: 0, equalWidth: true },
    pageNumbering: {},
    titlePage: false,
    headers: {},
    footers: {},
  } satisfies SectionInfo];
  const pages: PageBox[] = [];
  const firstSection = layoutSections[0]!;
  let sectionIndex = Math.max(0, firstSection.index);
  let current: PageBox | undefined;
  let nextNumber = firstSection.pageNumbering?.start ?? 1;

  const sectionAt = (index: number): SectionInfo => {
    return layoutSections.find((section) => section.index === index) ?? firstSection;
  };
  const sizeFor = (index: number) => usableSize(sectionAt(index));
  const startPage = (index: number, number = nextNumber): PageBox => {
    const page: PageBox = {
      index: pages.length,
      number,
      section: index,
      items: [],
      contentHeightPx: 0,
    };
    pages.push(page);
    current = page;
    nextNumber = number + 1;
    return page;
  };
  const ensurePage = () => current ?? startPage(sectionIndex);
  const newPage = (index = sectionIndex) => startPage(index);
  const hasContent = (page: PageBox) => page.items.length > 0;
  const heightLimit = () => sizeFor(sectionIndex).height;
  const width = () => sizeFor(sectionIndex).width;

  const append = (items: FlowItem[], height: number) => {
    const page = ensurePage();
    page.items.push(...items);
    page.contentHeightPx += height;
  };
  const putLines = (paragraph: ParagraphInfo, lines: LineBox[]) => {
    if (lines.length === 0) return;
    const limit = heightLimit();
    let offset = 0;
    while (offset < lines.length) {
      const page = ensurePage();
      const start = offset;
      let used = 0;
      while (offset < lines.length && used + Math.max(0, lines[offset]!.heightPx) <= Math.max(0, limit - page.contentHeightPx)) {
        used += Math.max(0, lines[offset]!.heightPx);
        offset++;
      }
      if (offset > start) {
        append(lines.slice(start, offset).map((line) => ({ type: 'line', paragraph: paragraph.index, line })), used);
      }
      // A line taller than a page is still progress: place it alone.
      if (offset === 0 || (used === 0 && page.contentHeightPx === 0)) {
        const line = lines[offset++]!;
        append([{ type: 'line', paragraph: paragraph.index, line }], Math.max(0, line.heightPx));
      }
      if (offset < lines.length) newPage();
    }
  };
  const paragraphLines = (paragraph: ParagraphInfo) =>
    measurer.measureParagraph(paragraph, width(), context) ?? [];

  const addParagraph = (paragraph: ParagraphInfo) => {
    const lines = paragraphLines(paragraph);
    if (lines.length === 0) return;
    if (paragraph.pageBreakBefore && hasContent(ensurePage())) newPage();
    const total = lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0);
    const limit = heightLimit();
    const keepWhole = paragraph.keepLines === true;
    const widow = paragraph.widowControl !== false && lines.length > 1;
    const page = ensurePage();
    if (keepWhole || total <= Math.max(0, limit - page.contentHeightPx)) {
      if (keepWhole && total > Math.max(0, limit - page.contentHeightPx) && hasContent(page)) newPage();
      append(lines.map((line) => ({ type: 'line', paragraph: paragraph.index, line })), total);
      return;
    }
    let fit = 0;
    let used = 0;
    while (fit < lines.length && used + Math.max(0, lines[fit]!.heightPx) <= Math.max(0, limit - page.contentHeightPx)) {
      used += Math.max(0, lines[fit]!.heightPx);
      fit++;
    }
    if (fit === 0 || (widow && (fit === 1 || lines.length - fit === 1))) {
      if (hasContent(page)) newPage();
      putLines(paragraph, lines);
      return;
    }
    append(lines.slice(0, fit).map((line) => ({ type: 'line', paragraph: paragraph.index, line })), used);
    newPage();
    putLines(paragraph, lines.slice(fit));
  };

  const addParagraphGroup = (paragraphs: ParagraphInfo[]) => {
    if (paragraphs.length === 1) {
      addParagraph(paragraphs[0]!);
      return;
    }
    const measured = paragraphs.map((paragraph) => ({ paragraph, lines: paragraphLines(paragraph) }));
    const total = measured.reduce((sum, item) => sum + item.lines.reduce((n, line) => n + Math.max(0, line.heightPx), 0), 0);
    const page = ensurePage();
    if (measured.some((item) => item.paragraph.pageBreakBefore) && hasContent(page)) newPage();
    if (total > Math.max(0, heightLimit() - page.contentHeightPx) && hasContent(page)) newPage();
    if (total > heightLimit()) {
      for (const item of paragraphs) addParagraph(item);
      return;
    }
    for (const item of measured) {
      append(item.lines.map((line) => ({ type: 'line', paragraph: item.paragraph.index, line })), item.lines.reduce((n, line) => n + Math.max(0, line.heightPx), 0));
    }
  };

  let blockIndex = 0;
  while (blockIndex < blocks.length) {
    const block = blocks[blockIndex]!;
    if (block.type === 'paragraph') {
      const group = [block.paragraph];
      while (blockIndex + group.length < blocks.length) {
        const next = blocks[blockIndex + group.length]!;
        if (next.type !== 'paragraph' || !group[group.length - 1]!.keepNext) break;
        group.push(next.paragraph);
      }
      addParagraphGroup(group);
      blockIndex += group.length;
      continue;
    }
    if (block.type === 'pageBreak') {
      append([{ type: 'break', kind: 'page' }], 0);
      newPage();
      blockIndex++;
      continue;
    }
    if (block.type === 'table') {
      const rows = block.rows.map((row, rowIndex) => ({
        row,
        rowIndex,
        height: Math.max(0, measurer.measureTableRow(row, width(), context) || 0),
      }));
      const total = rows.reduce((sum, row) => sum + row.height, 0);
      if (total > Math.max(0, heightLimit() - ensurePage().contentHeightPx) && hasContent(ensurePage())) newPage();
      append(rows.map(({ rowIndex, height }) => ({ type: 'tableRow', table: blockIndex, row: rowIndex, heightPx: height })), total);
      blockIndex++;
      continue;
    }
    append([{ type: 'break', kind: 'section' }], 0);
    const target = block.section + 1;
    const nextSection = sectionAt(target);
    if (block.breakType !== 'continuous') {
      if (block.breakType === 'evenPage' || block.breakType === 'oddPage') {
        const wanted = block.breakType === 'evenPage' ? 0 : 1;
        if ((nextNumber % 2) !== wanted) newPage(sectionIndex);
      }
      newPage(target);
    } else {
      sectionIndex = target;
      if (current) {
        current.section = target;
      }
    }
    sectionIndex = target;
    if (nextSection.pageNumbering?.start !== undefined && current) {
      current.number = nextSection.pageNumbering.start;
      nextNumber = current.number + 1;
    }
    blockIndex++;
  }
  return pages;
}
