import type {
  DocumentBlock,
  ImageInfo,
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

export interface WrapExclusion {
  widthPx: number;
  heightPx: number;
  wrap: 'square' | 'tight' | 'through' | 'topAndBottom';
  carried: boolean;
}

export interface ParagraphMeasureArea {
  widthPx: number;
  wraps: WrapExclusion[];
}

export type LayoutTable = Extract<DocumentBlock, { type: 'table' }>;

export interface LayoutMeasurer {
  measureParagraph(paragraph: ParagraphInfo, area: ParagraphMeasureArea, context: MeasureContext): LineBox[];
  measureTableRow(table: LayoutTable, row: TableRowInfo, rowIndex: number, widthPx: number, context: MeasureContext): number;
}

export type FlowItem =
  | { type: 'line'; paragraph: number; line: LineBox; column?: number }
  | { type: 'tableRow'; table: number; row: number; heightPx: number; column?: number }
  | { type: 'break'; kind: 'page' | 'section'; column?: number };

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

function columnWidths(section: SectionInfo): number[] {
  const count = Math.max(1, Math.floor(section.columns.count) || 1);
  const total = usableSize(section).width;
  const gap = Math.max(0, section.columns.space * TWIPS_TO_PX);
  const available = Math.max(0, total - gap * (count - 1));
  const explicit = section.columns.equalWidth === false && section.columns.widths?.length === count
    ? section.columns.widths.map((width) => Math.max(0, width * TWIPS_TO_PX))
    : undefined;
  if (explicit && explicit.some((width) => width > 0)) return explicit;
  return Array.from({ length: count }, () => available / count);
}

function wrappingImages(paragraph: ParagraphInfo): ImageInfo[] {
  return paragraph.images.filter((image) =>
    image.placement === 'floating' && image.wrap !== undefined && image.wrap !== 'none');
}

/**
 * Lay out the block stream without touching the DOM. A row-spanning cell and
 * every row it covers form one pagination group, so a merge is moved intact
 * rather than rendered as disconnected cells across a column or page.
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
  let currentColumn = 0;
  let columnHeights: number[] = [];
  let wrapsByColumn: WrapExclusion[][] = [];
  let nextNumber = firstSection.pageNumbering?.start ?? 1;

  const sectionAt = (index: number): SectionInfo =>
    layoutSections.find((section) => section.index === index) ?? firstSection;
  const heightLimit = () => usableSize(sectionAt(sectionIndex)).height;
  const widths = () => columnWidths(sectionAt(sectionIndex));
  const width = () => widths()[Math.min(currentColumn, widths().length - 1)] ?? 0;
  const usedHeight = () => columnHeights[currentColumn] ?? 0;
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
    currentColumn = 0;
    columnHeights = Array.from({ length: columnWidths(sectionAt(index)).length }, () => 0);
    wrapsByColumn = columnHeights.map(() => []);
    nextNumber = number + 1;
    return page;
  };
  const ensurePage = () => current ?? startPage(sectionIndex);
  const newPage = (index = sectionIndex) => startPage(index);
  const hasContent = (page: PageBox, column?: number) =>
    page.items.some((item) => column === undefined || (item.column ?? 0) === column);
  const advanceColumn = () => {
    ensurePage();
    if (currentColumn + 1 < widths().length) {
      currentColumn++;
    } else {
      newPage();
    }
  };
  const append = (items: FlowItem[], height: number) => {
    const page = ensurePage();
    page.items.push(...items.map((item) => ({ ...item, column: currentColumn })));
    columnHeights[currentColumn] = usedHeight() + height;
    page.contentHeightPx = Math.max(...columnHeights, 0);
  };
  const remainingHeight = () => Math.max(0, heightLimit() - usedHeight());

  const measureParagraph = (paragraph: ParagraphInfo) => {
    const carried = wrapsByColumn[currentColumn] ?? [];
    const own = wrappingImages(paragraph).map((image): WrapExclusion => ({
      widthPx: Math.max(0, image.widthPx),
      heightPx: Math.max(0, image.heightPx),
      wrap: image.wrap as WrapExclusion['wrap'],
      carried: false,
    }));
    return measurer.measureParagraph(paragraph, { widthPx: width(), wraps: [...carried, ...own] }, context) ?? [];
  };
  const updateWraps = (paragraph: ParagraphInfo, consumedHeight: number) => {
    const carried = (wrapsByColumn[currentColumn] ?? [])
      .map((wrap) => ({ ...wrap, heightPx: wrap.heightPx - consumedHeight, carried: true }))
      .filter((wrap) => wrap.heightPx > 0);
    const own = wrappingImages(paragraph)
      .map((image): WrapExclusion => ({
        widthPx: Math.max(0, image.widthPx),
        heightPx: Math.max(0, image.heightPx - consumedHeight),
        wrap: image.wrap as WrapExclusion['wrap'],
        carried: true,
      }))
      .filter((wrap) => wrap.heightPx > 0);
    wrapsByColumn[currentColumn] = [...carried, ...own];
  };
  const putLines = (paragraph: ParagraphInfo, lines: LineBox[]) => {
    if (lines.length === 0) return;
    let offset = 0;
    while (offset < lines.length) {
      const start = offset;
      let used = 0;
      while (offset < lines.length &&
        used + Math.max(0, lines[offset]!.heightPx) <= remainingHeight()) {
        used += Math.max(0, lines[offset]!.heightPx);
        offset++;
      }
      if (offset > start) {
        append(lines.slice(start, offset).map((line) => ({
          type: 'line',
          paragraph: paragraph.index,
          line,
        })), used);
        updateWraps(paragraph, used);
      }
      // A line taller than a column is still progress: place it alone.
      if (offset === start) {
        if (hasContent(ensurePage(), currentColumn)) {
          advanceColumn();
          continue;
        }
        const line = lines[offset++]!;
        const height = Math.max(0, line.heightPx);
        append([{ type: 'line', paragraph: paragraph.index, line }], height);
        updateWraps(paragraph, height);
      }
      if (offset < lines.length) advanceColumn();
    }
  };
  const addParagraph = (paragraph: ParagraphInfo) => {
    if (paragraph.pageBreakBefore && hasContent(ensurePage())) newPage();
    let lines = measureParagraph(paragraph);
    if (lines.length === 0) return;
    let total = lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0);
    const keepWhole = paragraph.keepLines === true;
    const widow = paragraph.widowControl !== false && lines.length > 1;
    if ((keepWhole || total <= remainingHeight())) {
      if (keepWhole && total > remainingHeight() && hasContent(ensurePage(), currentColumn)) {
        advanceColumn();
        lines = measureParagraph(paragraph);
        total = lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0);
      }
      append(lines.map((line) => ({ type: 'line', paragraph: paragraph.index, line })), total);
      updateWraps(paragraph, total);
      return;
    }
    let fit = 0;
    let used = 0;
    while (fit < lines.length && used + Math.max(0, lines[fit]!.heightPx) <= remainingHeight()) {
      used += Math.max(0, lines[fit]!.heightPx);
      fit++;
    }
    if (fit === 0 || (widow && (fit === 1 || lines.length - fit === 1))) {
      if (hasContent(ensurePage(), currentColumn)) {
        advanceColumn();
        lines = measureParagraph(paragraph);
      }
      putLines(paragraph, lines);
      return;
    }
    append(lines.slice(0, fit).map((line) => ({ type: 'line', paragraph: paragraph.index, line })), used);
    updateWraps(paragraph, used);
    advanceColumn();
    putLines(paragraph, lines.slice(fit));
  };
  const addParagraphGroup = (paragraphs: ParagraphInfo[]) => {
    if (paragraphs.length === 1) {
      addParagraph(paragraphs[0]!);
      return;
    }
    let measured = paragraphs.map((paragraph) => ({ paragraph, lines: measureParagraph(paragraph) }));
    let total = measured.reduce((sum, item) =>
      sum + item.lines.reduce((height, line) => height + Math.max(0, line.heightPx), 0), 0);
    if (total > remainingHeight() && hasContent(ensurePage(), currentColumn)) {
      advanceColumn();
      measured = paragraphs.map((paragraph) => ({ paragraph, lines: measureParagraph(paragraph) }));
      total = measured.reduce((sum, item) =>
        sum + item.lines.reduce((height, line) => height + Math.max(0, line.heightPx), 0), 0);
    }
    if (total > heightLimit()) {
      for (const paragraph of paragraphs) addParagraph(paragraph);
      return;
    }
    for (const item of measured) {
      const height = item.lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0);
      append(item.lines.map((line) => ({ type: 'line', paragraph: item.paragraph.index, line })), height);
      updateWraps(item.paragraph, height);
    }
  };
  const addTable = (table: LayoutTable, blockIndex: number) => {
    const measuredRows = table.rows.map((row, rowIndex) => ({
      rowIndex,
      height: Math.max(0, measurer.measureTableRow(table, row, rowIndex, width(), context) || 0),
    }));
    let headerCount = 0;
    while (headerCount < table.rows.length && table.rows[headerCount]!.format?.header) headerCount++;
    const appendHeaders = () => {
      for (let index = 0; index < headerCount; index++) {
        const measured = measuredRows[index]!;
        append([{ type: 'tableRow', table: blockIndex, row: index, heightPx: measured.height }], measured.height);
      }
    };
    if (headerCount === table.rows.length) {
      appendHeaders();
      return;
    }
    let index = headerCount;
    let fragmentStarted = false;
    while (index < table.rows.length) {
      let end = index + 1;
      for (let scan = index; scan < end; scan++) {
        const span = Math.max(1, ...table.rows[scan]!.cells.map((cell) => Math.max(1, cell.rowSpan)));
        end = Math.max(end, Math.min(table.rows.length, scan + span));
      }
      const group = measuredRows.slice(index, end);
      const groupHeight = group.reduce((sum, row) => sum + row.height, 0);
      const headerHeight = measuredRows.slice(0, headerCount).reduce((sum, row) => sum + row.height, 0);
      if ((!fragmentStarted && headerHeight + groupHeight > remainingHeight() && hasContent(ensurePage(), currentColumn)) ||
          (fragmentStarted && groupHeight > remainingHeight())) {
        advanceColumn();
        fragmentStarted = false;
      }
      if (!fragmentStarted) {
        appendHeaders();
        fragmentStarted = true;
      }
      for (const measured of group) {
        append([{
          type: 'tableRow',
          table: blockIndex,
          row: measured.rowIndex,
          heightPx: measured.height,
        }], measured.height);
      }
      index = end;
    }
  };

  let blockIndex = 0;
  while (blockIndex < blocks.length) {
    const block = blocks[blockIndex]!;
    if (block.type === 'paragraph') {
      const group = [block.paragraph];
      while (blockIndex + group.length < blocks.length) {
        const next = blocks[blockIndex + group.length]!;
        if (next.type !== 'paragraph' || !group[group.length - 1]!.keepNext || next.paragraph.pageBreakBefore) break;
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
      addTable(block, blockIndex);
      blockIndex++;
      continue;
    }
    append([{ type: 'break', kind: 'section' }], 0);
    const target = block.section + 1;
    const nextSection = sectionAt(target);
    if (block.breakType === 'nextColumn') {
      sectionIndex = target;
      if (current) current.section = target;
      advanceColumn();
    } else if (block.breakType !== 'continuous') {
      if (block.breakType === 'evenPage' || block.breakType === 'oddPage') {
        const wanted = block.breakType === 'evenPage' ? 0 : 1;
        if ((nextNumber % 2) !== wanted) newPage(sectionIndex);
      }
      newPage(target);
    } else {
      sectionIndex = target;
      if (current) current.section = target;
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
