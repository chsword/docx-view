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

/** Two adjacent paragraph margins use the larger margin, like CSS margin collapse. */
export function paragraphSpacingPx(
  previous: ParagraphInfo | undefined,
  current: ParagraphInfo,
): { beforePx: number; afterPx: number } {
  const format = current.effective ?? current;
  const previousFormat = previous?.effective ?? previous;
  const before = format.spacingBeforeAuto ? 0 : Math.max(0, format.spacingBefore ?? 0) * TWIPS_TO_PX;
  const after = format.spacingAfterAuto ? 0 : Math.max(0, format.spacingAfter ?? 0) * TWIPS_TO_PX;
  const previousAfter = previousFormat?.spacingAfterAuto
    ? 0 : Math.max(0, previousFormat?.spacingAfter ?? 0) * TWIPS_TO_PX;
  const sameStyle = previous !== undefined && previousFormat?.style === format.style;
  return {
    beforePx: sameStyle && format.contextualSpacing ? before : Math.max(previousAfter, before),
    afterPx: after,
  };
}

export function snapLineHeightPx(naturalHeightPx: number, docGrid: SectionInfo['docGrid']): number {
  if (!docGrid || !['lines', 'linesAndChars', 'snapToChars'].includes(docGrid.type) ||
      !Number.isFinite(docGrid.linePitch) || (docGrid.linePitch ?? 0) <= 0) {
    return naturalHeightPx;
  }

  const pitchPx = docGrid.linePitch! * TWIPS_TO_PX;
  if (!Number.isFinite(pitchPx) || pitchPx <= 0 || !Number.isFinite(naturalHeightPx)) return naturalHeightPx;
  return Math.ceil(Math.max(0, naturalHeightPx) / pitchPx) * pitchPx;
}

/** 给已分页的行列表编号。返回与输入等长的数组，null 表示该行不显示行号。 */
export function lineNumbersFor(
  pages: PageBox[],
  paragraphs: ParagraphInfo[],
  sections: SectionInfo[],
): Array<Array<number | null>> {
  const paragraphByIndex = new Map(paragraphs.map((paragraph) => [paragraph.index, paragraph]));
  let count = 0;
  let numberBase = 1;
  let previousSection: number | undefined;
  return pages.map((page) => {
    const section = sections.find((entry) => entry.index === page.section) ?? sections[0];
    const numbering = section?.lineNumbering;
    const restart = numbering?.restart ?? 'newPage';
    const start = Number.isFinite(numbering?.start) && (numbering?.start ?? 0) > 0 ? numbering!.start! : 1;
    const countBy = Number.isFinite(numbering?.countBy) && (numbering?.countBy ?? 0) > 0
      ? numbering!.countBy! : 1;
    const resetHere = previousSection === undefined ||
      restart === 'newPage' ||
      (restart === 'newSection' && previousSection !== page.section);
    if (resetHere) {
      count = 0;
      numberBase = start;
    }
    previousSection = page.section;
    const result: Array<number | null> = [];
    for (const item of page.items) {
      if (item.type !== 'line') {
        result.push(null);
        continue;
      }
      const paragraph = paragraphByIndex.get(item.paragraph);
      if (!numbering || paragraph?.effective?.suppressLineNumbers === true) {
        result.push(null);
        continue;
      }
      const number = numberBase + count;
      result.push(number % countBy === 0 ? number : null);
      count++;
    }
    return result;
  });
}

function usableSize(section: SectionInfo): { width: number; height: number } {
  const width = (section.pageWidth - section.margins.left - section.margins.right) * TWIPS_TO_PX;
  const height = (section.pageHeight - section.margins.top - section.margins.bottom) * TWIPS_TO_PX;
  return { width: Math.max(0, width), height: Math.max(0, height) };
}

export interface PageBoxPx {
  widthPx: number;
  heightPx: number;
  padding: { top: number; right: number; bottom: number; left: number };
}

export function pageBoxPx(section: Pick<SectionInfo, 'pageWidth' | 'pageHeight' | 'margins'>): PageBoxPx {
  // Same arithmetic as the page CSS it replaced; `twips * TWIPS_TO_PX` differs in the last ulp for many inputs.
  const toPx = (twips: number) => twips * 96 / 1440;
  return {
    widthPx: toPx(section.pageWidth),
    heightPx: toPx(section.pageHeight),
    padding: {
      top: toPx(section.margins.top),
      right: toPx(section.margins.right),
      bottom: toPx(section.margins.bottom),
      left: toPx(section.margins.left),
    },
  };
}

export function columnWidthsPx(section: SectionInfo): number[] {
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
  let previousParagraph: ParagraphInfo | undefined;
  let previousAfterPx = 0;
  let nextNumber = firstSection.pageNumbering?.start ?? 1;

  const sectionAt = (index: number): SectionInfo =>
    layoutSections.find((section) => section.index === index) ?? firstSection;
  const heightLimit = () => usableSize(sectionAt(sectionIndex)).height;
  const widths = () => columnWidthsPx(sectionAt(sectionIndex));
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
    previousParagraph = undefined;
    previousAfterPx = 0;
    columnHeights = Array.from({ length: columnWidthsPx(sectionAt(index)).length }, () => 0);
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
      previousParagraph = undefined;
      previousAfterPx = 0;
    } else {
      newPage();
      previousParagraph = undefined;
      previousAfterPx = 0;
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
    return (measurer.measureParagraph(paragraph, { widthPx: width(), wraps: [...carried, ...own] }, context) ?? [])
      .map((line) => ({ ...line, heightPx: snapLineHeightPx(line.heightPx, sectionAt(sectionIndex).docGrid) }));
  };
  const updateWraps = (paragraph: ParagraphInfo, consumedHeight: number, includeOwn = true) => {
    const carried = (wrapsByColumn[currentColumn] ?? [])
      .map((wrap) => ({ ...wrap, heightPx: wrap.heightPx - consumedHeight, carried: true }))
      .filter((wrap) => wrap.heightPx > 0);
    const own = includeOwn ? wrappingImages(paragraph)
      .map((image): WrapExclusion => ({
        widthPx: Math.max(0, image.widthPx),
        heightPx: Math.max(0, image.heightPx - consumedHeight),
        wrap: image.wrap as WrapExclusion['wrap'],
        carried: true,
      }))
      .filter((wrap) => wrap.heightPx > 0) : [];
    wrapsByColumn[currentColumn] = [...carried, ...own];
  };
  const putLines = (paragraph: ParagraphInfo, lines: LineBox[], includeOwn = true) => {
    if (lines.length === 0) return;
    let offset = 0;
    let includeOwnWraps = includeOwn;
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
        updateWraps(paragraph, used, includeOwnWraps);
        includeOwnWraps = false;
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
        updateWraps(paragraph, height, includeOwnWraps);
        includeOwnWraps = false;
      }
      if (offset < lines.length) advanceColumn();
    }
  };
  const addParagraph = (paragraph: ParagraphInfo, nextParagraph?: ParagraphInfo) => {
    if (paragraph.pageBreakBefore && hasContent(ensurePage())) {
      newPage();
      previousParagraph = undefined;
      previousAfterPx = 0;
    }
    let lines = measureParagraph(paragraph);
    if (lines.length === 0) return;
    const spacing = paragraphSpacingPx(previousParagraph, paragraph);
    const spacingContribution = spacing.beforePx + (nextParagraph ? 0 : spacing.afterPx);
    let total = lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0) + Math.max(0, spacingContribution);
    const keepWhole = paragraph.keepLines === true;
    const widow = paragraph.widowControl !== false && lines.length > 1;
    if ((keepWhole || total <= remainingHeight())) {
      if (keepWhole && total > remainingHeight() && hasContent(ensurePage(), currentColumn)) {
        advanceColumn();
        lines = measureParagraph(paragraph);
        total = lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0) + Math.max(0, spacingContribution);
      }
      append(lines.map((line) => ({ type: 'line', paragraph: paragraph.index, line })), total);
      updateWraps(paragraph, total);
      previousParagraph = paragraph;
      previousAfterPx = spacing.afterPx;
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
      previousParagraph = paragraph;
      previousAfterPx = spacing.afterPx;
      return;
    }
    append(lines.slice(0, fit).map((line) => ({ type: 'line', paragraph: paragraph.index, line })), used);
    updateWraps(paragraph, used);
    advanceColumn();
    putLines(paragraph, lines.slice(fit), false);
    previousParagraph = paragraph;
    previousAfterPx = spacing.afterPx;
  };
  const addParagraphGroup = (paragraphs: ParagraphInfo[], following?: ParagraphInfo) => {
    if (paragraphs.length === 1) {
      addParagraph(paragraphs[0]!, following);
      return;
    }
    let measured = paragraphs.map((paragraph) => ({ paragraph, lines: measureParagraph(paragraph) }));
    let groupPrevious = previousParagraph;
    let total = measured.reduce((sum, item) => {
      const spacing = paragraphSpacingPx(groupPrevious, item.paragraph);
      const nextParagraph = measured[measured.indexOf(item) + 1]?.paragraph ?? following;
      const contribution = spacing.beforePx + (nextParagraph ? 0 : spacing.afterPx);
      groupPrevious = item.paragraph;
      return sum + item.lines.reduce((height, line) => height + Math.max(0, line.heightPx), 0) + Math.max(0, contribution);
    }, 0);
    if (total > remainingHeight() && hasContent(ensurePage(), currentColumn)) {
      advanceColumn();
      measured = paragraphs.map((paragraph) => ({ paragraph, lines: measureParagraph(paragraph) }));
      groupPrevious = previousParagraph;
      total = measured.reduce((sum, item) => {
        const spacing = paragraphSpacingPx(groupPrevious, item.paragraph);
        const nextParagraph = measured[measured.indexOf(item) + 1]?.paragraph ?? following;
        const contribution = spacing.beforePx + (nextParagraph ? 0 : spacing.afterPx);
        groupPrevious = item.paragraph;
        return sum + item.lines.reduce((height, line) => height + Math.max(0, line.heightPx), 0) + Math.max(0, contribution);
      }, 0);
    }
    if (total > heightLimit()) {
      for (let index = 0; index < paragraphs.length; index++) {
        addParagraph(paragraphs[index]!, paragraphs[index + 1] ?? (index === paragraphs.length - 1 ? following : undefined));
      }
      return;
    }
    for (const item of measured) {
      const height = item.lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0);
      const spacing = paragraphSpacingPx(previousParagraph, item.paragraph);
      const nextParagraph = measured[measured.indexOf(item) + 1]?.paragraph ?? following;
      const contribution = Math.max(0, spacing.beforePx + (nextParagraph ? 0 : spacing.afterPx));
      append(item.lines.map((line) => ({ type: 'line', paragraph: item.paragraph.index, line })), height + contribution);
      updateWraps(item.paragraph, height + contribution);
      previousParagraph = item.paragraph;
      previousAfterPx = spacing.afterPx;
    }
  };
  const addTable = (table: LayoutTable, blockIndex: number) => {
    // A table row is the smallest table FlowItem, so cantSplit rows are always
    // atomic; rowSpan extends that atomic group through every covered row.
    const heightCache = new Map<number, Map<number, number>>();
    const rowHeight = (rowIndex: number) => {
      const widthPx = width();
      let byWidth = heightCache.get(rowIndex);
      if (!byWidth) {
        byWidth = new Map();
        heightCache.set(rowIndex, byWidth);
      }
      const cached = byWidth.get(widthPx);
      if (cached !== undefined) return cached;
      const height = Math.max(0, measurer.measureTableRow(
        table, table.rows[rowIndex]!, rowIndex, widthPx, context,
      ) || 0);
      byWidth.set(widthPx, height);
      return height;
    };
    let headerCount = 0;
    while (headerCount < table.rows.length && table.rows[headerCount]!.format?.header) headerCount++;
    const appendHeaders = () => {
      for (let index = 0; index < headerCount; index++) {
        const height = rowHeight(index);
        append([{ type: 'tableRow', table: blockIndex, row: index, heightPx: height }], height);
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
      let group = Array.from({ length: end - index }, (_, offset) => ({
        rowIndex: index + offset,
        height: rowHeight(index + offset),
      }));
      let groupHeight = group.reduce((sum, row) => sum + row.height, 0);
      let headerHeight = Array.from({ length: headerCount }, (_, headerIndex) => rowHeight(headerIndex))
        .reduce((sum, height) => sum + height, 0);
      if ((!fragmentStarted && headerHeight + groupHeight > remainingHeight() && hasContent(ensurePage(), currentColumn)) ||
          (fragmentStarted && groupHeight > remainingHeight())) {
        advanceColumn();
        fragmentStarted = false;
        group = Array.from({ length: end - index }, (_, offset) => ({
          rowIndex: index + offset,
          height: rowHeight(index + offset),
        }));
        groupHeight = group.reduce((sum, row) => sum + row.height, 0);
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
      const following = blocks[blockIndex + group.length];
      addParagraphGroup(group, following?.type === 'paragraph' && !following.paragraph.pageBreakBefore
        ? following.paragraph : undefined);
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
      previousParagraph = undefined;
      previousAfterPx = 0;
      blockIndex++;
      continue;
    }
    append([{ type: 'break', kind: 'section' }], 0);
    const target = block.section + 1;
    const nextSection = sectionAt(target);
    const currentWidths = columnWidthsPx(sectionAt(sectionIndex));
    const nextWidths = columnWidthsPx(nextSection);
    const sameColumns = currentWidths.length === nextWidths.length &&
      currentWidths.every((value, index) => Math.abs(value - nextWidths[index]!) < 0.01);
    if (block.breakType === 'nextColumn') {
      sectionIndex = target;
      previousParagraph = undefined;
      previousAfterPx = 0;
      if (sameColumns) {
        if (current) current.section = target;
        advanceColumn();
      } else {
        newPage(target);
      }
    } else if (block.breakType !== 'continuous') {
      if (block.breakType === 'evenPage' || block.breakType === 'oddPage') {
        const wanted = block.breakType === 'evenPage' ? 0 : 1;
        if ((nextNumber % 2) !== wanted) newPage(sectionIndex);
      }
      newPage(target);
    } else {
      sectionIndex = target;
      if (sameColumns) {
        if (current) current.section = target;
      } else {
        newPage(target);
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
