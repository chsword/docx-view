import type {
  CompatibilitySettings,
  DocumentBlock,
  EastAsianLayout,
  EquationNode,
  ParagraphFrame,
  TableFloatingPosition,
  ImageInfo,
  ParagraphInfo,
  RubyInfo,
  SectionInfo,
  TableRowInfo,
} from './types.js';

export interface MeasureContext {
  defaultTabStopTwips: number;
  compatibilitySettings?: CompatibilitySettings;
}

export interface LineBox {
  heightPx: number;
  startOffset: number;
  endOffset: number;
  /** 这一行实际占的宽度；量不到时没有。带 `w:framePr` 的段落靠它算排除区的宽度。 */
  widthPx?: number;
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
  compatibilitySettings?: CompatibilitySettings,
): { beforePx: number; afterPx: number } {
  const format = current.effective ?? current;
  const previousFormat = previous?.effective ?? previous;
  const autoSpacingDisabled = compatibilitySettings?.doNotUseHTMLParagraphAutoSpacing === true;
  const before = format.spacingBeforeAuto && !autoSpacingDisabled ? 0 : Math.max(0, format.spacingBefore ?? 0) * TWIPS_TO_PX;
  const after = format.spacingAfterAuto && !autoSpacingDisabled ? 0 : Math.max(0, format.spacingAfter ?? 0) * TWIPS_TO_PX;
  const previousAfter = previousFormat?.spacingAfterAuto && !autoSpacingDisabled
    ? 0 : Math.max(0, previousFormat?.spacingAfter ?? 0) * TWIPS_TO_PX;
  const sameStyle = previous !== undefined && previousFormat?.style === format.style;
  return {
    beforePx: sameStyle && format.contextualSpacing ? before : Math.max(previousAfter, before),
    afterPx: after,
  };
}

export function effectiveKinsoku(value: boolean | null | undefined, compatibilitySettings?: CompatibilitySettings): boolean | undefined {
  return compatibilitySettings?.doNotUseEastAsianBreakRules === true ? undefined : value ?? undefined;
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
    // 带 w:framePr 的段落浮出正常流：自己不占纵向高度，只给后面的内容留出排除区。首字下沉
    // 和定位文本框都走这里。
    const frame = frameWrapExclusion(paragraph.frame, {
      widthPx: lines.reduce<number | undefined>((widest, line) => line.widthPx === undefined ? widest
        : Math.max(widest ?? 0, line.widthPx), undefined),
      heightPx: lines.reduce((sum, line) => sum + Math.max(0, line.heightPx), 0),
    });
    if (frame) {
      // 排除区放不进本栏剩余高度、而本栏已经有内容时先换栏，别让它跨到下一栏去挤正文。
      if (frame.heightPx > remainingHeight() && hasContent(ensurePage(), currentColumn)) advanceColumn();
      append(lines.map((line) => ({ type: 'line' as const, paragraph: paragraph.index, line })), 0);
      wrapsByColumn[currentColumn] = [...(wrapsByColumn[currentColumn] ?? []), { ...frame, carried: true }];
      return;
    }
    const spacing = paragraphSpacingPx(previousParagraph, paragraph, context.compatibilitySettings);
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
      const spacing = paragraphSpacingPx(groupPrevious, item.paragraph, context.compatibilitySettings);
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
        const spacing = paragraphSpacingPx(groupPrevious, item.paragraph, context.compatibilitySettings);
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
      const spacing = paragraphSpacingPx(previousParagraph, item.paragraph, context.compatibilitySettings);
      const nextParagraph = measured[measured.indexOf(item) + 1]?.paragraph ?? following;
      const contribution = Math.max(0, spacing.beforePx + (nextParagraph ? 0 : spacing.afterPx));
      append(item.lines.map((line) => ({ type: 'line', paragraph: item.paragraph.index, line })), height + contribution);
      updateWraps(item.paragraph, height + contribution);
      previousParagraph = item.paragraph;
      previousAfterPx = spacing.afterPx;
    }
  };
  const addTable = (table: LayoutTable, blockIndex: number) => {
    const floating = (table.effective ?? table.format)?.floatingPosition;
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
    if (floating) {
      // 浮动表格整块脱离正常流：自己不占纵向高度，只给后面的内容留出排除区。宽度优先用
      // w:tblW，没有就用网格列宽合计（gridCol 的 w:w 是缇）。
      const heightPx = table.rows.reduce((sum, _row, index) => sum + rowHeight(index), 0);
      const widthTwips = (table.effective ?? table.format)?.width?.type === 'dxa'
        ? (table.effective ?? table.format)!.width!.value
        : table.grid.reduce((sum, column) => sum + Math.max(0, column), 0);
      const exclusion = tableWrapExclusion(floating, { widthPx: widthTwips / 15, heightPx });
      if (exclusion) {
        // 排除区放不进本栏剩余高度、而本栏已经有内容时先换栏，和带 framePr 的段落一样。
        if (exclusion.heightPx > remainingHeight() && hasContent(ensurePage(), currentColumn)) advanceColumn();
        append(table.rows.map((_row, index) => ({
          type: 'tableRow' as const, table: blockIndex, row: index, heightPx: rowHeight(index),
        })), 0);
        wrapsByColumn[currentColumn] = [...(wrapsByColumn[currentColumn] ?? []), { ...exclusion, carried: true }];
        return;
      }
    }
    let headerCount = 0;
    // 表头行可能是表格样式的 firstRow 条件给的 w:tblHeader，不只是行自己写的。
    while (headerCount < table.rows.length
      && (table.rows[headerCount]!.effective ?? table.rows[headerCount]!.format)?.header) headerCount++;
    const appendHeaders = () => {
      for (let index = 0; index < headerCount; index++) {
        const height = rowHeight(index);
        append([{ type: 'tableRow', table: blockIndex, row: index, heightPx: height }], height);
      }
    };
    const wrappedTable = context.compatibilitySettings?.doNotBreakWrappedTables === true &&
      table.rows.some((row) => row.cells.some((cell) => cell.blocks.some((block) =>
        block.type === 'paragraph' && block.paragraph.images.some((image) =>
          image.placement === 'floating' && image.wrap !== undefined && image.wrap !== 'none'))));
    if (headerCount === table.rows.length) {
      appendHeaders();
      return;
    }
    let index = headerCount;
    let fragmentStarted = false;
    while (index < table.rows.length) {
      let end = wrappedTable ? table.rows.length : index + 1;
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

/**
 * 「双行合一」（`w:eastAsianLayout w:combine`）：Word 把这段文字平分成上下两行，一起占一行的
 * 高度。按码点切，不按 UTF-16 码元，否则会把代理对劈成两半。
 */
export function combinedTextLines(text: string): [string, string] {
  const characters = Array.from(text);
  const cut = Math.ceil(characters.length / 2);
  return [characters.slice(0, cut).join(''), characters.slice(cut).join('')];
}

/** 双行合一两侧的括号。CSS 没有对应能力，所以当装饰画，不进 `readText()`。 */
export function combineBracketChars(style: EastAsianLayout['combineBrackets']): [string, string] | undefined {
  switch (style) {
    case 'round': return ['\uff08', '\uff09'];
    case 'square': return ['\uff3b', '\uff3d'];
    case 'angle': return ['\u3008', '\u3009'];
    case 'curly': return ['\uff5b', '\uff5d'];
    default: return undefined;
  }
}

/**
 * `w:rubyAlign` → CSS `ruby-align`。CSS 只有 start / center / space-between / space-around 四个值，
 * `right` 与 `rightVertical` 没有对应项，返回 `undefined` 让浏览器用默认值，不硬凑一个近似的。
 */
export function rubyAlignToCss(align: RubyInfo['align']): string | undefined {
  switch (align) {
    case 'center': return 'center';
    case 'distributeLetter': return 'space-between';
    case 'distributeSpace': return 'space-around';
    case 'left': return 'start';
    default: return undefined;
  }
}

/** `EQ \o(…)` 叠印出来的一层。 */
export interface OverstrikeLayer {
  text: string;
  /**
   * 这段文字是不是域结果里的内容。不是的话，它只存在于指令里（「带圈字符」的那个圈就是
   * 这种），只能当装饰画出来——否则编辑正文时会把它写回文档。
   */
  content: boolean;
  /** 垂直位移，单位磅；正数向上，`\do` 来的是负数。 */
  raisePoints: number;
}

function equationText(node: EquationNode): string {
  return node.text ?? (node.parts ?? []).map(equationText).join('');
}

/**
 * 把 `EQ \o(…)` 拆成要叠印的各层，并判断每层是域结果里的内容还是指令里的装饰字形。
 *
 * 判断办法是按顺序去对域结果：对得上的那段是内容，对不上的是装饰。「合并字符」的两组文字
 * 拼起来正好是域结果，所以两层都是内容；「带圈字符」的圈只在指令里，所以是装饰。
 *
 * 各层拼起来覆盖不住整个域结果时返回 `undefined`（指令和缓存结果不一致，例如域是脏的）——
 * 此时调用方按普通行内文字画，宁可不叠印也不能把结果里的文字丢掉。
 */
export function overstrikeLayers(equation: EquationNode | undefined, result: string): OverstrikeLayer[] | undefined {
  if (equation?.switch !== 'o' || !equation.parts?.length) return undefined;
  const layers: OverstrikeLayer[] = [];
  let cursor = 0;
  for (const part of equation.parts) {
    const text = equationText(part);
    if (!text) continue;
    const content = result.startsWith(text, cursor);
    if (content) cursor += text.length;
    layers.push({ text, content, raisePoints: part.raisePoints ?? 0 });
  }
  if (cursor !== result.length) return undefined;
  return layers.length ? layers : undefined;
}

/**
 * `w:framePr` 的段落脱离正常流，给后面的内容留出一块排除区——这正是浏览器 float 的行为，
 * 所以**首字下沉和段落定位走的是同一条路**，只是尺寸来源不同。
 *
 * 下沉字的尺寸不从 `w:lines` 反推：Word 已经把那个字的 `w:sz` 调到正好跨 `lines` 行，所以
 * 量出来的字框就是排除区，和「读 Word 已经算好的结果」一个路子。定位文本框优先用 `w:w` /
 * `w:h`，缺的那一维用量出来的。
 *
 * `w:wrap` 的取值正好能对上现有的排除区类型：`none` 不产生排除区（正文不绕它排），
 * `notBeside` 是「旁边不许有文字」，也就是 `topAndBottom`。
 */
/**
 * 浮动表格（`w:tblpPr`）的排除区。和 `frameWrapExclusion` 是同一件事：这一块脱离正常流，
 * 给后面的内容留出空间，环绕交给浏览器 `float`。
 *
 * `w:tblpPr` 上没有 `w:wrap` —— 浮动表格在 Word 里一定绕排，这正是它的用途，所以排除区
 * 类型固定是 `square`；`w:tblOverlap` 管的是能否与其他浮动对象重叠，与正文绕排无关。
 */
export function tableWrapExclusion(position: TableFloatingPosition | null | undefined,
  measured: { widthPx?: number; heightPx: number }): WrapExclusion | undefined {
  if (!position) return undefined;
  const { widthPx, heightPx } = measured;
  // 宽度算不出来时不编一个：不环绕只是少个效果，编错了会把正文挤歪。
  if (widthPx === undefined || !(widthPx > 0) || !(heightPx > 0)) return undefined;
  return { widthPx, heightPx, wrap: 'square', carried: false };
}

export function frameWrapExclusion(frame: ParagraphFrame | null | undefined,
  measured: { widthPx?: number; heightPx: number }): WrapExclusion | undefined {
  if (!frame) return undefined;
  const wrap = frame.wrap ?? 'auto';
  if (wrap === 'none') return undefined;
  // 文档明写了 w:w / w:h 就照它来，没写才用量出来的。下沉字正是「没写」的那种：Word 不给
  // 它写尺寸，而是把那个字的 w:sz 调到正好跨 lines 行，所以量出来的字框就是排除区。
  const widthPx = frame.widthTwips === undefined ? measured.widthPx : frame.widthTwips / 15;
  const heightPx = frame.heightTwips === undefined ? measured.heightPx : frame.heightTwips / 15;
  // 宽度量不到又没给 w:w 时不编一个出来：没有排除区也只是不环绕，编错了会把正文挤歪。
  if (widthPx === undefined || !(widthPx > 0) || !(heightPx > 0)) return undefined;
  return {
    widthPx,
    heightPx,
    wrap: wrap === 'tight' ? 'tight' : wrap === 'through' ? 'through'
      : wrap === 'notBeside' ? 'topAndBottom' : 'square',
    carried: false,
  };
}
