import { DocxDocument } from './document.js';
import type {
  BorderFormat,
  CompatibilitySettings,
  ClipboardFragment,
  ClipboardRun,
  BordersFormat,
  CellFormat,
  ChartSeriesInfo,
  ShapeShadow,
  ShapeTextParagraph,
  EquationNode,
  TableException,
  WebDivInfo,
  DocumentRange,
  DocumentBlock,
  DocumentSnapshot,
  FieldInfo,
  FieldKind,
  ImageInfo,
  MathInfo,
  MathMlNode,
  ParagraphFormat,
  ParagraphInfo,
  PaginationInfo,
  DocumentStatistics,
  ShapeInfo,
  ReviewerFilterAuthor,
  RunFormat,
  RunInfo,
  SectionInfo,
  TabStop,
  TableCellLocation,
  TableFormat,
  TableRowInfo,
  WidthFormat,
} from './types.js';
import { contentTypeForExtension, dataUrlForBytes, isBrowserRenderableContentType, pxToEmu } from './drawing.js';
import { customGeometryPath, presetGeometryIsOpen, presetGeometryPath } from './geometry.js';
import { axisTicks, barRects, errorBarRanges, pieSlicePath, radarPoint, trendlinePoints, valueToPx } from './chart.js';
import { isSafeHyperlinkUrl } from './hyperlink.js';
import { reviewerBucketKey, reviewerBucketOf } from './revisions.js';
import { eighthPointsToPx, normalizeColor, normalizeWidth, twipsToPx } from './table.js';
import { assertText, sanitizeText, sanitizeTextWithInfo } from './xml.js';
import { columnWidthsPx, combineBracketChars, combinedTextLines, effectiveKinsoku, floatOffset, lineNumbersFor, overstrikeLayers, pageBoxPx, paginate, paragraphSpacingPx, rubyAlignToCss } from './layout.js';
import type { FloatOffset, FlowItem, LayoutTable, LineBox, MeasureContext, PageBox, ParagraphMeasureArea } from './layout.js';
import { equationToMathMl, formatPageNumber, pageFieldResult } from './fields.js';
import { webDivIndents } from './web-divs.js';

export { formatPageNumber };

export const MAX_FIELD_UPDATE_ITERATIONS = 5;
const MATH_RENDER_INDEX = Symbol('mathRenderIndex');

export function updateFieldsUntilStable(update: (iteration: number) => boolean,
  maxIterations = MAX_FIELD_UPDATE_ITERATIONS): { updated: boolean; iterations: number } {
  let updated = false;
  for (let iteration = 0; iteration < maxIterations; iteration++) {
    const changed = update(iteration);
    updated ||= changed;
    if (!changed) return { updated, iterations: iteration + 1 };
  }
  return { updated, iterations: maxIterations };
}

export function paginationInfoFromPages(pages: PageBox[], blocks: DocumentBlock[]): PaginationInfo {
  const pageByParagraph = new Map<number, number>();
  const addBlocks = (items: DocumentBlock[], pageIndex: number): void => {
    for (const block of items) {
      if (block.type === 'paragraph') {
        if (!pageByParagraph.has(block.paragraph.index)) pageByParagraph.set(block.paragraph.index, pageIndex);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) addBlocks(cell.blocks, pageIndex);
      }
    }
  };
  const addRowParagraphs = (block: DocumentBlock | undefined, rowIndex: number, pageIndex: number): void => {
    if (block?.type !== 'table') return;
    const row = block.rows[rowIndex];
    if (row) for (const cell of row.cells) addBlocks(cell.blocks, pageIndex);
  };
  for (const page of pages) {
    for (const item of page.items) {
      if (item.type === 'line') {
        if (!pageByParagraph.has(item.paragraph)) pageByParagraph.set(item.paragraph, page.index);
      } else if (item.type === 'tableRow') addRowParagraphs(blocks[item.table], item.row, page.index);
    }
  }
  // 行数：正文的每一行算一行；表格行按一行算（单元格里的行没有单独测量）——是近似值，
  // 文字多的表格会比 Word 少算。
  const lineCount = pages.reduce((total, page) =>
    total + page.items.filter((item) => item.type === 'line' || item.type === 'tableRow').length, 0);
  return {
    pageCount: pages.length,
    lineCount,
    pageOfParagraph: paragraph => pageByParagraph.get(paragraph),
    numberOfPage: pageIndex => pages[pageIndex]?.number ?? 0,
  };
}

function twipsToPoints(value: number | null | undefined): string | undefined {
  return value !== undefined && value !== null ? `${value / 20}pt` : undefined;
}

function highlightColor(value: string): string {
  return {
    darkBlue: '#000080',
    darkCyan: '#008080',
    darkGray: '#808080',
    darkGreen: '#008000',
    darkMagenta: '#800080',
    darkRed: '#800000',
    darkYellow: '#808000',
    lightGray: '#D3D3D3',
    magenta: '#FF00FF',
  }[value] ?? value;
}

function underlineStyleToCss(value: string): string {
  return {
    single: 'solid',
    words: 'solid',
    thick: 'solid',
    dotted: 'dotted',
    dash: 'dashed',
    dashed: 'dashed',
    dashDotHeavy: 'dashed',
    dashLong: 'dashed',
    dashLongHeavy: 'dashed',
    dotDash: 'dashed',
    dotDotDash: 'dashed',
    double: 'double',
    doubleAccounting: 'double',
    wave: 'wavy',
    wavyDouble: 'wavy',
    wavyHeavy: 'wavy',
  }[value] ?? 'solid';
}

function borderStyle(value: string): string {
  return {
    single: 'solid',
    double: 'double',
    thick: 'solid',
    dashed: 'dashed',
    dotted: 'dotted',
    wave: 'wavy',
    none: 'none',
    nil: 'none',
  }[value] ?? 'solid';
}

const DOCX_CLIPBOARD_MIME = 'application/x-docx-view+json';
const MAX_REVIEW_FILTER_AUTHORS = 1_000;

export interface EditorReviewFilter {
  authors?: ReviewerFilterAuthor[];
  showRevisions?: boolean;
  showComments?: boolean;
  revisionView?: 'final' | 'original' | 'markup';
}

interface NormalizedReviewFilter {
  authors?: ReviewerFilterAuthor[];
  showRevisions: boolean;
  showComments: boolean;
  revisionView: 'final' | 'original' | 'markup';
}

interface ReviewRenderContext {
  authors?: Set<string>;
  deletedTextByRun: Map<string, string>;
  revisionColors: Map<string, string>;
}

const REVISION_COLOR_PALETTE = ['#2E75B6', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646', '#1F497D', '#843C0C'];

function revisionColorIndexForKey(key: string): number {
  let hash = 2166136261;
  for (let i = 0; i < key.length; i++) {
    hash ^= key.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0) % REVISION_COLOR_PALETTE.length;
}

function normalizeReviewFilterAuthor(author: ReviewerFilterAuthor, path = 'reviewFilter.authors[]'): ReviewerFilterAuthor {
  if (!author || typeof author !== 'object' || Array.isArray(author)) throw new Error(`${path} must be an object.`);
  if (!['named', 'unattributed', 'empty', 'blank'].includes(author.kind)) {
    throw new Error(`${path}.kind must be one of: named, unattributed, empty, blank.`);
  }
  const rawAuthor = (author as { author?: string }).author;
  if (author.kind === 'unattributed') {
    if (rawAuthor !== undefined) throw new Error(`${path}.author must be omitted for unattributed kind.`);
    return { kind: 'unattributed' };
  }
  if (rawAuthor === undefined) throw new Error(`${path}.author is required for ${author.kind} kind.`);
  assertText(rawAuthor, `${path}.author`);
  if (author.kind === 'named') {
    if (!rawAuthor.trim()) throw new Error(`${path}.author must be non-empty for named kind.`);
    return { kind: 'named', author: rawAuthor };
  }
  if (author.kind === 'empty') {
    if (rawAuthor !== '') throw new Error(`${path}.author must be an empty string for empty kind.`);
    return { kind: 'empty', author: '' };
  }
  if (rawAuthor.trim()) throw new Error(`${path}.author must be whitespace-only for blank kind.`);
  return { kind: 'blank', author: rawAuthor };
}

function normalizeReviewFilter(filter: EditorReviewFilter | undefined): NormalizedReviewFilter {
  if (filter !== undefined && (!filter || typeof filter !== 'object' || Array.isArray(filter))) {
    throw new Error('reviewFilter must be an object.');
  }
  if (filter?.authors !== undefined && !Array.isArray(filter.authors)) throw new Error('reviewFilter.authors must be an array.');
  if (filter?.authors && filter.authors.length > MAX_REVIEW_FILTER_AUTHORS) {
    throw new Error(`reviewFilter.authors must contain at most ${MAX_REVIEW_FILTER_AUTHORS} items.`);
  }
  if (filter?.showRevisions !== undefined && typeof filter.showRevisions !== 'boolean') {
    throw new Error('reviewFilter.showRevisions must be boolean.');
  }
  if (filter?.showComments !== undefined && typeof filter.showComments !== 'boolean') {
    throw new Error('reviewFilter.showComments must be boolean.');
  }
  if (filter?.revisionView !== undefined && !['final', 'original', 'markup'].includes(filter.revisionView)) {
    throw new Error('reviewFilter.revisionView must be one of: final, original, markup.');
  }
  const authors = filter?.authors?.map((author) => normalizeReviewFilterAuthor(author))
    .filter((author, index, all) =>
      all.findIndex((entry) => reviewerBucketKey(entry) === reviewerBucketKey(author)) === index);
  return {
    ...(authors?.length ? { authors } : {}),
    showRevisions: filter?.showRevisions ?? true,
    showComments: filter?.showComments ?? true,
    revisionView: filter?.revisionView ?? 'markup',
  };
}

function reviewFilterEqual(a: NormalizedReviewFilter, b: NormalizedReviewFilter): boolean {
  const authorsA = a.authors ?? [];
  const authorsB = b.authors ?? [];
  return a.showRevisions === b.showRevisions &&
    a.showComments === b.showComments &&
    a.revisionView === b.revisionView &&
    authorsA.length === authorsB.length &&
  authorsA.every((value, index) => reviewerBucketKey(value) === reviewerBucketKey(authorsB[index]!));
}

type HeaderFooterKind = keyof SectionInfo['headers'];

/**
 * Choose which header/footer variant a physical page shows and which part
 * backs it. Without a page number (continuous view) the caller's selected
 * kind is kept; a missing variant falls back to the default part.
 */
export function selectHeaderFooter(
  parts: SectionInfo['headers'],
  options: { selectedKind: HeaderFooterKind; pageNumber?: number; firstPhysicalPage?: boolean; titlePage?: boolean },
): { kind: HeaderFooterKind; part: string | undefined } {
  const kind: HeaderFooterKind = options.pageNumber === undefined
    ? options.selectedKind
    : options.firstPhysicalPage && options.titlePage ? 'first' : options.pageNumber % 2 === 0 ? 'even' : 'default';
  return { kind, part: parts[kind] ?? parts.default };
}

/** Replace the result runs of PAGE / NUMPAGES fields with the physical page values; other runs are untouched. */
export function replacePageFields(blocks: DocumentBlock[], pageNumber: number, pageCount: number | undefined, format?: string): DocumentBlock[] {
  return blocks.map((block) => {
    if (block.type !== 'paragraph') return block;
    return {
      ...block,
      paragraph: {
        ...block.paragraph,
        runs: block.paragraph.runs.map((run) => {
          const kind = run.field?.role === 'result' ? run.field.kind : undefined;
          const result = kind ? pageFieldResult(kind, pageNumber, pageCount ?? 0, format) : undefined;
          return result !== undefined
            ? { ...run, text: result }
            : run;
        }),
      },
    };
  });
}

/**
 * `w:framePr` 的段落浮出正常流，正文绕着它排。环绕交给浏览器的 float——测量分页时也是用
 * float 占位（见 measureParagraphForPagination），两边是同一套行为，不会各算一套。
 *
 * `w:x` / `w:y` 那套按页面或页边距定位的绝对坐标本期不实现：那需要相对页框定位，而连续视图
 * 没有页框。浮动方向按 `xAlign` 取左右，其余照左浮。
 */
/**
 * `offset`：分页布局按 x / y 算出的偏移（见 layout.ts 的 `floatOffset`），叠在 hSpace / vSpace 的外边距
 * 上；`null` 表示调用方（分页视图）已经算过、这一块没有偏移；不传时按连续视图处理——那里没有页，
 * 只认相对正文（text 锚点）的偏移。
 */
/**
 * 偏移是用外边距挪的，但浮动块的外边距区域默认也挡字——上方和左侧会空出一大片。`shape-outside`
 * 把挡字的范围收回到块本身，外边距那片空白照常排字。与分页测量的占位（同样的 inset）一致。
 */
function applyOffsetShape(element: HTMLElement, offset: FloatOffset): void {
  const x = Math.max(0, offset.xPx);
  const y = Math.max(0, offset.yPx);
  if (x || y) element.style.shapeOutside = `inset(${y}px 0 0 ${x}px)`;
}

function applyParagraphFrameStyle(element: HTMLElement, frame: ParagraphInfo['frame'], offset?: FloatOffset | null): void {
  if (!frame) return;
  const resolved = offset === undefined ? floatOffset({
    horizontalAnchor: frame.horizontalAnchor, x: frame.xTwips, xAlign: frame.xAlign,
    verticalAnchor: frame.verticalAnchor, y: frame.yTwips, yAlign: frame.yAlign,
  }, undefined) : offset ?? undefined;
  const wrap = frame.wrap ?? 'auto';
  element.dataset.docxFrame = frame.dropCap === 'drop' || frame.dropCap === 'margin' ? 'dropCap' : 'frame';
  if (frame.widthTwips !== undefined) element.style.width = `${frame.widthTwips / 15}px`;
  if (frame.heightTwips !== undefined) {
    const height = `${frame.heightTwips / 15}px`;
    if (frame.heightRule === 'exact') element.style.height = height;
    else element.style.minHeight = height;
  }
  // hSpace / vSpace 是到周围文字的距离。定位了 x / y 时块本身就该在那个位置，间距只留在
  // 朝向正文的右侧与下侧，不能再把块往里推（否则比文档指定的位置偏出一个 hSpace）。
  const horizontal = (frame.horizontalSpaceTwips ?? 0) / 15;
  const vertical = (frame.verticalSpaceTwips ?? 0) / 15;
  if (frame.horizontalSpaceTwips !== undefined || resolved?.xPx) {
    element.style.marginLeft = `${resolved?.xPx ? resolved.xPx : horizontal}px`;
    element.style.marginRight = `${horizontal}px`;
  }
  if (frame.verticalSpaceTwips !== undefined || resolved?.yPx) {
    element.style.marginTop = `${resolved?.yPx ? resolved.yPx : vertical}px`;
    element.style.marginBottom = `${vertical}px`;
  }
  if (resolved) {
    element.dataset.docxFloatOffset = `${resolved.xPx},${resolved.yPx}`;
    applyOffsetShape(element, resolved);
  }
  // notBeside 是「旁边不许有文字」，不浮动才是对的；none 也不浮。
  if (wrap === 'none' || wrap === 'notBeside') return;
  element.style.cssFloat = frame.xAlign === 'right' ? 'right' : 'left';
}

export function deriveLineBoxes(
  rects: Array<{ top: number; height: number; start: number; end: number; width?: number }>,
): LineBox[] {
  const lines: Array<{ top: number; height: number; start: number; end: number; width?: number }> = [];
  for (const rect of rects) {
    const line = lines.find((entry) => Math.abs(entry.top - rect.top) < 1);
    if (line) {
      line.start = Math.min(line.start, rect.start);
      line.end = Math.max(line.end, rect.end);
      line.height = Math.max(line.height, rect.height);
      if (rect.width !== undefined) line.width = Math.max(line.width ?? 0, rect.width);
    } else lines.push({ ...rect });
  }
  return lines.sort((a, b) => a.start - b.start).map(({ height, start, end, width }) => ({
    heightPx: height, startOffset: start, endOffset: end,
    // 量不到宽度时不写这个字段，让调用方看得出是「没量到」而不是「宽度为 0」。
    ...(width === undefined ? {} : { widthPx: width }),
  }));
}

function applyParagraphStyle(element: HTMLElement, paragraph: ParagraphInfo, previous?: ParagraphInfo, includeAfter = true,
  compatibilitySettings?: CompatibilitySettings, divIndents?: Map<number, { left: number; right: number }>): void {
  const effective = paragraph.effective ?? paragraph;
  const spacing = paragraphSpacingPx(previous, { ...paragraph, ...effective }, compatibilitySettings);
  const kinsoku = effectiveKinsoku(effective.kinsoku, compatibilitySettings);
  if (kinsoku !== undefined) {
    element.style.lineBreak = kinsoku ? 'strict' : 'auto';
  }
  if (effective.alignment) element.style.textAlign = ['both', 'distribute'].includes(effective.alignment) ? 'justify' : effective.alignment;
  // w:divId：HTML div 的累计边距加在段落自己的缩进之外（邮件里层层引用的回复就是这么缩进的）。
  const div = effective.divId !== undefined && effective.divId !== null ? divIndents?.get(effective.divId) : undefined;
  const indentLeft = div?.left ? (effective.indentLeft ?? 0) + div.left : effective.indentLeft;
  const indentRight = div?.right ? (effective.indentRight ?? 0) + div.right : effective.indentRight;
  if (indentLeft !== undefined && indentLeft !== null) element.style.marginLeft = twipsToPoints(indentLeft)!;
  if (indentRight !== undefined && indentRight !== null) element.style.marginRight = twipsToPoints(indentRight)!;
  if (spacing.beforePx > 0) element.style.marginTop = `${spacing.beforePx}px`;
  if (includeAfter && spacing.afterPx > 0) element.style.marginBottom = `${spacing.afterPx}px`;
  if ((effective.indentFirstLine !== undefined && effective.indentFirstLine !== null) ||
      (effective.indentHanging !== undefined && effective.indentHanging !== null)) {
    const indent = (effective.indentFirstLine ?? 0) - (effective.indentHanging ?? 0);
    element.style.textIndent = twipsToPoints(indent)!;
  }
  if (effective.lineSpacing !== undefined && effective.lineSpacing !== null) {
    element.style.lineHeight = (effective.lineSpacingRule ?? 'auto') === 'auto'
      ? String(effective.lineSpacing / 240)
      : `${effective.lineSpacing / 20}pt`;
  }
  if (effective.shading?.fill && effective.shading.fill !== 'auto' && /^[0-9a-f]{6}$/i.test(effective.shading.fill)) {
    element.style.backgroundColor = `#${effective.shading.fill}`;
  }
  for (const [side, css] of [
    ['top', 'borderTop'],
    ['left', 'borderLeft'],
    ['right', 'borderRight'],
    ['bottom', 'borderBottom'],
    ['bar', 'borderLeft'],
  ] as const) {
    const border = effective.borders?.[side];
    if (!border || ['none', 'nil'].includes(border.style)) continue;
    const width = `${Math.max(1, border.size) / 8}pt`;
    const color = border.color === 'auto' ? '#000' : /^[0-9a-f]{6}$/i.test(border.color) ? `#${border.color}` : '#000';
    (element.style as CSSStyleDeclaration)[css] = `${width} ${borderStyle(border.style)} ${color}`;
  }
}

/**
 * `w:textAlignment` → CSS `vertical-align`。它管的是一行里字符的**垂直**位置（字号不一时
 * 怎么对齐），和 `w:jc` 的左右对齐是两件事。
 *
 * `center` 只是近似：CSS 的 `middle` 对的是基线加半个 x-height，Word 对的是行的正中。
 * `auto` 是 Word 的默认，不设任何东西、交给浏览器。
 */
const LINE_TEXT_ALIGNMENT_CSS: Record<string, string> = {
  top: 'top', center: 'middle', bottom: 'bottom', baseline: 'baseline',
};

function applyRunStyle(span: HTMLElement, run: RunInfo, lineTextAlignment?: ParagraphFormat['textAlignment']): void {
  const effective = run.effective ?? run;
  // 复杂文种的 run 用 w:bCs / w:iCs，其余用 w:b / w:i —— ECMA-376 把这两组分开管，Word 里
  // 一个只写了 <w:b/> 的阿拉伯文 run 是不加粗的。「复杂文种」这里只认 w:cs / w:rtl 两个显式标记，
  // 不按 Unicode 文种逐字符判断。
  const complex = effective.complexScript === true || effective.rtl === true;
  const bold = complex ? effective.boldComplexScript : effective.bold;
  const italic = complex ? effective.italicComplexScript : effective.italic;
  if (bold !== undefined && bold !== null) span.style.fontWeight = bold ? '700' : '400';
  if (italic !== undefined && italic !== null) span.style.fontStyle = italic ? 'italic' : 'normal';
  // w:noProof：浏览器自己会给 contenteditable 画拼写波浪线，关掉就是它的事。
  if (effective.noProof) span.spellcheck = false;
  if (effective.emphasisMark && effective.emphasisMark !== 'none') {
    span.style.textEmphasisStyle = effective.emphasisMark === 'comma' ? 'sesame' : effective.emphasisMark === 'underDot' ? 'dot' : effective.emphasisMark;
    span.style.textEmphasisPosition = `${effective.emphasisMark === 'underDot' ? 'under' : 'over'} right`;
  }
  const textDecorations = [
    effective.underline ? 'underline' : '',
    effective.strike || effective.doubleStrike ? 'line-through' : '',
  ].filter(Boolean);
  if (textDecorations.length) span.style.textDecoration = textDecorations.join(' ');
  else if (effective.underline === false || effective.strike === false || effective.doubleStrike === false) span.style.textDecoration = 'none';
  if (effective.underlineStyle) span.style.textDecorationStyle = underlineStyleToCss(effective.underlineStyle);
  if (effective.underlineColor && /^[0-9a-f]{6}$/i.test(effective.underlineColor)) span.style.textDecorationColor = `#${effective.underlineColor}`;
  if (effective.fontSize !== undefined) span.style.fontSize = `${effective.fontSize}pt`;
  if (effective.fontFamily || effective.fontFamilyEastAsia) {
    span.style.fontFamily = [effective.fontFamily, effective.fontFamilyEastAsia]
      .filter((name, index, all): name is string => Boolean(name) && all.indexOf(name) === index)
      .map((name) => `"${name}"`).join(', ');
  }
  if (effective.color && /^[0-9a-f]{6}$/i.test(effective.color)) span.style.color = `#${effective.color}`;
  // 段落的 textAlignment 是这一行的默认，**run 自己的上下标 / position 压过它**：两者都落在
  // 同一个 vertical-align 上，而 run 级是更具体的那一层，所以先垫段落级再让 run 级覆盖。
  if (lineTextAlignment && LINE_TEXT_ALIGNMENT_CSS[lineTextAlignment]) {
    span.style.verticalAlign = LINE_TEXT_ALIGNMENT_CSS[lineTextAlignment]!;
  }
  if (effective.verticalAlign === 'subscript' || effective.verticalAlign === 'superscript') span.style.verticalAlign = effective.verticalAlign;
  if (effective.position !== undefined && effective.position !== null) span.style.verticalAlign = `${effective.position / 2}pt`;
  if (effective.smallCaps || effective.allCaps) span.style.fontVariantCaps = effective.allCaps ? 'all-small-caps' : 'small-caps';
  if (effective.allCaps) span.style.textTransform = 'uppercase';
  if (effective.highlight && effective.highlight !== 'none') span.style.backgroundColor = highlightColor(effective.highlight);
  if (effective.characterSpacing !== undefined && effective.characterSpacing !== null) span.style.letterSpacing = `${effective.characterSpacing / 20}pt`;
  if (effective.shading?.fill && effective.shading.fill !== 'auto' && /^[0-9a-f]{6}$/i.test(effective.shading.fill)) {
    span.style.backgroundColor = `#${effective.shading.fill}`;
  }
  if (effective.border && !['none', 'nil'].includes(effective.border.style)) {
    const color = effective.border.color === 'auto'
      ? '#000'
      : /^[0-9a-f]{6}$/i.test(effective.border.color) ? `#${effective.border.color}` : '#000';
    span.style.border = `${Math.max(1, effective.border.size) / 8}pt ${borderStyle(effective.border.style)} ${color}`;
    span.style.paddingInline = '0.05em';
  }
  if (effective.textOutline) {
    const originalColor = effective.color && /^[0-9a-f]{6}$/i.test(effective.color)
      ? `#${effective.color}`
      : span.style.color || 'inherit';
    span.style.webkitTextStroke = '1px currentColor';
    span.style.webkitTextStrokeColor = originalColor;
    span.style.color = 'transparent';
  }
  if (effective.kerning !== undefined && effective.kerning !== null &&
      effective.fontSize !== undefined && effective.fontSize !== null) {
    span.style.fontKerning = effective.fontSize * 2 >= effective.kerning ? 'normal' : 'none';
  }
  // 纵中横（w:eastAsianLayout w:vert）就是 CSS 的 text-combine-upright。横排视图里它不产生
  // 可见差别——Word 也一样，只有竖排时才显示，所以照写，不另做近似。
  if (effective.eastAsianLayout?.vert) span.style.textCombineUpright = 'all';
  if (effective.textShadow || effective.emboss || effective.imprint) {
    const textShadows = [
      ...(effective.textShadow ? ['1px 1px 2px rgba(0, 0, 0, 0.45)'] : []),
      ...(effective.emboss ? ['-1px -1px 1px rgba(255, 255, 255, 0.9)', '1px 1px 1px rgba(0, 0, 0, 0.65)'] : []),
      ...(effective.imprint ? ['1px 1px 1px rgba(255, 255, 255, 0.9)', '-1px -1px 1px rgba(0, 0, 0, 0.65)'] : []),
    ];
    span.style.textShadow = textShadows.join(', ');
  }
}

export interface DocxEditorOptions {
  onChange?: (snapshot: DocumentSnapshot) => void;
  onError?: (error: Error, context: { paragraph: number }) => void;
  showFormattingMarks?: boolean;
  showFieldShading?: boolean;
  showHiddenText?: boolean;
  reviewFilter?: EditorReviewFilter;
  viewMode?: 'continuous' | 'paginated';
}

/** A browser-only, editable view of the supported DOCX paragraph/run/table subset. */
export class DocxEditor {
  private document: DocxDocument;
  private readonly root: HTMLDivElement;
  private readonly options: DocxEditorOptions;
  private readonly paragraphs = new Map<number, {
    element: HTMLParagraphElement;
    content: HTMLSpanElement;
    text: string;
    failed: boolean;
  }>();
  private headerKind: 'default' | 'first' | 'even' = 'default';
  private footerKind: 'default' | 'first' | 'even' = 'default';
  private selected: number | null = null;
  private selectedImageInfo: ImageInfo | null = null;
  private selectedTableCellInfo: TableCellLocation | null = null;
  private selectedRangeInfo: { range: DocumentRange; format: RunFormat } | null = null;
  private composing = false;
  private renderAfterComposition = false;
  private destroyed = false;
  private reviewFilter: NormalizedReviewFilter;
  private readonly metrics: CanvasRenderingContext2D | null;
  private commentRunIds = new Map<string, number[]>();
  private commentParagraphIds = new Map<number, number[]>();
  private revisionRunIds = new Map<string, HTMLElement[]>();
  private revisionParagraphIds = new Map<number, HTMLElement[]>();
  private renderShapeInfos: ShapeInfo[] = [];
  private renderFieldInfos = new Map<number, FieldInfo>();
  private activeRevisionId: number | null = null;
  private viewMode: 'continuous' | 'paginated';
  private measuring = false;
  private compatibilitySettings: CompatibilitySettings = {};
  private divIndents = new Map<number, { left: number; right: number }>();

  /** 同 readCompatibilitySettings：宿主通过 setDocument() 传入的文档对象未必实现它。 */
  private readDivIndents(): Map<number, { left: number; right: number }> {
    const getter = (this.document as DocxDocument & { getWebDivs?: () => WebDivInfo[] }).getWebDivs;
    return typeof getter === 'function' ? webDivIndents(getter.call(this.document)) : new Map();
  }

  private readCompatibilitySettings(): CompatibilitySettings {
    const getter = (this.document as DocxDocument & { getCompatibilitySettings?: () => CompatibilitySettings }).getCompatibilitySettings;
    return typeof getter === 'function' ? getter.call(this.document) : {};
  }

  private dispatchLinkClick(target: HTMLElement): void {
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (!EventClass) return;
    this.root.dispatchEvent(new EventClass('docx-linkclick', {
      bubbles: true,
      detail: {
        url: target.dataset.docxUrl,
        anchor: target.dataset.docxAnchor,
        unsafe: target.dataset.docxUnsafe === 'true',
      },
    }));
  }

  private dispatchCommentClick(target: HTMLElement): void {
    const ids = (target.dataset.docxCommentIds ?? '')
      .split(',')
      .map((value) => Number(value))
      .filter((value) => Number.isSafeInteger(value) && value >= 0);
    if (!ids.length) return;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (!EventClass) return;
    this.root.dispatchEvent(new EventClass('docx-commentclick', {
      bubbles: true,
      detail: { ids },
    }));
  }

  private linkTargetFromSelection(): HTMLElement | null {
    const node = this.root.ownerDocument.getSelection()?.anchorNode;
    if (!node || !this.root.contains(node)) return null;
    const element = node.nodeType === 1 ? node as Element : node.parentElement;
    const target = element?.closest<HTMLElement>('[data-docx-link="1"]');
    return target && this.root.contains(target) ? target : null;
  }

  constructor(container: HTMLElement, document: DocxDocument, options: DocxEditorOptions = {}) {
    this.document = document;
    this.options = options;
    this.viewMode = options.viewMode ?? 'continuous';
    this.reviewFilter = normalizeReviewFilter(options.reviewFilter);
    this.root = container.ownerDocument.createElement('div');
    this.root.className = 'docx-editor';
    this.root.setAttribute('aria-label', '文档编辑区域');
    container.append(this.root);
    this.metrics = this.root.ownerDocument.createElement('canvas').getContext('2d');
    this.compatibilitySettings = this.readCompatibilitySettings();
    this.divIndents = this.readDivIndents();
    this.root.ownerDocument.addEventListener('selectionchange', this.handleSelection);
    this.root.addEventListener('keydown', this.handleRootKeydown);
    this.render();
  }

  get selectedParagraph(): number | null {
    return this.selected;
  }

  get selectedImage(): ImageInfo | null {
    return this.selectedImageInfo;
  }

  get selectedTableCell(): TableCellLocation | null {
    return this.selectedTableCellInfo ? { ...this.selectedTableCellInfo } : null;
  }

  get selectedRange(): DocumentRange | null {
    return this.selectedRangeInfo ? {
      start: { ...this.selectedRangeInfo.range.start },
      end: { ...this.selectedRangeInfo.range.end },
    } : null;
  }

  get selectedRangeFormat(): RunFormat | null {
    return this.selectedRangeInfo ? { ...this.selectedRangeInfo.format } : null;
  }

  getViewMode(): 'continuous' | 'paginated' {
    return this.viewMode;
  }

  setViewMode(mode: 'continuous' | 'paginated'): void {
    if (this.destroyed || mode === this.viewMode) return;
    if (mode === 'paginated') this.flush();
    else this.paragraphs.clear();
    this.viewMode = mode;
    this.render();
  }

  private isMarkupReviewView(): boolean {
    return (this.reviewFilter ?? normalizeReviewFilter(this.options?.reviewFilter)).revisionView === 'markup';
  }

  /** Commit visible text before an external API operation or an export. */
  flush(): void {
    if (this.destroyed) return;
    if (this.viewMode === 'paginated') return;
    if (!this.isMarkupReviewView()) return;
    let changed = false;
    for (const [index, entry] of this.paragraphs) {
      const readSegments = this.readTextSegments(entry.content);
      const rawSegments = readSegments.length > 1 ? readSegments : [this.readText(entry.content)];
      const sanitizedSegments = rawSegments.map((segment) => sanitizeTextWithInfo(segment));
      const sanitizedText = sanitizedSegments.map((segment) => segment.text).join('');
      const sanitized = {
        text: sanitizedText,
        truncated: sanitizedSegments.some((segment) => segment.truncated),
        truncatedAt: sanitizedSegments.find((segment) => segment.truncated)?.truncatedAt,
      };
      if (sanitized.text === entry.text) {
        entry.failed = false;
        continue;
      }
      if (sanitized.truncated && !entry.failed) {
        this.reportError(new Error(`Paragraph text was truncated at ${sanitized.truncatedAt ?? sanitized.text.length} characters.`), index);
      }
      try {
        if (rawSegments.length > 1) this.document.setParagraphTextPreservingHiddenRuns(index, sanitizedSegments.map((segment) => segment.text));
        else this.document.setParagraphText(index, sanitized.text);
        entry.text = sanitized.text;
        entry.failed = false;
        changed = true;
      } catch (error) {
        entry.failed = true;
        this.reportError(error, index);
      }
    }
    if (changed) this.options.onChange?.(this.document.getSnapshot());
  }

  setDocument(document: DocxDocument): void {
    if (this.destroyed) return;
    this.flush();
    this.compatibilitySettings = this.readCompatibilitySettings();
    this.divIndents = this.readDivIndents();
    this.document = document;
    this.selected = null;
    this.selectedImageInfo = null;
    this.selectedTableCellInfo = null;
    this.selectedRangeInfo = null;
    this.composing = false;
    this.renderAfterComposition = false;
    this.paragraphs.clear();
    this.render();
  }

  setReviewFilter(filter: EditorReviewFilter = {}): void {
    if (this.destroyed) return;
    this.reviewFilter ??= normalizeReviewFilter(this.options?.reviewFilter);
    const next = normalizeReviewFilter({
      ...(this.reviewFilter.authors !== undefined ? { authors: this.reviewFilter.authors } : {}),
      showRevisions: this.reviewFilter.showRevisions,
      showComments: this.reviewFilter.showComments,
      revisionView: this.reviewFilter.revisionView,
      ...filter,
    });
    if (reviewFilterEqual(this.reviewFilter, next)) return;
    this.flush();
    this.reviewFilter = next;
    this.render();
  }

  private filteredRevisionIds(): number[] {
    if (!this.reviewFilter.showRevisions) return [];
    const authors = this.reviewFilter.authors ? new Set(this.reviewFilter.authors.map((author) => reviewerBucketKey(author))) : null;
    return this.document
      .getRevisions()
      .filter((revision) => !authors || authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))
      .map((revision) => revision.id);
  }

  private setActiveRevision(id: number | null): void {
    this.activeRevisionId = id;
    if (typeof this.root.querySelectorAll !== 'function') return;
    for (const node of this.root.querySelectorAll<HTMLElement>('[data-docx-revision-ids]')) {
      const ids = (node.dataset.docxRevisionIds ?? '')
        .split(',')
        .map((value) => Number(value))
        .filter((value) => Number.isSafeInteger(value));
      node.classList.toggle('docx-revision-active', id !== null && ids.includes(id));
    }
  }

  private applyRevisionAction(action: () => void): boolean {
    const previous = this.document.revision;
    action();
    if (this.document.revision === previous) return false;
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
    return true;
  }

  acceptRevision(id: number): boolean {
    this.flush();
    const accepted = this.applyRevisionAction(() => this.document.acceptRevision(id));
    if (accepted) this.setActiveRevision(null);
    return accepted;
  }

  rejectRevision(id: number): boolean {
    this.flush();
    const rejected = this.applyRevisionAction(() => this.document.rejectRevision(id));
    if (rejected) this.setActiveRevision(null);
    return rejected;
  }

  acceptAllRevisions(filter: { authors?: ReviewerFilterAuthor[] } = {}): boolean {
    this.flush();
    const accepted = this.applyRevisionAction(() => {
      if (!filter.authors?.length) {
        this.document.acceptAllRevisions();
        return;
      }
      const authors = new Set(filter.authors.map((author) => reviewerBucketKey(author)));
      const operations = this.document
        .getRevisions()
        .filter((revision) => authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))
        .map((revision) => ({ type: 'acceptRevision' as const, id: revision.id }));
      if (!operations.length) return;
      this.document.applyOperations({ operations });
    });
    if (accepted) this.setActiveRevision(null);
    return accepted;
  }

  rejectAllRevisions(filter: { authors?: ReviewerFilterAuthor[] } = {}): boolean {
    this.flush();
    const rejected = this.applyRevisionAction(() => {
      if (!filter.authors?.length) {
        this.document.rejectAllRevisions();
        return;
      }
      const authors = new Set(filter.authors.map((author) => reviewerBucketKey(author)));
      const operations = this.document
        .getRevisions()
        .filter((revision) => authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))
        .map((revision) => ({ type: 'rejectRevision' as const, id: revision.id }));
      if (!operations.length) return;
      this.document.applyOperations({ operations });
    });
    if (rejected) this.setActiveRevision(null);
    return rejected;
  }

  focusNextRevision(): number | null {
    const ids = this.filteredRevisionIds();
    if (!ids.length) {
      this.setActiveRevision(null);
      return null;
    }
    const current = this.activeRevisionId === null ? -1 : ids.indexOf(this.activeRevisionId);
    const nextId = ids[(current + 1) % ids.length]!;
    this.setActiveRevision(nextId);
    const target = this.revisionRunIds.get(String(nextId))?.[0] ?? this.revisionParagraphIds.get(nextId)?.[0];
    target?.focus();
    target?.scrollIntoView({ block: 'nearest' });
    return nextId;
  }

  focusPreviousRevision(): number | null {
    const ids = this.filteredRevisionIds();
    if (!ids.length) {
      this.setActiveRevision(null);
      return null;
    }
    const current = this.activeRevisionId === null ? ids.length : ids.indexOf(this.activeRevisionId);
    const previousId = ids[(current - 1 + ids.length) % ids.length]!;
    this.setActiveRevision(previousId);
    const target = this.revisionRunIds.get(String(previousId))?.[0] ?? this.revisionParagraphIds.get(previousId)?.[0];
    target?.focus();
    target?.scrollIntoView({ block: 'nearest' });
    return previousId;
  }

  focusRevision(id: number): boolean {
    if (!Number.isSafeInteger(id) || id < 0) throw new Error('revision id must be a non-negative integer.');
    if (!this.filteredRevisionIds().includes(id)) return false;
    this.setActiveRevision(id);
    const target = this.revisionRunIds.get(String(id))?.[0] ?? this.revisionParagraphIds.get(id)?.[0];
    target?.focus();
    target?.scrollIntoView({ block: 'nearest' });
    return true;
  }

  /**
   * 重算并写入 `docProps/app.xml` 的统计值。页数与行数要靠排版，所以在这里分页测量一次再交给
   * `DocxDocument.updateDocumentStatistics()`；只要字数的话直接调文档那个方法就够了。
   */
  updateDocumentStatistics(): DocumentStatistics {
    this.flush();
    const blocks = this.document.getBlocks();
    const sections = this.document.getSections();
    let defaultTabStopTwips = 720;
    try {
      defaultTabStopTwips = Math.max(1, Number(this.document.getSettings().defaultTabStop) || 720);
    } catch {
      defaultTabStopTwips = 720;
    }
    const pages = this.paginateDocument(blocks, sections, defaultTabStopTwips);
    return this.document.updateDocumentStatistics({ pagination: paginationInfoFromPages(pages, blocks) });
  }

  updateFields(options: { kinds?: FieldKind[]; now?: Date; filename?: string } = {}): boolean {
    const result = updateFieldsUntilStable(() => {
      const blocks = this.document.getBlocks();
      const sections = this.document.getSections();
      let defaultTabStopTwips = 720;
      try {
        defaultTabStopTwips = Math.max(1, Number(this.document.getSettings().defaultTabStop) || 720);
      } catch {
        defaultTabStopTwips = 720;
      }
      const pages = this.paginateDocument(blocks, sections, defaultTabStopTwips);
      return this.document.updateFields({
        ...options,
        pagination: paginationInfoFromPages(pages, blocks),
      });
    });
    if (result.updated) this.render();
    return result.updated;
  }

  render(): void {
    if (this.destroyed) return;
    this.reviewFilter ??= normalizeReviewFilter(this.options?.reviewFilter);
    if (this.composing) {
      this.renderAfterComposition = true;
      return;
    }
    const range = this.captureDocumentRange();
    const activeImageId = (this.root.ownerDocument.activeElement as HTMLElement | null)?.closest<HTMLElement>('[data-image]')?.dataset.image
      ?? this.selectedImageInfo?.id
      ?? null;
    this.flush();
    try {
      this.renderFieldInfos = new Map(this.document.getFields().map((field) => [field.index, field]));
    } catch {
      this.renderFieldInfos = new Map();
    }
    if (this.viewMode === 'continuous') this.applyPageSetup();
    this.paragraphs.clear();
    this.commentRunIds ??= new Map();
    this.commentParagraphIds ??= new Map();
    this.commentRunIds.clear();
    this.commentParagraphIds.clear();
    this.revisionRunIds ??= new Map();
    this.revisionParagraphIds ??= new Map();
    this.revisionRunIds.clear();
    this.revisionParagraphIds.clear();
    const reviewContext: ReviewRenderContext = {
      ...(this.reviewFilter.authors ? { authors: new Set(this.reviewFilter.authors.map((author) => reviewerBucketKey(author))) } : {}),
      deletedTextByRun: new Map(),
      revisionColors: new Map(),
    };
    try {
      for (const revision of this.document.getRevisions({ kinds: ['deletion', 'move'] })) {
        if (revision.paragraph < 0 || revision.run === undefined || !revision.deletedText) continue;
        if (revision.kind === 'move' && revision.move?.side !== 'from') continue;
        if (reviewContext.authors && !reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(revision.author)))) continue;
        const key = `${revision.paragraph}:${revision.run}`;
        reviewContext.deletedTextByRun.set(key, `${reviewContext.deletedTextByRun.get(key) ?? ''}${revision.deletedText}`);
      }
      for (const comment of this.reviewFilter.showComments ? this.document.getComments() : []) {
        if (reviewContext.authors && !reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(comment.author)))) continue;
        if (!comment.anchor || comment.anchor.sourcePartPath !== this.document.mainDocumentPath) continue;
        if ('runs' in comment.anchor) {
          for (const run of comment.anchor.runs) {
            const key = `${comment.anchor.paragraph}:${run}`;
            const ids = this.commentRunIds.get(key) ?? [];
            if (!ids.includes(comment.id)) ids.push(comment.id);
            this.commentRunIds.set(key, ids);
          }
        } else {
          for (let paragraph = comment.anchor.startParagraph; paragraph <= comment.anchor.endParagraph; paragraph++) {
            const ids = this.commentParagraphIds.get(paragraph) ?? [];
            if (!ids.includes(comment.id)) ids.push(comment.id);
            this.commentParagraphIds.set(paragraph, ids);
          }
        }
      }
      const fragment = this.root.ownerDocument.createDocumentFragment();
      const canRenderHeaderFooter = typeof this.root.ownerDocument.createElement === 'function';
      let defaultTabStopTwips = 720;
      try {
        defaultTabStopTwips = Math.max(1, Number(this.document.getSettings().defaultTabStop) || 720);
      } catch {
        defaultTabStopTwips = 720;
      }
      this.renderShapeInfos = this.document.getShapes();
      if (this.viewMode === 'paginated') {
        this.renderPaginated(fragment, reviewContext, defaultTabStopTwips);
      } else {
        const blocks = this.document.getBlocks();
        const sections = this.continuousSections();
        if (sections.length > 1) {
          this.appendContinuousSections(fragment, blocks, sections, defaultTabStopTwips, reviewContext, canRenderHeaderFooter);
        } else {
          if (canRenderHeaderFooter) fragment.append(this.makeHeaderFooter('header'));
          this.appendBlocks(fragment, blocks, defaultTabStopTwips, reviewContext);
          if (canRenderHeaderFooter) fragment.append(this.makeHeaderFooter('footer'));
        }
      }
      this.root.replaceChildren(fragment);
      if (this.selected !== null && !this.paragraphs.has(this.selected)) this.selected = null;
      const nextSelected = activeImageId ? this.document.getImages().find((image) => image.id === activeImageId) ?? null : null;
      this.selectImage(nextSelected);
      if (activeImageId) {
        Array.from(this.root.querySelectorAll<HTMLElement>('[data-image]'))
          .find((node) => node.dataset.image === activeImageId)
          ?.focus({ preventScroll: true });
      }
      if (range) this.restoreDocumentRange(range);
      this.updateRangeSelection(this.captureDocumentRange());
      this.setActiveRevision(this.activeRevisionId);
    } finally {
      this.renderShapeInfos = [];
      reviewContext.deletedTextByRun.clear();
      reviewContext.revisionColors.clear();
    }
  }

  destroy(): void {
    if (this.destroyed) return;
    this.flush();
    this.destroyed = true;
    this.root.ownerDocument.removeEventListener('selectionchange', this.handleSelection);
    this.root.removeEventListener('keydown', this.handleRootKeydown);
    this.root.remove();
    this.paragraphs.clear();
    this.selectedRangeInfo = null;
  }

  private readText(element: HTMLElement): string {
    return this.readTextSegments(element).join('').replace(/\n$/, '');
  }

  private readTextSegments(element: HTMLElement): string[] {
    const segments = [''];
    const walk = (node: Node): string => {
      if (node.nodeType === 3) {
        segments[segments.length - 1] += node.textContent ?? '';
        return '';
      }
      if (node.nodeType !== 1) return '';
      const current = node as HTMLElement;
      if (current.dataset.image || current.dataset.docxMark !== undefined) return '';
      if (current.dataset.docxDeleted !== undefined) return '';
      if (current.dataset.docxHiddenPreserved !== undefined) {
        segments.push('');
        return '';
      }
      if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return '';
      if (current.tagName === 'BR') {
        segments[segments.length - 1] += '\n';
        return '\n';
      }
      const startLength = segments[segments.length - 1]!.length;
      for (const child of Array.from(current.childNodes)) walk(child);
      const text = segments[segments.length - 1]!.slice(startLength);
      if (['DIV', 'P'].includes(current.tagName) && text) segments[segments.length - 1] += '\n';
      return text;
    };
    walk(element);
    if (segments[segments.length - 1]?.endsWith('\n')) {
      segments[segments.length - 1] = segments[segments.length - 1]!.slice(0, -1);
    }
    return segments;
  }

  private twipsToPx(value: number | undefined): number | undefined {
    return value === undefined ? undefined : value / 15;
  }

  private focusContent(element: HTMLElement): void {
    element.focus({ preventScroll: true });
    const selection = this.root.ownerDocument.getSelection();
    if (!selection) return;
    const range = this.root.ownerDocument.createRange();
    range.selectNodeContents(element);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  private borderCss(border: BorderFormat | undefined): string | undefined {
    if (!border) return undefined;
    if (border.none || ['nil', 'none'].includes(border.style ?? '')) return 'none';
    const width = border.size !== undefined ? `${Math.max(1, eighthPointsToPx(border.size))}px` : '1px';
    const color = border.color && /^[0-9a-f]{6}$/i.test(border.color) ? `#${border.color}` : '#dbe3ed';
    return `${width} solid ${color}`;
  }

  /**
   * `w:tl2br` / `w:tr2bl`：单元格里的对角线。CSS 没有对角边框，但角落关键字的
   * `linear-gradient` 正好能画：规范规定 `to bottom left` 的渐变线 50% 处**恰好穿过
   * 左上和右下两个角**（`to bottom right` 则穿过右上与左下），所以一个硬色标就是一条
   * 精确的对角线，而且不管单元格是不是正方形都对。
   *
   * 用背景而不是插一个元素，还顺带躲开了第 4 条：背景不是节点，没法流回文档，也不需要
   * `contentEditable=false`。代价是线型只能画成实线——渐变表达不了 dashed / dotted。
   */
  private diagonalCss(border: BorderFormat | undefined, direction: 'to bottom left' | 'to bottom right'): string | undefined {
    if (!border) return undefined;
    if (border.none || ['nil', 'none'].includes(border.style ?? '')) return undefined;
    const half = Math.max(0.5, (border.size !== undefined ? Math.max(1, eighthPointsToPx(border.size)) : 1) / 2);
    const color = border.color && /^[0-9a-f]{6}$/i.test(border.color) ? `#${border.color}` : '#dbe3ed';
    const from = `calc(50% - ${half}px)`;
    const to = `calc(50% + ${half}px)`;
    return `linear-gradient(${direction}, transparent ${from}, ${color} ${from}, ${color} ${to}, transparent ${to})`;
  }

  private cellBorder(side: 'top' | 'right' | 'bottom' | 'left', table: TableFormat | undefined, cell: CellFormat | undefined,
    row: number, col: number, rowSpan: number, colSpan: number, rowCount: number, colCount: number): string | undefined {
    const explicit = cell?.borders?.[side];
    if (explicit) return this.borderCss(explicit);
    const borders = table?.borders;
    if (!borders) return undefined;
    if (side === 'top' && row > 0 && borders.insideH) return this.borderCss(borders.insideH);
    if (side === 'bottom' && row + rowSpan < rowCount && borders.insideH) return this.borderCss(borders.insideH);
    if (side === 'left' && col > 0 && borders.insideV) return this.borderCss(borders.insideV);
    if (side === 'right' && col + colSpan < colCount && borders.insideV) return this.borderCss(borders.insideV);
    return this.borderCss(borders[side]);
  }

  private widthCss(width: WidthFormat | undefined): string | undefined {
    return normalizeWidth(width);
  }

  private paddingCss(width: WidthFormat | undefined): string | undefined {
    if (!width) return undefined;
    if (width.type === 'pct') return `${width.value / 50}%`;
    if (width.type === 'dxa') return `${twipsToPx(width.value)}px`;
    return undefined;
  }

  private applyTableStyle(table: HTMLTableElement, format: TableFormat | undefined, offset?: FloatOffset | null): void {
    if (!format) return;
    table.style.tableLayout = format.layout === 'fixed' ? 'fixed' : 'auto';
    const width = this.widthCss(format.width);
    if (width) table.style.width = width;
    if (format.alignment === 'center') table.style.marginInline = 'auto';
    if (format.alignment === 'right') { table.style.marginLeft = 'auto'; table.style.marginRight = '0'; }
    if (format.alignment === 'left') { table.style.marginLeft = '0'; table.style.marginRight = 'auto'; }
    if (format.indent !== undefined && (!format.alignment || format.alignment === 'left')) table.style.marginLeft = `${twipsToPx(format.indent)}px`;
    if (format.shading?.fill) table.style.backgroundColor = `#${format.shading.fill}`;
    // w:tblCellSpacing：ECMA-376 把「单元格之间」和「单元格与表格边缘之间」说成同一个值，
    // 这正是 CSS border-spacing 的语义，所以一对一映射。有间距就必须 separate —— collapse
    // 下 border-spacing 被忽略。w:trPr 上的行级 cellSpacing 这里落不下来：CSS 的
    // border-spacing 只能加在表格上，没有行级对应物。
    const spacing = this.paddingCss(format.cellSpacing);
    if (spacing) {
      table.style.borderCollapse = 'separate';
      table.style.borderSpacing = spacing;
    }
    // w:tblpPr：浮动表格，正文绕着它排。和 framePr 一样把环绕交给浏览器 float —— 分页测量
    // 也是用 float 占位（见 measureParagraphForPagination），两边是同一套行为。
    // w:tblpX / w:tblpY 那套按页面或页边距定位的绝对坐标本期不实现：需要相对页框定位，而
    // 连续视图没有页框；浮动方向按 tblpXSpec 取左右，其余照左浮。
    const floating = format.floatingPosition;
    if (floating) {
      table.dataset.docxFloating = '1';
      table.style.cssFloat = floating.xSpec === 'right' ? 'right' : 'left';
      // 先清掉 jc 可能设下的 marginInline:auto（它会抵消浮动），再设间距 —— margin-inline 是
      // marginLeft / marginRight 的简写，顺序反了会把刚设的间距一起清掉。
      table.style.marginInline = '';
      const gap = (value: number | undefined) => (value === undefined ? undefined : `${twipsToPx(value)}px`);
      const left = gap(floating.leftFromText);
      const right = gap(floating.rightFromText);
      const top = gap(floating.topFromText);
      const bottom = gap(floating.bottomFromText);
      if (right) table.style.marginRight = right;
      if (bottom) table.style.marginBottom = bottom;
      // x / y 定位，含义同 applyParagraphFrameStyle 的 offset：偏移取代那一侧的 from-text 间距。
      const resolved = offset === undefined ? floatOffset({
        horizontalAnchor: floating.horizontalAnchor, x: floating.x, xAlign: floating.xSpec,
        verticalAnchor: floating.verticalAnchor, y: floating.y, yAlign: floating.ySpec,
      }, undefined) : offset ?? undefined;
      // 定位了 x / y 时间距不再把表推离指定位置（同 applyParagraphFrameStyle）。
      if (left || resolved?.xPx) table.style.marginLeft = resolved?.xPx ? `${resolved.xPx}px` : left!;
      if (top || resolved?.yPx) table.style.marginTop = resolved?.yPx ? `${resolved.yPx}px` : top!;
      if (resolved) {
        table.dataset.docxFloatOffset = `${resolved.xPx},${resolved.yPx}`;
        applyOffsetShape(table, resolved);
      }
    }
    if (format.caption) {
      const caption = table.createCaption();
      caption.textContent = format.caption;
    }
  }

  private applyRowStyle(tr: HTMLTableRowElement, row: TableRowInfo): void {
    // 优先用算上表格样式之后的行格式；没有（调用方没注入解析器）才退回直接格式。
    const format = row.effective ?? row.format;
    if (!format) return;
    if (format.height) tr.style.height = `${twipsToPx(format.height.value)}px`;
    if (format.header) tr.dataset.header = 'true';
  }

  private applyCellStyle(td: HTMLTableCellElement, cell: CellFormat | undefined, table: TableFormat | undefined,
    row: number, col: number, rowSpan: number, colSpan: number, rowCount: number, colCount: number): void {
    td.style.borderTop = this.cellBorder('top', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderTop;
    td.style.borderRight = this.cellBorder('right', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderRight;
    td.style.borderBottom = this.cellBorder('bottom', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderBottom;
    td.style.borderLeft = this.cellBorder('left', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderLeft;
    if (cell?.shading?.fill) td.style.backgroundColor = `#${cell.shading.fill}`;
    // 两条对角线叠在底纹之上：background-image 画在 background-color 上面。
    const diagonals = [
      this.diagonalCss(cell?.borders?.tl2br, 'to bottom left'),
      this.diagonalCss(cell?.borders?.tr2bl, 'to bottom right'),
    ].filter((entry): entry is string => Boolean(entry));
    if (diagonals.length) td.style.backgroundImage = diagonals.join(', ');
    if (cell?.verticalAlign) td.style.verticalAlign = cell.verticalAlign;
    if (cell?.width) td.style.width = this.widthCss(cell.width) ?? '';
    if (cell?.margin?.top) td.style.paddingTop = this.paddingCss(cell.margin.top) ?? '';
    if (cell?.margin?.right) td.style.paddingRight = this.paddingCss(cell.margin.right) ?? '';
    if (cell?.margin?.bottom) td.style.paddingBottom = this.paddingCss(cell.margin.bottom) ?? '';
    if (cell?.margin?.left) td.style.paddingLeft = this.paddingCss(cell.margin.left) ?? '';
    if (cell?.noWrap) td.style.whiteSpace = 'nowrap';
    if (cell?.textDirection?.toLowerCase().includes('tb') || cell?.textDirection?.toLowerCase().includes('bt')) td.style.writingMode = 'vertical-rl';
  }

  /**
   * `w:tblPrEx`：这一行的表格属性例外。优先级是 单元格自己的 `tcPr` > 行的 `tblPrEx` > 表格的
   * `tblPr` 与表格样式，所以判断「单元格有没有自己的值」要看**直接格式**——有效格式里已经叠进了
   * 样式给的边框和边距，拿它判断的话例外永远压不过样式。
   *
   * 能落到 CSS 的是边框、底纹、单元格边距这三样；`jc` / `tblInd` / `tblW` / `tblCellSpacing`
   * 是整行平移或加宽，HTML 表格的行没有这种能力，只读写、不渲染。
   */
  private applyTableException(td: HTMLTableCellElement, exception: TableException, direct: CellFormat | undefined,
    row: number, col: number, rowSpan: number, colSpan: number, rowCount: number, colCount: number): void {
    if (exception.borders) {
      const sides = { top: 'borderTop', right: 'borderRight', bottom: 'borderBottom', left: 'borderLeft' } as const;
      for (const [side, property] of Object.entries(sides) as Array<[keyof typeof sides, typeof sides[keyof typeof sides]]>) {
        if (direct?.borders?.[side]) continue;
        const css = this.cellBorder(side, { borders: exception.borders }, undefined, row, col, rowSpan, colSpan, rowCount, colCount);
        if (css) td.style[property] = css;
      }
    }
    if (exception.shading?.fill && !direct?.shading?.fill && /^[0-9a-f]{6}$/i.test(exception.shading.fill)) {
      td.style.backgroundColor = `#${exception.shading.fill}`;
    }
    if (exception.cellMargin) {
      const sides = { top: 'paddingTop', right: 'paddingRight', bottom: 'paddingBottom', left: 'paddingLeft' } as const;
      for (const [side, property] of Object.entries(sides) as Array<[keyof typeof sides, typeof sides[keyof typeof sides]]>) {
        const margin = exception.cellMargin[side];
        if (!margin || direct?.margin?.[side]) continue;
        const css = this.paddingCss(margin);
        if (css) td.style[property] = css;
      }
    }
  }

  private makeTable(block: LayoutTable, rowIndices: number[], defaultTabStopTwips: number,
    reviewContext: ReviewRenderContext, floatPlacement?: FloatOffset | null): HTMLTableElement {
    const table = this.root.ownerDocument.createElement('table');
    table.className = 'docx-table';
    // 优先用算上表格样式之后的表格级格式；没有（调用方没注入解析器）才退回直接格式。
    this.applyTableStyle(table, block.effective ?? block.format, floatPlacement);
    const body = table.createTBody();
    for (const rowIndex of rowIndices) {
      const row = block.rows[rowIndex];
      if (!row) continue;
      const tr = body.insertRow();
      this.applyRowStyle(tr, row);
      const rowFormat = row.effective ?? row.format;
      // w:gridBefore 跳过的网格列也占列号。从 0 起算的话，这一行所有 gridStart / gridEnd
      // 都会偏小，宿主按下标定位单元格就会落到错的列上。
      let colIndex = rowFormat?.gridBefore ?? 0;
      // 跳过的那块在 Word 里是实打实的空白，用一个空 td 占住。它不是文档内容，所以标成
      // contentEditable=false，也不画边框。
      const gridSkip = (span: number, width: WidthFormat | undefined) => {
        const td = tr.insertCell();
        td.colSpan = Math.max(1, span);
        td.contentEditable = 'false';
        td.dataset.docxGridSkip = String(span);
        td.setAttribute('aria-hidden', 'true');
        td.style.borderStyle = 'none';
        const css = this.widthCss(width);
        if (css) td.style.width = css;
      };
      if (rowFormat?.gridBefore) gridSkip(rowFormat.gridBefore, rowFormat.widthBefore);
      for (const cell of row.cells) {
        const logicalStart = colIndex;
        colIndex += Math.max(1, cell.colSpan);
        if (cell.isMergeContinuation) continue;
        const td = tr.insertCell();
        td.dataset.tableCell = 'true';
        td.dataset.gridStart = String(logicalStart);
        td.dataset.gridEnd = String(logicalStart + Math.max(1, cell.colSpan));
        td.dataset.rowStart = String(rowIndex);
        td.dataset.rowEnd = String(rowIndex + Math.max(1, cell.rowSpan));
        td.colSpan = Math.max(1, cell.colSpan);
        if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
        // 优先用算上表格样式之后的格式；没有（调用方没注入解析器）才退回直接格式。
        this.applyCellStyle(td, cell.effective ?? cell.format, block.effective ?? block.format, rowIndex, logicalStart,
          Math.max(1, cell.rowSpan), Math.max(1, cell.colSpan), block.rows.length, block.grid.length);
        if (rowFormat?.tableException) {
          this.applyTableException(td, rowFormat.tableException, cell.format, rowIndex, logicalStart,
            Math.max(1, cell.rowSpan), Math.max(1, cell.colSpan), block.rows.length, block.grid.length);
        }
        this.appendBlocks(td, cell.blocks, defaultTabStopTwips, reviewContext);
      }
      if (rowFormat?.gridAfter) gridSkip(rowFormat.gridAfter, rowFormat.widthAfter);
    }
    return table;
  }

  private appendBlocks(parent: Node, blocks: DocumentBlock[], defaultTabStopTwips: number, reviewContext: ReviewRenderContext): void {
    let previous: ParagraphInfo | undefined;
    for (let index = 0; index < blocks.length; index++) {
      const block = blocks[index]!;
      if (block.type === 'paragraph') {
        const next = blocks[index + 1];
        parent.appendChild(this.makeParagraph(block.paragraph, defaultTabStopTwips, reviewContext, undefined, previous,
          next?.type !== 'paragraph'));
        previous = block.paragraph;
      } else if (block.type === 'table') {
        const table = this.makeTable(block, block.rows.map((_, index) => index), defaultTabStopTwips, reviewContext);
        parent.appendChild(table);
        if (block.format?.description) {
          const description = this.root.ownerDocument.createElement('div');
          description.className = 'sr-only';
          description.id = `docx-table-desc-${Math.random().toString(36).slice(2)}`;
          description.textContent = block.format.description;
          table.setAttribute('aria-describedby', description.id);
          parent.appendChild(description);
        }
      } else {
        const marker = this.root.ownerDocument.createElement('div');
        marker.className = 'docx-break-marker';
        marker.textContent = block.type === 'pageBreak'
          ? '—— 分页符 ——'
          : `—— 分节符（${block.breakType}）——`;
        marker.setAttribute('role', 'note');
        marker.setAttribute('aria-label', marker.textContent);
        parent.appendChild(marker);
      }
    }
  }

  private withMeasuring<T>(render: () => T): T {
    const previous = this.measuring;
    this.measuring = true;
    try {
      return render();
    } finally {
      this.measuring = previous;
    }
  }

  private measureParagraphForPagination(paragraph: ParagraphInfo, area: ParagraphMeasureArea, context: MeasureContext): LineBox[] {
    const host = this.root.ownerDocument.createElement('div');
    host.style.cssText = `position:absolute;visibility:hidden;left:-100000px;width:${Math.max(1, area.widthPx)}px;`;
    for (const wrap of area.wraps.filter((candidate) => candidate.carried)) {
      const spacer = this.root.ownerDocument.createElement('div');
      spacer.style.width = wrap.wrap === 'topAndBottom' ? '100%' : `${Math.min(area.widthPx, wrap.widthPx)}px`;
      spacer.style.height = `${wrap.heightPx}px`;
      if (wrap.wrap !== 'topAndBottom') spacer.style.cssFloat = 'left';
      if (wrap.offsetXPx || wrap.offsetYPx) spacer.style.shapeOutside = `inset(${wrap.offsetYPx ?? 0}px 0 0 ${wrap.offsetXPx ?? 0}px)`;
      host.append(spacer);
    }
    const measured = this.withMeasuring(() => this.makeParagraph(paragraph, context.defaultTabStopTwips, {
      authors: undefined,
      deletedTextByRun: new Map(),
      revisionColors: new Map(),
    }));
    host.append(measured);
    this.root.append(host);
    const textNodes: Array<{ node: Text; start: number }> = [];
    const collect = (node: Node, start: number): number => {
      if (node.nodeType === 3) {
        const text = node as Text;
        textNodes.push({ node: text, start });
        return start + text.data.length;
      }
      if (node.nodeType !== 1 || (node as HTMLElement).dataset.docxMark !== undefined) return start;
      let offset = start;
      for (const child of Array.from(node.childNodes)) offset = collect(child, offset);
      return offset;
    };
    collect(measured.querySelector('.docx-paragraph-content') ?? measured, 0);
    const content = measured.querySelector('.docx-paragraph-content') ?? measured;
    const paragraphRange = this.root.ownerDocument.createRange();
    paragraphRange.selectNodeContents(content);
    const lineRects = Array.from(paragraphRange.getClientRects()).filter((rect, index, all) =>
      index === all.findIndex((entry) => Math.abs(entry.top - rect.top) < 1));
    const nodeAt = (offset: number): { node: Text; offset: number } | undefined => {
      const entry = textNodes.find((item, index) => offset <= item.start + item.node.data.length && (index === textNodes.length - 1 || offset < textNodes[index + 1]!.start));
      return entry ? { node: entry.node, offset: Math.max(0, Math.min(entry.node.data.length, offset - entry.start)) } : undefined;
    };
    const charRect = (offset: number): DOMRect | undefined => {
      const start = nodeAt(offset);
      const end = nodeAt(Math.min(offset + 1, paragraph.text.length));
      if (!start || !end || start === end) return undefined;
      const range = this.root.ownerDocument.createRange();
      range.setStart(start.node, start.offset);
      range.setEnd(end.node, end.offset);
      return Array.from(range.getClientRects())[0];
    };
    const rects: Array<{ top: number; height: number; start: number; end: number; width?: number }> = [];
    for (const line of lineRects) {
      let start = 0;
      let end = paragraph.text.length;
      while (start < end) {
        const middle = Math.floor((start + end) / 2);
        const rect = charRect(middle);
        if (!rect || rect.top < line.top - 1) start = middle + 1; else end = middle;
      }
      const lineStart = start;
      start = lineStart;
      end = paragraph.text.length;
      while (start < end) {
        const middle = Math.ceil((start + end) / 2);
        const rect = charRect(Math.max(lineStart, middle - 1));
        if (rect && Math.abs(rect.top - line.top) < 1) start = middle; else end = middle - 1;
      }
      if (start > lineStart) {
        rects.push({ top: line.top, height: line.height, width: line.width, start: lineStart, end: start });
      }
    }
    host.remove();
    if (!rects.length) return paragraph.text ? [{ heightPx: 18, startOffset: 0, endOffset: paragraph.text.length }] : [];
    return deriveLineBoxes(rects);
  }

  private measureTableRowForPagination(table: LayoutTable, rowIndex: number, widthPx: number, context: MeasureContext): number {
    const host = this.root.ownerDocument.createElement('div');
    host.style.cssText = `position:absolute;visibility:hidden;left:-100000px;width:${Math.max(1, widthPx)}px;`;
    const rendered = this.withMeasuring(() => this.makeTable(table, [rowIndex], context.defaultTabStopTwips, {
      authors: undefined,
      deletedTextByRun: new Map(),
      revisionColors: new Map(),
    }));
    rendered.style.width = '100%';
    host.append(rendered);
    this.root.append(host);
    const tr = rendered.rows[0];
    const cellHeight = tr
      ? Array.from(tr.cells).reduce((height, cell) => Math.max(height, cell.getBoundingClientRect().height), 0)
      : 0;
    const height = Math.max(cellHeight, tr?.getBoundingClientRect().height ?? 0);
    host.remove();
    return Math.max(0, height);
  }

  private sliceParagraph(paragraph: ParagraphInfo, start: number, end: number): ParagraphInfo {
    let offset = 0;
    const includedRunIndexes = new Set<number>();
    const runs = paragraph.runs.flatMap((run) => {
      const runStart = offset;
      offset += run.text.length;
      const from = Math.max(start, runStart);
      const to = Math.min(end, offset);
      if (!run.text.length && runStart >= start && (runStart < end || (end === paragraph.text.length && runStart === end))) return [{ ...run }];
      if (to <= from) return [];
      includedRunIndexes.add(run.index);
      return [{ ...run, text: run.text.slice(from - runStart, to - runStart) }];
    });
    const math = paragraph.math?.flatMap((info, index) =>
      includedRunIndexes.has(info.runOffset) ||
      (end === paragraph.text.length && info.runOffset >= paragraph.runs.length)
        ? [{ ...info, [MATH_RENDER_INDEX]: (info as MathInfo & { [MATH_RENDER_INDEX]?: number })[MATH_RENDER_INDEX] ?? index }]
        : []);
    return { ...paragraph, text: paragraph.text.slice(start, end), runs, ...(math ? { math } : {}) };
  }

  private makePageContent(page: PageBox, section: SectionInfo, blocks: DocumentBlock[], paragraphs: ParagraphInfo[],
    defaultTabStopTwips: number, reviewContext: ReviewRenderContext, lineNumbers: Array<number | null> = []): HTMLElement {
    const body = this.root.ownerDocument.createElement('div');
    body.className = 'docx-page-content';
    const gapPx = Math.max(0, section.columns.space * 96 / 1440);
    const columnWidths = columnWidthsPx(section);
    body.style.display = 'grid';
    body.style.gridTemplateColumns = columnWidths.map((value) => `${value}px`).join(' ');
    body.style.columnGap = `${gapPx}px`;
    body.style.alignItems = 'start';
    if (section.verticalAlignment === 'center' || section.verticalAlignment === 'bottom') {
      body.style.minHeight = `${Math.max(0, pageBoxPx(section).heightPx - pageBoxPx(section).padding.top - pageBoxPx(section).padding.bottom)}px`;
      body.style.alignContent = section.verticalAlignment === 'center' ? 'center' : 'end';
    }
    const columns = columnWidths.map((_, index) => {
      const column = this.root.ownerDocument.createElement('div');
      column.className = 'docx-page-column';
      column.dataset.column = String(index);
      body.append(column);
      return column;
    });
    const renderedTables = new Set<string>();
    const previousByColumn = new Map<number, ParagraphInfo>();
    for (let itemIndex = 0; itemIndex < page.items.length; itemIndex++) {
      const item = page.items[itemIndex]!;
      const columnIndex = Math.min(item.column ?? 0, columns.length - 1);
      const column = columns[columnIndex]!;
      if (item.type === 'line') {
        const paragraph = paragraphs.find((entry) => entry.index === item.paragraph);
        if (paragraph) {
          // paginate() 交出来的行高已经按 docGrid（和段落的 snapToGrid）吸附过了——它每个 line 项都出自
          // measureParagraph，吸附就在那里。这里原先再吸附一次：幂等所以看不出来，但它是第二份真相，
          // 段落级开关漏传到这一处时没有任何用例会红。所以不再吸附，照用。
          const lineHeightPx = item.line.heightPx;
          const firstLine = item.line.startOffset === 0;
          const lastLine = item.line.endOffset >= paragraph.text.length;
          const hasFollowingParagraph = page.items.slice(itemIndex + 1).some((candidate) =>
            candidate.type === 'line' && candidate.paragraph !== paragraph.index);
          const fragment = {
            ...paragraph,
            ...(firstLine ? {} : { spacingBefore: null }),
            ...(lastLine ? {} : { spacingAfter: null }),
            ...(paragraph.effective ? {
              effective: {
                ...paragraph.effective,
                ...(firstLine ? {} : { spacingBefore: null }),
                ...(lastLine ? {} : { spacingAfter: null }),
              },
            } : {}),
          };
          // 浮动段落一行一个浮动元素往下堆：x 偏移每行都要，y 偏移只给第一行，后面的行自然堆在它下面。
          const placement = paragraph.frame
            ? (item.floatOffset ? { xPx: item.floatOffset.xPx, yPx: firstLine ? item.floatOffset.yPx : 0 } : null)
            : undefined;
          const paragraphElement = this.makeParagraph(
            this.sliceParagraph(fragment, item.line.startOffset, item.line.endOffset),
            defaultTabStopTwips,
            reviewContext,
            lineHeightPx,
            firstLine ? previousByColumn.get(columnIndex) : undefined,
            lastLine && !hasFollowingParagraph,
            placement,
          );
          const lineNumber = lineNumbers[itemIndex];
          if (lineNumber !== null && lineNumber !== undefined) {
            const marker = this.root.ownerDocument.createElement('span');
            marker.className = 'docx-line-number';
            marker.contentEditable = 'false';
            marker.setAttribute('data-docx-mark', '1');
            marker.dataset.docxLineNumber = String(lineNumber);
            marker.setAttribute('aria-hidden', 'true');
            marker.textContent = String(lineNumber);
            marker.style.position = 'absolute';
            marker.style.top = '0';
            marker.style.whiteSpace = 'nowrap';
            const distance = Math.max(0, section.lineNumbering?.distance ?? 0) * 96 / 1440;
            const rtl = paragraph.effective?.bidi === true;
            marker.style[rtl ? 'right' : 'left'] = `-${distance + 2}px`;
            paragraphElement.style.position = 'relative';
            paragraphElement.append(marker);
          }
          column.append(paragraphElement);
          if (lastLine) previousByColumn.set(columnIndex, paragraph);
        }
      } else if (item.type === 'tableRow' && !renderedTables.has(`${columnIndex}:${item.table}`)) {
        const block = blocks[item.table];
        if (block?.type === 'table') {
          const rowIndices = page.items
            .filter((candidate): candidate is Extract<FlowItem, { type: 'tableRow' }> =>
              candidate.type === 'tableRow' && candidate.table === item.table && (candidate.column ?? 0) === columnIndex)
            .map((candidate) => candidate.row);
          column.append(this.makeTable(block, rowIndices, defaultTabStopTwips, reviewContext,
            (block.effective ?? block.format)?.floatingPosition ? (item.floatOffset ?? null) : undefined));
          renderedTables.add(`${columnIndex}:${item.table}`);
        }
      }
    }
    return body;
  }

  private paginateDocument(blocks: DocumentBlock[], sections: SectionInfo[], defaultTabStopTwips: number): PageBox[] {
    this.compatibilitySettings = this.readCompatibilitySettings();
    this.divIndents = this.readDivIndents();
    const measurer = {
      measureParagraph: (paragraph: ParagraphInfo, area: ParagraphMeasureArea, context: MeasureContext) =>
        this.measureParagraphForPagination(paragraph, area, context),
      measureTableRow: (table: LayoutTable, _row: TableRowInfo, rowIndex: number, widthPx: number, context: MeasureContext) =>
        this.measureTableRowForPagination(table, rowIndex, widthPx, context),
    };
    return paginate(blocks, sections, measurer, { defaultTabStopTwips, compatibilitySettings: this.compatibilitySettings });
  }

  private renderPaginated(parent: Node, reviewContext: ReviewRenderContext, defaultTabStopTwips: number): void {
    const blocks = this.document.getBlocks();
    const sections = this.document.getSections();
    const paragraphs = blocks.filter((block): block is Extract<DocumentBlock, { type: 'paragraph' }> => block.type === 'paragraph')
      .map((block) => block.paragraph);
    const pages = this.paginateDocument(blocks, sections, defaultTabStopTwips);
    const lineNumbers = lineNumbersFor(pages, paragraphs, sections);
    for (const page of pages) {
      const section = this.document.getSection(page.section);
      const firstPhysicalPage = page.index === 0 || pages[page.index - 1]?.section !== page.section;
      const pageElement = this.root.ownerDocument.createElement('section');
      pageElement.className = 'docx-page';
      pageElement.contentEditable = 'false';
      pageElement.dataset.page = String(page.number);
      const box = pageBoxPx(section);
      pageElement.style.width = `${box.widthPx}px`;
      pageElement.style.minHeight = `${box.heightPx}px`;
      pageElement.style.padding = `${box.padding.top}px ${box.padding.right}px ${box.padding.bottom}px ${box.padding.left}px`;
      pageElement.style.boxSizing = 'border-box';
      const borders = section.pageBorders;
      const borderProperties = {
        top: 'borderTop',
        right: 'borderRight',
        bottom: 'borderBottom',
        left: 'borderLeft',
      } as const;
      const showBorder = borders && (borders.display === undefined || borders.display === 'allPages' ||
        (borders.display === 'firstPage' ? firstPhysicalPage : !firstPhysicalPage));
      if (showBorder && borders?.offsetFrom !== 'text') {
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const border = borders?.[side];
          if (!border || border.style === 'none' || border.style === 'nil') continue;
          const width = Math.max(0, border.size) * 96 / 5760;
          pageElement.style[borderProperties[side]] =
            `${Math.max(1, width)}px ${border.style === 'single' ? 'solid' : border.style} ${border.color === 'auto' ? 'currentColor' : `#${border.color}`}`;
        }
      }
      pageElement.append(this.makeHeaderFooter('header', page.section, page.number, pages.length, firstPhysicalPage));
      const body = this.makePageContent(page, section, blocks, paragraphs, defaultTabStopTwips, reviewContext, lineNumbers[page.index] ?? []);
      if (showBorder && borders?.offsetFrom === 'text') {
        for (const side of ['top', 'right', 'bottom', 'left'] as const) {
          const border = borders[side];
          if (!border || border.style === 'none' || border.style === 'nil') continue;
          const width = Math.max(0, border.size) * 96 / 5760;
          body.style[borderProperties[side]] =
            `${Math.max(1, width)}px ${border.style === 'single' ? 'solid' : border.style} ${border.color === 'auto' ? 'currentColor' : `#${border.color}`}`;
        }
      }
      pageElement.append(body, this.makeHeaderFooter('footer', page.section, page.number, pages.length, firstPhysicalPage));
      (parent as DocumentFragment).append(pageElement);
    }
  }

  private continuousSections(): SectionInfo[] {
    try {
      return this.document.getSections();
    } catch {
      return [];
    }
  }

  /**
   * 多节文档的连续视图：每一节放进自己的 `div.docx-section`，用那一节的纸张宽度、左右页边距和
   * 分栏。原先整篇只套第一节，于是「纵向正文 + 一节横向宽表格」里横向那节被压在纵向的宽度里。
   *
   * 分节靠 `getBlocks()` 里的 `sectionBreak` 标记切：标记的 `section` 是它**结束**的那一节，
   * 标记之后的内容属于下一节（第 23 条的那个坑）。标记本身留在它结束的那一节末尾。页眉页脚
   * 仍只画第一节的——连续视图没有页，「每页的页眉」无从谈起；按节看页眉请用分页视图。
   */
  private appendContinuousSections(parent: Node, blocks: DocumentBlock[], sections: SectionInfo[], defaultTabStopTwips: number,
    reviewContext: ReviewRenderContext, withHeaderFooter: boolean): void {
    const groups: DocumentBlock[][] = [[]];
    for (const block of blocks) {
      groups.at(-1)!.push(block);
      if (block.type === 'sectionBreak') groups.push([]);
    }
    if (!groups.at(-1)!.length && groups.length > 1) groups.pop();
    const toPx = (twips: number) => `${Math.max(0, twips * 96 / 1440)}px`;
    const wrap = (section: SectionInfo | undefined, index: number): HTMLElement => {
      const element = this.root.ownerDocument.createElement('div');
      element.className = 'docx-section';
      element.dataset.section = String(index);
      if (!section) return element;
      element.dataset.orientation = section.orientation;
      element.style.boxSizing = 'border-box';
      element.style.maxWidth = toPx(section.pageWidth);
      element.style.marginLeft = 'auto';
      element.style.marginRight = 'auto';
      element.style.paddingLeft = toPx(section.margins.left);
      element.style.paddingRight = toPx(section.margins.right);
      if (section.columns.count > 1) {
        element.style.columnCount = String(section.columns.count);
        element.style.columnGap = toPx(section.columns.space);
      }
      return element;
    };
    if (withHeaderFooter) {
      const header = wrap(sections[0], 0);
      header.append(this.makeHeaderFooter('header'));
      parent.appendChild(header);
    }
    groups.forEach((group, index) => {
      const element = wrap(sections[index] ?? sections.at(-1), index);
      this.appendBlocks(element, group, defaultTabStopTwips, reviewContext);
      parent.appendChild(element);
    });
    if (withHeaderFooter) {
      const footer = wrap(sections[0], 0);
      footer.append(this.makeHeaderFooter('footer'));
      parent.appendChild(footer);
    }
  }

  private applyPageSetup(): void {
    const rootStyle = (this.root as unknown as { style?: CSSStyleDeclaration }).style;
    const paper = this.root.parentElement as HTMLElement | null;
    if (rootStyle) {
      rootStyle.columnCount = '';
      rootStyle.columnGap = '';
    }
    if (paper) {
      paper.style.maxWidth = '';
      paper.style.paddingTop = '';
      paper.style.paddingRight = '';
      paper.style.paddingBottom = '';
      paper.style.paddingLeft = '';
      delete paper.dataset.orientation;
    }
    let section: SectionInfo | undefined;
    try { section = this.document.getSection(0); } catch { section = undefined; }
    if (!section) return;
    const toPx = (twips: number) => `${Math.max(0, twips * 96 / 1440)}px`;
    const sections = this.continuousSections();
    if (sections.length > 1) {
      // 多节：纸张取最宽那一节的宽度，左右页边距与分栏交给各节自己的容器（appendContinuousSections）。
      if (paper) {
        paper.style.maxWidth = toPx(Math.max(...sections.map((entry) => entry.pageWidth)));
        paper.style.paddingTop = toPx(section.margins.top);
        paper.style.paddingRight = '0px';
        paper.style.paddingBottom = toPx(sections.at(-1)!.margins.bottom);
        paper.style.paddingLeft = '0px';
        paper.dataset.orientation = section.orientation;
      }
      return;
    }
    if (paper) {
      paper.style.maxWidth = toPx(section.pageWidth);
      paper.style.paddingTop = toPx(section.margins.top);
      paper.style.paddingRight = toPx(section.margins.right);
      paper.style.paddingBottom = toPx(section.margins.bottom);
      paper.style.paddingLeft = toPx(section.margins.left);
      paper.dataset.orientation = section.orientation;
    }
    if (rootStyle) {
      rootStyle.columnCount = String(Math.max(1, section.columns.count));
      rootStyle.columnGap = toPx(section.columns.space);
    }
  }

  private makeHeaderFooter(type: 'header' | 'footer', sectionIndex = 0, pageNumber?: number, pageCount?: number, firstPhysicalPage = false): HTMLElement {
    const selectedKind = type === 'header' ? this.headerKind : this.footerKind;
    let kind = selectedKind;
    let part: string | undefined;
    let format: string | undefined;
    try {
      const section = this.document.getSection(sectionIndex);
      ({ kind, part } = selectHeaderFooter(type === 'header' ? section.headers : section.footers, {
        selectedKind, pageNumber, firstPhysicalPage, titlePage: section.titlePage,
      }));
      format = section.pageNumbering?.format;
    } catch {
      part = undefined;
    }
    let blocks = type === 'header'
      ? this.document.getHeaderBlocks(sectionIndex, kind)
      : this.document.getFooterBlocks(sectionIndex, kind);
    if (pageNumber !== undefined) blocks = replacePageFields(blocks, pageNumber, pageCount, format);
    const partXml = part ? this.document.getPartXml(part) : '';
    const plainEditable = !!part && !/<w:(tbl|fldSimple|fldChar|drawing|hyperlink|object|pict|sdt|customXml|smartTag|ins|del)\b/.test(partXml);
    const area = this.root.ownerDocument.createElement('div');
    area.className = `docx-${type}`;
    const label = this.root.ownerDocument.createElement('div');
    label.className = 'docx-header-footer-label';
    label.id = `docx-${type}-${kind}-label`;
    label.textContent = `${type === 'header' ? '页眉' : '页脚'}（${kind}）`;
    const editable = this.root.ownerDocument.createElement('div');
    editable.contentEditable = pageNumber === undefined && plainEditable ? 'true' : 'false';
    editable.className = 'docx-header-footer-text';
    editable.setAttribute('role', 'textbox');
    editable.setAttribute('aria-multiline', 'true');
    editable.setAttribute('aria-labelledby', label.id);
    const renderedText = blocks.flatMap(block => block.type === 'paragraph' ? [block.paragraph.text] : []).join('\n');
    const normalizedRenderedText = renderedText.replace(/\r\n?/g, '\n').trimEnd();
    editable.textContent = renderedText;
    if (pageNumber !== undefined || !plainEditable) editable.setAttribute('aria-readonly', 'true');
    editable.addEventListener('blur', () => {
      if (pageNumber !== undefined || !plainEditable) return;
      const text = editable.innerText.replace(/\r\n?/g, '\n').trimEnd();
      if (text === normalizedRenderedText) return;
      if (type === 'header') this.document.setHeaderText(0, text, kind);
      else this.document.setFooterText(0, text, kind);
      this.render();
      this.options.onChange?.(this.document.getSnapshot());
    });
    area.append(label, editable);
    return area;
  }

  setHeaderKind(kind: 'default' | 'first' | 'even'): void {
    this.headerKind = kind;
    this.render();
  }

  setFooterKind(kind: 'default' | 'first' | 'even'): void {
    this.footerKind = kind;
    this.render();
  }

  private makeParagraph(paragraph: ParagraphInfo, defaultTabStopTwips: number, reviewContext: ReviewRenderContext,
    lineHeightPx?: number, previous?: ParagraphInfo, includeAfter = true, floatPlacement?: FloatOffset | null): HTMLParagraphElement {
    const element = this.root.ownerDocument.createElement('p');
    const content = this.root.ownerDocument.createElement('span');
    element.className = 'docx-paragraph';
    element.dataset.paragraph = String(paragraph.index);
    element.style.whiteSpace = 'pre-wrap';
    element.style.minHeight = '1.5em';
    applyParagraphFrameStyle(element, paragraph.frame, floatPlacement);
    applyParagraphStyle(element, paragraph, previous, includeAfter, this.compatibilitySettings, this.divIndents);
    if (lineHeightPx !== undefined) {
      element.style.lineHeight = `${lineHeightPx}px`;
      element.style.minHeight = `${lineHeightPx}px`;
    }
    if (paragraph.style) element.dataset.style = paragraph.style;
    if (paragraph.numbering) {
      const marker = this.root.ownerDocument.createElement('span');
      marker.className = 'docx-numbering';
      marker.contentEditable = 'false';
      marker.setAttribute('data-docx-mark', '1');
      marker.setAttribute('aria-hidden', 'true');
      marker.style.userSelect = 'none';
      marker.style.pointerEvents = 'none';
      let bulletImage: HTMLImageElement | undefined;
      if (paragraph.numbering.image) {
        try {
          bulletImage = this.root.ownerDocument.createElement('img');
          bulletImage.src = this.document.getImageDataUrl(paragraph.numbering.image);
          bulletImage.alt = '';
          bulletImage.setAttribute('aria-hidden', 'true');
          bulletImage.style.width = `${paragraph.numbering.image.widthPx}px`;
          bulletImage.style.height = `${paragraph.numbering.image.heightPx}px`;
        } catch {
          bulletImage = undefined;
        }
      }
      if (bulletImage) marker.append(bulletImage);
      else marker.textContent = paragraph.numbering.text;
      marker.dataset.suffix = paragraph.numbering.suffix;
      if (paragraph.numbering.runFormat?.fontFamily) marker.style.fontFamily = paragraph.numbering.runFormat.fontFamily;
      if (paragraph.numbering.runFormat?.bold !== undefined) marker.style.fontWeight = paragraph.numbering.runFormat.bold ? '700' : '400';
      if (paragraph.numbering.runFormat?.italic !== undefined) marker.style.fontStyle = paragraph.numbering.runFormat.italic ? 'italic' : 'normal';
      if (paragraph.numbering.runFormat?.underline !== undefined) marker.style.textDecoration = paragraph.numbering.runFormat.underline ? 'underline' : 'none';
      if (paragraph.numbering.runFormat?.fontSize !== undefined) marker.style.fontSize = `${paragraph.numbering.runFormat.fontSize}pt`;
      if (paragraph.numbering.runFormat?.color && /^[0-9a-f]{6}$/i.test(paragraph.numbering.runFormat.color)) marker.style.color = `#${paragraph.numbering.runFormat.color}`;
      element.append(marker);
      const left = this.twipsToPx(paragraph.numbering.indentLeft);
      const hanging = this.twipsToPx(paragraph.numbering.indentHanging);
      if (left !== undefined) element.style.marginLeft = `${left}px`;
      if (hanging !== undefined) element.style.textIndent = `${-hanging}px`;
      element.dataset.numberingLevel = String(paragraph.numbering.level);
      element.dataset.numberingFormat = paragraph.numbering.format;
    }
    content.className = 'docx-paragraph-content';
    content.contentEditable = this.isMarkupReviewView() ? 'true' : 'false';
    content.spellcheck = false;
    content.setAttribute('role', 'textbox');
    content.setAttribute('aria-multiline', 'true');
    content.setAttribute('aria-label', `第 ${paragraph.index + 1} 段`);
    if (!this.isMarkupReviewView()) content.setAttribute('aria-readonly', 'true');
    if (paragraph.numbering) content.setAttribute('aria-description', `列表项 ${paragraph.numbering.text}，级别 ${paragraph.numbering.level + 1}`);
    const runRevisionIds = new Set<number>();
    const hasRunRevision = this.reviewFilter.showRevisions && paragraph.runs.some((run) =>
      run.revisions?.some((revision) => {
        const visible = !reviewContext.authors || reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(revision.author)));
        if (visible) runRevisionIds.add(revision.id);
        return visible;
      }));
    const visibleParagraphRevision = this.reviewFilter.showRevisions && paragraph.paragraphRevision &&
      (!reviewContext.authors || reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(paragraph.paragraphRevision.author))))
      ? paragraph.paragraphRevision
      : undefined;
    if ((hasRunRevision || visibleParagraphRevision) && this.reviewFilter.revisionView === 'markup') {
      const marker = this.makeMark('▎', '修订变更条');
      marker.classList.add('docx-change-bar');
      marker.style.position = 'absolute';
      marker.style.left = '-0.9em';
      marker.style.top = '0';
      marker.style.bottom = '0';
      marker.style.display = 'flex';
      marker.style.alignItems = 'stretch';
      marker.style.color = this.reviewColor(visibleParagraphRevision?.author, reviewContext);
      marker.style.opacity = '0.9';
      marker.textContent = '│';
      marker.title = '';
      marker.setAttribute('role', 'presentation');
      marker.setAttribute('aria-hidden', 'true');
      element.append(marker);
      const markerRevisionIds = new Set<number>(runRevisionIds);
      if (visibleParagraphRevision) markerRevisionIds.add(visibleParagraphRevision.id);
      if (markerRevisionIds.size) {
        element.dataset.docxRevisionIds = [...markerRevisionIds].join(',');
      }
      if (!this.measuring) {
        for (const id of markerRevisionIds) {
          const list = this.revisionParagraphIds.get(id) ?? [];
          list.push(element);
          this.revisionParagraphIds.set(id, list);
        }
      }
    }
    const paragraphCommentIds = this.commentParagraphIds.get(paragraph.index);
    if (paragraphCommentIds?.length) {
      element.classList.add('docx-comment-anchor');
      element.dataset.docxCommentIds = paragraphCommentIds.join(',');
    }
    element.addEventListener('mousedown', (event) => {
      const target = event.target as Node | null;
      if (target && content.contains(target)) return;
      event.preventDefault();
      this.focusContent(content);
    });
    let currentLineOffsetPx = 0;
    const math = paragraph.math ?? [];
    const appendMath = (info: MathInfo, index: number): void => {
      content.append(this.makeMath(info, (info as MathInfo & { [MATH_RENDER_INDEX]?: number })[MATH_RENDER_INDEX] ?? index));
    };
    for (const run of paragraph.runs) {
      math.forEach((info, index) => { if (info.runOffset === run.index) appendMath(info, index); });
      const visibleRun = this.reviewScopedRun(paragraph.index, run, reviewContext);
      currentLineOffsetPx = this.appendRun(content, paragraph, visibleRun, reviewContext, defaultTabStopTwips, currentLineOffsetPx);
      const hiddenRun = this.runIsHidden(run) && !this.options.showHiddenText;
      if (!hiddenRun) this.appendDeletedRunVisualization(content, paragraph.index, visibleRun, reviewContext);
      if (!hiddenRun && run.noteReference) {
        const marker = this.root.ownerDocument.createElement('sup');
        marker.className = 'docx-note-ref';
        marker.contentEditable = 'false';
        marker.setAttribute('data-docx-mark', '1');
        marker.textContent = run.noteReference.marker;
        marker.setAttribute('aria-label', `${run.noteReference.kind} reference ${run.noteReference.marker}`);
        content.append(marker);
      }
      if (!hiddenRun) {
        for (const image of run.images ?? (run.image ? [run.image] : [])) content.append(this.makeImage(paragraph.index, image));
        for (const shape of this.renderShapeInfos.filter((item) => item.paragraph === paragraph.index && item.run === run.index)) {
          content.append(this.makeShape(shape, defaultTabStopTwips, reviewContext));
        }
      }
    }
    // 末尾公式 = 没有任何 run 的索引等于它的 runOffset。用 runs.length 判断会随分页切片变化，
    // 导致 runOffset 命中切片内某个 run 时这里再渲染一次。
    math.forEach((info, index) => {
      if (!paragraph.runs.some(run => run.index === info.runOffset)) appendMath(info, index);
    });

    if (!paragraph.runs.length && !math.length) content.textContent = paragraph.text;
    if (this.options.showFormattingMarks) content.append(this.makeMark('¶', '段落标记'));
    element.append(content);
    if (!this.measuring) this.paragraphs.set(paragraph.index, { element, content, text: sanitizeText(this.readText(content)), failed: false });
    content.addEventListener('focus', () => this.selectParagraph(paragraph.index));
    content.addEventListener('blur', () => { if (!this.composing) this.flush(); });
    content.addEventListener('compositionstart', () => { this.composing = true; });
    content.addEventListener('compositionend', () => {
      this.composing = false;
      if (this.renderAfterComposition) {
        this.renderAfterComposition = false;
        this.render();
      } else if (this.root.ownerDocument.activeElement !== content) {
        this.flush();
      }
    });
    content.addEventListener('copy', (event) => this.handleClipboardCopy(event));
    content.addEventListener('cut', (event) => this.handleClipboardCut(event, content));
    content.addEventListener('paste', (event) => this.handleClipboardPaste(event, content));
    // Do not allow rich HTML or embedded objects from drag-and-drop either.
    content.addEventListener('drop', (event) => { event.preventDefault(); });
    content.addEventListener('keydown', (event) => {
      if (this.handleHistoryShortcut(event)) return;
      if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        const target = this.linkTargetFromSelection();
        if (target) {
          event.preventDefault();
          this.dispatchLinkClick(target);
          return;
        }
      }
      if (event.key === 'Enter' && !event.isComposing && !this.composing) {
        event.preventDefault();
        this.insertText(content, '\n');
      }
      if (event.key === 'Tab' && !event.isComposing && !this.composing && paragraph.numbering) {
        this.flush();
        const current = this.document.getParagraph?.(paragraph.index)
          ?? this.document.getParagraphs().find((item) => item.index === paragraph.index);
        if (!current?.numbering) return;
        const nextLevel = current.numbering.level + (event.shiftKey ? -1 : 1);
        if (nextLevel < 0 || nextLevel > 8) return;
        event.preventDefault();
        this.document.setParagraphLevel(paragraph.index, event.shiftKey ? -1 : 1);
        this.render();
        this.options.onChange?.(this.document.getSnapshot());
      }
      if (event.key === 'Tab') {
        if (this.moveToAdjacentCell(element, event.shiftKey ? -1 : 1)) event.preventDefault();
      }
      const caret = this.caretIn(element);
      if (event.key === 'ArrowLeft' && caret?.start === 0 && caret.end === 0) {
        if (this.moveToAdjacentCell(element, -1)) event.preventDefault();
      }
      if (event.key === 'ArrowRight' && caret && caret.start === caret.end && caret.end === this.readText(element).length) {
        if (this.moveToAdjacentCell(element, 1)) event.preventDefault();
      }
      if (event.key === 'ArrowUp' && caret?.start === 0 && caret.end === 0) {
        if (this.moveVerticalCell(element, -1)) event.preventDefault();
      }
      if (event.key === 'ArrowDown' && caret && caret.start === caret.end && caret.end === this.readText(element).length) {
        if (this.moveVerticalCell(element, 1)) event.preventDefault();
      }
      if ((event.ctrlKey || event.metaKey) && ['b', 'i', 'u'].includes(event.key.toLowerCase())) {
        event.preventDefault();
      }
    });
    content.addEventListener('click', (event) => {
      const target = (event.target as Element | null)?.closest<HTMLElement>('[data-docx-link="1"]');
      if (!target || !this.root.contains(target)) return;
      event.preventDefault();
      if (!(event.ctrlKey || event.metaKey)) return;
      this.dispatchLinkClick(target);
    });
    content.addEventListener('click', (event) => {
      const target = (event.target as Element | null)?.closest<HTMLElement>('[data-docx-comment-ids]');
      if (!target || !this.root.contains(target)) return;
      this.dispatchCommentClick(target);
    });
    content.addEventListener('beforeinput', (event) => {
      if (event.inputType === 'historyUndo') {
        event.preventDefault();
        this.applyHistory('undo');
        return;
      }
      if (event.inputType === 'historyRedo') {
        event.preventDefault();
        this.applyHistory('redo');
        return;
      }
      if (this.selectionTouchesField(content)) {
        event.preventDefault();
        return;
      }
      if (this.preventFieldDeletion(content, event)) return;
      if (!event.isComposing && ['insertParagraph', 'insertLineBreak'].includes(event.inputType)) {
        event.preventDefault();
        this.insertText(content, '\n');
      }
      if (event.inputType.startsWith('format')) event.preventDefault();
    });
    return element;
  }

  private makeShape(shape: ShapeInfo, defaultTabStopTwips: number, reviewContext: ReviewRenderContext): HTMLElement {
    const wrapper = this.root.ownerDocument.createElement(shape.placement === 'floating' ? 'div' : 'span');
    wrapper.className = `docx-shape docx-shape-${shape.kind}`;
    wrapper.contentEditable = 'false';
    wrapper.dataset.docxShape = shape.id;
    wrapper.style.display = shape.placement === 'floating' ? 'block' : 'inline-block';
    wrapper.style.width = `${Math.max(48, shape.widthPx || 160)}px`;
    wrapper.style.minHeight = `${Math.max(48, shape.heightPx || 90)}px`;
    wrapper.style.maxWidth = '100%';
    wrapper.style.margin = shape.placement === 'floating' ? '8px 12px 8px 0' : '0 2px';
    wrapper.style.verticalAlign = 'text-bottom';
    wrapper.style.position = 'relative';
    wrapper.style.boxSizing = 'border-box';
    wrapper.setAttribute('aria-label', shape.alt ?? shape.geometry ?? shape.kind);
    const width = Math.max(48, shape.widthPx || 160);
    const height = Math.max(48, shape.heightPx || 90);
    const svgNs = 'http://www.w3.org/2000/svg';
    const svg = this.root.ownerDocument.createElementNS(svgNs, 'svg');
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`);
    svg.setAttribute('width', '100%');
    svg.setAttribute('height', '100%');
    svg.setAttribute('aria-hidden', 'true');
    svg.style.position = 'absolute';
    svg.style.inset = '0';
    svg.style.overflow = 'visible';
    if (shape.chart && shape.chart.kind !== 'unsupported') {
      this.renderChart(svg, shape.chart, width, height);
      wrapper.append(svg);
      return wrapper;
    }
    const pathData = shape.customGeometry
      ? customGeometryPath(shape.customGeometry, width, height)
      : shape.geometry
        ? presetGeometryPath(shape.geometry, width, height, new Map(shape.adjustments?.map(({ name, value }) => [name, value])))
        : undefined;
    const path = this.root.ownerDocument.createElementNS(svgNs, pathData ? 'path' : 'rect');
    if (pathData) path.setAttribute('d', pathData);
    else {
      path.setAttribute('x', '0');
      path.setAttribute('y', '0');
      path.setAttribute('width', String(width));
      path.setAttribute('height', String(height));
    }
    const fill = shape.fill;
    if (fill?.type === 'none') path.setAttribute('fill', 'none');
    else if (fill?.type === 'gradient' && fill.stops?.length) {
      const defs = this.root.ownerDocument.createElementNS(svgNs, 'defs');
      const gradient = this.root.ownerDocument.createElementNS(svgNs, 'linearGradient');
      const gradientId = `shape-gradient-${String(shape.id).replace(/[^a-zA-Z0-9_-]/g, '-')}`;
      gradient.setAttribute('id', gradientId);
      gradient.setAttribute('x1', '0%');
      gradient.setAttribute('y1', '0%');
      gradient.setAttribute('x2', '100%');
      gradient.setAttribute('y2', '0%');
      if (fill.angle) gradient.setAttribute('gradientTransform', `rotate(${fill.angle} 0.5 0.5)`);
      for (const stop of fill.stops) {
        const stopElement = this.root.ownerDocument.createElementNS(svgNs, 'stop');
        stopElement.setAttribute('offset', `${stop.position * 100}%`);
        stopElement.setAttribute('stop-color', stop.color);
        gradient.appendChild(stopElement);
      }
      defs.appendChild(gradient);
      svg.appendChild(defs);
      path.setAttribute('fill', `url(#${gradientId})`);
    } else if (fill?.type === 'picture' && fill.imagePartPath) {
      try {
        const contentType = contentTypeForExtension(fill.imagePartPath.split('.').pop() ?? '');
        if (isBrowserRenderableContentType(contentType)) {
          const bytes = this.document.getPartBytes(fill.imagePartPath);
          const pattern = this.root.ownerDocument.createElementNS(svgNs, 'pattern');
          const patternId = `shape-pattern-${String(shape.id).replace(/[^a-zA-Z0-9_-]/g, '-')}`;
          pattern.setAttribute('id', patternId);
          pattern.setAttribute('width', '100%');
          pattern.setAttribute('height', '100%');
          pattern.setAttribute('patternContentUnits', 'objectBoundingBox');
          const image = this.root.ownerDocument.createElementNS(svgNs, 'image');
          image.setAttribute('width', '1');
          image.setAttribute('height', '1');
          image.setAttribute('preserveAspectRatio', 'xMidYMid slice');
          image.setAttribute('href', dataUrlForBytes(bytes, contentType!));
          pattern.appendChild(image);
          const defs = this.root.ownerDocument.createElementNS(svgNs, 'defs');
          defs.appendChild(pattern);
          svg.appendChild(defs);
          path.setAttribute('fill', `url(#${patternId})`);
        } else {
          path.setAttribute('fill', '#f7f9fd');
          path.setAttribute('stroke-dasharray', shape.line?.dash ?? '6 4');
        }
      } catch {
        path.setAttribute('fill', '#f7f9fd');
        path.setAttribute('stroke-dasharray', shape.line?.dash ?? '6 4');
      }
    } else if (fill?.type === 'picture') {
      path.setAttribute('fill', '#f7f9fd');
      path.setAttribute('stroke-dasharray', shape.line?.dash ?? '6 4');
    } else path.setAttribute('fill', fill?.color ?? 'none');
    if (shape.line?.color) path.setAttribute('stroke', shape.line.color);
    else path.setAttribute('stroke', fill?.type === 'picture' ? '#c7d3e5' : 'none');
    if (shape.line?.widthPx !== undefined) path.setAttribute('stroke-width', String(shape.line.widthPx));
    if (shape.line?.dash) path.setAttribute('stroke-dasharray', shape.line.dash);
    // 线、连接线、弧、括号这类开放路径只描边：SVG 会把开放路径首尾连起来填色。没写线条颜色时
    // 它们就整个看不见了，所以给一个黑色兜底（Word 的默认线条色）。
    if (!shape.customGeometry && shape.geometry && presetGeometryIsOpen(shape.geometry)) {
      path.setAttribute('fill', 'none');
      if (!shape.line?.color) path.setAttribute('stroke', '#000000');
    }
    if (shape.shadow) path.setAttribute('filter', `url(#${this.appendShadowFilter(svg, shape.shadow, `shape-shadow-${shape.id}`)})`);
    if (shape.rotation || shape.flipH || shape.flipV) {
      const transforms = [`translate(${width / 2} ${height / 2})`];
      if (shape.rotation) transforms.push(`rotate(${shape.rotation})`);
      transforms.push(`scale(${shape.flipH ? -1 : 1} ${shape.flipV ? -1 : 1})`);
      transforms.push(`translate(${-width / 2} ${-height / 2})`);
      path.setAttribute('transform', transforms.join(' '));
    }
    svg.appendChild(path);
    for (const [index, child] of (shape.children ?? []).entries()) {
      const childWidth = Math.max(1, child.widthPx || 1);
      const childHeight = Math.max(1, child.heightPx || 1);
      const childPathData = child.customGeometry
        ? customGeometryPath(child.customGeometry, childWidth, childHeight)
        : child.geometry
          ? presetGeometryPath(child.geometry, childWidth, childHeight)
          : undefined;
      const childPath = this.root.ownerDocument.createElementNS(svgNs, childPathData ? 'path' : 'rect');
      if (childPathData) childPath.setAttribute('d', childPathData);
      else {
        childPath.setAttribute('x', '0');
        childPath.setAttribute('y', '0');
        childPath.setAttribute('width', String(childWidth));
        childPath.setAttribute('height', String(childHeight));
      }
      childPath.setAttribute('data-docx-shape-child', String(index));
      if (child.fill?.type === 'none') childPath.setAttribute('fill', 'none');
      else childPath.setAttribute('fill', child.fill?.color ?? 'none');
      if (child.line?.color) childPath.setAttribute('stroke', child.line.color);
      else childPath.setAttribute('stroke', 'none');
      if (child.line?.widthPx !== undefined) childPath.setAttribute('stroke-width', String(child.line.widthPx));
      if (child.line?.dash) childPath.setAttribute('stroke-dasharray', child.line.dash);
      if (child.shadow) childPath.setAttribute('filter', `url(#${this.appendShadowFilter(svg, child.shadow, `shape-shadow-${shape.id}-${index}`)})`);
      const transforms = [`translate(${child.offsetXPx} ${child.offsetYPx})`];
      if (child.rotation || child.flipH || child.flipV) {
        transforms.push(`translate(${childWidth / 2} ${childHeight / 2})`);
        if (child.rotation) transforms.push(`rotate(${child.rotation})`);
        transforms.push(`scale(${child.flipH ? -1 : 1} ${child.flipV ? -1 : 1})`);
        transforms.push(`translate(${-childWidth / 2} ${-childHeight / 2})`);
      }
      childPath.setAttribute('transform', transforms.join(' '));
      svg.appendChild(childPath);
      if (child.paragraphs?.length) {
        svg.appendChild(this.makeShapeChildText(child.paragraphs, child.offsetXPx, child.offsetYPx, childWidth, childHeight));
      } else if (child.text) {
        const text = this.root.ownerDocument.createElementNS(svgNs, 'text');
        text.setAttribute('x', String(child.offsetXPx + childWidth / 2));
        text.setAttribute('y', String(child.offsetYPx + childHeight / 2));
        text.setAttribute('text-anchor', 'middle');
        text.setAttribute('dominant-baseline', 'middle');
        text.textContent = child.text;
        svg.appendChild(text);
      }
    }
    wrapper.append(svg);
    if (shape.hasTextContent) {
      for (const paragraph of this.document.getShapeParagraphs(shape.id)) {
        const element = this.root.ownerDocument.createElement('p');
        element.className = 'docx-shape-paragraph';
        element.style.whiteSpace = 'pre-wrap';
        element.style.position = 'relative';
        element.style.zIndex = '1';
        applyParagraphStyle(element, paragraph, undefined, true, this.compatibilitySettings, this.divIndents);
        let offset = 0;
        // 文本框里的公式和正文一样按 runOffset 插在对应 run 之前；没有对应 run 的落在段尾。
        const math = paragraph.math ?? [];
        for (const run of paragraph.runs) {
          for (const info of math) if (info.runOffset === run.index) element.append(this.makeMath(info, undefined));
          offset = this.appendRun(element, paragraph, run, reviewContext, defaultTabStopTwips, offset);
        }
        for (const info of math) {
          if (!paragraph.runs.some((run) => run.index === info.runOffset)) element.append(this.makeMath(info, undefined));
        }
        if (!paragraph.runs.length && !math.length) element.textContent = paragraph.text;
        wrapper.append(element);
      }
    }
    return wrapper;
  }

  private renderChart(svg: SVGElement, chart: NonNullable<ShapeInfo['chart']>, width: number, height: number): void {
    const doc = this.root.ownerDocument;
    const ns = 'http://www.w3.org/2000/svg';
    const element = (name: string) => doc.createElementNS(ns, name);
    const add = (name: string, attrs: Record<string, string>, parent = svg) => {
      const node = element(name);
      for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
      parent.appendChild(node);
      return node;
    };
    if (chart.kind === 'radar') {
      this.renderRadarChart(add, chart, width, height);
      return;
    }
    const titleHeight = chart.title ? 18 : 4;
    const seriesType = (series: ChartSeriesInfo) => series.type ?? chart.kind;
    const primarySeries = chart.series.filter((series) => series.axis !== 'secondary');
    const secondarySeries = chart.series.filter((series) => series.axis === 'secondary');
    const barSeries = chart.series.filter((series) => seriesType(series) === 'bar');
    const horizontalBars = barSeries.length > 0 && chart.barDirection === 'bar';
    const left = (horizontalBars ? chart.axes?.category?.visible : chart.axes?.value?.visible) === false ? 8 : 34;
    // 有序列名时图例画在最底下一行，要单独留出 14px；原先图例和类目标签挤在同一条带里，互相压字。
    const legendHeight = chart.kind !== 'pie' && chart.kind !== 'doughnut' && chart.series.some((series) => series.name) ? 14 : 0;
    const bottom = ((horizontalBars ? chart.axes?.value?.visible : chart.axes?.category?.visible) === false ? 8 : 24) + legendHeight;
    const pieLegendColumns = Math.max(1, Math.floor(Math.max(1, width - left - 8) / 70));
    const pieLegendRows = chart.kind === 'pie' || chart.kind === 'doughnut'
      ? Math.max(1, Math.ceil(chart.categories.length / pieLegendColumns))
      : 1;
    const pieLegendExtraHeight = (pieLegendRows - 1) * 12;
    // 次坐标轴的刻度画在右边，要留出地方。
    const right = secondarySeries.length && chart.axes?.secondaryValue?.visible !== false ? 34 : 8;
    const plot = { x: left, y: titleHeight, width: Math.max(1, width - left - right), height: Math.max(1, height - titleHeight - bottom - pieLegendExtraHeight) };
    if (chart.title) add('text', { x: String(width / 2), y: '13', 'text-anchor': 'middle', 'font-size': '12', fill: '#222' }).textContent = chart.title;
    const isPercentBars = barSeries.length > 0 && chart.grouping === 'percentStacked';
    const isStackedBars = barSeries.length > 0 && (chart.grouping === 'stacked' || isPercentBars);
    // 每根数值轴各算各的刻度：次坐标轴的序列（常见是「金额 + 增长率」）量级完全不同，合到一根
    // 轴上会把其中一条压成一条直线。误差线的端点也算进范围，否则会画出绘图区。
    const scaleFor = (members: ChartSeriesInfo[]) => {
      const values = members.flatMap((series) => [
        ...series.values.filter((value): value is number => value !== null && Number.isFinite(value)),
        ...(series.errorBars ?? []).filter((bars) => bars.direction === 'y').flatMap((bars) =>
          errorBarRanges(series.values, bars).flatMap((range) => (range ? [range.low, range.high] : []))),
      ]);
      const bars = members.filter((series) => seriesType(series) === 'bar');
      const stackedTotals = isStackedBars && bars.length ? bars[0]!.values.map((_, index) =>
        bars.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0)) : [];
      const minValue = isPercentBars && bars.length ? -100 : stackedTotals.length ? Math.min(0, ...stackedTotals, ...values) : values.length ? Math.min(0, ...values) : 0;
      const maxValue = isPercentBars && bars.length ? 100 : stackedTotals.length ? Math.max(1, ...stackedTotals, ...values) : values.length ? Math.max(0, ...values) : 1;
      return axisTicks(minValue, chart.kind === 'bubble' ? maxValue + (maxValue - minValue) * 0.15 : maxValue, 5);
    };
    const scale = scaleFor(primarySeries.length ? primarySeries : chart.series);
    const secondaryScale = secondarySeries.length ? scaleFor(secondarySeries) : undefined;
    const scaleOf = (series: ChartSeriesInfo) => (series.axis === 'secondary' && secondaryScale ? secondaryScale : scale);
    if (chart.kind === 'pie' || chart.kind === 'doughnut') {
      const total = chart.series.reduce((sum, series) => sum + series.values.reduce<number>((part, value) => part + (value !== null && value > 0 ? value : 0), 0), 0);
      const radius = Math.max(1, Math.min(plot.width, plot.height) / 2 - 2);
      let angle = -Math.PI / 2;
      for (const [index, value] of (chart.series[0]?.values ?? []).entries()) {
        const amount = value !== null && value > 0 ? value : 0;
        const next = angle + (total ? amount / total : 0) * Math.PI * 2;
        const fill = chart.series[0]?.pointFills?.[index]?.color ?? chart.series[0]?.fill?.color ?? 'none';
        const path = add('path', { d: pieSlicePath(angle, next, plot.x + plot.width / 2, plot.y + plot.height / 2, radius, chart.kind === 'doughnut' ? radius * 0.5 : 0), fill, stroke: '#fff', 'stroke-width': '1' });
        path.setAttribute('data-docx-chart-series', '0');
        angle = next;
      }
      chart.categories.forEach((category, index) => {
        if (!category) return;
        const column = index % pieLegendColumns;
        const row = Math.floor(index / pieLegendColumns);
        const x = plot.x + column * 70;
        const y = height - 4 - (pieLegendRows - 1 - row) * 12;
        add('text', { x: String(x + 11), y: String(y), 'font-size': '9', fill: '#444' }).textContent = category;
        add('rect', { x: String(x), y: String(y - 8), width: '8', height: '8', fill: chart.series[0]?.pointFills?.[index]?.color ?? chart.series[0]?.fill?.color ?? 'none' });
      });
      return;
    }
    if (chart.kind === 'stock') {
      this.renderStockChart(add, chart, plot, scale);
    }
    const xyChart = chart.kind === 'scatter' || chart.kind === 'bubble';
    const allXValues = chart.series.flatMap((series) => series.xValues?.filter((value): value is number => value !== null && Number.isFinite(value)) ?? []);
    // 气泡有半径，数据点贴着轴画会被切掉一半：两侧各留 15% 的余量。
    const xPadding = chart.kind === 'bubble' && allXValues.length ? (Math.max(...allXValues) - Math.min(...allXValues) || 1) * 0.15 : 0;
    const xScale = xyChart && allXValues.length
      ? axisTicks(Math.min(...allXValues) - xPadding, Math.max(...allXValues) + xPadding, 5)
      : undefined;
    const categoryCount = Math.max(chart.categories.length, ...chart.series.map((series) => series.values.length), 0);
    if (!xyChart && chart.axes?.category?.visible !== false) {
      chart.categories.forEach((category, index) => {
        if (!category) return;
        // 柱子画在每个类目带的中间，标签也要在中间；只有折线 / 面积图才两端对齐。
        const x = horizontalBars
          ? plot.x - 4
          : plot.x + (barSeries.length || chart.kind === 'stock'
            ? ((index + 0.5) / Math.max(1, categoryCount)) * plot.width
            : (categoryCount > 1 ? index / (categoryCount - 1) : 0.5) * plot.width);
        const y = horizontalBars
          ? plot.y + (index + 0.5) / Math.max(1, categoryCount) * plot.height + 4
          : plot.y + plot.height + 14;
        const label = add('text', { x: String(x), y: String(y), 'text-anchor': horizontalBars ? 'end' : 'middle', 'font-size': '9', fill: '#555', 'data-docx-chart-category': '1' });
        label.textContent = category;
      });
    }
    if (chart.axes?.value?.majorGridlines) {
      for (const tick of scale.ticks) {
        if (horizontalBars) {
          const x = plot.x + valueToPx(tick, scale, plot.width);
          add('line', { x1: String(x), x2: String(x), y1: String(plot.y), y2: String(plot.y + plot.height), stroke: '#e5e7eb', 'data-docx-chart-gridline': '1' });
        } else {
          const y = plot.y + plot.height - valueToPx(tick, scale, plot.height);
          add('line', { x1: String(plot.x), x2: String(plot.x + plot.width), y1: String(y), y2: String(y), stroke: '#e5e7eb', 'data-docx-chart-gridline': '1' });
        }
      }
    }
    if (chart.axes?.value?.visible !== false) add('line', horizontalBars
      ? { x1: String(plot.x), x2: String(plot.x + plot.width), y1: String(plot.y + plot.height), y2: String(plot.y + plot.height), stroke: '#555', 'data-docx-chart-axis': 'value' }
      : { x1: String(plot.x), x2: String(plot.x), y1: String(plot.y), y2: String(plot.y + plot.height), stroke: '#555', 'data-docx-chart-axis': 'value' });
    if (chart.axes?.category?.visible !== false) add('line', horizontalBars
      ? { x1: String(plot.x), x2: String(plot.x), y1: String(plot.y), y2: String(plot.y + plot.height), stroke: '#555', 'data-docx-chart-axis': 'category' }
      : { x1: String(plot.x), x2: String(plot.x + plot.width), y1: String(plot.y + plot.height), y2: String(plot.y + plot.height), stroke: '#555', 'data-docx-chart-axis': 'category' });
    for (const tick of scale.ticks) {
      const x = horizontalBars
        ? plot.x + valueToPx(tick, scale, plot.width)
        : plot.x - 4;
      const y = horizontalBars
        ? plot.y + plot.height + 14
        : plot.y + plot.height - valueToPx(tick, scale, plot.height) + 4;
      const label = add('text', { x: String(x), y: String(y), 'text-anchor': horizontalBars ? 'middle' : 'end', 'font-size': '9', fill: '#555', 'data-docx-chart-tick': '1' });
      label.textContent = String(tick);
    }
    if (secondaryScale && chart.axes?.secondaryValue?.visible !== false && !horizontalBars) {
      const axisX = plot.x + plot.width;
      add('line', { x1: String(axisX), x2: String(axisX), y1: String(plot.y), y2: String(plot.y + plot.height), stroke: '#555', 'data-docx-chart-axis': 'secondary' });
      for (const tick of secondaryScale.ticks) {
        const label = add('text', { x: String(axisX + 4), y: String(plot.y + plot.height - valueToPx(tick, secondaryScale, plot.height) + 4), 'text-anchor': 'start', 'font-size': '9', fill: '#555', 'data-docx-chart-tick': 'secondary' });
        label.textContent = String(tick);
      }
    }
    // 有柱子时折线点落在每个类目带的中间（Excel 的 crossBetween=between），否则沿用两端对齐。
    const categoryX = (index: number, count: number) => plot.x + (barSeries.length
      ? ((index + 0.5) / Math.max(1, count)) * plot.width
      : (count > 1 ? index / (count - 1) : 0.5) * plot.width);
    const seriesXOf = (series: ChartSeriesInfo, index: number): number => {
      const xValue = series.xValues?.[index];
      if (xyChart && xScale && xValue !== null && xValue !== undefined) return plot.x + valueToPx(xValue, xScale, plot.width);
      return categoryX(index, chart.series[0]?.values.length ?? 0);
    };
    const yOf = (series: ChartSeriesInfo, value: number) => plot.y + plot.height - valueToPx(value, scaleOf(series), plot.height);
    if (barSeries.length) {
      // 主、次坐标轴上的柱子各按自己的刻度算。
      for (const members of [barSeries.filter((series) => series.axis !== 'secondary'), barSeries.filter((series) => series.axis === 'secondary')]) {
        if (!members.length) continue;
        const rects = barRects(members.map((series) => series.values.map((value) => value ?? 0)), scaleOf(members[0]!), plot.width, plot.height, { grouping: chart.grouping ?? 'clustered', direction: chart.barDirection });
        rects.forEach((rectsOfSeries, memberIndex) => {
          const series = members[memberIndex]!;
          const seriesIndex = chart.series.indexOf(series);
          rectsOfSeries.forEach((rect) => add('rect', { x: String(plot.x + rect.x), y: String(plot.y + rect.y), width: String(rect.width), height: String(rect.height), fill: series.fill?.color ?? 'none', 'data-docx-chart-series': String(seriesIndex) }));
        });
      }
    }
    if (chart.kind === 'bubble') {
      const sizes = chart.series.flatMap((series) => series.bubbleSizes?.filter((size): size is number => size !== null && size > 0) ?? []);
      const largest = Math.max(1e-9, ...sizes);
      // 气泡按面积缩放：半径正比于 √size，最大的气泡直径约为绘图区短边的 1/4（Excel 默认 bubbleScale 100）。
      const maxRadius = Math.min(plot.width, plot.height) / 8;
      chart.series.forEach((series, seriesIndex) => series.values.forEach((value, index) => {
        const size = series.bubbleSizes?.[index];
        if (value === null || !Number.isFinite(value) || size === null || size === undefined || size <= 0) return;
        add('circle', {
          cx: String(seriesXOf(series, index)), cy: String(yOf(series, value)), r: String(Math.sqrt(size / largest) * maxRadius),
          fill: series.fill?.color ?? 'none', 'fill-opacity': '0.75', stroke: series.line?.color ?? series.fill?.color ?? 'none',
          'data-docx-chart-series': String(seriesIndex),
        });
      }));
    } else if (chart.kind !== 'stock') {
      chart.series.forEach((series, seriesIndex) => {
        const type = seriesType(series);
        if (type === 'bar') return;
        const segments: Array<Array<{ x: number; y: number }>> = [];
        let current: Array<{ x: number; y: number }> = [];
        const zeroY = plot.y + plot.height - valueToPx(0, scaleOf(series), plot.height);
        series.values.forEach((value, index) => {
          if (value === null || !Number.isFinite(value)) {
            if (current.length) segments.push(current);
            current = [];
            return;
          }
          current.push({ x: seriesXOf(series, index), y: yOf(series, value) });
        });
        if (current.length) segments.push(current);
        const path = segments.map((points) => {
          const line = points.map((point, pointIndex) => `${pointIndex ? 'L' : 'M'} ${point.x} ${point.y} `).join('');
          if (type !== 'area') return line;
          const last = points[points.length - 1]!;
          const first = points[0]!;
          return last.x === first.x
            ? `${line}L ${first.x} ${zeroY} Z`
            : `${line}L ${last.x} ${zeroY} L ${first.x} ${zeroY} Z`;
        }).join(' ');
        if (path) add('path', { d: path, fill: type === 'area' ? (series.fill?.color ?? 'none') : 'none', 'fill-opacity': type === 'area' ? '0.35' : '1', stroke: series.line?.color ?? series.fill?.color ?? 'none', 'data-docx-chart-series': String(seriesIndex) });
      });
    }
    // 趋势线：类目图的 x 按 Excel 的约定取 1..n，散点 / 气泡图取缓存的 x。
    chart.series.forEach((series, seriesIndex) => {
      const color = series.line?.color ?? series.fill?.color ?? '#444';
      const count = series.values.length;
      for (const trendline of series.trendlines ?? []) {
        const points = series.values.flatMap((value, index) => {
          if (value === null || !Number.isFinite(value)) return [];
          const x = xyChart ? series.xValues?.[index] : index + 1;
          return x === null || x === undefined || !Number.isFinite(x) ? [] : [{ x, y: value }];
        });
        const fitted = trendlinePoints(points, trendline.type, { order: trendline.order, period: trendline.period });
        const toPx = (x: number) => (xyChart && xScale ? plot.x + valueToPx(x, xScale, plot.width) : categoryX(x - 1, count));
        const d = fitted.map((point, index) => `${index ? 'L' : 'M'} ${toPx(point.x)} ${yOf(series, point.y)}`).join(' ');
        if (d) add('path', { d, fill: 'none', stroke: color, 'stroke-dasharray': '4 3', 'data-docx-chart-trendline': String(seriesIndex) });
      }
      // 误差线：竖向用数值轴刻度，横向（只有散点 / 气泡图有）用 x 刻度；两端各画一小段横 / 竖线当帽子。
      for (const bars of series.errorBars ?? []) {
        const source = bars.direction === 'x' ? series.xValues ?? [] : series.values;
        errorBarRanges(source, bars).forEach((range, index) => {
          const value = series.values[index];
          if (!range || value === null || value === undefined || !Number.isFinite(value)) return;
          const x = seriesXOf(series, index);
          const y = yOf(series, value);
          if (bars.direction === 'x') {
            if (!xScale) return;
            const x1 = plot.x + valueToPx(range.low, xScale, plot.width);
            const x2 = plot.x + valueToPx(range.high, xScale, plot.width);
            add('path', { d: `M ${x1} ${y} L ${x2} ${y} M ${x1} ${y - 3} L ${x1} ${y + 3} M ${x2} ${y - 3} L ${x2} ${y + 3}`, stroke: color, fill: 'none', 'data-docx-chart-error-bar': String(seriesIndex) });
          } else {
            const y1 = yOf(series, range.low);
            const y2 = yOf(series, range.high);
            add('path', { d: `M ${x} ${y1} L ${x} ${y2} M ${x - 3} ${y1} L ${x + 3} ${y1} M ${x - 3} ${y2} L ${x + 3} ${y2}`, stroke: color, fill: 'none', 'data-docx-chart-error-bar': String(seriesIndex) });
          }
        });
      }
    });
    chart.series.forEach((series, index) => {
      if (!series.name) return;
      const x = plot.x + index * 70;
      add('rect', { x: String(x), y: String(height - 12), width: '8', height: '8', fill: series.fill?.color ?? 'none' });
      add('text', { x: String(x + 11), y: String(height - 4), 'font-size': '9', fill: '#444' }).textContent = series.name;
    });
  }

  /** 雷达图：每个类目一根辐条，序列连成闭合多边形；`filled` 样式填色，其余只描边。 */
  private renderRadarChart(add: (name: string, attrs: Record<string, string>) => Element, chart: NonNullable<ShapeInfo['chart']>,
    width: number, height: number): void {
    const titleHeight = chart.title ? 18 : 4;
    if (chart.title) add('text', { x: String(width / 2), y: '13', 'text-anchor': 'middle', 'font-size': '12', fill: '#222' }).textContent = chart.title;
    const count = Math.max(chart.categories.length, ...chart.series.map((series) => series.values.length), 0);
    if (!count) return;
    const cx = width / 2;
    const cy = titleHeight + (height - titleHeight - 14) / 2;
    const radius = Math.max(1, Math.min(width / 2 - 30, (height - titleHeight - 14) / 2 - 10));
    const values = chart.series.flatMap((series) => series.values.filter((value): value is number => value !== null && Number.isFinite(value)));
    const scale = axisTicks(Math.min(0, ...values), Math.max(1, ...values), 4);
    for (const tick of scale.ticks) {
      if (tick <= scale.min) continue;
      const r = valueToPx(tick, scale, radius);
      const ring = Array.from({ length: count }, (_, index) => radarPoint(index, count, r, cx, cy));
      add('path', { d: `${ring.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')} Z`, fill: 'none', stroke: '#e5e7eb', 'data-docx-chart-gridline': '1' });
    }
    for (let index = 0; index < count; index++) {
      const end = radarPoint(index, count, radius, cx, cy);
      add('line', { x1: String(cx), y1: String(cy), x2: String(end.x), y2: String(end.y), stroke: '#cfd5dd', 'data-docx-chart-axis': 'category' });
      const label = radarPoint(index, count, radius + 10, cx, cy);
      const text = chart.categories[index];
      if (text) add('text', { x: String(label.x), y: String(label.y + 3), 'text-anchor': 'middle', 'font-size': '9', fill: '#555', 'data-docx-chart-category': '1' }).textContent = text;
    }
    chart.series.forEach((series, seriesIndex) => {
      const points = series.values.map((value, index) => radarPoint(index, count, valueToPx(value ?? scale.min, scale, radius), cx, cy));
      const color = series.line?.color ?? series.fill?.color ?? '#444';
      add('path', {
        d: `${points.map((point, index) => `${index ? 'L' : 'M'} ${point.x} ${point.y}`).join(' ')} Z`,
        fill: chart.radarStyle === 'filled' ? (series.fill?.color ?? color) : 'none',
        'fill-opacity': chart.radarStyle === 'filled' ? '0.4' : '1',
        stroke: color, 'data-docx-chart-series': String(seriesIndex),
      });
      if (chart.radarStyle === 'marker') {
        for (const point of points) add('circle', { cx: String(point.x), cy: String(point.y), r: '2.5', fill: color, 'data-docx-chart-marker': String(seriesIndex) });
      }
    });
  }

  /**
   * 股价图。序列按 Excel 的约定排：三条是「最高—最低—收盘」，四条是「开盘—最高—最低—收盘」。
   * 每个类目一根最高—最低竖线、收盘价一个向右的小横；有开盘价且开了涨跌柱时，开盘到收盘画一根柱，
   * 涨（收 ≥ 开）白、跌黑。序列自己的连线在股价图里默认不画（Word 写的是 noFill）。
   */
  private renderStockChart(add: (name: string, attrs: Record<string, string>) => Element, chart: NonNullable<ShapeInfo['chart']>,
    plot: { x: number; y: number; width: number; height: number }, scale: { min: number; max: number }): void {
    const series = chart.series;
    if (series.length < 3) return;
    const [open, high, low, close] = series.length >= 4 ? [series[0], series[1], series[2], series[3]] : [undefined, series[0], series[1], series[2]];
    const count = Math.max(...series.map((entry) => entry.values.length));
    const band = plot.width / Math.max(1, count);
    const y = (value: number) => plot.y + plot.height - valueToPx(value, scale, plot.height);
    for (let index = 0; index < count; index++) {
      const x = plot.x + (index + 0.5) * band;
      const hi = high?.values[index];
      const lo = low?.values[index];
      const cl = close?.values[index];
      const op = open?.values[index];
      if (hi !== null && hi !== undefined && lo !== null && lo !== undefined && chart.stock?.hiLowLines !== false) {
        add('line', { x1: String(x), x2: String(x), y1: String(y(hi)), y2: String(y(lo)), stroke: '#333', 'data-docx-chart-stock': 'hilow' });
      }
      if (op !== null && op !== undefined && cl !== null && cl !== undefined && chart.stock?.upDownBars) {
        const top = y(Math.max(op, cl));
        add('rect', { x: String(x - band * 0.25), y: String(top), width: String(band * 0.5), height: String(Math.max(1, y(Math.min(op, cl)) - top)),
          fill: cl >= op ? '#ffffff' : '#333333', stroke: '#333', 'data-docx-chart-stock': cl >= op ? 'up' : 'down' });
      } else if (cl !== null && cl !== undefined) {
        add('line', { x1: String(x), x2: String(x + band * 0.2), y1: String(y(cl)), y2: String(y(cl)), stroke: '#333', 'data-docx-chart-stock': 'close' });
      }
    }
  }

  private reviewScopedRun(paragraphIndex: number, run: RunInfo, reviewContext: ReviewRenderContext): RunInfo {
    if (!run.revisions?.length) return run;
    let revisions = run.revisions;
    if (!this.reviewFilter.showRevisions) revisions = [];
    if (reviewContext.authors) {
      revisions = revisions.filter((revision) => reviewContext.authors!.has(reviewerBucketKey(reviewerBucketOf(revision.author))));
    }
    const hasInsertion = revisions.some((revision) => revision.kind === 'insertion');
    const hasDeletion = revisions.some((revision) => revision.kind === 'deletion');
    const hasMoveFrom = revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'from');
    const hasMoveTo = revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'to');
    const view = this.reviewFilter.revisionView;
    let text = run.text;
    if (view === 'original') {
      if (hasInsertion || hasMoveTo) text = '';
      else if (hasDeletion || hasMoveFrom) text = reviewContext.deletedTextByRun.get(`${paragraphIndex}:${run.index}`) ?? run.text;
    } else if (view === 'final') {
      if (hasDeletion || hasMoveFrom) text = '';
    }
    if (text === run.text && revisions.length === run.revisions.length && revisions.every((entry, index) => entry === run.revisions![index])) {
      return run;
    }
    return { ...run, text, revisions: revisions.length ? revisions : undefined };
  }

  private reviewColor(author: string | undefined, reviewContext: ReviewRenderContext): string {
    const key = reviewerBucketKey(reviewerBucketOf(author));
    let color = reviewContext.revisionColors.get(key);
    if (!color) {
      color = REVISION_COLOR_PALETTE[revisionColorIndexForKey(key)]!;
      reviewContext.revisionColors.set(key, color);
    }
    return color;
  }

  private revisionAriaDescription(revisions: RunInfo['revisions']): string {
    if (!revisions?.length) return '';
    const labels = [];
    if (revisions.some((revision) => revision.kind === 'insertion')) labels.push('插入');
    if (revisions.some((revision) => revision.kind === 'deletion')) labels.push('删除');
    if (revisions.some((revision) => revision.kind === 'move')) labels.push('移动');
    const author = revisions.find((revision) => revision.author !== undefined)?.author;
    return author ? `修订：${labels.join(' / ') || '变更'}，作者 ${author}` : `修订：${labels.join(' / ') || '变更'}`;
  }

  private registerRevisionNode(ids: number[], node: HTMLElement): void {
    if (this.measuring) return;
    if (!ids.length) return;
    node.dataset.docxRevisionIds = ids.join(',');
    node.tabIndex = -1;
    for (const id of ids) {
      const key = String(id);
      const list = this.revisionRunIds.get(key) ?? [];
      list.push(node);
      this.revisionRunIds.set(key, list);
    }
  }

  private appendDeletedRunVisualization(
    paragraphElement: HTMLElement,
    paragraphIndex: number,
    run: RunInfo,
    reviewContext: ReviewRenderContext,
  ): void {
    if (this.reviewFilter.revisionView !== 'markup' || !run.revisions?.some((revision) =>
      revision.kind === 'deletion' || (revision.kind === 'move' && revision.move?.side === 'from'))) return;
    const deletedText = reviewContext.deletedTextByRun.get(`${paragraphIndex}:${run.index}`);
    if (!deletedText) return;
    const marker = this.root.ownerDocument.createElement('span');
    marker.className = 'docx-deleted-text';
    marker.dataset.docxDeleted = '1';
    marker.contentEditable = 'false';
    marker.textContent = deletedText;
    const hasMoveFrom = run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'from');
    marker.style.textDecoration = hasMoveFrom ? 'line-through underline' : 'line-through';
    if (hasMoveFrom) marker.style.textDecorationStyle = 'solid double';
    marker.style.opacity = '0.85';
    const authorRevision = run.revisions.find((revision) =>
      revision.kind === 'deletion' || (revision.kind === 'move' && revision.move?.side === 'from'));
    marker.style.color = this.reviewColor(authorRevision?.author, reviewContext);
    marker.setAttribute('aria-label', hasMoveFrom ? `修订移动来源文本：${deletedText}` : `修订删除文本：${deletedText}`);
    if (hasMoveFrom) marker.append(this.makeMark('↤', '移动来源'));
    this.registerRevisionNode(run.revisions.map((revision) => revision.id), marker);
    paragraphElement.append(marker);
  }

  private runIsHidden(run: RunInfo): boolean {
    const format = run.effective ?? run;
    return format.hidden === true || format.webHidden === true;
  }

  private appendHiddenRunVisualization(parent: HTMLElement, run: RunInfo): void {
    const marker = this.root.ownerDocument.createElement('span');
    marker.className = 'docx-hidden-text';
    marker.dataset.docxHidden = '1';
    marker.dataset.docxHiddenPreserved = '1';
    marker.contentEditable = 'false';
    marker.textContent = run.text;
    marker.style.display = 'none';
    marker.setAttribute('aria-hidden', 'true');
    parent.append(marker);
  }

  private makeMark(text: string, label: string): HTMLElement {
    const mark = this.root.ownerDocument.createElement('span');
    mark.className = 'docx-mark';
    mark.textContent = text;
    mark.contentEditable = 'false';
    mark.setAttribute('aria-hidden', 'true');
    mark.setAttribute('data-docx-mark', '1');
    mark.title = label;
    mark.style.userSelect = 'none';
    mark.style.pointerEvents = 'none';
    mark.style.opacity = '0.6';
    return mark;
  }

  private leader(value: TabStop['leader'] | undefined): string {
    switch (value) {
      case 'dot': return '.';
      case 'hyphen': return '-';
      case 'underscore': return '_';
      case 'heavy': return '━';
      case 'middleDot': return '·';
      default: return '';
    }
  }

  private measure(text: string, run: RunInfo): number {
    if (!this.metrics || !text) return 0;
    const effective = run.effective ?? run;
    const style = effective.italic ? 'italic' : 'normal';
    const weight = effective.bold ? '700' : '400';
    const size = `${effective.fontSize ?? 11}pt`;
    const family = [effective.fontFamily, effective.fontFamilyEastAsia, 'Arial', 'sans-serif'].filter(Boolean).join(', ');
    this.metrics.font = `${style} ${weight} ${size} ${family}`;
    return this.metrics.measureText(text).width;
  }

  private nextTabStop(tabs: TabStop[], currentPx: number): TabStop | undefined {
    const currentTwips = currentPx * 1440 / 96;
    return [...tabs]
      .filter((tab) => Number.isFinite(tab.position))
      .sort((a, b) => a.position - b.position)
      .find((tab) => tab.position > currentTwips);
  }

  private makeTabSpan(
    paragraph: ParagraphInfo,
    run: RunInfo,
    currentPx: number,
    following: string,
    defaultTabStopTwips: number,
  ): HTMLSpanElement {
    const tab = this.root.ownerDocument.createElement('span');
    tab.className = 'docx-tab';
    tab.textContent = '\t';
    tab.style.display = 'inline-block';
    const stop = this.nextTabStop(paragraph.effective?.tabs ?? paragraph.tabs ?? [], currentPx);
    const defaultTab = defaultTabStopTwips * 96 / 1440;
    const target = stop ? Math.max(0, stop.position) * 96 / 1440 : (Math.floor(currentPx / defaultTab) + 1) * defaultTab;
    const nextWidth = this.measure(following, run);
    const decimalMatch = /[.,，．]/.exec(following);
    const decimalLeft = decimalMatch ? this.measure(following.slice(0, decimalMatch.index), run) : nextWidth;
    const alignment = stop?.alignment ?? 'left';
    const rawWidth = alignment === 'center' ? target - currentPx - nextWidth / 2
      : alignment === 'right' ? target - currentPx - nextWidth
        : alignment === 'decimal' ? target - currentPx - decimalLeft
          : target - currentPx;
    const width = Math.max(0, rawWidth);
    tab.style.width = `${width}px`;
    if (alignment === 'bar') tab.style.borderLeft = '1px solid currentColor';
    const leader = this.leader(stop?.leader);
    if (leader) {
      const visual = this.makeMark(leader.repeat(Math.max(1, Math.floor(Math.max(width, 8) / Math.max(1, this.measure(leader, run))))), '制表位前导符');
      visual.style.position = 'absolute';
      visual.style.inset = '0';
      visual.style.whiteSpace = 'nowrap';
      visual.style.overflow = 'hidden';
      tab.style.position = 'relative';
      tab.style.width = `${Math.max(width, 8)}px`;
      tab.append(visual);
    }
    return tab;
  }

  private appendRun(
    paragraphElement: HTMLElement,
    paragraph: ParagraphInfo,
    run: RunInfo,
    reviewContext: ReviewRenderContext,
    defaultTabStopTwips: number,
    currentLineOffsetPx: number,
  ): number {
    const hidden = this.runIsHidden(run);
    if (hidden && !this.options.showHiddenText) {
      this.appendHiddenRunVisualization(paragraphElement, run);
      return currentLineOffsetPx;
    }
    const unsafe = run.hyperlink?.unsafe ?? false;
    const hasSafeLink = !!(run.hyperlink && !unsafe && (run.hyperlink.url || run.hyperlink.anchor));
    const runSpan = this.root.ownerDocument.createElement(hasSafeLink ? 'a' : 'span');
    runSpan.dataset.docxRun = String(run.index);
    if (hidden) {
      runSpan.dataset.docxHidden = '1';
      runSpan.classList.add('docx-hidden-text');
    }
    if (run.field) {
      runSpan.dataset.docxField = String(run.field.index);
      runSpan.dataset.docxFieldRole = run.field.role;
      runSpan.contentEditable = 'false';
      if (run.field.role === 'result') {
        runSpan.dataset.docxContent = '1';
        if (this.options.showFieldShading !== false) runSpan.classList.add('docx-field-shading');
      }
    }
    const fieldInfo = run.field ? this.renderFieldInfos?.get(run.field.index) : undefined;
    // EQ 域除顶层 \o（下面按层叠画）以外的开关：按指令结构画成 MathML。画出来的公式是**装饰**
    // ——指令才是内容，所以 contentEditable=false 加 data-docx-mark，readText() 跳过它。Word
    // 缓存的域结果（常常为空，有时是压平的文字）照样留在 DOM 里让 readText() 读到，只是不显示：
    // 拿掉它的话 flush() 会以为用户删了这段文字，把它从文档里删掉。
    const equation = fieldInfo?.kind === 'EQ' && fieldInfo.equation && fieldInfo.equation.switch !== 'o'
      ? fieldInfo.equation : undefined;
    const equationAnchor = equation && fieldInfo
      ? fieldInfo.resultRuns[0] ?? Math.max(...fieldInfo.runs)
      : undefined;
    if (equation && run.index === equationAnchor) paragraphElement.append(this.makeEquation(equation));
    const hideEquationResult = equation !== undefined && run.field?.role === 'result';
    const startOffsetPx = currentLineOffsetPx;
    if (hideEquationResult) {
      runSpan.dataset.docxEquationResult = '1';
      runSpan.style.display = 'none';
    }
    if (run.revisions?.length) runSpan.dataset.docxRevisionIds = run.revisions.map((revision) => revision.id).join(',');
    const commentIds = [...new Set([...(this.commentParagraphIds.get(paragraph.index) ?? []), ...(this.commentRunIds.get(`${paragraph.index}:${run.index}`) ?? [])])];
    if (commentIds.length) {
      runSpan.classList.add('docx-comment-anchor');
      runSpan.dataset.docxCommentIds = commentIds.join(',');
    }
    if (hasSafeLink && run.hyperlink) {
      const link = runSpan as HTMLAnchorElement;
      runSpan.dataset.docxLink = '1';
      runSpan.dataset.docxUnsafe = 'false';
      if (run.hyperlink.url) {
        link.href = run.hyperlink.url;
        runSpan.dataset.docxUrl = run.hyperlink.url;
      } else if (run.hyperlink.anchor) {
        link.href = `#${run.hyperlink.anchor}`;
      }
      if (run.hyperlink.anchor) runSpan.dataset.docxAnchor = run.hyperlink.anchor;
      if (run.hyperlink.tooltip) runSpan.title = run.hyperlink.tooltip;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      runSpan.style.color = '#0563C1';
      runSpan.style.textDecoration = 'underline';
    } else if (run.hyperlink) {
      runSpan.dataset.docxLink = '1';
      runSpan.dataset.docxUnsafe = String(unsafe);
      if (run.hyperlink.url) runSpan.dataset.docxUrl = run.hyperlink.url;
      if (run.hyperlink.anchor) runSpan.dataset.docxAnchor = run.hyperlink.anchor;
      if (run.hyperlink.tooltip) runSpan.title = run.hyperlink.tooltip;
      if (unsafe) runSpan.style.textDecoration = 'underline wavy red';
    }
    applyRunStyle(runSpan, run, (paragraph.effective ?? paragraph).textAlignment ?? undefined);
    if (run.revisions?.length && this.reviewFilter.showRevisions && this.reviewFilter.revisionView === 'markup') {
      const hasInsertion = run.revisions.some((revision) => revision.kind === 'insertion');
      const hasDeletion = run.revisions.some((revision) => revision.kind === 'deletion');
      const hasMoveFrom = run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'from');
      const hasMoveTo = run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'to');
      const author = run.revisions.find((revision) => revision.author !== undefined)?.author;
      const revisionColor = this.reviewColor(author, reviewContext);
      const textDecoration = [hasInsertion || hasMoveTo ? 'underline' : '', hasDeletion || hasMoveFrom ? 'line-through' : ''].filter(Boolean).join(' ');
      if (textDecoration) runSpan.style.textDecoration = textDecoration;
      if (hasMoveTo || hasMoveFrom) runSpan.style.textDecorationStyle = 'double';
      if (textDecoration) runSpan.style.textDecorationColor = revisionColor;
      const description = this.revisionAriaDescription(run.revisions);
      if (description) runSpan.setAttribute('aria-description', description);
      this.registerRevisionNode(run.revisions.map((revision) => revision.id), runSpan);
    }
    if (hidden) {
      if (!runSpan.style.textDecoration?.includes('underline')) {
        runSpan.style.textDecoration = [runSpan.style.textDecoration, 'underline'].filter(Boolean).join(' ');
      }
      runSpan.style.textDecorationStyle = 'dashed';
    }
    // 文本平时直接进 runSpan;注音要把基字符放进 <ruby> 里，双行合一要放进上下两行的盒子里。
    let textTarget: HTMLElement = runSpan;
    let rubyAnnotation: HTMLElement | undefined;
    if (run.ruby) {
      const ruby = this.root.ownerDocument.createElement('ruby');
      ruby.dataset.docxRuby = '1';
      runSpan.append(ruby);
      textTarget = ruby;
      // 注音是文档内容，但不在段落的阅读顺序里。标成 contentEditable=false，readText() 就会
      // 跳过它——否则注音会被当成正文追加进 paragraph.text，把文本写坏。
      const annotation = this.root.ownerDocument.createElement('rt');
      annotation.contentEditable = 'false';
      annotation.dataset.docxRubyText = '1';
      annotation.style.userSelect = 'none';
      if (run.ruby.sizeHalfPoints !== undefined) annotation.style.fontSize = `${run.ruby.sizeHalfPoints / 2}pt`;
      const rubyAlign = rubyAlignToCss(run.ruby.align);
      if (rubyAlign) annotation.style.rubyAlign = rubyAlign;
      if (run.ruby.language) annotation.lang = run.ruby.language;
      annotation.textContent = run.ruby.text;
      rubyAnnotation = annotation;
    }
    // EQ \o 的叠印只在「整个域结果都在这一个 run 里」时接管渲染；结果跨多个 run 时按原样
    // 画，宁可不叠印也不去切分结果文本。
    const equationLayers = run.field?.role === 'result' && fieldInfo?.kind === 'EQ'
      && fieldInfo.resultRuns[0] === run.index && run.text === fieldInfo.result
      ? overstrikeLayers(fieldInfo.equation, fieldInfo.result)
      : undefined;
    const combine = run.effective?.eastAsianLayout ?? run.eastAsianLayout;
    if (combine?.combine) {
      // CSS 画不出双行合一，所以自己堆：两个 display:block 的 span（不能用 <br>，readText()
      // 会把它读成换行）。括号是装饰，contentEditable=false 且不进 readText()。
      const brackets = combineBracketChars(combine.combineBrackets);
      const box = this.root.ownerDocument.createElement('span');
      box.dataset.docxCombine = '1';
      box.style.display = 'inline-flex';
      box.style.flexDirection = 'column';
      box.style.verticalAlign = 'middle';
      box.style.fontSize = '50%';
      box.style.lineHeight = '1';
      box.style.textAlign = 'center';
      const lines = combinedTextLines(run.text);
      if (brackets) textTarget.append(this.makeMark(brackets[0], '双行合一左括号'));
      for (const line of lines) {
        const row = this.root.ownerDocument.createElement('span');
        row.style.display = 'block';
        row.append(this.root.ownerDocument.createTextNode(line));
        box.append(row);
      }
      textTarget.append(box);
      if (brackets) textTarget.append(this.makeMark(brackets[1], '双行合一右括号'));
      // 盒子是 font-size:50% 的上下两行，占的宽度是较长那行的一半，不是整段文字的宽度；
      // 按整段算会高估近一倍，让这种 run 过早换行。括号按原字号另计。
      currentLineOffsetPx += Math.max(...lines.map((line) => this.measure(line, run))) / 2;
      if (brackets) currentLineOffsetPx += brackets.reduce((total, bracket) => total + this.measure(bracket, run), 0);
    } else if (equationLayers) {
      // EQ 的 \o 是叠印：各层占同一个位置，各自按 \s\up / \s\do 上下错开。用 inline-grid
      // 把每层都放进 1/1 这个格子——容器宽度自然取最宽那层；换成绝对定位宽度会塌成 0。
      const box = this.root.ownerDocument.createElement('span');
      box.dataset.docxEquation = 'o';
      box.style.display = 'inline-grid';
      // \o\ac 是居中对齐。
      box.style.justifyItems = fieldInfo?.equation?.options?.includes('ac') ? 'center' : 'start';
      for (const layer of equationLayers) {
        const cell = this.root.ownerDocument.createElement('span');
        cell.style.gridArea = '1 / 1';
        if (layer.raisePoints) cell.style.transform = `translateY(${-layer.raisePoints}pt)`;
        if (layer.content) {
          cell.append(this.root.ownerDocument.createTextNode(layer.text));
        } else {
          // 这一层只在指令里，域结果里没有（「带圈字符」的那个圈）。必须当装饰：
          // readText() 要跳过它，否则编辑正文会把它写回文档。
          cell.contentEditable = 'false';
          cell.dataset.docxEquationGlyph = '1';
          cell.setAttribute('aria-hidden', 'true');
          cell.style.userSelect = 'none';
          cell.style.pointerEvents = 'none';
          cell.textContent = layer.text;
        }
        box.append(cell);
      }
      textTarget.append(box);
      currentLineOffsetPx += Math.max(...equationLayers.map((layer) => this.measure(layer.text, run)));
    } else {
    const segments = (run.field?.role === 'instruction' ? '' : run.text).split(/(\t|\n)/);
    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i]!;
      if (!segment) continue;
      if (segment === '\n') {
        textTarget.append(this.root.ownerDocument.createElement('br'));
        if (this.options.showFormattingMarks) textTarget.append(this.makeMark('↵', '换行符'));
        currentLineOffsetPx = 0;
        continue;
      }
      if (segment === '\t') {
        const nextText = segments.slice(i + 1).find((part) => part !== '\t' && part !== '\n') ?? '';
        const tab = this.makeTabSpan(paragraph, run, currentLineOffsetPx, nextText, defaultTabStopTwips);
        textTarget.append(tab);
        currentLineOffsetPx += Number.parseFloat(tab.style.width || '0');
        if (this.options.showFormattingMarks) textTarget.append(this.makeMark('→', '制表符'));
        continue;
      }
      if (this.options.showFormattingMarks && segment.includes(' ')) {
        const parts = segment.split(/( )/);
        for (const part of parts) {
          if (!part) continue;
          if (part === ' ') {
            textTarget.append(this.root.ownerDocument.createTextNode(' '));
            textTarget.append(this.makeMark('·', '空格'));
          } else textTarget.append(this.root.ownerDocument.createTextNode(part));
        }
      } else {
        textTarget.append(this.root.ownerDocument.createTextNode(segment));
      }
      currentLineOffsetPx += this.measure(segment, run);
    }
    }
    // 注音挂在基字符后面，<ruby> 的子元素顺序就是基字符在前、<rt> 在后。
    if (rubyAnnotation) textTarget.append(rubyAnnotation);
    if (run.revisions?.length && this.reviewFilter.showRevisions && this.reviewFilter.revisionView === 'markup' &&
        run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'to')) {
      runSpan.append(this.makeMark('↦', '移动目标'));
    }
    if (run.field?.role === 'result' && fieldInfo?.result === '' &&
        fieldInfo.resultRuns[0] === run.index && fieldInfo.formField) {
      const formField = fieldInfo.formField;
      if (formField.kind === 'checkBox') {
        const checkbox = this.root.ownerDocument.createElement('span');
        checkbox.className = 'docx-form-checkbox';
        checkbox.contentEditable = 'false';
        checkbox.setAttribute('data-docx-mark', '1');
        checkbox.setAttribute('aria-hidden', 'true');
        checkbox.style.userSelect = 'none';
        checkbox.style.pointerEvents = 'none';
        if (formField.checkBox?.sizePt) checkbox.style.fontSize = `${formField.checkBox.sizePt}pt`;
        checkbox.textContent = formField.checkBox?.checked ? '☑' : '☐';
        runSpan.append(checkbox);
      } else {
        const fallback = formField.kind === 'text'
          ? formField.text?.default
          : formField.dropDown?.entries[formField.dropDown.default ?? -1];
        if (fallback) {
          const placeholder = this.root.ownerDocument.createElement('span');
          placeholder.className = 'docx-form-field-default';
          placeholder.contentEditable = 'false';
          placeholder.setAttribute('data-docx-mark', '1');
          placeholder.textContent = fallback;
          runSpan.append(placeholder);
        }
      }
    }
    paragraphElement.append(runSpan);
    return hideEquationResult ? startOffsetPx : currentLineOffsetPx;
  }

  /**
   * SmartArt 子形状的文字：每一行一个 `<tspan>` 行，行内每个 run 一个 `<tspan>` 带格式。整块在框里
   * 竖直居中（SmartArt 的默认锚点），水平位置按段落对齐。SVG 的 text 不会自动换行，Word 预渲染时
   * 已经按框宽把字号缩好，这里不再折行。
   */
  private makeShapeChildText(paragraphs: ShapeTextParagraph[], x: number, y: number, width: number, height: number): Element {
    const svgNs = 'http://www.w3.org/2000/svg';
    const text = this.root.ownerDocument.createElementNS(svgNs, 'text');
    text.setAttribute('data-docx-shape-text', '1');
    const lines = paragraphs.flatMap((paragraph) => paragraph.lines.map((line) => ({ line, alignment: paragraph.alignment ?? 'center' })));
    const lineHeight = (line: typeof lines[number]['line']) => Math.max(12, ...line.map((run) => (run.fontSize ?? 10.5) * (4 / 3))) * 1.2;
    const total = lines.reduce((sum, entry) => sum + lineHeight(entry.line), 0);
    let cursor = y + (height - total) / 2;
    for (const { line, alignment } of lines) {
      const advance = lineHeight(line);
      cursor += advance;
      const row = this.root.ownerDocument.createElementNS(svgNs, 'tspan');
      const anchor = alignment === 'left' ? 'start' : alignment === 'right' ? 'end' : 'middle';
      row.setAttribute('x', String(alignment === 'left' ? x + 4 : alignment === 'right' ? x + width - 4 : x + width / 2));
      // 基线在行底往上 1/5 行高处，近似字体的下沉部分。
      row.setAttribute('y', String(cursor - advance * 0.2));
      row.setAttribute('text-anchor', anchor);
      for (const run of line) {
        const span = this.root.ownerDocument.createElementNS(svgNs, 'tspan');
        if (run.bold) span.setAttribute('font-weight', 'bold');
        if (run.italic) span.setAttribute('font-style', 'italic');
        const decorations = [run.underline ? 'underline' : '', run.strike ? 'line-through' : ''].filter(Boolean).join(' ');
        if (decorations) span.setAttribute('text-decoration', decorations);
        if (run.fontSize) span.setAttribute('font-size', `${run.fontSize}pt`);
        if (run.fontFamily) span.setAttribute('font-family', run.fontFamily);
        if (run.color) span.setAttribute('fill', run.color);
        span.textContent = run.text;
        row.appendChild(span);
      }
      text.appendChild(row);
    }
    return text;
  }

  /**
   * 外阴影画成 SVG 的 `feDropShadow`。SVG 的 stdDeviation 大约是模糊半径的一半；滤镜区域放大，
   * 免得阴影被形状的包围盒裁掉。返回滤镜 id。
   */
  private appendShadowFilter(svg: Element, shadow: ShapeShadow, rawId: string): string {
    const svgNs = 'http://www.w3.org/2000/svg';
    const id = rawId.replace(/[^a-zA-Z0-9_-]/g, '-');
    const defs = this.root.ownerDocument.createElementNS(svgNs, 'defs');
    const filter = this.root.ownerDocument.createElementNS(svgNs, 'filter');
    filter.setAttribute('id', id);
    for (const [name, value] of [['x', '-50%'], ['y', '-50%'], ['width', '200%'], ['height', '200%']] as const) filter.setAttribute(name, value);
    const drop = this.root.ownerDocument.createElementNS(svgNs, 'feDropShadow');
    drop.setAttribute('dx', String(shadow.dxPx));
    drop.setAttribute('dy', String(shadow.dyPx));
    drop.setAttribute('stdDeviation', String(shadow.blurPx / 2));
    drop.setAttribute('flood-color', shadow.color);
    filter.appendChild(drop);
    defs.appendChild(filter);
    svg.appendChild(defs);
    return id;
  }

  /** OMML 公式渲染成 MathML 节点。正文段落和文本框 / 形状里的段落共用。 */
  private makeMath(info: MathInfo, index: number | undefined): HTMLElement {
    const ns = 'http://www.w3.org/1998/Math/MathML';
    const build = (node: MathMlNode): Element => {
      const element = this.root.ownerDocument.createElementNS(ns, node.tag);
      for (const [name, value] of Object.entries(node.attrs ?? {})) element.setAttribute(name, value);
      if (node.text !== undefined) element.textContent = node.text;
      for (const child of node.children ?? []) element.append(build(child));
      return element;
    };
    const root = build(info.mathMl) as HTMLElement;
    root.setAttribute('aria-label', info.linear);
    root.contentEditable = 'false';
    root.dataset.docxMath = '1';
    // 正文公式带索引，宿主拿它调 setMath()；文本框 / 形状里的公式不在正文的段落索引里，不给索引，
    // 改标 data-docx-shape-math，免得宿主拿段内序号去改正文里的另一个公式。
    if (index === undefined) root.dataset.docxShapeMath = '1';
    else root.dataset.docxMathIndex = String(index);
    if (info.display === 'block') root.style.display = 'block';
    const semantics = this.root.ownerDocument.createElementNS(ns, 'semantics');
    const annotation = this.root.ownerDocument.createElementNS(ns, 'annotation');
    annotation.setAttribute('encoding', 'text/plain');
    annotation.textContent = info.linear;
    semantics.append(annotation);
    root.append(semantics);
    return root;
  }

  private makeEquation(equation: EquationNode): HTMLElement {
    const ns = 'http://www.w3.org/1998/Math/MathML';
    const build = (node: MathMlNode): Element => {
      const element = this.root.ownerDocument.createElementNS(ns, node.tag);
      for (const [name, value] of Object.entries(node.attrs ?? {})) element.setAttribute(name, value);
      if (node.text !== undefined) element.textContent = node.text;
      for (const child of node.children ?? []) element.append(build(child));
      return element;
    };
    const root = this.root.ownerDocument.createElementNS(ns, 'math') as unknown as HTMLElement;
    root.append(build(equationToMathMl(equation)));
    root.contentEditable = 'false';
    root.dataset.docxEquation = 'math';
    root.setAttribute('data-docx-mark', '1');
    root.style.userSelect = 'none';
    return root;
  }

  private makeImage(paragraph: number, image: ImageInfo): HTMLElement {
    const wrapper = this.root.ownerDocument.createElement(image.placement === 'floating' ? 'div' : 'span');
    wrapper.className = `docx-image${this.selectedImageInfo?.id === image.id ? ' selected' : ''}`;
    wrapper.contentEditable = 'false';
    wrapper.tabIndex = 0;
    wrapper.dataset.image = image.id;
    wrapper.dataset.paragraph = String(paragraph);
    wrapper.style.position = 'relative';
    wrapper.style.display = image.placement === 'floating' ? 'block' : 'inline-block';
    wrapper.style.width = `${Math.max(1, image.widthPx || 1)}px`;
    wrapper.style.height = `${Math.max(1, image.heightPx || 1)}px`;
    wrapper.style.maxWidth = '100%';
    wrapper.style.verticalAlign = 'text-bottom';
    wrapper.style.margin = image.placement === 'floating' ? '8px 12px 8px 0' : '0 2px';
    wrapper.style.overflow = 'hidden';
    if (image.placement === 'floating') {
      if (['square', 'tight', 'through'].includes(image.wrap ?? '')) wrapper.style.cssFloat = 'left';
      else if (image.wrap === 'topAndBottom') { wrapper.style.margin = '12px auto'; }
      else if (image.wrap === 'none') { wrapper.style.position = 'absolute'; wrapper.style.right = '0'; }
      wrapper.style.zIndex = image.behindDoc ? '0' : '1';
    }
    const viewport = this.root.ownerDocument.createElement('span');
    viewport.style.display = 'block';
    viewport.style.width = '100%';
    viewport.style.height = '100%';
    viewport.style.overflow = 'hidden';
    const stage = this.root.ownerDocument.createElement('span');
    stage.style.display = 'block';
    stage.style.width = '100%';
    stage.style.height = '100%';
    const img = this.root.ownerDocument.createElement('img');
    img.src = this.document.getImageDataUrl(image);
    img.alt = image.alt ?? '';
    img.draggable = false;
    img.style.width = '100%';
    img.style.height = '100%';
    img.style.display = 'block';
    if (image.crop) {
      const scaleX = 1 / Math.max(0.01, 1 - image.crop.left - image.crop.right);
      const scaleY = 1 / Math.max(0.01, 1 - image.crop.top - image.crop.bottom);
      img.style.width = `${scaleX * 100}%`;
      img.style.height = `${scaleY * 100}%`;
      img.style.transformOrigin = 'top left';
      img.style.transform = `translate(${-image.crop.left * 100}%, ${-image.crop.top * 100}%)`;
    }
    const transforms = [
      image.rotation ? `rotate(${image.rotation}deg)` : '',
      image.flipH ? 'scaleX(-1)' : '',
      image.flipV ? 'scaleY(-1)' : '',
    ].filter(Boolean);
    if (transforms.length) stage.style.transform = transforms.join(' ');
    stage.append(img);
    viewport.append(stage);
    wrapper.append(viewport);
    for (const handle of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
      const node = this.root.ownerDocument.createElement('span');
      node.className = `docx-image-handle docx-image-handle-${handle}`;
      node.dataset.handle = handle;
      node.addEventListener('mousedown', (event) => this.startResize(event, wrapper, image, handle));
      wrapper.append(node);
    }
    wrapper.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.selectParagraph(paragraph);
      this.selectImage(image);
      wrapper.focus();
    });
    wrapper.addEventListener('keydown', (event) => {
      const keyEvent = event as KeyboardEvent;
      if (['Delete', 'Backspace'].includes(keyEvent.key)) {
        keyEvent.preventDefault();
        keyEvent.stopPropagation();
        this.document.deleteImage(image);
        this.render();
        this.options.onChange?.(this.document.getSnapshot());
        return;
      }
      if (!keyEvent.altKey || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(keyEvent.key)) return;
      keyEvent.preventDefault();
      const step = keyEvent.shiftKey ? 16 : 8;
      const delta = keyEvent.key === 'ArrowLeft' || keyEvent.key === 'ArrowUp' ? -step : step;
      if (keyEvent.key === 'ArrowLeft' || keyEvent.key === 'ArrowRight') {
        this.document.resizeImage(image, { widthEmu: pxToEmu(Math.max(1, (image.widthPx || 1) + delta)), keepAspect: keyEvent.shiftKey });
      } else {
        this.document.resizeImage(image, { heightEmu: pxToEmu(Math.max(1, (image.heightPx || 1) + delta)), keepAspect: keyEvent.shiftKey });
      }
      this.render();
      this.options.onChange?.(this.document.getSnapshot());
    });
    return wrapper;
  }

  private startResize(event: MouseEvent, wrapper: HTMLElement, image: ImageInfo, handle: string): void {
    event.preventDefault();
    event.stopPropagation();
    this.selectImage(image);
    const startX = event.clientX;
    const startY = event.clientY;
    const startWidth = Math.max(1, image.widthPx || 1);
    const startHeight = Math.max(1, image.heightPx || 1);
    const move = (next: MouseEvent): void => {
      const horizontal = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
      const vertical = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
      const width = Math.max(1, startWidth + (next.clientX - startX) * horizontal);
      const height = Math.max(1, startHeight + (next.clientY - startY) * vertical);
      wrapper.style.width = `${width}px`;
      wrapper.style.height = `${height}px`;
    };
    const up = (next: MouseEvent): void => {
      this.root.ownerDocument.removeEventListener('mousemove', move);
      this.root.ownerDocument.removeEventListener('mouseup', up);
      const widthEmu = pxToEmu(parseFloat(wrapper.style.width));
      const heightEmu = pxToEmu(parseFloat(wrapper.style.height));
      const resize = next.shiftKey
        ? (Math.abs(next.clientX - startX) >= Math.abs(next.clientY - startY)
          ? { widthEmu, keepAspect: true }
          : { heightEmu, keepAspect: true })
        : { widthEmu, heightEmu };
      this.document.resizeImage(image, resize);
      this.render();
      this.options.onChange?.(this.document.getSnapshot());
    };
    this.root.ownerDocument.addEventListener('mousemove', move);
    this.root.ownerDocument.addEventListener('mouseup', up);
  }

  private insertText(element: HTMLElement, text: string): void {
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount || this.selectionTouchesField(element)) return;
    const range = selection.getRangeAt(0);
    if (!element.contains(range.startContainer) || !element.contains(range.endContainer)) return;
    range.deleteContents();
    const node = this.root.ownerDocument.createTextNode(sanitizeText(text).replace(/\r\n?/g, '\n'));
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  private currentDocumentRange(): DocumentRange | null {
    const range = this.captureDocumentRange();
    return range ? this.documentRange(range) : null;
  }

  private parseClipboardFragment(raw: string | null | undefined): ClipboardFragment | null {
    if (!raw) return null;
    try {
      const value = JSON.parse(raw) as ClipboardFragment;
      if (!value || value.version !== 1 || typeof value.text !== 'string' || !Array.isArray(value.paragraphs)) return null;
      return value;
    } catch {
      return null;
    }
  }

  private writeClipboardFragment(data: DataTransfer | null, fragment: ClipboardFragment): boolean {
    if (!data) return false;
    const json = JSON.stringify(fragment);
    data.setData(DOCX_CLIPBOARD_MIME, json);
    data.setData('text/plain', fragment.text);
    data.setData('text/html', fragment.text);
    return true;
  }

  private mapExternalHtmlFragment(html: string, plainText: string): ClipboardFragment {
    const parser = this.root.ownerDocument.defaultView?.DOMParser;
    if (!parser) return { version: 1, text: sanitizeText(plainText), paragraphs: [{ runs: [{ text: sanitizeText(plainText) }] }] };
    const document = new parser().parseFromString(html, 'text/html');
    const paragraphs: ClipboardFragment['paragraphs'] = [];
    const blocks: NonNullable<ClipboardFragment['blocks']> = [];
    let nextListId = 1;
    type State = { format: RunFormat; hyperlink?: ClipboardRun['hyperlink'] };
    const pushParagraph = (runs: ClipboardRun[]): void => {
      const normalized = runs
        .map((run) => ({ ...run, ...(run.text ? { text: sanitizeText(run.text) } : {}) }))
        .filter((run) => run.text || run.images?.length);
      if (normalized.length) paragraphs.push({ runs: normalized });
    };
    const sameFormat = (left: RunFormat | undefined, right: RunFormat | undefined): boolean => {
      const keys = new Set([...Object.keys(left ?? {}), ...Object.keys(right ?? {})]);
      for (const key of keys) {
        if ((left as Record<string, unknown> | undefined)?.[key] !== (right as Record<string, unknown> | undefined)?.[key]) return false;
      }
      return true;
    };
    const sameHyperlink = (left: ClipboardRun['hyperlink'], right: ClipboardRun['hyperlink']): boolean =>
      (left?.url ?? '') === (right?.url ?? '') &&
      (left?.anchor ?? '') === (right?.anchor ?? '') &&
      (left?.tooltip ?? '') === (right?.tooltip ?? '');
    const pushRun = (runs: ClipboardRun[], run: ClipboardRun): void => {
      const last = runs.at(-1);
      if (last && !last.images?.length && !run.images?.length &&
          sameFormat(last.format, run.format) &&
          sameHyperlink(last.hyperlink, run.hyperlink)) {
        last.text = `${last.text ?? ''}${run.text ?? ''}`;
      } else runs.push(run);
    };
    const walkInline = (node: Node, state: State, runs: ClipboardRun[], depth = 0): void => {
      if (depth > 100) return;
      if (node.nodeType === 3) {
        const text = node.textContent ?? '';
        if (text) pushRun(runs, { text, format: { ...state.format }, ...(state.hyperlink ? { hyperlink: { ...state.hyperlink } } : {}) });
        return;
      }
      if (node.nodeType !== 1) return;
      const element = node as HTMLElement;
      const tag = element.tagName.toLowerCase();
      if (['script', 'style', 'noscript'].includes(tag)) return;
      if (tag === 'br') {
        pushRun(runs, { text: '\n', format: { ...state.format }, ...(state.hyperlink ? { hyperlink: { ...state.hyperlink } } : {}) });
        return;
      }
      if (tag === 'img') {
        const src = element.getAttribute('src') ?? '';
        const match = src.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/=]+)$/);
        if (!match || !['image/png', 'image/jpeg', 'image/gif', 'image/bmp'].includes(match[1]!.toLowerCase())) return;
        pushRun(runs, {
          images: [{
            contentType: match[1]!,
            bytes: match[2]!,
            alt: element.getAttribute('alt') ?? undefined,
          }],
          format: { ...state.format },
          ...(state.hyperlink ? { hyperlink: { ...state.hyperlink } } : {}),
        });
        return;
      }
      const nextState: State = { format: { ...state.format }, hyperlink: state.hyperlink ? { ...state.hyperlink } : undefined };
      if (['b', 'strong'].includes(tag)) nextState.format.bold = true;
      if (['i', 'em'].includes(tag)) nextState.format.italic = true;
      if (tag === 'u') nextState.format.underline = true;
      if (['s', 'del'].includes(tag)) nextState.format.strike = true;
      if (tag === 'a') {
        const href = element.getAttribute('href') ?? '';
        if (href && isSafeHyperlinkUrl(href)) nextState.hyperlink = { url: href };
        else nextState.hyperlink = undefined;
      }
      for (const child of Array.from(node.childNodes)) walkInline(child, nextState, runs, depth + 1);
    };
    const walkList = (list: HTMLElement, kind: 'bullet' | 'decimal', level: number, listId: number, depth: number): void => {
      if (depth > 100) return;
      const items = Array.from(list.children).filter((child) => child.tagName.toLowerCase() === 'li');
      for (const item of items) {
        const runs: ClipboardRun[] = [];
        for (const child of Array.from(item.childNodes)) {
          if (child.nodeType === 1) {
            const tag = (child as HTMLElement).tagName.toLowerCase();
            if (tag === 'ul') { walkList(child as HTMLElement, 'bullet', Math.min(8, level + 1), listId, depth + 1); continue; }
            if (tag === 'ol') { walkList(child as HTMLElement, 'decimal', Math.min(8, level + 1), listId, depth + 1); continue; }
          }
          walkInline(child, { format: {} }, runs, depth + 1);
        }
        pushParagraph(runs);
        if (paragraphs.length) {
          const paragraph = paragraphs[paragraphs.length - 1]!;
          paragraph.numbering = { kind, level, listId };
          blocks.push({ type: 'paragraph', paragraph });
        }
      }
    };
    const walkBlocks = (node: Node, depth = 0): void => {
      if (depth > 100) return;
      if (node.nodeType === 3) {
        const text = node.textContent?.trim();
        if (!text) return;
        const paragraph = { runs: [{ text, format: {} }] };
        pushParagraph(paragraph.runs);
        blocks.push({ type: 'paragraph', paragraph });
        return;
      }
      if (node.nodeType !== 1) return;
      const element = node as HTMLElement;
      const tag = element.tagName.toLowerCase();
      if (['script', 'style', 'noscript'].includes(tag)) return;
      if (['p', 'div'].includes(tag)) {
        const runs: ClipboardRun[] = [];
        for (const child of Array.from(element.childNodes)) walkInline(child, { format: {} }, runs, depth + 1);
        pushParagraph(runs);
        if (paragraphs.length) blocks.push({ type: 'paragraph', paragraph: paragraphs[paragraphs.length - 1]! });
        return;
      }
      if (tag === 'ul' || tag === 'ol') {
        walkList(element, tag === 'ul' ? 'bullet' : 'decimal', 0, nextListId++, depth + 1);
        return;
      }
      if (tag === 'table') {
        const tableRows: ClipboardFragment['paragraphs'][] = [];
        const rows = Array.from(element.children).flatMap((child) => {
          const name = child.tagName.toLowerCase();
          if (name === 'tr') return [child];
          if (['thead', 'tbody', 'tfoot'].includes(name)) {
            return Array.from(child.children).filter((item) => item.tagName.toLowerCase() === 'tr');
          }
          return [];
        });
        for (const row of rows) {
          const cells = Array.from(row.children).filter((child) => ['th', 'td'].includes(child.tagName.toLowerCase()));
          const tableRow: ClipboardFragment['paragraphs'] = [];
          for (const cell of cells) {
            const runs: ClipboardRun[] = [];
            for (const child of Array.from(cell.childNodes)) walkInline(child, { format: {} }, runs, depth + 1);
            const cellParagraph: ClipboardFragment['paragraphs'][number] = { runs };
            tableRow.push(cellParagraph);
          }
          if (tableRow.length) tableRows.push(tableRow);
        }
        if (tableRows.length) blocks.push({ type: 'table', table: { rows: tableRows } });
        return;
      }
      const runs: ClipboardRun[] = [];
      for (const child of Array.from(element.childNodes)) walkInline(child, { format: {} }, runs, depth + 1);
      pushParagraph(runs);
      if (paragraphs.length) blocks.push({ type: 'paragraph', paragraph: paragraphs[paragraphs.length - 1]! });
    };
    for (const child of Array.from(document.body.childNodes)) walkBlocks(child, 0);
    if (!blocks.length) {
      const text = sanitizeText(plainText || document.body.textContent || '');
      return { version: 1, text, paragraphs: [{ runs: [{ text }] }] };
    }
    const allParagraphs = blocks.flatMap((block) => block.type === 'paragraph' ? [block.paragraph] : block.table.rows.flat());
    const text = blocks.map((block) => {
      if (block.type === 'paragraph') return block.paragraph.runs.map((run) => run.text ?? '').join('');
      return block.table.rows
        .map((row) => row.map((cell) => cell.runs.map((run) => run.text ?? '').join('')).join('\t'))
        .join('\n');
    }).join('\n');
    return { version: 1, text, paragraphs: allParagraphs, blocks };
  }

  private handleClipboardCopy(event: ClipboardEvent): void {
    const range = this.currentDocumentRange();
    if (!range) return;
    const fragment = this.document.copyClipboardFragment(range);
    if (!this.writeClipboardFragment(event.clipboardData, fragment)) return;
    event.preventDefault();
  }

  private handleClipboardCut(event: ClipboardEvent, content: HTMLElement): void {
    if (this.selectionTouchesField(content)) {
      event.preventDefault();
      return;
    }
    const range = this.currentDocumentRange();
    if (!range) return;
    if (range.start.paragraph !== range.end.paragraph || range.start.offset === range.end.offset) {
      event.preventDefault();
      return;
    }
    const fragment = this.document.copyClipboardFragment(range);
    if (!this.writeClipboardFragment(event.clipboardData, fragment)) return;
    const before = this.document.revision;
    this.document.beginHistoryGroup('cut');
    let handled = false;
    try {
      handled = this.document.pasteClipboardFragment(range, { version: 1, text: '', paragraphs: [] });
    } finally {
      this.document.endHistoryGroup();
    }
    event.preventDefault();
    if (!handled) return;
    if (this.document.revision === before) {
      this.insertText(content, '');
      return;
    }
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
  }

  private handleClipboardPaste(event: ClipboardEvent, content: HTMLElement): void {
    if (this.selectionTouchesField(content)) {
      event.preventDefault();
      return;
    }
    const data = event.clipboardData;
    if (!data) return;
    const range = this.currentDocumentRange();
    if (!range) return;
    const plain = data.getData('text/plain') ?? '';
    const own = this.parseClipboardFragment(data.getData(DOCX_CLIPBOARD_MIME));
    if (own) {
      try {
        const before = this.document.revision;
        if (this.document.pasteClipboardFragment(range, own)) {
          event.preventDefault();
          if (this.document.revision !== before) {
            this.render();
            this.options.onChange?.(this.document.getSnapshot());
          }
          return;
        }
      } catch (error) {
        this.reportError(error, range.start.paragraph);
      }
    }
    const html = data.getData('text/html');
    if (html) {
      try {
        const mapped = this.mapExternalHtmlFragment(html, plain);
        const before = this.document.revision;
        if (this.document.pasteClipboardFragment(range, mapped)) {
          event.preventDefault();
          if (this.document.revision !== before) {
            this.render();
            this.options.onChange?.(this.document.getSnapshot());
          }
          return;
        }
      } catch (error) {
        this.reportError(error, range.start.paragraph);
      }
    }
    event.preventDefault();
    this.insertText(content, plain);
  }

  private reportError(error: unknown, paragraph: number): void {
    const normalized = error instanceof Error ? error : new Error(String(error));
    try {
      if (this.options.onError) this.options.onError(normalized, { paragraph });
      else console.error(normalized);
    } catch (reportError) {
      console.error(normalized);
      console.error(reportError);
    }
  }

  private caretIn(element: HTMLElement): { start: number; end: number } | null {
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount) return null;
    const range = selection.getRangeAt(0);
    if (!element.contains(range.startContainer) || !element.contains(range.endContainer)) return null;
    const prefix = range.cloneRange();
    prefix.selectNodeContents(element);
    prefix.setEnd(range.startContainer, range.startOffset);
    const start = prefix.toString().length;
    return { start, end: start + range.toString().length };
  }

  private focusParagraphInCell(cell: HTMLTableCellElement | null): void {
    const paragraph = cell?.querySelector<HTMLElement>('[data-paragraph]');
    paragraph?.focus();
  }

  private moveToAdjacentCell(element: HTMLElement, delta: number): boolean {
    const cell = element.closest<HTMLTableCellElement>('td[data-table-cell="true"]');
    const table = cell?.closest('table');
    if (!cell || !table) return false;
    const cells = Array.from(table.querySelectorAll<HTMLTableCellElement>('td[data-table-cell="true"]'));
    const index = cells.indexOf(cell);
    const target = cells[index + delta] ?? null;
    if (!target) return false;
    this.focusParagraphInCell(target);
    return true;
  }

  private moveVerticalCell(element: HTMLElement, delta: number): boolean {
    const cell = element.closest<HTMLTableCellElement>('td[data-table-cell="true"]');
    const table = cell?.closest('table');
    if (!cell || !table) return false;
    const currentCol = Number(cell.dataset.gridStart ?? 0);
    const targetRow = delta < 0 ? Number(cell.dataset.rowStart ?? 0) - 1 : Number(cell.dataset.rowEnd ?? 0);
    const target = Array.from(table.querySelectorAll<HTMLTableCellElement>('td[data-table-cell="true"]')).find((candidate) => {
      const rowStart = Number(candidate.dataset.rowStart ?? -1);
      const rowEnd = Number(candidate.dataset.rowEnd ?? -1);
      const colStart = Number(candidate.dataset.gridStart ?? -1);
      const colEnd = Number(candidate.dataset.gridEnd ?? -1);
      return rowStart <= targetRow && rowEnd > targetRow && colStart <= currentCol && colEnd > currentCol;
    });
    if (!target) return false;
    this.focusParagraphInCell(target);
    return true;
  }

  private textLength(node: Node): number {
    if (node.nodeType === 3) return Array.from(node.textContent ?? '').length;
    if (node.nodeType !== 1) return 0;
    const current = node as HTMLElement;
    if (current.dataset.image || current.dataset.docxMark !== undefined) return 0;
    if (current.dataset.docxDeleted !== undefined) return 0;
    if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return 0;
    if (current.tagName === 'BR') return 1;
    return Array.from(current.childNodes).reduce((total, child) => total + this.textLength(child), 0);
  }

  private paragraphText(index: number): string {
    // 选区换算每次要取两个段落。原先用 getParagraphs().find(...)，等于为一个段落重建整篇读模型
    // （1500 段约 30 ms × 2）。
    return this.document.getParagraph?.(index)?.text
      ?? this.document.getParagraphs().find((item) => item.index === index)?.text ?? '';
  }

  private codeUnitsFromCodePoints(text: string, points: number): number {
    let units = 0;
    let count = 0;
    for (const char of text) {
      if (count >= points) break;
      units += char.length;
      count++;
    }
    return units;
  }

  private documentRange(range: DocumentRange): DocumentRange {
    return {
      start: {
        paragraph: range.start.paragraph,
        offset: this.codeUnitsFromCodePoints(this.paragraphText(range.start.paragraph), range.start.offset),
      },
      end: {
        paragraph: range.end.paragraph,
        offset: this.codeUnitsFromCodePoints(this.paragraphText(range.end.paragraph), range.end.offset),
      },
    };
  }

  private offsetWithin(root: HTMLElement, target: Node, targetOffset: number): number {
    let offset = 0;
    const walk = (node: Node): boolean => {
      if (node === target) {
        if (node.nodeType === 3) {
          offset += Array.from((node.textContent ?? '').slice(0, Math.max(0, targetOffset))).length;
          return true;
        }
        if (node.nodeType !== 1) return true;
        const current = node as HTMLElement;
        if (current.dataset.image || current.dataset.docxMark !== undefined) return true;
        if (current.dataset.docxDeleted !== undefined) return true;
        if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return true;
        if (current.tagName === 'BR') {
          offset += targetOffset > 0 ? 1 : 0;
          return true;
        }
        const childNodes = Array.from(node.childNodes);
        for (let i = 0; i < Math.min(targetOffset, childNodes.length); i++) {
          offset += this.textLength(childNodes[i]!);
        }
        return true;
      }
      if (node.nodeType === 3) {
        offset += Array.from(node.textContent ?? '').length;
        return false;
      }
      if (node.nodeType !== 1) return false;
      const current = node as HTMLElement;
      if (current.dataset.image || current.dataset.docxMark !== undefined) return false;
      if (current.dataset.docxDeleted !== undefined) return false;
      if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return false;
      if (current.tagName === 'BR') {
        offset += 1;
        return false;
      }
      for (const child of Array.from(node.childNodes)) {
        if (walk(child)) return true;
      }
      return false;
    };
    walk(root);
    return offset;
  }

  private positionFromOffset(root: HTMLElement, offset: number): [Node, number] {
    const total = this.textLength(root);
    let remaining = Math.max(0, Math.min(offset, total));
    const locate = (node: Node): [Node, number] | null => {
      if (node.nodeType === 3) {
        const text = node.textContent ?? '';
        const textLength = Array.from(text).length;
        if (remaining <= textLength) return [node, this.codeUnitsFromCodePoints(text, remaining)];
        remaining -= textLength;
        return null;
      }
      if (node.nodeType !== 1) return null;
      const current = node as HTMLElement;
      if (current.dataset.image || current.dataset.docxMark !== undefined) return null;
      if (current.dataset.docxDeleted !== undefined) return null;
      if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return null;
      if (current.tagName === 'BR') {
        if (remaining <= 1) {
          const parent = node.parentNode as Node;
          const index = Array.prototype.indexOf.call(parent.childNodes, node);
          return [parent, remaining === 0 ? index : index + 1];
        }
        remaining -= 1;
        return null;
      }
      for (const child of Array.from(node.childNodes)) {
        const found = locate(child);
        if (found) return found;
      }
      return null;
    };
    return locate(root) ?? [root, root.childNodes.length];
  }

  private captureDocumentRange(): DocumentRange | null {
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount) return null;
    const raw = selection.getRangeAt(0);
    if (!this.root.contains(raw.startContainer) || !this.root.contains(raw.endContainer)) return null;
    const startElement = raw.startContainer.nodeType === 1 ? raw.startContainer as Element : raw.startContainer.parentElement;
    const endElement = raw.endContainer.nodeType === 1 ? raw.endContainer as Element : raw.endContainer.parentElement;
    const startContent = startElement?.closest<HTMLElement>('.docx-paragraph-content');
    const endContent = endElement?.closest<HTMLElement>('.docx-paragraph-content');
    const startParagraph = startContent?.closest<HTMLElement>('[data-paragraph]');
    const endParagraph = endContent?.closest<HTMLElement>('[data-paragraph]');
    if (!startContent || !endContent || !startParagraph || !endParagraph) return null;
    const start = {
      paragraph: Number(startParagraph.dataset.paragraph),
      offset: this.offsetWithin(startContent, raw.startContainer, raw.startOffset),
    };
    const end = {
      paragraph: Number(endParagraph.dataset.paragraph),
      offset: this.offsetWithin(endContent, raw.endContainer, raw.endOffset),
    };
    if (!Number.isSafeInteger(start.paragraph) || !Number.isSafeInteger(end.paragraph)) return null;
    const ordered = start.paragraph > end.paragraph || (start.paragraph === end.paragraph && start.offset > end.offset)
      ? { start: end, end: start }
      : { start, end };
    return ordered;
  }

  private restoreDocumentRange(range: DocumentRange): void {
    const start = this.paragraphs.get(range.start.paragraph)?.content;
    const end = this.paragraphs.get(range.end.paragraph)?.content;
    if (!start || !end) return;
    const selection = this.root.ownerDocument.getSelection();
    if (!selection) return;
    const startPoint = this.positionFromOffset(start, range.start.offset);
    const endPoint = this.positionFromOffset(end, range.end.offset);
    const domRange = this.root.ownerDocument.createRange();
    domRange.setStart(...startPoint);
    domRange.setEnd(...endPoint);
    selection.removeAllRanges();
    selection.addRange(domRange);
  }

  private updateRangeSelection(range: DocumentRange | null): void {
    const key = range
      ? `${range.start.paragraph}:${range.start.offset}-${range.end.paragraph}:${range.end.offset}`
      : '';
    const previous = this.selectedRangeInfo
      ? `${this.selectedRangeInfo.range.start.paragraph}:${this.selectedRangeInfo.range.start.offset}-${this.selectedRangeInfo.range.end.paragraph}:${this.selectedRangeInfo.range.end.offset}`
      : '';
    if (key === previous) return;
    if (!range) {
      const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
      this.selectedRangeInfo = null;
      if (EventClass && previous) this.root.dispatchEvent(new EventClass('docx-rangechange', { bubbles: true, detail: null }));
      return;
    }
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    let format: RunFormat = {};
    try {
      format = this.document.getDocumentRangeFormat(this.documentRange(range));
    } catch (error) {
      if (!(error instanceof Error) || !/Cross-container document ranges are not supported/.test(error.message)) throw error;
      format = {};
    }
    this.selectedRangeInfo = {
      range: {
        start: { ...range.start },
        end: { ...range.end },
      },
      format,
    };
    if (EventClass) {
      this.root.dispatchEvent(new EventClass('docx-rangechange', {
        bubbles: true,
        detail: { range: this.selectedRange, format: this.selectedRangeFormat ?? {} },
      }));
    }
  }

  private selectParagraph(index: number): void {
    const cell = this.document.getTableCellAt(index);
    const previous = this.selectedTableCellInfo;
    const paragraphChanged = this.selected !== index;
    const tableCellChanged =
      (previous === null) !== (cell === null)
      || (previous !== null && cell !== null && (
        previous.table !== cell.table || previous.row !== cell.row || previous.col !== cell.col
      ));
    this.selected = index;
    this.selectedTableCellInfo = cell;
    if (!paragraphChanged && !tableCellChanged) return;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (EventClass && paragraphChanged) {
      this.root.dispatchEvent(new EventClass('docx-selectionchange', { bubbles: true, detail: { index } }));
    }
    if (EventClass && tableCellChanged) {
      this.root.dispatchEvent(new EventClass('docx-tablecellchange', {
        bubbles: true,
        detail: cell ? { cell: this.selectedTableCell } : null,
      }));
    }
  }

  private selectImage(image: ImageInfo | null): void {
    this.selectedImageInfo = image;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (EventClass) this.root.dispatchEvent(new EventClass('docx-imageselectionchange', { bubbles: true, detail: image ? { image } : null }));
  }

  private readonly handleSelection = (): void => {
    const selection = this.root.ownerDocument.getSelection();
    const node = selection?.anchorNode;
    if (!node || !this.root.contains(node)) {
      this.updateRangeSelection(null);
      return;
    }
    this.expandFieldSelection(selection);
    const element = node.nodeType === 1 ? node as Element : node.parentElement;
    const paragraph = element?.closest<HTMLElement>('[data-paragraph]');
    if (paragraph && this.root.contains(paragraph)) this.selectParagraph(Number(paragraph.dataset.paragraph));
    const image = element?.closest<HTMLElement>('[data-image]');
    if (!image) this.selectImage(null);
    this.updateRangeSelection(this.captureDocumentRange());
  };

  private expandFieldSelection(selection: Selection): void {
    if (!selection.rangeCount) return;
    const range = selection.getRangeAt(0);
    const fields = Array.from(this.root.querySelectorAll<HTMLElement>('[data-docx-field-role="result"]'));
    const anchor = selection.anchorNode;
    const anchorElement = anchor?.nodeType === 1 ? anchor as Element : anchor?.parentElement;
    const touched = fields.filter((node) => range.intersectsNode(node) &&
      (!range.collapsed || anchorElement?.closest('[data-docx-field-role="result"]') === node));
    if (!touched.length) return;
    const ids = new Set(touched.map((node) => node.dataset.docxField));
    const selected = fields.filter((node) => ids.has(node.dataset.docxField));
    const expanded = range.cloneRange();
    const first = selected[0]!;
    const last = selected[selected.length - 1]!;
    const bounds = range.cloneRange();
    bounds.selectNode(first);
    if (range.compareBoundaryPoints(Range.START_TO_START, bounds) > 0) expanded.setStartBefore(first);
    bounds.selectNode(last);
    if (range.compareBoundaryPoints(Range.END_TO_END, bounds) < 0) expanded.setEndAfter(last);
    if (expanded.startContainer === range.startContainer && expanded.startOffset === range.startOffset &&
        expanded.endContainer === range.endContainer && expanded.endOffset === range.endOffset) return;
    selection.removeAllRanges();
    selection.addRange(expanded);
  }

  private selectionTouchesField(content: HTMLElement): boolean {
    const selection = this.root?.ownerDocument.getSelection?.();
    if (!selection?.rangeCount) return false;
    const range = selection.getRangeAt(0);
    const anchor = selection.anchorNode;
    const element = anchor?.nodeType === 1 ? anchor as Element : anchor?.parentElement;
    return Array.from(content?.querySelectorAll?.<HTMLElement>('[data-docx-field]') ?? [])
      .some((node) => range.intersectsNode(node) &&
        (!range.collapsed || element === node || node.contains(element ?? null)));
  }

  private deletionTouchesField(content: HTMLElement, direction: 'backward' | 'forward'): boolean {
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount) return false;
    const range = selection.getRangeAt(0);
    if (!range.collapsed || !content.contains(range.startContainer)) return false;
    let node: Node = range.startContainer;
    let offset = range.startOffset;
    while (true) {
      if (node.nodeType === 3) {
        if (direction === 'backward' ? offset > 0 : offset < (node.textContent ?? '').length) return false;
      } else if (node.nodeType === 1) {
        const siblings = node.childNodes;
        while (direction === 'backward' ? offset > 0 : offset < siblings.length) {
          const adjacent = siblings[direction === 'backward' ? --offset : offset++]!;
          if (adjacent.nodeType === 1 && (adjacent as HTMLElement).dataset.docxFieldRole === 'result') return true;
          if (this.textLength(adjacent) > 0) return false;
        }
      }
      if (node === content || !node.parentNode) return false;
      const parent = node.parentNode;
      offset = Array.prototype.indexOf.call(parent.childNodes, node) + (direction === 'forward' ? 1 : 0);
      node = parent;
    }
  }

  private preventFieldDeletion(content: HTMLElement, event: InputEvent): boolean {
    const direction = /^delete(?:Content|Word|SoftLine|HardLine)(Backward|Forward)$/.exec(event.inputType)?.[1];
    if (direction && this.deletionTouchesField(content, direction === 'Backward' ? 'backward' : 'forward')) {
      event.preventDefault();
      return true;
    }
    return false;
  }

  private readonly handleRootKeydown = (event: KeyboardEvent): void => {
    if (this.handleHistoryShortcut(event)) return;
    if (!this.selectedImageInfo || !['Delete', 'Backspace'].includes(event.key)) return;
    const active = this.root.ownerDocument.activeElement as HTMLElement | null;
    const image = active?.closest('[data-image]') as HTMLElement | null;
    if (!image || image.dataset.image !== this.selectedImageInfo.id) return;
    event.preventDefault();
    this.document.deleteImage(this.selectedImageInfo);
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
  };

  private handleHistoryShortcut(event: Pick<KeyboardEvent, 'isComposing' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey' | 'key' | 'preventDefault' | 'target'>): boolean {
    if (event.isComposing || this.composing || !(event.ctrlKey || event.metaKey) || event.altKey) return false;
    const node = event.target as Node | null;
    const target = node?.nodeType === 1 ? node as Element : node?.parentElement ?? null;
    const content = target?.closest<HTMLElement>('.docx-paragraph-content');
    if (!content || !this.root.contains(content)) return false;
    const key = event.key.toLowerCase();
    if (key === 'z' && !event.shiftKey) {
      event.preventDefault();
      this.applyHistory('undo');
      return true;
    }
    if (key === 'y' || (key === 'z' && event.shiftKey)) {
      event.preventDefault();
      this.applyHistory('redo');
      return true;
    }
    return false;
  }

  private applyHistory(direction: 'undo' | 'redo'): void {
    if (this.destroyed) return;
    const previousSelected = this.selected;
    this.flush();
    const beforeRevision = this.document.revision;
    const snapshot = direction === 'undo' ? this.document.undo() : this.document.redo();
    if (snapshot.revision === beforeRevision) return;
    this.render();
    const fallback = snapshot.paragraphs[0]?.index ?? null;
    const targetIndex = previousSelected !== null && this.paragraphs.has(previousSelected) ? previousSelected : fallback;
    if (targetIndex !== null) {
      const paragraph = this.paragraphs.get(targetIndex)?.content;
      if (paragraph) this.focusContent(paragraph);
    }
    this.options.onChange?.(snapshot);
  }

  private captureCaret(): { index: number; start: number; end: number } | null {
    if (!this.root.contains(this.root.ownerDocument.activeElement)) return null;
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount || this.selected === null) return null;
    const paragraph = this.paragraphs.get(this.selected)?.content;
    const range = selection.getRangeAt(0);
    if (!paragraph?.contains(range.startContainer) || !paragraph.contains(range.endContainer)) return null;
    const prefix = range.cloneRange();
    prefix.selectNodeContents(paragraph);
    prefix.setEnd(range.startContainer, range.startOffset);
    const start = prefix.toString().length;
    return { index: this.selected, start, end: start + range.toString().length };
  }

  private restoreCaret(caret: { index: number; start: number; end: number }): void {
    const paragraph = this.paragraphs.get(caret.index)?.content;
    if (!paragraph) return;
    paragraph.focus({ preventScroll: true });
    const range = this.root.ownerDocument.createRange();
    const position = (offset: number): [Node, number] => {
      const walker = this.root.ownerDocument.createTreeWalker(paragraph, 4);
      let node = walker.nextNode();
      while (node) {
        const length = node.textContent?.length ?? 0;
        if (offset <= length) return [node, offset];
        offset -= length;
        node = walker.nextNode();
      }
      return [paragraph, paragraph.childNodes.length];
    };
    range.setStart(...position(caret.start));
    range.setEnd(...position(caret.end));
    const selection = this.root.ownerDocument.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }
}
