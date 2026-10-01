import type { Element } from '@xmldom/xmldom';
import type {
  BorderFormat, BordersFormat, CellBordersFormat, CellFormat, DocumentBlock, MarginFormat, RowFormat, TableCellInfo,
  TableFormat, WidthFormat,
} from './types.js';
import { children, childrenThroughTransparent, WORD_NS, wordValue } from './xml.js';

export const TWIPS_PER_INCH = 1440;
export const EIGHTH_POINTS_PER_POINT = 8;
const PX_PER_INCH = 96;
const DEFAULT_GRID_WIDTH = 2250;

export interface TableCellPosition {
  cell: Element;
  start: number;
  span: number;
  format?: CellFormat;
  vMerge?: 'restart' | 'continue';
  hMerge?: 'restart' | 'continue';
}

export function twipsToPx(value: number): number {
  return value * PX_PER_INCH / TWIPS_PER_INCH;
}

export function eighthPointsToPx(value: number): number {
  return value * PX_PER_INCH / (72 * EIGHTH_POINTS_PER_POINT);
}

export function normalizeWidth(value?: WidthFormat): string | undefined {
  if (!value) return undefined;
  if (value.type === 'pct') return `${value.value / 50}%`;
  if (value.type === 'dxa') return `${twipsToPx(value.value)}px`;
  return value.value === 0 ? 'auto' : undefined;
}

function number(attribute: string | null): number | undefined {
  if (attribute === null || attribute === '') return undefined;
  const value = Number(attribute);
  return Number.isFinite(value) ? value : undefined;
}

function widthOf(element: Element | undefined): WidthFormat | undefined {
  if (!element) return undefined;
  const type = element.getAttributeNS(WORD_NS, 'type') ?? element.getAttribute('w:type') ?? 'dxa';
  const value = number(element.getAttributeNS(WORD_NS, 'w') ?? element.getAttribute('w:w')) ?? 0;
  if (!['auto', 'dxa', 'pct'].includes(type)) return undefined;
  return { type: type as WidthFormat['type'], value };
}

function boolValue(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  const value = wordValue(element);
  return !['0', 'false', 'off'].includes(value ?? '1');
}

function mergeValue(element: Element | undefined): 'restart' | 'continue' | undefined {
  if (!element) return undefined;
  const value = wordValue(element);
  if (!value || value === 'continue') return 'continue';
  return value === 'restart' ? 'restart' : undefined;
}

function borderOf(element: Element | undefined): BorderFormat | undefined {
  if (!element) return undefined;
  const style = wordValue(element) ?? undefined;
  const none = ['nil', 'none'].includes(style ?? '');
  return {
    style,
    size: number(element.getAttributeNS(WORD_NS, 'sz') ?? element.getAttribute('w:sz')),
    space: number(element.getAttributeNS(WORD_NS, 'space') ?? element.getAttribute('w:space')),
    color: normalizeColor(element.getAttributeNS(WORD_NS, 'color') ?? element.getAttribute('w:color') ?? undefined),
    none,
  };
}

export function normalizeColor(color: string | undefined): string | undefined {
  if (!color || color === 'auto') return undefined;
  return /^[a-f\d]{6}$/i.test(color) ? color.toUpperCase() : undefined;
}

function bordersOf(element: Element | undefined): BordersFormat | undefined {
  if (!element) return undefined;
  const borders: BordersFormat = {};
  for (const side of ['top', 'right', 'bottom', 'left', 'insideH', 'insideV'] as const) {
    const border = borderOf(children(element, side)[0]);
    if (border) borders[side] = border;
  }
  return Object.keys(borders).length ? borders : undefined;
}

/** `w:tcBorders`：四边 + 内侧，再加只有单元格才有的两条对角线。 */
function cellBordersOf(element: Element | undefined): CellBordersFormat | undefined {
  if (!element) return undefined;
  const borders: CellBordersFormat = { ...bordersOf(element) };
  for (const side of ['tl2br', 'tr2bl'] as const) {
    const border = borderOf(children(element, side)[0]);
    if (border) borders[side] = border;
  }
  return Object.keys(borders).length ? borders : undefined;
}

function shadingOf(element: Element | undefined): TableFormat['shading'] | CellFormat['shading'] {
  if (!element) return undefined;
  const fill = normalizeColor(element.getAttributeNS(WORD_NS, 'fill') ?? element.getAttribute('w:fill') ?? undefined);
  const color = normalizeColor(element.getAttributeNS(WORD_NS, 'color') ?? element.getAttribute('w:color') ?? undefined);
  const value = wordValue(element) ?? undefined;
  return fill || color || value ? { fill, color, value } : undefined;
}

function marginsOf(element: Element | undefined): MarginFormat | undefined {
  if (!element) return undefined;
  const margin: MarginFormat = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const child = children(element, side)[0];
    if (child) margin[side] = widthOf(child);
  }
  return Object.keys(margin).length ? margin : undefined;
}

const FLOAT_ANCHORS = ['margin', 'page', 'text'];
const FLOAT_X_SPEC = ['center', 'inside', 'left', 'outside', 'right'];
const FLOAT_Y_SPEC = ['bottom', 'center', 'inside', 'inline', 'outside', 'top'];

/** `w:tblpPr` 的信息全在属性上，没有 `w:val` 子元素。 */
function parseFloatingPosition(element: Element | undefined): TableFormat['floatingPosition'] {
  if (!element) return undefined;
  const attr = (name: string) =>
    element.getAttributeNS(WORD_NS, name) ?? element.getAttribute(`w:${name}`) ?? undefined;
  const pick = (name: string, allowed: string[]) => {
    const raw = attr(name);
    return raw !== undefined && allowed.includes(raw) ? raw : undefined;
  };
  const position = {
    leftFromText: number(attr('leftFromText') ?? null),
    rightFromText: number(attr('rightFromText') ?? null),
    topFromText: number(attr('topFromText') ?? null),
    bottomFromText: number(attr('bottomFromText') ?? null),
    verticalAnchor: pick('vertAnchor', FLOAT_ANCHORS),
    horizontalAnchor: pick('horzAnchor', FLOAT_ANCHORS),
    xSpec: pick('tblpXSpec', FLOAT_X_SPEC),
    x: number(attr('tblpX') ?? null),
    ySpec: pick('tblpYSpec', FLOAT_Y_SPEC),
    y: number(attr('tblpY') ?? null),
  } as NonNullable<TableFormat['floatingPosition']>;
  // w:tblpPr 在场就是浮动表格，即便一个属性都没给（都取默认值），所以不因为空对象而返回
  // undefined —— 那会把「浮动」这件事本身丢掉。
  return Object.fromEntries(Object.entries(position)
    .filter(([, value]) => value !== undefined)) as NonNullable<TableFormat['floatingPosition']>;
}

/** `w:tblOverlap`，只有 never / overlap 两个值；别的照未设置处理。 */
function overlapOf(element: Element | undefined): TableFormat['overlap'] {
  const value = wordValue(element);
  return value === 'never' || value === 'overlap' ? value : undefined;
}

export function parseTableFormat(tblPr: Element | undefined): TableFormat | undefined {
  if (!tblPr) return undefined;
  const alignment = wordValue(children(tblPr, 'jc')[0]);
  // w:tblLayout 的值在 w:type 上，不是 w:val —— Word 写的就是 <w:tblLayout w:type="fixed"/>。
  // 只读 w:val 的话这个属性永远读不出来，Word 按固定列宽排的表格在这里会退成 auto。
  const layoutElement = children(tblPr, 'tblLayout')[0];
  const layout = layoutElement?.getAttributeNS(WORD_NS, 'type') ?? layoutElement?.getAttribute('w:type')
    ?? wordValue(layoutElement);
  const format: TableFormat = {
    width: widthOf(children(tblPr, 'tblW')[0]),
    alignment: ['left', 'center', 'right'].includes(alignment ?? '') ? alignment as TableFormat['alignment'] : undefined,
    indent: number(children(tblPr, 'tblInd')[0]?.getAttributeNS(WORD_NS, 'w') ?? children(tblPr, 'tblInd')[0]?.getAttribute('w:w') ?? null),
    borders: bordersOf(children(tblPr, 'tblBorders')[0]),
    shading: shadingOf(children(tblPr, 'shd')[0]),
    cellMargin: marginsOf(children(tblPr, 'tblCellMar')[0]),
    layout: ['fixed', 'autofit'].includes(layout ?? '') ? layout as TableFormat['layout'] : undefined,
    cellSpacing: widthOf(children(tblPr, 'tblCellSpacing')[0]),
    overlap: overlapOf(children(tblPr, 'tblOverlap')[0]),
    style: wordValue(children(tblPr, 'tblStyle')[0]) ?? undefined,
    look: wordValue(children(tblPr, 'tblLook')[0]) ?? undefined,
    caption: wordValue(children(tblPr, 'tblCaption')[0]) ?? undefined,
    description: wordValue(children(tblPr, 'tblDescription')[0]) ?? undefined,
    bidiVisual: boolValue(children(tblPr, 'bidiVisual')[0]),
    floatingPosition: parseFloatingPosition(children(tblPr, 'tblpPr')[0]),
  };
  return Object.values(format).some(value => value !== undefined) ? format : undefined;
}

/** `w:gridBefore` / `w:gridAfter` 跳过的网格列数；负数和非整数按未设置处理。 */
function gridSkip(trPr: Element, name: string): number | undefined {
  const element = children(trPr, name)[0];
  if (!element) return undefined;
  const raw = element.getAttributeNS(WORD_NS, 'val') ?? element.getAttribute('w:val');
  const value = number(raw);
  return value !== undefined && Number.isSafeInteger(value) && value > 0 ? value : undefined;
}

export function parseRowFormat(trPr: Element | undefined): RowFormat | undefined {
  if (!trPr) return undefined;
  const height = children(trPr, 'trHeight')[0];
  const alignment = wordValue(children(trPr, 'jc')[0]);
  const rule = height?.getAttributeNS(WORD_NS, 'hRule') ?? height?.getAttribute('w:hRule') ?? undefined;
  const format: RowFormat = {
    height: height ? {
      value: number(height.getAttributeNS(WORD_NS, 'val') ?? height.getAttribute('w:val')) ?? 0,
      rule: rule === 'exact' || rule === 'atLeast' ? rule : undefined,
    } : undefined,
    cellSpacing: widthOf(children(trPr, 'tblCellSpacing')[0]),
    cantSplit: boolValue(children(trPr, 'cantSplit')[0]),
    header: boolValue(children(trPr, 'tblHeader')[0]),
    alignment: ['left', 'center', 'right'].includes(alignment ?? '') ? alignment as RowFormat['alignment'] : undefined,
    deleted: Boolean(children(trPr, 'del')[0]),
    inserted: Boolean(children(trPr, 'ins')[0]),
    gridBefore: gridSkip(trPr, 'gridBefore'),
    widthBefore: widthOf(children(trPr, 'wBefore')[0]),
    gridAfter: gridSkip(trPr, 'gridAfter'),
    widthAfter: widthOf(children(trPr, 'wAfter')[0]),
  };
  return Object.values(format).some(value => value !== undefined) ? format : undefined;
}

export function parseCellFormat(tcPr: Element | undefined): CellFormat | undefined {
  if (!tcPr) return undefined;
  const vAlign = wordValue(children(tcPr, 'vAlign')[0]);
  const format: CellFormat = {
    width: widthOf(children(tcPr, 'tcW')[0]),
    borders: cellBordersOf(children(tcPr, 'tcBorders')[0]),
    shading: shadingOf(children(tcPr, 'shd')[0]),
    margin: marginsOf(children(tcPr, 'tcMar')[0]),
    verticalAlign: ['top', 'center', 'bottom'].includes(vAlign ?? '') ? vAlign as CellFormat['verticalAlign'] : undefined,
    textDirection: wordValue(children(tcPr, 'textDirection')[0]) ?? undefined,
    noWrap: boolValue(children(tcPr, 'noWrap')[0]),
    fitText: boolValue(children(tcPr, 'tcFitText')[0]),
    hideMark: boolValue(children(tcPr, 'hideMark')[0]),
    hMerge: mergeValue(children(tcPr, 'hMerge')[0]),
    vMerge: mergeValue(children(tcPr, 'vMerge')[0]),
  };
  // 第 9 条：`w:val="0"` 是**显式关闭**，不是未设置。把 false 一并算成「什么都没有」的话，
  // 只写了 `<w:noWrap w:val="0"/>` 的单元格读出来是 undefined，调用方分不出「关掉」和
  // 「没说」，原样写回时那条显式关闭就消失了。
  return Object.values(format).some(value => value !== undefined) ? format : undefined;
}

export function tableGrid(table: Element): number[] {
  const grid = children(children(table, 'tblGrid')[0] ?? table, 'gridCol').map((column) => (
    number(column.getAttributeNS(WORD_NS, 'w') ?? column.getAttribute('w:w')) ?? DEFAULT_GRID_WIDTH
  ));
  if (grid.length) return grid;
  const columns = childrenThroughTransparent(table, 'tr')
    .reduce((max, row) => Math.max(max, rowCells(row).reduce((count, cell) => count + cell.span, 0)), 0);
  return Array.from({ length: columns }, () => DEFAULT_GRID_WIDTH);
}

export function cellSpan(cell: Element): number {
  const tcPr = children(cell, 'tcPr')[0];
  return Math.max(1, number(children(tcPr ?? cell, 'gridSpan')[0]?.getAttributeNS(WORD_NS, 'val') ?? children(tcPr ?? cell, 'gridSpan')[0]?.getAttribute('w:val') ?? null) ?? 1);
}

export function rowCells(row: Element): TableCellPosition[] {
  // w:gridBefore 跳过的网格列也占位置。从 0 起算的话，带 gridBefore 的行里每个单元格的网格
  // 列号都会偏小，而跨行合并是靠网格起始列匹配的（vMerge continue 要对上上面那个 restart），
  // 列号一偏合并就断掉——这比少画一块空白严重得多。
  let start = parseRowFormat(children(row, 'trPr')[0])?.gridBefore ?? 0;
  return childrenThroughTransparent(row, 'tc').map((cell) => {
    const tcPr = children(cell, 'tcPr')[0];
    const span = cellSpan(cell);
    const position: TableCellPosition = {
      cell,
      start,
      span,
      format: parseCellFormat(tcPr),
      vMerge: mergeValue(children(tcPr ?? cell, 'vMerge')[0]),
      hMerge: mergeValue(children(tcPr ?? cell, 'hMerge')[0]),
    };
    start += span;
    return position;
  });
}

export function readTable(table: Element, walk: (parent: Element) => DocumentBlock[],
  /**
   * 算上表格样式之后的单元格格式。表格读取本身不碰样式（这个模块是结构读取的叶子），所以
   * 由调用方注入；不给就只有单元格自己的直接格式。
   */
  effectiveCellFormat?: (cell: Element, direct: CellFormat | undefined) => CellFormat | undefined,
  /** 同上，行格式。 */
  effectiveRowFormat?: (row: Element, direct: RowFormat | undefined) => RowFormat | undefined,
  /** 同上，表格级格式。 */
  effectiveTableFormat?: (table: Element, direct: TableFormat | undefined) => TableFormat | undefined,
): Extract<DocumentBlock, { type: 'table' }> {
  const grid = tableGrid(table);
  const active = new Map<number, { start: number; end: number; master: TableCellInfo }>();
  const rows = childrenThroughTransparent(table, 'tr').map((row) => {
    const nextActive = new Map<number, { start: number; end: number; master: TableCellInfo }>();
    const cells = rowCells(row).map<TableCellInfo>((position) => {
      const colSpan = position.span;
      const cell: TableCellInfo = {
        blocks: walk(position.cell),
        colSpan,
        rowSpan: 1,
        isMergeContinuation: position.hMerge === 'continue' || false,
        format: position.format,
        ...(() => {
          const effective = effectiveCellFormat?.(position.cell, position.format);
          return effective ? { effective } : {};
        })(),
      };
      const master = active.get(position.start);
      const sameMerge = master && master.start === position.start && master.end === position.start + position.span;
      if (position.vMerge === 'continue' && sameMerge) {
        master.master.rowSpan += 1;
        cell.rowSpan = 0;
        cell.isMergeContinuation = true;
        nextActive.set(position.start, master);
      } else if (position.vMerge === 'restart') {
        const merge = { start: position.start, end: position.start + position.span, master: cell };
        nextActive.set(position.start, merge);
      }
      return cell;
    });
    for (const [start, merge] of nextActive) active.set(start, merge);
    for (const [start] of [...active]) if (!nextActive.has(start)) active.delete(start);
    const format = parseRowFormat(children(row, 'trPr')[0]);
    const effective = effectiveRowFormat?.(row, format);
    return { cells, format, ...(effective ? { effective } : {}) };
  });
  const tableFormat = parseTableFormat(children(table, 'tblPr')[0]);
  const effectiveTable = effectiveTableFormat?.(table, tableFormat);
  return {
    type: 'table', rows, format: tableFormat,
    ...(effectiveTable ? { effective: effectiveTable } : {}), grid,
  };
}
