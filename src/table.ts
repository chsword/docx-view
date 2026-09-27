import type { Element } from '@xmldom/xmldom';
import type {
  BorderFormat, BordersFormat, CellFormat, DocumentBlock, MarginFormat, RowFormat, TableCellInfo, TableFormat, WidthFormat,
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

export function parseTableFormat(tblPr: Element | undefined): TableFormat | undefined {
  if (!tblPr) return undefined;
  const alignment = wordValue(children(tblPr, 'jc')[0]);
  const layout = wordValue(children(tblPr, 'tblLayout')[0]);
  const format: TableFormat = {
    width: widthOf(children(tblPr, 'tblW')[0]),
    alignment: ['left', 'center', 'right'].includes(alignment ?? '') ? alignment as TableFormat['alignment'] : undefined,
    indent: number(children(tblPr, 'tblInd')[0]?.getAttributeNS(WORD_NS, 'w') ?? children(tblPr, 'tblInd')[0]?.getAttribute('w:w') ?? null),
    borders: bordersOf(children(tblPr, 'tblBorders')[0]),
    shading: shadingOf(children(tblPr, 'shd')[0]),
    cellMargin: marginsOf(children(tblPr, 'tblCellMar')[0]),
    layout: ['fixed', 'autofit'].includes(layout ?? '') ? layout as TableFormat['layout'] : undefined,
    style: wordValue(children(tblPr, 'tblStyle')[0]) ?? undefined,
    look: wordValue(children(tblPr, 'tblLook')[0]) ?? undefined,
    caption: wordValue(children(tblPr, 'tblCaption')[0]) ?? undefined,
    description: wordValue(children(tblPr, 'tblDescription')[0]) ?? undefined,
  };
  return Object.values(format).some(value => value !== undefined) ? format : undefined;
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
    cantSplit: boolValue(children(trPr, 'cantSplit')[0]),
    header: boolValue(children(trPr, 'tblHeader')[0]),
    alignment: ['left', 'center', 'right'].includes(alignment ?? '') ? alignment as RowFormat['alignment'] : undefined,
    deleted: Boolean(children(trPr, 'del')[0]),
    inserted: Boolean(children(trPr, 'ins')[0]),
  };
  return Object.values(format).some(value => value !== undefined && value !== false) ? format : undefined;
}

export function parseCellFormat(tcPr: Element | undefined): CellFormat | undefined {
  if (!tcPr) return undefined;
  const vAlign = wordValue(children(tcPr, 'vAlign')[0]);
  const format: CellFormat = {
    width: widthOf(children(tcPr, 'tcW')[0]),
    borders: bordersOf(children(tcPr, 'tcBorders')[0]),
    shading: shadingOf(children(tcPr, 'shd')[0]),
    margin: marginsOf(children(tcPr, 'tcMar')[0]),
    verticalAlign: ['top', 'center', 'bottom'].includes(vAlign ?? '') ? vAlign as CellFormat['verticalAlign'] : undefined,
    textDirection: wordValue(children(tcPr, 'textDirection')[0]) ?? undefined,
    noWrap: boolValue(children(tcPr, 'noWrap')[0]),
    hideMark: boolValue(children(tcPr, 'hideMark')[0]),
    hMerge: mergeValue(children(tcPr, 'hMerge')[0]),
    vMerge: mergeValue(children(tcPr, 'vMerge')[0]),
  };
  return Object.values(format).some(value => value !== undefined && value !== false) ? format : undefined;
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
  let start = 0;
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

export function readTable(table: Element, walk: (parent: Element) => DocumentBlock[]): Extract<DocumentBlock, { type: 'table' }> {
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
    return { cells, format: parseRowFormat(children(row, 'trPr')[0]) };
  });
  return { type: 'table', rows, format: parseTableFormat(children(table, 'tblPr')[0]), grid };
}
