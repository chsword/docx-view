import JSZip from 'jszip';
import type { Document, Element, Node } from '@xmldom/xmldom';
import type {
  AgentRequest, CellFormat, DocumentBlock, DocumentSnapshot, ParagraphFormat, ParagraphInfo, RowFormat, RunFormat, RunInfo, TableFormat, TableInfo,
} from './types.js';
import {
  assertText, children, CONTENT_TYPES_NS, descendants, OFFICE_DOCUMENT_REL, parseXml, REL_NS,
  serializeXml, setWordValue, validatePath, WORD_NS, wordElement, wordValue,
} from './xml.js';
import { assertIndex, validateParagraphFormat, validateRequest, validateRows, validateRunFormat } from './operations.js';
import { cellSpan, parseCellFormat, parseTableFormat, readTable, rowCells, tableGrid } from './table.js';

const MAX_ARCHIVE = 50 * 1024 * 1024;
const MAX_PART = 16 * 1024 * 1024;
const MAX_TOTAL = 64 * 1024 * 1024;
const MAX_PARTS = 2048;
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const MAIN_TYPE = `${DOCX_TYPE}.main+xml`;
const encoder = new TextEncoder();

function decodeXml(bytes: Uint8Array): string {
  const utf16le = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0);
  const utf16be = (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c);
  return new TextDecoder(utf16le ? 'utf-16le' : utf16be ? 'utf-16be' : 'utf-8', { fatal: true }).decode(bytes);
}

function encodeXml(xml: string): Uint8Array {
  return encoder.encode(xml.replace(/^(<\?xml\b[^?]*\bencoding\s*=\s*)(["'])[^"']*\2/i, '$1"UTF-8"'));
}

function bodyOf(document: Document): Element {
  const root = document.documentElement;
  if (!root || root.namespaceURI !== WORD_NS || root.localName !== 'document') {
    throw new Error('Only transitional WordprocessingML documents are supported.');
  }
  const body = children(root, 'body');
  if (body.length !== 1) throw new Error('Main document must have exactly one w:body.');
  return body[0]!;
}

function paragraphAt(document: Document, index: number): Element {
  assertIndex(index);
  const paragraph = descendants(bodyOf(document), 'p')[index];
  if (!paragraph) throw new Error(`Paragraph ${index} does not exist.`);
  return paragraph;
}

function textElements(element: Element): Element[] {
  const result: Element[] = [];
  function walk(node: Node): void {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.namespaceURI === WORD_NS) {
        if (element.localName === 'p') continue;
        if (['t', 'tab', 'br', 'cr'].includes(element.localName ?? '')) {
          result.push(element);
          continue;
        }
      }
      walk(element);
    }
  }
  walk(element);
  return result;
}

function elementText(element: Element): string {
  return element.localName === 't' ? element.textContent ?? '' : element.localName === 'tab' ? '\t' : '\n';
}

function textOf(element: Element): string {
  return textElements(element).map(elementText).join('');
}

function appendText(parent: Element, text: string, before: Node | null = null): void {
  const document = parent.ownerDocument!;
  for (const chunk of text.split(/(\t|\r\n|\r|\n)/)) {
    if (!chunk) continue;
    const element = wordElement(document, chunk === '\t' ? 'tab' : /^[\r\n]+$/.test(chunk) ? 'br' : 't');
    if (element.localName === 't') {
      element.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      element.appendChild(document.createTextNode(chunk));
    }
    parent.insertBefore(element, before);
  }
}

function newParagraph(document: Document, text: string): Element {
  const paragraph = wordElement(document, 'p');
  const run = wordElement(document, 'r');
  appendText(run, text);
  paragraph.appendChild(run);
  return paragraph;
}

// Edit only text-bearing nodes. Drawings, bookmarks, field codes and other XML survive.
function replaceSpan(paragraph: Element, start: number, end: number, replacement: string): void {
  const elements = textElements(paragraph);
  if (elements.length === 0) {
    const run = children(paragraph, 'r')[0] ?? wordElement(paragraph.ownerDocument!, 'r');
    if (!run.parentNode) paragraph.appendChild(run);
    appendText(run, replacement);
    return;
  }
  let offset = 0;
  let inserted = false;
  for (const element of elements) {
    const oldText = elementText(element);
    const next = offset + oldText.length;
    const insertHere = !inserted && start >= offset && (start === end ? start <= next : start < next);
    const overlaps = offset < end && next > start;
    if (insertHere || overlaps) {
      if (insertHere && !overlaps && element.localName !== 't') {
        appendText(element.parentNode as Element, replacement, start === offset ? element : element.nextSibling);
        inserted = true;
        offset = next;
        continue;
      }
      const prefix = oldText.slice(0, Math.max(0, Math.min(oldText.length, start - offset)));
      const suffix = oldText.slice(Math.max(0, Math.min(oldText.length, end - offset)));
      const text = prefix + (insertHere ? replacement : '') + suffix;
      const parent = element.parentNode as Element;
      appendText(parent, text, element);
      parent.removeChild(element);
      if (insertHere) inserted = true;
    }
    offset = next;
  }
}

function properties(element: Element, name: 'pPr' | 'rPr'): Element {
  let result = children(element, name)[0];
  if (!result) {
    result = wordElement(element.ownerDocument!, name);
    element.insertBefore(result, element.firstChild);
  }
  return result;
}

// Known property order keeps generated pPr/rPr conformant without dropping unknown properties.
const PROPERTY_ORDER = {
  pPr: ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
    'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
    'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
    'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap',
    'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId',
    'cnfStyle', 'rPr', 'sectPr', 'pPrChange'],
  rPr: ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden',
    'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect',
    'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
    'specVanish', 'oMath', 'rPrChange'],
  tblPr: ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
    'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar',
    'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'],
  trPr: ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
    'tblHeader', 'jc', 'hidden', 'ins', 'del', 'trPrChange'],
  tcPr: ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
    'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'],
  tblBorders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'],
  tcBorders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'],
  tblCellMar: ['top', 'left', 'bottom', 'right'],
  tcMar: ['top', 'left', 'bottom', 'right'],
};

function property(parent: Element, name: string): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    const order = PROPERTY_ORDER[parent.localName as keyof typeof PROPERTY_ORDER] ?? [];
    const position = order.indexOf(name);
    const following = position === -1 ? undefined : children(parent).find(child => {
      const childPosition = order.indexOf(child.localName!);
      return childPosition > position;
    });
    parent.insertBefore(result, following ?? null);
  }
  return result;
}

function ownRuns(paragraph: Element): Element[] {
  return descendants(paragraph, 'r').filter(run => {
    let parent = run.parentNode;
    while (parent && parent !== paragraph) {
      if (parent.nodeType === 1 && (parent as Element).namespaceURI === WORD_NS &&
          (parent as Element).localName === 'p') return false;
      parent = parent.parentNode;
    }
    return true;
  });
}

function readRun(run: Element, index: number): RunInfo {
  const props = children(run, 'rPr')[0];
  const get = (name: string) => props ? children(props, name)[0] : undefined;
  const toggle = (name: string) => get(name) ? !['0', 'false', 'off'].includes(wordValue(get(name)) ?? '') : undefined;
  const size = wordValue(get('sz'));
  const underline = get('u');
  const color = wordValue(get('color'));
  return {
    index, text: textOf(run), bold: toggle('b'), italic: toggle('i'),
    underline: underline ? !['none', '0', 'false'].includes(wordValue(underline) ?? '') : undefined,
    fontSize: size && Number.isFinite(Number(size)) ? Number(size) / 2 : undefined,
    fontFamily: get('rFonts')?.getAttributeNS(WORD_NS, 'ascii') ?? undefined,
    color: color && /^[a-f\d]{6}$/i.test(color) ? color : undefined,
  };
}

function readParagraph(paragraph: Element, index: number): ParagraphInfo {
  const props = children(paragraph, 'pPr')[0];
  const alignment = props ? wordValue(children(props, 'jc')[0]) : undefined;
  return {
    index, text: textOf(paragraph), runs: ownRuns(paragraph).map(readRun),
    style: props ? wordValue(children(props, 'pStyle')[0]) : undefined,
    alignment: ['left', 'center', 'right', 'both'].includes(alignment ?? '')
      ? alignment as ParagraphFormat['alignment'] : undefined,
  };
}

function bodyBlocks(body: Element): Element[] {
  return children(body).filter(child => ['p', 'tbl'].includes(child.localName ?? ''));
}

function tableAt(document: Document, index: number): Element {
  assertIndex(index);
  const table = children(bodyOf(document), 'tbl')[index];
  if (!table) throw new Error(`Table ${index} does not exist.`);
  return table;
}

function bodyBlockAt(document: Document, index: number): Element {
  assertIndex(index);
  const block = bodyBlocks(bodyOf(document))[index];
  if (!block) throw new Error(`Block ${index} does not exist.`);
  return block;
}

function tableProperty(parent: Element, name: 'tblPr' | 'trPr' | 'tcPr'): Element {
  let result = children(parent, name)[0];
  if (!result) {
    result = wordElement(parent.ownerDocument!, name);
    parent.insertBefore(result, parent.firstChild);
  }
  return result;
}

function ensureTableGrid(table: Element): Element {
  let grid = children(table, 'tblGrid')[0];
  if (!grid) {
    grid = wordElement(table.ownerDocument!, 'tblGrid');
    const tableProps = children(table, 'tblPr')[0];
    table.insertBefore(grid, tableProps?.nextSibling ?? children(table, 'tr')[0] ?? null);
    for (const width of tableGrid(table)) grid.appendChild(gridCol(table.ownerDocument!, width));
  } else if (!children(grid, 'gridCol').length) {
    for (const width of tableGrid(table)) grid.appendChild(gridCol(table.ownerDocument!, width));
  }
  return grid;
}

function removeWordChildren(parent: Element | undefined, ...names: string[]): void {
  if (!parent) return;
  for (const child of children(parent)) {
    if (names.includes(child.localName!)) parent.removeChild(child);
  }
}

function ensureCellParagraph(cell: Element): void {
  if (!children(cell).some(child => child.localName === 'p')) cell.appendChild(newParagraph(cell.ownerDocument!, ''));
}

function blankCell(document: Document): Element {
  const cell = wordElement(document, 'tc');
  cell.appendChild(newParagraph(document, ''));
  return cell;
}

function gridCol(document: Document, width = 2250): Element {
  const column = wordElement(document, 'gridCol');
  column.setAttributeNS(WORD_NS, 'w:w', String(width));
  return column;
}

function widthValue(parent: Element, name: string, value: { type: 'auto' | 'dxa' | 'pct'; value: number } | undefined): void {
  removeWordChildren(parent, name);
  if (!value) return;
  const width = property(parent, name);
  width.setAttributeNS(WORD_NS, 'w:type', value.type);
  width.setAttributeNS(WORD_NS, 'w:w', String(value.value));
}

function boolValue(parent: Element, name: string, value: boolean | undefined): void {
  removeWordChildren(parent, name);
  if (value === undefined) return;
  const element = property(parent, name);
  if (!value) setWordValue(element, '0');
}

function valueElement(parent: Element, name: string, value: string | undefined): void {
  removeWordChildren(parent, name);
  if (value === undefined) return;
  const element = property(parent, name);
  setWordValue(element, value);
}

function mergeElement(parent: Element, name: 'gridSpan' | 'vMerge' | 'hMerge', value: number | 'restart' | 'continue' | undefined): void {
  removeWordChildren(parent, name);
  if (value === undefined) return;
  const element = property(parent, name);
  if (typeof value === 'number') setWordValue(element, String(value));
  else if (value === 'restart') setWordValue(element, value);
}

function setBorders(parent: Element, name: 'tblBorders' | 'tcBorders', borders: TableFormat['borders'] | CellFormat['borders'] | undefined): void {
  removeWordChildren(parent, name);
  if (!borders) return;
  const element = property(parent, name);
  for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const) {
    const border = borders[side];
    if (!border) continue;
    const child = property(element, side);
    if (border.style) setWordValue(child, border.style);
    if (border.size !== undefined) child.setAttributeNS(WORD_NS, 'w:sz', String(border.size));
    if (border.space !== undefined) child.setAttributeNS(WORD_NS, 'w:space', String(border.space));
    if (border.color) child.setAttributeNS(WORD_NS, 'w:color', border.color);
    if (border.none && !border.style) setWordValue(child, 'nil');
  }
}

function setShading(parent: Element, shading: TableFormat['shading'] | CellFormat['shading'] | undefined): void {
  removeWordChildren(parent, 'shd');
  if (!shading) return;
  const element = property(parent, 'shd');
  if (shading.fill) element.setAttributeNS(WORD_NS, 'w:fill', shading.fill);
  if (shading.color) element.setAttributeNS(WORD_NS, 'w:color', shading.color);
  if (shading.value) setWordValue(element, shading.value);
}

function setMargins(parent: Element, name: 'tblCellMar' | 'tcMar', margin: TableFormat['cellMargin'] | CellFormat['margin'] | undefined): void {
  removeWordChildren(parent, name);
  if (!margin) return;
  const element = property(parent, name);
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const value = margin[side];
    if (!value) continue;
    const child = property(element, side);
    child.setAttributeNS(WORD_NS, 'w:type', value.type);
    child.setAttributeNS(WORD_NS, 'w:w', String(value.value));
  }
}

function setTableFormat(tbl: Element, format: TableFormat): void {
  const props = tableProperty(tbl, 'tblPr');
  if (format.width !== undefined) widthValue(props, 'tblW', format.width);
  if (format.alignment !== undefined) valueElement(props, 'jc', format.alignment);
  if (format.indent !== undefined) {
    removeWordChildren(props, 'tblInd');
    const ind = property(props, 'tblInd');
    ind.setAttributeNS(WORD_NS, 'w:w', String(format.indent));
    ind.setAttributeNS(WORD_NS, 'w:type', 'dxa');
  }
  if (format.borders !== undefined) setBorders(props, 'tblBorders', format.borders);
  if (format.shading !== undefined) setShading(props, format.shading);
  if (format.cellMargin !== undefined) setMargins(props, 'tblCellMar', format.cellMargin);
  if (format.layout !== undefined) valueElement(props, 'tblLayout', format.layout);
  if (format.style !== undefined) valueElement(props, 'tblStyle', format.style);
  if (format.look !== undefined) valueElement(props, 'tblLook', format.look);
  if (format.caption !== undefined) valueElement(props, 'tblCaption', format.caption);
  if (format.description !== undefined) valueElement(props, 'tblDescription', format.description);
}

function setRowFormat(row: Element, format: RowFormat): void {
  const props = tableProperty(row, 'trPr');
  if (format.height !== undefined) {
    removeWordChildren(props, 'trHeight');
    const height = property(props, 'trHeight');
    height.setAttributeNS(WORD_NS, 'w:val', String(format.height.value));
    if (format.height.rule) height.setAttributeNS(WORD_NS, 'w:hRule', format.height.rule);
  }
  if (format.cantSplit !== undefined) boolValue(props, 'cantSplit', format.cantSplit);
  if (format.header !== undefined) boolValue(props, 'tblHeader', format.header);
  if (format.alignment !== undefined) valueElement(props, 'jc', format.alignment);
  if (format.deleted !== undefined) boolValue(props, 'del', format.deleted);
  if (format.inserted !== undefined) boolValue(props, 'ins', format.inserted);
}

function setCellFormat(cell: Element, format: CellFormat): void {
  const props = tableProperty(cell, 'tcPr');
  if (format.width !== undefined) widthValue(props, 'tcW', format.width);
  if (format.borders !== undefined) setBorders(props, 'tcBorders', format.borders);
  if (format.shading !== undefined) setShading(props, format.shading);
  if (format.margin !== undefined) setMargins(props, 'tcMar', format.margin);
  if (format.verticalAlign !== undefined) valueElement(props, 'vAlign', format.verticalAlign);
  if (format.textDirection !== undefined) valueElement(props, 'textDirection', format.textDirection);
  if (format.noWrap !== undefined) boolValue(props, 'noWrap', format.noWrap);
  if (format.hideMark !== undefined) boolValue(props, 'hideMark', format.hideMark);
  if (format.hMerge !== undefined) mergeElement(props, 'hMerge', format.hMerge);
  if (format.vMerge !== undefined) mergeElement(props, 'vMerge', format.vMerge);
}

type XmlCellRef = { cell: Element; row: number; start: number; colSpan: number; rowSpan: number; isContinuation: boolean };

function tableModel(table: Element): { rows: Element[]; grid: number[]; matrix: XmlCellRef[][]; refs: XmlCellRef[] } {
  const rows = children(table, 'tr');
  const grid = tableGrid(table);
  const matrix: XmlCellRef[][] = Array.from({ length: rows.length }, () => []);
  const refs: XmlCellRef[] = [];
  const active = new Map<number, { end: number; master: XmlCellRef }>();
  rows.forEach((row, rowIndex) => {
    const nextActive = new Map<number, { end: number; master: XmlCellRef }>();
    for (const position of rowCells(row)) {
      const ref: XmlCellRef = {
        cell: position.cell,
        row: rowIndex,
        start: position.start,
        colSpan: position.span,
        rowSpan: 1,
        isContinuation: false,
      };
      const activeMerge = active.get(position.start);
      const sameMerge = activeMerge && activeMerge.end === position.start + position.span;
      if (position.vMerge === 'continue' && sameMerge) {
        activeMerge.master.rowSpan += 1;
        ref.rowSpan = 0;
        ref.isContinuation = true;
        nextActive.set(position.start, activeMerge);
        for (let col = position.start; col < position.start + position.span; col++) matrix[rowIndex]![col] = activeMerge.master;
      } else {
        refs.push(ref);
        if (position.vMerge === 'restart') nextActive.set(position.start, { end: position.start + position.span, master: ref });
        for (let col = position.start; col < position.start + position.span; col++) matrix[rowIndex]![col] = ref;
      }
    }
    active.clear();
    for (const [start, merge] of nextActive) active.set(start, merge);
  });
  return { rows, grid, matrix, refs };
}

function cellAt(table: Element, row: number, col: number): XmlCellRef {
  assertIndex(row); assertIndex(col);
  const model = tableModel(table);
  const cell = model.matrix[row]?.[col];
  if (!cell) throw new Error(`Cell ${row},${col} does not exist.`);
  return cell;
}

function clearCellContent(cell: Element): void {
  for (let child = cell.firstChild; child;) {
    const next = child.nextSibling;
    if (!(child.nodeType === 1 && (child as Element).namespaceURI === WORD_NS && (child as Element).localName === 'tcPr')) {
      cell.removeChild(child);
    }
    child = next;
  }
  ensureCellParagraph(cell);
}

function appendCellContent(target: Element, source: Element): void {
  for (const child of children(source).filter(node => node.localName !== 'tcPr')) target.appendChild(child);
  ensureCellParagraph(target);
}

function buildTable(document: Document, rows: number, cols: number, format?: TableFormat, texts?: string[][]): Element {
  const table = wordElement(document, 'tbl');
  const grid = wordElement(document, 'tblGrid');
  for (let i = 0; i < cols; i++) grid.appendChild(gridCol(document, Math.floor(9000 / cols)));
  const tableProps = format ? wordElement(document, 'tblPr') : null;
  if (tableProps) table.appendChild(tableProps);
  table.appendChild(grid);
  if (format) setTableFormat(table, format);
  for (let rowIndex = 0; rowIndex < rows; rowIndex++) {
    const tr = wordElement(document, 'tr');
    for (let colIndex = 0; colIndex < cols; colIndex++) {
      const cell = blankCell(document);
      const paragraph = children(cell, 'p')[0]!;
      replaceSpan(paragraph, 0, 0, texts?.[rowIndex]?.[colIndex] ?? '');
      tr.appendChild(cell);
    }
    table.appendChild(tr);
  }
  return table;
}

function repairVerticalMerges(table: Element): void {
  const active = new Map<number, { end: number }>();
  for (const row of children(table, 'tr')) {
    const nextActive = new Map<number, { end: number }>();
    for (const position of rowCells(row)) {
      const props = tableProperty(position.cell, 'tcPr');
      const merge = children(props, 'vMerge')[0];
      const continuing = merge && (wordValue(merge) ?? 'continue') === 'continue';
      const sameMerge = active.get(position.start)?.end === position.start + position.span;
      if (continuing && !sameMerge) {
        mergeElement(props, 'vMerge', 'restart');
        nextActive.set(position.start, { end: position.start + position.span });
      } else if (merge && (wordValue(merge) === 'restart' || (continuing && sameMerge))) {
        nextActive.set(position.start, { end: position.start + position.span });
      }
    }
    active.clear();
    for (const [start, merge] of nextActive) active.set(start, merge);
  }
}

async function readEntry(entry: JSZip.JSZipObject, limit: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let size = 0;
    // JSZip exposes internalStream publicly, but omits it from JSZipObject's typings.
    const stream = (entry as JSZip.JSZipObject & {
      internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
    }).internalStream('uint8array');
    stream.on('data', (chunk: Uint8Array) => {
      size += chunk.length;
      if (size > limit) {
        stream.pause();
        reject(new Error('DOCX exceeds the expanded package/part size limit.'));
        return;
      }
      chunks.push(chunk);
    });
    stream.on('error', reject);
    stream.on('end', () => {
      const result = new Uint8Array(size);
      let offset = 0;
      for (const chunk of chunks) { result.set(chunk, offset); offset += chunk.length; }
      resolve(result);
    });
    stream.resume();
  });
}

export class DocxDocument {
  private parts: Map<string, Uint8Array>;
  private mainPath: string;
  private currentRevision = 0;

  private constructor(parts: Map<string, Uint8Array>) {
    this.parts = parts;
    this.mainPath = this.validatePackage();
  }

  static create(): DocxDocument {
    const files: Record<string, string> = {
      '[Content_Types].xml': `<Types xmlns="${CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN_TYPE}"/></Types>`,
      '_rels/.rels': `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>`,
      'word/document.xml': `<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
    };
    return new DocxDocument(new Map(Object.entries(files).map(([path, xml]) => [path, encoder.encode(xml)])));
  }

  static async load(input: Uint8Array | ArrayBuffer | Blob): Promise<DocxDocument> {
    const size = input instanceof Blob ? input.size : input.byteLength;
    if (size > MAX_ARCHIVE) throw new Error('DOCX archive exceeds 50 MiB.');
    const bytes = input instanceof Blob ? await input.arrayBuffer() : input;
    const zip = await JSZip.loadAsync(bytes, { createFolders: false });
    const entries = Object.values(zip.files) as JSZip.JSZipObject[];
    if (entries.length > MAX_PARTS) throw new Error('DOCX contains too many ZIP entries.');
    const parts = new Map<string, Uint8Array>();
    let total = 0;
    for (const entry of entries) {
      if (entry.dir) continue;
      validatePath(entry.name);
      if (entry.unsafeOriginalName !== undefined && entry.unsafeOriginalName !== entry.name) {
        throw new Error('Unsafe ZIP entry path.');
      }
      const data = await readEntry(entry, Math.min(MAX_PART, MAX_TOTAL - total));
      total += data.byteLength;
      parts.set(entry.name, data);
    }
    return new DocxDocument(parts);
  }

  get revision(): number { return this.currentRevision; }
  get mainDocumentPath(): string { return this.mainPath; }

  listParts(): string[] { return [...this.parts.keys()].sort(); }

  getPartBytes(path: string): Uint8Array {
    validatePath(path);
    const bytes = this.parts.get(path);
    if (!bytes) throw new Error(`Package part not found: ${path}`);
    return Uint8Array.from(bytes);
  }

  getPartXml(path: string): string { return decodeXml(this.getPartBytes(path)); }

  /** Returns a detached DOM; use updatePartXml to persist changes. */
  getPartDocument(path: string): Document { return parseXml(this.getPartXml(path)); }

  updatePartXml(path: string, update: (document: Document) => void): void {
    const document = this.getPartDocument(path);
    update(document);
    this.setPartXml(path, serializeXml(document));
  }

  setPartXml(path: string, xml: string): void {
    if (typeof xml !== 'string' || xml.length > MAX_PART) throw new Error('XML part exceeds size limit.');
    parseXml(xml);
    this.setPartBytes(path, encodeXml(xml));
  }

  /** Replaces an existing part. Relationships/content types remain under caller control. */
  setPartBytes(path: string, bytes: Uint8Array): void {
    validatePath(path);
    if (!this.parts.has(path)) throw new Error('Use addPart with a content type to create a new part.');
    const next = new Map(this.parts);
    next.set(path, Uint8Array.from(bytes));
    this.commitParts(next);
  }

  addPart(path: string, bytes: Uint8Array, contentType: string): void {
    validatePath(path);
    assertText(contentType, 'contentType');
    if (!contentType || this.parts.has(path)) throw new Error('New part requires a unique path and content type.');
    const types = this.getPartDocument('[Content_Types].xml');
    const override = types.createElementNS(CONTENT_TYPES_NS, 'Override');
    override.setAttribute('PartName', `/${path}`);
    override.setAttribute('ContentType', contentType);
    types.documentElement!.appendChild(override);
    const next = new Map(this.parts);
    next.set('[Content_Types].xml', encodeXml(serializeXml(types)));
    next.set(path, Uint8Array.from(bytes));
    this.commitParts(next);
  }

  private commitParts(parts: Map<string, Uint8Array>): void {
    const draft = new DocxDocument(parts);
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
  }

  private validatePackage(): string {
    if (this.parts.size > MAX_PARTS) throw new Error('Too many package parts.');
    let total = 0;
    for (const [path, bytes] of this.parts) {
      validatePath(path);
      total += bytes.byteLength;
      if (bytes.byteLength > MAX_PART || total > MAX_TOTAL) throw new Error('Package exceeds size limits.');
    }
    const rels = this.getPartDocument('_rels/.rels').documentElement;
    if (rels?.namespaceURI !== REL_NS || rels.localName !== 'Relationships') {
      throw new Error('Invalid root package relationships.');
    }
    const main = children(rels, 'Relationship', REL_NS).filter(rel => rel.getAttribute('Type') === OFFICE_DOCUMENT_REL);
    if (main.length !== 1 || main[0]!.getAttribute('TargetMode') === 'External') {
      throw new Error('Package requires one internal officeDocument relationship.');
    }
    const path = decodeURIComponent(main[0]!.getAttribute('Target') ?? '').replace(/^\//, '');
    validatePath(path);
    const types = this.getPartDocument('[Content_Types].xml').documentElement;
    if (types?.namespaceURI !== CONTENT_TYPES_NS || types.localName !== 'Types') {
      throw new Error('Invalid content types part.');
    }
    const overrides = children(types, 'Override', CONTENT_TYPES_NS);
    const type = overrides.find(type => type.getAttribute('PartName') === `/${path}`)?.getAttribute('ContentType')
      ?? children(types, 'Default', CONTENT_TYPES_NS)
        .find(type => type.getAttribute('Extension') === path.split('.').pop())?.getAttribute('ContentType');
    if (type !== MAIN_TYPE) throw new Error('Package is not a supported .docx document (macros/strict OOXML are not supported).');
    bodyOf(this.getPartDocument(path));
    return path;
  }

  getParagraphs(): ParagraphInfo[] {
    return descendants(bodyOf(this.getPartDocument(this.mainPath)), 'p').map(readParagraph);
  }

  getBlocks(): DocumentBlock[] {
    const body = bodyOf(this.getPartDocument(this.mainPath));
    const indices = new Map(descendants(body, 'p').map((p, i) => [p, i]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, indices.get(child)!) }];
      if (child.localName === 'tbl') return [readTable(child, walk)];
      if (['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '')) return walk(child);
      return [];
    });
    return walk(body);
  }

  getSnapshot(): DocumentSnapshot {
    return { revision: this.revision, paragraphs: this.getParagraphs(), blocks: this.getBlocks(), parts: this.listParts() };
  }

  setParagraphText(index: number, text: string): void {
    assertText(text);
    const normalized = text.replace(/\r\n?/g, '\n');
    this.updatePartXml(this.mainPath, document => {
      const paragraph = paragraphAt(document, index);
      const old = textOf(paragraph);
      let start = 0;
      while (start < old.length && start < normalized.length && old[start] === normalized[start]) start++;
      let end = old.length;
      let replacementEnd = normalized.length;
      while (end > start && replacementEnd > start && old[end - 1] === normalized[replacementEnd - 1]) {
        end--; replacementEnd--;
      }
      // Never cut a surrogate pair at a shared UTF-16 prefix/suffix.
      if (start > 0 && /[\ud800-\udbff]/.test(old[start - 1]!)) start--;
      if (end < old.length && /[\udc00-\udfff]/.test(old[end]!)) { end++; replacementEnd++; }
      if (old !== normalized) replaceSpan(paragraph, start, end, normalized.slice(start, replacementEnd));
    });
  }

  insertParagraph(text: string, before?: number): void {
    assertText(text);
    this.updatePartXml(this.mainPath, document => {
      const paragraph = newParagraph(document, text);
      if (before !== undefined) {
        const target = paragraphAt(document, before);
        target.parentNode!.insertBefore(paragraph, target);
      } else {
        const body = bodyOf(document);
        body.insertBefore(paragraph, children(body, 'sectPr')[0] ?? null);
      }
    });
  }

  deleteParagraph(index: number): void {
    this.updatePartXml(this.mainPath, document => {
      const paragraph = paragraphAt(document, index);
      const parent = paragraph.parentNode as Element;
      // A cell must end with a paragraph, and section properties must not be silently lost.
      if (children(paragraph, 'pPr').some(props => children(props, 'sectPr').length)) {
        throw new Error('Cannot delete a section-break paragraph; edit its XML explicitly.');
      }
      parent.removeChild(paragraph);
      const last = children(parent).filter(child => child.localName !== 'sectPr').at(-1);
      if ((parent.localName === 'tc' && last?.localName !== 'p') ||
          (parent.localName === 'body' && !children(parent, 'p').length)) {
        parent.insertBefore(newParagraph(document, ''), children(parent, 'sectPr')[0] ?? null);
      }
    });
  }

  formatParagraph(index: number, format: ParagraphFormat): void {
    validateParagraphFormat(format);
    this.updatePartXml(this.mainPath, document => {
      const props = properties(paragraphAt(document, index), 'pPr');
      if (format.style !== undefined) setWordValue(property(props, 'pStyle'), format.style);
      if (format.alignment !== undefined) setWordValue(property(props, 'jc'), format.alignment);
    });
  }

  formatRun(paragraph: number, run: number, format: RunFormat): void {
    assertIndex(run);
    validateRunFormat(format);
    this.updatePartXml(this.mainPath, document => {
      const element = ownRuns(paragraphAt(document, paragraph))[run];
      if (!element) throw new Error(`Run ${run} does not exist.`);
      const props = properties(element, 'rPr');
      for (const [key, tag] of [['bold', 'b'], ['italic', 'i'], ['underline', 'u']] as const) {
        if (format[key] !== undefined) {
          setWordValue(property(props, tag), key === 'underline' ? (format[key] ? 'single' : 'none') : (format[key] ? '1' : '0'));
        }
      }
      if (format.fontSize !== undefined) {
        setWordValue(property(props, 'sz'), String(format.fontSize * 2));
        setWordValue(property(props, 'szCs'), String(format.fontSize * 2));
      }
      if (format.color !== undefined) setWordValue(property(props, 'color'), format.color);
      if (format.fontFamily !== undefined) {
        const fonts = property(props, 'rFonts');
        for (const name of ['ascii', 'hAnsi', 'eastAsia', 'cs']) fonts.setAttributeNS(WORD_NS, `w:${name}`, format.fontFamily);
      }
    });
  }

  replaceText(search: string, replacement: string): void {
    assertText(search, 'search'); assertText(replacement, 'replacement');
    if (!search) throw new Error('search must not be empty.');
    this.updatePartXml(this.mainPath, document => {
      for (const paragraph of descendants(bodyOf(document), 'p')) {
        const text = textOf(paragraph);
        const matches: number[] = [];
        for (let i = text.indexOf(search); i !== -1; i = text.indexOf(search, i + search.length)) matches.push(i);
        for (const start of matches.reverse()) replaceSpan(paragraph, start, start + search.length, replacement);
      }
    });
  }

  insertTable(rows: string[][]): void {
    validateRows(rows);
    this.updatePartXml(this.mainPath, document => {
      const body = bodyOf(document);
      const section = children(body, 'sectPr')[0] ?? null;
      const table = buildTable(document, rows.length, Math.max(...rows.map(row => row.length)), undefined, rows);
      body.insertBefore(table, section);
      body.insertBefore(newParagraph(document, ''), section);
    });
  }

  insertTableAt(rows: number, cols: number, before?: number, format?: TableFormat): void {
    assertIndex(rows); assertIndex(cols);
    if (rows < 1 || cols < 1) throw new Error('Table must contain at least one row and one column.');
    this.updatePartXml(this.mainPath, document => {
      const table = buildTable(document, rows, cols, format);
      const body = bodyOf(document);
      const section = children(body, 'sectPr')[0] ?? null;
      if (before !== undefined) {
        const target = bodyBlockAt(document, before);
        const previous = target.previousSibling?.nodeType === 1 ? target.previousSibling as Element : null;
        if (target.localName === 'tbl' && previous?.namespaceURI === WORD_NS && previous.localName === 'p' && textOf(previous) === '') {
          body.insertBefore(table, previous);
        } else {
          body.insertBefore(table, target);
          if (target.localName !== 'p') body.insertBefore(newParagraph(document, ''), target);
        }
      } else {
        body.insertBefore(table, section);
        body.insertBefore(newParagraph(document, ''), section);
      }
    });
  }

  getTable(index: number): TableInfo {
    const document = this.getPartDocument(this.mainPath);
    const body = bodyOf(document);
    const indices = new Map(descendants(body, 'p').map((p, i) => [p, i]));
    const walk = (parent: Element): DocumentBlock[] => children(parent).flatMap((child): DocumentBlock[] => {
      if (child.localName === 'p') return [{ type: 'paragraph', paragraph: readParagraph(child, indices.get(child)!) }];
      if (child.localName === 'tbl') return [readTable(child, walk)];
      if (['sdt', 'sdtContent', 'customXml'].includes(child.localName ?? '')) return walk(child);
      return [];
    });
    const table = readTable(tableAt(document, index), walk);
    return { index, rows: table.rows, format: table.format, grid: table.grid };
  }

  insertTableRow(table: number, at: number): void {
    assertIndex(table);
    assertIndex(at);
    this.updatePartXml(this.mainPath, document => {
      const element = tableAt(document, table);
      const rows = children(element, 'tr');
      if (at > rows.length) throw new Error(`Row ${at} does not exist.`);
      const tr = wordElement(document, 'tr');
      for (let i = 0; i < tableGrid(element).length; i++) tr.appendChild(blankCell(document));
      element.insertBefore(tr, rows[at] ?? null);
    });
  }

  deleteTableRow(table: number, at: number): void {
    this.updatePartXml(this.mainPath, document => {
      const element = tableAt(document, table);
      const rows = children(element, 'tr');
      const row = rows[at];
      if (!row) throw new Error(`Row ${at} does not exist.`);
      if (rows.length <= 1) throw new Error('Cannot delete the only table row.');
      const next = rows[at + 1];
      if (next) {
        const nextPositions = new Map(rowCells(next).map(position => [position.start, position]));
        for (const position of rowCells(row)) {
          if (position.vMerge === 'restart') {
            const continuation = nextPositions.get(position.start);
            if (continuation?.vMerge === 'continue') {
              const props = tableProperty(continuation.cell, 'tcPr');
              mergeElement(props, 'vMerge', 'restart');
            }
          }
        }
      }
      element.removeChild(row);
      repairVerticalMerges(element);
    });
  }

  insertTableColumn(table: number, at: number): void {
    assertIndex(table);
    assertIndex(at);
    this.updatePartXml(this.mainPath, document => {
      const element = tableAt(document, table);
      const grid = ensureTableGrid(element);
      const widths = tableGrid(element);
      if (at > widths.length) throw new Error(`Column ${at} does not exist.`);
      grid.insertBefore(gridCol(document, widths[Math.max(0, Math.min(at, widths.length - 1))] ?? 2250), children(grid, 'gridCol')[at] ?? null);
      const model = tableModel(element);
      const handled = new Set<Element>();
      for (const [rowIndex, row] of children(element, 'tr').entries()) {
        const covering = at < widths.length ? model.matrix[rowIndex]?.[at] : undefined;
        if (covering && covering.rowSpan > 1) {
          if (handled.has(covering.cell)) continue;
          handled.add(covering.cell);
          for (let index = covering.row; index < covering.row + covering.rowSpan; index++) {
            const chain = rowCells(model.rows[index]!).find(position => position.start === covering.start);
            if (!chain) continue;
            const props = tableProperty(chain.cell, 'tcPr');
            mergeElement(props, 'gridSpan', chain.span + 1 > 1 ? chain.span + 1 : undefined);
          }
          continue;
        }
        const positions = rowCells(row);
        let inserted = false;
        for (const position of positions) {
          if (at > position.start && at < position.start + position.span) {
            const props = tableProperty(position.cell, 'tcPr');
            mergeElement(props, 'gridSpan', position.span + 1 > 1 ? position.span + 1 : undefined);
            inserted = true;
            break;
          }
          if (at === position.start) {
            row.insertBefore(blankCell(document), position.cell);
            inserted = true;
            break;
          }
        }
        if (!inserted) row.appendChild(blankCell(document));
      }
    });
  }

  deleteTableColumn(table: number, at: number): void {
    this.updatePartXml(this.mainPath, document => {
      const element = tableAt(document, table);
      const grid = ensureTableGrid(element);
      const columns = children(grid, 'gridCol');
      if (!columns[at]) throw new Error(`Column ${at} does not exist.`);
      if (columns.length <= 1) throw new Error('Cannot delete the only table column.');
      const model = tableModel(element);
      grid.removeChild(columns[at]!);
      const handled = new Set<Element>();
      for (const [rowIndex, row] of children(element, 'tr').entries()) {
        const covering = model.matrix[rowIndex]?.[at];
        if (covering && covering.rowSpan > 1) {
          if (handled.has(covering.cell)) continue;
          handled.add(covering.cell);
          for (let index = covering.row; index < covering.row + covering.rowSpan; index++) {
            const chain = rowCells(model.rows[index]!).find(position => position.start === covering.start);
            if (!chain) continue;
            if (chain.span > 1) {
              const props = tableProperty(chain.cell, 'tcPr');
              mergeElement(props, 'gridSpan', chain.span - 1 > 1 ? chain.span - 1 : undefined);
            } else {
              model.rows[index]!.removeChild(chain.cell);
            }
          }
          continue;
        }
        const position = rowCells(row).find(cell => at >= cell.start && at < cell.start + cell.span);
        if (!position) throw new Error(`Column ${at} does not exist.`);
        if (position.span > 1) {
          const props = tableProperty(position.cell, 'tcPr');
          mergeElement(props, 'gridSpan', position.span - 1 > 1 ? position.span - 1 : undefined);
        } else {
          row.removeChild(position.cell);
        }
      }
    });
  }

  mergeCells(table: number, range: { row: number; col: number; rowSpan: number; colSpan: number }): void {
    assertIndex(range.row); assertIndex(range.col); assertIndex(range.rowSpan); assertIndex(range.colSpan);
    if (range.rowSpan < 1 || range.colSpan < 1) throw new Error('merge range must be at least 1 × 1.');
    this.updatePartXml(this.mainPath, document => {
      const element = tableAt(document, table);
      const model = tableModel(element);
      const master = model.matrix[range.row]?.[range.col];
      if (!master || master.row !== range.row || master.start !== range.col) throw new Error('mergeCells must start at a visible top-left cell.');
      if (range.row + range.rowSpan > model.rows.length || range.col + range.colSpan > model.grid.length) {
        throw new Error('merge range exceeds table bounds.');
      }
      for (let rowIndex = range.row; rowIndex < range.row + range.rowSpan; rowIndex++) {
        const positions = rowCells(model.rows[rowIndex]!);
        const covered = positions.filter(position => position.start >= range.col && position.start + position.span <= range.col + range.colSpan);
        if (!covered.length || covered[0]!.start !== range.col) throw new Error('mergeCells requires a rectangular, grid-aligned selection.');
        for (const position of positions) {
          const intersects = position.start < range.col + range.colSpan && position.start + position.span > range.col;
          const contained = position.start >= range.col && position.start + position.span <= range.col + range.colSpan;
          if (intersects && !contained) throw new Error('mergeCells cannot partially cover existing merged cells.');
        }
        if (rowIndex === range.row) {
          for (const position of covered) {
            if (position.cell === master.cell) continue;
            appendCellContent(master.cell, position.cell);
            model.rows[rowIndex]!.removeChild(position.cell);
          }
        } else {
          const continuation = covered[0]!;
          appendCellContent(master.cell, continuation.cell);
          for (const position of covered.slice(1)) {
            appendCellContent(master.cell, position.cell);
            model.rows[rowIndex]!.removeChild(position.cell);
          }
          clearCellContent(continuation.cell);
          const props = tableProperty(continuation.cell, 'tcPr');
          mergeElement(props, 'gridSpan', range.colSpan > 1 ? range.colSpan : undefined);
          mergeElement(props, 'vMerge', 'continue');
        }
      }
      const masterProps = tableProperty(master.cell, 'tcPr');
      mergeElement(masterProps, 'gridSpan', range.colSpan > 1 ? range.colSpan : undefined);
      mergeElement(masterProps, 'vMerge', range.rowSpan > 1 ? 'restart' : undefined);
      ensureCellParagraph(master.cell);
    });
  }

  splitCell(table: number, row: number, col: number, rows: number, cols: number): void {
    assertIndex(row); assertIndex(col); assertIndex(rows); assertIndex(cols);
    this.updatePartXml(this.mainPath, document => {
      const element = tableAt(document, table);
      const model = tableModel(element);
      const master = model.matrix[row]?.[col];
      if (!master || master.row !== row || master.start !== col) throw new Error('splitCell must target a visible top-left cell.');
      if (rows !== master.rowSpan || cols !== master.colSpan) {
        throw new Error('splitCell currently supports restoring a merged cell to its original grid span.');
      }
      const masterProps = tableProperty(master.cell, 'tcPr');
      mergeElement(masterProps, 'gridSpan', undefined);
      mergeElement(masterProps, 'vMerge', undefined);
      for (let rowIndex = row; rowIndex < row + rows; rowIndex++) {
        const rowElement = model.rows[rowIndex]!;
        const position = rowCells(rowElement).find(item => item.start === col);
        const cell = rowIndex === row ? master.cell : position?.cell;
        if (!cell) throw new Error('splitCell found an invalid merged-cell structure.');
        const anchor = rowCells(rowElement).find(item => item.start >= col + cols)?.cell ?? null;
        const props = tableProperty(cell, 'tcPr');
        mergeElement(props, 'gridSpan', undefined);
        mergeElement(props, 'vMerge', undefined);
        if (rowIndex !== row) clearCellContent(cell);
        const existingCells = rowIndex === row ? 1 : 1;
        for (let i = existingCells; i < cols; i++) {
          const extra = blankCell(document);
          rowElement.insertBefore(extra, anchor);
        }
      }
    });
  }

  formatTable(table: number, format: TableFormat): void {
    this.updatePartXml(this.mainPath, document => setTableFormat(tableAt(document, table), format));
  }

  formatTableRow(table: number, row: number, format: RowFormat): void {
    this.updatePartXml(this.mainPath, document => {
      const element = children(tableAt(document, table), 'tr')[row];
      if (!element) throw new Error(`Row ${row} does not exist.`);
      setRowFormat(element, format);
    });
  }

  formatCell(table: number, row: number, col: number, format: CellFormat): void {
    this.updatePartXml(this.mainPath, document => setCellFormat(cellAt(tableAt(document, table), row, col).cell, format));
  }

  setCellText(table: number, row: number, col: number, text: string): void {
    assertText(text);
    this.updatePartXml(this.mainPath, document => {
      const cell = cellAt(tableAt(document, table), row, col).cell;
      ensureCellParagraph(cell);
      const paragraph = children(cell, 'p')[0];
      if (!paragraph) throw new Error('Cell paragraph does not exist.');
      const indices = new Map(descendants(bodyOf(document), 'p').map((item, index) => [item, index]));
      const index = indices.get(paragraph);
      if (index === undefined) throw new Error('Cell paragraph index does not exist.');
      const old = textOf(paragraph);
      let start = 0;
      while (start < old.length && start < text.length && old[start] === text[start]) start++;
      let end = old.length;
      let replacementEnd = text.length;
      while (end > start && replacementEnd > start && old[end - 1] === text[replacementEnd - 1]) { end--; replacementEnd--; }
      if (start > 0 && /[\ud800-\udbff]/.test(old[start - 1]!)) start--;
      if (start > 0 && /[\ud800-\udbff]/.test(text[start - 1]!)) start--;
      if (end < old.length && /[\udc00-\udfff]/.test(old[end]!)) { end++; replacementEnd++; }
      if (replacementEnd < text.length && /[\udc00-\udfff]/.test(text[replacementEnd]!)) replacementEnd++;
      if (old !== text) replaceSpan(paragraph, start, end, text.slice(start, replacementEnd));
      readParagraph(paragraph, index);
    });
  }

  /** All operations succeed together, or the original package/revision is unchanged. */
  applyOperations(request: AgentRequest): DocumentSnapshot {
    validateRequest(request);
    if (request.expectedRevision !== undefined && request.expectedRevision !== this.revision) {
      throw new Error(`Revision conflict: expected ${request.expectedRevision}, current ${this.revision}.`);
    }
    if (!request.operations.length) return this.getSnapshot();
    const draft = new DocxDocument(new Map(this.parts));
    for (const operation of request.operations) {
      switch (operation.type) {
        case 'setParagraphText': draft.setParagraphText(operation.index, operation.text); break;
        case 'insertParagraph': draft.insertParagraph(operation.text, operation.before); break;
        case 'deleteParagraph': draft.deleteParagraph(operation.index); break;
        case 'formatParagraph': draft.formatParagraph(operation.index, operation.format); break;
        case 'formatRun': draft.formatRun(operation.paragraph, operation.run, operation.format); break;
        case 'replaceText': draft.replaceText(operation.search, operation.replacement); break;
        case 'insertTable': draft.insertTable(operation.rows); break;
        case 'insertTableAt': draft.insertTableAt(operation.rows, operation.cols, operation.before, operation.format); break;
        case 'insertTableRow': draft.insertTableRow(operation.table, operation.at); break;
        case 'deleteTableRow': draft.deleteTableRow(operation.table, operation.at); break;
        case 'insertTableColumn': draft.insertTableColumn(operation.table, operation.at); break;
        case 'deleteTableColumn': draft.deleteTableColumn(operation.table, operation.at); break;
        case 'mergeCells': draft.mergeCells(operation.table, operation.range); break;
        case 'splitCell': draft.splitCell(operation.table, operation.row, operation.col, operation.rows, operation.cols); break;
        case 'formatTable': draft.formatTable(operation.table, operation.format); break;
        case 'formatTableRow': draft.formatTableRow(operation.table, operation.row, operation.format); break;
        case 'formatCell': draft.formatCell(operation.table, operation.row, operation.col, operation.format); break;
        case 'setCellText': draft.setCellText(operation.table, operation.row, operation.col, operation.text); break;
        case 'setPartXml': draft.setPartXml(operation.path, operation.xml); break;
      }
    }
    const snapshot = draft.getSnapshot();
    this.parts = draft.parts;
    this.mainPath = draft.mainPath;
    this.currentRevision++;
    return { ...snapshot, revision: this.revision };
  }

  async toUint8Array(): Promise<Uint8Array> {
    const zip = new JSZip();
    for (const [path, bytes] of this.parts) zip.file(path, bytes, { createFolders: false });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async toBlob(): Promise<Blob> {
    const bytes = await this.toUint8Array();
    return new Blob([bytes as Uint8Array<ArrayBuffer>], { type: DOCX_TYPE });
  }
}
