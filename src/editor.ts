import { DocxDocument } from './document.js';
import type { BorderFormat, BordersFormat, CellFormat, DocumentBlock, DocumentSnapshot, ParagraphInfo, RunInfo, TableFormat, TableRowInfo, WidthFormat } from './types.js';
import { eighthPointsToPx, normalizeColor, normalizeWidth, twipsToPx } from './table.js';

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

function applyParagraphStyle(element: HTMLElement, paragraph: ParagraphInfo): void {
  const effective = paragraph.effective ?? paragraph;
  if (effective.alignment) element.style.textAlign = ['both', 'distribute'].includes(effective.alignment) ? 'justify' : effective.alignment;
  if (effective.indentLeft !== undefined && effective.indentLeft !== null) element.style.marginLeft = twipsToPoints(effective.indentLeft)!;
  if (effective.indentRight !== undefined && effective.indentRight !== null) element.style.marginRight = twipsToPoints(effective.indentRight)!;
  if (effective.spacingBefore !== undefined && effective.spacingBefore !== null) element.style.marginTop = twipsToPoints(effective.spacingBefore)!;
  if (effective.spacingAfter !== undefined && effective.spacingAfter !== null) element.style.marginBottom = twipsToPoints(effective.spacingAfter)!;
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
}

function applyRunStyle(span: HTMLSpanElement, run: RunInfo): void {
  const effective = run.effective ?? run;
  if (effective.bold !== undefined) span.style.fontWeight = effective.bold ? '700' : '400';
  if (effective.italic !== undefined) span.style.fontStyle = effective.italic ? 'italic' : 'normal';
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
  if (effective.verticalAlign === 'subscript' || effective.verticalAlign === 'superscript') span.style.verticalAlign = effective.verticalAlign;
  if (effective.smallCaps || effective.allCaps) span.style.fontVariantCaps = effective.allCaps ? 'all-small-caps' : 'small-caps';
  if (effective.allCaps) span.style.textTransform = 'uppercase';
  if (effective.highlight && effective.highlight !== 'none') span.style.backgroundColor = highlightColor(effective.highlight);
  if (effective.characterSpacing !== undefined && effective.characterSpacing !== null) span.style.letterSpacing = `${effective.characterSpacing / 20}pt`;
}

export interface DocxEditorOptions {
  onChange?: (snapshot: DocumentSnapshot) => void;
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
  }>();
  private selected: number | null = null;
  private composing = false;
  private renderAfterComposition = false;
  private destroyed = false;

  constructor(container: HTMLElement, document: DocxDocument, options: DocxEditorOptions = {}) {
    this.document = document;
    this.options = options;
    this.root = container.ownerDocument.createElement('div');
    this.root.className = 'docx-editor';
    this.root.setAttribute('aria-label', '文档编辑区域');
    container.append(this.root);
    this.root.ownerDocument.addEventListener('selectionchange', this.handleSelection);
    this.render();
  }

  get selectedParagraph(): number | null {
    return this.selected;
  }

  /** Commit visible text before an external API operation or an export. */
  flush(): void {
    if (this.destroyed) return;
    let changed = false;
    for (const [index, entry] of this.paragraphs) {
      const text = this.readText(entry.content);
      if (text !== entry.text) {
        this.document.setParagraphText(index, text);
        entry.text = text;
        changed = true;
      }
    }
    if (changed) this.options.onChange?.(this.document.getSnapshot());
  }

  setDocument(document: DocxDocument): void {
    if (this.destroyed) return;
    this.flush();
    this.document = document;
    this.selected = null;
    this.composing = false;
    this.renderAfterComposition = false;
    this.paragraphs.clear();
    this.render();
  }

  render(): void {
    if (this.destroyed) return;
    if (this.composing) {
      this.renderAfterComposition = true;
      return;
    }
    const caret = this.captureCaret();
    this.flush();
    this.paragraphs.clear();
    const fragment = this.root.ownerDocument.createDocumentFragment();
    this.appendBlocks(fragment, this.document.getBlocks());
    this.root.replaceChildren(fragment);
    if (this.selected !== null && !this.paragraphs.has(this.selected)) this.selected = null;
    if (caret) this.restoreCaret(caret);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.flush();
    this.destroyed = true;
    this.root.ownerDocument.removeEventListener('selectionchange', this.handleSelection);
    this.root.remove();
    this.paragraphs.clear();
  }

  private readText(element: HTMLElement): string {
    // Native editing can introduce line-break elements (e.g. via mobile keyboards).
    if (!element.querySelector('br, div, p')) return element.textContent ?? '';
    const text = element.innerText.replace(/\r\n?/g, '\n');
    return text === '\n' && !element.textContent ? '' : text;
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

  private applyTableStyle(table: HTMLTableElement, format: TableFormat | undefined): void {
    if (!format) return;
    table.style.tableLayout = format.layout === 'fixed' ? 'fixed' : 'auto';
    const width = this.widthCss(format.width);
    if (width) table.style.width = width;
    if (format.alignment === 'center') table.style.marginInline = 'auto';
    if (format.alignment === 'right') { table.style.marginLeft = 'auto'; table.style.marginRight = '0'; }
    if (format.alignment === 'left') { table.style.marginLeft = '0'; table.style.marginRight = 'auto'; }
    if (format.indent !== undefined && (!format.alignment || format.alignment === 'left')) table.style.marginLeft = `${twipsToPx(format.indent)}px`;
    if (format.shading?.fill) table.style.backgroundColor = `#${format.shading.fill}`;
    if (format.caption) {
      const caption = table.createCaption();
      caption.textContent = format.caption;
    }
  }

  private applyRowStyle(tr: HTMLTableRowElement, row: TableRowInfo): void {
    if (!row.format) return;
    if (row.format.height) tr.style.height = `${twipsToPx(row.format.height.value)}px`;
    if (row.format.header) tr.dataset.header = 'true';
  }

  private applyCellStyle(td: HTMLTableCellElement, cell: CellFormat | undefined, table: TableFormat | undefined,
    row: number, col: number, rowSpan: number, colSpan: number, rowCount: number, colCount: number): void {
    td.style.borderTop = this.cellBorder('top', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderTop;
    td.style.borderRight = this.cellBorder('right', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderRight;
    td.style.borderBottom = this.cellBorder('bottom', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderBottom;
    td.style.borderLeft = this.cellBorder('left', table, cell, row, col, rowSpan, colSpan, rowCount, colCount) ?? td.style.borderLeft;
    if (cell?.shading?.fill) td.style.backgroundColor = `#${cell.shading.fill}`;
    if (cell?.verticalAlign) td.style.verticalAlign = cell.verticalAlign;
    if (cell?.width) td.style.width = this.widthCss(cell.width) ?? '';
    if (cell?.margin?.top) td.style.paddingTop = this.paddingCss(cell.margin.top) ?? '';
    if (cell?.margin?.right) td.style.paddingRight = this.paddingCss(cell.margin.right) ?? '';
    if (cell?.margin?.bottom) td.style.paddingBottom = this.paddingCss(cell.margin.bottom) ?? '';
    if (cell?.margin?.left) td.style.paddingLeft = this.paddingCss(cell.margin.left) ?? '';
    if (cell?.noWrap) td.style.whiteSpace = 'nowrap';
    if (cell?.textDirection?.toLowerCase().includes('tb') || cell?.textDirection?.toLowerCase().includes('bt')) td.style.writingMode = 'vertical-rl';
  }

  private appendBlocks(parent: Node, blocks: DocumentBlock[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        parent.appendChild(this.makeParagraph(block.paragraph));
      } else {
        const table = this.root.ownerDocument.createElement('table');
        table.className = 'docx-table';
        this.applyTableStyle(table, block.format);
        const body = table.createTBody();
        for (const [rowIndex, row] of block.rows.entries()) {
          const tr = body.insertRow();
          this.applyRowStyle(tr, row);
          let colIndex = 0;
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
            this.applyCellStyle(td, cell.format, block.format, rowIndex, logicalStart, Math.max(1, cell.rowSpan), Math.max(1, cell.colSpan), block.rows.length, block.grid.length);
            this.appendBlocks(td, cell.blocks);
          }
        }
        parent.appendChild(table);
        if (block.format?.description) {
          const description = this.root.ownerDocument.createElement('div');
          description.className = 'sr-only';
          description.id = `docx-table-desc-${Math.random().toString(36).slice(2)}`;
          description.textContent = block.format.description;
          table.setAttribute('aria-describedby', description.id);
          parent.appendChild(description);
        }
      }
    }
  }

  private makeParagraph(paragraph: ParagraphInfo): HTMLParagraphElement {
    const element = this.root.ownerDocument.createElement('p');
    const content = this.root.ownerDocument.createElement('span');
    element.className = 'docx-paragraph';
    element.dataset.paragraph = String(paragraph.index);
    element.style.whiteSpace = 'pre-wrap';
    element.style.minHeight = '1.5em';
    applyParagraphStyle(element, paragraph);
    if (paragraph.style) element.dataset.style = paragraph.style;
    if (paragraph.numbering) {
      const marker = this.root.ownerDocument.createElement('span');
      marker.className = 'docx-numbering';
      marker.contentEditable = 'false';
      marker.setAttribute('aria-hidden', 'true');
      marker.textContent = paragraph.numbering.text;
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
    content.contentEditable = 'true';
    content.spellcheck = false;
    content.setAttribute('role', 'textbox');
    content.setAttribute('aria-multiline', 'true');
    content.setAttribute('aria-label', `第 ${paragraph.index + 1} 段`);
    if (paragraph.numbering) content.setAttribute('aria-description', `列表项 ${paragraph.numbering.text}，级别 ${paragraph.numbering.level + 1}`);
    element.addEventListener('mousedown', (event) => {
      const target = event.target as Node | null;
      if (target && content.contains(target)) return;
      event.preventDefault();
      this.focusContent(content);
    });
    for (const run of paragraph.runs) {
      const span = this.root.ownerDocument.createElement('span');
      span.textContent = run.text;
      applyRunStyle(span, run);
      content.append(span);
    }
    if (!paragraph.runs.length) content.textContent = paragraph.text;
    element.append(content);
    this.paragraphs.set(paragraph.index, { element, content, text: this.readText(content) });
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
    content.addEventListener('paste', (event) => {
      event.preventDefault();
      this.insertText(content, event.clipboardData?.getData('text/plain') ?? '');
    });
    // Do not allow rich HTML or embedded objects from drag-and-drop either.
    content.addEventListener('drop', (event) => { event.preventDefault(); });
    content.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && !event.isComposing && !this.composing) {
        event.preventDefault();
        this.insertText(content, '\n');
      }
      if (event.key === 'Tab' && !event.isComposing && !this.composing && paragraph.numbering) {
        this.flush();
        const current = this.document.getParagraphs().find((item) => item.index === paragraph.index);
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
    content.addEventListener('beforeinput', (event) => {
      if (!event.isComposing && ['insertParagraph', 'insertLineBreak'].includes(event.inputType)) {
        event.preventDefault();
        this.insertText(content, '\n');
      }
      if (event.inputType.startsWith('format')) event.preventDefault();
    });
    return element;
  }

  private insertText(element: HTMLElement, text: string): void {
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount) return;
    const range = selection.getRangeAt(0);
    if (!element.contains(range.startContainer) || !element.contains(range.endContainer)) return;
    range.deleteContents();
    const node = this.root.ownerDocument.createTextNode(text.replace(/\r\n?/g, '\n'));
    range.insertNode(node);
    range.setStartAfter(node);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
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

  private selectParagraph(index: number): void {
    if (this.selected === index) return;
    this.selected = index;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (EventClass) this.root.dispatchEvent(new EventClass('docx-selectionchange', { bubbles: true, detail: { index } }));
  }

  private readonly handleSelection = (): void => {
    const selection = this.root.ownerDocument.getSelection();
    const node = selection?.anchorNode;
    if (!node || !this.root.contains(node)) return;
    const element = node.nodeType === 1 ? node as Element : node.parentElement;
    const paragraph = element?.closest<HTMLElement>('[data-paragraph]');
    if (paragraph && this.root.contains(paragraph)) this.selectParagraph(Number(paragraph.dataset.paragraph));
  };

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
