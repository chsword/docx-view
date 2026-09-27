import { DocxDocument } from './document.js';
import type { BorderFormat, BordersFormat, CellFormat, DocumentBlock, DocumentSnapshot, ParagraphInfo, TableFormat, TableRowInfo, WidthFormat } from './types.js';
import { eighthPointsToPx, normalizeColor, normalizeWidth, twipsToPx } from './table.js';

export interface DocxEditorOptions {
  onChange?: (snapshot: DocumentSnapshot) => void;
}

/** A browser-only, editable view of the supported DOCX paragraph/run/table subset. */
export class DocxEditor {
  private document: DocxDocument;
  private readonly root: HTMLDivElement;
  private readonly options: DocxEditorOptions;
  private readonly paragraphs = new Map<number, { element: HTMLParagraphElement; text: string }>();
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
      const text = this.readText(entry.element);
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

  private borderCss(border: BorderFormat | undefined): string | undefined {
    if (!border) return undefined;
    if (border.none || ['nil', 'none'].includes(border.style ?? '')) return 'none';
    const width = border.size !== undefined ? `${Math.max(1, eighthPointsToPx(border.size))}px` : '1px';
    const color = border.color && /^[0-9a-f]{6}$/i.test(border.color) ? `#${border.color}` : '#dbe3ed';
    return `${width} solid ${color}`;
  }

  private cellBorder(side: 'top' | 'right' | 'bottom' | 'left', table: TableFormat | undefined, cell: CellFormat | undefined, row: number, col: number, rowCount: number, colCount: number): string | undefined {
    const explicit = cell?.borders?.[side];
    if (explicit) return this.borderCss(explicit);
    const borders = table?.borders;
    if (!borders) return undefined;
    if (side === 'top' && row > 0 && borders.insideH) return this.borderCss(borders.insideH);
    if (side === 'bottom' && row < rowCount - 1 && borders.insideH) return this.borderCss(borders.insideH);
    if (side === 'left' && col > 0 && borders.insideV) return this.borderCss(borders.insideV);
    if (side === 'right' && col < colCount - 1 && borders.insideV) return this.borderCss(borders.insideV);
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
    if (format.indent !== undefined) table.style.marginLeft = `${twipsToPx(format.indent)}px`;
    if (format.shading?.fill) table.style.backgroundColor = `#${format.shading.fill}`;
    if (format.caption) {
      const caption = table.createCaption();
      caption.textContent = format.caption;
    }
    if (format.description) table.setAttribute('aria-label', format.description);
  }

  private applyRowStyle(tr: HTMLTableRowElement, row: TableRowInfo): void {
    if (!row.format) return;
    if (row.format.height) tr.style.height = `${twipsToPx(row.format.height.value)}px`;
    if (row.format.header) tr.dataset.header = 'true';
  }

  private applyCellStyle(td: HTMLTableCellElement, cell: CellFormat | undefined, table: TableFormat | undefined, row: number, col: number, rowCount: number, colCount: number): void {
    td.style.borderTop = this.cellBorder('top', table, cell, row, col, rowCount, colCount) ?? td.style.borderTop;
    td.style.borderRight = this.cellBorder('right', table, cell, row, col, rowCount, colCount) ?? td.style.borderRight;
    td.style.borderBottom = this.cellBorder('bottom', table, cell, row, col, rowCount, colCount) ?? td.style.borderBottom;
    td.style.borderLeft = this.cellBorder('left', table, cell, row, col, rowCount, colCount) ?? td.style.borderLeft;
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
            this.applyCellStyle(td, cell.format, block.format, rowIndex, logicalStart, block.rows.length, block.grid.length);
            this.appendBlocks(td, cell.blocks);
          }
        }
        parent.appendChild(table);
      }
    }
  }

  private makeParagraph(paragraph: ParagraphInfo): HTMLParagraphElement {
    const element = this.root.ownerDocument.createElement('p');
    element.className = 'docx-paragraph';
    element.contentEditable = 'true';
    element.spellcheck = false;
    element.dataset.paragraph = String(paragraph.index);
    element.setAttribute('role', 'textbox');
    element.setAttribute('aria-multiline', 'true');
    element.setAttribute('aria-label', `第 ${paragraph.index + 1} 段`);
    element.style.whiteSpace = 'pre-wrap';
    element.style.minHeight = '1.5em';
    if (paragraph.alignment) element.style.textAlign = paragraph.alignment === 'both' ? 'justify' : paragraph.alignment;
    if (paragraph.style) element.dataset.style = paragraph.style;
    for (const run of paragraph.runs) {
      const span = this.root.ownerDocument.createElement('span');
      span.textContent = run.text;
      if (run.bold !== undefined) span.style.fontWeight = run.bold ? '700' : '400';
      if (run.italic !== undefined) span.style.fontStyle = run.italic ? 'italic' : 'normal';
      if (run.underline !== undefined) span.style.textDecoration = run.underline ? 'underline' : 'none';
      if (run.fontSize !== undefined) span.style.fontSize = `${run.fontSize}pt`;
      if (run.fontFamily) span.style.fontFamily = run.fontFamily;
      if (run.color && /^[0-9a-f]{6}$/i.test(run.color)) span.style.color = `#${run.color}`;
      element.append(span);
    }
    if (!paragraph.runs.length) element.textContent = paragraph.text;
    this.paragraphs.set(paragraph.index, { element, text: this.readText(element) });
    element.addEventListener('focus', () => this.selectParagraph(paragraph.index));
    element.addEventListener('blur', () => { if (!this.composing) this.flush(); });
    element.addEventListener('compositionstart', () => { this.composing = true; });
    element.addEventListener('compositionend', () => {
      this.composing = false;
      if (this.renderAfterComposition) {
        this.renderAfterComposition = false;
        this.render();
      } else if (this.root.ownerDocument.activeElement !== element) {
        this.flush();
      }
    });
    element.addEventListener('paste', (event) => {
      event.preventDefault();
      this.insertText(element, event.clipboardData?.getData('text/plain') ?? '');
    });
    // Do not allow rich HTML or embedded objects from drag-and-drop either.
    element.addEventListener('drop', (event) => { event.preventDefault(); });
    element.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && !event.isComposing && !this.composing) {
        event.preventDefault();
        this.insertText(element, '\n');
      }
      if (event.key === 'Tab') {
        event.preventDefault();
        this.moveToAdjacentCell(element, event.shiftKey ? -1 : 1);
      }
      const caret = this.caretIn(element);
      if (event.key === 'ArrowLeft' && caret?.start === 0 && caret.end === 0) {
        event.preventDefault();
        this.moveToAdjacentCell(element, -1);
      }
      if (event.key === 'ArrowRight' && caret && caret.start === caret.end && caret.end === this.readText(element).length) {
        event.preventDefault();
        this.moveToAdjacentCell(element, 1);
      }
      if (event.key === 'ArrowUp') {
        event.preventDefault();
        this.moveVerticalCell(element, -1);
      }
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        this.moveVerticalCell(element, 1);
      }
      if ((event.ctrlKey || event.metaKey) && ['b', 'i', 'u'].includes(event.key.toLowerCase())) {
        event.preventDefault();
      }
    });
    element.addEventListener('beforeinput', (event) => {
      if (!event.isComposing && ['insertParagraph', 'insertLineBreak'].includes(event.inputType)) {
        event.preventDefault();
        this.insertText(element, '\n');
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

  private moveToAdjacentCell(element: HTMLElement, delta: number): void {
    const cell = element.closest<HTMLTableCellElement>('td[data-table-cell="true"]');
    const table = cell?.closest('table');
    if (!cell || !table) return;
    const cells = Array.from(table.querySelectorAll<HTMLTableCellElement>('td[data-table-cell="true"]'));
    const index = cells.indexOf(cell);
    this.focusParagraphInCell(cells[index + delta] ?? null);
  }

  private moveVerticalCell(element: HTMLElement, delta: number): void {
    const cell = element.closest<HTMLTableCellElement>('td[data-table-cell="true"]');
    const table = cell?.closest('table');
    if (!cell || !table) return;
    const currentCol = Number(cell.dataset.gridStart ?? 0);
    const targetRow = delta < 0 ? Number(cell.dataset.rowStart ?? 0) - 1 : Number(cell.dataset.rowEnd ?? 0);
    const target = Array.from(table.querySelectorAll<HTMLTableCellElement>('td[data-table-cell="true"]')).find((candidate) => {
      const rowStart = Number(candidate.dataset.rowStart ?? -1);
      const rowEnd = Number(candidate.dataset.rowEnd ?? -1);
      const colStart = Number(candidate.dataset.gridStart ?? -1);
      const colEnd = Number(candidate.dataset.gridEnd ?? -1);
      return rowStart <= targetRow && rowEnd > targetRow && colStart <= currentCol && colEnd > currentCol;
    });
    this.focusParagraphInCell(target ?? null);
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
    const paragraph = this.paragraphs.get(this.selected)?.element;
    const range = selection.getRangeAt(0);
    if (!paragraph?.contains(range.startContainer) || !paragraph.contains(range.endContainer)) return null;
    const prefix = range.cloneRange();
    prefix.selectNodeContents(paragraph);
    prefix.setEnd(range.startContainer, range.startOffset);
    const start = prefix.toString().length;
    return { index: this.selected, start, end: start + range.toString().length };
  }

  private restoreCaret(caret: { index: number; start: number; end: number }): void {
    const paragraph = this.paragraphs.get(caret.index)?.element;
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
