import { DocxDocument } from './document.js';
import type { DocumentBlock, DocumentSnapshot, ParagraphInfo, RunInfo, TabStop } from './types.js';

export interface DocxEditorOptions {
  onChange?: (snapshot: DocumentSnapshot) => void;
  showFormattingMarks?: boolean;
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
  private readonly metrics: CanvasRenderingContext2D | null;

  constructor(container: HTMLElement, document: DocxDocument, options: DocxEditorOptions = {}) {
    this.document = document;
    this.options = options;
    this.root = container.ownerDocument.createElement('div');
    this.root.className = 'docx-editor';
    this.root.setAttribute('aria-label', '文档编辑区域');
    container.append(this.root);
    this.metrics = this.root.ownerDocument.createElement('canvas').getContext('2d');
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
    let defaultTabStopTwips = 720;
    try {
      defaultTabStopTwips = Math.max(1, Number(this.document.getSettings().defaultTabStop) || 720);
    } catch {
      defaultTabStopTwips = 720;
    }
    const fragment = this.root.ownerDocument.createDocumentFragment();
    this.appendBlocks(fragment, this.document.getBlocks(), defaultTabStopTwips);
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
    const marks = Array.from(element.querySelectorAll<HTMLElement>('[data-docx-mark]'));
    const locations = marks.map((mark) => ({ mark, parent: mark.parentNode, next: mark.nextSibling }));
    for (const mark of marks) {
      mark.parentNode?.removeChild(mark);
    }
    try {
      // Native editing can introduce line-break elements (e.g. via mobile keyboards).
      if (!element.querySelector('br, div, p')) return element.textContent ?? '';
      const text = element.innerText.replace(/\r\n?/g, '\n');
      return text === '\n' && !element.textContent ? '' : text;
    } finally {
      for (const { mark, parent, next } of locations) parent?.insertBefore(mark, next);
    }
  }

  private appendBlocks(parent: Node, blocks: DocumentBlock[], defaultTabStopTwips: number): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        parent.appendChild(this.makeParagraph(block.paragraph, defaultTabStopTwips));
      } else {
        const table = this.root.ownerDocument.createElement('table');
        table.className = 'docx-table';
        const body = table.createTBody();
        for (const row of block.rows) {
          const tr = body.insertRow();
          for (const cell of row.cells) {
            const td = tr.insertCell();
            this.appendBlocks(td, cell.blocks, defaultTabStopTwips);
          }
        }
        parent.appendChild(table);
      }
    }
  }

  private makeParagraph(paragraph: ParagraphInfo, defaultTabStopTwips: number): HTMLParagraphElement {
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
    if (paragraph.shading?.fill && paragraph.shading.fill !== 'auto') element.style.backgroundColor = `#${paragraph.shading.fill}`;
    if (paragraph.borders) {
      for (const [key, side] of Object.entries(paragraph.borders)) {
        if (!side || ['none', 'nil'].includes(side.style)) continue;
        const width = `${Math.min(24, Math.max(0, side.size)) / 8}pt`;
        const color = side.color === 'auto' ? '#000' : `#${side.color}`;
        const style = this.borderStyle(side.style);
        if (key === 'between') continue;
        if (key === 'bar') element.style.borderLeft = `${width} ${style} ${color}`;
        if (key === 'top') element.style.borderTop = `${width} ${style} ${color}`;
        if (key === 'left') element.style.borderLeft = `${width} ${style} ${color}`;
        if (key === 'right') element.style.borderRight = `${width} ${style} ${color}`;
        if (key === 'bottom') element.style.borderBottom = `${width} ${style} ${color}`;
      }
    }
    let currentLineOffsetPx = 0;
    for (const run of paragraph.runs) {
      currentLineOffsetPx = this.appendRun(element, paragraph, run, defaultTabStopTwips, currentLineOffsetPx);
    }
    if (!paragraph.runs.length) element.textContent = paragraph.text;
    if (this.options.showFormattingMarks) element.append(this.makeMark('¶', '段落标记'));
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

  private appendRun(
    paragraphElement: HTMLElement,
    paragraph: ParagraphInfo,
    run: RunInfo,
    defaultTabStopTwips: number,
    currentLineOffsetPx: number,
  ): number {
    const runSpan = this.root.ownerDocument.createElement('span');
    this.applyRunStyle(runSpan, run);
    const segments = run.text.split(/(\t|\n)/);
    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i]!;
      if (!segment) continue;
      if (segment === '\n') {
        runSpan.append(this.root.ownerDocument.createElement('br'));
        if (this.options.showFormattingMarks) runSpan.append(this.makeMark('↵', '换行符'));
        currentLineOffsetPx = 0;
        continue;
      }
      if (segment === '\t') {
        const nextText = segments.slice(i + 1).find((part) => part !== '\t' && part !== '\n') ?? '';
        const tab = this.makeTabSpan(paragraph, run, currentLineOffsetPx, nextText, defaultTabStopTwips);
        runSpan.append(tab);
        currentLineOffsetPx += Number.parseFloat(tab.style.width || '0');
        if (this.options.showFormattingMarks) runSpan.append(this.makeMark('→', '制表符'));
        continue;
      }
      runSpan.append(this.root.ownerDocument.createTextNode(segment));
      currentLineOffsetPx += this.measure(segment, run);
    }
    paragraphElement.append(runSpan);
    return currentLineOffsetPx;
  }

  private makeTabSpan(
    paragraph: ParagraphInfo,
    run: RunInfo,
    currentPx: number,
    following: string,
    defaultTabStopTwips: number,
  ): HTMLSpanElement {
    const span = this.root.ownerDocument.createElement('span');
    span.className = 'docx-tab';
    span.contentEditable = 'false';
    span.setAttribute('aria-hidden', 'true');
    span.textContent = '\t';
    const stop = this.nextTabStop(paragraph.tabs ?? [], currentPx);
    const defaultTab = defaultTabStopTwips * 96 / 1440;
    const target = stop ? Math.max(0, stop.position) * 96 / 1440 : (Math.floor(currentPx / defaultTab) + 1) * defaultTab;
    const nextWidth = this.measure(following, run);
    const decimalMatch = /[.,，．]/.exec(following);
    const decimalLeft = decimalMatch ? this.measure(following.slice(0, decimalMatch.index), run) : nextWidth;
    const alignment = stop ? stop.alignment : 'left';
    const rawWidth = alignment === 'center' ? target - currentPx - nextWidth / 2
      : alignment === 'right' ? target - currentPx - nextWidth
        : alignment === 'decimal' ? target - currentPx - decimalLeft : target - currentPx;
    const width = Math.max(0, rawWidth);
    span.style.display = 'inline-block';
    span.style.width = `${Math.max(0, width)}px`;
    if (alignment === 'bar') span.style.borderLeft = '1px solid currentColor';
    const leader = this.leader(stop?.leader);
    if (leader) {
      const leaderWidth = Math.max(8, width);
      span.style.width = `${leaderWidth}px`;
      const leaderVisual = this.makeMark(
        leader.repeat(Math.max(1, Math.floor(leaderWidth / Math.max(2, this.measure(leader, run))))),
        '制表位前导符',
      );
      leaderVisual.style.position = 'absolute';
      leaderVisual.style.inset = '0';
      leaderVisual.style.overflow = 'hidden';
      leaderVisual.style.whiteSpace = 'nowrap';
      leaderVisual.style.opacity = '0.7';
      span.style.position = 'relative';
      span.append(leaderVisual);
    }
    return span;
  }

  private nextTabStop(tabs: TabStop[], currentPx: number): TabStop | undefined {
    const currentTwips = currentPx * 1440 / 96;
    return [...tabs].filter((tab) => Number.isFinite(tab.position))
      .sort((a, b) => a.position - b.position)
      .find((tab) => tab.position > currentTwips);
  }

  private leader(value: string | undefined): string {
    switch (value) {
      case 'none': return '';
      case 'dot': return '.';
      case 'hyphen': return '-';
      case 'underscore': return '_';
      case 'heavy': return '━';
      case 'middleDot': return '·';
      default: return '';
    }
  }

  private makeMark(text: string, label: string): HTMLElement {
    const mark = this.root.ownerDocument.createElement('span');
    mark.textContent = text;
    mark.className = 'docx-mark';
    mark.contentEditable = 'false';
    mark.setAttribute('data-docx-mark', '1');
    mark.title = label;
    mark.setAttribute('aria-hidden', 'true');
    mark.style.userSelect = 'none';
    mark.style.pointerEvents = 'none';
    mark.style.opacity = '0.55';
    return mark;
  }

  private applyRunStyle(span: HTMLElement, run: RunInfo): void {
    if (run.bold !== undefined) span.style.fontWeight = run.bold ? '700' : '400';
    if (run.italic !== undefined) span.style.fontStyle = run.italic ? 'italic' : 'normal';
    if (run.underline !== undefined) span.style.textDecoration = run.underline ? 'underline' : 'none';
    if (run.fontSize !== undefined) span.style.fontSize = `${run.fontSize}pt`;
    if (run.fontFamily) span.style.fontFamily = run.fontFamily;
    if (run.color && /^[0-9a-f]{6}$/i.test(run.color)) span.style.color = `#${run.color}`;
    if (run.shading?.fill && run.shading.fill !== 'auto') span.style.backgroundColor = `#${run.shading.fill}`;
    if (run.border && !['none', 'nil'].includes(run.border.style)) {
      span.style.border = `${Math.max(0.5, run.border.size / 8)}pt ${this.borderStyle(run.border.style)} ${run.border.color === 'auto' ? '#000' : `#${run.border.color}`}`;
      span.style.padding = '0 0.05em';
    }
  }

  private borderStyle(style: string): string {
    switch (style) {
      case 'double': return 'double';
      case 'thick': return 'solid';
      case 'dashed': return 'dashed';
      case 'dotted': return 'dotted';
      case 'wave': return 'wavy';
      case 'none':
      case 'nil':
        return 'none';
      default:
        return 'solid';
    }
  }

  private measure(text: string, run: RunInfo): number {
    if (!this.metrics || !text) return 0;
    const fontStyle = run.italic ? 'italic' : 'normal';
    const fontWeight = run.bold ? '700' : '400';
    const fontSize = `${run.fontSize ?? 11}pt`;
    const fontFamily = run.fontFamily ?? 'Arial, sans-serif';
    this.metrics.font = `${fontStyle} ${fontWeight} ${fontSize} ${fontFamily}`;
    return this.metrics.measureText(text).width;
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
