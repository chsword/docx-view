import { DocxDocument } from './document.js';
import type { DocumentBlock, DocumentSnapshot, ParagraphInfo } from './types.js';

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

  private appendBlocks(parent: Node, blocks: DocumentBlock[]): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        parent.appendChild(this.makeParagraph(block.paragraph));
      } else {
        const table = this.root.ownerDocument.createElement('table');
        table.className = 'docx-table';
        const body = table.createTBody();
        for (const row of block.rows) {
          const tr = body.insertRow();
          for (const cell of row.cells) {
            const td = tr.insertCell();
            this.appendBlocks(td, cell.blocks);
          }
        }
        parent.appendChild(table);
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
    if (paragraph.alignment) element.style.textAlign = paragraph.alignment === 'both' ? 'justify' : paragraph.alignment;
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
    content.setAttribute('aria-label', paragraph.numbering
      ? `第 ${paragraph.index + 1} 段，列表项 ${paragraph.numbering.text}，级别 ${paragraph.numbering.level + 1}`
      : `第 ${paragraph.index + 1} 段`);
    for (const run of paragraph.runs) {
      const span = this.root.ownerDocument.createElement('span');
      span.textContent = run.text;
      if (run.bold !== undefined) span.style.fontWeight = run.bold ? '700' : '400';
      if (run.italic !== undefined) span.style.fontStyle = run.italic ? 'italic' : 'normal';
      if (run.underline !== undefined) span.style.textDecoration = run.underline ? 'underline' : 'none';
      if (run.fontSize !== undefined) span.style.fontSize = `${run.fontSize}pt`;
      if (run.fontFamily) span.style.fontFamily = run.fontFamily;
      if (run.color && /^[0-9a-f]{6}$/i.test(run.color)) span.style.color = `#${run.color}`;
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
        event.preventDefault();
        this.flush();
        const current = this.document.getParagraphs().find((item) => item.index === paragraph.index);
        if (!current?.numbering) return;
        this.document.setParagraphLevel(paragraph.index, event.shiftKey ? -1 : 1);
        this.render();
        this.options.onChange?.(this.document.getSnapshot());
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
