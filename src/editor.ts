import { DocxDocument } from './document.js';
import type { DocumentBlock, DocumentSnapshot, ParagraphInfo, RunInfo } from './types.js';

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
    element.className = 'docx-paragraph';
    element.contentEditable = 'true';
    element.spellcheck = false;
    element.dataset.paragraph = String(paragraph.index);
    element.setAttribute('role', 'textbox');
    element.setAttribute('aria-multiline', 'true');
    element.setAttribute('aria-label', `第 ${paragraph.index + 1} 段`);
    element.style.whiteSpace = 'pre-wrap';
    element.style.minHeight = '1.5em';
    applyParagraphStyle(element, paragraph);
    if (paragraph.style) element.dataset.style = paragraph.style;
    for (const run of paragraph.runs) {
      const span = this.root.ownerDocument.createElement('span');
      span.textContent = run.text;
      applyRunStyle(span, run);
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
