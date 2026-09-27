import { DocxDocument } from './document.js';
import { pxToEmu } from './drawing.js';
import type { DocumentBlock, DocumentSnapshot, ImageInfo, ParagraphInfo } from './types.js';

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
  private selectedImageInfo: ImageInfo | null = null;
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
    this.root.addEventListener('keydown', this.handleRootKeydown);
    this.render();
  }

  get selectedParagraph(): number | null {
    return this.selected;
  }

  get selectedImage(): ImageInfo | null {
    return this.selectedImageInfo;
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
    this.selectedImageInfo = null;
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
    this.selectedImageInfo = null;
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
    this.root.removeEventListener('keydown', this.handleRootKeydown);
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
    if (paragraph.alignment) element.style.textAlign = paragraph.alignment === 'both' ? 'justify' : paragraph.alignment;
    if (paragraph.style) element.dataset.style = paragraph.style;
    for (const run of paragraph.runs) {
      if (run.text) {
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
      if (run.image) element.append(this.makeImage(paragraph.index, run.image));
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

  private makeImage(paragraph: number, image: ImageInfo): HTMLElement {
    const wrapper = this.root.ownerDocument.createElement(image.placement === 'floating' ? 'div' : 'span');
    wrapper.className = `docx-image${this.selectedImageInfo?.relationshipId === image.relationshipId ? ' selected' : ''}`;
    wrapper.contentEditable = 'false';
    wrapper.tabIndex = 0;
    wrapper.dataset.image = image.relationshipId;
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
    if (transforms.length) img.style.transform = `${img.style.transform ? `${img.style.transform} ` : ''}${transforms.join(' ')}`.trim();
    wrapper.append(img);
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

  private selectImage(image: ImageInfo | null): void {
    this.selectedImageInfo = image;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (EventClass) this.root.dispatchEvent(new EventClass('docx-imageselectionchange', { bubbles: true, detail: image ? { image } : null }));
  }

  private readonly handleSelection = (): void => {
    const selection = this.root.ownerDocument.getSelection();
    const node = selection?.anchorNode;
    if (!node || !this.root.contains(node)) return;
    const element = node.nodeType === 1 ? node as Element : node.parentElement;
    const paragraph = element?.closest<HTMLElement>('[data-paragraph]');
    if (paragraph && this.root.contains(paragraph)) this.selectParagraph(Number(paragraph.dataset.paragraph));
    const image = element?.closest<HTMLElement>('[data-image]');
    if (!image) this.selectImage(null);
  };

  private readonly handleRootKeydown = (event: KeyboardEvent): void => {
    if (!this.selectedImageInfo || !['Delete', 'Backspace'].includes(event.key)) return;
    const active = this.root.ownerDocument.activeElement as HTMLElement | null;
    const image = active?.closest('[data-image]') as HTMLElement | null;
    if (!image || image.dataset.image !== this.selectedImageInfo.relationshipId) return;
    event.preventDefault();
    this.document.deleteImage(this.selectedImageInfo);
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
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
