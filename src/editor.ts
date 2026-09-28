import { DocxDocument } from './document.js';
import type {
  BorderFormat,
  BordersFormat,
  CellFormat,
  DocumentBlock,
  DocumentSnapshot,
  ImageInfo,
  ParagraphInfo,
  RunInfo,
  SectionInfo,
  TabStop,
  TableFormat,
  TableRowInfo,
  WidthFormat,
} from './types.js';
import { pxToEmu } from './drawing.js';
import { eighthPointsToPx, normalizeColor, normalizeWidth, twipsToPx } from './table.js';
import { sanitizeText, sanitizeTextWithInfo } from './xml.js';

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

function applyRunStyle(span: HTMLElement, run: RunInfo): void {
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
}

export interface DocxEditorOptions {
  onChange?: (snapshot: DocumentSnapshot) => void;
  onError?: (error: Error, context: { paragraph: number }) => void;
  showFormattingMarks?: boolean;
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
  private composing = false;
  private renderAfterComposition = false;
  private destroyed = false;
  private readonly metrics: CanvasRenderingContext2D | null;

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
    this.root = container.ownerDocument.createElement('div');
    this.root.className = 'docx-editor';
    this.root.setAttribute('aria-label', '文档编辑区域');
    container.append(this.root);
    this.metrics = this.root.ownerDocument.createElement('canvas').getContext('2d');
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
      const sanitized = sanitizeTextWithInfo(this.readText(entry.content));
      if (sanitized.text === entry.text) {
        entry.failed = false;
        continue;
      }
      if (sanitized.truncated && !entry.failed) {
        this.reportError(new Error(`Paragraph text was truncated at ${sanitized.truncatedAt ?? sanitized.text.length} characters.`), index);
      }
      try {
        this.document.setParagraphText(index, sanitized.text);
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
    const activeImageId = (this.root.ownerDocument.activeElement as HTMLElement | null)?.closest<HTMLElement>('[data-image]')?.dataset.image
      ?? this.selectedImageInfo?.id
      ?? null;
    this.flush();
    this.applyPageSetup();
    this.paragraphs.clear();
    const fragment = this.root.ownerDocument.createDocumentFragment();
    const canRenderHeaderFooter = typeof this.root.ownerDocument.createElement === 'function';
    if (canRenderHeaderFooter) fragment.append(this.makeHeaderFooter('header'));
    let defaultTabStopTwips = 720;
    try {
      defaultTabStopTwips = Math.max(1, Number(this.document.getSettings().defaultTabStop) || 720);
    } catch {
      defaultTabStopTwips = 720;
    }
    this.appendBlocks(fragment, this.document.getBlocks(), defaultTabStopTwips);
    if (canRenderHeaderFooter) fragment.append(this.makeHeaderFooter('footer'));
    this.root.replaceChildren(fragment);
    if (this.selected !== null && !this.paragraphs.has(this.selected)) this.selected = null;
    const nextSelected = activeImageId ? this.document.getImages().find((image) => image.id === activeImageId) ?? null : null;
    this.selectImage(nextSelected);
    if (activeImageId) {
      Array.from(this.root.querySelectorAll<HTMLElement>('[data-image]'))
        .find((node) => node.dataset.image === activeImageId)
        ?.focus({ preventScroll: true });
    }
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
    const walk = (node: Node): string => {
      if (node.nodeType === 3) return node.textContent ?? '';
      if (node.nodeType !== 1) return '';
      const current = node as HTMLElement;
      if (current.dataset.image || current.contentEditable === 'false') return '';
      if (current.tagName === 'BR') return '\n';
      const text = Array.from(current.childNodes).map(walk).join('');
      if (['DIV', 'P'].includes(current.tagName)) return text ? `${text}\n` : '';
      return text;
    };
    return walk(element).replace(/\n$/, '');
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

  private appendBlocks(parent: Node, blocks: DocumentBlock[], defaultTabStopTwips: number): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        parent.appendChild(this.makeParagraph(block.paragraph, defaultTabStopTwips));
      } else if (block.type === 'table') {
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
            this.appendBlocks(td, cell.blocks, defaultTabStopTwips);
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

  private makeHeaderFooter(type: 'header' | 'footer'): HTMLElement {
    const kind = type === 'header' ? this.headerKind : this.footerKind;
    let map: Partial<Record<'default' | 'first' | 'even', string>> = {};
    try {
      const section = this.document.getSection(0);
      map = type === 'header' ? section.headers : section.footers;
    } catch {
      map = {};
    }
    const part = map[kind] ?? map.default;
    const blocks = type === 'header'
      ? this.document.getHeaderBlocks(0, kind)
      : this.document.getFooterBlocks(0, kind);
    const partXml = part ? this.document.getPartXml(part) : '';
    const plainEditable = !!part && !/<w:(tbl|fldSimple|fldChar|drawing|hyperlink|object|pict|sdt|customXml|smartTag|ins|del)\b/.test(partXml);
    const area = this.root.ownerDocument.createElement('div');
    area.className = `docx-${type}`;
    const label = this.root.ownerDocument.createElement('div');
    label.className = 'docx-header-footer-label';
    label.id = `docx-${type}-${kind}-label`;
    label.textContent = `${type === 'header' ? '页眉' : '页脚'}（${kind}）`;
    const editable = this.root.ownerDocument.createElement('div');
    editable.contentEditable = plainEditable ? 'true' : 'false';
    editable.className = 'docx-header-footer-text';
    editable.setAttribute('role', 'textbox');
    editable.setAttribute('aria-multiline', 'true');
    editable.setAttribute('aria-labelledby', label.id);
    const renderedText = blocks.flatMap(block => block.type === 'paragraph' ? [block.paragraph.text] : []).join('\n');
    const normalizedRenderedText = renderedText.replace(/\r\n?/g, '\n').trimEnd();
    editable.textContent = renderedText;
    if (!plainEditable) editable.setAttribute('aria-readonly', 'true');
    editable.addEventListener('blur', () => {
      if (!plainEditable) return;
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

  private makeParagraph(paragraph: ParagraphInfo, defaultTabStopTwips: number): HTMLParagraphElement {
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
    let currentLineOffsetPx = 0;
    for (const run of paragraph.runs) {
      currentLineOffsetPx = this.appendRun(content, paragraph, run, defaultTabStopTwips, currentLineOffsetPx);
      for (const image of run.images ?? (run.image ? [run.image] : [])) content.append(this.makeImage(paragraph.index, image));
    }
    if (!paragraph.runs.length) content.textContent = paragraph.text;
    if (this.options.showFormattingMarks) content.append(this.makeMark('¶', '段落标记'));
    element.append(content);
    this.paragraphs.set(paragraph.index, { element, content, text: sanitizeText(this.readText(content)), failed: false });
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
    content.addEventListener('click', (event) => {
      const target = (event.target as Element | null)?.closest<HTMLElement>('[data-docx-link="1"]');
      if (!target || !this.root.contains(target)) return;
      event.preventDefault();
      if (!(event.ctrlKey || event.metaKey)) return;
      this.dispatchLinkClick(target);
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
    defaultTabStopTwips: number,
    currentLineOffsetPx: number,
  ): number {
    const unsafe = run.hyperlink?.unsafe ?? false;
    const hasSafeLink = !!(run.hyperlink && !unsafe && (run.hyperlink.url || run.hyperlink.anchor));
    const runSpan = this.root.ownerDocument.createElement(hasSafeLink ? 'a' : 'span');
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
    applyRunStyle(runSpan, run);
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
      if (this.options.showFormattingMarks && segment.includes(' ')) {
        const parts = segment.split(/( )/);
        for (const part of parts) {
          if (!part) continue;
          if (part === ' ') {
            runSpan.append(this.root.ownerDocument.createTextNode(' '));
            runSpan.append(this.makeMark('·', '空格'));
          } else runSpan.append(this.root.ownerDocument.createTextNode(part));
        }
      } else {
        runSpan.append(this.root.ownerDocument.createTextNode(segment));
      }
      currentLineOffsetPx += this.measure(segment, run);
    }
    paragraphElement.append(runSpan);
    return currentLineOffsetPx;
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
    if (!selection?.rangeCount) return;
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
    if (!image || image.dataset.image !== this.selectedImageInfo.id) return;
    event.preventDefault();
    this.document.deleteImage(this.selectedImageInfo);
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
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
