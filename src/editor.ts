import { DocxDocument } from './document.js';
import type {
  BorderFormat,
  ClipboardFragment,
  ClipboardRun,
  BordersFormat,
  CellFormat,
  DocumentRange,
  DocumentBlock,
  DocumentSnapshot,
  ImageInfo,
  ParagraphInfo,
  ShapeInfo,
  ReviewerFilterAuthor,
  RunFormat,
  RunInfo,
  SectionInfo,
  TabStop,
  TableCellLocation,
  TableFormat,
  TableRowInfo,
  WidthFormat,
} from './types.js';
import { placeholderDataUrl, pxToEmu } from './drawing.js';
import { isSafeHyperlinkUrl } from './hyperlink.js';
import { reviewerBucketKey, reviewerBucketOf } from './revisions.js';
import { eighthPointsToPx, normalizeColor, normalizeWidth, twipsToPx } from './table.js';
import { assertText, sanitizeText, sanitizeTextWithInfo } from './xml.js';

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

const DOCX_CLIPBOARD_MIME = 'application/x-docx-view+json';
const MAX_REVIEW_FILTER_AUTHORS = 1_000;

export interface EditorReviewFilter {
  authors?: ReviewerFilterAuthor[];
  showRevisions?: boolean;
  showComments?: boolean;
  revisionView?: 'final' | 'original' | 'markup';
}

interface NormalizedReviewFilter {
  authors?: ReviewerFilterAuthor[];
  showRevisions: boolean;
  showComments: boolean;
  revisionView: 'final' | 'original' | 'markup';
}

interface ReviewRenderContext {
  authors?: Set<string>;
  deletedTextByRun: Map<string, string>;
  revisionColors: Map<string, string>;
}

const REVISION_COLOR_PALETTE = ['#2E75B6', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646', '#1F497D', '#843C0C'];

function revisionColorIndexForKey(key: string): number {
  let hash = 2166136261;
  for (let i = 0; i < key.length; i++) {
    hash ^= key.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0) % REVISION_COLOR_PALETTE.length;
}

function normalizeReviewFilterAuthor(author: ReviewerFilterAuthor, path = 'reviewFilter.authors[]'): ReviewerFilterAuthor {
  if (!author || typeof author !== 'object' || Array.isArray(author)) throw new Error(`${path} must be an object.`);
  if (!['named', 'unattributed', 'empty', 'blank'].includes(author.kind)) {
    throw new Error(`${path}.kind must be one of: named, unattributed, empty, blank.`);
  }
  const rawAuthor = (author as { author?: string }).author;
  if (author.kind === 'unattributed') {
    if (rawAuthor !== undefined) throw new Error(`${path}.author must be omitted for unattributed kind.`);
    return { kind: 'unattributed' };
  }
  if (rawAuthor === undefined) throw new Error(`${path}.author is required for ${author.kind} kind.`);
  assertText(rawAuthor, `${path}.author`);
  if (author.kind === 'named') {
    if (!rawAuthor.trim()) throw new Error(`${path}.author must be non-empty for named kind.`);
    return { kind: 'named', author: rawAuthor };
  }
  if (author.kind === 'empty') {
    if (rawAuthor !== '') throw new Error(`${path}.author must be an empty string for empty kind.`);
    return { kind: 'empty', author: '' };
  }
  if (rawAuthor.trim()) throw new Error(`${path}.author must be whitespace-only for blank kind.`);
  return { kind: 'blank', author: rawAuthor };
}

function normalizeReviewFilter(filter: EditorReviewFilter | undefined): NormalizedReviewFilter {
  if (filter !== undefined && (!filter || typeof filter !== 'object' || Array.isArray(filter))) {
    throw new Error('reviewFilter must be an object.');
  }
  if (filter?.authors !== undefined && !Array.isArray(filter.authors)) throw new Error('reviewFilter.authors must be an array.');
  if (filter?.authors && filter.authors.length > MAX_REVIEW_FILTER_AUTHORS) {
    throw new Error(`reviewFilter.authors must contain at most ${MAX_REVIEW_FILTER_AUTHORS} items.`);
  }
  if (filter?.showRevisions !== undefined && typeof filter.showRevisions !== 'boolean') {
    throw new Error('reviewFilter.showRevisions must be boolean.');
  }
  if (filter?.showComments !== undefined && typeof filter.showComments !== 'boolean') {
    throw new Error('reviewFilter.showComments must be boolean.');
  }
  if (filter?.revisionView !== undefined && !['final', 'original', 'markup'].includes(filter.revisionView)) {
    throw new Error('reviewFilter.revisionView must be one of: final, original, markup.');
  }
  const authors = filter?.authors?.map((author) => normalizeReviewFilterAuthor(author))
    .filter((author, index, all) =>
      all.findIndex((entry) => reviewerBucketKey(entry) === reviewerBucketKey(author)) === index);
  return {
    ...(authors?.length ? { authors } : {}),
    showRevisions: filter?.showRevisions ?? true,
    showComments: filter?.showComments ?? true,
    revisionView: filter?.revisionView ?? 'markup',
  };
}

function reviewFilterEqual(a: NormalizedReviewFilter, b: NormalizedReviewFilter): boolean {
  const authorsA = a.authors ?? [];
  const authorsB = b.authors ?? [];
  return a.showRevisions === b.showRevisions &&
    a.showComments === b.showComments &&
    a.revisionView === b.revisionView &&
    authorsA.length === authorsB.length &&
  authorsA.every((value, index) => reviewerBucketKey(value) === reviewerBucketKey(authorsB[index]!));
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
  showFieldShading?: boolean;
  reviewFilter?: EditorReviewFilter;
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
    hasFields?: boolean;
  }>();
  private headerKind: 'default' | 'first' | 'even' = 'default';
  private footerKind: 'default' | 'first' | 'even' = 'default';
  private selected: number | null = null;
  private selectedImageInfo: ImageInfo | null = null;
  private selectedTableCellInfo: TableCellLocation | null = null;
  private selectedRangeInfo: { range: DocumentRange; format: RunFormat } | null = null;
  private composing = false;
  private renderAfterComposition = false;
  private destroyed = false;
  private reviewFilter: NormalizedReviewFilter;
  private readonly metrics: CanvasRenderingContext2D | null;
  private commentRunIds = new Map<string, number[]>();
  private commentParagraphIds = new Map<number, number[]>();
  private revisionRunIds = new Map<string, HTMLElement[]>();
  private revisionParagraphIds = new Map<number, HTMLElement[]>();
  private renderShapeInfos: ShapeInfo[] = [];
  private activeRevisionId: number | null = null;

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

  private dispatchCommentClick(target: HTMLElement): void {
    const ids = (target.dataset.docxCommentIds ?? '')
      .split(',')
      .map((value) => Number(value))
      .filter((value) => Number.isSafeInteger(value) && value >= 0);
    if (!ids.length) return;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (!EventClass) return;
    this.root.dispatchEvent(new EventClass('docx-commentclick', {
      bubbles: true,
      detail: { ids },
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
    this.reviewFilter = normalizeReviewFilter(options.reviewFilter);
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

  get selectedTableCell(): TableCellLocation | null {
    return this.selectedTableCellInfo ? { ...this.selectedTableCellInfo } : null;
  }

  get selectedRange(): DocumentRange | null {
    return this.selectedRangeInfo ? {
      start: { ...this.selectedRangeInfo.range.start },
      end: { ...this.selectedRangeInfo.range.end },
    } : null;
  }

  get selectedRangeFormat(): RunFormat | null {
    return this.selectedRangeInfo ? { ...this.selectedRangeInfo.format } : null;
  }

  private isMarkupReviewView(): boolean {
    return (this.reviewFilter ?? normalizeReviewFilter(this.options?.reviewFilter)).revisionView === 'markup';
  }

  /** Commit visible text before an external API operation or an export. */
  flush(): void {
    if (this.destroyed) return;
    if (!this.isMarkupReviewView()) return;
    let changed = false;
    for (const [index, entry] of this.paragraphs) {
      if (entry.hasFields) continue;
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
    this.selectedTableCellInfo = null;
    this.selectedRangeInfo = null;
    this.composing = false;
    this.renderAfterComposition = false;
    this.paragraphs.clear();
    this.render();
  }

  setReviewFilter(filter: EditorReviewFilter = {}): void {
    if (this.destroyed) return;
    this.reviewFilter ??= normalizeReviewFilter(this.options?.reviewFilter);
    const next = normalizeReviewFilter({
      ...(this.reviewFilter.authors !== undefined ? { authors: this.reviewFilter.authors } : {}),
      showRevisions: this.reviewFilter.showRevisions,
      showComments: this.reviewFilter.showComments,
      revisionView: this.reviewFilter.revisionView,
      ...filter,
    });
    if (reviewFilterEqual(this.reviewFilter, next)) return;
    this.flush();
    this.reviewFilter = next;
    this.render();
  }

  private filteredRevisionIds(): number[] {
    if (!this.reviewFilter.showRevisions) return [];
    const authors = this.reviewFilter.authors ? new Set(this.reviewFilter.authors.map((author) => reviewerBucketKey(author))) : null;
    return this.document
      .getRevisions()
      .filter((revision) => !authors || authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))
      .map((revision) => revision.id);
  }

  private setActiveRevision(id: number | null): void {
    this.activeRevisionId = id;
    if (typeof this.root.querySelectorAll !== 'function') return;
    for (const node of this.root.querySelectorAll<HTMLElement>('[data-docx-revision-ids]')) {
      const ids = (node.dataset.docxRevisionIds ?? '')
        .split(',')
        .map((value) => Number(value))
        .filter((value) => Number.isSafeInteger(value));
      node.classList.toggle('docx-revision-active', id !== null && ids.includes(id));
    }
  }

  private applyRevisionAction(action: () => void): boolean {
    const previous = this.document.revision;
    action();
    if (this.document.revision === previous) return false;
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
    return true;
  }

  acceptRevision(id: number): boolean {
    this.flush();
    const accepted = this.applyRevisionAction(() => this.document.acceptRevision(id));
    if (accepted) this.setActiveRevision(null);
    return accepted;
  }

  rejectRevision(id: number): boolean {
    this.flush();
    const rejected = this.applyRevisionAction(() => this.document.rejectRevision(id));
    if (rejected) this.setActiveRevision(null);
    return rejected;
  }

  acceptAllRevisions(filter: { authors?: ReviewerFilterAuthor[] } = {}): boolean {
    this.flush();
    const accepted = this.applyRevisionAction(() => {
      if (!filter.authors?.length) {
        this.document.acceptAllRevisions();
        return;
      }
      const authors = new Set(filter.authors.map((author) => reviewerBucketKey(author)));
      const operations = this.document
        .getRevisions()
        .filter((revision) => authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))
        .map((revision) => ({ type: 'acceptRevision' as const, id: revision.id }));
      if (!operations.length) return;
      this.document.applyOperations({ operations });
    });
    if (accepted) this.setActiveRevision(null);
    return accepted;
  }

  rejectAllRevisions(filter: { authors?: ReviewerFilterAuthor[] } = {}): boolean {
    this.flush();
    const rejected = this.applyRevisionAction(() => {
      if (!filter.authors?.length) {
        this.document.rejectAllRevisions();
        return;
      }
      const authors = new Set(filter.authors.map((author) => reviewerBucketKey(author)));
      const operations = this.document
        .getRevisions()
        .filter((revision) => authors.has(reviewerBucketKey(reviewerBucketOf(revision.author))))
        .map((revision) => ({ type: 'rejectRevision' as const, id: revision.id }));
      if (!operations.length) return;
      this.document.applyOperations({ operations });
    });
    if (rejected) this.setActiveRevision(null);
    return rejected;
  }

  focusNextRevision(): number | null {
    const ids = this.filteredRevisionIds();
    if (!ids.length) {
      this.setActiveRevision(null);
      return null;
    }
    const current = this.activeRevisionId === null ? -1 : ids.indexOf(this.activeRevisionId);
    const nextId = ids[(current + 1) % ids.length]!;
    this.setActiveRevision(nextId);
    const target = this.revisionRunIds.get(String(nextId))?.[0] ?? this.revisionParagraphIds.get(nextId)?.[0];
    target?.focus();
    target?.scrollIntoView({ block: 'nearest' });
    return nextId;
  }

  focusPreviousRevision(): number | null {
    const ids = this.filteredRevisionIds();
    if (!ids.length) {
      this.setActiveRevision(null);
      return null;
    }
    const current = this.activeRevisionId === null ? ids.length : ids.indexOf(this.activeRevisionId);
    const previousId = ids[(current - 1 + ids.length) % ids.length]!;
    this.setActiveRevision(previousId);
    const target = this.revisionRunIds.get(String(previousId))?.[0] ?? this.revisionParagraphIds.get(previousId)?.[0];
    target?.focus();
    target?.scrollIntoView({ block: 'nearest' });
    return previousId;
  }

  focusRevision(id: number): boolean {
    if (!Number.isSafeInteger(id) || id < 0) throw new Error('revision id must be a non-negative integer.');
    if (!this.filteredRevisionIds().includes(id)) return false;
    this.setActiveRevision(id);
    const target = this.revisionRunIds.get(String(id))?.[0] ?? this.revisionParagraphIds.get(id)?.[0];
    target?.focus();
    target?.scrollIntoView({ block: 'nearest' });
    return true;
  }

  render(): void {
    if (this.destroyed) return;
    this.reviewFilter ??= normalizeReviewFilter(this.options?.reviewFilter);
    if (this.composing) {
      this.renderAfterComposition = true;
      return;
    }
    const range = this.captureDocumentRange();
    const activeImageId = (this.root.ownerDocument.activeElement as HTMLElement | null)?.closest<HTMLElement>('[data-image]')?.dataset.image
      ?? this.selectedImageInfo?.id
      ?? null;
    this.flush();
    this.applyPageSetup();
    this.paragraphs.clear();
    this.commentRunIds ??= new Map();
    this.commentParagraphIds ??= new Map();
    this.commentRunIds.clear();
    this.commentParagraphIds.clear();
    this.revisionRunIds ??= new Map();
    this.revisionParagraphIds ??= new Map();
    this.revisionRunIds.clear();
    this.revisionParagraphIds.clear();
    const reviewContext: ReviewRenderContext = {
      ...(this.reviewFilter.authors ? { authors: new Set(this.reviewFilter.authors.map((author) => reviewerBucketKey(author))) } : {}),
      deletedTextByRun: new Map(),
      revisionColors: new Map(),
    };
    try {
      for (const revision of this.document.getRevisions({ kinds: ['deletion', 'move'] })) {
        if (revision.paragraph < 0 || revision.run === undefined || !revision.deletedText) continue;
        if (revision.kind === 'move' && revision.move?.side !== 'from') continue;
        if (reviewContext.authors && !reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(revision.author)))) continue;
        const key = `${revision.paragraph}:${revision.run}`;
        reviewContext.deletedTextByRun.set(key, `${reviewContext.deletedTextByRun.get(key) ?? ''}${revision.deletedText}`);
      }
      for (const comment of this.reviewFilter.showComments ? this.document.getComments() : []) {
        if (reviewContext.authors && !reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(comment.author)))) continue;
        if (!comment.anchor || comment.anchor.sourcePartPath !== this.document.mainDocumentPath) continue;
        if ('runs' in comment.anchor) {
          for (const run of comment.anchor.runs) {
            const key = `${comment.anchor.paragraph}:${run}`;
            const ids = this.commentRunIds.get(key) ?? [];
            if (!ids.includes(comment.id)) ids.push(comment.id);
            this.commentRunIds.set(key, ids);
          }
        } else {
          for (let paragraph = comment.anchor.startParagraph; paragraph <= comment.anchor.endParagraph; paragraph++) {
            const ids = this.commentParagraphIds.get(paragraph) ?? [];
            if (!ids.includes(comment.id)) ids.push(comment.id);
            this.commentParagraphIds.set(paragraph, ids);
          }
        }
      }
      const fragment = this.root.ownerDocument.createDocumentFragment();
      const canRenderHeaderFooter = typeof this.root.ownerDocument.createElement === 'function';
      if (canRenderHeaderFooter) fragment.append(this.makeHeaderFooter('header'));
      let defaultTabStopTwips = 720;
      try {
        defaultTabStopTwips = Math.max(1, Number(this.document.getSettings().defaultTabStop) || 720);
      } catch {
        defaultTabStopTwips = 720;
      }
      this.renderShapeInfos = this.document.getShapes();
      this.appendBlocks(fragment, this.document.getBlocks(), defaultTabStopTwips, reviewContext);
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
      if (range) this.restoreDocumentRange(range);
      this.updateRangeSelection(this.captureDocumentRange());
      this.setActiveRevision(this.activeRevisionId);
    } finally {
      this.renderShapeInfos = [];
      reviewContext.deletedTextByRun.clear();
      reviewContext.revisionColors.clear();
    }
  }

  destroy(): void {
    if (this.destroyed) return;
    this.flush();
    this.destroyed = true;
    this.root.ownerDocument.removeEventListener('selectionchange', this.handleSelection);
    this.root.removeEventListener('keydown', this.handleRootKeydown);
    this.root.remove();
    this.paragraphs.clear();
    this.selectedRangeInfo = null;
  }

  private readText(element: HTMLElement): string {
    const walk = (node: Node): string => {
      if (node.nodeType === 3) return node.textContent ?? '';
      if (node.nodeType !== 1) return '';
      const current = node as HTMLElement;
      if (current.dataset.image || current.dataset.docxMark !== undefined) return '';
      if (current.dataset.docxDeleted !== undefined) return '';
      if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return '';
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

  private appendBlocks(parent: Node, blocks: DocumentBlock[], defaultTabStopTwips: number, reviewContext: ReviewRenderContext): void {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        parent.appendChild(this.makeParagraph(block.paragraph, defaultTabStopTwips, reviewContext));
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
            this.appendBlocks(td, cell.blocks, defaultTabStopTwips, reviewContext);
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

  private makeParagraph(paragraph: ParagraphInfo, defaultTabStopTwips: number, reviewContext: ReviewRenderContext): HTMLParagraphElement {
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
      marker.setAttribute('data-docx-mark', '1');
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
    const hasFields = paragraph.runs.some((run) => run.field !== undefined);
    content.contentEditable = this.isMarkupReviewView() && !hasFields ? 'true' : 'false';
    content.spellcheck = false;
    content.setAttribute('role', 'textbox');
    content.setAttribute('aria-multiline', 'true');
    content.setAttribute('aria-label', `第 ${paragraph.index + 1} 段`);
    if (!this.isMarkupReviewView()) content.setAttribute('aria-readonly', 'true');
    if (paragraph.numbering) content.setAttribute('aria-description', `列表项 ${paragraph.numbering.text}，级别 ${paragraph.numbering.level + 1}`);
    const runRevisionIds = new Set<number>();
    const hasRunRevision = this.reviewFilter.showRevisions && paragraph.runs.some((run) =>
      run.revisions?.some((revision) => {
        const visible = !reviewContext.authors || reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(revision.author)));
        if (visible) runRevisionIds.add(revision.id);
        return visible;
      }));
    const visibleParagraphRevision = this.reviewFilter.showRevisions && paragraph.paragraphRevision &&
      (!reviewContext.authors || reviewContext.authors.has(reviewerBucketKey(reviewerBucketOf(paragraph.paragraphRevision.author))))
      ? paragraph.paragraphRevision
      : undefined;
    if ((hasRunRevision || visibleParagraphRevision) && this.reviewFilter.revisionView === 'markup') {
      const marker = this.makeMark('▎', '修订变更条');
      marker.classList.add('docx-change-bar');
      marker.style.position = 'absolute';
      marker.style.left = '-0.9em';
      marker.style.top = '0';
      marker.style.bottom = '0';
      marker.style.display = 'flex';
      marker.style.alignItems = 'stretch';
      marker.style.color = this.reviewColor(visibleParagraphRevision?.author, reviewContext);
      marker.style.opacity = '0.9';
      marker.textContent = '│';
      marker.title = '';
      marker.setAttribute('role', 'presentation');
      marker.setAttribute('aria-hidden', 'true');
      element.append(marker);
      const markerRevisionIds = new Set<number>(runRevisionIds);
      if (visibleParagraphRevision) markerRevisionIds.add(visibleParagraphRevision.id);
      if (markerRevisionIds.size) {
        element.dataset.docxRevisionIds = [...markerRevisionIds].join(',');
      }
      for (const id of markerRevisionIds) {
        const list = this.revisionParagraphIds.get(id) ?? [];
        list.push(element);
        this.revisionParagraphIds.set(id, list);
      }
    }
    const paragraphCommentIds = this.commentParagraphIds.get(paragraph.index);
    if (paragraphCommentIds?.length) {
      element.classList.add('docx-comment-anchor');
      element.dataset.docxCommentIds = paragraphCommentIds.join(',');
    }
    element.addEventListener('mousedown', (event) => {
      const target = event.target as Node | null;
      if (target && content.contains(target)) return;
      event.preventDefault();
      this.focusContent(content);
    });
    let currentLineOffsetPx = 0;
    for (const run of paragraph.runs) {
      const visibleRun = this.reviewScopedRun(paragraph.index, run, reviewContext);
      currentLineOffsetPx = this.appendRun(content, paragraph, visibleRun, reviewContext, defaultTabStopTwips, currentLineOffsetPx);
      this.appendDeletedRunVisualization(content, paragraph.index, visibleRun, reviewContext);
      if (run.noteReference) {
        const marker = this.root.ownerDocument.createElement('sup');
        marker.className = 'docx-note-ref';
        marker.contentEditable = 'false';
        marker.setAttribute('data-docx-mark', '1');
        marker.textContent = run.noteReference.marker;
        marker.setAttribute('aria-label', `${run.noteReference.kind} reference ${run.noteReference.marker}`);
        content.append(marker);
      }
      for (const image of run.images ?? (run.image ? [run.image] : [])) content.append(this.makeImage(paragraph.index, image));
      for (const shape of this.renderShapeInfos.filter((item) => item.paragraph === paragraph.index && item.run === run.index)) {
        content.append(this.makeShape(shape, defaultTabStopTwips, reviewContext));
      }
    }

    if (!paragraph.runs.length) content.textContent = paragraph.text;
    if (this.options.showFormattingMarks) content.append(this.makeMark('¶', '段落标记'));
    element.append(content);
    this.paragraphs.set(paragraph.index, { element, content, text: sanitizeText(this.readText(content)), failed: false, hasFields });
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
    content.addEventListener('copy', (event) => this.handleClipboardCopy(event));
    content.addEventListener('cut', (event) => this.handleClipboardCut(event, content));
    content.addEventListener('paste', (event) => this.handleClipboardPaste(event, content));
    // Do not allow rich HTML or embedded objects from drag-and-drop either.
    content.addEventListener('drop', (event) => { event.preventDefault(); });
    content.addEventListener('keydown', (event) => {
      if (this.handleHistoryShortcut(event)) return;
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
    content.addEventListener('click', (event) => {
      const target = (event.target as Element | null)?.closest<HTMLElement>('[data-docx-comment-ids]');
      if (!target || !this.root.contains(target)) return;
      this.dispatchCommentClick(target);
    });
    content.addEventListener('beforeinput', (event) => {
      if (event.inputType === 'historyUndo') {
        event.preventDefault();
        this.applyHistory('undo');
        return;
      }
      if (event.inputType === 'historyRedo') {
        event.preventDefault();
        this.applyHistory('redo');
        return;
      }
      if (this.selectionTouchesField(content)) {
        event.preventDefault();
        return;
      }
      if (!event.isComposing && ['insertParagraph', 'insertLineBreak'].includes(event.inputType)) {
        event.preventDefault();
        this.insertText(content, '\n');
      }
      if (event.inputType.startsWith('format')) event.preventDefault();
    });
    return element;
  }

  private makeShape(shape: ShapeInfo, defaultTabStopTwips: number, reviewContext: ReviewRenderContext): HTMLElement {
    const wrapper = this.root.ownerDocument.createElement(shape.placement === 'floating' ? 'div' : 'span');
    wrapper.className = `docx-shape docx-shape-${shape.kind}`;
    wrapper.contentEditable = 'false';
    wrapper.dataset.docxShape = shape.id;
    wrapper.style.display = shape.placement === 'floating' ? 'block' : 'inline-block';
    wrapper.style.width = `${Math.max(48, shape.widthPx || 160)}px`;
    wrapper.style.minHeight = `${Math.max(48, shape.heightPx || 90)}px`;
    wrapper.style.maxWidth = '100%';
    wrapper.style.margin = shape.placement === 'floating' ? '8px 12px 8px 0' : '0 2px';
    wrapper.style.verticalAlign = 'text-bottom';
    wrapper.style.border = '1px solid #c7d3e5';
    wrapper.style.background = '#f7f9fd';
    wrapper.style.padding = '4px';
    wrapper.style.boxSizing = 'border-box';
    wrapper.setAttribute('aria-label', shape.alt ?? shape.geometry ?? shape.kind);
    if (shape.hasTextContent) {
      for (const paragraph of this.document.getShapeParagraphs(shape.id)) {
        const element = this.root.ownerDocument.createElement('p');
        element.className = 'docx-shape-paragraph';
        element.style.whiteSpace = 'pre-wrap';
        applyParagraphStyle(element, paragraph);
        let offset = 0;
        for (const run of paragraph.runs) offset = this.appendRun(element, paragraph, run, reviewContext, defaultTabStopTwips, offset);
        if (!paragraph.runs.length) element.textContent = paragraph.text;
        wrapper.append(element);
      }
    } else {
      const label = shape.alt ?? shape.geometry ?? ({
        smartArt: 'SmartArt', chart: '图表', textbox: '文本框', shape: '形状', ole: 'OLE', unknown: '对象',
      }[shape.kind] ?? shape.kind);
      const image = this.root.ownerDocument.createElement('img');
      image.src = placeholderDataUrl(label, shape.widthPx || 160, shape.heightPx || 90);
      image.alt = label;
      image.draggable = false;
      image.style.width = '100%';
      image.style.height = '100%';
      image.style.display = 'block';
      wrapper.append(image);
    }
    return wrapper;
  }

  private reviewScopedRun(paragraphIndex: number, run: RunInfo, reviewContext: ReviewRenderContext): RunInfo {
    if (!run.revisions?.length) return run;
    let revisions = run.revisions;
    if (!this.reviewFilter.showRevisions) revisions = [];
    if (reviewContext.authors) {
      revisions = revisions.filter((revision) => reviewContext.authors!.has(reviewerBucketKey(reviewerBucketOf(revision.author))));
    }
    const hasInsertion = revisions.some((revision) => revision.kind === 'insertion');
    const hasDeletion = revisions.some((revision) => revision.kind === 'deletion');
    const hasMoveFrom = revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'from');
    const hasMoveTo = revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'to');
    const view = this.reviewFilter.revisionView;
    let text = run.text;
    if (view === 'original') {
      if (hasInsertion || hasMoveTo) text = '';
      else if (hasDeletion || hasMoveFrom) text = reviewContext.deletedTextByRun.get(`${paragraphIndex}:${run.index}`) ?? run.text;
    } else if (view === 'final') {
      if (hasDeletion || hasMoveFrom) text = '';
    }
    if (text === run.text && revisions.length === run.revisions.length && revisions.every((entry, index) => entry === run.revisions![index])) {
      return run;
    }
    return { ...run, text, revisions: revisions.length ? revisions : undefined };
  }

  private reviewColor(author: string | undefined, reviewContext: ReviewRenderContext): string {
    const key = reviewerBucketKey(reviewerBucketOf(author));
    let color = reviewContext.revisionColors.get(key);
    if (!color) {
      color = REVISION_COLOR_PALETTE[revisionColorIndexForKey(key)]!;
      reviewContext.revisionColors.set(key, color);
    }
    return color;
  }

  private revisionAriaDescription(revisions: RunInfo['revisions']): string {
    if (!revisions?.length) return '';
    const labels = [];
    if (revisions.some((revision) => revision.kind === 'insertion')) labels.push('插入');
    if (revisions.some((revision) => revision.kind === 'deletion')) labels.push('删除');
    if (revisions.some((revision) => revision.kind === 'move')) labels.push('移动');
    const author = revisions.find((revision) => revision.author !== undefined)?.author;
    return author ? `修订：${labels.join(' / ') || '变更'}，作者 ${author}` : `修订：${labels.join(' / ') || '变更'}`;
  }

  private registerRevisionNode(ids: number[], node: HTMLElement): void {
    if (!ids.length) return;
    node.dataset.docxRevisionIds = ids.join(',');
    node.tabIndex = -1;
    for (const id of ids) {
      const key = String(id);
      const list = this.revisionRunIds.get(key) ?? [];
      list.push(node);
      this.revisionRunIds.set(key, list);
    }
  }

  private appendDeletedRunVisualization(
    paragraphElement: HTMLElement,
    paragraphIndex: number,
    run: RunInfo,
    reviewContext: ReviewRenderContext,
  ): void {
    if (this.reviewFilter.revisionView !== 'markup' || !run.revisions?.some((revision) =>
      revision.kind === 'deletion' || (revision.kind === 'move' && revision.move?.side === 'from'))) return;
    const deletedText = reviewContext.deletedTextByRun.get(`${paragraphIndex}:${run.index}`);
    if (!deletedText) return;
    const marker = this.root.ownerDocument.createElement('span');
    marker.className = 'docx-deleted-text';
    marker.dataset.docxDeleted = '1';
    marker.contentEditable = 'false';
    marker.textContent = deletedText;
    const hasMoveFrom = run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'from');
    marker.style.textDecoration = hasMoveFrom ? 'line-through underline' : 'line-through';
    if (hasMoveFrom) marker.style.textDecorationStyle = 'solid double';
    marker.style.opacity = '0.85';
    const authorRevision = run.revisions.find((revision) =>
      revision.kind === 'deletion' || (revision.kind === 'move' && revision.move?.side === 'from'));
    marker.style.color = this.reviewColor(authorRevision?.author, reviewContext);
    marker.setAttribute('aria-label', hasMoveFrom ? `修订移动来源文本：${deletedText}` : `修订删除文本：${deletedText}`);
    if (hasMoveFrom) marker.append(this.makeMark('↤', '移动来源'));
    this.registerRevisionNode(run.revisions.map((revision) => revision.id), marker);
    paragraphElement.append(marker);
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
    reviewContext: ReviewRenderContext,
    defaultTabStopTwips: number,
    currentLineOffsetPx: number,
  ): number {
    const unsafe = run.hyperlink?.unsafe ?? false;
    const hasSafeLink = !!(run.hyperlink && !unsafe && (run.hyperlink.url || run.hyperlink.anchor));
    const runSpan = this.root.ownerDocument.createElement(hasSafeLink ? 'a' : 'span');
    runSpan.dataset.docxRun = String(run.index);
    if (run.field) {
      runSpan.dataset.docxField = String(run.field.index);
      runSpan.dataset.docxFieldRole = run.field.role;
      runSpan.contentEditable = 'false';
      if (run.field.role === 'result') {
        runSpan.dataset.docxContent = '1';
        if (this.options.showFieldShading !== false) runSpan.classList.add('docx-field-shading');
      }
    }
    if (run.revisions?.length) runSpan.dataset.docxRevisionIds = run.revisions.map((revision) => revision.id).join(',');
    const commentIds = [...new Set([...(this.commentParagraphIds.get(paragraph.index) ?? []), ...(this.commentRunIds.get(`${paragraph.index}:${run.index}`) ?? [])])];
    if (commentIds.length) {
      runSpan.classList.add('docx-comment-anchor');
      runSpan.dataset.docxCommentIds = commentIds.join(',');
    }
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
    if (run.revisions?.length && this.reviewFilter.showRevisions && this.reviewFilter.revisionView === 'markup') {
      const hasInsertion = run.revisions.some((revision) => revision.kind === 'insertion');
      const hasDeletion = run.revisions.some((revision) => revision.kind === 'deletion');
      const hasMoveFrom = run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'from');
      const hasMoveTo = run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'to');
      const author = run.revisions.find((revision) => revision.author !== undefined)?.author;
      const revisionColor = this.reviewColor(author, reviewContext);
      const textDecoration = [hasInsertion || hasMoveTo ? 'underline' : '', hasDeletion || hasMoveFrom ? 'line-through' : ''].filter(Boolean).join(' ');
      if (textDecoration) runSpan.style.textDecoration = textDecoration;
      if (hasMoveTo || hasMoveFrom) runSpan.style.textDecorationStyle = 'double';
      if (textDecoration) runSpan.style.textDecorationColor = revisionColor;
      const description = this.revisionAriaDescription(run.revisions);
      if (description) runSpan.setAttribute('aria-description', description);
      this.registerRevisionNode(run.revisions.map((revision) => revision.id), runSpan);
    }
    const segments = (run.field?.role === 'instruction' ? '' : run.text).split(/(\t|\n)/);
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
    if (run.revisions?.length && this.reviewFilter.showRevisions && this.reviewFilter.revisionView === 'markup' &&
        run.revisions.some((revision) => revision.kind === 'move' && revision.move?.side === 'to')) {
      runSpan.append(this.makeMark('↦', '移动目标'));
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
    if (!selection?.rangeCount || this.selectionTouchesField(element)) return;
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

  private currentDocumentRange(): DocumentRange | null {
    const range = this.captureDocumentRange();
    return range ? this.documentRange(range) : null;
  }

  private parseClipboardFragment(raw: string | null | undefined): ClipboardFragment | null {
    if (!raw) return null;
    try {
      const value = JSON.parse(raw) as ClipboardFragment;
      if (!value || value.version !== 1 || typeof value.text !== 'string' || !Array.isArray(value.paragraphs)) return null;
      return value;
    } catch {
      return null;
    }
  }

  private writeClipboardFragment(data: DataTransfer | null, fragment: ClipboardFragment): boolean {
    if (!data) return false;
    const json = JSON.stringify(fragment);
    data.setData(DOCX_CLIPBOARD_MIME, json);
    data.setData('text/plain', fragment.text);
    data.setData('text/html', fragment.text);
    return true;
  }

  private mapExternalHtmlFragment(html: string, plainText: string): ClipboardFragment {
    const parser = this.root.ownerDocument.defaultView?.DOMParser;
    if (!parser) return { version: 1, text: sanitizeText(plainText), paragraphs: [{ runs: [{ text: sanitizeText(plainText) }] }] };
    const document = new parser().parseFromString(html, 'text/html');
    const paragraphs: ClipboardFragment['paragraphs'] = [];
    const blocks: NonNullable<ClipboardFragment['blocks']> = [];
    let nextListId = 1;
    type State = { format: RunFormat; hyperlink?: ClipboardRun['hyperlink'] };
    const pushParagraph = (runs: ClipboardRun[]): void => {
      const normalized = runs
        .map((run) => ({ ...run, ...(run.text ? { text: sanitizeText(run.text) } : {}) }))
        .filter((run) => run.text || run.images?.length);
      if (normalized.length) paragraphs.push({ runs: normalized });
    };
    const sameFormat = (left: RunFormat | undefined, right: RunFormat | undefined): boolean => {
      const keys = new Set([...Object.keys(left ?? {}), ...Object.keys(right ?? {})]);
      for (const key of keys) {
        if ((left as Record<string, unknown> | undefined)?.[key] !== (right as Record<string, unknown> | undefined)?.[key]) return false;
      }
      return true;
    };
    const sameHyperlink = (left: ClipboardRun['hyperlink'], right: ClipboardRun['hyperlink']): boolean =>
      (left?.url ?? '') === (right?.url ?? '') &&
      (left?.anchor ?? '') === (right?.anchor ?? '') &&
      (left?.tooltip ?? '') === (right?.tooltip ?? '');
    const pushRun = (runs: ClipboardRun[], run: ClipboardRun): void => {
      const last = runs.at(-1);
      if (last && !last.images?.length && !run.images?.length &&
          sameFormat(last.format, run.format) &&
          sameHyperlink(last.hyperlink, run.hyperlink)) {
        last.text = `${last.text ?? ''}${run.text ?? ''}`;
      } else runs.push(run);
    };
    const walkInline = (node: Node, state: State, runs: ClipboardRun[], depth = 0): void => {
      if (depth > 100) return;
      if (node.nodeType === 3) {
        const text = node.textContent ?? '';
        if (text) pushRun(runs, { text, format: { ...state.format }, ...(state.hyperlink ? { hyperlink: { ...state.hyperlink } } : {}) });
        return;
      }
      if (node.nodeType !== 1) return;
      const element = node as HTMLElement;
      const tag = element.tagName.toLowerCase();
      if (['script', 'style', 'noscript'].includes(tag)) return;
      if (tag === 'br') {
        pushRun(runs, { text: '\n', format: { ...state.format }, ...(state.hyperlink ? { hyperlink: { ...state.hyperlink } } : {}) });
        return;
      }
      if (tag === 'img') {
        const src = element.getAttribute('src') ?? '';
        const match = src.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/=]+)$/);
        if (!match || !['image/png', 'image/jpeg', 'image/gif', 'image/bmp'].includes(match[1]!.toLowerCase())) return;
        pushRun(runs, {
          images: [{
            contentType: match[1]!,
            bytes: match[2]!,
            alt: element.getAttribute('alt') ?? undefined,
          }],
          format: { ...state.format },
          ...(state.hyperlink ? { hyperlink: { ...state.hyperlink } } : {}),
        });
        return;
      }
      const nextState: State = { format: { ...state.format }, hyperlink: state.hyperlink ? { ...state.hyperlink } : undefined };
      if (['b', 'strong'].includes(tag)) nextState.format.bold = true;
      if (['i', 'em'].includes(tag)) nextState.format.italic = true;
      if (tag === 'u') nextState.format.underline = true;
      if (['s', 'del'].includes(tag)) nextState.format.strike = true;
      if (tag === 'a') {
        const href = element.getAttribute('href') ?? '';
        if (href && isSafeHyperlinkUrl(href)) nextState.hyperlink = { url: href };
        else nextState.hyperlink = undefined;
      }
      for (const child of Array.from(node.childNodes)) walkInline(child, nextState, runs, depth + 1);
    };
    const walkList = (list: HTMLElement, kind: 'bullet' | 'decimal', level: number, listId: number, depth: number): void => {
      if (depth > 100) return;
      const items = Array.from(list.children).filter((child) => child.tagName.toLowerCase() === 'li');
      for (const item of items) {
        const runs: ClipboardRun[] = [];
        for (const child of Array.from(item.childNodes)) {
          if (child.nodeType === 1) {
            const tag = (child as HTMLElement).tagName.toLowerCase();
            if (tag === 'ul') { walkList(child as HTMLElement, 'bullet', Math.min(8, level + 1), listId, depth + 1); continue; }
            if (tag === 'ol') { walkList(child as HTMLElement, 'decimal', Math.min(8, level + 1), listId, depth + 1); continue; }
          }
          walkInline(child, { format: {} }, runs, depth + 1);
        }
        pushParagraph(runs);
        if (paragraphs.length) {
          const paragraph = paragraphs[paragraphs.length - 1]!;
          paragraph.numbering = { kind, level, listId };
          blocks.push({ type: 'paragraph', paragraph });
        }
      }
    };
    const walkBlocks = (node: Node, depth = 0): void => {
      if (depth > 100) return;
      if (node.nodeType === 3) {
        const text = node.textContent?.trim();
        if (!text) return;
        const paragraph = { runs: [{ text, format: {} }] };
        pushParagraph(paragraph.runs);
        blocks.push({ type: 'paragraph', paragraph });
        return;
      }
      if (node.nodeType !== 1) return;
      const element = node as HTMLElement;
      const tag = element.tagName.toLowerCase();
      if (['script', 'style', 'noscript'].includes(tag)) return;
      if (['p', 'div'].includes(tag)) {
        const runs: ClipboardRun[] = [];
        for (const child of Array.from(element.childNodes)) walkInline(child, { format: {} }, runs, depth + 1);
        pushParagraph(runs);
        if (paragraphs.length) blocks.push({ type: 'paragraph', paragraph: paragraphs[paragraphs.length - 1]! });
        return;
      }
      if (tag === 'ul' || tag === 'ol') {
        walkList(element, tag === 'ul' ? 'bullet' : 'decimal', 0, nextListId++, depth + 1);
        return;
      }
      if (tag === 'table') {
        const tableRows: ClipboardFragment['paragraphs'][] = [];
        const rows = Array.from(element.children).flatMap((child) => {
          const name = child.tagName.toLowerCase();
          if (name === 'tr') return [child];
          if (['thead', 'tbody', 'tfoot'].includes(name)) {
            return Array.from(child.children).filter((item) => item.tagName.toLowerCase() === 'tr');
          }
          return [];
        });
        for (const row of rows) {
          const cells = Array.from(row.children).filter((child) => ['th', 'td'].includes(child.tagName.toLowerCase()));
          const tableRow: ClipboardFragment['paragraphs'] = [];
          for (const cell of cells) {
            const runs: ClipboardRun[] = [];
            for (const child of Array.from(cell.childNodes)) walkInline(child, { format: {} }, runs, depth + 1);
            const cellParagraph: ClipboardFragment['paragraphs'][number] = { runs };
            tableRow.push(cellParagraph);
          }
          if (tableRow.length) tableRows.push(tableRow);
        }
        if (tableRows.length) blocks.push({ type: 'table', table: { rows: tableRows } });
        return;
      }
      const runs: ClipboardRun[] = [];
      for (const child of Array.from(element.childNodes)) walkInline(child, { format: {} }, runs, depth + 1);
      pushParagraph(runs);
      if (paragraphs.length) blocks.push({ type: 'paragraph', paragraph: paragraphs[paragraphs.length - 1]! });
    };
    for (const child of Array.from(document.body.childNodes)) walkBlocks(child, 0);
    if (!blocks.length) {
      const text = sanitizeText(plainText || document.body.textContent || '');
      return { version: 1, text, paragraphs: [{ runs: [{ text }] }] };
    }
    const allParagraphs = blocks.flatMap((block) => block.type === 'paragraph' ? [block.paragraph] : block.table.rows.flat());
    const text = blocks.map((block) => {
      if (block.type === 'paragraph') return block.paragraph.runs.map((run) => run.text ?? '').join('');
      return block.table.rows
        .map((row) => row.map((cell) => cell.runs.map((run) => run.text ?? '').join('')).join('\t'))
        .join('\n');
    }).join('\n');
    return { version: 1, text, paragraphs: allParagraphs, blocks };
  }

  private handleClipboardCopy(event: ClipboardEvent): void {
    const range = this.currentDocumentRange();
    if (!range) return;
    const fragment = this.document.copyClipboardFragment(range);
    if (!this.writeClipboardFragment(event.clipboardData, fragment)) return;
    event.preventDefault();
  }

  private handleClipboardCut(event: ClipboardEvent, content: HTMLElement): void {
    if (this.selectionTouchesField(content)) {
      event.preventDefault();
      return;
    }
    const range = this.currentDocumentRange();
    if (!range) return;
    if (range.start.paragraph !== range.end.paragraph || range.start.offset === range.end.offset) {
      event.preventDefault();
      return;
    }
    const fragment = this.document.copyClipboardFragment(range);
    if (!this.writeClipboardFragment(event.clipboardData, fragment)) return;
    const before = this.document.revision;
    this.document.beginHistoryGroup('cut');
    let handled = false;
    try {
      handled = this.document.pasteClipboardFragment(range, { version: 1, text: '', paragraphs: [] });
    } finally {
      this.document.endHistoryGroup();
    }
    event.preventDefault();
    if (!handled) return;
    if (this.document.revision === before) {
      this.insertText(content, '');
      return;
    }
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
  }

  private handleClipboardPaste(event: ClipboardEvent, content: HTMLElement): void {
    if (this.selectionTouchesField(content)) {
      event.preventDefault();
      return;
    }
    const data = event.clipboardData;
    if (!data) return;
    const range = this.currentDocumentRange();
    if (!range) return;
    const plain = data.getData('text/plain') ?? '';
    const own = this.parseClipboardFragment(data.getData(DOCX_CLIPBOARD_MIME));
    if (own) {
      try {
        const before = this.document.revision;
        if (this.document.pasteClipboardFragment(range, own)) {
          event.preventDefault();
          if (this.document.revision !== before) {
            this.render();
            this.options.onChange?.(this.document.getSnapshot());
          }
          return;
        }
      } catch (error) {
        this.reportError(error, range.start.paragraph);
      }
    }
    const html = data.getData('text/html');
    if (html) {
      try {
        const mapped = this.mapExternalHtmlFragment(html, plain);
        const before = this.document.revision;
        if (this.document.pasteClipboardFragment(range, mapped)) {
          event.preventDefault();
          if (this.document.revision !== before) {
            this.render();
            this.options.onChange?.(this.document.getSnapshot());
          }
          return;
        }
      } catch (error) {
        this.reportError(error, range.start.paragraph);
      }
    }
    event.preventDefault();
    this.insertText(content, plain);
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

  private textLength(node: Node): number {
    if (node.nodeType === 3) return Array.from(node.textContent ?? '').length;
    if (node.nodeType !== 1) return 0;
    const current = node as HTMLElement;
    if (current.dataset.image || current.dataset.docxMark !== undefined) return 0;
    if (current.dataset.docxDeleted !== undefined) return 0;
    if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return 0;
    if (current.tagName === 'BR') return 1;
    return Array.from(current.childNodes).reduce((total, child) => total + this.textLength(child), 0);
  }

  private paragraphText(index: number): string {
    return this.document.getParagraphs().find((item) => item.index === index)?.text ?? '';
  }

  private codeUnitsFromCodePoints(text: string, points: number): number {
    let units = 0;
    let count = 0;
    for (const char of text) {
      if (count >= points) break;
      units += char.length;
      count++;
    }
    return units;
  }

  private documentRange(range: DocumentRange): DocumentRange {
    return {
      start: {
        paragraph: range.start.paragraph,
        offset: this.codeUnitsFromCodePoints(this.paragraphText(range.start.paragraph), range.start.offset),
      },
      end: {
        paragraph: range.end.paragraph,
        offset: this.codeUnitsFromCodePoints(this.paragraphText(range.end.paragraph), range.end.offset),
      },
    };
  }

  private offsetWithin(root: HTMLElement, target: Node, targetOffset: number): number {
    let offset = 0;
    const walk = (node: Node): boolean => {
      if (node === target) {
        if (node.nodeType === 3) {
          offset += Array.from((node.textContent ?? '').slice(0, Math.max(0, targetOffset))).length;
          return true;
        }
        if (node.nodeType !== 1) return true;
        const current = node as HTMLElement;
        if (current.dataset.image || current.dataset.docxMark !== undefined) return true;
        if (current.dataset.docxDeleted !== undefined) return true;
        if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return true;
        if (current.tagName === 'BR') {
          offset += targetOffset > 0 ? 1 : 0;
          return true;
        }
        const childNodes = Array.from(node.childNodes);
        for (let i = 0; i < Math.min(targetOffset, childNodes.length); i++) {
          offset += this.textLength(childNodes[i]!);
        }
        return true;
      }
      if (node.nodeType === 3) {
        offset += Array.from(node.textContent ?? '').length;
        return false;
      }
      if (node.nodeType !== 1) return false;
      const current = node as HTMLElement;
      if (current.dataset.image || current.dataset.docxMark !== undefined) return false;
      if (current.dataset.docxDeleted !== undefined) return false;
      if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return false;
      if (current.tagName === 'BR') {
        offset += 1;
        return false;
      }
      for (const child of Array.from(node.childNodes)) {
        if (walk(child)) return true;
      }
      return false;
    };
    walk(root);
    return offset;
  }

  private positionFromOffset(root: HTMLElement, offset: number): [Node, number] {
    const total = this.textLength(root);
    let remaining = Math.max(0, Math.min(offset, total));
    const locate = (node: Node): [Node, number] | null => {
      if (node.nodeType === 3) {
        const text = node.textContent ?? '';
        const textLength = Array.from(text).length;
        if (remaining <= textLength) return [node, this.codeUnitsFromCodePoints(text, remaining)];
        remaining -= textLength;
        return null;
      }
      if (node.nodeType !== 1) return null;
      const current = node as HTMLElement;
      if (current.dataset.image || current.dataset.docxMark !== undefined) return null;
      if (current.dataset.docxDeleted !== undefined) return null;
      if (current.contentEditable === 'false' && current.dataset.docxContent === undefined) return null;
      if (current.tagName === 'BR') {
        if (remaining <= 1) {
          const parent = node.parentNode as Node;
          const index = Array.prototype.indexOf.call(parent.childNodes, node);
          return [parent, remaining === 0 ? index : index + 1];
        }
        remaining -= 1;
        return null;
      }
      for (const child of Array.from(node.childNodes)) {
        const found = locate(child);
        if (found) return found;
      }
      return null;
    };
    return locate(root) ?? [root, root.childNodes.length];
  }

  private captureDocumentRange(): DocumentRange | null {
    const selection = this.root.ownerDocument.getSelection();
    if (!selection?.rangeCount) return null;
    const raw = selection.getRangeAt(0);
    if (!this.root.contains(raw.startContainer) || !this.root.contains(raw.endContainer)) return null;
    const startElement = raw.startContainer.nodeType === 1 ? raw.startContainer as Element : raw.startContainer.parentElement;
    const endElement = raw.endContainer.nodeType === 1 ? raw.endContainer as Element : raw.endContainer.parentElement;
    const startContent = startElement?.closest<HTMLElement>('.docx-paragraph-content');
    const endContent = endElement?.closest<HTMLElement>('.docx-paragraph-content');
    const startParagraph = startContent?.closest<HTMLElement>('[data-paragraph]');
    const endParagraph = endContent?.closest<HTMLElement>('[data-paragraph]');
    if (!startContent || !endContent || !startParagraph || !endParagraph) return null;
    const start = {
      paragraph: Number(startParagraph.dataset.paragraph),
      offset: this.offsetWithin(startContent, raw.startContainer, raw.startOffset),
    };
    const end = {
      paragraph: Number(endParagraph.dataset.paragraph),
      offset: this.offsetWithin(endContent, raw.endContainer, raw.endOffset),
    };
    if (!Number.isSafeInteger(start.paragraph) || !Number.isSafeInteger(end.paragraph)) return null;
    const ordered = start.paragraph > end.paragraph || (start.paragraph === end.paragraph && start.offset > end.offset)
      ? { start: end, end: start }
      : { start, end };
    return ordered;
  }

  private restoreDocumentRange(range: DocumentRange): void {
    const start = this.paragraphs.get(range.start.paragraph)?.content;
    const end = this.paragraphs.get(range.end.paragraph)?.content;
    if (!start || !end) return;
    const selection = this.root.ownerDocument.getSelection();
    if (!selection) return;
    const startPoint = this.positionFromOffset(start, range.start.offset);
    const endPoint = this.positionFromOffset(end, range.end.offset);
    const domRange = this.root.ownerDocument.createRange();
    domRange.setStart(...startPoint);
    domRange.setEnd(...endPoint);
    selection.removeAllRanges();
    selection.addRange(domRange);
  }

  private updateRangeSelection(range: DocumentRange | null): void {
    const key = range
      ? `${range.start.paragraph}:${range.start.offset}-${range.end.paragraph}:${range.end.offset}`
      : '';
    const previous = this.selectedRangeInfo
      ? `${this.selectedRangeInfo.range.start.paragraph}:${this.selectedRangeInfo.range.start.offset}-${this.selectedRangeInfo.range.end.paragraph}:${this.selectedRangeInfo.range.end.offset}`
      : '';
    if (key === previous) return;
    if (!range) {
      const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
      this.selectedRangeInfo = null;
      if (EventClass && previous) this.root.dispatchEvent(new EventClass('docx-rangechange', { bubbles: true, detail: null }));
      return;
    }
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    let format: RunFormat = {};
    try {
      format = this.document.getDocumentRangeFormat(this.documentRange(range));
    } catch (error) {
      if (!(error instanceof Error) || !/Cross-container document ranges are not supported/.test(error.message)) throw error;
      format = {};
    }
    this.selectedRangeInfo = {
      range: {
        start: { ...range.start },
        end: { ...range.end },
      },
      format,
    };
    if (EventClass) {
      this.root.dispatchEvent(new EventClass('docx-rangechange', {
        bubbles: true,
        detail: { range: this.selectedRange, format: this.selectedRangeFormat ?? {} },
      }));
    }
  }

  private selectParagraph(index: number): void {
    const cell = this.document.getTableCellAt(index);
    const previous = this.selectedTableCellInfo;
    const paragraphChanged = this.selected !== index;
    const tableCellChanged =
      (previous === null) !== (cell === null)
      || (previous !== null && cell !== null && (
        previous.table !== cell.table || previous.row !== cell.row || previous.col !== cell.col
      ));
    this.selected = index;
    this.selectedTableCellInfo = cell;
    if (!paragraphChanged && !tableCellChanged) return;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (EventClass && paragraphChanged) {
      this.root.dispatchEvent(new EventClass('docx-selectionchange', { bubbles: true, detail: { index } }));
    }
    if (EventClass && tableCellChanged) {
      this.root.dispatchEvent(new EventClass('docx-tablecellchange', {
        bubbles: true,
        detail: cell ? { cell: this.selectedTableCell } : null,
      }));
    }
  }

  private selectImage(image: ImageInfo | null): void {
    this.selectedImageInfo = image;
    const EventClass = this.root.ownerDocument.defaultView?.CustomEvent;
    if (EventClass) this.root.dispatchEvent(new EventClass('docx-imageselectionchange', { bubbles: true, detail: image ? { image } : null }));
  }

  private readonly handleSelection = (): void => {
    const selection = this.root.ownerDocument.getSelection();
    const node = selection?.anchorNode;
    if (!node || !this.root.contains(node)) {
      this.updateRangeSelection(null);
      return;
    }
    this.expandFieldSelection(selection);
    const element = node.nodeType === 1 ? node as Element : node.parentElement;
    const paragraph = element?.closest<HTMLElement>('[data-paragraph]');
    if (paragraph && this.root.contains(paragraph)) this.selectParagraph(Number(paragraph.dataset.paragraph));
    const image = element?.closest<HTMLElement>('[data-image]');
    if (!image) this.selectImage(null);
    this.updateRangeSelection(this.captureDocumentRange());
  };

  private expandFieldSelection(selection: Selection): void {
    if (!selection.rangeCount) return;
    const range = selection.getRangeAt(0);
    const fields = Array.from(this.root.querySelectorAll<HTMLElement>('[data-docx-field-role="result"]'));
    const anchor = selection.anchorNode;
    const anchorElement = anchor?.nodeType === 1 ? anchor as Element : anchor?.parentElement;
    const touched = fields.filter((node) => range.intersectsNode(node) &&
      (!range.collapsed || anchorElement?.closest('[data-docx-field-role="result"]') === node));
    if (!touched.length) return;
    const ids = new Set(touched.map((node) => node.dataset.docxField));
    const selected = fields.filter((node) => ids.has(node.dataset.docxField));
    const expanded = range.cloneRange();
    const first = selected[0]!;
    const last = selected[selected.length - 1]!;
    const bounds = range.cloneRange();
    bounds.selectNode(first);
    if (range.compareBoundaryPoints(Range.START_TO_START, bounds) > 0) expanded.setStartBefore(first);
    bounds.selectNode(last);
    if (range.compareBoundaryPoints(Range.END_TO_END, bounds) < 0) expanded.setEndAfter(last);
    if (expanded.startContainer === range.startContainer && expanded.startOffset === range.startOffset &&
        expanded.endContainer === range.endContainer && expanded.endOffset === range.endOffset) return;
    selection.removeAllRanges();
    selection.addRange(expanded);
  }

  private selectionTouchesField(content: HTMLElement): boolean {
    const selection = this.root?.ownerDocument.getSelection?.();
    if (!selection?.rangeCount) return false;
    const range = selection.getRangeAt(0);
    const anchor = selection.anchorNode;
    const element = anchor?.nodeType === 1 ? anchor as Element : anchor?.parentElement;
    return Array.from(content?.querySelectorAll?.<HTMLElement>('[data-docx-field]') ?? [])
      .some((node) => range.intersectsNode(node) &&
        (!range.collapsed || element === node || node.contains(element ?? null)));
  }

  private readonly handleRootKeydown = (event: KeyboardEvent): void => {
    if (this.handleHistoryShortcut(event)) return;
    if (!this.selectedImageInfo || !['Delete', 'Backspace'].includes(event.key)) return;
    const active = this.root.ownerDocument.activeElement as HTMLElement | null;
    const image = active?.closest('[data-image]') as HTMLElement | null;
    if (!image || image.dataset.image !== this.selectedImageInfo.id) return;
    event.preventDefault();
    this.document.deleteImage(this.selectedImageInfo);
    this.render();
    this.options.onChange?.(this.document.getSnapshot());
  };

  private handleHistoryShortcut(event: Pick<KeyboardEvent, 'isComposing' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey' | 'key' | 'preventDefault' | 'target'>): boolean {
    if (event.isComposing || this.composing || !(event.ctrlKey || event.metaKey) || event.altKey) return false;
    const node = event.target as Node | null;
    const target = node?.nodeType === 1 ? node as Element : node?.parentElement ?? null;
    const content = target?.closest<HTMLElement>('.docx-paragraph-content');
    if (!content || !this.root.contains(content)) return false;
    const key = event.key.toLowerCase();
    if (key === 'z' && !event.shiftKey) {
      event.preventDefault();
      this.applyHistory('undo');
      return true;
    }
    if (key === 'y' || (key === 'z' && event.shiftKey)) {
      event.preventDefault();
      this.applyHistory('redo');
      return true;
    }
    return false;
  }

  private applyHistory(direction: 'undo' | 'redo'): void {
    if (this.destroyed) return;
    const previousSelected = this.selected;
    this.flush();
    const beforeRevision = this.document.revision;
    const snapshot = direction === 'undo' ? this.document.undo() : this.document.redo();
    if (snapshot.revision === beforeRevision) return;
    this.render();
    const fallback = snapshot.paragraphs[0]?.index ?? null;
    const targetIndex = previousSelected !== null && this.paragraphs.has(previousSelected) ? previousSelected : fallback;
    if (targetIndex !== null) {
      const paragraph = this.paragraphs.get(targetIndex)?.content;
      if (paragraph) this.focusContent(paragraph);
    }
    this.options.onChange?.(snapshot);
  }

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
