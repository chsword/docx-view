import { DocxDocument, DocxEditor, WORD_NS, linearToMathMl } from '../src/index.js';
import { contentTypeForExtension, decodeBase64 } from '../src/index.js';
import type {
  AgentRequest,
  DocumentRange,
  DocumentSnapshot,
  HyperlinkInfo,
  ImageInfo,
  MathMlNode,
  OutlineNode,
  ParagraphFormat,
  ReviewerFilterAuthor,
  ReviewerInfo,
  RunFormat,
  StyleInfo,
  TableCellLocation,
} from '../src/index.js';
import { reviewerBucketKey, reviewerBucketLabel, reviewerBucketOf } from '../src/revisions.js';
import { findReusableNumberingId } from '../src/numbering.js';
import { RibbonContextState } from './context-tabs.js';
import {
  createCommandRegistry,
  createCommandContextBuilder,
  createExampleCommandDescriptors,
  getCommandControlState,
  type CommandContext,
} from './commands.js';
import { initializeRibbon } from './ribbon.js';
import { initializeContextMenu } from './context-menu.js';
import { styleFormValues, stylePatchFromForm, uniqueStyleId, withoutNulls, type StyleFormValues, type TriState } from './style-editor.js';
import './style.css';

const SAMPLE_IMAGE = decodeBase64('iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAQAAAD8fJRsAAAAC0lEQVR42mP8/x8AAusB9WnM0iEAAAAASUVORK5CYII=');
const CLEAR_RUN_FORMAT: RunFormat = {
  style: null,
  bold: null,
  italic: null,
  underline: null,
  underlineStyle: null,
  underlineColor: null,
  fontSize: null,
  fontFamily: null,
  fontFamilyEastAsia: null,
  color: null,
  strike: null,
  doubleStrike: null,
  verticalAlign: null,
  smallCaps: null,
  allCaps: null,
  highlight: null,
  characterSpacing: null,
  border: null,
  shading: null,
};
const CLEAR_PARAGRAPH_FORMAT: ParagraphFormat = {
  alignment: null,
  indentLeft: null,
  indentRight: null,
  indentFirstLine: null,
  indentHanging: null,
  spacingBefore: null,
  spacingAfter: null,
  lineSpacing: null,
  lineSpacingRule: null,
  keepNext: null,
  keepLines: null,
  pageBreakBefore: null,
  widowControl: null,
  suppressLineNumbers: null,
  suppressAutoHyphens: null,
  outlineLevel: null,
  tabs: null,
  borders: null,
  shading: null,
};

function element<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`找不到界面元素：${id}`);
  return node as T;
}

const ribbonContextState = new RibbonContextState(
  ['ribbon-tab-home', 'ribbon-tab-insert', 'ribbon-tab-layout', 'ribbon-tab-review', 'ribbon-tab-view'],
  ['ribbon-tab-table', 'ribbon-tab-image'],
);
const ribbon = initializeRibbon(element('ribbon'), {
  onManualActivate: (tabId) => ribbonContextState.manualActivate(tabId),
  onContextActivate: (tabId) => ribbonContextState.activateContext(tabId),
});

const recentNumbering = new Map<'bullet' | 'decimal', number>();

function applyNumbering(kind: 'bullet' | 'decimal'): void {
  const index = selectedIndex();
  const paragraph = doc.getParagraphs().find((item) => item.index === index)!;
  const numId = findReusableNumberingId(doc.getParagraphs(), doc.getNumberingDefinitions(), kind, index, recentNumbering.get(kind)) ?? doc.createNumbering(kind);
  doc.setParagraphNumbering(index, numId, paragraph.numbering?.level ?? 0);
  recentNumbering.set(kind, numId);
  editor.render();
  refresh();
  message(kind === 'bullet' ? '已应用项目符号列表。' : '已应用编号列表。');
}

function changeNumberingLevel(delta: number): void {
  const index = selectedIndex();
  doc.setParagraphLevel(index, delta);
  editor.render();
  refresh();
  message(delta > 0 ? '已提高列表级别。' : '已降低列表级别。');
}

let doc = createSample();
let filename = '产品计划.docx';
let xmlRevision = -1;
let imageAction: 'insert' | 'replace' = 'insert';
let imageActionTarget: ImageInfo | null = null;
let selectedRange: DocumentRange | null = null;
let selectedRangeFormat: RunFormat | null = null;
let selectedCommentId: number | null = null;
let selectedRevisionId: number | null = null;
let formatPainter: { format: RunFormat; locked: boolean } | null = null;
let applyingFormatPainter = false;
let selectedReviewerAuthors: ReviewerFilterAuthor[] | undefined;
let reviewRevisionView: 'final' | 'original' | 'markup' = 'markup';
let reviewShowRevisions = true;
let reviewShowComments = true;
let outlineDrag: { paragraph: number; end: number } | null = null;
const status = element('status');
const host = element('editor');
const agentInput = element<HTMLTextAreaElement>('agent-input');
const xmlInput = element<HTMLTextAreaElement>('xml-input');
const xmlPart = element<HTMLSelectElement>('xml-part');
const editor = new DocxEditor(host, doc, {
  onChange: refresh,
  reviewFilter: { revisionView: reviewRevisionView, showComments: reviewShowComments, showRevisions: reviewShowRevisions },
});

function reviewerSelectionKey(author: ReviewerFilterAuthor): string {
  return reviewerBucketKey(author);
}

function reviewerSelectionFrom(reviewer: Pick<ReviewerInfo, 'author' | 'kind'>): ReviewerFilterAuthor {
  return reviewer.kind === 'named'
    ? { kind: reviewer.kind, author: reviewer.author ?? '' }
    : reviewer.kind === 'blank'
      ? { kind: reviewer.kind, author: reviewer.author ?? ' ' }
      : reviewer.kind === 'empty'
        ? { kind: reviewer.kind, author: '' }
        : { kind: reviewer.kind };
}

function applyReviewFilter(): void {
  editor.setReviewFilter({
    ...(selectedReviewerAuthors?.length ? { authors: selectedReviewerAuthors } : {}),
    showComments: reviewShowComments,
    showRevisions: reviewShowRevisions,
    revisionView: reviewRevisionView,
  });
}

function isMarkupView(): boolean {
  return reviewRevisionView === 'markup';
}

function ensureMarkupView(): void {
  if (!isMarkupView()) throw new Error('请先切回“标记视图”后再编辑样式或大纲。');
}

function defaultParagraphStyle(): StyleInfo | undefined {
  return doc.getStyles().find((style) => style.type === 'paragraph' && style.isDefault)
    ?? doc.getStyles().find((style) => style.type === 'paragraph');
}

function toolbarParagraphStyles(): StyleInfo[] {
  const styles = doc.getStyleGallery().filter((style) => style.type === 'paragraph');
  const defaultStyle = defaultParagraphStyle();
  if (!defaultStyle || styles.some((style) => style.id === defaultStyle.id)) return styles;
  return [defaultStyle, ...styles];
}

function flattenOutline(nodes: OutlineNode[]): OutlineNode[] {
  const result: OutlineNode[] = [];
  const walk = (items: OutlineNode[]) => {
    for (const item of items) {
      result.push(item);
      walk(item.children);
    }
  };
  walk(nodes);
  return result;
}

function outlineSectionEnd(paragraph: number, flat = flattenOutline(doc.getOutline())): number {
  const node = flat.find((entry) => entry.paragraph === paragraph);
  if (!node) return paragraph;
  const next = flat.find((entry) => entry.paragraph > paragraph && entry.level <= node.level);
  return next ? next.paragraph - 1 : doc.getParagraphs().at(-1)?.index ?? paragraph;
}

function moveOutlineWithKeyboard(paragraph: number, direction: -1 | 1): void {
  ensureMarkupView();
  const outline = flattenOutline(doc.getOutline());
  const index = outline.findIndex((entry) => entry.paragraph === paragraph);
  if (index === -1) return;
  const current = outline[index]!;
  const currentEnd = outlineSectionEnd(current.paragraph, outline);
  if (direction < 0) {
    const previous = outline[index - 1];
    if (!previous) return;
    doc.moveOutlineSection(current.paragraph, previous.paragraph);
    editor.render();
    refresh();
    focusParagraph(previous.paragraph);
    message('已上移当前标题。');
    return;
  }
  const next = outline[index + 1];
  if (!next) return;
  const target = outlineSectionEnd(next.paragraph, outline) + 1;
  doc.moveOutlineSection(current.paragraph, target);
  const movedParagraph = target > currentEnd ? target - (currentEnd - current.paragraph + 1) : target;
  editor.render();
  refresh();
  focusParagraph(movedParagraph);
  message('已下移当前标题。');
}

function setOutlineLevelWithKeyboard(paragraph: number, delta: -1 | 1): void {
  ensureMarkupView();
  const node = flattenOutline(doc.getOutline()).find((entry) => entry.paragraph === paragraph);
  if (!node) return;
  doc.setOutlineLevel(paragraph, Math.max(0, Math.min(8, node.level + delta)));
  editor.render();
  refresh();
  focusParagraph(paragraph);
  message(delta < 0 ? '已提升标题层级。' : '已降低标题层级。');
}

function focusParagraph(paragraph: number): void {
  const target = host.querySelector<HTMLElement>(`.docx-paragraph[data-paragraph="${paragraph}"] .docx-paragraph-content`);
  if (!target) return;
  target.scrollIntoView({ block: 'nearest' });
  target.focus();
}

function createSample(): DocxDocument {
  const sample = DocxDocument.create();
  sample.defineStyle({ id: 'Title', name: '标题', type: 'paragraph', quickFormat: true, paragraph: { spacingAfter: 160 }, run: { bold: true, fontSize: 26, color: '223855' } });
  sample.defineStyle({ id: 'Heading1', name: '标题 1', type: 'paragraph', quickFormat: true, paragraph: { spacingBefore: 120, spacingAfter: 80 }, run: { bold: true, fontSize: 16, color: '3567D6' } });
  sample.defineStyle({ id: 'BodyText', name: '正文', type: 'paragraph', quickFormat: true, paragraph: { spacingAfter: 80 }, run: { fontSize: 11, color: '394A61' } });
  sample.defineStyle({ id: 'Quote', name: '引用', type: 'paragraph', quickFormat: true, paragraph: { indentLeft: 360, spacingAfter: 120 }, run: { italic: true, color: '788597' } });
  sample.setParagraphText(0, '把想法，写成下一步。');
  sample.formatParagraph(0, { style: 'Title' }, { validateStyle: true });
  sample.insertParagraph('纸间工作室  /  产品共创计划  /  2026');
  sample.formatRun(1, 0, { fontSize: 10, color: '788597' });
  sample.insertParagraph('01  项目愿景');
  sample.formatParagraph(2, { style: 'Heading1' }, { validateStyle: true });
  sample.insertParagraph('让每一份文档都能自由流转。我们希望把熟悉的文字编辑，与透明的文档结构、可靠的自动化连接起来。');
  sample.formatParagraph(3, { style: 'BodyText' }, { validateStyle: true });
  sample.insertParagraph('点击任意段落开始编辑，也可以在右侧运行一组 Agent 指令。所有处理都发生在你的浏览器里，文件不会上传。');
  sample.formatParagraph(4, { style: 'BodyText' }, { validateStyle: true });
  sample.insertImage({ bytes: SAMPLE_IMAGE, contentType: 'image/png', paragraph: 3, alt: '示例图片' });
  sample.insertParagraph('02  从一个小计划开始');
  sample.formatParagraph(5, { style: 'Heading1' }, { validateStyle: true });
  const bullet = sample.createNumbering('bullet');
  sample.insertParagraph('梳理文档结构与保留策略');
  sample.setParagraphNumbering(6, bullet);
  sample.insertParagraph('验证浏览器内可视化编辑体验');
  sample.setParagraphNumbering(7, bullet);
  const decimal = sample.createNumbering('multilevel');
  sample.insertParagraph('第一阶段：实现编号与项目符号');
  sample.setParagraphNumbering(8, decimal, 0);
  sample.insertParagraph('解析 numbering.xml 与多级编号');
  sample.setParagraphNumbering(9, decimal, 1);
  sample.insertTable([['阶段', '交付内容', '状态'], ['探索', '梳理需求与文档结构', '已完成'], ['共创', '编辑体验与自动化接口', '进行中'], ['发布', '验证 DOCX 导出与兼容性', '下一步']]);
  sample.formatTable(0, {
    layout: 'fixed',
    width: { type: 'pct', value: 5000 },
    borders: {
      top: { style: 'single', size: 8, color: '9AA8BA' },
      right: { style: 'single', size: 8, color: '9AA8BA' },
      bottom: { style: 'single', size: 8, color: '9AA8BA' },
      left: { style: 'single', size: 8, color: '9AA8BA' },
      insideH: { style: 'single', size: 8, color: 'D7DFEA' },
      insideV: { style: 'single', size: 8, color: 'D7DFEA' },
    },
  });
  sample.formatTableRow(0, 0, { header: true, height: { value: 520, rule: 'atLeast' } });
  sample.formatCell(0, 0, 0, { shading: { fill: 'EEF3FF' }, verticalAlign: 'center' });
  sample.insertParagraph('好的工具，让内容成为主角。');
  const last = sample.getParagraphs().at(-1)!;
  sample.formatParagraph(last.index, { style: 'Quote' }, { validateStyle: true });
  sample.insertParagraph('域结果（只读）：');
  sample.insertField(sample.getParagraphs().at(-1)!.index, ' SEQ 演示 ', '1');
  return sample;
}

function message(text: string, error = false): void {
  status.textContent = text;
  status.classList.toggle('error', error);
  status.setAttribute('role', error ? 'alert' : 'status');
}

function reportError(error: unknown): void {
  message(error instanceof Error ? error.message : String(error), true);
}

function run(action: () => void | Promise<void>): void {
  void Promise.resolve().then(action).catch(reportError);
}

function refresh(snapshot: DocumentSnapshot = doc.getSnapshot()): void {
  element('revision').textContent = String(snapshot.revision);
  element('paragraph-count').textContent = String(snapshot.paragraphs.length);
  element('snapshot-output').textContent = JSON.stringify(snapshot, null, 2);
  let orientation: 'portrait' | 'landscape' = 'portrait';
  let pageSize: 'A4' | 'Letter' = 'A4';
  try {
    const first = doc.getSection(0);
    orientation = first.orientation;
    const longEdge = Math.max(first.pageWidth, first.pageHeight);
    pageSize = longEdge > 16300 ? 'A4' : 'Letter';
  } catch {
    orientation = 'portrait';
    pageSize = 'A4';
  }
  element<HTMLSelectElement>('page-orientation').value = orientation;
  element<HTMLSelectElement>('page-size').value = pageSize;
  loadStyleOptions();
  refreshOutline();
  updateSelection();
  updateImageSelection();
  refreshReviewers();
  refreshComments(snapshot);
  refreshRevisions();
}

function loadStyleOptions(): void {
  const select = element<HTMLSelectElement>('paragraph-style');
  const previous = select.value;
  select.replaceChildren(...toolbarParagraphStyles().map((style) => {
    const option = document.createElement('option');
    option.value = style.id;
    option.textContent = style.name;
    return option;
  }));
  if (Array.from(select.options).some((option) => option.value === previous)) select.value = previous;
}

function refreshOutline(): void {
  const list = element<HTMLUListElement>('outline-list');
  const outline = flattenOutline(doc.getOutline());
  if (!outline.length) {
    const empty = document.createElement('li');
    empty.className = 'hint';
    empty.textContent = '当前文档还没有标题层级。';
    list.replaceChildren(empty);
    return;
  }
  const editable = isMarkupView();
  list.replaceChildren(...outline.map((node) => {
    const item = document.createElement('li');
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'outline-item';
    if (editor.selectedParagraph === node.paragraph) button.classList.add('active');
    button.draggable = editable;
    button.style.paddingLeft = `${10 + node.level * 18}px`;
    button.dataset.paragraph = String(node.paragraph);
    button.dataset.level = String(node.level);
    button.append(
      Object.assign(document.createElement('span'), { className: 'outline-level', textContent: `H${node.level + 1}` }),
      Object.assign(document.createElement('span'), { className: 'outline-text', textContent: node.text || '（空标题）' }),
    );
    button.addEventListener('click', () => focusParagraph(node.paragraph));
    button.addEventListener('keydown', (event) => run(() => {
      if (!editable || !event.altKey) return;
      if (event.key === 'ArrowUp') {
        event.preventDefault();
        moveOutlineWithKeyboard(node.paragraph, -1);
      } else if (event.key === 'ArrowDown') {
        event.preventDefault();
        moveOutlineWithKeyboard(node.paragraph, 1);
      } else if (event.key === 'ArrowLeft') {
        event.preventDefault();
        setOutlineLevelWithKeyboard(node.paragraph, -1);
      } else if (event.key === 'ArrowRight') {
        event.preventDefault();
        setOutlineLevelWithKeyboard(node.paragraph, 1);
      }
    }));
    button.addEventListener('dragstart', (event) => {
      if (!editable) return;
      outlineDrag = { paragraph: node.paragraph, end: outlineSectionEnd(node.paragraph, outline) };
      event.dataTransfer?.setData('text/plain', String(node.paragraph));
      event.dataTransfer!.effectAllowed = 'move';
    });
    button.addEventListener('dragend', () => { outlineDrag = null; button.classList.remove('drag-over'); });
    button.addEventListener('dragover', (event) => {
      if (!editable || !outlineDrag) return;
      event.preventDefault();
      button.classList.add('drag-over');
      event.dataTransfer!.dropEffect = 'move';
    });
    button.addEventListener('dragleave', () => button.classList.remove('drag-over'));
    button.addEventListener('drop', (event) => run(() => {
      button.classList.remove('drag-over');
      if (!editable || !outlineDrag) return;
      ensureMarkupView();
      event.preventDefault();
      const targetParagraph = node.paragraph;
      let movedParagraph = outlineDrag.paragraph;
      if (targetParagraph !== outlineDrag.paragraph) {
        doc.moveOutlineSection(outlineDrag.paragraph, targetParagraph);
        movedParagraph = targetParagraph > outlineDrag.end
          ? targetParagraph - (outlineDrag.end - outlineDrag.paragraph + 1)
          : targetParagraph;
      }
      const rect = button.getBoundingClientRect();
      const targetLevel = Math.max(0, Math.min(8, Math.floor(Math.max(0, event.clientX - rect.left - 12) / 24)));
      doc.setOutlineLevel(movedParagraph, targetLevel);
      outlineDrag = null;
      editor.render();
      refresh();
      focusParagraph(movedParagraph);
      message('已更新大纲顺序/层级。');
    }));
    item.append(button);
    return item;
  }));
}

function updateSelection(): void {
  const index = editor.selectedParagraph;
  const paragraph = doc.getParagraphs().find((item) => item.index === index);
  const hasRange = !!selectedRange;
  element('selection-label').textContent = hasRange
    ? `已选择范围：第 ${selectedRange!.start.paragraph + 1} 段 ${selectedRange!.start.offset} 到 第 ${selectedRange!.end.paragraph + 1} 段 ${selectedRange!.end.offset}`
    : paragraph ? `已选择第 ${paragraph.index + 1} 段` : '点击正文选择段落';
  const size = element<HTMLSelectElement>('font-size');
  const color = element<HTMLInputElement>('font-color');
  const style = element<HTMLSelectElement>('paragraph-style');
  const alignment = element<HTMLSelectElement>('alignment');
  size.value = selectedRangeFormat?.fontSize ? String(selectedRangeFormat.fontSize) : '';
  const runColor = selectedRangeFormat?.color;
  color.value = runColor && /^[0-9a-f]{6}$/i.test(runColor) ? `#${runColor}` : '#25334a';
  const currentStyle = paragraph?.style ? doc.getStyle(paragraph.style) : defaultParagraphStyle();
  if (currentStyle && !Array.from(style.options).some((option) => option.value === currentStyle.id)) {
    style.append(Object.assign(document.createElement('option'), { value: currentStyle.id, textContent: currentStyle.name }));
  }
  style.value = currentStyle?.id ?? '';
  alignment.value = paragraph?.effective?.alignment ?? paragraph?.alignment ?? 'left';
  element('effective-format').textContent = paragraph ? JSON.stringify(paragraph.effective ?? {}, null, 2) : '点击正文选择段落';
  updateTableSelection(editor.selectedTableCell, false);
  syncCommandState();
}

function updateImageSelection(): void {
  const image = editor.selectedImage;
  setContextTabVisible('ribbon-tab-image', image !== null);
  element('image-selection-label').textContent = image ? `已选择图片 · ${image.alt ?? image.name ?? image.relationshipId}` : '未选中图片';
  element<HTMLInputElement>('image-alt').value = image?.alt ?? '';
  syncCommandState();
}

function setContextTabVisible(tabId: string, visible: boolean): void {
  const activateTabId = ribbonContextState.setContextVisible(tabId, visible);
  ribbon.setTabVisible(tabId, visible);
  if (activateTabId) ribbon.activate(activateTabId);
}

function updateTableSelection(cell: TableCellLocation | null = editor.selectedTableCell, syncCommands = true): void {
  setContextTabVisible('ribbon-tab-table', cell !== null);
  element<HTMLElement>('table-context-hint').hidden = !cell?.nested;
  if (syncCommands) syncCommandState();
}

function refreshComments(snapshot: DocumentSnapshot = doc.getSnapshot()): void {
  const authorFilter = element<HTMLInputElement>('comment-author-filter').value.trim();
  const resolvedFilter = element<HTMLSelectElement>('comment-resolved-filter').value;
  const list = element<HTMLUListElement>('comment-list');
  const selectedAuthors = selectedReviewerAuthors ? new Set(selectedReviewerAuthors.map(reviewerSelectionKey)) : undefined;
  const comments = (reviewShowComments ? snapshot.comments : []).filter((comment) => {
    if (selectedAuthors && !selectedAuthors.has(reviewerBucketKey(reviewerBucketOf(comment.author)))) return false;
    if (authorFilter && !(comment.author ?? '').includes(authorFilter)) return false;
    if (resolvedFilter === 'resolved' && comment.resolved !== true) return false;
    if (resolvedFilter === 'open' && comment.resolved === true) return false;
    return true;
  });
  if (selectedCommentId !== null && !comments.some((comment) => comment.id === selectedCommentId)) selectedCommentId = null;
  list.replaceChildren(...comments.map((comment) => {
    const item = document.createElement('li');
    const button = document.createElement('button');
    button.className = selectedCommentId === comment.id ? 'active' : '';
    button.addEventListener('click', () => {
      selectedCommentId = comment.id;
      refreshComments();
    });
    const meta = document.createElement('div');
    meta.className = 'comment-meta';
    meta.textContent = `${comment.author ?? '匿名'} · #${comment.id}${comment.resolved ? ' · 已解决' : ''}${comment.parentId ? ` · 回复 #${comment.parentId}` : ''}`;
    const body = document.createElement('div');
    body.className = 'comment-body';
    body.textContent = comment.text || '(空批注)';
    button.append(meta, body);
    item.append(button);
    return item;
  }));
  for (const node of host.querySelectorAll<HTMLElement>('[data-docx-comment-ids]')) {
    const ids = (node.dataset.docxCommentIds ?? '').split(',').map((value) => Number(value));
    node.classList.toggle('docx-comment-active', selectedCommentId !== null && ids.includes(selectedCommentId));
  }
  syncCommandState();
}

function revisionSummaryText(revision: ReturnType<DocxDocument['getRevisions']>[number]): string {
  if (revision.kind === 'deletion') return revision.deletedText ? `删除：“${revision.deletedText}”` : '删除';
  if (revision.kind === 'insertion') return '插入';
  if (revision.kind === 'runFormatChange') return '文字格式修订';
  if (revision.kind === 'paragraphFormatChange') return '段落格式修订';
  if (revision.kind === 'tableFormatChange') return '表格格式修订';
  if (revision.kind === 'sectionFormatChange') return '页面设置修订';
  if (revision.kind === 'rowFormatChange') return '行格式修订';
  return '单元格格式修订';
}

function visibleRevisions() {
  const selectedAuthors = selectedReviewerAuthors ? new Set(selectedReviewerAuthors.map(reviewerSelectionKey)) : undefined;
  return (reviewShowRevisions ? doc.getRevisions() : []).filter((revision) =>
    !selectedAuthors || selectedAuthors.has(reviewerBucketKey(reviewerBucketOf(revision.author))));
}

function refreshRevisions(): void {
  const list = element<HTMLUListElement>('revision-list');
  const revisions = visibleRevisions();
  if (selectedRevisionId !== null && !revisions.some((revision) => revision.id === selectedRevisionId)) selectedRevisionId = null;
  list.replaceChildren(...revisions.map((revision) => {
    const item = document.createElement('li');
    const button = document.createElement('button');
    button.className = selectedRevisionId === revision.id ? 'active' : '';
    button.addEventListener('click', () => {
      selectedRevisionId = revision.id;
      editor.focusRevision(revision.id);
      refreshRevisions();
    });
    const meta = document.createElement('div');
    meta.className = 'comment-meta';
    meta.textContent = `${revision.author ?? '匿名'} · #${revision.id} · 第 ${revision.paragraph + 1} 段${revision.run !== undefined ? ` / run ${revision.run}` : ''}`;
    const body = document.createElement('div');
    body.className = 'comment-body';
    body.textContent = revisionSummaryText(revision);
    const actions = document.createElement('div');
    actions.className = 'comment-panel-actions';
    const accept = document.createElement('button');
    accept.type = 'button';
    accept.className = 'button subtle';
    accept.textContent = '接受';
    accept.addEventListener('click', (event) => {
      event.stopPropagation();
      run(() => {
        editor.acceptRevision(revision.id);
        selectedRevisionId = null;
        refresh();
      });
    });
    const reject = document.createElement('button');
    reject.type = 'button';
    reject.className = 'button subtle';
    reject.textContent = '拒绝';
    reject.addEventListener('click', (event) => {
      event.stopPropagation();
      run(() => {
        editor.rejectRevision(revision.id);
        selectedRevisionId = null;
        refresh();
      });
    });
    actions.append(accept, reject);
    button.append(meta, body, actions);
    item.append(button);
    return item;
  }));
  syncCommandState();
}

function refreshReviewers(): void {
  const reviewers = doc.getReviewers();
  const list = element<HTMLUListElement>('reviewer-list');
  const previousSelection = JSON.stringify((selectedReviewerAuthors ?? []).map(reviewerSelectionKey));
  const available = new Set(reviewers.map((reviewer) => reviewerBucketKey(reviewerSelectionFrom(reviewer))));
  if (selectedReviewerAuthors?.length) {
    selectedReviewerAuthors = selectedReviewerAuthors.filter((author) => available.has(reviewerSelectionKey(author)));
    if (!selectedReviewerAuthors.length) selectedReviewerAuthors = undefined;
  }
  if (previousSelection !== JSON.stringify((selectedReviewerAuthors ?? []).map(reviewerSelectionKey))) applyReviewFilter();
  const selected = new Set((selectedReviewerAuthors ?? reviewers.map((reviewer) => reviewerSelectionFrom(reviewer))).map(reviewerSelectionKey));
  list.replaceChildren(...reviewers.map((reviewer) => {
    const reviewerAuthor = reviewerSelectionFrom(reviewer);
    const key = reviewerSelectionKey(reviewerAuthor);
    const item = document.createElement('li');
    const label = document.createElement('label');
    label.className = 'reviewer-item';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = selected.has(key);
    checkbox.addEventListener('change', () => {
      const allEntries = reviewers.map((entry) => reviewerSelectionFrom(entry));
      const selectedEntries = new Map((selectedReviewerAuthors ?? allEntries).map((entry) => [reviewerSelectionKey(entry), entry]));
      if (checkbox.checked) selectedEntries.set(key, reviewerAuthor);
      else selectedEntries.delete(key);
      selectedReviewerAuthors = selectedEntries.size === reviewers.length ? undefined : [...selectedEntries.values()];
      applyReviewFilter();
      refreshComments();
      refreshRevisions();
    });
    const name = document.createElement('span');
    name.textContent = `${reviewerBucketLabel(reviewerSelectionFrom(reviewer))}${reviewer.initials ? ` (${reviewer.initials})` : ''}`;
    const count = document.createElement('span');
    count.className = 'reviewer-count';
    count.textContent = `修订 ${reviewer.revisionCount} · 批注 ${reviewer.commentCount} · 未解决 ${reviewer.unresolvedCommentCount}`;
    label.append(checkbox, name, count);
    item.append(label);
    return item;
  }));
}

let styleDialogMode: 'modify' | 'new' = 'modify';
let styleDialogInitial: StyleFormValues | null = null;
let styleDialogParagraph: number | null = null;

const BLANK_STYLE_FORM: StyleFormValues = {
  name: '', basedOn: '', next: '', fontFamily: '', fontSize: '', bold: '', italic: '', underline: '', color: '',
  alignment: '', spacingBefore: '', spacingAfter: '', lineSpacing: '', indentLeft: '', indentFirstLine: '',
  quickFormat: true,
};

function styleOption(value: string, label: string): HTMLOptionElement {
  return Object.assign(document.createElement('option'), { value, textContent: label });
}

function styleTypeLabel(style: StyleInfo): string {
  return style.type === 'character' ? '字符' : '段落';
}

/** 「基于」只能选同类型的样式；成环的选择由 `updateStyle()` 拒绝并在对话框里报出来。 */
function fillStyleRelations(type: StyleInfo['type'], self: string | null): void {
  const styles = doc.getStyles();
  element<HTMLSelectElement>('style-dialog-based-on').replaceChildren(styleOption('', '（无）'),
    ...styles.filter((style) => style.type === type && style.id !== self).map((style) => styleOption(style.id, style.name)));
  element<HTMLSelectElement>('style-dialog-next').replaceChildren(styleOption('', '（同本样式）'),
    ...styles.filter((style) => style.type === 'paragraph').map((style) => styleOption(style.id, style.name)));
}

function writeStyleForm(values: StyleFormValues, type: StyleInfo['type']): void {
  element<HTMLInputElement>('style-dialog-name').value = values.name;
  element<HTMLSelectElement>('style-dialog-based-on').value = values.basedOn;
  element<HTMLSelectElement>('style-dialog-next').value = values.next;
  element<HTMLSelectElement>('style-dialog-next').disabled = type !== 'paragraph';
  element<HTMLInputElement>('style-dialog-font-family').value = values.fontFamily;
  element<HTMLInputElement>('style-dialog-font-size').value = values.fontSize;
  element<HTMLSelectElement>('style-dialog-bold').value = values.bold;
  element<HTMLSelectElement>('style-dialog-italic').value = values.italic;
  element<HTMLSelectElement>('style-dialog-underline').value = values.underline;
  element<HTMLInputElement>('style-dialog-color').value = values.color;
  element<HTMLInputElement>('style-dialog-color-picker').value = `#${values.color || '25334A'}`;
  element<HTMLSelectElement>('style-dialog-alignment').value = values.alignment;
  element<HTMLInputElement>('style-dialog-line-spacing').value = values.lineSpacing;
  element<HTMLInputElement>('style-dialog-spacing-before').value = values.spacingBefore;
  element<HTMLInputElement>('style-dialog-spacing-after').value = values.spacingAfter;
  element<HTMLInputElement>('style-dialog-indent-left').value = values.indentLeft;
  element<HTMLInputElement>('style-dialog-indent-first-line').value = values.indentFirstLine;
  element<HTMLInputElement>('style-dialog-quick-format').checked = values.quickFormat;
  element('style-dialog-paragraph').hidden = type === 'character';
}

function readStyleForm(): StyleFormValues {
  return {
    name: element<HTMLInputElement>('style-dialog-name').value,
    basedOn: element<HTMLSelectElement>('style-dialog-based-on').value,
    next: element<HTMLSelectElement>('style-dialog-next').value,
    fontFamily: element<HTMLInputElement>('style-dialog-font-family').value,
    fontSize: element<HTMLInputElement>('style-dialog-font-size').value,
    bold: element<HTMLSelectElement>('style-dialog-bold').value as TriState,
    italic: element<HTMLSelectElement>('style-dialog-italic').value as TriState,
    underline: element<HTMLSelectElement>('style-dialog-underline').value as TriState,
    color: element<HTMLInputElement>('style-dialog-color').value,
    alignment: element<HTMLSelectElement>('style-dialog-alignment').value as StyleFormValues['alignment'],
    lineSpacing: element<HTMLInputElement>('style-dialog-line-spacing').value,
    spacingBefore: element<HTMLInputElement>('style-dialog-spacing-before').value,
    spacingAfter: element<HTMLInputElement>('style-dialog-spacing-after').value,
    indentLeft: element<HTMLInputElement>('style-dialog-indent-left').value,
    indentFirstLine: element<HTMLInputElement>('style-dialog-indent-first-line').value,
    quickFormat: element<HTMLInputElement>('style-dialog-quick-format').checked,
  };
}

function loadStyleIntoDialog(id: string): void {
  const style = doc.getStyle(id);
  if (!style) throw new Error(`找不到样式 ${id}。`);
  fillStyleRelations(style.type, style.id);
  styleDialogInitial = styleFormValues(style);
  writeStyleForm(styleDialogInitial, style.type);
  const remove = element<HTMLButtonElement>('style-dialog-delete');
  remove.disabled = style.isDefault === true;
  remove.title = style.isDefault ? '默认样式不能删除' : '';
}

function showStyleDialogError(text: string | null): void {
  const error = element('style-dialog-error');
  error.hidden = !text;
  error.textContent = text ?? '';
}

function openStyleDialog(mode: 'modify' | 'new'): void {
  ensureMarkupView();
  editor.flush();
  styleDialogMode = mode;
  styleDialogParagraph = editor.selectedParagraph;
  const paragraph = styleDialogParagraph === null ? undefined
    : doc.getParagraphs().find((item) => item.index === styleDialogParagraph);
  const currentStyle = (paragraph?.style ? doc.getStyle(paragraph.style) : undefined) ?? defaultParagraphStyle();
  const target = element<HTMLSelectElement>('style-dialog-target');
  showStyleDialogError(null);
  if (mode === 'new') {
    if (!paragraph) throw new Error('请先点击要套用新样式的段落。');
    target.replaceChildren(styleOption('', '（新样式）'));
    target.disabled = true;
    fillStyleRelations('paragraph', null);
    styleDialogInitial = { ...BLANK_STYLE_FORM, basedOn: currentStyle?.id ?? '' };
    writeStyleForm(styleDialogInitial, 'paragraph');
    element<HTMLButtonElement>('style-dialog-delete').disabled = true;
    element('style-dialog-title').textContent = '新建样式';
  } else {
    const editable = doc.getStyles().filter((style) => style.type === 'paragraph' || style.type === 'character');
    if (!editable.length) throw new Error('这份文档还没有可编辑的段落或字符样式。');
    target.replaceChildren(...editable.map((style) => styleOption(style.id, `${style.name}（${styleTypeLabel(style)}）`)));
    target.disabled = false;
    target.value = currentStyle && editable.some((style) => style.id === currentStyle.id) ? currentStyle.id : editable[0]!.id;
    loadStyleIntoDialog(target.value);
    element('style-dialog-title').textContent = '修改样式';
  }
  element<HTMLDialogElement>('style-dialog').showModal();
  element<HTMLInputElement>('style-dialog-name').focus();
}

function saveStyleDialog(): void {
  const current = readStyleForm();
  if (styleDialogMode === 'new') {
    const name = current.name.trim();
    if (!name) throw new Error('请填写样式名称。');
    if (doc.getStyles().some((style) => style.name === name)) throw new Error(`已有名为「${name}」的样式。`);
    const patch = stylePatchFromForm(BLANK_STYLE_FORM, current, 'paragraph');
    const id = uniqueStyleId(name, doc.getStyles().map((style) => style.id));
    const paragraph = styleDialogParagraph;
    if (paragraph === null) throw new Error('请先点击要套用新样式的段落。');
    // 一个批次：定义样式 + 套用到当前段落，撤销时一步回去。
    doc.applyOperations({ operations: [
      { type: 'defineStyle', style: {
        id, name, type: 'paragraph', quickFormat: current.quickFormat,
        ...(current.basedOn ? { basedOn: current.basedOn } : {}),
        ...(current.next ? { next: current.next } : {}),
        ...(withoutNulls(patch.paragraph) ? { paragraph: withoutNulls(patch.paragraph) } : {}),
        ...(withoutNulls(patch.run) ? { run: withoutNulls(patch.run) } : {}),
      } },
      { type: 'applyParagraphStyle', index: paragraph, styleId: id },
    ] });
    message(`已新建样式「${name}」并套用到当前段落。`);
  } else {
    const id = element<HTMLSelectElement>('style-dialog-target').value;
    const style = doc.getStyle(id);
    if (!style || !styleDialogInitial) throw new Error('样式已不存在，请重新打开对话框。');
    const patch = stylePatchFromForm(styleDialogInitial, current, style.type);
    if (Object.keys(patch).length) doc.updateStyle(id, patch);
    message(Object.keys(patch).length ? `已修改样式「${current.name.trim() || style.name}」，所有使用它的段落都已更新。` : '样式没有改动。');
  }
  element<HTMLDialogElement>('style-dialog').close();
  editor.render();
  refresh();
}

function deleteStyleFromDialog(): void {
  const id = element<HTMLSelectElement>('style-dialog-target').value;
  const style = doc.getStyle(id);
  if (!style) return;
  if (!window.confirm(`删除样式「${style.name}」？使用它的内容会改用默认样式，继承它的样式改为继承「${style.basedOn ? doc.getStyle(style.basedOn)?.name ?? style.basedOn : '（无）'}」。`)) return;
  doc.deleteStyle(id);
  element<HTMLDialogElement>('style-dialog').close();
  editor.render();
  refresh();
  message(`已删除样式「${style.name}」。`);
}

/** 公式对话框：编辑既有公式时记下它在正文里的位置；插入时只记段落。 */
let mathTarget: { paragraph: number; index: number | null } | null = null;

function renderMathPreview(source: string): void {
  const preview = element('math-dialog-preview');
  const error = element('math-dialog-error');
  preview.replaceChildren();
  error.hidden = true;
  if (!source.trim()) return;
  try {
    const ns = 'http://www.w3.org/1998/Math/MathML';
    const build = (node: MathMlNode): Element => {
      const created = document.createElementNS(ns, node.tag);
      for (const [name, value] of Object.entries(node.attrs ?? {})) created.setAttribute(name, value);
      if (node.text !== undefined) created.textContent = node.text;
      for (const child of node.children ?? []) created.append(build(child));
      return created;
    };
    const root = build(linearToMathMl(source));
    if (element<HTMLInputElement>('math-dialog-display').checked) root.setAttribute('display', 'block');
    preview.append(root);
  } catch (problem) {
    error.hidden = false;
    error.textContent = problem instanceof Error ? problem.message : String(problem);
  }
}

function openMathDialog(target?: { paragraph: number; index: number }): void {
  ensureMarkupView();
  editor.flush();
  const input = element<HTMLTextAreaElement>('math-dialog-input');
  const display = element<HTMLInputElement>('math-dialog-display');
  if (target) {
    const info = doc.getMath().filter((item) => item.paragraph === target.paragraph)[target.index];
    if (!info) throw new Error('找不到这个公式，文档可能已经改过了。');
    mathTarget = target;
    input.value = info.linear;
    display.checked = info.display === 'block';
    // 已有公式的行内 / 独立形式由它在文档里的元素决定（oMath / oMathPara），这里只读。
    display.disabled = true;
    element('math-dialog-title').textContent = '编辑公式';
    element<HTMLButtonElement>('math-dialog-delete').hidden = false;
  } else {
    mathTarget = { paragraph: selectedIndex(), index: null };
    input.value = '';
    display.checked = false;
    display.disabled = false;
    element('math-dialog-title').textContent = '插入公式（段落末尾）';
    element<HTMLButtonElement>('math-dialog-delete').hidden = true;
  }
  renderMathPreview(input.value);
  element<HTMLDialogElement>('math-dialog').showModal();
  input.focus();
}

function saveMathDialog(): void {
  const source = element<HTMLTextAreaElement>('math-dialog-input').value.trim();
  if (!mathTarget) return;
  if (!source) throw new Error('公式不能为空；要删掉它请用「删除公式」。');
  if (mathTarget.index === null) {
    doc.insertMath(mathTarget.paragraph, { linear: source },
      { display: element<HTMLInputElement>('math-dialog-display').checked ? 'block' : 'inline' });
    message('已插入公式。');
  } else {
    doc.setMath(mathTarget.paragraph, mathTarget.index, { linear: source });
    message('已更新公式。');
  }
  element<HTMLDialogElement>('math-dialog').close();
  editor.render();
  refresh();
}

function selectedIndex(): number {
  editor.flush();
  const index = editor.selectedParagraph;
  if (index === null || !doc.getParagraphs().some((paragraph) => paragraph.index === index)) {
    throw new Error('请先点击正文中的一个段落，再使用格式工具。');
  }
  return index;
}

function selectedCommentRange() {
  if (selectedRange) return toDocumentRange(selectedRange);
  const index = selectedIndex();
  const text = doc.getParagraphs().find((paragraph) => paragraph.index === index)?.text ?? '';
  return { paragraph: index, start: 0, end: text.length };
}

function selectedCell(): TableCellLocation {
  editor.flush();
  const cell = editor.selectedTableCell;
  if (!cell) throw new Error('请先把光标放进一个表格单元格，再使用表格工具。');
  if (cell.nested) {
    throw new Error('当前演示的结构化表格工具仅支持正文顶层表格，不支持嵌套表格。');
  }
  return cell;
}

function currentTableContext(): TableCellLocation | null {
  return editor.selectedTableCell;
}

function currentSelectionElement(): HTMLElement | null {
  const selection = document.getSelection();
  const anchor = selection?.anchorNode;
  if (anchor instanceof HTMLElement) return anchor;
  if (anchor?.parentElement) return anchor.parentElement;
  return document.activeElement instanceof HTMLElement ? document.activeElement : null;
}

function currentCommentIdsAtSelection(): number[] {
  const node = currentSelectionElement()?.closest<HTMLElement>('[data-docx-comment-ids]');
  if (!node) return [];
  return (node.dataset.docxCommentIds ?? '')
    .split(',')
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value));
}

function idsAtTarget(target: HTMLElement, attribute: 'docxCommentIds' | 'docxRevisionIds'): number[] {
  const node = target.closest<HTMLElement>(`[data-${attribute.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`)}]`);
  return (node?.dataset[attribute] ?? '')
    .split(',')
    .map(Number)
    .filter((value) => Number.isSafeInteger(value) && value >= 0);
}

function hyperlinkAtTarget(target: HTMLElement): HyperlinkInfo | null {
  const run = target.closest<HTMLElement>('[data-docx-run]');
  const paragraph = run?.closest<HTMLElement>('[data-paragraph]');
  const paragraphIndex = Number(paragraph?.dataset.paragraph);
  const runIndex = Number(run?.dataset.docxRun);
  if (!Number.isSafeInteger(paragraphIndex) || !Number.isSafeInteger(runIndex)) return null;
  return doc.getHyperlinks().find((link) => link.paragraph === paragraphIndex && link.runs.includes(runIndex)) ?? null;
}

function toDocumentRange(range: DocumentRange): DocumentRange {
  const paragraphByIndex = new Map(doc.getParagraphs().map((paragraph) => [paragraph.index, paragraph.text]));
  const toOffset = (paragraph: number, points: number): number => {
    const text = paragraphByIndex.get(paragraph) ?? '';
    let offset = 0;
    let count = 0;
    for (const char of text) {
      if (count >= points) break;
      offset += char.length;
      count++;
    }
    return offset;
  };
  return {
    start: { paragraph: range.start.paragraph, offset: toOffset(range.start.paragraph, range.start.offset) },
    end: { paragraph: range.end.paragraph, offset: toOffset(range.end.paragraph, range.end.offset) },
  };
}

function formatRuns(format: RunFormat): void {
  const range = selectedRange;
  if (range) doc.formatDocumentRange(toDocumentRange(range), format);
  else {
    const index = selectedIndex();
    const paragraph = doc.getParagraphs().find((item) => item.index === index)!;
    if (!paragraph.runs.length) throw new Error('请先在空段落中输入文字，再设置文字格式。');
    doc.formatDocumentRange({
      start: { paragraph: index, offset: 0 },
      end: { paragraph: index, offset: paragraph.text.length },
    }, format);
  }
  editor.render();
  refresh();
  message('已更新文字格式。');
}

function selectedTextRange() {
  const range = selectedRange;
  if (range && range.start.paragraph === range.end.paragraph) {
    const converted = toDocumentRange(range);
    return { paragraph: converted.start.paragraph, start: converted.start.offset, end: converted.end.offset };
  }
  const index = selectedIndex();
  const paragraph = doc.getParagraphs().find((item) => item.index === index)!;
  return { paragraph: index, start: 0, end: paragraph.text.length };
}

function activateFormatPainter(locked: boolean): void {
  const range = selectedTextRange();
  const format = doc.copyFormat(range);
  if (!Object.keys(format).length) throw new Error('当前选区没有可复制的直接文字格式。');
  formatPainter = { format, locked };
  updateSelection();
  message(locked ? '格式刷已锁定，连续点击可多次应用；按 Esc 取消。' : '格式刷已启用，下一次选区应用后自动关闭。');
}

function cancelFormatPainter(silent = false): void {
  if (!formatPainter) return;
  formatPainter = null;
  updateSelection();
  if (!silent) message('已取消格式刷。');
}

function loadXmlParts(): void {
  const previous = xmlPart.value;
  xmlPart.replaceChildren(...doc.listParts().filter((path) => /\.(xml|rels)$/i.test(path)).map((path) => {
    const option = document.createElement('option');
    option.value = option.textContent = path;
    return option;
  }));
  xmlPart.value = Array.from(xmlPart.options).some((option) => option.value === previous) ? previous : doc.mainDocumentPath;
  readXml();
}

function readXml(): void {
  editor.flush();
  xmlInput.value = doc.getPartXml(xmlPart.value);
  xmlRevision = doc.revision;
}

function resetAgent(): void {
  editor.flush();
  agentInput.value = JSON.stringify({
    expectedRevision: doc.revision,
    operations: [
      { type: 'insertParagraph', text: '来自 Agent 的建议：让下一步清晰、具体、可执行。' },
      { type: 'formatParagraph', index: editor.selectedParagraph ?? 0, format: { alignment: 'left' } },
    ],
  }, null, 2);
}

function setDocument(next: DocxDocument, name: string): void {
  editor.flush();
  editor.setDocument(next);
  doc = next;
  selectedRange = null;
  selectedRangeFormat = null;
  selectedCommentId = null;
  selectedRevisionId = null;
  recentNumbering.clear();
  filename = name;
  element('document-name').textContent = filename;
  refresh();
  loadXmlParts();
  resetAgent();
}

const commandRegistry = createCommandRegistry(createExampleCommandDescriptors({
  actions: {
    toggleViewMode() {
      editor.setViewMode(editor.getViewMode() === 'continuous' ? 'paginated' : 'continuous');
      element<HTMLButtonElement>('toggle-view-mode').textContent = editor.getViewMode() === 'continuous' ? '分页预览' : '连续视图';
      message(editor.getViewMode() === 'continuous' ? '已切换到连续视图。' : '已切换到只读分页预览。');
    },
    clipboard(action) {
      if (!document.execCommand(action)) message(`浏览器未允许${action === 'copy' ? '复制' : action === 'cut' ? '剪切' : '粘贴'}，请使用键盘快捷键。`, true);
    },
    toggleRunFormat(kind) {
      formatRuns({ [kind]: !selectedRangeFormat?.[kind] });
    },
    activateFormatPainter,
    clearFormat() {
      ensureMarkupView();
      const paragraph = doc.getParagraphs().find((item) => item.index === selectedIndex());
      if (!paragraph && !selectedRange) throw new Error('请先选择要清除格式的段落或文本。');
      if (selectedRange) doc.formatDocumentRange(selectedRange, CLEAR_RUN_FORMAT);
      if (paragraph) {
        const styleId = paragraph.style ?? defaultParagraphStyle()?.id;
        if (styleId) doc.applyParagraphStyle(paragraph.index, styleId, { clearDirectFormat: true });
        else doc.formatParagraph(paragraph.index, CLEAR_PARAGRAPH_FORMAT);
        if (!selectedRange && paragraph.text.length) {
          doc.formatDocumentRange({
            start: { paragraph: paragraph.index, offset: 0 },
            end: { paragraph: paragraph.index, offset: paragraph.text.length },
          }, CLEAR_RUN_FORMAT);
        }
      }
      editor.render();
      refresh();
      message('已清除直接格式。');
    },
    setFontSize(value) {
      formatRuns({ fontSize: value });
    },
    setFontColor(color) {
      formatRuns({ color });
    },
    openStyleDialog,
    openMathDialog: () => openMathDialog(),
    applyParagraphStyle(style) {
      ensureMarkupView();
      doc.applyParagraphStyle(selectedIndex(), style);
      editor.render();
      refresh();
      message(`已应用段落样式 ${doc.getStyle(style)?.name ?? style}。`);
    },
    setAlignment(alignment) {
      doc.formatParagraph(selectedIndex(), { alignment });
      editor.render();
      refresh();
      message('已更新段落对齐方式。');
    },
    applyNumbering,
    changeNumberingLevel,
    insertTableRow() {
      const cell = selectedCell();
      doc.insertTableRow(cell.table, cell.row + 1);
      editor.render();
      refresh();
      message('已在当前单元格下方插入一行。');
    },
    insertTableRowAt(where) {
      const cell = selectedCell();
      doc.insertTableRow(cell.table, where === 'above' ? cell.row : cell.row + cell.rowSpan);
      editor.render();
      refresh();
      message(`已在当前单元格${where === 'above' ? '上方' : '下方'}插入一行。`);
    },
    deleteTableRow() {
      const cell = selectedCell();
      doc.deleteTableRow(cell.table, cell.row);
      editor.render();
      refresh();
      message('已删除当前行。');
    },
    insertTableColumn() {
      const cell = selectedCell();
      doc.insertTableColumn(cell.table, cell.col + cell.colSpan);
      editor.render();
      refresh();
      message('已在当前单元格右侧插入一列。');
    },
    insertTableColumnAt(where) {
      const cell = selectedCell();
      doc.insertTableColumn(cell.table, where === 'left' ? cell.col : cell.col + cell.colSpan);
      editor.render();
      refresh();
      message(`已在当前单元格${where === 'left' ? '左侧' : '右侧'}插入一列。`);
    },
    deleteTableColumn() {
      const cell = selectedCell();
      doc.deleteTableColumn(cell.table, cell.col);
      editor.render();
      refresh();
      message('已删除当前列。');
    },
    mergeCells() {
      const cell = selectedCell();
      doc.mergeCells(cell.table, { row: cell.row, col: cell.col, rowSpan: 2, colSpan: 2 });
      editor.render();
      refresh();
      message('已尝试从当前单元格开始合并 2 × 2 区域。');
    },
    splitCell() {
      const cell = selectedCell();
      doc.splitCell(cell.table, cell.row, cell.col, cell.rowSpan, cell.colSpan);
      editor.render();
      refresh();
      message('已按当前跨度拆分单元格。');
    },
    applyCellStyle(fill) {
      const cell = selectedCell();
      doc.formatCell(cell.table, cell.row, cell.col, {
        shading: { fill: fill.slice(1).toUpperCase() },
        borders: {
          top: { style: 'single', size: 8, color: '7A8AA0' },
          right: { style: 'single', size: 8, color: '7A8AA0' },
          bottom: { style: 'single', size: 8, color: '7A8AA0' },
          left: { style: 'single', size: 8, color: '7A8AA0' },
        },
      });
      editor.render();
      refresh();
      message('已应用当前单元格边框与底纹。');
    },
    unavailableTableAction() {
      message('当前版本尚未提供此表格操作。', true);
    },
    startInsertImage() {
      imageAction = 'insert';
      imageActionTarget = null;
      element<HTMLInputElement>('image-file-input').click();
    },
    startReplaceImage(image) {
      imageActionTarget = image ?? editor.selectedImage;
      if (!imageActionTarget) return;
      imageAction = 'replace';
      element<HTMLInputElement>('image-file-input').click();
    },
    deleteImage(image) {
      const target = image ?? editor.selectedImage;
      if (!target) throw new Error('请先选择一张图片。');
      doc.deleteImage(target);
      editor.render();
      refresh();
      message('已删除图片。');
    },
    setImageAlt(text, image) {
      const target = image ?? editor.selectedImage;
      if (!target) throw new Error('请先选择一张图片。');
      doc.setImageAlt(target, text);
      editor.render();
      refresh();
      message('已更新图片替代文本。');
    },
    focusPreviousRevision() {
      const id = editor.focusPreviousRevision();
      selectedRevisionId = id;
      refreshRevisions();
    },
    focusNextRevision() {
      const id = editor.focusNextRevision();
      selectedRevisionId = id;
      refreshRevisions();
    },
    acceptAllRevisions() {
      editor.acceptAllRevisions(selectedReviewerAuthors?.length ? { authors: selectedReviewerAuthors } : {});
      selectedRevisionId = null;
      refresh();
    },
    rejectAllRevisions() {
      editor.rejectAllRevisions(selectedReviewerAuthors?.length ? { authors: selectedReviewerAuthors } : {});
      selectedRevisionId = null;
      refresh();
    },
    insertHyperlink(ctx) {
      const range = ctx.selection.range;
      if (!range || range.start.paragraph !== range.end.paragraph || ctx.selection.collapsed) {
        throw new Error('请先选择同一段落内的文字。');
      }
      const url = window.prompt('输入超链接地址');
      if (url === null) return;
      const converted = toDocumentRange(range);
      doc.insertHyperlink({
        paragraph: converted.start.paragraph,
        start: converted.start.offset,
        end: converted.end.offset,
      }, { url });
      editor.render();
      refresh();
    },
    openHyperlink(link) {
      const target = link.url ?? (link.anchor ? `#${link.anchor}` : '');
      if (target) window.open(target, '_blank', 'noopener,noreferrer');
    },
    copyHyperlinkAddress(link) {
      const target = link.url ?? (link.anchor ? `#${link.anchor}` : '');
      if (target) void navigator.clipboard?.writeText(target).catch(reportError);
    },
    editHyperlink(link) {
      const url = window.prompt('编辑超链接地址', link.url ?? (link.anchor ? `#${link.anchor}` : ''));
      if (url === null) return;
      doc.updateHyperlink(link, url.startsWith('#') ? { anchor: url.slice(1) } : { url });
      editor.render();
      refresh();
    },
    removeHyperlink(link) {
      doc.removeHyperlink(link, { keepText: true });
      editor.render();
      refresh();
    },
    acceptRevisions(ctx, paragraph) {
      const revisions = paragraph
        ? doc.getRevisions().filter((revision) => revision.paragraph === ctx.selection.paragraph)
        : ctx.revisionsAtPoint;
      if (revisions.length) doc.applyOperations({ operations: revisions.map((revision) => ({ type: 'acceptRevision' as const, id: revision.id })) });
      editor.render();
      refresh();
    },
    rejectRevisions(ctx, paragraph) {
      const revisions = paragraph
        ? doc.getRevisions().filter((revision) => revision.paragraph === ctx.selection.paragraph)
        : ctx.revisionsAtPoint;
      if (revisions.length) doc.applyOperations({ operations: revisions.map((revision) => ({ type: 'rejectRevision' as const, id: revision.id })) });
      editor.render();
      refresh();
    },
    openReviewOptions() {
      element('ribbon-tab-review').click();
    },
    addComment() {
      const text = window.prompt('输入批注内容');
      if (text === null) return;
      const id = doc.addComment(selectedCommentRange(), { text });
      selectedCommentId = id;
      editor.render();
      refresh();
      message('已添加批注。');
    },
    replyComment(commentId) {
      const text = window.prompt('输入回复内容');
      if (text === null) return;
      selectedCommentId = doc.replyComment(commentId, { text });
      refresh();
      message('已添加回复。');
    },
    toggleCommentResolved(commentId) {
      const current = doc.getComments().find((comment) => comment.id === commentId);
      if (!current) throw new Error('找不到当前批注。');
      doc.setCommentResolved(current.id, !current.resolved);
      refresh();
      message(current.resolved ? '已取消解决状态。' : '已标记为解决。');
    },
    deleteComment(commentId) {
      doc.deleteComment(commentId);
      selectedCommentId = null;
      editor.render();
      refresh();
      message('已删除批注。');
    },
    focusComment(commentId) {
      selectedCommentId = commentId;
      refreshComments();
      element('comment-list').querySelector<HTMLElement>('button.active')?.focus();
    },
  },
  getParagraphInfo(paragraph) {
    if (paragraph === null) return null;
    return doc.getParagraphs().find((item) => item.index === paragraph) ?? null;
  },
  getVisibleRevisionCount() {
    return visibleRevisions().length;
  },
  isFormatPainterActive() {
    return Boolean(formatPainter);
  },
  getFontSizeValue() {
    return element<HTMLSelectElement>('font-size').value;
  },
  getFontColorValue() {
    return element<HTMLInputElement>('font-color').value.slice(1);
  },
  getParagraphStyleValue() {
    return element<HTMLSelectElement>('paragraph-style').value;
  },
  getAlignmentValue() {
    return element<HTMLSelectElement>('alignment').value as ParagraphFormat['alignment'];
  },
  getCellFillValue() {
    return element<HTMLInputElement>('cell-fill').value;
  },
  getImageAltValue() {
    return element<HTMLInputElement>('image-alt').value;
  },
  isCommentResolved(commentId) {
    return doc.getComments().find((comment) => comment.id === commentId)?.resolved === true;
  },
}));

const commandControls: Array<{ elementId: string; commandId: string; pressed?: boolean }> = [
  { elementId: 'toggle-view-mode', commandId: 'view.mode' },
  { elementId: 'format-bold', commandId: 'format.bold', pressed: true },
  { elementId: 'format-italic', commandId: 'format.italic', pressed: true },
  { elementId: 'format-underline', commandId: 'format.underline', pressed: true },
  { elementId: 'format-painter', commandId: 'format.painter', pressed: true },
  { elementId: 'clear-format', commandId: 'format.clear' },
  { elementId: 'font-size', commandId: 'format.fontSize' },
  { elementId: 'font-color', commandId: 'format.fontColor' },
  { elementId: 'paragraph-style', commandId: 'paragraph.style' },
  { elementId: 'modify-style', commandId: 'style.modify' },
  { elementId: 'new-style', commandId: 'style.new' },
  { elementId: 'insert-math', commandId: 'math.insert' },
  { elementId: 'alignment', commandId: 'paragraph.alignment' },
  { elementId: 'list-bullet', commandId: 'list.bullet', pressed: true },
  { elementId: 'list-decimal', commandId: 'list.decimal', pressed: true },
  { elementId: 'list-indent', commandId: 'list.indent' },
  { elementId: 'list-outdent', commandId: 'list.outdent' },
  { elementId: 'insert-row', commandId: 'table.insertRow' },
  { elementId: 'delete-row', commandId: 'table.deleteRow' },
  { elementId: 'insert-col', commandId: 'table.insertColumn' },
  { elementId: 'delete-col', commandId: 'table.deleteColumn' },
  { elementId: 'merge-cells', commandId: 'table.mergeCells' },
  { elementId: 'split-cell', commandId: 'table.splitCell' },
  { elementId: 'apply-cell-style', commandId: 'table.applyCellStyle' },
  { elementId: 'replace-image', commandId: 'image.replace' },
  { elementId: 'delete-image', commandId: 'image.delete' },
  { elementId: 'image-alt', commandId: 'image.setAlt' },
  { elementId: 'reply-comment', commandId: 'comment.reply' },
  { elementId: 'resolve-comment', commandId: 'comment.toggleResolved' },
  { elementId: 'delete-comment', commandId: 'comment.delete' },
  { elementId: 'review-prev-revision', commandId: 'review.previousRevision' },
  { elementId: 'review-next-revision', commandId: 'review.nextRevision' },
  { elementId: 'review-accept-all', commandId: 'review.acceptAll' },
  { elementId: 'review-reject-all', commandId: 'review.rejectAll' },
] as const;

const commandContextBuilder = createCommandContextBuilder({
  getRevisionView: () => reviewRevisionView,
  getSelection(target) {
    const range = editor.selectedRange ?? selectedRange;
    const targetParagraph = Number(target?.closest<HTMLElement>('[data-paragraph]')?.dataset.paragraph);
    const paragraph = Number.isSafeInteger(targetParagraph)
      ? targetParagraph
      : doc.getParagraphs().find((item) => item.index === editor.selectedParagraph)?.index ?? null;
    return {
      paragraph,
      range,
      format: selectedRangeFormat,
      collapsed: !range || (range.start.paragraph === range.end.paragraph && range.start.offset === range.end.offset),
    };
  },
  getTable(_target, paragraph) {
    return paragraph === null ? currentTableContext() : doc.getTableCellAt(paragraph);
  },
  getImage(target) {
    const imageId = target?.closest<HTMLElement>('[data-image]')?.dataset.image;
    return imageId
      ? doc.getParagraphs().flatMap((item) => item.runs.flatMap((run) => run.images ?? [])).find((item) => item.id === imageId) ?? null
      : editor.selectedImage ?? null;
  },
  getShape(target) {
    const shapeId = target?.closest<HTMLElement>('[data-docx-shape]')?.dataset.docxShape;
    return shapeId ? doc.getShapes().find((shape) => shape.id === shapeId) ?? null : null;
  },
  getHyperlink: (target) => target ? hyperlinkAtTarget(target) : null,
  getRevisionsAtPoint(target) {
    const revisionIds = target ? idsAtTarget(target, 'docxRevisionIds') : [];
    return doc.getRevisions().filter((revision) => revisionIds.includes(revision.id));
  },
  getCommentsAtPoint: (target) => target ? idsAtTarget(target, 'docxCommentIds') : currentCommentIdsAtSelection(),
  getActiveCommentId: () => selectedCommentId,
  getClipboard: () => 'unknown',
});

function buildCommandContext(
  source: CommandContext['source'] = 'ribbon',
  target: HTMLElement | null = currentSelectionElement(),
): CommandContext {
  return commandContextBuilder(source, target);
}

function moveCaretToPoint(x: number, y: number): void {
  const owner = host.ownerDocument;
  const legacy = owner as Document & { caretRangeFromPoint?: (left: number, top: number) => Range | null };
  const modern = owner as Document & { caretPositionFromPoint?: (left: number, top: number) => { offsetNode: Node; offset: number } | null };
  let range = legacy.caretRangeFromPoint?.(x, y) ?? null;
  if (!range) {
    const position = modern.caretPositionFromPoint?.(x, y);
    if (position) {
      range = owner.createRange();
      range.setStart(position.offsetNode, position.offset);
      range.collapse(true);
    }
  }
  if (!range || !host.contains(range.startContainer)) return;
  const selection = owner.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
  owner.dispatchEvent(new Event('selectionchange'));
}

const editorRoot = host.querySelector<HTMLElement>('.docx-editor');
if (!editorRoot) throw new Error('找不到文档编辑区。');
initializeContextMenu({
  editorRoot,
  registry: commandRegistry,
  buildContext: (source, target) => buildCommandContext(source, target),
  moveCaretToPoint,
  beforeRun: () => editor.flush(),
  onError: reportError,
});

function syncCommandState(): void {
  const ctx = buildCommandContext();
  for (const { elementId, commandId, pressed } of commandControls) {
    const control = element<HTMLElement>(elementId);
    const command = commandRegistry.get(commandId);
    const state = getCommandControlState(command, ctx, { pressed });
    if ('disabled' in control) (control as HTMLButtonElement | HTMLSelectElement | HTMLInputElement).disabled = state.disabled;
    if (pressed) control.setAttribute('aria-pressed', String(Boolean(state.pressed)));
  }
}

function runCommand(id: string, source: CommandContext['source'] = 'ribbon', options?: { immediate?: boolean }): void {
  const action = () => commandRegistry.run(id, buildCommandContext(source));
  if (options?.immediate) {
    void action().catch(reportError);
    return;
  }
  run(action);
}

host.addEventListener('docx-selectionchange', updateSelection);
host.addEventListener('docx-rangechange', (event) => {
  const detail = (event as CustomEvent<{ range: DocumentRange; format: RunFormat } | null>).detail;
  selectedRange = detail?.range ?? null;
  selectedRangeFormat = detail?.format ?? null;
  if (formatPainter && detail?.range && !applyingFormatPainter) {
    const converted = toDocumentRange(detail.range);
    applyingFormatPainter = true;
    try {
      doc.beginHistoryGroup('format painter');
      try {
        doc.applyFormat(converted, formatPainter.format);
      } finally {
        doc.endHistoryGroup();
      }
      editor.render();
      refresh();
      message(formatPainter.locked ? '已应用格式（锁定模式）。' : '已应用格式。');
      if (!formatPainter.locked) formatPainter = null;
    } finally {
      applyingFormatPainter = false;
    }
  }
  updateSelection();
});
host.addEventListener('docx-imageselectionchange', updateImageSelection);
host.addEventListener('docx-tablecellchange', (event) => {
  const detail = (event as CustomEvent<{ cell: TableCellLocation } | null>).detail;
  updateTableSelection(detail?.cell ?? null);
});
host.addEventListener('docx-commentclick', (event) => {
  const ids = (event as CustomEvent<{ ids: number[] }>).detail?.ids ?? [];
  selectedCommentId = ids[0] ?? null;
  refreshComments();
});
element<HTMLInputElement>('toggle-comments').addEventListener('change', (event) => {
  host.classList.toggle('comments-hidden', !(event.target as HTMLInputElement).checked);
});
element<HTMLSelectElement>('review-revision-view').addEventListener('change', (event) => {
  reviewRevisionView = (event.target as HTMLSelectElement).value as 'final' | 'original' | 'markup';
  applyReviewFilter();
  syncCommandState();
});
element<HTMLInputElement>('review-show-revisions').addEventListener('change', (event) => {
  reviewShowRevisions = (event.target as HTMLInputElement).checked;
  applyReviewFilter();
  refreshRevisions();
});
element<HTMLInputElement>('review-show-comments').addEventListener('change', (event) => {
  reviewShowComments = (event.target as HTMLInputElement).checked;
  applyReviewFilter();
  refreshComments();
});
element('review-select-all').addEventListener('click', () => {
  selectedReviewerAuthors = undefined;
  applyReviewFilter();
  refreshReviewers();
  refreshComments();
  refreshRevisions();
});
element('review-clear-authors').addEventListener('click', () => {
  selectedReviewerAuthors = [];
  applyReviewFilter();
  refreshReviewers();
  refreshComments();
  refreshRevisions();
});
element('comment-author-filter').addEventListener('input', () => refreshComments());
element('comment-resolved-filter').addEventListener('change', () => refreshComments());
element('new-comment').addEventListener('click', () => runCommand('comment.new'));
element('toggle-view-mode').addEventListener('click', () => runCommand('view.mode'));
element('review-prev-revision').addEventListener('click', () => runCommand('review.previousRevision'));
element('review-next-revision').addEventListener('click', () => runCommand('review.nextRevision'));
element('review-accept-all').addEventListener('click', () => runCommand('review.acceptAll'));
element('review-reject-all').addEventListener('click', () => runCommand('review.rejectAll'));
element('reply-comment').addEventListener('click', () => runCommand('comment.reply'));
element('resolve-comment').addEventListener('click', () => runCommand('comment.toggleResolved'));
element('delete-comment').addEventListener('click', () => runCommand('comment.delete'));
for (const key of ['bold', 'italic', 'underline'] as const) {
  element(`format-${key}`).addEventListener('click', () => runCommand(`format.${key}`));
}
element('format-painter').addEventListener('click', () => runCommand('format.painter'));
element('format-painter').addEventListener('dblclick', (event) => run(() => {
  event.preventDefault();
  activateFormatPainter(true);
}));
element('clear-format').addEventListener('click', () => runCommand('format.clear'));
document.addEventListener('keydown', (event) => {
  if (event.key !== 'Escape') return;
  cancelFormatPainter();
});
element<HTMLSelectElement>('font-size').addEventListener('change', (event) => {
  const value = (event.target as HTMLSelectElement).value;
  if (value) runCommand('format.fontSize');
});
element<HTMLInputElement>('font-color').addEventListener('change', (event) => {
  const color = (event.target as HTMLInputElement).value.slice(1);
  if (color) runCommand('format.fontColor');
});
element<HTMLSelectElement>('alignment').addEventListener('change', (event) => {
  if ((event.target as HTMLSelectElement).value) runCommand('paragraph.alignment');
});
element('list-bullet').addEventListener('click', () => runCommand('list.bullet'));
element('list-decimal').addEventListener('click', () => runCommand('list.decimal'));
element('list-indent').addEventListener('click', () => runCommand('list.indent'));
element('list-outdent').addEventListener('click', () => runCommand('list.outdent'));
element<HTMLSelectElement>('paragraph-style').addEventListener('change', (event) => {
  const style = (event.target as HTMLSelectElement).value;
  if (style) runCommand('paragraph.style');
});
element('insert-math').addEventListener('click', () => runCommand('math.insert'));
// 点正文里的公式打开编辑对话框。公式节点是 contentEditable=false 的装饰，点它不会落光标。
element('editor').addEventListener('click', (event) => {
  const math = (event.target as Element | null)?.closest?.('[data-docx-math-index]') as HTMLElement | null;
  const paragraph = math?.closest<HTMLElement>('[data-paragraph]');
  if (!math || !paragraph || !isMarkupView()) return;
  run(() => openMathDialog({ paragraph: Number(paragraph.dataset.paragraph), index: Number(math.dataset.docxMathIndex) }));
});
element<HTMLTextAreaElement>('math-dialog-input').addEventListener('input', (event) => renderMathPreview((event.target as HTMLTextAreaElement).value));
element<HTMLInputElement>('math-dialog-display').addEventListener('change', () => renderMathPreview(element<HTMLTextAreaElement>('math-dialog-input').value));
element<HTMLFormElement>('math-form').addEventListener('submit', (event) => {
  event.preventDefault();
  try {
    saveMathDialog();
  } catch (problem) {
    const error = element('math-dialog-error');
    error.hidden = false;
    error.textContent = problem instanceof Error ? problem.message : String(problem);
  }
});
element('math-dialog-cancel').addEventListener('click', () => element<HTMLDialogElement>('math-dialog').close());
element('math-dialog-delete').addEventListener('click', () => run(() => {
  if (!mathTarget || mathTarget.index === null) return;
  doc.deleteMath(mathTarget.paragraph, mathTarget.index);
  element<HTMLDialogElement>('math-dialog').close();
  editor.render();
  refresh();
  message('已删除公式。');
}));
element('modify-style').addEventListener('click', () => runCommand('style.modify'));
element('new-style').addEventListener('click', () => runCommand('style.new'));
element<HTMLSelectElement>('style-dialog-target').addEventListener('change', (event) => run(() => {
  showStyleDialogError(null);
  loadStyleIntoDialog((event.target as HTMLSelectElement).value);
}));
element<HTMLInputElement>('style-dialog-color-picker').addEventListener('input', (event) => {
  element<HTMLInputElement>('style-dialog-color').value = (event.target as HTMLInputElement).value.slice(1).toUpperCase();
});
element<HTMLFormElement>('style-form').addEventListener('submit', (event) => {
  event.preventDefault();
  try {
    saveStyleDialog();
  } catch (error) {
    showStyleDialogError(error instanceof Error ? error.message : String(error));
  }
});
element('style-dialog-cancel').addEventListener('click', () => element<HTMLDialogElement>('style-dialog').close());
element('style-dialog-delete').addEventListener('click', () => {
  try {
    deleteStyleFromDialog();
  } catch (error) {
    showStyleDialogError(error instanceof Error ? error.message : String(error));
  }
});
element('add-paragraph').addEventListener('click', () => run(() => {
  editor.flush();
  doc.insertParagraph('在这里写下新的想法。');
  editor.render();
  refresh();
  message('已在文档末尾添加段落。');
}));
element('add-table').addEventListener('click', () => run(() => {
  editor.flush();
  const rows = Number(element<HTMLSelectElement>('table-rows').value);
  const cols = Number(element<HTMLSelectElement>('table-cols').value);
  doc.insertTableAt(rows, cols, undefined, { layout: 'fixed' });
  doc.setCellText(doc.getBlocks().filter((block) => block.type === 'table').length - 1, 0, 0, '项目');
  if (cols > 1) doc.setCellText(doc.getBlocks().filter((block) => block.type === 'table').length - 1, 0, 1, '说明');
  if (rows > 1) doc.setCellText(doc.getBlocks().filter((block) => block.type === 'table').length - 1, 1, 0, '新项目');
  if (rows > 1 && cols > 1) doc.setCellText(doc.getBlocks().filter((block) => block.type === 'table').length - 1, 1, 1, '点击单元格中的文字即可编辑');
  editor.render();
  refresh();
  message(`已在文档末尾添加 ${rows} × ${cols} 表格。`);
}));
element('insert-row').addEventListener('click', () => runCommand('table.insertRow'));
element('delete-row').addEventListener('click', () => runCommand('table.deleteRow'));
element('insert-col').addEventListener('click', () => runCommand('table.insertColumn'));
element('delete-col').addEventListener('click', () => runCommand('table.deleteColumn'));
element('merge-cells').addEventListener('click', () => runCommand('table.mergeCells'));
element('split-cell').addEventListener('click', () => runCommand('table.splitCell'));
element('apply-cell-style').addEventListener('click', () => runCommand('table.applyCellStyle'));
element('insert-image').addEventListener('click', () => runCommand('image.insert', 'ribbon', { immediate: true }));
element('replace-image').addEventListener('click', () => runCommand('image.replace', 'ribbon', { immediate: true }));
element('delete-image').addEventListener('click', () => runCommand('image.delete'));
element<HTMLInputElement>('image-alt').addEventListener('change', () => runCommand('image.setAlt'));
element('apply-page-setup').addEventListener('click', () => run(() => {
  const size = element<HTMLSelectElement>('page-size').value;
  const orientation = element<HTMLSelectElement>('page-orientation').value as 'portrait' | 'landscape';
  const portrait = size === 'Letter' ? { pageWidth: 12240, pageHeight: 15840 } : { pageWidth: 11906, pageHeight: 16838 };
  const setup = orientation === 'landscape'
    ? { pageWidth: portrait.pageHeight, pageHeight: portrait.pageWidth }
    : portrait;
  doc.setPageSetup(0, { ...setup, orientation });
  editor.render();
  refresh();
  message('已更新页面设置。');
}));
element('insert-section-break').addEventListener('click', () => run(() => {
  const index = selectedIndex();
  doc.insertSectionBreak(index, 'nextPage');
  editor.render();
  refresh();
  message('已插入分节符。');
}));
element('insert-page-break').addEventListener('click', () => run(() => {
  const index = selectedIndex();
  doc.updatePartXml(doc.mainDocumentPath, xml => {
    const paragraph = xml.getElementsByTagNameNS(WORD_NS, 'p')[index];
    if (!paragraph) throw new Error('找不到目标段落。');
    const run = xml.createElementNS(WORD_NS, 'w:r');
    const br = xml.createElementNS(WORD_NS, 'w:br');
    br.setAttributeNS(WORD_NS, 'w:type', 'page');
    run.appendChild(br);
    paragraph.appendChild(run);
  });
  editor.render();
  refresh();
  message('已插入分页符。');
}));
element<HTMLSelectElement>('header-kind').addEventListener('change', (event) => {
  editor.setHeaderKind((event.target as HTMLSelectElement).value as 'default' | 'first' | 'even');
});
element<HTMLSelectElement>('footer-kind').addEventListener('change', (event) => {
  editor.setFooterKind((event.target as HTMLSelectElement).value as 'default' | 'first' | 'even');
});
element('add-footnote').addEventListener('click', () => run(() => {
  editor.flush();
  const index = editor.selectedParagraph ?? 0;
  const run = doc.getParagraphs().find(item => item.index === index)?.runs.length ?? 0;
  doc.insertFootnote(index, run, '示例脚注内容');
  editor.render();
  refresh();
  message('已插入脚注。');
}));
element('add-endnote').addEventListener('click', () => run(() => {
  editor.flush();
  const index = editor.selectedParagraph ?? 0;
  const run = doc.getParagraphs().find(item => item.index === index)?.runs.length ?? 0;
  doc.insertEndnote(index, run, '示例尾注内容');
  editor.render();
  refresh();
  message('已插入尾注。');
}));
element('new-document').addEventListener('click', () => run(() => {
  if (!window.confirm('新建会替换当前工作区。请先下载需要保留的文档，是否继续？')) return;
  setDocument(DocxDocument.create(), '未命名.docx');
  message('已创建空白文档。点击纸张中的空段落开始输入。');
}));
element('open-document').addEventListener('click', () => element<HTMLInputElement>('file-input').click());
element<HTMLInputElement>('file-input').addEventListener('change', (event) => run(async () => {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (!file) return;
  if (!window.confirm('打开文件将替换当前工作区。请确认已下载需要保留的修改。')) return;
  message(`正在读取 ${file.name}…`);
  const next = await DocxDocument.load(file);
  setDocument(next, file.name);
  const macros = next.getPackageKind().hasMacros
    ? '文件含 VBA 宏：不会运行，保存时原样保留。'
    : '';
  message(`已打开 ${file.name}。${macros}未支持的版式可能不会显示，原始部件会保留。`);
}));
element<HTMLInputElement>('image-file-input').addEventListener('change', (event) => run(async () => {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = '';
  if (!file) return;
  const bytes = new Uint8Array(await file.arrayBuffer());
  const contentType = file.type || contentTypeForExtension(file.name.split('.').pop() ?? '') || 'application/octet-stream';
  if (!contentType.startsWith('image/')) throw new Error('请选择图片文件。');
  editor.flush();
  if (imageAction === 'replace') {
    if (!imageActionTarget) throw new Error('请先选择一张图片，再替换。');
    doc.replaceImageBytes(imageActionTarget, bytes, contentType);
    imageActionTarget = null;
    message(`已替换图片：${file.name}`);
  } else {
    doc.insertImage({ bytes, contentType, paragraph: editor.selectedParagraph ?? undefined, alt: file.name.replace(/\.[^.]+$/, '') });
    message(`已插入图片：${file.name}`);
  }
  editor.render();
  refresh();
}));
element('download-document').addEventListener('click', () => run(async () => {
  editor.flush();
  const blob = await doc.toBlob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  // 扩展名跟着包的种类走：.docm 存成 .docx 会让 Word 拒绝打开（内容类型与扩展名不符）。
  const extension = `.${doc.getPackageKind().extension}`;
  link.download = filename.toLowerCase().endsWith(extension) ? filename : `${filename.replace(/\.(docx|docm|dotx|dotm)$/i, '')}${extension}`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
  message(`已生成 ${extension.slice(1).toUpperCase()} 并发起下载。`);
}));
element('reset-agent').addEventListener('click', () => run(() => { resetAgent(); message('已填入使用当前修订号的请求示例。'); }));
element('apply-agent').addEventListener('click', () => run(() => {
  editor.flush();
  let request: unknown;
  try { request = JSON.parse(agentInput.value); } catch { throw new Error('JSON 格式不正确，请检查引号、逗号与括号。'); }
  if (!request || typeof request !== 'object' || !Array.isArray((request as AgentRequest).operations)) {
    throw new Error('请求必须是包含 operations 数组的 JSON 对象。');
  }
  const snapshot = doc.applyOperations(request as AgentRequest);
  editor.render();
  refresh(snapshot);
  message(`指令已执行，当前修订 ${snapshot.revision}。可切换「文档快照」查看结果。再次运行前请更新 expectedRevision。`);
}));
xmlPart.addEventListener('change', () => run(readXml));
element('reload-xml').addEventListener('click', () => run(() => { loadXmlParts(); message('已重新读取当前部件，未应用的 XML 修改已丢弃。'); }));
element('apply-xml').addEventListener('click', () => run(() => {
  editor.flush();
  if (xmlRevision !== doc.revision) throw new Error('文档已发生变化。请先保存你的 XML 草稿，再点击「重新读取」，避免覆盖较新的修改。');
  doc.setPartXml(xmlPart.value, xmlInput.value);
  editor.render();
  refresh();
  xmlRevision = doc.revision;
  message(`已应用 ${xmlPart.value}，修订 ${doc.revision}。`);
}));

const tabs = ['agent', 'xml', 'snapshot'];
function selectTab(name: string): void {
  for (const tab of tabs) {
    const active = tab === name;
    const button = element(`tab-${tab}`);
    button.classList.toggle('active', active);
    button.setAttribute('aria-selected', String(active));
    button.tabIndex = active ? 0 : -1;
    element(`panel-${tab}`).hidden = !active;
  }
}
for (const [index, name] of tabs.entries()) {
  const tab = element(`tab-${name}`);
  tab.addEventListener('click', () => selectTab(name));
  tab.addEventListener('keydown', (event) => {
    if (!['ArrowRight', 'ArrowLeft', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? tabs.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
    const nextName = tabs[next]!;
    selectTab(nextName);
    element(`tab-${nextName}`).focus();
  });
}

element<HTMLSelectElement>('review-revision-view').value = reviewRevisionView;
element<HTMLInputElement>('review-show-revisions').checked = reviewShowRevisions;
element<HTMLInputElement>('review-show-comments').checked = reviewShowComments;
applyReviewFilter();
refresh();
loadXmlParts();
resetAgent();
