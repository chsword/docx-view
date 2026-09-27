import { DocxDocument, DocxEditor } from '../src/index.js';
import type { AgentRequest, DocumentSnapshot, ParagraphFormat, RunFormat } from '../src/index.js';
import { findReusableNumberingId } from '../src/numbering.js';
import './style.css';

function element<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`找不到界面元素：${id}`);
  return node as T;
}

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
const status = element('status');
const host = element('editor');
const agentInput = element<HTMLTextAreaElement>('agent-input');
const xmlInput = element<HTMLTextAreaElement>('xml-input');
const xmlPart = element<HTMLSelectElement>('xml-part');
const editor = new DocxEditor(host, doc, { onChange: refresh });

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
  return sample;
}

function message(text: string, error = false): void {
  status.textContent = text;
  status.classList.toggle('error', error);
  status.setAttribute('role', error ? 'alert' : 'status');
}

function run(action: () => void | Promise<void>): void {
  void Promise.resolve().then(action).catch((error: unknown) => {
    message(error instanceof Error ? error.message : String(error), true);
  });
}

function refresh(snapshot: DocumentSnapshot = doc.getSnapshot()): void {
  element('revision').textContent = String(snapshot.revision);
  element('paragraph-count').textContent = String(snapshot.paragraphs.length);
  element('snapshot-output').textContent = JSON.stringify(snapshot, null, 2);
  loadStyleOptions();
  updateSelection();
}

function loadStyleOptions(): void {
  const select = element<HTMLSelectElement>('paragraph-style');
  const previous = select.value;
  select.replaceChildren(...[
    (() => {
      const option = document.createElement('option');
      option.value = '';
      option.textContent = '样式';
      return option;
    })(),
    ...doc.getStyles()
      .filter((style) => style.type === 'paragraph' && style.quickFormat)
      .map((style) => {
        const option = document.createElement('option');
        option.value = style.id;
        option.textContent = `${style.name} (${style.id})`;
        return option;
      }),
  ]);
  if (Array.from(select.options).some((option) => option.value === previous)) select.value = previous;
}

function updateSelection(): void {
  const index = editor.selectedParagraph;
  const paragraph = doc.getParagraphs().find((item) => item.index === index);
  element('selection-label').textContent = paragraph ? `已选择第 ${paragraph.index + 1} 段 · 格式应用于整段` : '点击正文选择段落';
  for (const key of ['bold', 'italic', 'underline'] as const) {
    const button = element<HTMLButtonElement>(`format-${key}`);
    button.disabled = !paragraph;
    button.setAttribute('aria-pressed', String(Boolean(paragraph?.runs.length && paragraph.runs.every((item) => item[key]))));
  }
  const size = element<HTMLSelectElement>('font-size');
  const color = element<HTMLInputElement>('font-color');
  const style = element<HTMLSelectElement>('paragraph-style');
  const alignment = element<HTMLSelectElement>('alignment');
  const bullet = element<HTMLButtonElement>('list-bullet');
  const decimal = element<HTMLButtonElement>('list-decimal');
  const indent = element<HTMLButtonElement>('list-indent');
  const outdent = element<HTMLButtonElement>('list-outdent');
  size.disabled = color.disabled = style.disabled = alignment.disabled = !paragraph;
  bullet.disabled = decimal.disabled = !paragraph;
  indent.disabled = !paragraph?.numbering || paragraph.numbering.level >= 8;
  outdent.disabled = !paragraph?.numbering || paragraph.numbering.level <= 0;
  bullet.setAttribute('aria-pressed', String(Boolean(paragraph?.numbering?.isBullet)));
  decimal.setAttribute('aria-pressed', String(Boolean(paragraph?.numbering && !paragraph.numbering.isBullet)));
  const firstRun = paragraph?.runs[0];
  const effectiveRun = firstRun?.effective ?? firstRun;
  size.value = effectiveRun?.fontSize ? String(effectiveRun.fontSize) : '';
  const runColor = effectiveRun?.color;
  color.value = runColor && /^[0-9a-f]{6}$/i.test(runColor) ? `#${runColor}` : '#25334a';
  style.value = paragraph?.style ?? '';
  alignment.value = paragraph?.effective?.alignment ?? paragraph?.alignment ?? 'left';
  element('effective-format').textContent = paragraph ? JSON.stringify(paragraph.effective ?? {}, null, 2) : '点击正文选择段落';
}

function selectedIndex(): number {
  editor.flush();
  const index = editor.selectedParagraph;
  if (index === null || !doc.getParagraphs().some((paragraph) => paragraph.index === index)) {
    throw new Error('请先点击正文中的一个段落，再使用格式工具。');
  }
  return index;
}

function selectedCell(): { table: number; row: number; col: number; rowSpan: number; colSpan: number } {
  editor.flush();
  const active = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const cell = active?.closest<HTMLTableCellElement>('td[data-table-cell="true"]');
  if (!cell) throw new Error('请先把光标放进一个表格单元格，再使用表格工具。');
  const tableElement = cell.closest('.docx-table');
  if (!tableElement) throw new Error('未找到当前表格。');
  if (tableElement.parentElement?.closest('.docx-table')) {
    throw new Error('当前演示的结构化表格工具仅支持正文顶层表格，不支持嵌套表格。');
  }
  const table = Array.from(host.querySelectorAll('.docx-editor > .docx-table')).indexOf(tableElement);
  if (table < 0) throw new Error('未找到当前表格。');
  return {
    table,
    row: Number(cell.dataset.rowStart ?? 0),
    col: Number(cell.dataset.gridStart ?? 0),
    rowSpan: Math.max(1, Number(cell.getAttribute('rowspan') ?? 1)),
    colSpan: Math.max(1, Number(cell.getAttribute('colspan') ?? 1)),
  };
}

function formatRuns(format: RunFormat): void {
  const index = selectedIndex();
  const paragraph = doc.getParagraphs().find((item) => item.index === index)!;
  if (!paragraph.runs.length) throw new Error('请先在空段落中输入文字，再设置文字格式。');
  doc.applyOperations({
    expectedRevision: doc.revision,
    operations: paragraph.runs.map((run) => ({ type: 'formatRun', paragraph: index, run: run.index, format })),
  });
  editor.render();
  refresh();
  message(`已更新第 ${index + 1} 段格式。`);
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
  recentNumbering.clear();
  filename = name;
  element('document-name').textContent = filename;
  refresh();
  loadXmlParts();
  resetAgent();
}

host.addEventListener('docx-selectionchange', updateSelection);
for (const key of ['bold', 'italic', 'underline'] as const) {
  element(`format-${key}`).addEventListener('click', () => run(() => {
    const index = selectedIndex();
    const paragraph = doc.getParagraphs().find((item) => item.index === index)!;
    formatRuns({ [key]: !paragraph.runs.every((item) => item[key]) });
  }));
}
element<HTMLSelectElement>('font-size').addEventListener('change', (event) => {
  const value = (event.target as HTMLSelectElement).value;
  if (value) run(() => formatRuns({ fontSize: Number(value) }));
});
element<HTMLInputElement>('font-color').addEventListener('change', (event) => {
  const color = (event.target as HTMLInputElement).value.slice(1);
  run(() => formatRuns({ color }));
});
element<HTMLSelectElement>('alignment').addEventListener('change', (event) => {
  const alignment = (event.target as HTMLSelectElement).value as ParagraphFormat['alignment'];
  run(() => {
    doc.formatParagraph(selectedIndex(), { alignment });
    editor.render();
    refresh();
    message('已更新段落对齐方式。');
  });
});
element('list-bullet').addEventListener('click', () => run(() => applyNumbering('bullet')));
element('list-decimal').addEventListener('click', () => run(() => applyNumbering('decimal')));
element('list-indent').addEventListener('click', () => run(() => changeNumberingLevel(1)));
element('list-outdent').addEventListener('click', () => run(() => changeNumberingLevel(-1)));
element<HTMLSelectElement>('paragraph-style').addEventListener('change', (event) => {
  const style = (event.target as HTMLSelectElement).value;
  if (!style) return;
  run(() => {
    doc.formatParagraph(selectedIndex(), { style }, { validateStyle: true });
    editor.render();
    refresh();
    message(`已应用段落样式 ${style}。`);
  });
});
element('list-bullet').addEventListener('click', () => run(() => applyNumbering('bullet')));
element('list-decimal').addEventListener('click', () => run(() => applyNumbering('decimal')));
element('list-indent').addEventListener('click', () => run(() => changeNumberingLevel(1)));
element('list-outdent').addEventListener('click', () => run(() => changeNumberingLevel(-1)));
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
element('insert-row').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  doc.insertTableRow(cell.table, cell.row + 1);
  editor.render();
  refresh();
  message('已在当前单元格下方插入一行。');
}));
element('delete-row').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  doc.deleteTableRow(cell.table, cell.row);
  editor.render();
  refresh();
  message('已删除当前行。');
}));
element('insert-col').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  doc.insertTableColumn(cell.table, cell.col + cell.colSpan);
  editor.render();
  refresh();
  message('已在当前单元格右侧插入一列。');
}));
element('delete-col').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  doc.deleteTableColumn(cell.table, cell.col);
  editor.render();
  refresh();
  message('已删除当前列。');
}));
element('merge-cells').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  doc.mergeCells(cell.table, { row: cell.row, col: cell.col, rowSpan: 2, colSpan: 2 });
  editor.render();
  refresh();
  message('已尝试从当前单元格开始合并 2 × 2 区域。');
}));
element('split-cell').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  doc.splitCell(cell.table, cell.row, cell.col, cell.rowSpan, cell.colSpan);
  editor.render();
  refresh();
  message('已按当前跨度拆分单元格。');
}));
element('apply-cell-style').addEventListener('click', () => run(() => {
  const cell = selectedCell();
  const fill = element<HTMLInputElement>('cell-fill').value.slice(1).toUpperCase();
  doc.formatCell(cell.table, cell.row, cell.col, {
    shading: { fill },
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
  message(`已打开 ${file.name}。未支持的版式可能不会显示，原始部件会保留。`);
}));
element('download-document').addEventListener('click', () => run(async () => {
  editor.flush();
  const blob = await doc.toBlob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename.toLowerCase().endsWith('.docx') ? filename : `${filename}.docx`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
  message('已生成 DOCX 并发起下载。');
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

refresh();
loadXmlParts();
resetAgent();
