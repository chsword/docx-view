import { DocxDocument, DocxEditor } from '../src/index.js';
import type { AgentRequest, DocumentSnapshot, ParagraphFormat, RunFormat } from '../src/index.js';
import './style.css';

function element<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`找不到界面元素：${id}`);
  return node as T;
}

let doc = createSample();
let filename = '产品计划.docx';
let xmlRevision = -1;
let showFormattingMarks = false;
const status = element('status');
const host = element('editor');
const agentInput = element<HTMLTextAreaElement>('agent-input');
const xmlInput = element<HTMLTextAreaElement>('xml-input');
const xmlPart = element<HTMLSelectElement>('xml-part');
let editor = new DocxEditor(host, doc, { onChange: refresh, showFormattingMarks });

function createSample(): DocxDocument {
  const sample = DocxDocument.create();
  sample.setParagraphText(0, '把想法，写成下一步。');
  sample.formatRun(0, 0, { bold: true, fontSize: 26, color: '223855' });
  sample.insertParagraph('纸间工作室  /  产品共创计划  /  2026');
  sample.formatRun(1, 0, { fontSize: 10, color: '788597' });
  sample.insertParagraph('01  项目愿景');
  sample.formatRun(2, 0, { bold: true, fontSize: 16, color: '3567D6' });
  sample.insertParagraph('让每一份文档都能自由流转。我们希望把熟悉的文字编辑，与透明的文档结构、可靠的自动化连接起来。');
  sample.insertParagraph('点击任意段落开始编辑，也可以在右侧运行一组 Agent 指令。所有处理都发生在你的浏览器里，文件不会上传。');
  sample.insertParagraph('02  从一个小计划开始');
  sample.formatRun(5, 0, { bold: true, fontSize: 16, color: '3567D6' });
  sample.insertParagraph('目录示例\t第一章\t1.1');
  sample.setParagraphTabs(6, [
    { position: 2200, alignment: 'left', leader: 'dot' },
    { position: 6200, alignment: 'right' },
  ]);
  sample.insertTable([['阶段', '交付内容', '状态'], ['探索', '梳理需求与文档结构', '已完成'], ['共创', '编辑体验与自动化接口', '进行中'], ['发布', '验证 DOCX 导出与兼容性', '下一步']]);
  sample.insertParagraph('好的工具，让内容成为主角。');
  const last = sample.getParagraphs().at(-1)!;
  sample.formatRun(last.index, 0, { italic: true, fontSize: 11, color: '788597' });
  return sample;
}

function rebuildEditor(): void {
  editor.destroy();
  editor = new DocxEditor(host, doc, { onChange: refresh, showFormattingMarks });
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
  updateSelection();
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
  const alignment = element<HTMLSelectElement>('alignment');
  size.disabled = color.disabled = alignment.disabled = !paragraph;
  size.value = paragraph?.runs[0]?.fontSize ? String(paragraph.runs[0].fontSize) : '';
  const runColor = paragraph?.runs[0]?.color;
  color.value = runColor && /^[0-9a-f]{6}$/i.test(runColor) ? `#${runColor}` : '#25334a';
  alignment.value = paragraph?.alignment ?? 'left';
}

function selectedIndex(): number {
  editor.flush();
  const index = editor.selectedParagraph;
  if (index === null || !doc.getParagraphs().some((paragraph) => paragraph.index === index)) {
    throw new Error('请先点击正文中的一个段落，再使用格式工具。');
  }
  return index;
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
  doc = next;
  filename = name;
  rebuildEditor();
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
element('add-paragraph').addEventListener('click', () => run(() => {
  editor.flush();
  doc.insertParagraph('在这里写下新的想法。');
  editor.render();
  refresh();
  message('已在文档末尾添加段落。');
}));
element('add-table').addEventListener('click', () => run(() => {
  editor.flush();
  doc.insertTable([['项目', '说明'], ['新项目', '点击单元格中的文字即可编辑']]);
  editor.render();
  refresh();
  message('已在文档末尾添加 2 × 2 表格。');
}));
element('set-tabs').addEventListener('click', () => run(() => {
  const index = selectedIndex();
  doc.setParagraphTabs(index, [
    { position: 2200, alignment: 'left', leader: 'dot' },
    { position: 6200, alignment: 'right' },
  ]);
  editor.render();
  refresh();
  message('已设置段落制表位。');
}));
element('clear-tabs').addEventListener('click', () => run(() => {
  doc.formatParagraph(selectedIndex(), { tabs: [] });
  editor.render();
  refresh();
  message('已清除段落制表位。');
}));
element('set-border-shading').addEventListener('click', () => run(() => {
  const index = selectedIndex();
  doc.setParagraphBorders(index, {
    top: { style: 'single', size: 8, space: 2, color: '3567D6' },
    bottom: { style: 'single', size: 8, space: 2, color: '3567D6' },
  });
  doc.setParagraphShading(index, { pattern: 'clear', fill: 'EEF3FF' });
  editor.render();
  refresh();
  message('已设置段落边框和底纹。');
}));
element('insert-page-break').addEventListener('click', () => run(() => {
  const index = selectedIndex();
  doc.insertBreak(index, 0, 'page');
  editor.render();
  refresh();
  message('已插入分页符。');
}));
element('insert-symbol').addEventListener('click', () => run(() => {
  const index = selectedIndex();
  doc.insertSymbol(index, 0, 'Wingdings', 0xF0FC);
  editor.render();
  refresh();
  message('已插入符号字符。');
}));
element<HTMLInputElement>('show-marks').addEventListener('change', (event) => run(() => {
  showFormattingMarks = (event.target as HTMLInputElement).checked;
  rebuildEditor();
  refresh();
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
