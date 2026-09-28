# docx-view

一个面向浏览器和 AI Agent 的 TypeScript / JavaScript DOCX 编辑组件库，包含无需后端的静态 `examples`。

**当前是可运行的基础版本，不是 Microsoft Word 排版引擎，也不等同于 .NET Open XML SDK 的完整实现。** 支持段落、文字格式、编号 / 项目符号列表、表格、常见图片、分节页面设置与页眉页脚编辑；对于更细粒度的操作，可以直接访问 DOCX 包中的部件、关系 XML 和命名空间感知的 OOXML DOM。

## 运行

需要 Node.js 22.12+。

```sh
npm ci
npm run dev
```

打开终端显示的地址。演示包含：新建 / 打开 / 下载 DOCX、正文 / 表格 / 图片编辑、整段文字格式工具、样式下拉框、当前段落有效格式面板、OOXML 部件编辑器、Agent JSON 请求和文档快照。

```sh
npm run check          # TypeScript 检查
npm test              # 构建库并运行 Node 内置测试
npm run test:perf     # 独立进程性能回归检查（预热 + 多轮取中位数）
npm run build         # dist/：ES modules、类型声明和 source maps
npm run build:examples # examples-dist/：可部署的纯静态站点
npm pack              # 打包库，供其他项目安装
```

将 `examples-dist/` 部署到任意静态 HTTP 服务即可；资源使用相对路径，支持子目录部署。不要直接用 `file://` 打开源码 HTML。

### 性能回归检查

默认 `npm test` **不运行**性能断言，避免它们与其余测试共享进程时受到机器负载、GC 和新增测试数量的影响。需要检查性能回归时，单独运行：

```sh
npm run test:perf
```

性能套件会在独立 Node 进程里对关键场景先预热，再跑 5 轮取中位数，并断言：

- `insertParagraph` 的单位成本不能明显偏离 `setParagraphText`
- 批量 `applyOperations` 插入必须继续明显快于逐次插入
- 1000 段批量改写必须继续保持近线性扩展

失败信息会同时打印实测中位数、各轮样本和对应阈值，便于判断是真退化还是环境噪声。CI 也会把这组检查作为独立步骤运行。

### GitHub Pages 演示站

部署成功后的地址：**https://chsword.github.io/docx-view/**

首次部署：

1. 在仓库 **Settings → Pages → Build and deployment → Source** 中选择 **GitHub Actions**。
2. 将部署工作流合并到 `main`，推送后会自动执行类型检查、测试、静态构建及部署。
3. 在 **Actions → Deploy examples to GitHub Pages** 查看结果；也可以通过 **Run workflow** 选择 `main` 手动部署。

工作流只允许从 `main` 发布，构建产物来自 `examples-dist/`，不需要额外的访问令牌或 `gh-pages` 分支。如果为 `github-pages` 环境配置了审批规则，需要批准后才能发布。上述链接在首次部署成功后才可访问。

## 库 API

以下示例供安装本项目打包产物的应用使用；库的核心 API 同时支持 Node.js 和浏览器，不依赖浏览器全局 `document`。

### 创建、编辑与导出

```ts
import { DocxDocument } from 'docx-view';

const doc = DocxDocument.create();
doc.setParagraphText(0, '你好，DOCX！');
doc.formatRange({ paragraph: 0, start: 3, end: 7 }, { bold: true, fontSize: 18, color: '2455AA' });
doc.formatParagraph(0, { alignment: 'center' });
doc.insertParagraph('支持段落、文字与表格编辑。');
doc.insertTable([['任务', '状态'], ['文档编辑', '完成']]);
doc.insertImage({ bytes: pngBytes, contentType: 'image/png', alt: '产品示意图' });

const bytes = await doc.toUint8Array(); // Node: 可传给 fs.writeFile
const blob = await doc.toBlob();       // 浏览器：可用于下载
const reopened = await DocxDocument.load(bytes); // 也接受 ArrayBuffer / Blob / File
console.log(reopened.getSnapshot());
```

| API | 用途 |
| --- | --- |
| `getParagraphs()` / `getBlocks()` / `getSnapshot()` | 段落 / 表格结构、直接格式、有效格式、修订标记、样式清单、部件列表和修订号 |
| `getStyles()` / `getStyle(id)` | 读取 `styles.xml` 中的段落 / 字符 / 表格 / 编号样式元数据 |
| `getStyleGallery()` | 读取常用（`qFormat`）样式，按 `uiPriority` 排序，适合工具栏样式面板 |
| `getEffectiveParagraphFormat(index)` / `getEffectiveRunFormat(paragraph, run)` | 读取 Word 样式层叠后的有效格式 |
| `setParagraphText(index, text)` | 修改段落文字，支持制表符和换行 |
| `insertParagraph(text, before?)` | 在指定段落前插入；省略 `before` 则追加到正文 |
| `deleteParagraph(index)` | 删除段落；保留正文 / 单元格必要的空段落，拒绝隐式删除分节符 |
| `formatParagraph(index, format, options?)` | 设置段落直接格式；`options.validateStyle` 可在写入前校验样式 ID；将某个字段设为 `null` 可清除该直接格式 |
| `applyParagraphStyle(index, styleId, options?)` | 严格应用段落样式；样式不存在时报错，`clearDirectFormat` 可清除与样式冲突的段落 / run 直接格式 |
| `getNumberingDefinitions()` | 读取 `word/numbering.xml` 中已解析的编号定义 |
| `setParagraphNumbering(index, numId, level?)` | 为段落绑定指定编号定义与级别（默认 0） |
| `clearParagraphNumbering(index)` | 清除段落上的直接编号绑定 |
| `createNumbering(kind)` | 创建新的项目符号 / 编号 / 多级编号定义并返回新的 `numId` |
| `setParagraphLevel(index, delta)` | 提高 / 降低段落列表级别，结果钳制在 `0..8` |
| `formatRun(paragraph, run, format)` | 设置 run 直接格式，包括字符样式、字号、颜色、下划线、删除线、上下标等常用字段；将某个字段设为 `null` 可回退到继承样式 |
| `formatRange(range, format)` / `clearRangeFormat(range, fields?)` | 按段落内字符偏移格式化任意文本范围，支持清除全部或指定 run 直接格式字段 |
| `applyCharacterStyle(range, styleId, options?)` | 严格应用字符样式；可在应用时清除与字符样式冲突的直接格式 |
| `getRangeFormat(range)` | 读取字符范围内一致的 run 直接格式；同一字段在范围内不一致时返回 `undefined` |
| `copyFormat(range)` / `applyFormat(range, format)` | 格式刷 API：复制段落内范围的 run 直接格式，并应用到跨段落选区 |
| `formatDocumentRange(range, format)` / `getDocumentRangeFormat(range)` | 支持跨段落选区：首段部分 + 中间整段 + 末段部分 |
| `copyClipboardFragment(range)` / `pasteClipboardFragment(range, fragment)` | 内部富文本剪贴板 API：run 直接格式、超链接、内嵌图片、段落格式与样式引用、编号（含多级）和表格；粘贴按落点切分段落，图片按目标文档关系与 media 部件重建（不复用源 `rId`） |
| `defineStyle(style)` | 创建或更新 `styles.xml` 样式定义；缺少部件时自动补内容类型与主文档关系 |
| `createStyleFromSelection(range, style)` | 从当前选区提炼段落 / 字符直接格式并写成新的段落样式定义 |
| `getOutline()` / `setOutlineLevel(index, level)` / `moveOutlineSection(from, to)` | 读取标题大纲、显式设置 `outlineLevel`，以及按大纲整体移动标题节 |
| `replaceText(search, replacement)` | 正文及表格段落内的字面替换，支持跨 run 匹配，不跨段落 |
| `DocxDocument.compare(base, revised, options?)` | 比较两份主文档并返回新的带修订实例；纯文本/格式差异细化到段落与 run，表格及含图片的段落按粗粒度删除 + 插入处理 |
| `getRevisions(filter?)` / `acceptRevision(id)` / `rejectRevision(id)` / `acceptAllRevisions(filter?)` / `rejectAllRevisions(filter?)` | 读取并逐条/批量接受或拒绝修订（支持按作者筛选）；`moveFrom` / `moveTo` 当前按删除 / 插入降级处理 |
| `getReviewers()` | 聚合主文档修订与主文档锚点批注的审阅者统计（修订数、批注数、未解决批注数、时间范围）；结果默认按 `revisionCount`、`commentCount`、`unresolvedCommentCount` 降序，再按 `author` 升序 |
| `insertTable(rows)` / `insertTableAt(rows, cols, before?, format?)` | 在正文中插入表格；支持空白表格、基础表格格式和正文块级定位 |
| `getTable(index)` | 读取正文中第 N 个表格的 grid、跨度和表格/行/单元格格式信息 |
| `insertTableRow()` / `deleteTableRow()` / `insertTableColumn()` / `deleteTableColumn()` | 行列编辑；同步维护 `w:tblGrid`，拒绝删成 0 行或 0 列 |
| `mergeCells()` / `splitCell()` | 基于逻辑网格合并/拆分单元格，读写 `w:gridSpan` / `w:vMerge` |
| `formatTable()` / `formatTableRow()` / `formatCell()` | 设置表格宽度、布局、边框、底纹、行高、标题行、单元格对齐和边距等显式属性 |
| `setCellText()` | 修改可见单元格文字，同时保留段落结构与其他未改内容 |
| `getImages()` / `getImageBytes()` / `getImageDataUrl()` | 读取主文档中的图片元数据、二进制内容和可直接渲染的 `data:` URL |
| `insertImage()` / `replaceImageBytes()` / `resizeImage()` / `setImageAlt()` / `deleteImage()` | 插入、替换、调整尺寸、更新替代文本和删除图片 |
| `getSections()` / `getSection(index)` | 读取分节类型、纸张、页边距、分栏、页眉页脚引用 |
| `setPageSetup(section, setup)` | 修改指定节的纸张方向、边距、分栏和页码起始等页面设置 |
| `insertSectionBreak(paragraph, type)` / `deleteSectionBreak(section)` | 插入/删除分节符（删除为显式 API，保留原有段落删除保护） |
| `getHeaderBlocks()` / `getFooterBlocks()` | 读取页眉页脚 block 结构（段落、表格等） |
| `createHeader()` / `createFooter()` / `setHeaderText()` / `setFooterText()` | 创建并写入页眉页脚部件，自动维护 rels 与 content-types |
| `insertPageNumberField(partPath, options)` | 在页眉/页脚部件写入 `PAGE` 或 `NUMPAGES` 域占位结构 |
| `getSettings()` / `setTrackChanges(enabled)` / `setRevisionAuthor(author)` | 读取常用文档设置（当前返回 `{ defaultTabStop, evenAndOddHeaders, trackChanges }`），显式开启/关闭 `w:trackChanges`（关闭时写 `w:val="0"`，不删除元素），并设置后续记录修订写入使用的作者名 |
| `revision` | 本实例的修订号；加载文件后从 0 开始，不持久化到 DOCX |

索引从 0 开始，包含主文档中的表格段落；结构变更后请重新读取快照。`insertSectionBreak(paragraph)` 的 `paragraph` 表示“该段落结束处插入分节”；`deleteSectionBreak(section)` 删除第 `section` 节末尾的分节符并与下一节合并。高层操作默认处理主文档，可通过页眉页脚 API 读写 `header*.xml` / `footer*.xml`；`getRevisions()`、`accept*`/`reject*`、`getReviewers()` 与 `DocxDocument.compare()` 当前都只作用于主文档。`compare()` 返回的是新实例（`revision` 从 0 开始），不会修改输入文档；它对纯段落文字/格式做细粒度修订，对表格和含图片的段落仅做粗粒度删除 + 插入，不比较页眉/页脚/脚注/尾注。超过 1000 个主文档段落的输入会直接抛错，避免段落级对齐的 O(N²) 内存开销。当段落中存在未接受的删除 (`w:del` / `w:moveFrom`) 时，其文本不会进入 `paragraph.text`，但对应 run 仍保留在 `runs[]` 中并以空字符串占位；删除内容请通过 `getRevisions().deletedText` 读取。记录修订开启后，文本/段落/图片/表格行等编辑会写入 `w:ins` / `w:del` / `rPrChange` / `pPrChange`；此时 `deleteParagraph()` 会保留原段落节点并把内容标记为删除，因此按索引循环删除时应在每步后重新读取段落列表。接受“仅段落标记删除（`w:pPr/w:rPr/w:del`）”时会优先合并到同容器中的下一段；若当前段/下一段含 `sectPr`、当前段后继不是段落、或已到容器末尾，则降级为仅移除该删除标记。若未显式调用 `setRevisionAuthor(author)`，默认作者名为 `docx-view`。
节范围是闭区间：`startParagraph <= i <= endParagraph`。当某节暂时没有段落时，返回 `endParagraph < startParagraph`（例如 `[1,0]`）表示空区间。
| `getFootnotes()` / `getEndnotes()` / `insertFootnote()` / `insertEndnote()` / `setNoteText()` / `deleteNote()` / `convertNote()` / `getNoteSettings()` / `setNoteSettings()` | 读取和编辑脚注/尾注、转换类型、调整编号设置 |
| `getComments()` / `addComment()` / `replyComment()` / `setCommentResolved()` / `setCommentText()` / `deleteComment()` | 读取和编辑批注、回复链与解决状态 |
| `revision` | 本实例的修订号；加载文件后从 0 开始，不持久化到 DOCX |
| `canUndo()` / `canRedo()` / `undo()` / `redo()` / `getHistory()` / `clearHistory()` | 访问撤销/重做栈；撤销与重做都会生成新的 `revision` |
| `beginHistoryGroup(label?)` / `endHistoryGroup()` | 将多次编辑显式合并为一个撤销步（例如格式刷批量操作） |

索引从 0 开始，包含主文档中的表格段落；结构变更后请重新读取快照。高层操作只处理主文档，页眉、页脚等部件请使用底层 API。
注释（脚注/尾注）`blocks` 里的段落 `index` 固定为 `-1`，不属于正文索引命名空间；注释内容请使用 `setNoteText(kind, id, text)` 编辑。`insertFootnote` / `insertEndnote` 为匹配 Word 常见显示，会在标记后以保留空格写入正文文本 run（例如读回 `" 内容"`）。
批注锚点通过 `CommentInfo.anchor.sourcePartPath` 指明所属部件；`paragraph` / `startParagraph` / `endParagraph` 都是该部件内部的局部顺序，不可直接拿去调用正文 `setParagraphText()` 之类的 API。回复批注会复用父批注锚点；没有正文锚点或没有 `comments.xml` 条目的记录会被标记为 `isOrphan: true`。
当前写入型批注 API（`addComment` / `replyComment` / `setCommentResolved` / `setCommentText` / `deleteComment`）以主文档为编辑入口；`getComments()` 会同时读出正文、页眉、页脚、脚注和尾注中的批注锚点。`getReviewers()` 按 `author` 聚合时使用以下作者桶常量：缺失作者为 `"(unattributed)"`、空串作者为 `"(empty author)"`、仅空白作者为 `"(blank author)"`；同一作者出现多个 `initials` 时，取出现次数最多的值，若并列则取最早出现的值（`initials` 仅来自批注；仅有修订而没有批注的作者不会带 `initials`）。当前 `getReviewers()` 只统计主文档修订与主文档锚点批注，页眉/页脚/脚注/尾注锚点的批注不计入。
逐次调用公开方法会按次记录历史（每步一次快照），连续编辑同一段落的文字会在短时间窗口内合并为一步；批量修改请走 `applyOperations`，一个批次只记录一步、也只拍一次快照——实测 600 次逐个 `insertParagraph` 明显慢于同样内容的单批操作，大批量场景请优先使用批次接口。若一次写入调用（包括非空 `applyOperations` 批次）最终没有产生任何部件字节变化，则该调用视为 no-op：不会推进 `revision`、不会新增撤销历史。撤销历史默认最多保留 50 步，且总快照字节默认上限 64 MiB；超过上限时会丢弃最旧步骤。`revision` 表示“变更次数”而不是“文档版本号”：执行 `undo()` / `redo()` 时 `revision` 依然单调递增，不会回退。

### 可视化组件

```ts
import { DocxDocument, DocxEditor } from 'docx-view';

const container = document.getElementById('editor')!;
const doc = DocxDocument.create();
const editor = new DocxEditor(container, doc, {
  onChange: snapshot => console.log(snapshot.revision),
  reviewFilter: { showRevisions: true, showComments: true, revisionView: 'markup' },
});

// 在外部 API 操作和导出前提交当前正在输入的内容。
editor.flush();
doc.insertParagraph('由 API 添加');
editor.render();

// 卸载组件时清理监听器。
// editor.destroy();
```

输入在段落失焦或调用 `flush()` 时提交；`onChange` 通知组件提交的修改。外部 API 修改后调用 `render()` 刷新。列表编号/项目符号和脚注/尾注引用标记会作为不可编辑的前缀渲染，段落正文文本本身不包含这些前缀；在演示界面中也可以通过 Tab / Shift+Tab 调整列表级别。不要在未 `flush()` 的情况下修改同一个文档的段落结构；也应避免在输入法组合输入期间切换文档或执行外部编辑。注意：`flush()` 仅在 `revisionView: 'markup'` 时提交文本，在 `'final'` / `'original'` 只读预览视图下会跳过提交。

组件还提供 `selectedParagraph`、`selectedRange`、`setDocument(doc)`、`setReviewFilter(filter)` 和 `destroy()`。`reviewFilter` / `setReviewFilter()` 支持按作者过滤审阅内容，并切换 `showRevisions`、`showComments`、`revisionView: 'final' | 'original' | 'markup'`（纯渲染状态，不修改文档）。其中 `revisionView: 'final' | 'original'` 为只读预览模式；编辑前请切回 `'markup'`。当前 `'original'` 视图还原插入/删除文本，不还原 `rPrChange` / `pPrChange` 的格式快照。`docx-selectionchange` 冒泡事件的 `detail.index` 是当前段落索引；`docx-rangechange` 的 `detail` 包含 `{ range, format }`（跨段落 `DocumentRange` 与 `getDocumentRangeFormat` 结果，可用于三态工具栏）。当选区跨越不同容器（如正文与表格单元格）时，`format` 会降级为空对象 `{}`。

组件使用 `.docx-editor`、`.docx-paragraph`、`.docx-table`、`.docx-image` 类名，不强制注入全局 CSS；宿主可以自行设置纸张外观、表格边框等，参考 `examples/style.css`。视图优先使用样式解析后的**有效格式**渲染常用字体、字号、颜色、加粗 / 斜体 / 下划线 / 删除线、上下标、大小写、高亮、字间距及段落缩进 / 间距 / 行距 / 对齐；图片的尺寸、旋转、翻转和裁剪也由组件渲染，浮动环绕采用简化布局。

### OOXML 细粒度操作

```ts
import { DocxDocument, WORD_NS } from 'docx-view';

const doc = DocxDocument.create();
console.log(doc.listParts());

// 主文档路径从包根关系解析，不假定为 word/document.xml。
doc.updatePartXml(doc.mainDocumentPath, xml => {
  const size = xml.getElementsByTagNameNS(WORD_NS, 'pgSz')[0];
  size?.setAttributeNS(WORD_NS, 'w:w', '16838');
  size?.setAttributeNS(WORD_NS, 'w:h', '11906');
});

// 任意现有 XML 部件（如 styles.xml、numbering.xml 或 .rels）均可读写。
const xml = doc.getPartXml(doc.mainDocumentPath);
doc.setPartXml(doc.mainDocumentPath, xml);

// 新增任意部件并登记 content type；关系由调用方在 .rels 中显式维护。
doc.addPart(
  'customXml/item1.xml',
  new TextEncoder().encode('<metadata xmlns="urn:example"><status>review</status></metadata>'),
  'application/xml',
);
```

- `getPartBytes(path)` 返回副本；`setPartBytes(path, bytes)` 替换已有部件，适用于图片等二进制数据。
- `getPartDocument(path)` 返回分离的 XML DOM；直接修改它不会更新包。使用 `updatePartXml` 回调提交修改；回调抛错则不提交。
- XML DOM 来自 `@xmldom/xmldom`，使用 `getElementsByTagNameNS`、`createElementNS` 等标准 DOM 操作；不提供浏览器 `querySelector` 或完整 SDK 强类型元素模型。
- 未修改部件的解压后字节原样保留；被修改的 XML 部件会重新序列化。ZIP 压缩结果和元数据不保证逐字节相同，数字签名不保证继续有效。
- 高层文字编辑只调整文字、制表和换行节点，尽量保留 run 属性、书签、绘图等未知内容；在 `trackChanges` 开启时支持自动写入修订标记，并可逐条或批量接受 / 拒绝修订。
- 底层 API 校验 XML 良构性及关键包结构，不进行完整 ECMA-376 模式校验。关系 ID、目标、content type、样式引用和合法元素顺序由调用方负责。

## AI Agent 接入

无需操作浏览器 DOM，Agent 可以读取结构化快照并提交声明式操作。通过导出的 `AGENT_OPERATION_SCHEMA` 为工具调用提供 JSON Schema，将模型返回的 JSON 交给运行时验证；**不执行 JavaScript 或模型生成的代码**。

```ts
import { DocxDocument, AGENT_OPERATION_SCHEMA } from 'docx-view';

const doc = DocxDocument.create();
const snapshot = doc.getSnapshot();
const tool = {
  name: 'edit_docx',
  description: '读取快照后，原子地修改当前 DOCX 文档',
  parameters: AGENT_OPERATION_SCHEMA,
};

const result = doc.applyOperations({
  expectedRevision: snapshot.revision,
  operations: [
    { type: 'setParagraphText', index: 0, text: '由 Agent 更新的标题' },
    { type: 'formatRun', paragraph: 0, run: 0, format: { bold: true } },
    { type: 'insertParagraph', text: '下一步：人工确认并导出。' },
  ],
});
console.log(tool, result.revision);
```

支持的操作类型：`setTrackChanges`、`setRevisionAuthor`、`acceptRevision`、`rejectRevision`、`acceptAllRevisions`、`rejectAllRevisions`、`setParagraphText`、`insertParagraph`、`deleteParagraph`、`formatParagraph`、`applyParagraphStyle`、`setParagraphNumbering`、`clearParagraphNumbering`、`setParagraphLevel`、`formatRun`、`formatRange`、`applyCharacterStyle`、`clearRangeFormat`、`formatDocumentRange`、`setOutlineLevel`、`moveOutlineSection`、`setParagraphTabs`、`setParagraphBorders`、`setParagraphShading`、`insertBreak`、`insertSymbol`、`replaceText`、`insertTable`、`insertTableAt`、`insertTableRow`、`deleteTableRow`、`insertTableColumn`、`deleteTableColumn`、`mergeCells`、`splitCell`、`formatTable`、`formatTableRow`、`formatCell`、`setCellText`、`insertHyperlink`、`updateHyperlink`、`removeHyperlink`、`insertBookmark`、`deleteBookmark`、`insertImage`、`replaceImageBytes`、`resizeImage`、`setImageAlt`、`deleteImage`、`setPartXml`、`insertFootnote`、`insertEndnote`、`setNoteText`、`deleteNote`、`convertNote`、`addComment`、`replyComment`、`setCommentResolved`、`setCommentText`、`deleteComment`、`undo`、`redo`。

- 请求中的所有操作在副本上顺序执行；任一操作失败，原文档和修订号不变。
- 成功的非空批次只增加一次修订号；空批次不增加。
- 每个操作的索引相对于该操作执行前的状态，前面的插入 / 删除可能改变后续索引。
- 推荐始终提供 `expectedRevision`，避免覆盖其他编辑者的修改；冲突时重新获取快照再生成请求。
- 这是本地文档实例 API，不自带 LLM 服务、MCP 服务、远程鉴权或多人协作。若包装为服务器工具，宿主应自行实现权限、资源隔离、审核与持久化；不要把文档内容当作可信指令。

## 设计文档

- [`docs/review-features.md`](docs/review-features.md) —— 审阅功能（修订与批注）的设计说明：现状实测、架构决策、issue 拆分与依赖、与各模块的集成点、以及最容易踩的工程约定。

## 支持范围与安全边界

当前可视化视图支持正文段落、常用样式继承、主题字体 / 主题色、段落与 run 的常见有效格式、基于 `numbering.xml` 的项目符号 / 编号列表、带 `w:gridSpan` / `w:vMerge`、显式边框 / 底纹、固定列宽、行高和单元格对齐的表格、常见 `w:drawing` / `w:pict` 图片、批注高亮与列表，以及分节页面设置近似和页眉页脚（默认 / 首页 / 偶数页）编辑；**不承诺与 Word 像素级一致或精确分页**。列表计数只在主文档正文（含表格单元格）内计算，支持常见 `numFmt`，未知格式回退为十进制。表格样式参与常用条件格式（`firstRow` / `lastRow` / `firstCol` / `lastCol` / `band1Horz` / `band2Horz`）的格式计算，其余条件样式尚未实现。浮动图片使用简化的浏览器布局：四周型 / 紧密型 / 穿越型映射为浮动，`topAndBottom` 映射为块级，`wrapNone` 映射为绝对定位；外部链接图片显示占位框且不会主动联网加载。当前已支持读取修订（插入、删除、`rPrChange` / `pPrChange` / `tblPrChange` / `trPrChange` / `tcPrChange`）、`trackChanges` 开关，以及常见文本 / 段落 / 图片 / 表格行编辑自动写入修订；这些部件 / XML 会尽量保留。尚未支持接受 / 拒绝修订或完整域值计算，低层 API 仍可操作。

支持普通 Transitional OOXML `.docx`，不支持加密文件、`.docm` 宏文档或 Strict OOXML。导入限制：ZIP 不超过 50 MiB、最多 2048 个条目、单部件解压后不超过 16 MiB、总解压大小不超过 64 MiB。批次最多 1000 个操作，单个文本参数最多 1,000,000 字符，表格最多 10,000 个单元格。剪贴板片段最多 1000 个段落、10,000 个 run、200 张图片，单个 run 文本最多 1,000,000 字符。

XML 禁止 DTD / 自定义实体声明，ZIP 路径禁止目录穿越。视图通过 DOM 文本节点和 `data:` URL 图片渲染，不将文档 XML 当作 HTML；编辑器剪贴板支持内部富文本与外部 HTML 映射，但 HTML 仅在分离文档中解析：`<script>/<style>`、事件属性、`javascript:` / `vbscript:` / `file:` / `data:` 链接都会被丢弃，`<img>` 仅接受 `data:` 形式的 PNG / JPEG / GIF / BMP（其余类型跳过该图，不影响同段其余内容），**不会主动请求外部 URL**。超链接一律经 `isSafeHyperlinkUrl()` 白名单（仅 http / https / mailto），编辑器与 `pasteClipboardFragment()` 公开 API 共用同一道校验。内部剪贴板携带段落的**样式 ID**但不迁移样式定义：跨文档粘贴时若目标文档未定义该样式，样式引用会悬空、显示回落到默认格式（run 的直接格式不受影响）；需要保真时请先在目标文档 `defineStyle()`。保留原始部件**不等于清除恶意内容**；下载文件中的外部链接、嵌入对象等仍需使用者按来源谨慎处理。大文档或不可信输入建议在 Web Worker / 隔离服务中处理。

测试覆盖 DOCX 往返、未修改部件保留、跨 run 替换、Unicode、样式链与主题解析、编号解析与创建、多级编号、style `numPr`、legal numbering、表格跨度解析、行列编辑、单元格合并 / 拆分、显式表格格式、分节、格式顺序、DOM 编辑、事务回滚、版本冲突、XML 校验、UTF-16、非标准主文档路径和 ZIP 解压限制。
