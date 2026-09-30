# docx-view

一个面向浏览器和 AI Agent 的 TypeScript / JavaScript DOCX 编辑组件库，包含无需后端的静态 `examples`。

**当前是可运行的基础版本，不是 Microsoft Word 排版引擎，也不等同于 .NET Open XML SDK 的完整实现。** 支持段落、文字格式、编号 / 项目符号列表、表格、常见图片、分节页面设置与页眉页脚编辑；对于更细粒度的操作，可以直接访问 DOCX 包中的部件、关系 XML 和命名空间感知的 OOXML DOM。OMML 公式支持读取和原生 MathML 渲染，并通过 API 提供线性文本兜底；不支持公式编辑。

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
| `getParagraphs()` / `getBlocks()` / `getSnapshot()` | 段落 / 表格结构、直接格式、有效格式、修订标记、样式清单、部件列表和修订号；`ParagraphInfo.text` 保留全部文字，隐藏 run 存在时另提供不含隐藏文字的 `visibleText` |
| `getContentControls()` | 读取内容控件类型、标题 / 标签、锁定状态、占位符、列表项、数据绑定及正文段落索引 |
| `setContentControlText(id, text)` / `setContentControlChecked(id, checked)` / `setContentControlProperties(id, patch)` / `removeContentControl(id, options?)` | 修改内容控件值或属性、移除控件包装；不提供新建内容控件 API |
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
| `restartNumbering(index, options?)` | 从段落处重新开始编号（默认 1；可用 `options.start` 指定起始值），并将后续连续列表项（含更深层级）切换到新编号实例；遇到不同 `numId` 或更高层级即停止；重复调用会新建实例并重新应用起始值。普通段落夹在列表中间不构成停止边界（与 Word 一致）；但**若中间插入了另一个列表的项，段落段在此处终止，其后的同列表项将继续原编号实例而非新实例——此处与 Word 不同**（仅在指定 `options.start` 时可观察到差异） |
| `continueNumbering(index)` | 将段落及后续连续列表项切回前面同抽象编号、同层级段落的编号实例；不存在前项时不做修改 |
| `formatRun(paragraph, run, format)` | 设置 run 直接格式，包括字符样式、字号、颜色、下划线、删除线、上下标等常用字段；将某个字段设为 `null` 可回退到继承样式 |
| `formatRange(range, format)` / `clearRangeFormat(range, fields?)` | 按段落内字符偏移格式化任意文本范围，支持清除全部或指定 run 直接格式字段 |
| `applyCharacterStyle(range, styleId, options?)` | 严格应用字符样式；可在应用时清除与字符样式冲突的直接格式 |
| `getRangeFormat(range)` | 读取字符范围内一致的 run 直接格式；同一字段在范围内不一致时返回 `undefined` |
| `copyFormat(range)` / `applyFormat(range, format)` | 格式刷 API：复制段落内范围的 run 直接格式，并应用到跨段落选区 |
| `formatDocumentRange(range, format)` / `getDocumentRangeFormat(range)` | 支持跨段落选区：首段部分 + 中间整段 + 末段部分 |
| `getEditableRegions()` / `addEditableRegion(range, options)` / `removeEditableRegion(id)` | 读取、创建和移除正文区域编辑权限标记；授权方式为编辑组或用户 ID |
| `copyClipboardFragment(range)` / `pasteClipboardFragment(range, fragment)` | 内部富文本剪贴板 API：run 直接格式、超链接、内嵌图片、段落格式与样式引用、编号（含多级）和表格；粘贴按落点切分段落，图片按目标文档关系与 media 部件重建（不复用源 `rId`） |
| `defineStyle(style)` | 创建或更新 `styles.xml` 样式定义；缺少部件时自动补内容类型与主文档关系 |
| `createStyleFromSelection(range, style)` | 从当前选区提炼段落 / 字符直接格式并写成新的段落样式定义 |
| `getOutline()` / `setOutlineLevel(index, level)` / `moveOutlineSection(from, to)` | 读取标题大纲、显式设置 `outlineLevel`，以及按大纲整体移动标题节 |
| `replaceText(search, replacement)` | 正文及表格段落内的字面替换，支持跨 run 匹配，不跨段落 |
| `DocxDocument.compare(base, revised, options?)` | 比较两份主文档并返回新的带修订实例；纯文本/格式差异细化到段落与 run，表格及含图片的段落按粗粒度删除 + 插入处理 |
| `getRevisions(filter?)` / `acceptRevision(id)` / `rejectRevision(id)` / `acceptAllRevisions(filter?)` / `rejectAllRevisions(filter?)` | 读取并逐条/批量接受或拒绝修订（支持按作者筛选）；`moveFrom` / `moveTo` 以 `kind: 'move'` 返回，并暴露 `{ move: { name, side, pairedId? } }` |
| `getReviewers()` | 聚合主文档修订与主文档锚点批注的审阅者统计（修订数、批注数、未解决批注数、时间范围）；每项返回 `{ kind, author? }`（`kind`: `named` / `unattributed` / `empty` / `blank`，四种都可实际构造），结果默认按计数降序，再按 `kind+author` 稳定排序 |
| `insertTable(rows)` / `insertTableAt(rows, cols, before?, format?)` | 在正文中插入表格；支持空白表格、基础表格格式和正文块级定位 |
| `getTable(index)` | 读取正文中第 N 个表格的 grid、跨度和表格/行/单元格格式信息 |
| `insertTableRow()` / `deleteTableRow()` / `insertTableColumn()` / `deleteTableColumn()` | 行列编辑；同步维护 `w:tblGrid`，拒绝删成 0 行或 0 列 |
| `mergeCells()` / `splitCell()` | 基于逻辑网格合并/拆分单元格，读写 `w:gridSpan` / `w:vMerge` |
| `formatTable()` / `formatTableRow()` / `formatCell()` | 设置表格宽度、布局、边框、底纹、行高、标题行、单元格对齐和边距等显式属性 |
| `setCellText()` | 修改可见单元格文字，同时保留段落结构与其他未改内容 |
| `getImages()` / `getImageBytes()` / `getImageDataUrl()` | 读取主文档中的图片元数据、二进制内容和可直接渲染的 `data:` URL |
| `getShapes()` / `getShapeParagraphs(shapeId)` | 读取文本框、纯形状、SmartArt、图表等元数据及文本框独立文字流；形状由 SVG 绘制，文字流仍独立只读 |
| `insertImage()` / `replaceImageBytes()` / `resizeImage()` / `setImageAlt()` / `deleteImage()` | 插入、替换、调整尺寸、更新替代文本和删除图片 |
| `getSections()` / `getSection(index)` | 读取分节类型、纸张、页边距、分栏、页眉页脚引用 |
| `setPageSetup(section, setup)` | 修改指定节的纸张方向、边距、分栏和页码起始等页面设置 |
| `insertSectionBreak(paragraph, type)` / `deleteSectionBreak(section)` | 插入/删除分节符（删除为显式 API，保留原有段落删除保护） |
| `getHeaderBlocks()` / `getFooterBlocks()` | 读取页眉页脚 block 结构（段落、表格等） |
| `createHeader()` / `createFooter()` / `setHeaderText()` / `setFooterText()` | 创建并写入页眉页脚部件，自动维护 rels 与 content-types |
| `insertPageNumberField(partPath, options)` | 在页眉/页脚部件写入 `PAGE` 或 `NUMPAGES` 域占位结构 |
| `getFields(partPath?)` / `updateFields(options?)` | 按部件读取域；文档层可接收调用方提供的分页结果，写回 `PAGE` / `NUMPAGES` / `PAGEREF` 与受限的 `TOC` 域结果 |
| `getDocumentProperties()` / `setDocumentProperties(patch)` | 读取和按字段更新 `docProps/core.xml` / `docProps/app.xml` 中的常用文档属性；缺失部件时自动补包级关系与 content-type，未知元素和未改字段原样保留 |
| `getDocumentProtection()` / `setDocumentProtection(value)` | 读取和写入 `settings.xml` 中的 `w:documentProtection`（如只读 / 仅批注 / 仅修订 / 表单）；仅修改 `edit` / `enforcement`，保留已有 hash/salt 等密码相关属性 |
| `getSettings()` / `setTrackChanges(enabled)` / `setRevisionAuthor(author)` | 读取常用文档设置（当前返回 `{ defaultTabStop, evenAndOddHeaders, trackChanges }`），显式开启/关闭 `w:trackChanges`（关闭时写 `w:val="0"`，不删除元素），并设置后续记录修订写入使用的作者名 |
| `revision` | 本实例的修订号；加载文件后从 0 开始，不持久化到 DOCX |

**索引与作用域**

`getParagraphs()` 和 `getBlocks()` 只包含正文段落；文本框（`w:txbxContent`）里的段落属于独立的只读文字流，通过 `getShapeParagraphs(shapeId)` 访问，不占用正文段落下标。正文按下标写入不会穿透到文本框内部。

形状 SVG 渲染支持 `rect`、`roundRect`、`ellipse`、`triangle`、`rtTriangle`、`diamond`、`parallelogram`、`trapezoid`、`pentagon`、`hexagon`、`star5`、`rightArrow`、`leftArrow`、`upArrow`、`downArrow`、`leftRightArrow`、`line`、`straightConnector1`、`wedgeRectCallout` 和 `cloudCallout` 这 20 种 DrawingML 预设几何。其它预设以及含未支持指令的自定义几何以矩形绘制，但保留已读出的填充和线条；自定义路径目前处理首个 `a:path` 中的 `moveTo`、`lnTo`、`cubicBezTo` 与 `close`。支持实体色、线性渐变、透明度 `alpha`、线宽/虚线、旋转和翻转；未指定填充或线条的形状使用浅色填充与细描边作为可见兜底。形状样式只取 `wps:style` 的 `fillRef` / `lnRef` 颜色，忽略 `idx` 指向的主题格式表渐变和效果预设；`lumMod`、`lumOff` 等其他颜色变换不处理。

图表 SVG 支持柱状 / 条形、折线、饼图、圆环、面积和散点图，数值只读取 `c:numCache` / `c:numLit` / `c:strCache`，不解析公式或内嵌工作簿，也不会联网读取 `c:externalData`。横向条形图将类目标签绘制在左侧、数值刻度绘制在底部。折线与面积图在缺失数据点处断开，面积填充按每一段连续数据各自向零值线闭合，缺口区间不填充。缺失缓存的序列和不支持的 3D、雷达、曲面、股价、气泡、趋势线、误差线图表退化为占位；组合图和次坐标轴中的序列会合并到同一主坐标轴绘制，因此是近似结果。类目轴绘制缓存中的类目标签；散点图使用缓存的 `c:xVal` 数值范围定位横坐标。图表填充和线条沿用 DrawingML 外观与主题色解析，未指定序列外观时使用当前文档主题的 `accent1`~`accent6`。

VML 支持 `v:shape`、`v:rect`、`v:oval`、`v:line` 的基础填充色/线条，以及内置类型 1、2、3、4、5、202、203 的常见几何映射；`v:path` 不解析，退化为矩形并保留颜色。形状 `blipFill` 只显示包内可渲染图片；外部链接图片不联网，以浅色虚线框占位。

SmartArt 使用 Word 预渲染的 `diagrams/drawing*.xml` 形状绘制；缺少该部件时显示占位，子形状文字仅支持纯文本、不支持富文本。

`getFields(partPath)` 的段落与 run 索引只在指定部件内有效；页眉、页脚的索引不能用于正文数组。`DocxDocument.updateFields({ pagination })` 接受由调用方计算的页数、段落页索引与显示页码映射，核心文档 API 不依赖浏览器排版。

- 索引从 0 开始，包含主文档中的表格段落；结构变更后请重新读取快照。
- 高层操作默认处理主文档，可通过页眉页脚 API 读写 `header*.xml` / `footer*.xml`；`getRevisions()`、`accept*`/`reject*`、`getReviewers()` 与 `DocxDocument.compare()` 当前都只作用于主文档。
- `insertSectionBreak(paragraph)` 的 `paragraph` 表示“该段落结束处插入分节”；`deleteSectionBreak(section)` 删除第 `section` 节末尾的分节符并与下一节合并。节范围是闭区间：`startParagraph <= i <= endParagraph`。当某节暂时没有段落时，返回 `endParagraph < startParagraph`（例如 `[1,0]`）表示空区间。

**修订的读取**

- 当段落中存在未接受的删除 (`w:del` / `w:moveFrom`) 时，其文本不会进入 `paragraph.text`，但对应 run 仍保留在 `runs[]` 中并以空字符串占位；删除内容请通过 `getRevisions().deletedText` 读取。
- `ParagraphInfo.text` 始终包含 `w:vanish` / `w:webHidden` 文字，以保留 `setParagraphText(index, paragraph.text)` 往返；存在隐藏文字且可见内容不同时，`visibleText` 提供排除隐藏 run 后的文字。`w:vanish` / `w:webHidden` 只是显示属性，不是安全机制或访问控制。
- `getRevisions({ kinds })` 对 `kind` 严格匹配：移动修订只会命中 `kinds: ['move']`，不再包含在 `insertion` / `deletion` 过滤中。
- `getRevisions().author` 保留修订标记里的原始 `w:author`：缺失时为 `undefined`，空串为 `''`，仅空白字符串按原样保留；审阅者身份分桶时，`named` 会用 `author.trim()` 归一（例如 `' Alice '` 与 `'Alice'` 归为同一作者），`empty` / `blank` / `unattributed` 规则不变；`getRevisions({ authors })` / `getComments({ authors })` / `acceptAllRevisions({ authors })` / `rejectAllRevisions({ authors })` 使用同一归一规则。
- `RevisionInfo.move.pairedId` 与接受/拒绝逻辑使用同一配对规则（优先范围标记，其次文档顺序配对同名 `moveFrom`/`moveTo`）。

**修订的写入与接受 / 拒绝**

- 记录修订开启后，文本/段落/图片/表格行等编辑会写入 `w:ins` / `w:del` / `rPrChange` / `pPrChange`；此时 `deleteParagraph()` 会保留原段落节点并把内容标记为删除，因此按索引循环删除时应在每步后重新读取段落列表。
- 若未显式调用 `setRevisionAuthor(author)`，默认作者名为 `docx-view`。
- `acceptRevision(id)` / `rejectRevision(id)` 传入移动修订任一半时会成对处理 `moveFrom` + `moveTo`；批量按作者筛选时，只要配对中的任一半命中过滤条件，整对都会一起处理。
- 接受“仅段落标记删除（`w:pPr/w:rPr/w:del`）”时会优先合并到同容器中的下一段；若当前段/下一段含 `sectPr`、当前段后继不是段落、或已到容器末尾，则降级为仅移除该删除标记。

**文档比较**

- `compare()` 返回的是新实例（`revision` 从 0 开始），不会修改输入文档；它对纯段落文字/格式做细粒度修订，对表格和含图片的段落仅做粗粒度删除 + 插入，不比较页眉/页脚/脚注/尾注。
- 超过 1000 个主文档段落的输入会直接抛错，避免段落级对齐的 O(N²) 内存开销。

**文档属性与保护**

- `getDocumentProperties().revisionNumber` / `setDocumentProperties({ revisionNumber })` 读写的是 DOCX `cp:revision` 文档属性，**与**实例级 `doc.revision`（内存中的变更计数，不持久化到 DOCX）互不联动。
- `documentProtection` 只是文档内声明：库会如实读写它，但**不会**因为 `readOnly` / `comments` / `trackedChanges` 而禁用编辑 API；若宿主需要据此调整按钮或 UI，请自行在外层实现。
- 内容控件的 `w:lock` 同样只是文档内声明；只有 `setContentControlText()` 在 `contentLocked` / `sdtContentLocked` 时拒绝写入。它不会禁用 `setParagraphText()` 等既有编辑 API。
- `w:permStart` / `w:permEnd` 区域编辑权限同样只是文档内声明，不是安全机制；`getEditableRegions()` 如实读取声明，但不会禁用 `setParagraphText()` 或其他编辑 API。调用方可自行据此调整 UI，不能把这些标记当作访问控制。
- `getEditableRegions()` 按正文标记文档顺序以 `w:id` 配对；重复 ID 按出现顺序配对，只有一端的标记会以 `unpaired: 'startOnly'` 或 `'endOnly'` 返回。

| `getFootnotes()` / `getEndnotes()` / `insertFootnote()` / `insertEndnote()` / `setNoteText()` / `deleteNote()` / `convertNote()` / `getNoteSettings()` / `setNoteSettings()` | 读取和编辑脚注/尾注、转换类型、调整编号设置 |
| `getComments()` / `addComment()` / `replyComment()` / `setCommentResolved()` / `setCommentText()` / `deleteComment()` | 读取和编辑批注、回复链与解决状态 |
| `revision` | 本实例的修订号；加载文件后从 0 开始，不持久化到 DOCX |
| `canUndo()` / `canRedo()` / `undo()` / `redo()` / `getHistory()` / `clearHistory()` | 访问撤销/重做栈；撤销与重做都会生成新的 `revision` |
| `beginHistoryGroup(label?)` / `endHistoryGroup()` | 将多次编辑显式合并为一个撤销步（例如格式刷批量操作） |

索引从 0 开始，包含主文档中的表格段落；结构变更后请重新读取快照。内容控件的 `paragraphs` 沿用同一正文索引，嵌套控件各自返回且内层标记 `nested: true`。高层操作只处理主文档，页眉、页脚等部件请使用底层 API。
注释（脚注/尾注）`blocks` 里的段落 `index` 固定为 `-1`，不属于正文索引命名空间；注释内容请使用 `setNoteText(kind, id, text)` 编辑。`insertFootnote` / `insertEndnote` 为匹配 Word 常见显示，会在标记后以保留空格写入正文文本 run（例如读回 `" 内容"`）。
批注锚点通过 `CommentInfo.anchor.sourcePartPath` 指明所属部件；`paragraph` / `startParagraph` / `endParagraph` 都是该部件内部的局部顺序，不可直接拿去调用正文 `setParagraphText()` 之类的 API。回复批注会复用父批注锚点；没有正文锚点或没有 `comments.xml` 条目的记录会被标记为 `isOrphan: true`。
**批注的读写范围**

- 当前写入型批注 API（`addComment` / `replyComment` / `setCommentResolved` / `setCommentText` / `deleteComment`）以主文档为编辑入口；`getComments()` 会同时读出正文、页眉、页脚、脚注和尾注中的批注锚点。

**审阅者聚合**

- `getReviewers()` 返回原始 `author`（可为 `undefined` / `''` / 空白字符串）+ `kind` 判别字段，不再用字符串哨兵替代作者值；修订与批注两侧都会按同一规则产出 `named` / `unattributed` / `empty` / `blank`，其中 `named` 会按 `trim()` 归一作者身份（`' Alice '` 与 `'Alice'` 合并为同一审阅者），因此真实作者名即使等于 `"(unattributed)"` / `"(empty author)"` / `"(blank author)"` 也不会与缺失/空串/空白作者合并。
- 同一作者出现多个 `initials` 时，取出现次数最多的值，若并列则取最早出现的值（`initials` 仅来自批注；仅有修订而没有批注的作者不会带 `initials`）。
- 对成对的移动修订（`moveFrom` + `moveTo`），`getReviewers().revisionCount` 按 1 条计数。
- 当前 `getReviewers()` 只统计主文档修订与主文档锚点批注，页眉/页脚/脚注/尾注锚点的批注不计入。

**筛选入参（破坏性变更）**

- `DocxEditor.setReviewFilter({ authors })` 也同步改为传 `ReviewerFilterAuthor[]`：`named` 必须带非空 `author`，`unattributed` 不带 `author`，`empty` 必须传 `author: ''`，`blank` 必须传仅空白的 `author`。
- 旧的字符串数组写法需迁移。
逐次调用公开方法会按次记录历史（每步一次快照），连续编辑同一段落的文字会在短时间窗口内合并为一步；`setParagraphText()` 的历史项使用段落级增量，避免为纯文字编辑序列化整篇文档；其他公开写入方法仍按完整部件快照记录。批量修改请走 `applyOperations`，一个批次只记录一步、也只拍一次快照——实测 600 次逐个 `insertParagraph` 明显慢于同样内容的单批操作，大批量场景请优先使用批次接口。若一次写入调用（包括非空 `applyOperations` 批次）最终没有产生任何部件字节变化，则该调用视为 no-op：不会推进 `revision`、不会新增撤销历史。撤销历史默认最多保留 50 步，且总快照字节默认上限 64 MiB；超过上限时会丢弃最旧步骤。`revision` 表示“变更次数”而不是“文档版本号”：执行 `undo()` / `redo()` 时 `revision` 依然单调递增，不会回退。

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

**组件成员**

- 组件还提供 `selectedParagraph`、`selectedRange`、`setDocument(doc)`、`setReviewFilter(filter)`、`acceptRevision(id)`、`rejectRevision(id)`、`acceptAllRevisions(filter?)`、`rejectAllRevisions(filter?)`、`focusRevision(id)`、`focusNextRevision()`、`focusPreviousRevision()` 和 `destroy()`。

**文字显示**

- `showHiddenText` 缺省为 `false`，隐藏 run 默认不显示；启用后会以虚线下划线标记并显示。隐藏 run 在默认不显示时仍保留，并在编辑同段落时原样写回。

**审阅筛选与视图**

- `reviewFilter` / `setReviewFilter()` 支持按作者过滤审阅内容，并切换 `showRevisions`、`showComments`、`revisionView: 'final' | 'original' | 'markup'`（纯渲染状态，不修改文档）。
- 其中 `revisionView: 'final' | 'original'` 为只读预览模式；需要直接编辑正文文字时请切回 `'markup'`。
- 接受/拒绝修订会直接修改文档，因此在 `'final'` / `'original'` 视图下同样可用（与可编辑正文无关）。
- 当前 `'original'` / `'final'` 视图除插入/删除外，也会对移动修订显示对应一侧（`original` 显示 `moveFrom`，`final` 显示 `moveTo`）；`'original'` 仍不还原 `rPrChange` / `pPrChange` 的格式快照。

**只读预览模式下的界面行为**

- 演示功能区中的正文格式、表格、图片和批注写入命令，以及右键菜单中的写入命令，在 `'final'` / `'original'` 下禁用；审阅筛选、修订导航、复制/打开链接、批注定位等浏览操作仍可用。
- 接受/拒绝修订在功能区和右键菜单中仍可用；这些操作直接修改文档，而不依赖正文是否可编辑。上下文选项卡仍按选区显示，其写入命令禁用。
- 切换 `revisionView` 只改变预览与可用控件，不修改文档或触发 `onChange`。

**选区事件**

- `docx-selectionchange` 冒泡事件的 `detail.index` 是当前段落索引；`docx-rangechange` 的 `detail` 包含 `{ range, format }`（跨段落 `DocumentRange` 与 `getDocumentRangeFormat` 结果，可用于三态工具栏）。
- 当选区跨越不同容器（如正文与表格单元格）时，`format` 会降级为空对象 `{}`。
- 示例编辑器的右键菜单遵循 Word 的选区规则：右键落在现有选区内时保留选区，落在选区外时先把插入点移动到点击处。
- 编辑区内默认使用自定义菜单；`Shift` + 右键放行浏览器原生菜单。菜单也可通过 `Shift+F10` 或 `ContextMenu` 键打开，关闭后恢复编辑区焦点与选区。

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

支持的操作类型（按 `AGENT_OPERATION_SCHEMA` 顺序）：`setTrackChanges`、`setRevisionAuthor`、`acceptRevision`、`rejectRevision`、`acceptAllRevisions`、`rejectAllRevisions`、`setParagraphText`、`insertParagraph`、`deleteParagraph`、`formatParagraph`、`applyParagraphStyle`、`setParagraphNumbering`、`clearParagraphNumbering`、`setParagraphLevel`、`restartNumbering`、`continueNumbering`、`formatRun`、`formatRange`、`applyCharacterStyle`、`clearRangeFormat`、`formatDocumentRange`、`setOutlineLevel`、`moveOutlineSection`、`setParagraphTabs`、`setParagraphBorders`、`setParagraphShading`、`insertBreak`、`insertSymbol`、`replaceText`、`insertTable`、`insertTableAt`、`insertTableRow`、`deleteTableRow`、`insertTableColumn`、`deleteTableColumn`、`mergeCells`、`splitCell`、`formatTable`、`formatTableRow`、`formatCell`、`setCellText`、`insertHyperlink`、`updateHyperlink`、`removeHyperlink`、`insertBookmark`、`deleteBookmark`、`updateFields`、`insertField`、`setContentControlText`、`setContentControlChecked`、`setContentControlProperties`、`removeContentControl`、`addEditableRegion`、`removeEditableRegion`、`insertImage`、`replaceImageBytes`、`resizeImage`、`setImageAlt`、`deleteImage`、`setPartXml`、`insertFootnote`、`insertEndnote`、`setNoteText`、`deleteNote`、`convertNote`、`addComment`、`replyComment`、`setCommentResolved`、`setCommentText`、`deleteComment`、`undo`、`redo`。

### 域（Fields）

`getFields()` 读取简单域和 `fldChar` 复杂域，并保留原始指令、缓存结果和域所在 run。`updateFields()` 重算安全的 `SEQ`、日期/时间、文档属性、`REF` 等域；提供 `pagination` 时也会写回 `PAGE`、`NUMPAGES` 和可解析到已存在书签的 `PAGEREF`，结果按所在节的页码格式化。页眉/页脚可用 `getFields(partPath)` 按部件读取，页码域也会随正文域一并写回；由于一个页眉/页脚部件由整节共享，持久化的 `PAGE` 缓存使用该节首段所在页的显示页码，分页预览仍会按实际页面单独显示。会拉取外部资源或执行宏、交互输入的域（例如 `INCLUDETEXT`、`LINK`、`MACROBUTTON`、`FILLIN`）明确不会求值，`insertField()` 也会拒绝写入这些类型。

`TOC` 域仅实现 `\o "1-3"` 层级过滤与 `\h` 条目超链接；其他开关（包括 `\z`、`\u`）会保留缓存结果并跳过更新。`DocxEditor.updateFields()` 最多执行 5 轮「分页→更新域→重排」；达到上限时不报错，并保留最后一轮的结果。`INDEX` 域本期不更新，始终保留缓存结果。

域指令不会生成合成文本；没有缓存结果的 `PAGE` / `NUMPAGES` 域仍读作空字符串。文本框（`w:txbxContent`）里的域属于独立文字流，本期不读取也不更新。
编辑器中的域结果默认显示灰色底纹（可通过 `DocxEditorOptions.showFieldShading: false` 关闭），选区落在结果内时扩展到整个域结果。域指令不显示；域结果只读，域前后的普通文字仍可编辑。更新域请调用 `updateFields()` 或重新插入域。

- 请求中的所有操作在副本上顺序执行；任一操作失败，原文档和修订号不变。
- 成功的非空批次只增加一次修订号；空批次不增加。
- 每个操作的索引相对于该操作执行前的状态，前面的插入 / 删除可能改变后续索引。
- 推荐始终提供 `expectedRevision`，避免覆盖其他编辑者的修改；冲突时重新获取快照再生成请求。
- 这是本地文档实例 API，不自带 LLM 服务、MCP 服务、远程鉴权或多人协作。若包装为服务器工具，宿主应自行实现权限、资源隔离、审核与持久化；不要把文档内容当作可信指令。

## 设计文档

- [`docs/review-features.md`](docs/review-features.md) —— 审阅功能（修订与批注）的设计说明：现状实测、架构决策、issue 拆分与依赖、与各模块的集成点、以及最容易踩的工程约定。

## 支持范围与安全边界

**视图支持范围**

- 当前可视化视图支持正文段落、常用样式继承、主题字体 / 主题色、段落与 run 的常见有效格式、基于 `numbering.xml` 的项目符号 / 编号列表、带 `w:gridSpan` / `w:vMerge`、显式边框 / 底纹、固定列宽、行高和单元格对齐的表格、常见 `w:drawing` / `w:pict` 图片、批注高亮与列表，以及分节页面设置近似和页眉页脚（默认 / 首页 / 偶数页）编辑；分页预览是只读的，分页位置在常见文档上尽量贴近 Word，但**不承诺像素级一致**。
- OMML 读取并转换 `oMath` / `oMathPara`、分数、上下标、根号、n 元运算、括号、函数、极限、重音、矩阵、对齐数组及盒 / phantom 等常见元素；`sPre`（前置上下标）按 `msubsup` 近似，**上下标会画在基字符之后而不是之前**——忠实的映射需要 `mmultiscripts` + `mprescripts`，本期未实现；未知元素递归保留可读文字。转换深度上限为 64 层，公式依赖浏览器原生 MathML，`getMath()` 同时提供线性文本和结构化 MathML 数据。文本框 / 形状内公式目前不纳入 `getMath()`，也不在形状文字渲染中显示。
- 分页预览按栏宽重新度量内容，支持等宽 / 指定宽度分栏与 `nextColumn`，并按表格行跨页 / 跨栏拆分；连续的 `w:tblHeader` 标题行会在每个片段重复，`cantSplit` 行保持完整。
- run 着重号支持 `w:em` 的 `dot`、`comma`、`circle`、`underDot`（分别使用浏览器原生 `text-emphasis`）；显式 `none` 可关闭继承的着重号。不按竖排文字方向调整着重号位置。
- 读取节的 `w:docGrid` `type`、`linePitch` 与 `charSpace`；`lines`、`linesAndChars`、`snapToChars` 的 `linePitch` 用于分页行高吸附，`default` 不吸附。`charSpace` 目前只读取并保留，不参与字符宽度计算。
- 分页预览按节套用页面设置：每一页使用所在节的纸张宽度、页边距、方向、分栏与页眉页脚。连续视图（`viewMode: 'continuous'`，默认）是一条不分页的滚动流，整篇文档**只套用第一节**（`getSection(0)`）的纸张宽度、页边距、方向与分栏数 / 栏间距，页眉页脚也只显示第一节的，其余节的页面设置不会反映在连续视图中——例如「纵向正文 + 一节横向宽表格」的文档，横向那节会按第一节的宽度渲染，分栏数不同的节也按第一节的栏数排。查看多节文档的真实版式请切换到分页视图（`setViewMode('paginated')`）；文档中各节的页面设置本身不受影响，导出时原样保留。

**列表与表格**

- 列表计数只在主文档正文（含表格单元格）内计算，支持常见 `numFmt`，未知格式回退为十进制。
- 表格样式参与常用条件格式（`firstRow` / `lastRow` / `firstCol` / `lastCol` / `band1Horz` / `band2Horz`）的格式计算，其余条件样式尚未实现。
- 跨越分页或分栏边界的 `rowSpan` 合并区域会连同其覆盖的行整体移到下一页或下一栏，不会拆成断开的单元格；超高的单行或合并区域仍会完整放置并允许超出可用高度。

**图片布局**

- 浮动图片使用简化的浏览器布局：四周型 / 紧密型 / 穿越型映射为浮动并缩窄相邻文字行，`topAndBottom` 映射为占满整行的块级区域，`wrapNone` 映射为绝对定位且不影响分页；外部链接图片显示占位框且不会主动联网加载。

**修订与域**

- 当前已支持读取修订（插入、删除、`rPrChange` / `pPrChange` / `tblPrChange` / `trPrChange` / `tcPrChange`）、`trackChanges` 开关，以及常见文本 / 段落 / 图片 / 表格行编辑自动写入修订；这些部件 / XML 会尽量保留。
- 接受 / 拒绝修订已支持（逐条、批量、按作者筛选，移动修订成对处理）。域值计算见上文「域」一节：`SEQ`、日期 / 时间、文档属性、`REF` 会重算，提供 `pagination` 时 `PAGE` / `NUMPAGES` / `PAGEREF` 也会写回，`TOC` 仅实现 `\o` 层级过滤与 `\h` 超链接；`INDEX` 始终保留缓存结果。低层 API 仍可直接操作。
- 编辑器标记域 run 的 `data-docx-field` / `data-docx-field-role`；结果灰底可关闭，结果不可直接编辑（不是“支持编辑域”）。

支持普通 Transitional OOXML `.docx`，不支持加密文件、`.docm` 宏文档或 Strict OOXML。导入限制：ZIP 不超过 50 MiB、最多 2048 个条目、单部件解压后不超过 16 MiB、总解压大小不超过 64 MiB。批次最多 1000 个操作，单个文本参数最多 1,000,000 字符，表格最多 10,000 个单元格。剪贴板片段最多 1000 个段落、10,000 个 run、200 张图片，单个 run 文本最多 1,000,000 字符。`setDocumentProperties()` 仅校验并写入常用 `docProps` 字段；`app.xml` 中 Pages / Words / Characters / Lines / Paragraphs 等统计值不会自动重算。

**解析边界**

- XML 禁止 DTD / 自定义实体声明，ZIP 路径禁止目录穿越。

**渲染与剪贴板**

- 视图通过 DOM 文本节点和 `data:` URL 图片渲染，不将文档 XML 当作 HTML；编辑器剪贴板支持内部富文本与外部 HTML 映射，但 HTML 仅在分离文档中解析：`<script>/<style>`、事件属性、`javascript:` / `vbscript:` / `file:` / `data:` 链接都会被丢弃，`<img>` 仅接受 `data:` 形式的 PNG / JPEG / GIF / BMP（其余类型跳过该图，不影响同段其余内容），**不会主动请求外部 URL**。
- 超链接一律经 `isSafeHyperlinkUrl()` 白名单（仅 http / https / mailto），编辑器与 `pasteClipboardFragment()` 公开 API 共用同一道校验。
- 内部剪贴板携带段落的**样式 ID**但不迁移样式定义：跨文档粘贴时若目标文档未定义该样式，样式引用会悬空、显示回落到默认格式（run 的直接格式不受影响）；需要保真时请先在目标文档 `defineStyle()`。

**使用者的责任**

- 保留原始部件**不等于清除恶意内容**；下载文件中的外部链接、嵌入对象等仍需使用者按来源谨慎处理。
- 大文档或不可信输入建议在 Web Worker / 隔离服务中处理。

测试覆盖 DOCX 往返、未修改部件保留、跨 run 替换、Unicode、样式链与主题解析、编号解析与创建、多级编号、style `numPr`、legal numbering、表格跨度解析、行列编辑、单元格合并 / 拆分、显式表格格式、分节、格式顺序、DOM 编辑、事务回滚、版本冲突、XML 校验、UTF-16、非标准主文档路径和 ZIP 解压限制。
