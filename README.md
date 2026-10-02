# docx-view

一个面向浏览器和 AI Agent 的 TypeScript / JavaScript DOCX 编辑组件库，包含无需后端的静态 `examples`。

**当前是可运行的基础版本，不是 Microsoft Word 排版引擎，也不等同于 .NET Open XML SDK 的完整实现。** 支持段落、文字格式、编号 / 项目符号列表、表格、常见图片、分节页面设置与页眉页脚编辑；对于更细粒度的操作，可以直接访问 DOCX 包中的部件、关系 XML 和命名空间感知的 OOXML DOM。OMML 公式支持读取、原生 MathML 渲染，以及通过 API 写入结构化 MathML 数据或受限的线性文本；不支持在公式内部直接进行 WYSIWYG 编辑。

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

### 把存盘的压缩步骤搬到 Worker

存盘有两步：把改过的部件序列化成字节，再把字节打成 ZIP。**序列化必须留在主线程**——它要读那棵活的 XML 树，而 xmldom 节点不是结构化可克隆的，传不进 Worker。**ZIP（DEFLATE）只依赖字节，可以搬**。1500 段文档实测：整条链约 138 ms，其中 DEFLATE 约 92 ms（占 37%~67%，随文档而变）。

`toUint8Array()` / `toBlob()` 接受 `{ zip }`，由宿主决定在哪儿压。库**不替宿主创建 Worker**：那要替它决定打包器与 CSP 策略。宿主侧的 worker 就这几行：

```js
// zip-worker.js
import { zipParts } from 'docx-view';
self.onmessage = async (event) => {
  self.postMessage(await zipParts(event.data));
};
```

```js
const worker = new Worker(new URL('./zip-worker.js', import.meta.url), { type: 'module' });
const bytes = await doc.toUint8Array({
  zip: (parts) => new Promise((resolve) => {
    worker.onmessage = (event) => resolve(event.data);
    worker.postMessage(parts);   // 结构化克隆会复制字节，本文档的部件不受影响
  }),
});
```

注意三点：

- **不要转移 `parts` 里的 `ArrayBuffer`**（`postMessage(parts, [buffer])`）。那会把主线程这边的字节置为分离状态，文档随后就用不了了。结构化克隆的复制开销很小（实测交接约 0.4~10 ms）。
- 收益是**主线程占用减半**，不是存盘变快：1500 段上主线程占用从约 57 ms 降到约 27~36 ms（剩下的是必须留在主线程的序列化），而墙钟因为多了一次往返反而略长。要的是「存盘时界面不卡」，不是「存盘更快」。
- `zip` 必须返回 `Uint8Array`，否则 `toUint8Array()` 直接抛错——不然问题会推迟到宿主写文件时才暴露。

### 其它可以搬进 Worker 的批处理

`compareDocxBytes(base, revised, options?)`（比较两份 .docx，返回带修订的结果文档字节）、`readDocxSnapshot(bytes)`（只读打开，交回快照）、`searchDocxText(bytes, query, options?)`（全文搜索）都是字节进、纯数据或字节出，两头都过得了结构化克隆，宿主在 Worker 里 import 后照上面的方式收发即可。可编辑的文档对象搬不了：编辑模型是活的 XML 树，传不过线程边界；分页测量要 `getClientRects()`，只有主线程有。


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

分页与连续视图使用同一套段落间距计算：相邻段落的 `before` / `after` 取较大值（首段的 `before` 和末段的 `after` 仍保留）。`contextualSpacing` 在相邻段落样式相同处抑制该间距；`beforeLines` / `afterLines` 和 `autospacing` 会无损读取并写回，但当前不参与排版（行单位需要实际行高，autospacing 的 Word 算法也不在布局度量器中）。分页视图读取节级行号和页面边框；行号按分页后的行顺序（跨栏按流项目顺序）计算，`suppressLineNumbers` 的段落跳过且不占号，连续视图不显示行号。节的 `vAlign` 支持 `top`、`center`、`bottom`；`both` 当前退化为 `top`。

`w:framePr` 读写为 `ParagraphFormat.frame`（全部属性，传 `null` 清除；非法枚举与非数字按未设置处理）。它盖着 Word 里两个看起来无关的功能——**首字下沉**（`dropCap` 为 `drop` / `margin`，下沉字自成一段）和**段落定位**（DrawingML 之前的浮动做法）——但在排版上是同一件事：这一段脱离正常流，给后面的内容留出一块排除区。实现上就走同一条路：渲染时给该段加 `float`，环绕交给浏览器；分页测量用同样的 float 占位，所以两边不会各算一套。

尺寸的来源只有一条规则:文档明写了 `w:w` / `w:h` 就照它来，没写才用量出来的那一维。下沉字正是「没写」的那种——所以它的尺寸**不从 `w:lines` 反推**:Word 已经把那个字的 `w:sz` 调到正好跨 `lines` 行，量出来的字框就是排除区。宽度既量不到又没有 `w:w` 时不产生排除区（不环绕只是少个效果，尺寸编错了会把正文挤歪）。`w:wrap` 映射为排除区类型：`around` / `auto` → `square`，`tight`、`through` 原样，`notBeside` → `topAndBottom`（字面意思就是旁边不许有文字，因此也不浮动），`none` 不产生排除区也不浮动。`w:hRule` 为 `exact` 时用固定高度，否则用最小高度。排除区放不进本栏剩余高度而本栏已有内容时先换栏。

**`w:x` / `w:y` 那套按页面或页边距定位的绝对坐标本期不实现**：那需要相对页框定位，而连续视图没有页框；这些值如实读取并原样写回，浮动方向按 `xAlign` 取左右，其余照左浮。`w:anchorLock`、`w:yAlign`、`w:vAnchor` / `w:hAnchor` 同样只读取保留，不参与排版。

`w:textAlignment` 读写为 `ParagraphFormat.textAlignment`（`auto` / `baseline` / `bottom` / `center` / `top`），管的是**一行里字符的垂直对齐**（字号不一时怎么在行内对齐），和 `w:jc` 的左右对齐是两件事。渲染落在 run 的 `vertical-align` 上——段落是块级元素，`vertical-align` 设在它身上没有效果。**run 自己的上下标与 `w:position` 压过它**：三者落在同一个 CSS 属性上，而 run 级是更具体的那一层。`center` 只是近似（CSS 的 `middle` 对的是基线加半个 x-height，Word 对的是行的正中），`auto` 是 Word 的默认、不设任何东西。

另外三个段落属性**只做读写保真、不参与渲染**，理由各自不同：`w:adjustRightInd`（用文档网格时自动调整右缩进）—— 这里的 `w:docGrid` 只做了行高吸附，没有横向字符格，也就没有那个可调的量；`w:suppressOverlap`（禁止 `w:framePr` 文本框互相重叠）—— 和表格的 `w:tblOverlap` 同理，环绕用浏览器 `float`，浮动块本来就不互相重叠；`w:textboxTightWrap`（周围段落按文本框内容的实际行宽紧密绕排）—— 那需要逐行量文本框内容，而这里的排除区是矩形，与 `w:framePr` / `w:tblpPr` 一样的 `square` 简化。四个属性的非法枚举值都按未设置处理，写入时则拒绝。

`clearDirectFormat` 与剪贴板覆盖**全部**段落格式字段。这两处原先用的是一份更早建立的字段名单，少了 `kinsoku` / `wordWrap` / `overflowPunct` / `topLinePunct` / `autoSpaceDE` / `autoSpaceDN` / `bidi` / `textDirection` 八项，于是与样式冲突的东亚排版设置清不掉、剪贴板也带不走；现在名单只有一份。

Agent JSON Schema 里 `formatParagraph` / `formatRun` / `formatRange` / `formatDocumentRange` / `formatTable` / `formatTableRow` / `formatCell` 的 `format` 属性集与运行时校验**共用同一份字段名**（`operations.ts` 导出的 `PARAGRAPH_FORMAT_KEYS` / `RUN_FORMAT_FIELDS` / `TABLE_FORMAT_KEYS` / `ROW_FORMAT_KEYS` / `CELL_FORMAT_KEYS`），嵌套的 `frame` / `eastAsianLayout` / `floatingPosition` 也各有一份。此前 `frame`、`eastAsianLayout`、`floatingPosition` 和行的 `gridBefore` / `wBefore` / `gridAfter` / `wAfter` 都是校验收、schema 不声明——schema 比校验窄，等于这些字段对 agent 不存在，而两边都是绿的。

`getWebDivs()` 读取 `webSettings.xml`（经主文档关系解析）里的 `w:divs`：文档经 HTML 往返（另存为网页、邮件回复）时 Word 把 `<div>` / `<blockquote>` 的嵌套结构记在这里。树形结构摊平成数组，`parentId` 还原层级，边距单位为缇。段落与表格行的 `w:divId` 读写为 `ParagraphFormat.divId` / `RowFormat.divId`（`null` 清除）；段落所在 div 及其所有祖先的左右边距累加在段落自己的缩进之外渲染，悬空的 id 按没有 div 处理。div 的上下边距与框线（邮件引用的那条竖线）只读不渲染：它们属于一组连续段落而不是单个段落。`divId` 只在本文档内有意义，所以剪贴板与「从选区建样式」不带它。

`getCompatibilitySettings()` 读取 `settings.xml` 的 `w:compat` 声明：四个已接入的标志会暴露为明确字段，并分别影响自动段间距、东亚断行、环绕表格分页和表格条件样式规则；`compatSetting` 三元组通过 `compatSettings` 暴露，其余标志收集在 `other` 中。兼容性声明只被读取，不会放松文本、ZIP/XML、路径或其它安全校验，也不会执行文档内容。

`getThemeSettings()` 读取 `settings.xml` 的 `w:clrSchemeMapping` 和 `w:themeFontLang`，以及 `styles.xml` 的 `w:latentStyles`。颜色槽位映射参与主题色解析；替换主题关系指向的主题部件（通常名为 `theme1.xml`）即可切换当前主题。`themeFontLang` 与 `latentStyles` 只暴露为元数据，不参与字体选择或排版；主题字体仍按主题中声明的脚本槽位解析。
| `getNumberingDefinitions()` | 读取 `word/numbering.xml` 中已解析的编号定义 |
| `setParagraphNumbering(index, numId, level?)` | 为段落绑定指定编号定义与级别（默认 0） |
| `clearParagraphNumbering(index)` | 清除段落上的直接编号绑定 |
| `createNumbering(kind)` | 创建新的项目符号 / 编号 / 多级编号定义并返回新的 `numId` |
| `setParagraphLevel(index, delta)` | 提高 / 降低段落列表级别，结果钳制在 `0..8` |
| `restartNumbering(index, options?)` | 从段落处重新开始编号（默认 1；可用 `options.start` 指定起始值），并将后续连续列表项（含更深层级）切换到新编号实例；遇到不同 `numId` 或更高层级即停止；重复调用会新建实例并重新应用起始值。普通段落夹在列表中间不构成停止边界（与 Word 一致）；但**若中间插入了另一个列表的项，段落段在此处终止，其后的同列表项将继续原编号实例而非新实例——此处与 Word 不同**（仅在指定 `options.start` 时可观察到差异） |
| `continueNumbering(index)` | 将段落及后续连续列表项切回前面同抽象编号、同层级段落的编号实例；不存在前项时不做修改 |
| `formatRun(paragraph, run, format)` | 设置 run 直接格式，包括字符样式、字号、颜色、下划线、删除线、上下标与文字效果属性；将某个字段设为 `null` 可回退到继承样式 |
| `formatRange(range, format)` / `clearRangeFormat(range, fields?)` | 按段落内字符偏移格式化任意文本范围，支持清除全部或指定 run 直接格式字段 |
| `applyCharacterStyle(range, styleId, options?)` | 严格应用字符样式；可在应用时清除与字符样式冲突的直接格式 |
| `getRangeFormat(range)` | 读取字符范围内一致的 run 直接格式；同一字段在范围内不一致时返回 `undefined` |
| `copyFormat(range)` / `applyFormat(range, format)` | 格式刷 API：复制段落内范围的 run 直接格式，并应用到跨段落选区 |
| `formatDocumentRange(range, format)` / `getDocumentRangeFormat(range)` | 支持跨段落选区：首段部分 + 中间整段 + 末段部分 |
| `getEditableRegions()` / `addEditableRegion(range, options)` / `removeEditableRegion(id)` | 读取、创建和移除正文区域编辑权限标记；授权方式为编辑组或用户 ID |
| `copyClipboardFragment(range)` / `pasteClipboardFragment(range, fragment)` | 内部富文本剪贴板 API：run 直接格式、超链接、内嵌图片、段落格式与样式引用、编号（含多级）和表格；粘贴按落点切分段落，图片按目标文档关系与 media 部件重建（不复用源 `rId`） |
| `defineStyle(style)` | 创建或更新 `styles.xml` 样式定义；缺少部件时自动补内容类型与主文档关系 整体重写该样式的 `pPr` / `rPr`；进入撤销历史 |
| `updateStyle(id, patch)` | 按补丁修改已有样式：只动补丁里出现的字段，样式里读模型不认识的属性（`numPr`、主题字体槽位等）原样保留；`basedOn` / `next` / `link` / `aliases` / `uiPriority` 与格式字段传 `null` 清除。拒绝成环的 `basedOn`、类型不符的 `basedOn` / `next` / `link`、重名，以及 `paragraph.style` / `run.style`（继承请用 `basedOn`） |
| `deleteStyle(id)` | 删除样式：继承它的样式改为继承它的 `basedOn`（不折入它的格式，与 Word 一致），其他样式指向它的 `next` / `link` 删除，正文、页眉页脚、脚注尾注、批注里的 `pStyle` / `rStyle` / `tblStyle` 引用删除并回落到默认样式。默认样式不能删 |
| `createStyleFromSelection(range, style)` | 从当前选区提炼段落 / 字符直接格式并写成新的段落样式定义 |
| `getOutline()` / `setOutlineLevel(index, level)` / `moveOutlineSection(from, to)` | 读取标题大纲、显式设置 `outlineLevel`，以及按大纲整体移动标题节 |
| `replaceText(search, replacement)` | 正文及表格段落内的字面替换，支持跨 run 匹配，不跨段落 |
| `DocxDocument.compare(base, revised, options?)` | 比较两份主文档并返回新的带修订实例；纯文本/格式差异细化到段落与 run，表格及含图片的段落按粗粒度删除 + 插入处理 |
| `findText(query, options?)` | 在正文段落文字里查找，返回 `{ paragraph, start, end }`，可直接拼成 `DocumentRange`；不跨段落匹配，默认不区分大小写，命中数默认最多 1000（`maxResults` 可到 100000） |
| `getRevisions(filter?)` / `acceptRevision(id)` / `rejectRevision(id)` / `acceptAllRevisions(filter?)` / `rejectAllRevisions(filter?)` | 读取并逐条/批量接受或拒绝修订（支持按作者筛选）；`moveFrom` / `moveTo` 以 `kind: 'move'` 返回，并暴露 `{ move: { name, side, pairedId? } }` |
| `getReviewers()` | 聚合主文档修订与主文档锚点批注的审阅者统计（修订数、批注数、未解决批注数、时间范围）；每项返回 `{ kind, author? }`（`kind`: `named` / `unattributed` / `empty` / `blank`，四种都可实际构造），结果默认按计数降序，再按 `kind+author` 稳定排序 |
| `insertTable(rows)` / `insertTableAt(rows, cols, before?, format?)` | 在正文中插入表格；支持空白表格、基础表格格式和正文块级定位 |
| `getTable(index)` | 读取正文中第 N 个表格的 grid、跨度和表格/行/单元格格式信息 |
| `insertTableRow()` / `deleteTableRow()` / `insertTableColumn()` / `deleteTableColumn()` | 行列编辑；同步维护 `w:tblGrid`，拒绝删成 0 行或 0 列 |
| `mergeCells()` / `splitCell()` | 基于逻辑网格合并/拆分单元格，读写 `w:gridSpan` / `w:vMerge` |
| `formatTable()` / `formatTableRow()` / `formatCell()` | 设置表格宽度、布局、边框（含单元格对角线 `tl2br` / `tr2bl`）、底纹、单元格间距、浮动定位、行高、标题行、网格跳过、单元格对齐和边距等显式属性 |
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
| `getMath(partPath?)` / `insertMath(paragraph, source, options?)` / `setMath(paragraph, mathIndex, source)` / `deleteMath(paragraph, mathIndex)` | 读取公式，按段落 run 偏移插入、替换或删除；`source` 接受 `MathMlNode` 数据或下列线性文本子集，不接受 MathML 标记字符串 |
| `getDocumentProperties()` / `setDocumentProperties(patch)` | 读取和按字段更新 `docProps/core.xml` / `docProps/app.xml` 中的常用文档属性；缺失部件时自动补包级关系与 content-type，未知元素和未改字段原样保留 |
| `getDocumentStatistics()` / `updateDocumentStatistics(options?)` | 按 Word 的口径统计正文（含表格）：东亚文字每个字算一个词，其余按空白切分，东亚标点不算词；字符按码点计；隐藏文字与已删除修订不算，段落只数有文字的。`update…` 把 Words / Characters / CharactersWithSpaces / Paragraphs 写进 `docProps/app.xml`（没有就建），传入 `pagination` 时再写 Pages / Lines；页数与行数取决于排版，编辑器的 `editor.updateDocumentStatistics()` 会先分页测量。行数里表格行按一行算，是近似值 |
| `getDocumentProtection()` / `setDocumentProtection(value)` | 读取和写入 `settings.xml` 中的 `w:documentProtection`（如只读 / 仅批注 / 仅修订 / 表单）；仅修改 `edit` / `enforcement`，保留已有 hash/salt 等密码相关属性 |
| `getSettings()` / `setTrackChanges(enabled)` / `setRevisionAuthor(author)` | 读取常用文档设置（当前返回 `{ defaultTabStop, evenAndOddHeaders, trackChanges }`），显式开启/关闭 `w:trackChanges`（关闭时写 `w:val="0"`，不删除元素），并设置后续记录修订写入使用的作者名 |
| `revision` | 本实例的修订号；加载文件后从 0 开始，不持久化到 DOCX |

**索引与作用域**

`getParagraphs()` 和 `getBlocks()` 只包含正文段落；文本框（`w:txbxContent`）里的段落属于独立的只读文字流，通过 `getShapeParagraphs(shapeId)` 访问，不占用正文段落下标。正文按下标写入不会穿透到文本框内部。

形状 SVG 渲染支持 113 种 DrawingML 预设几何（`supportedPresetGeometries()` 返回完整清单）：基本形状、正多边形与 4~32 角星、箭头、流程图、标注（尾巴位置按 `adj1` / `adj2`）、括号与大括号、弧 / 扇形 / 弦 / 空心弧、连接线、数学符号、波形等；调整值（`avLst`）按 ECMA-376 的默认值与取值范围解释，角度类调整值以 1/60000 度为单位。线、连接线、弧、括号是开放路径，只描边不填充（`presetGeometryIsOpen()`），没写线条颜色时用黑色兜底。其它预设以及含未支持指令的自定义几何以矩形绘制，但保留已读出的填充和线条；自定义路径目前处理首个 `a:path` 中的 `moveTo`、`lnTo`、`cubicBezTo` 与 `close`。支持实体色、线性渐变、透明度 `alpha`、线宽/虚线、旋转和翻转；图案填充取前景色画成实心。未指定填充或线条的形状使用浅色填充与细描边作为可见兜底。

DrawingML 颜色读 `srgbClr`、`schemeClr`、`sysClr`（取 `lastClr`）、`prstClr`、`scrgbClr`（线性 → sRGB）、`hslClr`，并按子元素在文档里的**顺序**套用变换：`lumMod` / `lumOff` / `lum`（Word「主题颜色」面板的淡色 / 深色就是这两个）、`satMod` / `satOff` / `sat`、`hueMod` / `hueOff` / `hue`、`shade` / `tint`、`comp`、`inv`、`gray`、`red` / `green` / `blue` 及其 `Mod` / `Off`。HSL 往返有 ±1 的取整差，Office 自己的几处色板之间也差 1。`wps:style` 的 `fillRef` / `lnRef` 按 `idx` 查主题格式表（`a:fmtScheme`）：`fillRef` 的 1~999 指 `fillStyleLst`、1001 起指 `bgFillStyleLst`、0 是无填充，`lnRef` 指 `lnStyleLst`、0 是无线条；表里的占位色 `phClr` 换成引用处给的颜色，再叠上表里写的变换。`spPr` 里的 `a:ln` 与主题线条**按属性合并**（只写了宽度时颜色和线型仍来自主题），显式 `<a:noFill/>` 的线条就是没有线。主题没有格式表时退回引用处的颜色当实心填充。`effectRef`（阴影、发光等效果预设）不渲染。

图表 SVG 支持柱状 / 条形、折线、饼图、圆环、面积、散点、雷达（`standard` / `marker` / `filled`）、气泡（按面积缩放）和股价图（三条序列是「最高—最低—收盘」，四条是「开盘—最高—最低—收盘」，画最高—最低线与涨跌柱）；`bar3DChart` 等 3D 图按对应的 2D 图画（`threeD: true`），透视与深度不画，复合饼图按普通饼图画。数值只读取 `c:numCache` / `c:numLit` / `c:strCache`，不解析公式或内嵌工作簿，也不会联网读取 `c:externalData`。横向条形图将类目标签绘制在左侧、数值刻度绘制在底部。折线与面积图在缺失数据点处断开，面积填充按每一段连续数据各自向零值线闭合，缺口区间不填充。**组合图**按每条序列自己的图表组类型画（`ChartSeriesInfo.type`），有柱子时折线点与类目标签都落在类目带中间；**次坐标轴**上的序列（`axis: 'secondary'`，按图表组指向的数值轴判断）有自己的刻度、画在右侧。趋势线支持线性、指数、对数、多项式（2~6 次）、乘幂与滑动平均，拟合方式与 Excel 一致（类目图的 x 取 1..n）；误差线支持定值、百分比、标准偏差（以序列均值为中心，与 Excel 相同）、标准误差与自定义，散点 / 气泡图还画横向误差线；误差线端点计入数值轴范围。缺失缓存的序列与曲面图（等高线）退化为占位。图表填充和线条沿用 DrawingML 外观与主题色解析，未指定序列外观时按 `c:idx` 取当前文档主题的 `accent1`~`accent6`。

VML 支持 `v:shape`、`v:rect`、`v:oval`、`v:line` 的基础填充色/线条，以及内置类型 1、2、3、4、5、202、203 的常见几何映射；`v:path` 不解析，退化为矩形并保留颜色。形状 `blipFill` 只显示包内可渲染图片；外部链接图片不联网，以浅色虚线框占位。

SmartArt 使用 Word 预渲染的 `diagrams/drawing*.xml` 形状绘制；缺少该部件时显示占位。子形状文字读成 `ShapeChildInfo.paragraphs`：段落对齐（`a:pPr/@algn`）、`a:br` 段内换行，run 的加粗 / 斜体 / 下划线 / 删除线、字号、`a:latin` 字体（主题字体占位 `+mn-lt` 之类不当字体名）与实心填充颜色；渲染为 SVG 的 `<tspan>` 行，整块竖直居中。Word 预渲染时已按框大小缩好字号，这里不再折行。`text` 仍给出全部文字连成的一串。

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
- `compare()` 的块级对齐先把首尾完全相同的块直接配上，再以两边**各只出现一次**的相同块为锚点（取位置的最长递增子序列），动态规划只跑在锚点之间不同的那一段——改动散在几处的长文档也能比较，3000 段改三处约 0.6 秒。上限有两道：每份输入最多 20000 段；单段要做动态规划的不同区域最多 400 万格（约 2000 × 2000 块），超过时报错说明是哪一段太大。
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

支持的操作类型（按 `AGENT_OPERATION_SCHEMA` 顺序）：`setTrackChanges`、`setRevisionAuthor`、`acceptRevision`、`rejectRevision`、`acceptAllRevisions`、`rejectAllRevisions`、`setParagraphText`、`insertParagraph`、`deleteParagraph`、`formatParagraph`、`applyParagraphStyle`、`setParagraphNumbering`、`clearParagraphNumbering`、`setParagraphLevel`、`restartNumbering`、`continueNumbering`、`formatRun`、`formatRange`、`applyCharacterStyle`、`clearRangeFormat`、`formatDocumentRange`、`setOutlineLevel`、`moveOutlineSection`、`setParagraphTabs`、`setParagraphBorders`、`setParagraphShading`、`insertBreak`、`insertSymbol`、`replaceText`、`insertTable`、`insertTableAt`、`insertTableRow`、`deleteTableRow`、`insertTableColumn`、`deleteTableColumn`、`mergeCells`、`splitCell`、`formatTable`、`formatTableRow`、`formatCell`、`setCellText`、`insertHyperlink`、`updateHyperlink`、`removeHyperlink`、`insertBookmark`、`deleteBookmark`、`updateFields`、`insertField`、`setContentControlText`、`setContentControlChecked`、`setContentControlProperties`、`removeContentControl`、`addEditableRegion`、`removeEditableRegion`、`insertImage`、`replaceImageBytes`、`resizeImage`、`setImageAlt`、`deleteImage`、`setPartXml`、`insertFootnote`、`insertEndnote`、`setNoteText`、`deleteNote`、`convertNote`、`addComment`、`replyComment`、`setCommentResolved`、`setCommentText`、`deleteComment`、`defineStyle`、`updateStyle`、`deleteStyle`、`undo`、`redo`。

### 域（Fields）

`getFields()` 读取简单域和 `fldChar` 复杂域，并保留原始指令、缓存结果和域所在 run。`updateFields()` 重算安全的 `SEQ`、日期/时间、文档属性、`REF` 等域；提供 `pagination` 时也会写回 `PAGE`、`NUMPAGES` 和可解析到已存在书签的 `PAGEREF`，结果按所在节的页码格式化。页眉/页脚可用 `getFields(partPath)` 按部件读取，页码域也会随正文域一并写回；由于一个页眉/页脚部件由整节共享，持久化的 `PAGE` 缓存使用该节首段所在页的显示页码，分页预览仍会按实际页面单独显示。会拉取外部资源或执行宏、交互输入的域（例如 `INCLUDETEXT`、`LINK`、`MACROBUTTON`、`FILLIN`）明确不会求值，`insertField()` 也会拒绝写入这些类型。

`EQ`（公式域）会分类并把指令按**括号语法**解析成 `FieldInfo.equation` 结构树（`switch` / `options` / `raisePoints` / `values` / `characters` / `parts` / `text`，嵌套上限 16 层，超限处的文字原样保留）。EQ 用的是 `\o\ac(\s\up 10(○),甲)` 这类括号加分隔符的写法，不是「空白分词 + 反斜杠开关」那一套，所以它的 `switches` 为空数组、不提供 `argument`；分隔符按 Word 的区域设置可能是 `,` 也可能是 `;`，两者都认。EQ 是排版域而不是取值域，`evaluable` 为 `false`，`updateFields()` 不重算它——结果就是 Word 缓存的那段文字。

Word 的「合并字符」与「带圈字符」都是同一个机制：`EQ \o` 是叠印（overstrike），把括号里各部分画在同一处，各部分再用 `\s\up N` / `\s\do N` 上下错开（`\do` 在 `raisePoints` 里记为负数），所以实现的是 `\o` 本身，两个功能一并落地。渲染用 `inline-grid` 把各层放进同一个网格单元，容器宽度自然取最宽那层。**只在指令里出现、域结果里没有的那一层**（「带圈字符」的那个圈就是）标为 `contentEditable="false"` 并被 `readText()` 跳过，不会写回文档；能对上域结果的那些层仍是可编辑正文。各层拼不满整个域结果时（指令与缓存结果不一致，例如域是脏的）退回普通行内文字，不做叠印，也不会丢掉结果里的文字；域结果跨多个 run 时同样退回，避免首个 run 重复画出整个结果。其余 EQ 开关按指令结构转成 MathML 交给浏览器排版（`equationToMathMl()`，和 OMML 公式同一条路）：`\f` 分数、`\r` 根号（两个参数时**第一个是根指数**）、`\i` 积分（`\su` / `\pr` 换成求和 / 求积并把上下限放到符号上下，`\in` 放成行内上下标，`\fc` / `\vc` 自定义符号）、`\b` 括号（`\lc` / `\rc` / `\bc` 指定字符，只给一侧时另一侧为空）、`\a` 数组（`\co` 列数，`\al` / `\ac` / `\ar` 对齐）、`\l` 逗号列表、`\s\up` / `\s\do` 上下移、`\x` 方框（`\to` / `\bo` / `\le` / `\ri` 只画对应的边）、`\d\fo` / `\d\ba` 前移 / 后退；不认识的开关退化为参数并排。带数值与带字符参数的子开关分别记在 `values` / `characters` 里，转义字符（`\,`）是文字本身；一个参数里几个元素并排（`x\s\up8(2)`）读成一个组。画出来的公式是装饰（`contentEditable="false"` + `data-docx-mark`），不进 `readText()`；Word 缓存的域结果仍留在 DOM 里给 `readText()` 读到、只是不显示，所以 `flush()` 不会把它当成被删掉的文字。嵌在别的开关里的 `\o` 退化为并排。

`MERGEFIELD` 会分类并读取合并域名称，但不访问外部数据源或求值。旧式 `FORMTEXT`、`FORMCHECKBOX`、`FORMDROPDOWN` 读取 `w:ffData` 元数据，不提供填写交互；缓存结果缺失时显示默认值，复选框显示只读渲染方框。表单域的 `entryMacro` / `exitMacro` 仅作为文档声明读取，文档里的宏名声明不会被执行，也不会改变任何行为。上述用户填写或外部数据提供的域值都不会由 `updateFields()` 重算。

### 公式写入

`insertMath()` / `setMath()` 接受 `{ mathMl: MathMlNode }` 或 `{ linear: string }`。结构化 MathML 数据是主要输入；线性文本只支持以下明确子集，超出时会抛错：

| 写法 | 结果 |
| --- | --- |
| `a/b` | `m:f` 分数 |
| `a^b` | `m:sSup` |
| `a_b` | `m:sSub` |
| `a_b^c` | `m:sSubSup` |
| `√(a)` 或 `sqrt(a)` | `m:rad` |
| `(…)` | `m:d`，同时用于分组 |
| `∑_(a)^(b) c` / `∫_(a)^(b) c` | `m:nary` |

`MathMlNode` 只接受已支持的结构化标签与属性，不接受 MathML 标记字符串。前置上下标使用 `mmultiscripts` / `mprescripts` 数据映射为 OMML `m:sPre`；不支持完整 UnicodeMath。公式以不可编辑的 MathML 节点渲染，宿主可通过 `data-docx-math-index` 识别公式，并在自有界面编辑后调用 `setMath()`；不支持在公式内部直接输入或进行 WYSIWYG 编辑。演示页就是这么做的：点正文里的公式打开对话框，改线性写法、实时预览，保存走 `setMath()`，「插入 → 公式」走 `insertMath()`。

线性写法与 `getMath().linear` 互为往返：写出时凡是读回来不是单个记号的运算数都加括号（`(a+b)/(c−d)`、`x^(2n)`、`(a+b)^2`），读入时分数的分子分母与上下标外面那层括号只是分组、不画出来（UnicodeMath 的约定），作为上下标的底时括号照画；`[`、`]`、`{`、`}` 当普通分隔符。所以在对话框里不改直接保存是空操作。原先的写法在多记号的分母和上标上读不回来（`(a+b)/c−d` 读成 (a+b)/c 再减 d），读入也不认 `(a+b)/(c-d)` 这种最常见的分数写法。

MathML 的一个标签对应多个 OMML 元素（`mover` 可能来自 `m:bar` / `m:acc` / `m:groupChr` / `m:limUpp`，`mrow` 可能来自 `m:d` / `m:func` / `m:box` / `m:nary`，`mtable` 可能来自 `m:m` / `m:eqArr`），因此 `getMath()` 读出的节点会在 `MathMlNode.source` 上带出原本的 OMML 元素名，`setMath()` 写回时据此还原，把读出来原样写回当作不改动处理。手工构造的节点和线性文本没有这个字段，此时按标签推断：首尾为 `stretchy` 的 `mo` 的 `mrow` 写成 `m:d`，`munderover` 后跟底数写成 `m:nary`，其余 `mrow` 由外层容器直接承载。行内文本按 MathML 语义拆分为 `mi` / `mn` / `mo`，因此写回时 run 的切分可能比原文更细，元素结构和渲染结果不变。

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
- 图片项目符号使用包内图片渲染；外部图片只显示占位图、不联网，图片不可用时退回编号级别中的文字标记。
- OMML 读取并转换 `oMath` / `oMathPara`、分数、上下标、根号、n 元运算、括号、函数、极限、重音、横线、组合字符、矩阵、对齐数组、前置上下标（`mmultiscripts` / `mprescripts`）及盒 / 边框 / phantom 等常见元素；矩阵一行中的每个 `m:e` 各自成为一个单元格；未知元素递归保留可读文字。转换深度上限为 64 层，公式依赖浏览器原生 MathML，`getMath()` 同时提供线性文本和结构化 MathML 数据。文本框 / 形状内的公式从 `getShapeParagraphs(shapeId)` 返回的段落 `math` 读取，并在形状文字里渲染；它们不属于锚定文本框的正文段落，所以不进 `getMath()`，`setMath()` / `deleteMath()` 也按正文索引碰不到它们（渲染节点不带 `data-docx-math-index`，改标 `data-docx-shape-math`）。
- 分页预览按栏宽重新度量内容，支持等宽 / 指定宽度分栏与 `nextColumn`，并按表格行跨页 / 跨栏拆分；连续的 `w:tblHeader` 标题行会在每个片段重复，`cantSplit` 行保持完整。
- run 着重号支持 `w:em` 的 `dot`、`comma`、`circle`、`underDot`（分别使用浏览器原生 `text-emphasis`）；显式 `none` 可关闭继承的着重号。不按竖排文字方向调整着重号位置。
- 注音（`w:ruby`）读取为 `RunInfo.ruby`，含注音文字、基字符、`w:rubyAlign`、`w:hps` / `w:hpsRaise` / `w:hpsBaseText`（半磅）与 `w:lid`。**注音不计入段落正文**：`paragraph.text` 与 run 的 `text` 都只含基字符，与 Word 的阅读顺序一致；注音排在基字符上方，通过浏览器原生 `<ruby>` / `<rt>` 渲染，`<rt>` 标为 `contentEditable="false"` 且被 `readText()` 跳过，因此编辑正文不会把注音追加进文本。`w:rubyAlign` 只映射 CSS `ruby-align` 真正支持的 `center` / `distributeLetter` / `distributeSpace` / `left`，`right` 与 `rightVertical` 没有对应值，交给浏览器默认行为。注音内容本期不提供写入 API。
- 复杂文种的粗体 / 斜体（`w:bCs` / `w:iCs`）读写为 `RunFormat.boldComplexScript` / `italicComplexScript`。ECMA-376 把一个 run 的字符分成两类：`w:b` / `w:i` 管非复杂文种，`w:bCs` / `w:iCs` 管复杂文种（阿拉伯文、希伯来文、泰文……），所以 Word 里只写了 `<w:b/>` 的阿拉伯文 run **不加粗**。渲染照此：run 被标成复杂文种（`w:cs` 或 `w:rtl` 为真）就用 `bCs` / `iCs`，否则用 `b` / `i`；复杂文种 run 没写 `bCs` 时按「没设」处理，**不回退到 `b`**。**没有按 Unicode 文种逐字符判断**——含阿拉伯字符却没标 `w:cs` / `w:rtl` 的 run 仍按 `b` 画。
- `w:noProof`（不检查拼写与语法）读写为 `RunFormat.noProof`，渲染落在 run 的 `spellcheck="false"` 上，波浪线的事交给浏览器。
- `w:snapToGrid` 有两级：**段落级**（`ParagraphFormat.snapToGrid`）是这段的行是否吸附到 `w:docGrid` 的行距，默认开，显式 `false` 的段落在分页与渲染里都不吸附——两处走的是同一个 `snapLineHeightPx`，开关只判断一次；**run 级**（`RunFormat.snapToGrid`）管的是字符间距对齐到横向字符格，这里没有横向字符格（同 `adjustRightInd`），只读写保真。
- `w:specVanish`（「特殊隐藏」的段落标记，Word 用它做样式分隔符：两段并成一行、各用各的样式）读写为 `RunFormat.specVanish`，只做读写保真——渲染成一行要把两段合并排，不做。
- `w:eastAsianLayout` 读取为 `RunFormat.eastAsianLayout`（`id` / `combine` / `combineBrackets` / `vert` / `vertCompress`），可通过 `formatRun()` 写入，传 `null` 清除。「双行合一」（`combine`）CSS 没有对应能力，渲染时自行把文字按**码点**平分成上下两行（奇数时上一行多一个字），括号按 `combineBrackets` 用全角 `（）` / `［］` / `〈〉` / `｛｝` 绘制——括号是渲染装饰，`contentEditable="false"` 且被 `readText()` 跳过，不会写回文档。分页测量按较长那行的一半计宽（加上按原字号计的括号），不按整段文字计宽。「纵中横」（`vert`）映射为 CSS `text-combine-upright: all`，横排视图下不产生可见差别，与 Word 一致（Word 也只在竖排时显示）；`vertCompress` 只读取保留。双行合一的 run 内若含制表符或换行，不再走制表位计算。
- run 文字效果按映射质量分档：`w:position` 以半磅映射到 `vertical-align`；`w:outline` 使用 `-webkit-text-stroke: 1px currentColor` 并透明化文字填充，描边保留原文字颜色；`w:shadow` 使用 `1px 1px 2px rgba(0, 0, 0, 0.45)` 阴影。
- `w:emboss` / `w:imprint` 以双阴影近似：阳文使用 `-1px -1px 1px rgba(255, 255, 255, 0.9)` 与 `1px 1px 1px rgba(0, 0, 0, 0.65)`；阴文反转两道阴影的方向。该效果是浏览器 CSS 近似，不保证与 Word 像素一致。
- `w:kern` 以半磅阈值近似映射为 `font-kerning: normal`（当前字号达到阈值）或 `none`（未达到）；CSS 不支持 Word 的字号阈值语义，且未解析到字号时不额外设置字距。
- `w:w`、`w:fitText` 与 `w:effect` 读取为 `characterScale`、`fitTextWidth`、`textEffect`，也可通过 `formatRun()` 写入 / 清除，但不映射到 CSS：字符缩放与按宽度压缩会破坏行内布局 / 分页测量，动画效果已废弃且没有可靠静态映射。
- 读取节的 `w:docGrid` `type`、`linePitch` 与 `charSpace`；`lines`、`linesAndChars`、`snapToChars` 的 `linePitch` 用于分页行高吸附，`default` 不吸附，段落显式 `w:snapToGrid w:val="0"` 也不吸附。`charSpace` 目前只读取并保留，不参与字符宽度计算。
- 分页预览按节套用页面设置：每一页使用所在节的纸张宽度、页边距、方向、分栏与页眉页脚。连续视图（`viewMode: 'continuous'`，默认）是一条不分页的滚动流：单节文档照旧把第一节的纸张宽度、页边距与分栏套在纸张容器上；**多节文档**每一节放进自己的 `div.docx-section`（`data-section`、`data-orientation`），用该节的纸张宽度、左右页边距与分栏，纸张容器取最宽那一节的宽度、左右内边距归零——「纵向正文 + 一节横向宽表格」里横向那节按横向的宽度排。连续视图没有页，页眉页脚只显示第一节的；按节看页眉页脚与真实分页请切换到分页视图（`setViewMode('paginated')`）。文档中各节的页面设置本身不受影响，导出时原样保留。

**列表与表格**

- 列表计数只在主文档正文（含表格单元格）内计算，支持常见 `numFmt`，未知格式回退为十进制。
- 表格样式参与全部条件格式的格式计算：`firstRow` / `lastRow` / `firstCol` / `lastCol`、`band1Horz` / `band2Horz`、`band1Vert` / `band2Vert`，以及四个角单元格 `nwCell` / `neCell` / `swCell` / `seCell`。
- 条件格式按 ECMA-376 的优先级套用，从低到高：`wholeTable` → `band*Vert` → `band*Horz` → `firstCol` / `lastCol` → `firstRow` / `lastRow` → 四个角单元格。也就是说表头行横贯整行（含第一列），第一列的特殊格式压过带状，横向带压过纵向带，角单元格压过所有行列条件。角单元格要求对应的行、列条件都在 `tblLook` 里开着——关掉「第一列」就没有特殊的第一列，左上角也退回 `firstRow`。1×1 表格四个标志全开时四条角条件都命中，按 `nwCell` → `neCell` → `swCell` → `seCell` 的顺序最后一条胜出（Word 在这种退化情形下的行为没有明确定义，这里取确定的顺序）。
- 表格样式的单元格格式（`tcPr` 的底纹、边框、边距、垂直对齐等）会解析为 `TableCellInfo.effective`，渲染即用它。叠加顺序按级进行，每一级都是「先垫这一级 `tblPr` 给的区域默认值（`tblBorders` / `tblCellMar`），再叠这一级的 `tcPr`」：样式自身一级（`tblBorders` 就是 Word 的「Table Grid」定全框线的地方）→ 各命中条件各一级，**顺序与段落 / 文字格式用的是同一套条件和同一套优先级** → 最后是单元格自己的 `tcPr`。`TableCellInfo.format` 的语义不变，仍只含单元格自己的直接格式。
- 条件格式里的 `tblPr` 描述的是**该条件匹配到的区域**，所以它的 `tblBorders` 与 `tblCellMar` 落在那些单元格上，而不是整张表上。
- 表格元素自身的格式（宽度、对齐、缩进、布局、底纹、单元格边距、框线）解析为 `TableInfo.effective`，由样式链的 `wholeTable` 层加表格自己的 `tblPr` 决定，**条件不参与**。合并时会跳过样式里的 `tblStyle` 与 `tblLook`：前者会反过来改写表格引用的样式 id，后者会把带状 / 首行这些开关搅乱，它们只能来自表格自己。
- `w:tblLayout` 的值在 `w:type` 上（Word 写的是 `<w:tblLayout w:type="fixed"/>`），读写都按 `w:type`。
- 浮动表格（`w:tblpPr`）读写为 `TableFormat.floatingPosition`（`leftFromText` / `rightFromText` / `topFromText` / `bottomFromText`、`verticalAnchor` / `horizontalAnchor`、`xSpec` / `x`、`ySpec` / `y`；传 `null` 清除，表格回到正常流）。它是表格版的 `w:framePr`：表格脱离正常流、正文绕着它排，所以走的是同一条路——渲染给表格加 `float`，环绕交给浏览器；分页测量用同样的 float 占位，两边不会各算一套。排除区宽度优先用 `w:tblW`，没有就用网格列宽合计。
- **`w:tblpPr` 上没有 `w:wrap`**：浮动表格在 Word 里一定绕排，这正是它的用途，所以排除区类型固定为 `square`。`w:tblOverlap` 管的是能否与**其他浮动对象**重叠，与正文是否绕排无关：读写为 `TableFormat.overlap`（`never` / `overlap`，其余值按未设置处理），但**不影响排版**——环绕用的是浏览器 `float`，而浮动块本来就不互相重叠，所以 `never` 天然成立、`overlap` 无法实现。
- `<w:tblpPr/>` 即便一个属性都没有也是浮动表格（全取默认值），所以读出来是空对象而不是 `undefined`——退化成 `undefined` 会把「浮动」这件事本身丢掉。非法枚举与非数字按未设置处理，但仍然是浮动表格。
- **`w:tblpX` / `w:tblpY` 那套按页面或页边距定位的绝对坐标本期不实现**（与 `w:framePr` 同理：需要相对页框定位，而连续视图没有页框）。这些值如实读写，浮动方向按 `tblpXSpec` 取左右，其余照左浮。
- 行与单元格的读模型带两份条件格式：`conditions` 是按 `tblLook` 与位置**算出来**的（`effective` 与渲染用的就是它），`recordedConditions` 是 Word 存盘时写进 `w:cnfStyle` 的（具名属性优先，没有就读 Word 2007 的 12 位 `w:val` 位串）。`cnfStyle` 是缓存，Word 打开时会重算，所以以算出来的为准；两者不一致说明文件被别的程序改过结构或开关而没刷新缓存。`recordedConditions` 只读，不进格式对象；段落 `pPr` 里的 `cnfStyle` 不读。
- 行级表格属性例外 `w:tblPrEx` 读写为 `RowFormat.tableException`（`tblPr` 的子集：宽度、对齐、缩进、框线、底纹、单元格边距、布局、单元格间距、`tblLook`；传 `null` 删除）。它不在 `w:trPr` 里，而是 `w:tr` 的第一个子元素，写入时排在 `trPr` 之前。优先级是单元格自己的 `tcPr` > 行的 `tblPrEx` > 表格的 `tblPr` 与表格样式：框线、底纹、单元格边距按这个顺序渲染到这一行的单元格上；对齐、缩进、宽度、单元格间距是整行平移或加宽，HTML 表格的行做不到，只读写、不渲染。
- 行的网格跳过（`w:gridBefore` / `w:wBefore` / `w:gridAfter` / `w:wAfter`）读写为 `RowFormat.gridBefore` / `widthBefore` / `gridAfter` / `widthAfter`；`0`、负数与非整数按未设置处理。**跳过的列算进网格列号**——跨行合并是靠网格起始列匹配的（`vMerge` 的 continue 要对上上面那个 restart），不算进去的话带 `gridBefore` 的行里合并会断掉，单元格的 `gridStart` / `gridEnd` 也会偏小。渲染时跳过的那块用一个空单元格占位，宽度取 `wBefore` / `wAfter`，标为 `contentEditable="false"` 且不画边框。
- 单元格的两条对角线（`w:tl2br` / `w:tr2bl`）读写为 `CellFormat.borders.tl2br` / `.tr2bl`，并渲染出来。**它们只存在于 `w:tcBorders`**，`w:tblBorders` 的 schema 里没有，所以 `formatTable()` 传对角线会被拒绝，不会悄悄落到表格级边框上。渲染用角落关键字的 `linear-gradient`：CSS 规范让 `to bottom left` 的渐变线 50% 处恰好穿过左上和右下两个角（`to bottom right` 穿过右上与左下），所以一个硬色标就是一条精确的对角线，单元格不是正方形也对。用背景而不是插节点，装饰就不可能流回文档；代价是线型只能画成实线，渐变表达不了 dashed / dotted。
- `w:tblCellSpacing` 读写为 `TableFormat.cellSpacing` 与 `RowFormat.cellSpacing`。ECMA-376 把「单元格之间」和「单元格与表格边缘之间」说成同一个值，这正是 CSS `border-spacing` 的语义，所以一对一映射；有间距时会同时设 `border-collapse: separate`（`collapse` 下 `border-spacing` 被忽略）。**行级的那个落不到渲染上**：CSS 的 `border-spacing` 只能加在表格上，没有行级对应物，读写仍然保真。
- `w:tcFitText`（把单元格文字压缩 / 拉伸到正好占满格宽）读写为 `CellFormat.fitText`，**不参与渲染**——那需要按实测文本宽度反算字间距。它和 run 上的 `w:fitText`（`RunFormat.fitTextWidth`）是同一件事的两个层级，两者都只做读写保真。
- `tcPr` / `trPr` 里只写了显式关闭（如 `<w:noWrap w:val="0"/>`）的单元格和行，读出来是带 `false` 的格式对象，不是 `undefined`——把 `false` 一并算成「什么都没有」的话，调用方分不出「关掉」和「没说」，原样写回时那条显式关闭就消失了。
- 合并时 `borders` 与 `margin` 按**边**合并，`shading` 作为整体替换（`w:shd` 本就是单个元素）。按边合并是必须的：「整表定四边 + `firstRow` 只定下边框」是表格样式里最常见的组合，整块替换会把其余三边抹掉。
- 表格样式的行格式（`trPr`：行高、`cantSplit`、`tblHeader` 等）解析为 `TableRowInfo.effective`，渲染与分页都用它。**只有由行位置决定的条件参与**：`firstRow` / `lastRow` / `band1Horz` / `band2Horz`；`firstCol` 之类是单元格范围的条件，对整行没有意义，其 `trPr` 不生效。叠加顺序同样是样式自身的 `trPr` → 条件的 `trPr` → 行自己的 `trPr`（最高）；`height` 对应单个 `w:trHeight` 元素，整体替换。因此由样式的 `firstRow` 条件提供 `w:tblHeader` 的表格，跨页时也会重复表头。`TableRowInfo.format` 的语义不变，仍只含行自己的直接格式。
- `tblStyleRowBandSize` / `tblStyleColBandSize` 从**样式**的 `tblPr` 读取（Word 写在那里），表格实例上的同名属性作为直接格式优先；样式链里取最靠近的那个，都没有则为 1。首行 / 首列（以及末行 / 末列，当 `tblLook` 开了对应标志时）不参与带状计数，否则带的相位会偏一格；`useWord2002TableStyleRules` 下按 Word 的旧规则把它们一起算进去。
- `w:tblLook` 的两种写法都读：Word 2010+ 的具名属性优先，缺失时回退到 Word 2007 的 `w:val` 十六进制位掩码（`firstRow` 0x0020、`lastRow` 0x0040、`firstColumn` 0x0080、`lastColumn` 0x0100、`noHBand` 0x0200、`noVBand` 0x0400）。只认具名属性的话，只写掩码的旧文档表头行会被当成普通带状行。
- 带状按单元格在行内的序号计数，不按网格列号；含横向合并（`w:gridSpan`）的表格里带的相位可能与 Word 不一致。
- 跨越分页或分栏边界的 `rowSpan` 合并区域会连同其覆盖的行整体移到下一页或下一栏，不会拆成断开的单元格；超高的单行或合并区域仍会完整放置并允许超出可用高度。

**图片布局**

- 浮动图片使用简化的浏览器布局：四周型 / 紧密型 / 穿越型映射为浮动并缩窄相邻文字行，`topAndBottom` 映射为占满整行的块级区域，`wrapNone` 映射为绝对定位且不影响分页；外部链接图片显示占位框且不会主动联网加载。

**修订与域**

- 当前已支持读取修订（插入、删除、`rPrChange` / `pPrChange` / `tblPrChange` / `trPrChange` / `tcPrChange`）、`trackChanges` 开关，以及常见文本 / 段落 / 图片 / 表格行编辑自动写入修订；这些部件 / XML 会尽量保留。
- 节属性修订 `w:sectPrChange` 读作 `sectionFormatChange`，挂在**结束这一节的段落**上（正文末尾的 `sectPr` 对应最后一段），旧版面属性在 `previousSection`。拒绝时版面属性换回快照，页眉页脚引用保留（快照是 CT_SectPrBase，本来就不含它们）。`trackChanges` 开启时 `setPageSetup()` 自动记修订：快照始终是第一次修改前的属性，再改只刷新 id / 作者；改回原值则删掉这条修订，没有实际变化的调用不留修订。
- 接受 / 拒绝修订已支持（逐条、批量、按作者筛选，移动修订成对处理）。域值计算见上文「域」一节：`SEQ`、日期 / 时间、文档属性、`REF` 会重算，提供 `pagination` 时 `PAGE` / `NUMPAGES` / `PAGEREF` 也会写回，`TOC` 仅实现 `\o` 层级过滤与 `\h` 超链接；`INDEX` 始终保留缓存结果。低层 API 仍可直接操作。
- 编辑器标记域 run 的 `data-docx-field` / `data-docx-field-role`；结果灰底可关闭，结果不可直接编辑（不是“支持编辑域”）。

支持普通 Transitional OOXML `.docx`，不支持加密文件、`.docm` 宏文档或 Strict OOXML。导入限制：ZIP 不超过 50 MiB、最多 2048 个条目、单部件解压后不超过 16 MiB、总解压大小不超过 64 MiB。批次最多 1000 个操作，单个文本参数最多 1,000,000 字符，表格最多 10,000 个单元格。剪贴板片段最多 1000 个段落、10,000 个 run、200 张图片，单个 run 文本最多 1,000,000 字符。`setDocumentProperties()` 仅校验并写入常用 `docProps` 字段；`app.xml` 的统计值不随编辑自动更新，需要时调 `updateDocumentStatistics()`（见「文档属性」）。

**解析边界**

- XML 禁止 DTD / 自定义实体声明，ZIP 路径禁止目录穿越。

**渲染与剪贴板**

- 视图通过 DOM 文本节点和 `data:` URL 图片渲染，不将文档 XML 当作 HTML；编辑器剪贴板支持内部富文本与外部 HTML 映射，但 HTML 仅在分离文档中解析：`<script>/<style>`、事件属性、`javascript:` / `vbscript:` / `file:` / `data:` 链接都会被丢弃，`<img>` 仅接受 `data:` 形式的 PNG / JPEG / GIF / BMP（其余类型跳过该图，不影响同段其余内容），**不会主动请求外部 URL**。
- 超链接一律经 `isSafeHyperlinkUrl()` 白名单（仅 http / https / mailto），编辑器与 `pasteClipboardFragment()` 公开 API 共用同一道校验。
- 内部剪贴板携带片段引用到的样式定义（`ClipboardFragment.styles`：段落 / 字符样式及其 `basedOn` 链与 `link` 配对，不带 `isDefault`，`next` / `link` 指向集合外时去掉）。跨文档粘贴时只定义目标文档**没有**的样式，同 ID 的以目标文档为准（Word 的默认「使用目标样式」）；样式定义与粘贴是同一次提交、一步撤销。载荷是不可信输入：最多 200 个样式，每个按 `defineStyle` 操作的校验检查，不合法整批拒绝。

**使用者的责任**

- 保留原始部件**不等于清除恶意内容**；下载文件中的外部链接、嵌入对象等仍需使用者按来源谨慎处理。
- 大文档或不可信输入建议在 Web Worker / 隔离服务中处理。

测试覆盖 DOCX 往返、未修改部件保留、跨 run 替换、Unicode、样式链与主题解析、编号解析与创建、多级编号、style `numPr`、legal numbering、表格跨度解析、行列编辑、单元格合并 / 拆分、显式表格格式、分节、格式顺序、DOM 编辑、事务回滚、版本冲突、XML 校验、UTF-16、非标准主文档路径和 ZIP 解压限制。
