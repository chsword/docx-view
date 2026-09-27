# docx-view

一个面向浏览器和 AI Agent 的 TypeScript / JavaScript DOCX 编辑组件库，包含无需后端的静态 `examples`。

**当前是可运行的基础版本，不是 Microsoft Word 排版引擎，也不等同于 .NET Open XML SDK 的完整实现。** 支持段落、文字格式和基础表格的可视化编辑；对于更细粒度的操作，可以直接访问 DOCX 包中的部件、关系 XML 和命名空间感知的 OOXML DOM。

## 运行

需要 Node.js 22.12+。

```sh
npm ci
npm run dev
```

打开终端显示的地址。演示包含：新建 / 打开 / 下载 DOCX、正文与表格编辑、整段文字格式工具、OOXML 部件编辑器、Agent JSON 请求和文档快照。

```sh
npm run check          # TypeScript 检查
npm test              # 构建库并运行 Node 内置测试
npm run build         # dist/：ES modules、类型声明和 source maps
npm run build:examples # examples-dist/：可部署的纯静态站点
npm pack              # 打包库，供其他项目安装
```

将 `examples-dist/` 部署到任意静态 HTTP 服务即可；资源使用相对路径，支持子目录部署。不要直接用 `file://` 打开源码 HTML。

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
doc.formatRun(0, 0, { bold: true, fontSize: 18, color: '2455AA' });
doc.formatParagraph(0, { alignment: 'center' });
doc.insertParagraph('支持段落、文字与表格编辑。');
doc.insertTable([['任务', '状态'], ['文档编辑', '完成']]);

const bytes = await doc.toUint8Array(); // Node: 可传给 fs.writeFile
const blob = await doc.toBlob();       // 浏览器：可用于下载
const reopened = await DocxDocument.load(bytes); // 也接受 ArrayBuffer / Blob / File
console.log(reopened.getSnapshot());
```

| API | 用途 |
| --- | --- |
| `getParagraphs()` / `getBlocks()` / `getSnapshot()` | 段落 / 表格结构、文字格式、部件列表和修订号 |
| `setParagraphText(index, text)` | 修改段落文字，支持制表符和换行 |
| `insertParagraph(text, before?)` | 在指定段落前插入；省略 `before` 则追加到正文 |
| `deleteParagraph(index)` | 删除段落；保留正文 / 单元格必要的空段落，拒绝隐式删除分节符 |
| `formatParagraph(index, format)` | 对齐方式和样式 ID；不会自动创建样式定义 |
| `formatRun(paragraph, run, format)` | 粗体、斜体、下划线、字体、字号（磅）、六位十六进制颜色 |
| `replaceText(search, replacement)` | 正文及表格段落内的字面替换，支持跨 run 匹配，不跨段落 |
| `insertTable(rows)` | 在正文末尾插入表格，较短行补为空单元格 |
| `getFootnotes()` / `getEndnotes()` | 读取脚注 / 尾注内容、引用位置与计算后的显示编号 |
| `insertFootnote()` / `insertEndnote()` | 在指定段落 run 位置插入引用并创建注释内容（支持自定义标记） |
| `setNoteText()` / `deleteNote()` / `convertNote()` | 修改、删除或在脚注/尾注之间转换注释 |
| `getNoteSettings()` / `setNoteSettings()` | 读取与更新 `settings.xml` 中的脚注 / 尾注编号设置 |
| `revision` | 本实例的修订号；加载文件后从 0 开始，不持久化到 DOCX |

索引从 0 开始，包含主文档中的表格段落；结构变更后请重新读取快照。`NoteInfo.blocks` 里的段落索引按**单条注释内部**从 0 计数，不与主文档段落索引共用命名空间。高层操作只处理主文档，页眉、页脚等部件请使用底层 API。

### 可视化组件

```ts
import { DocxDocument, DocxEditor } from 'docx-view';

const container = document.getElementById('editor')!;
const doc = DocxDocument.create();
const editor = new DocxEditor(container, doc, {
  onChange: snapshot => console.log(snapshot.revision),
});

// 在外部 API 操作和导出前提交当前正在输入的内容。
editor.flush();
doc.insertParagraph('由 API 添加');
editor.render();

// 卸载组件时清理监听器。
// editor.destroy();
```

输入在段落失焦或调用 `flush()` 时提交；`onChange` 通知组件提交的修改。外部 API 修改后调用 `render()` 刷新。不要在未 `flush()` 的情况下修改同一个文档的段落结构；也应避免在输入法组合输入期间切换文档或执行外部编辑。

组件还提供 `selectedParagraph`、`setDocument(doc)` 和 `destroy()`。`docx-selectionchange` 冒泡事件的 `detail.index` 是当前段落索引。格式工具栏由宿主实现；演示工具栏作用于整段，而不是任意选中的字符范围。

组件使用 `.docx-editor`、`.docx-paragraph`、`.docx-table` 类名，不强制注入全局 CSS；宿主可以自行设置纸张外观、表格边框等，参考 `examples/style.css`。DOCX 中支持的显式 run 格式和对齐方式由组件渲染。

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
- 高层文字编辑只调整文字、制表和换行节点，尽量保留 run 属性、书签、绘图等未知内容；这并不保证复杂域、修订标记或其他高级结构的语义不受影响。
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

支持的操作类型：`setParagraphText`、`insertParagraph`、`deleteParagraph`、`formatParagraph`、`formatRun`、`replaceText`、`insertTable`、`setPartXml`、`insertFootnote`、`insertEndnote`、`setNoteText`、`deleteNote`、`convertNote`。

- 请求中的所有操作在副本上顺序执行；任一操作失败，原文档和修订号不变。
- 成功的非空批次只增加一次修订号；空批次不增加。
- 每个操作的索引相对于该操作执行前的状态，前面的插入 / 删除可能改变后续索引。
- 推荐始终提供 `expectedRevision`，避免覆盖其他编辑者的修改；冲突时重新获取快照再生成请求。
- 这是本地文档实例 API，不自带 LLM 服务、MCP 服务、远程鉴权或多人协作。若包装为服务器工具，宿主应自行实现权限、资源隔离、审核与持久化；不要把文档内容当作可信指令。

## 支持范围与安全边界

当前可视化视图支持正文段落、显式 run 格式、段落对齐和基础表格，并显示脚注/尾注引用标记；**不承诺与 Word 像素级一致或精确分页**。样式继承、编号列表、图片显示、合并单元格、复杂版式、页眉页脚、修订及域计算尚未实现；这些部件 / XML 会尽量保留，低层 API 仍可操作。`w:style` ID 的修改会保存，但视图不会解析样式继承。

支持普通 Transitional OOXML `.docx`，不支持加密文件、`.docm` 宏文档或 Strict OOXML。导入限制：ZIP 不超过 50 MiB、最多 2048 个条目、单部件解压后不超过 16 MiB、总解压大小不超过 64 MiB。批次最多 1000 个操作，单个文本参数最多 1,000,000 字符，表格最多 10,000 个单元格。

XML 禁止 DTD / 自定义实体声明，ZIP 路径禁止目录穿越。视图通过 DOM 文本节点渲染，不将文档 XML 当作 HTML；粘贴仅接受纯文本，不主动加载外部关系目标。保留原始部件**不等于清除恶意内容**；下载文件中的外部链接、嵌入对象等仍需使用者按来源谨慎处理。大文档或不可信输入建议在 Web Worker / 隔离服务中处理。

测试覆盖 DOCX 往返、未修改部件保留、跨 run 替换、Unicode、表格与分节、格式顺序、DOM 编辑、事务回滚、版本冲突、XML 校验、UTF-16、非标准主文档路径和 ZIP 解压限制。