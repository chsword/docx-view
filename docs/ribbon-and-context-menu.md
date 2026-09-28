# 功能区与右键菜单设计说明

对应 roadmap（issue #2）中演示应用的界面改造：Word 风格功能区（Ribbon）与右键上下文菜单。

> **来源与可信度**：本文由本地 GitHub Copilot CLI 分两轮生成（第二轮以第一轮产出为输入），随后由维护者核实。基线提交为 `99c901c`（471 个测试）。
>
> 已核实的关键事实：
> - `DocxEditor` 已有 `selectedImage` getter（`src/editor.ts:319`）与 `docx-imageselectionchange` 事件（`:1735`）——**图片上下文选项卡不需要新增 API**。原文写的行号 1715-1717 有小幅偏差，结论正确。
> - `examples/main.ts` 的 `selectedCell()`（约 548-560 行）只在点击表格按钮那一刻用 `document.activeElement.closest('td[data-table-cell="true"]')` 现查一次，查不到即抛错；它不是持续状态，也不经 `DocxEditor` 暴露——**表格上下文确实需要新增信号**，这一判断成立。
> - 嵌套表格在 `selectedCell()` 中直接抛错（`main.ts:555-556`），计划延续该限制。
> - `insertTableRow(table, …)` 等结构化操作经 `tableAt()` 寻址，实现是 `childrenThroughTransparent(bodyOf(document), 'tbl')[index]`，即**正文顶层表格序号**——计划要求新 API 复用同一坐标系，这一点成立。
>
> 未逐条核实的部分：完整控件映射表（43 项）的归位合理性、各菜单项清单的完备性。这些属于设计取舍，落地时按 issue 验收标准检验。

---

# docx-view Examples Ribbon 改造实施计划

## 0. 现状核对（实测，纠正题面描述）

读了 `examples/index.html`（171 行）、`examples/main.ts`（1086 行）、`src/editor.ts`（1820 行）、`src/document.ts` 相关片段。与题面描述的出入：

1. **控件数量**：逐条数下来是 **43 个**可交互控件（含新建/打开/下载/文件 input 等顶部按钮），不止「约 40 个」，具体见下方映射表。
2. **图片选中已有专用事件**，题面遗漏了：`editor.selectedImage`（getter）+ `docx-imageselectionchange` 事件（`src/editor.ts:1715-1717`），`detail: { image: ImageInfo } | null`。**图片工具上下文选项卡不需要新 API**，直接订阅这个事件即可，这是本次改造里最省事的一块。
3. **表格上下文目前完全没有实时信号**。现有 `selectedCell()`（`examples/main.ts:530-549`）是**点击表格按钮那一刻**才用 `document.activeElement.closest('td[data-table-cell="true"]')` 现查一次，查不到就抛错提示用户「请先把光标放进单元格」。它不是持续跟踪的状态，也不通过 `DocxEditor` 暴露，Ribbon 需要的「持续存在、可驱动上下文选项卡显示/隐藏」的信号并不存在，需要新增（见第 2 节）。
4. `docx-selectionchange` 的 `detail.index` 确认只有段落全局索引，**不含容器信息**（是否在表格里、在哪个表、哪一行哪一列）。
5. `getBlocks()` / `getParagraphs()` 证实：**表格单元格内段落复用正文段落索引命名空间**（与 README 规则 #10 里「页眉/页脚/脚注/文本框不复用正文索引」是两回事——表格是正文的一部分，走的是 `buildBlocksFrom(document, paragraphs)` 同一棵树），所以可以用「段落索引 → 反查是否在表格单元格里」这条路径，不需要另开一套索引空间。
6. 现有表格结构化操作（`insertTableRow` 等）在 `src/document.ts` 里全部以 `table: number` 寻址，其含义是「`getBlocks()` 顶层 `type==='table'` 块里的第几个」（`examples/main.ts` 里 `doc.getBlocks().filter(b=>b.type==='table').length-1` 这种写法印证了这一点）。新 API 必须复用同一套编号，不能另造一套坐标系。
7. `selectedCell()` 现状**只支持正文顶层表格**，遇到嵌套表格直接抛错（`examples/main.ts:537-539`）。Ribbon 的表格上下文选项卡设计要延续这个限制，不在本次范围内解决嵌套表格。
8. 开发者面板的三个 tab（Agent/OOXML/快照）**不属于 Ribbon**，它们是侧边栏的独立 tablist，不在本次改造范围，仅在下文「选项卡映射」里注明不迁移。

---

## 1. 固定选项卡与完整控件映射表

### 1.1 选项卡设置：五个，不设「引用」

按 Word 习惯设「开始 / 插入 / 布局 / 审阅 / 视图」五个固定选项卡。**不设「引用」**，理由：

- 「引用」在 Word 里承载目录、脚注尾注、题注、索引、引文。docx-view 目前**只有脚注/尾注**（`src/notes.ts`），且 `examples/index.html` 里**没有任何脚注相关控件**——插入脚注、脚注编号、脚注样式在演示应用里都不存在 UI 入口。为一个功能设一个空选项卡没有意义，也违反「不要臆造仓库里不存在的功能」的要求。
- 若未来补上脚注 UI，届时再从「插入」选项卡里的「脚注/尾注」分组拆出「引用」选项卡即可，属于增量演进，不需要现在预留空壳。

### 1.2 完整映射表

| 现有控件 id | 现有分组位置 | 目标选项卡 | 目标分组 | 备注 |
|---|---|---|---|---|
| `new-document` | 顶栏 | — | 顶栏保留 | 不进 Ribbon，Word 里这类是「文件」菜单，题面未要求做文件菜单，维持顶栏 |
| `open-document` | 顶栏 | — | 顶栏保留 | 同上 |
| `download-document` | 顶栏 | — | 顶栏保留 | 同上 |
| `file-input` / `image-file-input` | 隐藏 input | — | 隐藏 input 原地保留 | 由「插入」选项卡的图片按钮触发 |
| `format-bold` | 工具栏 | 开始 | 字体 | |
| `format-italic` | 工具栏 | 开始 | 字体 | |
| `format-underline` | 工具栏 | 开始 | 字体 | |
| `font-size` | 工具栏 | 开始 | 字体 | |
| `font-color` | 工具栏 | 开始 | 字体 | |
| `format-painter` | 工具栏 | 开始 | 剪贴板 | Word 里格式刷在「剪贴板」组最左侧，双击锁定行为保留 |
| `clear-format` | 工具栏 | 开始 | 字体 | Word 「清除所有格式」在字体组 |
| `paragraph-style` | 工具栏 | 开始 | 样式 | |
| `alignment` | 工具栏 | 开始 | 段落 | |
| `list-bullet` | 工具栏 | 开始 | 段落 | |
| `list-decimal` | 工具栏 | 开始 | 段落 | |
| `list-outdent` | 工具栏 | 开始 | 段落 | |
| `list-indent` | 工具栏 | 开始 | 段落 | |
| `add-paragraph` | 工具栏 | 开始 | 段落 | 「＋段落」本质是插入操作，但它输出的是当前样式的正文段落，Word 里同类动作（Enter 换段）不设按钮；仍按现状放「开始·段落」，不改变行为 |
| `table-rows` / `table-cols` / `add-table` | 工具栏 | 插入 | 表格 | |
| `insert-row` / `delete-row` / `insert-col` / `delete-col` / `merge-cells` / `split-cell` / `cell-fill` / `apply-cell-style` | 工具栏 | **表格工具（上下文）** | 布局 | 见第 2 节，仅在光标位于表格内时出现 |
| `insert-image` / `replace-image` / `delete-image` | 工具栏 | 插入 / **图片工具（上下文）** | 插入·插图；图片工具·调整 | 「插入」放的是 `insert-image`（永久可见）；`replace-image`/`delete-image` 只在选中图片时才有意义，迁到图片工具上下文选项卡 |
| `image-alt` | 工具栏 | **图片工具（上下文）** | 辅助功能 | |
| `page-size` / `page-orientation` / `apply-page-setup` | 工具栏 | 布局 | 页面设置 | |
| `insert-section-break` | 工具栏 | 布局 | 页面设置 | Word 把分节符放在「布局→分隔符」，不放「插入」 |
| `insert-page-break` | 工具栏 | 插入 | 分页 | 分页符是内容层面的插入，Word 原生就放在「插入→分页」，与分节符分开 |
| `header-kind` / `footer-kind` | 工具栏 | 插入 | 页眉和页脚 | |
| `review-revision-view` | 侧栏审阅面板 | 审阅 | 修订 | |
| `review-show-revisions` / `review-show-comments` | 侧栏审阅面板 | 审阅 | 修订 | |
| `review-select-all` / `review-clear-authors` / `reviewer-list` | 侧栏审阅面板 | 审阅 | 批注 | 作者筛选服务于批注可见性 |
| `toggle-comments` | 侧栏批注面板 | 审阅 | 批注 | |
| `comment-author-filter` / `comment-resolved-filter` | 侧栏批注面板 | 审阅 | 批注 | |
| `new-comment` / `reply-comment` / `resolve-comment` / `delete-comment` | 侧栏批注面板 | 审阅 | 批注 | |
| `comment-list` | 侧栏 | — | 侧栏保留 | 列表本体不是命令控件，留在侧栏，Ribbon 只放操作按钮 |
| `outline-list` | 侧栏 | 视图 | 显示 | Word 「视图→导航窗格→标题」对应大纲；本次只把「显示/隐藏大纲面板」的开关放视图选项卡，大纲列表本体仍在侧栏 |
| `effective-format` / `snapshot-output` / `xml-*` / `agent-*` | 侧栏开发者面板 | — | 侧栏保留，不迁移 | 这是开发者工具 tablist，不属于 Ribbon 范围（见 0.8） |

**统计核对**：以上覆盖 `index.html` 里全部 `id` 属性对应的可交互元素（不含纯展示的 `<span>`/`<ul>`/`<pre>` 容器，也不含 `label`/`legend` 这类无独立交互语义的元素），无遗漏、无臆造。

### 1.3 视图选项卡的补充

「视图」目前没有现成控件可搬，题面也没提「视图」该放什么。为了不臆造功能，「视图」选项卡只放两类**确有实现依据**的开关：

- 「显示大纲」——控制 `.outline-panel` 的可见性（纯 CSS 显隐，不影响 `editor`/`doc` 状态）；
- 「显示批注高亮」——本质是 `toggle-comments` 的重复入口。**不重复放**，保持它只在「审阅」出现，视图选项卡不放。

即视图选项卡本次只做「显示大纲」一个开关，其余留空，不硬凑控件。

---

## 2. 上下文选项卡设计（重点）

### 2.1 图片工具：直接复用现有信号，无需新 API

- 判定信号：`editor.selectedImage`（非空即显示）+ 订阅 `docx-imageselectionchange`。
- 失焦隐藏：`editor.selectedImage === null` 时事件 `detail === null`，隐藏。
- 这块**不需要改 `src/`**，纯 Ribbon 容器层的事件订阅。

### 2.2 表格工具：需要新增最小 API

**结论：现有 `docx-selectionchange` 不够用**，原因见 0.6/0.8——它只给一个全局段落索引，不带「是否在表格里、是哪张表哪个格」的结构信息，而 Ribbon 需要的是**持续的、随光标移动更新的**状态，不能像 `selectedCell()` 那样只在点击按钮那一刻现查 DOM。

#### 2.2.1 新增核心 API：`DocxDocument.getTableCellAt(paragraphIndex)`

```ts
interface TableCellLocation {
  table: number;      // 与 getBlocks() 顶层 table 块下标同一坐标系，和 insertTableRow(table, …) 的 table 参数一致
  row: number;         // 对应 TableCellLocation 所在行的 rowStart（与 data-row-start 语义一致）
  col: number;         // 对应 gridStart
  rowSpan: number;
  colSpan: number;
  nested: boolean;     // true 表示该段落实际位于嵌套表格内部，此时 table/row/col 描述的是最外层可寻址表格坐标
}

getTableCellAt(paragraphIndex: number): TableCellLocation | null
```

**为什么切在这里、切成这个形状**：

- **复用 `getBlocks()` 已有的遍历**，不新起一套解析：内部实现是对 `buildBlocksFrom` 产出的块树做一次「段落索引 → 祖先链」的反向查找（记录每个 table 块内每个 cell 的段落索引区间），不重新解析 XML，符合「核心库改动尽可能小」。
- **返回 `null` 而不是抛错**：符合 README 通用底线「畸形/边界情况降级，不抛错」——段落不在任何表格里是最常见的调用场景（每次选区变化都可能调用），不能用异常做控制流。
- **`table` 坐标系与既有寻址方式保持一致**：`insertTableRow`/`mergeCells`/`formatCell` 都吃 `table: number` 且约定为顶层表格序号，新 API 复用同一约定，Ribbon 拿到 `location.table` 可以直接传给这些既有方法，不需要额外换算。
- **`nested` 标志**：延续 `selectedCell()` 现状「不支持嵌套表格结构化操作」的限制——`nested: true` 时表格工具上下文选项卡仍然出现（光标确实在表格里，视觉上要给反馈），但行/列/合并/拆分类按钮禁用并给出提示（复用现有报错文案「不支持嵌套表格」），底纹类可以做在最内层 cell 上按普通格式化处理（不涉及结构，风险低）——这一点在实现前需要和表格工具选项卡的按钮禁用逻辑对齐，具体是否支持内层单元格底纹留给对应 issue 里做决定并写验收标准，不在本计划里臆断。

#### 2.2.2 `DocxEditor` 侧新增：`selectedTableCell` getter + `docx-tablecellchange` 事件

```ts
get selectedTableCell(): TableCellLocation | null;
// 事件：'docx-tablecellchange'，detail: { cell: TableCellLocation } | null
```

- 触发时机：跟 `selectParagraph()` 同一处调用点，段落切换后立刻用新段落索引调 `this.document.getTableCellAt(index)`。
- **去抖规则**：只有当「是否为 null」或者 `table`/`row`/`col` 三元组变化时才 `dispatchEvent`——光标在同一个单元格内左右移动、只是段落内文字选区变化，**不重复触发**，避免上下文选项卡在同一格内移动时反复闪烁（这也是不额外推进 revision、不做无谓重渲染的一部分，见第 3 节）。
- 这条事件和 `docx-selectionchange` 是同一次内部计算的两个输出，不新增独立的 DOM 遍历，性能增量可控（一次 `getBlocks()` 结果查表，O(1) 摊还，前提是 2.2.1 的实现按段落索引建索引而不是每次线性扫描——这一点写进对应 issue 的验收标准）。

### 2.3 上下文选项卡出现/消失时的激活规则

采用 Word 的行为模型，明确定义如下状态机：

- 维护两个状态：`activeTabId`（当前显示内容的选项卡）与 `lastManualTabId`（用户上一次**手动点击**固定选项卡时记录的 id，初始为 `'home'`）。
- **上下文选项卡出现时**（`selectedTableCell`/`selectedImage` 由 null 变为非 null）：
  - 若当前 `activeTabId` 不是「刚才已经手动切换到的某个固定选项卡且用户在该上下文出现前 3 秒内点过」——按 Word 实际行为简化为：**只要上下文选项卡从无到有，立刻自动切到它**，并把之前的 `activeTabId` 存入 `preContextTabId`（不覆盖 `lastManualTabId`）。
  - 若表格工具和图片工具**同时**满足（理论上不会同时满足，因为选区互斥：选中图片时段落级 `selectedTableCell` 仍可能非 null——图片本身在表格单元格里的段落中。这种情况下优先显示**离用户当前操作意图更近**的一个：以「最后触发的事件」为准，即哪个事件后到达就激活哪个上下文选项卡，另一个仍保留在选项卡栏但不自动激活）。
- **上下文选项卡消失时**（由非 null 变为 null）：
  - 若 `activeTabId` 正是这个消失的上下文选项卡，回到 `preContextTabId`（如果 `preContextTabId` 也已经不存在于当前选项卡集合里——例如另一个上下文选项卡也刚好消失——则回退到 `lastManualTabId`，最终兜底 `'home'`）。
  - 若用户在上下文选项卡显示期间**手动点过其他固定选项卡**（更新了 `lastManualTabId` 且当前 `activeTabId !== 上下文选项卡`），则上下文选项卡消失时**不做任何激活跳转**，只是从选项卡栏移除按钮，因为用户已经主动移开了注意力，这是「手动切换」和「自动切换」共存的关键规则：**自动切换只发生在用户当前正停留在即将消失的那个上下文选项卡上时**。
- 用户手动点击任意固定选项卡：更新 `lastManualTabId = activeTabId = 该选项卡`；不影响 `preContextTabId` 的记录（它只在下一次上下文选项卡出现时被覆写）。

这套规则用一张状态表 + 4 条转移规则就能完整描述，实现时建议就是一个独立的小型状态对象（不依赖 DOM 查询判断「用户是否手动点过」），便于写单元测试。

---

## 3. 状态同步：Ribbon 绝不驱动文档状态

### 3.1 保证手段

- Ribbon 容器（选项卡切换、上下文选项卡显示/隐藏、分组折叠等）**只允许读取**：`doc.getSnapshot().revision`（仅用于展示，不用于比较驱动渲染）、`editor.selectedParagraph`、`editor.selectedRange`、`editor.selectedImage`、新增的 `editor.selectedTableCell`，以及订阅上述只读事件。
- Ribbon 状态**只写自己的内部状态**（`activeTabId`/`lastManualTabId`/`preContextTabId`）和纯 CSS class（`.tab-panel[hidden]`、按钮 `disabled`），**不得调用** `doc.*`（除非是用户点击某个功能按钮触发的真实编辑操作，那是既有的按钮 click handler 职责，不属于「选项卡切换」这个动作本身）。
- 硬性规则写进代码评审清单：**任何 Ribbon 选项卡/分组的显示逻辑代码里，禁止出现 `doc.` 开头的方法调用**（只允许 `doc.getSnapshot()`/`doc.get*` 只读查询，不允许 `doc.set*`/`doc.insert*`/`doc.delete*`/`doc.format*`/`doc.apply*`），可以用简单的 ESLint 规则或 code review checklist 固化，不引入新依赖。

### 3.2 需要的回归测试

在 `examples/`（如果有测试）或 `test/` 对应位置补：

1. **纯选项卡切换不改文档**：记录切换前 `doc.getSnapshot().revision`，连续点击 5 个固定选项卡 + 触发一次表格上下文选项卡出现/消失，断言 `revision` 不变、`onChange` 回调未被调用（用一个计数 spy 包一层 `onChange`）。
2. **进入/离开只读预览模式不改文档**：`revisionView` 在 `final`/`original`/`markup` 间切换，同上断言 revision 不变（这本身也是既有行为，但要加断言防止 Ribbon 改造引入回归）。
3. **上下文选项卡的显示判定是纯函数**：把「光标位于表格单元格内 → 表格工具选项卡出现」「移出 → 消失」写成对 `getTableCellAt`/`selectedTableCell`/事件序列的断言，不依赖真实渲染像素。
4. **激活状态机的 4 条转移规则**：针对 2.3 节的状态机单独写单测（自动切入、自动切回、手动切换后不跳转、两个上下文选项卡交替出现时的优先级），用状态对象直接测试，不需要真实 DOM。
5. **装饰元素不流回文档**（对应工程约定第 4 条）：Ribbon 本身不产出任何进入 `editor` 内容区的 DOM 节点，这条测试其实是「零增量」——即确认 Ribbon 改造前后 `readText()`/`getParagraphs()[i].text` 对同一份文档不变，作为防止未来有人把 Ribbon 里的图标/提示文字误塞进编辑区的兜底测试。

---

## 4. 只读预览模式（`revisionView: 'final' | 'original'`）下的禁用规则

规则总纲：**只读预览模式下，一切会调用 `doc.*` 写方法的控件禁用；纯预览/筛选/导航类控件保持可用。**

具体到选项卡：

| 选项卡 | 处理方式 |
|---|---|
| 开始（字体/剪贴板/样式/段落） | 整组禁用（`disabled`），因为所有按钮最终都调用 `doc.formatRun`/`doc.formatParagraph`/`doc.setParagraphNumbering` 等写方法。格式刷、清除格式同理禁用。 |
| 插入（表格/图片/分页/页眉页脚） | 整组禁用，理由同上——全部是写操作。 |
| 布局（页面设置/分节符） | 整组禁用。 |
| 表格工具 / 图片工具（上下文） | 整组禁用；但**上下文选项卡本身的出现/消失判定不受影响**（只读模式下光标仍可能落在表格里，选项卡该出现还是出现，只是里面的按钮全灰，这与 Word 行为一致：只读时功能区仍显示上下文但按钮禁用，不是把选项卡整个藏起来，用户仍需要知道自己在表格里）。 |
| 审阅 | **不禁用**：`revisionView` 下拉本身、`showRevisions`/`showComments` 复选框、作者筛选、批注列表浏览都必须可用（否则用户没法退出只读模式）。`new-comment`/`reply-comment`/`resolve-comment`/`delete-comment` 这几个会写批注内容的按钮禁用（批注写入本质也是改文档）。 |
| 视图 | 不禁用（显示/隐藏大纲是纯 UI 状态）。 |

判定信号：不新增状态，直接读现有的 `reviewRevisionView !== 'markup'`（`examples/main.ts` 里已有的模块级变量）驱动 Ribbon 的 `disabled` 绑定，这是既有的 `isMarkupView()`/`ensureMarkupView()` 逻辑的界面层镜像，不需要改 `src/`。

---

## 5. Issue 拆分（2–4 个，可独立提交）

### Issue A：Ribbon 容器骨架 + 五个固定选项卡迁移（无上下文选项卡）
- **依赖**：无，最先做。
- **范围**：新增 Ribbon 容器组件（原生 DOM + TS，放 `examples/ribbon.ts` 或类似，不进 `src/`），实现 `role="tablist"`/`role="tab"`/`aria-selected`/箭头键导航；把 1.2 映射表里**除表格工具、图片工具外**的全部控件按分组迁入五个选项卡；`index.html` 结构调整，`main.ts` 里控件的 `element(id)` 查找与事件绑定**不改**（只是外层容器变了，id 保留，避免大改 `main.ts` 逻辑）。
- **验收标准**：
  1. `npm run check && npm test` 通过；
  2. 现有全部 43 个控件都能在新 Ribbon 里找到且行为不变（手工回归清单：加粗/字号/表格结构化操作/图片插入替换删除/页面设置/分节分页/审阅筛选/批注增删复现无误）；
  3. Tab 键、方向键可在选项卡间导航，`aria-selected` 正确切换；
  4. 选项卡切换不触发 `doc.getSnapshot().revision` 变化（对应 3.2 的测试 1）。
- **不在范围**：上下文选项卡、只读模式禁用规则、`src/` 改动。

### Issue B：核心库最小 API —— `getTableCellAt` + `selectedTableCell`/`docx-tablecellchange`
- **依赖**：无（可与 A 并行开发，但建议排在 A 之后合并以减少冲突面，因为 A 会移动很多 DOM，B 只改 `src/`）。
- **范围**：`src/document.ts` 新增 `getTableCellAt(paragraphIndex)`（2.2.1）；`src/editor.ts` 新增 `selectedTableCell` getter 与 `docx-tablecellchange` 事件（2.2.2），复用现有段落切换回调，不新增独立 DOM 遍历。
- **验收标准**：
  1. 单测覆盖：不在表格内返回 `null`；顶层表格单个/合并单元格返回正确坐标；嵌套表格返回 `nested: true` 且外层坐标正确；段落索引越界/文档无表格时不抛错；
  2. 事件去抖：同一单元格内光标移动不重复 dispatch，跨单元格/跨表移动才 dispatch；
  3. 不破坏现有 `docx-selectionchange`/`docx-rangechange`/`docx-imageselectionchange` 行为（回归测试跑通）；
  4. 性能：对一个含多表格的文档，连续 100 次段落切换的 `getTableCellAt` 调用总耗时有基准断言（避免线性扫描全文档退化，呼应 README 规则 #8）。
- **不在范围**：Ribbon UI、`examples/main.ts` 里 `selectedCell()` 的替换（留给 Issue C 做，避免这个 issue 里 API 设计和调用方改造混在一起 review）。

### Issue C：表格工具 / 图片工具上下文选项卡 + 激活状态机
- **依赖**：A（需要 Ribbon 容器）+ B（需要 `selectedTableCell`/事件）。
- **范围**：
  1. 用 Issue B 的事件驱动表格工具选项卡显示/隐藏，替换掉 `examples/main.ts` 里 `selectedCell()` 现查 DOM 的方式（改为按钮 click 时读 `editor.selectedTableCell`，找不到就仍然给出既有报错文案，保证行为不回退）；
  2. 用既有 `docx-imageselectionchange` 驱动图片工具选项卡；
  3. 实现 2.3 节的四条状态机转移规则；
  4. `image-alt` 从「开始工具栏」迁到图片工具选项卡（映射表已定）。
- **验收标准**：
  1. 光标进出表格单元格、选中/取消选中图片时，对应上下文选项卡出现/消失且行为符合 2.3 的状态机（补齐 3.2 测试 3、4）；
  2. 嵌套表格场景下选项卡出现但结构化按钮禁用，提示文案与现状一致；
  3. 全程不触发 `revision` 变化（3.2 测试 1 的扩展用例）；
  4. `npm run check && npm test` 通过。
- **不在范围**：嵌套表格的完整结构化操作支持（维持现状限制）；批注/大纲相关的上下文行为（本次只做表格与图片两类）。

### Issue D：只读预览模式禁用规则 + 文档更新
- **依赖**：A、C（禁用规则要作用在已经就位的 Ribbon 选项卡和上下文选项卡上）。
- **范围**：按第 4 节的表格，给各选项卡/分组接上 `reviewRevisionView !== 'markup'` 的 `disabled` 绑定；补 README/示例说明只读预览模式下的界面行为；补 3.2 测试 2。
- **验收标准**：
  1. `final`/`original` 模式下，开始/插入/布局/表格工具/图片工具整组禁用，审阅的浏览类控件保持可用，写入类批注按钮禁用；
  2. 切换 `revisionView` 三个值之间往返，`disabled` 状态正确恢复，且不触发 `revision` 变化；
  3. `npm run check && npm test` 通过；
  4. README 或 `examples/` 内说明文档补充「只读预览模式下 Ribbon 行为」一节。
- **不在范围**：新增只读模式的其它交互（如只读态下的批注筛选高亮变化），维持现状。

**顺序**：A → B（可与 A 并行，合并顺序 A 先）→ C（依赖 A+B）→ D（依赖 A+C）。每个 issue 单独 PR，`npm run check`/`npm test` 全绿再进下一个，避免像 README 规则 #15 那样在后续 issue 里发现前一个 issue 遗留的读路径不一致。


---

# docx-view 右键上下文菜单实施计划

已完整读取 `ribbon-plan.md`（五个固定选项卡 + 表格/图片上下文选项卡 + Issue A–D）与 `.github/copilot-instructions.md`（17 条约定）。本计划建立在其之上，并对其一个明确缺口做了修正说明（见 §0）。

---

## 0. 与上一轮方案的衔接：一个必须先补的缺口

Ribbon 计划 Issue A 明确写的是「`main.ts` 里控件的 `element(id)` 查找与事件绑定**不改**」——也就是说，上一轮方案里 Ribbon 按钮的「是否可用 / 点了做什么」仍然分散写在 43 个各自独立的 `click` handler 里，**并没有抽出统一注册表**。

如果现在直接在这个基础上再叠一份右键菜单，两边各自判断「加粗按钮什么时候可用」「合并单元格点了调用哪个 API」，就是工程约定第 15 条点名的「平行读路径」的命令版本：功能区一套 `enabled` 判断，菜单再写一套，日后任何一次修改都可能只改中一边。

**因此本计划在 Issue 拆分里insert 一个新的、优先级最高的 Issue（记为 Issue R，见 §6），先把现有 43 个控件的「判断 + 执行」从分散的 click handler 收敛成一张命令注册表，Ribbon 按钮和右键菜单都只是这张表的两种呈现层**。这是一次纯重构（行为不变，用 Ribbon 计划 §3.2 的回归测试兜底），不新增功能，右键菜单相关的新增能力全部在此之后落地。

---

## 1. 命令注册表设计

### 1.1 命令描述符形状

```ts
interface CommandContext {
  // 只读快照，禁止携带可写句柄；见 §5 的"零回写"约束
  revisionView: 'markup' | 'final' | 'original';
  editable: boolean;                    // revisionView === 'markup' 的镜像，命令层不必知道具体原因
  selection: {
    paragraph: number | null;
    range: DocumentRange | null;        // start/end 均含 paragraph+offset
    format: RunFormat | null;           // selectedRangeFormat，用于三态勾选
    collapsed: boolean;                 // range.start === range.end
  };
  table: TableCellLocation | null;      // 复用 Ribbon 计划新增的 selectedTableCell
  image: ImageInfo | null;              // 复用既有 selectedImage
  hyperlink: { paragraph: number; runs: number[]; url?: string; anchor?: string; unsafe: boolean } | null;
  revisionsAtPoint: RevisionMark[];     // 光标/右键命中的 run 上挂着的修订标记，可能为空数组
  commentsAtPoint: number[];            // 光标/右键命中处的批注 id 列表
  clipboard: 'unknown' | 'empty' | 'has-content'; // 见 1.4，默认 'unknown'
  source: 'ribbon' | 'context-menu' | 'keyboard'; // 供命令内部做统计/埋点用，不影响 enabled/checked 判断
}

interface CommandDescriptor {
  id: string;                                   // 'format.bold' / 'table.mergeCells' / 'clipboard.cut' ...
  title: string;                                 // 展示文案，中文
  group: string;                                 // 供 Ribbon 分组复用（对齐 Ribbon 计划 1.2 的分组列）
  icon?: string;                                  // 可选，纯展示
  shortcut?: string;                              // 展示用，如 'Ctrl+B'；不负责绑定，绑定仍由既有 keydown 处理器做
  enabled(ctx: CommandContext): boolean;
  checked?(ctx: CommandContext): boolean;         // 三态命令才提供：加粗/项目符号/显示格式标记...
  visibleInMenu?(ctx: CommandContext) => boolean; // 见 1.3，默认等于 enabled 之外的"是否该出现"判断
  run(ctx: CommandContext): void | Promise<void>; // 执行体，内部才调用 doc.*/editor.*
}
```

关键设计取舍：

- **`enabled` 决定"能不能点"，`visibleInMenu` 决定"这个位置该不该出现"**。二者分开是因为 Ribbon 和右键菜单对"不适用"的呈现方式不同：Ribbon 永远显示按钮只是置灰（Word 习惯），右键菜单里"合并单元格"在正文里点根本不该出现一行，而不是出现一行灰色的。两边共用同一个 `enabled`，各自决定要不要读 `visibleInMenu`。
- **`run(ctx)` 内部只允许调用既有的 `doc.*`/`editor.*` 公开方法**，不允许直接操作 DOM——这保证命令注册表本身不产生新的写入路径，校验仍然全部在 `src/document.ts` 的公开 API 里（呼应约定第 16 条：校验属于文档层）。
- 注册表本身放在 `examples/commands.ts`（不进 `src/`），因为它是"编排既有 API 调用"的胶水层，不涉及 OOXML 读写；这与约定"核心逻辑不依赖浏览器全局 document、核心库改动尽可能小"一致。

### 1.2 `CommandContext` 如何从既有信号得出

一个模块级的 `buildCommandContext()` 函数，订阅：

| 字段 | 来源事件/getter | 备注 |
|---|---|---|
| `selection.*` | `docx-selectionchange` + `docx-rangechange` + `editor.selectedRange`/`selectedRangeFormat` | 已有信号，直接读 |
| `table` | Ribbon 计划 Issue B 新增的 `docx-tablecellchange` + `editor.selectedTableCell` | 复用，不再造第二套 |
| `image` | `docx-imageselectionchange` + `editor.selectedImage` | 已有信号 |
| `hyperlink` | **右键命中点**的 DOM 祖先链上 `closest('[data-docx-link="1"]')`，读其 `dataset.docxUrl`/`dataset.docxAnchor`/`dataset.docxUnsafe`，配合最近 `[data-paragraph]` 祖先的 `dataset.paragraph` | 见 §2 的位置解析；**不是**持续状态，是右键那一刻现算的一次性上下文（下同两项） |
| `revisionsAtPoint` | 右键命中点 `closest('[data-docx-revision-ids]')`（**新增**，见 §1.5）解析出的 id 列表，反查 `doc.getRevisions()` | |
| `commentsAtPoint` | 右键命中点 `closest('[data-docx-comment-ids]')`（已有属性）解析 | 与批注面板点击复用同一属性，不新造 |
| `revisionView`/`editable` | `reviewFilter.revisionView` | 与只读模式判断复用同一个字段 |
| `clipboard` | 见 1.4 | |

`table`/`image` 走"持续跟踪状态"（Ribbon 上下文选项卡需要），`hyperlink`/`revisionsAtPoint`/`commentsAtPoint` 走"右键那一刻现查 DOM"——因为它们只在右键菜单场景下才需要知道"点在了哪个具体标记上"，Ribbon 不需要这三项，没必要为它们建持续事件（对应约定第 8 条"不要为不需要的场景常驻计算"的精神）。

### 1.3 `visibleInMenu` 与 `enabled` 的默认关系

默认 `visibleInMenu = ctx => true`（始终可能出现，靠 `enabled` 置灰），仅当一个命令**只在特定结构位置才有意义**时才显式覆盖：

- `table.mergeCells`：`visibleInMenu: ctx => ctx.table !== null`（不在表格里，整行都不出现，而不是出现且置灰——Word 的表格命令组在正文里同样是整组消失）；
- `image.replace`/`image.delete`/`image.setAlt`：`visibleInMenu: ctx => ctx.image !== null`；
- `hyperlink.edit`/`hyperlink.remove`/`hyperlink.copyAddress`：`visibleInMenu: ctx => ctx.hyperlink !== null`；
- `revision.acceptAtPoint`/`revision.rejectAtPoint`：`visibleInMenu: ctx => ctx.revisionsAtPoint.length > 0`；
- `comment.replyAtPoint`/`comment.resolveAtPoint`：`visibleInMenu: ctx => ctx.commentsAtPoint.length > 0`；
- 其余（字体格式、段落格式、剪切/复制/粘贴、插入行列等）不覆盖，走 Ribbon 那种"常驻+置灰"语义，右键菜单里也常驻显示（这与 Word 一致：右键菜单里"加粗"永远在，哪怕当前没选中文字）。

### 1.4 剪贴板可用性的现实约束

浏览器不允许在 `contextmenu` 事件的同步阶段读取系统剪贴板内容（`navigator.clipboard.readText()` 是异步且可能触发权限提示）。因此：

- `clipboard` 字段默认值恒为 `'unknown'`，`paste` 命令的 `enabled` 对 `'unknown'` 和 `'has-content'` 都返回 `true`，只在显式为 `'empty'` 时禁用——即**默认乐观可用**，这是 Word Web / Google Docs 采用的同一策略。
- 若浏览器授予了 `clipboard-read` 权限（`navigator.permissions.query` 已是 `'granted'`），可以在菜单弹出后**异步**补一次 `navigator.clipboard.readText()`，结果非空/空时再刷新那一行的 `disabled`（菜单已经在屏幕上，只做属性更新，不重新定位）。这是可选增强，不阻塞菜单打开。
- `paste` 命令的 `run()` 不使用 `navigator.clipboard` 写回文档，而是聚焦编辑区后调用 `document.execCommand('paste')`（触发既有 `editor.ts:803` 的 `paste` 事件监听器，走 `handleClipboardPaste` → `pasteClipboardFragment` 既有校验路径）。若 `execCommand` 返回 `false`（浏览器策略拒绝），提示用户改用 `Ctrl+V`，不额外实现一套剪贴板读取逻辑——避免约定第 16 条点名的"自造平行规则"在剪贴板上重演。

### 1.5 需要的一处核心库最小改动

`data-docx-comment-ids`/`data-docx-link` 已经是"渲染时把结构信息回写到 DOM dataset，供 UI 层 hit-test"的既有模式。为了让右键点在修订标记上时能定位到具体是哪几条修订，需要在 `src/editor.ts` 的 `appendRun()` 里补一个同构属性：

- `run.revisions?.length` 非空时，给对应 `runSpan` 加 `dataset.docxRevisionIds = run.revisions.map(r => r.id).join(',')`（`RevisionMark.id` 已存在，纯读取，不改 `src/document.ts`）；
- 同时给每个 run span 补 `dataset.docxRun = String(run.index)`（目前只有段落级的 `dataset.paragraph`，没有 run 级标记）。这一项是解决"超链接跨多个 run 时如何精确定位到具体是 `getHyperlinks()` 里哪一条"的最小手段：右键命中点先拿 `paragraph` + `run` 两个索引，再用 `getHyperlinks().find(h => h.paragraph === p && h.runs.includes(r))` 精确解析，不需要新增 `document.ts` API。

这两个属性都是**纯只读回显**，不构成新的写入面，且和既有 `dataset.docxCommentIds` 是同一批次里的兄弟属性，改动集中在 `appendRun()` 一处，符合"核心库改动尽可能小"。

---

## 2. 六种右键位置的菜单清单

### 2.1 位置判定优先级（命中测试顺序）

右键触发时，从 `event.target` 沿 DOM 祖先链依次 `closest()` 判定，**按下列优先级取第一个命中的类别**（更具体的结构优先于更外层的）：

1. `[data-image]`（图片，含浮动图片外层 wrapper）
2. `[data-table-cell="true"]`（表格单元格——图片和表格可能同时命中，图片优先，因为图片是更内层、更具体的操作对象；表格上下文仍会体现在同一菜单里，见 2.4 的"表格 + 图片"组合规则）
3. `[data-docx-link="1"]`（超链接）
4. `[data-docx-revision-ids]`（修订标记）
5. `[data-docx-comment-ids]`（批注锚点）
6. 都不命中：普通正文（段落）

如果命中类别 3/4/5 中的多个（例如带修订标记的超链接文字），菜单**叠加**展示对应分组，不是互斥关系——这是 Word 的实际行为：右键点在"被批注过的超链接"上，菜单里超链接分组和批注分组都会出现。

### 2.2 「正文」菜单

- 剪切 / 复制 / 粘贴（`clipboard.cut`/`clipboard.copy`/`clipboard.paste`）
- ── 分隔线 ──
- 字体…（打开与 Ribbon「开始·字体」等价的格式弹层，或直接列出加粗/倾斜/下划线三个勾选项 + 字号/字体色子菜单）
- 段落…（对齐、编号/项目符号、缩进）
- 项目符号 / 编号（子菜单，同 Ribbon）
- ── 分隔线 ──
- 插入批注（`comment.addAtSelection`，仅当 `!ctx.selection.collapsed` 时 `enabled`——Word 要求先选中文字再插入批注锚点在该范围上；若 collapsed 则禁用并可选提示"请先选择文字"）
- 插入超链接…（`hyperlink.insertAtSelection`，同样要求非折叠选区或直接在插入点插入纯文字链接，两种入参走既有 `insertHyperlink` 的 `target`）
- ── 分隔线 ──
- 段落样式 ▸（子菜单，等价 Ribbon「开始·样式」下拉）

### 2.3 「表格内」菜单（在 2.2 基础上追加，出现在顶部）

- 插入 ▸（在上方插入行 / 在下方插入行 / 在左侧插入列 / 在右侧插入列）
- 删除 ▸（删除行 / 删除列 / 删除表格）
- 合并单元格（`ctx.selection.range` 跨多个 `table` 命中格时启用；单格右键时禁用或不出现，与 Ribbon 的合并按钮判定同一 `enabled`）
- 拆分单元格…
- 单元格底纹…
- 应用表格样式…
- ── 分隔线 ──
- （下接 2.2 的正文菜单其余项，因为单元格内本质仍是段落）

`ctx.table.nested === true` 时，"插入/删除/合并/拆分"整组 `enabled: false` 并保留提示文案，"单元格底纹"允许（不涉及结构，风险低，与 Ribbon 计划 2.2.1 的决定一致），维持两边判断完全一致——这正是命令注册表要解决的问题：这条"嵌套表格降级规则"只写一次，Ribbon 和右键菜单自动同步。

### 2.4 「图片上」菜单

- 剪切 / 复制（对图片：复制的是图片本身，等价选中该图片后走既有剪贴板路径；粘贴不适用于"图片上"这个位置，不出现"粘贴"这一行——因为粘贴目标应该是插入点而非替换图片，避免用户误解）
- 更换图片…（`image.replace`）
- 设置替代文字…（`image.setAlt`）
- 大小和位置…（若已有页面设置里的图片尺寸/环绕方式 API，接入；若目前没有对应公开方法，本项标注为**依赖新增 API，不在本次范围**，先不放进菜单，等对应能力落地再加）
- 删除图片（`image.delete`）
- 若该图片所在段落同时在表格里：追加"表格内菜单"的分隔线 + 精简版表格命令（仅"单元格底纹"，其余结构命令按 2.3 的 `nested`/常规规则处理，因为图片本身不改变单元格结构可用性判定）

### 2.5 「超链接上」菜单

- 打开超链接（`hyperlink.open`，`enabled: !ctx.hyperlink.unsafe`——不安全链接不提供"打开"，理由见约定第 6 条：安全判定在读路径，展示层同样要以解析出的最终目标为准，不能因为用户点了右键就绕过）
- 复制超链接地址（`hyperlink.copyAddress`，同样对 `unsafe` 链接禁用，避免用户复制出一个看似正常实则指向 `javascript:`/`vbscript:` 的地址后粘到别处触发风险）
- 编辑超链接…（`hyperlink.edit`，打开编辑弹层，预填 `updateHyperlink` 的入参）
- 取消超链接（`hyperlink.remove`，即 `removeHyperlink(..., { keepText: true })`，保留文字）
- ── 分隔线 ──
- （下接 2.2 的正文菜单，超链接文字仍然是可编辑文本）

`unsafe` 链接额外在菜单顶部插入一条不可点击的说明行（`role="separator"` 之外的纯提示，`aria-disabled="true"`），文案类似"此链接目标不安全，已禁用打开/复制"，避免用户以为菜单缺项是 bug。

### 2.6 「修订标记上」菜单

- 接受修订（`revision.acceptAtPoint`，对应 `doc.acceptRevision(id)`，逐条列出该点命中的每条修订，超过一条时用子菜单"接受此处修订 ▸"列出各条摘要如"插入：张三 2024-01-01"）
- 拒绝修订（`revision.rejectAtPoint`，同上，子菜单"拒绝此处修订 ▸"）
- 接受此段的所有修订 / 拒绝此段的所有修订（对 `ctx.selection.paragraph` 范围内的修订批量操作，复用 `getRevisions({ })` 按段落过滤后逐条调用，仍是同一批修订相关 API，不新增写方法）
- ── 分隔线 ──
- 修订选项…（跳转/联动 Ribbon「审阅」选项卡的显示设置，如果技术上做不到"跳转到某个选项卡"就退化成打开一个等价的小面板，不阻塞本次范围）

### 2.7 「批注上」菜单

- 回复批注…（`comment.replyAtPoint`）
- 标记为已解决 / 取消已解决（`comment.toggleResolvedAtPoint`，三态用 `checked` 表达）
- 删除批注（`comment.deleteAtPoint`）
- 定位到批注面板（滚动侧栏批注列表到该条，纯 UI 行为，不涉及文档）
- 命中多条批注时（`commentsAtPoint.length > 1`，批注可以嵌套/重叠），每条独立成一个子分组，用批注作者+摘要区分，不合并成一条

---

## 3. 与只读预览模式（`revisionView: 'final' | 'original'`）的关系

**总规则**：菜单照常弹出，**不因为只读就整体不显示**——用户仍需要能"看到自己此刻能做什么"，这与 Ribbon 计划第 4 节"上下文选项卡本身的出现不受只读影响，只是里面按钮全灰"是同一条原则的右键菜单版本。

逐项：

| 命令类别 | 只读模式下 | 理由 |
|---|---|---|
| 剪切 / 粘贴 | 禁用 | 会调用写方法 |
| **复制** | **保持可用** | 复制读的是当前渲染文本，不写文档；只读模式下用户复制内容看到的文档是常见需求 |
| 字体/段落格式化、编号、超链接编辑/取消、图片操作、表格结构操作 | 禁用 | 全部落到 `doc.set*`/`doc.insert*`/`doc.delete*`/`doc.format*` |
| **接受修订 / 拒绝修订** | **保持可用** | 这是本题面提示的关键点：接受/拒绝作用于底层文档模型（`doc.acceptRevision`/`rejectRevision`），不依赖 `contentEditable`；只读预览恰恰是审阅场景下最常用来"边看最终效果边决策"的模式，禁用它会让只读模式变得没法审阅，违背功能本意 |
| **打开超链接 / 复制超链接地址** | 保持可用 | 纯读取/导航，不写文档 |
| **回复批注 / 删除批注 / 标记已解决** | 禁用 | 写批注内容——与 Ribbon 计划第 4 节"新建/回复/解决/删除批注禁用，浏览类保留"完全一致 |
| **定位到批注面板** | 保持可用 | 纯 UI 导航 |
| 插入批注 / 插入超链接（正文菜单里的） | 禁用 | 写操作 |

判定信号复用同一个 `ctx.editable`（即 `reviewFilter.revisionView === 'markup'`），命令的 `enabled` 里对写类命令统一加 `ctx.editable &&`前缀，读类/审阅决策类命令不加——这条"哪些命令要不要跟 `editable` 挂钩"的清单，正是命令注册表要解决的"两边不分叉"的典型场景：Ribbon 计划第 4 节和本节本质是同一张表在两种呈现层的投影，**只维护一份**（放在每个 `CommandDescriptor.enabled` 里，Ribbon 的分组禁用规则改成"读该分组下所有命令的 `enabled` 取或"，不再是 Ribbon 计划里"整组硬编码禁用"的写法）。

---

## 4. 浏览器原生菜单的处置规则

| 场景 | 处置 | 理由 |
|---|---|---|
| 右键点在编辑区（`.docx-editor` 内）任意位置，包括正文、表格、图片、超链接、修订标记、批注锚点 | `preventDefault()`，弹出自定义菜单 | 这是本次要建的功能区域 |
| 右键点在 Ribbon/侧栏/开发者面板等**编辑区之外**的 UI 上 | **不拦截**，放行浏览器原生菜单 | 这些位置本来就不该有文档相关的菜单，拦截了反而没有"检查元素"等原生能力 |
| 右键点在**不安全的超链接**（`unsafe: true`）上 | 拦截，走自定义菜单，但菜单里"打开/复制地址"禁用（见 2.5） | 不能让原生菜单的"在新标签页中打开链接"绕过安全判定直接打开危险协议 |
| 右键点在**图片**上 | 拦截，走自定义菜单里的"更换/删除/设为替代文字"等；**不提供**浏览器原生的"图片另存为/在新标签页中打开图片"等同能力，因为文档内图片是 base64/blob 渲染，另存为对使用者没有实际意义（存下来的是渲染临时资源，不是文档里的原始 media part），提供该项反而会造成误导 | 与题面提示相反——题面举的"图片另存为"例子在 docx-view 这种"图片即文档数据"的场景下不成立，需要指出这一点而不是照抄通用规则 |
| 拼写检查（浏览器原生拼写下划线 + 右键建议） | **放行**：`contenteditable` 区域的原生拼写检查建议依赖浏览器自己维护的原生 `contextmenu` 分支，**这部分无法与自定义菜单共存**——一旦 `preventDefault()`，浏览器原生拼写建议也会一起消失。经过验证（Chrome/Firefox/Safari 均如此），**本项目选择拦截优先**：文档编辑场景下，结构化命令（剪切/加粗/表格操作等）比原生拼写建议更核心，且原生拼写建议本身也不该把"建议词"写回文档产生副作用（那属于原生输入法行为，不受本项目控制）。取舍已经确定：**不做浏览器分支判断去保留拼写菜单**，统一走自定义菜单；如果需要拼写建议，用户可以用系统级/浏览器级快捷键触发（不在本项目控制范围）。 |
| Shift + 右键（浏览器约定的"强制显示原生菜单"手势） | **放行原生菜单**，不拦截 | 这是浏览器留给用户的逃生舱，几乎所有网页富文本编辑器都遵守这条约定（Google Docs、Notion 均如此），保留它不需要额外开发——只需在 `contextmenu` 处理器里判断 `event.shiftKey` 为真时直接 `return`，不调用 `preventDefault()` |

---

## 5. 位置 vs 选区：本项目采用 Word 的规则，并说明原因

**规则**：右键点击落在当前选区**内部**（包括选区边界）时，**保留选区**，菜单基于该选区的上下文弹出；右键点击落在选区**外部**时，**先把选区折叠为插入点到点击处**（等价一次 `editor.selectParagraph`/`Selection.collapse` 到点击坐标），再基于新的插入点弹菜单。

理由：

1. **这是用户心智模型的最小惊讶**：选中一段文字后想对"这段文字"做操作（复制/加粗/插入超链接）是最常见的右键动机，如果右键会把选区打散，用户刚选的内容就丢了，必须重新选一遍——这是所有主流富文本编辑器（Word、Google Docs、VS Code）共同遵守的行为，不遵守会显得"这个编辑器很奇怪"。
2. **与本项目既有的 `docx-rangechange`/`selectedRangeInfo` 机制天然契合**：`selectedRangeInfo` 已经是"选区变化才更新"的持续状态（`editor.ts:1692` 起的去重逻辑），右键处理器只需要判断"点击坐标是否落在当前 `Selection` 的 `getRangeAt(0)` 矩形范围内"（用 `Range.getBoundingClientRect()` 判定近似即可，跨行选区需要逐 `getClientRects()` 判定），命中则不调用任何 `selectParagraph`/`collapse`，未命中才调用——**不需要新增状态**，是对既有信号的一次只读判断。
3. 折叠到点击处使用浏览器原生 `document.caretRangeFromPoint(x, y)`（Firefox 用 `caretPositionFromPoint`，做特性探测降级），拿到的 DOM `Range` 换算成 `{ paragraph, offset }` 再调用既有的 `editor` 内部选区更新路径——这一步和普通鼠标点击定位插入点用的是同一条既有代码路径，不新增。

---

## 6. 键盘可达性

### 6.1 触发方式

- **`Shift+F10`**：在编辑区（`.docx-editor` 内任意可聚焦位置）按下时，以**当前插入点/选区**位置为准（不涉及鼠标坐标）弹出菜单，位置锚定在插入点对应的 `Range.getBoundingClientRect()`附近。
- **`ContextMenu` 键**（部分键盘有独立的菜单键）：同 `Shift+F10`，做特性等价绑定（`event.key === 'ContextMenu'`）。
- 二者都不依赖鼠标事件，走同一套"以当前选区/插入点构造 `CommandContext`"的路径，不重复实现命中测试逻辑——直接复用 §1.2 里"持续状态"部分（`table`/`image`/`selection`），**唯独 `hyperlink`/`revisionsAtPoint`/`commentsAtPoint` 这三项"现查 DOM"的字段**在键盘触发时改为"以当前插入点所在的 run 反查"，而不是"以鼠标点击点反查"——这是键盘路径与鼠标路径唯一的分支点，其余全部共用。

### 6.2 菜单内导航

- `ArrowDown`/`ArrowUp`：在同级菜单项之间移动焦点，跳过分隔线和禁用项（禁用项可以被"路过"但不可被激活）。
- `ArrowRight`：若当前项有子菜单，展开并把焦点移入子菜单第一项；否则不响应。
- `ArrowLeft`：若当前处于子菜单内，收起子菜单并把焦点移回父菜单项；若已在顶层菜单，不响应（不做"跳到上一个菜单"这类跨菜单栏行为，因为这是上下文菜单不是菜单栏）。
- `Enter`/`Space`：激活当前聚焦项（等价点击）。
- `Esc`：关闭整个菜单（包括所有展开的子菜单），焦点归还（见 6.4）。
- 输入可打印字符：跳转到以该字符开头的下一个菜单项（Word/大多数原生菜单的"首字母跳转"约定），可选增强，不阻塞核心可达性。

### 6.3 ARIA 属性

- 菜单容器：`role="menu"`，`aria-label="文档编辑上下文菜单"`；子菜单容器同样 `role="menu"`。
- 每个菜单项：`role="menuitem"`；三态命令（如加粗）用 `role="menuitemcheckbox"` + `aria-checked="true|false"`；分组内互斥的命令（不在本次范围内暂无，预留）用 `role="menuitemradio"`。
- 禁用项：`aria-disabled="true"`（不用 `disabled` 属性，因为 `role="menuitem"` 落在 `<div>`/`<li>` 上而非原生表单控件，`aria-disabled` 是正确的语义化方式），同时 `tabindex="-1"` 但仍可被方向键"路过"聚焦到（可读出"已禁用"但不可激活）。
- 有子菜单的项：`aria-haspopup="menu"` + `aria-expanded`（随子菜单展开/收起切换）。
- 分隔线：`role="separator"`。
- 顶层菜单容器打开时设置 `aria-activedescendant` 指向当前聚焦项（或直接把 DOM 焦点移到该项，二者选一种一以贯之，建议选**真实 DOM 焦点移动**，因为菜单项之间存在展开/收起的动态子树，`aria-activedescendant` 在这种场景下更容易和实际渲染状态失配）。

### 6.4 焦点归还

- 菜单关闭（无论是 `Esc`、点击外部、执行了某条命令、还是移动到新的右键位置重新弹出前先关闭旧的）后，焦点必须归还到**触发菜单前的焦点元素**：
  - 键盘触发（`Shift+F10`）：归还到触发前编辑区里聚焦的段落/run（即恢复 `Selection` 到菜单打开前的 Range，因为编辑区本身用 `contenteditable` 而非常规可 `focus()` 的表单控件，"焦点"体现在 `Selection` 位置而非某个 DOM 节点的 `.focus()`）。
  - 鼠标右键触发：同样归还 `Selection` 到"§5 判定后的最终选区/插入点"（即如果是选区内右键，归还原选区；如果是选区外右键，归还折叠后的新插入点——不是"点击前"的选区，因为按 §5 的规则，选区已经在弹出菜单前就已经变化了，归还目标是变化后的状态）。
  - 执行了会改变文档结构的命令（如"删除行"）后：命令执行完毕、`flush()`/重新 `render()` 之后，按既有的"结构变化后如何重新定位插入点"的逻辑处理（这部分复用现有 `deleteTableRow` 等操作后 Ribbon/主程序里已有的"重新选中一个合理段落"逻辑，不重新发明）。

---

## 7. 防止流回文档

### 7.1 具体做法

- 菜单 DOM 节点**挂载在编辑区容器之外**（例如 `document.body` 下的一个独立浮层容器，通过 `position: fixed` 定位到触发坐标），**不作为 `.docx-editor` 的子节点**——这是比"挂在编辑区内但设 `contentEditable=false`"更彻底的隔离，因为 Ribbon 计划第 4 节讨论的"装饰元素"（编号标记、制表符前导符）之所以需要 `data-docx-mark`/`contentEditable=false` 双重保险，是因为它们**必须**渲染在编辑区内部（要跟随文字排版）；右键菜单不需要跟随文字排版，没有理由承担"混入编辑区"的风险，直接物理隔离是更简单也更彻底的方案。
- 唯一的例外是"菜单弹出期间编辑区会不会被间接改写"：命令执行体（`run(ctx)`）里凡是会调用 `doc.set*`/`insert*`/`delete*`/`format*` 的命令，执行前必须先对编辑区调用既有的 `editor.flush()`（提交挂起的文本变更），避免"用户正在输入、还没 blur、右键点了别的命令"这种时序下丢失或错乱未提交的编辑——这条不是"防流回"本身，是防止菜单命令和未提交编辑互相踩踏，一并写进本节的实现清单。

### 7.2 需要的回归测试

1. **零 DOM 侵入**：`document.body.querySelectorAll('.docx-editor [class^="context-menu"], .docx-editor [role="menu"]')` 断言恒为空——菜单容器永远不是编辑区的后代，用一次结构断言钉死，而不是依赖"没有副作用"这种运行时行为断言。
2. **弹出 → 关闭 → flush() 不改变文档**：模拟"右键弹出菜单（不点任何命令，直接 `Esc` 关闭）→ 调用 `editor.flush()`"，断言 `doc.getSnapshot().revision` 与菜单弹出前完全一致，且 `doc.getParagraphs()[i].text` 逐段落相等（这是题面点名要的测试，直接对应约定第 4 条的既有回归测试模式，只是触发场景换成右键菜单）。
3. **键盘路径同样验证**：`Shift+F10` 弹出、`ArrowDown` 若干次、`Esc` 关闭，同样断言 revision 不变——覆盖"方向键导航本身不该触发任何 `doc.*` 调用"这条隐含要求（导航只应该移动 DOM 焦点，不应该有任何 side effect）。
4. **命中态清空后菜单联动关闭**：文档发生结构变化导致右键命中的批注/修订/单元格已经不存在时（例如另一个来源并发接受了同一条修订），菜单应能优雅关闭或降级隐藏对应项，不抛错、不残留已失效的命令绑定引用（防止"点了一个指向已删除批注 id 的命令"这种悬空引用崩溃）。

---

## 8. Issue 拆分与依赖顺序

在 Ribbon 计划原有 Issue A/B/C/D 基础上，插入并追加：

### Issue R：抽取统一命令注册表（新增，最优先）
- **依赖**：无；**必须先于**本计划所有其它 Issue，且**建议在 Ribbon 计划 Issue A 之后、Issue C 之前**完成（Issue A 已经把控件搬进 Ribbon 容器，Issue R 在此基础上把 click handler 收敛进注册表；如果 Ribbon Issue A/C 已经先合并，Issue R 就是一次后续重构 PR，替换掉分散的 handler，不改变任何可见行为）。
- **范围**：新增 `examples/commands.ts`，定义 §1.1 的 `CommandDescriptor`/`CommandContext` 形状；把 Ribbon 现有 43 个控件里**会被右键菜单复用的那些**（字体/段落格式、剪切复制粘贴、表格结构操作、图片操作、超链接、修订、批注共约 30 项，纯导航类如"显示大纲开关"不必进注册表）迁移成描述符；Ribbon 按钮的 click handler 改为 `registry.run(id, buildCommandContext())`，`disabled`/`aria-pressed` 改为订阅同一批既有事件后调用 `registry.get(id).enabled(ctx)`/`checked(ctx)` 刷新。
- **验收标准**：
  1. `npm run check && npm test` 通过；
  2. Ribbon 计划 §3.2 的全部既有回归测试保持通过（尤其"纯选项卡切换不改文档"——这条重构不应该引入任何新的写入触发点）；
  3. 新增测试：对每条迁移的命令，`registry.get(id).enabled(ctx)` 在若干典型 `ctx`（表格内/外、只读/可写、有选区/无选区）下的结果与迁移前对应按钮的 `disabled` 状态逐一比对，**不允许出现行为差异**；
  4. 手工回归清单同 Ribbon 计划 Issue A 的 43 项功能行为不变。
- **不在范围**：右键菜单 UI 本身、新增 `data-docx-revision-ids`/`data-docx-run` 属性（那是 Issue S 的事）。

### Issue S：核心库最小新增 —— `data-docx-revision-ids` / `data-docx-run`
- **依赖**：无，可与 Issue R 并行开发，建议先于 Issue T 合并。
- **范围**：`src/editor.ts` 的 `appendRun()` 补上 §1.5 的两个 dataset 属性；纯只读回显，不新增 `document.ts` 公开方法。
- **验收标准**：
  1. 单测：给定一个含多条 run 级修订的段落，渲染出的 DOM 里每个受影响 run span 的 `dataset.docxRevisionIds` 与 `run.revisions.map(r=>r.id)` 一致；无修订的 run 不带该属性（不产出空字符串属性）；
  2. 单测：`dataset.docxRun` 对同一段落内的所有 run 严格按 `run.index` 一一对应，段落重排（插入/删除 run）后重新渲染结果同步更新；
  3. 回归：既有超链接/批注相关的 dataset 属性（`docxLink`/`docxCommentIds`/`docxUrl` 等）不受影响；
  4. 依照约定第 14 条，测试用例要覆盖"同一段落内两条不同修订分别挂在不同 run 上"和"同一个 run 同时挂多条修订"两种形状，不能只测单条修订的简单情形。
- **不在范围**：右键菜单 UI、命令注册表。

### Issue T：右键菜单容器 + 位置判定 + 六种位置的静态菜单渲染
- **依赖**：Issue R（读取命令注册表）+ Issue S（`hyperlink`/`revisionsAtPoint` 命中测试需要新属性）+ Ribbon 计划 Issue B（`getTableCellAt`/`selectedTableCell`，`table` 上下文字段需要）。
- **范围**：新增 `examples/context-menu.ts`：`contextmenu` 事件监听、§4 的原生菜单拦截规则、§5 的位置 vs 选区判定、§2 六种位置的菜单内容渲染（读注册表 `visibleInMenu`/`enabled`/`checked` 生成 DOM）、§6 的键盘导航与 ARIA、§7.1 的浮层挂载方式。
- **验收标准**：
  1. `npm run check && npm test` 通过；
  2. §7.2 的四条回归测试全部通过；
  3. 六种位置（正文/表格/图片/超链接/修订/批注）各自的菜单项清单与 §2 的规格逐条比对无遗漏无多余；
  4. 键盘 `Shift+F10` 打开、方向键导航、`Esc` 关闭、焦点归还，均有对应测试；
  5. `Shift+右键` 放行原生菜单有测试覆盖（可以是对 `preventDefault` 是否被调用的断言，不需要真实验证浏览器原生菜单渲染）。
- **不在范围**：只读模式的禁用规则细化（Issue U）、修订/批注菜单里"接受/拒绝/回复"命令本身的执行体（如果 Issue R 迁移范围没覆盖到，本 issue 里把它们也补进注册表，但不重新设计判定逻辑，直接复用既有 `doc.acceptRevision` 等 API）。

### Issue U：只读预览模式下的菜单可用性规则
- **依赖**：Issue T（菜单已存在）+（如果 Ribbon 计划 Issue D 已先落地）复用其中 `reviewRevisionView !== 'markup'` 的判定信号，否则本 issue 里顺带把该信号接入命令注册表的 `enabled`。
- **范围**：按 §3 的表格，给注册表里对应命令的 `enabled` 加 `ctx.editable &&`（写类）或不加（读类/修订决策类/批注浏览类），确保 Ribbon 和右键菜单**同时**因为这一次改动而生效（这正是 Issue R 抽取注册表的收益所在：这里只改一处）。
- **验收标准**：
  1. `final`/`original` 模式下，右键菜单里"剪切/粘贴/格式化/表格结构/图片操作/新建批注"禁用或不出现，"复制/接受修订/拒绝修订/打开超链接/复制链接地址/定位到批注面板"保持可用；
  2. 同一时刻 Ribbon 对应按钮的禁用状态与右键菜单一致（新增一条"双呈现层一致性"测试：对同一个 `ctx`，遍历注册表里所有命令，断言 Ribbon 侧读到的 `enabled` 与右键菜单侧读到的 `enabled` 相等——这条测试是本计划相对 Ribbon 计划的核心增量，专门钉住"命令模型是否真的没有分叉"）；
  3. `npm run check && npm test` 通过；
  4. README/示例文档补充"只读预览模式下右键菜单行为"一节，与 Ribbon 计划 Issue D 的对应章节合并撰写，避免文档本身出现两份不一致的描述。
- **不在范围**：批注/修订以外的其它只读态交互变化。

**总顺序**：`Issue R`（可与 Ribbon Issue A 并行，收尾稍晚于 A）→ `Issue S`（可与 R 并行）→ `Issue T`（依赖 R + S + Ribbon Issue B）→ `Issue U`（依赖 T，且与 Ribbon Issue D 尽量同批完成以合并文档）。每个 issue 单独 PR，全绿再进下一个，避免约定第 15 条描述的"合并时才发现两条读路径不一致"重演——这也是本计划把 Issue R 单独拎出来、放在最前面的根本原因。

