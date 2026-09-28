# 审阅功能设计说明

对应 roadmap（issue #2）阶段 4 的「修订」与「批注」两大块，拆分为 issue #36 ~ #38、#40、#42 ~ #44。

> **来源与可信度**：本文初稿由本地 GitHub Copilot CLI 基于当时的代码库生成，随后由维护者逐条核实与修正。文中标注「已实测」的结论都在 `main` 上跑过验证脚本；原稿中一处描述有误（行级修订写成 `w:val="1"`，实际产出是裸元素）已纠正；原稿中若干「需要先确认」的存疑已查实并写成结论。基线提交为 `a7fe86d`。

## 1. 现状（已实测）

| 事实 | 证据 |
| --- | --- |
| **完全没有 run / 段落级修订支持** | `w:ins` / `w:del` / `w:rPrChange` / `w:pPrChange` 仅出现在 `PROPERTY_ORDER` 的顺序表里，无任何解析逻辑 |
| **完全没有批注支持** | `src/` 下 `comment` 字样一处都没有 |
| **`w:ins` 包裹的 run 已被当普通 run 读出** | 段落含 `<w:ins><w:r><w:t>inserted</w:t></w:r></w:ins><w:r><w:t>keep</w:t></w:r>` 时，`text` 读为 `"insertedkeep"`、`runs.length === 2` |
| **`w:delText` 不在文本白名单** | `textElements()` 只认 `t/tab/br/cr/noBreakHyphen/softHyphen/sym`；已跟踪删除的 run 读出空串，但**仍占 `runs` 索引** |
| **`isTransparentWordWrapper()` 不含 `ins` / `del`** | `TRANSPARENT_WORD_WRAPPERS` 只有 `sdt` / `sdtContent` / `customXml` |
| **已有一份不完整的行级修订实现（缺陷）** | 见下节 |

### 1.1 发现的既有缺陷

**（a）行级修订标记缺 `w:id`** —— 已并入 #36 修复。

`RowFormat.deleted` / `inserted` 通过通用的 `boolValue()` 写入：

```js
d.insertTable([['a','b']]);
d.formatTableRow(0, 0, { deleted: true, inserted: true });
// 实际产出：<w:trPr><w:ins/><w:del/></w:trPr>
```

`w:trPr/w:ins` 与 `w:trPr/w:del` 的类型是 **`CT_TrackChange`**：`w:id` 是必需属性，通常还应带 `w:author` / `w:date`，并且**没有 `w:val` 属性**。当前产出的裸元素缺 `w:id`，Word 会判定为无法读取的内容。这份实现随 PR #9 进入代码库，当时的 review 未发现。

**（b）settings 顺序表拼写错误** —— 已并入 #36 修复。

`src/notes.ts` 的 settings 顺序表把开关写成 `trackRevisions`，规范中的元素名是 `w:trackChanges`。今天无害（无代码写它），但落地「记录修订」开关时会让 `orderedProperty()` 找不到位置而退化为追加到末尾。

## 2. 架构决策

### 2.1 修订标记写入器只能有一份

`markRevision()`（#36 建立）是 `w:ins` / `w:del` / `*Change` 的**唯一写入入口**，`w:id` 分配、`w:author`、`w:date` 都由它负责；`table.ts` 现有的行级修订要收敛进来，#37（写入）、#38（接受/拒绝）、#43（比较）都必须复用它，不得另造平行实现。

`w:id` 分配**扫描文档内已用最大值 + 1**，不用全局自增变量（多次编辑后会冲突或不确定）；书签 id 分配已有同款实现可参照。

### 2.2 已跟踪删除的文本与正文文本必须分开

- `w:ins` 包裹的内容**属于正文**（Word 里也显示），继续计入 `text` 与 `runs`，但 `RunInfo` 上要能看出它带插入标记；
- `w:del` 包裹的内容**不属于 `text`**，单独字段暴露。

理由：若把两者混进同一个 `text` 字符串，写回时无法区分，直接违反工程约定第 5 条「读出来的格式/内容要能原样写回」。

### 2.3 「遍历要穿透」不等于「读文本要跳过」

`w:ins` / `w:del` 在遍历意义上是透明包裹（`ownRuns()`、图片读取、表格行列遍历都要能穿透），因此要纳入 `isTransparentWordWrapper()`；但**读文本时 `w:del` 内的内容要排除、`w:ins` 内的要保留**。两件事必须分开处理，不能用同一个判定糊过去。

### 2.4 装饰 vs 内容：审阅功能上的质变

工程约定第 4 条原本是「界面装饰元素绝不能流回文档」，靠 `readText()` 跳过 `contentEditable === 'false'` 与 `dataset.image` 实现。审阅功能打破了这个二分：

| 类别 | 例子 | 处理 |
| --- | --- | --- |
| 纯渲染装饰 | 变更条、作者色底纹、批注气泡与侧栏、「已解决」角标、接受/拒绝按钮 | 走 `makeMark()`：`contentEditable="false"` + `data-docx-mark`，`readText()` 跳过 |
| **文档内容** | `w:ins` / `w:del` 包裹的文字、`w:commentReference`、`w:commentRangeStart` / `End` | **不能当装饰过滤掉**，否则 `flush()` 后修订标记与批注锚点会消失 |

最危险的一条组合：`w:del` 的旧文本在「显示标记」视图里要**显示出来**（带删除线），但它**不属于 `text`**——渲染这段文字时必须让 `readText()` 跳过，否则一次 `flush()` 就会把已删除的旧文本写回正文。#44 需要为此专门写测试。

### 2.5 视图状态不得修改文档

`revisionView: 'final' | 'original' | 'markup'` 是**纯渲染状态**：切换它不推进 `revision`、不触发 `onChange`、不改变 `getParagraphs()[i].text`、`getPartXml()` 完全相同。

`'final'` 表示「假装所有修订已接受」来显示，极易被实现成「临时调用 `acceptAllRevisions()` 再渲染」——**不要那样做**。

## 3. Issue 拆分与依赖

```
#36 修订基础设施（只读解析 + 记录修订开关 + markRevision + 修复 1.1 的两处缺陷）
 ├── #37 修订写入（编辑操作产生 w:ins / w:del / rPrChange）
 │    └── #43 文档比较（复用 markRevision 生成差异标记）
 ├── #38 接受与拒绝修订（可用手写 fixture 并行开发，不强依赖 #37）
 └── #42 审阅者聚合与筛选视图（需要 getRevisions）
#40 批注（comments.xml / commentsExtended.xml）
 └── #42 审阅者聚合与筛选视图（需要 getComments）
#44 审阅渲染层（依赖以上全部，最后做）
```

刻意让 #38 只依赖读取能力，这样它不会被 #37 的进度卡住。

## 4. 与现有模块的集成点

| 模块 | 需要做什么 |
| --- | --- |
| `document.ts` | `ownRuns()` / `textOf()` / `readRun()` / `readParagraph()` 感知 `w:ins` / `w:del`；`textOf` 处理 `w:delText`（独立字段，见 2.2）；`PROPERTY_ORDER` 里 `rPrChange` / `pPrChange` 位置已在末尾且符合序列，补一条「是否升序」自测即可 |
| `table.ts` | `RowFormat` / `CellFormat` 增加修订字段，现有 `boolValue(props,'del'/'ins')` 换成 `markRevision`；`tcPr` 的 `cellIns` / `cellDel` / `cellMerge` 一并接上，不要留成「行有修订、单元格没有」的半成品 |
| `styles.ts` | 基本不用改。`rPrChange` / `pPrChange` 存的是**直接格式快照**，不涉及样式定义 |
| `numbering.ts` | 不用改，但要有回归测试：**接受/拒绝修订不能打乱编号计数**（列表计数依赖段落顺序） |
| `drawing.ts` | 图片所在 run 被 `w:ins` / `w:del` 包裹时，`readRunImages` 的遍历要能穿透（见 2.3） |
| `notes.ts` | 脚注正文段落也可能带修订；`getFootnotes()` 的 `blocks` 走同一套 `readParagraph`，不要为脚注单写一套修订解析 |
| `section.ts` | `sectPrChange` 本期明确不支持，在 README 写明 |
| `hyperlink.ts` | `isUnsafeHyperlink` 判定不变；但遍历时要注意 `w:ins` / `w:del` 包裹层 |
| 缓存 | 新增与 `stylesCache` / `numberingContextCache` / `noteStateCache` 同级的修订与批注缓存，失效键仍是 `revision`；无修订 / 无批注时走快路径（参照 #27 的做法） |

## 5. 最容易踩的约定（按风险排序）

1. **第 2 条（在 run 的真实父节点插入）**：修订标记与批注锚点的插入点是全功能里最复杂的——一个 run 可能同时嵌在 `w:hyperlink > w:ins > w:sdtContent` 三层里。这批 PR 已有三次同类事故（#10 `insertImage`、#24 `insertFootnote`、#25 `insertHyperlink` 都曾用 `paragraph.insertBefore` 抛 `child not in parent`）。
2. **第 4 条（装饰 vs 内容）**：见 2.4，本功能特有的新坑。
3. **第 1 条（部件路径从关系解析）**：`comments.xml` / `commentsExtended.xml` 的文件名只是惯例；新建前先查既存关系——`ensureNotePart` 就因为没查而在 #24 出过「新增重复关系、Word 拒绝」的事故。
4. **第 5 条（格式对象双向一致）**：`rPrChange` 是「读出来的格式反过来当输入」的极端形式，一次可能快照几十个字段。落地 #37 前先跑通 `formatRun(0, 0, readBack(rPrChange))` 这条往返。
5. **第 7 条（一次事务一次提交）**：`acceptAllRevisions` / `rejectAllRevisions` / `compare()` / `addComment` 都是「内部循环或多部件」的方法，最容易写成多次提交。
6. **第 9 条（`w:val="0"` 是显式关闭）**：`w:trackChanges`、`w15:done` 都是布尔元素，缺失与显式 `false` 语义不同，实现前定策略并写测试。
7. **第 10 条（段落索引唯一命名空间）**：批注范围可能跨正文与单元格、或落在脚注 / 页眉里，不能用正文段落 index 硬套。#23 与 #24 已各踩过一次同类问题。

## 6. 与既有语义的冲突点

| 既有机制 | 冲突点 | 界定 |
| --- | --- | --- |
| `revision` 计数 | 接受/拒绝修订、切换记录修订开关算不算一次编辑？ | 只要文档字节实际变化就 +1；`acceptAllRevisions()` 内部处理 N 条也只 +1；`setTrackChanges(相同值)` 是 no-op 不 +1（另见 #35：目前 14/15 个编辑方法在 no-op 时仍会推进 revision，待统一收敛） |
| Agent 事务 | 一个批次里 `acceptRevision` 解包 `w:ins` 会改变后续操作的 run 索引 | 沿用既有规则「每个操作的索引相对于该操作执行前的状态」，但要专门测一条「接受修订后索引位移」的批次用例——这是本功能特有的新型位移来源 |
| 撤销重做 | 记录修订开启时，一次文字编辑会产生更多 DOM 变化（生成 `w:ins` + 可能拆分相邻 run） | 撤销粒度以**公开方法调用**为单位，不以 DOM 节点变化为单位。撤销栈本身在 issue #29 / PR #33 进行中（本文初稿生成时尚不存在） |
| `deleteParagraph` 的保护语义 | 拒绝一条「插入段落」的修订会让该段落整体消失 | 必须复用既有 `deleteParagraph` 的「破坏不变量时就地清空、不抛错」语义，不要另写删除路径 |

## 7. 通用验收要点

- 每个 issue **独立可验收**：#38 用手写的带 `w:ins` / `w:del` 的 fixture XML 直接测，不必等 #37。
- **畸形输入降级**：缺 `w:id`、`author` 为空、`date` 非法、`commentReference` 悬空——都必须安全读出而不抛错。
- **往返**：`accept(read())` / `reject(read())` 不抛错；`rPrChange` 快照能被现有 setter 直接接受。
- **顺序自测**：生成的 `w:ins` / `w:del` / `*Change` 子元素顺序与规范 sequence 做「是否升序」断言。
- **视图开关不改文档**：切换显示状态前后 `getPartXml()` 完全相同、`text` 不变、反复 `flush()` 不让标记叠加。
- **安全**：`author` 等自由文本走 `assertText`；批注正文不解析为富文本或 HTML，不新增可执行内容面。
- **跨模块叠加**：至少覆盖「表格单元格内的段落同时有编号、样式、图片，且被跟踪修订包裹」这类组合——只有集成阶段才测得到。
- **compare 端到端**：比较结果全部接受后等价于 `revised`，全部拒绝后等价于 `base`。
