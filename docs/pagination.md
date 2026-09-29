# 分页与打印视图（roadmap A）设计方案

> 状态：方案，未开工。
> 基线核对于 `main` = `1bf28af`（2026-09-29）。下面「现状」一节的每条都是读代码确认过的，不是推测。

## 为什么这是最大的一块

`src/` 里**没有一处代码知道「第几页」**。这不只是少一个功能：它挡住了域与目录里所有依赖页码的部分（`PAGE`、`NUMPAGES`、`PAGEREF`、`TOC`、`INDEX`），这些在 #88 里被明确排除掉了，等的就是这里。

## 现状

### 已经有的（比预想的多）

| 能力 | 位置 | 说明 |
| --- | --- | --- |
| 完整页面设置（读） | `src/section.ts` | `pageWidth` / `pageHeight` / `orientation` / `margins{top,right,bottom,left,header,footer,gutter}` / `columns{count,space,equalWidth,widths}` / `pageNumbering{start,format}` / `titlePage` / `headers` / `footers` |
| 节边界 | `collectSections()` | 每节的 `startParagraph` / `endParagraph` / `type`（`nextPage` / `continuous` / `evenPage` / `oddPage` / `nextColumn`） |
| 分页相关段落属性（读） | `src/types.ts:56-60` | `keepNext` / `keepLines` / `pageBreakBefore` / `widowControl` / `suppressLineNumbers` —— **全都已经能读，只是没有任何代码在用** |
| 分页符 / 分节符标记 | `DocumentBlock` | `{ type: 'pageBreak' }` 与 `{ type: 'sectionBreak', section, breakType }` 已经在块流里 |
| 文本度量 | `editor.ts:1183` | canvas `measureText`，目前只服务于制表位 |

### 还没有的

- **行盒与页面切分**：一行有多高、一页装得下几行，全无。
- **页眉页脚不随页重复**：`render()` 里在整份内容的前后各插一个，不是每页一份。
- **`applyPageSetup()` 只渲染一张连续的「纸」**：设 `max-width` + `padding` + CSS `columnCount`，没有页的概念。
- **而且它只读 `getSection(0)`**（`editor.ts:801`）。多节文档里第二节及以后的页面设置被完全忽略——比如一份中间插入横向页的文档，横向那节会按第一节的纵向尺寸渲染。这是一个**现存缺陷**，A2 顺带修掉。
- `pageBreak` / `sectionBreak` 现在只渲染成一行文字标记（`—— 分页符 ——`）。

## 架构约束：分页不能做成 `DocxDocument` 的方法

roadmap 全局约束第 3 条：**核心 API 不依赖浏览器全局 `document`**，只有 `src/editor.ts` 可以用真实 DOM。

分页必须有字体度量，而字体度量在 Node 里拿不到。如果把 `getPageCount()` 之类直接挂到 `DocxDocument` 上，要么破坏这条约束，要么这段逻辑**在 Node 里完全无法测试**——而本仓库 1098 个测试全部跑在 Node 里，`test/editor.test.mjs` 更是用 `Object.create(DocxEditor.prototype)` 手搭替身、根本没有真 DOM。

所以：

```
src/layout.ts          纯逻辑：块流 + 度量 → 页面。不 import 任何 DOM。
    ↑ 注入 LayoutMeasurer
src/editor.ts          浏览器实现：用真实排版量行盒
test/layout.test.mjs   确定性假实现：给定合成度量，切分结果必须逐字节确定
```

```ts
export interface LayoutMeasurer {
  // 把一段内容排进 widthPx 宽的盒子，回报每个行盒的高度与它吃掉的内容范围
  measureParagraph(paragraph: ParagraphInfo, widthPx: number, context: MeasureContext): LineBox[];
  measureTableRow(row: TableRowInfo, widthPx: number, context: MeasureContext): number;
}
```

**假实现必须是一等公民，不是测试的权宜之计。** 「12 号字每行 40 个字符、每行 16px 高」这种规则化度量，能让页面切分的所有边界条件（`keepNext` 连锁、寡行控制、表格跨页）在 Node 里被精确断言。真实浏览器度量反而测不了这些——它不可复现。

## 测量方式：优先浏览器排版，而不是 canvas

现有的 `measure()` 用 canvas `measureText` 量一整个 run 的宽度。用它做分页会撞上一堆问题：**它不做断行**，不懂 CJK 的禁则处理（#84 刚做完 `kinsoku` / `wordWrap` / `overflowPunct` 的读取）、不懂双向文字、不懂连字与字距调整。自己实现这些等于重写一个排版引擎。

**浏览器已经免费提供了这一切。** 把段落排进一个固定宽度的盒子，再用 `Range.getClientRects()` 读回行盒——断行、CJK 禁则、bidi、连字全部由浏览器处理好了。

所以：

- **浏览器实现**：隐藏的量度容器 + `Range.getClientRects()`，这是生产路径。
- **canvas `measureText`**：保留给制表位（现有用途），**不**作为分页的度量来源。
- **Node 假实现**：规则化度量，服务于测试。

## 分阶段

### A1 — 布局模型与页面切分（纯逻辑，Node 可测）

新建 `src/layout.ts`，不碰渲染。

输入：`DocumentBlock[]` + `SectionInfo[]` + `LayoutMeasurer`。
输出：

```ts
export interface PageBox {
  index: number;            // 0 起
  number: number;           // 显示用页码，受 pgNumType.start 与 evenPage/oddPage 影响
  section: number;
  items: FlowItem[];        // 落在本页的行盒 / 表格行 / 图片
  contentHeightPx: number;
}
```

必须处理的切分规则：

- 显式分页符（`w:br w:type="page"`）与 `pageBreakBefore`
- 分节符：`nextPage` / `evenPage` / `oddPage`（后两者可能要插一张空白页）/ `continuous`（不换页但换页面设置）/ `nextColumn`
- `keepNext`（与下一段同页，**连锁**：A keepNext B keepNext C 要一起走）
- `keepLines`（段落不拆页）
- `widowControl`（孤行寡行：默认开启，段落不得只留 1 行在页首或页尾）
- 每节的页面设置可能不同 —— 可用高度 = `pageHeight - margins.top - margins.bottom`

**防死循环**：一个行盒比整页可用高度还高时（巨大的图片、超大字号），必须放在单独一页并继续，不能无限找不到落点。这类退化情形要有明确测试。

### A2 — 只读分页预览渲染

编辑器按 A1 的结果渲染成 N 个页面元素。

- 修掉 `applyPageSetup()` 只读 `getSection(0)` 的缺陷，**每节用自己的页面设置**
- 页眉页脚**每页一份**，正确选择 `first` / `even` / `default`（`titlePg` 决定是否启用 first）
- 页码：`pgNumType.start` 与格式（`decimal` / `upperRoman` / `chineseCounting` …）
- 浏览器 `LayoutMeasurer` 实现
- **只读**。分页视图下不可编辑，容器 `contentEditable = 'false'`

**这一条必须是只读的。** 参考 #50 的教训：把不该编辑的内容渲染进可编辑区域，用户一打字就会产生伪造的数据。分页视图下的编辑是 A5，单独做。

### A3 — 表格跨页与分栏

- 表格跨页拆行；`w:tblHeader` 标题行在每页重复；`cantSplit` 的行不拆
- 分栏（`w:cols`）：先把列填满再换页
- 浮动图片与环绕对行盒的影响

### A4 — 页码相关的域

A1 出了页码之后，#88 里明确排除的那批才能做：`PAGE` / `NUMPAGES` / `PAGEREF` / `TOC` / `INDEX`。
注意 **`TOC` 的更新会改变文档长度，从而改变页码**——需要「排版 → 算域 → 重排」的收敛循环，并设迭代上限（Word 也是这么做的，且同样会不收敛）。

### A5 — 分页视图下的编辑

最难的一块，**可以一直推迟**：跨页边界的光标移动、每次击键的增量重排、选区跨页。
在 A2 提供了可用的只读预览之后，这一条的紧迫性会大幅下降。

## 与在飞 PR 的冲突面

| 阶段 | 主要文件 | 与 #90 / #91 / #92 的重叠 |
| --- | --- | --- |
| **A1** | 新建 `src/layout.ts`、`test/layout.test.mjs` | **无重叠，现在就能开** |
| A2 | `src/editor.ts` 大改 | 与 #91（`editor.ts` +52）重叠，**等 #91 合并** |
| A3 | `src/layout.ts` + `editor.ts` | 等 A2 |
| A4 | `src/fields.ts` | 等 #90（它创建这个文件）与 A1 |
| A5 | `src/editor.ts` | 等 A2 |

## 不做什么

- **不自己实现排版引擎**。断行、CJK 禁则、bidi、连字一律交给浏览器。
- **不追求与 Word 逐像素一致**。字体回退不同就不可能一致。目标是页面切分位置在常见文档上与 Word 吻合，而不是每一行的像素宽度相同。
- **不做打印对话框**：artifact 与浏览器沙箱里 `window.print()` 不可靠；导出 PDF 属于另一个话题。
