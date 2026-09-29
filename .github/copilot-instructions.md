# docx-view 工程约定

这份约定来自对历次 PR 的逐条 review。下面每条规则都对应真实发生过、而且**在多个 PR 里重复发生**的缺陷。开工前读一遍，能省掉一整轮返工。

第 1–15 条来自前 12 个 PR（#7–#10、#19–#22、#23–#26），按「踩坑代价」排序，不按代码结构排序。第 16 条起是后续批次新增的，**追加在末尾而不插队**——issue 与 review 里会按编号引用，编号必须保持稳定。新增的两条代价并不低，请一并读完。

---

## 1. 部件路径必须从关系解析，不能硬编码

**重复了 5 次**：#8（`numbering.xml`）、#10（`media/*`）、#24（`footnotes.xml`）、#7（`styles.xml`），以及 #10 合并进 `main` 时再次回归（畸形 `Target` 让 `getParagraphs()` 抛错）。

`word/numbering.xml`、`word/footnotes.xml`、`word/media/image1.png` 这些名字只是**惯例**，不是规范。真实文档里关系的 `Target` 可以是任何相对路径（`fn.xml`、`sub/numbering.xml`、带 URL 编码的名字）。

- 通过所在部件的 `.rels` 解析 `r:id` → 真实部件路径。
- **每个部件有自己的 `.rels`**：主文档是 `word/_rels/document.xml.rels`，页眉是 `word/_rels/header1.xml.rels`，脚注是 `word/_rels/footnotes.xml.rels`。所以解析函数必须接受「源部件路径」参数，不能写死主文档。
- `Target` 畸形（`..`、`sub\x.xml`、`num%bering.xml`、绝对路径）时**降级**，不要让 `getParagraphs()` / `getSnapshot()` 抛错——#8 最初就是这样让整个文档打不开的。
- 新建部件前先查有没有既存关系，否则会出现两条同类型关系，Word 直接拒绝（#24 实际发生）。

## 2. 在 run 的真实父节点上插入，不是段落上

**重复了 3 次**：#10（`insertImage`）、#24（`insertFootnote`）、#25（`insertHyperlink`）。

典型错误：

```ts
const runs = ownRuns(paragraph);          // descendants，可能来自 w:hyperlink / w:ins / w:sdtContent
paragraph.insertBefore(newRun, runs[i]);  // ✗ DOMException: child not in parent
```

`ownRuns()` / `descendants()` 返回的 run 可能嵌在 `w:hyperlink`、`w:ins`、`w:del`、`w:sdtContent`、`w:smartTag` 里。插入时用 `runs[i].parentNode.insertBefore(...)`，并且要想清楚新节点应该在包裹元素**内**还是**外**（给已有超链接的文字加脚注：引用应该在链接外；给链接文字改格式：在链接内）。

## 3. OOXML 子元素顺序是强约束，用 `PROPERTY_ORDER` / `property()`

**违反 4 次**（#7 `defineStyle`、#9 `tblPr`/`tcPr`、#24 `footnotePr`/`pos`、#10 `wp:anchor` 必需属性），**正确 2 次**（#23 `sectPr`、#26 `pPr`）。

`CT_PPr`、`CT_RPr`、`CT_TblPr`、`CT_TcPr`、`CT_SectPr`、`CT_Style`、`CT_Settings`、`CT_FtnProps` 全是 XSD `sequence`。顺序错了 Word 报「无法读取的内容」，用户看到的是文件损坏提示。

- 仓库里已有 `PROPERTY_ORDER` + `property()` 机制专门解决这个问题。**新增字段补进对应的顺序表**，不要 `appendChild`。
- 新增一类属性容器（如 `sectPr`、`tblPr`）时，为它建一张顺序表，照 `PROPERTY_ORDER` 的形状写。
- DrawingML 侧还有 `use="required"` 的属性（`wp:anchor` 的 `locked`、`layoutInCell`、`behindDoc`、`relativeHeight`、`allowOverlap`，`a:CT_Point2D` 的 `x`/`y`）。缺一个就是修复提示。
- 自测方式：把生成的容器子元素名抽出来，和规范 sequence 求一次「是否升序」，比肉眼看可靠。

## 4. 界面装饰元素绝不能流回文档

**重复了 3 次**：#8（编号标记，已修）、#24（脚注引用 `<sup>`）、#26（制表位前导符）。

编号、项目符号、脚注标记、制表位前导符、编辑标记（`¶` `→` `·`）都是**渲染产物**，不是文档内容。一旦它们进了 contenteditable 的可读文本，`readText()` → `flush()` → `setParagraphText()` 会把它们当字面文本写进 DOCX，而且**每编辑一次叠加一次**（#24 实测 `body` → `ody1` → `ody11`）。

统一做法（PR #8 已落地，照搬，不要另写）：

- 装饰节点带 `contentEditable="false"`、`data-docx-mark`、CSS `user-select: none; pointer-events: none`；
- `readText()` 显式跳过 `dataset.image` / `data-docx-mark` / `contentEditable === 'false'` 的节点；
- 回归测试断言：装饰开关切换前后 `getParagraphs()[i].text` 不变。

## 5. 读出来的格式对象必须能原样写回

**#26 一个 PR 里三处**（`borders.shadow`、`shading.color`、`tabs.leader`）。

根因：读侧无条件写 `shadow: undefined`，校验侧用 `'shadow' in value` 判存在——`undefined` 也算存在，于是校验拒绝自己刚产出的对象：

```ts
d.setParagraphBorders(0, d.getParagraphs()[0].borders);  // ✗ border.shadow must be boolean
```

「读出来 → 改一个字段 → 写回去」是格式工具栏、格式刷、Agent 批量改格式的基本模式。二选一并贯彻到所有读函数：

- 读侧用条件展开，不产出 `undefined` 键；或
- 校验侧统一用 `value.x !== undefined`。

每类格式对象都要有一条**往返测试**：`set(read())` 不抛错且结果等价。这一条测试能兜住以后新增的所有字段。

## 6. 读路径才是安全闸门

**#25 的安全缺陷**：写路径把 `javascript:` / `data:` / `vbscript:` / `file:` 拦得很干净，但 `isUnsafeHyperlink` 一看到 `w:anchor` 有值就返回 `false`，不检查 `r:id` 解析出的 target。构造一个带恶意 `Target` 的 `.docx` 完全不需要经过写路径。

威胁模型是「**文档内容是不可信输入**」，用户不会用你的 API 攻击自己：

- 安全判定放在**读路径**，以最终解析出的 target 为依据，与其他字段是否存在无关；
- 渲染前再判一次，`unsafe` 的不生成 `href` / 不加载 / 不执行；
- 外部关系目标（`TargetMode="External"`、`a:blip@r:link`）**不主动请求**，渲染占位；
- 测试必须**直接构造恶意 XML / rels**，写路径的测试不能替代。

## 7. 一个事务一次提交、一次 revision

**违反 4 次**：#9（`insertTable` 每个单元格提交一次，40×10 耗时 11.9 秒、revision +401）、#10（VML 图片上 no-op 仍 +2）、#24（`ensureNotePart` 在前置检查之前就提交了部件和关系，失败后留下垃圾）、#25（`insertHyperlink` 一次 +3，失败仍留下孤儿关系）。

- 公开方法内部无论做多少步，**只提交一次、revision 只 +1**；
- **所有前置检查做完再提交**，不要「先建部件再校验」——失败必须让文档和 revision 完全不变；
- 真正的 no-op 不要推进 revision（推进了就是骗调用方缓存失效）；
- `applyOperations` 的「全成功或全回滚」语义不能被子步骤的提交打破。

## 8. 不要在每次读取时重新解析部件

**重复 5 次**：#7（有效格式按 run 重算、`styles.xml` 解析三遍）、#8（每次读重新解析 numbering + styles + rels）、#10（渲染 O(images²)）、#26（`getSettings()` 每段落一次）、#19（读路径完全没缓存，`getParagraphs()` 每次 47ms）。

编辑器在每次变更后都 `render()`，每次 `blur` 都 `flush()`。「一次读取解析一个部件」在这种调用频率下直接变成卡顿。

- 解析结果按部件缓存，**失效键包含 `revision` 和该部件字节**；
- 缓存后要有一条测试断言「改了字节之后读到的是新值」（#10 的 dataUrl 缓存就差这条）；
- 公开的 `getPartDocument()` 承诺返回**分离的 DOM**，所以它要返回缓存的深拷贝，不能把内部 DOM 暴露出去。

## 9. `w:val="0"` 是显式关闭，缺失才是未设置

`w:b`、`w:i`、`w:u` 这类 toggle 属性，`w:val="0"` / `"false"` 表示**显式关闭**，会覆盖样式链上继承来的 `true`；元素缺失才表示「未设置、继续继承」。把两者混为一谈是样式解析最常见的错（#7 专门测过这条）。

## 10. 段落索引只有一套命名空间

**#24**（脚注正文段落报 `index: 2`，与正文 `p2` 撞号，Agent 改错段落）、**#23**（`collectSections` 自己数段落、跳过 `w:txbxContent`，导致 `endParagraph` 偏移、`insertSectionBreak` 对合法段落抛错）。

- 正文段落索引的唯一来源是 `descendants(body, 'p')` 的顺序，**不要另起一套遍历**；
- 页眉、页脚、脚注、尾注、文本框里的段落**不复用**正文索引命名空间，用独立寻址（或干脆不暴露 index），并在 README 写明。

## 11. 透明包裹元素要穿透

`w:sdt` / `w:sdtContent` / `w:customXml` 可以包在 block、`w:tr`、`w:tc`、run 任何一级上，模板类文档里非常常见。判定「直接父节点是不是 `w:body` / `w:tc`」的代码遇到它们一律失效（#16/#20 的段落被删空、#17/#21 的表格行不可见）。

- 用统一的 `isTransparentWordWrapper()` / `childrenThroughTransparent()`，不要每处写一套；
- 判断结构不变量时向上找**最近的非透明祖先**，不看直接父节点。

## 12. 破坏结构不变量时就地清空，不要抛错

`deleteParagraph` 的既有语义（README 也这么写）是「**保留**正文 / 单元格必要的空段落」——删掉内容、留下空 `w:p`，不是拒绝操作。#20 把 body 一侧改成抛错，结果：与 README 矛盾、破坏既有 API、而且对 `applyOperations` 是实质破坏（一个「清空最后一段」的操作会让整批回滚，模型无法预判）。

抛错只留给真正无法修复的情况：非法索引、以及既有的「拒绝隐式删除分节符」。

## 13. Schema 和运行时校验必须一致

#8（schema 允许 `numId: 0`，运行时拒绝）、#26（border 的 `required: []` 与运行时要求 style/size/space/color 矛盾，且段落侧和 run 侧两套形状不一致）。

`AGENT_OPERATION_SCHEMA` 是给模型看的契约。schema 放过而运行时拒绝，意味着一个「合法」的工具调用会在批次中途触发整批回滚，模型没有任何办法预判。新增 Agent 操作时，schema 和校验函数一起改，并加一条「schema 合法的输入运行时必接受」的测试。

## 14. 测试不要写成「恰好能过的形状」

真实发生过的盲区：

- #9 的单元格测试**全都没有 `w:tcPr`**，于是「替换段落插到 `tcPr` 之前」没被发现（带 `tcPr` 的单元格才是常态）；
- #10 的多图测试用了**两张图共用一个 `rId`**，恰好是唯一能过的形状，不同 rId 立刻抛错；
- #7 的测试和实现犯了**同一个拼写错误**（`firstCol` vs 规范的 `firstColumn`），互相掩盖；
- #24 的 `convertNote` 测试用小写 `/footnoteRef/` grep，漏掉了残留的 `rStyle`。

写测试时问一句：**这个用例是不是恰好避开了实现的薄弱处？** 至少覆盖带完整可选属性的形状、多实例且标识不同的形状、以及规范里的精确拼写。

## 15. rebase 时要重新实现，不是文本上解冲突

阶段 1 四个 PR 合并的实测教训：

- **在地基 PR 之前分叉的分支，会自带一份同名函数的旧副本。** PR #10 早于 #7 分叉，于是它有自己的 `readRun` / `readParagraph`，里面是一份旧的 run 属性解析器。正确做法是**丢掉自己那份、保留 `main` 的实现，只把自己的新能力缝进去**（#10 最终只需给 `readRun` 加一个 `paragraphIndex` 和可选的图片上下文）。两份都留会得到两条行为不一致的读路径。
- **不要留下平行的读路径。** #10 原本有 `paragraphsWithRelationships()`，`main` 有 `buildParagraphs()`，两者都能产出 `ParagraphInfo`。合并时把前者收敛成后者的代理，否则同一份数据有两种读法，缓存和新字段只会落在其中一条上。
- **冲突边界经常落在函数体中间**，两侧结尾都缺一个闭合括号，而共享后缀只能闭合其中一侧。机械拼接会得到 `'}' expected`。拼接前先看冲突之后的第一行是什么。
- **集成完成的判据是跑起来，不是能编译。** 合并 #10 时我们引入过一个回归：接上它的关系解析后，畸形 `Target`（如 `..`）会让 `getParagraphs()` 抛错，正好打破规则 1 的既有回归测试。是测试发现的，不是类型检查。**合并后必须跑全量测试**，并额外验证几个特性叠加的场景（例如「表格单元格内的段落同时有样式、编号和图片」）——这类场景在任何单个分支上都无法测到。
- 合并提交要写清**冲突是怎么解的**：哪一侧被丢弃、为什么、以及做了哪些签名改动。下一个 rebase 的人靠这个判断基线。

## 16. 校验写在公开 API 里，不能只写在编辑器层

**PR #48（剪贴板）一次性踩了四处。** `pasteClipboardFragment()` 是 `DocxDocument` 上的公开方法，但它把剪贴板片段里的值直接写进 XML，而仓库里现成的校验函数一个都没调用：

| 既有写入路径 | 结果 | 粘贴路径 | 结果 |
| --- | --- | --- | --- |
| `insertHyperlink({url:'javascript:alert(1)'})` | 拒绝（`isSafeHyperlinkUrl` 白名单） | `pasteClipboardFragment(…hyperlink:{url:'javascript:alert(1)'})` | **写进 `Target=… TargetMode="External"`** |
| `formatRun(0,0,{color:'nope!!'})` | 拒绝（`validateRunFormat`） | 同上带 `format:{color:'nope!!',fontSize:-5}` | **写出 `<w:color w:val="nope!!"/><w:sz w:val="-10"/>`** |
| `setParagraphText(0,'a\u0000b')` | 拒绝（`assertText`） | 同上带 `text:'a\u0000b'` | **裸 U+0000 进了 `document.xml`** |
| Agent 批次 | `maxItems: 1000` | 片段段落数 | **无上限**，5 万段落跑了 498 秒 |

当时编辑器层**确实**写了一套过滤（`isUnsafeHtmlHref()`），但它是自造的黑名单，比仓库已有的白名单弱（漏掉 `ms-msdt:` 这类 Word 里真正危险的协议），而且公开 API 根本不经过它。

具体要求：

- **校验属于文档层，不属于 UI 层。** 任何新的公开写入方法，第一件事是把入参过一遍既有的 `assertText` / `sanitizeText` / `validateRunFormat` / `assertHyperlinkInput` / `validateTabs`，而不是自己现写一套判断。发现已有校验不够用，就**改那一处**，让所有调用方一起受益。
- **不要自造平行规则。** 同一件事出现两套判断（一个白名单 + 一个黑名单），弱的那套迟早成为绕过口。#48 最终把编辑器里那个函数整个删掉，只留 `isSafeHyperlinkUrl()`。
- **剪贴板 / 拖放 / postMessage / 自定义 MIME 都不是可信信道。** `application/x-docx-view+json` 听起来像自有格式，但任何网页在自己的 `copy` 事件里 `setData()` 就能投放同名载荷。**「我们自己写出去的数据」在读回来时是不可信输入。**
- **新的批量入口要有自己的上限**，和既有的对齐（ZIP 50 MiB / 单部件 16 MiB / 批次 1000 操作 / 单文本 1,000,000 字符 / 表格 10,000 单元格），并写进 README 的限制段落。
- 校验失败时整个事务回滚，不要留下半个部件或半条关系。

## 17. 修订与其它叠加状态，测试必须做「第二次」

**PR #47（记录修订写入）的四个阻塞缺陷，全部只在第二次操作时暴露**，第一版 265 行新测试全绿，一个都没抓到：

- 同段落**两次**跟踪插入：`markRevision()` 复用了第一次那个已经装着内容的 `w:ins`，`insertBefore` 把它连同里面的 run 一起搬走 —— 期望 `XABCYDEF`，实得 **`ABCXYDEF`**，可见正文被改错；
- 跨超链接的段落删除（run 被分成两组）：删除内容顺序变成 `BBB, AAA, CCC`；
- **两次**跟踪改格式：`rPrChange` 里的原始格式快照被第二次覆盖成中间态，拒绝后回不到原样；
- 跟踪删除段落：不含 `w:del` 后代的子元素被整体清掉，`bookmarkStart` / `commentRangeStart` 直接消失且不可恢复。

根因是同一个：`orderedRevisionChild()` 的「存在就复用」对**属性容器**（`trPr` / `pPr/rPr` / `tcPr`，同名子元素只能有一个）是对的，对**行内包装**（`w:p` / `w:hyperlink` 里的 `w:ins` / `w:del`，可以有任意多个）是错的。修法是拆成两个函数：`createRevisionWrapper()` 每次新建，`markRevision()` 才复用。

具体要求：

- **每个修订写入至少测一次「在已有修订标记之上再做一次」。** 一次操作的路径太顺，测不出复用、覆盖、搬移这三类问题。
- **删除类修订的判据是「放回去等于原文」**，不只是「可见文本对」。#47 有一个中间状态可见文本完全正确，但删除内容已经从段首挪到了段尾——等接受/拒绝功能上来就会还原成另一份文档。写断言时按文档顺序把 `delText` 拼回去和原文比。
- **格式类修订的快照必须始终是最初的未修订格式**，重复修改只刷新 `w:id` / `w:author` / `w:date`，不重写快照。
- **范围标记不是内容**：`bookmarkStart` / `bookmarkEnd` / `commentRangeStart` / `commentRangeEnd` / `permStart` 在跟踪删除时原样保留，不要因为「它们里面没有 run」就清掉。
- **`w:pPr/w:rPr` 是 `CT_ParaRPr`，不是 `CT_RPr`**：`ins` / `del` / `moveFrom` / `moveTo` 必须排在最前（`EG_ParaRPrTrackChanges` 在 `EG_RPrBase` 之前），而 `CT_RPr` 里**根本不允许**这几个。两者都叫 `rPr`，按父元素名索引顺序表会串。
- 同理适用于其它叠加状态：批注、编号覆盖、样式继承链——凡是「第二次写入要叠在第一次之上」的地方，都按这条写测试。

## 18. 文档的段落级冲突，和代码冲突不是一回事

第 15 条讲的是**代码**冲突。README 这类文档还有另一种冲突形态，`git` 帮不上忙，而且**机械合并会静默丢内容**。

**实际发生过 4 次**：

| | 情形 | 后果 |
| --- | --- | --- |
| 阶段 1 合并 | 用「过滤掉含『尚未实现』的句子」来合并 | 连带删掉了表格样式条件格式那句 |
| 阶段 1 合并 | 用正则从 README 提取操作清单 | 漏掉 10 个操作（超链接、制表位相关） |
| PR #54 | 两侧各改了**同一个段落的不同句子** | 冲突块为「整段 vs 整段」，机械取一侧就丢另一侧的说明 |
| PR #60 | 两侧各向**同一段落追加**了不同句子 | 共有中段与结尾完全相同，正确解法是四段拼接 |

后两种是同一类：冲突块看起来是「HEAD 整段 vs MAIN 整段」，但真实差异只是**段落内的一两句**。取任一侧都会丢掉对方新增的说明，而且**编译和测试都不会报错**——README 不参与构建。

具体要求：

- **解这类冲突时按句子对齐，不要按行取舍。** 先找出两侧的公共前缀/后缀，把各自独有的句子都留下，再按语义排序拼回去。
- **解完必须校验没丢句子**：把冲突前两侧的句子逐条在结果里搜一遍。这一步是机械的，值得真的做——上面两次丢内容都是「看起来对了」。
- **README 的操作清单、能力表这类列举，永远从代码重新生成，不要手工增删**。`AGENT_OPERATION_SCHEMA` 就是操作清单的唯一事实来源，按 schema 顺序生成后断言与 schema 完全一致（数量、内容、顺序）。
- **写文档时就要防冲突**：一个段落只讲一件事。一旦某段塞进十几条互不相关的说明，任何两个并行 PR 碰它都必然冲突——README 里那段「索引与作用域 / 修订 / 比较 / 属性与保护」混在一起的 6 行长段落，就是因此被拆成了分组条目。**新增说明请加进对应的分组，不要继续往长段落里追加。**

## 19. `examples/` 里的引用没有任何东西兜底

`src/` 有 `tsc` 兜底，`examples/` 只有一半——**TS 里的 id 字面量与 HTML 里的 `id` 之间不参与类型检查，也不在任何单元测试的覆盖路径上**。这是本仓库唯一一条「编译干净、测试全绿、功能完全不可用」的缝。

**实际发生过一次，而且持续了 33 个 PR**：

`examples/main.ts` 顶层执行 `element('add-footnote').addEventListener(…)`，而 `element()` 的实现是找不到就 `throw`：

```ts
function element<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`找不到界面元素：${id}`);
  return node as T;
}
```

`examples/index.html` 里**从来没有**这个 id（`git log -S'add-footnote' -- examples/index.html` 只有修复那一个提交）。于是自「脚注尾注支持」合并起，**示例应用一加载就抛错、整个初始化中断**，此后 33 个 PR、611 个测试全绿，而演示页根本打不开。

没人发现的原因很简单：**没有任何测试加载 `examples/`**。

具体要求：

- **`test/ribbon.test.mjs` 里那条结构性断言必须保持通过**：`examples/main.ts` 中所有 `element('…')` 字面量与 `commandControls` 表的 `elementId`，都要在 `examples/index.html` 里有对应 `id`。它在 `main` 上会直接报出上面那两个缺失 id，注入一个不存在的引用也会如期失败——**有牙，别绕过它**。
- **界面元素优先静态写在 `index.html` 里**（能被上面那条断言覆盖）。确实需要运行时动态创建的（菜单项、列表项），**不要用 `element()` 去找，用局部变量持有引用**——不要为了「让断言过」往 HTML 里塞占位元素，那是绕过而非遵守。
- 改 `examples/` 时，把「演示页能否正常初始化」当成一项验收，而不是假定测试覆盖了它。

## 20. 类型的重复定义比逻辑的重复更难自查

第 15 条讲的是逻辑上的平行路径。**类型层面的重复更隐蔽**：逻辑重复迟早跑出不同结果，类型重复只在「有人用到那个多出来的字段」时才暴露。

**实际发生过一次**：#63（命令注册表）与 #65（`getTableCellAt`）并行开发，各自定义了 `TableCellLocation`：

```ts
// src/types.ts（权威）        … nested: boolean;   ← 有
// examples/commands.ts        …                    ← 缺
```

`CommandContext.table` 用的是 `examples/` 那份，于是**命令的 `enabled` 拿不到 `nested`**。两份在合并当天就已经不一致，但**没有任何测试会失败**——直到 #67 真的需要用 `nested` 去禁用嵌套表格下的按钮，缺失才变成阻塞。

具体要求：

- **`examples/` 里出现与 `src/types.ts` 同名的 `interface` / `type`，直接当缺陷处理**，从库里 import。
- 并行开发多个 issue 时，新增公开类型前先 `grep -rn "interface <名字>" src/ examples/` 一遍。
- 同一个概念只能有一处定义——这条对类型和对函数同样成立。

---

## 21. 基准测量的设计错误，会把真实缺陷藏起来

`test/perf.mjs` 里 `batched 1000-paragraph updates scale near-linearly` 这条断言存在了很久，**一直通过**，而且在 `main` 上间歇性失败（#87）。排查它的时候才发现：**被它断言「近似线性」的那段代码根本不是线性的**——按下标寻址的编辑操作是 O(文档规模)，4000 段文档上单次 `setParagraphText` 要 12.9 ms（#92）。

一条写错的基准测试比没有测试更糟：它给出了「已经验过」的假象。五条具体要求：

**(1) 两个测量必须配对交替，每轮还要换序。**
`5 轮 A → 5 轮 B` 是错的。两段测量落在不同时间窗口里，任何跨窗口漂移都会被整份记进比值。实测同一份代码，`main` 上的比值从 2.72 漂到 4.03（阈值 4.00），因为 JIT 预热让先跑的那侧付了账。换序是因为堆状态也有方向性：干净堆上先跑的那侧更快。

比值取**「每轮配对比值的中位数」**，不是「两个中位数的比值」——只有前者抵消轮间波动。

**(2) 计时区间里不能有 setup。**
播种文档、构造操作数组都算 setup。原来的 `measureSingleInsert(200, 200)` 把 200 次播种插入算进了耗时，等于在测 400 次插入再除以 200。

**(3) 跨度要大到能分辨假设。**
500 → 1000 时线性期望是 2、二次期望是 4，而噪声区间横跨 4——**这个跨度根本无法区分线性和二次**，所以它从来没有验证过任何东西。要么把跨度拉开（250 → 1000 是 4 vs 16），要么换成「等量工作」的设计：同一份文档，1000 次操作一批 vs 拆成两批 500，工作量相同，线性 ≈ 1、二次 ≈ 2。

**(4) 被测区间要长到高出计时噪声。**
`getTableCellAt` 那条测 100 次缓存命中，预热充分后只要 0.03 ms，比值变成 `[7.42, 5.33, 6.81, 0.82, 1.11]` 的纯噪声。查 20 000 次才测得出来。

**(5) 跨规模的比较必须每档单独开进程 —— 配对交替救不了它。**
第 (1) 条的配对交替只能抵消**时间上**的漂移。两侧**工作集大小不同**时（一侧 500 段文档、另一侧 2000 段），小的那侧会因为 JIT 越跑越快、又会被大的那侧留下的垃圾拖慢，这种不对称配对抵消不掉。

我按第 (1) 条改完之后，仍然用同进程配对去比 500 段和 2000 段，结果配对比值在 **3.8 到 11.0 之间跳，5 次里有 1 次误报**；改成每档开独立子进程（`test/perf-worker.mjs`），同一个量测出来稳定在 **2.8** 上下。

同一个坑我还栽过一次：靠同进程测量先得出「是 O(N²)」，隔离后又得出「是线性的」——**两次都错**。正确做法是固定操作次数、只变文档规模、每档独立进程，才看出单次操作成本随文档规模线性增长。

判断标准很简单：**两侧的内存/数据规模不同 → 开子进程；规模相同、只是做法不同 → 同进程配对交替就够。**

**如果一条性能断言的阈值刚好卡在实测值边缘，那它测的很可能不是你以为的东西。** 先弄清真值分布，再定阈值；不要为了让它变绿而放宽阈值。

---

## 通用底线

- `npm run check` 与 `npm test` 必须通过；每个修复配一条回归测试。
- **无损优先**：未理解的 XML 原样保留，往返不丢部件 / 属性 / 未知元素。
- 不破坏现有 API 签名与行为；新能力走新方法或可选参数。
- 核心逻辑不依赖浏览器全局 `document`（用 `@xmldom/xmldom`）；只有 `src/editor.ts` 可以用真实 DOM。
- 不新增运行时依赖；确有必要在 PR 描述里说明理由。
- 安全边界不放松：不执行文档内脚本、不把文档 XML 当 HTML、不主动请求外部 URL，保留 ZIP / XML / 文本长度限制与路径穿越校验。性能优化**不是**削减校验的理由（#19 把良构性闸门拆了，结果能提交出让 `load()` 拒绝的文档）。
- 畸形输入**降级**，不要让读取方法抛错：悬空 `r:id`、缺失部件、未知枚举值、`0` 或负的尺寸、声明数量与实际不符。
- 多 PR 并行时**只改自己范围内的文件**。发现别人范围内的 bug，在 PR 描述里指出，不要顺手改——那些 PR 正在被独立 review，顺手改会造成连环冲突。
