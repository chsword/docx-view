export interface RunFormat {
  style?: string | null;
  bold?: boolean | null;
  italic?: boolean | null;
  hidden?: boolean | null;
  webHidden?: boolean | null;
  emphasisMark?: 'dot' | 'comma' | 'circle' | 'underDot' | 'none' | null;
  underline?: boolean | null;
  underlineStyle?: string | null;
  underlineColor?: string | null;
  fontSize?: number | null;
  fontFamily?: string | null;
  fontFamilyEastAsia?: string | null;
  color?: string | null;
  strike?: boolean | null;
  doubleStrike?: boolean | null;
  rtl?: boolean | null;
  complexScript?: boolean | null;
  /**
   * `w:bCs` / `w:iCs`：**复杂文种**字符的粗体 / 斜体。ECMA-376 把一个 run 的字符分成两类：
   * `w:b` / `w:i` 管非复杂文种，`w:bCs` / `w:iCs` 管复杂文种（阿拉伯文、希伯来文、泰文……）。
   * 渲染时 run 被当成复杂文种（`complexScript` 或 `rtl` 为真）就用这两个，否则用 `bold` / `italic`。
   * **没有按 Unicode 文种逐字符判断**：含阿拉伯字符却没标 `w:cs` / `w:rtl` 的 run 仍按 `bold` 画。
   */
  boldComplexScript?: boolean | null;
  italicComplexScript?: boolean | null;
  /** `w:noProof`：不检查拼写与语法。渲染落在 `spellcheck="false"` 上，交给浏览器。 */
  noProof?: boolean | null;
  /**
   * `w:snapToGrid`（run 级）：用文档网格的**字符间距**。**不参与渲染**——这里的 `w:docGrid` 只做了
   * 行高吸附，没有横向字符格（同 `ParagraphFormat.adjustRightInd`）。
   */
  snapToGrid?: boolean | null;
  /**
   * `w:specVanish`：「特殊隐藏」的段落标记，Word 用它做样式分隔符（两段并成一行、各用各的样式，
   * 目录条目常见）。只做读写保真；渲染成一行要把两段合并排，这里不做。
   */
  specVanish?: boolean | null;
  verticalAlign?: 'baseline' | 'subscript' | 'superscript' | null;
  smallCaps?: boolean | null;
  allCaps?: boolean | null;
  highlight?: string | null;
  characterSpacing?: number | null;
  position?: number | null;
  eastAsianLayout?: EastAsianLayout | null;
  characterScale?: number | null;
  kerning?: number | null;
  fitTextWidth?: number | null;
  textEffect?: string | null;
  textOutline?: boolean | null;
  textShadow?: boolean | null;
  emboss?: boolean | null;
  imprint?: boolean | null;
  border?: BorderSide | null;
  shading?: Shading | null;
}

export interface TabStop {
  position: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num';
  leader?: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
}

export interface BorderSide {
  style: string;
  size: number;
  space: number;
  color: string;
  shadow?: boolean;
}

export interface Shading {
  pattern: string;
  fill: string;
  color?: string;
}

export interface ParagraphFormat {
  alignment?: 'left' | 'center' | 'right' | 'both' | 'distribute' | null;
  /** `w:framePr`；首字下沉与段落定位都在这里。 */
  frame?: ParagraphFrame | null;
  style?: string | null;
  indentLeft?: number | null;
  indentRight?: number | null;
  indentFirstLine?: number | null;
  indentHanging?: number | null;
  spacingBefore?: number | null;
  spacingAfter?: number | null;
  spacingBeforeLines?: number | null;
  spacingAfterLines?: number | null;
  spacingBeforeAuto?: boolean | null;
  spacingAfterAuto?: boolean | null;
  contextualSpacing?: boolean | null;
  mirrorIndents?: boolean | null;
  lineSpacing?: number | null;
  lineSpacingRule?: 'auto' | 'atLeast' | 'exact' | null;
  keepNext?: boolean | null;
  keepLines?: boolean | null;
  pageBreakBefore?: boolean | null;
  widowControl?: boolean | null;
  suppressLineNumbers?: boolean | null;
  suppressAutoHyphens?: boolean | null;
  kinsoku?: boolean | null;
  wordWrap?: boolean | null;
  overflowPunct?: boolean | null;
  topLinePunct?: boolean | null;
  autoSpaceDE?: boolean | null;
  autoSpaceDN?: boolean | null;
  bidi?: boolean | null;
  textDirection?: string | null;
  /**
   * `w:textAlignment`：一行里字符的**垂直**对齐（字号不一的字符怎么在行内对齐），
   * 不是 `w:jc` 的左右对齐。渲染落在 run 的 `vertical-align` 上。
   */
  textAlignment?: 'auto' | 'baseline' | 'bottom' | 'center' | 'top' | null;
  /**
   * `w:adjustRightInd`：用文档网格时自动调整右缩进，使行正好容纳整数个网格字符。
   * **不参与渲染**：这里的 `w:docGrid` 只做了行高吸附（`snapLineHeightPx`），没有横向的
   * 字符格，没有可调的那个量。
   */
  adjustRightInd?: boolean | null;
  /**
   * `w:suppressOverlap`：禁止这段（`w:framePr` 的文本框）与其他框重叠。和表格的
   * `w:tblOverlap` 是同一回事：环绕用浏览器 `float`，浮动块本来就不互相重叠，所以这条
   * 天然成立，**不参与渲染**。
   */
  suppressOverlap?: boolean | null;
  /**
   * `w:textboxTightWrap`：允许周围段落按文本框**内容的实际行宽**紧密绕排。
   * **不参与渲染**：那需要逐行量文本框里的内容，而这里的排除区是个矩形——与 `w:framePr`、
   * `w:tblpPr` 一样的 `square` 简化。
   */
  textboxTightWrap?: 'none' | 'allLines' | 'firstAndLastLine' | 'firstLineOnly' | 'lastLineOnly' | null;
  /**
   * `w:divId`：这一段属于 `webSettings.xml` 里哪个 HTML `<div>`（见 `getWebDivs()`）。div 的
   * 左右边距累加在段落自己的缩进之外渲染；悬空的 id 按没有 div 处理。
   */
  divId?: number | null;
  /**
   * `w:snapToGrid`（段落级）：这段的行是否吸附到文档网格的行距（`w:docGrid`）。默认开；
   * 显式关掉的段落在分页与渲染里都**不做**行高吸附——这是 `snapLineHeightPx` 的开关。
   */
  snapToGrid?: boolean | null;
  outlineLevel?: number | null;
  tabs?: TabStop[] | null;
  borders?: Partial<Record<'top' | 'left' | 'bottom' | 'right' | 'between' | 'bar', BorderSide>> | null;
  shading?: Shading | null;
}

/** `webSettings.xml` 的 `w:divs` 里的一个 div，见 `getWebDivs()`。边距单位是缇。 */
export interface WebDivInfo {
  id: number;
  /** 外层 div 的 id（`w:divsChild` 嵌套）；顶层 div 没有。 */
  parentId?: number;
  /** HTML `<blockquote>`。 */
  blockQuote: boolean;
  /** HTML `<body>` 本身。 */
  bodyDiv: boolean;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  borders?: Partial<Record<'top' | 'left' | 'bottom' | 'right', BorderFormat>>;
}

export interface StyleInfo {
  id: string;
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  basedOn?: string;
  next?: string;
  link?: string;
  aliases?: string[];
  isDefault?: boolean;
  uiPriority?: number;
  quickFormat?: boolean;
  paragraph?: ParagraphFormat;
  run?: RunFormat;
}

/**
 * `updateStyle()` 的补丁：只改给出的字段，没给的原样保留（包括读模型不认识的子元素）。
 * `basedOn` / `next` / `link` / `aliases` / `uiPriority` 传 `null` 清除；`paragraph` / `run`
 * 与 `formatParagraph()` / `formatRun()` 同一语义——字段为 `null` 即从样式里删掉该属性。
 */
export interface StylePatch {
  name?: string;
  basedOn?: string | null;
  next?: string | null;
  link?: string | null;
  aliases?: string[] | null;
  uiPriority?: number | null;
  quickFormat?: boolean;
  paragraph?: ParagraphFormat;
  run?: RunFormat;
}

export interface OutlineNode {
  paragraph: number;
  level: number;
  text: string;
  styleId?: string;
  children: OutlineNode[];
}

export interface NumberingLevelDefinition {
  level: number;
  start?: number;
  format: string;
  text: string;
  justification?: string;
  suffix: 'tab' | 'space' | 'nothing';
  isLegal?: boolean;
  restart?: number;
  paragraphStyle?: string;
  indentLeft?: number;
  indentHanging?: number;
  runFormat?: RunFormat;
  image?: ImageInfo;
}

export interface NumberingDefinition {
  numId: number;
  abstractNumId: number;
  multiLevelType?: string;
  nsid?: string;
  tmpl?: string;
  styleLink?: string;
  numStyleLink?: string;
  levels: NumberingLevelDefinition[];
}

export interface NumberingInfo {
  numId: number;
  level: number;
  format: string;
  text: string;
  isBullet: boolean;
  indentLeft?: number;
  indentHanging?: number;
  suffix: 'tab' | 'space' | 'nothing';
  runFormat?: RunFormat;
  image?: ImageInfo;
}

/**
 * `w:eastAsianLayout` —— Word 的「双行合一」与「纵中横」。两者都只是显示方式,文本本身不变。
 */
export interface EastAsianLayout {
  id?: number;
  /** 双行合一：把这一段文字压成上下两行，占一行的高度。 */
  combine?: boolean;
  /** 双行合一两侧的括号样式。 */
  combineBrackets?: 'none' | 'round' | 'square' | 'angle' | 'curly';
  /** 纵中横：竖排文本里把这一段横过来排。横排视图下 Word 也不显示差别。 */
  vert?: boolean;
  /** 纵中横时压缩字宽以适应行宽。 */
  vertCompress?: boolean;
}

/**
 * `w:ruby` —— 注音（拼音 / 振假名）。注音文字排在基字符上方，**不进段落正文**：
 * Word 的阅读顺序只含基字符，所以 `RunInfo.text` 是基字符，注音在这里单独给出。
 */
export interface RubyInfo {
  /** 注音文字（`w:rt`）。 */
  text: string;
  /** 基字符（`w:rubyBase`），与所在 run 的 `text` 相同。 */
  base: string;
  align?: 'center' | 'distributeLetter' | 'distributeSpace' | 'left' | 'right' | 'rightVertical';
  /** 注音字号，半磅。 */
  sizeHalfPoints?: number;
  /** 注音相对基线抬升，半磅。 */
  raiseHalfPoints?: number;
  /** 基字符字号，半磅。 */
  baseSizeHalfPoints?: number;
  /** `w:lid`，注音所用语言。 */
  language?: string;
}

export interface RunInfo extends RunFormat {
  index: number;
  text: string;
  /** 注音；注音文字不算正文，所以不在 `text` 里。 */
  ruby?: RubyInfo;
  revisions?: RevisionMark[];
  effective?: RunFormat;
  hyperlink?: { url?: string; anchor?: string; tooltip?: string; unsafe: boolean };
  image?: ImageInfo;
  images?: ImageInfo[];
  noteReference?: { kind: 'footnote' | 'endnote'; id: number; number: number; marker: string };
  /** Field index is scoped to the source part; use kind/instruction across parts. */
  field?: { index: number; role: 'instruction' | 'result'; kind?: FieldKind; instruction?: string };
}

export interface MathMlNode {
  tag: string;
  attrs?: Record<string, string>;
  children?: MathMlNode[];
  text?: string;
  /**
   * 读出来时记下原本的 OMML 元素名。MathML 的一个标签对应多个 OMML 元素——mover 可能来自
   * bar / acc / groupChr / limUpp，mrow 可能来自 d / func / box / nary，mtable 可能来自
   * m / eqArr——写回时只看标签必然猜错一部分，所以把来处带上。手工构造的节点没有这个字段，
   * 写入侧照旧按标签猜。
   */
  source?: string;
}

export function assertText(text: unknown, name = 'text'): asserts text is string {
  if (typeof text !== 'string' || text.length > 1_000_000 ||
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/u.test(text) ||
      /[\ud800-\udfff]/u.test(text)) {
    throw new Error(`${name} must be valid XML text of at most 1,000,000 characters.`);
  }
}

export type MathSource =
  | { mathMl: MathMlNode }
  | { linear: string };

export interface MathInfo {
  runOffset: number;
  display: 'inline' | 'block';
  linear: string;
  mathMl: MathMlNode;
  truncated?: boolean;
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  visibleText?: string;
  runs: RunInfo[];
  paragraphRevision?: RevisionMark;
  effective?: ParagraphFormat;
  numbering?: NumberingInfo;
  images: ImageInfo[];
  math?: MathInfo[];
}

export type FieldKind =
  | 'SEQ' | 'DATE' | 'TIME' | 'CREATEDATE' | 'SAVEDATE' | 'PRINTDATE'
  | 'AUTHOR' | 'TITLE' | 'SUBJECT' | 'KEYWORDS' | 'COMMENTS' | 'LASTSAVEDBY'
  | 'DOCPROPERTY' | 'FILENAME' | 'REF' | 'PAGE' | 'NUMPAGES' | 'PAGEREF'
  | 'TOC' | 'INDEX' | 'INCLUDETEXT' | 'INCLUDEPICTURE' | 'LINK' | 'DDE' | 'DDEAUTO'
  | 'MACROBUTTON' | 'GOTOBUTTON' | 'FILLIN' | 'ASK' | 'DATABASE' | 'AUTOTEXT'
  | 'AUTOTEXTLIST' | 'HYPERLINK' | 'IF' | 'MERGEFIELD' | 'FORMTEXT' | 'FORMCHECKBOX'
  | 'FORMDROPDOWN' | 'EQ' | 'unknown';

export interface FieldSwitch {
  name: string;
  value?: string;
}

/**
 * `EQ`（公式域）指令的结构。Word 的「合并字符」与「带圈字符」都是同一个机制：
 * `\o` 是重叠排版（overstrike），把括号里各部分叠印在一处，各部分再用 `\s\up N` /
 * `\s\do N` 上下位移——所以「合并字符」是两组文字一上一下叠出来的，「带圈字符」是
 * 一个圈和一个字叠出来的。按这个机制实现，两个功能都落地，不必各自特判。
 */
/**
 * `w:framePr` —— 段落文本框。它盖着 Word 里两个看起来无关的功能：
 *
 * - **首字下沉**：`dropCap` 为 `drop`（落在正文里）或 `margin`（落到页边距外），下沉的那个字
 *   自成一段，Word 同时把它的 `w:sz` 调大到正好跨 `lines` 行。
 * - **段落定位**：DrawingML 之前的浮动做法——整段按 `x` / `y` 或 `xAlign` / `yAlign` 定位，
 *   正文按 `wrap` 绕着它排。
 *
 * 两者在排版上是同一件事：这一段脱离正常流，给后面的内容留出一块排除区。
 */
export interface ParagraphFrame {
  dropCap?: 'none' | 'drop' | 'margin';
  /** 下沉字跨几行。 */
  lines?: number;
  widthTwips?: number;
  heightTwips?: number;
  heightRule?: 'auto' | 'exact' | 'atLeast';
  wrap?: 'around' | 'auto' | 'none' | 'notBeside' | 'through' | 'tight';
  verticalAnchor?: 'margin' | 'page' | 'text';
  horizontalAnchor?: 'margin' | 'page' | 'text';
  xTwips?: number;
  yTwips?: number;
  xAlign?: 'center' | 'inside' | 'left' | 'outside' | 'right';
  yAlign?: 'bottom' | 'center' | 'inline' | 'inside' | 'outside' | 'top';
  horizontalSpaceTwips?: number;
  verticalSpaceTwips?: number;
  anchorLock?: boolean;
}

export interface EquationNode {
  /** 开关名，如 `o`（重叠）、`s`（升降）、`f`（分数）；纯文字节点没有。 */
  switch?: string;
  /** 开关后紧跟的子开关，如 `\o\ac` 的 `ac`（居中对齐）、`\s\up` 的 `up`。 */
  options?: string[];
  /** `\up` / `\do` 的位移量，单位磅；`do` 记为负数。 */
  raisePoints?: number;
  /** 带数值的开关与子开关各自的数值，如 `\a\co2\hs3` → `{ co: 2, hs: 3 }`。 */
  values?: Record<string, number>;
  /** 带字符参数的子开关，如 `\b\lc\{` → `{ lc: '{' }`；`\i\fc\∮` → `{ fc: '∮' }`。 */
  characters?: Record<string, string>;
  /** 括号里以分隔符隔开的各部分。 */
  parts?: EquationNode[];
  /** 纯文字。 */
  text?: string;
}

export interface FieldInfo {
  index: number;
  paragraph: number;
  runs: number[];
  resultRuns: number[];
  form: 'simple' | 'complex';
  kind: FieldKind;
  instruction: string;
  argument?: string;
  /** `EQ` 域的指令结构；其他域没有。`EQ` 的括号语法不是开关语法，所以 `switches` 对它为空。 */
  equation?: EquationNode;
  mergeFieldName?: string;
  formField?: {
    name?: string;
    enabled?: boolean;
    helpText?: string;
    statusText?: string;
    entryMacro?: string;
    exitMacro?: string;
    kind: 'text' | 'checkBox' | 'dropDown';
    text?: { default?: string; maxLength?: number; format?: string; type?: string };
    checkBox?: { default?: boolean; checked?: boolean; sizeAuto?: boolean; sizePt?: number };
    dropDown?: { default?: number; result?: number; entries: string[] };
  };
  switches: FieldSwitch[];
  result: string;
  requiresPagination: boolean;
  evaluable: boolean;
  locked: boolean;
  dirty: boolean;
  nestedIn?: number;
}

export interface PaginationInfo {
  pageCount: number;
  pageOfParagraph: (paragraph: number) => number | undefined;
  numberOfPage: (pageIndex: number) => number;
}

export interface BorderFormat {
  style?: string;
  size?: number;
  space?: number;
  color?: string;
  none?: boolean;
}

export interface BordersFormat {
  top?: BorderFormat;
  right?: BorderFormat;
  bottom?: BorderFormat;
  left?: BorderFormat;
  insideH?: BorderFormat;
  insideV?: BorderFormat;
}

/**
 * 单元格边框。比 `w:tblBorders` 多两条对角线，而且**只有** `w:tcBorders` 有——所以类型上
 * 分开，不要把它塞进 `BordersFormat`：写进 `tblBorders` 的对角线是 schema 不合法的。
 */
export interface CellBordersFormat extends BordersFormat {
  /** `w:tl2br`：左上到右下。 */
  tl2br?: BorderFormat;
  /** `w:tr2bl`：右上到左下。 */
  tr2bl?: BorderFormat;
}

export interface WidthFormat {
  type: 'auto' | 'dxa' | 'pct';
  value: number;
}

export interface ShadingFormat {
  fill?: string;
  color?: string;
  value?: string;
}

export interface MarginFormat {
  top?: WidthFormat;
  right?: WidthFormat;
  bottom?: WidthFormat;
  left?: WidthFormat;
}

/**
 * `w:tblpPr` —— 浮动表格定位。它是表格版的 `w:framePr`：表格脱离正常流，正文绕着它排，
 * 所以排版上走的是同一条路（给后面的内容留出一块排除区，环绕交给浏览器 `float`）。
 *
 * `w:tblpPr` 上没有 `w:wrap`：浮动表格在 Word 里一定是绕排的，这正是它的用途；
 * `w:tblOverlap` 管的是能否与**其他浮动对象**重叠，不是正文是否绕排。
 */
export interface TableFloatingPosition {
  /** 与周围正文的间距，单位缇。 */
  leftFromText?: number;
  rightFromText?: number;
  topFromText?: number;
  bottomFromText?: number;
  verticalAnchor?: 'margin' | 'page' | 'text';
  horizontalAnchor?: 'margin' | 'page' | 'text';
  xSpec?: 'center' | 'inside' | 'left' | 'outside' | 'right';
  /** 绝对横坐标，单位缇。 */
  x?: number;
  ySpec?: 'bottom' | 'center' | 'inside' | 'inline' | 'outside' | 'top';
  /** 绝对纵坐标，单位缇。 */
  y?: number;
}

export interface TableFormat {
  width?: WidthFormat;
  alignment?: 'left' | 'center' | 'right';
  indent?: number;
  borders?: BordersFormat;
  shading?: ShadingFormat;
  cellMargin?: MarginFormat;
  layout?: 'fixed' | 'autofit';
  /**
   * `w:tblCellSpacing`：单元格之间、以及单元格与表格边缘之间的间距。ECMA-376 把这两处
   * 说成同一个值，所以它和 CSS `border-spacing` 是一对一的。
   */
  cellSpacing?: WidthFormat;
  /**
   * `w:tblOverlap`：这张浮动表格能否和**其他浮动对象**重叠。渲染用的是浏览器 `float`，
   * 而浮动块本来就不互相重叠，所以 `never` 天然成立、`overlap` 没法实现——读写忠实保留，
   * 排版上不起作用。
   */
  overlap?: 'never' | 'overlap';
  style?: string;
  look?: string;
  caption?: string;
  description?: string;
  bidiVisual?: boolean | null;
  /** `w:tblpPr`；有它就是浮动表格。 */
  floatingPosition?: TableFloatingPosition | null;
}

/**
 * `w:tblPrEx`：表格属性的**行级例外**。它不在 `w:trPr` 里，而是 `w:tr` 的第一个子元素，
 * 内容是 `w:tblPr` 的一个子集，含义是「这一行按这些值覆盖表格自己的 `tblPr`」——Word 合并
 * 两张格式不同的表格时就会产出它（下半张的边框、底纹、边距落到各行的 `tblPrEx` 上）。
 */
export type TableException = Pick<TableFormat,
  'width' | 'alignment' | 'indent' | 'borders' | 'shading' | 'cellMargin' | 'layout' | 'cellSpacing' | 'look'>;

export interface RowFormat {
  /** `w:trPr/w:divId`：这一行属于哪个 HTML `<div>`，同 `ParagraphFormat.divId`。只读写，不渲染。 */
  divId?: number | null;
  /** 见 `TableException`。写入时 `null` 删除整个 `w:tblPrEx`。 */
  tableException?: TableException | null;
  height?: { value: number; rule?: 'atLeast' | 'exact' };
  /** `w:gridBefore`：这一行开头跳过的网格列数，视觉上就是整行缩进。 */
  gridBefore?: number;
  /** `w:wBefore`：开头跳过那块空间的宽度。 */
  widthBefore?: WidthFormat;
  /** `w:gridAfter`：这一行末尾跳过的网格列数。 */
  gridAfter?: number;
  /** `w:wAfter`：末尾跳过那块空间的宽度。 */
  widthAfter?: WidthFormat;
  /** `w:tblCellSpacing`：这一行的单元格间距，覆盖表格级的那个。 */
  cellSpacing?: WidthFormat;
  cantSplit?: boolean;
  header?: boolean;
  alignment?: 'left' | 'center' | 'right';
  deleted?: boolean;
  inserted?: boolean;
  revision?: { author?: string; date?: string };
}

export interface CellFormat {
  width?: WidthFormat;
  borders?: CellBordersFormat;
  shading?: ShadingFormat;
  margin?: MarginFormat;
  verticalAlign?: 'top' | 'center' | 'bottom';
  textDirection?: string;
  noWrap?: boolean;
  /**
   * `w:tcFitText`：把这一格的文字压缩 / 拉伸到正好占满单元格宽度。和 run 上的
   * `w:fitText`（`RunFormat.fitTextWidth`）是同一件事的两个层级，两者都只做读写保真，
   * 不参与渲染——那需要按实测文本宽度反算字间距。
   */
  fitText?: boolean;
  hideMark?: boolean;
  hMerge?: 'restart' | 'continue';
  vMerge?: 'restart' | 'continue';
}

/**
 * 表格样式的条件格式名（`w:tblStylePr/@w:type` 的取值，去掉 `wholeTable`）。数组里的顺序就是
 * 套用顺序，后面的覆盖前面的：带状 → 首末列 → 首末行 → 四个角。
 */
export type TableConditionName = 'firstRow' | 'lastRow' | 'firstCol' | 'lastCol'
  | 'band1Horz' | 'band2Horz' | 'band1Vert' | 'band2Vert'
  | 'nwCell' | 'neCell' | 'swCell' | 'seCell';

export interface TableCellInfo {
  blocks: DocumentBlock[];
  colSpan: number;
  rowSpan: number;
  isMergeContinuation: boolean;
  /** 单元格自己的 `w:tcPr`，不含表格样式带来的部分。 */
  format?: CellFormat;
  /**
   * 算上表格样式之后的单元格格式：样式自身的 `tcPr`、各条件格式（`firstRow` / 带状 / 角单元格
   * 等）的 `tcPr`，最后叠上单元格自己的直接格式。边框与边距按**边**合并——
   * 「整表定四边 + firstRow 只定下边框」是最常见的组合，整块替换会把其余三边抹掉。
   */
  effective?: CellFormat;
  /**
   * 按 `tblLook` 与单元格位置**算出来**的条件，`effective` 用的就是它。表格没有样式时为空数组
   * 也照样给出（开关与位置仍然有定义，只是没有格式可套）。
   */
  conditions?: TableConditionName[];
  /**
   * `w:tcPr/w:cnfStyle`：Word 存盘时**记下**的条件。它是缓存——Word 打开时按 `tblLook` 重算，
   * 所以渲染以 `conditions` 为准；两者不一致说明文件被别的程序改过结构或开关而没刷新这个缓存。
   * 只读：写回一个过期缓存没有意义。
   */
  recordedConditions?: TableConditionName[];
}

export interface TableRowInfo {
  cells: TableCellInfo[];
  /** 行自己的 `w:trPr`，不含表格样式带来的部分。 */
  format?: RowFormat;
  /**
   * 算上表格样式之后的行格式。只有**由行位置决定**的条件参与（`firstRow` / `lastRow` /
   * `band*Horz`）——`firstCol` 之类是单元格范围的条件，对整行没有意义。
   */
  effective?: RowFormat;
  /** 同 `TableCellInfo.conditions`，只含由行位置决定的条件。 */
  conditions?: TableConditionName[];
  /** 同 `TableCellInfo.recordedConditions`，读 `w:trPr/w:cnfStyle`。 */
  recordedConditions?: TableConditionName[];
}

export interface TableInfo {
  index: number;
  rows: TableRowInfo[];
  /** 表格自己的 `w:tblPr`，不含表格样式带来的部分。 */
  format?: TableFormat;
  /**
   * 算上表格样式之后的表格级格式。**条件不参与**——`firstRow` 之类描述的是表格里的一块区域，
   * 对表格元素自身没有意义；条件 `tblPr` 的边框与边距落在 `TableCellInfo.effective` 上。
   */
  effective?: TableFormat;
  grid: number[];
}

export type ContentControlKind =
  | 'text' | 'richText' | 'dropDownList' | 'comboBox'
  | 'date' | 'checkbox' | 'picture' | 'group' | 'unknown';

export interface ContentControlInfo {
  id?: number;
  kind: ContentControlKind;
  alias?: string;
  tag?: string;
  lock?: 'sdtLocked' | 'contentLocked' | 'sdtContentLocked' | 'unlocked';
  showingPlaceholder: boolean;
  placeholderDocPart?: string;
  items?: { displayText: string; value: string }[];
  checked?: boolean;
  dateFormat?: string;
  dataBinding?: { prefixMappings?: string; xpath?: string; storeItemId?: string };
  paragraphs: number[];
  nested: boolean;
  text: string;
}

export interface TableCellLocation {
  table: number;
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
  nested: boolean;
}

export interface ImageInfo {
  id: string;
  paragraph: number;
  run: number;
  ordinal?: number;
  sourcePartPath?: string;
  relationshipId: string;
  partPath?: string;
  contentType?: string;
  widthEmu: number;
  heightEmu: number;
  widthPx: number;
  heightPx: number;
  name?: string;
  alt?: string;
  title?: string;
  placement: 'inline' | 'floating';
  wrap?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  crop?: { left: number; top: number; right: number; bottom: number };
  isExternal: boolean;
  behindDoc?: boolean;
}

export type ShapeKind = 'textbox' | 'shape' | 'smartArt' | 'chart' | 'ole' | 'unknown';

export type CustomGeometryCommand =
  | { type: 'moveTo' | 'lnTo'; x: number; y: number }
  | { type: 'cubicBezTo'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { type: 'close' };

export interface CustomGeometry {
  width: number;
  height: number;
  commands: CustomGeometryCommand[];
}

export interface ShapeChildInfo {
  offsetXPx: number;
  offsetYPx: number;
  widthPx: number;
  heightPx: number;
  geometry?: string;
  customGeometry?: CustomGeometry;
  fill?: ShapeInfo['fill'];
  line?: ShapeInfo['line'];
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  text?: string;
}

export interface ChartInfo {
  kind: 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'unsupported';
  title?: string;
  categories: string[];
  series: Array<{ name?: string; values: Array<number | null>; xValues?: Array<number | null>; pointFills?: Array<ShapeInfo['fill'] | undefined>; fill?: ShapeInfo['fill']; line?: ShapeInfo['line'] }>;
  barDirection?: 'col' | 'bar';
  grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  legend?: { position: 'l' | 'r' | 't' | 'b' | 'tr' };
  axes?: { category?: { visible: boolean }; value?: { visible: boolean; majorGridlines: boolean } };
}

export interface ShapeInfo {
  id: string;
  paragraph: number;
  run: number;
  kind: ShapeKind;
  form: 'drawingml' | 'vml';
  name?: string;
  alt?: string;
  title?: string;
  widthPx: number;
  heightPx: number;
  placement: 'inline' | 'floating';
  wrap?: ImageInfo['wrap'];
  hasTextContent: boolean;
  geometry?: string;
  fill?: {
    type: 'none' | 'solid' | 'gradient' | 'picture';
    color?: string;
    stops?: Array<{ position: number; color: string }>;
    angle?: number;
    imagePartPath?: string;
  };
  line?: { color?: string; widthPx?: number; dash?: string };
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  adjustments?: Array<{ name: string; value: number }>;
  customGeometry?: CustomGeometry;
  children?: ShapeChildInfo[];
  chart?: ChartInfo;
}

export type SectionType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';

export interface SectionInfo {
  index: number;
  startParagraph: number;
  endParagraph: number;
  isImplicit?: boolean;
  type: SectionType;
  pageWidth: number;
  pageHeight: number;
  orientation: 'portrait' | 'landscape';
  margins: { top: number; right: number; bottom: number; left: number; header: number; footer: number; gutter: number };
  columns: { count: number; space: number; equalWidth: boolean; widths?: number[] };
  docGrid?: {
    type: 'default' | 'lines' | 'linesAndChars' | 'snapToChars';
    linePitch?: number;
    charSpace?: number;
  };
  pageNumbering?: { start?: number; format?: string };
  lineNumbering?: {
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: 'continuous' | 'newPage' | 'newSection';
  };
  pageBorders?: {
    display?: 'allPages' | 'firstPage' | 'notFirstPage';
    offsetFrom?: 'page' | 'text';
    top?: BorderSide;
    left?: BorderSide;
    bottom?: BorderSide;
    right?: BorderSide;
  };
  verticalAlignment?: 'top' | 'center' | 'both' | 'bottom';
  titlePage: boolean;
  headers: Partial<Record<'default' | 'first' | 'even', string>>;
  footers: Partial<Record<'default' | 'first' | 'even', string>>;
}

export interface PageSetup {
  type?: SectionType;
  pageWidth?: number;
  pageHeight?: number;
  orientation?: 'portrait' | 'landscape';
  margins?: Partial<SectionInfo['margins']>;
  columns?: Partial<SectionInfo['columns']>;
  pageNumbering?: SectionInfo['pageNumbering'];
  titlePage?: boolean;
}

export type DocumentBlock =
  | { type: 'paragraph'; paragraph: ParagraphInfo }
  | { type: 'table'; rows: TableRowInfo[]; format?: TableFormat; effective?: TableFormat; grid: number[] }
  | { type: 'sectionBreak'; section: number; breakType: SectionType }
  | { type: 'pageBreak' };

export interface HyperlinkInfo {
  paragraph: number;
  runs: number[];
  text: string;
  url?: string;
  anchor?: string;
  tooltip?: string;
  isExternal: boolean;
  unsafe: boolean;
  relationshipId?: string;
}

export interface BookmarkInfo {
  id: number;
  name: string;
  startParagraph: number;
  endParagraph: number;
  isInternal: boolean;
}

export type EditableRegionEditorGroup =
  | 'none'
  | 'everyone'
  | 'administrators'
  | 'contributors'
  | 'editors'
  | 'owners'
  | 'current';

export interface EditableRegionInfo {
  id: number;
  editorGroup?: EditableRegionEditorGroup;
  editorId?: string;
  rawEditorGroup?: string;
  start: { paragraph: number; offset: number };
  end: { paragraph: number; offset: number };
  unpaired?: 'startOnly' | 'endOnly';
  text: string;
}

export interface DocumentSnapshot {
  revision: number;
  paragraphs: ParagraphInfo[];
  blocks: DocumentBlock[];
  footnotes: NoteInfo[];
  endnotes: NoteInfo[];
  comments: CommentInfo[];
  parts: string[];
  styles: StyleInfo[];
  hyperlinks: HyperlinkInfo[];
  bookmarks: BookmarkInfo[];
}

export interface RevisionMark {
  id: number;
  kind: 'insertion' | 'deletion' | 'move' | 'runFormatChange' | 'paragraphFormatChange' | 'tableFormatChange' | 'rowFormatChange' | 'cellFormatChange';
  author?: string;
  date?: string;
  move?: {
    name: string;
    side: 'from' | 'to';
    pairedId?: number;
  };
}

export interface RevisionInfo extends RevisionMark {
  paragraph: number;
  run?: number;
  deletedText?: string;
  previousFormat?: RunFormat | ParagraphFormat;
}

export type ReviewerAuthorKind = 'named' | 'unattributed' | 'empty' | 'blank';

export interface ReviewerInfo {
  kind: ReviewerAuthorKind;
  author?: string;
  initials?: string;
  revisionCount: number;
  commentCount: number;
  unresolvedCommentCount: number;
  firstDate?: string;
  lastDate?: string;
}

export type ReviewerFilterAuthor =
  | { kind: 'named'; author: string }
  | { kind: 'unattributed'; author?: undefined }
  | { kind: 'empty'; author: '' }
  | { kind: 'blank'; author: string };

export interface HistoryEntry {
  revision: number;
  label?: string;
  at: number;
}

export interface DocumentProperties {
  title?: string;
  subject?: string;
  creator?: string;
  lastModifiedBy?: string;
  keywords?: string;
  description?: string;
  category?: string;
  created?: string;
  modified?: string;
  revisionNumber?: number;
  company?: string;
  manager?: string;
}

export interface DocumentProtection {
  enabled: boolean;
  edit?: 'readOnly' | 'comments' | 'trackedChanges' | 'forms' | 'none';
  enforced?: boolean;
}

export interface NoteSettingsValue {
  pos?: 'pageBottom' | 'beneathText' | 'sectEnd' | 'docEnd';
  numFmt?: string;
  numStart?: number;
  numRestart?: 'continuous' | 'eachSect';
}

export interface NoteSettings {
  footnote: NoteSettingsValue;
  endnote: NoteSettingsValue;
}

export interface CompatibilitySettings {
  doNotUseHTMLParagraphAutoSpacing?: boolean;
  doNotUseEastAsianBreakRules?: boolean;
  doNotBreakWrappedTables?: boolean;
  useWord2002TableStyleRules?: boolean;
  compatSettings?: Array<{ name: string; uri?: string; val?: string }>;
  other?: Record<string, boolean>;
}

export interface ColorSchemeMapping {
  [slot: string]: string;
}

export interface ThemeFontLanguages {
  val?: string;
  eastAsia?: string;
  bidi?: string;
}

export interface LatentStyleException {
  name: string;
  locked?: boolean;
  uiPriority?: number;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  qFormat?: boolean;
}

export interface LatentStyles {
  defaultLockedState?: boolean;
  defaultUiPriority?: number;
  defaultSemiHidden?: boolean;
  defaultUnhideWhenUsed?: boolean;
  defaultQFormat?: boolean;
  count?: number;
  exceptions: LatentStyleException[];
}

export interface ThemeSettings {
  clrSchemeMapping?: ColorSchemeMapping;
  themeFontLang?: ThemeFontLanguages;
  latentStyles?: LatentStyles;
}

export interface NoteInfo {
  id: number;
  kind: 'footnote' | 'endnote';
  number: number;
  marker: string;
  customMark?: string;
  blocks: DocumentBlock[];
  reference: { paragraph: number; run: number };
}

export type CommentAnchor =
  | { sourcePartPath: string; paragraph: number; runs: number[] }
  | { sourcePartPath: string; startParagraph: number; endParagraph: number };

export interface CommentInfo {
  id: number;
  author?: string;
  initials?: string;
  date?: string;
  text: string;
  blocks?: DocumentBlock[];
  anchor?: CommentAnchor;
  parentId?: number;
  resolved?: boolean;
  isOrphan: boolean;
}

/** Character offsets in the main-document paragraph namespace. */
export interface TextRange {
  paragraph: number;
  start: number;
  end: number;
}

/** Cross-paragraph offsets in the main-document paragraph namespace. */
export interface DocumentRange {
  start: { paragraph: number; offset: number };
  end: { paragraph: number; offset: number };
}

export interface ClipboardImage {
  bytes: string;
  contentType: string;
  widthEmu?: number;
  heightEmu?: number;
  alt?: string;
  placement?: 'inline' | 'floating';
}

export interface ClipboardRun {
  text?: string;
  format?: RunFormat;
  hyperlink?: { url?: string; anchor?: string; tooltip?: string };
  images?: ClipboardImage[];
}

export interface ClipboardParagraph {
  runs: ClipboardRun[];
  format?: ParagraphFormat;
  numbering?: { kind: 'bullet' | 'decimal'; level?: number; listId?: number };
}

export interface ClipboardTable {
  rows: ClipboardParagraph[][];
}

export type ClipboardBlock =
  | { type: 'paragraph'; paragraph: ClipboardParagraph }
  | { type: 'table'; table: ClipboardTable };

export interface ClipboardFragment {
  version: 1;
  text: string;
  paragraphs: ClipboardParagraph[];
  blocks?: ClipboardBlock[];
}

export type AgentOperation =
  | { type: 'setTrackChanges'; enabled: boolean }
  | { type: 'setRevisionAuthor'; author: string }
  | { type: 'acceptRevision'; id: number }
  | { type: 'rejectRevision'; id: number }
  | { type: 'acceptAllRevisions'; filter?: { authors?: string[] } }
  | { type: 'rejectAllRevisions'; filter?: { authors?: string[] } }
  | { type: 'setParagraphText'; index: number; text: string }
  | { type: 'insertParagraph'; text: string; before?: number }
  | { type: 'deleteParagraph'; index: number }
  | { type: 'formatParagraph'; index: number; format: ParagraphFormat }
  | { type: 'applyParagraphStyle'; index: number; styleId: string; options?: { clearDirectFormat?: boolean } }
  | { type: 'setParagraphNumbering'; index: number; numId: number; level?: number }
  | { type: 'clearParagraphNumbering'; index: number }
  | { type: 'setParagraphLevel'; index: number; delta: number }
  | { type: 'restartNumbering'; index: number; options?: { start?: number } }
  | { type: 'continueNumbering'; index: number }
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
  | { type: 'formatRange'; range: TextRange; format: RunFormat }
  | { type: 'applyCharacterStyle'; range: TextRange; styleId: string; options?: { clearDirectFormat?: boolean } }
  | { type: 'clearRangeFormat'; range: TextRange; fields?: (keyof RunFormat)[] }
  | { type: 'formatDocumentRange'; range: DocumentRange; format: RunFormat }
  | { type: 'setOutlineLevel'; index: number; level: number | null }
  | { type: 'moveOutlineSection'; from: number; to: number }
  | { type: 'setParagraphTabs'; index: number; tabs: TabStop[] }
  | {
    type: 'setParagraphBorders';
    index: number;
    borders: Partial<Record<'top' | 'left' | 'bottom' | 'right' | 'between' | 'bar', BorderSide>>;
  }
  | { type: 'setParagraphShading'; index: number; shading: Shading }
  | { type: 'insertBreak'; paragraph: number; run: number; breakType: 'textWrapping' | 'page' | 'column' }
  | { type: 'insertSymbol'; paragraph: number; run: number; font: string; charCode: number }
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'insertTableAt'; rows: number; cols: number; before?: number; format?: TableFormat }
  | { type: 'insertTableRow'; table: number; at: number }
  | { type: 'deleteTableRow'; table: number; at: number }
  | { type: 'insertTableColumn'; table: number; at: number }
  | { type: 'deleteTableColumn'; table: number; at: number }
  | { type: 'mergeCells'; table: number; range: { row: number; col: number; rowSpan: number; colSpan: number } }
  | { type: 'splitCell'; table: number; row: number; col: number; rows: number; cols: number }
  | { type: 'formatTable'; table: number; format: TableFormat }
  | { type: 'formatTableRow'; table: number; row: number; format: RowFormat }
  | { type: 'formatCell'; table: number; row: number; col: number; format: CellFormat }
  | { type: 'setCellText'; table: number; row: number; col: number; text: string }
  | { type: 'insertHyperlink'; target: { paragraph: number; start: number; end: number }; link: { url?: string; anchor?: string; tooltip?: string } }
  | { type: 'updateHyperlink'; hyperlink: number | { paragraph: number; runs: number[]; text: string }; link: { url?: string; anchor?: string; tooltip?: string } }
  | { type: 'removeHyperlink'; hyperlink: number | { paragraph: number; runs: number[]; text: string }; options?: { keepText?: boolean } }
  | { type: 'insertBookmark'; name: string; range: { startParagraph: number; endParagraph?: number } }
  | { type: 'updateFields'; kinds?: FieldKind[]; now?: string; filename?: string }
  | { type: 'insertField'; paragraph: number; instruction: string; result?: string }
  | { type: 'deleteBookmark'; name: string }
  | { type: 'setContentControlText'; id: number; text: string }
  | { type: 'setContentControlChecked'; id: number; checked: boolean }
  | { type: 'setContentControlProperties'; id: number; patch: { alias?: string | null; tag?: string | null; lock?: ContentControlInfo['lock'] } }
  | { type: 'removeContentControl'; id: number; options?: { keepContent?: boolean } }
  | { type: 'addEditableRegion'; range: DocumentRange; options: { editorGroup?: EditableRegionEditorGroup; editorId?: string } }
  | { type: 'removeEditableRegion'; id: number }
  | { type: 'insertImage'; bytes: string; contentType: string; paragraph?: number; run?: number; widthEmu?: number; heightEmu?: number; alt?: string; placement?: 'inline' | 'floating' }
  | { type: 'replaceImageBytes'; image: string; bytes: string; contentType?: string }
  | { type: 'resizeImage'; image: string; size: { widthEmu?: number; heightEmu?: number; keepAspect?: boolean } }
  | { type: 'setImageAlt'; image: string; alt: string; title?: string }
  | { type: 'deleteImage'; image: string }
  | { type: 'setPartXml'; path: string; xml: string }
  | { type: 'insertFootnote'; paragraph: number; run: number; text: string; customMark?: string }
  | { type: 'insertEndnote'; paragraph: number; run: number; text: string; customMark?: string }
  | { type: 'setNoteText'; kind: 'footnote' | 'endnote'; id: number; text: string }
  | { type: 'deleteNote'; kind: 'footnote' | 'endnote'; id: number }
  | { type: 'convertNote'; kind: 'footnote' | 'endnote'; id: number }
  | { type: 'addComment'; range: TextRange | DocumentRange; comment: { author?: string; initials?: string; text: string } }
  | { type: 'replyComment'; parentId: number; comment: { author?: string; initials?: string; text: string } }
  | { type: 'setCommentResolved'; id: number; resolved: boolean }
  | { type: 'setCommentText'; id: number; text: string }
  | { type: 'deleteComment'; id: number; options?: { withReplies?: boolean } }
  | { type: 'defineStyle'; style: StyleInfo }
  | { type: 'updateStyle'; id: string; patch: StylePatch }
  | { type: 'deleteStyle'; id: string }
  | { type: 'undo' }
  | { type: 'redo' };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
