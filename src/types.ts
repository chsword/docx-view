export interface RunFormat {
  style?: string | null;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  underlineStyle?: string | null;
  underlineColor?: string | null;
  fontSize?: number | null;
  fontFamily?: string | null;
  fontFamilyEastAsia?: string | null;
  color?: string | null;
  strike?: boolean | null;
  doubleStrike?: boolean | null;
  verticalAlign?: 'baseline' | 'subscript' | 'superscript' | null;
  smallCaps?: boolean | null;
  allCaps?: boolean | null;
  highlight?: string | null;
  characterSpacing?: number | null;
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
  style?: string | null;
  indentLeft?: number | null;
  indentRight?: number | null;
  indentFirstLine?: number | null;
  indentHanging?: number | null;
  spacingBefore?: number | null;
  spacingAfter?: number | null;
  lineSpacing?: number | null;
  lineSpacingRule?: 'auto' | 'atLeast' | 'exact' | null;
  keepNext?: boolean | null;
  keepLines?: boolean | null;
  pageBreakBefore?: boolean | null;
  widowControl?: boolean | null;
  suppressLineNumbers?: boolean | null;
  suppressAutoHyphens?: boolean | null;
  outlineLevel?: number | null;
  tabs?: TabStop[] | null;
  borders?: Partial<Record<'top' | 'left' | 'bottom' | 'right' | 'between' | 'bar', BorderSide>> | null;
  shading?: Shading | null;
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
}

export interface RunInfo extends RunFormat {
  index: number;
  text: string;
  revisions?: RevisionMark[];
  effective?: RunFormat;
  hyperlink?: { url?: string; anchor?: string; tooltip?: string; unsafe: boolean };
  image?: ImageInfo;
  images?: ImageInfo[];
  noteReference?: { kind: 'footnote' | 'endnote'; id: number; number: number; marker: string };
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
  paragraphRevision?: RevisionMark;
  effective?: ParagraphFormat;
  numbering?: NumberingInfo;
  images: ImageInfo[];
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

export interface TableFormat {
  width?: WidthFormat;
  alignment?: 'left' | 'center' | 'right';
  indent?: number;
  borders?: BordersFormat;
  shading?: ShadingFormat;
  cellMargin?: MarginFormat;
  layout?: 'fixed' | 'autofit';
  style?: string;
  look?: string;
  caption?: string;
  description?: string;
}

export interface RowFormat {
  height?: { value: number; rule?: 'atLeast' | 'exact' };
  cantSplit?: boolean;
  header?: boolean;
  alignment?: 'left' | 'center' | 'right';
  deleted?: boolean;
  inserted?: boolean;
  revision?: { author?: string; date?: string };
}

export interface CellFormat {
  width?: WidthFormat;
  borders?: BordersFormat;
  shading?: ShadingFormat;
  margin?: MarginFormat;
  verticalAlign?: 'top' | 'center' | 'bottom';
  textDirection?: string;
  noWrap?: boolean;
  hideMark?: boolean;
  hMerge?: 'restart' | 'continue';
  vMerge?: 'restart' | 'continue';
}

export interface TableCellInfo {
  blocks: DocumentBlock[];
  colSpan: number;
  rowSpan: number;
  isMergeContinuation: boolean;
  format?: CellFormat;
}

export interface TableRowInfo {
  cells: TableCellInfo[];
  format?: RowFormat;
}

export interface TableInfo {
  index: number;
  rows: TableRowInfo[];
  format?: TableFormat;
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
  pageNumbering?: { start?: number; format?: string };
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
  | { type: 'table'; rows: TableRowInfo[]; format?: TableFormat; grid: number[] }
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
  | { type: 'undo' }
  | { type: 'redo' };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
