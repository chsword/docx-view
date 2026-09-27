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
  outlineLevel?: number | null;
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
  quickFormat?: boolean;
  paragraph?: ParagraphFormat;
  run?: RunFormat;
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
  effective?: RunFormat;
  image?: ImageInfo;
  images?: ImageInfo[];
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
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

export interface DocumentSnapshot {
  revision: number;
  paragraphs: ParagraphInfo[];
  blocks: DocumentBlock[];
  parts: string[];
  styles: StyleInfo[];
}

export type AgentOperation =
  | { type: 'setParagraphText'; index: number; text: string }
  | { type: 'insertParagraph'; text: string; before?: number }
  | { type: 'deleteParagraph'; index: number }
  | { type: 'formatParagraph'; index: number; format: ParagraphFormat }
  | { type: 'setParagraphNumbering'; index: number; numId: number; level?: number }
  | { type: 'clearParagraphNumbering'; index: number }
  | { type: 'setParagraphLevel'; index: number; delta: number }
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
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
  | { type: 'insertImage'; bytes: string; contentType: string; paragraph?: number; run?: number; widthEmu?: number; heightEmu?: number; alt?: string; placement?: 'inline' | 'floating' }
  | { type: 'replaceImageBytes'; image: string; bytes: string; contentType?: string }
  | { type: 'resizeImage'; image: string; size: { widthEmu?: number; heightEmu?: number; keepAspect?: boolean } }
  | { type: 'setImageAlt'; image: string; alt: string; title?: string }
  | { type: 'deleteImage'; image: string }
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
