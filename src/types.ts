export interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
}

export interface ParagraphFormat {
  alignment?: 'left' | 'center' | 'right' | 'both';
  style?: string;
}

export interface RunInfo extends RunFormat {
  index: number;
  text: string;
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
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

export type DocumentBlock =
  | { type: 'paragraph'; paragraph: ParagraphInfo }
  | { type: 'table'; rows: TableRowInfo[]; format?: TableFormat; grid: number[] };

export interface DocumentSnapshot {
  revision: number;
  paragraphs: ParagraphInfo[];
  blocks: DocumentBlock[];
  parts: string[];
}

export type AgentOperation =
  | { type: 'setParagraphText'; index: number; text: string }
  | { type: 'insertParagraph'; text: string; before?: number }
  | { type: 'deleteParagraph'; index: number }
  | { type: 'formatParagraph'; index: number; format: ParagraphFormat }
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
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
