export interface RunFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
  border?: BorderSide | null;
  shading?: Shading | null;
}

export interface TabStop {
  position: number;
  alignment: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num';
  leader?: string;
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
  alignment?: 'left' | 'center' | 'right' | 'both';
  style?: string;
  tabs?: TabStop[] | null;
  borders?: Partial<Record<'top' | 'left' | 'bottom' | 'right' | 'between' | 'bar', BorderSide>> | null;
  shading?: Shading | null;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  suppressLineNumbers?: boolean;
  suppressAutoHyphens?: boolean;
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

export type DocumentBlock =
  | { type: 'paragraph'; paragraph: ParagraphInfo }
  | { type: 'table'; rows: { cells: { blocks: DocumentBlock[] }[] }[] };

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
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
