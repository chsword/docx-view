export interface RunFormat {
  style?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  underlineStyle?: string;
  underlineColor?: string;
  fontSize?: number;
  fontFamily?: string;
  fontFamilyEastAsia?: string;
  color?: string;
  strike?: boolean;
  doubleStrike?: boolean;
  verticalAlign?: 'baseline' | 'subscript' | 'superscript';
  smallCaps?: boolean;
  allCaps?: boolean;
  highlight?: string;
  characterSpacing?: number;
}

export interface ParagraphFormat {
  alignment?: 'left' | 'center' | 'right' | 'both' | 'distribute';
  style?: string;
  indentLeft?: number;
  indentRight?: number;
  indentFirstLine?: number;
  indentHanging?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  lineSpacing?: number;
  lineSpacingRule?: 'auto' | 'atLeast' | 'exact';
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  outlineLevel?: number;
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

export interface RunInfo extends RunFormat {
  index: number;
  text: string;
  effective?: RunFormat;
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
  effective?: ParagraphFormat;
}

export type DocumentBlock =
  | { type: 'paragraph'; paragraph: ParagraphInfo }
  | { type: 'table'; rows: { cells: { blocks: DocumentBlock[] }[] }[] };

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
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
