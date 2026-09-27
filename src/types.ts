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
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
  effective?: ParagraphFormat;
  numbering?: NumberingInfo;
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
  | { type: 'setParagraphNumbering'; index: number; numId: number; level?: number }
  | { type: 'clearParagraphNumbering'; index: number }
  | { type: 'setParagraphLevel'; index: number; delta: number }
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
