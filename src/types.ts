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
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
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
