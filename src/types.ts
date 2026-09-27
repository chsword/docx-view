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
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
