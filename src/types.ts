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

export type SectionType = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';

export interface SectionInfo {
  index: number;
  startParagraph: number;
  endParagraph: number;
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
  | { type: 'table'; rows: { cells: { blocks: DocumentBlock[] }[] }[] }
  | { type: 'sectionBreak'; section: number; breakType: SectionType }
  | { type: 'pageBreak' };

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
