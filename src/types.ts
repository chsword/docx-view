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
  noteReference?: { kind: 'footnote' | 'endnote'; id: number; number: number; marker: string };
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
  footnotes: NoteInfo[];
  endnotes: NoteInfo[];
  parts: string[];
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

export type AgentOperation =
  | { type: 'setParagraphText'; index: number; text: string }
  | { type: 'insertParagraph'; text: string; before?: number }
  | { type: 'deleteParagraph'; index: number }
  | { type: 'formatParagraph'; index: number; format: ParagraphFormat }
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'setPartXml'; path: string; xml: string }
  | { type: 'insertFootnote'; paragraph: number; run: number; text: string; customMark?: string }
  | { type: 'insertEndnote'; paragraph: number; run: number; text: string; customMark?: string }
  | { type: 'setNoteText'; kind: 'footnote' | 'endnote'; id: number; text: string }
  | { type: 'deleteNote'; kind: 'footnote' | 'endnote'; id: number }
  | { type: 'convertNote'; kind: 'footnote' | 'endnote'; id: number };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
