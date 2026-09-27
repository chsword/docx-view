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
  hyperlink?: { url?: string; anchor?: string; tooltip?: string; unsafe: boolean };
}

export interface ParagraphInfo extends ParagraphFormat {
  index: number;
  text: string;
  runs: RunInfo[];
}

export type DocumentBlock =
  | { type: 'paragraph'; paragraph: ParagraphInfo }
  | { type: 'table'; rows: { cells: { blocks: DocumentBlock[] }[] }[] };

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

export interface DocumentSnapshot {
  revision: number;
  paragraphs: ParagraphInfo[];
  blocks: DocumentBlock[];
  parts: string[];
  hyperlinks: HyperlinkInfo[];
  bookmarks: BookmarkInfo[];
}

export type AgentOperation =
  | { type: 'setParagraphText'; index: number; text: string }
  | { type: 'insertParagraph'; text: string; before?: number }
  | { type: 'deleteParagraph'; index: number }
  | { type: 'formatParagraph'; index: number; format: ParagraphFormat }
  | { type: 'formatRun'; paragraph: number; run: number; format: RunFormat }
  | { type: 'replaceText'; search: string; replacement: string }
  | { type: 'insertTable'; rows: string[][] }
  | { type: 'insertHyperlink'; target: { paragraph: number; start: number; end: number }; link: { url?: string; anchor?: string; tooltip?: string } }
  | { type: 'updateHyperlink'; hyperlink: HyperlinkInfo | number; link: { url?: string; anchor?: string; tooltip?: string } }
  | { type: 'removeHyperlink'; hyperlink: HyperlinkInfo | number; options?: { keepText?: boolean } }
  | { type: 'insertBookmark'; name: string; range: { startParagraph: number; endParagraph?: number } }
  | { type: 'deleteBookmark'; name: string }
  | { type: 'setPartXml'; path: string; xml: string };

export interface AgentRequest {
  expectedRevision?: number;
  operations: AgentOperation[];
}
